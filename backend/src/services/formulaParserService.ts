import { CellReference } from "../models/graph";
import {
  expandRange,
  normalizeCellAddress,
  normalizeFileName,
  normalizeSheetName,
  toNodeId
} from "../utils/cellUtils";

const RANGE_REGEX = /((?:'[^']+'|\[[^\]]+\][^!]+|[A-Za-z0-9_\.]+)!\$?[A-Z]{1,3}\$?[0-9]+|\$?[A-Z]{1,3}\$?[0-9]+):((?:'[^']+'|\[[^\]]+\][^!]+|[A-Za-z0-9_\.]+)!\$?[A-Z]{1,3}\$?[0-9]+|\$?[A-Z]{1,3}\$?[0-9]+)/g;
const REFERENCE_REGEX = /(?:'[^']+'|\[[^\]]+\][^!]+|[A-Za-z0-9_\.]+)!\$?[A-Z]{1,3}\$?[0-9]+|\$?[A-Z]{1,3}\$?[0-9]+/g;

function parseSheetToken(sheetToken: string, currentFileName: string): { file: string; sheet: string; external: boolean } {
  const cleaned = sheetToken.trim().replace(/^'|'$/g, "");
  const externalMatch = cleaned.match(/^(?:.*[\\/])?\[([^\]]+)\](.+)$/);
  if (externalMatch) {
    return {
      file: normalizeFileName(externalMatch[1]),
      sheet: normalizeSheetName(externalMatch[2]),
      external: true
    };
  }

  return {
    file: normalizeFileName(currentFileName),
    sheet: normalizeSheetName(cleaned),
    external: false
  };
}

function parseRef(ref: string, currentSheet: string, currentFileName: string): CellReference {
  const clean = ref.trim();
  if (clean.includes("!")) {
    const [sheetPart, cellPart] = clean.split("!");
    const sheetParsed = parseSheetToken(sheetPart, currentFileName);
    return {
      file: sheetParsed.file,
      sheet: sheetParsed.sheet,
      cell: normalizeCellAddress(cellPart),
      external: sheetParsed.external,
      original: clean
    };
  }

  return {
    file: normalizeFileName(currentFileName),
    sheet: currentSheet,
    cell: normalizeCellAddress(clean),
    external: false,
    original: clean
  };
}

export class FormulaParserService {
  extractReferences(formula: string | undefined, currentSheet: string, currentFileName: string): CellReference[] {
    if (!formula || !formula.startsWith("=")) {
      return [];
    }

    const body = formula.slice(1);
    const refs = new Map<string, CellReference>();

    // Parse ranges first so we can avoid reprocessing the same refs as single tokens.
    for (const match of body.matchAll(RANGE_REGEX)) {
      const left = parseRef(match[1], currentSheet, currentFileName);
      const right = parseRef(match[2], currentSheet, currentFileName);

      if (left.file !== right.file || left.sheet !== right.sheet) {
        refs.set(`${left.file}::${left.sheet}::${left.cell}`, left);
        refs.set(`${right.file}::${right.sheet}::${right.cell}`, right);
        continue;
      }

      const expanded = expandRange(left.cell, right.cell);
      for (const cell of expanded) {
        const expandedRef: CellReference = {
          file: left.file,
          sheet: left.sheet,
          cell,
          external: left.external,
          original: match[0]
        };
        refs.set(`${expandedRef.file}::${expandedRef.sheet}::${expandedRef.cell}`, expandedRef);
      }
    }

    const bodyWithoutRanges = body.replace(RANGE_REGEX, "");
    for (const token of bodyWithoutRanges.match(REFERENCE_REGEX) ?? []) {
      const parsed = parseRef(token, currentSheet, currentFileName);
      refs.set(`${parsed.file}::${parsed.sheet}::${parsed.cell}`, parsed);
    }

    return [...refs.values()];
  }

  extractDependencies(formula: string | undefined, currentSheet: string, currentFileName: string): string[] {
    return this.extractReferences(formula, currentSheet, currentFileName).map((ref) =>
      toNodeId(ref.file, ref.sheet, ref.cell)
    );
  }
}
