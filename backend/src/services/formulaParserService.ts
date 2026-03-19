import { expandRange, normalizeCellAddress, normalizeSheetName, toNodeId } from "../utils/cellUtils";

const RANGE_REGEX = /((?:'[^']+'|[A-Za-z0-9_]+)!\$?[A-Z]{1,3}\$?[0-9]+|\$?[A-Z]{1,3}\$?[0-9]+):((?:'[^']+'|[A-Za-z0-9_]+)!\$?[A-Z]{1,3}\$?[0-9]+|\$?[A-Z]{1,3}\$?[0-9]+)/g;
const REFERENCE_REGEX = /(?:'[^']+'|[A-Za-z0-9_]+)!\$?[A-Z]{1,3}\$?[0-9]+|\$?[A-Z]{1,3}\$?[0-9]+/g;

function parseRef(ref: string, currentSheet: string): { sheet: string; cell: string } {
  const clean = ref.trim();
  if (clean.includes("!")) {
    const [sheet, cell] = clean.split("!");
    return {
      sheet: normalizeSheetName(sheet),
      cell: normalizeCellAddress(cell)
    };
  }

  return {
    sheet: currentSheet,
    cell: normalizeCellAddress(clean)
  };
}

export class FormulaParserService {
  extractDependencies(formula: string | undefined, currentSheet: string): string[] {
    if (!formula || !formula.startsWith("=")) {
      return [];
    }

    const body = formula.slice(1);
    const deps = new Set<string>();

    // Parse ranges first so we can avoid reprocessing the same refs as single tokens.
    for (const match of body.matchAll(RANGE_REGEX)) {
      const left = parseRef(match[1], currentSheet);
      const right = parseRef(match[2], currentSheet);

      if (left.sheet !== right.sheet) {
        deps.add(toNodeId(left.sheet, left.cell));
        deps.add(toNodeId(right.sheet, right.cell));
        continue;
      }

      const expanded = expandRange(left.cell, right.cell);
      for (const cell of expanded) {
        deps.add(toNodeId(left.sheet, cell));
      }
    }

    const bodyWithoutRanges = body.replace(RANGE_REGEX, "");
    for (const token of bodyWithoutRanges.match(REFERENCE_REGEX) ?? []) {
      const parsed = parseRef(token, currentSheet);
      deps.add(toNodeId(parsed.sheet, parsed.cell));
    }

    return [...deps];
  }
}
