import { GraphEdge, GraphNode, ParsedWorkbook } from "../models/graph";
import { FormulaParserService } from "./formulaParserService";
import { normalizeCellAddress, toNodeId } from "../utils/cellUtils";

export interface BuildResult {
  nodes: GraphNode[];
  edges: GraphEdge[];
  sheets: string[];
  files: Array<{
    fileName: string;
    role: "input" | "output" | "other";
    sheets: string[];
    uploadName: string;
  }>;
}

export class GraphBuilderService {
  constructor(private readonly formulaParserService: FormulaParserService) {}

  private resolveExternalFileToken(
    fileToken: string,
    currentFileName: string,
    files: Array<{ fileName: string; role: "input" | "output" | "other" }>
  ): string {
    if (!/^\d+$/.test(fileToken)) {
      return fileToken;
    }

    const rankedCandidates = files
      .filter((file) => file.fileName !== currentFileName)
      .sort((left, right) => {
        const rank = (role: "input" | "output" | "other"): number => {
          if (role === "input") return 0;
          if (role === "output") return 1;
          return 2;
        };

        const byRole = rank(left.role) - rank(right.role);
        if (byRole !== 0) {
          return byRole;
        }
        return left.fileName.localeCompare(right.fileName);
      });

    if (rankedCandidates.length === 0) {
      return fileToken;
    }

    const index = Number(fileToken) - 1;
    if (index >= 0 && index < rankedCandidates.length) {
      return rankedCandidates[index].fileName;
    }

    if (rankedCandidates.length === 1) {
      return rankedCandidates[0].fileName;
    }

    return fileToken;
  }

  buildFromWorkbooks(workbooks: ParsedWorkbook[]): BuildResult {
    const filesMeta = workbooks.map((workbook) => ({
      fileName: workbook.fileName,
      role: workbook.fileRole
    }));

    const nodes = workbooks.flatMap((workbook) =>
      workbook.cells.map((cell) => {
        const id = toNodeId(cell.fileName, cell.sheet, cell.cell);
        const extractedReferences = this.formulaParserService.extractReferences(
          cell.formula,
          cell.sheet,
          cell.fileName
        );
        const referenceDetails = extractedReferences.map((ref) => {
          if (!ref.external) {
            return ref;
          }

          return {
            ...ref,
            file: this.resolveExternalFileToken(ref.file, cell.fileName, filesMeta)
          };
        });

        return {
          id,
          fileName: cell.fileName,
          fileRole: cell.fileRole,
          sheet: cell.sheet,
          cell: normalizeCellAddress(cell.cell),
          formula: cell.formula,
          value: cell.value,
          referenceDetails,
          dependencies: referenceDetails.map((ref) => toNodeId(ref.file, ref.sheet, ref.cell))
        } satisfies GraphNode;
      })
    );

    const uniqueSheets = new Set<string>();
    for (const workbook of workbooks) {
      for (const sheet of workbook.sheets) {
        uniqueSheets.add(`${workbook.fileName}::${sheet}`);
      }
    }

    return {
      nodes,
      edges: this.buildEdges(nodes),
      sheets: [...uniqueSheets],
      files: workbooks.map((workbook) => ({
        fileName: workbook.fileName,
        role: workbook.fileRole,
        sheets: workbook.sheets,
        uploadName: workbook.uploadName
      }))
    };
  }

  rebuildFromNodes(nodes: GraphNode[], files: BuildResult["files"]): BuildResult {
    const filesMeta = files.map((file) => ({
      fileName: file.fileName,
      role: file.role
    }));

    const rebuiltNodes = nodes.map((node) => {
      const extractedReferences = this.formulaParserService.extractReferences(
        node.formula,
        node.sheet,
        node.fileName
      );
      const referenceDetails = extractedReferences.map((ref) => {
        if (!ref.external) {
          return ref;
        }

        return {
          ...ref,
          file: this.resolveExternalFileToken(ref.file, node.fileName, filesMeta)
        };
      });

      return {
        ...node,
        referenceDetails,
        dependencies: referenceDetails.map((ref) => toNodeId(ref.file, ref.sheet, ref.cell))
      };
    });

    const uniqueSheets = new Set<string>();
    for (const node of rebuiltNodes) {
      uniqueSheets.add(`${node.fileName}::${node.sheet}`);
    }

    return {
      nodes: rebuiltNodes,
      edges: this.buildEdges(rebuiltNodes),
      sheets: [...uniqueSheets],
      files
    };
  }

  private buildEdges(nodes: GraphNode[]): GraphEdge[] {
    const edges: GraphEdge[] = [];
    for (const node of nodes) {
      for (const dep of node.dependencies) {
        edges.push({
          source: dep,
          target: node.id
        });
      }
    }
    return edges;
  }
}
