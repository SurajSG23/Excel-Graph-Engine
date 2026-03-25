import { CellReference, GraphEdge, GraphNode, WorkbookGraph, WorkbookRole } from "../types/workbook";

const TOKEN_REGEX = /((?:'[^']+'|\[[^\]]+\][^!]+|[A-Za-z0-9_.]+)!\$?[A-Z]{1,3}\$?[0-9]+|\$?[A-Z]{1,3}\$?[0-9]+)(:((?:'[^']+'|\[[^\]]+\][^!]+|[A-Za-z0-9_.]+)!\$?[A-Z]{1,3}\$?[0-9]+|\$?[A-Z]{1,3}\$?[0-9]+))?/g;

export interface GroupInstanceMapping {
  outputNodeId: string;
  inputNodeIds: string[];
}

export interface GroupedGraphNode {
  id: string;
  nodeType: "group";
  fileName: string;
  fileRole: WorkbookRole;
  sheet: string;
  cell: string;
  formula?: string;
  formulaTemplate: string;
  formulaSignature: string;
  value?: string | number | boolean;
  dependencies: string[];
  referenceDetails: CellReference[];
  memberNodeIds: string[];
  inputs: string[];
  outputs: string[];
  inputOutputMapping: GroupInstanceMapping[];
}

export type GraphViewNode = GraphNode | GroupedGraphNode;

export interface GraphProjection {
  nodes: GraphViewNode[];
  edges: GraphEdge[];
  nodeToGroupId: Map<string, string>;
  groupsById: Map<string, GroupedGraphNode>;
}

interface FormulaSignature {
  signature: string;
  template: string;
}

interface ParsedAddress {
  column: number;
  row: number;
  colAbs: boolean;
  rowAbs: boolean;
}

interface ParsedToken {
  file: string;
  sheet: string;
  explicitSheet: boolean;
  external: boolean;
  address: ParsedAddress;
}

const projectionCache = new Map<string, GraphProjection>();

function parseAddress(addressToken: string): ParsedAddress | null {
  const match = addressToken.toUpperCase().match(/^(\$?)([A-Z]{1,3})(\$?)([0-9]+)$/);
  if (!match) {
    return null;
  }

  return {
    colAbs: Boolean(match[1]),
    column: columnToIndex(match[2]),
    rowAbs: Boolean(match[3]),
    row: Number(match[4])
  };
}

function columnToIndex(column: string): number {
  let value = 0;
  for (let i = 0; i < column.length; i += 1) {
    value = value * 26 + (column.charCodeAt(i) - 64);
  }
  return value;
}

function parseRefToken(rawToken: string, fallbackFile: string, fallbackSheet: string): ParsedToken | null {
  const clean = rawToken.trim();

  if (!clean.includes("!")) {
    const address = parseAddress(clean);
    if (!address) {
      return null;
    }

    return {
      file: fallbackFile,
      sheet: fallbackSheet,
      explicitSheet: false,
      external: false,
      address
    };
  }

  const [sheetPartRaw, cellPartRaw] = clean.split("!");
  const sheetPart = sheetPartRaw.trim().replace(/^'|'$/g, "");
  const address = parseAddress(cellPartRaw.trim());
  if (!address) {
    return null;
  }

  const externalMatch = sheetPart.match(/^(?:.*[\\/])?\[([^\]]+)\](.+)$/);
  if (externalMatch) {
    return {
      file: externalMatch[1].trim().toUpperCase(),
      sheet: externalMatch[2].trim().toUpperCase(),
      explicitSheet: true,
      external: true,
      address
    };
  }

  return {
    file: fallbackFile,
    sheet: sheetPart.toUpperCase(),
    explicitSheet: true,
    external: false,
    address
  };
}

function parseCellAddress(cell: string): ParsedAddress | null {
  return parseAddress(cell.toUpperCase());
}

function placeholderLabel(index: number): string {
  const alphabetIndex = index % 26;
  const suffix = index >= 26 ? String(Math.floor(index / 26)) : "";
  return `${String.fromCharCode(65 + alphabetIndex)}x${suffix}`;
}

function tokenPattern(token: ParsedToken, anchor: ParsedAddress): string {
  const rowDelta = token.address.row - anchor.row;
  const colDelta = token.address.column - anchor.column;
  const sheetPart = token.explicitSheet ? `|s:${token.sheet}` : "|s:_";
  const filePart = token.external ? `|f:${token.file}` : "|f:_";

  return [
    token.external ? "ext" : "local",
    sheetPart,
    filePart,
    `|dr:${rowDelta}`,
    `|dc:${colDelta}`,
    `|ra:${token.address.rowAbs ? 1 : 0}`,
    `|ca:${token.address.colAbs ? 1 : 0}`
  ].join("");
}

function computeFormulaSignature(node: GraphNode): FormulaSignature | null {
  const formula = node.formula?.trim();
  if (!formula || !formula.startsWith("=")) {
    return null;
  }

  const anchor = parseCellAddress(node.cell);
  if (!anchor) {
    return null;
  }

  const patternToIndex = new Map<string, number>();
  const body = formula.slice(1);

  const normalizedBody = body.replace(
    TOKEN_REGEX,
    (matched, firstToken: string, _rangePart: string, secondToken?: string): string => {
      const first = parseRefToken(firstToken, node.fileName.toUpperCase(), node.sheet.toUpperCase());
      if (!first) {
        return matched;
      }

      const firstPattern = tokenPattern(first, anchor);
      const finalPattern = secondToken
        ? (() => {
            const second = parseRefToken(secondToken, node.fileName.toUpperCase(), node.sheet.toUpperCase());
            if (!second) {
              return `single:${firstPattern}`;
            }
            return `range:${firstPattern}::${tokenPattern(second, anchor)}`;
          })()
        : `single:${firstPattern}`;

      if (!patternToIndex.has(finalPattern)) {
        patternToIndex.set(finalPattern, patternToIndex.size);
      }

      return placeholderLabel(patternToIndex.get(finalPattern) ?? 0);
    }
  );

  return {
    signature: normalizedBody.replace(/\s+/g, "").toUpperCase(),
    template: `=${normalizedBody}`
  };
}

function collectUnique<T>(items: T[]): T[] {
  return [...new Set(items)];
}

function mergeReferences(nodes: GraphNode[]): CellReference[] {
  const merged = new Map<string, CellReference>();
  for (const node of nodes) {
    for (const reference of node.referenceDetails) {
      const key = `${reference.file}::${reference.sheet}::${reference.cell}`;
      if (!merged.has(key)) {
        merged.set(key, reference);
      }
    }
  }
  return [...merged.values()];
}

function toGroupId(signature: string, firstNode: GraphNode): string {
  const compact = signature.replace(/[^A-Z0-9]/gi, "").slice(0, 24) || "PATTERN";
  return `group:${firstNode.fileName}::${firstNode.sheet}::${compact}`;
}

export function isGroupedNode(node: GraphViewNode): node is GroupedGraphNode {
  return "nodeType" in node && node.nodeType === "group";
}

export function projectGraphForFormulaGrouping(
  workbook: WorkbookGraph | null,
  groupSimilarFormulas: boolean
): GraphProjection {
  if (!workbook) {
    return {
      nodes: [],
      edges: [],
      nodeToGroupId: new Map(),
      groupsById: new Map()
    };
  }

  const cacheKey = `${workbook.workbookId}:${workbook.version}:${groupSimilarFormulas ? "group" : "flat"}`;
  const cached = projectionCache.get(cacheKey);
  if (cached) {
    return cached;
  }

  if (!groupSimilarFormulas) {
    const flatProjection: GraphProjection = {
      nodes: workbook.nodes,
      edges: workbook.edges,
      nodeToGroupId: new Map(),
      groupsById: new Map()
    };
    projectionCache.set(cacheKey, flatProjection);
    return flatProjection;
  }

  const signatureBuckets = new Map<string, GraphNode[]>();
  const templateBySignature = new Map<string, string>();

  for (const node of workbook.nodes) {
    const signature = computeFormulaSignature(node);
    if (!signature) {
      continue;
    }

    if (!signatureBuckets.has(signature.signature)) {
      signatureBuckets.set(signature.signature, []);
    }

    signatureBuckets.get(signature.signature)?.push(node);
    if (!templateBySignature.has(signature.signature)) {
      templateBySignature.set(signature.signature, signature.template);
    }
  }

  const nodeToGroupId = new Map<string, string>();
  const groupsById = new Map<string, GroupedGraphNode>();

  for (const [signature, nodes] of signatureBuckets) {
    if (nodes.length < 2) {
      continue;
    }

    const ordered = [...nodes].sort((left, right) => left.id.localeCompare(right.id));
    const anchor = ordered[0];
    const groupId = toGroupId(signature, anchor);
    const mergedReferences = mergeReferences(ordered);

    for (const node of ordered) {
      nodeToGroupId.set(node.id, groupId);
    }

    const mapping: GroupInstanceMapping[] = ordered.map((node) => ({
      outputNodeId: node.id,
      inputNodeIds: [...node.dependencies]
    }));

    groupsById.set(groupId, {
      id: groupId,
      nodeType: "group",
      fileName: anchor.fileName,
      fileRole: anchor.fileRole,
      sheet: anchor.sheet,
      cell: anchor.cell,
      formula: templateBySignature.get(signature) ?? anchor.formula,
      formulaTemplate: templateBySignature.get(signature) ?? anchor.formula ?? "",
      formulaSignature: signature,
      value: undefined,
      dependencies: [],
      referenceDetails: mergedReferences,
      memberNodeIds: ordered.map((node) => node.id),
      inputs: collectUnique(ordered.flatMap((node) => node.dependencies)),
      outputs: ordered.map((node) => node.id),
      inputOutputMapping: mapping
    });
  }

  if (groupsById.size === 0) {
    const fallbackProjection: GraphProjection = {
      nodes: workbook.nodes,
      edges: workbook.edges,
      nodeToGroupId: new Map(),
      groupsById: new Map()
    };
    projectionCache.set(cacheKey, fallbackProjection);
    return fallbackProjection;
  }

  const projectedNodes: GraphViewNode[] = [];
  for (const node of workbook.nodes) {
    if (!nodeToGroupId.has(node.id)) {
      projectedNodes.push(node);
    }
  }
  projectedNodes.push(...groupsById.values());

  const edgeMap = new Map<string, GraphEdge>();
  for (const edge of workbook.edges) {
    const mappedSource = nodeToGroupId.get(edge.source) ?? edge.source;
    const mappedTarget = nodeToGroupId.get(edge.target) ?? edge.target;

    if (mappedSource === mappedTarget) {
      continue;
    }

    const key = `${mappedSource}->${mappedTarget}`;
    if (!edgeMap.has(key)) {
      edgeMap.set(key, {
        source: mappedSource,
        target: mappedTarget
      });
    }
  }

  const projectedEdges = [...edgeMap.values()];
  const incomingDependencies = new Map<string, Set<string>>();
  for (const edge of projectedEdges) {
    if (!incomingDependencies.has(edge.target)) {
      incomingDependencies.set(edge.target, new Set());
    }
    incomingDependencies.get(edge.target)?.add(edge.source);
  }

  const hydratedNodes = projectedNodes.map((node) => {
    const deps = [...(incomingDependencies.get(node.id) ?? new Set<string>())];
    if (isGroupedNode(node)) {
      return {
        ...node,
        dependencies: deps
      };
    }

    return {
      ...node,
      dependencies: deps
    };
  });

  const projection: GraphProjection = {
    nodes: hydratedNodes,
    edges: projectedEdges,
    nodeToGroupId,
    groupsById
  };

  projectionCache.set(cacheKey, projection);

  if (projectionCache.size > 18) {
    const firstKey = projectionCache.keys().next().value;
    if (firstKey) {
      projectionCache.delete(firstKey);
    }
  }

  return projection;
}
