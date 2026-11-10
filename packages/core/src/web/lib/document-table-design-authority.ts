/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { Editor } from "@tiptap/core";
import { closeHistory } from "@tiptap/pm/history";
import type { Node as ProseMirrorNode } from "@tiptap/pm/model";
import type { Selection } from "@tiptap/pm/state";

import {
  officeWordTableStyleByName,
  officeWordTableStyleForAllocatedId,
  type OfficeWordDefinedStyle,
  type OfficeWordTableStyle,
} from "./office-table-style-word";

const EDGE_KEYS = [
  "top",
  "bottom",
  "left",
  "right",
  "insideH",
  "insideV",
] as const;
const CONDITION_KEYS = [
  "band1Horz",
  "band2Horz",
  "band1Vert",
  "band2Vert",
  "firstCol",
  "lastCol",
  "firstRow",
  "lastRow",
  "nwCell",
  "neCell",
  "swCell",
  "seCell",
] as const;

export type DocumentTableVisualEdge = (typeof EDGE_KEYS)[number];
export type DocumentTableVisualCondition = (typeof CONDITION_KEYS)[number];

export interface DocumentTableStyleVisualBorder {
  readonly style: string;
  readonly sizePt?: number;
  readonly color?: string;
}

export interface DocumentTableStyleVisualLayer {
  readonly fill?: string;
  /** Text colour the style's `w:rPr/w:color` gives the region's runs. */
  readonly color?: string;
  /** The style's `w:rPr/w:b` for the region; direct run formatting wins. */
  readonly bold?: boolean;
  /** The style's `w:rPr/w:sz` for the region, in points (R60-D15: a header
   * row at 16 pt); a run's own size wins. */
  readonly fontSizePt?: number;
  readonly tableBorders: Readonly<
    Partial<Record<DocumentTableVisualEdge, DocumentTableStyleVisualBorder>>
  >;
  readonly cellBorders: Readonly<
    Partial<Record<DocumentTableVisualEdge, DocumentTableStyleVisualBorder>>
  >;
  readonly clearedTableBorders: readonly DocumentTableVisualEdge[];
  readonly clearedCellBorders: readonly DocumentTableVisualEdge[];
}

export interface DocumentTableStyleVisualAuthority {
  readonly schemaVersion: 1;
  readonly styleId: string;
  readonly wordWrapOmissionChainSha256?: string;
  readonly rowBandSize: number;
  readonly columnBandSize: number;
  readonly wholeTable: DocumentTableStyleVisualLayer;
  readonly conditionals: Readonly<
    Partial<Record<DocumentTableVisualCondition, DocumentTableStyleVisualLayer>>
  >;
}

export interface DocumentTableStyleVisualCatalogEntry {
  readonly styleId: string;
  readonly name: string;
  readonly visual: DocumentTableStyleVisualAuthority;
  /** Drawn for tables that already use it, never offered by Table Design: a
   * custom style, or one whose `basedOn` names a style the file lacks. */
  readonly displayOnly?: true;
}

export interface DocumentTableLook {
  readonly firstRow: boolean;
  readonly lastRow: boolean;
  readonly firstColumn: boolean;
  readonly lastColumn: boolean;
  readonly bandedRows: boolean;
  readonly bandedColumns: boolean;
}

export interface DocumentTableDesignState {
  readonly styleId: string | null;
  readonly look: DocumentTableLook;
}

export interface DocumentTableDesignAuthority {
  readonly current: DocumentTableDesignState;
  readonly receipt: {
    readonly schemaVersion: 1;
    readonly sourceState: DocumentTableDesignState;
    readonly sourceTablePropertiesSha256: string;
    readonly sourceStyleChainSha256: string | null;
    readonly sourceEffectsStyleChainSha256: string | null;
    readonly sourceStyleCustom: boolean | null;
    readonly editable: boolean;
  };
}

interface DocumentTableSelectionIdentity {
  readonly kind: string;
  readonly from: number;
  readonly to: number;
  readonly anchor: number;
  readonly head: number;
}

export interface DocumentTableDesignTarget {
  readonly doc: ProseMirrorNode;
  /** Exact immutable selection object captured at command-open time. */
  readonly selectionRef: Selection;
  readonly selection: DocumentTableSelectionIdentity;
  readonly tablePos: number;
  readonly tableNode: ProseMirrorNode;
  readonly authority: DocumentTableDesignAuthority;
  readonly authorityJson: string;
}

const EDGE_SET = new Set<string>(EDGE_KEYS);
const CONDITION_SET = new Set<string>(CONDITION_KEYS);
const HEX = /^#[0-9a-f]{6}$/;
const VISUAL_BORDER_COLOR = /^(?:#[0-9a-f]{6}|currentColor)$/;
const STYLE = /^[A-Za-z][A-Za-z0-9]{0,31}$/;
const VISUAL_BORDER_STYLES = new Set([
  "single",
  "double",
  "dotted",
  "dashed",
  "dashSmallGap",
  "dashDotStroked",
  "dotDash",
  "dotDotDash",
]);
const SHA = /^[0-9a-f]{64}$/;

function exactObject(value: unknown): Record<string, unknown> | null {
  return value && typeof value === "object" && !Array.isArray(value)
    ? (value as Record<string, unknown>)
    : null;
}

function exactKeys(
  value: Record<string, unknown>,
  allowed: readonly string[],
): boolean {
  return Object.keys(value).every((key) => allowed.includes(key));
}

function normalizeBorder(
  value: unknown,
): DocumentTableStyleVisualBorder | null {
  const raw = exactObject(value);
  if (!raw || !exactKeys(raw, ["style", "sizePt", "color"])) return null;
  if (
    typeof raw.style !== "string" ||
    !STYLE.test(raw.style) ||
    !VISUAL_BORDER_STYLES.has(raw.style)
  ) {
    return null;
  }
  const sizePt = raw.sizePt;
  if (
    sizePt !== undefined &&
    (typeof sizePt !== "number" ||
      !Number.isFinite(sizePt) ||
      sizePt < 0 ||
      sizePt > 96)
  ) {
    return null;
  }
  if (
    raw.color !== undefined &&
    (typeof raw.color !== "string" || !VISUAL_BORDER_COLOR.test(raw.color))
  ) {
    return null;
  }
  return {
    style: raw.style,
    ...(typeof sizePt === "number" ? { sizePt } : {}),
    ...(typeof raw.color === "string" ? { color: raw.color } : {}),
  };
}

function normalizeBorderMap(
  value: unknown,
): Partial<
  Record<DocumentTableVisualEdge, DocumentTableStyleVisualBorder>
> | null {
  const raw = exactObject(value);
  if (!raw || Object.keys(raw).length > EDGE_KEYS.length) return null;
  const result: Partial<
    Record<DocumentTableVisualEdge, DocumentTableStyleVisualBorder>
  > = {};
  for (const [key, candidate] of Object.entries(raw)) {
    if (!EDGE_SET.has(key)) return null;
    const border = normalizeBorder(candidate);
    if (!border) return null;
    result[key as DocumentTableVisualEdge] = border;
  }
  return result;
}

function normalizeClearedEdges(
  value: unknown,
): DocumentTableVisualEdge[] | null {
  if (!Array.isArray(value) || value.length > EDGE_KEYS.length) return null;
  if (
    !value.every(
      (item): item is DocumentTableVisualEdge =>
        typeof item === "string" && EDGE_SET.has(item),
    )
  ) {
    return null;
  }
  return new Set(value).size === value.length ? [...value] : null;
}

function normalizeLayer(value: unknown): DocumentTableStyleVisualLayer | null {
  const raw = exactObject(value);
  if (
    !raw ||
    !exactKeys(raw, [
      "fill",
      "color",
      "bold",
      "fontSizePt",
      "tableBorders",
      "cellBorders",
      "clearedTableBorders",
      "clearedCellBorders",
    ])
  ) {
    return null;
  }
  for (const key of ["fill", "color"] as const) {
    if (
      raw[key] !== undefined &&
      (typeof raw[key] !== "string" || !HEX.test(raw[key] as string))
    ) {
      return null;
    }
  }
  if (raw.bold !== undefined && typeof raw.bold !== "boolean") return null;
  if (
    raw.fontSizePt !== undefined &&
    (typeof raw.fontSizePt !== "number" || !Number.isFinite(raw.fontSizePt) || raw.fontSizePt < 1 || raw.fontSizePt > 1638)
  ) {
    return null;
  }
  const tableBorders = normalizeBorderMap(raw.tableBorders ?? {});
  const cellBorders = normalizeBorderMap(raw.cellBorders ?? {});
  const clearedTableBorders = normalizeClearedEdges(
    raw.clearedTableBorders ?? [],
  );
  const clearedCellBorders = normalizeClearedEdges(
    raw.clearedCellBorders ?? [],
  );
  if (
    !tableBorders ||
    !cellBorders ||
    !clearedTableBorders ||
    !clearedCellBorders
  ) {
    return null;
  }
  if (
    clearedTableBorders.some((edge) => tableBorders[edge]) ||
    clearedCellBorders.some((edge) => cellBorders[edge])
  ) {
    return null;
  }
  return {
    ...(typeof raw.fill === "string" ? { fill: raw.fill } : {}),
    ...(typeof raw.color === "string" ? { color: raw.color } : {}),
    ...(typeof raw.bold === "boolean" ? { bold: raw.bold } : {}),
    ...(typeof raw.fontSizePt === "number" ? { fontSizePt: raw.fontSizePt } : {}),
    tableBorders,
    cellBorders,
    clearedTableBorders,
    clearedCellBorders,
  };
}

export function normalizeDocumentTableStyleVisualAuthority(
  value: unknown,
): DocumentTableStyleVisualAuthority | null {
  const raw = exactObject(value);
  if (
    !raw ||
    !exactKeys(raw, [
      "schemaVersion",
      "styleId",
      "rowBandSize",
      "columnBandSize",
      "wholeTable",
      "conditionals",
      "wordWrapOmissionChainSha256",
    ]) ||
    raw.schemaVersion !== 1 ||
    (raw.wordWrapOmissionChainSha256 != null &&
      (typeof raw.wordWrapOmissionChainSha256 !== "string" ||
        !/^[0-9a-f]{64}$/.test(raw.wordWrapOmissionChainSha256))) ||
    typeof raw.styleId !== "string" ||
    !raw.styleId ||
    raw.styleId.length > 128 ||
    raw.styleId.trim() !== raw.styleId ||
    !Number.isInteger(raw.rowBandSize) ||
    (raw.rowBandSize as number) < 0 ||
    (raw.rowBandSize as number) > 3 ||
    !Number.isInteger(raw.columnBandSize) ||
    (raw.columnBandSize as number) < 0 ||
    (raw.columnBandSize as number) > 3
  ) {
    return null;
  }
  const wholeTable = normalizeLayer(raw.wholeTable);
  const conditions = exactObject(raw.conditionals);
  if (
    !wholeTable ||
    !conditions ||
    Object.keys(conditions).length > CONDITION_KEYS.length
  ) {
    return null;
  }
  const normalizedConditions: Partial<
    Record<DocumentTableVisualCondition, DocumentTableStyleVisualLayer>
  > = {};
  for (const [key, candidate] of Object.entries(conditions)) {
    if (!CONDITION_SET.has(key)) return null;
    const layer = normalizeLayer(candidate);
    if (!layer) return null;
    normalizedConditions[key as DocumentTableVisualCondition] = layer;
  }
  return {
    schemaVersion: 1,
    styleId: raw.styleId,
    rowBandSize: raw.rowBandSize as number,
    columnBandSize: raw.columnBandSize as number,
    wholeTable,
    conditionals: normalizedConditions,
    ...(typeof raw.wordWrapOmissionChainSha256 === "string"
      ? { wordWrapOmissionChainSha256: raw.wordWrapOmissionChainSha256 }
      : {}),
  };
}

export function normalizeDocumentTableStyleVisualCatalog(
  value: unknown,
): readonly DocumentTableStyleVisualCatalogEntry[] | null {
  if (!Array.isArray(value) || value.length > 512) return null;
  const result: DocumentTableStyleVisualCatalogEntry[] = [];
  const ids = new Set<string>();
  for (const candidate of value) {
    const raw = exactObject(candidate);
    if (!raw || !exactKeys(raw, ["styleId", "name", "visual", "displayOnly"])) return null;
    const visual = normalizeDocumentTableStyleVisualAuthority(raw.visual);
    if (
      typeof raw.styleId !== "string" ||
      typeof raw.name !== "string" ||
      !raw.name ||
      raw.name.length > 256 ||
      !visual ||
      visual.styleId !== raw.styleId ||
      ids.has(raw.styleId) ||
      (raw.displayOnly !== undefined && raw.displayOnly !== true)
    ) {
      return null;
    }
    ids.add(raw.styleId);
    result.push({
      styleId: raw.styleId,
      name: raw.name,
      visual,
      ...(raw.displayOnly === true ? { displayOnly: true as const } : {}),
    });
  }
  return result;
}

function normalizeLook(value: unknown): DocumentTableLook | null {
  const raw = exactObject(value);
  const keys = [
    "firstRow",
    "lastRow",
    "firstColumn",
    "lastColumn",
    "bandedRows",
    "bandedColumns",
  ] as const;
  if (!raw || Object.keys(raw).length !== keys.length || !exactKeys(raw, keys))
    return null;
  if (!keys.every((key) => typeof raw[key] === "boolean")) return null;
  return Object.fromEntries(
    keys.map((key) => [key, raw[key]]),
  ) as unknown as DocumentTableLook;
}

export function normalizeDocumentTableDesignState(
  value: unknown,
): DocumentTableDesignState | null {
  const raw = exactObject(value);
  if (!raw || !exactKeys(raw, ["styleId", "look"])) return null;
  if (
    raw.styleId !== null &&
    (typeof raw.styleId !== "string" ||
      !raw.styleId ||
      raw.styleId.length > 128)
  ) {
    return null;
  }
  const look = normalizeLook(raw.look);
  return look ? { styleId: raw.styleId as string | null, look } : null;
}

export function normalizeDocumentTableDesignAuthority(
  value: unknown,
): DocumentTableDesignAuthority | null {
  const raw = exactObject(value);
  if (!raw || !exactKeys(raw, ["current", "receipt"])) return null;
  const current = normalizeDocumentTableDesignState(raw.current);
  const receipt = exactObject(raw.receipt);
  if (
    !current ||
    !receipt ||
    !exactKeys(receipt, [
      "schemaVersion",
      "sourceState",
      "sourceTablePropertiesSha256",
      "sourceStyleChainSha256",
      "sourceEffectsStyleChainSha256",
      "sourceStyleCustom",
      "editable",
    ]) ||
    receipt.schemaVersion !== 1 ||
    typeof receipt.sourceTablePropertiesSha256 !== "string" ||
    !SHA.test(receipt.sourceTablePropertiesSha256) ||
    typeof receipt.editable !== "boolean"
  ) {
    return null;
  }
  const sourceState = normalizeDocumentTableDesignState(receipt.sourceState);
  if (!sourceState) return null;
  for (const key of [
    "sourceStyleChainSha256",
    "sourceEffectsStyleChainSha256",
  ] as const) {
    const candidate = receipt[key];
    if (
      candidate !== undefined &&
      candidate !== null &&
      (typeof candidate !== "string" || !SHA.test(candidate))
    ) {
      return null;
    }
  }
  if (
    receipt.sourceStyleCustom !== undefined &&
    receipt.sourceStyleCustom !== null &&
    typeof receipt.sourceStyleCustom !== "boolean"
  ) {
    return null;
  }
  return {
    current,
    receipt: {
      schemaVersion: 1,
      sourceState,
      sourceTablePropertiesSha256: receipt.sourceTablePropertiesSha256,
      sourceStyleChainSha256:
        typeof receipt.sourceStyleChainSha256 === "string"
          ? receipt.sourceStyleChainSha256
          : null,
      sourceEffectsStyleChainSha256:
        typeof receipt.sourceEffectsStyleChainSha256 === "string"
          ? receipt.sourceEffectsStyleChainSha256
          : null,
      sourceStyleCustom:
        typeof receipt.sourceStyleCustom === "boolean"
          ? receipt.sourceStyleCustom
          : null,
      editable: receipt.editable,
    },
  };
}

function mergeLayer(
  base: DocumentTableStyleVisualLayer,
  override: DocumentTableStyleVisualLayer,
): DocumentTableStyleVisualLayer {
  const mergeBorders = (
    baseBorders: DocumentTableStyleVisualLayer["tableBorders"],
    baseCleared: readonly DocumentTableVisualEdge[],
    overrideBorders: DocumentTableStyleVisualLayer["tableBorders"],
    overrideCleared: readonly DocumentTableVisualEdge[],
  ) => {
    const borders = { ...baseBorders };
    const cleared = new Set(baseCleared);
    for (const edge of overrideCleared) {
      delete borders[edge];
      cleared.add(edge);
    }
    for (const [edge, line] of Object.entries(overrideBorders)) {
      borders[edge as DocumentTableVisualEdge] = line;
      cleared.delete(edge as DocumentTableVisualEdge);
    }
    return { borders, cleared: [...cleared] };
  };
  const table = mergeBorders(
    base.tableBorders,
    base.clearedTableBorders,
    override.tableBorders,
    override.clearedTableBorders,
  );
  const cell = mergeBorders(
    base.cellBorders,
    base.clearedCellBorders,
    override.cellBorders,
    override.clearedCellBorders,
  );
  const color = override.color ?? base.color;
  const bold = override.bold ?? base.bold;
  const fontSizePt = override.fontSizePt ?? base.fontSizePt;
  return {
    ...(override.fill !== undefined
      ? { fill: override.fill }
      : base.fill
        ? { fill: base.fill }
        : {}),
    ...(color !== undefined ? { color } : {}),
    ...(bold !== undefined ? { bold } : {}),
    ...(fontSizePt !== undefined ? { fontSizePt } : {}),
    tableBorders: table.borders,
    cellBorders: cell.borders,
    clearedTableBorders: table.cleared,
    clearedCellBorders: cell.cleared,
  };
}

/** Which outer edges of a conditional region one cell lies on. */
interface DocumentTableStyleRegionEdges {
  readonly top: boolean;
  readonly right: boolean;
  readonly bottom: boolean;
  readonly left: boolean;
}

/**
 * A style region's cell borders on one cell. Word reads `w:tcBorders` of a
 * table style region (the whole table, a band, the header row, a column, a
 * corner) region-relative: `top`/`left`/`bottom`/`right` are the region's
 * outer edges and `insideH`/`insideV` the rules between its cells. Measured
 * in Word: Themed Style 1's header row (`left`/`right` single, `insideV`
 * nil) draws no rule between its cells, only at the table's edges.
 */
function placeRegionCellBorders(
  layer: DocumentTableStyleVisualLayer,
  region: DocumentTableStyleRegionEdges,
): DocumentTableStyleVisualLayer {
  const cellBorders: Partial<
    Record<DocumentTableVisualEdge, DocumentTableStyleVisualBorder>
  > = {};
  const clearedCellBorders: DocumentTableVisualEdge[] = [];
  for (const side of ["top", "right", "bottom", "left"] as const) {
    const source: DocumentTableVisualEdge = region[side]
      ? side
      : side === "top" || side === "bottom"
        ? "insideH"
        : "insideV";
    if (layer.clearedCellBorders.includes(source)) {
      clearedCellBorders.push(side);
      continue;
    }
    const line = layer.cellBorders[source];
    if (line) cellBorders[side] = line;
  }
  return { ...layer, cellBorders, clearedCellBorders };
}

export function resolveDocumentTableStyleCellVisual(
  visual: DocumentTableStyleVisualAuthority,
  look: DocumentTableLook,
  geometry: {
    rowStart: number;
    rowEnd: number;
    columnStart: number;
    columnEnd: number;
    rows: number;
    columns: number;
  },
): DocumentTableStyleVisualLayer | null {
  const { rowStart, rowEnd, columnStart, columnEnd, rows, columns } = geometry;
  if (
    !Number.isInteger(rowStart) ||
    !Number.isInteger(rowEnd) ||
    !Number.isInteger(columnStart) ||
    !Number.isInteger(columnEnd) ||
    rowStart < 0 ||
    rowStart > rowEnd ||
    rowEnd >= rows ||
    columnStart < 0 ||
    columnStart > columnEnd ||
    columnEnd >= columns
  ) {
    return null;
  }
  const atTop = rowStart === 0;
  const atBottom = rowEnd === rows - 1;
  const atLeft = columnStart === 0;
  const atRight = columnEnd === columns - 1;
  let layer = placeRegionCellBorders(visual.wholeTable, {
    top: atTop,
    right: atRight,
    bottom: atBottom,
    left: atLeft,
  });
  const apply = (
    key: DocumentTableVisualCondition,
    region: DocumentTableStyleRegionEdges,
  ) => {
    const candidate = visual.conditionals[key];
    if (candidate) {
      layer = mergeLayer(layer, placeRegionCellBorders(candidate, region));
    }
  };
  const bodyRowStart = look.firstRow && rows > 1 ? 1 : 0;
  const bodyRowEnd = rows - (look.lastRow && rows > bodyRowStart ? 1 : 0);
  if (
    look.bandedRows &&
    visual.rowBandSize > 0 &&
    rowStart >= bodyRowStart &&
    rowStart < bodyRowEnd
  ) {
    const band = ((rowStart - bodyRowStart) / visual.rowBandSize) | 0;
    const bandStart = bodyRowStart + band * visual.rowBandSize;
    const bandEnd = Math.min(bandStart + visual.rowBandSize, bodyRowEnd) - 1;
    apply(band % 2 === 0 ? "band1Horz" : "band2Horz", {
      top: rowStart === bandStart,
      right: atRight,
      bottom: rowEnd >= bandEnd,
      left: atLeft,
    });
  }
  const bodyColumnStart = look.firstColumn && columns > 1 ? 1 : 0;
  const bodyColumnEnd =
    columns - (look.lastColumn && columns > bodyColumnStart ? 1 : 0);
  if (
    look.bandedColumns &&
    visual.columnBandSize > 0 &&
    columnStart >= bodyColumnStart &&
    columnStart < bodyColumnEnd
  ) {
    const band = ((columnStart - bodyColumnStart) / visual.columnBandSize) | 0;
    const bandStart = bodyColumnStart + band * visual.columnBandSize;
    const bandEnd =
      Math.min(bandStart + visual.columnBandSize, bodyColumnEnd) - 1;
    apply(band % 2 === 0 ? "band1Vert" : "band2Vert", {
      top: atTop,
      right: columnEnd >= bandEnd,
      bottom: atBottom,
      left: columnStart === bandStart,
    });
  }
  if (look.firstColumn && columnStart === 0) {
    apply("firstCol", {
      top: atTop,
      right: columnEnd === 0,
      bottom: atBottom,
      left: true,
    });
  }
  if (look.lastColumn && columnEnd === columns - 1) {
    apply("lastCol", {
      top: atTop,
      right: true,
      bottom: atBottom,
      left: columnStart === columns - 1,
    });
  }
  if (look.firstRow && rowStart === 0) {
    apply("firstRow", {
      top: true,
      right: atRight,
      bottom: rowEnd === 0,
      left: atLeft,
    });
  }
  if (look.lastRow && rowEnd === rows - 1) {
    apply("lastRow", {
      top: rowStart === rows - 1,
      right: atRight,
      bottom: true,
      left: atLeft,
    });
  }
  const corner = { top: true, right: true, bottom: true, left: true };
  if (look.firstRow && look.firstColumn && rowStart === 0 && columnStart === 0)
    apply("nwCell", corner);
  if (
    look.firstRow &&
    look.lastColumn &&
    rowStart === 0 &&
    columnEnd === columns - 1
  )
    apply("neCell", corner);
  if (
    look.lastRow &&
    look.firstColumn &&
    rowEnd === rows - 1 &&
    columnStart === 0
  )
    apply("swCell", corner);
  if (
    look.lastRow &&
    look.lastColumn &&
    rowEnd === rows - 1 &&
    columnEnd === columns - 1
  )
    apply("seCell", corner);
  return layer;
}

export function documentTableStyleCatalogFromDoc(
  doc: ProseMirrorNode,
): readonly DocumentTableStyleVisualCatalogEntry[] {
  return (
    normalizeDocumentTableStyleVisualCatalog(
      doc.attrs.tableStyleVisualCatalog,
    ) ?? []
  );
}

/** The table styles a document's own `styles.xml` defines, by id and name. */
export function documentDefinedTableStyles(
  doc: ProseMirrorNode,
): readonly OfficeWordDefinedStyle[] {
  const defined = new Map<string, OfficeWordDefinedStyle>();
  const styles = exactObject(doc.attrs.styleDefinitions)?.styles;
  if (Array.isArray(styles)) {
    for (const raw of styles) {
      const style = exactObject(raw);
      if (
        style &&
        typeof style.styleId === "string" &&
        typeof style.name === "string"
      ) {
        defined.set(style.styleId, {
          styleId: style.styleId,
          name: style.name,
          ...(typeof style.type === "string" ? { type: style.type } : {}),
        });
      }
    }
  }
  for (const entry of documentTableStyleCatalogFromDoc(doc)) {
    if (!defined.has(entry.styleId)) {
      defined.set(entry.styleId, {
        styleId: entry.styleId,
        name: entry.name,
        type: "table",
      });
    }
  }
  return [...defined.values()];
}

/**
 * The Office gallery style a table style id stands for in this document:
 * a style the document defines is that gallery style when it carries the
 * gallery NAME (Word's own style identity); an id the document does not
 * define is one the gallery allocates and the save will define.
 */
export function documentOfficeTableStyleForId(
  doc: ProseMirrorNode,
  styleId: unknown,
): OfficeWordTableStyle | null {
  if (typeof styleId !== "string" || !styleId) return null;
  const defined = documentDefinedTableStyles(doc).find(
    (style) => style.styleId === styleId,
  );
  if (defined) {
    return defined.type === undefined || defined.type === "table"
      ? officeWordTableStyleByName(defined.name)
      : null;
  }
  return officeWordTableStyleForAllocatedId(styleId);
}

function documentTableDesignAcceptsStyle(
  doc: ProseMirrorNode,
  styleId: string,
): boolean {
  return (
    documentTableDesignStyleCatalogFromDoc(doc).some(
      (entry) => entry.styleId === styleId,
    ) || documentOfficeTableStyleForId(doc, styleId) !== null
  );
}

/** The styles Table Design may apply: display-only entries excluded. */
export function documentTableDesignStyleCatalogFromDoc(
  doc: ProseMirrorNode,
): readonly DocumentTableStyleVisualCatalogEntry[] {
  return documentTableStyleCatalogFromDoc(doc).filter(
    (entry) => entry.displayOnly !== true,
  );
}

export function documentTableDesignAuthorityFromNode(
  table: ProseMirrorNode,
): DocumentTableDesignAuthority | null {
  const source = exactObject(table.attrs.source);
  return normalizeDocumentTableDesignAuthority(source?.tableDesignAuthority);
}

function selectionIdentity(
  selection: Selection,
): DocumentTableSelectionIdentity {
  return {
    kind: selection.constructor.name,
    from: selection.from,
    to: selection.to,
    anchor: selection.anchor,
    head: selection.head,
  };
}

function sameSelectionIdentity(
  left: DocumentTableSelectionIdentity,
  right: DocumentTableSelectionIdentity,
): boolean {
  return (
    left.kind === right.kind &&
    left.from === right.from &&
    left.to === right.to &&
    left.anchor === right.anchor &&
    left.head === right.head
  );
}

export function readDocumentTableDesignTarget(
  editor: Editor,
): DocumentTableDesignTarget | null {
  if (editor.isDestroyed) return null;
  const { doc, selection } = editor.state;
  let tablePos: number | null = null;
  let tableNode: ProseMirrorNode | null = null;
  for (let depth = selection.$from.depth; depth > 0; depth -= 1) {
    const node = selection.$from.node(depth);
    if (node.type.name !== "table") continue;
    tablePos = selection.$from.before(depth);
    tableNode = node;
    break;
  }
  if (!tableNode && doc.nodeAt(selection.from)?.type.name === "table") {
    tablePos = selection.from;
    tableNode = doc.nodeAt(selection.from);
  }
  if (tablePos === null || !tableNode || doc.nodeAt(tablePos) !== tableNode) {
    return null;
  }
  const authority = documentTableDesignAuthorityFromNode(tableNode);
  if (!authority) return null;
  return {
    doc,
    selectionRef: selection,
    selection: selectionIdentity(selection),
    tablePos,
    tableNode,
    authority,
    authorityJson: JSON.stringify(authority),
  };
}

export function authenticateDocumentTableDesignTarget(
  editor: Editor,
  expected: DocumentTableDesignTarget,
): DocumentTableDesignTarget | null {
  if (
    editor.isDestroyed ||
    !editor.isEditable ||
    editor.state.doc !== expected.doc ||
    editor.state.selection !== expected.selectionRef
  ) {
    return null;
  }
  const current = readDocumentTableDesignTarget(editor);
  if (
    !current ||
    current.tablePos !== expected.tablePos ||
    current.tableNode !== expected.tableNode ||
    !sameSelectionIdentity(current.selection, expected.selection) ||
    current.authorityJson !== expected.authorityJson
  ) {
    return null;
  }
  return current;
}

export function authenticateDocumentTableDesignPreview(
  editor: Editor,
  expected: DocumentTableDesignTarget,
  state: unknown,
): DocumentTableDesignState | null {
  const live = authenticateDocumentTableDesignTarget(editor, expected);
  const normalized = normalizeDocumentTableDesignState(state);
  if (!live || !normalized || !live.authority.receipt.editable) return null;
  if (normalized.styleId === null) return normalized;
  return documentTableDesignAcceptsStyle(editor.state.doc, normalized.styleId)
    ? normalized
    : null;
}

export function setDocumentTableDesignState(
  editor: Editor,
  expected: DocumentTableDesignTarget,
  nextState: DocumentTableDesignState,
): boolean {
  const live = authenticateDocumentTableDesignTarget(editor, expected);
  const normalizedNext = normalizeDocumentTableDesignState(nextState);
  if (!live || !normalizedNext || !live.authority.receipt.editable)
    return false;
  if (
    normalizedNext.styleId !== null &&
    !documentTableDesignAcceptsStyle(editor.state.doc, normalizedNext.styleId)
  ) {
    return false;
  }
  if (
    JSON.stringify(live.authority.current) === JSON.stringify(normalizedNext)
  ) {
    return false;
  }
  const source = exactObject(live.tableNode.attrs.source) ?? {};
  const tr = closeHistory(
    editor.state.tr.setNodeMarkup(live.tablePos, undefined, {
      ...live.tableNode.attrs,
      source: {
        ...source,
        tableDesignAuthority: {
          ...live.authority,
          current: normalizedNext,
        },
      },
    }),
  );
  editor.view.dispatch(tr);
  return true;
}
