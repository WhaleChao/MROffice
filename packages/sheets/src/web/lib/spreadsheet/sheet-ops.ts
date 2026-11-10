/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Sheet-tab operations beyond add/rename/delete (Excel/WPS parity): MOVE/REORDER
 * a tab and DUPLICATE (deep-copy) a sheet. Kept pure and generic so the array
 * math and the deep clone are unit-tested in isolation, away from the React
 * component's state plumbing.
 *
 *  - REORDER: move the sheet at `from` to index `to`, shifting the rest, exactly
 *    like dragging a tab or "Move Left / Move Right". The array is rebuilt
 *    (never mutated) and out-of-range / no-op moves return the input unchanged.
 *
 *  - DUPLICATE: deep-clone a sheet (every cell, style, comment, hyperlink, span
 *    AND every overlay — merges, validations, charts, named ranges, …) so the
 *    copy is fully independent: editing one never bleeds into the other. The
 *    clone gets a fresh, de-duplicated id + name and is inserted directly AFTER
 *    its source (Excel drops "Sheet (2)" right next to the original).
 */

import {
  BLANK_CELL_FORMATS_HINT,
  type BlankCellFormatGroup,
} from "@mrnote/core/web/lib/spreadsheet/blank-cell-formats.ts";
import {
  MAX_DATA_TABLE_ID_LENGTH,
  MAX_DATA_TABLE_NAME_LENGTH,
  isValidDataTableName,
  renameDataTableReferences,
} from "@mrnote/core/web/lib/spreadsheet/data-table.ts";
import pageEnvelope from "@mrnote/core/web/lib/spreadsheet/spreadsheet-page-envelope.shared.json";
import { renameSheetInFormula } from "@mrnote/core/web/lib/spreadsheet/formula-engine.ts";
import { renameSheetReferences } from "./sheet-formula-references";
import {
  MAX_SHEET_NAME_LENGTH,
  sheetNameProblem,
} from "./sheet-name-rules";

/** The minimal sheet shape these helpers need: a stable id + a display name. */
export interface SheetLike {
  id: string;
  name: string;
  [key: string]: unknown;
}

/**
 * The sheet clone crosses the same JSON-backed persistence boundary as a normal
 * spreadsheet save. Keep its resource ceilings explicit and deterministic so a
 * duplicate action cannot allocate an unbounded hostile object graph before the
 * autosave validator gets a chance to inspect it.
 *
 * These values align with the Spreadsheet model's largest legal string leaf
 * (an inline image) and the API's spreadsheet content-tree envelope
 * (SPREADSHEET_JSON_TREE_BUDGET, read from the page envelope file both trees
 * hold, `spreadsheet-page-envelope.shared.json`). They stay
 * here because clone/native identity and serializer ownership are
 * editor-specific.
 */
export const SHEET_CLONE_LIMITS = Object.freeze({
  maxDepth: pageEnvelope.maxDepth,
  // A sheet may use the whole workbook's node budget.
  maxNodes: pageEnvelope.maxNodes,
  maxArrayLength: pageEnvelope.maxArrayLength,
  // The stored cells of one sheet map.
  maxObjectKeys: pageEnvelope.maxCellsPerSheet,
  maxStringChars: 2_000_000,
  maxTotalStringChars: pageEnvelope.maxTotalStringChars,
});

export type SheetCloneErrorCode =
  | "invalid_identity"
  | "unsupported_value"
  | "non_finite_number"
  | "cyclic_reference"
  | "unsupported_prototype"
  | "unsafe_property"
  | "limit_exceeded";

/** Stable fail-closed error surfaced when a sheet cannot be copied losslessly. */
export class SheetCloneError extends Error {
  readonly name = "SheetCloneError";

  constructor(readonly code: SheetCloneErrorCode) {
    super(`Sheet clone rejected: ${code}`);
  }
}

interface SheetCloneBudgetUsage {
  nodeCount: number;
  totalStringChars: number;
}

interface SheetCloneSnapshot<T> {
  value: T;
  usage: SheetCloneBudgetUsage;
}

function rejectSheetClone(code: SheetCloneErrorCode): never {
  throw new SheetCloneError(code);
}

function addSheetCloneString(
  usage: SheetCloneBudgetUsage,
  value: string,
): void {
  if (value.length > SHEET_CLONE_LIMITS.maxStringChars) {
    rejectSheetClone("limit_exceeded");
  }
  usage.totalStringChars += value.length;
  if (usage.totalStringChars > SHEET_CLONE_LIMITS.maxTotalStringChars) {
    rejectSheetClone("limit_exceeded");
  }
}

function isCanonicalArrayIndex(key: string, length: number): boolean {
  if (!/^(?:0|[1-9]\d*)$/.test(key)) return false;
  const index = Number(key);
  return Number.isSafeInteger(index) && index >= 0 && index < length;
}

/**
 * Clone the exact persistence-safe value graph without invoking user code.
 *
 * Object data properties whose value is `undefined` are preserved because the
 * current normalized SheetModel intentionally owns some optional keys with that
 * value. Array `undefined` is rejected: JSON persistence would silently turn it
 * into null. Repeated acyclic references are independently materialized, which
 * matches a save/reload boundary and prevents aliases between clone branches.
 */
function cloneSheetSnapshot<T>(root: T): SheetCloneSnapshot<T> {
  if (root === null || typeof root !== "object" || Array.isArray(root)) {
    rejectSheetClone("invalid_identity");
  }
  const usage: SheetCloneBudgetUsage = {
    nodeCount: 0,
    totalStringChars: 0,
  };
  const ancestors = new WeakSet<object>();

  const visit = (
    current: unknown,
    depth: number,
    inArray: boolean,
  ): unknown => {
    if (depth > SHEET_CLONE_LIMITS.maxDepth) {
      rejectSheetClone("limit_exceeded");
    }
    usage.nodeCount += 1;
    if (usage.nodeCount > SHEET_CLONE_LIMITS.maxNodes) {
      rejectSheetClone("limit_exceeded");
    }

    if (current === null || typeof current === "boolean") return current;
    if (current === undefined) {
      if (inArray) rejectSheetClone("unsupported_value");
      return current;
    }
    if (typeof current === "string") {
      addSheetCloneString(usage, current);
      return current;
    }
    if (typeof current === "number") {
      if (!Number.isFinite(current)) rejectSheetClone("non_finite_number");
      return current;
    }
    if (typeof current !== "object") {
      rejectSheetClone("unsupported_value");
    }
    if (ancestors.has(current)) rejectSheetClone("cyclic_reference");

    const isArray = Array.isArray(current);
    const prototype = Object.getPrototypeOf(current);
    if (
      (isArray && prototype !== Array.prototype) ||
      (!isArray && prototype !== Object.prototype && prototype !== null)
    ) {
      rejectSheetClone("unsupported_prototype");
    }

    let arrayLength = 0;
    if (isArray) {
      const lengthDescriptor = Object.getOwnPropertyDescriptor(
        current,
        "length",
      );
      if (
        !lengthDescriptor ||
        !("value" in lengthDescriptor) ||
        !Number.isSafeInteger(lengthDescriptor.value) ||
        lengthDescriptor.value < 0
      ) {
        rejectSheetClone("unsafe_property");
      }
      arrayLength = lengthDescriptor.value;
      if (arrayLength > SHEET_CLONE_LIMITS.maxArrayLength) {
        rejectSheetClone("limit_exceeded");
      }
    }

    // An oversized array is rejected from its fixed own length descriptor
    // before ownKeys can allocate one key string per populated element.
    const keys = Reflect.ownKeys(current);
    if (isArray) {
      // Exactly one own key per element plus the built-in length property means
      // there are no holes, symbols, or extra array properties. Each key is
      // still checked below before any value is read.
      if (keys.length !== arrayLength + 1) {
        rejectSheetClone("unsafe_property");
      }
    } else if (keys.length > SHEET_CLONE_LIMITS.maxObjectKeys) {
      rejectSheetClone("limit_exceeded");
    }

    const clone: object = isArray
      ? new Array(arrayLength)
      : Object.create(prototype);
    ancestors.add(current);
    try {
      for (const key of keys) {
        if (key === "length" && isArray) continue;
        if (typeof key === "symbol") rejectSheetClone("unsafe_property");
        if (isArray && !isCanonicalArrayIndex(key, arrayLength)) {
          rejectSheetClone("unsafe_property");
        }
        // JSON arrays serialize positions, not their numeric property names;
        // only record keys consume the persisted string budget.
        if (!isArray) addSheetCloneString(usage, key);
        const descriptor = Object.getOwnPropertyDescriptor(current, key);
        if (!descriptor || !("value" in descriptor) || !descriptor.enumerable) {
          rejectSheetClone("unsafe_property");
        }
        const child = visit(descriptor.value, depth + 1, isArray);
        Object.defineProperty(clone, key, {
          value: child,
          enumerable: true,
          configurable: true,
          writable: true,
        });
      }
    } finally {
      ancestors.delete(current);
    }
    return clone;
  };

  return { value: visit(root, 0, false) as T, usage };
}

function isValidSheetIdentity(value: unknown): value is string {
  return (
    typeof value === "string" &&
    value.length <= SHEET_CLONE_LIMITS.maxStringChars &&
    value.trim().length > 0
  );
}

function readSheetIdentity(sheet: SheetLike): { id: string; name: string } {
  if (!sheet || typeof sheet !== "object") {
    rejectSheetClone("invalid_identity");
  }
  const prototype = Object.getPrototypeOf(sheet);
  if (prototype !== Object.prototype && prototype !== null) {
    rejectSheetClone("unsupported_prototype");
  }
  const idDescriptor = Object.getOwnPropertyDescriptor(sheet, "id");
  const nameDescriptor = Object.getOwnPropertyDescriptor(sheet, "name");
  if (
    !idDescriptor ||
    !("value" in idDescriptor) ||
    !idDescriptor.enumerable ||
    !nameDescriptor ||
    !("value" in nameDescriptor) ||
    !nameDescriptor.enumerable ||
    !isValidSheetIdentity(idDescriptor.value) ||
    !isValidSheetIdentity(nameDescriptor.value)
  ) {
    rejectSheetClone("invalid_identity");
  }
  return { id: idDescriptor.value, name: nameDescriptor.value };
}

function applyClonedSheetIdentity<T extends SheetLike>(
  snapshot: SheetCloneSnapshot<T>,
  id: string,
  name: string,
): T {
  const clone = snapshot.value;
  const sourceIdentity = readSheetIdentity(clone);
  if (!isValidSheetIdentity(id) || !isValidSheetIdentity(name)) {
    rejectSheetClone("invalid_identity");
  }
  const outputStringChars =
    snapshot.usage.totalStringChars -
    sourceIdentity.id.length -
    sourceIdentity.name.length +
    id.length +
    name.length;
  if (outputStringChars > SHEET_CLONE_LIMITS.maxTotalStringChars) {
    rejectSheetClone("limit_exceeded");
  }
  Object.defineProperty(clone, "id", {
    value: id,
    enumerable: true,
    configurable: true,
    writable: true,
  });
  Object.defineProperty(clone, "name", {
    value: name,
    enumerable: true,
    configurable: true,
    writable: true,
  });
  return clone;
}

/**
 * Move the element at `from` to `to`, shifting the others to fill the gap. A
 * no-op (`from === to`) or any out-of-range index returns the SAME array
 * reference so the caller can skip the commit. The input is never mutated.
 *
 * `to` is the destination index in the POST-removal array sense used by a drag
 * reorder: moving item 0 to index 2 in [A,B,C,D] yields [B,C,A,D]. It is clamped
 * to [0, length-1].
 */
export function reorderSheets<T>(sheets: T[], from: number, to: number): T[] {
  if (from < 0 || from >= sheets.length) return sheets;
  const clampedTo = Math.max(0, Math.min(to, sheets.length - 1));
  if (from === clampedTo) return sheets;
  const next = sheets.slice();
  const [moved] = next.splice(from, 1);
  next.splice(clampedTo, 0, moved);
  return next;
}

/** Adjacent-swap convenience for "Move Left" (dir -1) / "Move Right" (dir +1). */
export function moveSheet<T>(sheets: T[], index: number, dir: -1 | 1): T[] {
  return reorderSheets(sheets, index, index + dir);
}

/**
 * `base` cut to at most `length` UTF-16 units (the unit Excel's 31-character
 * sheet-name limit counts), never splitting a surrogate pair, with trailing
 * spaces dropped so the suffix is not preceded by a double space.
 */
function truncateSheetNameBase(base: string, length: number): string {
  let cut = base.slice(0, Math.max(0, length));
  const last = cut.charCodeAt(cut.length - 1);
  if (cut.length < base.length && last >= 0xd800 && last <= 0xdbff) {
    cut = cut.slice(0, -1);
  }
  return cut.trimEnd();
}

/**
 * The copy name for `baseName` with copy number `counter`: Excel's
 * "<base> (2)", the base shortened so the whole name keeps within Excel's
 * 31-character worksheet-name limit (a 31-character "Q3 …" copies as the first
 * 27 characters + " (2)", exactly as Excel does rather than overflowing).
 */
function duplicateSheetNameCandidate(baseName: string, counter: number): string {
  const suffix = ` (${counter})`;
  const base =
    truncateSheetNameBase(baseName, MAX_SHEET_NAME_LENGTH - suffix.length) ||
    "Sheet";
  return `${base}${suffix}`;
}

/**
 * Build a fresh sheet id/name pair for a duplicate, de-duplicated against the
 * existing tabs. The name follows Excel's "<base> (2)" convention, bumping the
 * counter until it is free; a long base is shortened so the name stays a legal
 * Excel worksheet name (sheet-name-rules), and uniqueness is checked after the
 * shortening, so two long names sharing a prefix still get distinct copies.
 * Every copy route (Duplicate, Move or Copy, Paste Worksheet, the Copilot
 * worksheet copy) names its sheet here. The id is a slug + counter,
 * independently bumped so a persisted id never collides (ids and names are
 * de-duped separately because a user-renamed tab can make a name taken while
 * its id slot is free, and vice versa).
 */
export function duplicateSheetNaming(
  baseName: string,
  takenIds: ReadonlySet<string>,
  takenNames: ReadonlySet<string>,
): { id: string; name: string } {
  let counter = 2;
  let name = duplicateSheetNameCandidate(baseName, counter);
  // sheetNameProblem compares without regard to case, as Excel does.
  while (sheetNameProblem(name, takenNames) === "duplicate") {
    counter += 1;
    name = duplicateSheetNameCandidate(baseName, counter);
  }
  // The id slug is derived from the base name but stays ASCII-safe and bounded;
  // an empty/garbage slug falls back to "sheet".
  const slug =
    baseName
      .toLowerCase()
      .replace(/[^a-z0-9]+/g, "-")
      .replace(/^-+|-+$/g, "")
      .slice(0, 40) || "sheet";
  let idCounter = 2;
  let id = `${slug}-copy-${idCounter}`;
  while (takenIds.has(id)) {
    idCounter += 1;
    id = `${slug}-copy-${idCounter}`;
  }
  return { id, name };
}

/**
 * Deep-clone a persistence-safe sheet so the copy shares NO references with its
 * source — every nested cell, style bag, comment thread, overlay array and
 * config object is fully independent. Accessors and values that JSON would
 * silently invoke, drop, or coerce are rejected before a clone can be committed.
 * Then overwrite the id + name with the supplied fresh pair.
 *
 * Returns a brand-new object; the input is never mutated. A rejected sheet
 * throws SheetCloneError and produces no externally reachable partial clone.
 */
export function cloneSheetWithIdentity<T extends SheetLike>(
  sheet: T,
  id: string,
  name: string,
): T {
  if (!isValidSheetIdentity(id) || !isValidSheetIdentity(name)) {
    rejectSheetClone("invalid_identity");
  }
  const { content, blankCellFormats } = detachSheetRuntimeHints(sheet);
  return reattachSheetRuntimeHints(
    applyClonedSheetIdentity(cloneSheetSnapshot(content), id, name),
    blankCellFormats,
  );
}

/**
 * Separate the editor's own runtime metadata from a sheet's persisted content
 * before the fail-closed cloner sees it. The only such metadata is the
 * blank-cell-format hint (BLANK_CELL_FORMATS_HINT): the ranges a saved sheet's
 * formatted blank cells were loaded from, kept under a symbol so it never
 * reaches JSON. The formatted cells themselves are ordinary `cells` entries
 * and are cloned like any other; the hint only lets the next save fold them
 * back into the same ranges.
 *
 * Exactly that one symbol, as a plain data property on the sheet itself, is
 * taken aside; the rest of the sheet is handed on through its own property
 * descriptors (no getter runs), so any other symbol or accessor is still
 * rejected by the cloner. The hint's value is itself cloned through the same
 * JSON-safe path, so a hostile hint fails closed too.
 */
function detachSheetRuntimeHints<T extends SheetLike>(
  sheet: T,
): { content: T; blankCellFormats?: readonly BlankCellFormatGroup[] } {
  if (!sheet || typeof sheet !== "object") return { content: sheet };
  const descriptor = Object.getOwnPropertyDescriptor(
    sheet,
    BLANK_CELL_FORMATS_HINT,
  );
  if (!descriptor || !("value" in descriptor)) return { content: sheet };
  const descriptors = Object.getOwnPropertyDescriptors(sheet) as Record<
    PropertyKey,
    PropertyDescriptor
  >;
  delete descriptors[BLANK_CELL_FORMATS_HINT as unknown as string];
  const content = Object.create(
    Object.getPrototypeOf(sheet),
    descriptors,
  ) as T;
  if (descriptor.value === undefined) return { content };
  if (!Array.isArray(descriptor.value)) rejectSheetClone("unsafe_property");
  const blankCellFormats = cloneSheetSnapshot({ groups: descriptor.value })
    .value.groups as readonly BlankCellFormatGroup[];
  return { content, blankCellFormats };
}

function reattachSheetRuntimeHints<T extends SheetLike>(
  clone: T,
  blankCellFormats: readonly BlankCellFormatGroup[] | undefined,
): T {
  if (blankCellFormats) {
    Object.defineProperty(clone, BLANK_CELL_FORMATS_HINT, {
      value: blankCellFormats,
      enumerable: true,
      configurable: true,
      writable: true,
    });
  }
  return clone;
}

/** A workbook-unique `${base}-copy-N` id no longer than `maxLength`. */
function uniqueCopyId(
  base: string,
  taken: Set<string>,
  maxLength: number,
): string {
  for (let counter = 2; ; counter += 1) {
    const suffix = `-copy-${counter}`;
    const id = `${base.slice(0, Math.max(1, maxLength - suffix.length))}${suffix}`;
    if (!taken.has(id)) {
      taken.add(id);
      return id;
    }
  }
}

/**
 * The name a copied table takes: Excel keeps table names unique across the
 * workbook and, copying a sheet, numbers the copy after its source (Sales →
 * Sales2, Table1 → Table12, the next free number). A numbered name that would
 * read as a cell reference (Tax → TAX2) takes an underscore (Tax_2). Table and
 * defined names share one namespace, so both count as taken.
 */
function uniqueCopiedTableName(base: string, taken: Set<string>): string {
  // A legacy name the table rules refuse cannot seed a legal copy name.
  const stem = isValidDataTableName(base) ? base.trim() : "Table";
  const candidate = (joint: string, digits: string) =>
    `${stem.slice(0, MAX_DATA_TABLE_NAME_LENGTH - joint.length - digits.length)}${joint}${digits}`;
  for (let counter = 2; ; counter += 1) {
    const digits = String(counter);
    const plain = candidate("", digits);
    const name = isValidDataTableName(plain) ? plain : candidate("_", digits);
    if (isValidDataTableName(name) && !taken.has(name.toLowerCase())) {
      taken.add(name.toLowerCase());
      return name;
    }
  }
}

type JsonRecord = Record<string, unknown>;

function isJsonRecord(value: unknown): value is JsonRecord {
  return value !== null && typeof value === "object" && !Array.isArray(value);
}

function mapRecordStrings(
  record: unknown,
  rewrite: (value: string) => string,
): unknown {
  if (!isJsonRecord(record)) return record;
  let changed = false;
  const next: JsonRecord = {};
  for (const [key, value] of Object.entries(record)) {
    const rewritten = typeof value === "string" ? rewrite(value) : value;
    if (rewritten !== value) changed = true;
    next[key] = rewritten;
  }
  return changed ? next : record;
}

/**
 * Give every table on a freshly cloned sheet its own identity, the way Excel
 * copies a worksheet: a new id and a workbook-unique name for each table, and
 * the copy's own structured references (`=SUM(Sales[Amount])` in its cells and
 * in its tables' calculated-column / total templates) renamed to follow it, so
 * the copy computes from its own table. Nothing outside the clone changes:
 * formulas on other sheets keep naming the original. Table slicers and
 * PivotTables on the copy that read one of its tables follow the new id.
 */
function reidentifyCopiedSheetTables<T extends SheetLike>(
  clone: T,
  workbook: readonly SheetLike[],
): T {
  const tables = clone.dataTables;
  if (!Array.isArray(tables) || tables.length === 0) return clone;
  const takenIds = new Set<string>();
  const takenNames = new Set<string>();
  const takenSlicerIds = new Set<string>();
  for (const sheet of workbook) {
    for (const table of Array.isArray(sheet.dataTables) ? sheet.dataTables : []) {
      if (!isJsonRecord(table)) continue;
      if (typeof table.id === "string") takenIds.add(table.id);
      if (typeof table.name === "string") takenNames.add(table.name.toLowerCase());
    }
    for (const entry of Array.isArray(sheet.namedRanges) ? sheet.namedRanges : []) {
      if (isJsonRecord(entry) && typeof entry.name === "string") {
        takenNames.add(entry.name.toLowerCase());
      }
    }
    for (const slicer of Array.isArray(sheet.slicers) ? sheet.slicers : []) {
      if (isJsonRecord(slicer) && typeof slicer.id === "string") {
        takenSlicerIds.add(slicer.id);
      }
    }
  }

  const idMap = new Map<string, string>();
  const renames: Array<{ from: string; to: string }> = [];
  let nextTables = tables.map((table) => {
    if (
      !isJsonRecord(table) ||
      typeof table.id !== "string" ||
      typeof table.name !== "string"
    ) {
      return table;
    }
    const id = uniqueCopyId(table.id, takenIds, MAX_DATA_TABLE_ID_LENGTH);
    const name = uniqueCopiedTableName(table.name, takenNames);
    idMap.set(table.id, id);
    if (name !== table.name) renames.push({ from: table.name, to: name });
    return { ...table, id, name };
  });
  if (idMap.size === 0) return clone;

  // All tables are renamed at once: through placeholders first, so a copy
  // holding both Sales and Sales2 never renames Sales → Sales2 → Sales22.
  const placeholder = (index: number) => `__mraiCopiedTable${index}__`;
  const renameReferences = (formula: string) => {
    if (!formula.startsWith("=")) return formula;
    let next = formula;
    renames.forEach(({ from }, index) => {
      next = renameDataTableReferences(next, from, placeholder(index));
    });
    renames.forEach(({ to }, index) => {
      next = renameDataTableReferences(next, placeholder(index), to);
    });
    return next;
  };

  const next: JsonRecord = { ...(clone as JsonRecord) };
  if (renames.length > 0) {
    nextTables = nextTables.map((table) => {
      if (!isJsonRecord(table)) return table;
      let rewritten = table;
      for (const field of ["calculatedFormulas", "totalFormulas"] as const) {
        const templates = mapRecordStrings(table[field], renameReferences);
        if (templates !== table[field]) rewritten = { ...rewritten, [field]: templates };
      }
      return rewritten;
    });
    const cells = clone.cells;
    if (isJsonRecord(cells)) {
      let changed = false;
      const nextCells: JsonRecord = {};
      for (const [key, cell] of Object.entries(cells)) {
        let nextCell = cell;
        if (isJsonRecord(cell)) {
          for (const field of ["value", "formula"] as const) {
            const value = (nextCell as JsonRecord)[field];
            if (typeof value !== "string") continue;
            const rewritten = renameReferences(value);
            if (rewritten === value) continue;
            if (nextCell === cell) nextCell = { ...cell };
            (nextCell as JsonRecord)[field] = rewritten;
          }
        }
        if (nextCell !== cell) changed = true;
        nextCells[key] = nextCell;
      }
      if (changed) next.cells = nextCells;
    }
  }
  next.dataTables = nextTables;

  if (Array.isArray(clone.slicers) && clone.slicers.length > 0) {
    next.slicers = clone.slicers.map((slicer) => {
      if (!isJsonRecord(slicer)) return slicer;
      const source = slicer.source;
      if (
        !isJsonRecord(source) ||
        source.kind !== "table" ||
        typeof source.tableId !== "string" ||
        !idMap.has(source.tableId)
      ) {
        return slicer;
      }
      return {
        ...slicer,
        ...(typeof slicer.id === "string"
          ? { id: uniqueCopyId(slicer.id, takenSlicerIds, MAX_DATA_TABLE_ID_LENGTH) }
          : {}),
        source: { ...source, tableId: idMap.get(source.tableId) },
      };
    });
  }

  // A PivotTable without its own source-sheet id reads the sheet it sits on,
  // which for the copy is the copy: its table source follows the copied table.
  const remapPivot = (pivot: unknown) =>
    isJsonRecord(pivot) &&
    pivot.sourceSheetId === undefined &&
    typeof pivot.sourceTableId === "string" &&
    idMap.has(pivot.sourceTableId)
      ? { ...pivot, sourceTableId: idMap.get(pivot.sourceTableId) }
      : pivot;
  if (clone.pivot !== undefined) next.pivot = remapPivot(clone.pivot);
  if (Array.isArray(clone.morePivots)) {
    next.morePivots = clone.morePivots.map(remapPivot);
  }
  return next as T;
}

/**
 * The one copy of `source` that every worksheet-copy route inserts (Duplicate,
 * Move or Copy with "Create a copy", Paste Worksheet): a fully independent deep
 * clone with a fresh id and a legal, unique "<name> (2)" name; the copy's
 * references to its source sheet by name redirected to the copy itself
 * (defined names excepted — workbook names are mirrored into every sheet and a
 * copy does not change them, as in Excel); and its tables given their own ids
 * and names, with the copy's structured references following them. The input
 * workbook and source are never mutated; a sheet the fail-closed cloner
 * rejects throws SheetCloneError.
 */
export function copySheetForWorkbook<T extends SheetLike>(
  source: T,
  workbook: readonly SheetLike[],
): T {
  const identities = workbook.map((candidate) => readSheetIdentity(candidate));
  const sourceName = readSheetIdentity(source).name;
  const { id, name } = duplicateSheetNaming(
    sourceName,
    new Set(identities.map((identity) => identity.id)),
    new Set(identities.map((identity) => identity.name)),
  );
  const clone = cloneSheetWithIdentity(source, id, name);
  const redirected = renameSheetReferences(clone, sourceName, name, {
    skipNamedRanges: true,
  });
  return reidentifyCopiedSheetDrawings(
    rebindCopiedPivotCharts(
      reidentifyCopiedSheetTables(
        rebindCopiedLocalNames(redirected, readSheetIdentity(source).id, sourceName, id, name),
        workbook,
      ),
      readSheetIdentity(source).id,
      id,
    ),
    workbook,
  );
}

/** The longest drawing id a chart keeps (lib/spreadsheet/chart MAX_CHART_ID_LENGTH). */
const MAX_COPIED_DRAWING_ID_LENGTH = 200;

/**
 * A copied sheet's charts, pictures, shapes and SmartArt are new objects, as
 * in Excel, where the copy's drawing objects take ids of their own: each takes
 * an id no chart or object of the workbook holds (`chart-1` → `chart-1-copy-2`;
 * a copy of that copy → `chart-1-copy-3`), so the copy's chart can later be
 * moved beside its source chart (S-082) and nothing addressed by id reaches
 * both. A chart listed both as a chart and as a drawing object keeps one id
 * across the two lists; a connector glued to a copied shape follows its new
 * id, and each group of the copy becomes a group of its own.
 */
function reidentifyCopiedSheetDrawings<T extends SheetLike>(
  clone: T,
  workbook: readonly SheetLike[],
): T {
  const charts = Array.isArray(clone.charts) ? clone.charts : null;
  const objects = Array.isArray(clone.objects) ? clone.objects : null;
  if (!charts?.length && !objects?.length) return clone;
  const takenIds = new Set<string>();
  const takenGroups = new Set<string>();
  for (const sheet of [...workbook, clone]) {
    for (const field of ["charts", "objects"] as const) {
      const items = sheet[field];
      if (!Array.isArray(items)) continue;
      for (const item of items) {
        if (!isJsonRecord(item)) continue;
        if (typeof item.id === "string") takenIds.add(item.id);
        if (typeof item.groupId === "string") takenGroups.add(item.groupId);
      }
    }
  }
  const stem = (value: string) => value.replace(/-copy-\d+$/, "") || value;
  const ids = new Map<string, string>();
  const groups = new Map<string, string>();
  const idFor = (old: string) => {
    let next = ids.get(old);
    if (next === undefined) {
      next = uniqueCopyId(stem(old), takenIds, MAX_COPIED_DRAWING_ID_LENGTH);
      ids.set(old, next);
    }
    return next;
  };
  const groupFor = (old: string) => {
    let next = groups.get(old);
    if (next === undefined) {
      next = uniqueCopyId(stem(old), takenGroups, MAX_COPIED_DRAWING_ID_LENGTH);
      groups.set(old, next);
    }
    return next;
  };
  const reidentify = (item: unknown): unknown => {
    if (!isJsonRecord(item)) return item;
    const next: JsonRecord = { ...item };
    if (typeof item.id === "string" && item.id) next.id = idFor(item.id);
    if (typeof item.groupId === "string" && item.groupId) {
      next.groupId = groupFor(item.groupId);
    }
    return next;
  };
  const nextCharts = charts?.map(reidentify);
  const nextObjects = objects?.map(reidentify).map((item) => {
    if (!isJsonRecord(item)) return item;
    let next = item;
    for (const field of ["startObjectId", "endObjectId"] as const) {
      const target = item[field];
      if (typeof target === "string" && ids.has(target)) {
        next = { ...next, [field]: ids.get(target) };
      }
    }
    return next;
  });
  return {
    ...clone,
    ...(nextCharts ? { charts: nextCharts } : {}),
    ...(nextObjects ? { objects: nextObjects } : {}),
  } as T;
}

/**
 * A copied sheet's PivotCharts of its own PivotTables plot the copy's tables,
 * as Excel's copy does; a chart of a PivotTable on another sheet keeps it.
 */
function rebindCopiedPivotCharts<T extends SheetLike>(
  sheet: T,
  sourceId: string,
  copyId: string,
): T {
  const record = sheet as unknown as {
    charts?: unknown;
    objects?: unknown;
  };
  const rebind = (item: unknown): unknown => {
    if (!item || typeof item !== "object") return item;
    const source = (item as { pivotSource?: unknown }).pivotSource;
    if (
      !source ||
      typeof source !== "object" ||
      (source as { sheetId?: unknown }).sheetId !== sourceId
    ) {
      return item;
    }
    return { ...item, pivotSource: { ...source, sheetId: copyId } };
  };
  const charts = Array.isArray(record.charts) ? record.charts.map(rebind) : undefined;
  const objects = Array.isArray(record.objects) ? record.objects.map(rebind) : undefined;
  const changed =
    (charts && charts.some((item, index) => item !== (record.charts as unknown[])[index])) ||
    (objects && objects.some((item, index) => item !== (record.objects as unknown[])[index]));
  if (!changed) return sheet;
  return {
    ...sheet,
    ...(charts ? { charts } : {}),
    ...(objects ? { objects } : {}),
  } as T;
}

/**
 * The source sheet's OWN worksheet-scoped names belong to its copy too (Excel
 * copies them with the sheet): they are rebound to the copy's id and a
 * reference of theirs into the source sheet now reads the copy (live audit
 * B113). Workbook names and other sheets' local names stay as they are.
 */
function rebindCopiedLocalNames<T extends SheetLike>(
  sheet: T,
  sourceId: string,
  sourceName: string,
  copyId: string,
  copyName: string,
): T {
  const names = sheet.namedRanges;
  if (!Array.isArray(names) || names.length === 0) return sheet;
  let changed = false;
  const rebound = names.map((entry) => {
    if (
      !isJsonRecord(entry) ||
      entry.scope !== "worksheet" ||
      entry.scopeSheetId !== sourceId
    ) {
      return entry;
    }
    changed = true;
    // The copy's local name is a new name of the copy: the file's own name it
    // was copied from stays the source sheet's (WALK R270-X19: two page names
    // claimed one native name, defined_name_native_source_ambiguous).
    const rest = Object.fromEntries(
      Object.entries(entry).filter(([key]) => key !== "nativeSourceName" && key !== "nativeSourceScopeSheetId"),
    );
    return {
      ...rest,
      scopeSheetId: copyId,
      ...(typeof entry.ref === "string"
        ? { ref: renameSheetInFormula(entry.ref, sourceName, copyName) }
        : {}),
    };
  });
  return changed ? { ...sheet, namedRanges: rebound } : sheet;
}

/**
 * Duplicate the sheet at `index`: deep-clone it (fully independent cells/styles/
 * overlays), give it a fresh de-duplicated id + name, and insert the copy
 * directly AFTER the source. Returns `{ sheets, insertedIndex }` so the caller can
 * select the new tab; an out-of-range index returns the input array unchanged
 * with `insertedIndex = -1`. The input array and its sheets are never mutated.
 */
export function duplicateSheet<T extends SheetLike>(
  sheets: T[],
  index: number,
): { sheets: T[]; insertedIndex: number } {
  if (index < 0 || index >= sheets.length) {
    return { sheets, insertedIndex: -1 };
  }
  const source = sheets[index];
  // Validate and detach the source before deriving identity. In particular,
  // this keeps a hostile accessor/toJSON hook from running during naming and
  // means failure cannot expose a half-built sheet to the caller.
  const { content, blankCellFormats } = detachSheetRuntimeHints(source);
  const snapshot = cloneSheetSnapshot(content);
  const sourceIdentity = readSheetIdentity(snapshot.value);
  const identities = sheets.map((candidate, candidateIndex) =>
    candidateIndex === index ? sourceIdentity : readSheetIdentity(candidate),
  );
  const takenIds = new Set(identities.map((identity) => identity.id));
  const takenNames = new Set(identities.map((identity) => identity.name));
  const { id, name } = duplicateSheetNaming(
    sourceIdentity.name,
    takenIds,
    takenNames,
  );
  const clone = reattachSheetRuntimeHints(
    applyClonedSheetIdentity(snapshot, id, name),
    blankCellFormats,
  );
  const next = sheets.slice();
  const insertedIndex = index + 1;
  next.splice(insertedIndex, 0, clone);
  return { sheets: next, insertedIndex };
}
