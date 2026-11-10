/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Pure helpers for spreadsheet named ranges (Excel/WPS parity).
 *
 * A named range maps a human-friendly name (e.g. "Sales") to a cell or range
 * reference (e.g. "B2:B10"), so formulas can read `=SUM(Sales)` instead of
 * `=SUM(B2:B10)`. The workbook/sheet model stores the names as a list; this
 * module keeps name validation, normalization and lookup out of the React
 * component and out of the formula engine so each can be unit-tested in
 * isolation and so the manager UI, the persistence layer and the formula
 * resolver all agree on what a valid name is and how it resolves.
 *
 * Resolution model: when the formula engine meets a bare identifier that is not
 * a function call and not TRUE/FALSE, it asks for a name resolution. A name
 * resolves to its stored reference, which is then parsed as a single cell or a
 * range exactly like a literal `B2` or `B2:B10` would be. Unknown names yield
 * a #NAME? error, matching Excel.
 */

import { isKnownFunction } from "./formula-engine";
import { storedNameRefFromTyped, typedNameRefHasRelativeAxis } from "./name-relative-refs";
import {
  BARE_NAME_PART_CLASS,
  characterAt,
  definedNameEnd,
  formatSheetNameForFormula,
  isBareNameSpelling,
  isDefinedNameCharacter,
  isDefinedNameSpelling,
  looksLikeCellReference,
} from "./name-syntax";

// Regex source of every character a name token can hold (`u` flag): the
// defined-name rule's later characters, `?` and `\` included (./name-syntax).
const NAME_TOKEN_CLASS = `${BARE_NAME_PART_CLASS}?\\\\`;

export type NamedRangeScope = "workbook" | "worksheet";

/** Scope selector used by mutations and formula resolution. */
export interface NamedRangeScopeContext {
  /** Missing scope keeps the legacy workbook-scoped behavior. */
  scope?: NamedRangeScope;
  /** Required for worksheet-scoped names; compared case-insensitively. */
  scopeSheetId?: string;
}

/** A defined name and the reference it expands to. */
export interface NamedRange {
  /** The name as the user typed it (case preserved for display). */
  name: string;
  /**
   * What the name refers to: an A1 cell ("B2"), a range ("B2:B10"), OR a FORMULA
   * / LAMBDA ("LAMBDA(x, x*2)", "0.2", "SUM(A:A)") — the named-formula / named-
   * LAMBDA "macro". Cell/range refs are uppercased; a formula ref keeps its case.
   */
  ref: string;
  /** Missing on legacy content and therefore semantically workbook-scoped. */
  scope?: NamedRangeScope;
  /** Stable editor sheet id for a worksheet-scoped name. */
  scopeSheetId?: string;
  /**
   * Original-package names with semantics outside the owned classic grammar
   * remain visible/read-only (`_xlnm` built-ins are not projected at all). The
   * native writer independently re-authenticates this marker; it is never a
   * client grant to edit opaque OOXML.
   */
  nativeBlocker?:
    | "external-reference"
    | "unsupported-attributes"
    | "ambiguous-identity";
  /**
   * Stable source identity for an editable name imported from workbook.xml.
   * Rename keeps this identity so every original-package save can replay the
   * same surgical source -> target mutation against immutable source bytes.
   */
  nativeSourceName?: string;
  nativeSourceScopeSheetId?: string;
  /**
   * Excel's `hidden` flag. It only keeps the name out of Name Manager (and the
   * other name lists); formulas still resolve it, so it is part of the
   * calculation catalog and is written back as `hidden="1"`.
   */
  hidden?: true;
  /**
   * Name Manager's Comment (XLS-A066): up to 255 characters about the name,
   * written to workbook.xml as the definedName `comment` attribute. Absent
   * when empty.
   */
  comment?: string;
}

/** Excel caps a name's comment at 255 characters. */
export const MAX_NAME_COMMENT_LENGTH = 255;

// A1 cell reference (optionally $-absolute), used both to reject names that look
// like cell refs and to detect single-cell vs. range references. Mirrors the
// formula engine's CELL_REF_RE.
const CELL_REF_RE = /^\$?[A-Za-z]{1,3}\$?[0-9]{1,7}$/;

// A valid range reference: two cell refs joined by a colon.
const RANGE_REF_RE = /^\$?[A-Za-z]{1,3}\$?[0-9]{1,7}:\$?[A-Za-z]{1,3}\$?[0-9]{1,7}$/;

function absoluteCellRef(ref: string): string | null {
  const match = /^\$?([A-Za-z]{1,3})\$?([0-9]{1,7})$/.exec(ref.trim());
  return match ? `$${match[1].toUpperCase()}$${match[2]}` : null;
}

// The boolean literals read as values, never as a name. (The R1C1 shorthands
// "R" and "C" are refused as cell references by looksLikeCellReference.)
const RESERVED_NAMES = new Set(["TRUE", "FALSE"]);

/** Maximum length of a defined name (Excel caps at 255). */
export const MAX_NAME_LENGTH = 255;

/**
 * Validate a candidate name against Microsoft Support, "Names in formulas". A
 * name is valid when it:
 *  - is non-empty and at most MAX_NAME_LENGTH characters,
 *  - starts with a letter of any script, an underscore or a backslash, and
 *    continues with letters, digits, underscores or periods (no spaces or
 *    operators) — the formula tokenizer reads exactly that set as one name,
 *  - is not a cell reference, A1 ("A1", "$B$3") or R1C1 ("R1C1", "RC", and
 *    the "R" / "C" shorthands),
 *  - is not TRUE / FALSE and not in Excel's reserved `_xlnm.` namespace.
 * Every comparison is case-insensitive, as Excel's names are.
 */
export function isValidName(name: string): boolean {
  return typeof name === "string" && definedNameProblem(name) === null;
}

/** Why a candidate name is not a valid name — one reason per rule. */
export type DefinedNameProblem =
  | "empty"
  | "too-long"
  | "start"
  | "character"
  | "cell-reference"
  | "reserved";

/**
 * The first rule of {@link isValidName} that `name` breaks, or null when it is
 * a valid name, so a dialog can say why rather than only refuse. A leading
 * character outside the set is "start"; a space or other character outside
 * the set later on is "character".
 */
export function definedNameProblem(name: string): DefinedNameProblem | null {
  const trimmed = name.trim();
  if (trimmed.length === 0) return "empty";
  if (trimmed.length > MAX_NAME_LENGTH) return "too-long";
  if (!isDefinedNameSpelling(trimmed)) {
    return definedNameEnd(trimmed, 0) === 0 ? "start" : "character";
  }
  if (trimmed.toLowerCase().startsWith("_xlnm.")) return "reserved";
  if (looksLikeCellReference(trimmed)) return "cell-reference";
  if (RESERVED_NAMES.has(trimmed.toUpperCase())) return "reserved";
  return null;
}

/**
 * The message key that says why a name was refused, worded after Excel's "The
 * name that you entered is not valid" / "already exists" prompts. `exists` is
 * the caller's own duplicate check (a name is unique within its scope).
 */
export function definedNameProblemMessageKey(
  problem: DefinedNameProblem | "exists",
): string {
  switch (problem) {
    case "empty":
      return "spreadsheet.names.nameEmpty";
    case "too-long":
      return "spreadsheet.names.nameTooLong";
    case "start":
      return "spreadsheet.names.nameInvalidStart";
    case "character":
      return "spreadsheet.names.nameInvalidCharacter";
    case "cell-reference":
      return "spreadsheet.names.nameIsReference";
    case "reserved":
      return "spreadsheet.names.nameReserved";
    case "exists":
      return "spreadsheet.names.nameExists";
  }
}

/** True when a reference string is a single A1 cell ("B2"/"$B$2"). */
export function isCellRefString(ref: string): boolean {
  return CELL_REF_RE.test(ref.trim());
}

/** True when a reference string is a range ("B2:B10"). */
export function isRangeRefString(ref: string): boolean {
  return RANGE_REF_RE.test(ref.trim());
}

/** Max length of a named FORMULA / LAMBDA "refers to" (guards persisted size). */
export const MAX_FORMULA_REF_LENGTH = 8192;

/**
 * True when a reference string is a FORMULA / LAMBDA rather than a plain cell or
 * range — Excel's "named formula" / named LAMBDA (e.g. `LAMBDA(x, x*2)`, `0.2`,
 * `SUM(A:A)`). The modern, VBA-free way to define a reusable macro function. Any
 * non-empty, length-bounded string that is NOT a cell/range qualifies; the
 * formula engine validates parseability at evaluation time (a bad formula reads
 * as #NAME?/#VALUE!, it can never corrupt anything).
 */
export function isFormulaRefString(ref: string): boolean {
  const trimmed = ref.trim();
  return (
    trimmed.length > 0 &&
    trimmed.length <= MAX_FORMULA_REF_LENGTH &&
    !isCellRefString(trimmed) &&
    !isRangeRefString(trimmed)
  );
}

/** True when a reference string is a valid cell, range, OR named formula. */
export function isValidRef(ref: string): boolean {
  return (
    isCellRefString(ref) || isRangeRefString(ref) || isFormulaRefString(ref)
  );
}

/**
 * Bind a UI-authored name's plain A1 target to its originating sheet, in the
 * form Excel stores and shows it: absolute and sheet-qualified, the sheet name
 * quoted only when a formula would quote it (`回款!$C$4:$C$10`,
 * `'Q1''s Plan'!$C$1`). Workbook names are copied into every sheet projection,
 * so leaving `C1` bare would make the same name point at a different cell on
 * each sheet. Constants and non-relative formulas/LAMBDAs stay unchanged. A
 * typed relative reference is read from `base`, the active cell, and stored
 * as Excel stores it (lib/spreadsheet/name-relative-refs); with no base it
 * fails closed.
 */
export function qualifyWorkbookNamedRangeRef(
  ref: string,
  sheetName: string,
  options: {
    authored?: boolean;
    /**
     * The cell active when the reference was typed. A relative axis the user
     * wrote is relative to it, as Excel's New / Edit Name reads it, and is
     * stored as its offset from A1 (lib/spreadsheet/name-relative-refs).
     */
    base?: { row: number; col: number };
  } = {},
): string | null {
  const trimmed = ref.trim();
  if (!sheetName || /[\u0000-\u001f]/.test(sheetName)) return null;
  const hasFormulaPrefix = trimmed.startsWith("=");
  const body = hasFormulaPrefix ? trimmed.slice(1).trim() : trimmed;
  const prefix = hasFormulaPrefix ? "=" : "";
  // A reference the user typed keeps the axes they wrote. A relative one is
  // read from the active cell; with no active cell to read it from (an AI
  // command naming only a sheet) it is refused, never made absolute.
  // With the active cell known, a typed formula's own references are stored
  // as Excel stores them too: each on its sheet, relative axes as offsets.
  if (options.authored) {
    const relative = typedNameRefHasRelativeAxis(body);
    if (relative && !options.base) return null;
    const plain = isCellRefString(body) || isRangeRefString(body);
    if (options.base && (relative || !plain)) {
      if (!isValidRef(body)) return null;
      return `${prefix}${storedNameRefFromTyped(body, `${formatSheetNameForFormula(sheetName)}!`, options.base)}`;
    }
  }
  const parts = body.split(":");
  if (parts.length === 1 && isCellRefString(body)) {
    const cell = absoluteCellRef(body);
    if (!cell) return null;
    return `${prefix}${formatSheetNameForFormula(sheetName)}!${cell}`;
  }
  if (parts.length === 2 && isRangeRefString(body)) {
    const from = absoluteCellRef(parts[0]);
    const to = absoluteCellRef(parts[1]);
    if (!from || !to) return null;
    return `${prefix}${formatSheetNameForFormula(sheetName)}!${from}:${to}`;
  }
  if (!isFormulaRefString(body)) return null;
  // We have no persisted base-cell authority for a workbook formula. Preserve
  // constants/LAMBDAs only when the authored expression has no token that can
  // be interpreted as a relative A1 cell, row range, or column range. A run of
  // name characters around it (`金额A1`, `\A1`) is one name, not a cell.
  //
  // Read as a formula reads it (O016): text inside "..." is a string, not a
  // cell ("A1" stays two characters), and a sheet-qualified ABSOLUTE
  // reference (Data!$A$1, Data!$A$1:$A$2, Data!$A:$A) needs no base cell --
  // it is the form Define Name itself stores a selection in. Anything else
  // that reads as a cell or a whole row/column still needs a base and is
  // refused.
  const blank = (literal: string) => " ".repeat(literal.length);
  const code = body
    .replace(/"(?:[^"]|"")*"/g, blank)
    // A quoted sheet name is a name ('A1 Data'!$B$2), its "!" kept.
    .replace(/'(?:[^']|'')*'(?=!)/g, blank);
  const cellToken = new RegExp(
    // LOG10( is a function, not cell LOG10.
    `(^|[^${NAME_TOKEN_CLASS}])(\\$?)[A-Za-z]{1,3}(\\$?)[0-9]{1,7}(?![${NAME_TOKEN_CLASS}(])`,
    "gu",
  );
  const wholeLineToken = new RegExp(
    `(^|[^${NAME_TOKEN_CLASS}])(?:(\\$?)[A-Za-z]{1,3}:(\\$?)[A-Za-z]{1,3}|(\\$?)[0-9]{1,7}:(\\$?)[0-9]{1,7})(?![${NAME_TOKEN_CLASS}])`,
    "gu",
  );
  /** Whether the token starting at `index` sits after a sheet qualifier. */
  const qualifiedAt = (index: number) => {
    const before = code.slice(0, index);
    return before.endsWith("!") || /![^,;()=+\-*/^&<>\s]*:$/.test(before);
  };
  for (const match of code.matchAll(cellToken)) {
    const tokenStart = (match.index ?? 0) + match[1].length;
    // Data!$A$1:$A$2 -- the second corner follows the colon of a qualified
    // first corner.
    const absolute = match[2] === "$" && match[3] === "$";
    if (!absolute || !qualifiedAt(tokenStart)) return null;
  }
  for (const match of code.matchAll(wholeLineToken)) {
    const tokenStart = (match.index ?? 0) + match[1].length;
    const dollars = match.slice(2).filter((part) => part !== undefined);
    if (!dollars.every((part) => part === "$") || !qualifiedAt(tokenStart)) {
      return null;
    }
  }
  return `${prefix}${body}`;
}

/**
 * Coerce a single raw entry into a normalized NamedRange, or null when the name
 * is invalid or the reference is not a cell/range. The stored ref is uppercased
 * (A1 letters are case-insensitive) and trimmed.
 */
export function normalizeNamedRange(raw: unknown): NamedRange | null {
  if (!raw || typeof raw !== "object") return null;
  const r = raw as Record<string, unknown>;
  const name = typeof r.name === "string" ? r.name.trim() : "";
  const rawRef = typeof r.ref === "string" ? r.ref.trim() : "";
  if (!isValidName(name)) return null;
  const hasExplicitScope = r.scope !== undefined;
  if (
    hasExplicitScope &&
    r.scope !== "workbook" &&
    r.scope !== "worksheet"
  ) {
    return null;
  }
  const scope: NamedRangeScope =
    r.scope === "worksheet" ? "worksheet" : "workbook";
  const scopeSheetId =
    typeof r.scopeSheetId === "string" ? r.scopeSheetId.trim() : "";
  if (scope === "worksheet" && !scopeSheetId) return null;
  // A workbook name carrying a sheet identity is ambiguous. Reject instead of
  // silently laundering it into a different scope.
  if (scope === "workbook" && scopeSheetId) return null;
  const scopeFields =
    scope === "worksheet"
      ? ({ scope, scopeSheetId } as const)
      : hasExplicitScope
        ? ({ scope } as const)
        : {};
  const rawNativeBlocker = r.nativeBlocker;
  const nativeBlocker: NamedRange["nativeBlocker"] =
    rawNativeBlocker === "external-reference" ||
    rawNativeBlocker === "unsupported-attributes" ||
    rawNativeBlocker === "ambiguous-identity"
      ? rawNativeBlocker
      : undefined;
  if (rawNativeBlocker !== undefined && nativeBlocker === undefined) return null;
  const nativeSourceName =
    typeof r.nativeSourceName === "string" ? r.nativeSourceName.trim() : "";
  if (nativeSourceName && !isValidName(nativeSourceName)) return null;
  if (r.nativeSourceName !== undefined && !nativeSourceName) return null;
  const nativeSourceScopeSheetId =
    typeof r.nativeSourceScopeSheetId === "string"
      ? r.nativeSourceScopeSheetId.trim()
      : "";
  if (
    r.nativeSourceScopeSheetId !== undefined &&
    (!nativeSourceName || !nativeSourceScopeSheetId || scope !== "worksheet")
  ) {
    return null;
  }
  if (scope === "workbook" && nativeSourceScopeSheetId) return null;
  if (r.hidden !== undefined && typeof r.hidden !== "boolean") return null;
  if (r.comment !== undefined && typeof r.comment !== "string") return null;
  const comment = typeof r.comment === "string" ? r.comment.trim() : "";
  if (comment.length > MAX_NAME_COMMENT_LENGTH) return null;
  const nativeFields = {
    ...(nativeBlocker ? { nativeBlocker } : {}),
    ...(nativeSourceName ? { nativeSourceName } : {}),
    ...(nativeSourceScopeSheetId ? { nativeSourceScopeSheetId } : {}),
    ...(r.hidden === true ? { hidden: true as const } : {}),
    ...(comment ? { comment } : {}),
  };
  // A cell / range ref is uppercased (A1 letters are case-insensitive). A FORMULA
  // ref keeps its case VERBATIM — uppercasing it would corrupt string literals
  // (="hi") and is unnecessary (the engine is case-insensitive for names / refs /
  // functions / LAMBDA params, which all match via toUpperCase).
  if (isCellRefString(rawRef) || isRangeRefString(rawRef)) {
    return {
      name,
      ref: rawRef.toUpperCase(),
      ...scopeFields,
      ...nativeFields,
    };
  }
  if (isFormulaRefString(rawRef)) {
    return { name, ref: rawRef, ...scopeFields, ...nativeFields };
  }
  return null;
}

/**
 * Normalize a raw namedRanges array off a persisted sheet/workbook: drop
 * invalid entries and de-duplicate by case-insensitive name (last one wins,
 * matching the add-overwrites-existing behavior). Returns a fresh array.
 */
export function normalizeNamedRanges(raw: unknown): NamedRange[] {
  if (!Array.isArray(raw)) return [];
  const byName = new Map<string, NamedRange>();
  for (const item of raw) {
    const entry = normalizeNamedRange(item);
    if (!entry) continue;
    byName.set(namedRangeIdentity(entry), entry);
  }
  return [...byName.values()];
}

function normalizedScopeSheetId(value: string | undefined): string {
  return value?.trim().toUpperCase() ?? "";
}

export function namedRangeIdentity(entry: NamedRange): string {
  const scope = entry.scope ?? "workbook";
  return `${entry.name.toUpperCase()}\u0000${scope}\u0000${
    scope === "worksheet"
      ? normalizedScopeSheetId(entry.scopeSheetId)
      : ""
  }`;
}

/**
 * Whether a name belongs in the lists a user browses — Name Manager, the
 * Navigation pane. A hidden name is left out of them, as Excel does, but it
 * stays in the catalog every formula resolves against.
 */
export function isListedNamedRange(entry: NamedRange): boolean {
  return entry.hidden !== true;
}

/**
 * Look up a name (case-insensitive) and return its stored reference, or null
 * when no such name is defined. Used by the formula engine's name resolver.
 */
export function resolveName(
  name: string,
  names: readonly NamedRange[],
  scopeSheetId?: string,
): string | null {
  const key = name.trim().toUpperCase();
  const sheetKey = normalizedScopeSheetId(scopeSheetId);
  if (sheetKey) {
    for (const entry of names) {
      if (
        !entry.nativeBlocker &&
        entry.name.toUpperCase() === key &&
        entry.scope === "worksheet" &&
        normalizedScopeSheetId(entry.scopeSheetId) === sheetKey
      ) {
        return entry.ref;
      }
    }
  } else {
    // Existing formula-engine callers receive the current sheet's name list but
    // do not yet pass its id. The importer only places that sheet's local names
    // in the list, so a unique local match remains authoritative and shadows
    // the workbook name. Ambiguous hand-authored arrays fail back to global.
    const localMatches = names.filter(
      (entry) =>
        !entry.nativeBlocker &&
        entry.name.toUpperCase() === key && entry.scope === "worksheet",
    );
    if (localMatches.length === 1) return localMatches[0].ref;
  }
  for (const entry of names) {
    if (
      !entry.nativeBlocker &&
      entry.name.toUpperCase() === key &&
      (entry.scope === undefined || entry.scope === "workbook")
    ) {
      return entry.ref;
    }
  }
  return null;
}

/**
 * Add (or overwrite) a named range. Throws if the name or ref is invalid so the
 * caller surfaces a validation error to the user. De-duplicates by
 * case-insensitive name. Returns a fresh array.
 */
export function addNamedRange(
  names: readonly NamedRange[],
  name: string,
  ref: string,
  context: NamedRangeScopeContext = {},
): NamedRange[] {
  const entry = normalizeNamedRange({ name, ref, ...context });
  if (!entry) {
    throw new Error("Invalid named range");
  }
  const key = namedRangeIdentity(entry);
  if (
    names.some((candidate) => {
      const normalized = normalizeNamedRange(candidate);
      return Boolean(
        normalized?.nativeBlocker && namedRangeIdentity(normalized) === key,
      );
    })
  ) {
    throw new Error("Native named range is read-only");
  }
  const kept = names.filter((candidate) => {
    const normalized = normalizeNamedRange(candidate);
    return !normalized || namedRangeIdentity(normalized) !== key;
  });
  return [...kept, entry];
}

/** Remove a named range by case-insensitive name. Returns a fresh array. */
export function removeNamedRange(
  names: readonly NamedRange[],
  name: string,
  context: NamedRangeScopeContext = {},
): NamedRange[] {
  const scope = context.scope ?? "workbook";
  const key = `${name.trim().toUpperCase()}\u0000${scope}\u0000${
    scope === "worksheet"
      ? normalizedScopeSheetId(context.scopeSheetId)
      : ""
  }`;
  return names.filter((candidate) => {
    const normalized = normalizeNamedRange(candidate);
    return (
      !normalized ||
      Boolean(normalized.nativeBlocker) ||
      namedRangeIdentity(normalized) !== key
    );
  });
}

/** True when a name (case-insensitive) is already defined. */
export function hasName(
  name: string,
  names: readonly NamedRange[],
  context: NamedRangeScopeContext = {},
): boolean {
  const scope = context.scope ?? "workbook";
  const key = `${name.trim().toUpperCase()}\u0000${scope}\u0000${
    scope === "worksheet"
      ? normalizedScopeSheetId(context.scopeSheetId)
      : ""
  }`;
  return names.some((candidate) => {
    const normalized = normalizeNamedRange(candidate);
    return Boolean(normalized && namedRangeIdentity(normalized) === key);
  });
}

/** Whether the character at `source[index]` continues a name token. */
function isNameTokenCharacter(source: string, index: number): boolean {
  return isDefinedNameCharacter(characterAt(source, index), false);
}

/**
 * Rename one exact defined-name reference without touching string literals,
 * quoted sheet names, structured-reference selectors, similarly prefixed
 * identifiers, or a worksheet name before `!`. The scanner deliberately works
 * on the formula source instead of display text so workbook-wide rename is one
 * atomic content transaction and remains case-insensitive like Excel.
 *
 * `qualifierSheetName` also rewrites an explicitly-qualified local reference
 * (`'Sheet 1'!LocalRate`). Unqualified local references are handled by the
 * caller only on their owning sheet; workbook references are handled on every
 * sheet that does not carry a same-name local shadow.
 */
export function renameDefinedNameReferences(
  source: string,
  oldName: string,
  newName: string,
  qualifierSheetName?: string,
  rewriteUnqualified = true,
): string {
  if (!source || !oldName || oldName === newName) return source;
  const oldLower = oldName.toLowerCase();
  const quotedQualifier = qualifierSheetName
    ? `'${qualifierSheetName.replace(/'/g, "''")}'!`
    : "";
  const plainQualifier =
    qualifierSheetName && isBareNameSpelling(qualifierSheetName)
      ? `${qualifierSheetName}!`
      : "";
  let output = "";
  let index = 0;
  let inString = false;
  let structuredDepth = 0;
  while (index < source.length) {
    const character = source[index];
    if (character === '"') {
      if (inString && source[index + 1] === '"') {
        output += '""';
        index += 2;
        continue;
      }
      inString = !inString;
      output += character;
      index += 1;
      continue;
    }
    if (inString) {
      output += character;
      index += 1;
      continue;
    }
    if (character === "[") structuredDepth += 1;
    if (character === "]" && structuredDepth > 0) structuredDepth -= 1;
    if (structuredDepth > 0) {
      output += character;
      index += 1;
      continue;
    }
    if (character === "'") {
      let end = index + 1;
      while (end < source.length) {
        if (source[end] !== "'") {
          end += 1;
          continue;
        }
        if (source[end + 1] === "'") {
          end += 2;
          continue;
        }
        end += 1;
        break;
      }
      const quoted = source.slice(index, end);
      const qualifier = source[end] === "!" ? `${quoted}!` : quoted;
      output += qualifier;
      index = end + (source[end] === "!" ? 1 : 0);
      if (
        quotedQualifier &&
        qualifier.toLowerCase() === quotedQualifier.toLowerCase() &&
        source.slice(index, index + oldName.length).toLowerCase() === oldLower &&
        !isNameTokenCharacter(source, index + oldName.length)
      ) {
        output += newName;
        index += oldName.length;
      }
      continue;
    }
    if (
      plainQualifier &&
      source.slice(index, index + plainQualifier.length).toLowerCase() ===
        plainQualifier.toLowerCase()
    ) {
      output += source.slice(index, index + plainQualifier.length);
      index += plainQualifier.length;
      if (
        source.slice(index, index + oldName.length).toLowerCase() === oldLower &&
        !isNameTokenCharacter(source, index + oldName.length)
      ) {
        output += newName;
        index += oldName.length;
      }
      continue;
    }
    // A name token is the same run the formula tokenizer reads as one name
    // (`金额`, `\Rate`, `Sales金额`), so a prefix of a longer name never matches.
    const end = definedNameEnd(source, index);
    if (end === index) {
      output += character;
      index += 1;
      continue;
    }
    const token = source.slice(index, end);
    let previous = index - 1;
    while (/\s/.test(source[previous] ?? "")) previous -= 1;
    let next = end;
    while (/\s/.test(source[next] ?? "")) next += 1;
    // A call of a built-in function spelled like the name (RATE( beside a
    // name Rate) is the function, not the name: Excel reads a name followed
    // by "(" as a function call, and renaming the name leaves it alone.
    const builtInCall = source[next] === "(" && isKnownFunction(token);
    output +=
      rewriteUnqualified &&
      token.toLowerCase() === oldLower &&
      source[previous] !== "!" &&
      source[next] !== "!" &&
      !builtInCall
        ? newName
        : token;
    index = end;
  }
  return output;
}
