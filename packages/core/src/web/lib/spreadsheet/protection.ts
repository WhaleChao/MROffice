/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Pure helpers for sheet & cell protection (Excel/WPS parity).
 *
 * Excel model, reproduced here: every cell is "locked" by default; the lock only
 * takes EFFECT once the sheet itself is protected. So a protected sheet is
 * read-only everywhere EXCEPT the cells a user has explicitly unlocked to carve
 * out an editable region. An unprotected sheet is always fully editable
 * regardless of any per-cell `locked` flag.
 *
 * SECURITY CAVEAT — this is a WORKFLOW guard, NOT a security boundary. It mirrors
 * Excel's sheet protection, which is likewise trivially bypassed: a determined
 * user can hand-edit the persisted JSON to flip `protected`/`locked`, and the
 * grid never enforces protection server-side. The point is UX / accident
 * prevention (stop a collaborator from clobbering a formula column), not access
 * control. Do not rely on it to keep anyone out of data they could otherwise
 * read or load.
 *
 * Model:
 *  - `sheet.protected === true`  ⇒ the sheet is in read-only-guard mode.
 *  - `cell.locked` is tri-state: `true`/absent ⇒ locked (default), `false` ⇒
 *    unlocked. Storing `locked: true` is redundant with the default, so callers
 *    persist ONLY `locked: false` (and prune the key otherwise) to keep an
 *    unprotected / never-touched sheet byte-for-byte identical.
 *
 * This module keeps the predicate + range mutation out of the React component so
 * the cell-commit / clear / paste / structural enforcement points can all share
 * one definition and it stays unit-testable in isolation.
 */

import { expandLineStyleRuns, type LineStyleMap } from "./line-styles";
import { legacyPasswordByteCandidates } from "./legacy-password-code-pages";
import {
  checkProtectionPasswordHash,
  normalizeProtectionPasswordHash,
  type ProtectionPasswordHash,
} from "./protection-password";
import { editRangeOpensCell, type ProtectedEditRange } from "./protected-ranges";

/** The minimal shape a cell needs for the lock predicate (a real cell has more). */
export interface ProtectableCell {
  locked?: boolean;
  /** Format Cells > Protection > Hidden: takes effect only while protected. */
  formulaHidden?: boolean;
  [key: string]: unknown;
}

/**
 * The stored Unprotect password verifier: Excel's ISO hash (what MRAI writes
 * for a new password) or the 16-bit legacy verifier an older file carries.
 * Never the plaintext.
 */
export type SheetProtectionPassword =
  | ({ kind: "hash" } & ProtectionPasswordHash)
  | { kind: "legacy"; verifier: string };

/**
 * Native worksheet permissions that remain meaningful while protected
 * (Excel's Protect Sheet dialog, "Allow all users of this worksheet to").
 *
 * Every key is ABSENT at Excel's dialog default, so a sheet protected with the
 * defaults stores nothing: the `allow*` permissions default to off, and the
 * two selection permissions default to on, so they are stored as `deny*`.
 */
export interface SheetProtectionOptions {
  /** Excel Allow Sort: sorting is still limited to fully unlocked ranges. */
  allowSort?: true;
  /** Excel "Format cells": formatting commands stay available while the
   *  sheet is protected. Absent, every format write is refused, an unlocked
   *  cell's included (unlocking grants entry, not formatting; live audit
   *  B138). */
  allowFormatCells?: true;
  /** Column width, AutoFit, standard width, hide/unhide columns. */
  allowFormatColumns?: true;
  /** Row height, AutoFit, hide/unhide rows. */
  allowFormatRows?: true;
  /** Insert whole columns (no lock condition). */
  allowInsertColumns?: true;
  /** Insert whole rows (no lock condition). */
  allowInsertRows?: true;
  /** New hyperlinks, in unlocked cells. */
  allowInsertHyperlinks?: true;
  /** Delete whole columns whose every cell is unlocked. */
  allowDeleteColumns?: true;
  /** Delete whole rows whose every cell is unlocked. */
  allowDeleteRows?: true;
  /** Change criteria of an existing AutoFilter (never add/remove it). */
  allowAutoFilter?: true;
  /** Use PivotTable & PivotChart. */
  allowPivotTables?: true;
  /** Shapes, pictures, charts, text boxes, slicers; notes. */
  allowEditObjects?: true;
  /** Scenario Manager add/edit/delete; hidden scenarios. */
  allowEditScenarios?: true;
  /** Select locked cells turned OFF. */
  denySelectLockedCells?: true;
  /** Select unlocked cells turned OFF. */
  denySelectUnlockedCells?: true;
  /** Unprotect Sheet asks for this password. */
  password?: SheetProtectionPassword;
}

/** One checkbox of Excel's Protect Sheet permission list. */
export type SheetProtectionPermission =
  | "selectLockedCells"
  | "selectUnlockedCells"
  | "formatCells"
  | "formatColumns"
  | "formatRows"
  | "insertColumns"
  | "insertRows"
  | "insertHyperlinks"
  | "deleteColumns"
  | "deleteRows"
  | "sort"
  | "autoFilter"
  | "pivotTables"
  | "editObjects"
  | "editScenarios";

export type SheetProtectionPermissions = Record<SheetProtectionPermission, boolean>;

/** Excel's dialog order. */
export const SHEET_PROTECTION_PERMISSION_ORDER: readonly SheetProtectionPermission[] = [
  "selectLockedCells",
  "selectUnlockedCells",
  "formatCells",
  "formatColumns",
  "formatRows",
  "insertColumns",
  "insertRows",
  "insertHyperlinks",
  "deleteColumns",
  "deleteRows",
  "sort",
  "autoFilter",
  "pivotTables",
  "editObjects",
  "editScenarios",
];

/** Excel's Protect Sheet defaults: everyone may select; nothing else. */
export const EXCEL_DEFAULT_SHEET_PROTECTION_PERMISSIONS: Readonly<SheetProtectionPermissions> =
  Object.freeze({
    selectLockedCells: true,
    selectUnlockedCells: true,
    formatCells: false,
    formatColumns: false,
    formatRows: false,
    insertColumns: false,
    insertRows: false,
    insertHyperlinks: false,
    deleteColumns: false,
    deleteRows: false,
    sort: false,
    autoFilter: false,
    pivotTables: false,
    editObjects: false,
    editScenarios: false,
  });

/** What an unprotected (or paused) sheet allows: everything. */
export const UNPROTECTED_SHEET_PERMISSIONS: Readonly<SheetProtectionPermissions> =
  Object.freeze(
    Object.fromEntries(
      SHEET_PROTECTION_PERMISSION_ORDER.map((permission) => [permission, true]),
    ) as SheetProtectionPermissions,
  );

type AllowPermission = Exclude<
  SheetProtectionPermission,
  "selectLockedCells" | "selectUnlockedCells"
>;

const ALLOW_OPTION_KEYS: ReadonlyArray<
  readonly [AllowPermission, keyof SheetProtectionOptions]
> = [
  ["formatCells", "allowFormatCells"],
  ["formatColumns", "allowFormatColumns"],
  ["formatRows", "allowFormatRows"],
  ["insertColumns", "allowInsertColumns"],
  ["insertRows", "allowInsertRows"],
  ["insertHyperlinks", "allowInsertHyperlinks"],
  ["deleteColumns", "allowDeleteColumns"],
  ["deleteRows", "allowDeleteRows"],
  ["sort", "allowSort"],
  ["autoFilter", "allowAutoFilter"],
  ["pivotTables", "allowPivotTables"],
  ["editObjects", "allowEditObjects"],
  ["editScenarios", "allowEditScenarios"],
];

/** Every permission resolved against Excel's defaults. */
export function sheetProtectionPermissions(
  options: SheetProtectionOptions | undefined,
): SheetProtectionPermissions {
  const permissions: SheetProtectionPermissions = {
    ...EXCEL_DEFAULT_SHEET_PROTECTION_PERMISSIONS,
  };
  for (const [permission, key] of ALLOW_OPTION_KEYS) {
    permissions[permission] = options?.[key] === true;
  }
  permissions.selectLockedCells = options?.denySelectLockedCells !== true;
  permissions.selectUnlockedCells = options?.denySelectUnlockedCells !== true;
  return permissions;
}

/** The minimal persisted options for `permissions` (+ password), or
 *  undefined when they are exactly Excel's defaults with no password. */
export function sheetProtectionOptionsFrom(
  permissions: SheetProtectionPermissions,
  password?: SheetProtectionPassword,
): SheetProtectionOptions | undefined {
  const options: Record<string, unknown> = {};
  for (const [permission, key] of ALLOW_OPTION_KEYS) {
    if (permissions[permission]) options[key] = true;
  }
  if (!permissions.selectLockedCells) options.denySelectLockedCells = true;
  if (!permissions.selectUnlockedCells) options.denySelectUnlockedCells = true;
  if (password) options.password = password;
  return Object.keys(options).length > 0
    ? (options as SheetProtectionOptions)
    : undefined;
}

/** Strictly normalize a stored sheet password verifier. */
export function normalizeSheetProtectionPassword(
  value: unknown,
): SheetProtectionPassword | undefined {
  if (!value || typeof value !== "object" || Array.isArray(value)) return undefined;
  const candidate = value as Record<string, unknown>;
  if (candidate.kind === "legacy") {
    const verifier = normalizeLegacyWorkbookPasswordVerifier(candidate.verifier);
    return verifier ? { kind: "legacy", verifier } : undefined;
  }
  if (candidate.kind === "hash") {
    const hash = normalizeProtectionPasswordHash(candidate);
    return hash ? { kind: "hash", ...hash } : undefined;
  }
  return undefined;
}

/** Outcome of checking a typed password against a stored verifier. */
export type ProtectionPasswordVerification =
  | "verified"
  | "invalid-password"
  | "unsupported-algorithm"
  | "unsupported-encoding"
  | "unsupported-length";

/** Check a typed password against either verifier kind. */
export async function verifyProtectionPassword(
  password: string,
  stored: SheetProtectionPassword,
): Promise<ProtectionPasswordVerification> {
  if (stored.kind === "legacy") {
    return verifyLegacyWorkbookPasswordResult(password, stored.verifier).status;
  }
  return checkProtectionPasswordHash(password, stored);
}

type ProtectionSheet = {
  protected?: boolean;
  protectionOptions?: SheetProtectionOptions;
};

/** True when `permission` is open on the sheet: always when unprotected. */
export function sheetProtectionAllows(
  sheet: ProtectionSheet,
  permission: SheetProtectionPermission,
): boolean {
  if (!sheet.protected) return true;
  return sheetProtectionPermissions(sheet.protectionOptions)[permission];
}

/** True when formats may be written on the sheet: unprotected, or protected
 *  with the Format cells permission. */
export function canFormatCells(sheet: ProtectionSheet): boolean {
  return sheetProtectionAllows(sheet, "formatCells");
}

/** Insert whole rows/columns: the permission alone (no lock condition). */
export function canInsertLines(sheet: ProtectionSheet, axis: "row" | "col"): boolean {
  return sheetProtectionAllows(sheet, axis === "row" ? "insertRows" : "insertColumns");
}

/** Width/height/AutoFit/hide/unhide of whole rows or columns. */
export function canFormatLines(sheet: ProtectionSheet, axis: "row" | "col"): boolean {
  return sheetProtectionAllows(sheet, axis === "row" ? "formatRows" : "formatColumns");
}

/**
 * Excel's Hidden flag in effect: the sheet is protected and the cell is
 * Hidden, so its contents stay out of the formula bar and edit box (formulas
 * and constants alike) and a copy carries its value, not its formula.
 */
export function isCellContentHidden(
  sheet: { protected?: boolean },
  cell: ProtectableCell | undefined,
): boolean {
  return sheet.protected === true && cell?.formulaHidden === true;
}

/**
 * Preserve-only receipt for a workbookProtection node that is outside the
 * passwordless / legacy-verifier subset owned by the editor. The value is
 * deliberately stable across browser and API projections so an unrelated
 * workbook edit can prove that the native authority was not changed.
 */
export const WORKBOOK_STRUCTURE_PROTECTION_NATIVE_BLOCKER =
  "workbook_structure_protection_native_authority_unsupported";

export type WorkbookStructureProtectionPasswordKind =
  | "none"
  | "legacy"
  | "modern"
  | "unsupported";

/** Top-level workbook structure protection; independent from sheet locks. */
export interface WorkbookStructureProtection {
  /** Protect Workbook > Structure: no sheet may be added, deleted, renamed,
   *  moved, hidden or unhidden. */
  locked: boolean;
  /**
   * Protect Workbook > Windows, as an imported file has it (Excel 2007/2010
   * and Excel for Mac offer it). The same password guards it, Unprotect
   * Workbook clears it with the structure lock, and the browser has no
   * workbook windows to hold in place, so it is carried, never set here.
   */
  windowsLocked?: true;
  passwordKind: WorkbookStructureProtectionPasswordKind;
  /** Normalized uppercase four-hex legacy verifier; never plaintext. */
  legacyVerifier?: string;
  /** Excel's ISO hash (workbookAlgorithmName/HashValue/SaltValue/SpinCount). */
  passwordHash?: ProtectionPasswordHash;
  /** Present when the native node must be preserved and cannot be changed. */
  nativeBlocker?: string;
}

/** Normalize persisted/imported workbook protection without inventing authority. */
export function normalizeWorkbookStructureProtection(
  value: unknown,
): WorkbookStructureProtection | undefined {
  if (!value || typeof value !== "object" || Array.isArray(value)) return undefined;
  const candidate = value as Record<string, unknown>;
  const locked = candidate.locked;
  const passwordKind = candidate.passwordKind;
  if (
    typeof locked !== "boolean" ||
    !["none", "legacy", "modern", "unsupported"].includes(String(passwordKind)) ||
    (candidate.windowsLocked !== undefined && candidate.windowsLocked !== true)
  ) {
    return undefined;
  }
  const windowsLocked = candidate.windowsLocked === true;
  if (windowsLocked && passwordKind === "unsupported") return undefined;
  const protectedAtAll = locked || windowsLocked;
  const legacyVerifier = normalizeLegacyWorkbookPasswordVerifier(
    candidate.legacyVerifier,
  );
  const passwordHash =
    candidate.passwordHash === undefined
      ? undefined
      : normalizeProtectionPasswordHash(candidate.passwordHash);
  if (candidate.passwordHash !== undefined && !passwordHash) return undefined;
  const nativeBlocker =
    typeof candidate.nativeBlocker === "string" &&
    candidate.nativeBlocker.length > 0 &&
    candidate.nativeBlocker.length <= 200
      ? candidate.nativeBlocker
      : undefined;
  if (passwordKind === "legacy" && !legacyVerifier) return undefined;
  if (passwordKind !== "legacy" && legacyVerifier) return undefined;
  if (passwordKind === "legacy" && !protectedAtAll) return undefined;
  if (passwordKind === "modern" && (!passwordHash || !protectedAtAll)) return undefined;
  if (passwordKind !== "modern" && passwordHash) return undefined;
  if (passwordKind === "unsupported" && !nativeBlocker) return undefined;
  if (passwordKind !== "unsupported" && nativeBlocker) return undefined;
  return {
    locked,
    ...(windowsLocked ? { windowsLocked: true as const } : {}),
    passwordKind: passwordKind as WorkbookStructureProtectionPasswordKind,
    ...(legacyVerifier ? { legacyVerifier } : {}),
    ...(passwordHash ? { passwordHash } : {}),
    ...(nativeBlocker ? { nativeBlocker } : {}),
  };
}

/**
 * Whether Protect Workbook is on in Excel's sense: the structure, the
 * windows, or both. Review > Protect Workbook then reads as pressed and asks
 * for Unprotect Workbook; only `locked` (the structure) gates sheet commands.
 */
export function workbookProtectionActive(
  protection: WorkbookStructureProtection | null | undefined,
): boolean {
  return protection?.locked === true || protection?.windowsLocked === true;
}

/** The verifier Unprotect Workbook checks, or undefined when none is set. */
export function workbookStructureProtectionPassword(
  protection: WorkbookStructureProtection | null | undefined,
): SheetProtectionPassword | undefined {
  if (protection?.passwordKind === "legacy" && protection.legacyVerifier) {
    return { kind: "legacy", verifier: protection.legacyVerifier };
  }
  if (protection?.passwordKind === "modern" && protection.passwordHash) {
    return { kind: "hash", ...protection.passwordHash };
  }
  return undefined;
}

/** Strictly normalize Excel's legacy workbookPassword verifier. */
export function normalizeLegacyWorkbookPasswordVerifier(
  value: unknown,
): string | undefined {
  if (typeof value !== "string") return undefined;
  const normalized = value.trim().toUpperCase();
  if (!/^[0-9A-F]{1,4}$/.test(normalized)) return undefined;
  return normalized.padStart(4, "0");
}

/** The byte-length field in MS-OFFCRYPTO Method1 is one byte. */
export const MAX_LEGACY_WORKBOOK_PASSWORD_BYTES = 255;
export type LegacyWorkbookPasswordVerification = Readonly<{
  status: "verified" | "invalid-password" | "unsupported-encoding" | "unsupported-length";
}>;

function legacyWorkbookPasswordInput(password: string):
  { value: string } | { status: "unsupported-encoding" | "unsupported-length" } {
  // SpreadsheetML starts from UTF-16LE and removes a leading BOM.
  const value = password.startsWith("\uFEFF") ? password.slice(1) : password;
  if (value.length > MAX_LEGACY_WORKBOOK_PASSWORD_BYTES) return { status: "unsupported-length" };
  // NUL terminates the native C-style input and cannot be an embedded byte.
  if (value.includes("\0")) return { status: "unsupported-encoding" };
  return { value };
}

function hashLegacyPasswordBytes(bytes: Uint8Array): string {
  // ECMA-376 Part4 §15.3.3, GetPasswordHash: unlike the Word-specific high-
  // word preprocessing, this loop consumes the complete byte string.
  if (bytes.length === 0) return "0000";
  let verifier = 0;
  for (let index = bytes.length - 1; index >= 0; index--) {
    verifier = ((verifier >>> 14) & 1) | ((verifier << 1) & 0x7fff);
    verifier ^= bytes[index]!;
  }
  verifier = ((verifier >>> 14) & 1) | ((verifier << 1) & 0x7fff);
  verifier ^= bytes.length ^ 0xce4b;
  return (verifier & 0xffff).toString(16).toUpperCase().padStart(4, "0");
}

/** Native legacy workbook verifier for an ASCII password (the only bytes
 * every system agrees on). Never used to manufacture protection: a new
 * password gets Excel's ISO hash. Not encryption. */
export function hashLegacyWorkbookPassword(password: string): string | undefined {
  const input = legacyWorkbookPasswordInput(password);
  if (!("value" in input) || /[^\x01-\x7F]/u.test(input.value)) return undefined;
  const [ascii] = legacyPasswordByteCandidates(input.value);
  return ascii ? hashLegacyPasswordBytes(ascii.bytes) : undefined;
}

/**
 * Check a typed password against a legacy 16-bit verifier, distinguishing
 * a password no Windows code page can spell (or one over the native
 * 255-byte length field) from a wrong one. Excel hashes the password's bytes
 * in its system's ANSI code page (MS-XLS 2.2.9); the page tries every code
 * page that can spell it (legacy-password-code-pages.ts).
 */
export function verifyLegacyWorkbookPasswordResult(
  password: string,
  verifier: unknown,
): LegacyWorkbookPasswordVerification {
  const normalized = normalizeLegacyWorkbookPasswordVerifier(verifier);
  if (normalized === undefined) return { status: "invalid-password" };
  const input = legacyWorkbookPasswordInput(password);
  if (!("value" in input)) return input;
  const candidates = legacyPasswordByteCandidates(input.value);
  if (candidates.length === 0) return { status: "unsupported-encoding" };
  const fitting = candidates.filter(
    (candidate) => candidate.bytes.length <= MAX_LEGACY_WORKBOOK_PASSWORD_BYTES,
  );
  if (fitting.length === 0) return { status: "unsupported-length" };
  return {
    status: fitting.some((candidate) => hashLegacyPasswordBytes(candidate.bytes) === normalized)
      ? "verified"
      : "invalid-password",
  };
}

/** Boolean adapter for callers that only need the verification verdict. */
export function verifyLegacyWorkbookPassword(password: string, verifier: unknown): boolean {
  return verifyLegacyWorkbookPasswordResult(password, verifier).status === "verified";
}

/** The minimal sheet shape the protection predicate reads. */
export interface ProtectableSheet {
  id?: string;
  protected?: boolean;
  protectionOptions?: SheetProtectionOptions;
  cells: Record<string, ProtectableCell>;
  /** Review ▸ Allow Edit Ranges (XLS-A028). */
  protectedRanges?: ProtectedEditRange[];
  /** A row's / column's default format, with its default protection (B501). */
  rowStyles?: Record<string, { locked?: unknown } | undefined>;
  columnStyles?: Record<string, { locked?: unknown } | undefined>;
}

/**
 * Whether the cell at (row, col) is locked: its own flag, else its row's
 * default, else its column's (a blank cell of an unlocked XLSX column is
 * writable in a protected template; B501), else locked.
 */
export function isCellLockedIn(sheet: ProtectableSheet, row: number, col: number): boolean {
  const cell = sheet.cells[`${row}:${col}`];
  if (cell?.locked !== undefined) return cell.locked !== false;
  const rowLock = sheet.rowStyles?.[String(row)]?.locked;
  if (rowLock !== undefined) return rowLock !== false;
  const columnStyles = sheet.columnStyles as LineStyleMap | undefined;
  const columnStyle = (columnStyles?.[String(col)] ?? expandLineStyleRuns(columnStyles)?.[String(col)]) as { locked?: unknown } | undefined;
  return columnStyle?.locked !== false;
}

/** A rectangular selection (inclusive bounds, any corner order). */
export interface ProtectionRange {
  r1: number;
  r2: number;
  c1: number;
  c2: number;
}

/**
 * True when the cell at (row, col) is LOCKED in Excel's sense: a cell is locked
 * by default (absent `locked` ⇒ locked); only an explicit `locked === false`
 * unlocks it. This is independent of whether the sheet is protected — it just
 * reports the cell's own flag.
 */
export function isCellLocked(cell: ProtectableCell | undefined): boolean {
  return cell?.locked !== false;
}

/**
 * True when the cell at (row, col) may be edited right now. Always true on an
 * unprotected sheet; on a protected sheet, true only for an explicitly UNLOCKED
 * cell. This is the single predicate every enforcement point funnels through.
 */
export function isCellEditable(
  sheet: ProtectableSheet,
  row: number,
  col: number,
  /** Password edit ranges unlocked this session (XLS-A028). */
  unlockedRanges?: ReadonlySet<string>,
): boolean {
  if (!sheet.protected) return true;
  if (!isCellLockedIn(sheet, row, col)) return true;
  return editRangeOpensCell(sheet, row, col, unlockedRanges);
}

/** Normalize a range so r1<=r2 and c1<=c2 regardless of the dragged corner. */
function normalizeRange(range: ProtectionRange): ProtectionRange {
  return {
    r1: Math.min(range.r1, range.r2),
    r2: Math.max(range.r1, range.r2),
    c1: Math.min(range.c1, range.c2),
    c2: Math.max(range.c1, range.c2),
  };
}

/**
 * Set the `locked` flag across a rectangular range. Returns a FRESH cells map
 * (never mutates the input). To keep serialization minimal:
 *  - `locked === true` (the default) is stored as the ABSENCE of the key, so a
 *    just-locked cell that carries no other data is pruned entirely.
 *  - `locked === false` is stored explicitly (it is the only meaningful state).
 * A cell that would otherwise be empty (no value, no other metadata) is dropped
 * when it ends up locked-by-default, mirroring updateCellValue's blank handling.
 */
export function setRangeLocked<C extends ProtectableCell>(
  cells: Record<string, C>,
  range: ProtectionRange,
  locked: boolean,
): Record<string, C> {
  const { r1, r2, c1, c2 } = normalizeRange(range);
  const next: Record<string, C> = { ...cells };
  for (let r = r1; r <= r2; r += 1) {
    for (let c = c1; c <= c2; c += 1) {
      const key = `${r}:${c}`;
      const existing = next[key];
      if (locked) {
        // Locked is the default — drop the explicit flag (and the whole cell if
        // it carried nothing else) so the serialized shape stays minimal.
        if (!existing || existing.locked === undefined) continue;
        const rest = { ...existing };
        delete rest.locked;
        if (cellIsEmpty(rest)) delete next[key];
        else next[key] = rest;
      } else {
        next[key] = { ...(existing ?? ({} as C)), value: existing?.value ?? "", locked: false };
      }
    }
  }
  return next;
}

/**
 * Set Format Cells > Protection > Hidden across a range, returning a FRESH
 * cells map. Like `setRangeLocked`, only the non-default state is stored
 * (`formulaHidden: true`); clearing it drops the key and prunes a cell that
 * carried nothing else.
 */
export function setRangeFormulaHidden<C extends ProtectableCell>(
  cells: Record<string, C>,
  range: ProtectionRange,
  hidden: boolean,
): Record<string, C> {
  const { r1, r2, c1, c2 } = normalizeRange(range);
  const next: Record<string, C> = { ...cells };
  for (let r = r1; r <= r2; r += 1) {
    for (let c = c1; c <= c2; c += 1) {
      const key = `${r}:${c}`;
      const existing = next[key];
      if (!hidden) {
        if (!existing || existing.formulaHidden === undefined) continue;
        const rest = { ...existing };
        delete rest.formulaHidden;
        if (cellIsEmpty(rest)) delete next[key];
        else next[key] = rest;
      } else {
        next[key] = {
          ...(existing ?? ({} as C)),
          value: existing?.value ?? "",
          formulaHidden: true,
        };
      }
    }
  }
  return next;
}

/**
 * True when a cell carries no content or metadata worth persisting — only the
 * keys a fresh blank cell would have (value/kind, both empty). Used so unlocking
 * then re-locking a blank cell prunes it instead of leaving a husk behind.
 */
function cellIsEmpty(cell: Record<string, unknown>): boolean {
  for (const key of Object.keys(cell)) {
    if (key === "kind") continue;
    if (key === "value") {
      const value = cell.value;
      if (value !== undefined && value !== "") return false;
      continue;
    }
    return false;
  }
  return true;
}

/**
 * True when ANY cell in the range is locked under active protection — i.e. a
 * mutation touching this range should be blocked. Always false on an unprotected
 * sheet. Used by structural / multi-cell guards (clear, paste, delete line).
 */
export function rangeHasLockedCell(
  sheet: ProtectableSheet,
  range: ProtectionRange,
  unlockedRanges?: ReadonlySet<string>,
): boolean {
  if (!sheet.protected) return false;
  const { r1, r2, c1, c2 } = normalizeRange(range);
  // A position with no cell entry is locked by default, so a range larger
  // than the sheet's whole cell map always holds one (whole-line checks),
  // unless an edit range may be what opens it.
  const area = (r2 - r1 + 1) * (c2 - c1 + 1);
  if (area > 4_096 && area > Object.keys(sheet.cells).length && !sheet.protectedRanges?.length) {
    return true;
  }
  for (let r = r1; r <= r2; r += 1) {
    for (let c = c1; c <= c2; c += 1) {
      if (!isCellEditable(sheet, r, c, unlockedRanges)) return true;
    }
  }
  return false;
}

/**
 * Whether Excel's protected-sheet sort contract permits rewriting `range`.
 * An unprotected sheet is unrestricted. A protected sheet needs the native
 * Allow Sort permission AND every destination/source cell the sort rewrites
 * must be explicitly unlocked; permission never bypasses the cell lock gate.
 */
export function canSortRange(
  sheet: ProtectableSheet,
  range: ProtectionRange,
): boolean {
  if (!sheet.protected) return true;
  if (sheet.protectionOptions?.allowSort !== true) return false;
  return !rangeHasLockedCell(sheet, range);
}

/**
 * Where Tab (or Shift+Tab) goes on a protected sheet: the next cell that can
 * be filled in, in reading order -- along the row, then on to the next row
 * (the previous row for Shift+Tab), wrapping past the sheet's end -- so a
 * form's input cells are visited in turn however the locked cells lie
 * between them (Microsoft Support, "Lock or unlock specific areas of a
 * protected worksheet": "you can press TAB to move between the unlocked
 * cells"; live audit S-110). Null when no other cell can be filled in, or
 * when the sheet is too large to look through within `limit` probes.
 */
export function nextEditableCellInTabOrder(
  from: { row: number; col: number },
  reverse: boolean,
  rowCount: number,
  columnCount: number,
  isEditable: (row: number, col: number) => boolean,
  limit = 500_000,
): { row: number; col: number } | null {
  if (rowCount <= 0 || columnCount <= 0) return null;
  const total = rowCount * columnCount;
  if (total > limit) return null;
  const start = from.row * columnCount + from.col;
  const step = reverse ? -1 : 1;
  for (let offset = 1; offset < total; offset += 1) {
    const index = (((start + step * offset) % total) + total) % total;
    const row = Math.floor(index / columnCount);
    const col = index % columnCount;
    if (isEditable(row, col)) return { row, col };
  }
  return null;
}

/**
 * Whether the selection may land on (row, col): Excel's Select locked cells /
 * Select unlocked cells. Always true on an unprotected sheet.
 */
export function canSelectCell(
  sheet: ProtectableSheet,
  row: number,
  col: number,
): boolean {
  if (!sheet.protected) return true;
  const permissions = sheetProtectionPermissions(sheet.protectionOptions);
  return isCellLockedIn(sheet, row, col)
    ? permissions.selectLockedCells
    : permissions.selectUnlockedCells;
}

/**
 * Delete whole rows/columns: the permission AND every cell of the deleted
 * lines unlocked (VBA Worksheet.Protect AllowDeletingRows/Columns). `extent`
 * is the size of the other axis (columns for rows, rows for columns).
 */
export function canDeleteLines(
  sheet: ProtectableSheet,
  axis: "row" | "col",
  start: number,
  end: number,
  extent: number,
): boolean {
  if (!sheet.protected) return true;
  if (!sheetProtectionAllows(sheet, axis === "row" ? "deleteRows" : "deleteColumns")) {
    return false;
  }
  const last = Math.max(0, extent - 1);
  return !rangeHasLockedCell(
    sheet,
    axis === "row"
      ? { r1: start, r2: end, c1: 0, c2: last }
      : { r1: 0, r2: last, c1: start, c2: end },
  );
}
