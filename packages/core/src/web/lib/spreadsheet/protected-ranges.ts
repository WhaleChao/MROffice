/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * XLS-A028 — Review ▸ Protect ▸ Allow Edit Ranges.
 *
 * Microsoft Support, "Lock or unlock specific areas of a protected
 * worksheet": ranges a protected sheet still lets people edit. Each has a
 * Title, the cells it Refers to and an optional Range password. A range
 * without a password is editable while the sheet is protected; a range with
 * one asks for it ("Unlock Range") the first time a cell of it is changed,
 * and stays unlocked for that session. Ranges are defined while the sheet is
 * unprotected (Excel greys the command otherwise). SpreadsheetML keeps them as
 * `<protectedRanges><protectedRange name sqref algorithmName hashValue
 * saltValue spinCount/>` after `<sheetProtection>`.
 *
 * Windows user/group permissions (the Permissions… button) need a Windows
 * domain account system, which a browser product has no access to.
 */

import {
  normalizeProtectionPasswordHash,
  type ProtectionPasswordHash,
} from "./protection-password";
import { spreadsheetNameFold } from "./name-fold";

export interface ProtectedEditRange {
  /** Excel's Title (protectedRange/@name): letters, digits, `_`, no spaces. */
  title: string;
  /** Excel's Refers to cells, as an sqref: `A1:B5 D3` (no `$`, no sheet). */
  ref: string;
  passwordHash?: ProtectionPasswordHash;
}

export const MAX_PROTECTED_EDIT_RANGES = 256;
const MAX_COLUMNS = 16_384;
const MAX_ROWS = 1_048_576;

interface Rect {
  r1: number;
  c1: number;
  r2: number;
  c2: number;
}

function columnIndex(letters: string): number {
  let index = 0;
  for (const letter of letters) index = index * 26 + (letter.charCodeAt(0) - 64);
  return index - 1;
}

function columnLetters(index: number): string {
  let out = "";
  let value = index + 1;
  while (value > 0) {
    const remainder = (value - 1) % 26;
    out = String.fromCharCode(65 + remainder) + out;
    value = Math.floor((value - 1) / 26);
  }
  return out;
}

function parseCell(text: string): { row: number; col: number } | null {
  const match = /^\$?([A-Za-z]{1,3})\$?([1-9]\d{0,6})$/.exec(text);
  if (!match) return null;
  const col = columnIndex(match[1]!.toUpperCase());
  const row = Number(match[2]) - 1;
  if (col >= MAX_COLUMNS || row >= MAX_ROWS) return null;
  return { row, col };
}

/** The rectangles of an sqref, or null when any part is not a cell range. */
export function protectedRangeRects(ref: string): Rect[] | null {
  const parts = ref.trim().split(/[\s,]+/).filter(Boolean);
  if (parts.length === 0) return null;
  const rects: Rect[] = [];
  for (const part of parts) {
    const [first, second, extra] = part.split(":");
    if (extra !== undefined) return null;
    const a = parseCell(first ?? "");
    const b = second === undefined ? a : parseCell(second);
    if (!a || !b) return null;
    rects.push({
      r1: Math.min(a.row, b.row),
      r2: Math.max(a.row, b.row),
      c1: Math.min(a.col, b.col),
      c2: Math.max(a.col, b.col),
    });
  }
  return rects;
}

/**
 * Refers to cells as typed: `=$A$1:$B$5,Sheet1!D3` and the like (the sheet
 * name, if any, must be this sheet's) → the stored sqref, or null.
 */
export function normalizeProtectedRangeRef(input: string, sheetName: string): string | null {
  const body = input.trim().replace(/^=/, "");
  if (!body) return null;
  const parts = body.split(",").map((part) => part.trim()).filter(Boolean);
  const out: string[] = [];
  for (const part of parts) {
    const bang = part.lastIndexOf("!");
    let address = part;
    if (bang >= 0) {
      const named = part.slice(0, bang).replace(/^'(.*)'$/, "$1").replace(/''/g, "'");
      if (named.toLocaleLowerCase() !== sheetName.toLocaleLowerCase()) return null;
      address = part.slice(bang + 1);
    }
    const rects = protectedRangeRects(address);
    if (!rects || rects.length !== 1) return null;
    const rect = rects[0]!;
    const start = `${columnLetters(rect.c1)}${rect.r1 + 1}`;
    out.push(
      rect.r1 === rect.r2 && rect.c1 === rect.c2 ? start : `${start}:${columnLetters(rect.c2)}${rect.r2 + 1}`,
    );
  }
  return out.length ? out.join(" ") : null;
}

/** Excel's Refers to cells text for a stored sqref: `=$A$1:$B$5,$D$3`. */
export function protectedRangeDisplayRef(ref: string): string {
  return `=${ref
    .split(/\s+/)
    .filter(Boolean)
    .map((part) =>
      part
        .split(":")
        .map((cell) => cell.replace(/^([A-Z]+)(\d+)$/, "$$$1$$$2"))
        .join(":"),
    )
    .join(",")}`;
}

/** Excel's Title rule: a name-like word (no spaces). */
export function validProtectedRangeTitle(title: string): boolean {
  return /^[\p{L}_][\p{L}\p{N}_.]{0,254}$/u.test(title);
}

/** Fail-closed reader of a sheet's `protectedRanges`. */
export function normalizeProtectedEditRanges(raw: unknown): ProtectedEditRange[] {
  if (!Array.isArray(raw)) return [];
  const out: ProtectedEditRange[] = [];
  const titles = new Set<string>();
  for (const entry of raw) {
    if (out.length >= MAX_PROTECTED_EDIT_RANGES) break;
    if (!entry || typeof entry !== "object" || Array.isArray(entry)) continue;
    const candidate = entry as Record<string, unknown>;
    const title = typeof candidate.title === "string" ? candidate.title : "";
    const ref = typeof candidate.ref === "string" ? candidate.ref.trim() : "";
    if (!validProtectedRangeTitle(title) || titles.has(spreadsheetNameFold(title)) || !protectedRangeRects(ref)) {
      continue;
    }
    let passwordHash: ProtectionPasswordHash | undefined;
    if (candidate.passwordHash !== undefined) {
      const normalized = normalizeProtectionPasswordHash(candidate.passwordHash);
      if (!normalized) continue;
      passwordHash = normalized;
    }
    titles.add(spreadsheetNameFold(title));
    out.push({ title, ref, ...(passwordHash ? { passwordHash } : {}) });
  }
  return out;
}

/** The session key of an unlocked password range. */
export function protectedRangeUnlockKey(sheetId: string, range: ProtectedEditRange): string {
  return `${sheetId}\u0000${spreadsheetNameFold(range.title)}`;
}

/** The edit ranges covering (row, col), in their order. */
export function protectedRangesAt(
  ranges: readonly ProtectedEditRange[] | undefined,
  row: number,
  col: number,
): ProtectedEditRange[] {
  if (!ranges?.length) return [];
  return ranges.filter((range) =>
    (protectedRangeRects(range.ref) ?? []).some(
      (rect) => row >= rect.r1 && row <= rect.r2 && col >= rect.c1 && col <= rect.c2,
    ),
  );
}

/**
 * A protected sheet lets (row, col) be edited through an edit range: one
 * without a password, or one already unlocked this session.
 */
export function editRangeOpensCell(
  sheet: { id?: string; protectedRanges?: readonly ProtectedEditRange[] },
  row: number,
  col: number,
  unlocked?: ReadonlySet<string>,
): boolean {
  return protectedRangesAt(sheet.protectedRanges, row, col).some(
    (range) => !range.passwordHash || (unlocked?.has(protectedRangeUnlockKey(sheet.id ?? "", range)) ?? false),
  );
}
