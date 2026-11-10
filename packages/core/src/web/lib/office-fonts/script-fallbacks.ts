/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
// The faces a run's characters fall back to, by script (script-fallbacks.json):
// one family behind the run's face and its open stand-ins -- "MRAI Script
// Fallback" (or its serif / mono cut), drawn by a served face of each script
// for the characters that face covers -- so a character the stand-in lacks is
// drawn by a face of its own script on every machine, the print page's
// included, rather than by whatever the reader's system picks.
import table from "./script-fallbacks.json";
import { ensureOfficeAliasFamily, officeFontSubstitutes } from "./substitutes";

export type ScriptFallbackClass = "sans" | "serif" | "mono";

interface ScriptRow {
  readonly script: string;
  readonly sans: readonly string[];
  readonly serif?: readonly string[];
  readonly mono?: readonly string[];
}

const rows = table.scripts as readonly ScriptRow[];
const classes = table.classes as Readonly<Record<"serif" | "mono", readonly string[]>>;
const GENERIC = new Set(["serif", "sans-serif", "monospace", "cursive", "fantasy", "system-ui", "ui-sans-serif", "ui-serif", "ui-monospace", "math", "emoji"]);

/** The family each class falls back to. */
export const SCRIPT_FALLBACK_FAMILY: Readonly<Record<ScriptFallbackClass, string>> = {
  sans: "MRAI Script Fallback",
  serif: "MRAI Script Fallback Serif",
  mono: "MRAI Script Fallback Mono",
};

const unquote = (name: string) => name.trim().replace(/^["']|["']$/g, "").trim();

/** Serif, sans or mono, by the stack's first face and its stand-ins' names. */
export function scriptFallbackClass(stack: string): ScriptFallbackClass {
  const first = unquote(stack.split(",")[0] ?? "");
  const names = [first, ...officeFontSubstitutes(first)].map((name) => name.toLowerCase());
  if (names.some((name) => classes.mono.some((word) => name.includes(word)))) return "mono";
  
  if (names.some((name) => !name.includes("sans") && !/\bsan[\s-]?serif/.test(name) && classes.serif.some((word) => name.includes(word)))) return "serif";
  return "sans";
}

const rowFaces = (row: ScriptRow, kind: ScriptFallbackClass) => (kind === "sans" ? row.sans : row[kind] ?? row.sans);

/** The faces drawing one class's fallback family, in priority order, less
 * the rows (scripts) named in `without`. */
export function scriptFallbackFaces(kind: ScriptFallbackClass, without: ReadonlySet<string> = new Set()): readonly string[] {
  const out: string[] = [];
  for (const row of rows) {
    if (without.has(row.script)) continue;
    for (const face of rowFaces(row, kind)) {
      if (!out.includes(face)) out.push(face);
    }
  }
  return out;
}

/**
 * `stack` with its class's fallback family behind its faces and ahead of any
 * generic family (not repeated). The family's faces are registered on first
 * use; only the unicode-range slices a page draws are fetched.
 */
export function withScriptFallbacks(stack: string): string {
  const parts = stack.split(",").map((part) => part.trim()).filter(Boolean);
  if (parts.some((part) => unquote(part).startsWith("MRAI Script Fallback"))) return stack;
  const kind = scriptFallbackClass(stack);
  const named = parts.filter((part) => !GENERIC.has(unquote(part).toLowerCase()));
  const generic = parts.filter((part) => GENERIC.has(unquote(part).toLowerCase()));
  const covered = slicedRowsDrawnBy(named, kind);
  const family = covered.length ? `${SCRIPT_FALLBACK_FAMILY[kind]} without ${covered.join(" ")}` : SCRIPT_FALLBACK_FAMILY[kind];
  ensureOfficeAliasFamily(family, scriptFallbackFaces(kind, new Set(covered)));
  return [...named, `"${family}"`, ...generic].join(", ");
}

/**
 * The rows whose faces are cut into hundreds of unicode-range slices (Noto
 * Sans SC and KR: about 450 at regular and bold).
 */
const SLICED_ROWS: ReadonlySet<string> = new Set(["east-asian"]);

/**
 * The sliced rows a stack's own families already draw: every face of the
 * row stands behind one of them (the spreadsheet's "Microsoft YaHei" and
 * "Malgun Gothic" are Noto Sans SC and Noto Sans KR, in the row's order).
 * Their characters are drawn by those same files before the fallback is
 * reached, so the stack's fallback family leaves the row out. The browser
 * resolves every face of each family it reaches, once for each new size,
 * weight and style, whether or not a slice is ever fetched: the row cost
 * 0.2-0.4 s of a cold Japanese workbook's first layouts (R116-X02).
 */
function slicedRowsDrawnBy(named: readonly string[], kind: ScriptFallbackClass): string[] {
  const drawn = new Set<string>();
  for (const part of named) {
    const face = unquote(part);
    drawn.add(face.toLowerCase());
    for (const standIn of officeFontSubstitutes(face)) drawn.add(standIn.toLowerCase());
  }
  return rows
    .filter((row) => SLICED_ROWS.has(row.script) && rowFaces(row, kind).every((face) => drawn.has(face.toLowerCase())))
    .map((row) => row.script);
}
