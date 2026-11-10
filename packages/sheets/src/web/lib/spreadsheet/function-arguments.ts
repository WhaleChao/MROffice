/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
// ---------------------------------------------------------------------------
// Spreadsheet "Function Arguments" wizard (Excel/WPS parity) — PURE logic.
//
// No React, no DOM. Powers the wizard the fx Insert Function dialog hands off
// to once a function is picked: it turns a FunctionSignature into one labelled
// input per argument, assembles the typed values back into a "=NAME(a,b,…)"
// formula, and produces the partial formula string to evaluate for the live
// result preview.
//
// The signature is the single source of truth (lib/spreadsheet/formula-
// intellisense → signatureFor), so the field set can never drift from the
// engine's actual arguments. The engine itself does the live evaluation; this
// module only assembles the strings to hand it.
// ---------------------------------------------------------------------------

import type { FunctionSignature } from "./formula-intellisense";
import { Parser, tokenize, type Node } from "@mrnote/core/web/lib/spreadsheet/formula-engine.ts";

/**
 * How many EXTRA repeat slots a variadic function shows beyond the named ones.
 * Excel's dialog grows the argument list as you fill the last slot; we render a
 * fixed, bounded set up front (the named args + this many repeats) which is
 * enough for the overwhelmingly common 2–4 operand call, keeps the dialog a
 * fixed size, and is itself extended on demand by `growRepeatFields`.
 */
export const DEFAULT_VARIADIC_REPEAT_SLOTS = 2;

/**
 * One labelled input the wizard renders.
 *
 *  - `label`   the human argument name shown beside the input, already stripped
 *              of the signature's "[optional]" brackets and never the bare "…"
 *              repeat token (a repeat slot is labelled from the arg it repeats).
 *  - `optional` true for "[bracketed]" args AND for every generated repeat slot
 *              (a variadic call is valid with just its required operands).
 *  - `repeating` true for the slots generated from a variadic "…" token, so the
 *              UI can mark them and offer to add more.
 *  - `placeholder` the raw signature token (with brackets / "…") for the input's
 *              placeholder, so the user still sees the original Excel arg syntax.
 */
export interface ArgField {
  readonly label: string;
  readonly optional: boolean;
  readonly repeating: boolean;
  readonly placeholder: string;
  /**
   * For a repeat slot: how many arguments repeat together. SUM repeats one
   * (number3, number4, …); SUMIFS repeats a pair (criteria_range3, criteria3,
   * criteria_range4, …), so the next label is read from the slot this many
   * places back.
   */
  readonly repeatGroupSize?: number;
}

/** A variadic signature's last arg label is the literal ellipsis token. */
const ELLIPSIS = "…";

/** Strip surrounding "[ ]" (optional marker) from a signature arg label. */
function unbracket(label: string): { name: string; optional: boolean } {
  if (label.startsWith("[") && label.endsWith("]") && label.length >= 2) {
    return { name: label.slice(1, -1), optional: true };
  }
  return { name: label, optional: false };
}

/** Split a label into its stem and trailing number ("criteria2" → ["criteria", 2]). */
function splitTrailingNumber(label: string): { stem: string; n: number | null } {
  const trailing = /^(.*?)(\d+)$/.exec(label);
  return trailing
    ? { stem: trailing[1], n: Number(trailing[2]) }
    : { stem: label, n: null };
}

/**
 * The group of arguments a variadic signature repeats: the run of optional
 * labels just before the "…" (SUM(number1, [number2], …) repeats [number2];
 * SUMIFS repeats [criteria_range2], [criteria2]). A signature with no optional
 * run repeats its last named label, and a "…"-only signature repeats "value".
 */
function repeatGroup(named: readonly string[]): string[] {
  const group: string[] = [];
  for (let i = named.length - 1; i >= 0; i -= 1) {
    const { name, optional } = unbracket(named[i]);
    if (!optional) break;
    group.unshift(name);
  }
  if (group.length > 0) return group;
  const last = named.length > 0 ? unbracket(named[named.length - 1]).name : "";
  return [last || "value"];
}

/** The label `steps` repetitions after `label` ("criteria2" + 1 → "criteria3"). */
function nextRepeatLabel(label: string, steps: number, fromOne: boolean): string {
  const { stem, n } = splitTrailingNumber(label);
  if (n !== null) return `${stem}${n + steps}`;
  // An unnumbered base ("argument") starts counting at 1 on its first repeat.
  return `${stem}${fromOne ? steps : steps + 1}`;
}

/**
 * Build the ordered list of input fields for a function's signature.
 *
 * Non-variadic: one field per arg, in order, optional iff "[bracketed]".
 *
 * Variadic (last arg === "…"): the named args before the ellipsis are emitted
 * as-is, then `repeatSlots` extra OPTIONAL repetitions of the repeating group
 * are appended, continuing its numbering (SUM → number1, number2, number3, …;
 * SUMIFS → …, criteria_range3, criteria3). A bare generic signature
 * (args ["argument","…"]) yields "argument1", "argument2", ….
 *
 * A nullary function (args []) yields no fields — the wizard inserts "=NAME()".
 */
export function argFieldsFor(
  sig: FunctionSignature,
  repeatSlots: number = DEFAULT_VARIADIC_REPEAT_SLOTS,
): ArgField[] {
  const fields: ArgField[] = [];
  const args = sig.args;
  if (!sig.variadic) {
    for (const raw of args) {
      const { name, optional } = unbracket(raw);
      fields.push({ label: name, optional, repeating: false, placeholder: raw });
    }
    return fields;
  }

  // Variadic: everything before the trailing ellipsis is a fixed named arg.
  // (The ellipsis is always the final token by the signature contract, but we
  // filter defensively so a stray "…" elsewhere never becomes a literal field.)
  const named = args.filter((a) => a !== ELLIPSIS);
  for (const raw of named) {
    const { name, optional } = unbracket(raw);
    fields.push({ label: name, optional, repeating: false, placeholder: raw });
  }

  // The repeat slots continue the numbering of the group the "…" repeats, a
  // whole group per slot: SUM → number3, number4; SUMIFS → criteria_range3,
  // criteria3, criteria_range4, criteria4. An unnumbered base ("argument")
  // counts from 1.
  const group = repeatGroup(named);
  const slots = Math.max(0, repeatSlots);
  for (let i = 0; i < slots; i += 1) {
    for (const base of group) {
      fields.push({
        label: nextRepeatLabel(base, i + 1, true),
        optional: true,
        repeating: true,
        placeholder: ELLIPSIS,
        repeatGroupSize: group.length,
      });
    }
  }
  return fields;
}

/**
 * Append one more repeat field to an existing field list for a variadic
 * signature (Excel grows the list as the last slot is filled). The new slot
 * continues the numbering of the last repeat field; on a non-variadic signature
 * (no repeat fields) the list is returned unchanged.
 */
export function growRepeatFields(fields: readonly ArgField[]): ArgField[] {
  const next = [...fields];
  const lastRepeat = [...fields].reverse().find((f) => f.repeating);
  if (!lastRepeat) return next;
  // One field at a time, continuing the group pattern: the new label is the
  // label one group back, advanced by one repetition.
  const groupSize = lastRepeat.repeatGroupSize ?? 1;
  const template = next[next.length - groupSize];
  const base = template?.repeating ? template.label : lastRepeat.label;
  const { stem, n } = splitTrailingNumber(base);
  next.push({
    label: n !== null ? `${stem}${n + 1}` : `${stem}${next.length + 1}`,
    optional: true,
    repeating: true,
    placeholder: ELLIPSIS,
    repeatGroupSize: groupSize,
  });
  return next;
}

/**
 * Trim trailing empty values so an unused optional / repeat slot never emits a
 * dangling comma. Interior empties are KEPT (a blank required middle arg is a
 * deliberate "omitted" placeholder, e.g. `=IF(A1,,0)`), only the run of empties
 * at the END is dropped. Values are compared after trimming surrounding spaces.
 */
function trimTrailingEmpty(values: readonly string[]): string[] {
  const out = [...values];
  while (out.length > 0 && out[out.length - 1].trim() === "") out.pop();
  return out;
}

/**
 * Assemble the final formula from a function name + the per-field raw values.
 *
 *   assembleFormula("SUM", ["A1:A3", "10", "", ""]) === "=SUM(A1:A3,10)"
 *   assembleFormula("PI", [])                        === "=PI()"
 *   assembleFormula("IF", ["A1>0", "", "0"])         === "=IF(A1>0,,0)"
 *
 * Trailing empties are dropped (so unused optional/repeat slots vanish); an
 * interior empty is preserved as an explicit omitted argument. Each value is
 * trimmed of surrounding whitespace. The name is upper-cased to match how the
 * engine and the fx list present it. The result is always a "=" formula.
 */
export function assembleFormula(
  name: string,
  values: readonly string[],
): string {
  const fn = name.trim().toUpperCase();
  const trimmed = values.map((v) => v.trim());
  const used = trimTrailingEmpty(trimmed);
  return `=${fn}(${used.join(",")})`;
}

/**
 * The formula string to feed the engine for the LIVE preview. Identical to the
 * assembled formula, but returns null when there is nothing meaningful to show
 * yet — i.e. a function that takes arguments but has none filled in — so the UI
 * can hide the preview row instead of flashing a spurious error while the user
 * has not typed anything. A nullary function (no fields at all) always previews
 * (e.g. =TODAY(), =PI()).
 *
 * @param name      the chosen function name.
 * @param values    the per-field raw values (same order as argFieldsFor).
 * @param hasFields whether the signature has any argument fields at all.
 */
export function previewFormulaFor(
  name: string,
  values: readonly string[],
  hasFields: boolean,
): string | null {
  if (hasFields && values.every((v) => v.trim() === "")) return null;
  return assembleFormula(name, values);
}

/**
 * Function Arguments puts quotes round an argument typed as plain words when
 * its box is left, as Excel's dialog does: IF's value_if_true typed as 通过
 * becomes "通过" instead of an unknown name that evaluates to #NAME?. Only an
 * entry that is no expression at all (words that do not parse) or a single
 * name the workbook does not define is quoted; a number, a logical, a
 * reference, a defined name, a call or an expression stays as typed
 * (audit P-FUNCTION-TEXT-QUOTE).
 */
export function autoQuotedFunctionArgument(
  value: string,
  isDefinedName: (name: string) => boolean,
): string {
  const text = value.trim();
  if (text === "" || text.startsWith('"') || text.startsWith("=")) return value;
  const quoted = `"${text.replace(/"/g, '""')}"`;
  const tokens = tokenize(text);
  let node: Node | null = null;
  if (tokens) {
    try {
      node = new Parser(tokens).parse();
    } catch {
      node = null;
    }
  }
  if (!node) {
    // Words, not a broken expression: no operator, bracket or separator.
    return /^[^=+\-*/^&<>(){}[\],;:!"'#$%@]+$/u.test(text) ? quoted : value;
  }
  // Words separated by spaces read as an intersection of names.
  const undefinedWords = (candidate: Node): boolean =>
    candidate.kind === "name"
      ? !isDefinedName(candidate.name)
      : candidate.kind === "isect" && undefinedWords(candidate.left) && undefinedWords(candidate.right);
  return undefinedWords(node) ? quoted : value;
}
