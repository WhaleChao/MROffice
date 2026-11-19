/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The rules (underline, strikethrough) of a scanned line: a line whose text
 * is an invisible OCR layer under the picture of the page (`3 Tr`, a
 * glyphless face). A reader who underlines such a line wants a line under
 * the scanned words; the page's own rule is written beside the invisible
 * run (lib/pdf/engine/native/invisible-text-rule.ts) and the scan is left
 * as it is -- taken into a text box for a rule, the line showed the OCR's
 * misreadings instead of the words (pdf-walk-1000, 0436: "PUfflH1").
 *
 * This file is what the editor and the window share about them, with no
 * planner in it: the rules standing in the page as stretches of the box's
 * text, the presses of a session over them, what is asked of the page, and
 * what the text box carries once the line is taken into it (a character
 * format pressed, words typed).
 */
import type { PdfNativeInvisibleTextRuleFormat } from "./engine/native/invisible-text-rule";
import {
  applyPdfTextPartFormat,
  pdfTextPartFormatsAfterEdit,
  type PdfTextPartFormat,
  type PdfTextPartFormats,
} from "./native-text-part-formats";

export type PdfNativeCoverRuleStyle = "underline" | "strikethrough";
export const PDF_NATIVE_COVER_RULE_STYLES: readonly PdfNativeCoverRuleStyle[] = ["underline", "strikethrough"];

/** Stretches of the box's text, in its characters. */
export type PdfNativeCoverRuleStretches = readonly Readonly<{ start: number; end: number }>[];

/**
 * What the window says of a scanned line's rules to its editor: `pending`
 * while the page is still read (a press is held meanwhile), `proven` with
 * the rules standing once every line of the box takes a rule beside its
 * invisible run. A line that does not is given none (undefined): its rules
 * are its text box's.
 */
export type PdfNativeCoverRules = Readonly<{
  state: "pending" | "proven";
  /** The box's text the stretches are of (the page's, as the runs read). */
  text: string;
  underline: PdfNativeCoverRuleStretches;
  strikethrough: PdfNativeCoverRuleStretches;
}>;

/** What the editor is given while the page is read. */
export const PDF_NATIVE_COVER_RULES_PENDING: PdfNativeCoverRules = Object.freeze({
  state: "pending", text: "", underline: Object.freeze([]), strikethrough: Object.freeze([]),
});

const trimmedEnd = (text: string) => text.replace(/[\s\u00a0]+$/u, "");

/**
 * Whether the stretches read on `ruleText` are stretches of the editor's
 * text: the same characters, a blank standing for any blank (the editor
 * shows a tab or a no-break space of the page as it types one) and blanks
 * at the end left out of both.
 */
export function pdfNativeCoverRulesTextFits(ruleText: string, editorText: string): boolean {
  const left = trimmedEnd(ruleText);
  const right = trimmedEnd(editorText);
  if (left.length !== right.length) return false;
  for (let at = 0; at < left.length; at += 1) {
    if (left[at] !== right[at] && (left[at].trim() || right[at].trim())) return false;
  }
  return true;
}

 
export type PdfNativeCoverRuleSegment = Readonly<{
  row: number; start: number; end: number; underline?: boolean; strikethrough?: boolean;
}>;

const RULE_KEYS: ReadonlySet<string> = new Set<string>(PDF_NATIVE_COVER_RULE_STYLES);
const SEGMENT_KEYS: ReadonlySet<string> = new Set(["row", "start", "end", "underline", "strikethrough"]);

/**
 * The format as a rules-only ask of the page: the whole box's rules and its
 * parts' (offsets in the box's text). Null when it asks anything else, or
 * nothing -- a colour, a size, an alignment or a part named by its row is
 * not a rule beside an invisible run, and such a format is not half written.
 */
export function pdfNativeCoverRulesAsk(format: unknown): PdfNativeInvisibleTextRuleFormat | null {
  if (!format || typeof format !== "object" || Array.isArray(format)) return null;
  const record = format as Record<string, unknown>;
  const out: { underline?: boolean; strikethrough?: boolean; segments?: PdfNativeCoverRuleSegment[] } = {};
  let asked = false;
  for (const [key, value] of Object.entries(record)) {
    if (value === undefined) continue;
    if (RULE_KEYS.has(key)) {
      if (typeof value !== "boolean") return null;
      out[key as PdfNativeCoverRuleStyle] = value;
      asked = true;
    } else if (key === "segments") {
      if (!Array.isArray(value)) return null;
      const segments: PdfNativeCoverRuleSegment[] = [];
      for (const entry of value as unknown[]) {
        if (!entry || typeof entry !== "object") return null;
        const segment = entry as Record<string, unknown>;
        if (Object.entries(segment).some(([name, field]) => field !== undefined && !SEGMENT_KEYS.has(name))) return null;
        const { row, start, end, underline, strikethrough } = segment;
        if (typeof row !== "number" || row >= 0 || typeof start !== "number" || typeof end !== "number" ||
          !Number.isInteger(start) || !Number.isInteger(end) || start < 0 || !(end > start)) return null;
        if ((underline !== undefined && typeof underline !== "boolean") ||
          (strikethrough !== undefined && typeof strikethrough !== "boolean")) return null;
        if (underline === undefined && strikethrough === undefined) return null;
        segments.push(Object.freeze({ row: -1, start, end,
          ...(underline === undefined ? {} : { underline }), ...(strikethrough === undefined ? {} : { strikethrough }) }));
      }
      if (segments.length) {
        out.segments = segments;
        asked = true;
      }
    } else {
      return null;
    }
  }
  return asked ? Object.freeze(out) : null;
}



export function pdfNativeCoverRuleSegments(parts: PdfTextPartFormats): readonly PdfNativeCoverRuleSegment[] {
  return Object.freeze(parts.flatMap((part): PdfNativeCoverRuleSegment[] =>
    part.underline === undefined && part.strikethrough === undefined ? [] : [Object.freeze({
      row: -1, start: part.start, end: part.end,
      ...(part.underline === undefined ? {} : { underline: part.underline }),
      ...(part.strikethrough === undefined ? {} : { strikethrough: part.strikethrough }),
    })]));
}

/** Whether the parts press a rule at all. */
export function pdfNativeCoverRulesPressed(parts: PdfTextPartFormats): boolean {
  return parts.some((part) => part.underline !== undefined || part.strikethrough !== undefined);
}

/** The characters of `text` a rule runs under: the ones standing, with the
 * parts pressed set over them. */
function ruleCharacters(
  standing: PdfNativeCoverRuleStretches,
  parts: PdfTextPartFormats,
  style: PdfNativeCoverRuleStyle,
  text: string,
): Uint8Array {
  const on = new Uint8Array(text.length);
  for (const stretch of standing) {
    for (let at = Math.max(0, stretch.start); at < Math.min(text.length, stretch.end); at += 1) on[at] = 1;
  }
  for (const part of parts) {
    const value = part[style];
    if (value === undefined) continue;
    for (let at = Math.max(0, part.start); at < Math.min(text.length, part.end); at += 1) on[at] = value ? 1 : 0;
  }
  return on;
}

/**
 * The stretches of `text` a rule runs under once the parts pressed are set
 * over the ones standing. Blanks between two stretches join them (Word
 * underlines the space between two underlined words), and no stretch starts
 * or ends on a blank.
 */
export function pdfNativeCoverRuleStretches(
  standing: PdfNativeCoverRuleStretches,
  parts: PdfTextPartFormats,
  style: PdfNativeCoverRuleStyle,
  text: string,
): PdfNativeCoverRuleStretches {
  const on = ruleCharacters(standing, parts, style, text);
  const out: { start: number; end: number }[] = [];
  let open: { start: number; end: number } | null = null;
  for (let at = 0; at < text.length; at += 1) {
    if (!text[at].trim()) continue;
    if (on[at]) {
      if (open) open.end = at + 1;
      else out.push(open = { start: at, end: at + 1 });
    } else {
      open = null;
    }
  }
  return Object.freeze(out.map((stretch) => Object.freeze(stretch)));
}

/**
 * Whether the button of a rule reads pressed over [start, end) of `text`:
 * every character there that is not a blank has the rule (Word). A caret
 * reads the character before it (the first, at the start), as the part
 * formats read (pdfTextPartFieldsOver).
 */
export function pdfNativeCoverRulePressed(
  standing: PdfNativeCoverRuleStretches,
  parts: PdfTextPartFormats,
  style: PdfNativeCoverRuleStyle,
  range: Readonly<{ start: number; end: number }> | null,
  text: string,
): boolean {
  let start = range ? Math.max(0, range.start) : 0;
  let end = range ? Math.min(text.length, range.end) : text.length;
  if (!(end > start)) {
    start = Math.max(0, start - 1);
    end = Math.min(text.length, start + 1);
    if (!(end > start)) return false;
  }
  const on = ruleCharacters(standing, parts, style, text);
  let seen = false;
  for (let at = start; at < end; at += 1) {
    if (!text[at].trim() && end - start > 1) continue;
    if (!on[at]) return false;
    seen = true;
  }
  return seen;
}

/** What a text box carries of a scanned line's rules once the line is its
 * own: the rules of its whole text, and the parts under a rule of their own. */
export type PdfNativeCoverRulesCarry = Readonly<{
  whole: Readonly<{ underline?: true; strikethrough?: true }>;
  parts: readonly Readonly<{ start: number; end: number; underline?: true; strikethrough?: true }>[];
}>;

/**
 * The rules a scanned line shows (the ones standing in the page beside its
 * invisible runs, read on `standingText`, and the ones pressed this session
 * on `text`), as its text box carries them from the moment it is one: a
 * character format pressed after a rule, or words typed into an underlined
 * line, must not lose the rule (the page's own is then not written, or
 * stands under the box's cover). Words typed since move the standing rules
 * with them (pdfTextPartFormatsAfterEdit). Null: no rule to carry.
 */
export function pdfNativeCoverRulesCarry(
  standing: Readonly<{ underline: PdfNativeCoverRuleStretches; strikethrough: PdfNativeCoverRuleStretches }> | null,
  standingText: string,
  parts: PdfTextPartFormats,
  text: string,
): PdfNativeCoverRulesCarry | null {
  const whole: { underline?: true; strikethrough?: true } = {};
  let carried: PdfTextPartFormats = [];
  let any = false;
  for (const style of PDF_NATIVE_COVER_RULE_STYLES) {
    let stood: PdfTextPartFormats = [];
    for (const stretch of standing?.[style] ?? []) {
      stood = applyPdfTextPartFormat(stood, stretch, { [style]: true });
    }
    if (standingText !== text) stood = pdfTextPartFormatsAfterEdit(stood, standingText, text);
    const stretches = pdfNativeCoverRuleStretches(
      stood.map((part) => ({ start: part.start, end: part.end })), parts, style, text);
    if (stretches.length === 0) continue;
    any = true;
    if (pdfNativeCoverRulePressed(stretches, [], style, { start: 0, end: text.length }, text) && text.trim()) {
      whole[style] = true;
    } else {
      for (const stretch of stretches) carried = applyPdfTextPartFormat(carried, stretch, { [style]: true });
    }
  }
  if (!any) return null;
  return Object.freeze({
    whole: Object.freeze(whole),
    parts: Object.freeze(carried.map((part: PdfTextPartFormat) => Object.freeze({
      start: part.start, end: part.end,
      ...(part.underline ? { underline: true as const } : {}),
      ...(part.strikethrough ? { strikethrough: true as const } : {}),
    }))),
  });
}
