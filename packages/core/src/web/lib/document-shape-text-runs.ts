/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { ptToPx } from "./office-units";
import { normalizeCellCharacterStyle } from "./presentation-cell-rich-text";
import type { RunStyle, TextRun } from "./presentation-runs";
import { normalizeOfficeTextPaint } from "./office-text-paint";

/** The run fields a shape's writers state (`w:b`, `w:i`, `w:u`, `w:strike`, `w:color`, `w:shd`, `w:sz`, `w:rFonts`, `w:caps` / `w:smallCaps`). */
export const DOCUMENT_SHAPE_RUN_STYLE_KEYS = [
  "bold",
  "italic",
  "underline",
  // The underline's line: the slide run model's kinds (a table cell's
  // `presentation-text-decoration`) and Word's `words`, what a Word run
  // states with `w:u` (dotted, dashed, wavy, thick, double, words only).
  "underlineStyle",
  "strike",
  "color",
  // The slide run's highlight (`#rrggbb` behind the words), what a Word run
  // states as character shading (`w:shd` clear with that fill).
  "highlight",
  "fontSize",
  "fontFamily",
  // Font > All caps / Small caps: the slide run's own `caps` (PPT-B362),
  // what a Word group member's run states with `w:caps` / `w:smallCaps`.
  "caps",
  // Font > Superscript / Subscript: the slide run's own two baselines, what a
  // Word run states with `w:vertAlign` (at most one is on).
  "superscript",
  "subscript",
  // WordArt on these words (Word's w14 run effects: text fill, outline,
  // glow, shadow, reflection, 3-D), what WordArt Styles, Text Fill, Text
  // Outline and Text Effects write on the selected words.
  "textPaint",
  // Font > Advanced > Spacing (`w:spacing`, twentieths of a point): the slide
  // run's own `letterSpacing`, in points (corpus R33-D30, a group member's
  // condensed title).
  "letterSpacing",
  // Font > Advanced > Scale (`w:w`, a whole percent 1-600): the body run's
  // own `charScale` (corpus R103-D03, a group's labels at 95-119%).
  "charScale",
  // Asian Layout > Fit Text (`w:fitText`): the width in twips (1-31680) and
  // the region (`w:id`) its neighbouring runs share -- a body run's fitText
  // (corpus R148-D14, a BODIK form's group labels).
  "fitTextTwips",
  "fitTextId",
  // Font > Advanced > Position (`w:position`, half-points raised or lowered):
  // the body run's own textPosition (corpus R158-D29).
  "positionHalfPoints",
  // Font > Hidden (`w:vanish`): Word neither shows nor prints these words;
  // the shape keeps them in its text and does not draw them -- the body
  // run's own hidden (corpus R475-D21, a group member's hidden words).
  "hidden",
] as const;

export type DocumentShapeRunStyleKey = (typeof DOCUMENT_SHAPE_RUN_STYLE_KEYS)[number];

const MAX_RUNS = 512;

/** The bound a page link holds (`MAX_PICTURE_LINK_HREF_CHARS`). */
const MAX_RUN_LINK_CHARS = 2_000;
/** The schemes the link policy lets a page open (document_link_policy.py). */
const CLICKABLE_LINK_SCHEMES = new Set(["http", "https", "mailto", "tel", "sms", "ftp", "ftps", "file"]);

/**
 * A run's hyperlink (`href`, the presentation run model's own field) as a
 * Docs shape keeps it, or `undefined`: a place in the document (`#name`), a
 * relative address, or a scheme the link policy opens. Only the words' runs
 * carry one; an inline atom's `runStyle` states no link. The server twin is
 * `shape_run_link_href` (document_shape_text_runs.py). Imported links are
 * carried; the shape's text editor does not author new ones.
 */
export function documentShapeRunLinkHref(value: unknown): string | undefined {
  if (typeof value !== "string" || !value || value !== value.trim()) return undefined;
  if (value.length > MAX_RUN_LINK_CHARS) return undefined;
  if (value.startsWith("#")) return /^#[^\s<>"']{1,255}$/.test(value) ? value : undefined;
  if (/[\u0000-\u0020\u007f]/.test(value)) return undefined;
  const scheme = /^([a-z][a-z0-9+\-.]*):/i.exec(value);
  if (!scheme) return value;
  return CLICKABLE_LINK_SCHEMES.has(scheme[1].toLowerCase()) ? value : undefined;
}

/**
 * One run's style restricted to what the Word run writers state, or
 * `undefined` when none of it is left. Also the vocabulary of a Docs
 * text-bearing inline atom's own formatting (`runStyle`, native-document/inlineAtomRunStyle.ts).
 */
/** Word's words-only underline (`w:u w:val="words"`), which the slide kinds lack. */
const WORDS_UNDERLINE = "words";

export function normalizeDocumentShapeRunStyle(value: unknown): RunStyle | undefined {
  if (!value || typeof value !== "object") return undefined;
  // A run whose only formatting is one the cell model lacks (`charScale`)
  // is still a run with formatting.
  const style = normalizeCellCharacterStyle(value) ?? ({} as RunStyle);
  const kept: Record<string, unknown> = {};
  for (const key of DOCUMENT_SHAPE_RUN_STYLE_KEYS) {
    if (key === "fitTextTwips" || key === "fitTextId") {
      // Read from the raw run: the cell model has no Fit Text (server twin:
      // document_shape_text_runs.normalize_document_shape_run_style). A
      // region without a width is no Fit Text.
      const raw = value as Record<string, unknown>;
      const twips = raw.fitTextTwips;
      if (!(typeof twips === "number" && Number.isInteger(twips) && twips >= 1 && twips <= 31680)) continue;
      if (key === "fitTextTwips") kept.fitTextTwips = twips;
      else if (typeof raw.fitTextId === "number" && Number.isInteger(raw.fitTextId)) kept.fitTextId = raw.fitTextId;
      continue;
    }
    if (key === "hidden") {
      // Read from the raw run: the cell model has no hidden words (server
      // twin: document_shape_text_runs.normalize_document_shape_run_style).
      if ((value as Record<string, unknown>).hidden === true) kept.hidden = true;
      continue;
    }
    if (key === "charScale") {
      // Read from the raw run: the cell model has no scale (server twin:
      // document_shape_text_runs.normalize_document_shape_run_style).
      const scale = (value as Record<string, unknown>).charScale;
      if (typeof scale === "number" && Number.isInteger(scale) && scale >= 1 && scale <= 600 && scale !== 100) {
        kept.charScale = scale;
      }
      continue;
    }
    if (key === "underlineStyle") {
      // A kind only on an underline. The cell model keeps the slide kinds;
      // Word's `words` is read from the raw run (server twin:
      // document_shape_text_runs.normalize_document_shape_run_style).
      const raw = (value as Record<string, unknown>).underlineStyle;
      if (style.underline === true && (style.underlineStyle || raw === WORDS_UNDERLINE)) {
        kept.underlineStyle = style.underlineStyle ?? WORDS_UNDERLINE;
      }
      continue;
    }
    const field = (style as Record<string, unknown>)[key];
    if (field === undefined) continue;
    if ((key === "color" || key === "highlight") && !(typeof field === "string" && /^#[0-9a-f]{6}$/i.test(field)))
      continue;
    if (key === "fontFamily" && !(typeof field === "string" && field.trim() && field.length <= 64)) continue;
    // 1-1638 pt, the shape writers' `w:sz` range, in CSS px.
    if (key === "fontSize" && !(typeof field === "number" && field >= ptToPx(1) && field <= ptToPx(1638))) continue;
    if (key === "caps" && field !== "all" && field !== "small") continue;
    if ((key === "superscript" || key === "subscript") && typeof field !== "boolean") continue;
    if (key === "letterSpacing" && !(typeof field === "number" && Number.isFinite(field))) continue;
    if (
      key === "positionHalfPoints" &&
      !(typeof field === "number" && Number.isInteger(field) && field !== 0 && Math.abs(field) <= 3168)
    )
      continue;
    if (key === "subscript" && field === true && kept.superscript === true) continue;
    if (key === "textPaint") {
      try {
        const paint = normalizeOfficeTextPaint(field);
        if (paint) kept.textPaint = paint;
      } catch {
        /* not a paint */
      }
      continue;
    }
    kept[key] = key === "color" || key === "highlight" ? (field as string).toLowerCase() : field;
  }
  return Object.keys(kept).length ? (kept as RunStyle) : undefined;
}

/**
 * The attributes of the body `contentControl` mark (the server's
 * `document_docx_content_controls.MARK_ATTRS`), which a group member's run
 * carries when its words stand in Word's content control (corpus R75-D17's
 * bound Plain Text control, R125-D30's cover Year date picker). The words'
 * runs only; the server checks each value as the body mark's.
 */
const RUN_CONTENT_CONTROL_KEYS = new Set([
  "controlType",
  "alias",
  "tag",
  "controlId",
  "lock",
  "listItems",
  "checked",
  "dateFormat",
  "dateFullDate",
  "multiline",
  "placeholder",
  "checkedState",
  "uncheckedState",
  "controlKey",
  "docPartGallery",
  "docPartUnique",
  "temporary",
]);

export function documentShapeRunContentControl(value: unknown): Record<string, unknown> | undefined {
  if (!value || typeof value !== "object" || Array.isArray(value)) return undefined;
  const record = value as Record<string, unknown>;
  if (typeof record.controlType !== "string") return undefined;
  if (Object.keys(record).some((key) => !RUN_CONTENT_CONTROL_KEYS.has(key))) return undefined;
  return JSON.parse(JSON.stringify(record)) as Record<string, unknown>;
}

/**
 * The shape's runs, or `undefined` when they carry no formatting or do not
 * spell `text` exactly (a stale reading is dropped, never applied to other
 * words).
 */
export function normalizeDocumentShapeTextRuns(
  value: unknown,
  text: string,
): TextRun[] | undefined {
  if (!Array.isArray(value) || value.length === 0 || value.length > MAX_RUNS) return undefined;
  const runs: TextRun[] = [];
  for (const raw of value) {
    if (!raw || typeof raw !== "object") return undefined;
    const record = raw as { text?: unknown; style?: unknown };
    if (typeof record.text !== "string") return undefined;
    if (!record.text) continue;
    const href = documentShapeRunLinkHref(
      record.style && typeof record.style === "object"
        ? (record.style as { href?: unknown }).href
        : undefined,
    );
    const formatted = normalizeDocumentShapeRunStyle(record.style);
    const control = documentShapeRunContentControl(
      record.style && typeof record.style === "object"
        ? (record.style as { contentControl?: unknown }).contentControl
        : undefined,
    );
    const linked = href ? ({ ...formatted, href } as RunStyle) : formatted;
    const style = control ? ({ ...linked, contentControl: control } as RunStyle) : linked;
    runs.push(style ? { text: record.text, style } : { text: record.text });
  }
  // Adjacent runs of one style are one run. The styles are already this
  // model's own (a link the presentation's run policy would not keep, a
  // `#bookmark`, included), so they are compared as they are.
  const merged: TextRun[] = [];
  for (const run of runs) {
    const last = merged.at(-1);
    if (last && JSON.stringify(last.style) === JSON.stringify(run.style)) last.text += run.text;
    else merged.push({ ...run });
  }
  if (merged.map((run) => run.text).join("") !== text) return undefined;
  return merged.some((run) => run.style !== undefined) ? merged : undefined;
}

/**
 * The runs without the given fields: what a whole-shape command leaves, since
 * formatting the selected shape formats all of its words (Word, the shape's
 * border selected) and a run's own value must no longer override it.
 */
export function documentShapeTextRunsWithout(
  runs: readonly TextRun[] | null | undefined,
  text: string,
  keys: readonly DocumentShapeRunStyleKey[],
): TextRun[] | undefined {
  if (!runs?.length) return undefined;
  return normalizeDocumentShapeTextRuns(
    runs.map((run) => {
      if (!run.style) return { text: run.text };
      const style = { ...run.style } as Record<string, unknown>;
      for (const key of keys) delete style[key];
      return { text: run.text, style };
    }),
    text,
  );
}

/** Points on the shape, CSS px in a run. */
export function documentShapeRunFontSizePt(style: RunStyle | undefined): number | undefined {
  return typeof style?.fontSize === "number" ? Math.round(((style.fontSize * 3) / 4) * 10) / 10 : undefined;
}
