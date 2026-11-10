/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Pure logic for PDF AcroForm interactive form-fill (Acrobat-style fillable
 * forms).
 *
 * pdf.js exposes a page's interactive form fields via `page.getAnnotations()`:
 * each `subtype === "Widget"` annotation carries a `fieldType` (`Tx` text, `Btn`
 * checkbox/radio/pushbutton, `Ch` choice/list/combo), a `fieldName`, a `rect`
 * (the widget box in PDF user space), and — for choices / radios — an option
 * list. This module turns those loosely-typed raw widgets into a typed,
 * sanitized list the viewer renders as an interactive overlay (one input /
 * checkbox / radio / select positioned over each widget rect, mapped to
 * page-relative percent like the existing link / highlight overlays), and holds
 * everything needed to PERSIST + RE-APPLY the user-entered values.
 *
 * The persisted shape is a SINGLE additive `pdf_form_values` data-item
 * annotation: a flat `{ [fieldName]: value }` map (value is a string for
 * text/single-choice, a string array for a multi-select list, a boolean for a
 * checkbox, or the selected export value for a radio group). This is additive
 * to the `pdf_highlight` / `pdf_note` / `pdf_ink`
 * annotation types — old PDFs with none of these keep loading unchanged. Storing
 * ONE map (rather than one annotation per field) means a re-apply on reload is a
 * single lookup keyed by `fieldName`, and an edited field overwrites in place.
 *
 * Everything here is pure (no DOM, no pdf.js runtime) so it can be unit-tested
 * without jsdom or a live document. The interactive RENDER + the real pdf.js
 * widget layout (jsdom has no pdf.js layout) is browser-gated; the viewer
 * converts each widget's PDF-space rect into a viewport rect with pdf.js
 * (rotation-aware) and then calls the pure {@link widgetRectToPercent} to
 * position the overlay — reusing the SAME `viewportRectToPercent` geometry the
 * link overlays use, so there is one definition of "rect -> page percent".
 */

import { isPdfImageFieldActions, PDF_EMAIL_VALIDATE_SCRIPT, pdfTextFieldFormatFromScripts, type PdfTextFieldFormat } from "./form-field-format";
import { sanitizeLinkUrl } from "./doc-search";
import { PDF_LINK_OVERLAY_MAX_URL, viewportRectToPercent } from "./links";
import {
  parseCalculationScript,
  parseFormatScript,
  parseKeystrokeScript,
  type PdfKeystrokeRule,
} from "./acroform-scripts";

/**
 * The interactive field kinds we recognize. `text` / `checkbox` / `radio` /
 * `choice` carry a fillable value persisted in the `pdf_form_values` map.
 * `signature` is a `Sig` widget: it carries NO form-fill value (it is never put
 * into the values map — {@link normalizeFieldValue} returns null for it). It is
 * recognized here only so the viewer can place a "Sign here" affordance at the
 * field's rect; the actual (visual, non-cryptographic) signature is stored as a
 * SEPARATE `pdf_signatures` annotation (lib/pdf/signatures.ts), exactly like
 * sticky notes / ink live outside the form-values map. `resetButton` is the
 * one push-button action we expose: pdf.js must prove it is an unrestricted
 * native `/ResetForm` action. `printButton` similarly requires pdf.js to prove
 * the standard named Print action. `navigationButton` is limited to standard
 * page and view-history Named actions; `viewButton` is limited to safe view
 * commands implemented by this reader. `saveButton` requires the standard
 * SaveAs action and routes through the reader's current-copy save policy.
 * `goToPageButton` requires the standard GoToPage action and focuses the
 * reader's existing page-number control instead of storing a destination.
 * `webLinkButton` requires a standard `/URI` action whose URL passes the same
 * bounded HTTP/HTTPS sanitizer as native link annotations. `imageField` is
 * Acrobat's Image Field: a push button whose one action is the call that asks
 * for a picture (it is never run; the reader places the picture itself).
 * Other push buttons remain fail-closed.
 */
export type AcroFieldType =
  | "text"
  | "checkbox"
  | "radio"
  | "choice"
  | "signature"
  | "resetButton"
  | "printButton"
  | "navigationButton"
  | "viewButton"
  | "saveButton"
  | "goToPageButton"
  | "webLinkButton"
  | "imageField";

export type PdfPageNavigationAction =
  | "FirstPage"
  | "PrevPage"
  | "NextPage"
  | "LastPage";

export type PdfViewHistoryAction = "GoBack" | "GoForward";

export type PdfNavigationButtonAction =
  | PdfPageNavigationAction
  | PdfViewHistoryAction;

export type PdfViewCommandAction =
  | "Find"
  | "FullScreen"
  | "ZoomIn"
  | "ZoomOut"
  | "FitPage"
  | "FitWidth";

const PDF_PAGE_NAVIGATION_ACTIONS = new Set<PdfPageNavigationAction>([
  "FirstPage",
  "PrevPage",
  "NextPage",
  "LastPage",
]);

const PDF_NAVIGATION_BUTTON_ACTIONS = new Set<PdfNavigationButtonAction>([
  ...PDF_PAGE_NAVIGATION_ACTIONS,
  "GoBack",
  "GoForward",
]);

const PDF_VIEW_COMMAND_ACTIONS = new Set<PdfViewCommandAction>([
  "Find",
  "FullScreen",
  "ZoomIn",
  "ZoomOut",
  "FitPage",
  "FitWidth",
]);

/** Accept only standard named actions that stay inside the current document. */
export function isPdfPageNavigationAction(
  value: unknown,
): value is PdfPageNavigationAction {
  return (
    typeof value === "string" &&
    PDF_PAGE_NAVIGATION_ACTIONS.has(value as PdfPageNavigationAction)
  );
}

/** Accept page and view-history actions that stay inside this reader. */
export function isPdfNavigationButtonAction(
  value: unknown,
): value is PdfNavigationButtonAction {
  return (
    typeof value === "string" &&
    PDF_NAVIGATION_BUTTON_ACTIONS.has(value as PdfNavigationButtonAction)
  );
}

/** Accept only safe Named view commands implemented by this reader. */
export function isPdfViewCommandAction(
  value: unknown,
): value is PdfViewCommandAction {
  return (
    typeof value === "string" &&
    PDF_VIEW_COMMAND_ACTIONS.has(value as PdfViewCommandAction)
  );
}

function normalizeWebLinkButtonUrl(value: unknown): string | null {
  return typeof value === "string" && value.length <= PDF_LINK_OVERLAY_MAX_URL
    ? sanitizeLinkUrl(value)
    : null;
}

/** Max length of a field NAME after normalization (defensive cap). */
export const ACROFORM_MAX_FIELD_NAME = 500;
/**
 * Max length of a TEXT / choice field VALUE after normalization, in UTF-16
 * code units (the unit `pdfFormValuesSchema`'s `z.string().max` counts). A
 * field's own `/MaxLen` is counted in characters instead; see
 * {@link normalizeFieldValue}.
 */
export const ACROFORM_MAX_TEXT = 4000;
/** Max decoded /TU alternate-name length retained for UI/accessibility. */
const ACROFORM_MAX_ALTERNATE_NAME = 512;
/** Max number of options kept for a choice field (defensive cap). */
export const ACROFORM_MAX_OPTIONS = 1000;
/** Max number of persisted field entries kept in one values map. */
export const ACROFORM_MAX_FIELDS = 16_384;
/** Largest native text-field `/DA` font size retained or authored, in points. */
export const ACROFORM_MAX_TEXT_FONT_SIZE = 512;
/** Default AcroForm text colour used when no explicit `/DA` colour exists. */
export const PDF_TEXT_FIELD_DEFAULT_COLOR = "#000000";
/** Default native Widget fill used for newly authored text fields. */
export const PDF_TEXT_FIELD_DEFAULT_BACKGROUND_COLOR = "#ffffff";






export const PDF_TEXT_FIELD_NO_COLOR = "transparent";

/** A fill or border colour: `#rrggbb`, or {@link PDF_TEXT_FIELD_NO_COLOR}. */
export function normalizePdfTextFieldPaint(value: unknown): string | null {
  return value === PDF_TEXT_FIELD_NO_COLOR
    ? PDF_TEXT_FIELD_NO_COLOR
    : normalizePdfTextFieldColor(value);
}
/** Default native Widget border used for newly authored text fields. */
export const PDF_TEXT_FIELD_DEFAULT_BORDER_COLOR = "#000000";
/** Default native Widget border width used for newly authored text fields. */
export const PDF_TEXT_FIELD_DEFAULT_BORDER_WIDTH = 1;
/** Defensive upper bound for an authored Widget border width, in points. */
export const ACROFORM_MAX_WIDGET_BORDER_WIDTH = 12;

/** One normalized value carried by the interactive form-fill pipeline. */
export type AcroFormValue = string | boolean | string[];
export type PdfFormScriptLoss =
  | "unsupported-calculate"
  | "unsupported-format"
  | "unsupported-validate"
   
  | "unsupported-keystroke"
  | "rich-text-value";

/**
 * A declarative numeric range the field's /AA /V script enforces
 * (`AFRange_Validate`): the only validation the reader runs, as a rule it
 * checks itself rather than a script it executes.
 */
export interface PdfFieldRangeValidation {
  /** Lower bound, inclusive, when the script checks one. */
  min: number | null;
  /** Upper bound, inclusive, when the script checks one. */
  max: number | null;
}

/** Native AcroForm text justification (`/Q`): left, centered, or right. */
export type PdfTextFieldAlignment = "left" | "center" | "right";
/** Native Widget annotation visibility/print contract for a text field. */
export type PdfTextFieldVisibility =
  | "visible"
  | "screenOnly"
  | "printOnly"
  | "hidden";

const PDF_ANNOTATION_INVISIBLE_FLAG = 1 << 0;
const PDF_ANNOTATION_HIDDEN_FLAG = 1 << 1;
const PDF_ANNOTATION_PRINT_FLAG = 1 << 2;
const PDF_ANNOTATION_NO_VIEW_FLAG = 1 << 5;

/** Resolve pdf.js's raw Widget `/F` bits without conflating view and print. */
export function normalizePdfTextFieldVisibility(
  value: unknown,
): PdfTextFieldVisibility {
  const flags =
    typeof value === "number" &&
    Number.isSafeInteger(value) &&
    value >= 0
      ? value
      : 0;
  if (
    (flags & PDF_ANNOTATION_INVISIBLE_FLAG) !== 0 ||
    (flags & PDF_ANNOTATION_HIDDEN_FLAG) !== 0
  ) {
    return "hidden";
  }
  const printable = (flags & PDF_ANNOTATION_PRINT_FLAG) !== 0;
  if ((flags & PDF_ANNOTATION_NO_VIEW_FLAG) !== 0) {
    return printable ? "printOnly" : "hidden";
  }
  return printable ? "visible" : "screenOnly";
}

/** Normalize an authored hex colour or pdf.js `/DA` RGB byte triplet. */
export function normalizePdfTextFieldColor(value: unknown): string | null {
  if (typeof value === "string") {
    return /^#[0-9a-f]{6}$/i.test(value) ? value.toLowerCase() : null;
  }

  const channels = Array.isArray(value)
    ? value
    : ArrayBuffer.isView(value) && !(value instanceof DataView)
      ? Array.from(value as unknown as ArrayLike<unknown>)
      : null;
  if (
    !channels ||
    channels.length < 3 ||
    channels.slice(0, 3).some(
      (channel) =>
        typeof channel !== "number" ||
        !Number.isInteger(channel) ||
        channel < 0 ||
        channel > 255,
    )
  ) {
    return null;
  }

  return `#${channels
    .slice(0, 3)
    .map((channel) => (channel as number).toString(16).padStart(2, "0"))
    .join("")}`;
}

// ISO 32000 field flags are one-based. ReadOnly is bit 1, Required is bit 2, a
// radio group's NoToggleToOff is bit 15; choice-field Combo is bit 18, Edit is
// bit 19, and MultiSelect is bit 22; text-field DoNotSpellCheck is bit 23 and
// DoNotScroll is bit 24.
// pdf.js exposes their raw aggregate as fieldFlags, while its annotation-level
// isEditable property describes PDF.js editor ownership and is not the
// AcroForm Edit flag.
const ACRO_FIELD_READ_ONLY_FLAG = 1 << 0;
const ACRO_FIELD_REQUIRED_FLAG = 1 << 1;
const ACRO_BUTTON_NO_TOGGLE_TO_OFF_FLAG = 1 << 14;
/** Radio widgets with the same export value are one option: selecting one
 * turns all of them on (ISO 32000-1 table 226, RadiosInUnison). */
const ACRO_BUTTON_RADIOS_IN_UNISON_FLAG = 1 << 25;
const ACRO_CHOICE_COMBO_FLAG = 1 << 17;
const ACRO_CHOICE_EDIT_FLAG = 1 << 18;
const ACRO_CHOICE_MULTISELECT_FLAG = 1 << 21;
const ACRO_TEXT_DO_NOT_SPELL_CHECK_FLAG = 1 << 22;
/** Tx field flag bit 26 (`RichText`): the value also lives in `/RV`. */
const ACRO_TEXT_RICH_TEXT_FLAG = 1 << 25;
/** Text field flag FileSelect (ISO 32000-1 table 228, bit 21). */
const ACRO_TEXT_FILE_SELECT_FLAG = 1 << 20;
const ACRO_TEXT_DO_NOT_SCROLL_FLAG = 1 << 23;

/**
 * Strip C0 control characters EXCEPT tab (U+0009) and newline (U+000A), plus
 * DEL (U+007F). Built from char codes so no literal control byte lives in the
 * source (keeps the file ASCII-clean + the lint rule happy). Carriage returns
 * are collapsed to LF before this runs, so they never reach here.
 */
const CONTROL_CHARS_REGEX = new RegExp(
  `[${String.fromCharCode(0)}-${String.fromCharCode(8)}` +
    `${String.fromCharCode(11)}-${String.fromCharCode(31)}` +
    `${String.fromCharCode(127)}]`,
  "g",
);

/**
 * Unpaired UTF-16 surrogates. With the `u` flag a well-formed pair is ONE
 * astral code point outside this class, so only lone halves match.
 */
const LONE_SURROGATE_REGEX = /[\ud800-\udfff]/gu;

/** A text value as the reader keeps it: CR/CRLF as LF, C0 controls (but tab
 *  and newline) and DEL removed, lone surrogates as U+FFFD. */
function cleanFormText(value: string): string {
  return value
    .replace(/\r\n?/g, "\n")
    .replace(CONTROL_CHARS_REGEX, "")
    .replace(LONE_SURROGATE_REGEX, "\ufffd");
}

/**
 * Cut `value` to at most `maxChars` Unicode code points AND at most
 * `maxUnits` UTF-16 code units, only ever on a code-point boundary, so the
 * result never ends in half of a surrogate pair. Lone surrogates must already
 * be replaced (see {@link LONE_SURROGATE_REGEX}).
 */
function truncateFormText(
  value: string,
  maxChars: number,
  maxUnits: number,
): string {
  // Code points never outnumber code units, so a short string fits both caps.
  if (value.length <= Math.min(maxChars, maxUnits)) return value;
  let end = 0;
  let chars = 0;
  while (end < value.length && chars < maxChars) {
    const unit = value.charCodeAt(end);
    const next = value.charCodeAt(end + 1);
    const width =
      unit >= 0xd800 && unit <= 0xdbff && next >= 0xdc00 && next <= 0xdfff
        ? 2
        : 1;
    if (end + width > maxUnits) break;
    end += width;
    chars += 1;
  }
  return value.slice(0, end);
}






export function formTextCharacterCount(value: string): number {
  return Array.from(value).length;
}

/**
 * HTML `maxLength` for a text field's input. The DOM attribute counts UTF-16
 * code units while PDF `/MaxLen` counts characters, so passing `maxLen`
 * straight through would refuse "A" + U+1F600 (three code units) in a
 * `/MaxLen 2` field that PDF viewers accept. Every code point is at most two
 * code units, so `2 * maxLen` never blocks a legal value;
 * {@link normalizeFieldValue} then enforces the exact character cap. Returns
 * undefined when the field has no usable `/MaxLen`.
 */
export function acroFormTextInputMaxLength(
  maxLen: number | null | undefined,
): number | undefined {
  if (typeof maxLen !== "number" || !Number.isInteger(maxLen) || maxLen <= 0) {
    return undefined;
  }
  return Math.min(maxLen * 2, ACROFORM_MAX_TEXT);
}

/** A raw Widget annotation as returned by pdf.js `getAnnotations()`. Loose. */
export interface RawWidgetAnnotation {
  /** Annotation subtype; only `"Widget"` is treated as a form field. */
  subtype?: unknown;
  /** AcroForm field type: `"Tx"` text, `"Btn"` button, `"Ch"` choice. */
  fieldType?: unknown;
  /** Fully-qualified field name (the persistence key). */
  fieldName?: unknown;
  /** Widget box in PDF user space: `[x1, y1, x2, y2]`. */
  rect?: unknown;
  /** The widget's /MK /R quarter turn, as pdf.js reads it. */
  rotation?: unknown;
  /** pdf.js id for the widget (used as a stable React key when present). */
  id?: unknown;
  /**
   * For `Btn` fields: a checkbox is a single button, a radio is part of a group.
   * pdf.js sets `radioButton: true` (and a shared `fieldName`) for radio members
   * and `checkBox: true` for checkboxes. `pushButton: true` marks a non-value
   * action button (e.g. submit/reset) which we never persist.
   */
  checkBox?: unknown;
  radioButton?: unknown;
  pushButton?: unknown;
  /** pdf.js's decoded native `/ResetForm` action, when present. */
  resetForm?: unknown;
  /** pdf.js's decoded named push-button action, when safely recognized. */
  action?: unknown;
  /** pdf.js's resolved `/URI` action target for a push button. */
  url?: unknown;
  /** Tooltip/alternate text authored on the widget. */
  alternativeText?: unknown;
  /** Current pdf.js field value (initial value baked into the PDF). */
  fieldValue?: unknown;
  /** The field's `/DV` (reset/default value), when pdf.js exposes it. */
  defaultFieldValue?: unknown;
  /** For a radio member: this widget's export ("on") value. */
  buttonValue?: unknown;
  /** For a checkbox: the non-Off appearance/export value exposed by pdf.js. */
  exportValue?: unknown;
  /** For `Ch` fields: the selectable options. */
  options?: unknown;
  /** Whether a `Ch` field is a dropdown rather than a list box. */
  combo?: unknown;
  /** Raw AcroForm field flags exposed by pdf.js. */
  fieldFlags?: unknown;
  /** Raw Widget annotation `/F` flags exposed by pdf.js. */
  annotationFlags?: unknown;
  /** Whether a `Ch` list field accepts multiple simultaneous selections. */
  multiSelect?: unknown;
  /** Whether the field is read-only (we render it disabled, never persist it). */
  readOnly?: unknown;
  /** Whether the field is required by the source AcroForm. */
  required?: unknown;
  /** Whether a `Tx` field accepts multiple lines. */
  multiLine?: unknown;
  /** Whether a `Tx` field masks its value while editing. */
  password?: unknown;
  /** Whether a `Tx` field divides its width into fixed character cells. */
  comb?: unknown;
  /** Native AcroForm `/Q` text justification: 0 left, 1 center, 2 right. */
  textAlignment?: unknown;
  /** Parsed inheritable `/DA` data exposed by pdf.js. */
  defaultAppearanceData?: unknown;
  /** Native Widget `/MK /BG` fill exposed by pdf.js as RGB bytes. */
  backgroundColor?: unknown;
  /** Native Widget `/MK /BC` border colour exposed by pdf.js as RGB bytes. */
  borderColor?: unknown;
  /** Native Widget `/BS` object exposed by pdf.js; only its width is used. */
  borderStyle?: unknown;
  /** pdf.js may expose a per-field max length for text fields. */
  maxLen?: unknown;
  /**
   * The field's additional actions, as pdf.js reports them: a map of trigger
   * name to the JavaScript lines the PDF carries. Only `Calculate` (/AA /C) and
   * `Format` (/AA /F) are read — the rest are ignored rather than run.
   */
  actions?: unknown;
  /** pdf.js: the widget draws its own normal appearance (/AP /N). */
  hasAppearance?: unknown;
}

/** A choice option: a stored export value + a human display label. */
export interface AcroFieldOption {
  /** The export value (`/Opt[i][0]`), possibly empty and possibly shared with
   * another option: a choice is identified by `index`, not by this string. */
  value: string;
  label: string;
  /** Position in the field's `/Opt` array (the identity `/I` refers to). */
  index: number;
}

/** A normalized, sanitized form field ready for rendering + filling. */
export interface AcroField {
  /** Stable id (pdf.js id when present, else a positional fallback). */
  id: string;
  /** The interactive field kind. */
  type: AcroFieldType;
  /** Fully-qualified field name, normalized + capped (the persistence key). */
  name: string;
  /** Native /TU alternate name used as the visible tooltip/accessibility label. */
  alternateName?: string | null;
  /** Widget box in PDF user space, normalized to `[x1,y1,x2,y2]`. */
  rect: [number, number, number, number];
  



  widgetRotation?: 0 | 90 | 180 | 270;
  /** True when the field is read-only (render disabled, never persist). */
  readOnly: boolean;
  /** True when the source AcroForm marks the field as required. */
  required: boolean;
  /** For text fields: whether the input is multi-line. */
  multiLine: boolean;
  /** For text fields: masks entry only; it does not encrypt the stored value. */
  password: boolean;
  /** For text fields: divides the field into maxLen character cells. */
  comb: boolean;
  /** For text fields: its Format as the editor writes it (form-field-format.ts),
   * "other" for Format / Keystroke / Validate scripts of another kind. */
  textFormat?: PdfTextFieldFormat | "other" | null;
  /** For text fields: whether readers should offer spell checking. */
  spellCheck: boolean;
  /** For text fields: whether content may scroll beyond the widget bounds. */
  scrollable: boolean;
  




  readOnlyReason?:
    | "name-too-long"
    | "value-too-long"
    | "options-too-many"
    


    | "file-select";
  /** For text fields: normalized native `/Q` text justification. */
  alignment: PdfTextFieldAlignment;
  





  alignmentInherited?: boolean;
  /** For text fields: fixed native `/DA` size in points; null means auto. */
  fontSize: number | null;
  /** For text fields: the `/DA` size was 0, which means auto-size: the
   * text is set as large as the widget's box allows. */
  autoFontSize?: boolean;
  /** For text fields: the CSS font family of the `/DA` font when it is a
   * standard face (Courier, Helvetica, Times, Symbol, ZapfDingbats), so the
   * live editor sets the value in the field's own face. */
  fontFamily?: string | null;
  /** For text fields: normalized native `/DA` foreground colour. */
  textColor: string | null;
  /** For text fields: normalized native Widget `/MK /BG` fill colour. */
  backgroundColor: string | null;
  /** For text fields: normalized native Widget `/MK /BC` border colour. */
  borderColor: string | null;
  /** For text fields: bounded native Widget `/BS /W` width in points. */
  borderWidth: number | null;
  /** For text fields: native Widget screen/print visibility semantics. */
  visibility: PdfTextFieldVisibility;
  /** Per-field text max length, when the PDF declares one (else null). */
  maxLen: number | null;
  /** For checkbox/radio widgets: this widget's export ("on") value. */
  exportValue: string | null;
  /** For radio groups: selecting the active member must leave it selected. */
  noToggleToOff: boolean;
  /** For radio groups: widgets sharing an export value are one option and
   * are all on or all off together (RadiosInUnison). Absent on other field
   * types and in fixtures that never name it. */
  radiosInUnison?: boolean;
  /** For choice fields: the allowed options. */
  options: AcroFieldOption[];
  /** For choice fields: whether this widget is a dropdown. */
  combo: boolean;
  /** For combo fields: whether the user may enter a value outside the list. */
  editable: boolean;
  /** For list fields: whether the native field accepts multiple selections. */
  multiSelect?: boolean;
  /** Value already stored in the source PDF, normalized by this field's rules. */
  initialValue?: AcroFormValue | null;
  




  defaultValue?: AcroFormValue | null;
  

  buttonLabel?: string;
  /** A push button's caption (`/MK /CA`), read from the file's bytes: the
   *  text drawn on a button that has no face of its own (user audit PDF
   *  0027). */
  buttonCaption?: string;
  /** True when a push button draws its own face (/AP): the page already
   *  shows the author's button, so the reader adds no visible text on it. */
  sourceFace?: boolean;
  /** Exact standard Named action carried by a supported action button. */
  buttonAction?: PdfNavigationButtonAction | PdfViewCommandAction;
  /** Sanitized HTTP/HTTPS target carried by a supported `/URI` push button. */
  buttonUrl?: string;
   
  resetScope?: PdfResetFormScope;
  /**
   * The field's /AA /C calculation script, when it has one. Carried so a
   * computed total can be recalculated while filling instead of showing the
   * value the author last saved.
   */
  calculateScript: string | null;
  /** The field's /AA /F format script, when it has one. */
  formatScript: string | null;
  /** The field's /AA /V validation script, when it has one. */
  validateScript?: string | null;
  /** The numeric range the validation script enforces, when it is the
   * standard `AFRange_Validate`; other validation scripts are reported as
   * losses and the field stays writable. */
  rangeValidation?: PdfFieldRangeValidation | null;
  /** The input rule the /AA /K keystroke script declares, when it is a
   * standard helper the reader enforces (ZIP, phone, number, date, mask);
   * other keystroke scripts are reported as losses. */
  keystrokeRule?: PdfKeystrokeRule | null;
  /** True for a text field whose value is rich text (`/RV`): the reader
   * shows it but does not rewrite it as plain text. */
  richText?: boolean;
  /** Declared Calculate/Format/Validate actions outside the reader's safe
   * allowlist, and rich text values the reader preserves rather than edits. */
  scriptLosses: PdfFormScriptLoss[];
}

const RANGE_VALIDATE_REGEX =
  /^\s*AFRange_Validate\s*\(\s*(true|false)\s*,\s*(-?\d+(?:\.\d+)?)\s*,\s*(true|false)\s*,\s*(-?\d+(?:\.\d+)?)\s*\)\s*;?\s*$/i;

/** The range an `AFRange_Validate(bGreater, nGreater, bLess, nLess)` script
 * enforces, or null when the script is anything else. */
export function parseRangeValidationScript(
  source: unknown,
): PdfFieldRangeValidation | null {
  if (typeof source !== "string") return null;
  const match = RANGE_VALIDATE_REGEX.exec(source);
  if (!match) return null;
  const min = match[1].toLowerCase() === "true" ? Number(match[2]) : null;
  const max = match[3].toLowerCase() === "true" ? Number(match[4]) : null;
  if ((min !== null && !Number.isFinite(min)) || (max !== null && !Number.isFinite(max))) {
    return null;
  }
  return { min, max };
}

/**
 * Whether `value` satisfies the field's declared numeric range. An empty
 * value passes (a required check is the field's own); a value that is not a
 * number fails the way Acrobat's AFRange_Validate does.
 */
export function fieldValueWithinRange(
  range: PdfFieldRangeValidation,
  value: string,
): boolean {
  const trimmed = value.trim();
  if (trimmed === "") return true;
  const numeric = Number(trimmed.replace(/,/g, "."));
  if (!Number.isFinite(numeric)) return false;
  if (range.min !== null && numeric < range.min) return false;
  if (range.max !== null && numeric > range.max) return false;
  return true;
}

/** Code points in `value` (what PDF `/MaxLen` counts). */
function countCodePoints(value: string): number {
  let count = 0;
  for (const _ of value) count += 1;
  return count;
}

/**
 * A full field takes new characters only by dropping the ones the user
 * typed, never the unselected tail. Given the value before the edit and the
 * value the input now holds, keep the untouched prefix and suffix and cut
 * the inserted middle to what still fits in `maxChars`; when nothing fits the
 * previous value comes back unchanged. Both ends are compared on code
 * points so a surrogate pair is never split.
 */
export function limitFormTextInsertion(
  previous: string,
  next: string,
  maxChars: number,
): string {
  if (countCodePoints(next) <= maxChars) return next;
  const before = Array.from(previous);
  const after = Array.from(next);
  let prefix = 0;
  while (
    prefix < before.length &&
    prefix < after.length &&
    before[prefix] === after[prefix]
  ) {
    prefix += 1;
  }
  let suffix = 0;
  while (
    suffix < before.length - prefix &&
    suffix < after.length - prefix &&
    before[before.length - 1 - suffix] === after[after.length - 1 - suffix]
  ) {
    suffix += 1;
  }
  const inserted = after.slice(prefix, after.length - suffix);
  const kept = prefix + suffix;
  const room = Math.max(0, maxChars - kept);
  if (room === 0) return previous;
  return [
    ...after.slice(0, prefix),
    ...inserted.slice(0, room),
    ...after.slice(after.length - suffix),
  ].join("");
}

/** Cap on a carried script; a pathological one is dropped, never truncated
 *  into something that might parse differently than the author wrote. */
const MAX_SCRIPT_LENGTH = 20_000;

/** Read one /AA trigger's JavaScript out of pdf.js's `actions` map. */
export function readFieldAction(
  actions: unknown,
  trigger: "Calculate" | "Format" | "Validate" | "Keystroke",
): string | null {
  if (!actions || typeof actions !== "object") return null;
  const value = (actions as Record<string, unknown>)[trigger];
  const lines = Array.isArray(value)
    ? value
    : typeof value === "string"
      ? [value]
      : null;
  if (!lines) return null;
  const joined = lines
    .filter((line): line is string => typeof line === "string")
    .join("\n")
    .trim();
  if (joined === "" || joined.length > MAX_SCRIPT_LENGTH) return null;
  return joined;
}

function hasDeclaredFieldAction(
  actions: unknown,
  trigger: "Calculate" | "Format" | "Validate" | "Keystroke",
): boolean {
  return (
    !!actions &&
    typeof actions === "object" &&
    Object.prototype.hasOwnProperty.call(actions, trigger)
  );
}

function isFiniteNumber(value: unknown): value is number {
  return typeof value === "number" && Number.isFinite(value);
}

/**
 * Normalize a field NAME: coerce to string + hard-cap length, keeping the name
 * EXACTLY as the PDF declares it. Returns "" for non-string / empty input (an
 * unnamed widget can't be persisted, so the caller drops it). This is NOT HTML
 * escaping — the name is only ever used as an object key / React key / aria
 * label rendered as a text node.
 *
 * The exact name MUST be preserved: the flatten exporter looks the field up in
 * pdf-lib by `getField(name)`, which matches the field's exact decoded name.
 * Collapsing interior whitespace or trimming ends (a tab / newline / trailing /
 * consecutive space is legal in an AcroForm `/T`) made the persisted name differ
 * from the real one, so the lookup missed and the user's filled value was
 * silently dropped from the downloaded PDF.
 */
export function normalizeFieldName(value: unknown): string {
  if (typeof value !== "string") return "";
  return value.slice(0, ACROFORM_MAX_FIELD_NAME);
}

/** 53-bit string hash (cyrb53), as 14 hex digits. */
function fieldNameDigest(value: string): string {
  let h1 = 0xdeadbeef;
  let h2 = 0x41c6ce57;
  for (let index = 0; index < value.length; index += 1) {
    const code = value.charCodeAt(index);
    h1 = Math.imul(h1 ^ code, 2654435761);
    h2 = Math.imul(h2 ^ code, 1597334677);
  }
  h1 = Math.imul(h1 ^ (h1 >>> 16), 2246822507) ^ Math.imul(h2 ^ (h2 >>> 13), 3266489909);
  h2 = Math.imul(h2 ^ (h2 >>> 16), 2246822507) ^ Math.imul(h1 ^ (h1 >>> 13), 3266489909);
  return (4294967296 * (2097151 & h2) + (h1 >>> 0)).toString(16).padStart(14, "0");
}











export function readerFieldIdentity(value: unknown): {
  name: string;
  truncated: boolean;
} {
  if (typeof value !== "string") return { name: "", truncated: false };
  if (value.length <= ACROFORM_MAX_FIELD_NAME) {
    return { name: value, truncated: false };
  }
  const digest = fieldNameDigest(value);
  return {
    name: `${value.slice(0, ACROFORM_MAX_FIELD_NAME - digest.length - 1)}#${digest}`,
    truncated: true,
  };
}







export function setFormValueEntry<T>(
  target: Record<string, T>,
  name: string,
  value: T,
): void {
  Object.defineProperty(target, name, {
    value,
    writable: true,
    enumerable: true,
    configurable: true,
  });
}

/**
 * Normalize raw pdf.js `rect` into an ordered `[x1,y1,x2,y2]` (x1<=x2, y1<=y2)
 * tuple, or `null` for non-4-finite / zero-area input. Mirrors `links.ts`
 * `normalizeRect` so widget boxes and link rects share one definition.
 */
export function normalizeWidgetRect(
  rect: unknown,
): [number, number, number, number] | null {
  if (!Array.isArray(rect) || rect.length < 4) return null;
  const [a, b, c, d] = rect;
  if (!isFiniteNumber(a) || !isFiniteNumber(b) || !isFiniteNumber(c) || !isFiniteNumber(d)) {
    return null;
  }
  const x1 = Math.min(a, c);
  const y1 = Math.min(b, d);
  const x2 = Math.max(a, c);
  const y2 = Math.max(b, d);
  if (x2 <= x1 || y2 <= y1) return null;
  return [x1, y1, x2, y2];
}

/**
 * Classify a raw widget into one of the supported {@link AcroFieldType}s, or
 * `null` when it is not a field we handle (unsupported push-button actions,
 * unknown field types,
 * non-Widget annotations).
 *
 *   - `Tx` -> "text"
 *   - `Btn` -> "checkbox" (single checkbox) or "radio" (group member);
 *     only unrestricted `/ResetForm`, named Print, and allowlisted in-reader
 *     navigation buttons are kept.
 *   - `Ch` -> "choice" (combo / list box)
 *   - `Sig` -> "signature" (a signature field; recognized so the viewer can
 *     offer a "Sign here" affordance at the rect, but it carries no form-fill
 *     value — see {@link AcroFieldType}). Previously these were dropped.
 */
/**
 * Which fields a `/ResetForm` resets (ISO 32000-1 12.7.5.3): its `/Fields`
 * by name and by reference, each with its descendants, reset or kept out by
 * the Include/Exclude flag. Absent on a reset of the whole form.
 */
export interface PdfResetFormScope {
  readonly fields: readonly string[];
  readonly refs: readonly string[];
  readonly include: boolean;
}

const RESET_SCOPE_MAX_TARGETS = 500;
const RESET_SCOPE_REF_RE = /^[1-9][0-9]{0,9}R(?:[0-9]{1,5})?$/u;






export function readResetFormAction(
  value: unknown,
): PdfResetFormScope | null | undefined {
  if (!value || typeof value !== "object" || Array.isArray(value)) return undefined;
  const action = value as Record<string, unknown>;
  if (
    typeof action.include !== "boolean" ||
    !Array.isArray(action.fields) ||
    !Array.isArray(action.refs)
  ) {
    return undefined;
  }
  // Without /Fields the flag is ignored and every field is reset.
  if (action.fields.length === 0 && action.refs.length === 0) return null;
  if (action.fields.length + action.refs.length > RESET_SCOPE_MAX_TARGETS) return undefined;
  const fields = action.fields.filter(
    (name): name is string => typeof name === "string" && name.length > 0 && name.length <= 1000,
  );
  const refs = action.refs.filter(
    (ref): ref is string => typeof ref === "string" && RESET_SCOPE_REF_RE.test(ref),
  );
  if (fields.length !== action.fields.length || refs.length !== action.refs.length) {
    return undefined;
  }
  return Object.freeze({
    fields: Object.freeze([...fields]),
    refs: Object.freeze([...refs]),
    include: action.include,
  });
}

export function classifyFieldType(raw: RawWidgetAnnotation): AcroFieldType | null {
  if (!raw || typeof raw !== "object") return null;
  if (raw.subtype !== "Widget") return null;
  const fieldType = typeof raw.fieldType === "string" ? raw.fieldType : "";
  if (fieldType === "Tx") return "text";
  if (fieldType === "Ch") return "choice";
  if (fieldType === "Sig") return "signature";
  if (fieldType === "Btn") {
    if (raw.pushButton === true) {
      if (isPdfImageFieldActions(raw.actions)) return "imageField";
      if (readResetFormAction(raw.resetForm) !== undefined) return "resetButton";
      if (raw.action === "Print") return "printButton";
      if (raw.action === "SaveAs") return "saveButton";
      if (raw.action === "GoToPage") return "goToPageButton";
      if (normalizeWebLinkButtonUrl(raw.url) !== null) return "webLinkButton";
      if (isPdfNavigationButtonAction(raw.action)) return "navigationButton";
      return isPdfViewCommandAction(raw.action) ? "viewButton" : null;
    }
    if (raw.radioButton === true) return "radio";
    // pdf.js sets checkBox for checkboxes; default a non-radio, non-push button
    // to a checkbox so a plainly-typed Btn still renders something fillable.
    return "checkbox";
  }
  return null;
}

/**
 * The CSS family of a `/DA` font resource name when it is one of the
 * standard 14 faces (or an alias readers give them); null for an embedded
 * or unknown face, which the live editor cannot set.
 */
export function standardFontFamily(fontName: unknown): string | null {
  if (typeof fontName !== "string" || !fontName) return null;
  const name = fontName.replace(/^\//, "").toLowerCase();
  if (/^(cour|courier)/.test(name)) return '"Courier New", Courier, monospace';
  if (/^(helv|helvetica|arial)/.test(name)) return 'Helvetica, Arial, sans-serif';
  if (/^(tiro|times)/.test(name)) return '"Times New Roman", Times, serif';
  if (/^(symb|symbol)/.test(name)) return 'Symbol, serif';
  if (/^(zadb|zapf)/.test(name)) return '"Zapf Dingbats", "ZapfDingbats", serif';
  return null;
}

/**
 * Normalize a raw `options` list (pdf.js exposes either `["label", …]` strings
 * or `{ exportValue, displayValue }` objects) into a typed, deduplicated,
 * capped list of `{ value, label }`. An option with an empty value is dropped.
 * Both value and label are length-capped to {@link ACROFORM_MAX_TEXT}.
 */
export function normalizeOptions(raw: unknown): AcroFieldOption[] {
  if (!Array.isArray(raw)) return [];
  const out: AcroFieldOption[] = [];
  // Every /Opt row is kept, in order, with its own index: ISO 32000-1 lets
  // two rows share an export value (they differ by display label and by
  // position, which /I names) and lets a row export the empty string (a
  // "clear" choice). Collapsing rows by export value, or replacing an empty
  // export value with its label, silently dropped selectable options and
  
  for (const opt of raw) {
    let value = "";
    let label = "";
    if (typeof opt === "string") {
      value = opt;
      label = opt;
    } else if (opt && typeof opt === "object") {
      const record = opt as Record<string, unknown>;
      const ev = record.exportValue;
      const dv = record.displayValue;
      if (typeof ev === "string") {
        value = ev;
        label = typeof dv === "string" ? dv : ev;
      } else if (typeof dv === "string") {
        // Only a display value: it is the export value too.
        value = dv;
        label = dv;
      } else {
        continue;
      }
    } else {
      continue;
    }
    value = value.slice(0, ACROFORM_MAX_TEXT);
    label = (label || value).slice(0, ACROFORM_MAX_TEXT);
    // A row with nothing to show and nothing to export cannot be selected.
    if (!value && !label) continue;
    out.push({ value, label, index: out.length });
    if (out.length >= ACROFORM_MAX_OPTIONS) break;
  }
  return out;
}

/** The option a stored value and (optionally) display label name, or null. A
 * display label disambiguates rows that share an export value. */
export function findAcroFieldOption(
  options: readonly AcroFieldOption[],
  value: string,
  displayValue?: string | null,
): AcroFieldOption | null {
  const matches = options.filter((option) => option.value === value);
  if (matches.length === 0) return null;
  if (matches.length > 1 && typeof displayValue === "string") {
    const labelled = matches.find((option) => option.label === displayValue);
    if (labelled) return labelled;
  }
  return matches[0];
}

/**
 * Normalize a page's raw widget annotations into a typed, sanitized list of
 * fillable {@link AcroField}s.
 *
 *   - keeps only `subtype === "Widget"` entries that classify to a supported
 *     field type (text / checkbox / radio / choice / signature) AND have a
 *     non-empty normalized name AND a valid rect,
 *   - drops unsupported push-button actions / unnamed / degenerate-rect widgets
 *     while retaining native reset-all buttons and signature placeholders,
 *   - assigns a stable id (pdf.js id when present, else a positional fallback
 *     `field-<pageNumber>-<index>` so positional ids stay unique across pages,
 *     mirroring the link overlay's id scheme).
 *
 * @param rawAnnotations The array from `page.getAnnotations()` (may be garbage).
 * @param pageNumber The 1-based page; namespaces the positional fallback id.
 */
export function normalizeFormFields(
  rawAnnotations: unknown,
  pageNumber: number = 0,
): AcroField[] {
  if (!Array.isArray(rawAnnotations)) return [];
  const pageKey = isFiniteNumber(pageNumber) ? pageNumber : 0;
  const out: AcroField[] = [];
  rawAnnotations.forEach((raw, index) => {
    if (!raw || typeof raw !== "object") return;
    const widget = raw as RawWidgetAnnotation;
    const type = classifyFieldType(widget);
    if (!type) return;
    const identity = readerFieldIdentity(widget.fieldName);
    const name = identity.name;
    if (!name) return;
    const rect = normalizeWidgetRect(widget.rect);
    if (!rect) return;

    const rawId = widget.id;
    const id =
      typeof rawId === "string" && rawId.length > 0
        ? `field-${rawId}`
        : `field-${pageKey}-${index}`;

    const rawExportValue =
      type === "radio"
        ? widget.buttonValue
        : type === "checkbox"
          ? widget.exportValue
          : null;
    const exportValue =
      typeof rawExportValue === "string"
        ? rawExportValue.slice(0, ACROFORM_MAX_TEXT)
        : null;
    const options = type === "choice" ? normalizeOptions(widget.options) : [];
    const fieldFlags =
      typeof widget.fieldFlags === "number" &&
      Number.isSafeInteger(widget.fieldFlags) &&
      widget.fieldFlags >= 0
        ? widget.fieldFlags
        : 0;
    const combo =
      type === "choice" &&
      (widget.combo === true ||
        (fieldFlags & ACRO_CHOICE_COMBO_FLAG) === ACRO_CHOICE_COMBO_FLAG);
    const editable =
      combo &&
      (fieldFlags & ACRO_CHOICE_EDIT_FLAG) === ACRO_CHOICE_EDIT_FLAG;
    const multiSelect =
      type === "choice" &&
      !combo &&
      (widget.multiSelect === true ||
        (fieldFlags & ACRO_CHOICE_MULTISELECT_FLAG) ===
          ACRO_CHOICE_MULTISELECT_FLAG);
    const actionButton =
      type === "resetButton" ||
      type === "printButton" ||
      type === "navigationButton" ||
      type === "viewButton" ||
      type === "saveButton" ||
      type === "goToPageButton" ||
      type === "webLinkButton";
    const buttonUrl =
      type === "webLinkButton" ? normalizeWebLinkButtonUrl(widget.url) : null;
    const alternateName =
      typeof widget.alternativeText === "string"
        ? widget.alternativeText
            .replace(/\r\n?/g, "\n")
            .replace(CONTROL_CHARS_REGEX, "")
            .trim()
            .slice(0, ACROFORM_MAX_ALTERNATE_NAME) || null
        : null;
    const buttonLabel = actionButton ? (alternateName ?? "") : "";
    const maxLenRaw = widget.maxLen;
    const maxLen =
      typeof maxLenRaw === "number" && Number.isInteger(maxLenRaw) && maxLenRaw > 0
        ? Math.min(maxLenRaw, ACROFORM_MAX_TEXT)
        : null;
    const defaultAppearanceData =
      widget.defaultAppearanceData &&
      typeof widget.defaultAppearanceData === "object"
        ? (widget.defaultAppearanceData as Record<string, unknown>)
        : null;
    const rawFontSize = defaultAppearanceData?.fontSize;
    const fontSize =
      type === "text" &&
      typeof rawFontSize === "number" &&
      Number.isFinite(rawFontSize) &&
      rawFontSize > 0 &&
      rawFontSize <= ACROFORM_MAX_TEXT_FONT_SIZE
        ? rawFontSize
        : null;
    // `/DA ... 0 Tf` is auto-size (PDF 32000-1 12.7.3.3): the reader sets
    // the text as large as the box allows, not at a size of its own.
    const autoFontSize = type === "text" && rawFontSize === 0;
    const fontFamily =
      type === "text"
        ? standardFontFamily(defaultAppearanceData?.fontName)
        : null;
    const textColor =
      type === "text"
        ? normalizePdfTextFieldColor(defaultAppearanceData?.fontColor)
        : null;
    const styledWidget =
      type === "text" || type === "checkbox" || type === "radio";
    const backgroundColor = styledWidget
      ? normalizePdfTextFieldColor(widget.backgroundColor)
      : null;
    const borderColor = styledWidget
      ? normalizePdfTextFieldColor(widget.borderColor)
      : null;
    const rawBorderStyle =
      widget.borderStyle && typeof widget.borderStyle === "object"
        ? (widget.borderStyle as Record<string, unknown>)
        : null;
    const rawBorderWidth = rawBorderStyle?.width;
    const borderWidth =
      styledWidget &&
      typeof rawBorderWidth === "number" &&
      Number.isFinite(rawBorderWidth) &&
      rawBorderWidth >= 0 &&
      rawBorderWidth <= ACROFORM_MAX_WIDGET_BORDER_WIDTH
        ? rawBorderWidth
        : null;
    // Every Widget honours the Hidden / NoView / Print annotation flags, not
    // only text fields: a hidden checkbox or choice must not be shown, nor
    
    const visibility = normalizePdfTextFieldVisibility(widget.annotationFlags);
    const calculateScript = readFieldAction(widget.actions, "Calculate");
    const formatScript = readFieldAction(widget.actions, "Format");
    const validateScript = readFieldAction(widget.actions, "Validate");
    const rangeValidation =
      type === "text" ? parseRangeValidationScript(validateScript) : null;
    const declared = (trigger: "Format" | "Keystroke" | "Validate") =>
      hasDeclaredFieldAction(widget.actions, trigger) ? readFieldAction(widget.actions, trigger) : undefined;
    const textFormat = type === "text"
      ? pdfTextFieldFormatFromScripts({ format: declared("Format"), keystroke: declared("Keystroke"), validate: declared("Validate") })
      : null;
    // An email field's rule is its validation (the editor's 电子邮件 script).
    const emailValidation = type === "text" && validateScript?.trim() === PDF_EMAIL_VALIDATE_SCRIPT;
    const keystrokeRule =
      type === "text"
        ? parseKeystrokeScript(readFieldAction(widget.actions, "Keystroke")) ??
          (emailValidation ? ({ kind: "email" } as const) : null)
        : null;
    const richText =
      type === "text" &&
      (fieldFlags & ACRO_TEXT_RICH_TEXT_FLAG) === ACRO_TEXT_RICH_TEXT_FLAG;
    const scriptLosses: PdfFormScriptLoss[] = [];
    if (
      hasDeclaredFieldAction(widget.actions, "Validate") &&
      rangeValidation === null &&
      !emailValidation
    ) {
      scriptLosses.push("unsupported-validate");
    }
    if (richText) scriptLosses.push("rich-text-value");
    // A field with an input rule the reader cannot enforce says so instead
    
    if (
      hasDeclaredFieldAction(widget.actions, "Keystroke") &&
      keystrokeRule === null
    ) {
      scriptLosses.push("unsupported-keystroke");
    }
    if (
      hasDeclaredFieldAction(widget.actions, "Calculate") &&
      (!calculateScript || parseCalculationScript(calculateScript) === null)
    ) {
      scriptLosses.push("unsupported-calculate");
    }
    if (
      hasDeclaredFieldAction(widget.actions, "Format") &&
      (!formatScript || parseFormatScript(formatScript) === null)
    ) {
      scriptLosses.push("unsupported-format");
    }

    // A text value longer than the editing cap is shown whole and read-only,
    
    const valueTooLong =
      type === "text" &&
      (maxLen === null ||
        (typeof maxLenRaw === "number" && maxLenRaw > ACROFORM_MAX_TEXT)) &&
      typeof widget.fieldValue === "string" &&
      widget.fieldValue.length > ACROFORM_MAX_TEXT;
    // A choice with more options than the model keeps is shown as the PDF
    // has it and read-only: with the list cut at the cap, a selection past it
    // read as "nothing selected", and the first key press replaced it with
    
    const optionsTooMany =
      type === "choice" &&
      Array.isArray(widget.options) &&
      widget.options.length > ACROFORM_MAX_OPTIONS;
    const readOnlyReason: AcroField["readOnlyReason"] = identity.truncated
      ? "name-too-long"
      : valueTooLong
        ? "value-too-long"
        : optionsTooMany
          ? "options-too-many"
          : type === "text" &&
              (fieldFlags & ACRO_TEXT_FILE_SELECT_FLAG) ===
                ACRO_TEXT_FILE_SELECT_FLAG
            ? "file-select"
            : undefined;

    const rawRotation = Number(widget.rotation);
    const widgetRotation =
      rawRotation === 90 || rawRotation === 180 || rawRotation === 270
        ? rawRotation
        : 0;
    out.push({
      id,
      type,
      name,
      alternateName,
      rect,
      ...(widgetRotation ? { widgetRotation } : {}),
      readOnly:
        readOnlyReason !== undefined ||
        widget.readOnly === true ||
        (fieldFlags & ACRO_FIELD_READ_ONLY_FLAG) === ACRO_FIELD_READ_ONLY_FLAG,
      ...(readOnlyReason ? { readOnlyReason } : {}),
      required:
        widget.required === true ||
        (fieldFlags & ACRO_FIELD_REQUIRED_FLAG) === ACRO_FIELD_REQUIRED_FLAG,
      multiLine: type === "text" && widget.multiLine === true,
      password: type === "text" && widget.password === true,
      comb: type === "text" && widget.comb === true,
      spellCheck:
        type === "text" &&
        (fieldFlags & ACRO_TEXT_DO_NOT_SPELL_CHECK_FLAG) === 0,
      scrollable:
        type === "text" && (fieldFlags & ACRO_TEXT_DO_NOT_SCROLL_FLAG) === 0,
      alignment:
        type === "text" && widget.textAlignment === 1
          ? "center"
          : type === "text" && widget.textAlignment === 2
            ? "right"
            : "left",
      ...(type === "text" &&
      widget.textAlignment !== 0 &&
      widget.textAlignment !== 1 &&
      widget.textAlignment !== 2
        ? { alignmentInherited: true }
        : {}),
      fontSize,
      autoFontSize,
      fontFamily,
      textColor,
      backgroundColor,
      borderColor,
      borderWidth,
      visibility,
      maxLen,
      exportValue,
      noToggleToOff:
        type === "radio" &&
        (fieldFlags & ACRO_BUTTON_NO_TOGGLE_TO_OFF_FLAG) ===
          ACRO_BUTTON_NO_TOGGLE_TO_OFF_FLAG,
      radiosInUnison:
        type === "radio" &&
        (fieldFlags & ACRO_BUTTON_RADIOS_IN_UNISON_FLAG) ===
          ACRO_BUTTON_RADIOS_IN_UNISON_FLAG,
      options,
      combo,
      editable,
      multiSelect,
      initialValue: valueTooLong
        ? cleanFormText(widget.fieldValue as string)
        : normalizeFieldValue(
            {
              type,
              maxLen,
              options,
              exportValue,
              editable,
              multiSelect,
            },
            widget.fieldValue,
          ),
      ...(widget.defaultFieldValue !== undefined &&
      widget.defaultFieldValue !== null
        ? {
            defaultValue: normalizeFieldValue(
              {
                type,
                maxLen,
                options,
                exportValue,
                editable,
                multiSelect,
              },
              widget.defaultFieldValue,
            ),
          }
        : {}),
      ...(actionButton && buttonLabel ? { buttonLabel } : {}),
      ...(actionButton && widget.hasAppearance === true
        ? { sourceFace: true }
        : {}),
      ...(type === "navigationButton" &&
        isPdfNavigationButtonAction(widget.action)
        ? { buttonAction: widget.action }
        : {}),
      ...(type === "viewButton" && isPdfViewCommandAction(widget.action)
        ? { buttonAction: widget.action }
        : {}),
      ...(buttonUrl ? { buttonUrl } : {}),
      ...(type === "resetButton" && readResetFormAction(widget.resetForm)
        ? { resetScope: readResetFormAction(widget.resetForm)! }
        : {}),
      calculateScript,
      formatScript,
      validateScript,
      rangeValidation,
      ...(keystrokeRule ? { keystrokeRule } : {}),
      ...(textFormat !== null ? { textFormat } : {}),
      richText,
      scriptLosses,
    });
  });
  return out;
}

/**
 * Normalize a single field's user-entered value to exactly what we persist +
 * render, per the field's kind:
 *
 *   - text: coerce to string, strip control chars except tab/newline, replace
 *     lone surrogates with U+FFFD, and cap length: at most `maxLen` CHARACTERS
 *     (Unicode code points, as PDF `/MaxLen` specifies) and at most
 *     {@link ACROFORM_MAX_TEXT} UTF-16 code units, never splitting a pair.
 *   - checkbox: a boolean (truthy -> true). Strings "true"/"on"/"yes"/"1" map
 *     to true; everything else to false.
 *   - radio: the value must be one of the group's export values OR "" (cleared);
 *     anything else returns "" (no selection) so a crafted value can't smuggle
 *     an arbitrary string into a radio group. Pass the group's member export
 *     values via `radioValues`.
 *   - single choice: pdf.js may expose the scalar `/V` as either a string or
 *     an exact one-string array. The value must still be one of the field's
 *     option values OR ""; empty, multi-value, or malformed arrays return
 *     "".
 *   - multi-select list: accepts one legacy string or a string array, retains
 *     only declared option values, de-duplicates in source order, and returns
 *     an array (including `[]` for an explicit clear).
 *   - signature/resetButton/printButton/navigationButton/viewButton/saveButton/
 *     goToPageButton/webLinkButton: returns `null` — none carries a fill value (the visual signature is a separate
 *     `pdf_signatures` annotation), so it never enters the persisted values
 *     map.
 *
 * Returns `null` for a signature/action button OR a field whose kind is unknown
 * (defensive). A null result is dropped by {@link normalizeFormValues}. Callers
 * persist the returned value verbatim, so this is the single source of truth for
 * "what a field value may be".
 */
export function normalizeFieldValue(
  field: Pick<AcroField, "type" | "maxLen" | "options" | "exportValue"> & {
    /** Editable combo boxes accept a bounded value outside their option list. */
    editable?: boolean;
    /** Multi-select list boxes carry an ordered array of export values. */
    multiSelect?: boolean;
    /** For a radio GROUP, all member export values (so any member is valid). */
    radioValues?: ReadonlyArray<string>;
     
    checkboxStates?: ReadonlyArray<string>;
  },
  value: unknown,
): AcroFormValue | null {
  switch (field.type) {
    case "text": {
      const str =
        typeof value === "string" ? value : value == null ? "" : String(value);
      // Collapse CR/CRLF to LF, strip C0 controls (keeping tab + newline) and
      // DEL, replace lone surrogates with U+FFFD, then cap. PDF `/MaxLen` is a
      // CHARACTER count (Acrobat / comb cells count code points, not UTF-16
      // units), while the global cap bounds persisted UTF-16 length; both cut
      // on a code-point boundary so the result is never ill-formed.
      const cleaned = cleanFormText(str);
      const maxChars =
        field.maxLen && field.maxLen > 0
          ? Math.min(field.maxLen, ACROFORM_MAX_TEXT)
          : ACROFORM_MAX_TEXT;
      return truncateFormText(cleaned, maxChars, ACROFORM_MAX_TEXT);
    }
    case "checkbox": {
      const states = field.checkboxStates ?? [];
      if (states.length >= 2) {
        
        // on-state names or "" (Off). A legacy boolean maps to the first
        // declared state, the state the boolean model used to write.
        if (typeof value === "string") {
          const name = value.startsWith("/") ? value.slice(1) : value;
          if (name === "" || name === "Off") return "";
          return states.includes(name) ? name : "";
        }
        if (typeof value === "boolean") return value ? states[0]! : "";
        return "";
      }
      if (typeof value === "boolean") return value;
      if (typeof value === "string") {
        if (field.exportValue && value === field.exportValue) return true;
        return ["true", "on", "yes", "1", "checked"].includes(
          value.trim().toLowerCase(),
        );
      }
      if (typeof value === "number") return value !== 0;
      return false;
    }
    case "radio": {
      if (typeof value !== "string") return "";
      if (value === "") return "";
      const allowed = new Set<string>(
        (field.radioValues && field.radioValues.length > 0
          ? field.radioValues
          : field.exportValue
            ? [field.exportValue]
            : []
        ).filter((v): v is string => typeof v === "string"),
      );
      return allowed.has(value) ? value : "";
    }
    case "choice": {
      const allowed = new Set(field.options.map((o) => o.value));
      if (field.multiSelect) {
        const rawValues = Array.isArray(value)
          ? value
          : typeof value === "string" && value !== ""
            ? [value]
            : [];
        const selected: string[] = [];
        const seen = new Set<string>();
        for (const raw of rawValues) {
          if (typeof raw !== "string" || !allowed.has(raw) || seen.has(raw)) {
            continue;
          }
          seen.add(raw);
          selected.push(raw);
          if (selected.length >= ACROFORM_MAX_OPTIONS) break;
        }
        return selected;
      }
      const scalarValue =
        typeof value === "string"
          ? value
          : Array.isArray(value) &&
              value.length === 1 &&
              typeof value[0] === "string"
            ? value[0]
            : null;
      if (scalarValue === null || scalarValue === "") return "";
      if (field.editable) {
        const cleaned = scalarValue
          .replace(/\r\n?/g, "\n")
          .replace(CONTROL_CHARS_REGEX, "")
          .replace(/[\t\n]/g, " ")
          .replace(LONE_SURROGATE_REGEX, "\ufffd");
        return truncateFormText(cleaned, ACROFORM_MAX_TEXT, ACROFORM_MAX_TEXT);
      }
      return allowed.has(scalarValue) ? scalarValue : "";
    }
    default:
      return null;
  }
}

/**
 * A merged per-field descriptor used to normalize a values map: one entry per
 * unique field NAME, with radio member export values collected into
 * `radioValues` so any member of the group is a valid value.
 */
export interface MergedField {
  type: AcroFieldType;
  maxLen: number | null;
  options: AcroFieldOption[];
  exportValue: string | null;
  editable: boolean;
  multiSelect: boolean;
  radioValues: string[];
  





  checkboxStates: string[];
  /** See {@link AcroField.richText}. */
  richText?: boolean;
  /** See {@link AcroField.rangeValidation}. */
  rangeValidation?: PdfFieldRangeValidation | null;
  /** See {@link AcroField.keystrokeRule}. */
  keystrokeRule?: PdfKeystrokeRule | null;
}

/** A checkbox widget's export value names an on-state when it is not Off. */
function isCheckboxOnState(value: string | null | undefined): value is string {
  return typeof value === "string" && value !== "" && value !== "Off";
}









export function multiStateCheckboxNames(
  fields: ReadonlyArray<Pick<AcroField, "type" | "name" | "exportValue">>,
): Set<string> {
  const statesByName = new Map<string, Set<string>>();
  for (const field of fields) {
    if (field.type !== "checkbox" || !isCheckboxOnState(field.exportValue)) {
      continue;
    }
    const states = statesByName.get(field.name) ?? new Set<string>();
    states.add(field.exportValue);
    statesByName.set(field.name, states);
  }
  const out = new Set<string>();
  for (const [name, states] of statesByName) {
    if (states.size >= 2) out.add(name);
  }
  return out;
}

/**
 * Collapse a flat field list (which has one entry PER widget, so a radio group
 * appears as several same-named members) into a name-keyed directory: one
 * {@link MergedField} per unique name, with all radio export values merged. This
 * is what {@link normalizeFormValues} uses to validate a values map, and what
 * the viewer uses to look up a field's kind by name.
 */
export function buildFieldDirectory(
  fields: ReadonlyArray<AcroField>,
): Map<string, MergedField> {
  const byName = new Map<string, MergedField>();
  for (const f of fields) {
    const existing = byName.get(f.name);
    if (!existing) {
      byName.set(f.name, {
        type: f.type,
        maxLen: f.maxLen,
        options: f.options,
        exportValue: f.exportValue,
        editable: f.editable,
        multiSelect: f.multiSelect === true,
        radioValues: f.type === "radio" && f.exportValue ? [f.exportValue] : [],
        checkboxStates:
          f.type === "checkbox" && isCheckboxOnState(f.exportValue)
            ? [f.exportValue]
            : [],
        richText: f.richText === true,
        rangeValidation: f.rangeValidation ?? null,
        keystrokeRule: f.keystrokeRule ?? null,
      });
      continue;
    }
    // Merge a same-named radio member's export value into the group.
    if (
      f.type === "radio" &&
      f.exportValue &&
      !existing.radioValues.includes(f.exportValue)
    ) {
      existing.radioValues.push(f.exportValue);
    }
    // A same-named checkbox widget with its own on-state makes the field
    
    if (
      f.type === "checkbox" &&
      isCheckboxOnState(f.exportValue) &&
      !existing.checkboxStates.includes(f.exportValue)
    ) {
      existing.checkboxStates.push(f.exportValue);
    }
    // If a later widget supplies more options (choice) keep the richer set.
    if (f.options.length > existing.options.length) {
      existing.options = f.options;
    }
    if (f.editable) existing.editable = true;
    if (f.multiSelect) existing.multiSelect = true;
  }
  return byName;
}

/**
 * Collapse the source PDF's current field values (`/V`, decoded by pdf.js as
 * `initialValue`) into one value per field name. This is a read-only display
 * baseline: reader-owned overrides are layered on top and only those overrides
 * are persisted, so merely opening a populated form never makes it dirty.
 *
 * Radio groups need special handling because each widget is normalized against
 * its own export value. Exactly one member therefore carries the group's
 * non-empty source value; empty siblings must not overwrite it.
 */
export function buildSourceFormValues(
  fields: ReadonlyArray<AcroField>,
): Record<string, AcroFormValue> {
  const out: Record<string, AcroFormValue> = {};
  const multiState = multiStateCheckboxNames(fields);
  let count = 0;
  for (const field of fields) {
    if (count >= ACROFORM_MAX_FIELDS) break;
    const value = field.initialValue;
    if (value === null || value === undefined) continue;
    if (field.type === "radio") {
      if (typeof value !== "string" || value === "") continue;
      if (!Object.hasOwn(out, field.name)) count += 1;
      setFormValueEntry(out, field.name, value);
      continue;
    }
    if (field.type === "checkbox" && multiState.has(field.name)) {
      // The field's value is the name of the one widget state that is on;
      
      if (!Object.hasOwn(out, field.name)) {
        setFormValueEntry(out, field.name, "");
        count += 1;
      }
      if (value === true && isCheckboxOnState(field.exportValue)) {
        setFormValueEntry(out, field.name, field.exportValue);
      }
      continue;
    }
    if (Object.hasOwn(out, field.name)) continue;
    setFormValueEntry(out, field.name, Array.isArray(value) ? [...value] : value);
    count += 1;
  }
  return out;
}










export function existingFieldDraftDefaultValue(
  field: Pick<AcroField, "defaultValue"> & { password?: boolean },
  options: { allowEmpty?: boolean } = {},
): { defaultValue?: string } {
  if (field.password) return {};
  const value = field.defaultValue;
  if (typeof value !== "string") return {};
  if (value === "" && !options.allowEmpty) return {};
  return { defaultValue: value };
}











export function calculationInputValues(
  fields: ReadonlyArray<Pick<AcroField, "type" | "name" | "exportValue">>,
  values: Readonly<Record<string, unknown>>,
): Map<string, unknown> {
  const out = new Map<string, unknown>(Object.entries(values));
  const exportValues = new Map<string, string>();
  for (const field of fields) {
    if (field.type !== "checkbox" || exportValues.has(field.name)) continue;
    if (isCheckboxOnState(field.exportValue)) {
      exportValues.set(field.name, field.exportValue);
    }
  }
  for (const [name, value] of out) {
    if (typeof value !== "boolean") continue;
    const exportValue = exportValues.get(name);
    if (!value) out.set(name, "Off");
    else if (exportValue !== undefined) out.set(name, exportValue);
  }
  return out;
}

/**
 * Validate + normalize a whole persisted `{ [fieldName]: value }` map against a
 * directory of known fields, dropping entries whose name is unknown / blank or
 * whose value normalizes away. The result is a clean map ready to PUT (so
 * persisted bytes match what we render) and to re-apply on reload.
 *
 * Each value is run through {@link normalizeFieldValue} for the matching field's
 * kind. A checkbox `false`, empty text `""`, and multi-select `[]` are KEPT (a
 * user may have explicitly cleared a value baked into the source). A literal
 * scalar radio / choice `""` is also kept as an explicit override of `/V`, but
 * any other invalid value that merely normalizes to `""` is dropped. The
 * overall map is capped to {@link ACROFORM_MAX_FIELDS} entries.
 *
 * @param map The raw values map (from the persisted annotation or live edits).
 * @param fields The known fields for the document. For radio groups, pass each
 *   member; members sharing a name are merged so any member's export value is
 *   accepted.
 */
export function normalizeFormValues(
  map: unknown,
  fields: ReadonlyArray<AcroField>,
): Record<string, AcroFormValue> {
  const out: Record<string, AcroFormValue> = {};
  if (!map || typeof map !== "object" || Array.isArray(map)) return out;

  const byName = buildFieldDirectory(fields);

  let count = 0;
  for (const [rawName, rawValue] of Object.entries(map as Record<string, unknown>)) {
    if (count >= ACROFORM_MAX_FIELDS) break;
    const name = normalizeFieldName(rawName);
    if (!name) continue;
    const field = byName.get(name);
    if (!field) continue;
    const normalized = normalizeFieldValue(field, rawValue);
    if (normalized === null) continue;
    // A literal empty string is a valid, explicit clear and must override a
    // non-empty source `/V`. Other malformed scalar values normalize to empty
    // too, but are dropped rather than being allowed to clear a field.
    if (
      (field.type === "radio" || field.type === "choice") &&
      normalized === "" &&
      rawValue !== ""
    ) {
      continue;
    }
    // A multi-state checkbox clears through "", "Off" or false; any other
    
    if (
      field.type === "checkbox" &&
      field.checkboxStates.length >= 2 &&
      normalized === "" &&
      rawValue !== "" &&
      rawValue !== "Off" &&
      rawValue !== "/Off" &&
      rawValue !== false
    ) {
      continue;
    }
    setFormValueEntry(out, name, normalized);
    count += 1;
  }
  return out;
}

/**
 * Convert a widget's PDF-space rect into page-wrapper percentages
 * `{x, y, width, height}` (top-left origin), reusing the SAME geometry as the
 * link overlays. The caller first maps the PDF-space `rect` into VIEWPORT pixels
 * with pdf.js (`viewport.convertToViewportRectangle(field.rect)`, which applies
 * rotation), then passes that viewport rect + the rendered page box dimensions
 * here. Returns `null` for degenerate inputs (so a bad widget paints nothing).
 *
 * This is a thin, intention-revealing wrapper over `links.ts`
 * `viewportRectToPercent` so form-field placement and link placement provably
 * share one rotation-aware definition.
 */
export function widgetRectToPercent(
  viewportRect: ReadonlyArray<number>,
  viewportWidth: number,
  viewportHeight: number,
): { x: number; y: number; width: number; height: number } | null {
  return viewportRectToPercent(viewportRect, viewportWidth, viewportHeight);
}

/**
 * Build the `{ [fieldName]: value }` payload to PUT for the whole form. Identical
 * to {@link normalizeFormValues} but named for the persist path so call sites
 * read clearly ("what we persist == what we render"). Kept separate so a future
 * persist-only transform (e.g. dropping defaults) has a home.
 */
export function buildFormValuesPayload(
  map: unknown,
  fields: ReadonlyArray<AcroField>,
): Record<string, AcroFormValue> {
  return normalizeFormValues(map, fields);
}
