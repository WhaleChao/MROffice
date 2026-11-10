/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A text field's Format, as Acrobat's Text Field Properties › Format and
 * ONLYOFFICE's special fields offer it: 数字, 百分比, 日期, 时间, 邮政编码,
 * 电话号码, 社会保障号, 信用卡, 电子邮件 and a custom mask. A PDF states a
 * field's format as the scripts in its /AA -- Format (/F), Keystroke (/K),
 * Validate (/V) -- calling Acrobat's built-in AF helpers with literal settings,
 * which every reader that runs form scripts knows; the editor writes exactly
 * those (and, for an email address, one fixed validation script), and the
 * server takes no other script (form_format_scripts.py, the same table).
 */

/** The date and time patterns offered (Acrobat's own lists). */
export const PDF_FIELD_DATE_PATTERNS = [
  "yyyy-mm-dd", "yyyy/mm/dd", "m/d/yyyy", "mm/dd/yyyy", "dd/mm/yyyy", "d-mmm-yyyy", "mmmm d, yyyy",
] as const;
export const PDF_FIELD_TIME_PATTERNS = ["HH:MM", "h:MM tt", "HH:MM:ss", "h:MM:ss tt"] as const;

export type PdfTextFieldFormat = Readonly<
  | { kind: "number"; decimals: number; separatorStyle: 0 | 1 | 2 | 3; currency: string; prependCurrency: boolean }
  | { kind: "percent"; decimals: number; separatorStyle: 0 | 1 | 2 | 3 }
  | { kind: "date"; pattern: (typeof PDF_FIELD_DATE_PATTERNS)[number] }
  | { kind: "time"; pattern: (typeof PDF_FIELD_TIME_PATTERNS)[number] }
  /** AFSpecial: 0 ZIP, 1 ZIP+4, 2 phone, 3 social security number. */
  | { kind: "special"; special: "zip" | "zip4" | "phone" | "ssn" }
  /** AFSpecial_KeystrokeEx: 9 digit, A letter, O letter or digit, X any. */
  | { kind: "mask"; mask: string }
  | { kind: "creditCard" }
  | { kind: "email" }
>;

/** 信用卡: sixteen digits in fours (ONLYOFFICE's Credit card field). */
export const PDF_CREDIT_CARD_MASK = "9999 9999 9999 9999";

/** 电子邮件: the one validation script the editor writes, word for word
 * (form_format_scripts.PDF_EMAIL_VALIDATE_SCRIPT). */
export const PDF_EMAIL_VALIDATE_SCRIPT =
  'if (event.value && !/^[^\\s@]+@[^\\s@]+\\.[^\\s@]+$/.test(event.value)) ' +
  '{ app.alert("Enter an email address."); event.rc = false; }';

const SPECIAL_INDEX = { zip: 0, zip4: 1, phone: 2, ssn: 3 } as const;
/** A mask or a pattern: the server's literal alphabet. */
const LITERAL = /^[A-Za-z0-9 :/\-.,()+#*]{1,64}$/u;

export function pdfTextFieldFormatValid(format: PdfTextFieldFormat): boolean {
  switch (format.kind) {
    case "number":
      return Number.isInteger(format.decimals) && format.decimals >= 0 && format.decimals <= 10 &&
        /^[^"\\\u0000-\u001f]{0,4}$/u.test(format.currency);
    case "percent":
      return Number.isInteger(format.decimals) && format.decimals >= 0 && format.decimals <= 10;
    case "date":
      return (PDF_FIELD_DATE_PATTERNS as readonly string[]).includes(format.pattern);
    case "time":
      return (PDF_FIELD_TIME_PATTERNS as readonly string[]).includes(format.pattern);
    case "mask":
      return LITERAL.test(format.mask) && /[9AOX]/u.test(format.mask);
    default:
      return true;
  }
}

/** The scripts the field's /AA holds for the format (Format, Keystroke,
 * Validate), as Acrobat's wizard writes them. */
export function pdfTextFieldFormatScripts(
  format: PdfTextFieldFormat,
): Readonly<{ format?: string; keystroke?: string; validate?: string }> {
  switch (format.kind) {
    case "number": {
      const args = `${format.decimals}, ${format.separatorStyle}, 0, 0, "${format.currency}", ${format.prependCurrency}`;
      return { format: `AFNumber_Format(${args});`, keystroke: `AFNumber_Keystroke(${args});` };
    }
    case "percent":
      return {
        format: `AFPercent_Format(${format.decimals}, ${format.separatorStyle});`,
        keystroke: `AFPercent_Keystroke(${format.decimals}, ${format.separatorStyle});`,
      };
    case "date":
      return { format: `AFDate_FormatEx("${format.pattern}");`, keystroke: `AFDate_KeystrokeEx("${format.pattern}");` };
    case "time":
      return { format: `AFTime_FormatEx("${format.pattern}");`, keystroke: `AFTime_KeystrokeEx("${format.pattern}");` };
    case "special":
      return {
        format: `AFSpecial_Format(${SPECIAL_INDEX[format.special]});`,
        keystroke: `AFSpecial_Keystroke(${SPECIAL_INDEX[format.special]});`,
      };
    case "mask":
      return { keystroke: `AFSpecial_KeystrokeEx("${format.mask}");` };
    case "creditCard":
      return { keystroke: `AFSpecial_KeystrokeEx("${PDF_CREDIT_CARD_MASK}");` };
    case "email":
      return { validate: PDF_EMAIL_VALIDATE_SCRIPT };
  }
}

/** A field's format read back from its scripts (the editor's own, or an
 * Acrobat form's of the same shape); null where it has none or another. */
export function pdfTextFieldFormatOf(
  scripts: Readonly<{ format?: string | null; keystroke?: string | null; validate?: string | null }>,
): PdfTextFieldFormat | null {
  const keystroke = scripts.keystroke?.trim() ?? "";
  const validate = scripts.validate?.trim() ?? "";
  if (validate === PDF_EMAIL_VALIDATE_SCRIPT && !keystroke) return { kind: "email" };
  let match = /^AFNumber_Keystroke\((\d{1,2}), ?([0-3]), ?[0-3], ?0, ?"([^"\\]{0,4})", ?(true|false)\);?$/u.exec(keystroke);
  if (match) {
    return { kind: "number", decimals: Number(match[1]), separatorStyle: Number(match[2]) as 0 | 1 | 2 | 3,
      currency: match[3]!, prependCurrency: match[4] === "true" };
  }
  match = /^AFPercent_Keystroke\((\d{1,2}), ?([0-3])\);?$/u.exec(keystroke);
  if (match) return { kind: "percent", decimals: Number(match[1]), separatorStyle: Number(match[2]) as 0 | 1 | 2 | 3 };
  match = /^AFDate_KeystrokeEx\("([^"\\]{1,64})"\);?$/u.exec(keystroke);
  if (match && (PDF_FIELD_DATE_PATTERNS as readonly string[]).includes(match[1]!)) {
    return { kind: "date", pattern: match[1] as PdfTextFieldFormat & { kind: "date" } extends { pattern: infer P } ? P : never };
  }
  match = /^AFTime_KeystrokeEx\("([^"\\]{1,64})"\);?$/u.exec(keystroke);
  if (match && (PDF_FIELD_TIME_PATTERNS as readonly string[]).includes(match[1]!)) {
    return { kind: "time", pattern: match[1] as PdfTextFieldFormat & { kind: "time" } extends { pattern: infer P } ? P : never };
  }
  match = /^AFSpecial_Keystroke\(([0-3])\);?$/u.exec(keystroke);
  if (match) return { kind: "special", special: (["zip", "zip4", "phone", "ssn"] as const)[Number(match[1])]! };
  match = /^AFSpecial_KeystrokeEx\("([^"\\]{1,64})"\);?$/u.exec(keystroke);
  if (match) return match[1] === PDF_CREDIT_CARD_MASK ? { kind: "creditCard" } : { kind: "mask", mask: match[1]! };
  return null;
}

/** The Format menu's entries, in Acrobat's order and ONLYOFFICE's special
 * fields after them. */
export const PDF_FORM_FIELD_FORMAT_CHOICES = [
  "none", "number", "percent", "date", "time", "zip", "zip4", "phone", "ssn", "creditCard", "email", "mask",
] as const;
export type PdfFormFieldFormatChoice = (typeof PDF_FORM_FIELD_FORMAT_CHOICES)[number];

/** The menu entry a format is. */
export function pdfFormFieldFormatChoice(format: PdfTextFieldFormat | null): PdfFormFieldFormatChoice {
  if (!format) return "none";
  return format.kind === "special" ? format.special : format.kind;
}

/** A menu entry's format, with Acrobat's defaults (two decimals, 1,234.56;
 * ISO date; 24-hour time). */
export function pdfFormFieldFormatFor(choice: PdfFormFieldFormatChoice): PdfTextFieldFormat | null {
  switch (choice) {
    case "none":
      return null;
    case "number":
      return { kind: "number", decimals: 2, separatorStyle: 0, currency: "", prependCurrency: true };
    case "percent":
      return { kind: "percent", decimals: 2, separatorStyle: 0 };
    case "date":
      return { kind: "date", pattern: "yyyy-mm-dd" };
    case "time":
      return { kind: "time", pattern: "HH:MM" };
    case "zip":
    case "zip4":
    case "phone":
    case "ssn":
      return { kind: "special", special: choice };
    case "creditCard":
      return { kind: "creditCard" };
    case "email":
      return { kind: "email" };
    case "mask":
      return { kind: "mask", mask: "999-999" };
  }
}

/** A field's Format, Keystroke and Validate scripts (undefined: none of that
 * trigger; null: one that is no string) read as the format the editor writes:
 * the format, null for none, or "other" for scripts of another kind or more
 * than it writes (which the editor never overwrites). */
export function pdfTextFieldFormatFromScripts(
  scripts: Readonly<{ format?: string | null; keystroke?: string | null; validate?: string | null }>,
): PdfTextFieldFormat | null | "other" {
  if (scripts.format === undefined && scripts.keystroke === undefined && scripts.validate === undefined) return null;
  if (scripts.format === null || scripts.keystroke === null || scripts.validate === null) return "other";
  const format = pdfTextFieldFormatOf(scripts);
  if (!format) return "other";
  const written = pdfTextFieldFormatScripts(format);
  const same = (read: string | undefined, write: string | undefined) => (read?.trim() || undefined) === write;
  return same(scripts.format, written.format) && same(scripts.keystroke, written.keystroke) &&
    same(scripts.validate, written.validate) ? format : "other";
}

/** A field's format scripts as its /AA triggers name them: F Format,
 *  K Keystroke, V Validate (the native authoring spec's `formatScripts`). */
export type PdfTextFieldFormatTriggers = Readonly<{ F?: string; K?: string; V?: string }>;

/** The /AA scripts a format writes, keyed by trigger; none for no format. */
export function pdfTextFieldFormatTriggers(format: PdfTextFieldFormat | null | undefined): PdfTextFieldFormatTriggers {
  if (!format) return {};
  const scripts = pdfTextFieldFormatScripts(format);
  return {
    ...(scripts.format ? { F: scripts.format } : {}),
    ...(scripts.keystroke ? { K: scripts.keystroke } : {}),
    ...(scripts.validate ? { V: scripts.validate } : {}),
  };
}

/** Whether `value` is the trigger scripts of exactly one format (or none). */
export function isPdfTextFieldFormatTriggers(value: unknown): value is PdfTextFieldFormatTriggers {
  if (!value || typeof value !== "object" || Array.isArray(value)) return false;
  const entries = Object.entries(value as Record<string, unknown>);
  if (entries.some(([key, script]) => (key !== "F" && key !== "K" && key !== "V") || typeof script !== "string")) {
    return false;
  }
  if (entries.length === 0) return true;
  const triggers = value as PdfTextFieldFormatTriggers;
  const format = pdfTextFieldFormatFromScripts({ format: triggers.F, keystroke: triggers.K, validate: triggers.V });
  return format !== null && format !== "other" && pdfTextFieldFormatValid(format);
}

/** 图片字段 (Acrobat's Image Field, ONLYOFFICE's picture form): the push
 * button's Mouse Up asks the reader to place a picture on it. Acrobat writes
 * this call and nothing else (form_format_scripts.PDF_IMAGE_FIELD_SCRIPT). */
export const PDF_IMAGE_FIELD_SCRIPT = "event.target.buttonImportIcon();";

/** Whether pdf.js's `actions` map of a push button is the image field's one
 * call (its `/A`, which pdf.js reports as "Action"), and nothing else. */
export function isPdfImageFieldActions(actions: unknown): boolean {
  if (!actions || typeof actions !== "object") return false;
  const entries = Object.entries(actions as Record<string, unknown>);
  if (entries.length !== 1 || entries[0][0] !== "Action") return false;
  const scripts = entries[0][1];
  return Array.isArray(scripts) && scripts.length === 1 &&
    typeof scripts[0] === "string" && scripts[0].trim() === PDF_IMAGE_FIELD_SCRIPT;
}
