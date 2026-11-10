/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Form field authoring (WPS PDF ▸ 表单).
 *
 * Filling an existing AcroForm already works; this is the other half — CREATING
 * fields. The rules that matter are the ones a PDF reader enforces later, so
 * they are checked before a field is written rather than discovered as a broken
 * form: names are unique leaves in a valid dotted hierarchy, a radio group
 * shares one name but needs distinct export values, a
 * checkbox needs an export value that is not `Off` (which PDF reserves for
 * "unchecked"), and a choice field needs options.
 *
 * Pure: a proposed field plus the fields already on the form, in; either the
 * normalized field or the reason it was refused, out.
 */

import { readPdfFormRole, type PdfFormRole } from "./form-roles";
import {
  isPdfTextFieldFormatTriggers,
  pdfTextFieldFormatValid,
  type PdfTextFieldFormat,
  type PdfTextFieldFormatTriggers,
} from "@mrnote/core/web/lib/pdf/form-field-format.ts";
import {
  ACROFORM_MAX_OPTIONS,
  ACROFORM_MAX_TEXT,
  ACROFORM_MAX_TEXT_FONT_SIZE,
  ACROFORM_MAX_WIDGET_BORDER_WIDTH,
  normalizePdfTextFieldColor,
  normalizePdfTextFieldPaint,
  isPdfNavigationButtonAction,
  isPdfViewCommandAction,
  type PdfTextFieldAlignment,
  type PdfTextFieldVisibility,
  type PdfNavigationButtonAction,
  type PdfViewCommandAction,
  type AcroFormValue,
} from "@mrnote/core/web/lib/pdf/acroform.ts";
import { PDF_LINK_OVERLAY_MAX_URL, sanitizeLinkUrl } from "@mrnote/core/web/lib/pdf/links.ts";

export type FormFieldKind =
  | "text"
  | "checkbox"
  | "radio"
  | "dropdown"
  | "listbox"
  | "button"
  | "signature"
  | "resetButton"
  | "printButton"
  | "navigationButton"
  | "viewButton"
  | "saveButton"
  | "goToPageButton"
  | "webLinkButton"
  | "imageField";

export interface FormFieldDraft {
  kind: FormFieldKind;
  name: string;
  /** Native AcroForm /TU: a human-facing alternate name and tooltip. */
  alternateName?: string;
  /** Who fills the field (form-roles.ts); absent: anyone. */
  role?: PdfFormRole;
  /** Radio buttons and checkboxes: the value stored when this one is chosen. */
  exportValue?: string;
  /** Choice fields: the options offered. */
  options?: readonly string[];
  defaultValue?: string;
  required?: boolean;
  readOnly?: boolean;
  /** Text fields only. 0 means unlimited, as PDF does. */
  maxLength?: number;
  /** Text fields only. Writes the native AcroForm multiline flag. */
  multiline?: boolean;
  /** Text fields only. Masks entry; the stored PDF value is not encrypted. */
  password?: boolean;
  /** Text fields only. Divides the field into maxLength character cells. */
  comb?: boolean;
  /** Text fields only. Defaults to true; false writes DoNotSpellCheck. */
  spellCheck?: boolean;
  /** Text fields only. Defaults to true; false writes DoNotScroll. */
  scrollable?: boolean;
  /** Text fields only. Writes native AcroForm `/Q`; defaults to left. */
  alignment?: PdfTextFieldAlignment;
  /** Text fields only. Native `/DA` size in points; 0 means automatic. */
  fontSize?: number;
  /** Text fields only. Native `/DA` foreground colour as canonical hex. */
  textColor?: string;
  /** Text fields only. Native Widget `/MK /BG` fill as canonical hex. */
  backgroundColor?: string;
  /** Text fields only. Native Widget `/MK /BC` border as canonical hex. */
  borderColor?: string;
  /** Text fields only. Native Widget `/BS /W` border width in points. */
  borderWidth?: number;
  /** Text fields only. Native Widget screen/print visibility contract. */
  visibility?: PdfTextFieldVisibility;
  /** One-based manual widget order on its page. */
  tabOrder?: number;
  /** Text fields only. Acrobat's Format (数字, 日期, 电话号码, 电子邮件...),
   * written as the standard scripts in the field's /AA (form-field-format.ts). */
  format?: PdfTextFieldFormat;
  /** Text fields only. The field's Format / Keystroke / Validate are scripts
   * of another kind: shown, never rewritten. */
  formatLocked?: boolean;
}

export interface ExistingFormField {
  kind: FormFieldKind;
  name: string;
  exportValue?: string;
}

export type FormFieldProblem =
  | "empty-name"
  | "name-too-long"
  | "invalid-name-path"
  | "hierarchy-name-conflict"
  | "alternate-name-too-long"
  | "invalid-alternate-name"
  | "duplicate-name"
  | "kind-conflict"
  | "missing-export-value"
  | "reserved-export-value"
  | "duplicate-export-value"
  | "missing-options"
  | "default-not-an-option"
  | "negative-max-length"
  | "incompatible-text-flags"
  | "comb-requires-max-length"
  | "password-default-value"
  | "invalid-text-alignment"
  | "invalid-text-font-size"
  | "invalid-text-color"
  | "invalid-text-background-color"
  | "invalid-text-border-color"
  | "invalid-text-format"
  | "invalid-text-border-width"
  | "invalid-text-visibility"
  | "invalid-tab-order";

/** PDF's practical limit for a field name. */
export const MAX_FIELD_NAME_LENGTH = 127;
/** Bounded native /TU value used for tooltips and accessible field names. */
export const MAX_FIELD_ALTERNATE_NAME_LENGTH = 512;
/** PDF reserves this export value for "unchecked". */
export const RESERVED_OFF_VALUE = "Off";
const FIELD_NAME_CONTROL_CHARACTER = /[\u0000-\u001f\u007f]/;

const CHOICE_KINDS: ReadonlySet<FormFieldKind> = new Set([
  "dropdown",
  "listbox",
]);
const TOGGLE_KINDS: ReadonlySet<FormFieldKind> = new Set(["checkbox", "radio"]);

export const NATIVE_PDF_FORM_AUTHORING_SCHEMA_VERSION =
  "mrai.pdf-native-form-authoring.v1" as const;
export const NATIVE_PDF_FORM_GROUP_PROPERTIES_SCHEMA_VERSION =
  "mrai.pdf-native-form-group-properties.v1" as const;
export const NATIVE_PDF_RADIO_GROUP_OPTIONS_SCHEMA_VERSION =
  "mrai.pdf-native-radio-group-options.v1" as const;
export const NATIVE_PDF_FORM_VALUE_SNAPSHOT_SCHEMA_VERSION =
  "mrai.pdf-native-form-value-snapshot.v1" as const;
export const NATIVE_PDF_FORM_WIDGET_GEOMETRY_SCHEMA_VERSION =
  "mrai.pdf-native-form-widget-geometry.v1" as const;
export const NATIVE_PDF_FORM_WIDGET_APPEARANCE_SCHEMA_VERSION =
  "mrai.pdf-native-form-widget-appearance.v1" as const;
/** 图片字段: the picture placed on an image field (native_edit_primitives). */
export const NATIVE_PDF_FORM_IMAGE_SCHEMA_VERSION =
  "mrai.pdf-native-form-image.v1" as const;
/** The picture as sent: the window shrinks it to this before it goes. */
export const NATIVE_PDF_FORM_IMAGE_MAX_BYTES = 3 * 1024 * 1024;
export const MAX_NATIVE_PDF_FORM_AUTHORING_OPERATIONS = 128;
export const MAX_NATIVE_PDF_FORM_TAB_ORDER = 4096;

export type NativePdfFormFieldKind =
  | "text"
  | "checkbox"
  | "radio"
  | "dropdown"
  | "listbox"
  | "signature"
  | "resetButton"
  | "printButton"
  | "navigationButton"
  | "viewButton"
  | "saveButton"
  | "goToPageButton"
  | "webLinkButton"
  | "imageField";

type NativePdfFormGroupPropertyKind = Exclude<
  NativePdfFormFieldKind,
  | "imageField"
  | "signature"
  | "resetButton"
  | "printButton"
  | "navigationButton"
  | "viewButton"
  | "saveButton"
  | "goToPageButton"
  | "webLinkButton"
>;

export interface NativePdfFormFieldSpec {
  readonly kind: NativePdfFormFieldKind;
  readonly name: string;
  readonly alternateName?: string;
  /** Who fills the field (form-roles.ts): /MRNoteRole on the field. */
  readonly role?: PdfFormRole;
  readonly page: number;
  readonly rect: Readonly<{
    x: number;
    y: number;
    width: number;
    height: number;
  }>;
  readonly tabOrder: number;
  readonly required: boolean;
  readonly readOnly: boolean;
  readonly defaultValue?: string;
  readonly exportValue?: string;
  readonly options?: readonly string[];
  readonly selectedValues?: readonly string[];
  readonly checked?: boolean;
  readonly editable?: boolean;
  readonly multiselect?: boolean;
  readonly offToggleable?: boolean;
  readonly maxLength?: number;
  readonly multiline?: boolean;
  readonly password?: boolean;
  readonly comb?: boolean;
  readonly spellCheck?: boolean;
  readonly scrollable?: boolean;
  readonly alignment?: PdfTextFieldAlignment;
  readonly fontSize?: number;
  readonly textColor?: string;
  readonly backgroundColor?: string;
  readonly borderColor?: string;
  readonly borderWidth?: number;
  readonly visibility?: PdfTextFieldVisibility;
  /** Text fields: the Format's /AA scripts (none clears it). Absent on an
   *  edit, the field's scripts stay as they are. */
  readonly formatScripts?: PdfTextFieldFormatTriggers;
  readonly buttonLabel?: string;
  readonly navigationAction?: PdfNavigationButtonAction;
  readonly viewAction?: PdfViewCommandAction;
  readonly buttonUrl?: string;
}

export interface NativePdfFormFieldExpectedIdentity {
  readonly name: string;
  readonly kind: NativePdfFormFieldKind;
  readonly page: number;
}

export interface NativePdfFormWidgetExpectedIdentity {
  readonly objectNumber: number;
  readonly generation: number;
  readonly page: number;
}

export interface NativePdfFormFieldGroupExpectedIdentity {
  readonly name: string;
  readonly kind: NativePdfFormGroupPropertyKind;
  readonly widgets: readonly NativePdfFormWidgetExpectedIdentity[];
}

export type NativePdfFormWidgetGeometryKind =
  | "text"
  | "checkbox"
  | "radio"
  | "dropdown"
  | "listbox"
  | "signature"
  | "pushbutton";

export interface NativePdfFormWidgetGeometryExpectedIdentity {
  readonly name: string;
  readonly kind: NativePdfFormWidgetGeometryKind;
  readonly widget: NativePdfFormWidgetExpectedIdentity;
  readonly rect: readonly [number, number, number, number];
}

export interface NativePdfFormWidgetAppearanceStyle {
  readonly backgroundColor: string;
  readonly borderColor: string;
  readonly borderWidth: number;
}

export interface NativePdfFormWidgetAppearanceExpectedIdentity
  extends NativePdfFormWidgetAppearanceStyle {
  readonly name: string;
  readonly kind: "checkbox" | "radio";
  readonly widget: NativePdfFormWidgetExpectedIdentity;
  readonly rect: readonly [number, number, number, number];
}

export interface NativePdfRadioGroupOptionMutation
  extends NativePdfFormWidgetExpectedIdentity {
  readonly originalValue: string;
  readonly replacementValue: string;
}

/** Parse pdf.js' stable indirect-reference id (`field-12R` / `12R3`). */
export function nativePdfFormWidgetExpectedIdentity(
  id: string,
  page: number,
): NativePdfFormWidgetExpectedIdentity | null {
  const match = /^(?:field-)?(?<object>[1-9]\d*)R(?<generation>\d*)$/.exec(
    id,
  );
  if (!match?.groups || !Number.isSafeInteger(page) || page < 1) return null;
  const objectNumber = Number(match.groups.object);
  const generation = match.groups.generation
    ? Number(match.groups.generation)
    : 0;
  return Number.isSafeInteger(objectNumber) &&
    objectNumber > 0 &&
    Number.isSafeInteger(generation) &&
    generation >= 0
    ? { objectNumber, generation, page }
    : null;
}

/** A check box's or radio button's on-state name: any text a PDF name holds
 * (no NUL; up to the 128 characters the authoring dialog takes, 512 bytes in
 * UTF-8), Off aside; spaces, slashes and # are escaped where the name is
 * written (WPS and Acrobat take them; the server's
 * _native_form_export_value_ok, user audit C0011). */
function nativePdfFormExportValueBytes(value: string): number {
  return new TextEncoder().encode(value).length;
}

export function isNativePdfFormExportValue(value: unknown): value is string {
  return (
    typeof value === "string" &&
    value !== RESERVED_OFF_VALUE &&
    !value.includes("\u0000") &&
    nativePdfFormExportValueBytes(value) >= 1 &&
    nativePdfFormExportValueBytes(value) <= 512
  );
}

export type NativePdfFormAuthoringOperation =
  | Readonly<{
      schemaVersion: typeof NATIVE_PDF_FORM_WIDGET_APPEARANCE_SCHEMA_VERSION;
      id: string;
      operation: "edit-widget-appearance";
      expected: NativePdfFormWidgetAppearanceExpectedIdentity;
      target: NativePdfFormWidgetAppearanceStyle;
    }>
  | Readonly<{
      schemaVersion: typeof NATIVE_PDF_FORM_WIDGET_GEOMETRY_SCHEMA_VERSION;
      id: string;
      operation: "edit-widget-geometry";
      expected: NativePdfFormWidgetGeometryExpectedIdentity;
      targetRect: readonly [number, number, number, number];
    }>
  | Readonly<{
      schemaVersion: typeof NATIVE_PDF_FORM_VALUE_SNAPSHOT_SCHEMA_VERSION;
      id: string;
      operation: "snapshot-value";
      expected:
        | NativePdfFormFieldExpectedIdentity
        | NativePdfFormFieldGroupExpectedIdentity;
      value: AcroFormValue;
    }>
  | Readonly<{
      schemaVersion: typeof NATIVE_PDF_FORM_IMAGE_SCHEMA_VERSION;
      id: string;
      operation: "fill-image";
      expected: NativePdfFormFieldExpectedIdentity;
      /** PNG or JPEG, base64; null takes the picture off. */
      image: Readonly<{ mime: "image/png" | "image/jpeg"; data: string }> | null;
    }>
  | Readonly<{
      schemaVersion: typeof NATIVE_PDF_FORM_AUTHORING_SCHEMA_VERSION;
      id: string;
      operation: "create";
      field: NativePdfFormFieldSpec;
    }>
  | Readonly<{
      schemaVersion: typeof NATIVE_PDF_FORM_AUTHORING_SCHEMA_VERSION;
      id: string;
      operation: "edit";
      expected: NativePdfFormFieldExpectedIdentity;
      field: NativePdfFormFieldSpec;
    }>
  | Readonly<{
      schemaVersion: typeof NATIVE_PDF_FORM_AUTHORING_SCHEMA_VERSION;
      id: string;
      operation: "delete";
      expected: NativePdfFormFieldExpectedIdentity;
    }>
  | Readonly<{
      schemaVersion: typeof NATIVE_PDF_FORM_GROUP_PROPERTIES_SCHEMA_VERSION;
      id: string;
      operation: "edit-group-properties";
      expected: NativePdfFormFieldGroupExpectedIdentity;
      required: boolean;
      readOnly: boolean;
    }>
  | Readonly<{
      schemaVersion: typeof NATIVE_PDF_RADIO_GROUP_OPTIONS_SCHEMA_VERSION;
      id: string;
      operation: "edit-radio-options";
      name: string;
      widgets: readonly NativePdfRadioGroupOptionMutation[];
      selectedValue: string | null;
      offToggleable: boolean;
    }>;

export interface NativePdfFormAuthoringMutation {
  readonly generation: string;
  readonly gestureId: string;
  readonly operation: NativePdfFormAuthoringOperation;
}

/**
 * Value-setting authoring commands own their final native /V. Remove only that
 * field's older fill override, or the later fill phase could undo the command.
 * Metadata-only edits deliberately retain their independently filled values.
 */
export function reconcilePdfFormAuthoringValues<T extends { id: string }>(
  fields: readonly T[],
  operation: NativePdfFormAuthoringOperation,
): T[] {
  let name: string | undefined;
  if (operation.operation === "snapshot-value") {
    name = operation.expected.name;
  } else if (operation.operation === "edit-radio-options") {
    name = operation.name;
  } else if (operation.operation === "create") {
    name = operation.field.name;
  } else if (operation.operation === "delete") {
    name = operation.expected.name;
  } else if (
    operation.operation === "edit" &&
    (operation.field.kind === "checkbox" ||
      operation.field.kind === "radio" ||
      operation.field.kind === "listbox")
  ) {
    name = operation.field.name;
  }
  return fields.filter((field) => field.id !== name);
}

/** Put a filled value into the ordered history before metadata can lock it or
 * change its option set. The native snapshot uses the existing value writer;
 * only the obsolete later fill override is removed from the page projection. */
export function planPdfFormAuthoringTransition<
  T extends { id: string; value: AcroFormValue },
>(
  fields: readonly T[],
  operation: NativePdfFormAuthoringOperation,
  snapshotId: string,
): { operations: NativePdfFormAuthoringOperation[]; fields: T[] } {
  const metadataOnly =
    (operation.operation === "edit-group-properties" && operation.readOnly) ||
    (operation.operation === "edit" &&
      (operation.field.kind === "dropdown" ||
        (operation.field.kind === "text" && operation.field.readOnly)));
  if (metadataOnly) {
    const field = fields.find((item) => item.id === operation.expected.name);
    if (field) {
      const expected =
        "widgets" in operation.expected
          ? {
              ...operation.expected,
              widgets: operation.expected.widgets.map((widget) => ({
                ...widget,
              })),
            }
          : { ...operation.expected };
      return {
        operations: [
          {
            schemaVersion: NATIVE_PDF_FORM_VALUE_SNAPSHOT_SCHEMA_VERSION,
            id: snapshotId,
            operation: "snapshot-value",
            expected,
            value: Array.isArray(field.value) ? [...field.value] : field.value,
          },
          operation,
        ],
        fields: fields.filter((item) => item.id !== field.id),
      };
    }
  }
  return {
    operations: [operation],
    fields: reconcilePdfFormAuthoringValues(fields, operation),
  };
}

export function isNativePdfFormFieldKind(
  value: FormFieldKind | string,
): value is NativePdfFormFieldKind {
  return (
    value === "text" ||
    value === "checkbox" ||
    value === "radio" ||
    value === "dropdown" ||
    value === "listbox" ||
    value === "signature" ||
    value === "resetButton" ||
    value === "printButton" ||
    value === "navigationButton" ||
    value === "viewButton" ||
    value === "saveButton" ||
    value === "goToPageButton" ||
    value === "webLinkButton" ||
    value === "imageField"
  );
}

function readNativePdfFormFieldSpec(
  value: unknown,
): NativePdfFormFieldSpec | null {
  if (!value || typeof value !== "object") return null;
  const candidate = value as Record<string, unknown>;
  const kind = candidate.kind;
  const rect = candidate.rect;
  if (
    typeof kind !== "string" ||
    !isNativePdfFormFieldKind(kind) ||
    !rect ||
    typeof rect !== "object"
  ) {
    return null;
  }
  const box = rect as Record<string, unknown>;
  const x = box.x;
  const y = box.y;
  const width = box.width;
  const height = box.height;
  const page = candidate.page;
  const tabOrder = candidate.tabOrder;
  if (
    typeof x !== "number" ||
    typeof y !== "number" ||
    typeof width !== "number" ||
    typeof height !== "number" ||
    ![x, y, width, height].every(Number.isFinite) ||
    x < 0 ||
    y < 0 ||
    width <= 0 ||
    height <= 0 ||
    x + width > 1 ||
    y + height > 1 ||
    !Number.isSafeInteger(page) ||
    (page as number) < 1 ||
    !Number.isSafeInteger(tabOrder) ||
    (tabOrder as number) < 1 ||
    (tabOrder as number) > MAX_NATIVE_PDF_FORM_TAB_ORDER
  ) {
    return null;
  }
  const draft = normalizeFormField(candidate as unknown as FormFieldDraft);
  if (!draft || draft.kind !== kind) return null;
  if (candidate.role !== undefined && !readPdfFormRole(candidate.role)) return null;
  if (
    (kind === "signature" ||
      kind === "imageField" ||
      kind === "resetButton" ||
      kind === "printButton" ||
      kind === "navigationButton" ||
      kind === "viewButton" ||
      kind === "saveButton" ||
      kind === "goToPageButton" ||
      kind === "webLinkButton") &&
    candidate.defaultValue !== undefined
  ) {
    return null;
  }
  if (
    (kind === "resetButton" ||
      kind === "printButton" ||
      kind === "navigationButton" ||
      kind === "viewButton" ||
      kind === "saveButton" ||
      kind === "goToPageButton" ||
      kind === "webLinkButton") &&
    (typeof candidate.buttonLabel !== "string" ||
      !candidate.buttonLabel.trim() ||
      candidate.buttonLabel.trim() !== candidate.buttonLabel ||
      candidate.buttonLabel.length > 512)
  ) {
    return null;
  }
  if (
    kind === "navigationButton" &&
    !isPdfNavigationButtonAction(candidate.navigationAction)
  ) {
    return null;
  }
  if (kind === "viewButton" && !isPdfViewCommandAction(candidate.viewAction)) {
    return null;
  }
  if (
    candidate.formatScripts !== undefined &&
    (kind !== "text" || !isPdfTextFieldFormatTriggers(candidate.formatScripts))
  ) {
    return null;
  }
  if (kind === "webLinkButton") {
    const sanitized = sanitizeLinkUrl(candidate.buttonUrl);
    if (
      !sanitized ||
      sanitized !== candidate.buttonUrl ||
      sanitized.length > PDF_LINK_OVERLAY_MAX_URL
    ) {
      return null;
    }
  }
  const options = candidate.options;
  const selectedValues = candidate.selectedValues;
  if (
    (options !== undefined &&
      (!Array.isArray(options) ||
        options.length > 200 ||
        options.some(
          (option) =>
            typeof option !== "string" || !option || option.length > 256,
        ))) ||
    (selectedValues !== undefined &&
      (!Array.isArray(selectedValues) ||
        selectedValues.length > 200 ||
        selectedValues.some((item) => typeof item !== "string")))
  ) {
    return null;
  }
  return JSON.parse(JSON.stringify(candidate)) as NativePdfFormFieldSpec;
}

function readNativePdfExpectedIdentity(
  value: unknown,
): NativePdfFormFieldExpectedIdentity | null {
  if (!value || typeof value !== "object") return null;
  const candidate = value as Record<string, unknown>;
  return typeof candidate.name === "string" &&
    candidate.name.trim() === candidate.name &&
    isValidFormFieldNamePath(candidate.name) &&
    typeof candidate.kind === "string" &&
    isNativePdfFormFieldKind(candidate.kind) &&
    Number.isSafeInteger(candidate.page) &&
    (candidate.page as number) >= 1
    ? {
        name: candidate.name,
        kind: candidate.kind,
        page: candidate.page as number,
      }
    : null;
}

function readNativePdfGroupExpectedIdentity(
  value: unknown,
): NativePdfFormFieldGroupExpectedIdentity | null {
  if (!value || typeof value !== "object") return null;
  const candidate = value as Record<string, unknown>;
  if (
    typeof candidate.name !== "string" ||
    candidate.name.trim() !== candidate.name ||
    !isValidFormFieldNamePath(candidate.name) ||
    typeof candidate.kind !== "string" ||
    !isNativePdfFormFieldKind(candidate.kind) ||
    candidate.kind === "signature" ||
    candidate.kind === "imageField" ||
    candidate.kind === "resetButton" ||
    candidate.kind === "printButton" ||
    candidate.kind === "navigationButton" ||
    candidate.kind === "viewButton" ||
    candidate.kind === "saveButton" ||
    candidate.kind === "goToPageButton" ||
    candidate.kind === "webLinkButton" ||
    !Array.isArray(candidate.widgets) ||
    candidate.widgets.length < 1 ||
    candidate.widgets.length > 200
  ) {
    return null;
  }
  const seen = new Set<string>();
  const widgets: NativePdfFormWidgetExpectedIdentity[] = [];
  for (const raw of candidate.widgets) {
    if (!raw || typeof raw !== "object") return null;
    const widget = raw as Record<string, unknown>;
    if (
      !Number.isSafeInteger(widget.objectNumber) ||
      Number(widget.objectNumber) < 1 ||
      !Number.isSafeInteger(widget.generation) ||
      Number(widget.generation) < 0 ||
      !Number.isSafeInteger(widget.page) ||
      Number(widget.page) < 1
    ) {
      return null;
    }
    const key = `${widget.objectNumber}:${widget.generation}:${widget.page}`;
    if (seen.has(key)) return null;
    seen.add(key);
    widgets.push({
      objectNumber: Number(widget.objectNumber),
      generation: Number(widget.generation),
      page: Number(widget.page),
    });
  }
  widgets.sort(
    (left, right) =>
      left.objectNumber - right.objectNumber ||
      left.generation - right.generation ||
      left.page - right.page,
  );
  return {
    name: candidate.name,
    kind: candidate.kind as NativePdfFormGroupPropertyKind,
    widgets,
  };
}

function readNativePdfFormWidgetGeometryRect(
  value: unknown,
): readonly [number, number, number, number] | null {
  if (
    !Array.isArray(value) ||
    value.length !== 4 ||
    value.some(
      (item) =>
        typeof item !== "number" ||
        !Number.isFinite(item) ||
        Math.abs(item) > 10_000_000,
    ) ||
    value[2] <= value[0] ||
    value[3] <= value[1]
  ) {
    return null;
  }
  return [value[0], value[1], value[2], value[3]];
}

function readNativePdfFormWidgetGeometryExpected(
  value: unknown,
): NativePdfFormWidgetGeometryExpectedIdentity | null {
  if (!value || typeof value !== "object") return null;
  const candidate = value as Record<string, unknown>;
  const kinds = new Set<NativePdfFormWidgetGeometryKind>([
    "text",
    "checkbox",
    "radio",
    "dropdown",
    "listbox",
    "signature",
    "pushbutton",
  ]);
  const widget = candidate.widget;
  const identity =
    widget && typeof widget === "object"
      ? (widget as Record<string, unknown>)
      : null;
  const rect = readNativePdfFormWidgetGeometryRect(candidate.rect);
  if (
    Object.keys(candidate).sort().join(",") !== "kind,name,rect,widget" ||
    typeof candidate.name !== "string" ||
    !isValidFormFieldNamePath(candidate.name) ||
    typeof candidate.kind !== "string" ||
    !kinds.has(candidate.kind as NativePdfFormWidgetGeometryKind) ||
    !identity ||
    Object.keys(identity).sort().join(",") !==
      "generation,objectNumber,page" ||
    !Number.isSafeInteger(identity.objectNumber) ||
    Number(identity.objectNumber) < 1 ||
    !Number.isSafeInteger(identity.generation) ||
    Number(identity.generation) < 0 ||
    !Number.isSafeInteger(identity.page) ||
    Number(identity.page) < 1 ||
    !rect
  ) {
    return null;
  }
  return {
    name: candidate.name,
    kind: candidate.kind as NativePdfFormWidgetGeometryKind,
    widget: {
      objectNumber: Number(identity.objectNumber),
      generation: Number(identity.generation),
      page: Number(identity.page),
    },
    rect,
  };
}

function readNativePdfFormWidgetAppearanceStyle(
  value: unknown,
): NativePdfFormWidgetAppearanceStyle | null {
  if (!value || typeof value !== "object") return null;
  const candidate = value as Record<string, unknown>;
  const backgroundColor = normalizePdfTextFieldColor(
    candidate.backgroundColor,
  );
  const borderColor = normalizePdfTextFieldColor(candidate.borderColor);
  const borderWidth = candidate.borderWidth;
  return Object.keys(candidate).sort().join(",") ===
    "backgroundColor,borderColor,borderWidth" &&
    backgroundColor !== null &&
    borderColor !== null &&
    typeof borderWidth === "number" &&
    Number.isFinite(borderWidth) &&
    borderWidth >= 0.1 &&
    borderWidth <= ACROFORM_MAX_WIDGET_BORDER_WIDTH
    ? { backgroundColor, borderColor, borderWidth }
    : null;
}

function readNativePdfFormWidgetAppearanceExpected(
  value: unknown,
): NativePdfFormWidgetAppearanceExpectedIdentity | null {
  if (!value || typeof value !== "object") return null;
  const candidate = value as Record<string, unknown>;
  const geometry = readNativePdfFormWidgetGeometryExpected({
    name: candidate.name,
    kind: candidate.kind,
    widget: candidate.widget,
    rect: candidate.rect,
  });
  const style = readNativePdfFormWidgetAppearanceStyle({
    backgroundColor: candidate.backgroundColor,
    borderColor: candidate.borderColor,
    borderWidth: candidate.borderWidth,
  });
  if (
    Object.keys(candidate).sort().join(",") !==
      "backgroundColor,borderColor,borderWidth,kind,name,rect,widget" ||
    !geometry ||
    (geometry.kind !== "checkbox" && geometry.kind !== "radio") ||
    !style
  ) {
    return null;
  }
  return { ...geometry, kind: geometry.kind, ...style };
}

export function readNativePdfFormAuthoringOperations(
  value: unknown,
): readonly NativePdfFormAuthoringOperation[] | null {
  if (value === undefined) return [];
  if (
    !Array.isArray(value) ||
    value.length > MAX_NATIVE_PDF_FORM_AUTHORING_OPERATIONS
  ) {
    return null;
  }
  const ids = new Set<string>();
  const operations: NativePdfFormAuthoringOperation[] = [];
  for (const raw of value) {
    if (!raw || typeof raw !== "object") return null;
    const candidate = raw as Record<string, unknown>;
    const id = candidate.id;
    const operation = candidate.operation;
    if (
      typeof id !== "string" ||
      !/^[A-Za-z0-9._:-]{1,256}$/.test(id) ||
      ids.has(id)
    ) {
      return null;
    }
    ids.add(id);
    if (operation === "edit-widget-appearance") {
      if (
        candidate.schemaVersion !==
          NATIVE_PDF_FORM_WIDGET_APPEARANCE_SCHEMA_VERSION ||
        Object.keys(candidate).sort().join(",") !==
          "expected,id,operation,schemaVersion,target"
      ) {
        return null;
      }
      const expected = readNativePdfFormWidgetAppearanceExpected(
        candidate.expected,
      );
      const target = readNativePdfFormWidgetAppearanceStyle(candidate.target);
      if (!expected || !target) return null;
      operations.push({
        schemaVersion: NATIVE_PDF_FORM_WIDGET_APPEARANCE_SCHEMA_VERSION,
        id,
        operation,
        expected,
        target,
      });
      continue;
    }
    if (operation === "edit-widget-geometry") {
      if (
        candidate.schemaVersion !==
          NATIVE_PDF_FORM_WIDGET_GEOMETRY_SCHEMA_VERSION ||
        Object.keys(candidate).sort().join(",") !==
          "expected,id,operation,schemaVersion,targetRect"
      ) {
        return null;
      }
      const expected = readNativePdfFormWidgetGeometryExpected(
        candidate.expected,
      );
      const targetRect = readNativePdfFormWidgetGeometryRect(
        candidate.targetRect,
      );
      if (!expected || !targetRect) return null;
      operations.push({
        schemaVersion: NATIVE_PDF_FORM_WIDGET_GEOMETRY_SCHEMA_VERSION,
        id,
        operation,
        expected,
        targetRect,
      });
      continue;
    }
    if (operation === "fill-image") {
      const expected = readNativePdfExpectedIdentity(candidate.expected);
      const image = candidate.image as Record<string, unknown> | null | undefined;
      if (
        candidate.schemaVersion !== NATIVE_PDF_FORM_IMAGE_SCHEMA_VERSION ||
        Object.keys(candidate).sort().join(",") !== "expected,id,image,operation,schemaVersion" ||
        !expected ||
        expected.kind !== "imageField" ||
        (image !== null &&
          (!image ||
            typeof image !== "object" ||
            Object.keys(image).sort().join(",") !== "data,mime" ||
            (image.mime !== "image/png" && image.mime !== "image/jpeg") ||
            typeof image.data !== "string" ||
            !image.data ||
            image.data.length > Math.ceil(NATIVE_PDF_FORM_IMAGE_MAX_BYTES / 3) * 4 ||
            !/^[A-Za-z0-9+/]+={0,2}$/.test(image.data)))
      )
        return null;
      operations.push({
        schemaVersion: NATIVE_PDF_FORM_IMAGE_SCHEMA_VERSION,
        id,
        operation,
        expected,
        image: image === null ? null : {
          mime: image.mime as "image/png" | "image/jpeg",
          data: image.data as string,
        },
      });
      continue;
    }
    if (operation === "snapshot-value") {
      if (
        candidate.schemaVersion !==
          NATIVE_PDF_FORM_VALUE_SNAPSHOT_SCHEMA_VERSION ||
        Object.keys(candidate).sort().join(",") !==
          "expected,id,operation,schemaVersion,value"
      )
        return null;
      const rawExpected = candidate.expected;
      const expected =
        rawExpected &&
        typeof rawExpected === "object" &&
        "widgets" in rawExpected
          ? readNativePdfGroupExpectedIdentity(rawExpected)
          : readNativePdfExpectedIdentity(rawExpected);
      if (
        !expected ||
        !["text", "checkbox", "radio", "dropdown", "listbox"].includes(
          expected.kind,
        )
      )
        return null;
      const value = candidate.value;
      const textValid = (text: unknown): text is string =>
        typeof text === "string" &&
        text.length <= ACROFORM_MAX_TEXT &&
        Array.from(text).every(
          (character) =>
            character.length !== 1 || !/[\uD800-\uDFFF]/.test(character),
        );
      const valueIsValid =
        (expected.kind === "checkbox" && typeof value === "boolean") ||
        ((expected.kind === "text" ||
          expected.kind === "dropdown" ||
          expected.kind === "radio" ||
          expected.kind === "listbox") &&
          textValid(value)) ||
        (expected.kind === "listbox" &&
          Array.isArray(value) &&
          value.length <= ACROFORM_MAX_OPTIONS &&
          value.every(textValid) &&
          new Set(value).size === value.length);
      if (!valueIsValid) return null;
      operations.push({
        schemaVersion: NATIVE_PDF_FORM_VALUE_SNAPSHOT_SCHEMA_VERSION,
        id,
        operation,
        expected,
        value: Array.isArray(value) ? [...value] : value,
      });
      continue;
    }
    if (operation === "edit-group-properties") {
      if (
        Object.keys(candidate).sort().join("\u0000") !==
          [
            "expected",
            "id",
            "operation",
            "readOnly",
            "required",
            "schemaVersion",
          ]
            .sort()
            .join("\u0000") ||
        candidate.schemaVersion !==
          NATIVE_PDF_FORM_GROUP_PROPERTIES_SCHEMA_VERSION ||
        typeof candidate.required !== "boolean" ||
        typeof candidate.readOnly !== "boolean"
      ) {
        return null;
      }
      const expected = readNativePdfGroupExpectedIdentity(candidate.expected);
      if (!expected) return null;
      operations.push({
        schemaVersion: NATIVE_PDF_FORM_GROUP_PROPERTIES_SCHEMA_VERSION,
        id,
        operation,
        expected,
        required: candidate.required,
        readOnly: candidate.readOnly,
      });
      continue;
    }
    if (operation === "edit-radio-options") {
      if (
        Object.keys(candidate).sort().join("\u0000") !==
          [
            "id",
            "name",
            "offToggleable",
            "operation",
            "schemaVersion",
            "selectedValue",
            "widgets",
          ]
            .sort()
            .join("\u0000") ||
        candidate.schemaVersion !==
          NATIVE_PDF_RADIO_GROUP_OPTIONS_SCHEMA_VERSION ||
        typeof candidate.name !== "string" ||
        candidate.name.trim() !== candidate.name ||
        !isValidFormFieldNamePath(candidate.name) ||
        typeof candidate.offToggleable !== "boolean" ||
        !Array.isArray(candidate.widgets) ||
        candidate.widgets.length < 2 ||
        candidate.widgets.length > 200
      ) {
        return null;
      }
      const seenIdentities = new Set<string>();
      const originalValues = new Set<string>();
      const replacementValues = new Set<string>();
      const widgets: NativePdfRadioGroupOptionMutation[] = [];
      for (const rawWidget of candidate.widgets) {
        if (!rawWidget || typeof rawWidget !== "object") return null;
        const widget = rawWidget as Record<string, unknown>;
        if (
          !Number.isSafeInteger(widget.objectNumber) ||
          Number(widget.objectNumber) < 1 ||
          !Number.isSafeInteger(widget.generation) ||
          Number(widget.generation) < 0 ||
          !Number.isSafeInteger(widget.page) ||
          Number(widget.page) < 1 ||
          !isNativePdfFormExportValue(widget.originalValue) ||
          !isNativePdfFormExportValue(widget.replacementValue)
        ) {
          return null;
        }
        const identity = `${widget.objectNumber}:${widget.generation}:${widget.page}`;
        if (
          seenIdentities.has(identity) ||
          originalValues.has(widget.originalValue) ||
          replacementValues.has(widget.replacementValue)
        ) {
          return null;
        }
        seenIdentities.add(identity);
        originalValues.add(widget.originalValue);
        replacementValues.add(widget.replacementValue);
        widgets.push({
          objectNumber: Number(widget.objectNumber),
          generation: Number(widget.generation),
          page: Number(widget.page),
          originalValue: widget.originalValue,
          replacementValue: widget.replacementValue,
        });
      }
      widgets.sort(
        (left, right) =>
          left.objectNumber - right.objectNumber ||
          left.generation - right.generation ||
          left.page - right.page,
      );
      const selectedValue = candidate.selectedValue;
      if (
        selectedValue !== null &&
        (typeof selectedValue !== "string" ||
          !replacementValues.has(selectedValue))
      ) {
        return null;
      }
      operations.push({
        schemaVersion: NATIVE_PDF_RADIO_GROUP_OPTIONS_SCHEMA_VERSION,
        id,
        operation,
        name: candidate.name,
        widgets,
        selectedValue,
        offToggleable: candidate.offToggleable,
      });
      continue;
    }
    if (candidate.schemaVersion !== NATIVE_PDF_FORM_AUTHORING_SCHEMA_VERSION) {
      return null;
    }
    if (operation === "create") {
      const field = readNativePdfFormFieldSpec(candidate.field);
      if (!field) return null;
      operations.push({
        schemaVersion: NATIVE_PDF_FORM_AUTHORING_SCHEMA_VERSION,
        id,
        operation,
        field,
      });
      continue;
    }
    const expected = readNativePdfExpectedIdentity(candidate.expected);
    if (!expected) return null;
    if (operation === "edit") {
      const field = readNativePdfFormFieldSpec(candidate.field);
      if (
        !field ||
        field.kind !== expected.kind ||
        field.name !== expected.name
      ) {
        return null;
      }
      operations.push({
        schemaVersion: NATIVE_PDF_FORM_AUTHORING_SCHEMA_VERSION,
        id,
        operation,
        expected,
        field,
      });
      continue;
    }
    if (operation !== "delete") return null;
    operations.push({
      schemaVersion: NATIVE_PDF_FORM_AUTHORING_SCHEMA_VERSION,
      id,
      operation,
      expected,
    });
  }
  return operations;
}

/** A full AcroForm name is a dotted path of non-empty partial-name segments. */
export function isValidFormFieldNamePath(value: string): boolean {
  if (!value) return false;
  return value.split(".").every(
    (segment) =>
      segment.length > 0 &&
      segment === segment.trim() &&
      !FIELD_NAME_CONTROL_CHARACTER.test(segment),
  );
}

/**
 * A terminal field cannot also be a non-terminal parent. Siblings under one
 * dotted parent are valid; an exact match is handled by the field-kind rules.
 */
export function formFieldNamesHaveHierarchyConflict(
  left: string,
  right: string,
): boolean {
  return (
    left !== right &&
    (left.startsWith(`${right}.`) || right.startsWith(`${left}.`))
  );
}

/**
 * Check a proposed field against the form it would join. Returns every problem
 * found, so a dialog can show them all at once instead of one per attempt.
 */
export function reviewFormField(
  draft: FormFieldDraft,
  existing: readonly ExistingFormField[] = [],
): FormFieldProblem[] {
  const problems: FormFieldProblem[] = [];
  const name = draft.name.trim();

  if (!name) problems.push("empty-name");
  if (name.length > MAX_FIELD_NAME_LENGTH) problems.push("name-too-long");
  if (name && !isValidFormFieldNamePath(name)) {
    problems.push("invalid-name-path");
  }

  const alternateName = draft.alternateName?.trim() ?? "";
  if (alternateName.length > MAX_FIELD_ALTERNATE_NAME_LENGTH) {
    problems.push("alternate-name-too-long");
  }
  if (
    [...alternateName].some((character) => {
      const code = character.charCodeAt(0);
      return code < 32 || code === 127;
    })
  ) {
    problems.push("invalid-alternate-name");
  }

  const sameName = existing.filter((field) => field.name === name);
  if (sameName.length > 0) {
    // Radio buttons SHARE a name on purpose — that is what groups them.
    if (draft.kind === "radio" && sameName.every((f) => f.kind === "radio")) {
      const taken = new Set(
        sameName.map((field) => field.exportValue).filter(Boolean),
      );
      if (draft.exportValue && taken.has(draft.exportValue)) {
        problems.push("duplicate-export-value");
      }
    } else if (sameName.some((field) => field.kind !== draft.kind)) {
      problems.push("kind-conflict");
    } else {
      problems.push("duplicate-name");
    }
  }
  if (
    existing.some((field) =>
      formFieldNamesHaveHierarchyConflict(name, field.name),
    )
  ) {
    problems.push("hierarchy-name-conflict");
  }

  if (TOGGLE_KINDS.has(draft.kind)) {
    const value = (draft.exportValue ?? "").trim();
    if (!value) problems.push("missing-export-value");
    else if (value === RESERVED_OFF_VALUE)
      problems.push("reserved-export-value");
  }

  if (CHOICE_KINDS.has(draft.kind)) {
    const options = (draft.options ?? []).filter(
      (option) => option.trim() !== "",
    );
    if (options.length === 0) problems.push("missing-options");
    else if (draft.defaultValue && !options.includes(draft.defaultValue)) {
      problems.push("default-not-an-option");
    }
  }

  if (
    draft.kind === "text" &&
    draft.maxLength !== undefined &&
    (!Number.isFinite(draft.maxLength) || draft.maxLength < 0)
  ) {
    problems.push("negative-max-length");
  }
  const enabledTextFlags = [draft.multiline, draft.password, draft.comb].filter(
    Boolean,
  ).length;
  if (draft.kind === "text" && enabledTextFlags > 1) {
    problems.push("incompatible-text-flags");
  }
  if (
    draft.kind === "text" &&
    draft.comb &&
    (!Number.isSafeInteger(draft.maxLength) || (draft.maxLength ?? 0) < 1)
  ) {
    problems.push("comb-requires-max-length");
  }
  if (draft.kind === "text" && draft.password && draft.defaultValue) {
    problems.push("password-default-value");
  }
  if (
    draft.kind === "text" &&
    draft.alignment !== undefined &&
    draft.alignment !== "left" &&
    draft.alignment !== "center" &&
    draft.alignment !== "right"
  ) {
    problems.push("invalid-text-alignment");
  }
  if (
    draft.kind === "text" &&
    draft.fontSize !== undefined &&
    (!Number.isFinite(draft.fontSize) ||
      draft.fontSize < 0 ||
      draft.fontSize > ACROFORM_MAX_TEXT_FONT_SIZE)
  ) {
    problems.push("invalid-text-font-size");
  }
  if (
    draft.kind === "text" &&
    draft.textColor !== undefined &&
    normalizePdfTextFieldColor(draft.textColor) === null
  ) {
    problems.push("invalid-text-color");
  }
  if (
    draft.kind === "text" &&
    draft.backgroundColor !== undefined &&
    normalizePdfTextFieldPaint(draft.backgroundColor) === null
  ) {
    problems.push("invalid-text-background-color");
  }
  if (
    draft.kind === "text" &&
    draft.borderColor !== undefined &&
    normalizePdfTextFieldPaint(draft.borderColor) === null
  ) {
    problems.push("invalid-text-border-color");
  }
  if (draft.kind === "text" && draft.format !== undefined && !pdfTextFieldFormatValid(draft.format)) {
    problems.push("invalid-text-format");
  }
  if (
    draft.kind === "text" &&
    draft.borderWidth !== undefined &&
    (!Number.isFinite(draft.borderWidth) ||
      draft.borderWidth < 0 ||
      draft.borderWidth > ACROFORM_MAX_WIDGET_BORDER_WIDTH)
  ) {
    problems.push("invalid-text-border-width");
  }
  if (
    draft.kind === "text" &&
    draft.visibility !== undefined &&
    draft.visibility !== "visible" &&
    draft.visibility !== "screenOnly" &&
    draft.visibility !== "printOnly" &&
    draft.visibility !== "hidden"
  ) {
    problems.push("invalid-text-visibility");
  }
  if (
    draft.tabOrder !== undefined &&
    (!Number.isSafeInteger(draft.tabOrder) ||
      draft.tabOrder < 1 ||
      draft.tabOrder > MAX_NATIVE_PDF_FORM_TAB_ORDER)
  ) {
    problems.push("invalid-tab-order");
  }

  return problems;
}

/** A field ready to write, or null when it was refused. */
export function normalizeFormField(
  draft: FormFieldDraft,
  existing: readonly ExistingFormField[] = [],
): FormFieldDraft | null {
  if (reviewFormField(draft, existing).length > 0) return null;
  const normalized: FormFieldDraft = {
    kind: draft.kind,
    name: draft.name.trim(),
  };
  if (draft.alternateName?.trim()) {
    normalized.alternateName = draft.alternateName.trim();
  }
  const role = draft.role === undefined ? null : readPdfFormRole(draft.role);
  if (role) normalized.role = role;
  if (draft.exportValue) normalized.exportValue = draft.exportValue.trim();
  if (draft.options) {
    normalized.options = draft.options
      .map((option) => option.trim())
      .filter((option) => option !== "");
  }
  if (draft.defaultValue) normalized.defaultValue = draft.defaultValue;
  if (draft.required) normalized.required = true;
  if (draft.readOnly) normalized.readOnly = true;
  if (draft.maxLength) normalized.maxLength = Math.floor(draft.maxLength);
  if (draft.kind === "text" && draft.multiline) normalized.multiline = true;
  if (draft.kind === "text" && draft.password) normalized.password = true;
  if (draft.kind === "text" && draft.comb) normalized.comb = true;
  if (draft.kind === "text" && typeof draft.spellCheck === "boolean") {
    normalized.spellCheck = draft.spellCheck;
  }
  if (draft.kind === "text" && typeof draft.scrollable === "boolean") {
    normalized.scrollable = draft.scrollable;
  }
  if (draft.kind === "text" && draft.alignment) {
    normalized.alignment = draft.alignment;
  }
  if (draft.kind === "text" && draft.fontSize !== undefined) {
    normalized.fontSize = draft.fontSize;
  }
  if (draft.kind === "text" && draft.textColor !== undefined) {
    normalized.textColor =
      normalizePdfTextFieldColor(draft.textColor) ?? undefined;
  }
  if (draft.kind === "text" && draft.backgroundColor !== undefined) {
    normalized.backgroundColor =
      normalizePdfTextFieldPaint(draft.backgroundColor) ?? undefined;
  }
  if (draft.kind === "text" && draft.borderColor !== undefined) {
    normalized.borderColor =
      normalizePdfTextFieldPaint(draft.borderColor) ?? undefined;
  }
  if (draft.kind === "text" && draft.borderWidth !== undefined) {
    normalized.borderWidth = draft.borderWidth;
  }
  if (draft.kind === "text" && draft.visibility !== undefined) {
    normalized.visibility = draft.visibility;
  }
  if (draft.kind === "text" && draft.format !== undefined) normalized.format = draft.format;
  if (draft.tabOrder !== undefined) normalized.tabOrder = draft.tabOrder;
  return normalized;
}

/** Fields a reader tabs through, in the order PDF visits them. */
export function tabOrder(
  fields: readonly ExistingFormField[],
  positions: ReadonlyMap<string, { row: number; col: number }>,
): ExistingFormField[] {
  return [...fields].sort((a, b) => {
    const left = positions.get(a.name);
    const right = positions.get(b.name);
    if (!left || !right) return 0;
    return left.row === right.row ? left.col - right.col : left.row - right.row;
  });
}
