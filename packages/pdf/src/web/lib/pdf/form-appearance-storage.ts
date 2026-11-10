/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Form widgets are painted by the page itself, in every mode.
 *
 * Like WPS and Acrobat, MRAI draws an AcroForm widget from the PDF's own
 * appearance stream whether the reader is browsing, filling or editing the
 * page. The HTML controls of fill mode are transparent hit targets over that
 * painting (the caret's field alone shows a live editor). So a value the user
 * changed must reach the painting: pdf.js re-generates a widget's appearance
 * from its `annotationStorage` entry when a page renders with
 * `AnnotationMode.ENABLE_STORAGE`, exactly as it does for printing a filled
 * form. This module keeps that storage in step with the viewer's value map.
 */
import type { AcroField, AcroFormValue } from "@mrnote/core/web/lib/pdf/acroform.ts";
import { formatWithSpec, parseFormatScript } from "@mrnote/core/web/lib/pdf/acroform-scripts.ts";

/** The subset of pdf.js `AnnotationStorage` the sync needs. */
export interface PdfAnnotationStorageLike {
  setValue(key: string, value: Record<string, unknown>): void;
  remove(key: string): void;
  /** pdf.js answers `Object.assign(defaultValue, entry)` when an entry
   * exists: the default must be an object (`{}`), never null or undefined,
   * or the read throws and takes the page's render down with it (user
   * audit PDF 0018: the window crashed on reaching a password field). */
  getValue?(key: string, defaultValue: Record<string, unknown>): unknown;
}

/** `AnnotationMode.ENABLE_STORAGE`: widgets render from `annotationStorage`. */
export const PDFJS_ANNOTATION_MODE_ENABLE_STORAGE = 3;

const POSITIONAL_FIELD_ID = /^\d+-\d+$/;

/**
 * The pdf.js annotation id behind a normalized field id, or null when the
 * id is the positional fallback (`field-<page>-<index>`) that no widget owns.
 */
export function pdfWidgetIdForField(fieldId: string): string | null {
  if (!fieldId.startsWith("field-")) return null;
  const id = fieldId.slice("field-".length);
  if (id.length === 0 || POSITIONAL_FIELD_ID.test(id)) return null;
  return id;
}

/**
 * The storage `value` pdf.js expects for one widget of a field holding
 * `value`, or null when the widget kind has no stored value (buttons).
 */
export function pdfWidgetStorageValue(
  field: Pick<AcroField, "type" | "exportValue" | "multiSelect">,
  value: AcroFormValue | undefined,
): boolean | string | string[] | null {
  switch (field.type) {
    case "checkbox":
      return value === true;
    case "radio":
      return (
        typeof value === "string" &&
        value !== "" &&
        field.exportValue !== null &&
        value === field.exportValue
      );
    case "text":
      return typeof value === "string" ? value : null;
    case "choice":
      if (field.multiSelect) {
        return Array.isArray(value)
          ? value.filter((entry): entry is string => typeof entry === "string")
          : typeof value === "string"
            ? [value]
            : null;
      }
      return typeof value === "string"
        ? value
        : Array.isArray(value) && typeof value[0] === "string"
          ? value[0]
          : null;
    default:
      return null;
  }
}

/** A list box's selected values in one order (a set of rows, whichever
 * order they were chosen in); any other value as it is. */
function normalizedSelection(
  field: Pick<AcroField, "type" | "multiSelect">,
  value: boolean | string | string[],
): boolean | string | string[] {
  return field.type === "choice" && field.multiSelect && Array.isArray(value) ? [...value].sort() : value;
}

function sameStorageValue(
  left: boolean | string | string[],
  right: boolean | string | string[],
): boolean {
  if (Array.isArray(left) || Array.isArray(right)) {
    return (
      Array.isArray(left) &&
      Array.isArray(right) &&
      left.length === right.length &&
      left.every((entry, index) => entry === right[index])
    );
  }
  return left === right;
}

/**
 * Hiding comments hides the page's markup, not its form. With every widget
 * painted by the page, the canvas cannot simply skip annotations while
 * comments are hidden; instead each non-widget annotation is flagged
 * `noView` in the storage (pdf.js honours the flag when it renders with
 * `ENABLE_STORAGE`) and the flag is lifted when comments show again.
 * `tracked` remembers the ids this viewer flagged so nothing else's entries
 * are touched. Returns whether any flag changed.
 */
export function syncPdfCommentVisibilityStorage(
  storage: PdfAnnotationStorageLike,
  annotations: ReadonlyArray<unknown>,
  hidden: boolean,
  tracked: Set<string>,
): boolean {
  let changed = false;
  for (const annotation of annotations) {
    if (!annotation || typeof annotation !== "object") continue;
    const { id, subtype } = annotation as { id?: unknown; subtype?: unknown };
    if (typeof id !== "string" || id.length === 0 || subtype === "Widget") {
      continue;
    }
    if (hidden) {
      if (tracked.has(id)) continue;
      storage.setValue(id, { noView: true });
      tracked.add(id);
      changed = true;
    } else if (tracked.has(id)) {
      storage.remove(id);
      tracked.delete(id);
      changed = true;
    }
  }
  return changed;
}

/**
 * The PDF's own copy of each annotation a held edit replaces (AutoSave off:
 * the new version drawn by the page's layers until Save,
 * lib/pdf/held-annotation-targets.ts) is flagged `noView`; the flag is lifted
 * when it is no longer held. `tracked` remembers what this flagged; an entry
 * another sync owns (`keep`: hidden comments, NoZoom drawings) stays when
 * lifted. Returns whether any flag changed.
 */
export function syncPdfHeldAnnotationStorage(
  storage: PdfAnnotationStorageLike,
  annotations: ReadonlyArray<unknown>,
  held: ReadonlySet<string>,
  tracked: Set<string>,
  keep: (id: string) => boolean,
): boolean {
  let changed = false;
  for (const annotation of annotations) {
    if (!annotation || typeof annotation !== "object") continue;
    const { id } = annotation as { id?: unknown };
    if (typeof id !== "string" || id.length === 0) continue;
    if (held.has(id)) {
      if (tracked.has(id)) continue;
      const previous = storage.getValue?.(id, {});
      storage.setValue(id, {
        ...(previous && typeof previous === "object" ? (previous as Record<string, unknown>) : {}),
        noView: true,
      });
      tracked.add(id);
      changed = true;
    } else if (tracked.has(id)) {
      tracked.delete(id);
      if (!keep(id)) storage.remove(id);
      changed = true;
    }
  }
  return changed;
}

/**
 * The pdf.js ids of the password text widgets among a page's annotations
 * (`getAnnotations()` data) or a document's field objects
 * (`getFieldObjects()` entries).
 */
export function pdfPasswordWidgetIds(
  annotations: ReadonlyArray<unknown>,
): string[] {
  const ids: string[] = [];
  for (const annotation of annotations) {
    if (!annotation || typeof annotation !== "object") continue;
    const { id, fieldType, type, password } = annotation as {
      id?: unknown;
      fieldType?: unknown;
      type?: unknown;
      password?: unknown;
    };
    if (
      typeof id === "string" &&
      id.length > 0 &&
      (fieldType === "Tx" || type === "text") &&
      password === true
    ) {
      ids.push(id);
    }
  }
  return ids;
}

/**
 * A password field's own appearance stream can hold its value as plain text
 * (whatever wrote the file drew it so), and pdf.js keeps that stream for a
 * password field whatever its storage entry says. So the page never paints a
 * password widget: each is flagged `noView` here, before the page renders,
 * and the field's control draws the field, masked, in every state (user
 * audit PDF 0006: the canvas showed the value under a `type=password`
 * control). Nothing lifts the flag; an export shows the field blank rather
 * than the value. Returns whether any flag was added.
 */
export function syncPdfPasswordWidgetStorage(
  storage: PdfAnnotationStorageLike,
  widgetIds: Iterable<string>,
): boolean {
  let changed = false;
  for (const id of widgetIds) {
    const current = storage.getValue?.(id, {}) as
      | { noView?: unknown }
      | undefined;
    if (current?.noView === true) continue;
    storage.setValue(id, { noView: true });
    changed = true;
  }
  return changed;
}

/**
 * Flags every password widget of a loaded document in its own storage, for
 * a view that renders its pages with `ENABLE_STORAGE` but has no form layer
 * of its own (a comparison, a preview): the field shows blank there, never
 * its value.
 */
export async function maskPdfDocumentPasswordWidgets(document_: {
  getFieldObjects?: () => Promise<Record<string, unknown[]> | null>;
  annotationStorage?: PdfAnnotationStorageLike;
}): Promise<void> {
  const storage = document_.annotationStorage;
  if (!storage || typeof document_.getFieldObjects !== "function") return;
  let fields: Record<string, unknown[]> | null = null;
  try {
    fields = await document_.getFieldObjects();
  } catch {
    return;
  }
  if (!fields) return;
  syncPdfPasswordWidgetStorage(
    storage,
    pdfPasswordWidgetIds(Object.values(fields).flat()),
  );
}

const NO_ZOOM_SHAPE_SUBTYPES = new Set([
  "Square",
  "Circle",
  "Line",
  "Polygon",
  "PolyLine",
]);










export function syncPdfNoZoomAppearanceStorage(
  storage: PdfAnnotationStorageLike,
  annotations: ReadonlyArray<unknown>,
  tracked: Set<string>,
): boolean {
  let changed = false;
  for (const annotation of annotations) {
    if (!annotation || typeof annotation !== "object") continue;
    const { id, subtype, annotationFlags } = annotation as {
      id?: unknown;
      subtype?: unknown;
      annotationFlags?: unknown;
    };
    if (
      typeof id !== "string" ||
      id.length === 0 ||
      typeof subtype !== "string" ||
      !NO_ZOOM_SHAPE_SUBTYPES.has(subtype) ||
      typeof annotationFlags !== "number" ||
      (annotationFlags & (1 << 3)) === 0
    ) {
      continue;
    }
    const current = storage.getValue?.(id, {}) as
      | { noView?: unknown }
      | undefined;
    if (current?.noView === true && tracked.has(id)) continue;
    storage.setValue(id, { noView: true });
    tracked.add(id);
    changed = true;
  }
  return changed;
}














export function withPdfNoZoomAppearanceShown<T>(
  storage: PdfAnnotationStorageLike | null | undefined,
  noZoomIds: ReadonlySet<string>,
  hiddenComments: ReadonlySet<string>,
  start: () => T,
): T {
  const lifted: string[] = [];
  if (storage) {
    for (const id of noZoomIds) {
      if (hiddenComments.has(id)) continue;
      storage.remove(id);
      lifted.push(id);
    }
  }
  try {
    return start();
  } finally {
    for (const id of lifted) storage?.setValue(id, { noView: true });
  }
}

/** The display text of a text field's value under its format script, when
 * it has one and the value is not blank; null when the page paints the value
 * itself. */
export function pdfWidgetFormattedValue(
  field: Pick<AcroField, "type" | "formatScript" | "password">,
  value: boolean | string | string[],
  locale?: string,
): string | null {
  if (field.type !== "text" || typeof value !== "string" || value === "" || field.password) return null;
  const spec = field.formatScript ? parseFormatScript(field.formatScript) : null;
  if (!spec) return null;
  const formatted = formatWithSpec(value, spec, locale ?? "en");
  return formatted === value ? null : formatted;
}

export interface PdfFormAppearanceSyncResult {
  /** Every widget id the storage now carries an entry for. */
  keys: Map<string, boolean | string | string[]>;
  /** Field ids whose widget entry was written or removed by this sync. */
  changedFieldIds: string[];
}

/**
 * Writes the current value of every overridden field into the pdf.js storage
 * and removes entries for fields that no longer carry an override, so a page
 * rendered with `ENABLE_STORAGE` paints exactly the viewer's values. Fields
 * without an override keep the file's own appearance (no entry).
 */
export function syncPdfFormAppearanceStorage(
  storage: PdfAnnotationStorageLike,
  fields: ReadonlyArray<AcroField>,
  values: Readonly<Record<string, AcroFormValue>>,
  previous: ReadonlyMap<string, boolean | string | string[]>,
  locale?: string,
): PdfFormAppearanceSyncResult {
  const keys = new Map<string, boolean | string | string[]>();
  const changedFieldIds: string[] = [];
  for (const field of fields) {
    const widgetId = pdfWidgetIdForField(field.id);
    if (!widgetId || !Object.hasOwn(values, field.name)) continue;
    // pdf.js never draws a password field from its value, and its entry
    // carries the `noView` flag that keeps the page from painting it
    // (syncPdfPasswordWidgetStorage): no value goes in, none is removed.
    if (field.password) continue;
    const value = pdfWidgetStorageValue(field, values[field.name]);
    if (value === null) continue;
    // The value the file already holds is painted by the field's own
    // appearance, with its fill and border: pdf.js redraws a list box from
    // a storage entry with neither, so a saved (or put back) selection
    // showed a white, borderless list (user audit PDF 0074).
    const held = field.initialValue === undefined || field.initialValue === null
      ? null : pdfWidgetStorageValue(field, field.initialValue);
    if (held !== null && sameStorageValue(normalizedSelection(field, held), normalizedSelection(field, value))) continue;
    keys.set(widgetId, value);
    const before = previous.get(widgetId);
    if (before !== undefined && sameStorageValue(before, value)) continue;
    // A field with a standard format script (AFNumber_Format,
    // AFPercent_Format, ...) stores the raw value as /V and shows the
    // formatted one: the page paints the storage entry's formattedValue
    // (pdf.js draws it in place of the value), so "$1,234.50" stays on the
    // page after 1234.5 is typed, as the reader's own scripting would show.
    const formatted = pdfWidgetFormattedValue(field, value, locale);
    storage.setValue(widgetId, formatted === null ? { value } : { value, formattedValue: formatted });
    changedFieldIds.push(field.id);
  }
  for (const widgetId of previous.keys()) {
    if (keys.has(widgetId)) continue;
    storage.remove(widgetId);
    const fieldId = fields.find(
      (field) => pdfWidgetIdForField(field.id) === widgetId,
    )?.id;
    if (fieldId) changedFieldIds.push(fieldId);
  }
  return { keys, changedFieldIds };
}
