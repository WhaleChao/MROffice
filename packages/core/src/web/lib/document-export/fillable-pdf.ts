/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A document's PDF made a fillable form: an AcroForm field over each of its
 * content controls.
 *
 * The server prints the document with each form control's words linked to
 * `https://form-field.mrai.invalid/#<field>` (`document_pdf_form_fields.py`),
 * the link's fragment the field's description; the print writes each line of
 * such a link as a `/Link` annotation with the line's rectangle. Here every
 * such link is replaced by the field it describes, written with pdf-lib as
 * the PDF editor writes its form fields (`lib/pdf/form-field-create`):
 * a text field, a check box, an option button of its group, a drop-down
 * (editable for a combo box), required where the control is required. The
 * field's white face covers the printed words under it, so a value shows
 * once. A value the standard font cannot draw (Chinese, say) is set without
 * an appearance and the reader is asked to draw it (`NeedAppearances`), as
 * Acrobat, PDFium (Chrome, Edge) and pdf.js (Firefox) all do.
 */

import {
  PDFArray,
  PDFBool,
  PDFDict,
  PDFDocument,
  PDFHexString,
  PDFName,
  PDFString,
  rgb,
  type PDFPage,
} from "pdf-lib";
import { pdfLibObjectsPerTick, pdfLibParseSpeed } from "../pdf/pdf-lib-pacing";

export const FIELD_LINK_PREFIX = "https://form-field.mrai.invalid/#";

interface FieldDescription {
  readonly i: number;
  readonly k: "text" | "richText" | "checkbox" | "dropDownList" | "comboBox" | "date";
  readonly n: string;
  readonly v?: string;
  readonly c?: boolean;
  readonly g?: string;
  readonly o?: readonly string[];
  readonly r?: boolean;
  readonly m?: boolean;
}

interface Placed {
  readonly field: FieldDescription;
  readonly page: PDFPage;
  readonly rects: { x: number; y: number; width: number; height: number }[];
}

function decodeField(uri: string): FieldDescription | null {
  if (!uri.startsWith(FIELD_LINK_PREFIX)) return null;
  try {
    const base64 = uri.slice(FIELD_LINK_PREFIX.length).replace(/-/g, "+").replace(/_/g, "/");
    const bytes = Uint8Array.from(atob(base64 + "=".repeat((4 - (base64.length % 4)) % 4)), (c) => c.charCodeAt(0));
    const parsed = JSON.parse(new TextDecoder().decode(bytes)) as FieldDescription;
    return typeof parsed?.i === "number" && typeof parsed.k === "string" && typeof parsed.n === "string"
      ? parsed
      : null;
  } catch {
    return null;
  }
}

function linkUri(annotation: PDFDict): string | null {
  if (annotation.lookup(PDFName.of("Subtype")) !== PDFName.of("Link")) return null;
  const action = annotation.lookup(PDFName.of("A"));
  if (!(action instanceof PDFDict)) return null;
  const uri = action.lookup(PDFName.of("URI"));
  return uri instanceof PDFString || uri instanceof PDFHexString ? uri.decodeText() : null;
}

/** The field links of every page, taken off their pages. */
function takeFieldLinks(doc: PDFDocument): Placed[] {
  const placed = new Map<number, Placed>();
  for (const page of doc.getPages()) {
    const annots = page.node.lookup(PDFName.of("Annots"));
    if (!(annots instanceof PDFArray)) continue;
    const kept: unknown[] = [];
    for (let index = 0; index < annots.size(); index += 1) {
      const ref = annots.get(index);
      const annotation = annots.lookup(index);
      const uri = annotation instanceof PDFDict ? linkUri(annotation) : null;
      const field = uri ? decodeField(uri) : null;
      if (!field || !(annotation instanceof PDFDict)) {
        kept.push(ref);
        continue;
      }
      const rect = annotation.lookup(PDFName.of("Rect"));
      if (rect instanceof PDFArray && rect.size() === 4) {
        const [x1, y1, x2, y2] = [0, 1, 2, 3].map((at) => Number((rect.lookup(at) as { asNumber?: () => number }).asNumber?.() ?? 0));
        const box = {
          x: Math.min(x1!, x2!),
          y: Math.min(y1!, y2!),
          width: Math.abs(x2! - x1!),
          height: Math.abs(y2! - y1!),
        };
        const known = placed.get(field.i);
        // A field is placed on the page its words begin on.
        if (!known) placed.set(field.i, { field, page, rects: [box] });
        else if (known.page === page) known.rects.push(box);
      }
    }
    if (kept.length !== annots.size()) {
      page.node.set(PDFName.of("Annots"), doc.context.obj(kept as never[]));
    }
  }
  return [...placed.values()].sort((left, right) => left.field.i - right.field.i);
}

const WIN_ANSI = /^[\x20-\x7e\xa0-\xff–—‘’“”•…€]*$/;

function uniqueName(name: string, used: Set<string>): string {
  // A partial field name holds no period (ISO 32000-1 12.7.3.2).
  const base = name.replace(/\./g, "·").trim() || "Field";
  let candidate = base;
  for (let copy = 2; used.has(candidate); copy += 1) candidate = `${base} ${copy}`;
  used.add(candidate);
  return candidate;
}

/** The first line's box (a multi-line field: every line's, together). */
function fieldBox(placed: Placed): { x: number; y: number; width: number; height: number } {
  const lines = placed.field.m ? placed.rects : placed.rects.slice(0, 1);
  const left = Math.min(...lines.map((rect) => rect.x));
  const bottom = Math.min(...lines.map((rect) => rect.y));
  const right = Math.max(...lines.map((rect) => rect.x + rect.width));
  const top = Math.max(...lines.map((rect) => rect.y + rect.height));
  return { x: left, y: bottom, width: Math.max(right - left, 4), height: Math.max(top - bottom, 4) };
}

function squareBox(box: { x: number; y: number; width: number; height: number }) {
  const side = Math.max(4, Math.min(box.width, box.height));
  return { x: box.x + (box.width - side) / 2, y: box.y + (box.height - side) / 2, width: side, height: side };
}

/**
 * `pdf` (the document's fillable print) with its field links made fields;
 * the bytes unchanged when it holds none.
 */
export async function fillablePdfFromFieldLinks(pdf: ArrayBuffer | Uint8Array): Promise<Uint8Array> {
  const doc = await PDFDocument.load(pdf, { parseSpeed: pdfLibParseSpeed(), updateMetadata: false });
  const placed = takeFieldLinks(doc);
  if (!placed.length) return new Uint8Array(pdf instanceof Uint8Array ? pdf : new Uint8Array(pdf));
  const form = doc.getForm();
  const used = new Set(form.getFields().map((field) => field.getName()));
  const face = { backgroundColor: rgb(1, 1, 1), borderColor: rgb(0.55, 0.55, 0.55), borderWidth: 0.5 };
  let needsAppearances = false;
  const groups = new Map<string, ReturnType<typeof form.createRadioGroup>>();
  const groupChosen = new Map<string, string>();
  const groupRequired = new Set<string>();

  for (const entry of placed) {
    const { field, page } = entry;
    const box = fieldBox(entry);
    if (field.k === "checkbox" && field.g) {
      let group = groups.get(field.g);
      if (!group) {
        group = form.createRadioGroup(uniqueName(field.g, used));
        group.enableMutualExclusion();
        groups.set(field.g, group);
      }
      const option = uniqueOption(group.getOptions(), field.n);
      group.addOptionToPage(option, page, { ...squareBox(box), ...face });
      if (field.c && !groupChosen.has(field.g)) groupChosen.set(field.g, option);
      if (field.r) groupRequired.add(field.g);
      continue;
    }
    if (field.k === "checkbox") {
      const box2 = squareBox(box);
      const checkBox = form.createCheckBox(uniqueName(field.n, used));
      checkBox.addToPage(page, { ...box2, ...face });
      if (field.c) checkBox.check();
      if (field.r) checkBox.enableRequired();
      continue;
    }
    // The widget first, drawn empty; then the value (whose appearance pdf-lib
    // draws on save when its standard font can, the reader otherwise).
    if (!WIN_ANSI.test(field.v ?? "") || (field.o ?? []).some((option) => !WIN_ANSI.test(option))) {
      needsAppearances = true;
    }
    if ((field.k === "dropDownList" || field.k === "comboBox") && field.o && field.o.length) {
      const dropdown = form.createDropdown(uniqueName(field.n, used));
      dropdown.setOptions([...field.o]);
      if (field.k === "comboBox") dropdown.enableEditing();
      dropdown.addToPage(page, { ...box, ...face });
      if (field.v && (field.o.includes(field.v) || field.k === "comboBox")) {
        dropdown.select(field.v);
      }
      if (field.r) dropdown.enableRequired();
      continue;
    }
    const text = form.createTextField(uniqueName(field.n, used));
    if (field.m) text.enableMultiline();
    text.addToPage(page, { ...box, ...face });
    if (field.v) text.setText(field.v);
    if (field.r) text.enableRequired();
  }
  for (const [name, group] of groups) {
    const chosen = groupChosen.get(name);
    if (chosen) group.select(chosen);
    if (groupRequired.has(name)) group.enableRequired();
  }
  if (needsAppearances) form.acroForm.dict.set(PDFName.of("NeedAppearances"), PDFBool.True);
  return doc.save({ objectsPerTick: pdfLibObjectsPerTick(), updateFieldAppearances: !needsAppearances });
}

function uniqueOption(existing: readonly string[], name: string): string {
  const base = name.trim() || "Option";
  let candidate = base;
  for (let copy = 2; existing.includes(candidate); copy += 1) candidate = `${base} ${copy}`;
  return candidate;
}
