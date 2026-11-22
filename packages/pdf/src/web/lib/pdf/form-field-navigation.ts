/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A PDF form's Previous Field / Next Field (Word's Forms tab has them, the
 * Docs editor's documentFormFill; ONLYOFFICE's mobile PDF editor puts them on
 * its toolbar): the fields in the form's own order -- the order of their
 * widgets in the file, which is the PDF viewer's DOM order and the Tab
 * order -- the one after (or before) the field being filled focused and
 * brought into view. From outside any field, Next is the first and Previous
 * the last. Fields of pages not drawn yet are not counted (the viewer draws
 * the pages near the one in view).
 */

export const PDF_FORM_FIELD_SELECTOR = ".pdf-viewer__form-field";

/** The fields a person can move to, in the form's order. */
export function pdfFormFields(root: ParentNode): HTMLElement[] {
  return Array.from(root.querySelectorAll<HTMLElement>(PDF_FORM_FIELD_SELECTOR)).filter(
    (field) =>
      !(field as HTMLInputElement).disabled &&
      field.getAttribute("aria-disabled") !== "true" &&
      field.tabIndex >= 0 &&
      field.getAttribute("aria-hidden") !== "true",
  );
}

/** Focuses the field after (1) or before (-1) the one being filled and
 * answers it, or null when the form has no field to move to. */
export function movePdfFormField(root: ParentNode, direction: 1 | -1): HTMLElement | null {
  const fields = pdfFormFields(root);
  if (!fields.length) return null;
  const active = typeof document === "undefined" ? null : document.activeElement;
  const at = active instanceof HTMLElement ? fields.findIndex((field) => field === active || field.contains(active)) : -1;
  const next = at < 0 ? (direction > 0 ? 0 : fields.length - 1) : (at + direction + fields.length) % fields.length;
  const field = fields[next]!;
  field.focus({ preventScroll: true });
  field.scrollIntoView?.({ block: "center", inline: "center", behavior: "smooth" });
  return field;
}

/** Whether the focus is in one of the form's fields (the keyboard's bar
 * then offers Previous / Next Field). */
export function focusInPdfFormField(root: ParentNode): boolean {
  const active = typeof document === "undefined" ? null : document.activeElement;
  return active instanceof HTMLElement && Boolean(active.closest(PDF_FORM_FIELD_SELECTOR)) && (root as Node).contains(active);
}
