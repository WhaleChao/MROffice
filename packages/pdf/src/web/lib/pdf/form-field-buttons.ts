/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The PDF form ribbon's field buttons as pure functions: what the editor
 * writes for a field a person adds through the field dialog. Shared by the
 * PDF viewer and the headless editor commands (lib/editor-commands/pdf).
 */

import type { PdfAuthoredFormFieldSelection } from "../../components/notebook/contents/pdf/PdfFormFieldAuthoringDialog";
import {
  isNativePdfFormFieldKind,
  type NativePdfFormFieldSpec,
} from "./form-field-authoring";
import { pdfTextFieldFormatTriggers } from "@mrnote/core/web/lib/pdf/form-field-format.ts";
import { sanitizeLinkUrl } from "@mrnote/core/web/lib/pdf/links.ts";

/** The native AcroForm field a confirmed "add field" dialog creates, or null
 *  when the field kind cannot be written natively or the draft is unnamed or
 *  has no tab position. */
export function nativeFormFieldSpec(
  selection: PdfAuthoredFormFieldSelection,
): NativePdfFormFieldSpec | null {
  if (!isNativePdfFormFieldKind(selection.kind)) return null;
  const tabOrder = selection.draft.tabOrder;
  if (
    !Number.isSafeInteger(tabOrder) ||
    (tabOrder ?? 0) < 1 ||
    !selection.draft.name.trim()
  ) {
    return null;
  }
  return {
    kind: selection.kind,
    name: selection.draft.name.trim(),
    ...(selection.draft.alternateName?.trim()
      ? { alternateName: selection.draft.alternateName.trim() }
      : {}),
    ...(selection.draft.role ? { role: { name: selection.draft.role.name, color: selection.draft.role.color } } : {}),
    page: selection.placement.page,
    rect: {
      x: selection.placement.x,
      y: selection.placement.y,
      width: selection.placement.width,
      height: selection.placement.height,
    },
    tabOrder: tabOrder!,
    required: selection.draft.required === true,
    readOnly: selection.draft.readOnly === true,
    ...(selection.draft.defaultValue !== undefined
      ? { defaultValue: selection.draft.defaultValue }
      : {}),
    ...(selection.draft.exportValue
      ? { exportValue: selection.draft.exportValue }
      : {}),
    ...(selection.kind === "radio" ||
    selection.kind === "dropdown" ||
    selection.kind === "listbox"
      ? {
          options:
            selection.kind === "radio"
              ? [...selection.options]
              : [...(selection.draft.options ?? [])],
        }
      : {}),
    ...(selection.kind === "listbox"
      ? { selectedValues: [...selection.selectedValues] }
      : {}),
    ...(selection.kind === "checkbox" ? { checked: selection.checked } : {}),
    ...(selection.kind === "dropdown" ? { editable: selection.editable } : {}),
    ...(selection.kind === "listbox"
      ? { multiselect: selection.multiselect }
      : {}),
    ...(selection.kind === "radio"
      ? {
          selectedValues: selection.selectedValue
            ? [selection.selectedValue]
            : [],
          offToggleable: selection.offToggleable,
        }
      : {}),
    ...(selection.kind === "text"
      ? {
          maxLength: selection.draft.maxLength ?? 0,
          multiline: selection.draft.multiline === true,
          password: selection.draft.password === true,
          comb: selection.draft.comb === true,
          spellCheck: selection.draft.spellCheck !== false,
          scrollable: selection.draft.scrollable !== false,
          alignment: selection.draft.alignment ?? "left",
          fontSize: selection.draft.fontSize ?? 0,
          textColor: selection.draft.textColor ?? "#000000",
          backgroundColor: selection.draft.backgroundColor ?? "#FFFFFF",
          borderColor: selection.draft.borderColor ?? "#000000",
          borderWidth: selection.draft.borderWidth ?? 1,
          visibility: selection.draft.visibility ?? "visible",
          // A format the dialog cannot show (formatLocked) is left as the
          // field has it.
          ...(selection.draft.formatLocked
            ? {}
            : { formatScripts: pdfTextFieldFormatTriggers(selection.draft.format) }),
        }
      : {}),
    ...(selection.kind === "resetButton" ||
    selection.kind === "printButton" ||
    selection.kind === "navigationButton" ||
    selection.kind === "viewButton" ||
    selection.kind === "saveButton" ||
    selection.kind === "goToPageButton" ||
    selection.kind === "webLinkButton"
      ? { buttonLabel: selection.buttonLabel }
      : {}),
    ...(selection.kind === "navigationButton"
      ? { navigationAction: selection.navigationAction }
      : {}),
    ...(selection.kind === "viewButton"
      ? { viewAction: selection.viewAction }
      : {}),
    ...(selection.kind === "webLinkButton"
      ? { buttonUrl: sanitizeLinkUrl(selection.buttonUrl) ?? "" }
      : {}),
  };
}

const FLAG_ONLY_RECT_TOLERANCE = 1e-9;













export function flagOnlyFormFieldEdit(
  initial: NativePdfFormFieldSpec,
  next: NativePdfFormFieldSpec,
): { required: boolean; readOnly: boolean } | null {
  const rest = (spec: NativePdfFormFieldSpec) => {
    const copy: Record<string, unknown> = { ...spec };
    delete copy.required;
    delete copy.readOnly;
    delete copy.rect;
    // The role apart: the key list below names the spec's own keys only, and
    // filtered the role's colour out -- a role recoloured read as unchanged.
    delete copy.role;
    return `${JSON.stringify(copy, Object.keys(copy).sort())}|${JSON.stringify(spec.role ?? null)}`;
  };
  const sameRect = (["x", "y", "width", "height"] as const).every(
    (key) =>
      Math.abs(initial.rect[key] - next.rect[key]) <= FLAG_ONLY_RECT_TOLERANCE,
  );
  if (!sameRect || rest(initial) !== rest(next)) return null;
  return { required: next.required, readOnly: next.readOnly };
}
