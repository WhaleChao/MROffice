/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Table Design > Table Color: recolor the selected table's style to the same
 * style family in another accent (or neutral), for a table whose look is a
 * native table-design authority as for one carrying a local Office style. The
 * ribbon's Table Color menu (DocumentOfficeTableColorMenu) and the AI's
 * document.table.style_color command call these, so both recolor the same way.
 */

import type { Editor } from "@tiptap/core";
import {
  documentOfficeTableStyleForId,
  readDocumentTableDesignTarget,
  setDocumentTableDesignState,
  type DocumentTableDesignState,
} from "@mrnote/core/web/lib/document-table-design-authority.ts";
import {
  officeWordTableStyleRecolor,
  type OfficeWordTableStyle,
  type OfficeWordTableStyleColor,
} from "@mrnote/core/web/lib/office-table-style-word.ts";
import {
  applyDocumentOfficeTableStyle,
  documentOfficeTableStyleIdForSlot,
  findSelectedDocumentTable,
  localOfficeTableStyleId,
} from "@mrnote/core/web/components/console/editor/native-document/documentOfficeTableStyles.ts";
import { setDocumentTableStylePreview } from "@mrnote/core/web/components/console/editor/native-document/tableBorders.ts";

export interface DocumentOfficeTableColorTarget {
  readonly style: OfficeWordTableStyle;
  readonly kind: "native" | "local";
}

export function readDocumentOfficeTableColorTarget(
  editor: Editor,
): DocumentOfficeTableColorTarget | null {
  const native = readDocumentTableDesignTarget(editor);
  if (native) {
    if (!native.authority.receipt.editable) return null;
    const style = documentOfficeTableStyleForId(
      native.doc,
      native.authority.current.styleId,
    );
    return style ? { style, kind: "native" } : null;
  }
  const table = findSelectedDocumentTable(editor);
  if (!table) return null;
  const styleId = localOfficeTableStyleId(editor.state.doc, table.node);
  const style = styleId
    ? documentOfficeTableStyleForId(editor.state.doc, styleId)
    : null;
  return style ? { style, kind: "local" } : null;
}

export function nativeDocumentTableColorState(
  editor: Editor,
  slot: number,
): { state: DocumentTableDesignState; apply: () => boolean; preview: () => boolean } | null {
  const target = readDocumentTableDesignTarget(editor);
  const styleId = documentOfficeTableStyleIdForSlot(editor.state.doc, slot);
  if (!target || !styleId) return null;
  const state = { ...target.authority.current, styleId };
  return {
    state,
    apply: () => setDocumentTableDesignState(editor, target, state),
    preview: () => setDocumentTableStylePreview(editor, target, state),
  };
}

export function applyDocumentOfficeTableColor(
  editor: Editor,
  color: OfficeWordTableStyleColor,
): boolean {
  const target = readDocumentOfficeTableColorTarget(editor);
  const next = target ? officeWordTableStyleRecolor(target.style, color) : null;
  if (!target || !next || next.slot === target.style.slot) return false;
  if (target.kind === "native") return nativeDocumentTableColorState(editor, next.slot)?.apply() ?? false;
  return applyDocumentOfficeTableStyle(editor, next.slot);
}
