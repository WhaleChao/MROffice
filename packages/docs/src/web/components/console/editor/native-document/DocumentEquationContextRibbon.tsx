"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Docs' contextual Equation tab: the shared Office Equation tab
 * (`OfficeEquationContextRibbon`) driven by the selected ProseMirror equation
 * node. Input format, layout and placement are one ProseMirror transaction
 * each (`applyDocumentEquationCommand`); symbols and structures go into the
 * node's live MathLive draft through the NodeView seam and commit with it.
 */

import { useMemo, useState } from "react";
import type { Editor } from "@tiptap/core";
import { closeHistory } from "@tiptap/pm/history";
import type { Node as ProseMirrorNode } from "@tiptap/pm/model";
import { NodeSelection } from "@tiptap/pm/state";
import { useTranslations } from "@mrnote/core/host/next-intl.tsx";

import {
  normalizeEquationInputMode,
  normalizeEquationLayoutMode,
  type EquationInputMode,
  type EquationLayoutMode,
  type EquationStructureId,
} from "@mrnote/core/web/components/console/editor/extensions/equationCatalog.ts";
import { latexToEquationText } from "@mrnote/core/web/components/console/editor/extensions/equationLinearFormat.ts";
import { latexToReversibleWordLinear } from "@mrnote/core/web/lib/unicode-math/index.ts";
import {
  equationPlacementOfNodeName,
  setEquationPlacement,
  type EquationPlacement,
} from "@mrnote/core/web/components/console/editor/extensions/equationPlacement.ts";
import {
  requestVisualMathEditorFocus,
  requestVisualMathNormalText,
  requestVisualMathStructureInsert,
  requestVisualMathStructureMemberInsert,
  requestVisualMathSymbolInsert,
} from "@mrnote/core/web/components/console/editor/extensions/visualMathEditor.ts";
import {
  OfficeEquationContextRibbon,
  type OfficeEquationRibbonAdapter,
} from "@mrnote/core/web/components/office-equation/OfficeEquationContextRibbon.tsx";

// The tab's pieces live with the shared tab; Docs' callers keep importing
// them from here.
export {
  EquationStructureFamilyFace,
  equationSymbolLabel,
} from "@mrnote/core/web/components/office-equation/OfficeEquationContextRibbon.tsx";
import { openEquationEditors } from "@mrnote/core/web/components/console/editor/extensions/openEquationEditors.ts";
import { OfficeEquationToolsGroup } from "@mrnote/core/web/components/office-equation/OfficeEquationToolsGroup.tsx";
import type { EquationInsertPayload } from "@mrnote/core/web/components/console/editor/native-document/DocumentEquationGalleryMenu.tsx";

const EQUATION_NODE_TYPES = new Set(["mathBlock", "inlineMath"]);
/**
 * Word's two input formats, as every editor's Equation tab offers them;
 * Word's "abc Text" is the Tools group's Normal Text. An equation saved in
 * the older plain-Text format keeps it listed, so it shows and can leave it.
 */
const DOCUMENT_EQUATION_INPUT_MODES: readonly EquationInputMode[] = ["unicode", "latex"];
const DOCUMENT_EQUATION_INPUT_MODES_WITH_TEXT: readonly EquationInputMode[] = ["unicode", "latex", "text"];

export interface DocumentEquationSelection {
  pos: number;
  nodeType: "mathBlock" | "inlineMath";
  latex: string;
  inputMode: EquationInputMode;
  layoutMode: EquationLayoutMode;
}

export type DocumentEquationCommand =
  | { type: "inputMode"; value: EquationInputMode }
  | { type: "layoutMode"; value: EquationLayoutMode }
  | { type: "placement"; value: EquationPlacement }
  | { type: "symbol"; id: string }
  | { type: "structure"; id: EquationStructureId }
  | { type: "structureMember"; id: string };

export type DocumentEquationCommandResult =
  | { ok: true; selection: DocumentEquationSelection }
  | {
      ok: false;
      reason:
        | "stale_selection"
        | "readonly"
        | "unknown_structure_member"
        | "unknown_symbol"
        | "unsupported_unicode"
        | "unsupported_text"
        | "unsupported_placement";
    };

/**
 * The equation the contextual controls act on: an explicit atom
 * NodeSelection, or the equation open in its editor. Word shows the Equation
 * tab while the insertion point is inside an equation ("Write an equation or
 * formula", Microsoft Support); typing in the open editor moves the browser
 * selection into it, which ProseMirror reads back as a caret beside the atom,
 * so the open editor itself is the evidence. A caret merely adjacent to an
 * inline equation, with no editor open, must not switch tabs.
 */
export function readDocumentEquationSelection(
  editor: Editor,
): DocumentEquationSelection | null {
  // The equation open in its editor comes first: double-clicking an
  // equation in a table cell opens its editor while ProseMirror's selection
  // can still be the equation selected before it, and the ribbon's commands
  // then rewrote that other equation (live Word audit WUA 046).
  const selection = editor.state.selection;
  const selected =
    openEquationEditorTarget(editor) ??
    (selection instanceof NodeSelection &&
    EQUATION_NODE_TYPES.has(selection.node.type.name)
      ? { node: selection.node, pos: selection.from }
      : null);
  if (!selected) return null;
  const node = selected.node;
  return {
    pos: selected.pos,
    nodeType: node.type.name as "mathBlock" | "inlineMath",
    latex: String(node.attrs.latex ?? ""),
    inputMode: normalizeEquationInputMode(node.attrs.inputMode),
    layoutMode: normalizeEquationLayoutMode(node.attrs.layoutMode),
  };
}

/** The equation whose visual editor is open (its node view says so). */
function openEquationEditorTarget(
  editor: Editor,
): { node: ProseMirrorNode; pos: number } | null {
  let open: Element | null = null;
  try {
    const candidates = openEquationEditors(editor.view.dom);
    // The one the keyboard is in, when more than one is open.
    const active = editor.view.dom.ownerDocument.activeElement;
    open =
      candidates.find((candidate) => active instanceof Element && candidate.contains(active)) ??
      candidates[0] ??
      null;
  } catch {
    return null;
  }
  if (!open) return null;
  // The open editor's own node, from its DOM: walking the document asked the
  // view for the DOM of every formula before it (thousands in a math book).
  try {
    const pos = editor.view.posAtDOM(open, 0) - 1;
    const node = pos >= 0 ? editor.state.doc.nodeAt(pos) : null;
    if (node && EQUATION_NODE_TYPES.has(node.type.name) && editor.view.nodeDOM(pos) === open) return { node, pos };
  } catch {
    // The walk below answers.
  }
  let found: { node: ProseMirrorNode; pos: number } | null = null;
  editor.state.doc.descendants((node, pos) => {
    if (found) return false;
    if (EQUATION_NODE_TYPES.has(node.type.name) && editor.view.nodeDOM(pos) === open) {
      found = { node, pos };
      return false;
    }
    return true;
  });
  return found;
}

function isCurrentDocumentEquationSelection(
  editor: Editor,
  expected: DocumentEquationSelection,
): boolean {
  const activeSelection = readDocumentEquationSelection(editor);
  if (
    !activeSelection ||
    activeSelection.pos !== expected.pos ||
    activeSelection.nodeType !== expected.nodeType ||
    activeSelection.latex !== expected.latex ||
    activeSelection.inputMode !== expected.inputMode ||
    activeSelection.layoutMode !== expected.layoutMode
  ) {
    return false;
  }
  const current = editor.state.doc.nodeAt(expected.pos);
  return Boolean(
    current &&
    current.type.name === expected.nodeType &&
    String(current.attrs.latex ?? "") === expected.latex &&
    normalizeEquationInputMode(current.attrs.inputMode) ===
      expected.inputMode &&
    normalizeEquationLayoutMode(current.attrs.layoutMode) ===
      expected.layoutMode,
  );
}

/**
 * Contextual structures belong to the live MathLive draft. Activating one
 * opens/reuses the exact selected NodeView but deliberately performs no
 * ProseMirror transaction; Done or legal blur owns the one commit.
 */
export function applyDocumentEquationLiveStructureCommand(
  editor: Editor,
  expected: DocumentEquationSelection,
  structureId: EquationStructureId,
): DocumentEquationCommandResult {
  if (!editor.isEditable) return { ok: false, reason: "readonly" };
  if (!isCurrentDocumentEquationSelection(editor, expected)) {
    return { ok: false, reason: "stale_selection" };
  }
  const result = requestVisualMathStructureInsert(
    editor,
    expected,
    structureId,
  );
  if (!result.ok) return result;
  return { ok: true, selection: expected };
}

/** Resolve a frozen gallery member without accepting caller-provided LaTeX. */
export function applyDocumentEquationLiveStructureMemberCommand(
  editor: Editor,
  expected: DocumentEquationSelection,
  memberId: string,
): DocumentEquationCommandResult {
  if (!editor.isEditable) return { ok: false, reason: "readonly" };
  if (!isCurrentDocumentEquationSelection(editor, expected)) {
    return { ok: false, reason: "stale_selection" };
  }
  const result = requestVisualMathStructureMemberInsert(
    editor,
    expected,
    memberId,
  );
  if (!result.ok) return result;
  return { ok: true, selection: expected };
}

/** Resolve a frozen symbol identity into the current MathLive draft/cursor. */
export function applyDocumentEquationLiveSymbolCommand(
  editor: Editor,
  expected: DocumentEquationSelection,
  symbolId: string,
): DocumentEquationCommandResult {
  if (!editor.isEditable) return { ok: false, reason: "readonly" };
  if (!isCurrentDocumentEquationSelection(editor, expected)) {
    return { ok: false, reason: "stale_selection" };
  }
  const result = requestVisualMathSymbolInsert(editor, expected, symbolId);
  if (!result.ok) return result;
  return { ok: true, selection: expected };
}

/** Preserve the Phase 2 Fraction entry point while sharing its live seam. */
export function applyDocumentEquationLiveFractionCommand(
  editor: Editor,
  expected: DocumentEquationSelection,
): DocumentEquationCommandResult {
  return applyDocumentEquationLiveStructureCommand(
    editor,
    expected,
    "fraction",
  );
}

/**
 * Apply a non-structure equation operation with one ProseMirror transaction.
 * Structure commands delegate to the live NodeView seam above. Every
 * conversion and identity check completes before dispatch, so a rejected
 * Unicode/Text conversion or stale selection is guaranteed to mutate nothing.
 */
export function applyDocumentEquationCommand(
  editor: Editor,
  expected: DocumentEquationSelection,
  command: DocumentEquationCommand,
): DocumentEquationCommandResult {
  if (command.type === "structure") {
    return applyDocumentEquationLiveStructureCommand(
      editor,
      expected,
      command.id,
    );
  }
  if (command.type === "structureMember") {
    return applyDocumentEquationLiveStructureMemberCommand(
      editor,
      expected,
      command.id,
    );
  }
  if (command.type === "symbol") {
    return applyDocumentEquationLiveSymbolCommand(editor, expected, command.id);
  }
  if (!editor.isEditable) return { ok: false, reason: "readonly" };
  if (!isCurrentDocumentEquationSelection(editor, expected)) {
    return { ok: false, reason: "stale_selection" };
  }
  const current = editor.state.doc.nodeAt(expected.pos);
  if (!current) return { ok: false, reason: "stale_selection" };

  if (command.type === "placement") {
    // Word's "Change to Inline" / "Change to Display": the same equation in
    // its other wrapper. The conversion keeps the equation selected, so the
    // live selection after dispatch is the new identity.
    if (!setEquationPlacement(editor.view, expected.pos, command.value)) {
      return { ok: false, reason: "unsupported_placement" };
    }
    const converted = readDocumentEquationSelection(editor);
    return converted
      ? { ok: true, selection: converted }
      : { ok: false, reason: "stale_selection" };
  }

  const latex = expected.latex;
  let inputMode = expected.inputMode;
  let layoutMode = expected.layoutMode;

  if (command.type === "inputMode") {
    if (command.value === "unicode" && !latexToReversibleWordLinear(latex)) {
      return { ok: false, reason: "unsupported_unicode" };
    }
    if (command.value === "text" && latexToEquationText(latex) === null) {
      return { ok: false, reason: "unsupported_text" };
    }
    inputMode = command.value;
  } else if (command.type === "layoutMode") {
    layoutMode = command.value;
  }

  if (inputMode === "unicode" && !latexToReversibleWordLinear(latex)) {
    return { ok: false, reason: "unsupported_unicode" };
  }
  if (inputMode === "text" && latexToEquationText(latex) === null) {
    return { ok: false, reason: "unsupported_text" };
  }

  const attrs = {
    ...current.attrs,
    latex,
    inputMode,
    layoutMode,
  };
  let transaction = editor.state.tr.setNodeMarkup(
    expected.pos,
    undefined,
    attrs,
  );
  transaction = transaction.setSelection(
    NodeSelection.create(transaction.doc, expected.pos),
  );
  transaction = closeHistory(transaction);
  editor.view.dispatch(transaction);
  return {
    ok: true,
    selection: {
      pos: expected.pos,
      nodeType: expected.nodeType,
      latex,
      inputMode,
      layoutMode,
    },
  };
}

export function DocumentEquationContextRibbon({
  editor,
  selection,
  locale,
  onInsertEquation,
  onInsertNewEquation,
}: {
  editor: Editor;
  selection: DocumentEquationSelection;
  locale: string;
  /** Equation ▾: a gallery equation, placed as Insert > Equation places it. */
  onInsertEquation: (payload: EquationInsertPayload) => void;
  onInsertNewEquation: () => void;
}) {
  const t = useTranslations("console-notebooks");
  const editable = !editor.isDestroyed && editor.isEditable;
  // Normal Text belongs to one open draft: another equation, or this one
  // written back, types math again.
  const draftKey = `${selection.pos}:${selection.layoutMode}:${selection.latex}`;
  const [normalTextDraft, setNormalTextDraft] = useState<string | null>(null);
  const normalText = normalTextDraft === draftKey;
  const adapter = useMemo<OfficeEquationRibbonAdapter>(
    () => ({
      editable,
      inputMode: selection.inputMode,
      inputModes:
        selection.inputMode === "text" ? DOCUMENT_EQUATION_INPUT_MODES_WITH_TEXT : DOCUMENT_EQUATION_INPUT_MODES,
      layoutMode: selection.layoutMode,
      placement: equationPlacementOfNodeName(selection.nodeType) ?? undefined,
      run: (command) => {
        const result = applyDocumentEquationCommand(editor, selection, command);
        return result.ok ? { ok: true } : result;
      },
      refocusDraft: () => requestVisualMathEditorFocus(editor, selection),
    }),
    [editable, editor, selection],
  );
  const tools = (
    <OfficeEquationToolsGroup
      t={t as (key: string, values?: Record<string, string>) => string}
      locale={locale}
      disabled={!editable}
      layoutMode={selection.layoutMode}
      selectedEquationLatex={selection.latex}
      onInsertEquation={(_latex, payload) => onInsertEquation(payload)}
      onInsertNewEquation={onInsertNewEquation}
      normalText={normalText}
      onToggleNormalText={() => {
        const active = requestVisualMathNormalText(editor, selection);
        if (active !== null) setNormalTextDraft(active ? draftKey : null);
      }}
      groupLabel={t("presentation.equation.toolsGroup")}
      equationLabel={t("equation.ribbon.equation")}
      normalTextLabel={t("presentation.equation.normalText")}
      controlIds={{ gallery: "equation.tools.gallery", normalText: "equation.tools.normalText" }}
      testIdPrefix="document-equation"
    />
  );
  return (
    <OfficeEquationContextRibbon
      adapter={adapter}
      t={t as (key: string, values?: Record<string, string>) => string}
      testIdPrefix="document-equation"
      symbolsMoreControlId="equation.symbols.more"
      leading={tools}
    />
  );
}
