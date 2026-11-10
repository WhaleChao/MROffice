"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * PowerPoint's contextual Equation tab (公式), shown while an equation on a
 * slide is open in place (an equation box), open in the editor, or selected
 * (PPT-A052): Docs' own Equation tab (`OfficeEquationContextRibbon` -- Input
 * Format Unicode / LaTeX, Layout Professional / Linear, Symbols, Structures)
 * driven through the deck's adapter, after PowerPoint's Tools group (the
 * Equation gallery, Normal Text).
 *
 * The adapter runs every command on the open draft (`runInEditor`); with no
 * draft open it opens the selected equation first -- an equation box in
 * place, an equation among words in the editor -- and runs once the draft is
 * live. Every control is part of the open draft
 * (`VISUAL_MATH_EDITOR_FOCUS_OWNER_ATTRIBUTE`): pressing one never commits it.
 *
 * Normal Text (普通文本) switches the built-up draft's caret between math and
 * ordinary text (MathLive's text mode, `\text{...}` in the LaTeX), as Office's
 * toggle does; it is not Word's "Text" input format (a linear plain-text
 * source), which PowerPoint does not have.
 */

import { useMemo } from "react";

import type {
  EquationInputMode,
  EquationLayoutMode,
} from "@mrnote/core/web/components/console/editor/extensions/equationCatalog.ts";
import { equationRibbonCopy } from "@mrnote/core/web/components/console/editor/native-document/equationRibbonCopy.ts";
import {
  OfficeEquationContextRibbon,
  type OfficeEquationRibbonAdapter,
} from "@mrnote/core/web/components/office-equation/OfficeEquationContextRibbon.tsx";
import { OfficeEquationToolsGroup } from "@mrnote/core/web/components/office-equation/OfficeEquationToolsGroup.tsx";
import {
  EQUATION_BOX_INPUT_MODES,
  officeEquationDraftRibbonAdapter,
} from "@mrnote/core/web/components/office-equation/officeEquationDraftAdapter.ts";

import type { EquationEditorCommand } from "./usePresentationEquations";

/** PowerPoint's two linear formats (it has no plain-Text one). */
const PRESENTATION_EQUATION_INPUT_MODES: readonly EquationInputMode[] = EQUATION_BOX_INPUT_MODES;

export interface PresentationEquationContextRibbonProps {
  t: (key: string, values?: Record<string, string>) => string;
  locale: string;
  disabled: boolean;
  /** The editor's layout while a draft is open, else the default. */
  layoutMode: EquationLayoutMode;
  /** The open draft's linear input format, else the window's. */
  inputMode: EquationInputMode;
  /** The selected equation's LaTeX (Save Selection to Equation Gallery). */
  selectedEquationLatex: string | null;
  onInsertEquation: (latex: string) => void;
  onInsertNewEquation: () => void;
  onLayoutMode: (mode: EquationLayoutMode) => void;
  /** False: the draft has no spelling in that format; undefined: still opening. */
  onInputMode: (mode: EquationInputMode) => boolean | undefined;
  /** Normal Text is on at the open draft's caret. */
  normalText: boolean;
  onToggleNormalText: () => void;
  runInEditor: (command: EquationEditorCommand) => void;
  /** The keyboard back into the open draft (nothing when none is open). */
  onRefocusDraft: () => void;
}

/**
 * The deck's adapter for the shared Equation tab: the shared draft adapter
 * (`officeEquationDraftRibbonAdapter`, every command on the open draft) in
 * PowerPoint's two linear formats.
 */
export function presentationEquationRibbonAdapter({
  disabled,
  inputMode,
  layoutMode,
  onInputMode,
  onLayoutMode,
  runInEditor,
  onRefocusDraft,
}: Pick<
  PresentationEquationContextRibbonProps,
  "disabled" | "inputMode" | "layoutMode" | "onInputMode" | "onLayoutMode" | "runInEditor" | "onRefocusDraft"
>): OfficeEquationRibbonAdapter {
  return officeEquationDraftRibbonAdapter({
    disabled,
    inputMode,
    inputModes: PRESENTATION_EQUATION_INPUT_MODES,
    layoutMode,
    onInputMode,
    onLayoutMode,
    runInEditor,
    onRefocusDraft,
  });
}

export function PresentationEquationContextRibbon({
  t,
  locale,
  disabled,
  layoutMode,
  inputMode,
  selectedEquationLatex,
  onInsertEquation,
  onInsertNewEquation,
  onLayoutMode,
  onInputMode,
  normalText,
  onToggleNormalText,
  runInEditor,
  onRefocusDraft,
}: PresentationEquationContextRibbonProps) {
  const copy = useMemo(() => equationRibbonCopy(t as (key: string) => string), [t]);
  const adapter = useMemo(
    () =>
      presentationEquationRibbonAdapter({
        disabled,
        inputMode,
        layoutMode,
        onInputMode,
        onLayoutMode,
        runInEditor,
        onRefocusDraft,
      }),
    [disabled, inputMode, layoutMode, onInputMode, onLayoutMode, runInEditor, onRefocusDraft],
  );
  const tools = (
    <OfficeEquationToolsGroup
      t={t}
      locale={locale}
      disabled={disabled}
      layoutMode={layoutMode}
      selectedEquationLatex={selectedEquationLatex}
      onInsertEquation={onInsertEquation}
      onInsertNewEquation={onInsertNewEquation}
      normalText={normalText}
      onToggleNormalText={onToggleNormalText}
      groupLabel={t("presentation.equation.toolsGroup")}
      equationLabel={copy.equation}
      normalTextLabel={t("presentation.equation.normalText")}
      controlIds={{
        gallery: "PPTCTL.CONTEXT.EQUATION.GALLERY",
        normalText: "PPTCTL.CONTEXT.EQUATION.NORMALTEXT",
      }}
      testIdPrefix="presentation-equation"
    />
  );
  return (
    <OfficeEquationContextRibbon
      adapter={adapter}
      t={t}
      testIdPrefix="presentation-equation"
      symbolsMoreControlId="PPTCTL.CONTEXT.EQUATION.SYMBOLS.MORE"
      leading={tools}
    />
  );
}
