"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Excel's contextual Equation tab (公式), shown while an equation text box is
 * selected or open in place: the shared Equation tab
 * (`OfficeEquationContextRibbon` -- Input Format Unicode / LaTeX, Layout
 * Professional / Linear, Symbols, Structures) driven through the worksheet's
 * draft adapter, after Office's Tools group (the Equation gallery, Normal
 * Text). The same tab a slide's equation box shows, under the worksheet's
 * control ids.
 */

import { useMemo } from "react";

import { equationRibbonCopy } from "@mrnote/core/web/components/console/editor/native-document/equationRibbonCopy.ts";
import { OfficeEquationContextRibbon } from "@mrnote/core/web/components/office-equation/OfficeEquationContextRibbon.tsx";
import { OfficeEquationToolsGroup } from "@mrnote/core/web/components/office-equation/OfficeEquationToolsGroup.tsx";
import { officeEquationDraftRibbonAdapter } from "@mrnote/core/web/components/office-equation/officeEquationDraftAdapter.ts";

import type { SpreadsheetEquations } from "./useSpreadsheetEquations";

export function SpreadsheetEquationContextRibbon({
  t,
  locale,
  disabled,
  equations,
}: {
  readonly t: (key: string, values?: Record<string, string>) => string;
  readonly locale: string;
  readonly disabled: boolean;
  readonly equations: SpreadsheetEquations;
}) {
  const copy = useMemo(() => equationRibbonCopy(t as (key: string) => string), [t]);
  const { inputMode, layoutMode, setEditorInputMode, setEditorLayoutMode, runInEditor, focusDraft } = equations;
  const adapter = useMemo(
    () =>
      officeEquationDraftRibbonAdapter({
        disabled,
        inputMode,
        layoutMode,
        onInputMode: setEditorInputMode,
        onLayoutMode: setEditorLayoutMode,
        runInEditor,
        onRefocusDraft: focusDraft,
      }),
    [disabled, focusDraft, inputMode, layoutMode, runInEditor, setEditorInputMode, setEditorLayoutMode],
  );
  const tools = (
    <OfficeEquationToolsGroup
      t={t}
      locale={locale}
      disabled={disabled}
      layoutMode={layoutMode}
      selectedEquationLatex={equations.selectedEquationLatex}
      onInsertEquation={equations.insertEquation}
      onInsertNewEquation={equations.insertNewEquation}
      normalText={equations.normalText}
      onToggleNormalText={equations.toggleNormalText}
      groupLabel={t("presentation.equation.toolsGroup")}
      equationLabel={copy.equation}
      normalTextLabel={t("presentation.equation.normalText")}
      controlIds={{
        gallery: "XLSCTL.CONTEXT.EQUATION.GALLERY",
        normalText: "XLSCTL.CONTEXT.EQUATION.NORMALTEXT",
      }}
      testIdPrefix="spreadsheet-equation"
    />
  );
  return (
    <OfficeEquationContextRibbon
      adapter={adapter}
      t={t}
      testIdPrefix="spreadsheet-equation"
      symbolsMoreControlId="XLSCTL.CONTEXT.EQUATION.SYMBOLS.MORE"
      leading={tools}
    />
  );
}
