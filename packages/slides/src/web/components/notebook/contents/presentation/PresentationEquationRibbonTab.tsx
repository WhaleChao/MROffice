"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The Equation contextual tab of the presentation Ribbon, drawn from the
 * window bindings it reads.
 *
 * Extracted verbatim from the render of
 * `components/notebook/contents/PresentationWindow.tsx`: the JSX is unchanged;
 * the window passes the bindings it used to close over as `ctl`, and the tab
 * condition and key stay at the call site. No hooks run here, so the window's
 * hook order is unchanged.
 */

import { PresentationEquationContextRibbon } from "./PresentationEquationContextRibbon";
import type { useTranslations } from "@mrnote/core/host/next-intl.tsx";
import type { embeddedPresentationAutosave } from "../../../../lib/presentation-embedded-session";
import type { PresentationEquationSpan } from "../../../../lib/presentation-equation-buttons";
import type { EquationEditorCommand, PresentationActiveEquation } from "./usePresentationEquations";
import type { EquationInputMode, EquationLayoutMode } from "@mrnote/core/web/components/console/editor/extensions/equationCatalog.ts";
import type { JSX } from "react/jsx-runtime";

/** The window bindings this tab reads, passed through unchanged. */
export interface PresentationEquationRibbonTabCtl {
  t: ReturnType<typeof useTranslations<"console-notebooks">>;
  locale: string;
  autosave: ReturnType<typeof embeddedPresentationAutosave>;
  presentationEquations: { insertEquation: (latex: string) => void; insertEquationInBox: (latex: string, box: { left: number; top: number; width: number; height: number; }) => void; insertNewEquation: () => void; editEquation: (blockId: string, span: PresentationEquationSpan) => void; activeEquation: PresentationActiveEquation | null; selectedEquationLatex: string | null; editorOpen: boolean; layoutMode: EquationLayoutMode; setEditorLayoutMode: (mode: EquationLayoutMode) => void; inputMode: EquationInputMode; setEditorInputMode: (mode: EquationInputMode) => boolean | undefined; focusDraft: () => void; normalText: boolean; toggleNormalText: () => void; runInEditor: (command: EquationEditorCommand) => void; dialog: JSX.Element | null; };
}

export function PresentationEquationRibbonTab({ ctl }: { ctl: PresentationEquationRibbonTabCtl }) {
  const {
    t,
    locale,
    autosave,
    presentationEquations,
  } = ctl;
  return (
                  <PresentationEquationContextRibbon
                    t={t as (key: string, values?: Record<string, string>) => string}
                    locale={locale}
                    disabled={autosave.readOnly}
                    layoutMode={presentationEquations.layoutMode}
                    inputMode={presentationEquations.inputMode}
                    onInputMode={presentationEquations.setEditorInputMode}
                    onRefocusDraft={presentationEquations.focusDraft}
                    selectedEquationLatex={presentationEquations.selectedEquationLatex}
                    onInsertEquation={presentationEquations.insertEquation}
                    onInsertNewEquation={presentationEquations.insertNewEquation}
                    onLayoutMode={presentationEquations.setEditorLayoutMode}
                    normalText={presentationEquations.normalText}
                    onToggleNormalText={presentationEquations.toggleNormalText}
                    runInEditor={presentationEquations.runInEditor}
                  />
  );
}
