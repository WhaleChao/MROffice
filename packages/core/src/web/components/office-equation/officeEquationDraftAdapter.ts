/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The shared Equation tab (`OfficeEquationContextRibbon`) driven through an
 * editor's open DRAFT: the adapter every host whose equations are edited in
 * the shared MathLive controller builds (a slide's equation box, a
 * worksheet's equation text box). Every command runs on the open draft
 * (`runInEditor`); with no draft open the host opens the selected equation
 * first and runs the command once the draft is live (`wasOpen` false: the
 * caret is then at the draft's end, not where the person left it).
 */

import type {
  EquationInputMode,
  EquationLayoutMode,
} from "../console/editor/extensions/equationCatalog";
import type {
  VisualMathEditorController,
  VisualMathStructureInsertResult,
} from "../console/editor/extensions/visualMathEditor";

import type { OfficeEquationRibbonAdapter, OfficeEquationRibbonResult } from "./OfficeEquationContextRibbon";

/** An Equation-tab command on the live draft. */
export type EquationDraftCommand = (controller: VisualMathEditorController, wasOpen: boolean) => void;

/** Office's two linear formats for an equation box (no plain-Text one). */
export const EQUATION_BOX_INPUT_MODES: readonly EquationInputMode[] = ["unicode", "latex"];

export interface OfficeEquationDraftAdapterOptions {
  readonly disabled: boolean;
  /** The open draft's linear input format, else the host's. */
  readonly inputMode: EquationInputMode;
  readonly inputModes?: readonly EquationInputMode[];
  readonly layoutMode: EquationLayoutMode;
  /** False: the draft has no spelling in that format; undefined: still opening. */
  readonly onInputMode: (mode: EquationInputMode) => boolean | undefined;
  readonly onLayoutMode: (mode: EquationLayoutMode) => void;
  readonly runInEditor: (command: EquationDraftCommand) => void;
  /** The keyboard back into the open draft (nothing when none is open). */
  readonly onRefocusDraft: () => void;
}

/** A host's adapter for the shared Equation tab: every command on the open draft. */
export function officeEquationDraftRibbonAdapter({
  disabled,
  inputMode,
  inputModes = EQUATION_BOX_INPUT_MODES,
  layoutMode,
  onInputMode,
  onLayoutMode,
  runInEditor,
  onRefocusDraft,
}: OfficeEquationDraftAdapterOptions): OfficeEquationRibbonAdapter {
  /** A catalog insertion into the draft; its answer when the draft was already open. */
  const insertIntoDraft = (
    insert: (controller: VisualMathEditorController, wasOpen: boolean) => VisualMathStructureInsertResult,
  ): OfficeEquationRibbonResult => {
    let result: OfficeEquationRibbonResult = undefined;
    runInEditor((controller, wasOpen) => {
      // A draft the tab just opened has its caret at the end.
      const answer = insert(controller, wasOpen);
      result = answer.ok ? { ok: true } : { ok: false, reason: answer.reason };
      controller.focus();
    });
    return result;
  };
  return {
    editable: !disabled,
    inputMode,
    inputModes,
    layoutMode,
    run: (command) => {
      switch (command.type) {
        case "inputMode": {
          const applied = onInputMode(command.value);
          if (applied === undefined) return;
          return applied ? { ok: true } : { ok: false, reason: "unsupported_unicode" };
        }
        case "layoutMode":
          onLayoutMode(command.value);
          return;
        case "symbol":
          return insertIntoDraft((controller, wasOpen) =>
            controller.insertSymbol(command.id, wasOpen ? "selection" : "end"),
          );
        case "structure":
          return insertIntoDraft((controller, wasOpen) =>
            controller.insertStructure(command.id, wasOpen ? "selection" : "end"),
          );
        case "structureMember":
          return insertIntoDraft((controller, wasOpen) =>
            controller.insertStructureMember(command.id, wasOpen ? "selection" : "end"),
          );
        case "placement":
          // An equation box's equation is its own line by definition.
          return { ok: false, reason: "unsupported_placement" };
      }
    },
    refocusDraft: onRefocusDraft,
  };
}
