/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
export type CodeSplitPane = "primary" | "secondary";

export interface CodeSplitViewState {
  readonly secondaryOpen: boolean;
  readonly activePane: CodeSplitPane;
}

const PRIMARY_ONLY_PANES = ["primary"] as const;
const SPLIT_PANES = ["primary", "secondary"] as const;

export function createCodeSplitViewState(): CodeSplitViewState {
  return { secondaryOpen: false, activePane: "primary" };
}

export function visibleCodeSplitPanes(
  state: CodeSplitViewState,
): readonly CodeSplitPane[] {
  return state.secondaryOpen ? SPLIT_PANES : PRIMARY_ONLY_PANES;
}

export function toggleCodeSplitView(
  state: CodeSplitViewState,
): CodeSplitViewState {
  return state.secondaryOpen
    ? { secondaryOpen: false, activePane: "primary" }
    : { secondaryOpen: true, activePane: state.activePane };
}

export function closeCodeSplitView(
  state: CodeSplitViewState,
): CodeSplitViewState {
  return state.secondaryOpen || state.activePane !== "primary"
    ? { secondaryOpen: false, activePane: "primary" }
    : state;
}

export function activateCodeSplitPane(
  state: CodeSplitViewState,
  pane: CodeSplitPane,
): CodeSplitViewState {
  if (pane === "secondary" && !state.secondaryOpen) return state;
  return state.activePane === pane ? state : { ...state, activePane: pane };
}

/**
 * The primary scope intentionally preserves the historical page key. The
 * secondary scope is stable across close/reopen, so each pane owns its own
 * selection, folds, and scroll offsets without forking document authority.
 */
export function codeSplitPaneViewStateKey(
  pageId: string,
  pane: CodeSplitPane,
): string {
  return pane === "primary"
    ? `page:${pageId}`
    : `page:${pageId}:split:secondary`;
}

export function codeSplitPaneEditorTestId(pane: CodeSplitPane): string {
  return pane === "primary" ? "code-editor" : "code-editor-secondary";
}
