/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { EditorState } from "@tiptap/pm/state";
import { CellSelection } from "@tiptap/pm/tables";

/**
 * Whether the table commands can apply to the selection at all. Each of them
 * (Merge Cells, Split Cells, Split Table) asks for a cell selection or a
 * table around the selection's start, and refuses anywhere else; a toolbar
 * read need not dry-run them outside a table. Each dry run built every
 * command of every extension again (a stroke landing on a 3,000-stroke note
 * reads the toolbar twice). `__mraiToolbarTableDryRun` answers true always.
 */
export function selectionMayUseTableCommands(state: EditorState): boolean {
  if ((globalThis as { __mraiToolbarTableDryRun?: boolean }).__mraiToolbarTableDryRun) return true;
  try {
    const selection = state.selection;
    if (selection instanceof CellSelection) return true;
    const { $from } = selection;
    for (let depth = $from.depth; depth >= 0; depth -= 1) {
      if ($from.node(depth).type.name === "table") return true;
    }
    return false;
  } catch {
    return true;
  }
}
