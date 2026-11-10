/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { useAutosavedPageDocument } from "@mrnote/core/web/hooks/useAutosavedPageDocument.ts";

/** A private native editor. The parent owns its draft, history and eventual
 * apply decision. It is never a page, collaboration room or save receipt. */
export interface EmbeddedPresentationSession {
  content: Record<string, unknown>;
  revision: number;
  onChange(content: Record<string, unknown>): boolean;
  readOnlyReason?: string;
  undo(): void;
  redo(): void;
  canUndo: boolean;
  canRedo: boolean;
  onReady?(): void;
  /** View-only isolation; retain the full native slide's layout dependencies. */
  previewBlockId?: string;
  /** Atomic native diagrams paint all of their members in one object viewport. */
  previewBlockIds?: readonly string[];
  /** User-triggered request to open the existing native SmartArt text pane. */
  previewTextEditRequest?: number;
  /** The slide a preview shows (a live edit follows the slide it changes). */
  previewFocusSlideId?: string;
  /** Laid over the real slide at its drawn size (the Copilot's in-place
   * review): the slide fills the viewport, with none of the window's fit
   * margin, so it lands on the slide it changes. */
  fillViewport?: boolean;
}

type PageAutosave = ReturnType<typeof useAutosavedPageDocument>;

/** The page hook must ALSO be disabled at construction. This wrapper is the
 * private callback boundary, not a way to downgrade an active page's ACL. */
export function embeddedPresentationAutosave(
  disabledPage: PageAutosave,
  readSession: () => EmbeddedPresentationSession,
): Omit<PageAutosave, "readOnlyReason"> & { readOnlyReason: string | null } {
  const session = readSession();
  return {
    ...disabledPage,
    page: null,
    loading: false,
    readOnly: Boolean(session.readOnlyReason),
    readOnlyReason: session.readOnlyReason ?? null,
    saveStatus: "unsaved",
    collaborationHistory: undefined,
    flushCollaboration: undefined,
    queueSave(draft) {
      const current = readSession();
      if (current.readOnlyReason || !draft.content_json) return;
      current.onChange(draft.content_json);
    },
    async saveNow(draft) {
      const current = readSession();
      if (current.readOnlyReason) return null;
      if (draft?.content_json && !current.onChange(draft.content_json))
        throw new Error("embedded_presentation_not_committed");
      return null;
    },
  };
}
