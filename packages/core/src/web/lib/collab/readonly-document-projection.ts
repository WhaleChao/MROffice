/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { yXmlFragmentToProsemirrorJSON } from "@tiptap/y-tiptap";
import type { CollabSession } from "./collab-provider";
import { flattenPagedDocumentForStorage } from "../../components/console/editor/native-document/paged/serialization";
import {
  COLLABORATIVE_DOCUMENT_ATTRS_MAP,
  readCollaborativeDocumentAttrs,
} from "../../components/console/editor/native-document/collaborationDocumentAttrs";

/** Read a room without lending its Y.Doc to a locally paginated editor. */
export function observeReadonlyDocumentProjection(
  session: CollabSession,
  onProjection: (content: Record<string, unknown>) => void,
): () => void {
  let closed = false;
  let queued = false;
  let lastProjection: string | null = null;
  const observedAttributeKeys = new Set<string>();
  const publish = () => {
    queued = false;
    if (
      closed ||
      session.network?.failed ||
      !session.network?.ready ||
      !session.network.hasState ||
      typeof session.ydoc.getXmlFragment !== "function"
    )
      return;
    const fragment = session.ydoc.getXmlFragment("default");
    if (fragment.length === 0) return;
    // Room page wrappers can change once per physical page during reflow.
    // Compare semantic content so that those events do not repeatedly replace
    // the viewer's document or disturb its local selection/scroll.
    const content = flattenPagedDocumentForStorage(
      yXmlFragmentToProsemirrorJSON(fragment),
      { paginationSettled: false },
    ) as Record<string, unknown>;
    const attrs = session.ydoc.getMap(COLLABORATIVE_DOCUMENT_ATTRS_MAP);
    const currentAttrs = readCollaborativeDocumentAttrs(attrs);
    for (const key of Object.keys(currentAttrs)) observedAttributeKeys.add(key);
    if (attrs.size > 0 || observedAttributeKeys.size > 0) {
      // Preserve an explicit remote removal instead of resurrecting the
      // original saved attribute when the caller overlays its source metadata.
      content.attrs = Object.fromEntries(
        [...observedAttributeKeys].map((key) => [
          key,
          currentAttrs[key] ?? null,
        ]),
      );
    }
    const encoded = JSON.stringify(content);
    if (encoded === lastProjection) return;
    lastProjection = encoded;
    onProjection(content);
  };
  const schedule = () => {
    if (closed || queued) return;
    queued = true;
    queueMicrotask(publish);
  };
  session.ydoc.on("update", schedule);
  const unsubscribeConnection =
    session.network?.subscribeConnection?.(schedule);
  void session.network?.whenReady.then(schedule).catch(() => undefined);
  schedule();
  return () => {
    closed = true;
    session.ydoc.off("update", schedule);
    unsubscribeConnection?.();
  };
}
