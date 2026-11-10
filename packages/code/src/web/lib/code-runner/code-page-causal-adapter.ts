/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { PageContentDocument } from "@mrnote/core/web/hooks/usePageContentSync.ts";
import { applyJsonDelta, jsonEquals, yToJson } from "@mrnote/core/web/lib/collab/json-crdt.ts";
import { CODE_COLLAB_TEXT_KEY } from "@mrnote/core/web/lib/collab/code-collab-model.ts";
import { createNativePageCausalAdapter, type NativePageCausalHost } from "@mrnote/core/web/lib/collab/native-page-causal-adapter.ts";

/** Code retains its original text/metadata semantics while sharing the
 * registration, custody and acknowledgement lifecycle with native Docs. */
export function createCodePageCausalAdapter(host: NativePageCausalHost) {
  return createNativePageCausalAdapter(host, {
    model: "code-text-v1",
    projection: doc => ({ ...(yToJson(doc.getMap("page-json")) as PageContentDocument),
      content: doc.getText(CODE_COLLAB_TEXT_KEY).toString() }),
    canMirror: content => typeof content.content === "string",
    apply(doc, content, baseline) {
      const root = doc.getMap("page-json");
      const { content: nextText, ...next } = content;
      const { content: priorText, ...prior } = baseline;
      if (typeof nextText !== "string") return false;
      let changed = false;
      changed = applyJsonDelta(root, next, prior);
      const text = doc.getText(CODE_COLLAB_TEXT_KEY);
      // The editor may already have installed this gesture synchronously.
      // A stale shell snapshot must not overwrite a newer native text value.
      if (nextText !== priorText && text.toString() === priorText) {
        changed = true;
        let start = 0,
          end = 0;
        const old = text.toString();
        while (
          start < old.length &&
          start < nextText.length &&
          old[start] === nextText[start]
        )
          start++;
        if (
          start > 0 &&
          /[\uDC00-\uDFFF]/.test(old[start] ?? "") &&
          /[\uD800-\uDBFF]/.test(old[start - 1])
        )
          start--;
        while (
          end < old.length - start &&
          end < nextText.length - start &&
          old[old.length - 1 - end] === nextText[nextText.length - 1 - end]
        )
          end++;
        if (
          end > 0 &&
          /[\uDC00-\uDFFF]/.test(old[old.length - end] ?? "") &&
          /[\uD800-\uDBFF]/.test(old[old.length - end - 1] ?? "")
        )
          end--;
        text.delete(start, old.length - start - end);
        text.insert(start, nextText.slice(start, nextText.length - end));
      }
      return changed;
    },
    shouldDeliver: (session, transaction) => !transaction.local ||
      [...transaction.changed.keys()].some(type => Object.is(type, session.ydoc.getMap("page-json"))),
    seedLegacy(session, initial, first) {
      const { content, ...metadata } = initial;
      const text = session.ydoc.getText(CODE_COLLAB_TEXT_KEY), root = session.ydoc.getMap("page-json");
      const empty = !session.network?.hasState && session.ydoc.store.clients.size === 0;
      if (empty && !first) throw Error("Only the elected first peer can seed");
      if (!empty && text.toString() !== content) throw Error("Room differs from durable source");
      if (root.size && !jsonEquals(yToJson(root), metadata)) throw Error("Room metadata differs from durable source");
      if (!root.size) applyJsonDelta(root, metadata, {});
      if (empty && typeof content === "string" && content) text.insert(0, content);
    },
  });
}
