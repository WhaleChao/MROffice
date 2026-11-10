/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { Node as ProseMirrorNode } from "@tiptap/pm/model";
import { EditorState, Plugin, PluginKey, type Transaction } from "@tiptap/pm/state";
import { Mapping, type StepMap } from "@tiptap/pm/transform";

let installed = false;

/** Install it, once per page (idempotent). */
export function installProseMirrorResilience(): void {
  if (installed) return;
  installed = true;
  installStaleTransactionRebase();
}

/** How many of the editor's last document versions a stale transaction may
 * be made on and still be rebased. */
const KEPT_VERSIONS = 32;

interface AppliedVersion {
  readonly before: ProseMirrorNode;
  readonly maps: readonly StepMap[];
}

const versionsKey = new PluginKey<readonly AppliedVersion[]>("staleTransactionRebase");

/** The versions an editor's state keeps for `installStaleTransactionRebase`. */
export function staleTransactionRebasePlugin(): Plugin<readonly AppliedVersion[]> {
  return new Plugin<readonly AppliedVersion[]>({
    key: versionsKey,
    state: {
      init: () => [],
      apply: (tr, versions) => tr.docChanged
        ? [...versions.slice(-(KEPT_VERSIONS - 1)), { before: tr.before, maps: tr.mapping.maps }]
        : versions,
    },
  });
}

/** `tr` (made on an earlier version `state` kept) made again on `state`. */
export function rebaseStaleTransaction(state: EditorState, tr: Transaction): Transaction | null {
  const versions = versionsKey.getState(state);
  if (!versions) return null;
  const from = versions.findIndex((version) => version.before === tr.before);
  if (from < 0) return null;
  const next = state.tr;
  const own = tr.steps.map((step) => step.getMap());
  // tr's document back to the version it was made on, then everything the
  // editor applied since: each step is mapped through the rest (its own
  // inverse mirrored, as prosemirror-collab rebases).
  const mapping = new Mapping();
  for (let index = own.length - 1; index >= 0; index -= 1) mapping.appendMap(own[index].invert());
  for (const version of versions.slice(from)) for (const map of version.maps) mapping.appendMap(map);
  tr.steps.forEach((step, index) => {
    const mapped = step.map(mapping.slice(own.length - index));
    if (mapped && !next.maybeStep(mapped).failed) mapping.appendMap(mapped.getMap(), own.length - index - 1);
  });
  const meta = (tr as unknown as { meta?: Record<string, unknown> }).meta ?? {};
  for (const [key, value] of Object.entries(meta)) next.setMeta(key, value);
  if (tr.selectionSet) {
    try { next.setSelection(tr.selection.map(next.doc, mapping)); } catch { /* the mapped selection stays */ }
  }
  if (tr.storedMarksSet) next.setStoredMarks(tr.storedMarks);
  if (tr.scrolledIntoView) next.scrollIntoView();
  next.setTime(tr.time);
  return next;
}

let reportedRebase = false;

function installStaleTransactionRebase(): void {
  const original = EditorState.prototype.applyTransaction;
  EditorState.prototype.applyTransaction = function applyTransaction(this: EditorState, rootTr: Transaction) {
    if (rootTr.before !== this.doc && !rootTr.before.eq(this.doc)) {
      const rebased = rebaseStaleTransaction(this, rootTr);
      if (rebased) {
        if (!reportedRebase) {
          reportedRebase = true;
          console.warn("[editor] a transaction made on an earlier document was rebased onto the current one");
        }
        return original.call(this, rebased);
      }
    }
    return original.call(this, rootTr);
  };
}
