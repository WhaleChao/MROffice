/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type * as Y from "yjs";

/**
 * A version of a Y.Doc's content, for reusing what was read from it.
 *
 * Every change of a Y.Doc runs in a transaction. The revision advances when
 * a transaction OPENS, not when it ends: Yjs calls observers (y-prosemirror's
 * sync, our observeDeep handlers) before `afterTransaction`, and a
 * transaction an observer starts is applied before its own cleanup, so a
 * revision bumped at the end let a read inside an observer reuse what was
 * read before the change. While a transaction is open the content is in
 * flux and there is no revision at all (null): nothing read then is reused.
 * The shared root types are part of the key: a root is created outside any
 * transaction, and a projection refuses an unknown root.
 */
const revisions = new WeakMap<Y.Doc, { value: number }>();

/** The current content version of `doc` (counting starts at the first call),
 * or null while one of its transactions is open. */
export function yDocRevision(doc: Y.Doc): string | null {
  let entry = revisions.get(doc);
  if (!entry) {
    const created = { value: 0 };
    doc.on("beforeTransaction", () => {
      created.value += 1;
    });
    revisions.set(doc, created);
    entry = created;
  }
  if ((doc as unknown as { _transaction: unknown })._transaction) return null;
  return `${entry.value}:${doc.share.size}`;
}

/**
 * A version of ONE shared type's content (a type and everything nested in
 * it), for reusing what was read from that type while the rest of the
 * document changes (typing does not touch the document attributes map).
 *
 * Like yDocRevision, a transaction takes a fresh version when it OPENS and
 * there is no version while it is open. When it ends without changing the
 * type, and no later transaction has opened since, the version it replaced
 * comes back: the type's content is what it was under that version. Yjs can
 * interleave transactions (one an observer starts opens before the first one
 * ends), which is why only the latest opened transaction restores.
 */
interface TypeRevisionState {
  value: number;
  opened: WeakMap<Y.Transaction, { before: number; own: number }>;
}
const typeRevisions = new WeakMap<object, TypeRevisionState>();
let nextTypeRevision = 1;

/** Whether `transaction` changed `type` or anything nested in it. */
function changedType(transaction: Y.Transaction, type: object): boolean {
  for (const changed of transaction.changed.keys()) if (changed === type) return true;
  for (const parent of transaction.changedParentTypes.keys()) if (parent === type) return true;
  return false;
}

export function yTypeRevision<EventType>(type: Y.AbstractType<EventType>): number | null {
  const doc = type.doc;
  if (!doc) return null;
  let state = typeRevisions.get(type);
  if (!state) {
    const created: TypeRevisionState = { value: nextTypeRevision++, opened: new WeakMap() };
    doc.on("beforeTransaction", (transaction: Y.Transaction) => {
      const own = nextTypeRevision++;
      created.opened.set(transaction, { before: created.value, own });
      created.value = own;
    });
    doc.on("afterTransaction", (transaction: Y.Transaction) => {
      const opened = created.opened.get(transaction);
      if (!opened || created.value !== opened.own) return;
      if (!changedType(transaction, type)) created.value = opened.before;
    });
    typeRevisions.set(type, created);
    state = created;
  }
  if ((doc as unknown as { _transaction: unknown })._transaction) return null;
  return state.value;
}
