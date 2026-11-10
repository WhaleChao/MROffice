/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { Transaction } from "@codemirror/state";

export type SignatureHelpTrigger = "open" | "close" | null;

/** What a single transaction means for signature help. */
export function signatureHelpTriggerFor(
  transaction: Transaction,
): SignatureHelpTrigger {
  if (!transaction.isUserEvent("input.type")) return null;
  const { state } = transaction;
  const main = state.selection.main;
  if (!main.empty || main.head === 0) return null;
  const typed = state.doc.sliceString(main.head - 1, main.head);
  if (typed !== "(" && typed !== "," && typed !== ")") return null;
  if (transaction.docChanged) {
    // the character before the caret must be one this transaction inserted
    let inserted = false;
    transaction.changes.iterChanges((_fromA, _toA, fromB, toB) => {
      if (fromB < main.head && main.head - 1 < toB) inserted = true;
    });
    if (!inserted) return null;
  } else if (typed !== ")") {
    // only closeBrackets' step over a `)` types without changing the text
    return null;
  }
  return typed === ")" ? "close" : "open";
}

/** The last trigger among an update's transactions (null when none). */
export function signatureHelpTriggerForUpdate(
  transactions: readonly Transaction[],
): SignatureHelpTrigger {
  let trigger: SignatureHelpTrigger = null;
  for (const transaction of transactions) {
    trigger = signatureHelpTriggerFor(transaction) ?? trigger;
  }
  return trigger;
}
