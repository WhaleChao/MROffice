/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { JsonValue } from "@mrnote/core/web/lib/collab/json-crdt.ts";
import type { PdfNativeListTextSelectionEvidence, PdfNativeTextPropertyMutationEvidence } from "./engine/native/list-text-fidelity";

/**
 * The journal record of the edit that last turned a run (changed the
 * linear part of its text matrix), when the run stands where that edit left
 * it: what a turn back needs to take back the shift that kept the turned
 * run on the page. Null when the run was moved, rewritten or never turned.
 */
export function previousNativeTextTurn(
  collection: JsonValue | undefined,
  selection: PdfNativeListTextSelectionEvidence,
): Readonly<{ mutation: PdfNativeTextPropertyMutationEvidence; replacementAdvance: number }> | null {
  if (!Array.isArray(collection)) return null;
  const item = selection.item;
  for (let index = collection.length - 1; index >= 0; index -= 1) {
    const record = collection[index];
    if (!record || typeof record !== "object" || Array.isArray(record)) continue;
    const fidelity = (record as { textFidelity?: unknown }).textFidelity;
    if (!fidelity || typeof fidelity !== "object" || Array.isArray(fidelity)) continue;
    const mutation = (fidelity as { propertyMutation?: unknown }).propertyMutation as
      | PdfNativeTextPropertyMutationEvidence
      | undefined;
    const replacementAdvance = (fidelity as { replacementAdvance?: unknown }).replacementAdvance;
    if (!mutation || typeof replacementAdvance !== "number") continue;
    const written = mutation.replacementSelection?.item;
    if (!written || written.str !== item.str) continue;
    // The run as this edit left it is the run opened now: same origin.
    if (
      Math.abs(Number(written.transform[4]) - Number(item.transform[4])) > 0.01 ||
      Math.abs(Number(written.transform[5]) - Number(item.transform[5])) > 0.01
    ) {
      continue;
    }
    const before = mutation.original.textMatrixLinear;
    const after = mutation.replacement.textMatrixLinear;
    if (before.some((value, position) => value !== after[position])) {
      return Object.freeze({ mutation, replacementAdvance });
    }
    // The same run, edited since without a turn: look further back.
  }
  return null;
}
