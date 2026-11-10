"use client";


















/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { Node as ProseMirrorNode } from "@tiptap/pm/model";

import {
  documentJsonSliceForAiContext,
  documentTextForAiContext,
} from "../../__m__/mc38c6fb1a2";

export function selectionText(
  doc: ProseMirrorNode,
  from: number,
  to: number,
  blockSeparator = "\n",
): string {
  return documentTextForAiContext({
    json: documentJsonSliceForAiContext(doc, from, to),
    // Plain prose has no semantic nodes and comes straight through here,
    // byte-for-byte as before.
    fallbackText: doc.textBetween(from, to, blockSeparator),
  });
}
