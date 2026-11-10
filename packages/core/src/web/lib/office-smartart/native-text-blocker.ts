/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Whether a preserved Office diagram's text can be edited, and the host's
 * sentence when it cannot. Host-neutral and browser-safe: Docs, Presentation
 * and Spreadsheet ask it per render, from the texts they already hold. The
 * data-model reader (`./native-outline`) runs on the server.
 */

/**
 * Why a preserved Office diagram's text cannot be edited here.
 *
 * - `placeholderTextOnly` — Office saved the diagram while every node still
 *   held its `[Text]` prompt. A prompt is not content: the package carries no
 *   `a:t` run anywhere in the diagram, so there is no text to change, and the
 *   write-back may not invent a run Office never wrote. 37 of the 39 diagrams
 *   in Office's own 39-layout deck are in exactly this state
 *   (`docs/office-graphics/SMARTART-STYLES.md` §7.5).
 * - `notAttributable` — the diagram does carry text, but this host could not
 *   bind it back to the data model, so no edit of it could be written either.
 */
export type NativeSmartArtTextBlocker =
  | "placeholderTextOnly"
  | "notAttributable";

/**
 * The blocker to state for one preserved diagram, or `null` when its text IS
 * editable. `texts` is every string the diagram carries — its outline items on
 * a slide, its member labels in a document or a sheet — so one host-neutral
 * answer serves all three.
 */
export function nativeSmartArtTextBlocker(input: {
  readonly textEditable: boolean;
  readonly texts: Iterable<unknown>;
}): NativeSmartArtTextBlocker | null {
  if (input.textEditable) return null;
  for (const text of input.texts) {
    if (typeof text === "string" && text.trim() !== "") return "notAttributable";
  }
  return "placeholderTextOnly";
}

/**
 * The host's own sentence for a blocker, or `null` when there is none. Each
 * host owns its i18n namespace and passes both strings; the mapping itself
 * stays here so Docs, Presentation and Spreadsheet state the same thing.
 */
export function nativeSmartArtTextBlockerLabel(
  blocker: NativeSmartArtTextBlocker | null,
  labels: {
    readonly placeholderTextOnly: string;
    readonly notAttributable: string;
  },
): string | null {
  return blocker === null ? null : labels[blocker];
}
