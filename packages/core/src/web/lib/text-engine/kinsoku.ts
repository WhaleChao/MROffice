/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Office's East Asian line breaking lists ("kinsoku"): the characters that
 * cannot start a line and the ones that cannot end one, per language.
 *
 * Source: ISO/IEC 29500-1 §17.3.1.16 (w:kinsoku), the defaults Word and the
 * other Office applications use when a document sets no
 * noLineBreaksBefore/noLineBreaksAfter of its own; the lists apply to
 * Simplified Chinese, Traditional Chinese and Japanese text (Korean has its
 * own pair, used by Word's Korean proofing). Shared by every editor: Docs'
 * Word-specific deltas (documentWordDelimiterBoundaries.ts) are measured
 * against browsers on top of these.
 */
export type KinsokuList = "zh-CN" | "zh-TW" | "ja" | "ko";

export interface KinsokuRules {
  readonly noStart: ReadonlySet<string>;
  readonly noEnd: ReadonlySet<string>;
}

const set = (chars: string): ReadonlySet<string> => new Set([...chars]);

export const OFFICE_KINSOKU: Readonly<Record<KinsokuList, KinsokuRules>> = {
  "zh-CN": {
    noStart: set("!%),.:;>?]}¢¨°·ˇˉ―‖’”…‰′″›℃∶、。〃〉》」』】〕〗〞︶︺︾﹀﹄﹚﹜﹞！＂％＇），．：；？］｀｜｝～￠"),
    noEnd: set("$([{£¥·‘“〈《「『【〔〖〝﹙﹛﹝＄（．［｛￡￥"),
  },
  "zh-TW": {
    noStart: set("!),.:;?]}¢·–—’”•‥…‧′╴、。〉》」』】〕〞︰︱︳︴︶︸︺︼︾﹀﹂﹄﹏﹐﹑﹒﹔﹕﹖﹗﹚﹜﹞！），．：；？］｜｝､"),
    noEnd: set("([{£¥‘“‵〈《「『【〔〝︵︷︹︻︽︿﹁﹃﹙﹛﹝（｛"),
  },
  ja: {
    noStart: set("!%),.:;?]}¢°’”‰′″℃、。々〉》」』】〕゛゜ゝゞ・ヽヾ！％），．：；？］｝｡｣､･ﾞﾟ￠"),
    noEnd: set("$([\\{£¥‘“〈《「『【〔＄（［｛｢￡￥"),
  },
  ko: {
    noStart: set("!%),.:;?]}¢°’”′″℃〉》」』】〕！％），．：；？］｝￠"),
    noEnd: set("$([\\{£¥‘“〈《「『【〔＄（［｛￡￥￦"),
  },
};

/** East Asian characters: the context in which the lists apply. */
export const KINSOKU_EAST_ASIAN = /[ᄀ-ᇿ⺀-鿿ꥠ-\ua97f가-\ud7af豈-\ufaff︰-﹏\uff00-\uffef]|[\u{20000}-\u{3134f}]/u;
