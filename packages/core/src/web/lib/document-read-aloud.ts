/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Pure helpers for the document "Read Aloud" feature (Word Review ▸ Read Aloud /
 * Edge Immersive Reader): split the text to speak into bounded, sentence-aligned
 * chunks. Browsers cap a single `SpeechSynthesisUtterance` (and some silently
 * stop on very long strings), so the player queues several short utterances;
 * chunking on sentence boundaries also keeps prosody natural and lets the user
 * stop promptly between chunks. PURE + dependency-free → unit-testable apart
 * from the Web Speech API.
 */

/** Default per-utterance character budget (kept well under browser limits). */
export const READ_ALOUD_MAX_CHUNK = 220;

/**
 * Split `text` into speakable chunks, each ≤ `maxLen` characters.
 * Splits on sentence punctuation (Latin and CJK: 。！？；) / hard line breaks,
 * packing consecutive short sentences together up to the budget; a single
 * sentence longer than the budget is wrapped at its clauses (，、;：), then on
 * whitespace (never mid-word), and a run with no whitespace at all (Chinese,
 * Japanese) is cut at the budget. Returns `[]` for blank input.
 */
export function splitIntoSpeechChunks(
  text: string,
  maxLen: number = READ_ALOUD_MAX_CHUNK,
): string[] {
  const normalized = text.replace(/\r\n/g, "\n").trim();
  if (!normalized) return [];
  const limit = Math.max(1, Math.floor(maxLen));

  const chunks: string[] = [];
  let buffer = "";
  const flush = () => {
    const trimmed = buffer.trim();
    if (trimmed) chunks.push(trimmed);
    buffer = "";
  };

  // Sentence-ish pieces: a run up to (and including) terminal punctuation and
  // the closing quote after it, OR a block of blank lines. The fallback keeps
  // the whole string as one piece.
  const pieces =
    normalized.match(/[^.!?。！？；\n]+[.!?。！？；]*[”’」』）)]*\n*|\n+/g) ?? [normalized];
  for (const piece of pieces) {
    if (piece.length > limit) {
      // An overlong sentence: flush what we have, then wrap it.
      flush();
      hardWrap(piece, limit, chunks);
      continue;
    }
    if (buffer.trim() && (buffer + piece).length > limit) flush();
    buffer += piece;
  }
  flush();
  return chunks;
}

/** An overlong sentence in chunks: at its clauses first, then on whitespace,
 *  and a run with no whitespace (a Chinese sentence without commas) at the
 *  budget itself, so no chunk is ever longer than `limit`. */
function hardWrap(piece: string, limit: number, out: string[]): void {
  let line = "";
  const push = (part: string) => {
    if (line.trim() && (line + part).length > limit) {
      out.push(line.trim());
      line = "";
    }
    line += part;
  };
  const clauses = piece.match(/[^，、,;：:]+[，、,;：:]*/g) ?? [piece];
  for (const clause of clauses) {
    if (clause.length <= limit) {
      push(clause);
      continue;
    }
    for (const token of clause.split(/(\s+)/)) {
      if (token.length <= limit) {
        push(token);
        continue;
      }
      const chars = Array.from(token);
      for (let i = 0; i < chars.length; i += limit) push(chars.slice(i, i + limit).join(""));
    }
  }
  if (line.trim()) out.push(line.trim());
}
