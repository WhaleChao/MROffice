/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Reading a document aloud.
 *
 * The browser's SpeechSynthesis does the speaking; what it does not do is
 * cope with real documents. Feed it a whole page and some engines go silent
 * partway through (the notorious long-utterance bug), garbage-collect the
 * utterance mid-sentence, or block the queue for minutes with no way to show
 * progress. So the text is cut into sentence-sized utterances first, spoken
 * one at a time, and the cutter is a pure function with its own tests —
 * because "where does a sentence end" is exactly the kind of question that
 * has a different answer in Chinese.
 *
 * The synthesis engine is injected. The controller never touches
 * `window.speechSynthesis` itself, which is what makes its sequencing —
 * advance, pause, resume, stop, skip — testable without a browser voice.
 */

/** Longest single utterance handed to the engine, in characters. */
export const SPEECH_MAX_CHUNK = 220;

/** Shorter than this, a chunk is merged forward rather than spoken alone. */
export const SPEECH_MIN_CHUNK = 12;

/** Sentence enders, Western and CJK. Kept explicit — \p{P} would split on
 *  commas and read a list like a countdown. */
const SENTENCE_END = /[.!?。！？…]+["'”’』」)】]?/y;

/** Fallback cut points inside an endless sentence, best first. */
const SOFT_BREAKS = ["；", ";", "，", ",", "、", " ", "\n"];
















export function chunkTextForSpeech(
  text: string,
  maxChunk = SPEECH_MAX_CHUNK,
): string[] {
  const normalized = text.replace(/\s+/g, " ").trim();
  if (!normalized) return [];
  const cap = Math.max(SPEECH_MIN_CHUNK, maxChunk);

  type Span = { start: number; end: number };
  const trimmed = (span: Span): Span => {
    let { start, end } = span;
    while (start < end && normalized[start] === " ") start += 1;
    while (end > start && normalized[end - 1] === " ") end -= 1;
    return { start, end };
  };

  // First pass: split at sentence enders.
  const sentences: Span[] = [];
  let start = 0;
  for (let i = 0; i < normalized.length; i += 1) {
    SENTENCE_END.lastIndex = i;
    const match = SENTENCE_END.exec(normalized);
    if (match && match.index === i) {
      const end = i + match[0].length;
      const next = normalized[end];
      if (/^\.+$/.test(match[0]) && next !== undefined && next !== " ") {
        i = end - 1;
        continue;
      }
      sentences.push(trimmed({ start, end }));
      start = end;
      i = end - 1;
    }
  }
  if (start < normalized.length) {
    sentences.push(trimmed({ start, end: normalized.length }));
  }

  // Second pass: pack sentences into chunks under the cap, splitting any
  // single sentence that is itself too long.
  const chunks: Span[] = [];
  let pending: Span | null = null;
  const flush = () => {
    if (pending && pending.end > pending.start) chunks.push(pending);
    pending = null;
  };
  for (const sentence of sentences) {
    if (sentence.end <= sentence.start) continue;
    if (pending && sentence.end - (pending as Span).start > cap) flush();
    if (sentence.end - sentence.start <= cap) {
      pending = pending
        ? { start: (pending as Span).start, end: sentence.end }
        : sentence;
      continue;
    }
    flush();
    // An endless sentence: cut at the softest break inside the cap.
    let rest = sentence.start;
    while (sentence.end - rest > cap) {
      const window = normalized.slice(rest, rest + cap);
      let cut = -1;
      for (const soft of SOFT_BREAKS) {
        const at = window.lastIndexOf(soft);
        if (at > cut) cut = at;
      }
      if (cut < SPEECH_MIN_CHUNK) cut = cap - 1;
      chunks.push(trimmed({ start: rest, end: rest + cut + 1 }));
      rest = trimmed({ start: rest + cut + 1, end: sentence.end }).start;
    }
    if (rest < sentence.end) pending = { start: rest, end: sentence.end };
  }
  flush();

  // Merge a trailing fragment into its neighbour so the reading does not end
  // on a two-word gasp.
  if (chunks.length >= 2) {
    const last = chunks[chunks.length - 1];
    const before = chunks[chunks.length - 2];
    if (last.end - last.start < SPEECH_MIN_CHUNK && last.end - before.start <= cap) {
      chunks.splice(chunks.length - 2, 2, { start: before.start, end: last.end });
    }
  }
  return chunks
    .map((span) => normalized.slice(span.start, span.end))
    .filter(Boolean);
}

/** The slice of SpeechSynthesis the controller uses, injectable for tests. */
export interface SpeechEngineLike {
  speak: (utterance: SpeechUtteranceLike) => void;
  cancel: () => void;
  pause: () => void;
  resume: () => void;
}

export interface SpeechUtteranceLike {
  text: string;
  lang?: string;
  rate?: number;
  voice?: unknown;
  onboundary?: ((event: SpeechBoundaryEventLike) => void) | null;
  onend: (() => void) | null;
  onerror: ((event: { error?: string }) => void) | null;
}

export interface SpeechBoundaryEventLike {
  readonly charIndex: number;
  readonly charLength?: number;
  readonly name?: string;
}

export interface SpeechWordBoundary {
  readonly charIndex: number;
  readonly charLength: number;
}

/**
 * Accept only a browser-provided, in-range word boundary. Engines that omit a
 * reliable length retain the existing chunk highlight instead of making the
 * viewer guess a word boundary from document text.
 */
export function readSpeechWordBoundary(
  text: string,
  event: SpeechBoundaryEventLike,
): SpeechWordBoundary | null {
  if (event.name && event.name !== "word") return null;
  if (
    !Number.isInteger(event.charIndex) ||
    !Number.isInteger(event.charLength) ||
    event.charIndex < 0 ||
    (event.charLength ?? 0) <= 0
  ) {
    return null;
  }
  const charLength = event.charLength as number;
  if (event.charIndex + charLength > text.length) return null;
  if (!text.slice(event.charIndex, event.charIndex + charLength).trim()) {
    return null;
  }
  return Object.freeze({ charIndex: event.charIndex, charLength });
}

export type SpeechStatus = "idle" | "speaking" | "paused" | "done" | "error";

export interface SpeechControllerOptions {
  engine: SpeechEngineLike;
  makeUtterance: (text: string) => SpeechUtteranceLike;
  lang?: string;
  rate?: number;
  voice?: unknown;
  /** Fired as each chunk starts, with its index and the total. */
  onProgress?: (index: number, total: number, text: string) => void;
  /** Fired only for a browser-proven word boundary inside the current chunk. */
  onWordProgress?: (
    index: number,
    total: number,
    text: string,
    boundary: SpeechWordBoundary,
  ) => void;
  onStatusChange?: (status: SpeechStatus) => void;
}

/**
 * Speaks chunks strictly one at a time.
 *
 * Each utterance is held on the instance until it ends: several engines
 * garbage-collect an unreferenced utterance and its callbacks with it, which
 * ends the reading silently after one chunk — the bug every web TTS thread
 * eventually finds.
 */
export class PdfSpeechController {
  private readonly options: SpeechControllerOptions;
  private chunks: string[] = [];
  private index = 0;
  private status: SpeechStatus = "idle";
  /** Kept referenced on purpose; see the class comment. */
  private current: SpeechUtteranceLike | null = null;
  /** Distinguishes our own cancel from an engine failure. */
  private stopping = false;
  /** Invalidates callbacks from utterances cancelled by navigation/restart. */
  private generation = 0;

  constructor(options: SpeechControllerOptions) {
    this.options = options;
  }

  getStatus(): SpeechStatus {
    return this.status;
  }

  /** 0-based chunk being spoken, and how many there are. */
  getPosition(): { index: number; total: number } {
    return { index: this.index, total: this.chunks.length };
  }

  private setStatus(status: SpeechStatus): void {
    if (this.status === status) return;
    this.status = status;
    this.options.onStatusChange?.(status);
  }

  /** Start reading `text` from the beginning, replacing any prior reading. */
  play(text: string): void {
    this.stop();
    this.chunks = chunkTextForSpeech(text);
    this.index = 0;
    if (this.chunks.length === 0) {
      this.setStatus("done");
      return;
    }
    this.setStatus("speaking");
    this.speakCurrent();
  }

  private speakCurrent(): void {
    if (this.index >= this.chunks.length) {
      this.current = null;
      this.setStatus("done");
      return;
    }
    const utterance = this.options.makeUtterance(this.chunks[this.index]);
    const generation = this.generation;
    if (this.options.lang) utterance.lang = this.options.lang;
    if (this.options.rate) utterance.rate = this.options.rate;
    if (this.options.voice) utterance.voice = this.options.voice;
    utterance.onboundary = (event) => {
      if (this.stopping || generation !== this.generation) return;
      const boundary = readSpeechWordBoundary(utterance.text, event);
      if (!boundary) return;
      this.options.onWordProgress?.(
        this.index,
        this.chunks.length,
        utterance.text,
        boundary,
      );
    };
    utterance.onend = () => {
      if (this.stopping || generation !== this.generation) return;
      this.index += 1;
      this.speakCurrent();
    };
    utterance.onerror = (event) => {
      if (this.stopping || generation !== this.generation) return;
      // "interrupted"/"canceled" arrive when the ENGINE was reset around us
      // (another tab, a route change). Anything else is a real failure.
      if (event.error === "interrupted" || event.error === "canceled") {
        this.setStatus("idle");
        return;
      }
      this.setStatus("error");
    };
    this.current = utterance;
    this.options.onProgress?.(
      this.index,
      this.chunks.length,
      this.chunks[this.index],
    );
    this.options.engine.speak(utterance);
  }

  pause(): void {
    if (this.status !== "speaking") return;
    this.options.engine.pause();
    this.setStatus("paused");
  }

  resume(): void {
    if (this.status !== "paused") return;
    this.options.engine.resume();
    this.setStatus("speaking");
  }

  /** Jump to the next chunk — "skip this paragraph". */
  skip(): void {
    if (this.status !== "speaking" && this.status !== "paused") return;
    // Cancelling fires the old utterance's events; `stopping` keeps them from
    // double-advancing, and the next speak re-arms cleanly.
    this.stopping = true;
    this.generation += 1;
    this.options.engine.cancel();
    this.stopping = false;
    this.index += 1;
    if (this.status === "paused") this.setStatus("speaking");
    this.speakCurrent();
  }

  /** Restart the previous chunk without reviving the cancelled utterance. */
  previous(): void {
    if (this.status !== "speaking" && this.status !== "paused") return;
    this.stopping = true;
    this.generation += 1;
    this.options.engine.cancel();
    this.stopping = false;
    this.index = Math.max(0, this.index - 1);
    if (this.status === "paused") this.setStatus("speaking");
    this.speakCurrent();
  }

  stop(): void {
    if (this.status === "idle") return;
    this.stopping = true;
    this.generation += 1;
    this.options.engine.cancel();
    this.stopping = false;
    this.current = null;
    this.chunks = [];
    this.index = 0;
    this.setStatus("idle");
  }
}
