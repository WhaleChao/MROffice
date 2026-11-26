/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Second-screen Presenter View — pure derivation + cross-window sync helpers.
 *
 * PowerPoint's "Presenter View" puts a speaker console on one display (current
 * slide large, the next slide as a thumbnail, the slide's speaker notes, an
 * elapsed timer and prev/next controls) while the audience sees only the clean
 * slide on the other display. This module is the side-effect-free, DOM-free
 * core of that feature: it
 *
 *   1. projects a deck slide into a serializable summary (title + text lines)
 *      so a SEPARATE browser window can render a faithful console without
 *      reaching back into the editor's React tree or its CSS;
 *   2. derives the full console view-model — `{ currentIndex, nextIndex, notes,
 *      elapsedMs, current, next, total }` — from `(deck, index, startTime,
 *      now)`, the single source of truth shared across windows;
 *   3. defines the tiny sync-message schema (live slide index + timer origin +
 *      running flag + deck snapshot) and its serialize/parse round-trip; and
 *   4. wraps a `BroadcastChannel` (with a `storage`-event fallback) behind a
 *      `post` / `subscribe` / `close` interface so the window-open shell in the
 *      component stays thin and untested logic stays out of the UI.
 *
 * Keeping all of this here means the elapsed math, the next-slide / end-of-deck
 * clamping, the message contract and the transport-selection fallback are all
 * exercised in isolation, and the presenter popup is a thin render of values
 * this module computed.
 *
 * The elapsed-timer formatting, next-slide index and slide-counter copy are
 * shared with the in-overlay presenter console via {@link
 * ./presentation-presenter}; this module re-exports / composes them so both the
 * single-window console and the second-window console report identical numbers.
 *
 * DEFERRED (intentionally out of scope for this presenter-view increment):
 *   - laser pointer (a movable pointer dot broadcast onto the audience slide);
 *   - ink / drawing annotations over the live slide;
 *   - rehearsal recording / timing capture (record per-slide dwell times).
 * Each would extend the sync message contract here (a new `dir`) plus an
 * audience-side render layer; none are implemented.
 */

import {
  formatElapsed,
  nextSlideIndex,
  presenterCounterLabel,
} from "./presentation-presenter";

export { formatElapsed, nextSlideIndex, presenterCounterLabel };

/**
 * The minimal shape this module needs from a deck slide. It is intentionally a
 * STRUCTURAL subset of the editor's richer `SlideModel` (so a `SlideModel[]`
 * passes without conversion) and dependency-free, so the projection can run in
 * the presenter popup without importing the multi-thousand-line editor module.
 */
export interface PresenterSlideInput {
  id: string;
  title?: string;
  notes?: string;
  blocks?: ReadonlyArray<PresenterBlockInput>;
}

/** The subset of a block this module reads to build the text projection. */
export interface PresenterBlockInput {
  type?: string;
  /** Flat text (text / text-box blocks). */
  text?: string;
  /** Image alt / caption, embed/chart title, etc. — any human-readable label. */
  alt?: string;
  caption?: string;
  title?: string;
  /** Table rows of cells (each cell carries `text`). */
  rows?: ReadonlyArray<ReadonlyArray<{ text?: string }>>;
}

/**
 * A fully serializable projection of one slide: its title plus the ordered
 * human-readable text lines of its blocks. This is what crosses the window
 * boundary (it must survive `structuredClone` / `JSON`), so it holds only
 * strings — no geometry, no React, no functions.
 */
export interface PresenterSlideSummary {
  id: string;
  title: string;
  /** Ordered, non-empty text lines projected from the slide's blocks. */
  lines: string[];
}

/** Upper bound on projected lines per slide (a console preview, not the deck). */
const MAX_SUMMARY_LINES = 40;
/** Upper bound on one projected line's length (defends the cross-window payload). */
const MAX_SUMMARY_LINE_LEN = 2000;

function pushLines(target: string[], raw: string | undefined): void {
  if (!raw) return;
  for (const piece of raw.split("\n")) {
    if (target.length >= MAX_SUMMARY_LINES) return;
    const line = piece.trim();
    if (line) target.push(line.slice(0, MAX_SUMMARY_LINE_LEN));
  }
}

/**
 * Project one block into zero or more human-readable text lines. Mirrors (in a
 * pared-down, dependency-free way) how the editor flattens a block to plain
 * text: flat `text` for text blocks, alt/caption/title labels for media, and
 * tab-joined cells for table rows. Shapes (purely visual) and unknown block
 * types contribute nothing. Pure.
 */
function projectBlockLines(block: PresenterBlockInput): string[] {
  const out: string[] = [];
  switch (block.type) {
    case "image":
      pushLines(out, block.alt);
      pushLines(out, block.caption);
      break;
    case "table":
      pushLines(out, block.caption);
      for (const row of block.rows ?? []) {
        const cells = row.map((cell) => (cell?.text ?? "").trim()).filter(Boolean);
        if (cells.length) pushLines(out, cells.join("\t"));
      }
      break;
    case "chart":
    case "embed":
      pushLines(out, block.title);
      pushLines(out, block.caption);
      break;
    case "shape":
      // Purely visual — no text content to surface in the console preview.
      break;
    default:
      // text / text-box / anything carrying flat `text`.
      pushLines(out, block.text);
      break;
  }
  return out;
}

/**
 * Project a slide into its serializable {@link PresenterSlideSummary}. The title
 * falls back to an empty string (never `undefined`) and the body lines are the
 * concatenated, trimmed, capped projections of each block. Pure + total: a
 * blocks-less slide yields `{ id, title, lines: [] }`.
 */
export function projectSlideSummary(
  slide: PresenterSlideInput,
): PresenterSlideSummary {
  const lines: string[] = [];
  for (const block of slide.blocks ?? []) {
    if (lines.length >= MAX_SUMMARY_LINES) break;
    for (const line of projectBlockLines(block)) {
      if (lines.length >= MAX_SUMMARY_LINES) break;
      lines.push(line);
    }
  }
  return { id: slide.id, title: (slide.title ?? "").trim(), lines };
}

/** Project an entire deck into the serializable snapshot the popup renders. */
export function projectDeck(
  slides: ReadonlyArray<PresenterSlideInput>,
): PresenterSlideSummary[] {
  return slides.map(projectSlideSummary);
}

/**
 * The fully-derived presenter console view-model for a given live index and
 * timer origin. Everything the popup (or the in-window console) needs to render
 * one frame, computed purely from the deck snapshot + index + clock.
 */
export interface PresenterViewModel {
  /** 0-based clamped index of the slide currently being presented. */
  currentIndex: number;
  /** 0-based index of the next slide, or `null` on the last slide. */
  nextIndex: number | null;
  /** The current slide's summary (or `null` for an empty deck). */
  current: PresenterSlideSummary | null;
  /** The next slide's summary (or `null` at end-of-deck / empty deck). */
  next: PresenterSlideSummary | null;
  /** The current slide's raw speaker notes (trimmed; empty string when none). */
  notes: string;
  /** Elapsed presentation time in milliseconds (≥ 0). */
  elapsedMs: number;
  /** Total slide count (≥ 0). */
  total: number;
  /** 1-based `{ current, total }` counter pair (mirrors the audience counter). */
  counter: { current: number; total: number };
}

/**
 * Compute the elapsed milliseconds from a timer origin. When `startTime` is
 * `null` the timer is paused and `elapsedMs` is whatever was banked (passed via
 * `accumulatedMs`); otherwise it is the banked time plus `(now - startTime)`.
 * Clamps to ≥ 0 so a clock that ran backwards never shows a negative timer.
 */
export function computeElapsedMs(
  startTime: number | null,
  now: number,
  accumulatedMs = 0,
): number {
  const banked = Number.isFinite(accumulatedMs) ? Math.max(0, accumulatedMs) : 0;
  if (startTime === null || !Number.isFinite(startTime)) return banked;
  if (!Number.isFinite(now)) return banked;
  return banked + Math.max(0, now - startTime);
}

/**
 * Derive the full presenter console view-model from a deck snapshot, the live
 * slide index, the running timer origin and the current wall-clock. This is the
 * heart of the feature: BOTH windows derive from the same inputs, so the
 * audience slide and the presenter console can never disagree about which slide
 * is live or how much time has elapsed.
 *
 * - `index` is clamped into `[0, total)`; an out-of-range or non-finite index
 *   resolves to the nearest valid slide (or `null` current for an empty deck).
 * - `nextIndex` is `null` on the last slide so the console can show an
 *   "end of slideshow" placeholder rather than a bogus preview.
 * - `notes` is the current slide's trimmed notes ("" when absent).
 * - `elapsedMs` is `computeElapsedMs(startTime, now, accumulatedMs)`.
 *
 * Pure + total — no DOM, no side effects.
 */
export function derivePresenterView(
  slides: ReadonlyArray<PresenterSlideSummary>,
  index: number,
  startTime: number | null,
  now: number,
  accumulatedMs = 0,
): PresenterViewModel {
  const total = slides.length;
  const counter = presenterCounterLabel(index, total);
  const elapsedMs = computeElapsedMs(startTime, now, accumulatedMs);
  if (total === 0) {
    return {
      currentIndex: 0,
      nextIndex: null,
      current: null,
      next: null,
      notes: "",
      elapsedMs,
      total: 0,
      counter,
    };
  }
  const currentIndex = Math.max(
    0,
    Math.min(Number.isFinite(index) ? Math.floor(index) : 0, total - 1),
  );
  const nextIdx = nextSlideIndex(currentIndex, total);
  const current = slides[currentIndex] ?? null;
  const next = nextIdx === null ? null : slides[nextIdx] ?? null;
  // The current slide's notes ride on the original deck, not the summary; but
  // the summary is all the popup has, so notes are carried on the snapshot (see
  // `buildSyncSnapshot`). When deriving from summaries alone there are no notes
  // to attach, so callers that need notes derive from the snapshot below.
  return {
    currentIndex,
    nextIndex: nextIdx,
    current,
    next,
    notes: "",
    elapsedMs,
    total,
    counter,
  };
}

// ── Cross-window sync message contract ──────────────────────────────────────

/** Discriminant tag identifying a presenter-view sync message. */
export const PRESENTER_MESSAGE_KIND = "mrnote.presenter.sync.v1" as const;

/**
 * The serializable snapshot the PRESENT window broadcasts to the presenter
 * popup: the deck projection (incl. per-slide notes), the live index, the timer
 * origin + banked time + running flag, and an `epoch` so the popup can ignore a
 * stale broadcast from a previous present session. `startTime` is `null` while
 * the timer is paused. This is the single payload that crosses the window
 * boundary; it must remain JSON-serializable.
 */
export interface PresenterSyncSnapshot {
  /** Projected slide summaries (title + body lines), parallel to `notes`. */
  deck: PresenterSlideSummary[];
  /** Per-slide trimmed speaker notes, index-aligned with `deck`. */
  notes: string[];
  /** 0-based live slide index. */
  index: number;
  /** Wall-clock origin of the running timer segment, or `null` when paused. */
  startTime: number | null;
  /** Milliseconds banked across prior pause/resume cycles (≥ 0). */
  accumulatedMs: number;
  /** Whether the elapsed timer is currently running. */
  running: boolean;
  /** Monotonic id of the present session (a new present run bumps it). */
  epoch: number;
  /**
   * The show's Random Effects seed (row 1041), fixed when the show starts:
   * every window of the show picks the same effect for each step from it.
   */
  randomEffectSeed?: number;
}

/**
 * A command the presenter POPUP sends back to the present window (the speaker
 * clicked prev/next or a timer control there). The present window applies it to
 * its authoritative state and re-broadcasts a fresh snapshot, so both windows
 * converge. `index` carries an explicit target for `goto`.
 */
export interface PresenterCommand {
  action:
    | "next"
    | "prev"
    | "first"
    | "last"
    | "goto"
    | "toggleTimer"
    | "resetTimer"
    // The magnifying glass (GAP 7): the console's Zoom into slide (armed / off), and its click on
    // the current slide at (fx, fy), fractions of the slide.
    | "toggleMagnify"
    | "magnifyAt";
  /** Target slide index for `goto` (ignored otherwise). */
  index?: number;
  /** The slide point of `magnifyAt` (fractions of the slide; ignored otherwise). */
  fx?: number;
  fy?: number;
}

/** The union of messages that flow over the presenter channel, both directions. */
export type PresenterMessage =
  | { kind: typeof PRESENTER_MESSAGE_KIND; dir: "snapshot"; snapshot: PresenterSyncSnapshot }
  | { kind: typeof PRESENTER_MESSAGE_KIND; dir: "command"; command: PresenterCommand }
  // A popup-ready handshake: the popup announces it is alive so the present
  // window immediately pushes the current snapshot (rather than waiting for the
  // next index change). Carries no payload.
  | { kind: typeof PRESENTER_MESSAGE_KIND; dir: "hello" }
  // The present window left present mode; the popup should close itself.
  | { kind: typeof PRESENTER_MESSAGE_KIND; dir: "bye" };

/** Build the broadcast snapshot from the live deck + present-mode state. Pure. */
export function buildSyncSnapshot(
  slides: ReadonlyArray<PresenterSlideInput>,
  index: number,
  timer: {
    startTime: number | null;
    accumulatedMs: number;
    running: boolean;
  },
  epoch: number,
  randomEffectSeed?: number,
): PresenterSyncSnapshot {
  return {
    ...(randomEffectSeed !== undefined && Number.isInteger(randomEffectSeed) ? { randomEffectSeed } : {}),
    deck: projectDeck(slides),
    notes: slides.map((s) => (s.notes ?? "").trim()),
    index: Number.isFinite(index) ? Math.max(0, Math.floor(index)) : 0,
    startTime:
      timer.startTime === null || !Number.isFinite(timer.startTime)
        ? null
        : timer.startTime,
    accumulatedMs: Number.isFinite(timer.accumulatedMs)
      ? Math.max(0, timer.accumulatedMs)
      : 0,
    running: !!timer.running,
    epoch: Number.isFinite(epoch) ? epoch : 0,
  };
}

/**
 * Derive the popup's full view-model from a received snapshot — the analogue of
 * {@link derivePresenterView} that also attaches the current slide's notes (the
 * snapshot carries them index-aligned). This is what the popup renders each
 * frame. Pure + total.
 */
export function derivePresenterViewFromSnapshot(
  snapshot: PresenterSyncSnapshot,
  now: number,
): PresenterViewModel {
  const vm = derivePresenterView(
    snapshot.deck,
    snapshot.index,
    snapshot.startTime,
    now,
    snapshot.accumulatedMs,
  );
  const notes =
    snapshot.deck.length === 0
      ? ""
      : (snapshot.notes[vm.currentIndex] ?? "").trim();
  return { ...vm, notes };
}

/** Serialize any presenter message to a JSON string for transport. */
export function serializePresenterMessage(message: PresenterMessage): string {
  return JSON.stringify(message);
}

/**
 * Parse a transport string/value back into a {@link PresenterMessage}, or
 * `null` when it is not a well-formed presenter message (wrong kind, malformed
 * JSON, unknown direction). Accepts either a raw JSON string (storage-event
 * path) or an already-parsed object (BroadcastChannel path).
 */
export function parsePresenterMessage(raw: unknown): PresenterMessage | null {
  let value: unknown = raw;
  if (typeof raw === "string") {
    try {
      value = JSON.parse(raw);
    } catch {
      return null;
    }
  }
  if (!value || typeof value !== "object") return null;
  const obj = value as Record<string, unknown>;
  if (obj.kind !== PRESENTER_MESSAGE_KIND) return null;
  switch (obj.dir) {
    case "snapshot":
      if (obj.snapshot && typeof obj.snapshot === "object") {
        return {
          kind: PRESENTER_MESSAGE_KIND,
          dir: "snapshot",
          snapshot: obj.snapshot as PresenterSyncSnapshot,
        };
      }
      return null;
    case "command":
      if (obj.command && typeof obj.command === "object") {
        return {
          kind: PRESENTER_MESSAGE_KIND,
          dir: "command",
          command: obj.command as PresenterCommand,
        };
      }
      return null;
    case "hello":
      return { kind: PRESENTER_MESSAGE_KIND, dir: "hello" };
    case "bye":
      return { kind: PRESENTER_MESSAGE_KIND, dir: "bye" };
    default:
      return null;
  }
}

// ── Channel transport (BroadcastChannel + storage fallback) ─────────────────

/** A minimal duplex channel: post messages, subscribe to inbound, and close. */
export interface PresenterChannel {
  /** Which transport this channel selected ("broadcast" | "storage" | "none"). */
  readonly transport: "broadcast" | "storage" | "none";
  /** Broadcast a message to the other window(s). */
  post(message: PresenterMessage): void;
  /** Subscribe to inbound messages; returns an unsubscribe function. */
  subscribe(handler: (message: PresenterMessage) => void): () => void;
  /** Tear down the channel (close the BroadcastChannel / remove listeners). */
  close(): void;
}

/**
 * The slice of the host environment the channel needs. Injected so the
 * transport selection + fallback logic can be unit-tested by handing in a
 * `BroadcastChannel`-less environment (forcing the storage path) or a fully
 * inert one (forcing "none"). Defaults to the real `globalThis` in the browser.
 */
export interface PresenterChannelEnv {
  BroadcastChannel?: typeof BroadcastChannel;
  localStorage?: Pick<Storage, "setItem" | "removeItem">;
  addEventListener?: (type: "storage", listener: (event: StorageEvent) => void) => void;
  removeEventListener?: (
    type: "storage",
    listener: (event: StorageEvent) => void,
  ) => void;
}

function resolveEnv(env?: PresenterChannelEnv): PresenterChannelEnv {
  if (env) return env;
  if (typeof globalThis === "undefined") return {};
  const g = globalThis as unknown as {
    BroadcastChannel?: typeof BroadcastChannel;
    localStorage?: Storage;
    addEventListener?: PresenterChannelEnv["addEventListener"];
    removeEventListener?: PresenterChannelEnv["removeEventListener"];
  };
  return {
    BroadcastChannel: g.BroadcastChannel,
    localStorage: g.localStorage,
    addEventListener: g.addEventListener
      ? (type, listener) =>
          (g.addEventListener as (t: string, l: unknown) => void)(type, listener)
      : undefined,
    removeEventListener: g.removeEventListener
      ? (type, listener) =>
          (g.removeEventListener as (t: string, l: unknown) => void)(type, listener)
      : undefined,
  };
}

/**
 * Open a presenter sync channel named `name`. Prefers `BroadcastChannel`; when
 * that is unavailable it falls back to a `storage`-event bus (write a key, let
 * the OTHER window's `storage` listener fire, then clear the key so it never
 * persists). When neither transport exists it returns an inert "none" channel
 * (post is a no-op, subscribe never fires) so callers need no extra guards.
 *
 * Same-window note: `BroadcastChannel` and `storage` events both fire only in
 * OTHER documents, never the poster's own — which is exactly the present↔popup
 * topology here, and lets a single-window unit test use two channel instances
 * to observe a round-trip.
 */
export function createPresenterChannel(
  name: string,
  env?: PresenterChannelEnv,
): PresenterChannel {
  const resolved = resolveEnv(env);

  // Preferred: BroadcastChannel.
  if (typeof resolved.BroadcastChannel === "function") {
    const bc = new resolved.BroadcastChannel(name);
    return {
      transport: "broadcast",
      post(message) {
        bc.postMessage(message);
      },
      subscribe(handler) {
        const listener = (event: MessageEvent) => {
          const parsed = parsePresenterMessage(event.data);
          if (parsed) handler(parsed);
        };
        bc.addEventListener("message", listener);
        return () => bc.removeEventListener("message", listener);
      },
      close() {
        bc.close();
      },
    };
  }

  // Fallback: storage events. The key is namespaced by channel name; each post
  // writes a fresh `{ seq, payload }` envelope (the `seq` guarantees the value
  // always changes, so even identical back-to-back messages fire `storage`),
  // then removes the key so nothing lingers in localStorage.
  const storage = resolved.localStorage;
  const add = resolved.addEventListener;
  const remove = resolved.removeEventListener;
  if (storage && add && remove) {
    const storageKey = `mrnote.presenter.channel.${name}`;
    let seq = 0;
    return {
      transport: "storage",
      post(message) {
        seq += 1;
        const envelope = JSON.stringify({
          seq,
          payload: serializePresenterMessage(message),
        });
        try {
          storage.setItem(storageKey, envelope);
          // Clear immediately; the value's presence already fired the event in
          // the other window. Leaving it would pollute storage + re-deliver on
          // a future read. (Removal itself fires a storage event with
          // newValue=null, which we ignore below.)
          storage.removeItem(storageKey);
        } catch {
          // Quota / disabled storage — degrade silently to no-op for this post.
        }
      },
      subscribe(handler) {
        const listener = (event: StorageEvent) => {
          if (event.key !== storageKey || event.newValue == null) return;
          let payload: unknown;
          try {
            payload = (JSON.parse(event.newValue) as { payload?: unknown }).payload;
          } catch {
            return;
          }
          const parsed = parsePresenterMessage(payload);
          if (parsed) handler(parsed);
        };
        add("storage", listener);
        return () => remove("storage", listener);
      },
      close() {
        // No persistent state to tear down beyond per-subscription listeners.
      },
    };
  }

  // No transport available (SSR / locked-down env): inert channel.
  return {
    transport: "none",
    post() {},
    subscribe() {
      return () => {};
    },
    close() {},
  };
}

/** Stable channel name for a given deck/page so both windows rendezvous. */
export function presenterChannelName(pageId: string): string {
  return `presenter:${pageId}`;
}

// ── Presenter popup document rendering (pure HTML generation) ────────────────

/**
 * Localized strings the popup console renders. Supplied by the opener (it reads
 * them from next-intl) so the second window honors the same i18n as the editor
 * without bundling a translator into the popup.
 */
export interface PresenterDocStrings {
  /** Window / document title. */
  title: string;
  /** "Current slide" section label. */
  current: string;
  /** "Next slide" section label. */
  next: string;
  /** "Notes" section label. */
  notes: string;
  /** Placeholder when the current slide has no notes. */
  noNotes: string;
  /** End-of-deck placeholder for the next-slide preview. */
  end: string;
  /** Previous-slide control label. */
  prev: string;
  /** Next-slide control label. */
  nextControl: string;
  /** Pause-timer control label. */
  pauseTimer: string;
  /** Reset-timer control label. */
  resetTimer: string;
  /** Elapsed-time region aria-label template ("Elapsed time {elapsed}"). */
  elapsedLabel: (elapsed: string) => string;
  /** Slide-counter aria-label template ("Slide {current} of {total}"). */
  counterLabel: (current: number, total: number) => string;
  /** The magnifying glass's label (PowerPoint's "Zoom into slide"); no glass without it. */
  magnify?: string;
}

/** The show's magnifying glass as the console draws it (GAP 7). */
export interface PresenterMagnifyState {
  /** The glass is chosen: a click on the current slide magnifies there. */
  readonly armed: boolean;
  /** The slide is magnified. */
  readonly on: boolean;
}

/** HTML-escape a string for safe interpolation into the popup document. */
export function escapeHtml(value: string): string {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

/** Render one slide summary's body as escaped `<p>` lines (or a title fallback). */
function renderSummaryBody(summary: PresenterSlideSummary | null): string {
  if (!summary) return "";
  const parts: string[] = [];
  if (summary.title) {
    parts.push(`<h2 class="ps-slide__title">${escapeHtml(summary.title)}</h2>`);
  }
  for (const line of summary.lines) {
    parts.push(`<p class="ps-slide__line">${escapeHtml(line)}</p>`);
  }
  return parts.join("");
}

/**
 * The slides' own pictures for the console (object URLs of the show's
 * drawing of each). A slide without one yet shows its words meanwhile.
 */
export interface PresenterSlidePictures {
  current?: string;
  next?: string;
}

/** The slide as a picture when there is one, its words otherwise. */
function renderSlideBody(
  summary: PresenterSlideSummary | null,
  picture: string | undefined,
): string {
  if (!summary) return "";
  if (!picture || !/^(?:blob:|data:image\/png;)/.test(picture)) return renderSummaryBody(summary);
  return `<img class="ps-picture" data-ps-picture="${escapeHtml(picture)}" src="${escapeHtml(
    picture,
  )}" alt="${escapeHtml(summary.title)}">`;
}

/**
 * The static `<style>` for the popup console. Self-contained (the popup is a
 * blank same-origin window with no app CSS), dark like PowerPoint's presenter
 * display, and responsive to the window size.
 */
export const PRESENTER_DOC_STYLE = `
:root{color-scheme:dark}
*{box-sizing:border-box}
html,body{margin:0;height:100%;background:#0b0d10;color:#e8eaed;
  font-family:system-ui,-apple-system,"Segoe UI",Roboto,Helvetica,Arial,sans-serif}
body{display:flex;flex-direction:column;padding:18px;gap:14px}
.ps-top{display:flex;flex:1 1 auto;gap:16px;min-height:0}
.ps-current{flex:2 1 0;display:flex;flex-direction:column;min-width:0;min-height:0}
.ps-side{flex:1 1 0;display:flex;flex-direction:column;gap:12px;min-width:0;min-height:0}
.ps-label{font-size:12px;letter-spacing:.04em;text-transform:uppercase;
  color:#9aa0a6;margin-bottom:6px}
.ps-stage{flex:1 1 auto;background:#16181c;border:1px solid #2a2d31;border-radius:10px;
  padding:20px;overflow:auto;min-height:0}
.ps-stage--next{flex:0 0 auto;max-height:34%}
.ps-stage[data-picture="true"]{padding:0;display:flex;align-items:center;justify-content:center;
  overflow:hidden;background:#000}
.ps-stage--next[data-picture="true"]{max-height:none}
.ps-picture{display:block;max-width:100%;max-height:100%;width:auto;height:auto;object-fit:contain}
.ps-stage[data-magnify="armed"] .ps-picture{cursor:zoom-in}
.ps-stage[data-magnify="on"] .ps-picture{cursor:grab}
.ps-btn[aria-pressed="true"]{background:#2b3a52;border-color:#5b9dff}
.ps-slide__title{margin:0 0 10px;font-size:22px;line-height:1.25}
.ps-stage--next .ps-slide__title{font-size:15px}
.ps-slide__line{margin:0 0 6px;font-size:16px;line-height:1.4;color:#cfd2d6}
.ps-stage--next .ps-slide__line{font-size:12px;color:#aeb2b7}
.ps-end{color:#7a7f85;font-style:italic;margin:0}
.ps-notes{flex:1 1 auto;background:#16181c;border:1px solid #2a2d31;border-radius:10px;
  padding:16px;overflow:auto;min-height:0}
.ps-notes__body{margin:0;white-space:pre-wrap;font-size:15px;line-height:1.5;color:#dfe2e6}
.ps-notes__body--empty{color:#7a7f85;font-style:italic}
.ps-bar{display:flex;align-items:center;gap:12px;flex:0 0 auto}
.ps-timer{font-variant-numeric:tabular-nums;font-size:28px;font-weight:600;
  min-width:96px}
.ps-counter{font-variant-numeric:tabular-nums;font-size:16px;color:#9aa0a6;
  margin-left:auto}
.ps-btn{appearance:none;border:1px solid #34373c;background:#1d2024;color:#e8eaed;
  border-radius:8px;padding:9px 14px;font-size:14px;cursor:pointer}
.ps-btn:hover{background:#262a2f}
.ps-btn:focus-visible{outline:2px solid #5b9dff;outline-offset:2px}
`.trim();

/**
 * Render the popup console's full inner-body HTML for a view-model + labels.
 * PURE: produces a string only (the imperative shell writes it into the popup
 * document and wires the `data-presenter-action` buttons). Every dynamic value
 * is HTML-escaped. The structure mirrors the in-window presenter console:
 * current slide (large), next slide (thumb) / end placeholder, speaker notes,
 * an elapsed timer with pause/reset, prev/next controls and a slide counter.
 */
export function renderPresenterConsoleHtml(
  vm: PresenterViewModel,
  strings: PresenterDocStrings,
  pictures: PresenterSlidePictures = {},
  magnify: PresenterMagnifyState = { armed: false, on: false },
): string {
  const glassPressed = magnify.armed || magnify.on;
  const glass = strings.magnify
    ? `
  <button class="ps-btn" type="button" data-presenter-action="toggleMagnify"
    aria-label="${escapeHtml(strings.magnify)}" title="${escapeHtml(strings.magnify)}"
    aria-pressed="${glassPressed ? "true" : "false"}" data-testid="ps-magnify">&#x1F50D;&#xFE0E;</button>`
    : "";
  const elapsed = formatElapsed(vm.elapsedMs);
  const nextInner =
    vm.next === null
      ? `<p class="ps-end" data-testid="ps-next-end">${escapeHtml(strings.end)}</p>`
      : renderSlideBody(vm.next, pictures.next);
  const notesText = vm.notes.trim();
  const notesBody = notesText
    ? `<p class="ps-notes__body" data-testid="ps-notes">${escapeHtml(notesText)}</p>`
    : `<p class="ps-notes__body ps-notes__body--empty" data-testid="ps-notes">${escapeHtml(
        strings.noNotes,
      )}</p>`;
  return `
<div class="ps-top">
  <section class="ps-current">
    <span class="ps-label">${escapeHtml(strings.current)}</span>
    <div class="ps-stage" data-testid="ps-current"${
      pictures.current ? ' data-picture="true"' : ""
    }${magnify.on ? ' data-magnify="on"' : magnify.armed ? ' data-magnify="armed"' : ""}>${renderSlideBody(vm.current, pictures.current)}</div>
  </section>
  <aside class="ps-side">
    <div>
      <span class="ps-label">${escapeHtml(strings.next)}</span>
      <div class="ps-stage ps-stage--next" data-testid="ps-next" data-empty="${
        vm.next === null ? "true" : "false"
      }"${pictures.next && vm.next !== null ? ' data-picture="true"' : ""}>${nextInner}</div>
    </div>
    <div style="display:flex;flex-direction:column;flex:1 1 auto;min-height:0">
      <span class="ps-label">${escapeHtml(strings.notes)}</span>
      <div class="ps-notes">${notesBody}</div>
    </div>
  </aside>
</div>
<div class="ps-bar">
  <button class="ps-btn" type="button" data-presenter-action="prev"
    aria-label="${escapeHtml(strings.prev)}" title="${escapeHtml(strings.prev)}">&larr;</button>
  <button class="ps-btn" type="button" data-presenter-action="next"
    aria-label="${escapeHtml(strings.nextControl)}" title="${escapeHtml(
      strings.nextControl,
    )}">&rarr;</button>
  <span class="ps-timer" data-testid="ps-timer"
    aria-label="${escapeHtml(strings.elapsedLabel(elapsed))}">${escapeHtml(elapsed)}</span>
  <button class="ps-btn" type="button" data-presenter-action="toggleTimer"
    aria-label="${escapeHtml(strings.pauseTimer)}" title="${escapeHtml(
      strings.pauseTimer,
    )}">⏸</button>
  <button class="ps-btn" type="button" data-presenter-action="resetTimer"
    aria-label="${escapeHtml(strings.resetTimer)}" title="${escapeHtml(
      strings.resetTimer,
    )}">↺</button>${glass}
  <span class="ps-counter" data-testid="ps-counter"
    data-current="${vm.counter.current}" data-total="${vm.counter.total}"
    aria-label="${escapeHtml(
      strings.counterLabel(vm.counter.current, vm.counter.total),
    )}">${vm.counter.current} / ${vm.counter.total}</span>
</div>`.trim();
}
