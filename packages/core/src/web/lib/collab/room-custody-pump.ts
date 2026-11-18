/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
export interface RoomCustodyEdits {
  /** What is sent; the same bytes again are not sent again. A part room's
   * custody (lib/collab/document-part-room) sends its updates by id and puts
   * their ids here. */
  update: Uint8Array;
  epoch: number;
  model: string;
}

export interface RoomCustodyState {
  /** The room's socket is admitted and confirms this window's updates. */
  live: boolean;
  /** Every update this window made is confirmed durable by the room, or
   * the room is live on a server that does not confirm (it relays and saves
   * as before: nothing for custody). */
  settled: boolean;
}

export interface RoomCustodyPumpOptions<E extends RoomCustodyEdits = RoomCustodyEdits> {
  /** What the document holds beyond its adopted checkpoint; null: nothing
   * (or the room confirmed everything). */
  read: () => E | null;
  state: () => RoomCustodyState;
  /** `signal` aborts a request a newer state supersedes (still waiting for a
   * connection, or for its answer). */
  send: (edits: E, signal: AbortSignal) => Promise<unknown>;
  /** How often a waiting change is looked at. */
  tickMs?: number;
  /** How long a live room's confirmation may lag before the edits go to
   * custody anyway. */
  unsettledMs?: number;
  /** The same for a window whose socket is not admitted (a handshake
   * usually completes within it). */
  offlineMs?: number;
  /** Requests kept on their way at once (see DEFAULT_MAX_IN_FLIGHT): a newer
   * state aborts the oldest beyond it. */
  maxInFlight?: number;
  now?: () => number;
  setTimer?: (run: () => void, ms: number) => unknown;
  clearTimer?: (timer: unknown) => void;
}

export interface RoomCustodyPump {
  /** The document changed (any transaction): looked at on the next tick. */
  noteChange: () => void;
  /** Stop: no further request is made. */
  dispose: () => void;
  /** For tests and diagnostics. */
  readonly sent: number;
  readonly stopped: string | null;
}

const DEFAULT_TICK_MS = 300;
const DEFAULT_UNSETTLED_MS = 1_200;
const DEFAULT_OFFLINE_MS = 600;









const DEFAULT_MAX_IN_FLIGHT = 2;
/** After a request failed for want of the server (a network error, 5xx). */
const RETRY_AFTER_MS = 2_000;












const LIVE_MIN_INTERVAL_MS = 2_000;

function sameBytes(left: Uint8Array | null, right: Uint8Array): boolean {
  if (!left || left.length !== right.length) return false;
  for (let index = 0; index < left.length; index += 1) if (left[index] !== right[index]) return false;
  return true;
}

/** A refusal that says the edits will never be taken (another epoch, a page
 * that is not a deck, edits that do not attach): stop, and leave the page to
 * the window's save and closing edits as before. */
function refusedForGood(error: unknown): boolean {
  const status = (error as { status?: unknown } | null)?.status;
  return typeof status === "number" && status >= 400 && status < 500 && status !== 408 && status !== 429;
}

/** Edits that did not attach to what the server holds now: the custody this
 * window adopted may not be readable there yet. Tried again a few times. */
const STALE_CODE = "page_causal_delta_base_stale";
const STALE_RETRIES = 3;
const STALE_RETRY_AFTER_MS = 5_000;



export function roomCustodyPumpEnabled(): boolean {
  try {
    return typeof window === "undefined" || window.localStorage?.getItem("mrai.deck.custodyPump") !== "off";
  } catch {
    return true;
  }
}

export function createRoomCustodyPump<E extends RoomCustodyEdits = RoomCustodyEdits>(
  options: RoomCustodyPumpOptions<E>,
): RoomCustodyPump {
  const tickMs = options.tickMs ?? DEFAULT_TICK_MS;
  const unsettledMs = options.unsettledMs ?? DEFAULT_UNSETTLED_MS;
  const offlineMs = options.offlineMs ?? DEFAULT_OFFLINE_MS;
  const maxInFlight = options.maxInFlight ?? DEFAULT_MAX_IN_FLIGHT;
  const now = options.now ?? (() => Date.now());
  const setTimer = options.setTimer ?? ((run, ms) => setTimeout(run, ms));
  const clearTimer = options.clearTimer ?? ((timer) => clearTimeout(timer as ReturnType<typeof setTimeout>));
  let dirty = false;
  let unsettledSince: number | null = null;
  let timer: unknown = null;
  // The requests on their way, oldest first.
  const inFlight: Array<{ controller: AbortController }> = [];
  let retryAt = 0;
  // The bytes of the last request the server took (or that is on its way):
  // the same document again is not sent again.
  let lastSent: Uint8Array | null = null;
  let sent = 0;
  let stopped: string | null = null;
  let staleRetries = 0;
  // The live pace: when the next request may leave while the room relays.
  let liveNextAt = 0;

  const schedule = () => {
    if (stopped || timer !== null) return;
    timer = setTimer(tick, tickMs);
  };

  const tick = () => {
    timer = null;
    if (stopped) return;
    let state: RoomCustodyState;
    try { state = options.state(); } catch { state = { live: false, settled: false }; }
    if (state.settled) {
      // The room holds every update of this window: nothing to hand over.
      dirty = false;
      unsettledSince = null;
      return;
    }
    const at = now();
    unsettledSince ??= at;
    // Watched until the room settles, so a lag is measured from when it began.
    if (!dirty) { schedule(); return; }
    if (at - unsettledSince < (state.live ? unsettledMs : offlineMs)) { schedule(); return; }
    if (at < retryAt) { schedule(); return; }
    // A live room's backstop: one on its way, paced by the last one's answer.
    if (state.live && (inFlight.length > 0 || at < liveNextAt)) { schedule(); return; }
    let edits: E | null = null;
    try { edits = options.read(); } catch { edits = null; }
    dirty = false;
    if (!edits || sameBytes(lastSent, edits.update)) { schedule(); return; }
    const bytes = edits.update;
    const previous = lastSent;
    lastSent = bytes;
    // The newest goes now; the oldest a newer one supersedes make room for it.
    while (inFlight.length >= Math.max(1, maxInFlight)) inFlight.shift()!.controller.abort();
    const entry = { controller: new AbortController() };
    inFlight.push(entry);
    const settle = () => { const at = inFlight.indexOf(entry); if (at >= 0) inFlight.splice(at, 1); };
    sent += 1;
    const sentAt = at;
    const answered = () => {
      const done = now();
      liveNextAt = Math.max(sentAt + LIVE_MIN_INTERVAL_MS, done + (done - sentAt));
    };
    let request: Promise<unknown>;
    try { request = options.send(edits, entry.controller.signal); } catch (error) { request = Promise.reject(error); }
    request.then((result) => {
      settle();
      answered();
      if (result === "superseded") {
        // The server took a newer request of this window instead (custody_gate):
        // this one was not kept. If nothing newer went since, it goes again.
        if (lastSent === bytes) lastSent = previous;
        dirty = true;
        schedule();
      }
    }, (error: unknown) => {
      settle();
      answered();
      // Superseded by a newer request: nothing to do.
      if (entry.controller.signal.aborted) return;
      const stale = (error as { code?: unknown } | null)?.code === STALE_CODE && staleRetries < STALE_RETRIES;
      if (stale) staleRetries += 1;
      if (!stale && refusedForGood(error)) {
        stopped = `refused:${(error as { status?: number }).status}`;
        if (timer !== null) { clearTimer(timer); timer = null; }
        return;
      }
      // Not taken: the same document goes again once the server answers --
      // when it said (429 Retry-After, live or not), else after the usual wait.
      if (lastSent === bytes) lastSent = previous;
      const retryAfter = (error as { status?: unknown; retryAfterMs?: unknown } | null);
      const serverSaid = retryAfter?.status === 429 && typeof retryAfter.retryAfterMs === "number" ? retryAfter.retryAfterMs : null;
      retryAt = now() + (serverSaid ?? (stale ? STALE_RETRY_AFTER_MS : RETRY_AFTER_MS));
      dirty = true;
      schedule();
    });
    schedule();
  };

  return {
    noteChange: () => {
      if (stopped) return;
      dirty = true;
      unsettledSince ??= now();
      schedule();
    },
    dispose: () => {
      stopped ??= "disposed";
      // Requests on their way are left to finish: each is a state the server may take.
      if (timer !== null) { clearTimer(timer); timer = null; }
    },
    get sent() { return sent; },
    get stopped() { return stopped; },
  };
}
