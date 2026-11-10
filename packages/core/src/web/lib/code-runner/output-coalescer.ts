/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Pure helpers for PROGRESSIVE (streaming) multi-buffer host cell output (Jupyter parity:
 * a cell's stdout/stderr renders as it is produced, not only after the run ends).
 *
 * A running cell can emit a flood of tiny chunks (a tight `print` loop). Pushing
 * each one straight to React state would thrash rendering and freeze the UI, and
 * buffering an unbounded amount would balloon memory. {@link OutputCoalescer}
 * sits between the runner's incremental `onOutput(chunk, stream)` callback and the
 * component's live-output state and:
 *   - splits each text chunk into per-line entries tagged by stream, the SAME
 *     `{logs, logStreams}` shape the final RunResult uses, so the live view and
 *     the reconciled final view render identically (no special-case renderer);
 *   - coalesces rapid chunks behind a minimum flush interval (rate cap) while
 *     flushing immediately once enough has buffered (so the first output appears
 *     promptly and a steady stream stays live);
 *   - caps the total buffered characters AND line count so a runaway loop can
 *     neither freeze the UI nor exhaust memory — once capped it stops appending
 *     and records a single truncation marker.
 *
 * Everything here is synchronous and side-effect free EXCEPT the optional timer
 * used for rate-limited flushing, which is fully injectable (so it is unit-tested
 * deterministically with fake timers and works in any environment).
 *
 * Backward-compatibility / fidelity contract:
 *   - `logs` is a plain `string[]`; the parallel OPTIONAL `logStreams` tags each
 *     line's stream by index (missing tag ⇒ stdout), exactly like RunResult.
 *   - The coalescer NEVER persists anything; it only shapes transient live state.
 *     The runner's final batched RunResult remains the source of truth that gets
 *     reconciled in + persisted, so streaming is purely additive.
 */

import type { LogStream } from "./cell-output-format";

/** A live, renderable snapshot of the output produced so far. */
export interface LiveOutput {
  /** Per-line text captured so far (one printed line per entry). */
  readonly logs: string[];
  /** Per-line stream tag aligned to `logs` by index ("stdout" | "stderr"). */
  readonly logStreams: LogStream[];
  /** True once a size/line cap dropped further output (a marker line is kept). */
  readonly truncated: boolean;
}

export interface OutputCoalescerOptions {
  /**
   * Minimum gap between flushes (ms). Chunks arriving within this window are
   * coalesced into the next flush rather than each triggering a render. A flush
   * is still forced immediately when {@link flushThresholdChars} is exceeded so
   * the very first output (and a steady stream) appears without waiting a full
   * interval. Default 50ms.
   */
  readonly flushIntervalMs?: number;
  /**
   * Buffered-since-last-flush character count that forces an immediate flush
   * (so output isn't held back when a lot arrives at once). Default 2048.
   */
  readonly flushThresholdChars?: number;
  /**
   * Hard cap on total buffered characters across all lines. Once reached the
   * coalescer stops appending and records a truncation marker. Default 1,000,000
   * (~1MB) — generous for real output, bounded against a runaway loop.
   */
  readonly maxTotalChars?: number;
  /** Hard cap on the number of buffered lines. Default 50,000. */
  readonly maxLines?: number;
  /**
   * Schedule a deferred flush. Injectable for deterministic tests; defaults to
   * `setTimeout`. Returns a handle passed to {@link clearTimer}.
   */
  readonly setTimer?: (fn: () => void, ms: number) => unknown;
  /** Cancel a scheduled flush. Defaults to `clearTimeout`. */
  readonly clearTimer?: (handle: unknown) => void;
  /** Monotonic clock for the rate gate (ms). Defaults to `Date.now`. */
  readonly now?: () => number;
}

export const DEFAULT_FLUSH_INTERVAL_MS = 50;
export const DEFAULT_FLUSH_THRESHOLD_CHARS = 2_048;
export const DEFAULT_MAX_TOTAL_CHARS = 1_000_000;
export const DEFAULT_MAX_LINES = 50_000;
/** Marker appended once when output is capped, so the user knows it was cut. */
export const TRUNCATION_MARKER =
  "… [output truncated: streaming cap reached — re-run to see the full output]";

/**
 * Split an incremental text chunk into per-line entries, preserving partial
 * lines across calls. A chunk rarely arrives on a clean line boundary (a single
 * `print` may be delivered with or without its trailing newline; a chunk may
 * carry several lines), so the coalescer keeps a per-stream "pending" remainder:
 *
 *   appendChunk("ab")        → carry "ab"            (no newline yet)
 *   appendChunk("c\nde")     → emit "abc", carry "de"
 *   appendChunk("\n")        → emit "de"
 *
 * Returns the COMPLETE lines to emit and the new carry. Pure (no state). A
 * trailing newline yields its line but no empty carry; an empty chunk is a
 * no-op. Exported for direct unit testing.
 */
export function splitChunkLines(
  carry: string,
  chunk: string,
): { lines: string[]; carry: string } {
  if (chunk === "") return { lines: [], carry };
  const combined = carry + chunk;
  // `split("\n")` yields N+1 parts for N newlines; the LAST part is the new
  // carry (the unterminated remainder — "" when the chunk ended with "\n").
  const parts = combined.split("\n");
  const nextCarry = parts.pop() ?? "";
  return { lines: parts, carry: nextCarry };
}

/**
 * Buffers incremental `(chunk, stream)` output, coalesces rapid chunks behind a
 * rate gate, caps total size, and notifies a single `onFlush(snapshot)` callback
 * with the latest renderable {@link LiveOutput}. Stateful but self-contained;
 * the only side effect (a deferred-flush timer) is injectable.
 *
 * Usage: construct with `onFlush` (push into React state), call `push(chunk,
 * stream)` from the runner's `onOutput`, and call `finalize()` once the run ends
 * to flush any pending carry/buffer synchronously. The component then reconciles
 * the authoritative final RunResult over the live snapshot.
 */
export class OutputCoalescer {
  private readonly logs: string[] = [];
  private readonly logStreams: LogStream[] = [];
  /** Per-stream unterminated remainder awaiting its newline. */
  private carryStdout = "";
  private carryStderr = "";
  private totalChars = 0;
  private truncated = false;
  private dirty = false;
  private timer: unknown = null;
  private lastFlushAt = -Infinity;

  private readonly flushIntervalMs: number;
  private readonly flushThresholdChars: number;
  private readonly maxTotalChars: number;
  private readonly maxLines: number;
  private readonly setTimerFn: (fn: () => void, ms: number) => unknown;
  private readonly clearTimerFn: (handle: unknown) => void;
  private readonly nowFn: () => number;
  /** Characters accumulated since the last flush (drives the threshold gate). */
  private bufferedSinceFlush = 0;

  constructor(
    private readonly onFlush: (snapshot: LiveOutput) => void,
    options: OutputCoalescerOptions = {},
  ) {
    this.flushIntervalMs = options.flushIntervalMs ?? DEFAULT_FLUSH_INTERVAL_MS;
    this.flushThresholdChars =
      options.flushThresholdChars ?? DEFAULT_FLUSH_THRESHOLD_CHARS;
    this.maxTotalChars = options.maxTotalChars ?? DEFAULT_MAX_TOTAL_CHARS;
    this.maxLines = options.maxLines ?? DEFAULT_MAX_LINES;
    this.setTimerFn =
      options.setTimer ??
      ((fn, ms) => setTimeout(fn, ms) as unknown);
    this.clearTimerFn =
      options.clearTimer ?? ((handle) => clearTimeout(handle as never));
    this.nowFn = options.now ?? (() => Date.now());
  }

  /**
   * Append an incremental chunk for a stream. Splits into complete lines (a
   * partial line is carried until its newline), enforces the size/line caps, and
   * schedules or forces a flush per the rate gate. A no-op once capped.
   */
  push(chunk: string, stream: LogStream): void {
    if (typeof chunk !== "string" || chunk === "") return;
    if (this.truncated) return;
    const isStderr = stream === "stderr";
    const carry = isStderr ? this.carryStderr : this.carryStdout;
    const { lines, carry: nextCarry } = splitChunkLines(carry, chunk);
    if (isStderr) this.carryStderr = nextCarry;
    else this.carryStdout = nextCarry;
    for (const line of lines) this.appendLine(line, stream);
    // Count the chunk's characters toward the rate-gate threshold.
    this.bufferedSinceFlush += chunk.length;
    this.scheduleFlush();
  }

  /** Append one complete line, enforcing the total-size and line caps. */
  private appendLine(line: string, stream: LogStream): void {
    if (this.truncated) return;
    if (
      this.logs.length >= this.maxLines ||
      this.totalChars + line.length > this.maxTotalChars
    ) {
      this.markTruncated();
      return;
    }
    this.logs.push(line);
    this.logStreams.push(stream);
    this.totalChars += line.length;
    this.dirty = true;
  }

  /** Record the one-time truncation marker line (idempotent). */
  private markTruncated(): void {
    if (this.truncated) return;
    this.truncated = true;
    // The marker is stderr-tagged so it stands out (Jupyter shows warnings red).
    this.logs.push(TRUNCATION_MARKER);
    this.logStreams.push("stderr");
    this.dirty = true;
  }

  /**
   * Flush now if the rate gate allows (enough time elapsed) OR enough has
   * buffered; otherwise arm a deferred flush so the pending output isn't lost.
   */
  private scheduleFlush(): void {
    if (!this.dirty) return;
    const elapsed = this.nowFn() - this.lastFlushAt;
    if (
      elapsed >= this.flushIntervalMs ||
      this.bufferedSinceFlush >= this.flushThresholdChars ||
      this.truncated
    ) {
      this.flush();
      return;
    }
    if (this.timer === null) {
      const wait = Math.max(0, this.flushIntervalMs - elapsed);
      this.timer = this.setTimerFn(() => {
        this.timer = null;
        this.flush();
      }, wait);
    }
  }

  /** Emit the current snapshot to `onFlush` (no-op when nothing changed). */
  private flush(): void {
    if (this.timer !== null) {
      this.clearTimerFn(this.timer);
      this.timer = null;
    }
    if (!this.dirty) return;
    this.dirty = false;
    this.lastFlushAt = this.nowFn();
    this.bufferedSinceFlush = 0;
    this.onFlush(this.snapshot());
  }

  /**
   * Flush any pending carry (an unterminated final line with no trailing
   * newline) and the buffer synchronously, then return the final live snapshot.
   * Call once when the run ends (before reconciling the authoritative result).
   */
  finalize(): LiveOutput {
    // Emit any unterminated remainder as its own final line per stream so the
    // last `print(..., end="")` / partial write isn't dropped from the live view.
    if (this.carryStdout !== "") {
      this.appendLine(this.carryStdout, "stdout");
      this.carryStdout = "";
    }
    if (this.carryStderr !== "") {
      this.appendLine(this.carryStderr, "stderr");
      this.carryStderr = "";
    }
    // Force a final flush regardless of the rate gate.
    if (this.dirty) {
      if (this.timer !== null) {
        this.clearTimerFn(this.timer);
        this.timer = null;
      }
      this.dirty = false;
      this.lastFlushAt = this.nowFn();
      this.bufferedSinceFlush = 0;
      this.onFlush(this.snapshot());
    }
    return this.snapshot();
  }

  /** Cancel any pending flush (e.g. the cell unmounted / run aborted). */
  dispose(): void {
    if (this.timer !== null) {
      this.clearTimerFn(this.timer);
      this.timer = null;
    }
  }

  /** A copy of the current renderable state (callers must not mutate it). */
  snapshot(): LiveOutput {
    return {
      logs: this.logs.slice(),
      logStreams: this.logStreams.slice(),
      truncated: this.truncated,
    };
  }

  /** Whether any output has been buffered (for a "running…" vs "produced" cue). */
  hasOutput(): boolean {
    return this.logs.length > 0 || this.carryStdout !== "" || this.carryStderr !== "";
  }
}

/**
 * The shape a runner's incremental output callback takes. A runner that supports
 * streaming calls this as each stdout/stderr chunk is produced; when absent the
 * runner stays fully backward-compatible (final batch only). Kept as a named
 * type so the runners and the component share one signature.
 */
export type OnOutput = (chunk: string, stream: LogStream) => void;
