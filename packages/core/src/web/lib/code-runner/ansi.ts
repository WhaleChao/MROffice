/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Minimal ANSI/VT escape handling for runner output (Jupyter parity).
 *
 * Python libraries like tqdm, rich and colorama emit SGR colour codes plus
 * carriage-return progress rewrites. Without handling, those render as
 * `\x1b[32m…` garbage in the output panel. This module is framework-free:
 *
 *   - {@link parseAnsi} turns one log line into styled text segments
 *     (SGR colours/bold/italic/underline/dim; 16-colour, 256-colour and
 *     24-bit truecolour forms).
 *   - Non-SGR escapes (cursor movement, erase-line, OSC titles…) are
 *     STRIPPED — they position a terminal cursor we don't have.
 *   - A bare carriage return mid-line keeps only the text after the LAST
 *     `\r` (terminal overwrite semantics), so a tqdm progress line shows
 *     its final state instead of every intermediate frame glued together.
 *
 * Rendering callers branch on {@link hasAnsiOrCr} so plain output keeps its
 * existing zero-cost `<pre>{line}</pre>` path byte-for-byte.
 */

export interface AnsiSegment {
  text: string;
  /** CSS colour for the foreground, when an SGR colour is active. */
  color?: string;
  /** CSS colour for the background, when an SGR background is active. */
  background?: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  dim?: boolean;
}

/** Fast pre-check: does this line need ANSI/CR-aware rendering at all? */
export function hasAnsiOrCr(text: string): boolean {
  return text.includes("\x1b[") || text.includes("\x1b]") || text.includes("\r");
}

// The standard 16-colour palette (VGA-ish hues tuned for light backgrounds —
// the same family Jupyter Lab uses; bright variants for 90–97/100–107).
const BASE_COLORS: readonly string[] = [
  "#3f3f3f", // black (softened so it stays visible on dark themes too)
  "#b91c1c", // red
  "#15803d", // green
  "#a16207", // yellow
  "#1d4ed8", // blue
  "#7e22ce", // magenta
  "#0e7490", // cyan
  "#6b7280", // white (rendered as gray on light backgrounds)
];
const BRIGHT_COLORS: readonly string[] = [
  "#6b7280", // bright black
  "#ef4444", // bright red
  "#22c55e", // bright green
  "#eab308", // bright yellow
  "#3b82f6", // bright blue
  "#a855f7", // bright magenta
  "#06b6d4", // bright cyan
  "#9ca3af", // bright white
];

/** xterm 256-colour index → CSS colour. */
function color256(index: number): string | undefined {
  if (!Number.isInteger(index) || index < 0 || index > 255) return undefined;
  if (index < 8) return BASE_COLORS[index];
  if (index < 16) return BRIGHT_COLORS[index - 8];
  if (index < 232) {
    // 6×6×6 colour cube; xterm levels 0,95,135,175,215,255.
    const n = index - 16;
    const level = (v: number) => (v === 0 ? 0 : 55 + v * 40);
    const r = level(Math.floor(n / 36));
    const g = level(Math.floor(n / 6) % 6);
    const b = level(n % 6);
    return `rgb(${r}, ${g}, ${b})`;
  }
  // Grayscale ramp 232–255: 8, 18, …, 238.
  const gray = 8 + (index - 232) * 10;
  return `rgb(${gray}, ${gray}, ${gray})`;
}

interface SgrState {
  color?: string;
  background?: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  dim?: boolean;
}

/** Apply one SGR parameter list (the numbers in `\x1b[…m`) to a style state. */
function applySgr(state: SgrState, params: number[]): SgrState {
  const next: SgrState = { ...state };
  for (let i = 0; i < params.length; i++) {
    const p = params[i];
    if (p === 0) {
      return {};
    } else if (p === 1) next.bold = true;
    else if (p === 2) next.dim = true;
    else if (p === 3) next.italic = true;
    else if (p === 4) next.underline = true;
    else if (p === 21 || p === 22) {
      delete next.bold;
      delete next.dim;
    } else if (p === 23) delete next.italic;
    else if (p === 24) delete next.underline;
    else if (p >= 30 && p <= 37) next.color = BASE_COLORS[p - 30];
    else if (p === 39) delete next.color;
    else if (p >= 40 && p <= 47) next.background = BASE_COLORS[p - 40];
    else if (p === 49) delete next.background;
    else if (p >= 90 && p <= 97) next.color = BRIGHT_COLORS[p - 90];
    else if (p >= 100 && p <= 107) next.background = BRIGHT_COLORS[p - 100];
    else if (p === 38 || p === 48) {
      // Extended colour: 38;5;n (256) or 38;2;r;g;b (truecolour).
      const target = p === 38 ? "color" : "background";
      const mode = params[i + 1];
      if (mode === 5) {
        const value = color256(params[i + 2]);
        if (value) next[target] = value;
        i += 2;
      } else if (mode === 2) {
        const [r, g, b] = [params[i + 2], params[i + 3], params[i + 4]];
        if ([r, g, b].every((v) => Number.isInteger(v) && v >= 0 && v <= 255)) {
          next[target] = `rgb(${r}, ${g}, ${b})`;
        }
        i += 4;
      }
    }
    // Anything else (blink/reverse/conceal/fonts) is ignored.
  }
  return next;
}

// CSI: ESC [ params final-byte. OSC: ESC ] … (BEL | ESC \). Plus lone ESC ops.
const CSI_RE = /^\x1b\[([0-9;]*)([@-~])/;
const OSC_RE = /^\x1b\][^\x07\x1b]*(?:\x07|\x1b\\)?/;

/**
 * Terminal overwrite semantics for one logical line: a bare `\r` (not part of
 * `\r\n` — runners already split lines) rewinds to column 0, so only the text
 * after the LAST `\r` survives. SGR codes BEFORE the cut still apply (they set
 * state, not text), so styling is resolved during the main parse instead.
 */
function lastCarriageReturnFrame(text: string): string {
  const lastCr = text.lastIndexOf("\r");
  return lastCr === -1 ? text : text.slice(lastCr + 1);
}

/** Cap parsed segments per line — a hostile line cannot allocate unbounded spans. */
const MAX_SEGMENTS = 1_000;

/**
 * Parse one log line into styled segments. Always returns at least one
 * segment unless the visible text is empty (then `[]`).
 */
export function parseAnsi(text: string): AnsiSegment[] {
  // Overwrite first so styling state from discarded frames cannot leak text,
  // BUT keep any SGR sequences from before the cut by re-scanning the full
  // string for state while only emitting post-cut text. Simpler equivalent:
  // tqdm-style lines restate their colours after \r, so cutting first is
  // correct in practice and keeps the parser single-pass.
  const frame = lastCarriageReturnFrame(text);
  const segments: AnsiSegment[] = [];
  let state: SgrState = {};
  let plain = "";
  const flush = () => {
    if (!plain) return;
    if (segments.length < MAX_SEGMENTS) {
      segments.push({ text: plain, ...state });
    } else {
      // Beyond the cap, merge into the last segment unstyled-appended rather
      // than dropping user output.
      segments[segments.length - 1].text += plain;
    }
    plain = "";
  };
  let i = 0;
  while (i < frame.length) {
    const ch = frame[i];
    if (ch === "\x1b") {
      const rest = frame.slice(i);
      const csi = CSI_RE.exec(rest);
      if (csi) {
        if (csi[2] === "m") {
          flush();
          const params = csi[1] === "" ? [0] : csi[1].split(";").map((n) => Number(n || "0"));
          state = applySgr(state, params);
        }
        // Non-SGR CSI (cursor moves, erase line…) is stripped silently.
        i += csi[0].length;
        continue;
      }
      const osc = OSC_RE.exec(rest);
      if (osc) {
        i += osc[0].length;
        continue;
      }
      // Lone ESC or unsupported two-byte escape: drop the ESC byte.
      i += 1;
      continue;
    }
    plain += ch;
    i += 1;
  }
  flush();
  return segments;
}

/** Visible text with all escapes stripped (for plain-text consumers/tests). */
export function stripAnsi(text: string): string {
  return parseAnsi(text)
    .map((segment) => segment.text)
    .join("");
}
