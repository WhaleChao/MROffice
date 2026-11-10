/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Logpoints: a breakpoint that reports instead of stopping.
 *
 * The oldest debugging habit is adding a print statement, and the reason every
 * debugger grew a logpoint is that doing it by hand means editing the file you
 * are trying to understand — then remembering to take it back out.
 *
 * A logpoint here is a message on a breakpoint line. The line still gets a
 * breakpoint (that is how the run captures locals there), but the panel does
 * not stop on it: it formats the message from the recorded locals, adds it to
 * the log, and continues.
 *
 * **Placeholders are names, for the same reason watches are.** Python here
 * records and replays, so at a stop there is no live interpreter to evaluate
 * `{len(items)}` against. `{name}` resolves from the frame's locals; anything
 * else stays LITERAL — a message reading `count is {len(items)}` visibly did
 * not substitute, which is a far better failure than a confidently wrong value
 * or a silently empty one.
 */

/** How long a logpoint message may be, and how many placeholders it may carry. */
export const MAX_LOGPOINT_MESSAGE_CHARS = 512;
export const MAX_LOGPOINT_PLACEHOLDERS = 20;
/** How many formatted lines one debug run keeps. */
export const MAX_LOGPOINT_LINES = 500;
/** How long one substituted value may be before it is clipped. */
export const MAX_LOGPOINT_VALUE_CHARS = 200;

/** `{name}` — one bare identifier, nothing to walk into and nothing to call. */
const PLACEHOLDER =
  /\{([$_\p{ID_Start}][$\u200c\u200d\p{ID_Continue}]*)\}/gu;

/** Whether a message is usable as a logpoint. Empty means "not a logpoint". */
export function isLogpointMessage(message: unknown): message is string {
  return (
    typeof message === "string" &&
    message.trim().length > 0 &&
    message.length <= MAX_LOGPOINT_MESSAGE_CHARS
  );
}

/**
 * Substitute `{name}` from the recorded locals.
 *
 * An unresolvable placeholder is left exactly as written. A frame whose locals
 * were never captured resolves nothing, which is the same visible outcome and
 * for the same reason: the message shows what it could not answer rather than
 * pretending it did.
 */
export function formatLogpoint(
  message: string,
  locals: Readonly<Record<string, string>> | null | undefined,
): string {
  if (!isLogpointMessage(message)) return "";
  let substitutions = 0;
  return message.replace(PLACEHOLDER, (whole, name: string) => {
    if (substitutions >= MAX_LOGPOINT_PLACEHOLDERS) return whole;
    if (!locals || !Object.prototype.hasOwnProperty.call(locals, name)) {
      return whole;
    }
    substitutions += 1;
    const value = locals[name];
    if (typeof value !== "string") return whole;
    return value.length > MAX_LOGPOINT_VALUE_CHARS
      ? `${value.slice(0, MAX_LOGPOINT_VALUE_CHARS)}…`
      : value;
  });
}

/** One reported line: which source line produced it, and what it said. */
export interface LogpointLine {
  readonly line: number;
  readonly text: string;
}

/**
 * Append a formatted line, keeping the log bounded.
 *
 * A logpoint inside a loop is the normal case, not the pathological one, so the
 * cap keeps the NEWEST lines: the interesting part of a run that logged ten
 * thousand times is almost always the end.
 */
export function appendLogpointLine(
  lines: readonly LogpointLine[],
  next: LogpointLine,
): LogpointLine[] {
  const appended = [...lines, next];
  return appended.length > MAX_LOGPOINT_LINES
    ? appended.slice(appended.length - MAX_LOGPOINT_LINES)
    : appended;
}
