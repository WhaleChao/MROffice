/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
export function requestQuietCallback(task: () => void, timeoutMs = 2_000): () => void {
  if (typeof window === "undefined") {
    const timer = setTimeout(task, 0);
    return () => clearTimeout(timer);
  }
  if (typeof window.requestIdleCallback === "function") {
    const handle = window.requestIdleCallback(() => task(), { timeout: timeoutMs });
    return () => window.cancelIdleCallback(handle);
  }
  let cancelled = false;
  let frame = 0;
  let calm = 0;
  let last = performance.now();
  const deadline = last + timeoutMs;
  const run = () => {
    if (cancelled) return;
    cancelled = true;
    window.setTimeout(task, 0);
  };
  const tick = (now: number) => {
    if (cancelled) return;
    calm = now - last <= 34 ? calm + 1 : 0;
    last = now;
    if (calm >= 2 || now >= deadline) {
      run();
      return;
    }
    frame = window.requestAnimationFrame(tick);
  };
  frame = window.requestAnimationFrame(tick);
  // A hidden page draws no frames: the deadline still runs the task.
  const fallback = window.setTimeout(run, timeoutMs + 50);
  return () => {
    cancelled = true;
    window.cancelAnimationFrame(frame);
    window.clearTimeout(fallback);
  };
}
