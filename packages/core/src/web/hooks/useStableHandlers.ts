"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { useLayoutEffect, useRef, useState } from "react";

// eslint-disable-next-line @typescript-eslint/no-explicit-any
type Handlers = Record<string, (...args: any[]) => any>;

/**
 * The same functions on every render, each calling the implementation of the
 * latest render: what a memoized list row takes in place of a host's inline
 * handlers, so typing in the host's composer (which re-creates every
 * handler) no longer draws every row again.
 *
 * The set of names is fixed by the first render. The functions are for
 * events and effects; a row may also call one while rendering to read the
 * host's latest value (a ref's content, say), which is the value of the
 * last committed render.
 */
export function useStableHandlers<H extends Handlers>(handlers: H): H {
  const latest = useRef(handlers);
  useLayoutEffect(() => {
    latest.current = handlers;
  });
  const [stable] = useState(() => {
    const out: Handlers = {};
    for (const name of Object.keys(handlers)) {
      out[name] = (...args: unknown[]) => latest.current[name](...args);
    }
    return out as H;
  });
  return stable;
}
