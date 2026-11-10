"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The last picture each chart was drawn with, by the chart's content key
 * (its request without the size), for what a host draws BESIDE the picture
 * from the same drawing: a data table's legend keys take the colours the
 * engine painted the series in. `OfficeChartPicture` posts every picture it
 * shows; a reader re-renders when the drawing of its chart changes.
 */

import { useSyncExternalStore } from "react";
import type { ChartRenderPicture } from "../engine-host/charts-types";

const MAX_ENTRIES = 64;
const board = new Map<string, ChartRenderPicture>();
const listeners = new Map<string, Set<() => void>>();

export function postChartPicture(contentKey: string, picture: ChartRenderPicture): void {
  if (board.get(contentKey) === picture) return;
  board.delete(contentKey);
  board.set(contentKey, picture);
  while (board.size > MAX_ENTRIES) board.delete(board.keys().next().value as string);
  for (const listener of listeners.get(contentKey) ?? []) listener();
}

export function postedChartPicture(contentKey: string): ChartRenderPicture | undefined {
  return board.get(contentKey);
}

export function useChartPicturePosted(contentKey: string | null): ChartRenderPicture | undefined {
  return useSyncExternalStore(
    (listener) => {
      if (contentKey === null) return () => {};
      const set = listeners.get(contentKey) ?? new Set();
      set.add(listener);
      listeners.set(contentKey, set);
      return () => {
        set.delete(listener);
        if (!set.size) listeners.delete(contentKey);
      };
    },
    () => (contentKey === null ? undefined : board.get(contentKey)),
    () => undefined,
  );
}
