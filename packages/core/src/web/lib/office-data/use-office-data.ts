"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * React's view of the Office reference tables (`./store`): a surface that
 * draws with a table re-renders when it lands, and an editor holds its first
 * reveal until its tables are answered.
 */
import { useCallback, useEffect, useState, useSyncExternalStore } from "react";
import {
  loadOfficeData,
  officeData,
  officeDataSettled,
  officeDataVersion,
  subscribeOfficeData,
  type OfficeDataName,
  type OfficeDataWholeName,
} from "./store";

/** Re-renders the caller whenever a table (or a part of one) arrives; with
 * a name, only when that table does. */
export function useOfficeDataVersion(name?: OfficeDataName): number {
  const read = useCallback(() => officeDataVersion(name), [name]);
  return useSyncExternalStore(subscribeOfficeData, read, read);
}

/** A whole table, asked for on mount; `undefined` until it is here. */
export function useOfficeData(name: OfficeDataWholeName): unknown {
  useOfficeDataVersion(name);
  useEffect(() => {
    void loadOfficeData(name);
  }, [name]);
  return officeData(name);
}

/** How long an editor holds its first reveal for its tables. Past this it
 * opens with its fallback and redraws when they land. */
export const OFFICE_DATA_REVEAL_WAIT_MS = 4000;

/**
 * Asks for the named tables (an editor opening) and reports whether they are
 * answered: true at once for a guest or when they are already here, else
 * when they land, fail, or {@link OFFICE_DATA_REVEAL_WAIT_MS} has passed.
 */
export function useOfficeDataSettled(names: readonly OfficeDataWholeName[], enabled = true): boolean {
  useOfficeDataVersion(names.length === 1 ? names[0] : undefined);
  const key = names.join("|");
  const [expiredFor, setExpiredFor] = useState<string | null>(null);
  const settled = !enabled || officeDataSettled(names);
  useEffect(() => {
    if (!enabled) return undefined;
    for (const name of key.split("|")) void loadOfficeData(name as OfficeDataWholeName);
    if (officeDataSettled(key.split("|") as OfficeDataWholeName[])) return undefined;
    const timer = window.setTimeout(() => setExpiredFor(key), OFFICE_DATA_REVEAL_WAIT_MS);
    return () => window.clearTimeout(timer);
  }, [enabled, key]);
  return settled || expiredFor === key;
}
