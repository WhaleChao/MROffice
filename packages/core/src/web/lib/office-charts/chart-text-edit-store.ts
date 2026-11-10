/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The chart text being typed in each editor's chart, by the host's key for
 * the chart (the key its Current Selection is kept under,
 * `chart-selection-store`): a host begins the edit (a click on the selected
 * title, a double-click, F2) and the chart's picture makes that text of its
 * own drawing editable, in Docs, the deck and the worksheet alike.
 */
import { useSyncExternalStore } from "react";

import type { ChartTextPart } from "../engine-host/charts-types";
import type { OfficeChartTextValue } from "./chart-text-editing";

/** A chart text typed in place: the title, an axis title, or one data label
 * (`label:series:point`, the chart's own series and point). */
export type OfficeChartEditablePart = ChartTextPart | `label:${number}:${number}`;

/** A title or an axis title (not a data label). */
export function isOfficeChartTitlePart(part: OfficeChartEditablePart): part is ChartTextPart {
  return part === "title" || part === "axisX" || part === "axisY";
}

export interface OfficeChartTextEdit {
  part: OfficeChartEditablePart;
  /** Where the click that began the edit landed: the caret goes there.
   * Absent: every character is selected (an edit begun by keyboard). */
  at?: { clientX: number; clientY: number };
  /** The editable text's accessible name and test id. */
  label: string;
  testId?: string;
  /** Called once when the edit ends: the words as typed (Escape keeps them
   * too, as Office's Esc leaves a text box with what was typed). `commit`
   * false: the words are as they were, nothing to write. */
  onFinish: (commit: boolean, value: OfficeChartTextValue) => void;
  /** Distinguishes two edits begun on the same text. */
  serial: number;
}

const edits = new Map<string, OfficeChartTextEdit>();
const listeners = new Set<() => void>();
let serial = 0;

function notify(): void {
  for (const listener of listeners) listener();
}

/** Begin typing in a chart text (replaces an edit under way on that chart). */
export function beginOfficeChartTextEdit(
  key: string | null | undefined,
  edit: Omit<OfficeChartTextEdit, "serial">,
): void {
  if (!key) return;
  serial += 1;
  edits.set(key, { ...edit, serial });
  notify();
}

/** The edit under way on the chart, if any. */
export function officeChartTextEditOf(key: string | null | undefined): OfficeChartTextEdit | null {
  return key ? (edits.get(key) ?? null) : null;
}

/** The edit is over (the picture calls it after `onFinish`). */
export function endOfficeChartTextEdit(key: string | null | undefined, edit?: OfficeChartTextEdit): void {
  if (!key) return;
  const current = edits.get(key);
  if (!current || (edit && current.serial !== edit.serial)) return;
  edits.delete(key);
  notify();
}

function subscribe(listener: () => void): () => void {
  listeners.add(listener);
  return () => {
    listeners.delete(listener);
  };
}

/** The chart text being typed, live. */
export function useOfficeChartTextEdit(key: string | null | undefined): OfficeChartTextEdit | null {
  return useSyncExternalStore(
    subscribe,
    () => officeChartTextEditOf(key),
    () => null,
  );
}
