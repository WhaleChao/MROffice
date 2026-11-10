/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Shrink text on overflow, measured on the slide's own drawing of the words.
 *
 * The size PowerPoint shrinks a body to is the first step of its ladder
 * (`presentation-autofit-ladder`) at which the laid-out words fit their box.
 * The editor used to estimate that height from an average glyph width; here
 * the words are laid out by the same renderer the slide uses, once per step,
 * in a hidden measuring layer (`PresentationAutofitMeasurer`), and the first
 * step that fits is kept for that block.
 *
 * A block object is replaced on every edit, so the block itself is the key:
 * a result belongs to exactly the words, runs, style and frame it was
 * measured for, under the theme and masters it was measured with. A host
 * asking for a block not measured yet gets nothing back (it draws its
 * estimate) and the block is queued; when the measure lands every subscribed
 * host draws again. Fonts arriving may change a measure: every block is
 * measured again, and only one whose measure changed draws again.
 */

import { useSyncExternalStore } from "react";

import type { PresentationBlock } from "@mrnote/core/web/components/console/editor/native-document/documentCommands.ts";
import type { SlideMaster } from "@mrnote/core/web/lib/presentation-master.ts";
import type { DeckTheme } from "@mrnote/core/web/lib/presentation-theme.ts";
import type { PresentationFontAliases } from "@mrnote/core/web/lib/presentation-font-faces.ts";

export type MeasuredAutofit = {
  readonly fontScale: number;
  readonly lineSpacingMultiplier: number;
};

/** What a host knows when it draws the block: enough for the measurer to draw it the same. */
export type AutofitMeasureContext = {
  readonly slide: { masterId?: string } | undefined;
  readonly masters: readonly SlideMaster[] | undefined;
  readonly theme: DeckTheme | undefined;
  readonly fontAliases: PresentationFontAliases;
};

export type AutofitMeasureRequest = AutofitMeasureContext & {
  readonly block: PresentationBlock;
  readonly bounds: { width: number; height: number };
};

type Entry = {
  readonly theme: unknown;
  readonly masters: unknown;
  readonly fit: MeasuredAutofit;
  /** What it was measured from, to measure it again when fonts arrive. */
  readonly request: AutofitMeasureRequest;
};

const results = new WeakMap<object, Entry>();
/** Blocks measured before fonts that arrived since: drawn with that measure
 * until it is taken again. */
const stale = new WeakSet<object>();
/**
 * The same results by what the words lay out by: a host that draws a copy of
 * the block (a show's inherited placeholder, a preview) finds its measure
 * here, rather than queueing a copy of every render for ever.
 */
let bySignature = new Map<string, Entry>();
const SIGNATURE_LIMIT = 4000;
const themeIds = new WeakMap<object, number>();
let nextThemeId = 1;

function contextId(value: unknown): number {
  if (typeof value !== "object" || value === null) return 0;
  let id = themeIds.get(value);
  if (id === undefined) {
    id = nextThemeId;
    nextThemeId += 1;
    themeIds.set(value, id);
  }
  return id;
}

function signatureOf(block: object, context: AutofitMeasureContext): string {
  const record = block as Record<string, unknown>;
  return JSON.stringify([
    record.text,
    record.runs,
    record.paragraphs,
    record.style,
    record.width,
    record.height,
    record.textInsets,
    record.placeholder,
    record.autofit,
    record.columnCount,
    (record.source as { nativeTextRect?: unknown } | undefined)?.nativeTextRect,
    contextId(context.theme),
    contextId(context.masters),
    (context.slide as { masterId?: unknown } | undefined)?.masterId,
  ]);
}
const pending = new Map<object, AutofitMeasureRequest>();
let measurers = 0;
let version = 0;
const listeners = new Set<() => void>();
const measurerListeners = new Set<() => void>();
let notifyScheduled = false;
let fontsWatched = false;

function notify(): void {
  if (notifyScheduled) return;
  notifyScheduled = true;
  queueMicrotask(() => {
    notifyScheduled = false;
    version += 1;
    for (const listener of Array.from(listeners)) listener();
  });
}

function wakeMeasurers(): void {
  queueMicrotask(() => {
    for (const listener of Array.from(measurerListeners)) listener();
  });
}

function watchFonts(): void {
  if (fontsWatched || typeof document === "undefined" || !document.fonts) return;
  fontsWatched = true;
  // A face that arrives after a measure may change the lines it was taken
  // on: every measured block is measured again, keeping its measure on the
  // slide meanwhile, and only one that comes out different draws again.
  // Dropping them all drew every slide and thumbnail of every deck open
  // again on each font that loaded anywhere on the desk (a sheet's Chinese
  // slices, a document's faces), twice: once on the estimate, once on the
  // new measure.
  document.fonts.addEventListener("loadingdone", () => {
    const again = new Map<string, AutofitMeasureRequest>();
    for (const [signature, entry] of bySignature) again.set(signature, entry.request);
    for (const request of again.values()) {
      stale.add(request.block);
      if (!pending.has(request.block)) pending.set(request.block, request);
    }
    if (pending.size > 0 && measurers > 0) wakeMeasurers();
  });
}

/** The measured step for this block, or undefined when it has not been measured here. */
export function measuredAutofit(block: object, context: AutofitMeasureContext): MeasuredAutofit | undefined {
  const entry = results.get(block);
  if (entry && entry.theme === context.theme && entry.masters === context.masters) return entry.fit;
  const same = bySignature.get(signatureOf(block, context));
  if (!same) return undefined;
  results.set(block, same);
  return same.fit;
}

/**
 * Queue a block for measuring. Only while a measurer is mounted: a surface
 * with none (a server render, a test) keeps drawing its estimate.
 */
export function requestAutofitMeasure(request: AutofitMeasureRequest): void {
  if (measurers === 0 || pending.has(request.block)) return;
  const signature = signatureOf(request.block, request);
  for (const queued of pending.values()) {
    if (signatureOf(queued.block, queued) === signature) return;
  }
  pending.set(request.block, request);
  wakeMeasurers();
}

/** The next requests to measure, oldest first, taken off the queue. */
export function takeAutofitMeasureRequests(limit: number): AutofitMeasureRequest[] {
  const taken: AutofitMeasureRequest[] = [];
  for (const [block, request] of pending) {
    if (taken.length >= limit) break;
    pending.delete(block);
    if (stale.has(block) || measuredAutofit(block, request) === undefined) taken.push(request);
  }
  return taken;
}

export function storeMeasuredAutofit(request: AutofitMeasureRequest, fit: MeasuredAutofit): void {
  const before = measuredAutofit(request.block, request);
  stale.delete(request.block);
  const entry = { theme: request.theme, masters: request.masters, fit, request };
  results.set(request.block, entry);
  if (bySignature.size >= SIGNATURE_LIMIT) bySignature = new Map();
  bySignature.set(signatureOf(request.block, request), entry);
  if (!before || before.fontScale !== fit.fontScale || before.lineSpacingMultiplier !== fit.lineSpacingMultiplier) {
    notify();
  }
}

/** A measurer mounts here; the returned function unmounts it. */
export function attachAutofitMeasurer(onRequests: () => void): () => void {
  measurers += 1;
  watchFonts();
  measurerListeners.add(onRequests);
  if (pending.size > 0) wakeMeasurers();
  return () => {
    measurers -= 1;
    measurerListeners.delete(onRequests);
    if (measurers === 0) pending.clear();
  };
}

function subscribe(listener: () => void): () => void {
  listeners.add(listener);
  return () => listeners.delete(listener);
}

const getVersion = () => version;

/** Draw again when a measure lands (a host's text reads its new size). */
export function useAutofitMeasureVersion(): number {
  return useSyncExternalStore(subscribe, getVersion, getVersion);
}
