"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { useEffect, useRef, useState, type CSSProperties, type RefObject } from "react";
import { contentKey } from "@mrnote/core/web/lib/content-hash.ts";
import { clampRowHeight } from "@mrnote/core/web/lib/spreadsheet/grid-sizing.ts";
import { SPREADSHEET_WRAP_LINE_HEIGHT } from "@mrnote/core/web/lib/spreadsheet/cell-format.ts";
import { measureNativeTextHeightBatch, measureNativeAutoRowHeight, measureNativeParagraphLineHeight, measureNativeTextHeight, type NativeTextLayoutRequest, type NativeTextPart } from "../../../../lib/spreadsheet/native-text-layout";
import { clampNativeAutomaticRowHeight, nativeTextAdvances, resolveSpreadsheetTextStyle } from "@mrnote/core/web/lib/spreadsheet/native-typography.ts";
import { useSpreadsheetTextEngineEpoch } from "./useSpreadsheetTextEngineEpoch";
import { fontLoadReach } from "@mrnote/core/web/lib/spreadsheet/font-load-reach.ts";

export interface AutoRowMeasurement {
  row: number;
  width: number;
  text: string;
  style: CSSProperties;
  dpi?: number;
  paddingX?: number;
  nativeParagraph?: boolean;
  baseHeightOnly?: boolean;
  /** A literal that does not wrap: measured as one line of its font, however
   * long it is (Excel grows the row for a large font, never for the width). */
  singleLine?: boolean;
  /** Counter-clockwise text rotation in degrees: the row fits the rotated
   *  single line (|sin|·text width + |cos|·line height), as Excel's does. */
  rotation?: number;
  /** Vertical Text (stacked upright letters): the row fits the column of
   *  letters, as Excel's does. */
  stacked?: boolean;
  /** Base-height calibration: the browser's line box for the workbook's default
   * font divided into the row height the workbook stores for it, so a cell in
   * that font measures exactly the stored default and a larger font scales
   * from it (Excel's own row height, not the browser's line box). */
  scale?: number;
  numericValue?: boolean;
  parts?: NativeTextPart[];
}

/**
 * What one cell's text asks of its row, by everything the measurement reads
 * (the text, its font and runs, the width, the calibration), not by the row
 * it sits in. An edit changes one cell; without this, every wrapped or
 * tall-font cell of the sheet was laid out again in the probe after each
 * entry, a forced layout per cell (1.7 s per Enter on a 2,300-cell sheet).
 * Dropped when a web font finishes loading, which changes what text measures.
 */
const measuredHeights = new Map<string, { height: number; families: readonly string[]; text: string }>();
const MEASURED_HEIGHTS_LIMIT = 50_000;

function cssFamilies(value: unknown, into: Set<string>): void {
  if (typeof value !== "string") return;
  for (const part of value.split(",")) {
    const family = part.trim().replace(/^["']|["']$/g, "").trim().toLowerCase();
    if (family) into.add(family);
  }
}

/** Every family a cell's text may be drawn in, in stack order: the one the
 * grid resolves it to (its face, the stand-ins, the fallbacks), its own, its
 * runs'. */
function itemFamilies(item: AutoRowMeasurement): string[] {
  const out = new Set<string>();
  cssFamilies(resolveSpreadsheetTextStyle(item.style, item.text, item.dpi).fontFamily, out);
  cssFamilies(item.style.fontFamily, out);
  for (const part of item.parts ?? []) cssFamilies((part.style as CSSProperties | undefined)?.fontFamily, out);
  return [...out];
}

/** All the text a cell's measurement draws. */
function itemText(item: AutoRowMeasurement): string {
  return item.parts?.length ? item.parts.map((part) => part.text).join("") : item.text;
}

if (typeof document !== "undefined") {
  // A face that finished changes what the text drawn in it measures: those
  // heights go. Dropping every height on every face kept a sheet whose
  // workbook registers thousands of faces, some loading at any moment,
  
  // first sheet, a million forced layouts in one screenshot).
  document.fonts?.addEventListener("loadingdone", (event) => {
    const reach = fontLoadReach(event);
    if (!reach) return;
    for (const [key, entry] of measuredHeights) if (reach.touches(entry.families, entry.text)) measuredHeights.delete(key);
  });
}

/** View-only measurements: never write browser/font-dependent heights to a workbook. */
/** The content key of `measurements`, its parsed copy in `into` when the key
 * changed (a stable identity for equal measurements). */
function readMeasurements(
  measurements: readonly AutoRowMeasurement[],
  into: { current: { key: string; list: AutoRowMeasurement[] } | null },
): string {
  const text = JSON.stringify(measurements);
  const measurementKey = contentKey(text);
  if (into.current?.key !== measurementKey) {
    into.current = { key: measurementKey, list: JSON.parse(text) as AutoRowMeasurement[] };
  }
  return measurementKey;
}

function autoRowParts(item: AutoRowMeasurement): NativeTextPart[] {
  // A trailing line break is a last, empty line. The browser lays out
  // no box for it, so the measured text carries a space on that line,
  // as the painted cell does.
  const measuredText = item.text.endsWith("\n") ? `${item.text}\u00a0` : item.text;
  return item.parts
    ? item.parts.map((part, index, all) =>
        index === all.length - 1 && part.text.endsWith("\n") ? { ...part, text: `${part.text}\u00a0` } : part)
    : [{ text: measuredText }];
}

/** The height one cell asks of its row, `text` being a laid-out probe; nothing
 * when the probe is not laid out. */
function measureAutoRowItem(item: AutoRowMeasurement, defaultHeight: number, text: HTMLElement): number | null {
  const measuredText = item.text.endsWith("\n") ? `${item.text}\u00a0` : item.text;
  const parts = autoRowParts(item);
  if (item.baseHeightOnly) {
    const raw = item.dpi ? measureNativeParagraphLineHeight(parts, item.style, item.dpi) : undefined;
    const measured = raw !== undefined && item.scale ? Math.round(raw * item.scale * 2) / 2 : raw;
    return measured !== undefined && measured > defaultHeight ? clampNativeAutomaticRowHeight(measured) : 0;
  }
  const automaticHeight = item.nativeParagraph && item.dpi
    ? measureNativeAutoRowHeight(parts, item.style, item.singleLine ? 1e6 : item.width - 5 * 96 / item.dpi, item.dpi, item.numericValue)
    : undefined;
  if (automaticHeight !== undefined) return automaticHeight;
  const nativeHeight = item.nativeParagraph && item.dpi
    ? measureNativeTextHeight(parts, item.style, item.singleLine ? 1e6 : item.width - 5 * 96 / item.dpi, item.dpi)
    : undefined;
  if (nativeHeight !== undefined) return clampRowHeight(Math.ceil(nativeHeight + 4), defaultHeight);
  text.style.cssText = "";
  const measuredStyle = resolveSpreadsheetTextStyle(item.style, item.text, item.dpi);
  Object.assign(text.style, {
    fontFamily: measuredStyle.fontFamily ?? "inherit",
    fontWeight: String(measuredStyle.fontWeight ?? "normal"),
    fontStyle: measuredStyle.fontStyle ?? "normal",
    fontSize: typeof measuredStyle.fontSize === "number" ? `${measuredStyle.fontSize}px` : measuredStyle.fontSize ?? "12px",
    // A wrapped cell is painted at the wrapped-text pitch, so it is
    // measured at it: measured at the font's "normal" pitch, nine
    // painted lines fit a row of eight and the first line was cut off
    // above a bottom-aligned paragraph (user audit S-001).
    // (A native paragraph is painted at "normal", and measured so.)
    lineHeight: item.singleLine || item.nativeParagraph ? "normal" : String(SPREADSHEET_WRAP_LINE_HEIGHT),
    whiteSpace: item.singleLine ? "pre" : "pre-wrap", overflowWrap: "anywhere",
    wordBreak: "break-word", boxSizing: "content-box",
    // Normal view has 9px padding on each side; Page Layout has less.
    // One shared height must fit both, independent of the active view.
    // The painted paragraph also loses the cell's 1px gridline.
    width: `${Math.max(1, item.width - 2 * (item.paddingX ?? 9) - (item.singleLine || item.nativeParagraph ? 0 : 1))}px`, height: "auto",
  });
  if (item.dpi) {
    const font = `${text.style.fontStyle} ${text.style.fontWeight} ${text.style.fontSize} ${text.style.fontFamily}`;
    const runs = nativeTextAdvances(measuredText,font,item.dpi);
    text.style.fontKerning = "none";
    text.replaceChildren(...runs.map(run => {
      const span = document.createElement("span"); span.textContent = run.text; span.style.letterSpacing = `${run.spacing}px`; return span;
    }));
  } else if (item.parts) {
    // A cell with formatted runs is measured with each run in its own
    // font (a 24 px run makes a taller line than the cell's base),
    // not as plain text in the base font.
    text.replaceChildren(...parts.map(part => {
      const span = document.createElement("span");
      span.textContent = part.text;
      for (const [property, value] of Object.entries(part.style ?? {})) {
        if (value === undefined || value === null) continue;
        const cssValue = typeof value === "number" && property !== "fontWeight" && property !== "lineHeight"
          ? `${value}px` : String(value);
        (span.style as unknown as Record<string, string>)[property] = cssValue;
      }
      return span;
    }));
  } else text.textContent = measuredText;
  if (item.stacked) {
    text.style.writingMode = "vertical-rl";
    text.style.textOrientation = "upright";
    text.style.whiteSpace = "pre";
    text.style.width = "auto";
  } else if (item.rotation) {
    // The turned line's own length, not the column's: a box as wide
    // as the column reports the column's width for a short line, and
    // a 90° row grew to the column's width.
    text.style.width = "max-content";
  }
  const measured = item.stacked
    ? text.scrollHeight
    : item.rotation
    ? Math.abs(Math.sin((item.rotation * Math.PI) / 180)) * text.scrollWidth +
      Math.abs(Math.cos((item.rotation * Math.PI) / 180)) * text.scrollHeight
    : text.scrollHeight;
  // A probe that is not laid out (a hidden window) measures nothing:
  // that is no answer to keep.
  return measured > 0 ? clampRowHeight(Math.ceil(measured + 4), defaultHeight) : null;
}

function createAutoRowProbe(parent: HTMLElement) {
  const probe = document.createElement("div");
  probe.setAttribute("aria-hidden", "true");
  probe.setAttribute("data-spreadsheet-auto-row-probe", "");
  Object.assign(probe.style, {
    position: "absolute", left: "0", top: "0", visibility: "hidden",
    pointerEvents: "none", contain: "layout style paint", width: "0", height: "0",
    overflow: "hidden",
  });
  const text = document.createElement("div");
  probe.appendChild(text);
  parent.appendChild(probe);
  return { probe, text };
}

/**
 * The automatic heights these cells ask of their rows, now, as the hook below
 * would arrive at them: for a preview, which cannot wait for its chunks.
 */
export function measureAutoRowsNow(parent: HTMLElement, items: readonly AutoRowMeasurement[], defaultHeight: number): Record<string, number> {
  const heights: Record<string, number> = {};
  if (items.length === 0) return heights;
  const { probe, text } = createAutoRowProbe(parent);
  try {
    measureNativeTextHeightBatch(items.flatMap((item) => !item.nativeParagraph || !item.dpi || item.baseHeightOnly ? [] : [
      { parts: autoRowParts(item), style: item.style, width: item.singleLine ? 1e6 : item.width - 5 * 96 / item.dpi, dpi: item.dpi },
    ]));
    for (const item of items) {
      const height = measureAutoRowItem(item, defaultHeight, text);
      if (height !== null && height > defaultHeight) heights[item.row] = Math.max(heights[item.row] ?? defaultHeight, height);
    }
  } finally {
    probe.remove();
  }
  return heights;
}

export function useSpreadsheetAutoRowHeights(
  root: RefObject<HTMLElement | null>,
  measurements: readonly AutoRowMeasurement[],
  defaultHeight: number,
  /** The sheet the measurements belong to. While that sheet re-measures after
   * an edit, its last heights stay in effect instead of dropping to the
   * defaults for a frame; another sheet never inherits them. */
  scope?: string,
) {
  // Evaluator callbacks can be recreated without any text/geometry change.
  // Stabilize the semantic input so measurement itself cannot trigger a loop.
  // The measurements of a large sheet serialize to tens of megabytes; the
  // hook keeps their content key and one parsed copy, never the text (no
  // closure below may name it, or the effect's would keep it alive).
  // Heights measured while the text engine's faces loaded are measured again
  // once they have (the same rows in every browser).
  const engineEpoch = useSpreadsheetTextEngineEpoch();
  const measuredRef = useRef<{ key: string; list: AutoRowMeasurement[] } | null>(null);
  const measurementKey = readMeasurements(measurements, measuredRef);
  const key = `${defaultHeight}:${measurementKey}`;
  const stableMeasurements = measuredRef.current!.list;
  const [result, setResult] = useState<{
    source: string;
    scope: string | undefined;
    heights: Record<string, number>;
    /** The rows measured so far of a sheet's first measurement (X-229). */
    partial?: true;
  } | null>(null);

  useEffect(() => {
    const parent = root.current;
    if (!parent || stableMeasurements.length === 0) return;
    let cancelled = false;
    let generation = 0;
    let timer: ReturnType<typeof setTimeout> | undefined;
    const { probe, text } = createAutoRowProbe(parent);

    const run = async () => {
      const current = ++generation;
      if (timer) clearTimeout(timer);
      // Late web-font completion reruns this disposable projection.
      // Reading fonts.ready lays the page out (a tenth of every entry's
      // cost on a large sheet); it is awaited only while a font is loading.
      if (document.fonts?.status === "loading") await document.fonts.ready;
      if (cancelled || current !== generation) return;
      let cursor = 0;
      const heights: Record<string, number> = {};
      // A sheet measured for the first time shows its rows as they are
      // measured, the first chunks (the top of the sheet) at once: a
      // 7,577-row price list took some 40 s to measure whole, and its
      // wrapped rows sat at the default height, their lines cut to a sliver,
      // all that time (r52 P01, X-229). Rows only grow on the way; the
      // finished measurement replaces what was shown. A sheet that already
      // shows a finished measurement keeps it until the new one is done.
      let published = Number.NEGATIVE_INFINITY;
      let publishEvery = 500;
      const publishPartial = () => {
        const now = performance.now();
        if (now - published < publishEvery) return;
        published = now;
        publishEvery = Math.min(4_000, publishEvery * 2);
        const snapshot = { ...heights };
        setResult((previous) => {
          if (previous !== null && previous.scope === scope && !previous.partial) return previous;
          const kept = previous !== null && previous.scope === scope ? previous.heights : {};
          const shown: Record<string, number> = { ...kept };
          for (const [row, height] of Object.entries(snapshot)) shown[row] = Math.max(shown[row] ?? 0, height);
          return { source: key, scope, heights: shown, partial: true };
        });
      };
      const chunk = () => {
        if (cancelled || current !== generation) return;
        const ask = (item: AutoRowMeasurement, height: number) => {
          if (height > defaultHeight) heights[item.row] = Math.max(heights[item.row] ?? defaultHeight, height);
        };
        const keyOf = (item: AutoRowMeasurement) => {
          const content: Partial<AutoRowMeasurement> = { ...item };
          delete content.row;
          return `${defaultHeight}:${engineEpoch}:${contentKey(JSON.stringify(content))}`;
        };
        // The chunk's unknown paragraphs are laid out together, under one
        // layout of the page, before any of them is measured.
        const CHUNK = 512;
        // The texts this chunk measures are laid out together first, each
        // once, and the chunk measures no further than they reach: a cell
        // past them was laid out alone, a forced layout of the page each.
        // Repeated texts (a coded column: one category a thousand times)
        // took places in the batch and were then answered from the cache, so
        // the chunk ran on past it into long comments laid out one by one
        
        const ahead: NativeTextLayoutRequest[] = [];
        const batched = new Set<string>();
        let aheadEnd = cursor;
        for (; aheadEnd < stableMeasurements.length && ahead.length < CHUNK; aheadEnd += 1) {
          const item = stableMeasurements[aheadEnd];
          if (!item.nativeParagraph || !item.dpi || item.baseHeightOnly) continue;
          const itemKey = keyOf(item);
          if (batched.has(itemKey) || measuredHeights.has(itemKey)) continue;
          batched.add(itemKey);
          ahead.push({ parts: autoRowParts(item), style: item.style, width: item.singleLine ? 1e6 : item.width - 5 * 96 / item.dpi, dpi: item.dpi });
        }
        measureNativeTextHeightBatch(ahead);
        // Cached answers cost nothing; the chunk's budget is real
        // measurements, so a sheet whose cells are all known finishes at once.
        let measuredNow = 0;
        for (; cursor < aheadEnd && measuredNow < CHUNK; cursor += 1) {
          const item = stableMeasurements[cursor];
          const cacheKey = keyOf(item);
          const known = measuredHeights.get(cacheKey);
          if (known !== undefined) {
            ask(item, known.height);
            continue;
          }
          measuredNow += 1;
          const remember = (height: number) => {
            if (measuredHeights.size >= MEASURED_HEIGHTS_LIMIT) measuredHeights.clear();
            // Kept while faces load too: the face a height was measured
            // without takes it away when it lands (the listener above).
            measuredHeights.set(cacheKey, { height, families: itemFamilies(item), text: itemText(item) });
            ask(item, height);
          };
          const height = measureAutoRowItem(item, defaultHeight, text);
          if (height !== null) remember(height);
        }
        if (cursor < stableMeasurements.length) {
          publishPartial();
          timer = setTimeout(chunk, 0);
        } else setResult((previous) => {
          // Equal heights keep their identity: rows key their paint on it.
          const same = previous !== null && previous.scope === scope &&
            Object.keys(previous.heights).length === Object.keys(heights).length &&
            Object.entries(heights).every(([row, height]) => previous.heights[row] === height);
          return { source: key, scope, heights: same ? previous.heights : heights };
        });
      };
      chunk();
    };
    void run();
    // Measured again only when a face its text is drawn in finished.
    const fontsLoaded = (event: Event) => {
      const reach = fontLoadReach(event);
      if (!reach) return;
      if (stableMeasurements.some((item) => reach.touches(itemFamilies(item), itemText(item)))) void run();
    };
    document.fonts?.addEventListener("loadingdone", fontsLoaded);
    return () => {
      cancelled = true;
      if (timer) clearTimeout(timer);
      document.fonts?.removeEventListener("loadingdone", fontsLoaded);
      probe.remove();
    };
  }, [root, stableMeasurements, key, defaultHeight, scope, engineEpoch]);

  if (!result) return undefined;
  if (result.source === key) return result.heights;
  return scope !== undefined && result.scope === scope && stableMeasurements.length > 0 ? result.heights : undefined;
}
