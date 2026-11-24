"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import {
  createContext,
  useCallback,
  useContext,
  useLayoutEffect,
  useMemo,
  useRef,
  useState,
  useSyncExternalStore,
} from "react";


















/** The groups collapsed on the band now, by their collapse id. */
export const PdfRibbonFitContext = createContext<ReadonlySet<string>>(new Set());

export function usePdfRibbonGroupCollapsed(id: string | undefined): boolean {
  const collapsed = useContext(PdfRibbonFitContext);
  return id !== undefined && collapsed.has(id);
}

export const PDF_RIBBON_COLLAPSIBLE_ATTRIBUTE = "data-pdf-ribbon-collapsible";

const NO_GROUPS: ReadonlySet<string> = new Set();

/** Each band's collapsed set, by its track: for a group drawn into the band
 * from another React tree (a portal from outside the ribbon -- the native
 * text editor's Text Editing tab), which the context does not reach. */
const bandFits = new WeakMap<
  Element,
  { value: ReadonlySet<string>; listeners: Set<() => void> }
>();

function bandFitOf(track: Element) {
  let fit = bandFits.get(track);
  if (!fit) {
    fit = { value: NO_GROUPS, listeners: new Set() };
    bandFits.set(track, fit);
  }
  return fit;
}









export function usePdfRibbonFitOfElement(element: Element | null): ReadonlySet<string> {
  const track = element?.closest(".note-editor-ribbon-track") ?? null;
  const subscribe = useCallback(
    (listener: () => void) => {
      if (!track) return () => {};
      const fit = bandFitOf(track);
      fit.listeners.add(listener);
      return () => {
        fit.listeners.delete(listener);
      };
    },
    [track],
  );
  return useSyncExternalStore(
    subscribe,
    () => (track ? bandFitOf(track).value : NO_GROUPS),
    () => NO_GROUPS,
  );
}

/** A pixel of rounding is not an overflow. */
const FIT_EPSILON = 1;

export type PdfRibbonFitStep =
  | { readonly kind: "collapse"; readonly id: string }
  | { readonly kind: "expand"; readonly id: string }
  | { readonly kind: "stay" };

















export function nextPdfRibbonFitStep(input: {
  readonly available: number;
  readonly content: number;
  readonly collapsed: readonly string[];
  readonly candidates: readonly string[];
  readonly openWidths: Readonly<Record<string, number | undefined>>;
  readonly foldedWidths: Readonly<Record<string, number | undefined>>;
}): PdfRibbonFitStep {
  const { available, content, collapsed, candidates, openWidths, foldedWidths } = input;
  if (available <= 0) return { kind: "stay" };
  if (content > available + FIT_EPSILON) {
    // The furthest group along that is still open.
    for (let index = candidates.length - 1; index >= 0; index -= 1) {
      const id = candidates[index]!;
      if (!collapsed.includes(id)) return { kind: "collapse", id };
    }
    return { kind: "stay" };
  }
  const room = available - content;
  for (let index = collapsed.length - 1; index >= 0; index -= 1) {
    const id = collapsed[index]!;
    if (!candidates.includes(id)) continue;
    const open = openWidths[id];
    const folded = foldedWidths[id];
    if (open === undefined || folded === undefined) continue;
    if (open - folded <= room + FIT_EPSILON) return { kind: "expand", id };
  }
  return { kind: "stay" };
}

/** The width the band's track takes with its content (the track is at
 * least as wide as the band, so its own width does not say), in layout
 * pixels: `scale` is the band's on-screen size over its layout size (the
 * window may be scaled by the notebook canvas). */
function trackContentWidth(track: HTMLElement, scale: number): number {
  const left = track.getBoundingClientRect().left;
  let right = left;
  const visit = (element: Element) => {
    for (const child of Array.from(element.children)) {
      if (child.getClientRects().length === 0) {
        // A wrapper drawn as its children (display: contents).
        if (getComputedStyle(child).display === "contents") visit(child);
        continue;
      }
      right = Math.max(right, child.getBoundingClientRect().right);
    }
  };
  visit(track);
  const paddingRight = Number.parseFloat(getComputedStyle(track).paddingRight) || 0;
  return (right - left) / scale + paddingRight;
}

interface FitState {
  readonly tab: string;
  readonly collapsed: readonly string[];
}

/**
 * The band's fit controller: the marker it returns goes inside the band's
 * track; the set goes to PdfRibbonFitContext. Every tab starts open and is
 * fitted again from there.
 */
export function usePdfRibbonFit(activeTab: string): {
  markerRef: (element: HTMLElement | null) => void;
  collapsed: ReadonlySet<string>;
} {
  const [marker, setMarker] = useState<HTMLElement | null>(null);
  const [state, setState] = useState<FitState>({ tab: activeTab, collapsed: [] });
  // Another tab is fitted from open: its own widths, none collapsed.
  const current = useMemo<FitState>(
    () => (state.tab === activeTab ? state : { tab: activeTab, collapsed: [] }),
    [activeTab, state],
  );
  /** Each group's width when last seen open, on the tab being fitted. */
  const openWidthsRef = useRef<{ tab: string; widths: Record<string, number> }>({ tab: activeTab, widths: {} });
  const stateRef = useRef(current);
  stateRef.current = current;

  const measure = useCallback(() => {
    if (!marker) return;
    const scroller = marker.closest<HTMLElement>(".note-editor-ribbon-body");
    const track = marker.parentElement;
    if (!scroller || !track) return;
    const { tab, collapsed } = stateRef.current;
    if (openWidthsRef.current.tab !== tab) openWidthsRef.current = { tab, widths: {} };
    const openWidths = openWidthsRef.current.widths;
    const scale = track.offsetWidth > 0 ? track.getBoundingClientRect().width / track.offsetWidth : 1;
    const foldedWidths: Record<string, number> = {};
    const candidates: string[] = [];
    for (const group of Array.from(
      track.querySelectorAll<HTMLElement>(`[${PDF_RIBBON_COLLAPSIBLE_ATTRIBUTE}]`),
    )) {
      // A group of another tab is hidden by the tab gating: no box.
      if (group.getClientRects().length === 0) continue;
      const id = group.getAttribute(PDF_RIBBON_COLLAPSIBLE_ATTRIBUTE);
      if (!id || candidates.includes(id)) continue;
      candidates.push(id);
      const width = group.getBoundingClientRect().width / scale;
      if (collapsed.includes(id)) foldedWidths[id] = width;
      else openWidths[id] = width;
    }
    const available = scroller.clientWidth;
    const content = scroller.scrollWidth > available + FIT_EPSILON
      ? scroller.scrollWidth
      : Math.min(available, trackContentWidth(track, scale));
    const step = nextPdfRibbonFitStep({
      available,
      content,
      collapsed,
      candidates,
      openWidths,
      foldedWidths,
    });
    if (step.kind === "collapse") {
      setState({ tab, collapsed: [...collapsed, step.id] });
    } else if (step.kind === "expand") {
      setState({ tab, collapsed: collapsed.filter((id) => id !== step.id) });
    }
  }, [marker]);

  // Measured after every change of what is collapsed, before the paint: the
  // band settles in one frame, never showing the wider state first.
  useLayoutEffect(() => {
    measure();
  }, [measure, current.collapsed, activeTab]);

  useLayoutEffect(() => {
    if (!marker || typeof ResizeObserver === "undefined") return;
    const scroller = marker.closest<HTMLElement>(".note-editor-ribbon-body");
    const track = marker.parentElement;
    if (!scroller || !track) return;
    let frame = 0;
    const observer = new ResizeObserver(() => {
      cancelAnimationFrame(frame);
      frame = requestAnimationFrame(measure);
    });
    observer.observe(scroller);
    observer.observe(track);
    return () => {
      cancelAnimationFrame(frame);
      observer.disconnect();
    };
  }, [marker, measure]);

  const collapsedSet = useMemo<ReadonlySet<string>>(
    () => new Set(current.collapsed),
    [current.collapsed],
  );
  // Published on the band's track for groups drawn into it from another
  // tree (usePdfRibbonFitOfElement).
  useLayoutEffect(() => {
    const track = marker?.parentElement;
    if (!track) return;
    const fit = bandFitOf(track);
    if (fit.value === collapsedSet) return;
    fit.value = collapsedSet;
    for (const listener of Array.from(fit.listeners)) listener();
  }, [collapsedSet, marker]);

  return { markerRef: setMarker, collapsed: collapsedSet };
}
