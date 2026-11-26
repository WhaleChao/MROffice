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

/**
 * Office's middle scale: a group that names a compact form (`compact` on
 * PdfFoldingRibbonGroup) is drawn in it before it folds -- its frequent
 * commands on the band, the rest in its pane (Format Shape). Marked with
 * this attribute; the set of groups drawn compact goes through
 * PdfRibbonCompactContext.
 */
export const PDF_RIBBON_COMPACTABLE_ATTRIBUTE = "data-pdf-ribbon-compactable";

/** The groups drawn compact on the band now, by their collapse id. */
export const PdfRibbonCompactContext = createContext<ReadonlySet<string>>(new Set());

export function usePdfRibbonGroupCompact(id: string | undefined): boolean {
  const compact = useContext(PdfRibbonCompactContext);
  return id !== undefined && compact.has(id);
}

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
  | { readonly kind: "compact"; readonly id: string }
  | { readonly kind: "uncompact"; readonly id: string }
  | { readonly kind: "stay" };
























export function nextPdfRibbonFitStep(input: {
  readonly available: number;
  readonly content: number;
  readonly collapsed: readonly string[];
  readonly compacted?: readonly string[];
  readonly candidates: readonly string[];
  readonly compactable?: readonly string[];
  readonly openWidths: Readonly<Record<string, number | undefined>>;
  readonly compactWidths?: Readonly<Record<string, number | undefined>>;
  readonly foldedWidths: Readonly<Record<string, number | undefined>>;
  /** Groups not to open again (nor back to full) now: see usePdfRibbonFit's
   * hysteresis. */
  readonly held?: ReadonlySet<string>;
}): PdfRibbonFitStep {
  const { available, content, collapsed, candidates, openWidths, foldedWidths } = input;
  const held = input.held;
  const compacted = input.compacted ?? [];
  const compactable = input.compactable ?? [];
  const compactWidths = input.compactWidths ?? {};
  if (available <= 0) return { kind: "stay" };
  if (content > available + FIT_EPSILON) {
    // The furthest group along with a compact form not yet drawn in it.
    for (let index = candidates.length - 1; index >= 0; index -= 1) {
      const id = candidates[index]!;
      if (compactable.includes(id) && !compacted.includes(id) && !collapsed.includes(id)) {
        return { kind: "compact", id };
      }
    }
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
    if (!candidates.includes(id) || held?.has(id)) continue;
    const open = compacted.includes(id) ? compactWidths[id] : openWidths[id];
    const folded = foldedWidths[id];
    if (open === undefined || folded === undefined) continue;
    if (open - folded <= room + FIT_EPSILON) return { kind: "expand", id };
  }
  for (let index = compacted.length - 1; index >= 0; index -= 1) {
    const id = compacted[index]!;
    if (!candidates.includes(id) || collapsed.includes(id) || held?.has(id)) continue;
    const full = openWidths[id];
    const compact = compactWidths[id];
    if (full === undefined || compact === undefined) continue;
    if (full - compact <= room + FIT_EPSILON) return { kind: "uncompact", id };
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
  readonly compacted: readonly string[];
}

/**
 * The band's fit controller: the marker it attaches goes inside the band's
 * track; the set goes to PdfRibbonFitContext. Every tab starts open and is
 * fitted again from there.
 */
export function usePdfRibbonFit(activeTab: string): {
  attachMarker: (element: HTMLElement | null) => void;
  collapsed: ReadonlySet<string>;
  compacted: ReadonlySet<string>;
} {
  const [marker, setMarker] = useState<HTMLElement | null>(null);
  const [state, setState] = useState<FitState>({ tab: activeTab, collapsed: [], compacted: [] });
  // Another tab is fitted from open: its own widths, none collapsed.
  const current = useMemo<FitState>(
    () => (state.tab === activeTab ? state : { tab: activeTab, collapsed: [], compacted: [] }),
    [activeTab, state],
  );
  /** Each group's width when last seen in full and compact, on the tab
   * being fitted. */
  const widthsRef = useRef<{ tab: string; open: Record<string, number>; compact: Record<string, number> }>(
    { tab: activeTab, open: {}, compact: {} },
  );
  const stateRef = useRef(current);
  stateRef.current = current;
  /**
   * When and at what band width each group folded or went compact, on the
   * tab being fitted: a group goes back only when the band is wider than it
   * was then, or another group folded or went compact after it (and still
   * is) -- the room that one left. Without it a group whose width was misread
   * opened, overflowed and folded again for ever.
   */
  const stampsRef = useRef<{
    tab: string;
    seq: number;
    stamps: Record<string, { available: number; seq: number }>;
  }>({ tab: activeTab, seq: 0, stamps: {} });
  /** A measure waiting for the band's groups to be drawn as the fit says. */
  const pendingRef = useRef<{ frame: number; tries: number }>({ frame: 0, tries: 0 });

  const measure = useCallback(() => {
    if (!marker) return;
    const scroller = marker.closest<HTMLElement>(".note-editor-ribbon-body");
    const track = marker.parentElement;
    if (!scroller || !track) return;
    const { tab, collapsed, compacted } = stateRef.current;
    if (widthsRef.current.tab !== tab) widthsRef.current = { tab, open: {}, compact: {} };
    if (stampsRef.current.tab !== tab) stampsRef.current = { tab, seq: 0, stamps: {} };
    const { open: openWidths, compact: compactWidths } = widthsRef.current;
    const scale = track.offsetWidth > 0 ? track.getBoundingClientRect().width / track.offsetWidth : 1;
    const foldedWidths: Record<string, number> = {};
    const candidates: string[] = [];
    const compactable: string[] = [];
    let drawnAsFitted = true;
    for (const group of Array.from(
      track.querySelectorAll<HTMLElement>(`[${PDF_RIBBON_COLLAPSIBLE_ATTRIBUTE}]`),
    )) {
      // A group of another tab is hidden by the tab gating: no box.
      if (group.getClientRects().length === 0) continue;
      const id = group.getAttribute(PDF_RIBBON_COLLAPSIBLE_ATTRIBUTE);
      if (!id || candidates.includes(id)) continue;
      candidates.push(id);
      const canCompact = group.hasAttribute(PDF_RIBBON_COMPACTABLE_ATTRIBUTE);
      if (canCompact) compactable.push(id);
      // Each width is read as the group is drawn, never as the fit wants it:
      // a group drawn into the band from another tree (the native text
      // editor's Text Editing tab) takes a new fit a render later, and read
      // as wanted, a folded button's width was taken for the open group's
      // -- the fit opened and folded it for ever, and React's update-depth
      
      // to edit at any window under ~1300 px).
      const drawnFolded = group.hasAttribute("data-pdf-ribbon-collapsed");
      const drawnCompact = !drawnFolded && group.hasAttribute("data-pdf-ribbon-compact");
      const wantFolded = collapsed.includes(id);
      const wantCompact = !wantFolded && canCompact && compacted.includes(id);
      if (drawnFolded !== wantFolded || drawnCompact !== wantCompact) drawnAsFitted = false;
      const width = group.getBoundingClientRect().width / scale;
      if (drawnFolded) foldedWidths[id] = width;
      else if (drawnCompact) compactWidths[id] = width;
      else openWidths[id] = width;
    }
    const pending = pendingRef.current;
    if (!drawnAsFitted) {
      // Measured again once it is drawn (a frame or two later).
      cancelAnimationFrame(pending.frame);
      if (pending.tries < 20) {
        pending.tries += 1;
        pending.frame = requestAnimationFrame(() => measureRef.current());
      }
      return;
    }
    pending.tries = 0;
    const available = scroller.clientWidth;
    const content = scroller.scrollWidth > available + FIT_EPSILON
      ? scroller.scrollWidth
      : Math.min(available, trackContentWidth(track, scale));
    const stamps = stampsRef.current;
    const held = new Set<string>();
    for (const id of [...collapsed, ...compacted]) {
      const stamp = stamps.stamps[id];
      if (!stamp || available > stamp.available + FIT_EPSILON) continue;
      const later = [...collapsed, ...compacted].some(
        (other) => other !== id && (stamps.stamps[other]?.seq ?? -1) > stamp.seq,
      );
      if (!later) held.add(id);
    }
    const step = nextPdfRibbonFitStep({
      available,
      content,
      collapsed,
      compacted,
      candidates,
      compactable,
      openWidths,
      compactWidths,
      foldedWidths,
      held,
    });
    const stamp = (id: string) => {
      stamps.seq += 1;
      stamps.stamps[id] = { available, seq: stamps.seq };
    };
    if (step.kind === "collapse") {
      stamp(step.id);
      setState({ tab, collapsed: [...collapsed, step.id], compacted });
    } else if (step.kind === "expand") {
      // Back to compact, if it went compact before it folded.
      if (compacted.includes(step.id)) stamp(step.id);
      else delete stamps.stamps[step.id];
      setState({ tab, collapsed: collapsed.filter((id) => id !== step.id), compacted });
    } else if (step.kind === "compact") {
      stamp(step.id);
      setState({ tab, collapsed, compacted: [...compacted, step.id] });
    } else if (step.kind === "uncompact") {
      delete stamps.stamps[step.id];
      setState({ tab, collapsed, compacted: compacted.filter((id) => id !== step.id) });
    }
  }, [marker]);
  const measureRef = useRef(measure);
  useLayoutEffect(() => {
    measureRef.current = measure;
  }, [measure]);
  useLayoutEffect(() => {
    const pending = pendingRef.current;
    return () => cancelAnimationFrame(pending.frame);
  }, []);

  // Measured after every change of what is collapsed, before the paint: the
  // band settles in one frame, never showing the wider state first.
  useLayoutEffect(() => {
    measure();
  }, [measure, current.collapsed, current.compacted, activeTab]);

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

  const compactedSet = useMemo<ReadonlySet<string>>(
    () => new Set(current.compacted),
    [current.compacted],
  );
  return { attachMarker: setMarker, collapsed: collapsedSet, compacted: compactedSet };
}
