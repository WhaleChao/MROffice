"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The slide show's See All Slides (GAP #6): every slide of the show as a grid
 * of thumbnails over the show, to jump to one.
 *
 * Microsoft Support:
 *  - "Use keyboard shortcuts to deliver PowerPoint presentations": in Slide
 *    Show view Ctrl+S views the All Slides dialog box; "Magnify a slide in
 *    Slide Show view": zooming out does not open the see-all-slides view, G
 *    does;
 *  - "Start the presentation and see your notes in Presenter view": See all
 *    slides shows thumbnails of all the slides, to jump to a specific slide;
 *  - "Play a slide show": right-click, See All Slides, then click the slide to
 *    show the audience.
 * Esc leaves the grid back on the slide the show was on; arrows move between
 * thumbnails, Enter or a click shows that slide. What the documents leave
 * open (sections, a jumped-to slide's builds, hidden slides) is in
 * pptx-feature-gap PROGRESS "GAP #6", to be checked against PowerPoint.
 *
 * The grid is the editor's Slide Sorter, read-only, a click opening the slide
 * (`openOnClick`); each thumbnail is drawn as the show draws a slide
 * (`PresentationSlideView`, its media as badges as the presenter's next-slide
 * preview has them), under whichever text engine the show runs, and under its
 * own test id so nothing waiting on the show's slide counts it.
 */

import { useCallback, useEffect, useMemo, useRef, useState, type ReactNode } from "react";

import type { SectionGroup } from "@mrnote/core/web/lib/presentation-sections.ts";

import {
  PresentationSlideSorter,
  type PresentationSlideSorterLabels,
} from "./PresentationSlideSorter";
import { PresentationSlideView } from "./PresentationSlideView";
import { RAIL_EAGER_MINIATURE_RADIUS } from "./PresentationSlideThumbnails";

type SlideViewProps = Parameters<typeof PresentationSlideView>[0];
type ShowSlide = SlideViewProps["slide"] & { readonly id: string; readonly title: string };

export interface PresentationShowAllSlidesProps<S extends ShowSlide> {
  /** The show's slides, in show order. */
  readonly slides: readonly S[];
  readonly groups: readonly SectionGroup<S>[];
  /** The slide the show is on. */
  readonly currentIndex: number;
  readonly theme: SlideViewProps["theme"];
  readonly masters: SlideViewProps["masters"];
  readonly randomEffectSeed?: number;
  readonly isHidden: (slide: S) => boolean;
  readonly labels: PresentationSlideSorterLabels;
  /** Show slide `index` (of `slides`); the grid closes. */
  readonly onPick: (index: number) => void;
  /** The back arrow: leave the grid on the slide the show is on (Esc). */
  readonly onBack?: () => void;
  /** The grid's own words: its Sections rail, back arrow and zoom. */
  readonly chrome?: {
    readonly sections: string;
    readonly back: string;
    readonly toggleSections: string;
    readonly zoomOut: string;
    readonly zoomIn: string;
  };
}

/** The thumbnail sizes the zoom slider steps through (its middle: the default). */
const ZOOM_MIN = 0.6;
const ZOOM_MAX = 1.8;
const ZOOM_STEP = 0.2;

/** A thumbnail drawn once it is near the screen: a long deck's grid is not
 *  every slide laid out at once. The slides about the current one (the
 *  thumbnail pane's eager radius) are drawn at once, as is every one where
 *  there is no observer. */
function LazyThumbnail({ eager, children }: { eager: boolean; children: () => ReactNode }) {
  const ref = useRef<HTMLSpanElement | null>(null);
  const [shown, setShown] = useState(() => eager || typeof IntersectionObserver === "undefined");
  useEffect(() => {
    if (shown) return;
    const node = ref.current;
    if (!node) return;
    const observer = new IntersectionObserver(
      (entries) => {
        if (entries.some((entry) => entry.isIntersecting)) {
          setShown(true);
          observer.disconnect();
        }
      },
      { rootMargin: "200px" },
    );
    observer.observe(node);
    return () => observer.disconnect();
  }, [shown]);
  return (
    <span ref={ref} className="presentation-show-all__thumb">
      {shown ? children() : null}
    </span>
  );
}

export function PresentationShowAllSlides<S extends ShowSlide>({
  slides,
  groups,
  currentIndex,
  theme,
  masters,
  randomEffectSeed,
  isHidden,
  labels,
  onPick,
  onBack,
  chrome,
}: PresentationShowAllSlidesProps<S>) {
  // PowerPoint's See All Slides (truth show-see-all): a dark full-screen
  // grid; a Sections rail on the left (collapsible), the section of the
  // focused slide lit; the back arrow at the top left; a zoom slider at the
  // bottom right; the current slide framed and numbered in the accent
  // colour, the focused one framed in white; a hidden slide greyed, its
  // number struck through.
  const rootRef = useRef<HTMLDivElement | null>(null);
  const [railOpen, setRailOpen] = useState(true);
  const [zoom, setZoom] = useState(1);
  const [focusedId, setFocusedId] = useState<string | null>(null);
  const order = useMemo(() => slides.map((slide) => slide.id), [slides]);
  const sorterGroups = useMemo(
    () => groups.map((group) => ({ section: group.section, slides: group.slides })),
    [groups],
  );
  const [collapsed, setCollapsed] = useState<ReadonlySet<string>>(() => new Set());
  const toggleSection = useCallback((sectionId: string) => {
    setCollapsed((previous) => {
      const next = new Set(previous);
      if (next.has(sectionId)) next.delete(sectionId);
      else next.add(sectionId);
      return next;
    });
  }, []);
  const currentId = order[currentIndex];
  const renderThumbnail = useCallback(
    (slide: S) => (
      <LazyThumbnail
        eager={Math.abs(order.indexOf(slide.id) - Math.max(0, order.indexOf(currentId ?? ""))) <= RAIL_EAGER_MINIATURE_RADIUS}
      >
        {() => (
          <PresentationSlideView
            testId="presentation-show-all-thumb"
            randomEffectSeed={randomEffectSeed}
            fitToParent
            slide={slide}
            theme={theme}
            masters={masters}
            staticMedia
          />
        )}
      </LazyThumbnail>
    ),
    [currentId, masters, order, randomEffectSeed, theme],
  );
  const pick = useCallback(
    (slideId: string) => {
      const index = order.indexOf(slideId);
      if (index >= 0) onPick(index);
    },
    [onPick, order],
  );
  const ignore = useCallback(() => {}, []);
  const focusedSlideId = focusedId ?? currentId ?? null;
  const litSection =
    groups.find((group) => group.slides.some((slide) => slide.id === focusedSlideId))?.section?.id ?? null;
  const goToSection = (sectionId: string) => {
    const root = rootRef.current;
    const section = root?.querySelector<HTMLElement>(`[data-testid="presentation-sorter-section-${CSS.escape(sectionId)}"]`);
    section?.scrollIntoView?.({ block: "start" });
    section?.querySelector<HTMLElement>('[data-testid="presentation-sorter-cell"]')?.focus({ preventScroll: true });
  };
  const sectioned = groups.some((group) => group.section !== null);
  // Opened, the grid shows the current slide's section from its heading:
  // PowerPoint scrolls it to the top (truth show-see-all 02: on slide 3,
  // Body's heading at the top, Intro scrolled away above it). Once, as it
  // opens: the focus moving on after that scrolls as little as it needs.
  // (The sorter's own focus on the current slide comes first: a child's
  // effect runs before its parent's.)
  const openedAtSectionRef = useRef(false);
  useEffect(() => {
    if (openedAtSectionRef.current) return;
    openedAtSectionRef.current = true;
    const sectionId = groups.find((group) => group.slides.some((slide) => slide.id === currentId))?.section?.id;
    if (!sectionId) return;
    rootRef.current
      ?.querySelector<HTMLElement>(`[data-testid="presentation-sorter-section-${CSS.escape(sectionId)}"]`)
      ?.scrollIntoView?.({ block: "start" });
  }, [currentId, groups]);
  return (
    <div
      ref={rootRef}
      className="presentation-show-all"
      data-testid="presentation-show-all-slides"
      data-rail={sectioned && railOpen ? "open" : "closed"}
      style={{ ["--show-all-zoom" as string]: String(zoom) }}
      onFocus={(event) => {
        const cell = (event.target as HTMLElement).closest<HTMLElement>("[data-sorter-slide-id]");
        if (cell) setFocusedId(cell.getAttribute("data-sorter-slide-id"));
      }}
      // The show's click-to-advance and its own right-click menu stay
      // behind the grid.
      onClick={(event) => event.stopPropagation()}
      onContextMenu={(event) => {
        event.preventDefault();
        event.stopPropagation();
      }}
    >
      {chrome && onBack ? (
        <button
          type="button"
          className="presentation-show-all__back"
          aria-label={chrome.back}
          title={chrome.back}
          data-testid="presentation-show-all-back"
          onClick={onBack}
        >
          <svg viewBox="0 0 24 24" width="24" height="24" aria-hidden="true">
            <circle cx="12" cy="12" r="11" fill="none" stroke="currentColor" strokeWidth="1.5" />
            <path d="M17 12H7m4-4-4 4 4 4" fill="none" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round" />
          </svg>
        </button>
      ) : null}
      {chrome && sectioned ? (
        <nav className="presentation-show-all__rail" aria-label={chrome.sections} data-testid="presentation-show-all-rail">
          <div className="presentation-show-all__rail-head">
            {railOpen ? <span className="presentation-show-all__rail-title">{chrome.sections}</span> : null}
            <button
              type="button"
              className="presentation-show-all__rail-toggle"
              aria-label={chrome.toggleSections}
              aria-expanded={railOpen}
              title={chrome.toggleSections}
              data-testid="presentation-show-all-rail-toggle"
              onClick={() => setRailOpen((open) => !open)}
            >
              {railOpen ? "\u2039" : "\u203a"}
            </button>
          </div>
          {railOpen
            ? groups.map((group) =>
                group.section ? (
                  <button
                    key={group.section.id}
                    type="button"
                    className="presentation-show-all__rail-item"
                    aria-current={group.section.id === litSection ? "true" : undefined}
                    data-testid={`presentation-show-all-rail-${group.section.id}`}
                    onClick={() => goToSection(group.section!.id)}
                  >
                    {group.section.name}
                  </button>
                ) : null,
              )
            : null}
        </nav>
      ) : null}
      <PresentationSlideSorter<S>
        groups={sorterGroups}
        order={order}
        activeId={order[currentIndex] ?? null}
        collapsedSections={collapsed}
        onToggleSection={toggleSection}
        isHidden={isHidden}
        renderThumbnail={renderThumbnail}
        labels={labels}
        readOnly
        openOnClick
        onActivate={ignore}
        onOpen={pick}
        onReorder={ignore}
        onDelete={ignore}
        onContextMenu={ignore}
      />
      {chrome ? (
        <div className="presentation-show-all__zoom" data-testid="presentation-show-all-zoom">
          <button type="button" aria-label={chrome.zoomOut} title={chrome.zoomOut} onClick={() => setZoom((value) => Math.max(ZOOM_MIN, Math.round((value - ZOOM_STEP) * 10) / 10))}>
            {"\u2212"}
          </button>
          <input
            type="range"
            min={ZOOM_MIN}
            max={ZOOM_MAX}
            step={ZOOM_STEP}
            value={zoom}
            aria-label={chrome.zoomIn}
            onChange={(event) => setZoom(Number(event.target.value))}
            data-testid="presentation-show-all-zoom-range"
          />
          <button type="button" aria-label={chrome.zoomIn} title={chrome.zoomIn} onClick={() => setZoom((value) => Math.min(ZOOM_MAX, Math.round((value + ZOOM_STEP) * 10) / 10))}>
            +
          </button>
        </div>
      ) : null}
    </div>
  );
}
