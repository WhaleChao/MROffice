"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The deck's Designer pane as PowerPoint runs it (SPEC §1.1–1.2):
 *  - Design > Designer opens the pane on the current slide; its ideas come
 *    from the layout engine (lib/presentation-designer-engine) and the
 *    ranking service only orders them;
 *  - clicking an idea changes the slide at once and the pane stays open;
 *    clicking another replaces it, and one Undo returns the slide as it was;
 *  - the pane follows the current slide and the slide's edits;
 *  - the button is grey while several slides are selected, the notes have
 *    the focus, the slide show runs or the deck cannot be edited;
 *  - with "Automatically show me design ideas" on, the pane opens after a
 *    picture, chart, table or list goes on the slide; with "... when I
 *    create a new presentation" on, once a new deck's title slide has words.
 */

import { useCallback, useEffect, useMemo, useRef, useState, useSyncExternalStore } from "react";

import type { SlideModel } from "@mrnote/core/web/lib/presentation/deck-model.ts";
import type { DeckTheme } from "@mrnote/core/web/lib/presentation-theme.ts";
import type { PresentationBlock } from "@mrnote/core/web/components/console/editor/native-document/documentCommands.ts";
import {
  DESIGNER_MAX_IDEAS,
  designerContentTransforms,
  designerGraphicSummary,
  designerIdeasForSlide,
  designerRankedIdeas,
  designerSlideSignature,
  designerSlideWithGraphic,
  designerThemeSummary,
  type DesignerContext,
  type DesignerEmptyReason,
  type DesignerIdea,
  type DesignerPictureFacts,
  type DesignerTransform,
} from "../presentation-designer-engine";
import { officeIconCatalog } from "@mrnote/core/web/lib/office-icons/catalog.ts";
import {
  presentationDesignerPreferences,
  presentationDesignerPreferencesServerSnapshot,
  subscribePresentationDesignerPreferences,
} from "../presentation-designer-preferences";
import { loadOfficeImageNaturalSize } from "@mrnote/core/web/lib/office-picture-fit.ts";
import { slideBounds } from "@mrnote/core/web/lib/presentation-object-buttons.ts";
import type { LayoutFrameSource } from "@mrnote/core/web/lib/presentation-layouts.ts";
import { DESIGNER_IDEA_COMMIT } from "./window-constants";

/**
 * An idea as the pane holds it. A title-slide design idea also carries the
 * theme it puts on the whole deck (one look across every slide) and the
 * slide as that theme draws it, for its card.
 */
export type PresentationDesignerIdea = DesignerIdea<SlideModel> & {
  theme?: DeckTheme;
  previewSlide?: SlideModel;
};

/** A title-slide design: a theme for the whole deck (SPEC §1.3 item 9). */
export interface PresentationDesignerThemeChoice {
  id: string;
  name: string;
  theme: DeckTheme;
  previewSlide: SlideModel;
}

/** Why the button is grey (PowerPoint's own reasons), or null. */
export type PresentationDesignerBlock =
  | "read-only"
  | "slide-show"
  | "multiple-slides"
  | "notes-focus"
  | null;

export interface PresentationDesignerRankRequest {
  slideId: string;
  bounds: { width: number; height: number };
  ideas: readonly PresentationDesignerIdea[];
  signal: AbortSignal;
}

export interface UsePresentationDesignerOptions {
  slidesRef: { readonly current: SlideModel[] };
  activeIndexRef: { readonly current: number };
  activeSlide: SlideModel;
  slideCount: number;
  commitSlides: (next: SlideModel[], coalesceKey?: string) => void;
  /** How many steps Undo holds: a local command adds one. */
  undoDepth: () => number;
  




  automaticStep?: (depth: number) => boolean;
  accents: readonly string[];
  layoutSource: LayoutFrameSource;
  /** Header & Footer fields are drawn on the current slide. */
  footerBand: boolean;
  editable: boolean;
  presentMode: boolean;
  notesFocused: boolean;
  multipleSlidesSelected: boolean;
  /** A picture source the browser can load (`mrai-picture:` tokens resolved), or null. */
  resolvePictureSrc?: (src: string) => string | null;
  /**
   * Designer is opening by itself (Automatically show me design ideas), not
   * at the user's command: the host keeps the pane the user is in in front.
   */
  onAutoShow?: () => void;
  /** Ask the ranking service for an order; null when it has none. */
  rankIdeas?: (request: PresentationDesignerRankRequest) => Promise<readonly string[] | null>;
  /** Build a graphics idea's diagram (the Convert to SmartArt engine call); null when it cannot. */
  buildGraphic?: (transform: DesignerTransform, slide: SlideModel, signal: AbortSignal) => Promise<PresentationBlock[] | null>;
  /** Extra analysis of a decoded picture (subject focus, accent colour). */
  analysePicture?: (src: string, natural: { width: number; height: number }) => Promise<Pick<DesignerPictureFacts, "focus" | "accent"> | null>;
  /** The deck's theme now. */
  getTheme?: () => DeckTheme | undefined;
  /** Title-slide designs for this slide (none on other slides). */
  themeChoices?: (slide: SlideModel) => PresentationDesignerThemeChoice[];
  /**
   * Commit an idea with a deck theme (`null`: no theme), restyling every
   * slide for it in the same history step.
   */
  commitWithTheme?: (next: SlideModel[], coalesceKey: string, theme: DeckTheme | null) => void;
}

interface Session {
  slideId: string;
  generation: number;
  baseSignature: string;
  appliedSignature: string | null;
  appliedId: string | null;
  /** The deck's theme when the ideas were made: a later idea puts it back. */
  baseTheme: DeckTheme | undefined;
  themeApplied: boolean;
}

const RANK_TIMEOUT_MS = 6_000;

function pictureBlocks(slide: SlideModel): Array<Record<string, unknown>> {
  return slide.blocks.filter((block) => block.type === "image") as unknown as Array<Record<string, unknown>>;
}

/** What the auto-show watches on a slide: pictures, charts, tables and lists. */
function insertedContentCounts(slide: SlideModel): [number, number, number, number] {
  let pictures = 0;
  let charts = 0;
  let tables = 0;
  let lists = 0;
  for (const raw of slide.blocks) {
    const block = raw as unknown as Record<string, unknown>;
    if (block.type === "image") pictures += 1;
    else if (block.type === "chart") charts += 1;
    else if (block.type === "table") tables += 1;
    else if (block.type === "text" && Array.isArray(block.paragraphs)) {
      const bulleted = (block.paragraphs as Array<Record<string, unknown>>).filter(
        (paragraph) => paragraph.bullet === "bullet" || paragraph.bullet === "number",
      ).length;
      if (bulleted >= 3) lists += 1;
    }
  }
  return [pictures, charts, tables, lists];
}

function titleText(slide: SlideModel): string {
  for (const raw of slide.blocks) {
    const block = raw as unknown as Record<string, unknown>;
    if (block.placeholder === "title" || block.role === "title") {
      return typeof block.text === "string" ? block.text.trim() : "";
    }
  }
  return "";
}

export function usePresentationDesigner(options: UsePresentationDesignerOptions) {
  const optionsRef = useRef(options);
  optionsRef.current = options;
  const preferences = useSyncExternalStore(
    subscribePresentationDesignerPreferences,
    presentationDesignerPreferences,
    presentationDesignerPreferencesServerSnapshot,
  );
  const [open, setOpen] = useState(false);
  const [busy, setBusy] = useState(false);
  const [ideas, setIdeas] = useState<PresentationDesignerIdea[]>([]);
  const [reason, setReason] = useState<DesignerEmptyReason | null>(null);
  const [appliedId, setAppliedId] = useState<string | null>(null);
  const sessionRef = useRef<Session | null>(null);
  const generationRef = useRef(0);
  const abortRef = useRef<AbortController | null>(null);
  const openRef = useRef(open);
  openRef.current = open;

  const blocked: PresentationDesignerBlock = true
    ? "read-only"
    : options.presentMode
      ? "slide-show"
      : options.multipleSlidesSelected
        ? "multiple-slides"
        : options.notesFocused
          ? "notes-focus"
          : null;

  /** Make the ideas for the current slide (and rank them). */
  const generate = useCallback(async () => {
    const live = optionsRef.current;
    const slide = live.slidesRef.current[live.activeIndexRef.current];
    abortRef.current?.abort();
    const generation = generationRef.current + 1;
    generationRef.current = generation;
    if (!slide) {
      setIdeas([]);
      setReason("empty");
      return false;
    }
    const controller = new AbortController();
    abortRef.current = controller;
    setBusy(true);
    setAppliedId(null);
    const baseSignature = designerSlideSignature(slide);
    sessionRef.current = {
      slideId: slide.id,
      generation,
      baseSignature,
      appliedSignature: null,
      appliedId: null,
      baseTheme: live.getTheme?.(),
      themeApplied: false,
    };
    try {
      // Pictures: their pixel size (Designer's 200 x 200 rule, the crop) and,
      // where the picture can be read, its subject and colour.
      const pictures = new Map<string, DesignerPictureFacts>();
      await Promise.all(
        pictureBlocks(slide).map(async (block) => {
          const id = String(block.id);
          const rawSrc = typeof block.src === "string" ? block.src : "";
          const src = rawSrc ? (live.resolvePictureSrc ? live.resolvePictureSrc(rawSrc) : rawSrc) : null;
          let width = Number(block.pixelWidth);
          let height = Number(block.pixelHeight);
          if (!(width > 0 && height > 0) && src) {
            const natural = await loadOfficeImageNaturalSize(src, 1_500);
            width = natural?.width ?? 0;
            height = natural?.height ?? 0;
          }
          if (!(width > 0 && height > 0)) return;
          const extra = src && live.analysePicture ? await live.analysePicture(src, { width, height }).catch(() => null) : null;
          pictures.set(id, { width, height, ...(extra ?? {}) });
        }),
      );
      if (generationRef.current !== generation) return false;
      const context: DesignerContext = {
        bounds: slideBounds(slide),
        accents: live.accents,
        layoutSource: live.layoutSource,
        footerBand: live.footerBand,
        pictures,
      };
      const result = designerIdeasForSlide(slide, context, { minimum: 0 });
      let ordered = result.ideas;
      // Lists, processes and dates as SmartArt; key terms with icons.
      if (live.buildGraphic && (result.reason === null || result.reason === "few-ideas")) {
        const transforms = designerContentTransforms(slide, context, officeIconCatalog()).slice(0, 5);
        const built = await Promise.all(
          transforms.map(async (transform) => {
            const nodes = await live.buildGraphic!(transform, slide, controller.signal).catch(() => null);
            if (!nodes?.length) return null;
            const ideaSlide = designerSlideWithGraphic(slide, transform.targetId, nodes);
            const idea: PresentationDesignerIdea = {
              id: transform.id,
              kind: transform.kind,
              layoutName: transform.layoutId,
              layoutNameIsKey: true,
              slide: ideaSlide,
              score: transform.score,
              summary: designerGraphicSummary(transform, ideaSlide),
            };
            return idea;
          }),
        );
        if (generationRef.current !== generation) return false;
        ordered = [...built.filter((idea): idea is PresentationDesignerIdea => idea !== null), ...ordered]
          .sort((a, b) => b.score - a.score)
          .slice(0, DESIGNER_MAX_IDEAS);
      }
      // A title slide: designs that give the whole deck one look.
      if (live.themeChoices && (result.reason === null || result.reason === "few-ideas")) {
        const designs = live.themeChoices(slide).map((choice, index): PresentationDesignerIdea => ({
          id: choice.id,
          kind: "theme",
          layoutName: choice.name,
          layoutNameIsKey: false,
          slide,
          previewSlide: choice.previewSlide,
          theme: choice.theme,
          score: 1.25 - index * 0.01,
          summary: designerThemeSummary(choice.id, choice.previewSlide),
        }));
        ordered = [...designs, ...ordered].sort((a, b) => b.score - a.score).slice(0, DESIGNER_MAX_IDEAS);
      }
      
      // empty slide shows none.
      const reason = ordered.length > 0 ? null : result.reason ?? "few-ideas";
      if (reason) ordered = [];
      if (ordered.length > 0 && live.rankIdeas) {
        const timeout = new Promise<null>((resolve) => setTimeout(() => resolve(null), RANK_TIMEOUT_MS));
        const ranking = await Promise.race([
          live.rankIdeas({ slideId: slide.id, bounds: context.bounds, ideas: ordered, signal: controller.signal }).catch(() => null),
          timeout,
        ]);
        if (generationRef.current !== generation) return false;
        ordered = designerRankedIdeas(ordered, ranking);
      }
      setIdeas(ordered);
      setReason(reason);
      return ordered.length > 0;
    } finally {
      if (generationRef.current === generation) {
        setBusy(false);
        abortRef.current = null;
      }
    }
  }, []);

  const show = useCallback(() => {
    setOpen(true);
    return generate();
  }, [generate]);

  const close = useCallback(() => {
    generationRef.current += 1;
    abortRef.current?.abort();
    abortRef.current = null;
    setBusy(false);
    setOpen(false);
  }, []);

  const toggle = useCallback(() => {
    if (openRef.current) {
      close();
      return true;
    }
    if (blocked) return false;
    void show();
    return true;
  }, [blocked, close, show]);

  /** Apply one idea: one Undo step however many ideas are tried in a row. */
  const apply = useCallback((idea: PresentationDesignerIdea) => {
    const live = optionsRef.current;
    const session = sessionRef.current;
    const slides = live.slidesRef.current;
    const index = live.activeIndexRef.current;
    const slide = slides[index];
    if (!session || !slide || slide.id !== session.slideId || !live.editable) return false;
    const signature = designerSlideSignature(slide);
    if (signature !== session.baseSignature && signature !== session.appliedSignature) {
      void generate();
      return false;
    }
    if (session.appliedId === idea.id && signature === session.appliedSignature) return false;
    const next: SlideModel = { ...slide, blocks: idea.slide.blocks };
    for (const key of ["layout", "layoutId", "masterId"] as const) {
      if (idea.slide[key] === undefined) delete (next as Record<string, unknown>)[key];
      else (next as Record<string, unknown>)[key] = idea.slide[key];
    }
    const nextSlides = slides.map((candidate, at) => (at === index ? next : candidate));
    const key = `${DESIGNER_IDEA_COMMIT}${session.slideId}:${session.generation}`;
    if (live.commitWithTheme && (idea.theme || session.themeApplied)) {
      // A title-slide design gives the whole deck its theme; the next idea
      // tried in its place puts the deck's own theme back.
      live.commitWithTheme(nextSlides, key, idea.theme ?? session.baseTheme ?? null);
      session.themeApplied = Boolean(idea.theme);
    } else {
      live.commitSlides(nextSlides, key);
    }
    const applied = live.slidesRef.current[index];
    session.appliedSignature = applied ? designerSlideSignature(applied) : null;
    session.appliedId = idea.id;
    setAppliedId(idea.id);
    return true;
  }, [generate]);

  // The pane follows the slide: another slide, or an edit that is not the
  // idea just applied, makes new ideas.
  const activeSignature = useMemo(
    () => (open ? designerSlideSignature(options.activeSlide) : ""),
    [open, options.activeSlide],
  );
  useEffect(() => {
    if (!open) return;
    const session = sessionRef.current;
    if (
      session &&
      session.slideId === options.activeSlide.id &&
      (activeSignature === session.baseSignature || activeSignature === session.appliedSignature)
    ) {
      return;
    }
    const timer = setTimeout(() => void generate(), 450);
    return () => clearTimeout(timer);
  }, [activeSignature, generate, open, options.activeSlide.id]);

  // Auto-show after a picture, chart, table or list goes on the slide.
  const watchRef = useRef<{ slideId: string; counts: [number, number, number, number]; undo: number } | null>(null);
  useEffect(() => {
    const live = optionsRef.current;
    const counts = insertedContentCounts(options.activeSlide);
    const undo = live.undoDepth();
    const previous = watchRef.current;
    watchRef.current = { slideId: options.activeSlide.id, counts, undo };
    if (!previous || previous.slideId !== options.activeSlide.id) return;
    if (!preferences.autoShow || openRef.current || blocked) return;
    // Only the user's own command (Undo grew), not a load or a collaborator.
    if (undo <= previous.undo || live.automaticStep?.(undo)) return;
    if (counts.some((count, at) => count > previous.counts[at])) {
      live.onAutoShow?.();
      void show();
    }
  }, [blocked, options.activeSlide, preferences.autoShow, show]);

  // A new deck: once its title slide has a title, after a pause in typing.
  const newDeckRef = useRef<"unknown" | "watching" | "done">("unknown");
  useEffect(() => {
    const slide = options.activeSlide;
    if (newDeckRef.current === "unknown") {
      const nativeSource = Boolean((slide.source as Record<string, unknown> | undefined)?.slideId);
      newDeckRef.current = options.slideCount === 1 && !nativeSource && titleText(slide) === "" &&
        slide.blocks.some((block) => (block as Record<string, unknown>).placeholder === "title")
        ? "watching"
        : "done";
      return;
    }
    if (newDeckRef.current !== "watching") return;
    // The deck the window first drew can be the empty slide it shows before
    // the page arrives: an imported deck that then loads is no new deck
    
    if (
      options.slideCount !== 1 ||
      Boolean((slide.source as Record<string, unknown> | undefined)?.slideId)
    ) {
      newDeckRef.current = "done";
      return;
    }
    if (!preferences.autoShowNewDeck || openRef.current || blocked || titleText(slide) === "") return;
    const timer = setTimeout(() => {
      newDeckRef.current = "done";
      if (openRef.current) return;
      optionsRef.current.onAutoShow?.();
      void show();
    }, 1_500);
    return () => clearTimeout(timer);
  }, [blocked, options.activeSlide, options.slideCount, preferences.autoShowNewDeck, show]);

  useEffect(() => () => abortRef.current?.abort(), []);

  // The slide show or a read-only switch closes nothing, but a grey button
  // with the pane open still lets it be closed.
  return {
    open,
    busy,
    ideas,
    reason,
    appliedId,
    blocked,
    preferences,
    toggle,
    show,
    close,
    apply,
    refresh: generate,
  };
}
