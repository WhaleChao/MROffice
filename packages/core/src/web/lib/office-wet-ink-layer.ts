/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { startTransition, useCallback, useLayoutEffect, useRef, useState } from "react";

import type { WetInkChange } from "./office-wet-ink";

const SVG_NS = "http://www.w3.org/2000/svg";

/** Paints one wet stroke as an element of the host's SVG; null draws nothing. */
export type WetInkPainter<Stroke> = (stroke: Stroke, ownerDocument: Document) => Element | null;

export class WetInkLayer<Stroke extends object> {
  private host: Element | null = null;
  private strokes: readonly Stroke[] = [];
  private readonly drawn = new Map<Stroke, Element>();
  /** Elements of strokes that have landed, kept until the host's commit shows them landed. */
  private leaving: Element[] = [];
  private paint: WetInkPainter<Stroke>;

  constructor(paint: WetInkPainter<Stroke>) {
    this.paint = paint;
  }

  /** The painter for strokes added from now on (a stroke already drawn keeps its drawing). */
  setPainter(paint: WetInkPainter<Stroke>): void {
    this.paint = paint;
  }

  /** How many strokes are drawn. */
  get size(): number {
    return this.drawn.size;
  }

  /**
   * The element the strokes are drawn in; null when the host's element goes
   * away (the strokes stay with the layer and are drawn into the next one).
   */
  mount(host: Element | null): void {
    if (host === this.host) return;
    this.host = host;
    if (!host) return;
    // A new element shows the wet strokes only: the landed ones are already
    // drawn by their document.
    this.dropLeaving();
    this.place();
  }

  /**
   * Show `strokes` (the batch as it stands, in drawing order). A stroke not
   * drawn yet is painted now; one no longer in the batch goes at once, or,
   * with `leaveAtCommit` (the batch just landed), with the host's next
   * {@link WetInkLayer.commit}.
   */
  sync(strokes: readonly Stroke[], options: { readonly leaveAtCommit?: boolean } = {}): void {
    this.strokes = strokes;
    if (this.drawn.size > 0) {
      const kept = new Set(strokes);
      for (const [stroke, element] of this.drawn) {
        if (kept.has(stroke)) continue;
        this.drawn.delete(stroke);
        if (options.leaveAtCommit && this.host && element.parentNode === this.host) this.leaving.push(element);
        else element.remove();
      }
    }
    this.place();
  }

  /** The host committed what it renders: the strokes that landed leave the layer. */
  commit(): void {
    this.dropLeaving();
  }

  private dropLeaving(): void {
    if (this.leaving.length === 0) return;
    for (const element of this.leaving) element.remove();
    this.leaving = [];
  }

  /** Each stroke's element in the host, in drawing order; only the new ones are painted. */
  private place(): void {
    const host = this.host;
    if (!host) return;
    let previous: Element | null = null;
    for (const stroke of this.strokes) {
      let element = this.drawn.get(stroke);
      if (!element) {
        element = this.paint(stroke, host.ownerDocument) ?? undefined;
        if (!element) continue;
        this.drawn.set(stroke, element);
      }
      if (element.parentNode !== host) {
        // After the stroke drawn before it: a batch only grows at its end,
        // so this is an append but for a stroke moved to a new host.
        host.insertBefore(element, previous ? previous.nextSibling : host.firstChild);
      }
      previous = element;
    }
  }
}

/** An SVG element of `tag` with `attrs` (null and undefined values are left out). */
export function wetInkSvgElement(
  ownerDocument: Document,
  tag: string,
  attrs: Readonly<Record<string, string | number | null | undefined>> = {},
): SVGElement {
  const element = ownerDocument.createElementNS(SVG_NS, tag);
  for (const [name, value] of Object.entries(attrs)) {
    if (value === null || value === undefined) continue;
    element.setAttribute(name, String(value));
  }
  return element;
}

export interface WetInkLayerState<Stroke extends object> {
  /** The layer the host's wet-ink element mounts ({@link WetInkLayer.mount}). */
  readonly layer: WetInkLayer<Stroke>;
  /** Whether any stroke is wet (the host's Undo is enabled for it). */
  readonly pending: boolean;
  /** The batch's change (a WetInkBatcher's `onChange`): drawn at once, no render of the host. */
  readonly show: (strokes: readonly Stroke[], change: WetInkChange) => void;
}

/**
 * A host's wet ink: the layer, and the one bit of React state it needs --
 * whether anything is wet -- set only when it flips. A lift used to render
 * the host (the deck's or the worksheet's whole window) for every stroke; it
 * now renders it at most once a batch, after the lift's own task (a
 * transition: the Undo button takes the stroke a moment later, the next
 * Pencil contact is answered first). The batch landing renders the host as
 * before, in the same render that takes the landed strokes, and that
 * render's commit lets the wet ones go.
 */
export function useWetInkLayer<Stroke extends object>(paint: WetInkPainter<Stroke>): WetInkLayerState<Stroke> {
  const [layer] = useState(() => new WetInkLayer<Stroke>(paint));
  layer.setPainter(paint);
  const [pending, setPending] = useState(false);
  const [landings, setLandings] = useState(0);
  const pendingRef = useRef(false);
  const show = useCallback(
    (strokes: readonly Stroke[], change: WetInkChange) => {
      layer.sync(strokes, { leaveAtCommit: change === "land" });
      // Always a render when a batch lands: its commit is when the wet
      // strokes go, landed or not (a deck replaced since drops them).
      if (change === "land") setLandings((count) => count + 1);
      const next = strokes.length > 0;
      if (next === pendingRef.current) return;
      pendingRef.current = next;
      if (next) startTransition(() => setPending(true));
      else setPending(false);
    },
    [layer],
  );
  useLayoutEffect(() => {
    layer.commit();
  }, [layer, landings]);
  return { layer, pending, show };
}
