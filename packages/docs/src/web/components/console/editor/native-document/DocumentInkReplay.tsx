"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Draw ▸ Ink Replay in Docs (Word's Ink Replay): the document's strokes
 * are drawn again in the order they were drawn, at a steady pace, with a
 * bar to play, pause and step; the strokes themselves are hidden while the
 * replay runs. The timeline, the bar and the stepping are the deck's
 * (lib/office-ink-replay, PresentationInkReplay); the strokes come from
 * the page column, drawn over the pages like the Draw tab's own surface.
 */

import { useEffect, useLayoutEffect, useMemo, useState } from "react";
import { createPortal } from "react-dom";
import type { EditorView } from "@tiptap/pm/view";

import {
  PresentationInkReplayBar,
  type PresentationInkReplayState,
} from "@mrnote/slides/web/components/notebook/contents/presentation/PresentationInkReplay.tsx";
import { inkReplayFrame, inkReplayTimeline } from "@mrnote/ink/web/lib/office-ink-replay.ts";
import { inkStrokePathData } from "@mrnote/core/web/lib/office-ink-smooth.ts";
import { InkEffectStroke } from "@mrnote/core/web/components/office-ribbon/InkEffectStroke.tsx";
import { normalizeDocumentInkEffect, type DocumentInkEffect } from "@mrnote/core/web/lib/document-shapes.ts";
import { documentInkFlightStrokes, documentInkStrokes, type DocumentInkClientPoint } from "@mrnote/core/web/components/console/editor/native-document/documentInk.ts";
import { surfaceLayout, type SurfaceLayout } from "./DocumentInkSurface";
import { documentViewScale } from "@mrnote/core/web/components/console/editor/native-document/documentObjectResize.ts";
import { setInkLightStrokeHidden } from "@mrnote/core/web/components/console/editor/native-document/paged/inkLightLayer.ts";

export type DocumentInkReplayState = PresentationInkReplayState;

interface ReplayStroke {
  readonly id: string;
  readonly pos: number;
  /** In the overlay's own pixels. */
  readonly points: DocumentInkClientPoint[];
  readonly color: string;
  readonly width: number;
  /** A pen's effect, drawn in its layers as the page draws it. */
  readonly effect?: DocumentInkEffect;
}

/** The strokes as they stand on the page column, in the overlay's pixels. */
function readStrokes(view: EditorView, layout: SurfaceLayout): ReplayStroke[] {
  const positions = documentInkStrokes(view.state.doc).map(({ pos }) => pos);
  const flight = documentInkFlightStrokes(view, positions);
  if (!flight) return [];
  const { host, bounds } = layout;
  const hostRect = host?.getBoundingClientRect();
  const { scaleX, scaleY } = host ? documentViewScale(host) : { scaleX: 1, scaleY: 1 };
  const sx = scaleX || 1;
  const sy = scaleY || 1;
  const toOverlay = (point: DocumentInkClientPoint): DocumentInkClientPoint =>
    host && hostRect
      ? {
          x: (point.x - hostRect.left) / sx - host.clientLeft - bounds.left,
          y: (point.y - hostRect.top) / sy - host.clientTop - bounds.top,
        }
      : { x: point.x - bounds.left, y: point.y - bounds.top };
  return flight.map((stroke, index) => ({
    id: String(positions[index]),
    pos: positions[index]!,
    points: stroke.points.map(toOverlay),
    color: stroke.color,
    width: stroke.width / sx,
    ...inkEffectOf(view, positions[index]!),
  }));
}

function inkEffectOf(view: EditorView, pos: number): { effect?: DocumentInkEffect } {
  const attrs = view.state.doc.nodeAt(pos)?.attrs as Record<string, unknown> | undefined;
  const effect = attrs?.inkTool === "pen" ? normalizeDocumentInkEffect(attrs.inkEffect) : undefined;
  return effect ? { effect } : {};
}

export default function DocumentInkReplay({
  view,
  state,
  setState,
  barHost,
  labels,
}: {
  view: EditorView;
  state: DocumentInkReplayState;
  setState: (next: DocumentInkReplayState | null) => void;
  /** Where the replay bar hangs (the window's body). */
  barHost: HTMLElement | null;
  labels: Parameters<typeof PresentationInkReplayBar>[0]["labels"];
}) {
  // The strokes and the overlay's place are read once, when the replay
  // starts: the document is not edited while it plays.
  const [layout] = useState<SurfaceLayout>(() => surfaceLayout(view));
  const [strokes] = useState<ReplayStroke[]>(() => readStrokes(view, layout));
  const timeline = useMemo(
    () => inkReplayTimeline(strokes.map((stroke) => ({ id: stroke.id, points: stroke.points }))),
    [strokes],
  );
  // The document's own strokes are hidden while the replay draws them.
  useLayoutEffect(() => {
    const hidden: HTMLElement[] = [];
    const hiddenInLayer: number[] = [];
    for (const stroke of strokes) {
      // A stroke its sheet's layer draws is hidden there (paged/inkLightLayer).
      if (setInkLightStrokeHidden(view, stroke.pos, true)) {
        hiddenInLayer.push(stroke.pos);
        continue;
      }
      const dom = view.nodeDOM(stroke.pos);
      if (dom instanceof HTMLElement) {
        dom.style.visibility = "hidden";
        hidden.push(dom);
      }
    }
    return () => {
      for (const dom of hidden) dom.style.visibility = "";
      for (const pos of hiddenInLayer) setInkLightStrokeHidden(view, pos, false);
    };
  }, [strokes, view]);
  // The clock: the position advances with the frames while playing.
  const playing = state.playing;
  const duration = timeline.duration;
  useEffect(() => {
    if (!playing) return;
    let frame = 0;
    let last = performance.now();
    let position = state.position;
    const tick = (now: number) => {
      position = Math.min(duration, position + (now - last));
      last = now;
      setState({ position, playing: position < duration });
      if (position < duration) frame = window.requestAnimationFrame(tick);
    };
    frame = window.requestAnimationFrame(tick);
    return () => window.cancelAnimationFrame(frame);
    // The position is carried in the closure while playing; a seek pauses.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [duration, playing, setState]);
  const frame = useMemo(() => inkReplayFrame(timeline, state.position), [state.position, timeline]);
  if (typeof document === "undefined") return null;
  const { host, bounds } = layout;
  return (
    <>
      {createPortal(
        <svg
          aria-hidden="true"
          data-testid="document-ink-replay-strokes"
          viewBox={`0 0 ${Math.max(1, bounds.width)} ${Math.max(1, bounds.height)}`}
          preserveAspectRatio="none"
          style={{
            position: host ? "absolute" : "fixed",
            left: bounds.left,
            top: bounds.top,
            width: bounds.width,
            height: bounds.height,
            overflow: "visible",
            pointerEvents: "none",
            zIndex: 30,
          }}
        >
          {strokes.map((stroke) => {
            const count = frame.get(stroke.id) ?? 0;
            if (count < 2) return null;
            if (stroke.effect) {
              return (
                <InkEffectStroke
                  key={stroke.id}
                  effect={stroke.effect}
                  gradientId={`document-ink-replay-${stroke.id}`}
                  geometry={{ points: stroke.points.slice(0, count), widthPx: stroke.width, pressured: false }}
                />
              );
            }
            return (
              <path
                key={stroke.id}
                fill="none"
                d={inkStrokePathData(stroke.points.slice(0, count))}
                stroke={stroke.color}
                strokeWidth={stroke.width}
                strokeLinecap="round"
                strokeLinejoin="round"
                vectorEffect="non-scaling-stroke"
              />
            );
          })}
        </svg>,
        host ?? document.body,
      )}
      <PresentationInkReplayBar
        host={barHost}
        timeline={timeline}
        state={state}
        setState={setState}
        labels={labels}
      />
    </>
  );
}
