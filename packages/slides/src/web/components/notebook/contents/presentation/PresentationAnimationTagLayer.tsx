"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The animation tags on the slide canvas: beside each animated object, the
 * number of the click that plays each of its effects, as PowerPoint draws
 * them in Normal view while the Animations tab or the Animation Pane is open.
 * A tag picks its effect (and its object), as a click on PowerPoint's does.
 * Drawn at the same size on screen at every zoom, never in the show, a
 * thumbnail or a print.
 */

import { useMemo } from "react";
import type { PresentationBlock } from "@mrnote/core/web/components/console/editor/native-document/documentCommands.ts";
import { presentationAnimationTags } from "../../../../lib/presentation/animation-tags";
import { morphGeoOfBlock } from "../../../../lib/presentation/selection-geometry";
import { isBlockNotDrawn } from "@mrnote/core/web/lib/presentation-selection-pane.ts";
import type { TimelineStep } from "@mrnote/core/web/lib/presentation-anim-timeline.ts";
import styles from "./PresentationAnimationTagLayer.module.css";

/** Screen px a column of tags needs left of its object (tag, gap, margin). */
const TAG_ROOM_PX = 26;

export interface PresentationAnimationTagLabels {
  group: string;
  click: (click: number) => string;
  onEntry: string;
  trigger: string;
}

function axisAlignedTopLeft(
  block: PresentationBlock,
  bounds: { width: number; height: number },
): { left: number; top: number } {
  const geo = morphGeoOfBlock(block, bounds);
  const turn = (geo.rotation * Math.PI) / 180;
  const cos = Math.abs(Math.cos(turn));
  const sin = Math.abs(Math.sin(turn));
  const halfWidth = (geo.width * cos + geo.height * sin) / 2;
  const halfHeight = (geo.width * sin + geo.height * cos) / 2;
  return {
    left: geo.left + geo.width / 2 - halfWidth,
    top: geo.top + geo.height / 2 - halfHeight,
  };
}

export function PresentationAnimationTagLayer({
  blocks,
  sequence,
  bounds,
  screenWidth,
  selectedIndex,
  labels,
  onPick,
}: {
  blocks: readonly PresentationBlock[];
  sequence: readonly TimelineStep[];
  bounds: { width: number; height: number };
  /** The slide's width on screen: where a tag has no room left of its object
   * (the slide's edge clips it), it stands just inside the object. */
  screenWidth: number;
  selectedIndex: number | null;
  labels: PresentationAnimationTagLabels;
  onPick: (blockId: string, stepIndex: number) => void;
}) {
  const tags = useMemo(() => presentationAnimationTags(sequence), [sequence]);
  if (tags.size === 0) return null;
  return (
    <div className={styles.layer} role="group" aria-label={labels.group} data-testid="presentation-animation-tags">
      {blocks.map((block) => {
        const own = tags.get(block.id);
        if (!own || isBlockNotDrawn(block)) return null;
        const corner = axisAlignedTopLeft(block, bounds);
        const left = Math.max(0, corner.left / Math.max(1, bounds.width));
        const top = Math.max(0, corner.top / Math.max(1, bounds.height));
        const inside = left * screenWidth < TAG_ROOM_PX;
        return (
          <div
            key={block.id}
            className={styles.anchor}
            data-animation-tags-for={block.id}
            style={{ left: `${left * 100}%`, top: `${top * 100}%` }}
          >
            <div className={styles.column} data-inside={inside || undefined}>
              {own.map((tag) => {
                const selected = selectedIndex !== null && tag.steps.includes(selectedIndex);
                const name =
                  tag.click === null ? labels.trigger : tag.click === 0 ? labels.onEntry : labels.click(tag.click);
                return (
                  <button
                    key={tag.steps[0]}
                    type="button"
                    tabIndex={-1}
                    className={styles.tag}
                    data-testid="presentation-animation-tag"
                    data-click={tag.click ?? "trigger"}
                    data-stacked={tag.steps.length > 1 || undefined}
                    data-selected={selected || undefined}
                    aria-pressed={selected}
                    aria-label={name}
                    title={name}
                    onPointerDown={(event) => event.stopPropagation()}
                    // The caret, a selection being typed in, keeps its focus.
                    onMouseDown={(event) => {
                      event.preventDefault();
                      event.stopPropagation();
                    }}
                    onClick={(event) => {
                      event.stopPropagation();
                      onPick(block.id, tag.steps[0]);
                    }}
                  >
                    {tag.click === null ? (
                      <svg viewBox="0 0 10 10" aria-hidden="true">
                        <path d="M6.2 0.6 1.8 5.6h2.9L3.6 9.4 8.2 4.2H5.3z" fill="currentColor" />
                      </svg>
                    ) : (
                      tag.click
                    )}
                  </button>
                );
              })}
            </div>
          </div>
        );
      })}
    </div>
  );
}
