"use client";
/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Reorder Objects / Reorder Overlapping Objects (PPT-A045, PPT-A119).
 *
 * PowerPoint for Mac shows the slide's objects as stacked layers, front
 * first, and the user drags a layer to a new depth; OK writes the order,
 * Cancel leaves the slide as it was. Here each layer is a card drawn with the
 * object alone on a slide-shaped tile (the host renders it), front at the
 * top, tilted into a stack. A layer moves by dragging it, or by selecting it
 * and pressing Bring Forward / Send Backward (Alt+Up / Alt+Down); a locked
 * object stays where it is. The order is applied in one step through
 * `applyReorderedLayers` by the host.
 */
import { useMemo, useRef, useState, type ReactNode, type RefObject } from "react";

import { OfficeDialogShell } from "@mrnote/core/web/components/office-ribbon/OfficeDialogShell.tsx";
import { OfficeSemanticIcon } from "@mrnote/core/web/components/office-ribbon/icons/OfficeSemanticIcon.tsx";
import { moveLayer } from "../presentation-reorder-objects";

export interface PresentationReorderLayerView {
  readonly objectId: string;
  readonly label: string;
  readonly hidden: boolean;
  readonly locked: boolean;
  readonly selected: boolean;
  readonly preview: ReactNode;
}

export interface PresentationReorderObjectsLabels {
  readonly title: string;
  readonly description: string;
  readonly ok: string;
  readonly cancel: string;
  readonly bringForward: string;
  readonly sendBackward: string;
  readonly locked: string;
  readonly hidden: string;
  readonly layer: (position: number, count: number, label: string) => string;
}

/** Whether moving a layer leaves every locked layer at its own depth. */
function keepsLockedLayers(
  layers: readonly PresentationReorderLayerView[],
  order: readonly string[],
): boolean {
  return layers.every(
    (layer, index) => !layer.locked || order[index] === layer.objectId,
  );
}

export function PresentationReorderObjectsDialog({
  open,
  layers,
  labels,
  onApply,
  onCancel,
  returnFocusRef,
}: {
  readonly open: boolean;
  /** Front first, as `reorderLayers` lists them. */
  readonly layers: readonly PresentationReorderLayerView[];
  readonly labels: PresentationReorderObjectsLabels;
  /** The new order, front first. */
  readonly onApply: (order: readonly string[]) => void;
  readonly onCancel: () => void;
  readonly returnFocusRef?: RefObject<HTMLElement | null>;
}) {
  const initialOrder = useMemo(() => layers.map((layer) => layer.objectId), [layers]);
  const [order, setOrder] = useState<readonly string[]>(initialOrder);
  const [active, setActive] = useState<string | null>(
    () => layers.find((layer) => layer.selected && !layer.locked)?.objectId ?? null,
  );
  const dragging = useRef<string | null>(null);
  const listRef = useRef<HTMLOListElement | null>(null);
  const byId = useMemo(
    () => new Map(layers.map((layer) => [layer.objectId, layer])),
    [layers],
  );
  const changed = order.some((id, index) => id !== initialOrder[index]);

  const move = (objectId: string, index: number) => {
    const layer = byId.get(objectId);
    if (!layer || layer.locked) return;
    const next = moveLayer(order, objectId, index);
    if (!keepsLockedLayers(layers, next)) return;
    setOrder(next);
  };
  // The nearest place a step can put the layer: a locked layer keeps its
  // depth, so the step goes past it. Next to a locked layer the buttons
  
  const stepTarget = (objectId: string | null, delta: -1 | 1): number | null => {
    if (!objectId || byId.get(objectId)?.locked) return null;
    const from = order.indexOf(objectId);
    for (let index = from + delta; index >= 0 && index < order.length; index += delta) {
      if (keepsLockedLayers(layers, moveLayer(order, objectId, index))) return index;
    }
    return null;
  };
  const step = (objectId: string | null, delta: -1 | 1) => {
    const index = stepTarget(objectId, delta);
    if (objectId && index !== null) move(objectId, index);
  };

  return (
    <OfficeDialogShell
      open={open}
      title={labels.title}
      description={labels.description}
      defaultActionLabel={labels.ok}
      cancelActionLabel={labels.cancel}
      onDefaultAction={() => (changed ? onApply(order) : onCancel())}
      onCancel={onCancel}
      returnFocusRef={returnFocusRef}
      testIdPrefix="presentation-reorder-objects"
      footerStart={
        <span className="inline-flex gap-1">
          <button
            type="button"
            className="rounded border px-2 py-1 text-xs disabled:opacity-50"
            data-testid="presentation-reorder-objects-forward"
            disabled={stepTarget(active, -1) === null}
            onClick={() => step(active, -1)}
          >
            <OfficeSemanticIcon semanticId="object.bringForward" size={14} />{" "}
            {labels.bringForward}
          </button>
          <button
            type="button"
            className="rounded border px-2 py-1 text-xs disabled:opacity-50"
            data-testid="presentation-reorder-objects-backward"
            disabled={stepTarget(active, 1) === null}
            onClick={() => step(active, 1)}
          >
            <OfficeSemanticIcon semanticId="object.sendBackward" size={14} />{" "}
            {labels.sendBackward}
          </button>
        </span>
      }
    >
      <ol
        ref={listRef}
        className="flex max-h-[52vh] flex-col gap-0 overflow-y-auto py-2"
        style={{ perspective: "900px" }}
        data-testid="presentation-reorder-objects-layers"
        role="listbox"
        aria-label={labels.title}
      >
        {order.map((objectId, index) => {
          const layer = byId.get(objectId);
          if (!layer) return null;
          const isActive = objectId === active;
          return (
            <li
              key={objectId}
              role="option"
              aria-selected={isActive}
              aria-label={labels.layer(index + 1, order.length, layer.label)}
              tabIndex={isActive || (!active && index === 0) ? 0 : -1}
              draggable={!layer.locked}
              data-testid={`presentation-reorder-objects-layer-${index}`}
              data-object-id={objectId}
              data-locked={layer.locked || undefined}
              className={`flex items-center gap-3 rounded-md border bg-background px-2 py-1.5 outline-none transition-transform ${
                isActive ? "border-primary ring-2 ring-primary/40" : "border-border"
              } ${layer.locked ? "opacity-70" : "cursor-grab"}`}
              style={{
                transform: "rotateX(18deg)",
                marginTop: index === 0 ? 0 : -6,
                zIndex: order.length - index,
                position: "relative",
              }}
              onClick={() => setActive(objectId)}
              onKeyDown={(event) => {
                if (event.altKey && (event.key === "ArrowUp" || event.key === "ArrowDown")) {
                  event.preventDefault();
                  step(objectId, event.key === "ArrowUp" ? -1 : 1);
                  return;
                }
                if (event.key === "ArrowUp" || event.key === "ArrowDown") {
                  event.preventDefault();
                  const next = order[index + (event.key === "ArrowUp" ? -1 : 1)];
                  if (!next) return;
                  setActive(next);
                  listRef.current
                    ?.querySelector<HTMLElement>(`[data-object-id="${CSS.escape(next)}"]`)
                    ?.focus();
                }
              }}
              onDragStart={(event) => {
                dragging.current = objectId;
                setActive(objectId);
                event.dataTransfer.effectAllowed = "move";
                event.dataTransfer.setData("text/plain", objectId);
              }}
              onDragOver={(event) => {
                if (dragging.current) event.preventDefault();
              }}
              onDrop={(event) => {
                event.preventDefault();
                const from = dragging.current;
                dragging.current = null;
                if (from) move(from, index);
              }}
              onDragEnd={() => {
                dragging.current = null;
              }}
            >
              <span
                className="block h-[54px] w-[96px] shrink-0 overflow-hidden rounded-sm border bg-white"
                aria-hidden="true"
              >
                {layer.preview}
              </span>
              <span className="min-w-0 flex-1 truncate text-sm">{layer.label}</span>
              {layer.hidden ? (
                <span className="text-xs text-muted-foreground">{labels.hidden}</span>
              ) : null}
              {layer.locked ? (
                <span className="text-xs text-muted-foreground">{labels.locked}</span>
              ) : null}
            </li>
          );
        })}
      </ol>
    </OfficeDialogShell>
  );
}
