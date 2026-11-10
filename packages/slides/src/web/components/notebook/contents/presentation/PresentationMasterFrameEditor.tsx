"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The placeholder frames of the master or layout shown at stage size in Slide
 * Master view (PPT-B030), movable and resizable as PowerPoint's are: drag a
 * frame to move it, a corner handle to resize it, or select it and nudge it
 * with the arrow keys; on a layout, Delete removes the selected placeholder
 * from it, and so does the frame's menu (a right-click, or a press and hold
 * on a touch screen, which has no Delete key). A gesture is previewed live
 * and committed once, on
 * release, as fractions of the slide; the window turns it into the master's
 * or the layout's frame, which every inheriting slide follows.
 */

import { useRef, useState, type PointerEvent as ReactPointerEvent } from "react";
import { useTranslations } from "@mrnote/core/host/next-intl.tsx";

import { PointerContextMenu } from "@mrnote/core/web/components/ui/PointerContextMenu.tsx";

export interface PresentationMasterFrame {
  readonly key: string;
  readonly label: string;
  readonly left: number;
  readonly top: number;
  readonly width: number;
  readonly height: number;
  /** Selectable and deletable, never moved or resized (the master's footers). */
  readonly fixed?: boolean;
  /** What the frame shows inside it (the footers' date, label or ‹#›). */
  readonly text?: string;
}

type Rect = { left: number; top: number; width: number; height: number };
type Handle = "move" | "nw" | "ne" | "sw" | "se";

const MIN_SIZE = 0.02;
const NUDGE = 0.005;

function clampRect(rect: Rect): Rect {
  const width = Math.max(MIN_SIZE, Math.min(1, rect.width));
  const height = Math.max(MIN_SIZE, Math.min(1, rect.height));
  return {
    width,
    height,
    left: Math.max(0, Math.min(1 - width, rect.left)),
    top: Math.max(0, Math.min(1 - height, rect.top)),
  };
}

function moved(start: Rect, handle: Handle, dx: number, dy: number): Rect {
  if (handle === "move") return clampRect({ ...start, left: start.left + dx, top: start.top + dy });
  let { left, top, width, height } = start;
  if (handle === "nw" || handle === "sw") {
    left += dx;
    width -= dx;
  } else {
    width += dx;
  }
  if (handle === "nw" || handle === "ne") {
    top += dy;
    height -= dy;
  } else {
    height += dy;
  }
  return clampRect({ left, top, width, height });
}

export function PresentationMasterFrameEditor({
  frames,
  disabled,
  onCommit,
  onDelete,
}: {
  frames: readonly PresentationMasterFrame[];
  disabled: boolean;
  onCommit: (key: string, rect: Rect) => void;
  /** Delete / Backspace on a selected frame removes it from the master or the layout (PPT-B030). */
  onDelete?: (key: string) => void;
}) {
  const hostRef = useRef<HTMLDivElement>(null);
  const gestureRef = useRef<{
    key: string;
    handle: Handle;
    startX: number;
    startY: number;
    start: Rect;
  } | null>(null);
  const [preview, setPreviewState] = useState<{ key: string; rect: Rect } | null>(null);
  // The gesture's last rect, read on release whatever render the handler is from.
  const previewRef = useRef<{ key: string; rect: Rect } | null>(null);
  const setPreview = (next: { key: string; rect: Rect } | null) => {
    previewRef.current = next;
    setPreviewState(next);
  };
  const [selected, setSelected] = useState<string | null>(null);
  const [menu, setMenu] = useState<{ key: string; label: string; x: number; y: number } | null>(null);
  const t = useTranslations("console-notebooks");

  const begin = (event: ReactPointerEvent<HTMLElement>, frame: PresentationMasterFrame, handle: Handle) => {
    if (disabled || event.button !== 0) return;
    event.preventDefault();
    event.stopPropagation();
    setSelected(frame.key);
    if (frame.fixed) {
      // Selected for Delete, as a click selects; nothing to drag.
      (event.currentTarget as HTMLElement).focus?.({ preventScroll: true });
      return;
    }
    event.currentTarget.setPointerCapture?.(event.pointerId);
    gestureRef.current = {
      key: frame.key,
      handle,
      startX: event.clientX,
      startY: event.clientY,
      start: { left: frame.left, top: frame.top, width: frame.width, height: frame.height },
    };
  };
  const move = (event: ReactPointerEvent<HTMLElement>) => {
    const gesture = gestureRef.current;
    const host = hostRef.current?.getBoundingClientRect();
    if (!gesture || !host || host.width <= 0 || host.height <= 0) return;
    const dx = (event.clientX - gesture.startX) / host.width;
    const dy = (event.clientY - gesture.startY) / host.height;
    setPreview({ key: gesture.key, rect: moved(gesture.start, gesture.handle, dx, dy) });
  };
  const end = () => {
    const gesture = gestureRef.current;
    gestureRef.current = null;
    const last = previewRef.current;
    if (gesture && last?.key === gesture.key) onCommit(gesture.key, last.rect);
    setPreview(null);
  };

  return (
    <div
      ref={hostRef}
      data-testid="presentation-master-frame-editor"
      style={{ position: "absolute", inset: 0 }}
      onPointerMove={move}
      onPointerUp={end}
      onPointerCancel={() => {
        gestureRef.current = null;
        setPreview(null);
      }}
    >
      {frames.map((frame) => {
        const rect = preview?.key === frame.key ? preview.rect : frame;
        const isSelected = selected === frame.key;
        return (
          <div
            key={frame.key}
            role="button"
            tabIndex={disabled ? -1 : 0}
            aria-label={frame.label}
            aria-pressed={isSelected}
            data-testid={`presentation-master-frame-${frame.key}`}
            onPointerDown={(event) => begin(event, frame, "move")}
            onFocus={() => setSelected(frame.key)}
            onContextMenu={(event) => {
              if (!onDelete || disabled) return;
              event.preventDefault();
              event.stopPropagation();
              setSelected(frame.key);
              setMenu({ key: frame.key, label: frame.label, x: event.clientX, y: event.clientY });
            }}
            onKeyDown={(event) => {
              if ((event.key === "Delete" || event.key === "Backspace") && onDelete && !disabled) {
                event.preventDefault();
                event.stopPropagation();
                setSelected(null);
                onDelete(frame.key);
                return;
              }
              const step = { ArrowLeft: [-NUDGE, 0], ArrowRight: [NUDGE, 0], ArrowUp: [0, -NUDGE], ArrowDown: [0, NUDGE] }[
                event.key
              ];
              if (!step || disabled || frame.fixed) return;
              event.preventDefault();
              onCommit(frame.key, clampRect({ ...frame, left: frame.left + step[0], top: frame.top + step[1] }));
            }}
            style={{
              position: "absolute",
              left: `${rect.left * 100}%`,
              top: `${rect.top * 100}%`,
              width: `${rect.width * 100}%`,
              height: `${rect.height * 100}%`,
              boxSizing: "border-box",
              border: isSelected ? "1px solid var(--ppt-accent, #c43e1c)" : "1px dashed rgba(15, 23, 42, 0.45)",
              cursor: disabled || frame.fixed ? "default" : "move",
              touchAction: "none",
              ...(frame.text
                ? {
                    display: "flex",
                    alignItems: "center",
                    justifyContent: "center",
                    overflow: "hidden",
                    fontSize: 11,
                    color: "rgba(71, 85, 105, 0.9)",
                    whiteSpace: "nowrap",
                  }
                : {}),
            }}
          >
            {frame.text ? <span aria-hidden="true">{frame.text}</span> : null}
            {isSelected && !disabled && !frame.fixed
              ? (["nw", "ne", "sw", "se"] as const).map((handle) => (
                  <span
                    key={handle}
                    data-testid={`presentation-master-frame-${frame.key}-${handle}`}
                    onPointerDown={(event) => begin(event, frame, handle)}
                    style={{
                      position: "absolute",
                      width: 8,
                      height: 8,
                      background: "Canvas",
                      border: "1px solid var(--ppt-accent, #c43e1c)",
                      left: handle.endsWith("w") ? -4 : undefined,
                      right: handle.endsWith("e") ? -4 : undefined,
                      top: handle.startsWith("n") ? -4 : undefined,
                      bottom: handle.startsWith("s") ? -4 : undefined,
                      cursor: handle === "nw" || handle === "se" ? "nwse-resize" : "nesw-resize",
                    }}
                  />
                ))
              : null}
          </div>
        );
      })}
      {menu && onDelete ? (
        <PointerContextMenu
          x={menu.x}
          y={menu.y}
          label={menu.label}
          heading={menu.label}
          testId="presentation-master-frame-menu"
          onClose={() => setMenu(null)}
          entries={[
            {
              id: "delete",
              label: t("presentation.contextMenu.delete"),
              testId: "presentation-master-frame-menu-delete",
              onSelect: () => {
                setSelected(null);
                onDelete(menu.key);
              },
            },
          ]}
        />
      ) : null}
    </div>
  );
}
