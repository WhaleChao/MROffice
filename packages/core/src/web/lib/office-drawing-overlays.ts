/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Framework-free drawing overlays shared by Docs (an imperative ProseMirror
 * node view), the deck and the worksheet (React, through the wrappers in
 * `components/office-ribbon/OfficeDrawingSelection.tsx`):
 *
 *  - the yellow adjustment handles of a preset (`a:avLst`),
 *  - Edit Points on the one custom-geometry model (`a:custGeom`),
 *  - the Freeform: Shape / Freeform: Scribble / Curve drawing tools.
 *
 * Each overlay is mounted INTO a host element that covers the shape's own box
 * (or the drawing surface) and is rotated with it but not mirrored; hosts own
 * selection, permission and persistence and receive plain values back.
 */

import { isMacPlatform } from "./mac-platform";
import {
  officeShapeAdjustDrag,
  officeShapeAdjustHandles,
  officeShapeHandleDisplayPoint,
  officeShapeLocalPoint,
} from "./office-shape-adjust-handles";
import {
  officeEditAddVertex,
  officeEditDeleteVertex,
  officeEditHitSegment,
  officeEditMoveHandle,
  officeEditMoveVertex,
  officeEditPathToSvg,
  officeEditPathVertices,
  officeEditSetPointType,
  officeEditSetSegmentKind,
  officeEditToggleClosed,
  officeEditVertexHandles,
  OFFICE_FREEFORM_MAX_POINTS,
  type OfficeEditPath,
  type OfficeEditPointType,
  type OfficePathPoint,
} from "./office-custom-path-edit";
import {
  buildCurveCustomGeometry,
  buildMixedFreeformCustomGeometry,
  buildScribbleCustomGeometry,
  type PresentationCustomGeometryTool,
  type PresentationCustomShapeResult,
  type PresentationFreeformDraftSegment,
} from "./presentation-freeform";

const SVG_NS = "http://www.w3.org/2000/svg";

export interface OfficeOverlayClassNames {
  readonly adjustLayer: string;
  readonly adjust: string;
  readonly editLayer?: string;
  readonly editVertex?: string;
  readonly editHandle?: string;
  readonly editMenu?: string;
  readonly drawLayer?: string;
}

interface OverlayFrame {
  readonly width: number;
  readonly height: number;
  readonly rotation?: number;
  readonly flipH?: boolean;
  readonly flipV?: boolean;
  /**
   * The host element already wears the flip (the deck mirrors its whole
   * block): points are drawn unmirrored, pointer math still undoes the flip.
   */
  readonly mirroredHost?: boolean;
}

/** A client point in the frame's own box pixels (the canvas zoom undone). */
function framePoint(layer: HTMLElement, frame: OverlayFrame, clientX: number, clientY: number): OfficePathPoint | null {
  const rect = layer.getBoundingClientRect();
  const rotation = frame.rotation ?? 0;
  const radians = (rotation * Math.PI) / 180;
  const span = frame.width * Math.abs(Math.cos(radians)) + frame.height * Math.abs(Math.sin(radians));
  const scale = span > 0 && rect.width > 0 ? rect.width / span : 1;
  if (!(scale > 0) || !Number.isFinite(scale)) return null;
  const cx = rect.left + rect.width / 2;
  const cy = rect.top + rect.height / 2;
  const point = officeShapeLocalPoint(
    { x: clientX, y: clientY },
    { left: cx - (frame.width * scale) / 2, top: cy - (frame.height * scale) / 2, width: frame.width * scale, height: frame.height * scale },
    rotation, frame.flipH, frame.flipV,
  );
  return { x: point.x / scale, y: point.y / scale };
}

function displayPoint(point: OfficePathPoint, frame: OverlayFrame): OfficePathPoint {
  return frame.mirroredHost
    ? point
    : officeShapeHandleDisplayPoint(point, frame.width, frame.height, frame.flipH, frame.flipV);
}

function stop(event: Event) {
  event.preventDefault();
  event.stopPropagation();
}

// ---------------------------------------------------------------------------
// Adjustment handles

export interface OfficeAdjustHandlesOptions extends OverlayFrame {
  readonly drawingPreset: string | null | undefined;
  readonly adjustments?: Readonly<Record<string, number>> | null;
  readonly label: string;
  readonly classNames: OfficeOverlayClassNames;
  readonly testIdPrefix?: string;
  readonly onPreview?: (next: Record<string, number>) => void;
  readonly onCommit: (next: Record<string, number>) => void;
  readonly onCancel?: () => void;
}

export interface OfficeOverlayController<Options> {
  update(options: Options): void;
  destroy(): void;
}

/** The yellow handles of one preset shape, mounted into `host`. */









function screenToLayerRatio(layer: HTMLElement): number {
  const screenWidth = layer.getBoundingClientRect().width;
  return screenWidth > 0 && layer.offsetWidth > 0 ? layer.offsetWidth / screenWidth : 1;
}

export function mountOfficeAdjustHandles(
  host: HTMLElement,
  initial: OfficeAdjustHandlesOptions,
): OfficeOverlayController<OfficeAdjustHandlesOptions> {
  let options = initial;
  let preview: Record<string, number> | null = null;
  let drag: { index: number; pointerId: number; last: Record<string, number> | null } | null = null;
  const layer = host.ownerDocument.createElement("div");
  layer.dataset.officeDrawingAdjustLayer = "true";
  host.append(layer);

  const cancel = () => {
    if (!drag && !preview) return;
    drag = null;
    preview = null;
    paint();
    options.onCancel?.();
  };
  const onKey = (event: KeyboardEvent) => {
    if (event.key === "Escape" && drag) {
      stop(event);
      cancel();
    }
  };
  host.ownerDocument.addEventListener("keydown", onKey, true);

  function paint() {
    layer.className = options.classNames.adjustLayer;
    const handles = options.drawingPreset
      ? officeShapeAdjustHandles(options.drawingPreset, options.width, options.height, preview ?? options.adjustments ?? null)
      : [];
    while (layer.children.length > handles.length) layer.lastElementChild?.remove();
    handles.forEach((handle, position) => {
      let button = layer.children[position] as HTMLButtonElement | undefined;
      if (!button) {
        button = host.ownerDocument.createElement("button");
        button.type = "button";
        button.tabIndex = -1;
        bind(button);
        layer.append(button);
      }
      const shown = displayPoint(handle, options);
      button.className = options.classNames.adjust;
      button.dataset.officeDrawingAdjust = String(handle.index);
      button.dataset.testid = `${options.testIdPrefix ?? "office-drawing-adjust"}-${handle.index}`;
      button.setAttribute("data-testid", button.dataset.testid);
      button.setAttribute("aria-label", options.label);
      button.title = options.label;
      button.style.left = `${(shown.x / options.width) * 100}%`;
      button.style.top = `${(shown.y / options.height) * 100}%`;
    });
  }

  function bind(button: HTMLButtonElement) {
    button.addEventListener("pointerdown", (event) => {
      if (event.button !== 0) return;
      stop(event);
      button.setPointerCapture?.(event.pointerId);
      drag = { index: Number(button.dataset.officeDrawingAdjust), pointerId: event.pointerId, last: null };
    });
    button.addEventListener("pointermove", (event) => {
      if (!drag || drag.pointerId !== event.pointerId || !options.drawingPreset) return;
      stop(event);
      const point = framePoint(layer, options, event.clientX, event.clientY);
      if (!point) return;
      const next = officeShapeAdjustDrag(options.drawingPreset, options.width, options.height,
        preview ?? options.adjustments ?? null, drag.index, point);
      if (!next) return;
      drag.last = next;
      preview = next;
      paint();
      options.onPreview?.(next);
    });
    const finish = (event: PointerEvent) => {
      if (!drag || drag.pointerId !== event.pointerId) return;
      stop(event);
      const last = drag.last;
      drag = null;
      preview = null;
      if (last) options.onCommit(last);
      else options.onCancel?.();
      paint();
    };
    button.addEventListener("pointerup", finish);
    button.addEventListener("pointercancel", () => cancel());
    button.addEventListener("click", (event) => event.stopPropagation());
    button.addEventListener("dblclick", (event) => event.stopPropagation());
    button.addEventListener("mousedown", (event) => event.stopPropagation());
  }

  paint();
  return {
    update(next) {
      options = next;
      if (!drag) preview = null;
      paint();
    },
    destroy() {
      host.ownerDocument.removeEventListener("keydown", onKey, true);
      layer.remove();
    },
  };
}

// ---------------------------------------------------------------------------
// Edit Points

export interface OfficeEditPointsLabels {
  readonly vertex: string;
  readonly handle: string;
  readonly addPoint: string;
  readonly deletePoint: string;
  readonly openPath: string;
  readonly closePath: string;
  readonly smoothPoint: string;
  readonly straightPoint: string;
  readonly cornerPoint: string;
  readonly straightSegment: string;
  readonly curvedSegment: string;
  readonly exit: string;
}

/** Office's English names, for a host with no translator (tests, defaults). */
export const OFFICE_EDIT_POINTS_LABELS: OfficeEditPointsLabels = {
  vertex: "Vertex", handle: "Vertex handle", addPoint: "Add Point", deletePoint: "Delete Point",
  openPath: "Open Path", closePath: "Close Path", smoothPoint: "Smooth Point", straightPoint: "Straight Point",
  cornerPoint: "Corner Point", straightSegment: "Straight Segment", curvedSegment: "Curved Segment",
  exit: "Exit Edit Points",
};

/** The same labels in the viewer's locale, from `console-notebooks` (`office.editPoints.*`). */
export function officeEditPointsLabels(
  t: (key: string) => string,
): OfficeEditPointsLabels {
  const ids = Object.keys(
    OFFICE_EDIT_POINTS_LABELS,
  ) as (keyof OfficeEditPointsLabels)[];
  return Object.fromEntries(
    ids.map((id) => [id, t(`office.editPoints.${id}`)]),
  ) as unknown as OfficeEditPointsLabels;
}

export interface OfficeEditPointsOptions extends OverlayFrame {
  readonly path: OfficeEditPath;
  readonly labels: OfficeEditPointsLabels;
  readonly classNames: OfficeOverlayClassNames;
  readonly testIdPrefix?: string;
  /** Every pointer move: the outline the shape should be drawn with. */
  readonly onPreview?: (path: OfficeEditPath) => void;
  /** One finished edit (a drag, a menu command): write it. */
  readonly onCommit: (path: OfficeEditPath) => void;
  /** Esc or Exit Edit Points. */
  readonly onExit: () => void;
}

type EditTarget =
  | { kind: "vertex"; figure: number; index: number }
  | { kind: "handle"; figure: number; index: number; side: "incoming" | "outgoing" };

/** Edit Points over one shape, mounted into `host` (the shape's own box). */
export function mountOfficeEditPoints(
  host: HTMLElement,
  initial: OfficeEditPointsOptions,
): OfficeOverlayController<OfficeEditPointsOptions> {
  let options = initial;
  let path = initial.path;
  let active: { figure: number; index: number } | null = null;
  let drag: { target: EditTarget; pointerId: number; moved: boolean } | null = null;
  const doc = host.ownerDocument;
  const layer = doc.createElement("div");
  layer.dataset.officeDrawingEditPoints = "true";
  layer.setAttribute("data-testid", `${initial.testIdPrefix ?? "office-edit-points"}-layer`);
  const svg = doc.createElementNS(SVG_NS, "svg");
  svg.setAttribute("overflow", "visible");
  Object.assign(svg.style, { position: "absolute", inset: "0", width: "100%", height: "100%", overflow: "visible" });
  const outline = doc.createElementNS(SVG_NS, "path");
  outline.setAttribute("fill", "none");
  outline.setAttribute("stroke", "#c00000");
  outline.setAttribute("stroke-width", "1");
  outline.setAttribute("vector-effect", "non-scaling-stroke");
  outline.style.pointerEvents = "none";
  const hit = doc.createElementNS(SVG_NS, "path");
  hit.setAttribute("fill", "none");
  hit.setAttribute("stroke", "transparent");
  hit.setAttribute("stroke-width", "10");
  hit.style.pointerEvents = "stroke";
  hit.style.cursor = "crosshair";
  const guides = doc.createElementNS(SVG_NS, "g");
  guides.style.pointerEvents = "none";
  svg.append(hit, outline, guides);
  layer.append(svg);
  const points = doc.createElement("div");
  Object.assign(points.style, { position: "absolute", inset: "0", pointerEvents: "none" });
  layer.append(points);
  host.append(layer);
  let menu: HTMLElement | null = null;

  const closeMenu = () => {
    menu?.remove();
    menu = null;
  };
  const commit = (next: OfficeEditPath) => {
    path = next;
    paint();
    options.onCommit(next);
  };
  const onKey = (event: KeyboardEvent) => {
    if (event.key !== "Escape") return;
    stop(event);
    if (menu) {
      closeMenu();
      return;
    }
    if (drag) {
      drag = null;
      path = options.path;
      paint();
      options.onPreview?.(path);
      return;
    }
    options.onExit();
  };
  doc.addEventListener("keydown", onKey, true);
  const onOutside = (event: PointerEvent) => {
    if (menu && !menu.contains(event.target as Node)) closeMenu();
  };
  doc.addEventListener("pointerdown", onOutside, true);

  function toDisplay(point: OfficePathPoint) {
    const shown = displayPoint(point, options);
    return { left: `${(shown.x / options.width) * 100}%`, top: `${(shown.y / options.height) * 100}%` };
  }

  function paint() {
    layer.className = options.classNames.editLayer ?? options.classNames.adjustLayer;
    const flip = options.mirroredHost
      ? ""
      : `${options.flipH ? `translate(${options.width} 0) scale(-1 1)` : ""} ${options.flipV ? `translate(0 ${options.height}) scale(1 -1)` : ""}`.trim();
    svg.setAttribute("viewBox", `0 0 ${options.width} ${options.height}`);
    svg.setAttribute("preserveAspectRatio", "none");
    const d = officeEditPathToSvg(path);
    outline.setAttribute("d", d);
    hit.setAttribute("d", d);
    outline.setAttribute("transform", flip);
    hit.setAttribute("transform", flip);
    guides.setAttribute("transform", flip);
    guides.replaceChildren();
    points.replaceChildren();
    for (const vertex of officeEditPathVertices(path)) {
      const button = doc.createElement("button");
      button.type = "button";
      button.tabIndex = -1;
      button.className = options.classNames.editVertex ?? options.classNames.adjust;
      button.dataset.officeEditVertex = `${vertex.figure}:${vertex.index}`;
      button.setAttribute("data-testid", `${options.testIdPrefix ?? "office-edit-points"}-vertex-${vertex.figure}-${vertex.index}`);
      button.setAttribute("aria-label", `${options.labels.vertex} ${vertex.index + 1}`);
      if (active && active.figure === vertex.figure && active.index === vertex.index) button.dataset.active = "true";
      Object.assign(button.style, toDisplay(vertex.point));
      bindTarget(button, { kind: "vertex", figure: vertex.figure, index: vertex.index });
      points.append(button);
    }
    if (active) {
      const figure = path.figures[active.figure];
      const vertex = figure ? officeEditPathVertices(path).find((candidate) => candidate.figure === active!.figure && candidate.index === active!.index) : null;
      if (vertex) {
        const handles = officeEditVertexHandles(path, active.figure, active.index);
        for (const side of ["incoming", "outgoing"] as const) {
          const handle = handles[side];
          if (!handle) continue;
          const line = doc.createElementNS(SVG_NS, "line");
          line.setAttribute("x1", String(vertex.point.x));
          line.setAttribute("y1", String(vertex.point.y));
          line.setAttribute("x2", String(handle.x));
          line.setAttribute("y2", String(handle.y));
          line.setAttribute("stroke", "#5f5f5f");
          line.setAttribute("stroke-width", "1");
          line.setAttribute("vector-effect", "non-scaling-stroke");
          guides.append(line);
          const button = doc.createElement("button");
          button.type = "button";
          button.tabIndex = -1;
          button.className = options.classNames.editHandle ?? options.classNames.adjust;
          button.setAttribute("data-testid", `${options.testIdPrefix ?? "office-edit-points"}-handle-${side}`);
          button.setAttribute("aria-label", options.labels.handle);
          Object.assign(button.style, toDisplay(handle));
          bindTarget(button, { kind: "handle", figure: active.figure, index: active.index, side });
          points.append(button);
        }
      }
    }
  }

  function pointType(event: PointerEvent): OfficeEditPointType {
    if (event.shiftKey) return "smooth";
    if (event.ctrlKey || event.metaKey) return "straight";
    return "corner";
  }

  function bindTarget(button: HTMLButtonElement, target: EditTarget) {
    button.addEventListener("pointerdown", (event) => {
      if (event.button !== 0) return;
      stop(event);
      closeMenu();
      // Ctrl+click on a vertex deletes it (Office).
      if (target.kind === "vertex" && (event.ctrlKey || event.metaKey)) {
        const next = officeEditDeleteVertex(path, target.figure, target.index);
        if (next !== path) {
          active = null;
          commit(next);
        }
        return;
      }
      if (target.kind === "vertex") active = { figure: target.figure, index: target.index };
      // The LAYER holds the capture: `paint()` rebuilds the vertex buttons,
      // and a capture on the pressed button went with it, so dragging a
      // vertex past the frame lost the pointer (PPT-B223).
      layer.setPointerCapture?.(event.pointerId);
      drag = { target, pointerId: event.pointerId, moved: false };
      paint();
    });
    button.addEventListener("contextmenu", (event) => {
      stop(event);
      if (target.kind !== "vertex") return;
      active = { figure: target.figure, index: target.index };
      paint();
      openMenu(event.clientX, event.clientY, { vertex: target });
    });
  }

  layer.addEventListener("pointermove", (event) => {
    if (!drag || drag.pointerId !== event.pointerId) return;
    stop(event);
    const point = framePoint(layer, options, event.clientX, event.clientY);
    if (!point) return;
    drag.moved = true;
    const base = options.path;
    const { target } = drag;
    const next = target.kind === "vertex"
      ? officeEditMoveVertex(base, target.figure, target.index, point)
      : officeEditMoveHandle(base, target.figure, target.index, target.side, point, pointType(event));
    path = next;
    paint();
    options.onPreview?.(next);
  }, true);
  const finish = (event: PointerEvent) => {
    if (!drag || drag.pointerId !== event.pointerId) return;
    stop(event);
    const moved = drag.moved;
    drag = null;
    if (moved) commit(path);
    else paint();
  };
  layer.addEventListener("pointerup", finish, true);
  layer.addEventListener("pointercancel", () => {
    drag = null;
    path = options.path;
    paint();
  }, true);

  // The outline itself: Ctrl+click adds a vertex; a drag adds one and moves it.
  hit.addEventListener("pointerdown", (event) => {
    if (event.button !== 0) return;
    stop(event);
    closeMenu();
    const point = framePoint(layer, options, event.clientX, event.clientY);
    if (!point) return;
    const found = officeEditHitSegment(path, point, 8);
    if (!found) return;
    const next = officeEditAddVertex(path, found.figure, found.segment, found.t);
    if (next === path) return;
    const index = found.segment + 1;
    active = { figure: found.figure, index };
    if (event.ctrlKey || event.metaKey) {
      commit(next);
      return;
    }
    path = next;
    // Hand the new vertex to the drag as if it had been pressed.
    drag = { target: { kind: "vertex", figure: found.figure, index }, pointerId: event.pointerId, moved: false };
    layer.setPointerCapture?.(event.pointerId);
    options = { ...options, path: next };
    paint();
  });
  hit.addEventListener("contextmenu", (event) => {
    stop(event);
    const point = framePoint(layer, options, event.clientX, event.clientY);
    const found = point ? officeEditHitSegment(path, point, 8) : null;
    if (found) openMenu(event.clientX, event.clientY, { segment: found });
  });

  function openMenu(
    clientX: number,
    clientY: number,
    at: { vertex?: { figure: number; index: number }; segment?: { figure: number; segment: number; t: number } },
  ) {
    closeMenu();
    const labels = options.labels;
    const items: [string, string, () => OfficeEditPath | "exit"][] = [];
    if (at.segment) {
      const { figure, segment, t } = at.segment;
      const kind = path.figures[figure]?.segments[segment]?.kind;
      items.push(["add-point", labels.addPoint, () => officeEditAddVertex(path, figure, segment, t)]);
      items.push(kind === "cubic"
        ? ["straight-segment", labels.straightSegment, () => officeEditSetSegmentKind(path, figure, segment, "line")]
        : ["curved-segment", labels.curvedSegment, () => officeEditSetSegmentKind(path, figure, segment, "cubic")]);
    }
    if (at.vertex) {
      const { figure, index } = at.vertex;
      const closed = path.figures[figure]?.closed;
      items.push(["delete-point", labels.deletePoint, () => officeEditDeleteVertex(path, figure, index)]);
      items.push(closed
        ? ["open-path", labels.openPath, () => officeEditToggleClosed(path, figure, index)]
        : ["close-path", labels.closePath, () => officeEditToggleClosed(path, figure, index)]);
      items.push(["smooth-point", labels.smoothPoint, () => officeEditSetPointType(path, figure, index, "smooth")]);
      items.push(["straight-point", labels.straightPoint, () => officeEditSetPointType(path, figure, index, "straight")]);
      items.push(["corner-point", labels.cornerPoint, () => officeEditSetPointType(path, figure, index, "corner")]);
    }
    items.push(["exit", labels.exit, () => "exit"]);
    const element = doc.createElement("div");
    element.setAttribute("role", "menu");
    element.className = options.classNames.editMenu ?? "";
    element.setAttribute("data-testid", `${options.testIdPrefix ?? "office-edit-points"}-menu`);
    Object.assign(element.style, { position: "fixed", left: `${clientX}px`, top: `${clientY}px`, zIndex: "2147483000" });
    for (const [id, label, run] of items) {
      const item = doc.createElement("button");
      item.type = "button";
      item.setAttribute("role", "menuitem");
      item.textContent = label;
      item.setAttribute("data-testid", `${options.testIdPrefix ?? "office-edit-points"}-menu-${id}`);
      item.addEventListener("pointerdown", (event) => event.stopPropagation());
      item.addEventListener("click", (event) => {
        stop(event);
        closeMenu();
        const result = run();
        if (result === "exit") {
          options.onExit();
          return;
        }
        if (result !== path) commit(result);
      });
      element.append(item);
    }
    doc.body.append(element);
    menu = element;
  }

  paint();
  return {
    update(next) {
      options = next;
      if (!drag) path = next.path;
      paint();
    },
    destroy() {
      closeMenu();
      doc.removeEventListener("keydown", onKey, true);
      doc.removeEventListener("pointerdown", onOutside, true);
      layer.remove();
    },
  };
}

// ---------------------------------------------------------------------------
// Freeform: Shape, Freeform: Scribble, Curve

export interface OfficeFreeformDrawOptions {
  readonly tool: PresentationCustomGeometryTool;
  /** CSS pixels per surface unit the host lays shapes out in (its zoom). */
  readonly scale?: number;
  readonly classNames: OfficeOverlayClassNames;
  readonly testId?: string;
  /** The finished outline, its box in surface units (before zoom). */
  readonly onComplete: (result: PresentationCustomShapeResult) => void;
  readonly onCancel: () => void;
  /**
   * The words of the Finish / Cancel bar a touch screen gets for an open
   * Freeform or Curve: a finger has no Esc or Enter, and a double tap is
   * not a reliable double-click. Without them, no bar.
   */
  readonly touchKeys?: { readonly finish: string; readonly cancel: string };
  /**
   * Esc ends the drawing, keeping what is drawn, as PowerPoint's Custom
   * motion path does ("To stop drawing a custom path, press Esc"); without
   * it, Esc cancels, as a shape's drawing does.
   */
  readonly escapeFinishes?: boolean;
}

const DRAW_EMU = 9525;

/**
 * The drawing tools over a host surface. Office's gestures (Microsoft Support,
 * "Draw or edit a freeform shape"): Freeform -- click for a straight segment,
 * drag for a drawn one; Scribble -- drag, release to finish; Curve -- click
 * each bend. Double-click leaves the figure open, a click back on the start
 * closes it, Esc cancels. The outline is built by the shared
 * `presentation-freeform` builders the deck uses and capped at the shared
 * 512-vertex limit.
 */
export function mountOfficeFreeformDraw(
  surface: HTMLElement,
  initial: OfficeFreeformDrawOptions,
): OfficeOverlayController<OfficeFreeformDrawOptions> {
  let options = initial;
  const doc = surface.ownerDocument;
  const layer = doc.createElement("div");
  layer.dataset.officeFreeformDraw = initial.tool;
  layer.setAttribute("data-testid", initial.testId ?? "office-freeform-draw");
  Object.assign(layer.style, { position: "absolute", inset: "0", cursor: "crosshair", zIndex: "40", touchAction: "none" });
  const svg = doc.createElementNS(SVG_NS, "svg");
  Object.assign(svg.style, { position: "absolute", inset: "0", width: "100%", height: "100%", overflow: "visible", pointerEvents: "none" });
  const preview = doc.createElementNS(SVG_NS, "path");
  preview.setAttribute("fill", "none");
  preview.setAttribute("stroke", "#1f4e79");
  preview.setAttribute("stroke-width", "1.5");
  svg.append(preview);
  layer.append(svg);
  surface.append(layer);

  // A touch screen's way to finish or cancel (see `touchKeys`): a bar at the
  // foot of the viewport, where the drawing surface's own scroll cannot take
  // it out of sight. A mouse keeps Office's double-click, Enter and Esc.
  const touchBar = doc.createElement("div");
  touchBar.className = "floating-toolbar";
  touchBar.setAttribute("role", "toolbar");
  touchBar.setAttribute("data-testid", `${initial.testId ?? "office-freeform-draw"}-touch-bar`);
  Object.assign(touchBar.style, {
    position: "fixed",
    left: "50%",
    bottom: "max(24px, env(safe-area-inset-bottom, 0px))",
    translate: "-50% 0",
    zIndex: "1602",
    gap: "6px",
  });
  const touchKey = (onPress: () => void) => {
    const button = doc.createElement("button");
    button.type = "button";
    button.className = "floating-toolbar-btn";
    Object.assign(button.style, { width: "auto", minWidth: "64px", height: "44px", padding: "0 16px", fontSize: "14px" });
    button.addEventListener("click", onPress);
    touchBar.append(button);
    return button;
  };
  const touchFinish = touchKey(() => finish(false));
  const touchCancel = touchKey(() => {
    reset();
    options.onCancel();
  });
  const coarse = Boolean(doc.defaultView?.matchMedia?.("(any-pointer: coarse)").matches);
  const placeTouchBar = () => {
    const keys = options.touchKeys;
    const wanted = Boolean(coarse && keys && options.tool !== "scribble");
    if (!wanted || !keys) {
      touchBar.remove();
      return;
    }
    touchFinish.textContent = keys.finish;
    touchCancel.textContent = keys.cancel;
    touchFinish.disabled = options.tool === "curve" ? knots.length < 2 : !start || segments.length === 0;
    if (!touchBar.isConnected) doc.body.append(touchBar);
  };

  // Surface units, i.e. CSS pixels / scale.
  let start: OfficePathPoint | null = null;
  let segments: PresentationFreeformDraftSegment[] = [];
  let knots: OfficePathPoint[] = [];
  let samples: OfficePathPoint[] = [];
  let pressed: { pointerId: number; origin: OfficePathPoint; dragging: boolean; curve: OfficePathPoint[] } | null = null;
  let hover: OfficePathPoint | null = null;

  const scale = () => (options.scale && options.scale > 0 ? options.scale : 1);
  const local = (event: PointerEvent | MouseEvent): OfficePathPoint => {
    const rect = layer.getBoundingClientRect();
    return { x: (event.clientX - rect.left) / scale(), y: (event.clientY - rect.top) / scale() };
  };
  const emu = (point: OfficePathPoint) => ({ x: point.x * DRAW_EMU, y: point.y * DRAW_EMU });
  const vertexCount = () =>
    options.tool === "curve" ? knots.length : options.tool === "scribble" ? samples.length : 1 + segments.reduce((sum, segment) => sum + (segment.type === "lineTo" ? 1 : segment.samples.length), 0);

  function paint() {
    const s = scale() * screenToLayerRatio(layer);
    const toPx = (point: OfficePathPoint) => `${point.x * s} ${point.y * s}`;
    let d = "";
    if (options.tool === "scribble") {
      d = samples.map((point, index) => `${index ? "L" : "M"}${toPx(point)}`).join(" ");
    } else if (options.tool === "curve") {
      const points = hover ? [...knots, hover] : knots;
      d = points.map((point, index) => `${index ? "L" : "M"}${toPx(point)}`).join(" ");
    } else if (start) {
      d = `M${toPx(start)}`;
      for (const segment of segments) {
        if (segment.type === "lineTo") d += ` L${toPx(segment.point)}`;
        else d += segment.samples.map((point) => ` L${toPx(point)}`).join("");
      }
      if (pressed?.dragging) d += pressed.curve.map((point) => ` L${toPx(point)}`).join("");
      else if (hover) d += ` L${toPx(hover)}`;
    }
    preview.setAttribute("d", d);
    placeTouchBar();
  }

  const reset = () => {
    start = null;
    segments = [];
    knots = [];
    samples = [];
    pressed = null;
    hover = null;
    paint();
  };

  const toSurface = (result: PresentationCustomShapeResult | null): PresentationCustomShapeResult | null =>
    result
      ? {
          left: result.left / DRAW_EMU,
          top: result.top / DRAW_EMU,
          width: result.width / DRAW_EMU,
          height: result.height / DRAW_EMU,
          customPath: result.customPath,
        }
      : null;

  const finish = (closed: boolean) => {
    let result: PresentationCustomShapeResult | null = null;
    if (options.tool === "curve") {
      result = buildCurveCustomGeometry(knots.map(emu), closed);
    } else if (options.tool === "scribble") {
      result = buildScribbleCustomGeometry(samples.map(emu));
    } else if (start) {
      result = buildMixedFreeformCustomGeometry(
        emu(start),
        segments.map((segment) => segment.type === "lineTo"
          ? { type: "lineTo", point: emu(segment.point) }
          : { type: "curveTo", samples: segment.samples.map(emu) }),
        closed,
      );
    }
    const surfaceResult = toSurface(result);
    reset();
    if (surfaceResult) options.onComplete(surfaceResult);
    else options.onCancel();
  };

  // A click back within 8 px of the first point closes the figure, once it
  // has at least three vertices (Office's "click near the starting point").
  const nearStart = (point: OfficePathPoint) => {
    const origin = options.tool === "curve" ? knots[0] : start;
    if (!origin) return false;
    const count = options.tool === "curve" ? knots.length : segments.length + 1;
    return count >= 3 && Math.hypot(origin.x - point.x, origin.y - point.y) * scale() <= 8;
  };

  layer.addEventListener("pointerdown", (event) => {
    if (event.button !== 0) return;
    stop(event);
    layer.setPointerCapture?.(event.pointerId);
    const point = local(event);
    if (options.tool === "scribble") {
      samples = [point];
      pressed = { pointerId: event.pointerId, origin: point, dragging: true, curve: [] };
      paint();
      return;
    }
    if (options.tool === "curve") {
      if (knots.length && nearStart(point)) {
        finish(true);
        return;
      }
      if (vertexCount() < OFFICE_FREEFORM_MAX_POINTS) knots.push(point);
      paint();
      return;
    }
    if (!start) {
      start = point;
      pressed = { pointerId: event.pointerId, origin: point, dragging: false, curve: [] };
      paint();
      return;
    }
    if (nearStart(point)) {
      finish(true);
      return;
    }
    pressed = { pointerId: event.pointerId, origin: point, dragging: false, curve: [] };
  });
  layer.addEventListener("pointermove", (event) => {
    const point = local(event);
    hover = point;
    if (pressed && pressed.pointerId === event.pointerId) {
      if (options.tool === "scribble") {
        if (samples.length < OFFICE_FREEFORM_MAX_POINTS * 8) samples.push(point);
      } else if (options.tool === "freeform") {
        const moved = Math.hypot(point.x - pressed.origin.x, point.y - pressed.origin.y) * scale() > 3;
        if (moved) pressed.dragging = true;
        if (pressed.dragging && vertexCount() + pressed.curve.length < OFFICE_FREEFORM_MAX_POINTS) pressed.curve.push(point);
      }
    }
    paint();
  });
  layer.addEventListener("pointerup", (event) => {
    if (!pressed || pressed.pointerId !== event.pointerId) return;
    stop(event);
    const point = local(event);
    if (options.tool === "scribble") {
      samples.push(point);
      pressed = null;
      finish(false);
      return;
    }
    if (options.tool === "freeform") {
      const origin = pressed.origin;
      if (pressed.dragging) {
        // The first press may itself start a drawn segment from `start`.
        segments.push({ type: "curveTo", samples: [...pressed.curve, point] });
      } else if (origin !== start) {
        segments.push({ type: "lineTo", point: origin });
      }
    }
    pressed = null;
    paint();
  });
  layer.addEventListener("dblclick", (event) => {
    stop(event);
    if (options.tool === "freeform" && segments.length > 0) {
      // The second click of the double-click added a zero-length segment.
      const last = segments[segments.length - 1];
      const previous = segments.length > 1 ? segments[segments.length - 2] : null;
      const before = previous?.type === "lineTo" ? previous.point : segments.length === 1 ? start : null;
      if (last?.type === "lineTo" && before &&
        Math.hypot(last.point.x - before.x, last.point.y - before.y) * scale() < 4) segments.pop();
    }
    if (options.tool === "curve" && knots.length > 1) {
      const a = knots[knots.length - 1];
      const b = knots[knots.length - 2];
      if (Math.hypot(a.x - b.x, a.y - b.y) * scale() < 4) knots.pop();
    }
    finish(false);
  });
  const onKey = (event: KeyboardEvent) => {
    if (event.key === "Escape") {
      stop(event);
      if (options.escapeFinishes) {
        finish(false);
        return;
      }
      reset();
      options.onCancel();
    } else if (event.key === "Enter") {
      stop(event);
      finish(false);
    }
  };
  doc.addEventListener("keydown", onKey, true);

  paint();
  return {
    update(next) {
      if (next.tool !== options.tool) reset();
      options = next;
      layer.dataset.officeFreeformDraw = next.tool;
      paint();
    },
    destroy() {
      doc.removeEventListener("keydown", onKey, true);
      layer.remove();
      touchBar.remove();
    },
  };
}

/** The frame a drawn insertion asked for, in surface units (before zoom). */
export interface OfficeShapeDrawFrame {
  readonly left: number;
  readonly top: number;
  readonly width: number;
  readonly height: number;
  /**
   * Where the press began and where it ended, in surface units: a line or
   * connector runs from `start` to `end` (the frame is their bounding box).
   * A plain click sets both to the click point.
   */
  readonly start: { readonly x: number; readonly y: number };
  readonly end: { readonly x: number; readonly y: number };
  /** The pointer never moved: the host's default size was placed at the click. */
  readonly clicked: boolean;
}

export interface OfficeShapeDrawOptions {
  /** CSS pixels per surface unit the host lays shapes out in (its zoom). */
  readonly scale?: number;
  /**
   * The surface's width in its own (design) px: the scale is then read from
   * the layer's on-screen width, so a view zoom on top of the host's
   * transform (PowerPoint's 104 %) is included. Wins over `scale`.
   */
  readonly designWidth?: number;
  /**
   * The size a plain click inserts, in surface units. Office inserts a
   * 1" x 1" shape (96 x 96 CSS px) with its top-left at the click.
   */
  readonly defaultSize: { readonly width: number; readonly height: number };
  /**
   * A line / connector: the drag is a segment from the press to the release
   * (any direction); Shift snaps it to 45-degree steps. Otherwise the drag
   * is a box and Shift keeps it square.
   */
  readonly line?: boolean;
  /** Lock Drawing Mode: stay armed after a shape is placed (Esc leaves). */
  readonly lock?: boolean;
  /**
   * The host's object snapping (lib/office-object-snap): the press and the
   * far corner go where this puts them, in surface units. Called with null
   * when the gesture ends, so the host can take its guides away.
   */
  readonly snapPoint?: (point: OfficePathPoint | null, event: PointerEvent | null) => OfficePathPoint | null;
  readonly classNames: OfficeOverlayClassNames;
  readonly testId?: string;
  readonly onComplete: (frame: OfficeShapeDrawFrame) => void;
  readonly onCancel: () => void;
}

/**
 * Office's drawing gesture for a gallery shape (Microsoft Support, "Add
 * shapes"): the pointer is a crosshair; "click in the workspace and drag to
 * draw the shape"; "press and hold Shift while you drag" for a square /
 * circle; Ctrl (Option on a Mac) draws it out from the press as its centre;
 * a plain click places the shape at its default size; Esc leaves the
 * mode (and Lock Drawing Mode). The host converts the frame to its own
 * units and inserts; the overlay never touches the document.
 */
export function mountOfficeShapeDraw(
  surface: HTMLElement,
  initial: OfficeShapeDrawOptions,
): OfficeOverlayController<OfficeShapeDrawOptions> {
  let options = initial;
  const doc = surface.ownerDocument;
  const layer = doc.createElement("div");
  layer.dataset.officeShapeDraw = initial.line ? "line" : "box";
  layer.setAttribute("data-testid", initial.testId ?? "office-shape-draw");
  Object.assign(layer.style, { position: "absolute", inset: "0", cursor: "crosshair", zIndex: "40", touchAction: "none" });
  const svg = doc.createElementNS(SVG_NS, "svg");
  Object.assign(svg.style, { position: "absolute", inset: "0", width: "100%", height: "100%", overflow: "visible", pointerEvents: "none" });
  const preview = doc.createElementNS(SVG_NS, "path");
  preview.setAttribute("fill", "none");
  preview.setAttribute("stroke", "#1f4e79");
  preview.setAttribute("stroke-width", "1.5");
  preview.setAttribute("stroke-dasharray", "4 3");
  svg.append(preview);
  layer.append(svg);
  surface.append(layer);

  let pressed: {
    pointerId: number;
    origin: OfficePathPoint;
    current: OfficePathPoint;
    shift: boolean;
    center: boolean;
  } | null = null;
  // Ctrl (Option on a Mac, where Control+click is the right click) draws
  // from the centre out: the press is the shape's centre (a line's
  // midpoint), as Office's drawing does.
  const fromCenterKey = (event: PointerEvent) => (isMacPlatform() ? event.altKey : event.ctrlKey);

  const scale = () => {
    if (options.designWidth && options.designWidth > 0) {
      const width = layer.getBoundingClientRect().width;
      if (width > 0) return width / options.designWidth;
    }
    return options.scale && options.scale > 0 ? options.scale : 1;
  };
  const local = (event: PointerEvent | MouseEvent): OfficePathPoint => {
    const rect = layer.getBoundingClientRect();
    return { x: (event.clientX - rect.left) / scale(), y: (event.clientY - rect.top) / scale() };
  };
  /** The pointer's place after the host's object snapping. */
  const snapped = (event: PointerEvent): OfficePathPoint => {
    const point = local(event);
    return options.snapPoint?.(point, event) ?? point;
  };

  /** The end point after Shift's constraint (square box / 45-degree line). */
  const constrained = (origin: OfficePathPoint, point: OfficePathPoint, shift: boolean): OfficePathPoint => {
    if (!shift) return point;
    const dx = point.x - origin.x;
    const dy = point.y - origin.y;
    if (options.line) {
      const step = Math.PI / 4;
      const angle = Math.round(Math.atan2(dy, dx) / step) * step;
      const length = Math.hypot(dx, dy);
      return { x: origin.x + Math.cos(angle) * length, y: origin.y + Math.sin(angle) * length };
    }
    const side = Math.max(Math.abs(dx), Math.abs(dy));
    return { x: origin.x + Math.sign(dx || 1) * side, y: origin.y + Math.sign(dy || 1) * side };
  };

  /** Where the drawn shape begins: the press, or (from the centre) the
   * press mirrored through itself. */
  const startOf = (origin: OfficePathPoint, end: OfficePathPoint, center: boolean): OfficePathPoint =>
    center ? { x: 2 * origin.x - end.x, y: 2 * origin.y - end.y } : origin;
  const frameOf = (origin: OfficePathPoint, end: OfficePathPoint, center = false): OfficeShapeDrawFrame => {
    const start = startOf(origin, end, center);
    return {
      left: Math.min(start.x, end.x),
      top: Math.min(start.y, end.y),
      width: Math.abs(end.x - start.x),
      height: Math.abs(end.y - start.y),
      start,
      end,
      clicked: false,
    };
  };

  function paint() {
    if (!pressed) {
      preview.setAttribute("d", "");
      return;
    }
    const s = scale() * screenToLayerRatio(layer);
    const end = constrained(pressed.origin, pressed.current, pressed.shift);
    if (options.line) {
      const start = startOf(pressed.origin, end, pressed.center);
      preview.setAttribute("d", `M${start.x * s} ${start.y * s} L${end.x * s} ${end.y * s}`);
      return;
    }
    const frame = frameOf(pressed.origin, end, pressed.center);
    preview.setAttribute(
      "d",
      `M${frame.left * s} ${frame.top * s} h${frame.width * s} v${frame.height * s} h${-frame.width * s} Z`,
    );
  }

  const reset = () => {
    pressed = null;
    options.snapPoint?.(null, null);
    paint();
  };

  layer.addEventListener("pointerdown", (event) => {
    if (event.button !== 0) return;
    stop(event);
    try {
      layer.setPointerCapture?.(event.pointerId);
    } catch {
      // A synthetic pointer (no id) cannot be captured; the drag still works.
    }
    const point = snapped(event);
    pressed = {
      pointerId: event.pointerId,
      origin: point,
      current: point,
      shift: event.shiftKey,
      center: fromCenterKey(event),
    };
    paint();
  });
  layer.addEventListener("pointermove", (event) => {
    if (!pressed || pressed.pointerId !== event.pointerId) return;
    pressed.current = snapped(event);
    pressed.shift = event.shiftKey;
    pressed.center = fromCenterKey(event);
    paint();
  });
  layer.addEventListener("pointerup", (event) => {
    if (!pressed || pressed.pointerId !== event.pointerId) return;
    stop(event);
    const origin = pressed.origin;
    const end = constrained(origin, snapped(event), event.shiftKey || pressed.shift);
    const center = fromCenterKey(event) || pressed.center;
    reset();
    // A click (no drag to speak of) places the default size at the press.
    const moved = Math.hypot(end.x - origin.x, end.y - origin.y) * scale() > 3;
    const frame: OfficeShapeDrawFrame = moved
      ? frameOf(origin, end, center)
      : {
          left: origin.x,
          top: origin.y,
          width: options.defaultSize.width,
          height: options.defaultSize.height,
          start: origin,
          end: options.line
            ? { x: origin.x + options.defaultSize.width, y: origin.y }
            : origin,
          clicked: true,
        };
    options.onComplete(frame);
  });
  layer.addEventListener("pointercancel", () => reset());
  const onKey = (event: KeyboardEvent) => {
    if (event.key === "Escape") {
      stop(event);
      reset();
      options.onCancel();
    }
  };
  doc.addEventListener("keydown", onKey, true);

  return {
    update(next) {
      options = next;
      layer.dataset.officeShapeDraw = next.line ? "line" : "box";
      paint();
    },
    destroy() {
      doc.removeEventListener("keydown", onKey, true);
      layer.remove();
    },
  };
}

