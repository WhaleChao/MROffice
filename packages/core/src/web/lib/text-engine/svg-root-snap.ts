/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
const EPSILON = 1e-3;

interface SnapEntry {
  readonly svg: SVGSVGElement;
  readonly group: SVGGElement;
  /** The fraction last written (x, y), local px. */
  applied: readonly [number, number];
  /** The ancestor the offset is counted from (its transform scales), or null for the page. */
  reference: Element | null | undefined;
  /** The scale on screen (x, y) its words were laid out at: read when it is first seen, and again each time it is laid out again. */
  laidOutScale: readonly [number, number] | undefined;
}

const entries = new Map<SVGSVGElement, SnapEntry>();
const dirty = new Set<SnapEntry>();
let scheduled = false;
let listening = false;
let resizeObserver: ResizeObserver | null = null;
const observed = new WeakSet<Element>();
/** Watches each reference's own `style` (its transform: the canvas's scale). */
let styleObserver: MutationObserver | null = null;
/** Roots drawn at another scale than their words were laid out at, to lay out again. */
const rescaled = new Set<SnapEntry>();
let relayoutTimer: ReturnType<typeof setTimeout> | undefined;
/** How long a new scale holds before the roots are laid out again (a window being dragged to size changes it every frame). */
const RELAYOUT_AFTER_MS = 120;

function isTranslation(transform: string): boolean {
  if (!transform || transform === "none") return true;
  const match = /^matrix\(([^)]+)\)$/.exec(transform);
  if (!match) return false; // matrix3d or anything else: a reference
  const [a, b, c, d] = match[1].split(",").map((value) => Number(value.trim()));
  return Math.abs(a - 1) < 1e-9 && Math.abs(b) < 1e-9 && Math.abs(c) < 1e-9 && Math.abs(d - 1) < 1e-9;
}

/** The nearest ancestor whose transform is more than a translation (null: the page). */
function referenceOf(svg: Element): Element | null {
  for (let element = svg.parentElement; element; element = element.parentElement) {
    if (!isTranslation(getComputedStyle(element).transform)) return element;
  }
  return null;
}

/** The root's offset in its reference's local px, as laid out (before any fraction written), or null when unknown. */
function localOffset(entry: SnapEntry, ctm: DOMMatrix | null): [number, number] | null {
  // (No layout to read -- a test's DOM -- leaves the root as drawn.)
  if (!ctm) return null;
  const [appliedX, appliedY] = entry.applied;
  const reference = entry.reference ?? null;
  if (!reference) {
    const view = entry.svg.ownerDocument.defaultView;
    return [ctm.e + (view?.scrollX ?? 0) + appliedX, ctm.f + (view?.scrollY ?? 0) + appliedY];
  }
  // A scale alone (the canvas, a thumbnail, a group scaled inside the canvas):
  // the border box's top-left is the local origin, and a local px is as wide on
  // screen as every scale above the root together -- the root's own screen
  
  // scaled 0.883 inside the canvas's 1.12 were counted in 0.883 px, their
  // fractions wrong, "citE1" still drawn a pixel off). Anything turned or
  // mirrored is left as drawn.
  const { a, b, c, d } = ctm as { a?: number; b?: number; c?: number; d?: number };
  const scaleX = typeof a === "number" ? a : new DOMMatrixReadOnly(getComputedStyle(reference).transform).a;
  const scaleY = typeof d === "number" ? d : new DOMMatrixReadOnly(getComputedStyle(reference).transform).d;
  if (Math.abs(b ?? 0) > 1e-9 || Math.abs(c ?? 0) > 1e-9 || !(scaleX > 0) || !(scaleY > 0)) return null;
  const box = reference.getBoundingClientRect();
  return [(ctm.e - box.left) / scaleX + appliedX, (ctm.f - box.top) / scaleY + appliedY];
}

/** The root's screen matrix, read once per root and pass (TRIAGE 181: a read is layout). */
function screenMatrix(svg: SVGSVGElement): DOMMatrix | null {
  return typeof svg.getScreenCTM === "function" ? svg.getScreenCTM() : null;
}

/** The root's scale on screen (its screen matrix's x and y scale), or null when unknown. */
function screenScale(svg: SVGSVGElement, read: DOMMatrix | null = screenMatrix(svg)): [number, number] | null {
  const ctm = read;
  if (!ctm) return null;
  const { a, b, c, d } = ctm as { a?: number; b?: number; c?: number; d?: number };
  if (typeof a !== "number" || typeof d !== "number") return null;
  return [Math.hypot(a, b ?? 0), Math.hypot(c ?? 0, d)];
}

function sameScale(one: readonly [number, number], other: readonly [number, number]): boolean {
  return Math.abs(one[0] - other[0]) <= 1e-4 * Math.max(1, other[0]) && Math.abs(one[1] - other[1]) <= 1e-4 * Math.max(1, other[1]);
}

/** A composition (an input method's unconfirmed letters) is under way somewhere on the page. */
let composing = false;

/**
 * Whether a root stands in the text being edited: inside the focused
 * editable host (a body's editing rows draw the engine's words over them), or
 * anywhere while a composition is under way. Moving a node in an editable
 * host while the caret is in it, or while an input method composes, may move
 * the caret or end the composition, and the letters being typed with it: such
 * a root waits until the editing leaves it.
 */
function inEditing(svg: SVGSVGElement): boolean {
  if (composing) return true;
  const active = svg.ownerDocument.activeElement as HTMLElement | null;
  if (!active || active === svg.ownerDocument.body) return false;
  if (active.contains(svg)) return true;
  const host = svg.parentElement?.closest<HTMLElement>("[contenteditable]:not([contenteditable=\"false\"])");
  return !!host && (host === active || host.contains(active));
}

/**
 * Lay the roots' words out again at the scale they are drawn at: each root
 * taken out and put back in its own place (the same node, the same
 * neighbours: nothing React holds changes), so Chromium sizes its glyphs for
 * the screen anew. A root in the text being edited keeps its place until the
 * editing leaves it (focus out, composition ended).
 */
function relayoutRescaled(): void {
  if (relayoutTimer !== undefined) {
    clearTimeout(relayoutTimer);
    relayoutTimer = undefined;
  }
  const batch = [...rescaled];
  rescaled.clear();
  const scales = batch.map((entry) => (entry.svg.isConnected ? screenScale(entry.svg) : null));
  batch.forEach((entry, index) => {
    const scale = scales[index];
    const parent = entry.svg.parentNode;
    if (!scale || !parent || entries.get(entry.svg) !== entry) return;
    if (inEditing(entry.svg)) {
      rescaled.add(entry);
      return;
    }
    parent.insertBefore(entry.svg, entry.svg.nextSibling);
    entry.laidOutScale = scale;
  });
}

function scheduleRelayout(): void {
  if (relayoutTimer !== undefined) clearTimeout(relayoutTimer);
  relayoutTimer = setTimeout(relayoutRescaled, RELAYOUT_AFTER_MS);
}

function fraction(value: number): number {
  const whole = Math.floor(value + EPSILON);
  const rest = value - whole;
  // To 1/10000 px: what the browser's own offsets hold.
  return rest < EPSILON ? 0 : Math.round(rest * 10_000) / 10_000;
}

function flush(secondLook = false, relayoutNow = false): void {
  scheduled = false;
  if (!dirty.size) return;
  const batch = [...dirty];
  dirty.clear();
  const started = typeof performance === "undefined" ? 0 : performance.now();
  // A root can still move after it was drawn without being drawn again (its
  // frame's anchoring or autofit settling in a later effect): the roots of
  
  // tp-markph2's printed bodies stood 0.2-0.35 pt off where a root had moved
  
  if (!secondLook && typeof requestAnimationFrame === "function") {
    requestAnimationFrame(() => {
      for (const entry of batch) if (entries.get(entry.svg) === entry) dirty.add(entry);
      flush(true);
    });
  }
  // Read every offset first.
  const writes: Array<[SnapEntry, number, number]> = [];
  for (const entry of batch) {
    if (!entry.svg.isConnected) continue;
    if (entry.reference === undefined) entry.reference = referenceOf(entry.svg);
    const ctm = screenMatrix(entry.svg);
    const scale = screenScale(entry.svg, ctm);
    if (scale) {
      if (!entry.laidOutScale) entry.laidOutScale = scale;
      else if (!sameScale(scale, entry.laidOutScale)) rescaled.add(entry);
      else rescaled.delete(entry);
    }
    const offset = localOffset(entry, ctm);
    const [x, y] = offset ? [fraction(offset[0]), fraction(offset[1])] : [0, 0];
    if (Math.abs(x - entry.applied[0]) > EPSILON || Math.abs(y - entry.applied[1]) > EPSILON) writes.push([entry, x, y]);
    const reference = entry.reference;
    if (reference && !observed.has(reference)) {
      observed.add(reference);
      resizeObserver?.observe(reference);
      styleObserver?.observe(reference, { attributes: true, attributeFilter: ["style"] });
    }
  }
  // Then write them.
  for (const [entry, x, y] of writes) {
    entry.applied = [x, y];
    entry.svg.style.left = x ? `${-x}px` : "0px";
    entry.svg.style.top = y ? `${-y}px` : "0px";
    if (x || y) entry.group.setAttribute("transform", `translate(${x} ${y})`);
    else entry.group.removeAttribute("transform");
  }
  noteFlush(batch.length, started);
  if (!rescaled.size) return;
  // A print reads the page now: its roots at their scale before it says it is ready.
  if (relayoutNow) relayoutRescaled();
  else scheduleRelayout();
}

/** The roots each pass read and its time (`window.__mraiSnapFlush`: passes, roots, ms, largest pass; the probes' view). */
function noteFlush(roots: number, started: number): void {
  if (typeof window === "undefined" || typeof performance === "undefined") return;
  const ms = performance.now() - started;
  const seen = ((window as unknown as { __mraiSnapFlush?: { passes: number; roots: number; ms: number; maxRoots: number; maxMs: number } }).__mraiSnapFlush ??= { passes: 0, roots: 0, ms: 0, maxRoots: 0, maxMs: 0 });
  seen.passes += 1;
  seen.roots += roots;
  seen.ms += ms;
  seen.maxRoots = Math.max(seen.maxRoots, roots);
  seen.maxMs = Math.max(seen.maxMs, ms);
}

function schedule(): void {
  if (scheduled) return;
  scheduled = true;
  // Before the frame is painted (and before a print reads the page).
  queueMicrotask(() => flush());
}

function markAll(forgetReferences: boolean): void {
  for (const entry of entries.values()) {
    if (forgetReferences) entry.reference = undefined;
    dirty.add(entry);
  }
  schedule();
}

function markReferences(references: readonly Element[]): void {
  const changed = new Set(references);
  for (const entry of entries.values()) {
    if (entry.reference && changed.has(entry.reference)) dirty.add(entry);
  }
  schedule();
}

function listen(): void {
  if (listening || typeof window === "undefined") return;
  listening = true;
  window.addEventListener("resize", () => markAll(true));
  const watchRatio = () => {
    const query = window.matchMedia?.(`(resolution: ${window.devicePixelRatio}dppx)`);
    query?.addEventListener?.("change", () => {
      markAll(true);
      watchRatio();
    }, { once: true });
  };
  watchRatio();
  // A watched canvas that changed size moves the roots counted from it, and only those: marking every root of the
  // page re-read thousands (a deck's thumbnails of 50-row tables, a cell a root) whenever the show observed its new
  
  if (typeof ResizeObserver !== "undefined") resizeObserver = new ResizeObserver((records) => markReferences(records.map((record) => record.target)));
  // Roots held back while their text was edited: laid out again once the
  // editing leaves them.
  document.addEventListener("compositionstart", () => { composing = true; }, true);
  document.addEventListener("compositionend", () => {
    composing = false;
    if (rescaled.size) scheduleRelayout();
  }, true);
  document.addEventListener("focusout", () => {
    // (A composition never outlives its editor's focus.)
    composing = false;
    if (rescaled.size) scheduleRelayout();
  }, true);
  // A reference's transform changed (the canvas scaled to a new fit): its
  // roots' places, and the scale their words are drawn at.
  if (typeof MutationObserver !== "undefined") styleObserver = new MutationObserver((records) => markReferences(records.map((record) => record.target as Element)));
  // Faces that came late move what was laid out in their stand-ins.
  document.fonts?.addEventListener?.("loadingdone", () => markAll(false));
}

/**
 * Registers an engine root and its inner group for the snap correction;
 * returns the function that forgets them. Call it whenever the root is drawn
 * (its words or its place may have changed).
 */
export function snapEngineSvgRoot(svg: SVGSVGElement, group: SVGGElement): () => void {
  listen();
  let entry = entries.get(svg);
  if (!entry || entry.group !== group) {
    entry = { svg, group, applied: [0, 0], reference: undefined, laidOutScale: undefined };
    entries.set(svg, entry);
  }
  dirty.add(entry);
  schedule();
  return () => {
    const current = entries.get(svg);
    if (current) {
      entries.delete(svg);
      dirty.delete(current);
    }
  };
}

/**
 * Every root read and corrected now, synchronously: what a print calls last,
 * after its page has stopped moving and before it says it is ready (the deck
 * print route), so no root stands where it stood when last drawn.
 */
export function settleEngineSvgRoots(): void {
  for (const entry of entries.values()) dirty.add(entry);
  flush(true, true);
  if (rescaled.size) relayoutRescaled();
}

/**
 * The scale the engine roots are drawn at may have changed without any of
 * them being drawn again (a deck window's slide fitted to a new size, the
 * view zoom): read them all again, and lay out again those whose scale moved.
 */
export function noteEngineSvgRootsRescaled(): void {
  if (!entries.size) return;
  markAll(true);
}

/** Tests: the corrections now, not at the end of the task. */
export function flushEngineSvgRootSnapForTests(): void {
  flush(true);
}

/** Tests: the roots whose scale moved laid out again now, not after the scale has held. */
export function relayoutRescaledEngineSvgRootsForTests(): void {
  relayoutRescaled();
}
