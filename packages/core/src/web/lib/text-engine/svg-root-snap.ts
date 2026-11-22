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
}

const entries = new Map<SVGSVGElement, SnapEntry>();
const dirty = new Set<SnapEntry>();
let scheduled = false;
let listening = false;
let resizeObserver: ResizeObserver | null = null;
const observed = new WeakSet<Element>();

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
function localOffset(entry: SnapEntry): [number, number] | null {
  // (No layout to read -- a test's DOM -- leaves the root as drawn.)
  const ctm = typeof entry.svg.getScreenCTM === "function" ? entry.svg.getScreenCTM() : null;
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

function fraction(value: number): number {
  const whole = Math.floor(value + EPSILON);
  const rest = value - whole;
  // To 1/10000 px: what the browser's own offsets hold.
  return rest < EPSILON ? 0 : Math.round(rest * 10_000) / 10_000;
}

function flush(secondLook = false): void {
  scheduled = false;
  if (!dirty.size) return;
  const batch = [...dirty];
  dirty.clear();
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
    const offset = localOffset(entry);
    const [x, y] = offset ? [fraction(offset[0]), fraction(offset[1])] : [0, 0];
    if (Math.abs(x - entry.applied[0]) > EPSILON || Math.abs(y - entry.applied[1]) > EPSILON) writes.push([entry, x, y]);
    const reference = entry.reference;
    if (reference && resizeObserver && !observed.has(reference)) {
      observed.add(reference);
      resizeObserver.observe(reference);
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
  if (typeof ResizeObserver !== "undefined") resizeObserver = new ResizeObserver(() => markAll(false));
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
    entry = { svg, group, applied: [0, 0], reference: undefined };
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
  flush(true);
}

/** Tests: the corrections now, not at the end of the task. */
export function flushEngineSvgRootSnapForTests(): void {
  flush(true);
}
