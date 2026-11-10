/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Pure accessibility helpers for the presentation slide editor: keyboard nudge
 * math + block accessible-name composition.
 *
 * Dependency-free (no React / no DOM types other than the structural shapes it
 * is handed) so the logic can be unit-tested in isolation and reused by both the
 * editor component and tests. The EMU coordinate space matches the rest of the
 * presentation stack (`left`/`top`/`width`/`height` on a block).
 *
 * Mirrors the conventions of the sibling pure modules (`presentation-drag.ts`,
 * `presentation-shapes.ts`): frozen constants, small composable functions, and a
 * single deterministic source of truth the component + tests share.
 */

import { OFFICE_NUDGE_STEP_PX } from "@mrnote/core/web/lib/office-object-nudge.ts";
import { canEditBlockContent } from "@mrnote/core/web/lib/presentation-text-case.ts";

/** EMU per document pixel (96 dpi). */
export const EMU_PER_DOCUMENT_PX = 9_525;

/**
 * Keyboard nudge step (EMU) for a single arrow key press: the shared Office
 * step (`lib/office-object-nudge`, 1/12 in) the three editors use, so an
 * exact integer and persisted geometry stays clean. The slide canvas reads
 * the whole key through `officeObjectKeyAction` (Ctrl = one screen pixel,
 * Shift = resize, Alt = rotate), as PowerPoint does.
 */
export const NUDGE_STEP_EMU = OFFICE_NUDGE_STEP_PX * EMU_PER_DOCUMENT_PX;

/** A coarse ~10x step, kept for callers that still ask for one. */
export const NUDGE_STEP_LARGE_EMU = NUDGE_STEP_EMU * 10;

/** The four arrow keys this module recognizes for nudging. */
export type NudgeArrowKey =
  "ArrowUp" | "ArrowDown" | "ArrowLeft" | "ArrowRight";

const ARROW_KEYS: ReadonlySet<string> = new Set<NudgeArrowKey>([
  "ArrowUp",
  "ArrowDown",
  "ArrowLeft",
  "ArrowRight",
]);

/** Type guard: is `key` one of the four nudge arrow keys? */
export function isNudgeArrowKey(key: string): key is NudgeArrowKey {
  return ARROW_KEYS.has(key);
}

/**
 * Pick the nudge step for a keydown: the LARGE step when Shift is held, else the
 * normal step. Pure — reads only the boolean flag.
 */
export function nudgeStepForShift(shiftKey: boolean): number {
  return shiftKey ? NUDGE_STEP_LARGE_EMU : NUDGE_STEP_EMU;
}

/**
 * Translate an arrow key + step into an EMU `(dx, dy)` move delta, or `null` for
 * a non-arrow key. Up/Down move the y axis, Left/Right the x axis; the magnitude
 * is always the given `step`. Deterministic, never NaN (a non-finite step is
 * coerced to 0).
 */
export function arrowNudgeDelta(
  key: string,
  step: number,
): { dx: number; dy: number } | null {
  const magnitude = Number.isFinite(step) ? step : 0;
  switch (key) {
    case "ArrowUp":
      return { dx: 0, dy: -magnitude };
    case "ArrowDown":
      return { dx: 0, dy: magnitude };
    case "ArrowLeft":
      return { dx: -magnitude, dy: 0 };
    case "ArrowRight":
      return { dx: magnitude, dy: 0 };
    default:
      return null;
  }
}

/**
 * Whether the given (focused) element is a live text-editing surface where arrow
 * keys must move the caret/selection, NOT the shape. True for `<input>`,
 * `<textarea>`, `<select>`, and any `contenteditable` host. Used as the
 * edit-mode guard so keyboard block-movement never hijacks typing.
 *
 * Accepts the loose `EventTarget | null` shape so a caller can pass
 * `document.activeElement` or `event.target` directly.
 */
export function isTextEditTarget(target: EventTarget | null): boolean {
  if (!target || typeof target !== "object") return false;
  const el = target as {
    tagName?: unknown;
    isContentEditable?: unknown;
  };
  const tag = typeof el.tagName === "string" ? el.tagName.toUpperCase() : "";
  if (tag === "INPUT" || tag === "TEXTAREA" || tag === "SELECT") return true;
  if (el.isContentEditable === true) return true;
  return false;
}

/**
 * Collapse whitespace and clip a candidate accessible-name string to a sane
 * length so a long paragraph doesn't produce a multi-kilobyte aria-label. Pure;
 * returns "" for non-strings / empties.
 */
export function shortAccessibleName(value: unknown, max = 80): string {
  if (typeof value !== "string") return "";
  const collapsed = value.replace(/\s+/g, " ").trim();
  if (collapsed.length <= max) return collapsed;
  return `${collapsed.slice(0, Math.max(0, max - 1)).trimEnd()}…`;
}

/**
 * The structural slice of a presentation block this module reads to build an
 * accessible name. A superset-compatible subset of `PresentationBlock` so the
 * helper stays decoupled from the full editor type (and trivially testable).
 */
export interface A11yBlockLike {
  type: string;
  kind?: unknown;
  text?: unknown;
  alt?: unknown;
  decorative?: unknown;
  caption?: unknown;
  title?: unknown;
  shape?: unknown;
  url?: unknown;
}

/**
 * Resolved, locale-aware labels the name builder composes with. The caller (the
 * editor, holding `useTranslations`) resolves these once and hands them in, so
 * this module stays free of any i18n dependency while still producing localized
 * names.
 *
 * `blockKind(type)` → e.g. "Image block"; `shapeKind(shape)` → e.g. "Rectangle".
 */
export interface A11yBlockLabels {
  blockKind: (type: string) => string;
  shapeKind: (shape: string) => string;
}

/**
 * Compose the accessible NAME for a block, used as its `aria-label` on the
 * canvas. Reading-order-stable and never empty:
 *  - text  → the block's text (clipped), else its type label.
 *  - image → decorative pictures expose only their object-kind label; otherwise
 *    their alt text (clipped), then caption, then type label.
 *  - shape → its shape-kind label (e.g. "Rectangle"), prefixed by the kind label.
 *  - chart → its title (clipped), else its type label.
 *  - table → its caption (clipped), else its type label.
 *  - embed → its title (clipped), else its type label.
 * Always falls back to the localized block-kind label so the name is non-empty
 * even for an empty/blank block. Pure: no DOM, no Date/random.
 */
export function blockAccessibleName(
  block: A11yBlockLike,
  labels: A11yBlockLabels,
): string {
  const kindLabel = labels.blockKind(block.type);
  switch (block.type) {
    case "image": {
      if (block.decorative === true) return kindLabel;
      const alt = shortAccessibleName(block.alt);
      if (alt) return alt;
      const caption = shortAccessibleName(block.caption);
      if (caption) return caption;
      return kindLabel;
    }
    case "shape": {
      const shape =
        typeof block.shape === "string" && block.shape.trim()
          ? block.shape.trim()
          : "rectangle";
      const shapeLabel = labels.shapeKind(shape);
      // e.g. "Shape block: Rectangle" — the kind first (so a screen reader
      // announces it's a shape) then the specific geometry.
      return shapeLabel ? `${kindLabel}: ${shapeLabel}` : kindLabel;
    }
    case "chart": {
      const title = shortAccessibleName(block.title);
      return title || kindLabel;
    }
    case "table": {
      const caption = shortAccessibleName(block.caption);
      return caption || kindLabel;
    }
    case "embed": {
      const title = shortAccessibleName(block.title);
      return title || kindLabel;
    }
    case "text":
    default: {
      const text = shortAccessibleName(block.text);
      return text || kindLabel;
    }
  }
}

/**
 * The set of routine editor operations the presentation slide editor announces
 * to assistive technology via its visually-hidden `aria-live="polite"` status
 * region. Each variant carries only the data needed to format a concise,
 * localized sentence — the live-region effect maps the editor's reactive state
 * (active slide index, selected block) onto one of these and feeds it to
 * `formatLiveAnnouncement`.
 *
 *  - `slide-changed`   → focus moved to a different slide ("Slide N of M[: title]").
 *  - `object-selected` → a shape/object became the selection (its accessible name).
 *  - `object-added`    → a block was inserted (its accessible name).
 *  - `object-deleted`  → a block was removed (its accessible name).
 *  - `present-entered` / `present-exited` → present (full-screen) mode toggled.
 */
export type LiveAnnouncementEvent =
  | {
      kind: "slide-changed";
      /** 0-based active slide index; rendered 1-based. */
      index: number;
      /** Total slide count. */
      total: number;
      /** The slide's title, if any (clipped/blank → omitted from the message). */
      title?: string;
    }
  | { kind: "object-selected"; label: string }
  | { kind: "object-added"; label: string }
  | { kind: "object-deleted"; label: string }
  | { kind: "present-entered" }
  | { kind: "present-exited" };

/**
 * The minimal translator shape `formatLiveAnnouncement` needs: a `next-intl`
 * `t(key, params?)`. Declaring it structurally keeps this module free of any
 * i18n dependency (and trivially testable with an echo translator) while the
 * editor passes its real `useTranslations("console-notebooks")` instance.
 */
export type LiveAnnouncementTranslate = (
  key: string,
  params?: Record<string, string | number>,
) => string;

/** i18n keys for each live-announcement variant (single source of truth). */
export const LIVE_ANNOUNCEMENT_KEYS = {
  slideChanged: "presentation.liveSlideChanged",
  slideChangedTitled: "presentation.liveSlideChangedTitled",
  objectSelected: "presentation.liveObjectSelected",
  objectAdded: "presentation.liveObjectAdded",
  objectDeleted: "presentation.liveObjectDeleted",
  presentEntered: "presentation.livePresentEntered",
  presentExited: "presentation.livePresentExited",
} as const;

/**
 * Format a single live-region announcement into a localized, screen-reader
 * sentence. Pure: no DOM, no Date/random — the editor resolves copy via the
 * injected translator so the same logic is exercised verbatim in unit tests
 * (with an echo translator) and at runtime (with `next-intl`).
 *
 * Slide changes prefer the titled template when a non-blank title is present,
 * falling back to the bare "Slide N of M" form otherwise. Object events clip
 * their label through {@link shortAccessibleName} so a long paragraph can't
 * produce a multi-kilobyte announcement.
 */
export function formatLiveAnnouncement(
  event: LiveAnnouncementEvent,
  translate: LiveAnnouncementTranslate,
): string {
  switch (event.kind) {
    case "slide-changed": {
      const current = event.index + 1;
      const total = event.total;
      const title = shortAccessibleName(event.title);
      if (title) {
        return translate(LIVE_ANNOUNCEMENT_KEYS.slideChangedTitled, {
          current,
          total,
          title,
        });
      }
      return translate(LIVE_ANNOUNCEMENT_KEYS.slideChanged, { current, total });
    }
    case "object-selected":
      return translate(LIVE_ANNOUNCEMENT_KEYS.objectSelected, {
        label: shortAccessibleName(event.label),
      });
    case "object-added":
      return translate(LIVE_ANNOUNCEMENT_KEYS.objectAdded, {
        label: shortAccessibleName(event.label),
      });
    case "object-deleted":
      return translate(LIVE_ANNOUNCEMENT_KEYS.objectDeleted, {
        label: shortAccessibleName(event.label),
      });
    case "present-entered":
      return translate(LIVE_ANNOUNCEMENT_KEYS.presentEntered);
    case "present-exited":
      return translate(LIVE_ANNOUNCEMENT_KEYS.presentExited);
    default: {
      // Exhaustiveness guard: a new variant must extend the switch above.
      const _never: never = event;
      return _never;
    }
  }
}

/* -------------------------------------------------------------------------- *
 * Accessibility Checker projection
 * -------------------------------------------------------------------------- */

/** PowerPoint's pane must stay useful without allowing an unbounded deck scan. */
export const MAX_PRESENTATION_ACCESSIBILITY_ISSUES = 500;

/**
 * Hard synchronous work budget for one checker snapshot. Every structural pass
 * over a slide, block, table row/cell, overlap candidate or reading-order item
 * consumes from this single budget. Exhaustion invalidates the partial snapshot
 * instead of exposing findings derived from only a prefix of the deck.
 */
export const MAX_PRESENTATION_ACCESSIBILITY_SCAN_WORK = 20_000;

export type PresentationAccessibilityIssueKind =
  | "imageMissingAlt"
  | "slideTitle"
  | "tableStructure"
  | "readingOrder"
  | "colorContrast"
  | "mediaCaptions";

export type PresentationAccessibilityIssueSeverity =
  | "error"
  | "warning"
  | "tip";

export type PresentationAccessibilityIssueReason =
  | "missingAltText"
  | "missingSlideTitle"
  | "duplicateSlideTitle"
  | "missingTableHeader"
  | "mergedTableCells"
  | "spatialReadingOrder"
  | "lowColorContrast"
  | "mediaCaptionsReview";

export type PresentationAccessibilityIssueAction =
  | "edit-slide-title"
  | "fix-table-header"
  | "navigate"
  | "edit-image-alt"
  | "open-selection-pane";

/**
 * Honest UI remediation exposed for one checker finding. `edit` opens an
 * existing editor and `manage` opens the existing ordering surface; neither is
 * an automatic repair. Findings outside these authenticated routes remain
 * manual review even though the generic Go To action can still locate them.
 */
export type PresentationAccessibilityRemediation =
  | { readonly kind: "locate"; readonly target: "slide-title" }
  | { readonly kind: "edit"; readonly target: "image-alt" }
  | { readonly kind: "fix"; readonly target: "table-header" }
  | { readonly kind: "manage"; readonly target: "reading-order" }
  | { readonly kind: "manual-review" };

/**
 * Structural deck slice read by the checker. It deliberately accepts the
 * Presentation model without importing its editor-owned union, keeping this
 * module dependency-free and making every unsupported/opaque field skippable.
 */
export interface PresentationAccessibilityBlockLike {
  readonly id?: unknown;
  readonly type?: unknown;
  readonly kind?: unknown;
  readonly text?: unknown;
  readonly title?: unknown;
  readonly caption?: unknown;
  readonly alt?: unknown;
  readonly decorative?: unknown;
  readonly source?: unknown;
  readonly rows?: unknown;
  readonly headerRow?: unknown;
  readonly role?: unknown;
  readonly placeholder?: unknown;
  readonly style?: unknown;
  readonly runs?: unknown;
  readonly paragraphs?: unknown;
  readonly fill?: unknown;
  readonly fillTo?: unknown;
  readonly fillImage?: unknown;
  readonly opacity?: unknown;
  readonly left?: unknown;
  readonly top?: unknown;
  readonly width?: unknown;
  readonly height?: unknown;
  readonly rotation?: unknown;
  readonly hidden?: unknown;
  readonly locked?: unknown;
}

export interface PresentationAccessibilitySlideLike {
  readonly id?: unknown;
  readonly title?: unknown;
  readonly blocks?: readonly PresentationAccessibilityBlockLike[];
  readonly backgroundColor?: unknown;
  readonly backgroundGradient?: unknown;
  readonly backgroundImage?: unknown;
  readonly backgroundPattern?: unknown;
}

export interface PresentationAccessibilityIssue {
  /** Snapshot-local id. Full authority lives in the complete signature below. */
  readonly id: string;
  readonly kind: PresentationAccessibilityIssueKind;
  readonly severity: PresentationAccessibilityIssueSeverity;
  readonly reason: PresentationAccessibilityIssueReason;
  readonly action: PresentationAccessibilityIssueAction;
  readonly slideId: string;
  readonly slideIndex: number;
  readonly blockId: string | null;
  readonly blockType: string | null;
  /** Bounded plain text rendered only as a React child. */
  readonly label: string;
  /** Complete deterministic rule context used for stale reauthentication. */
  readonly context: string;
}

/** Derive the only contextual action the checker may truthfully advertise. */
export function presentationAccessibilityRemediation(
  issue: PresentationAccessibilityIssue,
): PresentationAccessibilityRemediation {
  switch (issue.action) {
    case "edit-slide-title":
      return { kind: "locate", target: "slide-title" };
    case "fix-table-header":
      return { kind: "fix", target: "table-header" };
    case "edit-image-alt":
      return { kind: "edit", target: "image-alt" };
    case "open-selection-pane":
      return { kind: "manage", target: "reading-order" };
    default:
      return { kind: "manual-review" };
  }
}

type A11yRecord = Record<string, unknown>;

interface AccessibilityScanBudget {
  remaining: number;
  exhausted: boolean;
}

function consumeAccessibilityScanWork(
  budget: AccessibilityScanBudget,
  units = 1,
): boolean {
  if (!Number.isSafeInteger(units) || units < 0 || units > budget.remaining) {
    budget.remaining = 0;
    budget.exhausted = true;
    return false;
  }
  budget.remaining -= units;
  return true;
}

function a11yRecord(value: unknown): A11yRecord | null {
  return value !== null && typeof value === "object" && !Array.isArray(value)
    ? (value as A11yRecord)
    : null;
}

function stableIdentity(value: unknown): string | null {
  if (typeof value !== "string") return null;
  if (value.length > 128) return null;
  const identity = value.trim();
  return identity || null;
}

function boundedContext(value: string): string | null {
  return value.length <= 8_192 ? value : null;
}

function normalizedTitle(value: unknown): string | null {
  if (value !== undefined && value !== null && typeof value !== "string") {
    return null;
  }
  if (typeof value !== "string") return "";
  if (value.length > 2_000) return null;
  const title = value.replace(/\s+/g, " ").trim();
  return title.length <= 500 ? title.toLocaleLowerCase("en-US") : null;
}

/**
 * The slide's title as PowerPoint reads it: the words in its title
 * placeholder. A slide that has one is titled by it alone (emptying the
 * placeholder leaves the slide untitled whatever name it was given); only a
 * slide without a title placeholder falls back to its stored name. `blockId`
 * is the placeholder's, when it has a usable one, so a finding can lead to
 * the place the title is typed.
 */
function slideTitleReading(
  slide: PresentationAccessibilitySlideLike,
  budget: AccessibilityScanBudget,
): { readonly title: unknown; readonly blockId: string | null } | null {
  const blocks = Array.isArray(slide.blocks) ? slide.blocks : [];
  for (const block of blocks) {
    if (!consumeAccessibilityScanWork(budget)) return null;
    if (!a11yRecord(block) || block.type !== "text" || block.hidden === true) {
      continue;
    }
    if (
      block.role !== "title" &&
      block.role !== "ctrTitle" &&
      block.placeholder !== "title" &&
      block.placeholder !== "sectionTitle"
    ) {
      continue;
    }
    let text: unknown = block.text;
    if (Array.isArray(block.paragraphs) && block.paragraphs.length > 0) {
      const lines: string[] = [];
      for (const paragraph of block.paragraphs) {
        if (!consumeAccessibilityScanWork(budget)) return null;
        const line = a11yRecord(paragraph)?.text;
        if (typeof line === "string") lines.push(line);
      }
      text = lines.join(" ");
    }
    const blockId = stableIdentity(block.id);
    const unique =
      blockId !== null &&
      blocks.filter((candidate) => candidate.id === block.id).length === 1;
    return { title: text ?? "", blockId: unique ? blockId : null };
  }
  return { title: slide.title, blockId: null };
}

function issueId(
  kind: PresentationAccessibilityIssueKind,
  slideId: string,
  blockId: string | null,
  reason: PresentationAccessibilityIssueReason,
  discriminator = "",
): string {
  return JSON.stringify([kind, slideId, blockId, reason, discriminator]);
}

function shortScannerAccessibleName(value: unknown, max = 80): string {
  if (typeof value === "string" && value.length > 8_192) return "…";
  return shortAccessibleName(value, max);
}

function issueLabel(block: PresentationAccessibilityBlockLike): string {
  return (
    shortScannerAccessibleName(block.alt) ||
    shortScannerAccessibleName(block.title) ||
    shortScannerAccessibleName(block.caption) ||
    shortScannerAccessibleName(block.text) ||
    shortScannerAccessibleName(block.type) ||
    "Object"
  );
}

/**
 * The words of the deck that name a finding's object, or "" when the label
 * is only the scanner's stand-in (the object's type, "Object", "Slide N").
 * A stand-in is not shown: the pane states the slide and the rule in the
 * interface language, and an English type id beside them says nothing more.
 */
export function presentationAccessibilityIssueObjectName(
  issue: Pick<
    PresentationAccessibilityIssue,
    "label" | "blockType" | "slideIndex"
  >,
): string {
  const label = issue.label;
  if (
    label === "Object" ||
    label === `Slide ${issue.slideIndex + 1}` ||
    (issue.blockType !== null && label === issue.blockType)
  ) {
    return "";
  }
  return label;
}

function appendPresentationIssue(
  issues: PresentationAccessibilityIssue[],
  issue: PresentationAccessibilityIssue,
): void {
  if (issues.length < MAX_PRESENTATION_ACCESSIBILITY_ISSUES) {
    issues.push(Object.freeze(issue));
  }
}

function countIdentities<T>(
  values: readonly T[],
  identityOf: (value: T) => unknown,
  budget: AccessibilityScanBudget,
): Map<string, number> | null {
  const counts = new Map<string, number>();
  for (let index = 0; index < values.length; index += 1) {
    if (!consumeAccessibilityScanWork(budget)) return null;
    const identity = stableIdentity(identityOf(values[index]));
    if (!identity) continue;
    counts.set(identity, (counts.get(identity) ?? 0) + 1);
  }
  return counts;
}

function hasDeterministicGeometry(
  block: PresentationAccessibilityBlockLike,
): block is PresentationAccessibilityBlockLike & {
  readonly left: number;
  readonly top: number;
  readonly width: number;
  readonly height: number;
} {
  return (
    typeof block.left === "number" &&
    Number.isFinite(block.left) &&
    typeof block.top === "number" &&
    Number.isFinite(block.top) &&
    typeof block.width === "number" &&
    Number.isFinite(block.width) &&
    block.width > 0 &&
    typeof block.height === "number" &&
    Number.isFinite(block.height) &&
    block.height > 0
  );
}

type DeterministicGeometryBlock = PresentationAccessibilityBlockLike & {
  readonly left: number;
  readonly top: number;
  readonly width: number;
  readonly height: number;
};

function rectanglesOverlap(
  left: DeterministicGeometryBlock,
  right: DeterministicGeometryBlock,
): boolean {
  return (
    left.left < right.left + right.width &&
    right.left < left.left + left.width &&
    left.top < right.top + right.height &&
    right.top < left.top + left.height
  );
}

function hasAuthenticatedZeroRotation(
  block: PresentationAccessibilityBlockLike,
): boolean {
  return block.rotation === undefined || block.rotation === 0;
}

function tableRows(
  value: unknown,
  budget: AccessibilityScanBudget,
): readonly (readonly A11yRecord[])[] | null {
  if (!Array.isArray(value) || value.length === 0) return null;
  const rows: A11yRecord[][] = [];
  // The last row index a rowspan from above still covers: a row with no cell
  // of its own under such a span is a legal logical row (a 2x2 table merged
  // vertically keeps two rows, the second empty), not an invalid table that
  // escapes the whole scan (PPT-B374).
  let coveredThrough = -1;
  for (let rowIndex = 0; rowIndex < value.length; rowIndex += 1) {
    if (!consumeAccessibilityScanWork(budget)) return null;
    const row = value[rowIndex];
    if (!Array.isArray(row)) return null;
    if (row.length === 0 && rowIndex > coveredThrough) return null;
    const cells: A11yRecord[] = [];
    for (let columnIndex = 0; columnIndex < row.length; columnIndex += 1) {
      if (!consumeAccessibilityScanWork(budget)) return null;
      const cell = row[columnIndex];
      const record = a11yRecord(cell);
      if (!record) return null;
      const rowspan = record.rowspan;
      if (typeof rowspan === "number" && Number.isInteger(rowspan) && rowspan > 1) {
        coveredThrough = Math.max(coveredThrough, rowIndex + rowspan - 1);
      }
      cells.push(record);
    }
    rows.push(cells);
  }
  return rows;
}

function tableHasMergedCells(
  rows: readonly (readonly A11yRecord[])[],
  budget: AccessibilityScanBudget,
): boolean | null {
  for (let rowIndex = 0; rowIndex < rows.length; rowIndex += 1) {
    if (!consumeAccessibilityScanWork(budget)) return null;
    const row = rows[rowIndex];
    for (let columnIndex = 0; columnIndex < row.length; columnIndex += 1) {
      if (!consumeAccessibilityScanWork(budget)) return null;
      const cell = row[columnIndex];
      const colspan = cell.colspan;
      const rowspan = cell.rowspan;
      if (
        (typeof colspan === "number" &&
          Number.isInteger(colspan) &&
          colspan > 1) ||
        (typeof rowspan === "number" &&
          Number.isInteger(rowspan) &&
          rowspan > 1)
      ) {
        return true;
      }
    }
  }
  return false;
}

const HEX_COLOR = /^#[0-9a-fA-F]{6}$/;

function parseOpaqueHex(value: unknown): [number, number, number] | null {
  if (typeof value !== "string" || !HEX_COLOR.test(value)) return null;
  return [
    Number.parseInt(value.slice(1, 3), 16),
    Number.parseInt(value.slice(3, 5), 16),
    Number.parseInt(value.slice(5, 7), 16),
  ];
}

function relativeLuminance(rgb: readonly [number, number, number]): number {
  const channels = rgb.map((channel) => {
    const value = channel / 255;
    return value <= 0.04045
      ? value / 12.92
      : ((value + 0.055) / 1.055) ** 2.4;
  }) as [number, number, number];
  const [red, green, blue] = channels;
  return 0.2126 * red + 0.7152 * green + 0.0722 * blue;
}

function contrastRatio(
  foreground: readonly [number, number, number],
  background: readonly [number, number, number],
): number {
  const foregroundLuminance = relativeLuminance(foreground);
  const backgroundLuminance = relativeLuminance(background);
  const lighter = Math.max(foregroundLuminance, backgroundLuminance);
  const darker = Math.min(foregroundLuminance, backgroundLuminance);
  return (lighter + 0.05) / (darker + 0.05);
}

/** One stretch of an object's text with the paint in force on its glyphs. */
interface ContrastSpan {
  readonly color: unknown;
  readonly highlight: unknown;
  readonly fontSize: unknown;
  readonly bold: unknown;
}

interface ContrastFinding {
  readonly ratio: number;
  readonly required: number;
  readonly foreground: string;
  readonly background: string;
}

/**
 * The stretches of text an object draws, each with its effective paint: a
 * character run's own colour, size and weight stand over the object's
 * (PowerPoint's Font Color on a selection writes the runs and leaves the box's
 * default as it was, so the default alone is not what is on screen). An
 * object without runs is one stretch in its default paint. Returns null when
 * the scan budget ran out.
 */
function contrastSpans(
  owner: { readonly runs?: unknown; readonly paragraphs?: unknown },
  base: ContrastSpan,
  budget: AccessibilityScanBudget,
): ContrastSpan[] | null {
  const runLists: unknown[] = [];
  let plainText = false;
  if (Array.isArray(owner.paragraphs) && owner.paragraphs.length > 0) {
    for (const paragraph of owner.paragraphs) {
      if (!consumeAccessibilityScanWork(budget)) return null;
      const record = a11yRecord(paragraph);
      if (!record) continue;
      if (Array.isArray(record.runs) && record.runs.length > 0) {
        runLists.push(record.runs);
      } else if (typeof record.text === "string" && record.text.trim()) {
        plainText = true;
      }
    }
  } else if (Array.isArray(owner.runs) && owner.runs.length > 0) {
    runLists.push(owner.runs);
  } else {
    plainText = true;
  }
  const spans: ContrastSpan[] = plainText ? [base] : [];
  for (const runs of runLists) {
    for (const run of runs as readonly unknown[]) {
      if (!consumeAccessibilityScanWork(budget)) return null;
      const record = a11yRecord(run);
      if (!record || typeof record.text !== "string" || !record.text.trim()) {
        continue;
      }
      const style = a11yRecord(record.style);
      spans.push({
        color: style?.color ?? base.color,
        highlight: style?.highlight ?? base.highlight,
        fontSize: style?.fontSize ?? base.fontSize,
        bold: style?.bold ?? base.bold,
      });
    }
  }
  return spans;
}

/**
 * The stretch that misses its required contrast by the most, or null when
 * every stretch that can be judged passes. A stretch whose colour or
 * background is not one opaque colour is not judged.
 */
function worstContrast(
  spans: readonly ContrastSpan[],
  background: unknown,
): ContrastFinding | null {
  let worst: ContrastFinding | null = null;
  for (const span of spans) {
    const behind = span.highlight !== undefined ? span.highlight : background;
    const foreground = parseOpaqueHex(span.color);
    const backdrop = parseOpaqueHex(behind);
    if (!foreground || !backdrop) continue;
    const ratio = contrastRatio(foreground, backdrop);
    const fontSize =
      typeof span.fontSize === "number" ? span.fontSize : Number.NaN;
    const largeText =
      Number.isFinite(fontSize) &&
      (fontSize >= 18 || (fontSize >= 14 && span.bold === true));
    const required = largeText ? 3 : 4.5;
    if (ratio >= required) continue;
    if (!worst || ratio / required < worst.ratio / worst.required) {
      worst = {
        ratio,
        required,
        foreground: String(span.color).toLowerCase(),
        background: String(behind).toLowerCase(),
      };
    }
  }
  return worst;
}

function explicitSlideBackground(
  slide: PresentationAccessibilitySlideLike,
): [number, number, number] | null {
  if (
    slide.backgroundGradient !== undefined ||
    slide.backgroundImage !== undefined ||
    slide.backgroundPattern !== undefined
  ) {
    return null;
  }
  return parseOpaqueHex(slide.backgroundColor);
}

function explicitBlockBackground(
  block: PresentationAccessibilityBlockLike,
  slideBackground: [number, number, number] | null,
  allowSlideFallback: boolean,
): [number, number, number] | null {
  if (
    block.fillTo !== undefined ||
    block.fillImage !== undefined ||
    (block.opacity !== undefined && block.opacity !== 100)
  ) {
    return null;
  }
  const blockFill = parseOpaqueHex(block.fill);
  if (block.fill !== undefined && !blockFill) return null;
  return (
    blockFill ?? (allowSlideFallback ? slideBackground : null)
  );
}

function hasUnambiguousSlideBackground(
  block: PresentationAccessibilityBlockLike,
  blockIndex: number,
  blocks: readonly PresentationAccessibilityBlockLike[],
  budget: AccessibilityScanBudget,
): boolean {
  if (
    !hasDeterministicGeometry(block) ||
    !hasAuthenticatedZeroRotation(block)
  ) {
    return false;
  }
  for (let siblingIndex = 0; siblingIndex < blocks.length; siblingIndex += 1) {
    if (!consumeAccessibilityScanWork(budget)) return false;
    if (siblingIndex === blockIndex) continue;
    const sibling = blocks[siblingIndex];
    if (sibling.hidden === true) continue;
    if (
      (sibling.hidden !== undefined && typeof sibling.hidden !== "boolean") ||
      !hasAuthenticatedZeroRotation(sibling) ||
      !hasDeterministicGeometry(sibling) ||
      rectanglesOverlap(block, sibling)
    ) {
      return false;
    }
  }
  return true;
}

function appendBlockContrastIssue(
  issues: PresentationAccessibilityIssue[],
  slide: PresentationAccessibilitySlideLike,
  slideId: string,
  slideIndex: number,
  block: PresentationAccessibilityBlockLike,
  blockIndex: number,
  blocks: readonly PresentationAccessibilityBlockLike[],
  blockId: string,
  blockType: string,
  budget: AccessibilityScanBudget,
): void {
  if (blockType !== "text" && blockType !== "shape") return;
  const style = a11yRecord(block.style);
  if (!style) return;
  const slideBackground = explicitSlideBackground(slide);
  const hasExplicitBlockFill = parseOpaqueHex(block.fill) !== null;
  const allowSlideFallback =
    blockType === "text" &&
    slideBackground !== null &&
    !hasExplicitBlockFill &&
    hasUnambiguousSlideBackground(block, blockIndex, blocks, budget);
  if (budget.exhausted) return;
  const background = explicitBlockBackground(
    block,
    slideBackground,
    allowSlideFallback,
  );
  const spans = contrastSpans(
    block,
    {
      color: style.color,
      highlight: undefined,
      fontSize: style.fontSize,
      bold: style.bold,
    },
    budget,
  );
  if (!spans) return;
  const backgroundHex = !background
    ? undefined
    : typeof block.fill === "string"
      ? block.fill
      : slide.backgroundColor;
  const finding = worstContrast(spans, backgroundHex);
  if (!finding) return;
  const context = `foreground=${finding.foreground};background=${finding.background};ratio=${finding.ratio.toFixed(2)};required=${finding.required.toFixed(1)}`;
  appendPresentationIssue(issues, {
    id: issueId(
      "colorContrast",
      slideId,
      blockId,
      "lowColorContrast",
      "block",
    ),
    kind: "colorContrast",
    severity: "warning",
    reason: "lowColorContrast",
    action: "navigate",
    slideId,
    slideIndex,
    blockId,
    blockType,
    label: issueLabel(block),
    context,
  });
}

function appendTableContrastIssues(
  issues: PresentationAccessibilityIssue[],
  slideId: string,
  slideIndex: number,
  block: PresentationAccessibilityBlockLike,
  blockId: string,
  rows: readonly (readonly A11yRecord[])[],
  budget: AccessibilityScanBudget,
): void {
  for (let rowIndex = 0; rowIndex < rows.length; rowIndex += 1) {
    if (!consumeAccessibilityScanWork(budget)) return;
    const row = rows[rowIndex];
    for (let columnIndex = 0; columnIndex < row.length; columnIndex += 1) {
      if (
        issues.length >= MAX_PRESENTATION_ACCESSIBILITY_ISSUES ||
        !consumeAccessibilityScanWork(budget)
      ) {
        return;
      }
      const cell = row[columnIndex];
      const spans = contrastSpans(
        cell,
        {
          color: cell.color,
          highlight: undefined,
          fontSize: cell.fontSize,
          bold: cell.bold,
        },
        budget,
      );
      if (!spans) return;
      const finding = worstContrast(spans, cell.fill);
      if (!finding) continue;
      const discriminator = `cell-${rowIndex}-${columnIndex}`;
      appendPresentationIssue(issues, {
        id: issueId(
          "colorContrast",
          slideId,
          blockId,
          "lowColorContrast",
          discriminator,
        ),
        kind: "colorContrast",
        severity: "warning",
        reason: "lowColorContrast",
        action: "navigate",
        slideId,
        slideIndex,
        blockId,
        blockType: "table",
        label: issueLabel(block),
        context: `${discriminator};foreground=${finding.foreground};background=${finding.background};ratio=${finding.ratio.toFixed(2)};required=${finding.required.toFixed(1)}`,
      });
    }
  }
}

/**
 * Derive a strict, session-only Presentation accessibility projection. The
 * scanner reports only evidence represented unambiguously by the current
 * model; unsupported native metadata, duplicate identities, rich fills and
 * indeterminate geometry are skipped rather than guessed.
 */
export function collectPresentationAccessibilityIssues(
  slides: readonly PresentationAccessibilitySlideLike[],
): PresentationAccessibilityIssue[] {
  const budget: AccessibilityScanBudget = {
    remaining: MAX_PRESENTATION_ACCESSIBILITY_SCAN_WORK,
    exhausted: false,
  };
  const issues: PresentationAccessibilityIssue[] = [];
  const slideIdCounts = countIdentities(slides, (slide) => slide.id, budget);
  if (!slideIdCounts) return [];
  const titleCounts = new Map<string, number>();
  for (let slideIndex = 0; slideIndex < slides.length; slideIndex += 1) {
    if (!consumeAccessibilityScanWork(budget)) return [];
    const slide = slides[slideIndex];
    const slideId = stableIdentity(slide.id);
    if (!slideId || slideIdCounts.get(slideId) !== 1) continue;
    const reading = slideTitleReading(slide, budget);
    if (!reading) return [];
    const title = normalizedTitle(reading.title);
    if (title) titleCounts.set(title, (titleCounts.get(title) ?? 0) + 1);
  }

  for (let slideIndex = 0; slideIndex < slides.length; slideIndex += 1) {
    if (issues.length >= MAX_PRESENTATION_ACCESSIBILITY_ISSUES) break;
    if (!consumeAccessibilityScanWork(budget)) return [];
    const slide = slides[slideIndex];
    const slideId = stableIdentity(slide.id);
    if (!slideId || slideIdCounts.get(slideId) !== 1) continue;
    const titleReading = slideTitleReading(slide, budget);
    if (!titleReading) return [];
    const title = normalizedTitle(titleReading.title);
    const titleBlock = titleReading.blockId
      ? { blockId: titleReading.blockId, blockType: "text" }
      : { blockId: null, blockType: null };
    if (title === "") {
      appendPresentationIssue(issues, {
        id: issueId(
          "slideTitle",
          slideId,
          titleReading.blockId,
          "missingSlideTitle",
        ),
        kind: "slideTitle",
        severity: "tip",
        reason: "missingSlideTitle",
        action: "edit-slide-title",
        slideId,
        slideIndex,
        ...titleBlock,
        label: `Slide ${slideIndex + 1}`,
        context: "title=blank",
      });
    } else if (title !== null && (titleCounts.get(title) ?? 0) > 1) {
      appendPresentationIssue(issues, {
        id: issueId(
          "slideTitle",
          slideId,
          titleReading.blockId,
          "duplicateSlideTitle",
        ),
        kind: "slideTitle",
        severity: "tip",
        reason: "duplicateSlideTitle",
        action: "edit-slide-title",
        slideId,
        slideIndex,
        ...titleBlock,
        label:
          shortAccessibleName(titleReading.title) || `Slide ${slideIndex + 1}`,
        context: `title=${JSON.stringify(title)};count=${titleCounts.get(title)}`,
      });
    }

    const blocks = Array.isArray(slide.blocks) ? slide.blocks : [];
    const blockIdCounts = countIdentities(
      blocks,
      (block) => block.id,
      budget,
    );
    if (!blockIdCounts) return [];
    for (let blockIndex = 0; blockIndex < blocks.length; blockIndex += 1) {
      if (issues.length >= MAX_PRESENTATION_ACCESSIBILITY_ISSUES) break;
      if (!consumeAccessibilityScanWork(budget)) return [];
      const block = blocks[blockIndex];
      const blockId = stableIdentity(block.id);
      const blockType = typeof block.type === "string" ? block.type : null;
      if (!blockId || blockIdCounts.get(blockId) !== 1 || !blockType) continue;

      if (blockType === "image") {
        const source = a11yRecord(block.source);
        const opaqueAlt =
          source?.pictureAltTextDisposition === "preserve-only";
        const metadataShapeKnown =
          (block.source === undefined || source !== null) &&
          (block.alt === undefined || typeof block.alt === "string") &&
          (block.decorative === undefined ||
            typeof block.decorative === "boolean");
        if (
          metadataShapeKnown &&
          !opaqueAlt &&
          block.decorative !== true &&
          shortScannerAccessibleName(block.alt, 1_000) === ""
        ) {
          appendPresentationIssue(issues, {
            id: issueId(
              "imageMissingAlt",
              slideId,
              blockId,
              "missingAltText",
            ),
            kind: "imageMissingAlt",
            severity: "error",
            reason: "missingAltText",
            action:
              canEditBlockContent(block) ? "edit-image-alt" : "navigate",
            slideId,
            slideIndex,
            blockId,
            blockType,
            label: issueLabel(block),
            context: `alt=blank;decorative=${String(block.decorative)}`,
          });
        }
      }

      if (blockType === "table") {
        const rows = tableRows(block.rows, budget);
        if (budget.exhausted) return [];
        if (rows) {
          if (
            (block.headerRow === undefined ||
              typeof block.headerRow === "boolean") &&
            block.headerRow !== true
          ) {
            appendPresentationIssue(issues, {
              id: issueId(
                "tableStructure",
                slideId,
                blockId,
                "missingTableHeader",
              ),
              kind: "tableStructure",
              severity: "error",
              reason: "missingTableHeader",
              action:
                canEditBlockContent(block) ? "fix-table-header" : "navigate",
              slideId,
              slideIndex,
              blockId,
              blockType,
              label: issueLabel(block),
              context: `headerRow=${String(block.headerRow)};rows=${rows.length}`,
            });
          }
          if (issues.length >= MAX_PRESENTATION_ACCESSIBILITY_ISSUES) break;
          const hasMergedCells = tableHasMergedCells(rows, budget);
          if (budget.exhausted) return [];
          if (hasMergedCells) {
            appendPresentationIssue(issues, {
              id: issueId(
                "tableStructure",
                slideId,
                blockId,
                "mergedTableCells",
              ),
              kind: "tableStructure",
              severity: "warning",
              reason: "mergedTableCells",
              action: "navigate",
              slideId,
              slideIndex,
              blockId,
              blockType,
              label: issueLabel(block),
              context: `merged=true;rows=${rows.length}`,
            });
          }
          if (issues.length >= MAX_PRESENTATION_ACCESSIBILITY_ISSUES) break;
          appendTableContrastIssues(
            issues,
            slideId,
            slideIndex,
            block,
            blockId,
            rows,
            budget,
          );
          if (budget.exhausted) return [];
        }
      } else {
        appendBlockContrastIssue(
          issues,
          slide,
          slideId,
          slideIndex,
          block,
          blockIndex,
          blocks,
          blockId,
          blockType,
          budget,
        );
        if (budget.exhausted) return [];
      }

      if (
        blockType === "embed" &&
        (block.kind === "video" || block.kind === "audio")
      ) {
        appendPresentationIssue(issues, {
          id: issueId(
            "mediaCaptions",
            slideId,
            blockId,
            "mediaCaptionsReview",
          ),
          kind: "mediaCaptions",
          severity: "warning",
          reason: "mediaCaptionsReview",
          action: "navigate",
          slideId,
          slideIndex,
          blockId,
          blockType,
          label: issueLabel(block),
          context: `kind=${block.kind};manual-review=true`,
        });
      }
    }

    if (issues.length >= MAX_PRESENTATION_ACCESSIBILITY_ISSUES) break;

    const visibleBlocks: Array<{
      block: DeterministicGeometryBlock;
      index: number;
      id: string;
    }> = [];
    let deterministic = true;
    for (let blockIndex = 0; blockIndex < blocks.length; blockIndex += 1) {
      if (!consumeAccessibilityScanWork(budget)) return [];
      const block = blocks[blockIndex];
      if (block.hidden === true) continue;
      const blockId = stableIdentity(block.id);
      if (
        blockId !== null &&
        blockIdCounts.get(blockId) === 1 &&
        (block.hidden === undefined || typeof block.hidden === "boolean") &&
        hasDeterministicGeometry(block)
      ) {
        visibleBlocks.push({ block, index: blockIndex, id: blockId });
      } else {
        deterministic = false;
        break;
      }
    }
    if (!deterministic || visibleBlocks.length < 2) continue;

    // Reserve a deterministic quadratic ceiling for sorting/comparison work.
    // Native sort normally needs far less, but this guarantees it never starts
    // when the remaining synchronous budget cannot cover its worst useful size.
    if (
      !consumeAccessibilityScanWork(
        budget,
        visibleBlocks.length * visibleBlocks.length,
      )
    ) {
      return [];
    }
    const modelOrder = visibleBlocks.map(({ id }) => id);
    const visualOrder = visibleBlocks
      .slice()
      .sort((left, right) => {
        return (
          left.block.top - right.block.top ||
          left.block.left - right.block.left ||
          left.index - right.index
        );
      })
      .map(({ id }) => id);
    if (modelOrder.every((id, index) => id === visualOrder[index])) continue;
    const context = boundedContext(
      JSON.stringify({ modelOrder, visualOrder }),
    );
    if (!context) continue;
    const firstMismatch = modelOrder.find(
      (id, index) => id !== visualOrder[index],
    );
    appendPresentationIssue(issues, {
      id: issueId(
        "readingOrder",
        slideId,
        firstMismatch ?? null,
        "spatialReadingOrder",
      ),
      kind: "readingOrder",
      severity: "warning",
      reason: "spatialReadingOrder",
      action: "open-selection-pane",
      slideId,
      slideIndex,
      blockId: firstMismatch ?? null,
      blockType: null,
      label: `Slide ${slideIndex + 1}`,
      context,
    });
  }

  return issues;
}

/** Complete order-sensitive authority fingerprint for one rendered finding. */
export function presentationAccessibilityIssueSignature(
  issue: PresentationAccessibilityIssue,
): string {
  return JSON.stringify([
    issue.id,
    issue.kind,
    issue.severity,
    issue.reason,
    issue.action,
    issue.slideId,
    issue.slideIndex,
    issue.blockId,
    issue.blockType,
    issue.label,
    issue.context,
  ]);
}

/** Re-derive and uniquely authenticate a finding against the exact live deck. */
export function resolvePresentationAccessibilityIssue(
  slides: readonly PresentationAccessibilitySlideLike[],
  expected: PresentationAccessibilityIssue,
): PresentationAccessibilityIssue | null {
  const matches = collectPresentationAccessibilityIssues(slides).filter(
    (candidate) => candidate.id === expected.id,
  );
  if (matches.length !== 1) return null;
  return presentationAccessibilityIssueSignature(matches[0]) ===
    presentationAccessibilityIssueSignature(expected)
    ? matches[0]
    : null;
}







export function presentationDecorativeGroupIds(
  // Any slide block: a parameter of only optional fields is a "weak type"
  // that a block stating neither field cannot be passed as (TS2345).
  blocks: readonly object[],
): ReadonlySet<string> {
  const ids = new Set<string>();
  for (const candidate of blocks) {
    const block = candidate as { readonly groupId?: unknown; readonly groupDecorative?: unknown };
    if (block.groupDecorative !== true || typeof block.groupId !== "string") continue;
    const id = block.groupId.trim();
    if (id) ids.add(id);
  }
  return ids;
}
