/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { PaginationBreakKind } from "../../../../lib/document-export/page-paginate";

import {
  KNOWN_FORMULA_ATTR,
  mathLatexForRender,
  ensureKatexFonts,
  previewTypesetsAsDisplay,
  renderKatexCached,
} from "../extensions/lazyKatex";
import { clearPageEdgeMarks } from "./paged/paginationProjectionKeys";

interface MeasuredFlowLayout {
  heights: number[];
  flowSpacingBefore: number[];
}

/** Classify editor-only pagination controls; null means ordinary content. */
export function paginationControlKind(
  block: HTMLElement,
): PaginationBreakKind | "continuous" | null {
  const type = block.getAttribute("data-type");
  if (type === "page-break") return "nextPage";
  if (type !== "section-break") return null;
  const sectionKind = block.getAttribute("data-break-type");
  if (sectionKind === "continuous") return "continuous";
  if (sectionKind === "evenPage" || sectionKind === "oddPage") {
    return sectionKind;
  }
  return "nextPage";
}

/** Remove live-editor selection and resize chrome from a measurement clone. */
export function stripEditorChrome(clone: HTMLElement): void {
  const livePaginationClasses = [
    "note-editor-page-boundary-before",
    "note-editor-page-boundary-after",
  ];
  clone.removeAttribute("contenteditable");
  clearPageEdgeMarks(clone);
  clone.classList.remove(
    "is-selected",
    "ProseMirror-selectednode",
    ...livePaginationClasses,
  );
  clone
    .querySelectorAll<HTMLElement>(".note-editor-image-handle,[data-document-selection-chrome]")
    .forEach((handle) => handle.remove());
  clone
    .querySelectorAll<HTMLElement>(
      ".note-editor-editable-page-gap,.ProseMirror-gapcursor",
    )
    .forEach((decoration) => decoration.remove());
  clone
    .querySelectorAll<HTMLElement>(
      ".is-selected,.ProseMirror-selectednode,.note-editor-page-boundary-before,.note-editor-page-boundary-after",
    )
    .forEach((selected) =>
      selected.classList.remove(
        "is-selected",
        "ProseMirror-selectednode",
        ...livePaginationClasses,
      ),
    );
  clone
    .querySelectorAll<HTMLElement>("[contenteditable]")
    .forEach((editable) => editable.removeAttribute("contenteditable"));
}

/**
 * Real-layout measurement with collapsed flow spacing owned by the following
 * printable block. PagePreview uses the full flow result; structural pages
 * cache body heights and combine them with live sheet-local spacing.
 */
export function measureFlowWithLayout(
  blocks: HTMLElement[],
): MeasuredFlowLayout {
  const heights = blocks.map(() => 0);
  const flowSpacingBefore = blocks.map(() => 0);
  const printableBlocks = blocks
    .map((block, index) => ({ block, index }))
    .filter(({ block }) => paginationControlKind(block) === null);
  const rects = printableBlocks.map(({ block }) =>
    block.getBoundingClientRect(),
  );
  printableBlocks.forEach(({ block, index }, printableIndex) => {
    const rect = rects[printableIndex];
    const style = getComputedStyle(block);
    const previous = rects[printableIndex - 1];
    const flowSpacing = previous
      ? Math.max(0, rect.top - (previous.top + previous.height))
      : Number.parseFloat(style.marginTop) || 0;
    const trailingMargin =
      printableIndex === printableBlocks.length - 1
        ? Number.parseFloat(style.marginBottom) || 0
        : 0;
    flowSpacingBefore[index] = flowSpacing;
    heights[index] = flowSpacing + Math.max(0, rect.height) + trailingMargin;
  });
  return { heights, flowSpacingBefore };
}

/** Real-layout height-only adapter retained for PagePreview's injectable prop. */
export function measureWithLayout(blocks: HTMLElement[]): number[] {
  return measureFlowWithLayout(blocks).heights;
}

function measureBodyHeights(blocks: HTMLElement[]): number[] {
  return blocks.map((block) => {
    if (paginationControlKind(block) !== null) return 0;
    return Math.max(
      0,
      block.offsetHeight || block.getBoundingClientRect().height,
    );
  });
}

export type TypesetCloneMeasurement = "flow" | "body";

export interface TypesetCloneLayoutTarget {
  /** Visible laid-out parent for the temporary measurement host. */
  parent: HTMLElement;
  /** Explicit printable content width, when the live source is display:none. */
  width?: number;
  /**
   * Optional semantic block order. Structural editing nests blocks in
   * pageBlock sheets, while pagination must measure the flattened reading
   * order. Omit for PagePreview's historical direct-child behavior.
   */
  sourceBlocks?: readonly HTMLElement[];
  /** Optional rendered content width for every source block. Structural
   * pagination uses this when pages/sections have different printable or
   * newspaper-column widths. */
  sourceBlockWidths?: readonly number[];
  /** One-shot hook while a large fully-typeset clone is still mounted. */
  onTypesetBatch?: (blocks: readonly HTMLElement[]) => void;
  /** Each block's clone before anything in it is typeset: what marks a formula
   * {@link KNOWN_FORMULA_ATTR} leaves it to stand as a placeholder. */
  prepareClone?: (block: HTMLElement, sourceIndex: number) => void;
  /** Only `onTypesetBatch` is wanted (the formulas' typeset boxes), not the
   * blocks' heights: every block is typeset in one clone, whatever their
   * number, and no height is read. Laid out and read one block at a time,
   * the page ledger's formulas cost a layout per block: 11 s of the
   * 296-page book's open in Chromium. */
  geometryOnly?: boolean;
}

function typesetOnePendingMath(el: HTMLElement): void {
  const latex = el.textContent ?? "";
  el.removeAttribute("data-math-pending");
  if (!latex.trim()) return;
  const displayMode = previewTypesetsAsDisplay(el);
  renderKatexCached(
    el,
    mathLatexForRender(latex, el.getAttribute("data-math-style")),
    displayMode,
  );
}

/** Typeset every pending formula synchronously (small docs / previews). */
export function typesetPendingMath(clone: HTMLElement): void {
  clone
    .querySelectorAll<HTMLElement>("[data-math-pending]")
    .forEach(typesetOnePendingMath);
}

async function typesetPendingMathChunked(
  clone: HTMLElement,
  isCancelled: () => boolean,
): Promise<boolean> {
  const pending = [
    ...(clone.matches("[data-math-pending]") ? [clone] : []),
    ...Array.from(clone.querySelectorAll<HTMLElement>(`[data-math-pending]:not([${KNOWN_FORMULA_ATTR}])`)),
  ];
  let sliceStartedAt = performance.now();
  for (let index = 0; index < pending.length; index += 1) {
    if (isCancelled()) return false;
    typesetOnePendingMath(pending[index]);
    if (performance.now() - sliceStartedAt >= 10) {
      await new Promise((resolve) => setTimeout(resolve, 0));
      sliceStartedAt = performance.now();
    }
  }
  return !isCancelled();
}

/**
 * Measure a fully-typeset clone while leaving the live editor's lazy KaTeX DOM
 * untouched. `sourceBlocks` lets structural pageBlock editing flatten blocks
 * across sheets without making the clone depend on current page boundaries.
 */
export async function measureTypesetCloneBlockHeights(
  sourceDom: HTMLElement,
  isCancelled: () => boolean,
  blockIndices?: readonly number[],
  onMeasuredBlock?: (
    index: number,
    block: HTMLElement,
    flowSpacingBefore: number,
  ) => void,
  measurement: TypesetCloneMeasurement = "flow",
  layoutTarget?: TypesetCloneLayoutTarget,
): Promise<number[] | null> {
  const parent =
    layoutTarget?.parent ?? sourceDom.parentElement ?? document.body;

  const host = document.createElement("div");
  host.className = "note-editor-page-preview__measure";
  host.setAttribute("aria-hidden", "true");
  const explicitWidth = layoutTarget?.width;
  const sourceWidth =
    typeof explicitWidth === "number" &&
    Number.isFinite(explicitWidth) &&
    explicitWidth > 0
      ? explicitWidth
      : sourceDom.offsetWidth || sourceDom.clientWidth;
  if (sourceWidth > 0) host.style.width = `${sourceWidth}px`;

  const cloneShell = sourceDom.cloneNode(false) as HTMLElement;
  stripEditorChrome(cloneShell);
  const sourceBlocks = layoutTarget?.sourceBlocks
    ? Array.from(layoutTarget.sourceBlocks)
    : Array.from(sourceDom.children).filter(
        (child): child is HTMLElement =>
          child instanceof HTMLElement &&
          !child.classList.contains("note-editor-editable-page-gap") &&
          !child.classList.contains("ProseMirror-gapcursor"),
      );
  const sourceBlockWidths = layoutTarget?.sourceBlockWidths;
  const blockWidth = (index: number): number | null => {
    if (sourceBlockWidths?.length !== sourceBlocks.length) return null;
    const value = sourceBlockWidths[index];
    return typeof value === "number" && Number.isFinite(value) && value > 0
      ? value
      : null;
  };
  const requested = blockIndices
    ? new Set(blockIndices)
    : new Set(sourceBlocks.map((_, index) => index));
  const controlKinds = sourceBlocks.map(paginationControlKind);
  const previousPrintableIndex = new Array<number | null>(
    sourceBlocks.length,
  ).fill(null);
  const nextPrintableIndex = new Array<number | null>(sourceBlocks.length).fill(
    null,
  );
  let previousPrintable: number | null = null;
  for (let index = 0; index < sourceBlocks.length; index += 1) {
    previousPrintableIndex[index] = previousPrintable;
    if (controlKinds[index] === null) previousPrintable = index;
  }
  let nextPrintable: number | null = null;
  for (let index = sourceBlocks.length - 1; index >= 0; index -= 1) {
    nextPrintableIndex[index] = nextPrintable;
    if (controlKinds[index] === null) nextPrintable = index;
  }

  parent.append(host);
  host.append(cloneShell);
  try {
    await new Promise((resolve) => setTimeout(resolve, 0));
    await ensureKatexFonts();
    if (isCancelled()) return null;

    if (requested.size > 500 || layoutTarget?.geometryOnly) {
      cloneShell.replaceChildren();
      const clonedBlocks = sourceBlocks.map((sourceBlock, index) => {
        const block = sourceBlock.cloneNode(true) as HTMLElement;
        stripEditorChrome(block);
        const width = blockWidth(index);
        if (width !== null) block.style.width = `${width}px`;
        block
          .querySelectorAll<HTMLElement>(
            ".note-editor-editable-page-gap,.ProseMirror-gapcursor",
          )
          .forEach((decoration) => decoration.remove());
        layoutTarget?.prepareClone?.(block, index);
        if (paginationControlKind(block) !== null) {
          block.replaceChildren();
          block.setAttribute("aria-hidden", "true");
          block.style.setProperty("display", "none", "important");
          block.style.setProperty("margin", "0", "important");
        }
        cloneShell.append(block);
        return block;
      });
      const done = await typesetPendingMathChunked(cloneShell, isCancelled);
      if (!done || isCancelled()) return null;
      if (document.fonts?.ready) await document.fonts.ready;
      if (isCancelled()) return null;
      layoutTarget?.onTypesetBatch?.(clonedBlocks);
      if (layoutTarget?.geometryOnly) return new Array<number>(sourceBlocks.length).fill(Number.NaN);
      const flowLayout = measureFlowWithLayout(clonedBlocks);
      const measured =
        measurement === "body"
          ? measureBodyHeights(clonedBlocks)
          : flowLayout.heights;
      const heights = new Array<number>(sourceBlocks.length).fill(Number.NaN);
      requested.forEach((index) => {
        const block = clonedBlocks[index];
        if (!block) return;
        onMeasuredBlock?.(
          index,
          block,
          flowLayout.flowSpacingBefore[index] ?? 0,
        );
        heights[index] = measured[index] ?? Number.NaN;
      });
      return heights;
    }

    const heights = new Array<number>(sourceBlocks.length).fill(Number.NaN);
    let sliceStartedAt = performance.now();

    for (let index = 0; index < sourceBlocks.length; index += 1) {
      if (!requested.has(index)) continue;
      if (isCancelled()) return null;
      if (controlKinds[index] !== null) {
        heights[index] = 0;
        continue;
      }
      cloneShell.replaceChildren();
      const width = blockWidth(index);
      cloneShell.style.width = width !== null ? `${width}px` : "";

      const previousIndex = previousPrintableIndex[index];
      const previousContext =
        previousIndex !== null
          ? (sourceBlocks[previousIndex].cloneNode(false) as HTMLElement)
          : null;
      if (previousContext) stripEditorChrome(previousContext);
      if (previousContext) cloneShell.append(previousContext);
      const block = sourceBlocks[index].cloneNode(true) as HTMLElement;
      stripEditorChrome(block);
      block
        .querySelectorAll<HTMLElement>(
          ".note-editor-editable-page-gap,.ProseMirror-gapcursor",
        )
        .forEach((decoration) => decoration.remove());
      cloneShell.append(block);
      const followingIndex = nextPrintableIndex[index];
      if (followingIndex !== null) {
        const followingContext = sourceBlocks[followingIndex].cloneNode(
          false,
        ) as HTMLElement;
        stripEditorChrome(followingContext);
        cloneShell.append(followingContext);
      }

      const done = await typesetPendingMathChunked(block, isCancelled);
      if (!done || isCancelled()) return null;
      if (document.fonts?.ready) await document.fonts.ready;
      if (isCancelled()) return null;
      const style = getComputedStyle(block);
      const marginTop = Number.parseFloat(style.marginTop) || 0;
      const marginBottom = Number.parseFloat(style.marginBottom) || 0;
      const previousMarginBottom = previousContext
        ? Number.parseFloat(getComputedStyle(previousContext).marginBottom) || 0
        : 0;
      const bodyHeight =
        block.offsetHeight || block.getBoundingClientRect().height;
      const flowSpacing =
        previousIndex === null
          ? marginTop
          : Math.max(previousMarginBottom, marginTop);
      onMeasuredBlock?.(index, block, flowSpacing);
      heights[index] =
        measurement === "body"
          ? Math.max(0, bodyHeight)
          : Math.max(0, bodyHeight) +
            flowSpacing +
            (followingIndex === null ? marginBottom : 0);

      if (performance.now() - sliceStartedAt >= 10) {
        await new Promise((resolve) => setTimeout(resolve, 0));
        sliceStartedAt = performance.now();
      }
    }
    return heights;
  } finally {
    host.remove();
  }
}
