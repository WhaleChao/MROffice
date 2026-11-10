/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Pure helpers for "paste an image from the OS clipboard" on the presentation
 * canvas (PowerPoint / WPS parity: Ctrl+V a screenshot / a copied image → an
 * image block on the slide).
 *
 * The component owns the async clipboard read (`navigator.clipboard.read()` →
 * `ClipboardItem.getType()` → a Blob → a data URL via `FileReader`) and the deck
 * commit; this module owns the two side-effect-free decisions worth testing in
 * isolation:
 *
 *  1. {@link clipboardImageCandidates} — rank the supported image MIME payloads
 *     across every `ClipboardItem` (preferring PNG, the lossless format
 *     screenshots land as). The compatibility wrapper
 *     {@link firstClipboardImageType} applies the same policy to one item.
 *
 *  2. {@link imageBlockFromDataUrl} — given a data-URL produced from the blob,
 *     VALIDATE it (the same magic-byte + size check the rest of the stack uses
 *     via `isValidInlineImageSource`) and, when valid, build a minimal image
 *     block (fresh id; empty alt/caption like the "Add image" button). A
 *     non-image / oversized / malformed data URL yields `null` so a hostile or
 *     truncated clipboard payload never lands on the slide.
 *
 * Dependency-free beyond the shared inline-image validator: no React, no DOM, no
 * `crypto`/`Date` of its own (the id factory is INJECTED so a test can assert the
 * exact block). Inputs are never mutated.
 */

import { isValidInlineImageSource } from "@mrnote/core/web/lib/inline-image-source.ts";
import {
  captureElementPicture,
  rasterElementPicture,
  type ElementPicture,
} from "@mrnote/core/web/lib/export/element-picture.ts";

/** A minimal positioned image block (the slide schema's `image` member subset). */
export interface ClipboardImageBlock {
  id: string;
  type: "image";
  src: string;
  alt: string;
  caption: string;
}

/** Minimal ClipboardItem surface used by the pure MIME planner. */
export interface ClipboardImageItemLike {
  readonly types: ReadonlyArray<unknown>;
}

/** A deterministic image read candidate within a ClipboardItem array. */
export interface ClipboardImageCandidate {
  readonly itemIndex: number;
  readonly type: string;
}

export interface PresentationClipboardPicturePlan {
  readonly leftEmu: number;
  readonly topEmu: number;
  readonly widthEmu: number;
  readonly heightEmu: number;
  readonly leftPx: number;
  readonly topPx: number;
  readonly widthPx: number;
  readonly heightPx: number;
  readonly canvasWidthPx: number;
  readonly canvasHeightPx: number;
}

function finite(value: unknown): value is number {
  return typeof value === "number" && Number.isFinite(value);
}

/** EMU per design px of the slide canvas (a 10 in slide is 960 px wide). */
const EMU_PER_DESIGN_PX = 9_525;

/**
 * How far an object's outline reaches past its frame (EMU): half its width
 * (an outline is centred on the edge, in design px or as stated in EMU) and a
 * pixel of anti-aliasing; nothing for an object drawn without one.
 */
function presentationClipboardPictureOutsetEmu(block: Readonly<Record<string, unknown>>): number {
  const stroke = block.stroke;
  if (typeof stroke === "string" && /^(none|transparent)$/i.test(stroke.trim())) return 0;
  const widthEmu = finite(block.strokeWidthEmu) && block.strokeWidthEmu > 0
    ? block.strokeWidthEmu
    : finite(block.strokeWidth) && block.strokeWidth > 0
      ? block.strokeWidth * EMU_PER_DESIGN_PX
      : 0;
  return widthEmu > 0 ? Math.ceil(widthEmu / 2 + EMU_PER_DESIGN_PX) : 0;
}

/**
 * Map an authenticated block snapshot onto one clipped canvas rectangle.
 * `canvas` is the canvas's layout box in CSS px (its client size: where its
 * objects are placed), unscaled by any zoom.
 * Picture paste intentionally refuses rotated/malformed geometry rather than
 * rasterizing a clipped or misleading result.
 */
export function presentationClipboardPicturePlan(
  blocks: ReadonlyArray<Record<string, unknown>> | null | undefined,
  slide: { width: number; height: number },
  canvas: { width: number; height: number },
): PresentationClipboardPicturePlan | null {
  if (
    !Array.isArray(blocks) ||
    blocks.length === 0 ||
    !finite(slide.width) ||
    slide.width <= 0 ||
    !finite(slide.height) ||
    slide.height <= 0 ||
    !finite(canvas.width) ||
    canvas.width <= 0 ||
    !finite(canvas.height) ||
    canvas.height <= 0
  ) {
    return null;
  }
  let left = Number.POSITIVE_INFINITY;
  let top = Number.POSITIVE_INFINITY;
  let right = Number.NEGATIVE_INFINITY;
  let bottom = Number.NEGATIVE_INFINITY;
  for (const block of blocks) {
    if (
      !finite(block.left) ||
      !finite(block.top) ||
      !finite(block.width) ||
      block.width <= 0 ||
      !finite(block.height) ||
      block.height <= 0 ||
      (block.rotation !== undefined && block.rotation !== 0)
    ) {
      return null;
    }
    // An outline is drawn on the object's edge, half of it outside: the
    // picture takes all of it in, as PowerPoint's picture of a shape does.
    const outset = presentationClipboardPictureOutsetEmu(block);
    left = Math.min(left, block.left - outset);
    top = Math.min(top, block.top - outset);
    right = Math.max(right, block.left + block.width + outset);
    bottom = Math.max(bottom, block.top + block.height + outset);
  }
  left = Math.max(0, left);
  top = Math.max(0, top);
  right = Math.min(slide.width, right);
  bottom = Math.min(slide.height, bottom);
  if (right <= left || bottom <= top) return null;
  const scaleX = canvas.width / slide.width;
  const scaleY = canvas.height / slide.height;
  return Object.freeze({
    leftEmu: Math.round(left),
    topEmu: Math.round(top),
    widthEmu: Math.round(right - left),
    heightEmu: Math.round(bottom - top),
    leftPx: left * scaleX,
    topPx: top * scaleY,
    widthPx: Math.max(1, (right - left) * scaleX),
    heightPx: Math.max(1, (bottom - top) * scaleY),
    canvasWidthPx: canvas.width,
    canvasHeightPx: canvas.height,
  });
}

function blobAsDataUrl(blob: Blob): Promise<string | null> {
  return new Promise((resolve) => {
    const reader = new FileReader();
    reader.onerror = () => resolve(null);
    reader.onload = () =>
      resolve(typeof reader.result === "string" ? reader.result : null);
    reader.readAsDataURL(blob);
  });
}












export function capturePresentationClipboardPicture(
  canvas: HTMLElement,
  selectedIds: ReadonlySet<string>,
  plan: PresentationClipboardPicturePlan,
): Promise<ElementPicture | null> {
  if (!canvas || !canvas.isConnected || selectedIds.size === 0) return Promise.resolve(null);
  try {
    // The picture is what the user sees: the stylesheet's look written onto
    // each element (a `foreignObject` reads inline style only), the editing
    // surfaces (a text editor, block-order buttons) replaced by or removed
    // from the clone, so neither a bare textarea nor hidden buttons reach the
    // PNG (PPT-B225), and the clone serialized as XML: HTML's unclosed
    // `<img>` and `<br>` made the whole picture unparseable, so Paste as
    // Picture silently pasted nothing (PPT-B224).
    return captureElementPicture(canvas, {
      width: plan.widthPx,
      height: plan.heightPx,
      omit: EDITING_CHROME,
      edit: (clone) => {
        let retained = 0;
        for (const node of Array.from(clone.querySelectorAll<HTMLElement>("[data-block-id]"))) {
          const id = node.dataset.blockId;
          if (!id || !selectedIds.has(id)) {
            node.remove();
            continue;
          }
          retained += 1;
          // The selection frame the copied object wears while selected (the
          // Office drawing frame is the block's own outline) is not the
          // object's.
          node.style.outlineStyle = "none";
        }
        if (retained !== selectedIds.size) return false;
        // The objects alone, as PowerPoint pictures them: the slide's
        // background stays on the slide, around them the picture is clear.
        clone.style.background = "none";
        // Unscaled, in the canvas's own layout box: the stage's zoom (the
        // canvas's own transform and its ancestors') is not the picture's,
        // and the objects are placed in the box inside the canvas's border.
        // A copy of the transform without the ancestors' scale put the
        
        clone.style.transform = "none";
        clone.style.border = "none";
        clone.style.boxSizing = "content-box";
        clone.style.width = `${plan.canvasWidthPx}px`;
        clone.style.height = `${plan.canvasHeightPx}px`;
        clone.style.position = "absolute";
        clone.style.left = `${-plan.leftPx}px`;
        clone.style.top = `${-plan.topPx}px`;
        clone.style.margin = "0";
        clone.style.pointerEvents = "none";
      },
      frame: (clone, doc) => {
        const wrapper = doc.createElement("div");
        wrapper.setAttribute(
          "style",
          `position:relative;width:${plan.widthPx}px;height:${plan.heightPx}px;overflow:hidden`,
        );
        wrapper.appendChild(clone);
        return wrapper;
      },
    }).catch(() => null);
  } catch {
    return Promise.resolve(null);
  }
}

/** The picture taken at Copy as a self-contained PNG data URL, or null. */
export async function rasterizePresentationClipboardPicture(
  captured: Promise<ElementPicture | null>,
): Promise<string | null> {
  let png: Blob | null = null;
  try {
    const picture = await captured;
    if (!picture) return null;
    png = await rasterElementPicture(picture, { scale: 2, background: null });
  } catch {
    return null;
  }
  if (!png || png.type !== "image/png") return null;
  const dataUrl = await blobAsDataUrl(png);
  return dataUrl && isValidInlineImageSource(dataUrl) ? dataUrl : null;
}

/** The canvas's editing surfaces: handles, selection, block buttons, fields. */
const EDITING_CHROME = [
  '[data-testid*="resize"]',
  '[data-testid*="rotation"]',
  '[data-testid*="selection"]',
  ".presentation-window__marquee",
  "button",
  '[role="toolbar"]',
  '[data-testid*="block-order"]',
  ".presentation-window__text-move-band",
  // An empty placeholder's prompt is the editor's hint, never the slide's
  // words (the show and every export leave it out).
  ".presentation-window__placeholder-prompt",
  "input",
  "select",
].join(", ");

/**
 * Whether a picture is shown differently from its source pixels: an artistic
 * effect, a colour adjustment, a crop, a picture style, a border, a shadow,
 * a flip or a turn. Such a picture is published to other applications as it
 * looks, not as its source (PPT-B250).
 */
export function presentationImageHasVisibleEdits(block: Record<string, unknown>): boolean {
  if (block.type !== "image") return false;
  const crop = block.cropRect as { x?: number; y?: number; w?: number; h?: number } | undefined;
  const cropped =
    !!crop && !((crop.x ?? 0) === 0 && (crop.y ?? 0) === 0 && (crop.w ?? 1) === 1 && (crop.h ?? 1) === 1);
  return (
    cropped ||
    ["artisticEffect", "recolor", "pictureStylePreset", "borderColor", "borderWidth", "cornerRadius"].some(
      (key) => block[key] !== undefined && block[key] !== null && block[key] !== "none",
    ) ||
    ["brightness", "contrast", "transparency"].some(
      (key) => typeof block[key] === "number" && block[key] !== 0,
    ) ||
    block.shadow === true ||
    block.flipH === true ||
    block.flipV === true ||
    (typeof block.rotation === "number" && block.rotation % 360 !== 0)
  );
}

const CLIPBOARD_IMAGE_TYPE_PRIORITY = [
  "image/png",
  "image/jpeg",
  "image/jpg",
  "image/webp",
  "image/gif",
  "image/bmp",
] as const;

const CLIPBOARD_IMAGE_TYPE_RANK = new Map<string, number>(
  CLIPBOARD_IMAGE_TYPE_PRIORITY.map((type, index) => [type, index]),
);

/**
 * Build the global image MIME plan for a ClipboardItem list.
 *
 * MIME quality wins before browser item order, so a later lossless PNG is read
 * before an earlier JPEG. Unknown image subtypes (SVG, TIFF, vendor payloads)
 * are refused instead of being interpreted or falling into a misleading image
 * shell. Ties retain stable item/type order.
 */
export function clipboardImageCandidates(
  items: ReadonlyArray<ClipboardImageItemLike> | null | undefined,
): ClipboardImageCandidate[] {
  if (!Array.isArray(items)) return [];
  const candidates: Array<
    ClipboardImageCandidate & { normalizedType: string; typeIndex: number }
  > = [];
  items.forEach((item, itemIndex) => {
    if (!item || !Array.isArray(item.types)) return;
    item.types.forEach((value: unknown, typeIndex: number) => {
      if (typeof value !== "string") return;
      const normalizedType = value.toLowerCase();
      if (!CLIPBOARD_IMAGE_TYPE_RANK.has(normalizedType)) return;
      // Preserve the exact advertised MIME spelling for ClipboardItem.getType;
      // normalization is used only for ranking.
      candidates.push({
        itemIndex,
        type: value,
        normalizedType,
        typeIndex,
      });
    });
  });
  candidates.sort((left, right) => {
    const rankDelta =
      (CLIPBOARD_IMAGE_TYPE_RANK.get(left.normalizedType) ??
        Number.MAX_SAFE_INTEGER) -
      (CLIPBOARD_IMAGE_TYPE_RANK.get(right.normalizedType) ??
        Number.MAX_SAFE_INTEGER);
    return (
      rankDelta ||
      left.itemIndex - right.itemIndex ||
      left.typeIndex - right.typeIndex
    );
  });
  return candidates.map(({ itemIndex, type }) => ({ itemIndex, type }));
}

/**
 * Pick the best supported image MIME off one `ClipboardItem`'s advertised
 * `types`, or `null` when none is supported. PNG is preferred (screenshots /
 * "copy image" land as lossless PNG), followed by the explicit safe raster
 * whitelist. A missing/non-array/text-only list yields `null`.
 *
 * Pure: the input is not mutated.
 */
export function firstClipboardImageType(
  types: ReadonlyArray<unknown> | null | undefined,
): string | null {
  if (!Array.isArray(types)) return null;
  return clipboardImageCandidates([{ types }])[0]?.type ?? null;
}

/** Standard image representation for a copied, self-contained image block. */
export function clipboardImageBlobFromBlock(
  block: Record<string, unknown> | null | undefined,
): { readonly mimeType: string; readonly blob: Blob } | null {
  const src = block?.src;
  if (
    block?.type !== "image" ||
    typeof src !== "string" ||
    !isValidInlineImageSource(src) ||
    typeof atob !== "function" ||
    typeof Blob === "undefined"
  ) {
    return null;
  }
  const match = /^data:(image\/[a-z0-9.+-]+);base64,([a-z0-9+/=]+)$/i.exec(src);
  if (!match || !CLIPBOARD_IMAGE_TYPE_RANK.has(match[1].toLowerCase())) {
    return null;
  }
  try {
    const binary = atob(match[2]);
    const bytes = new Uint8Array(binary.length);
    for (let index = 0; index < binary.length; index += 1) {
      bytes[index] = binary.charCodeAt(index);
    }
    return Object.freeze({
      mimeType: match[1].toLowerCase(),
      blob: new Blob([bytes], { type: match[1].toLowerCase() }),
    });
  } catch {
    return null;
  }
}

/**
 * Build a slide image block from a clipboard-derived data URL, or `null` when the
 * URL isn't a valid inline image. The validation is {@link isValidInlineImageSource}
 * (a `data:image/<mime>;base64,…` whose decoded payload matches the MIME's magic
 * bytes and is within the size cap) — identical to the guard the editor / schema
 * use — so a `data:text/html` smuggle, a wrong-magic blob, or an oversized image
 * is rejected. A valid block carries a fresh id (injected factory) and the same
 * empty alt/caption the "Add image" button uses.
 *
 * Pure + deterministic given a deterministic `makeId`.
 */
export function imageBlockFromDataUrl(
  dataUrl: unknown,
  makeId: () => string,
): ClipboardImageBlock | null {
  if (typeof dataUrl !== "string" || !dataUrl) return null;
  // Only data: URLs are accepted here — a clipboard image is always read into a
  // data URL by the caller; anything else (a remote http src, a blob: URL) is
  // out of scope and rejected so nothing unexpected reaches the slide.
  if (!dataUrl.toLowerCase().startsWith("data:image/")) return null;
  if (!isValidInlineImageSource(dataUrl)) return null;
  return {
    id: makeId(),
    type: "image",
    src: dataUrl,
    alt: "",
    caption: "",
  };
}
