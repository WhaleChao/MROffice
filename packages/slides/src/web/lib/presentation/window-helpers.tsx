/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Small helpers of the presentation editor window that are not tied to one
 * feature: edit-failure messages, the OS-clipboard paste authority type, the
 * uniform-selection and layer-capability readers, frame and custom-show draft
 * types, the Selection Pane type icon, the file-picker cancel hook, picture
 * compression source.
 *
 * Moved verbatim out of `components/notebook/contents/PresentationWindow.tsx`
 * (structural split, no behaviour change).
 */

import { useCallback, useEffect, useRef, type RefObject } from "react";
import type { OfficeSemanticIconId } from "@mrnote/core/web/components/office-ribbon/icons/officeIconRegistry.ts";
import { isValidInlineImageSource } from "@mrnote/core/web/lib/inline-image-source.ts";
import { hasLayoutGeometry, isImportedPptxBlock } from "@mrnote/core/web/lib/presentation-object-buttons.ts";
import type {
  NativePageCommandRejectionReason,
  PresentationBlock,
  PresentationImageBlock,
} from "@mrnote/core/web/components/console/editor/native-document/documentCommands.ts";
import type { SlideModel } from "@mrnote/core/web/lib/presentation/deck-model.ts";

export function presentationEditFailureMessage(
  t: (key: string) => string,
  reason?: NativePageCommandRejectionReason,
): string {
  if (reason === "invalid_image_source") {
    return t("presentation.editRejectedImage");
  }
  if (reason === "payload_too_large") {
    return t("presentation.editRejectedTooLarge");
  }
  return t("presentation.editTargetNotFound");
}

/**
 * Immutable admission receipt for one async OS-clipboard paste. Clipboard
 * reads can outlive both the slide that initiated them and an intervening deck
 * edit, so every awaited boundary reauthenticates this exact deck/slide/block
 * snapshot before any object is appended.
 */
export interface SystemClipboardPasteAuthority {
  requestId: number;
  deckGeneration: number;
  slides: SlideModel[];
  slideIndex: number;
  slide: SlideModel;
  blocks: PresentationBlock[];
  bounds: { width: number; height: number };
}

export function uniformPresentationSelectionValue<T>(
  values: readonly T[],
): T | null {
  const [first, ...rest] = values;
  if (first === undefined) return null;
  return rest.every((value) => Object.is(value, first)) ? first : null;
}

export function presentationSelectionLayerCapability(
  blocks: readonly PresentationBlock[],
  selection: readonly PresentationBlock[],
): Readonly<{ canMoveBackward: boolean; canMoveForward: boolean }> {
  const selectedIds = new Set(selection.map((block) => block.id));
  let unselectedBefore = false;
  let canMoveBackward = false;
  for (const block of blocks) {
    if (selectedIds.has(block.id)) {
      if (unselectedBefore) canMoveBackward = true;
    } else {
      unselectedBefore = true;
    }
  }
  let unselectedAfter = false;
  let canMoveForward = false;
  for (let index = blocks.length - 1; index >= 0; index -= 1) {
    const block = blocks[index];
    if (!block) continue;
    if (selectedIds.has(block.id)) {
      if (unselectedAfter) canMoveForward = true;
    } else {
      unselectedAfter = true;
    }
  }
  return { canMoveBackward, canMoveForward };
}

export type PresentationFrameAxis = "width" | "height";

export const EMPTY_PRESENTATION_FRAME_SIZE_DRAFTS: Record<
  PresentationFrameAxis,
  string
> = {
  width: "",
  height: "",
};

export interface PresentationCustomShowDraft {
  readonly mode: "new" | "edit";
  readonly showId: number;
  readonly name: string;
  readonly slideIds: readonly string[];
  readonly availableSlideId: string;
  readonly selectedSequenceIndex: number;
}

/** The Selection Pane row's type icon: the Insert command's own glyph. */
export function selectionPaneTypeIcon(
  block: PresentationBlock,
  memberCount: number,
): OfficeSemanticIconId {
  if (memberCount > 1) return "object.group";
  switch (block.type) {
    case "text":
      return "textBox.insert";
    case "image":
      return "image.insert";
    case "table":
      return "table.insert";
    case "chart":
      return "chart.insert";
    case "embed":
      return "link.insert";
    default:
      return "shape.insert";
  }
}

export function isPresentationOfficeGraphic(
  block: PresentationBlock | null | undefined,
): block is PresentationImageBlock {
  return (
    block?.type === "image" &&
    block.officeGraphicKind === "svg" &&
    (block.mimeType === undefined || block.mimeType === "image/svg+xml") &&
    isValidInlineImageSource(block.src) &&
    (typeof block.graphicFallbackSrc === "string"
      ? isValidInlineImageSource(block.graphicFallbackSrc)
      : block.sourceKind !== "asset" &&
        block.source?.officeGraphicDisposition ===
          "native-svg-no-raster-fallback")
  );
}

/**
 * A native ``p:pic`` with ``a:stretch/a:fillRect`` paints the complete image
 * into the picture frame. Imported pictures carry that native source identity
 * plus fixed geometry; letting the browser keep the bitmap's intrinsic aspect
 * ratio letterboxes it inside the frame and makes every picture appear shifted
 * (most visibly downward). Keep ordinary/native-created images unchanged, but
 * make imported positioned pictures obey the DrawingML frame in editor and
 * present mode alike.
 */
/**
 * Bind a hidden file input's `cancel` event through a CALLBACK REF rather than
 * an effect that reads `ref.current`.
 *
 * The editor returns a loader on its first render (it has no deck yet), so none
 * of these inputs exist when effects first run; an effect that reads the ref and
 * depends only on a stable handler then runs exactly once, against `null`, and
 * never again. Four of the six pickers here -- Change Picture, Insert Audio,
 * Insert Captions, Transition Sound -- had no `cancel` listener at all for that
 * reason, and the other two attached only because their dependency happened to
 * change identity on every render. React calls a callback ref with the real node
 * the moment it mounts, and with `null` when it goes away, which is the only
 * thing that knows when a ref has been filled.
 */








export function useFilePickerFocusFallback(
  inputRef: RefObject<HTMLInputElement | null>,
  openRef: RefObject<boolean>,
  onCancel: () => void,
): void {
  const handlerRef = useRef(onCancel);
  handlerRef.current = onCancel;
  useEffect(() => {
    let returned = false;
    const onFocus = () => {
      if (openRef.current) returned = true;
    };
    const onPress = () => {
      if (!returned) return;
      returned = false;
      const input = inputRef.current;
      if (openRef.current && (!input?.files || input.files.length === 0)) {
        handlerRef.current();
      }
    };
    window.addEventListener("focus", onFocus);
    window.addEventListener("pointerdown", onPress, true);
    window.addEventListener("keydown", onPress, true);
    return () => {
      window.removeEventListener("focus", onFocus);
      window.removeEventListener("pointerdown", onPress, true);
      window.removeEventListener("keydown", onPress, true);
    };
  }, [inputRef, openRef]);
}

export function useFilePickerCancel(
  ref: RefObject<HTMLInputElement | null>,
  onCancel: () => void,
): (node: HTMLInputElement | null) => void {
  const handlerRef = useRef(onCancel);
  handlerRef.current = onCancel;
  const listenerRef = useRef<(() => void) | null>(null);
  if (listenerRef.current === null) {
    listenerRef.current = () => handlerRef.current();
  }
  return useCallback(
    (node: HTMLInputElement | null) => {
      const listener = listenerRef.current!;
      const previous = ref.current;
      if (previous === node) return;
      previous?.removeEventListener("cancel", listener);
      ref.current = node;
      node?.addEventListener("cancel", listener);
    },
    [ref],
  );
}

export function presentationPictureCompressionSource(
  block: PresentationImageBlock,
): string | null {
  if (isImportedPptxBlock(block) || !hasLayoutGeometry(block)) return null;
  const source = String(block.src || "").trim();
  return /^data:image\/(?:png|jpeg|webp);base64,[A-Za-z0-9+/=]+$/i.test(
    source,
  )
    ? source
    : null;
}

/** The deck's dock labels its tabs itself (usePresentationTaskPaneDock), so a
 * pane in its slot reports nothing back. */
export function ignoreTaskPaneDockLabel(): void {}

