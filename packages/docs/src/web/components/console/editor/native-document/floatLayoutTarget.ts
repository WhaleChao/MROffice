/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The one floating drawing the Layout dialog acts on, whichever kind is
 * selected: a chart or a picture (`chartPlacement`), a shape
 * (`shapeButtons`), a text box (`textBoxFrameCommands`). Each keeps its own
 * attrs; the dialog reads and writes the one anchor model.
 */

import type { Editor } from "@tiptap/react";

import type { FloatAnchor } from "@mrnote/core/web/lib/office-float-anchor.ts";
import type { FloatPaint } from "@mrnote/core/web/lib/office-float-anchor-edit.ts";

import {
  liveDocumentDrawingLayout,
  liveDocumentDrawingSize,
  setLiveDocumentDrawingAnchor,
  setLiveDocumentDrawingExtent,
} from "@mrnote/core/web/components/console/editor/native-document/chartPlacement.ts";
import {
  liveDocumentShapeLayout,
  liveDocumentShapeSize,
  setLiveDocumentShapeAnchor,
  setLiveDocumentShapeExtent,
  toggleLiveDocumentShapeAspectRatio,
} from "@mrnote/core/web/components/console/editor/native-document/shapeButtons.ts";
import { BLOCK_PICTURE_NODE_NAME, INLINE_PICTURE_NODE_NAME } from "@mrnote/core/web/components/console/editor/native-document/documentObjectNodes.ts";
import { setLiveImageAspectRatioLock } from "./pictureButtons";
import { liveTextBoxLayout, liveTextBoxSize, setLiveTextBoxAnchor, setLiveTextBoxExtent } from "./textBoxFrameCommands";

/** The Size tab's reading: the drawing's size and whether its proportions
 * are locked (a picture's or a shape's own Lock aspect ratio). */
export interface LiveFloatLayoutSize {
  readonly widthEmu: number;
  readonly heightEmu: number;
  readonly lockAspectRatio: boolean;
}

export interface LiveFloatLayoutTarget {
  readonly anchor: FloatAnchor;
  readonly paint: FloatPaint | null;
  readonly size: LiveFloatLayoutSize | null;
  readonly apply: (anchor: FloatAnchor) => boolean;
  readonly applySize: (widthEmu: number, heightEmu: number) => boolean;
  /** Lock aspect ratio kept on the drawing; absent where it keeps none. */
  readonly applyLock?: (locked: boolean) => boolean;
}

export function liveDocumentFloatLayoutTarget(editor: Editor): LiveFloatLayoutTarget | null {
  const drawing = liveDocumentDrawingLayout(editor);
  if (drawing) {
    return {
      ...drawing,
      size: liveDocumentDrawingSize(editor),
      apply: (anchor) => setLiveDocumentDrawingAnchor(editor, anchor),
      applySize: (widthEmu, heightEmu) => setLiveDocumentDrawingExtent(editor, widthEmu, heightEmu),
      ...(editor.isActive(BLOCK_PICTURE_NODE_NAME) || editor.isActive(INLINE_PICTURE_NODE_NAME)
        ? { applyLock: (locked: boolean) => setLiveImageAspectRatioLock(editor, locked) }
        : {}),
    };
  }
  const shape = liveDocumentShapeLayout(editor);
  if (shape) {
    return {
      ...shape,
      size: liveDocumentShapeSize(editor),
      apply: (anchor) => setLiveDocumentShapeAnchor(editor, anchor),
      applySize: (widthEmu, heightEmu) => setLiveDocumentShapeExtent(editor, widthEmu, heightEmu),
      applyLock: (locked: boolean) =>
        liveDocumentShapeSize(editor)?.lockAspectRatio === locked || toggleLiveDocumentShapeAspectRatio(editor),
    };
  }
  const textBox = liveTextBoxLayout(editor as never);
  if (textBox) {
    return {
      ...textBox,
      size: liveTextBoxSize(editor as never),
      apply: (anchor) => setLiveTextBoxAnchor(editor as never, anchor),
      applySize: (widthEmu, heightEmu) => setLiveTextBoxExtent(editor as never, widthEmu, heightEmu),
    };
  }
  return null;
}
