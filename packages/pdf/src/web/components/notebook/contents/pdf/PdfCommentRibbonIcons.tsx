/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { OfficeSemanticIcon } from "@mrnote/core/web/components/office-ribbon/icons/OfficeSemanticIcon.tsx";
import type { OfficeSemanticIconId } from "@mrnote/core/web/components/office-ribbon/icons/officeIconRegistry.ts";

import { PdfInsertRibbonIcon } from "./PdfInsertRibbonIcons";

export type PdfCommentRibbonIconName =
  | "highlight"
  | "areaHighlight"
  | "draw"
  | "note"
  | "attachment"
  | "stamp"
  | "measure"
  | "measureArea"
  | "measurePerimeter"
  | "shape"
  | "underline"
  | "strikeout"
  | "squiggly"
  | "replaceMark"
  | "insertCaret"
  | "commentMode"
  | "hideComments"
  | "manageComments"
  | "exportComments"
  | "excerptPage"
  | "aiSummary"
  | "importComments"
  | "readAloud"
  | "readAloudSettings"
  | "readAloudPrevious"
  | "readAloudNext"
  | "pause"
  | "stop"
  | "line"
  | "arrow"
  | "rectangle"
  | "ellipse"
  | "eraser"
  | "copy"
  | "paste"
  | "duplicate"
  | "delete";

/*
 * The comment tab draws a command with the icon every other ribbon draws it
 * with: the shared Office glyph where one says the command (the same glyph the
 * Edit tab's shapes, Home's AI summary and the selection toolbar's underline
 * use), the PDF tabs' own stamp, and -- for the PDF markups Office has no
 * command for -- a glyph in the same two-tone paper / ink / accent language.
 */
const SHARED_ICONS = {
  aiSummary: "ai.assistant",
  arrow: "nav.arrow.upRight",
  attachment: "file.attach",
  commentMode: "review.comment",
  copy: "edit.copy",
  delete: "edit.delete",
  draw: "draw.pen",
  duplicate: "edit.duplicate",
  ellipse: "shape.circle",
  eraser: "draw.eraser.stroke",
  exportComments: "file.export",
  hideComments: "view.hide",
  highlight: "draw.highlighter",
  importComments: "file.import",
  line: "action.remove",
  measure: "view.ruler",
  measureArea: "selection.marquee",
  note: "comment.add",
  paste: "edit.paste",
  pause: "media.pause",
  readAloud: "review.readAloud",
  readAloudNext: "media.next",
  readAloudSettings: "app.settings",
  rectangle: "shape.square",
  shape: "shape.insert",
  stop: "media.stop",
  strikeout: "text.strikethrough",
  underline: "text.underline",
} as const satisfies Partial<
  Record<PdfCommentRibbonIconName, OfficeSemanticIconId>
>;

type PdfCommentOwnGlyph =
  | "areaHighlight"
  | "squiggly"
  | "replaceMark"
  | "insertCaret"
  | "manageComments"
  | "excerptPage"
  | "measurePerimeter";

function PdfCommentGlyph({ name }: { name: PdfCommentOwnGlyph }) {
  switch (name) {
    case "areaHighlight":
      return (
        <g data-pdf-comment-glyph="area-highlight">
          <path className="docs-home-icon__paper" d="M4.5 3.2h15v17.6h-15z" />
          <path className="docs-home-icon__ink-soft" d="M7.4 6.6h9.2M7.4 17.6h6.2" />
          <path className="docs-home-icon__highlight-fill" d="M6.6 9.4h10.8v5.4H6.6z" />
          <path className="docs-home-icon__ink" d="M8.6 12.1h6.8" />
        </g>
      );
    case "squiggly":
      return (
        <g data-pdf-comment-glyph="squiggly-markup">
          <path className="docs-home-icon__ink-soft" d="M5 6h14M5 10h10.5" />
          <path
            className="docs-home-icon__accent"
            d="M4 16.2c1.3-2.1 2.6-2.1 3.9 0s2.6 2.1 3.9 0 2.6-2.1 3.9 0 2.6 2.1 3.9 0"
          />
        </g>
      );
    case "replaceMark":
      return (
        <g data-pdf-comment-glyph="replace-markup">
          <path className="docs-home-icon__ink-soft" d="M4.5 5.6h15M4.5 13.4h6.4" />
          <path className="docs-home-icon__danger" d="M3.6 9.5h16.8" />
          <path className="docs-home-icon__accent-soft-fill" d="M12.6 12.8h8.4v8.4h-8.4z" />
          <path className="docs-home-icon__accent" d="m14.4 19.6 2.4-5.2 2.4 5.2M15.3 17.8h3" />
        </g>
      );
    case "insertCaret":
      return (
        <g data-pdf-comment-glyph="insert-caret">
          <path className="docs-home-icon__ink-soft" d="M3.8 14.6h5.4M14.8 14.6h5.4" />
          <path className="docs-home-icon__accent" d="m8 19.8 4-6.2 4 6.2" />
          <path className="docs-home-icon__accent-soft-fill" d="M8 2.6h8v7.4H8z" />
          <path className="docs-home-icon__accent" d="M12 4.2v4.2M9.9 6.3h4.2" />
        </g>
      );
    case "manageComments":
      return (
        <g data-pdf-comment-glyph="comments-pane">
          <path className="docs-home-icon__paper" d="M3.4 3.6h17.2v16.8H3.4z" />
          <path className="docs-home-icon__ink-soft" d="M12.2 3.6v16.8M5.6 7.4h4.2M5.6 10.4h3.2M5.6 13.4h4.2" />
          <path className="docs-home-icon__accent-soft-fill" d="M14 6.6h5v5h-2.4l-1.8 1.6v-1.6H14z" />
          <path className="docs-home-icon__accent" d="M14 6.6h5v5h-2.4l-1.8 1.6v-1.6H14zM14 16.2h4.6" />
        </g>
      );
    case "measurePerimeter":
      return (
        <g data-pdf-comment-glyph="measure-perimeter">
          <path className="docs-home-icon__accent-soft-fill" d="m12 3.4 7.6 5.5-2.9 9.1H7.3L4.4 8.9z" />
          <path className="docs-home-icon__accent" d="m12 3.4 7.6 5.5-2.9 9.1H7.3L4.4 8.9z" />
          <path className="docs-home-icon__ink-soft" d="M7.3 21h9.4M7.3 19.8v2.4M16.7 19.8v2.4" />
        </g>
      );
    case "excerptPage":
      return (
        <g data-pdf-comment-glyph="excerpt-page">
          <path className="docs-home-icon__paper" d="M5.2 2.8h9.4l4.2 4.2v14.2H5.2z" />
          <path className="docs-home-icon__ink-soft" d="M14.6 2.8V7h4.2M8 17.8h6.6" />
          <path className="docs-home-icon__highlight-fill" d="M7.6 9.4h8.8v2.6H7.6z" />
          <path className="docs-home-icon__accent" d="M8 14.6h4.6" />
          <path className="docs-home-icon__accent-fill" d="M6.2 13.8h1.1v1.6H6.2z" />
        </g>
      );
  }
}

export default function PdfCommentRibbonIcon({
  name,
  size = 24,
}: {
  name: PdfCommentRibbonIconName;
  size?: number;
}) {
  if (name in SHARED_ICONS) {
    return (
      <OfficeSemanticIcon
        semanticId={SHARED_ICONS[name as keyof typeof SHARED_ICONS]}
        size={size}
      />
    );
  }
  if (name === "stamp") {
    // The stamp Insert and Edit draw for 添加图章.
    return <PdfInsertRibbonIcon name="stamp" size={size} />;
  }
  if (name === "readAloudPrevious") {
    // Previous is Next turned about, as a media bar draws the pair.
    return (
      <span
        aria-hidden="true"
        data-pdf-comment-icon={name}
        style={{ display: "inline-flex", transform: "scaleX(-1)" }}
      >
        <OfficeSemanticIcon semanticId="media.next" size={size} />
      </span>
    );
  }

  return (
    <svg
      aria-hidden="true"
      className="docs-home-icon"
      data-pdf-comment-icon={name}
      fill="none"
      focusable="false"
      height={size}
      shapeRendering="geometricPrecision"
      stroke="currentColor"
      strokeLinecap="round"
      strokeLinejoin="round"
      strokeWidth={1.25}
      viewBox="0 0 24 24"
      width={size}
      xmlns="http://www.w3.org/2000/svg"
    >
      <PdfCommentGlyph name={name as PdfCommentOwnGlyph} />
    </svg>
  );
}
