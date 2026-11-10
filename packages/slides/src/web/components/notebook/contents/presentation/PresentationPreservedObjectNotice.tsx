/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { PreservedObjectNotice } from "@mrnote/core/web/lib/presentation-preserved-object.ts";

/**
 * The frame of an object the file keeps and this page cannot draw.
 *
 * Deliberately an outline and a caption rather than any invented artwork: what
 * the page is missing IS the object's own drawing, so anything drawn in its
 * place would be a shape the file does not contain. The frame is the one thing
 * the slide really does say.
 *
 * Inert to the pointer — the block underneath stays selectable, movable and
 * deletable exactly as before.
 */
export function PresentationPreservedObjectNotice({
  notice,
  label,
}: {
  notice: PreservedObjectNotice;
  label: string;
}) {
  return (
    <div
      data-testid="presentation-preserved-object"
      data-preserved-kind={notice.kind}
      data-preserved-code={notice.code}
      title={label}
      aria-label={label}
      style={{
        position: "absolute",
        inset: 0,
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        padding: 8,
        boxSizing: "border-box",
        border: "1px dashed currentColor",
        borderRadius: 2,
        opacity: 0.55,
        pointerEvents: "none",
        overflow: "hidden",
      }}
    >
      <span
        style={{
          font: "500 11px/1.35 system-ui, sans-serif",
          textAlign: "center",
          maxWidth: "100%",
        }}
      >
        {label}
      </span>
    </div>
  );
}

export default PresentationPreservedObjectNotice;
