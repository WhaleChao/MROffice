"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The Reorder Objects dialog's picture of one layer for a host with no slide
 * thumbnail to draw (Docs, Spreadsheet): the object's box in its own
 * proportions, centred in the dialog's 96 x 54 preview slot.
 */
export function OfficeReorderLayerPreview({
  box,
}: {
  readonly box: { readonly width: number; readonly height: number } | null;
}) {
  const width = box && box.width > 0 ? box.width : 4;
  const height = box && box.height > 0 ? box.height : 3;
  const scale = Math.min(80 / width, 42 / height);
  const w = Math.max(2, width * scale);
  const h = Math.max(2, height * scale);
  return (
    <svg viewBox="0 0 96 54" width="96" height="54" aria-hidden="true">
      <rect
        x={(96 - w) / 2}
        y={(54 - h) / 2}
        width={w}
        height={h}
        rx="2"
        fill="#dbeafe"
        stroke="#2563eb"
        strokeWidth="1"
      />
    </svg>
  );
}
