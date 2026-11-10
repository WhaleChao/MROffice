"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import {
  AlertTriangle,
  Trash2,
} from "@mrnote/core/web/components/office-ribbon/icons/houseIcons.tsx";

export interface PdfPendingRedactionReviewItem {
  /** Opaque annotation identity used only for current-state callback lookup. */
  id: string;
  /** Original 1-based source page recorded by the persisted mark. */
  page: number;
  /** Exact structural-output targets for this original page. */
  outputPageCount: number;
  /** True only when the original page and its overlay are currently visible. */
  canNavigate: boolean;
}

export interface PdfPendingRedactionReviewCopy {
  summary: string;
  markLabel: string;
  pageLabel: (page: number) => string;
  mappedLabel: (count: number) => string;
  copyOnlyLabel: (count: number) => string;
  unmappedLabel: string;
  removeLabel: string;
}

interface PdfRedactionPendingReviewProps {
  items: ReadonlyArray<PdfPendingRedactionReviewItem>;
  busy: boolean;
  copy: PdfPendingRedactionReviewCopy;
  onNavigate: (annotationId: string) => void;
  onDelete: (annotationId: string) => void;
}

/**
 * Compact pre-Apply review for persisted redaction marks.
 *
 * Redactions deliberately do not enter the ordinary Comments pane: these rows
 * represent pending destructive intent, not review comments. The panel exposes
 * only a stable ordinal, original page and structural-output status; covered
 * text, geometry and annotation ids never appear in the UI. Mutation and
 * navigation authority remain in PDFViewer, which re-resolves each opaque id
 * against its current live state before acting.
 */
export default function PdfRedactionPendingReview({
  items,
  busy,
  copy,
  onNavigate,
  onDelete,
}: PdfRedactionPendingReviewProps) {
  const hasUnmapped = items.some((item) => item.outputPageCount === 0);

  return (
    <details
      className="pdf-viewer__redaction-review"
      data-testid="pdf-redaction-review"
    >
      <summary
        className="pdf-viewer__redaction-review-summary min-h-11 focus-visible:outline focus-visible:outline-2 focus-visible:outline-offset-2 focus-visible:outline-blue-600"
        data-testid="pdf-redaction-review-toggle"
      >
        {hasUnmapped ? <AlertTriangle size={13} aria-hidden="true" /> : null}
        <span>{copy.summary}</span>
        <span className="pdf-viewer__redaction-review-count">
          {items.length}
        </span>
      </summary>

      <ol
        className="pdf-viewer__redaction-review-list"
        data-testid="pdf-redaction-review-list"
      >
        {items.map((item, index) => {
          const mappingStatus =
            item.outputPageCount === 0
              ? "unmapped"
              : item.canNavigate
                ? "mapped"
                : "copy-only";
          const statusLabel =
            mappingStatus === "unmapped"
              ? copy.unmappedLabel
              : mappingStatus === "copy-only"
                ? copy.copyOnlyLabel(item.outputPageCount)
                : copy.mappedLabel(item.outputPageCount);
          const rowLabel = `${copy.markLabel} ${index + 1}, ${copy.pageLabel(item.page)}`;

          return (
            <li
              key={item.id}
              className={`pdf-viewer__redaction-review-item pdf-viewer__redaction-review-item--${mappingStatus} !grid-cols-[minmax(0,1fr)_44px]`}
              data-testid="pdf-redaction-review-item"
              data-mapping-status={mappingStatus}
              data-output-page-count={item.outputPageCount}
            >
              <button
                type="button"
                className="pdf-viewer__redaction-review-jump min-h-11 focus-visible:outline focus-visible:outline-2 focus-visible:outline-offset-2 focus-visible:outline-blue-600"
                onClick={() => onNavigate(item.id)}
                disabled={busy || !item.canNavigate}
                aria-label={rowLabel}
                title={rowLabel}
                data-testid="pdf-redaction-review-jump"
              >
                <span>
                  <strong>
                    {copy.markLabel} {index + 1}
                  </strong>
                  <span>{copy.pageLabel(item.page)}</span>
                </span>
                <small>{statusLabel}</small>
              </button>
              <button
                type="button"
                className="pdf-viewer__redaction-review-delete min-h-11 min-w-11 focus-visible:outline focus-visible:outline-2 focus-visible:outline-offset-2 focus-visible:outline-blue-600"
                onClick={() => onDelete(item.id)}
                disabled={busy}
                aria-label={`${copy.removeLabel}: ${rowLabel}`}
                title={`${copy.removeLabel}: ${rowLabel}`}
                data-testid="pdf-redaction-review-delete"
              >
                <Trash2 size={13} aria-hidden="true" />
              </button>
            </li>
          );
        })}
      </ol>
    </details>
  );
}
