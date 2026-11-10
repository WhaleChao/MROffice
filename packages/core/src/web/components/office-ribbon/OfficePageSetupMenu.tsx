"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { ReactNode } from "react";

import styles from "./OfficePageSetupMenu.module.css";

/**
 * Shared, presentation-only pieces for the Page Setup galleries (Margins,
 * Orientation, Page Size) and for Word's Breaks gallery, which uses the same
 * row skeleton with a description instead of a measurement list.
 *
 * These export the class names and the row's inner markup ONLY. The Radix
 * item, its `value`, its `onSelect`, its `data-testid` and every command stay
 * with the calling ribbon, so Docs and the Spreadsheet keep their own
 * behaviour and authority checks while sharing one look.
 */
export const officePageSetupMenuClassNames = {
  /** `DropdownMenuContent` for a Margins-style gallery. */
  presetMenu: styles.presetMenu,
  /** A radio row inside a Margins-style gallery. */
  presetItem: styles.presetItem,
  /** The preview glyph inside a preset row. */
  presetPreview: styles.presetPreview,
  /** The trailing "Custom …" row. */
  customItem: styles.customItem,
  /** `DropdownMenuContent` for a scrolling paper-size gallery. */
  paperMenu: styles.paperMenu,
  /** A radio row inside a paper-size gallery. */
  paperItem: styles.paperItem,

  /** `DropdownMenuContent` for a gallery of described rows (Breaks). */
  describedMenu: styles.describedMenu,
  /** A row carrying a bold title plus a wrapped grey description. */
  describedItem: styles.describedItem,
  /** A group heading inside a described gallery ("Page Breaks"). */
  groupLabel: styles.groupLabel,
} as const;

export interface OfficePageSetupPresetRowProps {
  /** The large preview glyph. Draw it with the caller's own icon set. */
  preview: ReactNode;
  /** Bold preset name, exactly as the host application words it. */
  title: string;
  /**
   * The measurements (Margins) or dimensions (Page Size) shown under the name.
   * Pass an array to lay several measurements out on one wrapped line.
   */
  details?: ReactNode;
  /** Id for the details element, so the row can `aria-describedby` it. */
  detailsId?: string;
  /**
   * `measurements` lays the details out as a wrapped list of nowrap spans
   * (Margins); `single` is one short line (Page Size); `description` is Word's
   * wrapped two-line sentence (Breaks).
   */
  detailsVariant?: "measurements" | "single" | "description";
}

const DETAILS_CLASS_NAMES = {
  measurements: styles.presetValues,
  single: styles.paperValues,
  description: styles.description,
} as const;

export interface OfficePaperSizePreviewProps {
  /** Sheet width in CSS pixels, already scaled against the gallery's tallest stock. */
  readonly width: number;
  /** Sheet height in CSS pixels, on the same scale. */
  readonly height: number;
}

/**
 * One paper stock drawn to scale in a Page Size gallery (Word's and Excel's
 * Size galleries): a white sheet with the same one-pixel edge on all four
 * sides, centred in the row's preview slot.
 *
 * The sheet is a CSS box sized in whole pixels. Browsers snap a box's border
 * to whole device pixels, so every edge of every stock is equally crisp. A
 * glyph stretched with `preserveAspectRatio="none"` to a fractional size
 * scaled its stroke separately per axis and left edges on half pixels, so
 * some stocks drew all four edges and others lost one or two.
 */
export function OfficePaperSizePreview({ width, height }: OfficePaperSizePreviewProps) {
  return (
    <span aria-hidden="true" className={styles.paperPreviewSlot}>
      <span
        className={styles.paperSheet}
        data-office-paper-sheet=""
        style={{
          width: `${Math.max(4, Math.round(width))}px`,
          height: `${Math.max(4, Math.round(height))}px`,
        }}
      />
    </span>
  );
}

/**
 * The inner markup of one gallery row: preview glyph, then a copy column with
 * the bold name and the grey details. Rendered as a fragment so the caller
 * supplies the Radix `DropdownMenuRadioItem`/`DropdownMenuItem` wrapper.
 */
export function OfficePageSetupPresetRow({
  preview,
  title,
  details,
  detailsId,
  detailsVariant = "measurements",
}: OfficePageSetupPresetRowProps) {
  return (
    <>
      {preview}
      <span className={styles.presetCopy}>
        <strong>{title}</strong>
        {details === undefined || details === null ? null : (
          <span className={DETAILS_CLASS_NAMES[detailsVariant]} id={detailsId}>
            {details}
          </span>
        )}
      </span>
    </>
  );
}
