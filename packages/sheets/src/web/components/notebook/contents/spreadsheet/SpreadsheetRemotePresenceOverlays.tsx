"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { memo, useMemo, useSyncExternalStore } from "react";

import { spreadsheetRemotePresenceForSheet } from "@mrnote/core/web/lib/collab/spreadsheet-workbook-collab.ts";
import { type WorkbookRemotePresence } from "@/lib/collab/workbook-presence";
import { type SparseAxisSizeResolver } from "@mrnote/core/web/lib/spreadsheet/grid-sizing.ts";
import { workbookRemotePresenceOverlayBoxes } from "./useSpreadsheetSmartArtBoxesAndPresence";

const NO_SUBSCRIPTION = () => () => {};
const NO_REMOTE: readonly WorkbookRemotePresence[] = [];

/**
 * The other members' active cells and ranges over the grid. It hears their
 * moves on its own (the presence controller's `subscribe`), so a peer's range
 * growing under their drag redraws these boxes, not the spreadsheet window.
 */
export const SpreadsheetRemotePresenceOverlays = memo(function SpreadsheetRemotePresenceOverlays({
  remote,
  read,
  subscribe,
  sheetId,
  rowCount,
  columnCount,
  colWidth,
  rowHeight,
  gridRtl,
  editingLabel,
  editingTag,
}: {
  /** The list the window last rendered with (the fallback without `read`). */
  remote?: readonly WorkbookRemotePresence[];
  read?: () => readonly WorkbookRemotePresence[];
  subscribe?: (listener: () => void) => () => void;
  sheetId: string;
  rowCount: number;
  columnCount: number;
  colWidth: SparseAxisSizeResolver;
  rowHeight: SparseAxisSizeResolver;
  gridRtl: boolean;
  /** Appended to a peer's accessible label while they edit their cell. */
  editingLabel: string;
  /** Appended to a peer's name tag while they edit their cell. */
  editingTag: string;
}) {
  const live = useSyncExternalStore(
    subscribe ?? NO_SUBSCRIPTION,
    read ?? (() => remote ?? NO_REMOTE),
    read ?? (() => remote ?? NO_REMOTE),
  );
  const overlays = useMemo(
    () =>
      workbookRemotePresenceOverlayBoxes(
        spreadsheetRemotePresenceForSheet(live, sheetId, rowCount, columnCount),
        { colWidth, gridRtl, rowHeight, sheet: { rowCount, columnCount } },
      ),
    [colWidth, columnCount, gridRtl, live, rowCount, rowHeight, sheetId],
  );
  return (
    <>
      {overlays.map((overlay) => (
        <div
          key={`${overlay.presence.user.id ?? "guest"}:${overlay.presence.clientId}`}
          role="img"
          aria-label={`${overlay.label}${overlay.presence.editing ? editingLabel : ""}`}
          data-testid="spreadsheet-remote-presence"
          data-client-id={overlay.presence.clientId}
          data-user-id={overlay.presence.user.id ?? undefined}
          data-editing={overlay.presence.editing ? "true" : undefined}
          style={{
            position: "absolute",
            ...overlay.range,
            zIndex: 7,
            pointerEvents: "none",
            boxSizing: "border-box",
            border: `2px solid ${overlay.color}`,
            background: overlay.background,
          }}
        >
          <span
            aria-hidden="true"
            style={{
              position: "absolute",
              ...overlay.active,
              boxSizing: "border-box",
              border: `2px solid ${overlay.color}`,
            }}
          />
          <span
            aria-hidden="true"
            style={{
              position: "absolute",
              insetInlineStart: -2,
              bottom: "100%",
              maxWidth: 180,
              overflow: "hidden",
              textOverflow: "ellipsis",
              whiteSpace: "nowrap",
              borderRadius: "4px 4px 0 0",
              background: overlay.color,
              color: "#ffffff",
              padding: "2px 6px",
              fontSize: 11,
              fontWeight: 600,
              lineHeight: 1.4,
            }}
          >
            {overlay.presence.user.name}
            {overlay.presence.editing ? editingTag : ""}
          </span>
        </div>
      ))}
    </>
  );
});
