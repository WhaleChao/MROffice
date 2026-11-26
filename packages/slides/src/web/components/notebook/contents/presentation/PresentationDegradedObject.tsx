"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { DegradedObject } from "@mrnote/core/web/lib/presentation-object-degraded.ts";

type Translate = (key: string, values?: Record<string, string | number>) => string;

/** The selected object's note: what the page could not do with it, and why. */
function DegradedObjectNote({ t, record }: { t: Translate; record: DegradedObject }) {
  return (
    <div
      role="note"
      data-testid="presentation-object-degraded-note"
      data-stage={record.stage}
      style={{
        position: "absolute",
        left: 0,
        top: "100%",
        marginTop: 4,
        zIndex: 4,
        maxWidth: 360,
        padding: "4px 8px",
        borderRadius: 4,
        fontSize: 12,
        lineHeight: 1.4,
        background: "var(--popover, #fff)",
        color: "var(--popover-foreground, #222)",
        border: "1px solid var(--border, #d0d0d0)",
        boxShadow: "0 1px 3px rgba(0,0,0,0.15)",
        pointerEvents: "none",
        whiteSpace: "normal",
      }}
    >
      <div>
        {t(
          record.stage === "placeholder"
            ? "presentation.objectDegraded.placeholderNote"
            : "presentation.objectDegraded.readOnlyNote",
        )}
      </div>
      {record.reason || record.path ? (
        <div style={{ marginTop: 2, opacity: 0.75, overflowWrap: "anywhere" }}>
          {t("presentation.objectDegraded.reason", {
            path: record.path || "-",
            reason: record.reason || "-",
          })}
        </div>
      ) : null}
    </div>
  );
}

/**
 * Over an object the page keeps read-only because the validator refused
 * something in it (TRIAGE 152): a placeholder frame's words, always; and,
 * once selected, the note with the reason. The frame itself (dashed, no fill)
 * is the block's own shape, drawn wherever the slide is.
 */
export function PresentationDegradedObjectOverlay({
  t,
  record,
  isSelected,
}: {
  t: Translate;
  record: DegradedObject;
  isSelected: boolean;
}) {
  return (
    <>
      {record.stage === "placeholder" ? (
        <div
          data-testid="presentation-object-placeholder"
          aria-label={t("presentation.objectDegraded.placeholder")}
          style={{
            position: "absolute",
            inset: 0,
            display: "flex",
            alignItems: "center",
            justifyContent: "center",
            padding: 4,
            overflow: "hidden",
            textAlign: "center",
            fontSize: 12,
            lineHeight: 1.3,
            color: "#595959",
            background: "rgba(127,127,127,0.06)",
            pointerEvents: "none",
            userSelect: "none",
          }}
        >
          {t("presentation.objectDegraded.placeholder")}
        </div>
      ) : null}
      {isSelected ? <DegradedObjectNote t={t} record={record} /> : null}
    </>
  );
}
