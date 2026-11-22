"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { lazy, Suspense, useCallback, useEffect, useRef, useState } from "react";
import { useTranslations } from "../../../../host/next-intl";
import type { Node as ProseMirrorNode } from "@tiptap/pm/model";
import type { Editor } from "@tiptap/react";
import {
  Ban,
  Check,
  CheckCheck,
  X,
} from "../../office-ribbon/icons/houseIcons";

import { OfficeTaskPane } from "../../office-ribbon/OfficeTaskPane";
import {
  getTrackedChanges,
  type TrackedChangeEntry,
} from "./native-document/trackChangesExtension";
import { embedSelfName, onlyofficeAllowsReview } from "../../../lib/embed/onlyoffice-frame";
import { revisionResolutionLocked } from "./native-document/documentProtection";
import type { TrackedObjectKind } from "./native-document/trackChanges";
import { pageInterfaceLocale } from "../../../lib/i18n/content-locale";
import { aiEditReviewEnabled, fetchAiEditReview } from "../../../__cut__/mdc1651b936";

interface TrackChangesPanelProps {
  editor: Editor;
  onClose: () => void;
  /** A native document page: the pane also lists its AI changes, one group
   * per AI instruction (`AiEditReviewPane`). */
  pageId?: string | null;
  /** Open on the AI changes tab, on this group (a new `nonce` each ask). */
  aiReviewRequest?: { groupId: string | null; nonce: number } | null;
}

const AiEditReviewPane = lazy(() => import("./AiEditReviewPane"));

type RuntimeEditor = Editor & {
  on?: (event: string, cb: () => void) => void;
  off?: (event: string, cb: () => void) => void;
};

interface TrackedChangesSnapshot {
  /** Exact immutable document generation that produced `changes`. */
  readonly doc: ProseMirrorNode;
  readonly changes: TrackedChangeEntry[];
}

/** Format an epoch-millis timestamp for display, or "" when absent. */
function formatTimestamp(ts: number | null): string {
  if (ts === null) return "";
  try {
    return new Date(ts).toLocaleString(pageInterfaceLocale());
  } catch {
    return "";
  }
}

/**
 * Compact, order-sensitive signature of a change list for cheap equality, so an
 * unrelated edit that does not touch any tracked change doesn't churn the React
 * tree. Includes the live positions + snippet because BOTH shift when an
 * adjacent change is accepted/rejected.
 */
function changesSignature(changes: TrackedChangeEntry[]): string {
  return changes
    .map((c) =>
      [
        c.kind,
        c.moveId ?? "",
        c.moveSegments
          ?.map((segment) => `${segment.role}:${segment.from}:${segment.to}`)
          .join(",") ?? "",
        c.from,
        c.to,
        c.author ?? "",
        c.snippet,
      ].join("|"),
    )
    .join("\n");
}

/**
 * Track-changes review pane (Word "Reviewing Pane" / WPS 修订 parity). Lists
 * EVERY tracked change in the document in document order; each row shows the
 * author, the change kind (insertion / deletion), and a bounded snippet of the
 * affected text, with per-change Accept / Reject buttons plus Accept-all /
 * Reject-all at the top.
 *
 * Source of truth: the SAME pure `listTrackedChanges` derivation the marks use,
 * re-derived from the LIVE document on every editor `update`. The list is NOT
 * denormalized — an insert / accept / reject is reflected without any persisted
 * copy to keep in sync, and the panel UNSUBSCRIBES on unmount (the documented
 * listener-leak / setState-after-unmount class this product guards against).
 *
 * STALE-POSITION STRATEGY: a row's `from` position is only valid against the
 * exact immutable ProseMirror doc that produced the rendered list. The panel
 * stores those together and every visible-row or bulk action re-authenticates
 * `editor.state.doc === snapshot.doc` immediately before focus/dispatch. A doc
 * replacement that happens before React paints the refreshed list therefore
 * makes old buttons zero-mutation rather than letting the same numeric position
 * target a different revision. The next editor `update` installs a fresh bound
 * snapshot; each accepted/rejected change still uses ONE undoable transaction.
 *
 * SECURITY: author + snippet are rendered as React children (TEXT), never via
 * dangerouslySetInnerHTML, and both are sanitized by the pure derivation
 * (control bytes stripped, bounded length), so a forged/imported/AI-authored
 * change can never inject markup through the panel.
 */
/** The Reviewing pane's name for each kind of deleted object. */
const OBJECT_LABEL_KEYS = {
  picture: "trackChanges.picture",
  chart: "trackChanges.chart",
  shape: "trackChanges.shape",
  textBox: "trackChanges.textBox",
  equation: "trackChanges.equation",
  pageBreak: "trackChanges.pageBreak",
  columnBreak: "trackChanges.columnBreak",
  sectionBreak: "trackChanges.sectionBreak",
  horizontalRule: "trackChanges.horizontalRule",
  noteReference: "trackChanges.noteReference",
  field: "trackChanges.field",
} as const satisfies Record<TrackedObjectKind, string>;

export default function TrackChangesPanel({
  editor,
  onClose,
  pageId: pageIdProp,
  aiReviewRequest,
}: TrackChangesPanelProps) {
  const [aiReviewOn] = useState(aiEditReviewEnabled);
  const pageId = aiReviewOn ? pageIdProp : null;
  const t = useTranslations("console-notebooks");
  const [tab, setTab] = useState<"revisions" | "ai">(() => (aiReviewRequest && pageId ? "ai" : "revisions"));
  const [aiOpenCount, setAiOpenCount] = useState<number | null>(null);
  // A new ask for the AI changes turns the pane to them (React's
  // "adjusting state when a prop changes").
  const [seenAiRequest, setSeenAiRequest] = useState(aiReviewRequest);
  if (aiReviewRequest !== seenAiRequest) {
    setSeenAiRequest(aiReviewRequest);
    if (aiReviewRequest && pageId) setTab("ai");
  }
  // The AI tab's count before it is opened, and the tab a document with no
  // revisions but open AI changes opens on.
  useEffect(() => {
    if (!pageId) return;
    let cancelled = false;
    void fetchAiEditReview(pageId)
      .then((groups) => {
        if (cancelled) return;
        const open = groups.reduce(
          (sum, group) => sum + group.changes.filter((c) => c.status === "pending" || c.status === "held").length,
          0,
        );
        setAiOpenCount((current) => current ?? open);
        if (open > 0 && getTrackedChanges(editor.state).length === 0) setTab("ai");
      })
      .catch(() => undefined);
    return () => {
      cancelled = true;
    };
  }, [editor, pageId]);
  const [snapshot, setSnapshot] = useState<TrackedChangesSnapshot>(() => ({
    doc: editor.state.doc,
    changes: getTrackedChanges(editor.state),
  }));
  const [editable, setEditable] = useState(
    () =>
      !editor.isDestroyed &&
      editor.isEditable &&
      !revisionResolutionLocked(editor.state.doc.attrs.documentProtection),
  );
  const changes = snapshot.changes;
  const acceptAllRef = useRef<HTMLButtonElement | null>(null);
  const closeButtonRef = useRef<HTMLButtonElement | null>(null);
  const initialHasChangesRef = useRef(changes.length > 0);

  useEffect(() => {
    let focusFrame = 0;
    const settleFrame = window.requestAnimationFrame(() => {
      focusFrame = window.requestAnimationFrame(() => {
        const target = initialHasChangesRef.current
          ? acceptAllRef.current
          : closeButtonRef.current;
        target?.focus({ preventScroll: true });
      });
    });
    return () => {
      window.cancelAnimationFrame(settleFrame);
      window.cancelAnimationFrame(focusFrame);
    };
  }, []);

  // Live-update: re-derive the change list whenever the document changes, and
  // UNSUBSCRIBE on unmount. Every new doc identity must refresh the action
  // authority even when the visible change signature stayed identical; in that
  // case we retain the previous array so unrelated edits don't churn row keys.
  useEffect(() => {
    const runtime = editor as RuntimeEditor;
    const refresh = () => {
      setEditable(
        !editor.isDestroyed &&
          editor.isEditable &&
          !revisionResolutionLocked(editor.state.doc.attrs.documentProtection),
      );
      const state = editor.state;
      const nextChanges = getTrackedChanges(state);
      setSnapshot((previous) => {
        if (previous.doc === state.doc) return previous;
        return {
          doc: state.doc,
          changes:
            changesSignature(previous.changes) === changesSignature(nextChanges)
              ? previous.changes
              : nextChanges,
        };
      });
    };
    refresh();
    // The host turns editing off and on again without a document update
    // (a reconnect, a load): the view's contenteditable is what changes, so
    // the pane follows it rather than keep a stale read-only state.
    let editableObserver: MutationObserver | null = null;
    try {
      const dom = editor.isDestroyed ? null : editor.view.dom;
      if (dom && typeof MutationObserver !== "undefined") {
        editableObserver = new MutationObserver(refresh);
        editableObserver.observe(dom, {
          attributes: true,
          attributeFilter: ["contenteditable"],
        });
      }
    } catch {
      editableObserver = null;
    }
    if (typeof runtime.on !== "function" || typeof runtime.off !== "function") {
      return () => editableObserver?.disconnect();
    }
    runtime.on("update", refresh);
    return () => {
      editableObserver?.disconnect();
      runtime.off?.("update", refresh);
    };
  }, [editor]);

  const renderedSnapshotIsCurrent = useCallback(
    () => editor.state.doc === snapshot.doc,
    [editor, snapshot.doc],
  );

  // Click a row → scroll to the change only while the rendered doc generation
  // remains current. Within that immutable doc the cached range is authoritative.
  const focusChange = useCallback(
    (change: TrackedChangeEntry) => {
      if (!renderedSnapshotIsCurrent()) return;
      const size = editor.state.doc.content.size;
      const target = Math.min(Math.max(change.from, 0), size);
      if (change.object && editor.state.doc.nodeAt(target)) {
        editor.chain().focus().setNodeSelection(target).scrollIntoView().run();
        return;
      }
      editor.chain().focus().setTextSelection(target).scrollIntoView().run();
    },
    [editor, renderedSnapshotIsCurrent],
  );

  // Accept / reject a single change. The exact-doc gate prevents a stale row
  // from resolving its numeric position against a different current revision.
  const acceptOne = useCallback(
    (change: TrackedChangeEntry) => {
      if (
        editor.isDestroyed ||
        !editor.isEditable ||
        revisionResolutionLocked(editor.state.doc.attrs.documentProtection) ||
        !renderedSnapshotIsCurrent()
      ) {
        return;
      }
      // An embedded ONLYOFFICE config's reviewGroups: another group's change stays.
      if (!onlyofficeAllowsReview(change.author, embedSelfName())) return;
      editor.commands.acceptChangeAtPos(change.from);
    },
    [editor, renderedSnapshotIsCurrent],
  );

  const rejectOne = useCallback(
    (change: TrackedChangeEntry) => {
      if (
        editor.isDestroyed ||
        !editor.isEditable ||
        revisionResolutionLocked(editor.state.doc.attrs.documentProtection) ||
        !renderedSnapshotIsCurrent()
      ) {
        return;
      }
      if (!onlyofficeAllowsReview(change.author, embedSelfName())) return;
      editor.commands.rejectChangeAtPos(change.from);
    },
    [editor, renderedSnapshotIsCurrent],
  );

  const acceptAll = useCallback(() => {
    if (
      editor.isDestroyed ||
      !editor.isEditable ||
      revisionResolutionLocked(editor.state.doc.attrs.documentProtection) ||
      !renderedSnapshotIsCurrent()
    ) {
      return;
    }
    // reviewGroups: all at once only when every change is this person's to resolve.
    if (!snapshot.changes.every((change) => onlyofficeAllowsReview(change.author, embedSelfName()))) return;
    editor.commands.acceptAllChanges();
  }, [editor, renderedSnapshotIsCurrent, snapshot.changes]);

  const rejectAll = useCallback(() => {
    if (
      editor.isDestroyed ||
      !editor.isEditable ||
      revisionResolutionLocked(editor.state.doc.attrs.documentProtection) ||
      !renderedSnapshotIsCurrent()
    ) {
      return;
    }
    if (!snapshot.changes.every((change) => onlyofficeAllowsReview(change.author, embedSelfName()))) return;
    editor.commands.rejectAllChanges();
  }, [editor, renderedSnapshotIsCurrent, snapshot.changes]);

  // Word's Reviewing Pane is a docked task pane (Review > Tracking >
  // Reviewing Pane, "Track changes in Word", Microsoft Support): the shared
  // Office task pane frame, placed by NoteEditor's side-pane column.
  return (
    <OfficeTaskPane
      title={t("trackChanges.title")}
      closeLabel={t("trackChanges.close")}
      onClose={onClose}
      testIdPrefix="track-changes"
      testId="track-changes-panel"
      closeRef={closeButtonRef}
      dockRight={false}
      className="ai-panel track-changes-panel"
      headerActions={
        <span className="track-changes-panel-count">{tab === "ai" ? aiOpenCount ?? "" : changes.length}</span>
      }
    >
      {pageId ? (
        <div className="track-changes-panel-tabs" role="tablist" aria-label={t("trackChanges.title")}>
          <button
            type="button"
            role="tab"
            aria-selected={tab === "revisions"}
            className="track-changes-panel-tab"
            data-testid="track-changes-tab-revisions"
            onClick={() => setTab("revisions")}
          >
            {t("aiReview.revisionsTab")}
            <span className="track-changes-panel-tab-count">{changes.length}</span>
          </button>
          {null}
        </div>
      ) : null}
      {pageId && tab === "ai" ? (
        null
      ) : (
      <>

      <div className="track-changes-panel-bulk">
        <button
          ref={acceptAllRef}
          type="button"
          className="track-changes-panel-bulk-accept"
          data-testid="track-changes-accept-all"
          disabled={!editable || changes.length === 0}
          onClick={acceptAll}
        >
          <CheckCheck size={13} />
          {t("trackChanges.acceptAll")}
        </button>
        <button
          type="button"
          className="track-changes-panel-bulk-reject"
          data-testid="track-changes-reject-all"
          disabled={!editable || changes.length === 0}
          onClick={rejectAll}
        >
          <Ban size={13} />
          {t("trackChanges.rejectAll")}
        </button>
      </div>

      <div
        className="track-changes-panel-body"
        data-testid="track-changes-body"
      >
        {changes.length === 0 ? (
          <div
            className="track-changes-panel-empty"
            data-testid="track-changes-empty"
          >
            {t("trackChanges.empty")}
          </div>
        ) : (
          <ul
            className="track-changes-panel-list"
            data-testid="track-changes-list"
          >
            {changes.map((change) => (
              <li
                key={`${change.kind}-${change.moveId ?? change.formatScope ?? "content"}-${change.nodePos ?? change.from}-${change.to}`}
                className={`track-changes-panel-item track-changes-panel-item--${change.kind}`}
                data-testid="track-changes-item"
                data-kind={change.kind}
                data-format-scope={change.formatScope}
                data-from={change.from}
              >
                <button
                  type="button"
                  className="track-changes-panel-item-jump"
                  data-testid="track-changes-item-jump"
                  onClick={() => focusChange(change)}
                  aria-label={t("trackChanges.goToChange")}
                  title={t("trackChanges.goToChange")}
                >
                  <span className="track-changes-panel-item-meta">
                    <span
                      className={`track-changes-panel-item-kind track-changes-panel-item-kind--${change.kind}`}
                    >
                      {change.kind === "move"
                        ? `${t("trackChanges.deletion")} → ${t("trackChanges.insertion")}`
                        : change.kind === "insertion"
                          ? t("trackChanges.insertion")
                          : change.kind === "deletion"
                            ? t("trackChanges.deletion")
                            : t("trackChanges.formatting")}
                    </span>
                    <span className="track-changes-panel-item-author">
                      {change.author || t("trackChanges.anonymous")}
                    </span>
                    {change.timestamp !== null ? (
                      <span className="track-changes-panel-item-time">
                        {formatTimestamp(change.timestamp)}
                      </span>
                    ) : null}
                  </span>
                  {/* Rendered as TEXT (React children) — never HTML. */}
                  <span
                    className="track-changes-panel-item-snippet"
                    data-testid="track-changes-item-snippet"
                  >
                    {change.object
                      ? `${t(OBJECT_LABEL_KEYS[change.object])}${change.snippet ? `: ${change.snippet}` : ""}`
                      : change.formatScope === "table"
                        ? `${t("trackChanges.formattedTable")}${change.snippet ? `: ${change.snippet}` : ""}`
                        : change.snippet || t("trackChanges.emptySnippet")}
                    {change.truncated ? "…" : ""}
                  </span>
                </button>

                <div className="track-changes-panel-item-actions">
                  <button
                    type="button"
                    className="track-changes-panel-item-accept"
                    data-testid="track-changes-item-accept"
                    disabled={!editable}
                    onClick={() => acceptOne(change)}
                    aria-label={t("trackChanges.accept")}
                    title={t("trackChanges.accept")}
                  >
                    <Check size={13} />
                    {t("trackChanges.accept")}
                  </button>
                  <button
                    type="button"
                    className="track-changes-panel-item-reject"
                    data-testid="track-changes-item-reject"
                    disabled={!editable}
                    onClick={() => rejectOne(change)}
                    aria-label={t("trackChanges.reject")}
                    title={t("trackChanges.reject")}
                  >
                    <X size={13} />
                    {t("trackChanges.reject")}
                  </button>
                </div>
              </li>
            ))}
          </ul>
        )}
      </div>
      </>
      )}
    </OfficeTaskPane>
  );
}
