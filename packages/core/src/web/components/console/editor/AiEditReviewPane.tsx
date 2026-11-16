"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Review > Reviewing Pane > AI changes: each AI instruction's changes to the
 * document, the latest instruction first, with the instruction's words.
 * Inside a group, first the sections it touched, then each change, accepted
 * or rejected on its own or the group whole (Word's Reviewing Pane and its
 * Accept / Reject, narrowed to one reviewer: "Track changes in Word",
 * Microsoft Support).
 *
 * Key fields (amounts, numbers, dates, ratios, names, links, citations) are
 * flagged on their change and never accepted by "Accept group": the group's
 * other changes are accepted and these stay open, listed, until each is
 * accepted or rejected, or the person presses the second, explicit "Accept
 * these N key-field changes too". Leaving them open loses nothing and is
 * undone by nothing, where a confirmation dialog is pressed through out of
 * habit. A polish instruction's key-field changes were not made at all
 * (held); each can be applied here.
 *
 * Rejecting puts back exactly what the place held before
 * (`aiEditReview.ts`); the decision is recorded first, so of two windows
 * pressing Reject on the same change only one writes it.
 */
import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { useTranslations } from "../../../../host/next-intl";
import type { Editor } from "@tiptap/react";

import {
  AlertTriangle,
  Ban,
  Check,
  CheckCheck,
  ChevronDown,
  ChevronRight,
  ShieldCheck,
  Undo2,
  X,
} from "../../office-ribbon/icons/houseIcons";
import {
  decideAiEditReview,
  fetchAiEditReview,
  type AiReviewChange,
  type AiReviewDecision,
  type AiReviewGroup,
  type AiReviewKeyKind,
} from "../../../__cut__/mdc1651b936";
import { pageInterfaceLocale } from "../../../lib/i18n/content-locale";
import {
  aiChangeTransaction,
  isKeyFieldChange,
  locateChange,
  type AiChangeLocation,
} from "../../../__cut__/m0fe6c7f917";
import {
  aiEditReviewHighlightsPlugin,
  aiEditReviewHighlightsKey,
  setAiReviewHighlights,
} from "../../../__cut__/mfb36e242cf";

interface AiEditReviewPaneProps {
  editor: Editor;
  pageId: string;
  editable: boolean;
  /** The group to open (the run card's "逐条审阅"); changes with `nonce`. */
  focusGroupId?: string | null;
  focusNonce?: number;
  /** How many changes wait for a decision (the tab's count). */
  onOpenCount?: (count: number) => void;
}

type RuntimeEditor = Editor & {
  on?: (event: string, cb: () => void) => void;
  off?: (event: string, cb: () => void) => void;
};

const POLL_MS = 6000;

interface DecideOptions {
  retry?: boolean;
  /** The edit the decision makes: a reject's is the place as it was
   * ("before"), an apply's the AI's ("after"); taking a decision back
   * makes the other one. */
  write?: "before" | "after";
}
const OPEN = new Set(["pending", "held"]);
const KEY_KIND_ORDER: AiReviewKeyKind[] = ["amount", "number", "date", "ratio", "entity", "link", "citation"];

function formatTime(value: string | null): string {
  if (!value) return "";
  try {
    return new Date(value).toLocaleString(pageInterfaceLocale(), {
      month: "numeric",
      day: "numeric",
      hour: "2-digit",
      minute: "2-digit",
    });
  } catch {
    return "";
  }
}

function clip(text: string, max: number): string {
  const chars = Array.from(text);
  return chars.length > max ? `${chars.slice(0, max).join("")}…` : text;
}

function tail(text: string, max: number): string {
  const chars = Array.from(text);
  return chars.length > max ? `…${chars.slice(chars.length - max).join("")}` : text;
}

/** The document's place of each change, worked out once per document. */
function useLocations(editor: Editor, groups: AiReviewGroup[]): Map<string, AiChangeLocation> {
  const [doc, setDoc] = useState(() => editor.state.doc);
  useEffect(() => {
    const runtime = editor as RuntimeEditor;
    let timer: ReturnType<typeof setTimeout> | null = null;
    const update = () => {
      if (timer) clearTimeout(timer);
      timer = setTimeout(() => {
        if (!editor.isDestroyed) setDoc(editor.state.doc);
      }, 200);
    };
    runtime.on?.("update", update);
    return () => {
      if (timer) clearTimeout(timer);
      runtime.off?.("update", update);
    };
  }, [editor]);
  return useMemo(() => {
    const found = new Map<string, AiChangeLocation>();
    for (const group of groups) {
      for (const change of group.changes) {
        try {
          found.set(change.id, locateChange(doc, change));
        } catch {
          found.set(change.id, { state: "missing" });
        }
      }
    }
    return found;
  }, [doc, groups]);
}

export default function AiEditReviewPane({
  editor,
  pageId,
  editable,
  focusGroupId,
  focusNonce,
  onOpenCount,
}: AiEditReviewPaneProps) {
  const t = useTranslations("console-notebooks");
  const [groups, setGroups] = useState<AiReviewGroup[] | null>(null);
  const [loadFailed, setLoadFailed] = useState(false);
  const [expanded, setExpanded] = useState<Record<string, boolean>>({});
  const [busy, setBusy] = useState<string | null>(null);
  const [notice, setNotice] = useState<{ group: string; text: string; keyIds?: string[] } | null>(null);
  const [focusedChange, setFocusedChange] = useState<string | null>(null);
  const loadSeq = useRef(0);

  const load = useCallback(async () => {
    const seq = ++loadSeq.current;
    try {
      const next = await fetchAiEditReview(pageId);
      if (seq !== loadSeq.current) return;
      setGroups(next);
      setLoadFailed(false);
    } catch {
      if (seq !== loadSeq.current) return;
      setLoadFailed(true);
    }
  }, [pageId]);

  // Read now, again every few seconds while the pane is open and the tab is
  // seen (another window's decisions, a new instruction's changes), and when
  // the window comes back to the front.
  useEffect(() => {
    // eslint-disable-next-line react-hooks/set-state-in-effect -- a read whose answer arrives later
    void load();
    const timer = setInterval(() => {
      if (typeof document === "undefined" || document.visibilityState === "visible") void load();
    }, POLL_MS);
    const onFocus = () => void load();
    window.addEventListener("focus", onFocus);
    return () => {
      clearInterval(timer);
      window.removeEventListener("focus", onFocus);
    };
  }, [load]);

  const locations = useLocations(editor, groups ?? []);

  // The latest group open at first; the one asked for when asked (each new
  // ask opens it again).
  const [seenFocus, setSeenFocus] = useState<string | null | undefined>(undefined);
  const focusKey = focusGroupId ? `${focusGroupId}:${focusNonce ?? 0}` : null;
  if (groups?.length && focusKey !== seenFocus) {
    const asked = focusGroupId && groups.some((group) => group.id === focusGroupId) ? focusGroupId : null;
    if (asked || !focusKey) {
      setSeenFocus(focusKey);
      setExpanded((current) =>
        asked ? { ...current, [asked]: true } : Object.keys(current).length ? current : { [groups[0].id]: true },
      );
    }
  }

  const openCount = useMemo(
    () => (groups ?? []).reduce((sum, group) => sum + group.changes.filter((c) => OPEN.has(c.status)).length, 0),
    [groups],
  );
  useEffect(() => {
    onOpenCount?.(openCount);
  }, [onOpenCount, openCount]);

  // The open groups' places shaded in the document while the pane is open.
  useEffect(() => {
    if (editor.isDestroyed) return;
    const plugin = aiEditReviewHighlightsPlugin();
    editor.registerPlugin(plugin);
    return () => {
      if (!editor.isDestroyed) editor.unregisterPlugin(aiEditReviewHighlightsKey);
    };
  }, [editor]);
  useEffect(() => {
    if (editor.isDestroyed || !aiEditReviewHighlightsKey.getState(editor.state)) return;
    const items = (groups ?? [])
      .filter((group) => expanded[group.id])
      .flatMap((group) => group.changes)
      .flatMap((change) => {
        const location = locations.get(change.id);
        const shown = change.status === "pending" ? location?.state === "after" : change.status === "held" && location?.state === "before";
        if (!shown || location?.from === undefined || location.to === undefined) return [];
        return [{ id: change.id, from: location.from, to: location.to, key: isKeyFieldChange(change), focused: change.id === focusedChange }];
      });
    editor.view.dispatch(setAiReviewHighlights(editor.state.tr, items));
  }, [editor, groups, expanded, locations, focusedChange]);

  const goTo = useCallback(
    (change: AiReviewChange) => {
      setFocusedChange(change.id);
      const location = locateChange(editor.state.doc, change);
      if (location.from === undefined) return;
      editor.chain().setTextSelection({ from: location.from, to: location.to ?? location.from }).scrollIntoView().run();
    },
    [editor],
  );

  const replaceChanges = useCallback((updated: AiReviewChange[]) => {
    const byId = new Map(updated.map((change) => [change.id, change]));
    setGroups((current) =>
      (current ?? []).map((group) => ({
        ...group,
        changes: group.changes.map((change) => {
          const next = byId.get(change.id);
          return next ? { ...change, status: next.status, decided_at: next.decided_at } : change;
        }),
      })),
    );
  }, []);

  /**
   * Decide `changes`; for a reject or an apply, write the place's edit for
   * each one this window won (another window that won it writes it there).
   * A change whose place is not as the decision needs is not decided.
   */
  const decide = useCallback(
    async (groupId: string, changes: AiReviewChange[], decision: AiReviewDecision, options: DecideOptions = {}) => {
      if (!changes.length || !editable) return { done: 0, skipped: 0 };
      const target = options.write ?? (decision === "reject" ? "before" : decision === "apply" ? "after" : null);
      const request = { retry: options.retry };
      // Only what can be made now: the AI's words for a reject, the earlier
      // ones for an apply; a place already as asked is only recorded.
      const writable: AiReviewChange[] = [];
      const recordOnly: AiReviewChange[] = [];
      let skipped = 0;
      for (const change of changes) {
        if (!target) {
          recordOnly.push(change);
          continue;
        }
        const state = locateChange(editor.state.doc, change).state;
        if (state === target) recordOnly.push(change);
        else if ((target === "before" && state === "after") || (target === "after" && state === "before")) {
          if (change.unrestorable) skipped += 1;
          else writable.push(change);
        } else skipped += 1;
      }
      setBusy(groupId);
      try {
        const results = [
          ...(recordOnly.length ? await decideAiEditReview(pageId, recordOnly.map((c) => c.id), decision, request) : []),
          ...(writable.length ? await decideAiEditReview(pageId, writable.map((c) => c.id), decision, request) : []),
        ];
        const won = new Set(results.filter((result) => result.won).map((result) => result.id));
        const failed: AiReviewChange[] = [];
        if (target) {
          for (const change of writable) {
            if (!won.has(change.id)) continue;
            const tr = aiChangeTransaction(editor.state, change, target);
            if (!tr) {
              failed.push(change);
              continue;
            }
            editor.view.dispatch(tr);
          }
        }
        if (failed.length) {
          // The place changed while the decision was on its way: the
          // decision is taken back, nothing was written.
          const undo: AiReviewDecision | null =
            decision === "reject" ? "reopen" : decision === "apply" ? "reopen_held" : null;
          const back = undo
            ? await decideAiEditReview(pageId, failed.map((change) => change.id), undo, { retry: true }).catch(() => [])
            : [];
          results.push(...back.map((result) => ({ ...result, won: false })));
        }
        replaceChanges(results);
        return { done: results.filter((r) => r.won).length, skipped: skipped + failed.length };
      } catch {
        setNotice({ group: groupId, text: t("aiReview.failed") });
        return { done: 0, skipped };
      } finally {
        setBusy(null);
        void load();
      }
    },
    [editable, editor, load, pageId, replaceChanges, t],
  );

  const acceptGroup = useCallback(
    async (group: AiReviewGroup) => {
      const open = group.changes.filter((change) => change.status === "pending");
      const plain = open.filter((change) => !isKeyFieldChange(change));
      const key = open.filter(isKeyFieldChange);
      if (plain.length) await decide(group.id, plain, "accept");
      setNotice(
        key.length
          ? {
              group: group.id,
              text: `${plain.length ? t("aiReview.acceptedSome", { count: plain.length }) : ""}${t("aiReview.keyLeft", { count: key.length })}`,
              keyIds: key.map((change) => change.id),
            }
          : null,
      );
    },
    [decide, t],
  );

  const acceptKeyToo = useCallback(
    async (group: AiReviewGroup, ids: string[]) => {
      const key = group.changes.filter((change) => ids.includes(change.id) && change.status === "pending");
      await decide(group.id, key, "accept");
      setNotice(null);
    },
    [decide],
  );

  const rejectGroup = useCallback(
    async (group: AiReviewGroup) => {
      // Last place first: an earlier place is found again after each edit.
      const open = group.changes.filter((change) => change.status === "pending").reverse();
      const { skipped } = await decide(group.id, open, "reject");
      setNotice(skipped ? { group: group.id, text: t("aiReview.skippedEdited", { count: skipped }) } : null);
    },
    [decide, t],
  );

  if (groups === null) {
    return (
      null
    );
  }
  if (groups.length === 0) {
    return (
      null
    );
  }

  return (
    null
  );
}

interface ChangeRowProps {
  change: AiReviewChange;
  location: AiChangeLocation;
  editable: boolean;
  focused: boolean;
  onGoTo: () => void;
  onDecide: (decision: AiReviewDecision, options?: DecideOptions) => void;
}

function ChangeRow({ change, location, editable, focused, onGoTo, onDecide }: ChangeRowProps) {
  const t = useTranslations("console-notebooks");
  const key = isKeyFieldChange(change);
  const kinds = KEY_KIND_ORDER.filter((kind) => change.key_kinds.includes(kind) || change.key_fields?.some((f) => f.kind === kind));
  const before = change.before_text ?? "";
  const after = change.after_text ?? "";
  const state = location.state;
  const status = change.status;
  let stateNote: string | null = null;
  if (change.unrestorable && OPEN.has(status)) stateNote = t("aiReview.state.unrestorable");
  else if (status === "pending" && state === "edited") stateNote = t("aiReview.state.edited");
  else if (status === "pending" && state === "missing") stateNote = t("aiReview.state.missing");
  else if (status === "pending" && state === "before") stateNote = t("aiReview.state.reverted");
  else if (status === "held" && (state === "edited" || state === "missing")) stateNote = t("aiReview.state.heldMoved");
  const canReject = status === "pending" && (state === "after" || state === "before") && !change.unrestorable;
  const canApply = status === "held" && (state === "before" || state === "after") && !change.unrestorable;
  return (
    null
  );
}
