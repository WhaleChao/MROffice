"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { EditorLoadingState } from "../../ui/editor-loading-state";
import {
  useCallback,
  useEffect,
  useLayoutEffect,
  useRef,
  useState,
} from "react";
import type { Editor } from "@tiptap/react";
import { useTranslations } from "../../../../host/next-intl";
import { Loader2 } from "lucide-react";
import { ArrowRightLeft } from "../../office-ribbon/icons/houseIcons";

import {
  OfficeTaskPane,
  officeTaskPaneParts,
} from "../../office-ribbon/OfficeTaskPane";

import {
  getPageVersion,
  listPageVersions,
  type PageVersionDetail,
  type PageVersionItem,
} from "../../../lib/notebook-sdk";
// The redline (word-level LCS, Word's Compare) is computed on the server
// (engine `tools`, `document.compare`); this panel picks the version, sends
// both documents and loads the answer. Combine is the same request with its
// options (`lib/document/combine.ts`).
import type { DocumentCombineFormatSource, DocumentCombineOutcome } from "@/lib/document/combine";
import type {
  DocumentCompareOutcome,
  DocumentCompareRefusal,
  PMNodeJSON,
} from "@/lib/document/compare";
import type { CompareResultSources } from "../../../lib/document/sourceDocuments";
import { useSelfCollabIdentity } from "../../../lib/collab/self-identity";
import { TOOLS_COMBINE_MAX_AUTHOR_LENGTH } from "../../../lib/engine-host/tools-types";
import {
  callEngine,
  EngineRefusedError,
  engineErrorMessageKey,
} from "../../../lib/engine-host/client";
import { documentPageContentSchema } from "../../../lib/native-page-schemas";
import { TRACK_CHANGES_SKIP_META } from "./native-document/trackChangesExtension";
import { flattenPagedDocumentForStorage } from "./native-document/paged/serialization";
import { wrapBlocksIntoPages } from "./native-document/paged/pageBlockExtension";
import { pageInterfaceLocale } from "../../../lib/i18n/content-locale";

/** The panel's word for each reason the server gives for not comparing. */
const COMPARE_REFUSAL_MESSAGE: Record<DocumentCompareRefusal, string> = {
  existing_changes: "compare.unresolvedChanges",
  structure_changed: "compare.structureChanged",
  too_large: "compare.tooLarge",
};

const MAX_COMPARE_JSON_NODES = 300_000;
const MAX_COMPARE_JSON_DEPTH = 256;
const MAX_COMPARE_JSON_ARRAY_LENGTH = 100_000;
const MAX_COMPARE_JSON_OBJECT_KEYS = 10_000;
const MAX_COMPARE_STRING_LENGTH = 1_048_576;
const MAX_COMPARE_TOTAL_STRING_LENGTH = 8_000_000;
const MAX_COMPARE_DOCUMENT_NODES = 120_000;
const MAX_COMPARE_DOCUMENT_DEPTH = 40;

function isJsonObject(value: unknown): value is Record<string, unknown> {
  if (value === null || typeof value !== "object" || Array.isArray(value)) {
    return false;
  }
  const prototype = Object.getPrototypeOf(value);
  return prototype === Object.prototype || prototype === null;
}

function hasOwn(value: object, key: string): boolean {
  return Object.prototype.hasOwnProperty.call(value, key);
}

/** Iterative whole-tree budget: malformed or adversarial JSON never recurses. */
function isBoundedJsonTree(root: unknown): boolean {
  const stack: Array<{ value: unknown; depth: number }> = [
    { value: root, depth: 0 },
  ];
  const seen = new WeakSet<object>();
  let nodes = 0;
  let totalStringLength = 0;

  const addString = (value: string): boolean => {
    if (value.length > MAX_COMPARE_STRING_LENGTH) return false;
    totalStringLength += value.length;
    return totalStringLength <= MAX_COMPARE_TOTAL_STRING_LENGTH;
  };

  while (stack.length > 0) {
    const entry = stack.pop();
    if (!entry || entry.depth > MAX_COMPARE_JSON_DEPTH) return false;
    nodes += 1;
    if (nodes > MAX_COMPARE_JSON_NODES) return false;

    const value = entry.value;
    if (typeof value === "string") {
      if (!addString(value)) return false;
      continue;
    }
    if (value === null || typeof value === "boolean") continue;
    if (typeof value === "number") {
      if (
        !Number.isFinite(value) ||
        (Number.isInteger(value) && !Number.isSafeInteger(value))
      ) {
        return false;
      }
      continue;
    }
    if (Array.isArray(value)) {
      if (seen.has(value) || value.length > MAX_COMPARE_JSON_ARRAY_LENGTH) {
        return false;
      }
      seen.add(value);
      for (let index = value.length - 1; index >= 0; index -= 1) {
        stack.push({ value: value[index], depth: entry.depth + 1 });
      }
      continue;
    }
    if (!isJsonObject(value) || seen.has(value)) return false;
    seen.add(value);
    const entries = Object.entries(value);
    if (entries.length > MAX_COMPARE_JSON_OBJECT_KEYS) return false;
    for (let index = entries.length - 1; index >= 0; index -= 1) {
      const [key, child] = entries[index];
      if (!addString(key)) return false;
      stack.push({ value: child, depth: entry.depth + 1 });
    }
  }
  return true;
}

function isComparableDoc(value: unknown): value is PMNodeJSON {
  if (!isBoundedJsonTree(value) || !isJsonObject(value)) return false;
  if (value.type !== "doc" || !Array.isArray(value.content)) return false;

  const stack: Array<{ value: unknown; depth: number }> = [{ value, depth: 0 }];
  let nodeCount = 0;

  while (stack.length > 0) {
    const entry = stack.pop();
    if (!entry || entry.depth > MAX_COMPARE_DOCUMENT_DEPTH) return false;
    if (!isJsonObject(entry.value)) return false;
    const node = entry.value;
    nodeCount += 1;
    if (nodeCount > MAX_COMPARE_DOCUMENT_NODES) return false;

    if (typeof node.type !== "string" || node.type.length === 0) {
      return false;
    }
    if (hasOwn(node, "attrs") && !isJsonObject(node.attrs)) return false;

    if (hasOwn(node, "marks")) {
      if (!Array.isArray(node.marks)) return false;
      for (const markValue of node.marks) {
        if (
          !isJsonObject(markValue) ||
          typeof markValue.type !== "string" ||
          markValue.type.length === 0 ||
          (hasOwn(markValue, "attrs") && !isJsonObject(markValue.attrs))
        ) {
          return false;
        }
      }
    }

    const hasContent = hasOwn(node, "content");
    if (hasContent && !Array.isArray(node.content)) return false;
    if (node.type === "text") {
      if (typeof node.text !== "string" || hasContent) return false;
    } else if (hasOwn(node, "text")) {
      return false;
    }

    if (Array.isArray(node.content)) {
      for (let index = node.content.length - 1; index >= 0; index -= 1) {
        stack.push({ value: node.content[index], depth: entry.depth + 1 });
      }
    }
  }
  // The native schema deliberately allows additive/custom node fields. Keep
  // those forward-compatible while applying the product's current semantic
  // checks (including bounded/safe image and link attributes) after the
  // iterative resource/shape pass above.
  return documentPageContentSchema.safeParse(value).success;
}

/**
 * The live document as a saved version holds it: the editor's JSON is a
 * live object, with attributes left `undefined` (a shape's unset outline
 * cap, lock, ink fields) and one default value object shared by several
 * nodes (a table's cells), and the budget pass above refuses both, as it
 * must for a stored value. Read through JSON, as the save writes it, the
 * same document compares. Comparing any version with a page that held a
 * shape or a table inside a table said only "Compare failed, try again"
 * (live Word audit WUA 085: v182 and v6 alike). Null when it cannot be
 * written as JSON at all.
 */
function liveDocumentAsStored(editor: Editor): unknown {
  try {
    return JSON.parse(JSON.stringify(editor.getJSON())) as unknown;
  } catch {
    return null;
  }
}

/**
 * A version may enter the live document only when both its persisted editor
 * type and its JSON root prove that it is a document snapshot. Legacy rows
 * without a recorded page type are inferable only from an explicit doc root;
 * snapshot_text is display-only and must never be promoted to editable JSON.
 */
function comparableDocumentSnapshot(
  editor: Editor,
  version: PageVersionDetail,
): PMNodeJSON | null {
  if (version.page_type !== "document" && version.page_type !== null) {
    return null;
  }

  const snapshot = version.snapshot_json;
  if (!isComparableDoc(snapshot)) {
    return null;
  }
  return comparableForm(editor, snapshot);
}









function comparableForm(editor: Editor, doc: PMNodeJSON): PMNodeJSON | null {
  try {
    const schema = (editor as Partial<Editor>).schema;
    let hydrated = doc;
    if (schema?.nodeFromJSON) {
      const paged =
        schema.nodes.pageBlock &&
        Array.isArray(doc.content) &&
        !doc.content.some((node) => node?.type === "pageBlock")
          ? { ...doc, content: wrapBlocksIntoPages(doc.content as never) }
          : doc;
      hydrated = schema.nodeFromJSON(paged).toJSON() as PMNodeJSON;
    }
    const stored = flattenPagedDocumentForStorage(hydrated);
    // The layout caches a save stamps on the root describe the pages, not the
    // document; and a stored object may come back with its keys reordered.
    const attrs = { ...((stored.attrs as Record<string, unknown> | undefined) ?? {}) };
    for (const key of LAYOUT_CACHE_ATTRS) delete attrs[key];
    return canonicalKeyOrder({ ...stored, attrs }) as PMNodeJSON;
  } catch {
    return null;
  }
}

const LAYOUT_CACHE_ATTRS = ["pageBlockSizes", "pageBlockLayoutKey", "pageRegionLayoutReceipt"];

/** The same value with every object's keys in one order. */
function canonicalKeyOrder(value: unknown): unknown {
  if (Array.isArray(value)) return value.map(canonicalKeyOrder);
  if (value === null || typeof value !== "object") return value;
  return Object.fromEntries(
    Object.keys(value as Record<string, unknown>)
      .sort()
      .map((key) => [key, canonicalKeyOrder((value as Record<string, unknown>)[key])]),
  );
}

/** A compare or combine result view's id (the Show Source Documents state). */
function newResultId(): string {
  return typeof crypto !== "undefined" && typeof crypto.randomUUID === "function"
    ? crypto.randomUUID()
    : `result-${Date.now()}-${Math.random().toString(36).slice(2)}`;
}

/** A combine answer that is not a refusal (refusals share Compare's shape). */
function isCombineOutcome(
  answer: DocumentCompareOutcome | DocumentCombineOutcome,
): answer is Extract<
  DocumentCombineOutcome,
  { outcome: "combined" | "needs_format_choice" }
> {
  return answer.outcome === "combined" || answer.outcome === "needs_format_choice";
}

/** Combine's source picker value for the live document. */
const CURRENT_DOCUMENT = "current";

/** Apply a prepared redline on one transaction that the plugin will not wrap. */
export function applyComparedDocument(
  editor: Editor,
  redlineDoc: PMNodeJSON,
): boolean {
  if (editor.isDestroyed || editor.isEditable === false) return false;
  // The comparison runs on the stored form; a paged editor holds its blocks
  // in sheets, which its paginator then lays out.
  const content =
    (editor as Partial<Editor>).schema?.nodes?.pageBlock &&
    Array.isArray(redlineDoc.content) &&
    !redlineDoc.content.some((node) => node?.type === "pageBlock")
      ? { ...redlineDoc, content: wrapBlocksIntoPages(redlineDoc.content as never) }
      : redlineDoc;
  return editor
    .chain()
    .command(({ tr }) => {
      tr.setMeta(TRACK_CHANGES_SKIP_META, true);
      tr.setMeta("addToHistory", false);
      return true;
    })
    .setContent(content as PMNodeJSON)
    .run();
}

/**
 * Compare Documents panel (Word / WPS "比较文档" / legal blackline parity).
 *
 * Shows the page's version history and lets the user pick a BASE version to
 * compare the CURRENT document against. On confirm, computes a word-level LCS
 * diff between the two and renders the differences INTO the live editor as
 * track-change marks:
 *   - Text present only in the current doc → trackInsertion (green)
 *   - Text present only in the base        → trackDeletion (red, struck-through)
 *
 * The user can then accept/reject individual changes via the existing
 * Track Changes panel / Accept All / Reject All toolbar actions.
 */
export default function CompareDocumentsPanel({
  editor,
  pageId,
  onClose,
  onResult,
  mode = "compare",
  dockClassName,
}: {
  editor: Editor;
  pageId: string;
  onClose: (reason?: "applied") => void;
  /** The result now in the editor and its two sources (Show Source Documents). */
  onResult?: (result: CompareResultSources) => void;
  /** Word's Compare, or Combine (two copies' revisions merged into one). */
  mode?: "compare" | "combine";
  /** Where the host docks the task pane (Docs: its right-hand dock). */
  dockClassName?: string;
}) {
  const t = useTranslations("console-notebooks");
  const self = useSelfCollabIdentity();
  const [versions, setVersions] = useState<PageVersionItem[] | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [listFailed, setListFailed] = useState(false);
  const [busy, setBusy] = useState(false);
  const compareControllerRef = useRef<AbortController | null>(null);
  const listRequestGeneration = useRef(0);
  const compareRequestGeneration = useRef(0);
  const activePageIdRef = useRef(pageId);
  const activeEditorRef = useRef(editor);
  const closeButtonRef = useRef<HTMLButtonElement | null>(null);
  const canEdit = !editor.isDestroyed && editor.isEditable !== false;

  useLayoutEffect(() => {
    if (canEdit) return;
    compareControllerRef.current?.abort();
    compareControllerRef.current = null;
    compareRequestGeneration.current += 1;
    setBusy(false);
  }, [canEdit]);

  useLayoutEffect(() => {
    let focusFrame = 0;
    const settleFrame = window.requestAnimationFrame(() => {
      focusFrame = window.requestAnimationFrame(() => {
        closeButtonRef.current?.focus({ preventScroll: true });
      });
    });
    return () => {
      window.cancelAnimationFrame(settleFrame);
      window.cancelAnimationFrame(focusFrame);
    };
  }, []);

  const refresh = useCallback(async () => {
    const requestedPageId = pageId;
    const generation = ++listRequestGeneration.current;
    try {
      const nextVersions = await listPageVersions(requestedPageId);
      if (
        generation !== listRequestGeneration.current ||
        activePageIdRef.current !== requestedPageId
      ) {
        return;
      }
      setVersions(nextVersions);
      setError(null);
      setListFailed(false);
    } catch {
      if (
        generation !== listRequestGeneration.current ||
        activePageIdRef.current !== requestedPageId
      ) {
        return;
      }
      setError(t("compare.loadFailed"));
      setListFailed(true);
    }
  }, [pageId, t]);

  useEffect(() => {
    setVersions(null);
    setListFailed(false);
    void refresh();
    return () => {
      listRequestGeneration.current += 1;
    };
  }, [refresh]);

  // A panel can be reused while the surrounding tab switches pages. Invalidate
  // an in-flight detail fetch immediately on page/editor identity changes (and
  // on unmount), otherwise an old snapshot can be compared with and written
  // into the newly-active document after its request resolves.
  useLayoutEffect(() => {
    // Page/editor identity is an async-write ownership boundary. Publish it
    // only for a committed render: render-time ref writes let an abandoned
    // concurrent navigation cancel the still-visible page, while a passive
    // effect leaves old responses one frame in which to write into the new one.
    activePageIdRef.current = pageId;
    activeEditorRef.current = editor;
    listRequestGeneration.current += 1;
    compareRequestGeneration.current += 1;
    compareControllerRef.current?.abort();
    compareControllerRef.current = null;
    setBusy(false);
    setError(null);
    setListFailed(false);
    return () => {
      listRequestGeneration.current += 1;
      compareRequestGeneration.current += 1;
      compareControllerRef.current?.abort();
      compareControllerRef.current = null;
    };
  }, [editor, pageId]);

  const retryList = useCallback(() => {
    setVersions(null);
    setError(null);
    setListFailed(false);
    void refresh();
  }, [refresh]);

  const sourceLabel = useCallback(
    (source: string): string => {
      switch (source) {
        case "manual":
          return t("versions.sourceManual");
        case "pre_restore":
          return t("versions.sourcePreRestore");
        case "pre_copilot_parts":
          return t("versions.sourcePreCopilotParts");
        case "room_recovery":
          return t("versions.sourceRoomRecovery");
        case "autosave":
          return t("versions.sourceAutosave");
        default:
          return source;
      }
    },
    [t],
  );

  /** A version as the source pickers and the source panes name it. */
  const versionLabel = useCallback(
    (version: PageVersionItem): string =>
      t("compare.versionOption", {
        version: version.version_no,
        date: new Date(version.created_at).toLocaleString(pageInterfaceLocale()),
        source: sourceLabel(version.source),
      }),
    [sourceLabel, t],
  );

  const handleCompare = useCallback(
    async (version: PageVersionItem) => {
      if (busy || editor.isDestroyed || editor.isEditable === false) return;
      compareControllerRef.current?.abort();
      const controller = new AbortController();
      compareControllerRef.current = controller;
      const generation = ++compareRequestGeneration.current;
      const isCurrentRequest = () =>
        generation === compareRequestGeneration.current &&
        compareControllerRef.current === controller &&
        !controller.signal.aborted &&
        activePageIdRef.current === pageId &&
        activeEditorRef.current === editor &&
        !editor.isDestroyed &&
        editor.isEditable !== false;
      setBusy(true);
      setError(null);
      try {
        const detail = await getPageVersion(
          pageId,
          version.version_no,
          version.id,
          controller.signal,
        );
        if (!isCurrentRequest()) return;
        if (
          detail.id !== version.id ||
          detail.page_id !== pageId ||
          detail.version_no !== version.version_no
        ) {
          setError(t("compare.applyFailed"));
          return;
        }
        const baseDocJson = comparableDocumentSnapshot(editor, detail);
        if (!baseDocJson) {
          setError(t("compare.applyFailed"));
          return;
        }

        // The live editor is an equally untrusted comparison input: extensions
        // and imported/custom nodes can produce malformed or oversized JSON.
        // Apply the same iterative shape/resource budget as the stored detail.
        const currentDocValue = liveDocumentAsStored(editor);
        if (!isComparableDoc(currentDocValue)) {
          setError(t("compare.applyFailed"));
          return;
        }
        const currentDocJson = comparableForm(editor, currentDocValue);
        if (!currentDocJson) {
          setError(t("compare.applyFailed"));
          return;
        }

        // Unresolved changes, then the structure, then the LCS budget, then
        // the redline: the server answers the first that stops it, so an
        // over-budget pair never allocates its matrix.
        if (!isCurrentRequest()) return;
        const answer = await callEngine(
          "tools",
          "document.compare",
          { base: baseDocJson, current: currentDocJson, timestamp: Date.now() },
          { signal: controller.signal },
        );
        if (!isCurrentRequest()) return;
        if (isCombineOutcome(answer)) {
          setError(t("compare.applyFailed"));
          return;
        }
        if (answer.outcome === "refused") {
          setError(t(COMPARE_REFUSAL_MESSAGE[answer.reason] ?? "compare.applyFailed"));
          return;
        }
        const redlineDoc = answer.doc;
        // A replacement can merge both documents' text (current insertion +
        // base tombstone) into one redline. Re-validate the result so two valid
        // inputs cannot produce an editor state larger than the server will
        // ever accept and leave the user stuck with an unsaveable document.
        if (!isComparableDoc(redlineDoc)) {
          setError(t("compare.tooLarge"));
          return;
        }

        // Count changes for user feedback.
        const { insertions, deletions } = answer;

        // Load the redline into the editor via a skip-tracked transaction so
        // the compare result itself does not get double-tracked.
        // The metadata and setContent must share the same chained transaction;
        // setting meta on editor.state.tr and then dispatching another command
        // lets suggesting mode double-track the entire redline replacement.
        if (!isCurrentRequest()) return;
        if (!applyComparedDocument(editor, redlineDoc)) {
          setError(t("compare.applyFailed"));
          return;
        }

        const summary = t("compare.appliedSummary", {
          insertions,
          deletions,
        });
        console.info("[CompareDocuments]", summary);

        // Word's compare result shows the original (the version) and the
        // revised document (the page as it was) beside it.
        onResult?.({
          id: newResultId(),
          kind: "compare",
          original: { label: versionLabel(version), doc: baseDocJson },
          revised: { label: t("compare.currentDocument"), doc: currentDocJson },
        });
        onClose("applied");
      } catch (err) {
        if (!isCurrentRequest()) return;
        console.error("[CompareDocuments] error", err);
        // Two documents over the request's bound are "too large"; no
        // connection says so; anything else failed to apply.
        setError(
          t(
            err instanceof EngineRefusedError && err.reason === "too_large"
              ? "compare.tooLarge"
              : err instanceof EngineRefusedError
                ? "compare.applyFailed"
                : engineErrorMessageKey(err) === "engineHost.needsConnection"
                  ? "engineHost.needsConnection"
                  : "compare.applyFailed",
          ),
        );
      } finally {
        if (isCurrentRequest()) {
          compareControllerRef.current = null;
          setBusy(false);
        }
      }
    },
    [busy, editor, onClose, onResult, pageId, t, versionLabel],
  );

  // ---- Combine (Word: Review > Compare > Combine) ----
  // Word's dialog: Original document, Revised document, "Label unmarked
  // changes with" and, under More, "Show changes in". MRAI's documents are
  // this page and its saved versions; the result is shown in this page, so
  // one of the two must be the current document (Word's "Show changes in:
  // Original document" or "Revised document", whichever it is).
  const [originalChoice, setOriginalChoice] = useState<string | null>(null);
  const [revisedChoice, setRevisedChoice] = useState<string>(CURRENT_DOCUMENT);
  const [authorLabel, setAuthorLabel] = useState("");
  const authorEditedRef = useRef(false);
  const [formatChoicePending, setFormatChoicePending] = useState(false);

  // Word fills "Label unmarked changes with" with the revised copy's author
  // (the oracle: its last editor). A saved version records no editor here, so
  // the label starts as the person combining, whom they can rename.
  useEffect(() => {
    if (!authorEditedRef.current && self.name) setAuthorLabel(self.name);
  }, [self.name]);

  // The original defaults to the newest saved version, the revised to the page.
  useEffect(() => {
    if (originalChoice !== null || !versions?.length) return;
    setOriginalChoice(versions[0].id);
  }, [originalChoice, versions]);

  useEffect(() => {
    setFormatChoicePending(false);
  }, [originalChoice, revisedChoice]);

  const combineChoiceError =
    originalChoice === null
      ? null
      : originalChoice === revisedChoice
        ? t("compare.combineSameDocument")
        : originalChoice !== CURRENT_DOCUMENT && revisedChoice !== CURRENT_DOCUMENT
          ? t("compare.combineNeedsCurrent")
          : null;

  const handleCombine = useCallback(
    async (formatFrom: DocumentCombineFormatSource | null) => {
      if (busy || editor.isDestroyed || editor.isEditable === false) return;
      if (originalChoice === null || combineChoiceError) return;
      compareControllerRef.current?.abort();
      const controller = new AbortController();
      compareControllerRef.current = controller;
      const generation = ++compareRequestGeneration.current;
      const isCurrentRequest = () =>
        generation === compareRequestGeneration.current &&
        compareControllerRef.current === controller &&
        !controller.signal.aborted &&
        activePageIdRef.current === pageId &&
        activeEditorRef.current === editor &&
        !editor.isDestroyed &&
        editor.isEditable !== false;
      setBusy(true);
      setError(null);
      try {
        // Each side, read and bounded like Compare's two inputs.
        const readSide = async (
          choice: string,
        ): Promise<{ label: string; doc: PMNodeJSON } | null> => {
          if (choice === CURRENT_DOCUMENT) {
            const value = liveDocumentAsStored(editor);
            if (!isComparableDoc(value)) return null;
            const doc = comparableForm(editor, value);
            return doc ? { label: t("compare.currentDocument"), doc } : null;
          }
          const version = versions?.find((item) => item.id === choice);
          if (!version) return null;
          const detail = await getPageVersion(
            pageId,
            version.version_no,
            version.id,
            controller.signal,
          );
          if (
            detail.id !== version.id ||
            detail.page_id !== pageId ||
            detail.version_no !== version.version_no
          ) {
            return null;
          }
          const doc = comparableDocumentSnapshot(editor, detail);
          return doc ? { label: versionLabel(version), doc } : null;
        };
        const original = await readSide(originalChoice);
        if (!isCurrentRequest()) return;
        const revised = await readSide(revisedChoice);
        if (!isCurrentRequest()) return;
        if (!original || !revised) {
          setError(t("compare.applyFailed"));
          return;
        }
        const answer = await callEngine(
          "tools",
          "document.compare",
          {
            base: original.doc,
            current: revised.doc,
            timestamp: Date.now(),
            combine: {
              author: authorLabel.trim().slice(0, TOOLS_COMBINE_MAX_AUTHOR_LENGTH) || null,
              formatFrom,
            },
          },
          { signal: controller.signal },
        );
        if (!isCurrentRequest()) return;
        if (!isCombineOutcome(answer)) {
          setError(
            answer.outcome === "refused"
              ? t(COMPARE_REFUSAL_MESSAGE[answer.reason] ?? "compare.applyFailed")
              : t("compare.applyFailed"),
          );
          return;
        }
        if (answer.outcome === "needs_format_choice") {
          // Word: "Word can only store one set of formatting changes in the
          // final merged document. Choose which set of formatting changes to
          // keep."
          setFormatChoicePending(true);
          return;
        }
        if (!isComparableDoc(answer.doc)) {
          setError(t("compare.tooLarge"));
          return;
        }
        if (!applyComparedDocument(editor, answer.doc)) {
          setError(t("compare.applyFailed"));
          return;
        }
        console.info(
          "[CombineDocuments]",
          t("compare.combinedSummary", {
            insertions: answer.insertions,
            deletions: answer.deletions,
          }),
        );
        onResult?.({
          id: newResultId(),
          kind: "combine",
          original,
          revised,
        });
        onClose("applied");
      } catch (err) {
        if (!isCurrentRequest()) return;
        console.error("[CombineDocuments] error", err);
        setError(
          t(
            err instanceof EngineRefusedError && err.reason === "too_large"
              ? "compare.tooLarge"
              : !(err instanceof EngineRefusedError) &&
                  engineErrorMessageKey(err) === "engineHost.needsConnection"
                ? "engineHost.needsConnection"
                : "compare.applyFailed",
          ),
        );
      } finally {
        if (isCurrentRequest()) {
          compareControllerRef.current = null;
          setBusy(false);
        }
      }
    },
    [
      authorLabel,
      busy,
      combineChoiceError,
      editor,
      onClose,
      onResult,
      originalChoice,
      pageId,
      revisedChoice,
      t,
      versionLabel,
      versions,
    ],
  );

  const sourceOptions = (testId: string, value: string | null, onChange: (next: string) => void, label: string) => (
    <label className={officeTaskPaneParts.field}>
      {label}
      <select
        data-testid={testId}
        value={value ?? ""}
        disabled={busy || versions === null}
        onChange={(event) => onChange(event.target.value)}
      >
        <option value={CURRENT_DOCUMENT}>{t("compare.currentDocument")}</option>
        {versions?.map((version) => (
          <option key={version.id} value={version.id}>
            {versionLabel(version)}
          </option>
        ))}
      </select>
    </label>
  );

  if (mode === "combine") {
    const destination =
      originalChoice === CURRENT_DOCUMENT
        ? t("compare.showChangesInOriginal")
        : revisedChoice === CURRENT_DOCUMENT
          ? t("compare.showChangesInRevised")
          : null;
    return (
      <OfficeTaskPane
        title={t("compare.combineTitle")}
        closeLabel={t("compare.combineClose")}
        onClose={onClose}
        testIdPrefix="document-combine"
        testId="document-combine-panel"
        closeTestId="document-combine-panel-close"
        closeRef={closeButtonRef}
        dockRight={false}
        className={dockClassName}
      >
        <div className={officeTaskPaneParts.body}>
          <div className={officeTaskPaneParts.section}>
            <p className={officeTaskPaneParts.hint}>{t("compare.combinePrompt")}</p>
            {sourceOptions(
              "document-combine-original",
              originalChoice,
              setOriginalChoice,
              t("compare.originalDocument"),
            )}
            {sourceOptions(
              "document-combine-revised",
              revisedChoice,
              setRevisedChoice,
              t("compare.revisedDocument"),
            )}
            <label className={officeTaskPaneParts.field}>
              {t("compare.labelUnmarkedChanges")}
              <input
                type="text"
                data-testid="document-combine-author"
                value={authorLabel}
                maxLength={TOOLS_COMBINE_MAX_AUTHOR_LENGTH}
                disabled={busy}
                onChange={(event) => {
                  authorEditedRef.current = true;
                  setAuthorLabel(event.target.value);
                }}
              />
            </label>
            {destination ? (
              <p className={officeTaskPaneParts.hint} data-testid="document-combine-destination">
                {t("compare.showChangesIn", { document: destination })}
              </p>
            ) : null}
            {combineChoiceError ? (
              <p
                className={officeTaskPaneParts.hint}
                role="status"
                data-testid="document-combine-choice-error"
              >
                {combineChoiceError}
              </p>
            ) : null}
            {versions?.length === 0 ? (
              <p className={officeTaskPaneParts.hint}>{t("compare.empty")}</p>
            ) : null}
            {error ? (
              <div role="alert" style={{ color: "var(--office-dialog-danger, #b91c1c)", fontSize: 12 }}>
                <span>{error}</span>
                {listFailed ? (
                  <button
                    className={officeTaskPaneParts.button}
                    type="button"
                    onClick={retryList}
                    data-testid="document-combine-retry"
                  >
                    {t("compare.retry")}
                  </button>
                ) : null}
              </div>
            ) : null}
          </div>
          {formatChoicePending ? (
            <div
              className={officeTaskPaneParts.section}
              role="group"
              aria-label={t("compare.formatConflictPrompt")}
              data-testid="document-combine-format-choice"
            >
              <p className={officeTaskPaneParts.hint}>{t("compare.formatConflictPrompt")}</p>
              <div className={officeTaskPaneParts.row}>
                <button
                  className={officeTaskPaneParts.button}
                  type="button"
                  disabled={busy || !canEdit}
                  data-testid="document-combine-format-original"
                  onClick={() => void handleCombine("original")}
                >
                  {t("compare.originalDocument")}
                </button>
                <button
                  className={officeTaskPaneParts.button}
                  type="button"
                  disabled={busy || !canEdit}
                  data-testid="document-combine-format-revised"
                  onClick={() => void handleCombine("revised")}
                >
                  {t("compare.revisedDocument")}
                </button>
              </div>
            </div>
          ) : null}
        </div>
        <div className={officeTaskPaneParts.footer}>
          <button
            className={officeTaskPaneParts.button}
            type="button"
            data-testid="document-combine-submit"
            disabled={
              busy ||
              !canEdit ||
              originalChoice === null ||
              Boolean(combineChoiceError) ||
              formatChoicePending
            }
            onClick={() => void handleCombine(null)}
          >
            {busy ? <Loader2 size={12} className="animate-spin" aria-hidden /> : null}
            {t("compare.combineAction")}
          </button>
        </div>
      </OfficeTaskPane>
    );
  }

  // Word's Compare is a separate view on the desktop; here the source
  // picker is a task pane in the Docs dock, the shared Office task pane.
  return (
    <OfficeTaskPane
      title={t("compare.title")}
      closeLabel={t("compare.close")}
      onClose={onClose}
      testIdPrefix="document-compare"
      testId="document-compare-panel"
      closeTestId="document-compare-panel-close"
      closeRef={closeButtonRef}
      dockRight={false}
      className={dockClassName}
      // Without a dock the pane sits in its host's flow, bordered.
      style={dockClassName ? undefined : {
        border: "1px solid var(--console-border, rgba(15, 23, 42, 0.12))",
        borderRadius: 8,
        maxHeight: 320,
      }}
    >
      <div
        className={officeTaskPaneParts.body}
        style={{
          display: "flex",
          flexDirection: "column",
          gap: 6,
          padding: "10px 14px",
          fontSize: 12.5,
        }}
      >

      <p
        style={{
          margin: 0,
          color: "var(--office-pane-muted, #4b5563)",
          fontSize: 12,
        }}
      >
        {t("compare.pickPrompt")}
      </p>

      {error ? (
        <div role="alert" style={{ color: "var(--office-dialog-danger, #b91c1c)", fontSize: 12 }}>
          <span>{error}</span>
          {listFailed ? (
            <button
              className={officeTaskPaneParts.button}
              type="button"
              onClick={retryList}
              data-testid="document-compare-retry"
            >
              {t("compare.retry")}
            </button>
          ) : null}
        </div>
      ) : null}

      {versions === null && !listFailed ? (
        <EditorLoadingState
          kind="list"
          minimal
          label={t("compare.loading")}
        />
      ) : versions?.length === 0 ? (
        <div
          data-testid="document-compare-panel-empty"
          style={{ color: "var(--office-pane-muted, #6b7280)" }}
        >
          {t("compare.empty")}
        </div>
      ) : (
        versions?.map((version) => (
          <div
            key={version.id}
            data-testid={`document-compare-version-row-${version.version_no}`}
            style={{
              display: "flex",
              alignItems: "center",
              gap: 8,
              padding: "4px 0",
              borderTop:
                "1px solid var(--office-pane-border, rgba(15, 23, 42, 0.06))",
            }}
          >
            <span style={{ fontWeight: 600, whiteSpace: "nowrap" }}>
              v{version.version_no}
            </span>
            <span
              style={{
                whiteSpace: "nowrap",
                color: "var(--office-pane-muted, #6b7280)",
              }}
            >
              {new Date(version.created_at).toLocaleString(pageInterfaceLocale())}
            </span>
            <span style={{ whiteSpace: "nowrap" }}>
              {sourceLabel(version.source)}
            </span>
            <button
              className={officeTaskPaneParts.button}
              type="button"
              onClick={() => void handleCompare(version)}
              disabled={busy || !canEdit}
              data-testid={`document-compare-select-${version.version_no}`}
              style={{
                display: "inline-flex",
                alignItems: "center",
                gap: 4,
                fontSize: 12,
                cursor: "pointer",
                whiteSpace: "nowrap",
                marginLeft: "auto",
              }}
            >
              {busy ? (
                <Loader2 size={12} className="animate-spin" aria-hidden />
              ) : (
                <ArrowRightLeft size={12} aria-hidden />
              )}
              {t("compare.compareWith")}
            </button>
          </div>
        ))
      )}
      </div>
    </OfficeTaskPane>
  );
}
