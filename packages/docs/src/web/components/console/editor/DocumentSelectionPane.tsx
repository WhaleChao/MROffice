"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { isDoubleTap } from "@mrnote/core/web/lib/double-tap.ts";
import {
  useCallback,
  useEffect,
  useMemo,
  useRef,
  useState,
  type KeyboardEvent as ReactKeyboardEvent,
  type RefObject,
} from "react";
import type { Editor } from "@tiptap/react";
import type { Node as ProseMirrorNode } from "@tiptap/pm/model";
import { closeHistory } from "@tiptap/pm/history";
import { NodeSelection } from "@tiptap/pm/state";
import {
  ArrowDownToLine,
  ArrowUpToLine,
  BarChart3,
  Boxes,
  FileText,
  Image as ImageIcon,
  LayoutPanelTop,
  ChevronDown,
  ChevronUp,
  Shapes,
  Workflow,
} from "@mrnote/core/web/components/office-ribbon/icons/houseIcons.tsx";
import { useTranslations } from "@mrnote/core/host/next-intl.tsx";
import { toast } from "@mrnote/core/web/hooks/use-toast.ts";

import { DocumentDialogIcon } from "@mrnote/core/web/components/console/editor/native-document/DocumentDialogIcon.tsx";
import {
  collectDocumentGoToTargets,
  goToDocumentTarget,
  type DocumentGoToTarget,
} from "./native-document/documentGoTo";
import {
  documentShapeIsGroup,
  normalizeDocumentShapeName,
  documentShapeUngroupAttrs,
  documentShapeZOrderPatches,
  resolveDocumentShapeStack,
  type DocumentShapeZOrderAction,
  documentStackPatchAttrs,
} from "@mrnote/core/web/lib/document-shapes.ts";
import {
  documentShapeMultiSelection,
  documentShapeMultiSelectionKey,
  groupDocumentShapes,
  documentObjectsGroupable,
  documentShapePickable,
  ungroupDocumentShape,
} from "@mrnote/core/web/components/console/editor/native-document/documentShapeExtension.ts";
import { OfficeTaskPane } from "@mrnote/core/web/components/office-ribbon/OfficeTaskPane.tsx";
import {
  OfficeSelectionPaneList,
  OfficeSelectionPaneRow,
} from "@mrnote/core/web/components/office-ribbon/OfficeSelectionPaneList.tsx";
import {
  canHideInSelectionPane,
  isSelectionPaneHidden,
  setObjectsHidden,
} from "@mrnote/core/web/components/console/editor/native-document/documentObjectVisibility.ts";
import dockStyles from "@mrnote/core/web/components/console/editor/native-document/DocumentShapeFormatPane.module.css";
import styles from "./DocumentSelectionPane.module.css";

interface DocumentSelectionPaneProps {
  editor: Editor;
  onClose: () => void;
  returnFocusRef: RefObject<HTMLButtonElement | null>;
  /** Tells the ribbon whether Group has two or more groupable shapes to act
   * on (the pane's Ctrl+click selection), so Shape Format > Group can be
   * enabled only then, as in Word. */
  onGroupAvailabilityChange?: (
    action: { readonly run: () => void } | null,
  ) => void;
}

type RuntimeEditor = Editor & {
  on?: (event: string, callback: () => void) => void;
  off?: (event: string, callback: () => void) => void;
};

interface SelectionPaneSnapshot {
  doc: ProseMirrorNode;
  targets: DocumentGoToTarget[];
  selectedId: string | null;
  shapeStackComplete: boolean;
}

function objectType(target: DocumentGoToTarget): string {
  return target.id.split(":", 3)[1] || "object";
}

function selectedObjectId(
  editor: Editor,
  targets: readonly DocumentGoToTarget[],
): string | null {
  const selection = editor.state.selection;
  if (!(selection instanceof NodeSelection)) return null;
  return (
    targets.find(
      (target) =>
        target.selection === "node" && target.from === selection.from,
    )?.id ?? null
  );
}









export function orderSelectionPaneTargets(
  targets: readonly DocumentGoToTarget[],
  stack: Pick<ReturnType<typeof resolveDocumentShapeStack>, "complete" | "backToFront">,
): DocumentGoToTarget[] {
  if (!stack.complete) return [...targets];
  const shapeTargets = new Map(
    targets
      .filter((target) => objectType(target) === "documentShape")
      .map((target) => [target.from, target]),
  );
  const orderedShapes = [...stack.backToFront]
    .reverse()
    .map((entry) => shapeTargets.get(entry.pos))
    .filter((target): target is DocumentGoToTarget => Boolean(target));
  const stacked = new Set(stack.backToFront.map((entry) => entry.pos));
  return [
    ...orderedShapes,
    ...targets.filter(
      (target) => objectType(target) !== "documentShape" || !stacked.has(target.from),
    ),
  ];
}

function readSnapshot(editor: Editor): SelectionPaneSnapshot {
  const doc = editor.state.doc;
  const targets = collectDocumentGoToTargets(doc).filter(
    (target) => target.kind === "object",
  );
  const stack = resolveDocumentShapeStack(doc);
  const orderedTargets = orderSelectionPaneTargets(targets, stack);
  return {
    doc,
    targets: orderedTargets,
    selectedId: selectedObjectId(editor, orderedTargets),
    shapeStackComplete: stack.complete,
  };
}

function typeIcon(type: string) {
  switch (type) {
    case "documentShape":
    case "inlineDocumentShape":
      return <Shapes size={17} aria-hidden />;
    case "image":
      return <ImageIcon size={17} aria-hidden />;
    case "chart":
    case "codeChart":
      return <BarChart3 size={17} aria-hidden />;
    case "smartArt":
    case "mermaid":
      return <Workflow size={17} aria-hidden />;
    case "whiteboard":
      return <LayoutPanelTop size={17} aria-hidden />;
    case "file":
      return <FileText size={17} aria-hidden />;
    default:
      return <Boxes size={17} aria-hidden />;
  }
}

/**
 * MRNote's session-only Docs Selection Pane extension.
 *
 * Word documents expose many inline objects whose visibility/z-order cannot be
 * truthfully changed by our current model. This pane therefore owns only the
 * closed contract: enumerate live object atoms, select one by stable document
 * identity, reorder a complete native-shape floating stack, and group multiple
 * fully-authored shapes. Authenticated imported WPG groups may move as one,
 * while ungroup remains limited to groups authored by this editor so native
 * member custody is never flattened implicitly.
 */
export default function DocumentSelectionPane({
  editor,
  onClose,
  returnFocusRef,
  onGroupAvailabilityChange,
}: DocumentSelectionPaneProps) {
  const t = useTranslations("console-notebooks");
  const [snapshot, setSnapshot] = useState<SelectionPaneSnapshot>(() =>
    readSnapshot(editor),
  );
  const [rovingId, setRovingId] = useState<string | null>(
    () => snapshot.selectedId ?? snapshot.targets[0]?.id ?? null,
  );
  const [selectedIds, setSelectedIds] = useState<ReadonlySet<string>>(
    () => new Set(snapshot.selectedId ? [snapshot.selectedId] : []),
  );
  const rowRefs = useRef(new Map<string, HTMLButtonElement>());
  // Word's Selection Pane rename: double-click (or F2) a name to type a new one.
  const [renaming, setRenaming] = useState<{ id: string; draft: string } | null>(null);
  const placedInitialFocusRef = useRef(false);
  const targetIds = useMemo(
    () => new Set(snapshot.targets.map((target) => target.id)),
    [snapshot.targets],
  );
  const effectiveRovingId =
    rovingId && targetIds.has(rovingId)
      ? rovingId
      : (snapshot.selectedId ?? snapshot.targets[0]?.id ?? null);
  const effectiveSelectedIds = useMemo(() => {
    const retained = new Set(
      [...selectedIds].filter((id) => targetIds.has(id)),
    );
    if (snapshot.selectedId && !retained.has(snapshot.selectedId)) {
      return new Set([snapshot.selectedId]);
    }
    return retained;
  }, [selectedIds, snapshot.selectedId, targetIds]);

  const close = useCallback(() => {
    onClose();
    window.requestAnimationFrame(() => {
      const trigger = returnFocusRef.current;
      if (trigger?.isConnected) {
        trigger.focus({ preventScroll: true });
        return;
      }
      // Picking an object here can hand the ribbon to that object's
      // contextual tab (a chart row activates Chart Design), and an inactive
      // tab no longer constructs its controls, so the trigger that opened the
      // pane may be gone. Return focus to the document holding the selected
      // object instead of dropping it on <body>.
      try {
        if (!editor.isDestroyed) editor.view.focus();
      } catch {
        // The editor view can disappear between the lifecycle check and focus.
      }
    });
  }, [editor, onClose, returnFocusRef]);

  useEffect(() => {
    const runtime = editor as RuntimeEditor;
    const refresh = () => setSnapshot(readSnapshot(editor));
    refresh();
    runtime.on?.("update", refresh);
    runtime.on?.("selectionUpdate", refresh);
    return () => {
      runtime.off?.("update", refresh);
      runtime.off?.("selectionUpdate", refresh);
    };
  }, [editor]);

  useEffect(() => {
    if (placedInitialFocusRef.current) return;
    const focusId = snapshot.selectedId ?? snapshot.targets[0]?.id;
    if (!focusId) return;
    placedInitialFocusRef.current = true;
    const frame = window.requestAnimationFrame(() => {
      setRovingId(focusId);
      rowRefs.current.get(focusId)?.focus({ preventScroll: true });
    });
    return () => window.cancelAnimationFrame(frame);
  }, [snapshot.selectedId, snapshot.targets]);

  const typeLabels = useMemo(
    () => ({
      documentShape: t("toolbar.shape"),
      inlineDocumentShape: t("toolbar.shape"),
      image: t("editor.image"),
      chart: t("toolbar.chart"),
      codeChart: t("toolbar.chart"),
      smartArt: t("toolbar.smartArt"),
      mermaid: t("toolbar.object"),
      whiteboard: t("editor.whiteboard_label"),
      file: t("toolbar.file"),
    }),
    [t],
  );
  const shapeOrder = useMemo(() => {
    const ids = snapshot.targets
      .filter((target) => objectType(target) === "documentShape")
      .map((target) => target.id);
    return {
      count: ids.length,
      indexById: new Map(ids.map((id, index) => [id, index])),
    };
  }, [snapshot.targets]);

  const labelFor = useCallback(
    (target: DocumentGoToTarget) => {
      const type = objectType(target);
      const typeLabel =
        typeLabels[type as keyof typeof typeLabels] ?? t("toolbar.object");
      const raw = target.label?.trim();
      return raw && raw !== type ? raw : `${typeLabel} ${target.ordinal}`;
    },
    [t, typeLabels],
  );

  const activate = useCallback(
    (target: DocumentGoToTarget, additive = false) => {
      const canAdd = additive && objectType(target) === "documentShape";
      let next: Set<string>;
      if (!canAdd) {
        next = new Set([target.id]);
      } else {
        next = new Set(effectiveSelectedIds);
        if (next.has(target.id) && next.size > 1) next.delete(target.id);
        else next.add(target.id);
      }
      setSelectedIds(next);
      // Ctrl-clicking a selected object takes it OUT of the selection: the
      // document then selects one that stays (the last one added), never
      // the one just removed -- which the pane would read back as "only B".
      const selectedTarget = next.has(target.id)
        ? target
        : (snapshot.targets.find(
            (candidate) => candidate.id === [...next].at(-1),
          ) ?? target);
      if (!goToDocumentTarget(editor, selectedTarget, snapshot.doc)) return;
      // The rows picked here are the document's selection too, as Shift +
      // click on the page makes it: Align, the ribbon's count and a drag act
      // on all of them, not only on the last one picked (WUA 188).
      const state = editor.state;
      const positions = snapshot.targets
        .filter(
          (candidate) =>
            next.has(candidate.id) &&
            // A floating chart is picked with the shapes: Group takes it.
            documentShapePickable(snapshot.doc.nodeAt(candidate.from)),
        )
        .map((candidate) => candidate.from);
      if (
        canAdd &&
        positions.length > 1 &&
        state.doc === snapshot.doc &&
        state.selection instanceof NodeSelection &&
        positions.includes(state.selection.from)
      ) {
        editor.view.dispatch(state.tr.setMeta(documentShapeMultiSelectionKey, { positions }));
      } else if (documentShapeMultiSelection(state).length > 0) {
        editor.view.dispatch(state.tr.setMeta(documentShapeMultiSelectionKey, { positions: [] }));
      }
      setRovingId(target.id);
      window.requestAnimationFrame(() => {
        rowRefs.current.get(target.id)?.focus({ preventScroll: true });
      });
    },
    [editor, effectiveSelectedIds, snapshot.doc, snapshot.targets],
  );

  const selectedShapeTargets = useMemo(
    () =>
      snapshot.targets.filter(
        (target) =>
          effectiveSelectedIds.has(target.id) &&
          (objectType(target) === "documentShape" ||
            documentShapePickable(snapshot.doc.nodeAt(target.from))),
      ),
    [effectiveSelectedIds, snapshot.doc, snapshot.targets],
  );
  const selectedShapeNodes = useMemo(
    () =>
      selectedShapeTargets.map((target) => ({
        target,
        node: snapshot.doc.nodeAt(target.from),
      })),
    [selectedShapeTargets, snapshot.doc],
  );
  // The same test Group itself applies (documentShapeGroupAttrs): a button
  // that lit for two shapes the group then refused did nothing at all.
  const canGroup =
    editor.isEditable &&
    snapshot.shapeStackComplete &&
    selectedShapeNodes.length >= 2 &&
    selectedShapeNodes.every(
      ({ node }) =>
        node?.type.name === "chart" ||
        (node?.type.name === "documentShape" &&
          !documentShapeIsGroup(node.attrs as Record<string, unknown>) &&
          node.attrs.source === null),
    ) &&
    // Shapes, and the floating charts picked with them (documentShapeGroupAttrs).
    documentObjectsGroupable(
      snapshot.doc,
      selectedShapeNodes.map(({ target }) => target.from),
    );
  const canUngroup =
    editor.isEditable &&
    selectedShapeNodes.length === 1 &&
    selectedShapeNodes[0]?.node?.type.name === "documentShape" &&
    documentShapeUngroupAttrs(
      selectedShapeNodes[0].node.attrs as Record<string, unknown>,
    ) !== null;

  const refreshAfterStructureChange = useCallback(() => {
    const next = readSnapshot(editor);
    setSnapshot(next);
    setSelectedIds(new Set(next.selectedId ? [next.selectedId] : []));
    setRovingId(next.selectedId ?? next.targets[0]?.id ?? null);
  }, [editor]);

  const groupSelectedShapes = useCallback(() => {
    if (!canGroup) return;
    const position = groupDocumentShapes(
      editor,
      snapshot.doc,
      selectedShapeTargets.map((target) => target.from),
    );
    if (position === null) {
      toast({ title: t("toolbar.onlyShapesDrawnInEditorCan") });
      return;
    }
    refreshAfterStructureChange();
  }, [
    canGroup,
    editor,
    refreshAfterStructureChange,
    selectedShapeTargets,
    snapshot.doc,
  ]);

  useEffect(() => {
    if (!onGroupAvailabilityChange) return;
    onGroupAvailabilityChange(canGroup ? { run: groupSelectedShapes } : null);
  }, [canGroup, groupSelectedShapes, onGroupAvailabilityChange]);
  useEffect(
    () => () => onGroupAvailabilityChange?.(null),
    [onGroupAvailabilityChange],
  );

  const ungroupSelectedShape = useCallback(() => {
    const target = selectedShapeTargets[0];
    if (
      !target ||
      !canUngroup ||
      !ungroupDocumentShape(editor, snapshot.doc, target.from)
    ) {
      return;
    }
    refreshAfterStructureChange();
  }, [
    canUngroup,
    editor,
    refreshAfterStructureChange,
    selectedShapeTargets,
    snapshot.doc,
  ]);

  const reorderShape = useCallback(
    (target: DocumentGoToTarget, action: DocumentShapeZOrderAction) => {
      if (
        editor.isDestroyed ||
        !editor.isEditable ||
        editor.state.doc !== snapshot.doc ||
        !snapshot.shapeStackComplete
      ) {
        return;
      }
      const patches = documentShapeZOrderPatches(
        resolveDocumentShapeStack(editor.state.doc),
        target.from,
        action,
      );
      if (!patches?.length) return;
      let transaction = editor.state.tr;
      for (const patch of patches) {
        const node = transaction.doc.nodeAt(patch.pos);
        const attrs = node ? documentStackPatchAttrs(node, patch.relativeHeight) : null;
        if (!node || !attrs) return;
        transaction = transaction.setNodeMarkup(patch.pos, undefined, attrs, node.marks);
      }
      transaction = closeHistory(transaction).setSelection(
        NodeSelection.create(transaction.doc, target.from),
      );
      if (!transaction.docChanged) return;
      editor.view.dispatch(transaction);
      editor.view.focus();
      setRovingId(target.id);
      window.requestAnimationFrame(() => {
        rowRefs.current.get(target.id)?.focus({ preventScroll: true });
      });
    },
    [editor, snapshot.doc, snapshot.shapeStackComplete],
  );

  /**
   * Word's Selection Pane eye and rename, written on the drawing's `wp:docPr`
   * (`name`, `hidden`): one history step each, the object stays selected.
   */
  const patchShapeDocPr = useCallback(
    (targets: readonly DocumentGoToTarget[], patch: { name?: string; hidden?: boolean }) => {
      if (editor.isDestroyed || !editor.isEditable || editor.state.doc !== snapshot.doc) return;
      let transaction = editor.state.tr;
      for (const target of targets) {
        const node = transaction.doc.nodeAt(target.from);
        if (!node || node.type.name !== "documentShape") continue;
        const next = { ...node.attrs };
        if (patch.name !== undefined) next.name = patch.name;
        if (patch.hidden !== undefined) next.hidden = patch.hidden ? true : undefined;
        transaction = transaction.setNodeMarkup(target.from, undefined, next, node.marks);
      }
      if (!transaction.docChanged) return;
      editor.view.dispatch(closeHistory(transaction));
    },
    [editor, snapshot.doc],
  );
  // The eye: a shape's is its `wp:docPr/@hidden` (`hidden`, written by
  // both DOCX writers); a picture's or text box's is the document-local
  // `selectionPaneHidden` flag (documentObjectVisibility.ts), which no
  // package writer carries.
  const canHideTarget = useCallback(
    (target: DocumentGoToTarget) => {
      const node = snapshot.doc.nodeAt(target.from);
      return node?.type.name === "documentShape" || canHideInSelectionPane(node);
    },
    [snapshot.doc],
  );
  const shapeHidden = useCallback(
    (target: DocumentGoToTarget) => {
      const node = snapshot.doc.nodeAt(target.from);
      return node?.type.name === "documentShape"
        ? node.attrs.hidden === true
        : isSelectionPaneHidden(node);
    },
    [snapshot.doc],
  );
  const hideableTargets = useMemo(
    () => snapshot.targets.filter(canHideTarget),
    [canHideTarget, snapshot.targets],
  );
  /** Show / Hide every target in one transaction (one history step). */
  const setVisibility = useCallback(
    (targets: readonly DocumentGoToTarget[], hidden: boolean) => {
      if (editor.isDestroyed || !editor.isEditable || editor.state.doc !== snapshot.doc) return;
      const transaction = editor.state.tr;
      const others: number[] = [];
      for (const target of targets) {
        const node = transaction.doc.nodeAt(target.from);
        if (!node) continue;
        if (node.type.name !== "documentShape") {
          others.push(target.from);
          continue;
        }
        if ((node.attrs.hidden === true) === hidden) continue;
        transaction.setNodeMarkup(
          target.from,
          undefined,
          { ...node.attrs, hidden: hidden ? true : undefined },
          node.marks,
        );
      }
      setObjectsHidden(transaction, others, hidden);
      if (!transaction.docChanged) return;
      editor.view.dispatch(closeHistory(transaction));
    },
    [editor, snapshot.doc],
  );
  // Word's Bring Forward / Send Backward at the top of the pane act on the
  // one selected object.
  const singleSelectedShape =
    effectiveSelectedIds.size === 1
      ? snapshot.targets.find(
          (target) =>
            effectiveSelectedIds.has(target.id) &&
            objectType(target) === "documentShape",
        )
      : undefined;
  const singleShapeIndex = singleSelectedShape
    ? (shapeOrder.indexById.get(singleSelectedShape.id) ?? -1)
    : -1;
  const canReorderSelected =
    editor.isEditable &&
    snapshot.shapeStackComplete &&
    shapeOrder.count > 1 &&
    singleShapeIndex >= 0;
  const commitRename = useCallback(() => {
    if (!renaming) return;
    const target = snapshot.targets.find((candidate) => candidate.id === renaming.id);
    const name = normalizeDocumentShapeName(renaming.draft);
    setRenaming(null);
    if (target && name) patchShapeDocPr([target], { name });
    window.requestAnimationFrame(() => rowRefs.current.get(renaming.id)?.focus({ preventScroll: true }));
  }, [patchShapeDocPr, renaming, snapshot.targets]);

  const handleRowKeyDown = useCallback(
    (
      event: ReactKeyboardEvent<HTMLButtonElement>,
      target: DocumentGoToTarget,
      index: number,
    ) => {
      const focusAt = (nextIndex: number) => {
        const next = snapshot.targets[nextIndex];
        if (!next) return;
        setRovingId(next.id);
        rowRefs.current.get(next.id)?.focus({ preventScroll: true });
      };
      switch (event.key) {
        case "ArrowDown":
          event.preventDefault();
          focusAt(Math.min(index + 1, snapshot.targets.length - 1));
          return;
        case "ArrowUp":
          event.preventDefault();
          focusAt(Math.max(index - 1, 0));
          return;
        case "Home":
          event.preventDefault();
          focusAt(0);
          return;
        case "End":
          event.preventDefault();
          focusAt(snapshot.targets.length - 1);
          return;
        case "Enter":
        case " ":
          event.preventDefault();
          activate(target, event.metaKey || event.ctrlKey);
          return;
        case "F2":
          if (objectType(target) === "documentShape" && editor.isEditable) {
            event.preventDefault();
            setRenaming({ id: target.id, draft: labelFor(target) });
          }
          return;
      }
    },
    [activate, editor.isEditable, labelFor, snapshot.targets],
  );

  // Word's Selection Pane (Home > Select > Selection Pane) is a docked task
  // pane: the shared Office task pane, in the Docs dock Format Shape uses.
  return (
    <OfficeTaskPane
      title={t("toolbar.selectionPane")}
      closeLabel={`${t("outline.close")} · ${t("toolbar.selectionPane")}`}
      onClose={close}
      testIdPrefix="document-selection"
      testId="document-selection-pane"
      closeTestId="document-selection-pane-close"
      dataAttributes={{ "data-product-scope": "mrai-extension" }}
      dockRight={false}
      className={dockStyles.pane}
      headerActions={
        <span className={styles.count}>{snapshot.targets.length}</span>
      }
    >
      <div
        className={styles.actionRail}
        role="toolbar"
        aria-label={t("toolbar.objectGroup")}
      >
        <button
          type="button"
          className={styles.actionButton}
          disabled={!canGroup}
          onClick={groupSelectedShapes}
          data-testid="document-selection-pane-group"
        >
          <DocumentDialogIcon icon="object.group" />
          {t("toolbar.objectGroup")}
        </button>
        <button
          type="button"
          className={styles.actionButton}
          disabled={!canUngroup}
          onClick={ungroupSelectedShape}
          data-testid="document-selection-pane-ungroup"
        >
          <DocumentDialogIcon icon="object.ungroup" />
          {t("presentation.group.ungroup")}
        </button>
        {/* Word's Selection Pane: Show All / Hide All, over every listed
            object that has an eye (shapes, pictures, text boxes). */}
        {hideableTargets.length > 0 ? (<>
        <button
          type="button"
          className={styles.actionButton}
          disabled={!editor.isEditable || !hideableTargets.some(shapeHidden)}
          onClick={() => setVisibility(hideableTargets, false)}
          data-testid="document-selection-pane-show-all"
        >
          <DocumentDialogIcon icon="view.show" />
          {t("presentation.selectionPane.showAll")}
        </button>
        <button
          type="button"
          className={styles.actionButton}
          disabled={!editor.isEditable || hideableTargets.every(shapeHidden)}
          onClick={() => setVisibility(hideableTargets, true)}
          data-testid="document-selection-pane-hide-all"
        >
          <DocumentDialogIcon icon="view.hide" />
          {t("presentation.selectionPane.hideAll")}
        </button>
        </>) : null}
        <button
          type="button"
          className={styles.actionButton}
          disabled={!canReorderSelected || singleShapeIndex === 0}
          onClick={() =>
            singleSelectedShape &&
            reorderShape(singleSelectedShape, "bringForward")
          }
          title={t("presentation.contextMenu.bringForward")}
          data-testid="document-selection-pane-move-up"
        >
          <ChevronUp size={13} aria-hidden />{" "}
          {t("presentation.contextMenu.bringForward")}
        </button>
        <button
          type="button"
          className={styles.actionButton}
          disabled={
            !canReorderSelected || singleShapeIndex === shapeOrder.count - 1
          }
          onClick={() =>
            singleSelectedShape &&
            reorderShape(singleSelectedShape, "sendBackward")
          }
          title={t("presentation.contextMenu.sendBackward")}
          data-testid="document-selection-pane-move-down"
        >
          <ChevronDown size={13} aria-hidden />{" "}
          {t("presentation.contextMenu.sendBackward")}
        </button>
      </div>
      {!snapshot.shapeStackComplete ? (
        <div className={styles.extensionNote} role="note">
          <span className={styles.extensionBadge}>
            {t("toolbar.mraiExtension")}
          </span>
          <span>{t("toolbar.selectionPaneUnavailable")}</span>
        </div>
      ) : null}
      <OfficeSelectionPaneList
        isEmpty={snapshot.targets.length === 0}
        emptyLabel={t("goTo.empty")}
        emptyTestId="document-selection-pane-empty"
        listLabel={t("goTo.kind.object")}
      >
        {snapshot.targets.map((target, index) => {
          const selected = effectiveSelectedIds.has(target.id);
          const type = objectType(target);
          const isShape = type === "documentShape";
          const shapeIndex = shapeOrder.indexById.get(target.id) ?? -1;
          const shapeCount = shapeOrder.count;
          const hideable = canHideTarget(target);
          const hidden = hideable && shapeHidden(target);
          const label = labelFor(target);
          const visibilityLabel = hidden
            ? t("presentation.selectionPane.showObject")
            : t("presentation.selectionPane.hideObject");
          const showReorder =
            isShape &&
            effectiveSelectedIds.size <= 1 &&
            snapshot.shapeStackComplete &&
            shapeCount > 1;
          return (
            <OfficeSelectionPaneRow
              key={target.id}
              rowRef={(element) => {
                if (element) rowRefs.current.set(target.id, element);
                else rowRefs.current.delete(target.id);
              }}
              icon={typeIcon(type)}
              label={label}
              ordinal={target.ordinal}
              selected={selected}
              current={target.id === snapshot.selectedId}
              hidden={hidden}
              reorderable={
                editor.isEditable &&
                isShape &&
                snapshot.shapeStackComplete &&
                shapeCount > 1
              }
              tabIndex={target.id === effectiveRovingId ? 0 : -1}
              testId="document-selection-pane-row"
              rowData={{ "data-object-type": type }}
              onFocus={() => setRovingId(target.id)}
              onClick={(event) => {
                activate(target, event.metaKey || event.ctrlKey);
                // a finger's double tap renames, as a double-click does
                
                if (
                  isShape &&
                  editor.isEditable &&
                  isDoubleTap(event.nativeEvent, target.id)
                ) {
                  setRenaming({ id: target.id, draft: label });
                }
              }}
              onKeyDown={(event) => handleRowKeyDown(event, target, index)}
              onLabelDoubleClick={(event) => {
                if (!isShape || !editor.isEditable) return;
                event.preventDefault();
                setRenaming({ id: target.id, draft: label });
              }}
              rename={
                renaming?.id === target.id
                  ? {
                      draft: renaming.draft,
                      inputLabel: t("presentation.selectionPane.objectName"),
                      testId: "document-selection-pane-rename-input",
                      onChange: (draft) =>
                        setRenaming({ id: target.id, draft }),
                      onCommit: commitRename,
                      onCancel: () => setRenaming(null),
                    }
                  : null
              }
              visibility={
                hideable
                  ? {
                      hidden,
                      disabled: !editor.isEditable,
                      title: visibilityLabel,
                      ariaLabel: `${visibilityLabel} · ${label}`,
                      testId: "document-selection-pane-visibility",
                      onToggle: () => setVisibility([target], !hidden),
                    }
                  : null
              }
              reorder={
                showReorder
                  ? {
                      label: t("toolbar.objectLayer"),
                      actions: (
                        [
                          ["bringToFront", ArrowUpToLine, shapeIndex === 0],
                          ["bringForward", ChevronUp, shapeIndex === 0],
                          [
                            "sendBackward",
                            ChevronDown,
                            shapeIndex === shapeCount - 1,
                          ],
                          [
                            "sendToBack",
                            ArrowDownToLine,
                            shapeIndex === shapeCount - 1,
                          ],
                        ] as const
                      ).map(([action, Icon, disabled]) => ({
                        key: action,
                        label: t(`presentation.contextMenu.${action}`),
                        icon: <Icon size={13} aria-hidden />,
                        disabled: !editor.isEditable || disabled,
                        testId: `document-selection-pane-${action}`,
                        onClick: () => reorderShape(target, action),
                      })),
                    }
                  : null
              }
            />
          );
        })}
      </OfficeSelectionPaneList>
    </OfficeTaskPane>
  );
}
