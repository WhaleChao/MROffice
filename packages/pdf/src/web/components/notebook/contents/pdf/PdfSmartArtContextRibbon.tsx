"use client";










/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { useTranslations } from "@mrnote/core/host/next-intl.tsx";

import { OfficeSmartArtDesignTab, type OfficeSmartArtDesignTabDiagram } from "@mrnote/core/web/components/office-ribbon/OfficeSmartArtDesignTab.tsx";
import { OfficeSmartArtTextPane } from "@mrnote/core/web/components/office-ribbon/OfficeSmartArtTextPane.tsx";
import {
  useOfficeSmartArtDesignLabels,
  useOfficeSmartArtPlaceholders,
  useOfficeSmartArtTextPaneLabels,
} from "@mrnote/core/web/components/office-ribbon/officeSmartArtPlaceholders.ts";
import { toast } from "@mrnote/core/web/hooks/use-toast.ts";
import { readDocumentSmartArtDiagram } from "@mrnote/core/web/lib/office-smartart/document-design-model.ts";
import {
  SMARTART_NO_OUTLINE_COMMANDS,
  smartArtOutlineAfterAdd,
  smartArtOutlineAfterDelete,
  smartArtOutlineAfterDemote,
  smartArtOutlineAfterInsert,
  smartArtOutlineAfterMerge,
  smartArtOutlineAfterMove,
  smartArtOutlineAfterPaste,
  smartArtOutlineAfterPromote,
  smartArtOutlineAfterTextEdit,
  smartArtToggledDirection,
} from "@mrnote/core/web/lib/office-smartart/design-model.ts";
import { smartArtCommandItemId, smartArtItemCommandState } from "@mrnote/core/web/lib/office-smartart/interaction-model.ts";
import { smartArtUnusedItemIds } from "@mrnote/core/web/lib/office-smartart/model.ts";
import { MAX_SMARTART_OUTLINE_TEXT_LENGTH } from "@mrnote/core/web/lib/office-smartart/outline.ts";
import { smartArtEngine } from "@mrnote/core/web/lib/office-smartart/remote.ts";
import type { PdfChart } from "../../../../lib/pdf/managed-chart";

type Diagram = NonNullable<ReturnType<typeof readDocumentSmartArtDiagram>>;
type Outline = Diagram["outline"];
type OutlineChange = { items: Outline extends readonly (infer T)[] ? T[] : never; selectedItemId: string } | null;

/** The group with an engine patch merged onto it, as the Docs host merges
 * one onto the node: a key the patch states as undefined goes. */
export function pdfSmartArtAttrsWithPatch(attrs: Record<string, unknown>,
  patch: Record<string, unknown>): Record<string, unknown> {
  const merged: Record<string, unknown> = { ...attrs };
  for (const [key, value] of Object.entries(patch)) {
    if (value === undefined) delete merged[key];
    else merged[key] = value;
  }
  return merged;
}

export interface PdfSmartArtContextRibbonProps {
  readonly smartArt: PdfChart;
  readonly writable: boolean;
  readonly onChange: (smartArt: PdfChart) => void;
  /** The graphic on the page, which the Text Pane stands beside. */
  readonly anchor: () => Element | null;
}

export function PdfSmartArtContextRibbon({ smartArt, writable, onChange, anchor }: PdfSmartArtContextRibbonProps) {
  const t = useTranslations("console-notebooks");
  const labels = useOfficeSmartArtDesignLabels();
  const paneLabels = useOfficeSmartArtTextPaneLabels();
  const placeholders = useOfficeSmartArtPlaceholders();
  const diagram = useMemo(() => readDocumentSmartArtDiagram(smartArt.attrs as never), [smartArt.attrs]);
  const [textPaneOpen, setTextPaneOpen] = useState(false);
  const [paneItemId, setPaneItemId] = useState<string | null>(null);
  const [busy, setBusy] = useState(false);
  const textPaneTriggerRef = useRef<HTMLButtonElement | null>(null);
  // The record a reply is merged onto: the latest, not the one asked from.
  const latestRef = useRef(smartArt);
  latestRef.current = smartArt;
  const callRef = useRef<AbortController | null>(null);
  useEffect(() => () => callRef.current?.abort(), []);

  /** One engine answer for the graphic, merged and written; the latest ask wins. */
  const commit = useCallback((ask: (diagram: Diagram, signal: AbortSignal) => Promise<{ patch: unknown } | null>) => {
    if (!writable || !diagram) return false;
    callRef.current?.abort();
    const controller = new AbortController();
    callRef.current = controller;
    setBusy(true);
    void ask(diagram, controller.signal)
      .then((answer) => {
        if (controller.signal.aborted) return;
        const patch = answer?.patch as Record<string, unknown> | null | undefined;
        if (!patch) {
          toast({ title: t("toolbar.smartartLayoutCannotInsertedAsNative"), variant: "destructive" });
          return;
        }
        const current = latestRef.current;
        if (current.id !== smartArt.id) return;
        onChange({ ...current, attrs: pdfSmartArtAttrsWithPatch(current.attrs, patch) });
      })
      .catch((error: unknown) => {
        if (controller.signal.aborted || (error instanceof DOMException && error.name === "AbortError")) return;
        toast({ title: t("pdf.smartArt.engineUnavailable"), variant: "destructive" });
      })
      .finally(() => { if (callRef.current === controller) setBusy(false); });
    return true;
  }, [diagram, onChange, smartArt.id, t, writable]);

  const rebuild = useCallback((outline: Outline, extra: Record<string, unknown> = {}) =>
    commit((current, signal) => smartArtEngine("document.rebuild",
      { attrs: current.attrs, kind: current.kind, outline: [...outline], ...extra } as never, { signal })),
  [commit]);
  /** A Text Pane or Design command's outline change, then the relayout it implies. */
  const commitOutline = useCallback((mutate: (diagram: Diagram, selectedItemId: string) => OutlineChange,
    command?: "promote" | "demote" | "addAfter" | "moveUp" | "moveDown") => {
    if (!diagram) return false;
    const selected = command
      ? smartArtCommandItemId({ outline: diagram.outline, command: command as never, paneItemId })
      : diagram.outline.find((item) => item.id === paneItemId)?.id ?? diagram.outline[0]?.id;
    if (!selected) return false;
    const result = mutate(diagram, selected);
    if (!result) return false;
    setPaneItemId(result.selectedItemId);
    return rebuild(result.items as unknown as Outline);
  }, [diagram, paneItemId, rebuild]);

  if (!diagram) return null;
  const commands = writable
    ? smartArtItemCommandState({ outline: diagram.outline, kind: diagram.kind, paneItemId })
    : SMARTART_NO_OUTLINE_COMMANDS;
  const tabDiagram: OfficeSmartArtDesignTabDiagram = {
    kind: diagram.kind,
    styleId: diagram.styleId,
    rightToLeft: diagram.direction === "rtl",
    members: diagram.members as never,
    layoutUrn: diagram.layoutUrn,
    colorsId: diagram.colorsId,
    colors: diagram.colors ?? undefined,
    writable: writable && !busy,
    hasOutline: diagram.outline.length > 0,
    textPaneOpen,
    canPromote: commands.promote,
    canDemote: commands.demote,
    onToggleTextPane: () => setTextPaneOpen((open) => !open),
    onPromote: () => { commitOutline((current, id) => smartArtOutlineAfterPromote(current.outline, id, current.kind) as OutlineChange, "promote"); },
    onDemote: () => { commitOutline((current, id) => smartArtOutlineAfterDemote(current.outline, id, current.kind) as OutlineChange, "demote"); },
    onRightToLeft: () => { rebuild(diagram.outline, { direction: smartArtToggledDirection(diagram.direction) }); },
    onChangeLayout: (kind, layoutUrn) => {
      commit((current, signal) => smartArtEngine("document.rebuild", {
        attrs: current.attrs, kind, outline: [...current.outline],
        ...(layoutUrn !== undefined ? { layoutUrn } : {}),
      } as never, { signal }));
    },
    onChangeColors: (colors, nodePaint, colorsId) => {
      commit((current, signal) => smartArtEngine("document.paint", {
        attrs: current.attrs, colors, authorPaint: true, nodePaint: nodePaint ? [...nodePaint] : null,
        ...(colorsId !== undefined ? { colorsId } : {}),
      } as never, { signal }));
    },
    onChangeStyle: (styleId) => {
      commit((current, signal) => smartArtEngine("document.paint", { attrs: current.attrs, styleId } as never, { signal }));
    },
    onReset: () => {
      commit((current, signal) => smartArtEngine("document.reset", { attrs: current.attrs } as never, { signal }));
    },
    // The graphic's alternative text is its words, written as the
    // annotation's contents (the writer's); no pane of its own here.
    onAltText: () => undefined,
    commands,
    onAddShape: (position) => {
      commitOutline((current, id) => smartArtOutlineAfterInsert(current.outline, id, position, "", current.kind) as OutlineChange, "addAfter");
    },
    onMoveUp: () => { commitOutline((current, id) => smartArtOutlineAfterMove(current.outline, id, "up", current.kind) as OutlineChange, "moveUp"); },
    onMoveDown: () => { commitOutline((current, id) => smartArtOutlineAfterMove(current.outline, id, "down", current.kind) as OutlineChange, "moveDown"); },
    hasSelectedItem: paneItemId !== null,
  };
  return (
    <>
      <div data-pdf-ribbon-tab="smartart-design" data-testid="pdf-smartart-design-ribbon" style={{ display: "contents" }}>
        <OfficeSmartArtDesignTab
          testIdPrefix="pdf"
          application="docs"
          labels={labels}
          diagram={tabDiagram}
          preserved={null}
          placeholders={placeholders}
          textPaneTriggerRef={textPaneTriggerRef}
        />
      </div>
      {textPaneOpen ? (
        <OfficeSmartArtTextPane
          testIdPrefix="pdf"
          labels={paneLabels}
          anchor={anchor}
          layoutId={diagram.kind}
          outline={diagram.outline}
          unusedItemIds={smartArtUnusedItemIds(diagram.members as never)}
          selectedItemId={paneItemId}
          maxTextLength={MAX_SMARTART_OUTLINE_TEXT_LENGTH}
          canAddItem={writable}
          busy={busy}
          itemState={(item) => ({
            canPromote: Boolean(writable && smartArtOutlineAfterPromote(diagram.outline, item.id, diagram.kind)),
            canDemote: Boolean(writable && smartArtOutlineAfterDemote(diagram.outline, item.id, diagram.kind)),
            canDelete: Boolean(writable && smartArtOutlineAfterDelete(diagram.outline, item.id)),
            canMoveUp: Boolean(writable && smartArtOutlineAfterMove(diagram.outline, item.id, "up", diagram.kind)),
            canMoveDown: Boolean(writable && smartArtOutlineAfterMove(diagram.outline, item.id, "down", diagram.kind)),
            readOnly: !writable,
          })}
          preventDeleteFocus
          onClose={() => {
            setTextPaneOpen(false);
            textPaneTriggerRef.current?.focus({ preventScroll: true });
          }}
          onSelect={setPaneItemId}
          onCommitText={(itemId, text) => {
            commitOutline((current) => smartArtOutlineAfterTextEdit(current.outline, itemId, text) as OutlineChange);
          }}
          onAddAfter={(itemId, text, nextText) => {
            commitOutline((current) => smartArtOutlineAfterAdd(current.outline, itemId, nextText ?? "", current.kind,
              { itemId, text }) as OutlineChange);
          }}
          onAddShape={(itemId, text) => {
            commitOutline((current) => smartArtOutlineAfterInsert(current.outline, itemId, "after", "", current.kind,
              { itemId, text }) as OutlineChange);
          }}
          onPromote={(itemId, text) => {
            commitOutline((current) => smartArtOutlineAfterPromote(current.outline, itemId, current.kind, { itemId, text }) as OutlineChange);
          }}
          onDemote={(itemId, text) => {
            commitOutline((current) => smartArtOutlineAfterDemote(current.outline, itemId, current.kind, { itemId, text }) as OutlineChange);
          }}
          onDelete={(itemId) => {
            commitOutline((current) => smartArtOutlineAfterDelete(current.outline, itemId) as OutlineChange);
          }}
          onMoveUp={(itemId, text) => {
            commitOutline((current) => smartArtOutlineAfterMove(current.outline, itemId, "up", current.kind, { itemId, text }) as OutlineChange);
          }}
          onMoveDown={(itemId, text) => {
            commitOutline((current) => smartArtOutlineAfterMove(current.outline, itemId, "down", current.kind, { itemId, text }) as OutlineChange);
          }}
          onMerge={(keepId, text, goneId) => {
            commitOutline((current) => smartArtOutlineAfterMerge(current.outline, keepId, goneId, text) as OutlineChange);
          }}
          onPasteLines={(itemId, text, lines) => {
            commitOutline((current) => smartArtOutlineAfterPaste(current.outline, itemId, text, lines) as OutlineChange);
          }}
        />
      ) : null}
    </>
  );
}
