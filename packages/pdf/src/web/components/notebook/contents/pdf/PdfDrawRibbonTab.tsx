"use client";























/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { ReactNode } from "react";

import { DropdownMenu, DropdownMenuTrigger } from "@mrnote/core/web/components/ui/dropdown-menu.tsx";
import { RibbonDropdownMenuContent } from "@mrnote/core/web/components/office-ribbon/CupertinoRibbonMenu.tsx";
import {
  OfficeHomeCommand,
  OfficeHomeCommandRow,
  OfficeHomeLargeSplitLayout,
  OfficeHomeMenuTrigger,
  OfficeHomeRibbonGroup,
} from "@mrnote/core/web/components/office-ribbon/OfficeHomeRibbon.tsx";
import { OfficeColorPalettePopover } from "@mrnote/core/web/components/office-ribbon/OfficeColorPalettePopover.tsx";
import {
  OfficeInkAddPenMenuItems,
  OfficeInkEraserMenuItems,
  OfficeInkHideRibbonCommand,
  OfficeInkPenGallery,
  OfficeInkTouchDrawCommand,
  OfficeInkTrackpadSwitch,
  useInkDrawTools,
} from "@mrnote/core/web/components/office-ribbon/OfficeInkRibbon.tsx";
import { OfficeSemanticIcon } from "@mrnote/core/web/components/office-ribbon/icons/OfficeSemanticIcon.tsx";
import { Plus, Undo2 } from "@mrnote/core/web/components/office-ribbon/icons/houseIcons.tsx";
import { IconInkRuler } from "@mrnote/core/web/components/notebook/contents/PresentationRibbonIcons.tsx";
import DocumentInkTrackpad from "@mrnote/docs/web/components/console/editor/native-document/DocumentInkTrackpad.tsx";
import { usePdfRibbonGroupShown } from "./PdfOfficeRibbon";
import {
  DOCUMENT_INK_MAX_PENS,
  type DocumentInkEraser,
  type DocumentInkPoint,
  type DocumentInkPreset,
} from "@mrnote/core/web/lib/document-ink.ts";
import { officeInkPenById } from "@mrnote/ink/web/lib/office-ink-pens.ts";
import { PDF_INK_PEN_TOOLS } from "../../../../lib/pdf/ink-pens";
import { useDeviceHasTouch } from "@mrnote/core/web/lib/draw-with-touch.ts";
import type { useTranslations } from "@mrnote/core/host/next-intl.tsx";
import { useRef, useState } from "react";

import type { PdfInkPenCommands } from "./window/usePdfInkPens";

/** The Draw tab's armed tool; null is Select. */
export type PdfInkTool = "draw" | "eraser" | "lasso";

/** Draw with Trackpad (Word for Mac): the pad over the page and where its strokes land. */
export interface PdfDrawRibbonTrackpad {
  readonly on: boolean;
  readonly toggle: () => void;
  readonly close: () => void;
  /** The pen the pad draws with. */
  readonly preset: DocumentInkPreset;
  /** The element the pages scroll in (the pad dims and frames it). */
  readonly anchor: () => Element | null;
  /** The ribbon above it, left lit. */
  readonly below: () => Element | null;
  /** The page in view, in client pixels: the pad never leaves it. */
  readonly page: (area: { left: number; top: number; right: number; bottom: number }) => {
    left: number;
    top: number;
    right: number;
    bottom: number;
  } | null;
  /** One finished stroke, in client pixels. */
  readonly onStroke: (points: readonly DocumentInkPoint[]) => void;
}

export interface PdfDrawRibbonModel {
  /** The page can be written (a read-only page keeps every tool off). */
  readonly writable: boolean;
  readonly disabledReason?: string;
  readonly tool: PdfInkTool | null;
  readonly setTool: (tool: PdfInkTool | null) => void;
  readonly eraserKind: DocumentInkEraser;
  /** Picks an eraser kind and arms the Eraser with it (Word). */
  readonly chooseEraserKind: (kind: DocumentInkEraser) => void;
  /** The page has ink the Eraser can take. */
  readonly inkPresent: boolean;
  /** Draw ▸ Lasso Select is offered (a host whose ink it can move); absent, no Lasso. */
  readonly lasso?: boolean;
  /** Draw ▸ Ruler, the straightedge a pen draws along (lib/office-ink-ruler); absent, no Ruler. */
  readonly ruler?: { readonly on: boolean; readonly toggle: () => void };
  readonly pens: PdfInkPenCommands;
  /** Undo Last Stroke, the PDF's own command beside the pens. */
  readonly undoLastStroke?: { readonly run: () => void; readonly disabled: boolean };
  readonly trackpad?: PdfDrawRibbonTrackpad;
  /** Pen-first drawing: folds the ribbon into the pen palette (null while the palette is not available). */
  readonly hideRibbon: (() => void) | null;
}

export interface PdfDrawRibbonTabProps {
  t: ReturnType<typeof useTranslations<"console-notebooks">>;
  model: PdfDrawRibbonModel;
  /** The ribbon tab token the groups belong to. */
  tab?: string;
  testIdPrefix?: string;
  /** Pen names as the host's labels say them. */
  penLabel?: (pen: DocumentInkPreset) => string;
  /** Extra commands after the pens (a host's own). */
  children?: ReactNode;
}

export function PdfDrawRibbonTab({
  t,
  model,
  tab = "draw",
  testIdPrefix = "pdf-draw",
  penLabel,
  children,
}: PdfDrawRibbonTabProps) {
  const {
    writable,
    disabledReason,
    tool,
    setTool,
    eraserKind,
    chooseEraserKind,
    inkPresent,
    lasso,
    ruler,
    pens,
    undoLastStroke,
    trackpad,
    hideRibbon,
  } = model;
  const deviceHasTouch = useDeviceHasTouch();
  const inkDrawTools = useInkDrawTools(tool === "draw");
  // Off its tab the PDF ribbon draws none of a group's commands (the
  // window's renders while scrolling do not rebuild every tab).
  const shown = usePdfRibbonGroupShown(tab);
  const deniedReason = writable ? undefined : disabledReason;
  // More Colors: the pen whose colour the shared palette picks, and its face.
  const [moreColors, setMoreColors] = useState<{
    penId: string;
    anchor: { current: HTMLButtonElement | null };
  } | null>(null);
  const penRefs = useRef(new Map<string, HTMLButtonElement | null>());
  const eraserDisabled = !writable || !inkPresent;
  const eraserReason = !writable ? deniedReason : !inkPresent ? t("toolbar.drawNoInk") : undefined;
  if (!shown) return null;

  return (
    <>
      <OfficeHomeRibbonGroup
        ribbon={tab}
        label={t("toolbar.ribbonDrawTools")}
        size="content"
        testId={`${testIdPrefix}-tools-group`}
      >
        <OfficeHomeCommandRow>
        <OfficeHomeCommand
          controlId="PDFCTL.RIBBON.DRAW.TOOL.SELECT"
          variant="large"
          active={tool === null}
          icon={<OfficeSemanticIcon semanticId="draw.select" size={24} />}
          label={t("toolbar.select")}
          onClick={() => setTool(null)}
          testId={`${testIdPrefix}-select`}
        />
        {lasso ? (
          <OfficeHomeCommand
            controlId="PDFCTL.RIBBON.DRAW.TOOL.LASSO"
            variant="large"
            active={tool === "lasso"}
            disabled={eraserDisabled}
            disabledReason={eraserReason}
            icon={<OfficeSemanticIcon semanticId="draw.lasso" size={24} />}
            label={t("toolbar.drawLasso")}
            onClick={() => setTool(tool === "lasso" ? null : "lasso")}
            testId={`${testIdPrefix}-lasso`}
          />
        ) : null}
        <OfficeHomeCommand
          controlId="PDFCTL.RIBBON.DRAW.TOOL.DRAW_MODE"
          variant="large"
          active={inkDrawTools.penDrawActive}
          disabled={!writable}
          disabledReason={deniedReason}
          icon={<OfficeSemanticIcon semanticId="draw.mode" size={24} />}
          label={t("toolbar.drawMode")}
          onClick={() => {
            // Draw with Touch lit: Draw keeps the draw mode on and only
            // stops the finger inking (Docs).
            if (tool === "draw" && !inkDrawTools.penDrawActive) {
              inkDrawTools.choosePenDraw();
              return;
            }
            if (tool === "draw") {
              setTool(null);
              return;
            }
            inkDrawTools.choosePenDraw();
            setTool("draw");
          }}
          testId={`${testIdPrefix}-toggle`}
        />
        <OfficeInkTouchDrawCommand
          inkDrawing={tool === "draw"}
          controlId="PDFCTL.RIBBON.DRAW.TOOL.DRAW_TOUCH"
          disabled={!writable}
          disabledReason={deniedReason}
          testId={`${testIdPrefix}-touch`}
          onArm={() => setTool("draw")}
          onExit={() => setTool(null)}
        />
        <DropdownMenu modal={false}>
          <OfficeHomeLargeSplitLayout
            testId={`${testIdPrefix}-eraser-split`}
            primary={
              <OfficeHomeCommand
                controlId="PDFCTL.RIBBON.DRAW.TOOL.ERASER"
                variant="large"
                active={tool === "eraser"}
                disabled={eraserDisabled}
                disabledReason={eraserReason}
                icon={<OfficeSemanticIcon semanticId="draw.eraser.stroke" size={24} />}
                label={t("toolbar.drawEraser")}
                onClick={() => setTool(tool === "eraser" ? null : "eraser")}
                testId={`${testIdPrefix}-eraser`}
              />
            }
            trigger={
              <DropdownMenuTrigger asChild>
                <OfficeHomeMenuTrigger
                  data-control-id="PDFCTL.RIBBON.DRAW.TOOL.ERASER_MENU"
                  disabled={eraserDisabled}
                  disabledReason={eraserReason}
                  icon={<OfficeSemanticIcon semanticId="menu.chevron.down" size={12} />}
                  label={t("toolbar.drawEraserKinds")}
                  testId={`${testIdPrefix}-eraser-menu`}
                  variant="icon"
                />
              </DropdownMenuTrigger>
            }
          />
          <RibbonDropdownMenuContent
            align="start"
            data-testid={`${testIdPrefix}-eraser-menu-content`}
            onCloseAutoFocus={(event) => event.preventDefault()}
          >
            <OfficeInkEraserMenuItems
              value={eraserKind}
              testIdPrefix={`${testIdPrefix}-eraser`}
              onChoose={chooseEraserKind}
            />
          </RibbonDropdownMenuContent>
        </DropdownMenu>
        </OfficeHomeCommandRow>
      </OfficeHomeRibbonGroup>

      <OfficeHomeRibbonGroup
        ribbon={tab}
        label={t("toolbar.ribbonDrawPens")}
        size="content"
        testId={`${testIdPrefix}-pens-group`}
      >
        <OfficeHomeCommandRow>
        <OfficeInkPenGallery
          pens={pens.profile.pens}
          selectedId={pens.selectedId}
          drawing={tool === "draw"}
          disabled={!writable}
          disabledReason={deniedReason}
          testIdPrefix={`${testIdPrefix}-pen`}
          controlIdFor={() => "PDFCTL.RIBBON.DRAW.PEN.GALLERY"}
          effects={false}
          // PDF ink always follows the pen's pressure.
          pressure={false}
          penLabel={penLabel}
          onChoose={pens.choosePen}
          onUpdate={pens.updatePen}
          onRemove={pens.removePen}
          onMoreColors={(penId) =>
            setMoreColors({
              penId,
              anchor: { current: penRefs.current.get(penId) ?? null },
            })
          }
          penRef={(id, element) => {
            penRefs.current.set(id, element);
          }}
        />
        <DropdownMenu modal={false}>
          <DropdownMenuTrigger asChild>
            <OfficeHomeMenuTrigger
              controlId="PDFCTL.RIBBON.DRAW.PEN.ADD"
              disabled={!writable || pens.profile.pens.length >= DOCUMENT_INK_MAX_PENS}
              disabledReason={deniedReason}
              displayLabel={t("toolbar.drawAddPen")}
              hasMenuIndicator
              icon={<Plus size={24} aria-hidden />}
              label={t("toolbar.drawAddPen")}
              testId={`${testIdPrefix}-add-pen`}
              variant="large"
            />
          </DropdownMenuTrigger>
          <RibbonDropdownMenuContent align="start" data-testid={`${testIdPrefix}-add-pen-menu`}>
            <OfficeInkAddPenMenuItems
              testIdPrefix={`${testIdPrefix}-add-pen`}
              tools={PDF_INK_PEN_TOOLS}
              onAdd={pens.addPen}
            />
          </RibbonDropdownMenuContent>
        </DropdownMenu>
        {undoLastStroke ? (
          <OfficeHomeCommand
            controlId="PDFCTL.RIBBON.DRAW.UNDO_STROKE"
            variant="large"
            disabled={undoLastStroke.disabled}
            icon={<Undo2 size={24} />}
            label={t("pdf.drawUndo")}
            onClick={undoLastStroke.run}
            testId={`${testIdPrefix}-undo`}
          />
        ) : null}
        {children}
        </OfficeHomeCommandRow>
        {moreColors && writable ? (
          <OfficeColorPalettePopover
            anchorRef={moreColors.anchor}
            value={officeInkPenById(pens.profile, moreColors.penId).color}
            labels={{
              themeColors: t("pdf.annotationFormat.themeColors"),
              standardColors: t("pdf.annotationFormat.standardColors"),
              recentColors: t("pdf.annotationFormat.recentColors"),
              moreColors: t("toolbar.drawMoreColors"),
            }}
            onPick={(color) => {
              const { penId } = moreColors;
              setMoreColors(null);
              if (color) pens.updatePen(penId, { color: color.toLowerCase() });
            }}
            onDismiss={() => setMoreColors(null)}
            testIdPrefix={`${testIdPrefix}-pen-palette`}
          />
        ) : null}
      </OfficeHomeRibbonGroup>

      {/* Draw ▸ Ruler (Stencils), as on the Docs and deck Draw tabs: the
          straightedge a pen draws along; the page rulers stay on View. */}
      {ruler ? (
        <OfficeHomeRibbonGroup
          ribbon={tab}
          label={t("presentation.draw.stencilsGroup")}
          size="content"
          testId={`${testIdPrefix}-stencils-group`}
        >
          <OfficeHomeCommandRow>
            <OfficeHomeCommand
              controlId="PDFCTL.RIBBON.DRAW.RULER"
              variant="large"
              active={ruler.on}
              disabled={!writable}
              disabledReason={deniedReason}
              icon={<IconInkRuler home size={24} />}
              label={t("toolbar.ruler")}
              onClick={ruler.toggle}
              testId={`${testIdPrefix}-ruler`}
            />
          </OfficeHomeCommandRow>
        </OfficeHomeRibbonGroup>
      ) : null}

      {/* The way a finger or a trackpad draws, last on the tab as on the
          Docs Draw tab: Draw with Touch stands among the tools on a touch
          screen; a Mac draws with its trackpad instead. */}
      {trackpad && !deviceHasTouch ? (
        <OfficeHomeRibbonGroup
          ribbon={tab}
          label={t("toolbar.drawTrackpad")}
          size="content"
          testId={`${testIdPrefix}-trackpad-group`}
        >
          <OfficeHomeCommandRow>
          <OfficeInkTrackpadSwitch
            on={trackpad.on}
            controlId="PDFCTL.RIBBON.DRAW.TRACKPAD"
            disabled={!writable}
            disabledReason={deniedReason}
            testId={`${testIdPrefix}-trackpad`}
            onToggle={trackpad.toggle}
          />
          <DocumentInkTrackpad
            open={trackpad.on && writable}
            preset={trackpad.preset}
            title={t("toolbar.drawTrackpad")}
            hint={t("toolbar.drawTrackpadHint")}
            anchor={trackpad.anchor}
            below={trackpad.below}
            page={trackpad.page}
            onStroke={trackpad.onStroke}
            onClose={trackpad.close}
          />
          </OfficeHomeCommandRow>
        </OfficeHomeRibbonGroup>
      ) : null}

      {/* The ribbon brought back over the pen palette folds into it again
          from here (lib/pen-first-draw). */}
      {hideRibbon ? (
        <OfficeHomeRibbonGroup
          ribbon={tab}
          label={t("toolbar.penPalette")}
          size="content"
          testId={`${testIdPrefix}-pen-palette-group`}
        >
          <OfficeHomeCommandRow>
            <OfficeInkHideRibbonCommand
              controlId="PDFCTL.RIBBON.DRAW.HIDE_RIBBON"
              testId={`${testIdPrefix}-hide-ribbon`}
              onHide={hideRibbon}
            />
          </OfficeHomeCommandRow>
        </OfficeHomeRibbonGroup>
      ) : null}
    </>
  );
}
