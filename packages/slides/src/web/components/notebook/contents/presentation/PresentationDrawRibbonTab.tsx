"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The Draw tab of the presentation Ribbon (Drawing Tools, Pens, Convert,
 * Replay, Stencils), drawn from the window bindings it reads.
 *
 * Extracted verbatim from the render of
 * `components/notebook/contents/PresentationWindow.tsx`: the JSX is unchanged;
 * the window passes the bindings it used to close over as `ctl`, and the tab
 * condition and key stay at the call site. The only hook here is this
 * component's own device check, so the window's hook order is unchanged.
 *
 * The groups stand in the order every MRAI Draw tab shares (Office's):
 * Drawing Tools · Pens · Convert · Replay · Stencils · Draw with Touch or
 * Draw with Trackpad (one per device).
 */

import {
  OfficeHomeRibbonGroup,
  OfficeHomeCommandRow,
  OfficeHomeCommandStack,
  OfficeHomeCommand,
  OfficeHomeLargeSplitLayout,
  OfficeHomeMenuTrigger,
} from "@mrnote/core/web/components/office-ribbon/OfficeHomeRibbon.tsx";
import presentationChrome from "@mrnote/core/web/lib/presentation/presentation-window-shell.module.css";
import {
  IconSelectMenu,
  IconDrawMode,
  IconInkEraser,
  IconInkRuler,
} from "@mrnote/core/web/components/notebook/contents/PresentationRibbonIcons.tsx";
import { OfficeSemanticIcon } from "@mrnote/core/web/components/office-ribbon/icons/OfficeSemanticIcon.tsx";
import { DropdownMenu, DropdownMenuTrigger } from "@mrnote/core/web/components/ui/dropdown-menu.tsx";
import { RibbonDropdownMenuContent as DropdownMenuContent, RibbonDropdownMenuContent } from "@mrnote/core/web/components/office-ribbon/CupertinoRibbonMenu.tsx";
import {
  OfficeInkEraserMenuItems,
  OfficeInkHideRibbonCommand,
  OfficeInkTouchDrawCommand,
  useInkDrawTools,
  OfficeInkTrackpadSwitch,
  OfficeInkPenGallery,
  OfficeInkAddPenMenuItems,
} from "@mrnote/core/web/components/office-ribbon/OfficeInkRibbon.tsx";
import { useDeviceHasTouch } from "@mrnote/core/web/lib/draw-with-touch.ts";
import type { OfficeInkConversionMode } from "@mrnote/core/web/lib/office-ink-conversion.ts";
import DocumentInkTrackpad from "@mrnote/docs/web/components/console/editor/native-document/DocumentInkTrackpad.tsx";
import { Plus } from "@mrnote/core/web/components/office-ribbon/icons/houseIcons.tsx";
import { PRESENTATION_INK_TOOLS } from "../../../../lib/presentation-ink-pens";
import type { useTranslations } from "@mrnote/core/host/next-intl.tsx";
import type { PresentationInkReplayState } from "./PresentationInkReplay";
import type { Dispatch, RefObject, SetStateAction } from "react";
import type { embeddedPresentationAutosave } from "../../../../lib/presentation-embedded-session";
import type { PresentationViewAidCommandAuthority, ResolvedPresentationFeatureCommand } from "@mrnote/core/web/components/notebook/contents/ribbon/commandRegistry.ts";
import type { PresentationInkTool } from "@/lib/presentation-ink-presets";
import type { OfficeInkPenProfile, withOfficeInkPen } from "@mrnote/ink/web/lib/office-ink-pens.ts";
import type { DocumentInkEraser, DocumentInkPreset } from "@mrnote/core/web/lib/document-ink.ts";
import type { DocumentInkTool } from "@mrnote/core/web/lib/document-shapes.ts";
import type { InkRulerState } from "@mrnote/core/web/lib/office-ink-ruler.ts";
import type { PresentationViewAidKind } from "@mrnote/core/web/lib/presentation-feature-capabilities.ts";

/** The window bindings this tab reads, passed through unchanged. */
export interface PresentationDrawRibbonTabCtl {
  t: ReturnType<typeof useTranslations<"console-notebooks">>;
  inkReplay: PresentationInkReplayState | null;
  setInkReplay: Dispatch<SetStateAction<PresentationInkReplayState | null>>;
  slideCanvasRef: RefObject<HTMLDivElement | null>;
  showRuler: boolean;
  /** Draw ▸ Ruler: the straightedge over the slide (not View ▸ Ruler's rulers). */
  inkRuler: InkRulerState | null;
  toggleInkRuler: () => void;
  autosave: ReturnType<typeof embeddedPresentationAutosave>;
  hasPersistedReviewInk: boolean;
  activeSlideInkLassoAvailable: boolean;
  viewAidCommandAuthority: PresentationViewAidCommandAuthority;
  inkTool: PresentationInkTool | null;
  inkInteractionMode: ("draw" | "eraser" | "lasso") | null;
  inkEraserKind: "small" | "stroke" | "medium";
  inkColor: string;
  inkWidthPt: number;
  inkTrackpadEnabled: boolean;
  writeInkTrackpadEnabled: (next: boolean) => void;
  placePresentationTrackpadStroke: (points: readonly { x: number; y: number; }[]) => void;
  /** Which of Ink to Shape / Ink to Math / Ink to Text is on (one at a time). */
  inkConversionMode: OfficeInkConversionMode | null;
  /** Ink to Math: converts the selected ink in place, else turns the mode on or off. */
  runPresentationInkToMath: () => void;
  /** Ink to Text: converts the selected ink to a text box, else turns the mode on or off. */
  runPresentationInkToText: () => void;
  inkPens: OfficeInkPenProfile;
  activeInkPresetId: string | null;
  writeInkInteractionMode: (next: ("draw" | "eraser" | "lasso") | null) => void;
  choosePresentationInkEraser: (kind: DocumentInkEraser) => void;
  selectInkPreset: (pen: DocumentInkPreset, enterDrawMode?: boolean) => void;
  addInkPen: (tool: DocumentInkTool) => void;
  updateInkPen: (id: string, patch: Parameters<typeof withOfficeInkPen>[2]) => void;
  removeInkPen: (id: string) => void;
  runPresentationInkToShape: () => void;
  presentationInkPenLabel: (pen: DocumentInkPreset) => string;
  togglePresentationViewAid: (kind: PresentationViewAidKind) => void;
  featureDecision: (featureId: string) => ResolvedPresentationFeatureCommand;
  executePresentationFeature: (featureId: string) => undefined;
  /** Pen-first drawing: folds the ribbon back into the pen palette (null while the palette is not available). */
  hideRibbon: (() => void) | null;
}

export function PresentationDrawRibbonTab({ ctl }: { ctl: PresentationDrawRibbonTabCtl }) {
  const {
    t,
    inkReplay,
    setInkReplay,
    slideCanvasRef,
    inkRuler,
    toggleInkRuler,
    autosave,
    hasPersistedReviewInk,
    activeSlideInkLassoAvailable,
    inkTool,
    inkInteractionMode,
    inkEraserKind,
    inkColor,
    inkWidthPt,
    inkTrackpadEnabled,
    writeInkTrackpadEnabled,
    placePresentationTrackpadStroke,
    inkConversionMode,
    runPresentationInkToMath,
    runPresentationInkToText,
    hideRibbon,
    inkPens,
    activeInkPresetId,
    writeInkInteractionMode,
    choosePresentationInkEraser,
    selectInkPreset,
    addInkPen,
    updateInkPen,
    removeInkPen,
    runPresentationInkToShape,
    presentationInkPenLabel,
    featureDecision,
    executePresentationFeature,
  } = ctl;
  // This component's own hook (the window's hook order is untouched).
  const deviceHasTouch = useDeviceHasTouch();
  const inkDrawTools = useInkDrawTools(inkInteractionMode === "draw");
  return (
                  <>

                    <OfficeHomeRibbonGroup
                      ribbon="draw"
                      // The group's name every MRAI Draw tab shares (Docs', Sheets').
                      label={t("toolbar.ribbonDrawTools")}
                      size="content"
                      className={presentationChrome.drawToolsGroup}
                    >
                      <OfficeHomeCommandRow>
                        <OfficeHomeCommand
                          variant="large"
                          active={inkInteractionMode === null}
                          icon={<IconSelectMenu home size={24} />}
                          label={t("toolbar.select")}
                          onClick={() => writeInkInteractionMode(null)}
                          testId="presentation-ribbon-draw-select"
                        />
                        <OfficeHomeCommand
                          controlId="PPTCTL.RIBBON.DRAW.TOOL.LASSO"
                          variant="large"
                          active={inkInteractionMode === "lasso"}
                          disabled={!activeSlideInkLassoAvailable}
                          icon={
                            <OfficeSemanticIcon
                              semanticId="draw.lasso"
                              size={24}
                            />
                          }
                          label={t("toolbar.drawLasso")}
                          onClick={() => writeInkInteractionMode("lasso")}
                          testId="presentation-ribbon-draw-lasso"
                        />
                        <OfficeHomeCommand
                          controlId="PPTCTL.RIBBON.DRAW.TOOL.DRAW_MODE"
                          variant="large"
                          active={inkDrawTools.penDrawActive}
                          disabled={!featureDecision("ink.drawMode").enabled}
                          icon={<IconDrawMode home size={24} />}
                          label={t("presentation.draw.drawMode")}
                          onClick={() => {
                            // Draw with Touch lit: Draw keeps the draw mode on
                            // and only stops the finger inking.
                            if (inkInteractionMode === "draw" && !inkDrawTools.penDrawActive) {
                              inkDrawTools.choosePenDraw();
                              return;
                            }
                            inkDrawTools.choosePenDraw();
                            executePresentationFeature("ink.drawMode");
                          }}
                          testId="presentation-ribbon-draw-mode"
                        />
                        {/* Draw with Touch, one of the five tools (useInkDrawTools). */}
                        <OfficeInkTouchDrawCommand
                          inkDrawing={inkInteractionMode === "draw"}
                          controlId="PPTCTL.RIBBON.DRAW.TOOL.DRAW_TOUCH"
                          disabled={!featureDecision("ink.drawMode").enabled}
                          testId="presentation-ribbon-draw-touch"
                          onArm={() => writeInkInteractionMode("draw")}
                          onExit={() => writeInkInteractionMode(null)}
                        />
                        <DropdownMenu modal={false}>
                          <OfficeHomeLargeSplitLayout
                            testId="presentation-ribbon-draw-eraser-split"
                            primary={
                              <OfficeHomeCommand
                                controlId="PPTCTL.RIBBON.DRAW.TOOL.ERASER"
                                variant="large"
                                active={inkInteractionMode === "eraser"}
                                disabled={!featureDecision("ink.erase").enabled}
                                icon={<IconInkEraser home size={24} />}
                                label={t("toolbar.drawEraser")}
                                onClick={() =>
                                  executePresentationFeature("ink.erase")
                                }
                                testId="presentation-ribbon-draw-eraser"
                              />
                            }
                            trigger={
                              <DropdownMenuTrigger asChild>
                                <OfficeHomeMenuTrigger
                                  data-control-id="PPTCTL.RIBBON.DRAW.TOOL.ERASER_MENU"
                                  disabled={!featureDecision("ink.erase").enabled}
                                  icon={<OfficeSemanticIcon semanticId="menu.chevron.down" size={12} />}
                                  label={t("toolbar.drawEraserKinds")}
                                  testId="presentation-ribbon-draw-eraser-menu"
                                  variant="icon"
                                />
                              </DropdownMenuTrigger>
                            }
                          />
                          <DropdownMenuContent
                            align="start"
                            data-testid="presentation-ribbon-draw-eraser-menu-content"
                            onCloseAutoFocus={(event) => event.preventDefault()}
                          >
                            <OfficeInkEraserMenuItems
                              value={inkEraserKind}
                              testIdPrefix="presentation-ribbon-draw-eraser"
                              onChoose={choosePresentationInkEraser}
                            />
                          </DropdownMenuContent>
                        </DropdownMenu>
                      </OfficeHomeCommandRow>
                    </OfficeHomeRibbonGroup>

                    <OfficeHomeRibbonGroup
                      controlId="PPTCTL.RIBBON.DRAW.PEN.GALLERY"
                      ribbon="draw"
                      label={t("presentation.draw.pensGroup")}
                      size="content"
                      className={presentationChrome.drawPensGroup}
                    >
                      <OfficeHomeCommandRow
                        className={presentationChrome.drawPensRow}
                      >
                        <OfficeInkPenGallery
                          pens={inkPens.pens}
                          selectedId={activeInkPresetId ?? ""}
                          drawing={inkInteractionMode === "draw"}
                          disabled={!featureDecision("ink.pen").enabled}
                          testIdPrefix="presentation-ribbon-draw"
                          controlIdFor={() => "PPTCTL.RIBBON.DRAW.PEN.GALLERY"}
                          penLabel={presentationInkPenLabel}
                          onChoose={(pen) => selectInkPreset(pen)}
                          onUpdate={updateInkPen}
                          onRemove={removeInkPen}
                        />
                        <DropdownMenu>
                          <DropdownMenuTrigger asChild>
                            <OfficeHomeMenuTrigger
                              controlId="PPTCTL.RIBBON.DRAW.PEN.ADD"
                              disabled={!featureDecision("ink.pen").enabled}
                              displayLabel={t("presentation.draw.addPen")}
                              hasMenuIndicator
                              icon={<Plus size={24} />}
                              label={t("presentation.draw.addPen")}
                              testId="presentation-ribbon-draw-add-pen"
                              variant="large"
                            />
                          </DropdownMenuTrigger>
                          <RibbonDropdownMenuContent
                            align="start"
                            data-testid="presentation-ribbon-draw-add-pen-menu"
                          >
                            <OfficeInkAddPenMenuItems
                              testIdPrefix="presentation-ribbon-draw-add-pen"
                              tools={PRESENTATION_INK_TOOLS}
                              onAdd={addInkPen}
                            />
                          </RibbonDropdownMenuContent>
                        </DropdownMenu>
                      </OfficeHomeCommandRow>
                    </OfficeHomeRibbonGroup>

                    {





}
                    <OfficeHomeRibbonGroup
                      ribbon="draw"
                      label={t("toolbar.drawConvert")}
                      size="content"
                      testId="presentation-ribbon-draw-convert-group"
                    >
                      <OfficeHomeCommandStack rows={3}>
                        <OfficeHomeCommand
                          controlId="PPTCTL.RIBBON.DRAW.CONVERT.INK_TO_SHAPE"
                          variant="compact"
                          active={inkConversionMode === "shape"}
                          disabled={
                            autosave.readOnly || !featureDecision("ink.pen").enabled
                          }
                          icon={<OfficeSemanticIcon semanticId="ink.toShape" size={16} />}
                          label={t("toolbar.drawInkToShape")}
                          onClick={runPresentationInkToShape}
                          testId="presentation-ribbon-draw-ink-to-shape"
                        />
                        {null}
                        {null}
                      </OfficeHomeCommandStack>
                    </OfficeHomeRibbonGroup>

                    <OfficeHomeRibbonGroup
                      ribbon="draw"
                      label={t("toolbar.drawReplay")}
                      size="content"
                      testId="presentation-ribbon-draw-replay-group"
                    >
                      <OfficeHomeCommand
                        controlId="PPTCTL.RIBBON.DRAW.REPLAY.INK"
                        variant="large"
                        active={inkReplay !== null}
                        disabled={!hasPersistedReviewInk}
                        icon={<OfficeSemanticIcon semanticId="media.play" size={24} />}
                        label={t("toolbar.drawInkReplay")}
                        onClick={() => {
                          writeInkInteractionMode(null);
                          setInkReplay({ position: 0, playing: true });
                        }}
                        testId="presentation-ribbon-draw-ink-replay"
                      />
                    </OfficeHomeRibbonGroup>

                    <OfficeHomeRibbonGroup
                      ribbon="draw"
                      label={t("presentation.draw.stencilsGroup")}
                      size="content"
                      className={presentationChrome.drawStencilGroup}
                    >
                      {/* PowerPoint's Draw ▸ Ruler is the straightedge a pen
                          draws along (lib/office-ink-ruler); the page rulers
                          stay on View ▸ Ruler. */}
                      <OfficeHomeCommand
                        variant="large"
                        active={inkRuler !== null}
                        disabled={!featureDecision("ink.pen").enabled}
                        icon={<IconInkRuler home size={24} />}
                        label={t("toolbar.ruler")}
                        onClick={toggleInkRuler}
                        testId="presentation-ribbon-draw-ruler"
                      />
                    </OfficeHomeRibbonGroup>

                    {/* The way a finger or a trackpad draws, last on the tab
                        as on the Docs and worksheet Draw tabs: one switch per
                        device (a touch screen draws with touch, a Mac with
                        its trackpad). */}
                    {deviceHasTouch ? null : (
                      <OfficeHomeRibbonGroup
                        ribbon="draw"
                        label={t("toolbar.drawTrackpad")}
                        size="content"
                        testId="presentation-ribbon-draw-trackpad-group"
                      >
                        <OfficeHomeCommandRow>
                          <OfficeInkTrackpadSwitch
                            on={inkTrackpadEnabled}
                            controlId="PPTCTL.RIBBON.DRAW.TRACKPAD"
                            disabled={!featureDecision("ink.trackpad").enabled}
                            testId="presentation-ribbon-draw-trackpad"
                            onToggle={() =>
                              executePresentationFeature("ink.trackpad")
                            }
                          />
                          <DocumentInkTrackpad
                            open={
                              inkTrackpadEnabled &&
                              featureDecision("ink.trackpad").enabled
                            }
                            preset={{
                              id: "presentation-trackpad",
                              tool: inkTool ?? "pen",
                              color: inkColor,
                              widthPt: inkWidthPt,
                            }}
                            title={t("toolbar.drawTrackpad")}
                            hint={t("toolbar.drawTrackpadHint")}
                            anchor={() => slideCanvasRef.current}
                            below={() => {
                              // The ribbon stays lit: only the deck under it dims.
                              for (
                                let host = slideCanvasRef.current?.parentElement;
                                host;
                                host = host.parentElement
                              ) {
                                const ribbon = host.querySelector(
                                  ".presentation-office-ribbon",
                                );
                                if (ribbon) return ribbon;
                              }
                              return null;
                            }}
                            page={() =>
                              slideCanvasRef.current?.getBoundingClientRect() ??
                              null
                            }
                            onStroke={placePresentationTrackpadStroke}
                            onClose={() => writeInkTrackpadEnabled(false)}
                          />
                        </OfficeHomeCommandRow>
                      </OfficeHomeRibbonGroup>
                    )}

                    {/* The ribbon brought back over the pen palette folds
                        into it again from here (lib/pen-first-draw). */}
                    {hideRibbon ? (
                      <OfficeHomeRibbonGroup
                        ribbon="draw"
                        label={t("toolbar.penPalette")}
                        size="content"
                        testId="presentation-ribbon-draw-pen-palette-group"
                      >
                        <OfficeHomeCommandRow>
                          <OfficeInkHideRibbonCommand
                            controlId="PPTCTL.RIBBON.DRAW.HIDE_RIBBON"
                            testId="presentation-ribbon-draw-hide-ribbon"
                            onHide={hideRibbon}
                          />
                        </OfficeHomeCommandRow>
                      </OfficeHomeRibbonGroup>
                    ) : null}
                  </>
  );
}
