/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The Arrange group of the Picture Format and Graphics Format tabs.
 *
 * Moved verbatim from the body of
 * `components/notebook/contents/PresentationWindow.tsx`: the window passes the
 * bindings it used to close over as `ctl` and calls this where the closure was.
 * No hooks run here.
 */

import {
  OfficeHomeRibbonGroup,
  OfficeHomeCommandRow,
  OfficeHomeSplitLayout,
  OfficeHomeCommand,
  OfficeHomeMenuTrigger,
} from "@mrnote/core/web/components/office-ribbon/OfficeHomeRibbon.tsx";
import presentationChrome from "@mrnote/core/web/lib/presentation/presentation-window-shell.module.css";
import {
  DropdownMenu,
  DropdownMenuTrigger,
  DropdownMenuItem,
  DropdownMenuLabel,
  DropdownMenuCheckboxItem,
  DropdownMenuSeparator,
} from "@mrnote/core/web/components/ui/dropdown-menu.tsx";
import { OfficeSemanticIcon } from "@mrnote/core/web/components/office-ribbon/icons/OfficeSemanticIcon.tsx";
import { PRESENTATION_PICTURE_BRING_FORWARD_CONTROL_ID, PRESENTATION_PICTURE_SEND_BACKWARD_CONTROL_ID } from "@mrnote/core/web/lib/presentation-feature-capabilities.ts";
import { RibbonDropdownMenuContent } from "@mrnote/core/web/components/office-ribbon/CupertinoRibbonMenu.tsx";
import { OfficePictureMenuIcon } from "@mrnote/core/web/components/office-ribbon/icons/OfficePictureMenuIcon.tsx";
import type { AlignEdge, DistributeAxis } from "@mrnote/core/web/lib/presentation-align.ts";
import type { useTranslations } from "@mrnote/core/host/next-intl.tsx";
import type { RefObject } from "react";
import type { PresentationImageBlock } from "@mrnote/core/web/components/console/editor/native-document/documentCommands.ts";
import type { QuarterTurnDirection } from "@mrnote/core/web/lib/presentation-rotate-quarter.ts";
import type { ResolvedPresentationFeatureCommand } from "@mrnote/core/web/components/notebook/contents/ribbon/commandRegistry.ts";

/** The window bindings this reads, passed through unchanged. */
export interface BuildPictureArrangeGroupCtl {
  t: ReturnType<typeof useTranslations<"console-notebooks">>;
  selectionPaneReturnFocusRef: RefObject<HTMLElement | null>;
  selectionPaneImageTriggerRef: RefObject<HTMLButtonElement | null>;
  alignToSlide: boolean;
  writeAlignToSlide: (next: boolean) => void;
  contextualImageBlocks: readonly PresentationImageBlock[];
  selectionPaneOpen: boolean;
  alignSelectedBlocks: (edge: AlignEdge) => void;
  distributeSelectedBlocks: (axis: DistributeAxis) => void;
  flipSelectedBlocks: (axis: "flipH" | "flipV") => void;
  quarterTurnSelectedBlocks: (direction: QuarterTurnDirection) => void;
  groupSelectedBlocks: () => void;
  ungroupSelectedBlocks: () => void;
  drawingImageSelectionIsEligible: boolean;
  canGroupSelectedImages: boolean;
  

  pictureGroupRefusalReason?: string;
  canUngroupSelectedImages: boolean;
  contextualImageCanMoveForward: boolean;
  contextualImageCanMoveBackward: boolean;
  contextualImageCanMoveToFront: boolean;
  contextualImageCanMoveToBack: boolean;
  bringBlockToFront: (explicitId?: string, explicitOnly?: boolean) => void;
  sendBlockToBack: (explicitId?: string, explicitOnly?: boolean) => void;
  invokeContextualImageArrangeStep: (direction: "forward" | "backward") => void;
  featureDecision: (featureId: string) => ResolvedPresentationFeatureCommand;
  executePresentationFeature: (featureId: string) => undefined;
}

export function buildPictureArrangeGroup(ctl: BuildPictureArrangeGroupCtl) {
  const {
    t,
    selectionPaneReturnFocusRef,
    selectionPaneImageTriggerRef,
    alignToSlide,
    writeAlignToSlide,
    contextualImageBlocks,
    selectionPaneOpen,
    alignSelectedBlocks,
    distributeSelectedBlocks,
    flipSelectedBlocks,
    quarterTurnSelectedBlocks,
    groupSelectedBlocks,
    ungroupSelectedBlocks,
    drawingImageSelectionIsEligible,
    canGroupSelectedImages,
    pictureGroupRefusalReason,
    canUngroupSelectedImages,
    contextualImageCanMoveForward,
    contextualImageCanMoveBackward,
    contextualImageCanMoveToFront,
    contextualImageCanMoveToBack,
    bringBlockToFront,
    sendBlockToBack,
    invokeContextualImageArrangeStep,
    featureDecision,
    executePresentationFeature,
  } = ctl;
  const renderPictureArrangeGroup = (kind: "image" | "graphic") => (
      <OfficeHomeRibbonGroup
        className={presentationChrome.pictureArrangeGroup}
        label={t("presentation.home.arrange")}
        size="standard"
        testId={`presentation-ribbon-${kind}-arrange-group`}
      >
        <OfficeHomeCommandRow className={presentationChrome.pictureArrangeRow}>
          <DropdownMenu>
            <OfficeHomeSplitLayout
              testId={`presentation-ribbon-${kind}-bring-forward-split`}
              primary={
                <OfficeHomeCommand
                  variant="compact"
                  icon={<OfficeSemanticIcon semanticId="object.bringForward" size={20} />}
                  label={t(
                    "presentation.contextMenu.bringForward",
                  )}
                  disabled={!contextualImageCanMoveForward}
                  controlId={
                    PRESENTATION_PICTURE_BRING_FORWARD_CONTROL_ID
                  }
                  onClick={() =>
                    invokeContextualImageArrangeStep("forward")
                  }
                  testId={`presentation-ribbon-${kind}-bring-forward`}
                />
              }
              trigger={
                <DropdownMenuTrigger asChild>
                  <OfficeHomeMenuTrigger
                    variant="icon"
                    icon={
                      <OfficeSemanticIcon
                        semanticId="menu.chevron.down"
                        size={12}
                      />
                    }
                    label={t(
                      "presentation.contextMenu.bringForward",
                    )}
                    disabled={!contextualImageCanMoveForward}
                    data-control-id="PPTCTL.CONTEXT.PICTURE.ARRANGE.BRING_FORWARD.MENU"
                    testId={`presentation-ribbon-${kind}-bring-forward-menu`}
                  />
                </DropdownMenuTrigger>
              }
            />
            <RibbonDropdownMenuContent
              align="start"
              data-testid={`presentation-ribbon-${kind}-bring-forward-menu-content`}
            >
              <DropdownMenuItem
                disabled={!contextualImageCanMoveForward}
                onSelect={() =>
                  invokeContextualImageArrangeStep("forward")
                }
                data-testid={`presentation-ribbon-${kind}-bring-forward-menu-step`}
              >
                <OfficeSemanticIcon semanticId="object.bringForward" size={14} />
                {t("presentation.contextMenu.bringForward")}
              </DropdownMenuItem>
              <DropdownMenuItem
                disabled={!contextualImageCanMoveToFront}
                onSelect={() => bringBlockToFront()}
                data-testid={`presentation-ribbon-${kind}-bring-front`}
              >
                <OfficeSemanticIcon semanticId="object.bringToFront" size={14} />
                {t("presentation.contextMenu.bringToFront")}
              </DropdownMenuItem>
            </RibbonDropdownMenuContent>
          </DropdownMenu>
          <DropdownMenu>
            <DropdownMenuTrigger asChild>
              <OfficeHomeMenuTrigger
                variant="icon"
                active={alignToSlide}
                displayLabel={t(
                  "presentation.ribbon.align",
                )}
                hasMenuIndicator
                icon={<OfficeSemanticIcon semanticId="object.align" size={20} />}
                label={t("presentation.ribbon.align")}
                disabled={!drawingImageSelectionIsEligible}
                data-control-id="PPTCTL.CONTEXT.PICTURE.ARRANGE.ALIGN"
                data-count={contextualImageBlocks.length}
                testId={`presentation-ribbon-${kind}-align`}
              />
            </DropdownMenuTrigger>
            <RibbonDropdownMenuContent
              align="start"
              data-testid={`presentation-ribbon-${kind}-align-menu`}
            >
              <DropdownMenuLabel>
                {t("presentation.ribbon.align")}
              </DropdownMenuLabel>
              <DropdownMenuCheckboxItem
                checked={alignToSlide}
                onCheckedChange={(checked) =>
                  writeAlignToSlide(checked === true)
                }
                data-testid={`presentation-ribbon-${kind}-align-slide`}
              >
                <OfficePictureMenuIcon icon={{ kind: "align-to-slide" }} />
                {t("presentation.align.toSlideLabel")}
              </DropdownMenuCheckboxItem>
              <DropdownMenuSeparator />
              {(
                [
                  ["left", "align-left"],
                  ["hcenter", "align-center"],
                  ["right", "align-right"],
                  ["top", "align-top"],
                  ["vcenter", "align-middle"],
                  ["bottom", "align-bottom"],
                ] as const satisfies ReadonlyArray<
                  readonly [AlignEdge, string]
                >
              ).map(([edge, iconKind]) => (
                <DropdownMenuItem
                  key={edge}
                  disabled={
                    !drawingImageSelectionIsEligible ||
                    contextualImageBlocks.length <
                      (alignToSlide ? 1 : 2)
                  }
                  onSelect={() => alignSelectedBlocks(edge)}
                  data-testid={`presentation-ribbon-${kind}-align-${edge}`}
                >
                  <OfficePictureMenuIcon icon={{ kind: iconKind }} />
                  {t(
                    `presentation.align.${edge}` as "presentation.align.left",
                  )}
                </DropdownMenuItem>
              ))}
              <DropdownMenuSeparator />
              {(
                [
                  ["horizontal", "distribute-horizontal"],
                  ["vertical", "distribute-vertical"],
                ] as const satisfies ReadonlyArray<
                  readonly [DistributeAxis, string]
                >
              ).map(([axis, iconKind]) => (
                <DropdownMenuItem
                  key={axis}
                  disabled={
                    !drawingImageSelectionIsEligible ||
                    contextualImageBlocks.length < (alignToSlide ? 1 : 3)
                  }
                  onSelect={() =>
                    distributeSelectedBlocks(axis)
                  }
                  data-testid={`presentation-ribbon-${kind}-distribute-${axis}`}
                >
                  <OfficePictureMenuIcon icon={{ kind: iconKind }} />
                  {t(
                    `presentation.distribute.${axis}` as "presentation.distribute.horizontal",
                  )}
                </DropdownMenuItem>
              ))}
            </RibbonDropdownMenuContent>
          </DropdownMenu>
          {/* No Auto Fix here: desktop PowerPoint's Picture Format
              has none (PPT-A130; it is PowerPoint for the web's
              Home > Arrange command). */}
        </OfficeHomeCommandRow>
        <OfficeHomeCommandRow className={presentationChrome.pictureArrangeRow}>
          <DropdownMenu>
            <OfficeHomeSplitLayout
              testId={`presentation-ribbon-${kind}-send-backward-split`}
              primary={
                <OfficeHomeCommand
                  variant="compact"
                  icon={<OfficeSemanticIcon semanticId="object.sendBackward" size={20} />}
                  label={t(
                    "presentation.contextMenu.sendBackward",
                  )}
                  disabled={!contextualImageCanMoveBackward}
                  controlId={
                    PRESENTATION_PICTURE_SEND_BACKWARD_CONTROL_ID
                  }
                  onClick={() =>
                    invokeContextualImageArrangeStep("backward")
                  }
                  testId={`presentation-ribbon-${kind}-send-backward`}
                />
              }
              trigger={
                <DropdownMenuTrigger asChild>
                  <OfficeHomeMenuTrigger
                    variant="icon"
                    icon={
                      <OfficeSemanticIcon
                        semanticId="menu.chevron.down"
                        size={12}
                      />
                    }
                    label={t(
                      "presentation.contextMenu.sendBackward",
                    )}
                    disabled={!contextualImageCanMoveBackward}
                    data-control-id="PPTCTL.CONTEXT.PICTURE.ARRANGE.SEND_BACKWARD.MENU"
                    testId={`presentation-ribbon-${kind}-send-backward-menu`}
                  />
                </DropdownMenuTrigger>
              }
            />
            <RibbonDropdownMenuContent
              align="start"
              data-testid={`presentation-ribbon-${kind}-send-backward-menu-content`}
            >
              <DropdownMenuItem
                disabled={!contextualImageCanMoveBackward}
                onSelect={() =>
                  invokeContextualImageArrangeStep("backward")
                }
                data-testid={`presentation-ribbon-${kind}-send-backward-menu-step`}
              >
                <OfficeSemanticIcon semanticId="object.sendBackward" size={14} />
                {t("presentation.contextMenu.sendBackward")}
              </DropdownMenuItem>
              <DropdownMenuItem
                disabled={!contextualImageCanMoveToBack}
                onSelect={() => sendBlockToBack()}
                data-testid={`presentation-ribbon-${kind}-send-back`}
              >
                <OfficeSemanticIcon semanticId="object.sendToBack" size={14} />
                {t("presentation.contextMenu.sendToBack")}
              </DropdownMenuItem>
            </RibbonDropdownMenuContent>
          </DropdownMenu>
          <DropdownMenu>
            <DropdownMenuTrigger asChild>
              <OfficeHomeMenuTrigger
                variant="icon"
                displayLabel={t("presentation.group.group")}
                hasMenuIndicator
                icon={<OfficeSemanticIcon semanticId="object.group" size={20} />}
                label={t("presentation.group.group")}
                disabled={
                  !canGroupSelectedImages &&
                  !canUngroupSelectedImages
                }
                // A closed Group names why on the button itself: its menu
                
                disabledReason={pictureGroupRefusalReason}
                data-control-id="PPTCTL.CONTEXT.PICTURE.ARRANGE.GROUP"
                testId={`presentation-ribbon-${kind}-group-menu`}
              />
            </DropdownMenuTrigger>
            <RibbonDropdownMenuContent
              align="start"
              data-testid={`presentation-ribbon-${kind}-group-menu-content`}
            >
              <DropdownMenuLabel>
                {t("presentation.group.group")}
              </DropdownMenuLabel>
              <DropdownMenuItem
                disabled={!canGroupSelectedImages}
                title={!canGroupSelectedImages ? pictureGroupRefusalReason : undefined}
                data-disabled-reason={!canGroupSelectedImages ? pictureGroupRefusalReason : undefined}
                onSelect={groupSelectedBlocks}
                data-testid={`presentation-ribbon-${kind}-group`}
              >
                <OfficeSemanticIcon semanticId="object.group" size={14} />
                {t("presentation.group.group")}
              </DropdownMenuItem>
              <DropdownMenuItem
                disabled={!canUngroupSelectedImages}
                onSelect={ungroupSelectedBlocks}
                data-testid={`presentation-ribbon-${kind}-ungroup`}
              >
                <OfficeSemanticIcon semanticId="object.ungroup" size={14} />
                {t("presentation.group.ungroup")}
              </DropdownMenuItem>
            </RibbonDropdownMenuContent>
          </DropdownMenu>
        </OfficeHomeCommandRow>
        <OfficeHomeCommandRow className={presentationChrome.pictureArrangeRow}>
          <OfficeHomeCommand
            active={selectionPaneOpen}
            buttonRef={selectionPaneImageTriggerRef}
            controlId="PPTCTL.CONTEXT.PICTURE.ARRANGE.SELECTION_PANE"
            disabled={!featureDecision("editing.selectionPane").enabled}
            icon={<OfficeSemanticIcon semanticId="object.selectionPane" size={20} />}
            label={t("toolbar.selectionPane")}
            onClick={() => {
              selectionPaneReturnFocusRef.current =
                selectionPaneImageTriggerRef.current;
              executePresentationFeature(
                "editing.selectionPane",
              );
            }}
            testId={`presentation-ribbon-${kind}-selection-pane`}
          />
          <DropdownMenu>
            <DropdownMenuTrigger asChild>
              <OfficeHomeMenuTrigger
                variant="icon"
                displayLabel={t(
                  "presentation.geometry.rotation",
                )}
                hasMenuIndicator
                icon={<OfficeSemanticIcon semanticId="object.rotate" size={20} />}
                label={t("presentation.geometry.rotation")}
                disabled={!drawingImageSelectionIsEligible}
                data-control-id="PPTCTL.CONTEXT.PICTURE.ARRANGE.ROTATE"
                testId={`presentation-ribbon-${kind}-rotate`}
              />
            </DropdownMenuTrigger>
            <RibbonDropdownMenuContent
              align="start"
              data-testid={`presentation-ribbon-${kind}-rotate-menu`}
            >
              <DropdownMenuLabel>
                {t("presentation.geometry.rotation")}
              </DropdownMenuLabel>
              <DropdownMenuItem
                disabled={!drawingImageSelectionIsEligible}
                onSelect={() =>
                  quarterTurnSelectedBlocks("right")
                }
                data-testid={`presentation-ribbon-${kind}-rotate-right`}
              >
                <OfficePictureMenuIcon icon={{ kind: "rotate-right-90" }} />
                {t("presentation.rotate.right")}
              </DropdownMenuItem>
              <DropdownMenuItem
                disabled={!drawingImageSelectionIsEligible}
                onSelect={() =>
                  quarterTurnSelectedBlocks("left")
                }
                data-testid={`presentation-ribbon-${kind}-rotate-left`}
              >
                <OfficePictureMenuIcon icon={{ kind: "rotate-left-90" }} />
                {t("presentation.rotate.left")}
              </DropdownMenuItem>
              <DropdownMenuSeparator />
              <DropdownMenuItem
                disabled={!drawingImageSelectionIsEligible}
                onSelect={() => flipSelectedBlocks("flipV")}
                data-testid={`presentation-ribbon-${kind}-flip-vertical`}
              >
                <OfficeSemanticIcon semanticId="object.flipVertical" size={16} />
                {t("presentation.flip.vertical")}
              </DropdownMenuItem>
              <DropdownMenuItem
                disabled={!drawingImageSelectionIsEligible}
                onSelect={() => flipSelectedBlocks("flipH")}
                data-testid={`presentation-ribbon-${kind}-flip-horizontal`}
              >
                <OfficeSemanticIcon semanticId="object.flipHorizontal" size={16} />
                {t("presentation.flip.horizontal")}
              </DropdownMenuItem>
            </RibbonDropdownMenuContent>
          </DropdownMenu>
        </OfficeHomeCommandRow>
      </OfficeHomeRibbonGroup>
  );
  return renderPictureArrangeGroup;
}
