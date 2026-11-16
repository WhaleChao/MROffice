"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { objectSizePixels } from "../../../../lib/spreadsheet/drawing-buttons";
import { OfficePictureStyleGallery } from "@mrnote/core/web/components/office-ribbon/OfficePictureStyleGallery.tsx";
import { type OfficePictureStyleId } from "@mrnote/core/web/lib/office-picture-styles/catalog.ts";
import { type RefObject } from "react";
import { useTranslations } from "@mrnote/core/host/next-intl.tsx";
import { OfficeHomeCommandStack, OfficeHomeSplitLayout } from "@mrnote/core/web/components/office-ribbon/OfficeHomeRibbon.tsx";
import { OfficeShapeSizeFields } from "@mrnote/core/web/components/office-ribbon/OfficeShapeSizeFields.tsx";
import { loadOfficeImageNaturalSize, officePictureFitCrop } from "@mrnote/core/web/lib/office-picture-fit.ts";
import { pictureObjectDisplaySrc } from "@mrnote/core/web/lib/page-picture-objects.ts";
import { OfficeArtisticEffectsMenu } from "@mrnote/core/web/components/office-ribbon/OfficeArtisticEffectsMenu.tsx";
import { normalizeOfficeArtisticEffect } from "@mrnote/core/web/lib/office-artistic-effects.ts";
import { OfficeSemanticIcon } from "@mrnote/core/web/components/office-ribbon/icons/OfficeSemanticIcon.tsx";
import spreadsheetHomeStyles from "./SpreadsheetHomeRibbon.module.css";
import { SpreadsheetPictureIcon } from "./SpreadsheetPictureRibbonIcons";
import { SpreadsheetRibbonMenuChevron } from "../../../../lib/spreadsheet/SpreadsheetRibbon";
import { DropdownMenu, DropdownMenuTrigger } from "@mrnote/core/web/components/ui/dropdown-menu.tsx";
import { type SpreadsheetObject, type SheetModel } from "./spreadsheetSheetModel";
import { OfficeHomeMenuTrigger, SpreadsheetRibbonAction, DropdownMenuContent, DropdownMenuItem } from "./SpreadsheetRibbonControls";
import { type PictureFormatWriterControlId } from "./spreadsheetMutationReceipts";
import { type SparseAxisSizeResolver } from "@mrnote/core/web/lib/spreadsheet/grid-sizing.ts";
import type { OfficeThemeGridColors } from "@mrnote/core/web/lib/office-theme-color-grid.ts";
import {
  SpreadsheetPictureColorMenu,
  SpreadsheetPictureCorrectionsMenu,
  SpreadsheetPictureTransparencyMenu,
} from "./SpreadsheetPictureAdjustMenus";
import { type Dispatch, type JSX, type ReactNode, type SetStateAction } from "react";

/** What spreadsheetPictureFormatRibbonGroups reads from SpreadsheetWindow's render scope. */
export interface SpreadsheetPictureFormatRibbonGroupsInputs {
   
  arrangeLane: ReactNode;
  colWidth: SparseAxisSizeResolver;
  executePictureFormatMutation: (controlId: PictureFormatWriterControlId, parameter: string | number, surface?: "visible" | "keyboard" | "navigation") => Promise<boolean | void>;
  executeVisibleCommand: (controlId: string) => Promise<boolean>;
  pictureArrangeControlIdByMove: { readonly front: "XLSCTL.CONTEXT.PICTURE.BRING_TO_FRONT.MRAI.APPLY"; readonly forward: "XLSCTL.CONTEXT.PICTURE.BRING_FORWARD.MRAI.APPLY"; readonly backward: "XLSCTL.CONTEXT.PICTURE.SEND_BACKWARD.MRAI.APPLY"; readonly back: "XLSCTL.CONTEXT.PICTURE.SEND_TO_BACK.MRAI.APPLY"; };
  pictureCropDisabledReason: string | undefined;
  pictureCropEditingId: string | null;
  pictureCropTriggerRef: RefObject<HTMLButtonElement | null>;
  pictureEffectsDisabledReason: string | undefined;
  pictureMetadataDisabledReason: string | undefined;
  pictureStyleDisabledReason: string | undefined;
  pictureTransformDisabledReason: string | undefined;
  /** The workbook's theme, which Color's Dark and Light tiles resolve in. */
  pictureRecolorTheme: OfficeThemeGridColors | null;
  /** "… Options..." under Corrections, Color and Transparency. */
  openPictureFormatPane: () => void;
  /** Picture Effects (the shared menu over the picture's `pictureEffects`). */
  renderPictureEffectsMenu: () => ReactNode;
  renderChartGroupMenu: (prefix?: string) => JSX.Element;
  renderChartRotateMenu: (prefix?: string) => JSX.Element;
  rowHeight: SparseAxisSizeResolver;
  selectedObjectIdRef: RefObject<string | null>;
  /** Excel: the layer commands grey only when the sheet holds one drawing. */
  selectedObjectArrangeDisabled: boolean;
  selectedPictureObject: SpreadsheetObject | null;
  setPictureStylePreview: Dispatch<SetStateAction<{ readonly sheetId: string; readonly objectId: string; readonly style: OfficePictureStyleId; } | null>>;
  sheet: SheetModel;
  t: ReturnType<typeof useTranslations<"console-notebooks">>;
  visibleCommandMenuProps: (controlId: string, accessibleLabel?: string, selectionReason?: string) => { "data-control-id": string; "aria-label": string | undefined; disabled: boolean; hidden: boolean; title: string | undefined; };
}

/** Moved verbatim from SpreadsheetWindow (called at the same position, so hook order is unchanged). */
export function spreadsheetPictureFormatRibbonGroups({
  arrangeLane,
  colWidth,
  executePictureFormatMutation,
  executeVisibleCommand,
  pictureArrangeControlIdByMove,
  pictureCropDisabledReason,
  pictureCropEditingId,
  pictureCropTriggerRef,
  pictureEffectsDisabledReason,
  pictureMetadataDisabledReason,
  pictureStyleDisabledReason,
  pictureTransformDisabledReason,
  pictureRecolorTheme,
  openPictureFormatPane,
  renderPictureEffectsMenu,
  renderChartGroupMenu,
  renderChartRotateMenu,
  rowHeight,
  selectedObjectIdRef,
  selectedObjectArrangeDisabled,
  selectedPictureObject,
  setPictureStylePreview,
  sheet,
  t,
  visibleCommandMenuProps,
}: SpreadsheetPictureFormatRibbonGroupsInputs) {
  const pictureColorStated =
    (selectedPictureObject?.source as { nativePictureColorStated?: unknown } | undefined)?.nativePictureColorStated === true;
  const pictureFormatRibbonGroups = selectedPictureObject
    ? [
        // Excel: Adjust | Picture Styles | Accessibility | Arrange | Size.
        {
          id: "picture-format-adjust",
          label: t("presentation.image.adjustGroup"),
          content: (
            <>
              {/* Excel's Picture Format > Adjust > Corrections and Color: the
                  shared galleries, drawn through the grid's own filter. A
                  picture read before the sheet read its blip's colour is
                  not offered them: its save could not tell them apart. */}
              <SpreadsheetPictureCorrectionsMenu
                object={selectedPictureObject}
                theme={pictureRecolorTheme}
                disabled={Boolean(pictureEffectsDisabledReason) || !pictureColorStated}
                disabledReason={pictureEffectsDisabledReason ?? (pictureColorStated ? undefined : t("spreadsheet.uiCopy.pictureColorNeedsReopen"))}
                onOptions={openPictureFormatPane}
                onCommit={(brightness, contrast) =>
                  void executePictureFormatMutation(
                    "XLSCTL.CONTEXT.PICTURE.CORRECTIONS.MRAI.APPLY",
                    `${brightness},${contrast}`,
                  )
                }
                trigger={
                  <OfficeHomeMenuTrigger
                    variant="large"
                    hasMenuIndicator
                    controlId="XLSCTL.CONTEXT.PICTURE.CORRECTIONS.MRAI.APPLY"
                    disabled={Boolean(pictureEffectsDisabledReason) || !pictureColorStated}
                    icon={<OfficeSemanticIcon semanticId="image.corrections" size={24} />}
                    label={t("presentation.pictureRibbon.correctionsMenu")}
                    testId="spreadsheet-ribbon-picture-corrections"
                  />
                }
              />
              <SpreadsheetPictureColorMenu
                object={selectedPictureObject}
                theme={pictureRecolorTheme}
                disabled={Boolean(pictureEffectsDisabledReason) || !pictureColorStated}
                disabledReason={pictureEffectsDisabledReason ?? (pictureColorStated ? undefined : t("spreadsheet.uiCopy.pictureColorNeedsReopen"))}
                onOptions={openPictureFormatPane}
                onCommit={(preset) =>
                  void executePictureFormatMutation("XLSCTL.CONTEXT.PICTURE.COLOR.MRAI.APPLY", preset ?? "none")
                }
                trigger={
                  <OfficeHomeMenuTrigger
                    variant="large"
                    hasMenuIndicator
                    controlId="XLSCTL.CONTEXT.PICTURE.COLOR.MRAI.APPLY"
                    disabled={Boolean(pictureEffectsDisabledReason) || !pictureColorStated}
                    icon={<OfficeSemanticIcon semanticId="image.color" size={24} />}
                    label={t("toolbar.color")}
                    testId="spreadsheet-ribbon-picture-color"
                  />
                }
              />
              {/* Excel's Picture Format > Adjust > Artistic Effects: a large
                  gallery button with its drop arrow, like its neighbours. */}
              <OfficeArtisticEffectsMenu
                value={normalizeOfficeArtisticEffect(selectedPictureObject.artisticEffect)}
                disabled={Boolean(pictureEffectsDisabledReason) || selectedPictureObject.artisticEffectRendered === true}
                disabledReason={
                  selectedPictureObject.artisticEffectRendered === true
                    ? t("spreadsheet.uiCopy.pictureArtisticEffectFromFile")
                    : pictureEffectsDisabledReason
                }
                onCommit={(id) =>
                  void executePictureFormatMutation(
                    "XLSCTL.CONTEXT.PICTURE.ARTISTIC_EFFECTS.MRAI.APPLY",
                    id,
                  )
                }
                testIdPrefix="spreadsheet-ribbon-picture-artistic-effects"
                trigger={
                  <OfficeHomeMenuTrigger
                    variant="large"
                    hasMenuIndicator
                    controlId="XLSCTL.CONTEXT.PICTURE.ARTISTIC_EFFECTS.MRAI.APPLY"
                    disabled={Boolean(pictureEffectsDisabledReason)}
                    disabledReason={pictureEffectsDisabledReason}
                    icon={<OfficeSemanticIcon semanticId="image.corrections" size={24} />}
                    label={t("pictureFormatPane.artisticEffects")}
                    testId="spreadsheet-ribbon-picture-artistic-effects"
                  />
                }
              />
              <SpreadsheetPictureTransparencyMenu
                object={selectedPictureObject}
                disabled={Boolean(pictureEffectsDisabledReason) || !pictureColorStated}
                disabledReason={pictureEffectsDisabledReason ?? (pictureColorStated ? undefined : t("spreadsheet.uiCopy.pictureColorNeedsReopen"))}
                onOptions={openPictureFormatPane}
                onCommit={(transparency) =>
                  void executePictureFormatMutation("XLSCTL.CONTEXT.PICTURE.TRANSPARENCY.MRAI.APPLY", transparency)
                }
                trigger={
                  <OfficeHomeMenuTrigger
                    variant="large"
                    hasMenuIndicator
                    controlId="XLSCTL.CONTEXT.PICTURE.TRANSPARENCY.MRAI.APPLY"
                    disabled={Boolean(pictureEffectsDisabledReason) || !pictureColorStated}
                    icon={<OfficeSemanticIcon semanticId="image.transparency" size={24} />}
                    label={t("toolbar.transparency")}
                    testId="spreadsheet-ribbon-picture-transparency"
                  />
                }
              />
            </>
          ),
        },
        {
          id: "picture-format-styles",
          label: t("presentation.image.stylesGroup"),
          content: (
            <span
              className="relative flex min-h-[64px] items-center px-1"
              data-control-id="XLSCTL.CONTEXT.PICTURE.STYLE_GALLERY.MRAI.APPLY"
            >
              <OfficePictureStyleGallery
                application="spreadsheet"
                value={selectedPictureObject.pictureStyle ?? null}
                disabled={Boolean(pictureStyleDisabledReason)}
                disabledReason={pictureStyleDisabledReason}
                labels={{
                  ariaLabel: t("presentation.image.stylesGroup"),
                  more: t("spreadsheet.uiCopy.moreStyles"),
                }}
                visibleItemCount={6}
                onPreviewChange={(id) =>
                  setPictureStylePreview(
                    id && !pictureStyleDisabledReason
                      ? {
                          sheetId: sheet.id,
                          objectId: selectedPictureObject.id,
                          style: id,
                        }
                      : null,
                  )
                }
                onCommit={(id) =>
                  void executePictureFormatMutation(
                    "XLSCTL.CONTEXT.PICTURE.STYLE_GALLERY.MRAI.APPLY",
                    id,
                  )
                }
                testIdPrefix="spreadsheet-picture-style-gallery"
              />
              {/* Excel's Picture Styles group: the gallery, then Picture Effects. */}
              <OfficeHomeCommandStack>{renderPictureEffectsMenu()}</OfficeHomeCommandStack>
            </span>
          ),
        },
        {
          id: "picture-format-accessibility",
          label: t("spreadsheet.review.accessibility"),
          content: (
            <SpreadsheetRibbonAction
              controlId="XLSCTL.CONTEXT.PICTURE.ALT_TEXT"
              disabled={Boolean(pictureMetadataDisabledReason)}
              disabledReason={pictureMetadataDisabledReason}
              homeVariant="large"
              icon={<SpreadsheetPictureIcon name="altText" />}
              label={t("spreadsheet.objectAlt")}
              onClick={() =>
                void executeVisibleCommand("XLSCTL.CONTEXT.PICTURE.ALT_TEXT")
              }
              testId="spreadsheet-ribbon-picture-alt-command"
            />
          ),
        },
        {
          id: "picture-format-arrange",
          label: t("spreadsheet.arrangeMenu"),
          content: (
            <>
              {selectedPictureObject?.groupId ? <OfficeHomeCommandStack>{renderChartGroupMenu("spreadsheet-picture-chart")}{renderChartRotateMenu("spreadsheet-picture-chart")}</OfficeHomeCommandStack> : null}
              <OfficeHomeSplitLayout
                className={spreadsheetHomeStyles.insertLargeSplit}
                testId="spreadsheet-ribbon-picture-bring-forward-split"
                primary={
                  <SpreadsheetRibbonAction
                    controlId={pictureArrangeControlIdByMove.forward}
                    disabled={
                      Boolean(pictureTransformDisabledReason) ||
                      selectedObjectArrangeDisabled
                    }
                    disabledReason={pictureTransformDisabledReason}
                    homeVariant="large"
                    icon={<SpreadsheetPictureIcon name="bringForward" />}
                    label={t("spreadsheet.bringForward")}
                    onClick={() =>
                      void executePictureFormatMutation(
                        pictureArrangeControlIdByMove.forward,
                        "forward",
                      )
                    }
                    testId="spreadsheet-ribbon-arrange-forward"
                  />
                }
                trigger={
                  <DropdownMenu modal={false}>
                    <DropdownMenuTrigger asChild>
                      <OfficeHomeMenuTrigger
                        disabled={Boolean(pictureTransformDisabledReason)}
                        disabledReason={pictureTransformDisabledReason}
                        icon={<SpreadsheetRibbonMenuChevron size={12} />}
                        label={
                          t("spreadsheet.uiCopy.bringForwardOptions")
                        }
                        testId="spreadsheet-ribbon-arrange-forward-menu"
                        variant="icon"
                      />
                    </DropdownMenuTrigger>
                    <DropdownMenuContent
                      align="start"
                      className="docs-cupertino-menu w-max max-w-[12rem]"
                    >
                      <DropdownMenuItem
                        {...visibleCommandMenuProps(
                          pictureArrangeControlIdByMove.forward,
                        )}
                        disabled={
                          Boolean(pictureTransformDisabledReason) ||
                          selectedObjectArrangeDisabled
                        }
                        data-testid="spreadsheet-ribbon-arrange-forward-menu-item"
                        onSelect={() =>
                          void executePictureFormatMutation(
                            pictureArrangeControlIdByMove.forward,
                            "forward",
                          )
                        }
                      >
                        <SpreadsheetPictureIcon name="bringForward" size={16} />
                        {t("spreadsheet.bringForward")}
                      </DropdownMenuItem>
                      <DropdownMenuItem
                        {...visibleCommandMenuProps(
                          pictureArrangeControlIdByMove.front,
                        )}
                        disabled={
                          Boolean(pictureTransformDisabledReason) ||
                          selectedObjectArrangeDisabled
                        }
                        data-testid="spreadsheet-ribbon-arrange-front"
                        onSelect={() =>
                          void executePictureFormatMutation(
                            pictureArrangeControlIdByMove.front,
                            "front",
                          )
                        }
                      >
                        <SpreadsheetPictureIcon name="bringToFront" size={16} />
                        {t("spreadsheet.bringToFront")}
                      </DropdownMenuItem>
                    </DropdownMenuContent>
                  </DropdownMenu>
                }
              />
              <OfficeHomeSplitLayout
                className={spreadsheetHomeStyles.insertLargeSplit}
                testId="spreadsheet-ribbon-picture-send-backward-split"
                primary={
                  <SpreadsheetRibbonAction
                    controlId={pictureArrangeControlIdByMove.backward}
                    disabled={
                      Boolean(pictureTransformDisabledReason) ||
                      selectedObjectArrangeDisabled
                    }
                    disabledReason={pictureTransformDisabledReason}
                    homeVariant="large"
                    icon={<SpreadsheetPictureIcon name="sendBackward" />}
                    label={t("spreadsheet.sendBackward")}
                    onClick={() =>
                      void executePictureFormatMutation(
                        pictureArrangeControlIdByMove.backward,
                        "backward",
                      )
                    }
                    testId="spreadsheet-ribbon-arrange-backward"
                  />
                }
                trigger={
                  <DropdownMenu modal={false}>
                    <DropdownMenuTrigger asChild>
                      <OfficeHomeMenuTrigger
                        disabled={Boolean(pictureTransformDisabledReason)}
                        disabledReason={pictureTransformDisabledReason}
                        icon={<SpreadsheetRibbonMenuChevron size={12} />}
                        label={
                          t("spreadsheet.uiCopy.sendBackwardOptions")
                        }
                        testId="spreadsheet-ribbon-arrange-backward-menu"
                        variant="icon"
                      />
                    </DropdownMenuTrigger>
                    <DropdownMenuContent
                      align="start"
                      className="docs-cupertino-menu w-max max-w-[12rem]"
                    >
                      <DropdownMenuItem
                        {...visibleCommandMenuProps(
                          pictureArrangeControlIdByMove.backward,
                        )}
                        disabled={
                          Boolean(pictureTransformDisabledReason) ||
                          selectedObjectArrangeDisabled
                        }
                        data-testid="spreadsheet-ribbon-arrange-backward-menu-item"
                        onSelect={() =>
                          void executePictureFormatMutation(
                            pictureArrangeControlIdByMove.backward,
                            "backward",
                          )
                        }
                      >
                        <SpreadsheetPictureIcon name="sendBackward" size={16} />
                        {t("spreadsheet.sendBackward")}
                      </DropdownMenuItem>
                      <DropdownMenuItem
                        {...visibleCommandMenuProps(
                          pictureArrangeControlIdByMove.back,
                        )}
                        disabled={
                          Boolean(pictureTransformDisabledReason) ||
                          selectedObjectArrangeDisabled
                        }
                        data-testid="spreadsheet-ribbon-arrange-back"
                        onSelect={() =>
                          void executePictureFormatMutation(
                            pictureArrangeControlIdByMove.back,
                            "back",
                          )
                        }
                      >
                        <SpreadsheetPictureIcon name="sendToBack" size={16} />
                        {t("spreadsheet.sendToBack")}
                      </DropdownMenuItem>
                    </DropdownMenuContent>
                  </DropdownMenu>
                }
              />
              {arrangeLane}
            </>
          ),
        },
        {
          id: "picture-format-size",
          label: t("presentation.smartartContext.sizeGroup"),
          content: (
            <div className="flex min-h-[72px] items-center gap-2">
              <SpreadsheetRibbonAction
                active={pictureCropEditingId === selectedPictureObject.id}
                buttonRef={pictureCropTriggerRef}
                controlId="XLSCTL.CONTEXT.PICTURE.CROP"
                disabled={Boolean(pictureCropDisabledReason)}
                disabledReason={pictureCropDisabledReason}
                homeVariant="large"
                icon={<OfficeSemanticIcon semanticId="image.crop" size={24} />}
                label={t("presentation.cameo.crop")}
                onClick={() =>
                  void executeVisibleCommand("XLSCTL.CONTEXT.PICTURE.CROP")
                }
                testId="spreadsheet-ribbon-picture-crop"
              />
              {/* Excel's Crop > Fill / Fit: the kept source rectangle from the
                  shared math (`office-picture-fit`), written by the same crop
                  writer a crop drag uses. */}
              <div className="flex flex-col gap-1">
                {(["fill", "fit"] as const).map((mode) => (
                  <SpreadsheetRibbonAction
                    key={mode}
                    controlId="XLSCTL.CONTEXT.PICTURE.CROP.MRAI.APPLY"
                    disabled={Boolean(pictureCropDisabledReason)}
                    disabledReason={pictureCropDisabledReason}
                    homeVariant="compact"
                    icon={<OfficeSemanticIcon semanticId="image.crop" size={16} />}
                    label={t(`presentation.crop.${mode}` as "presentation.crop.fill")}
                    onClick={() => {
                      const target = selectedPictureObject;
                      const frame = objectSizePixels(target, colWidth, rowHeight);
                      // A stored picture is measured from its page's picture route.
                      void loadOfficeImageNaturalSize(pictureObjectDisplaySrc(target.src)).then((natural) => {
                        const rect = natural ? officePictureFitCrop(mode, natural, frame) : null;
                        if (!rect || selectedObjectIdRef.current !== target.id) return;
                        void executePictureFormatMutation(
                          "XLSCTL.CONTEXT.PICTURE.CROP.MRAI.APPLY",
                          JSON.stringify(rect),
                        );
                      });
                    }}
                    testId={`spreadsheet-ribbon-picture-crop-${mode}`}
                  />
                ))}
              </div>
              {/* Excel's Size group: Height and Width, the SAME shared fields
                  the Shape Format tab and the other two editors use. */}
              {(() => {
                const size = objectSizePixels(selectedPictureObject, colWidth, rowHeight);
                const cm = (px: number) => ((px * 2.54) / 96).toFixed(2);
                return (
                  <OfficeShapeSizeFields
                    values={{ width: cm(size.width), height: cm(size.height) }}
                    labels={{
                      width: t("spreadsheet.chartFormatCopy.size.width"),
                      height: t("spreadsheet.chartFormatCopy.size.height"),
                    }}
                    disabled={Boolean(pictureTransformDisabledReason)}
                    disabledReason={pictureTransformDisabledReason}
                    heightFirst
                    onCommit={(axis, value) =>
                      void executePictureFormatMutation(
                        axis === "width"
                          ? "XLSCTL.CONTEXT.PICTURE.WIDTH.MRAI.APPLY"
                          : "XLSCTL.CONTEXT.PICTURE.HEIGHT.MRAI.APPLY",
                        String(Math.max(1, Math.round((Number(value) * 96) / 2.54))),
                      )
                    }
                    testIdPrefix="spreadsheet-ribbon-picture-size"
                  />
                );
              })()}
            </div>
          ),
        },
      ]
    : [];
  return {
    pictureFormatRibbonGroups,
  };
}
