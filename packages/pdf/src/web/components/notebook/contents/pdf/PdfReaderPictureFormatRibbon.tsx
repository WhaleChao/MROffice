"use client";

















/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { useState, type ReactNode } from "react";
import { useTranslations } from "@mrnote/core/host/next-intl.tsx";

import { RibbonDropdownMenuContent as DropdownMenuContent, RibbonDropdownMenuSubContent as DropdownMenuSubContent } from "@mrnote/core/web/components/office-ribbon/CupertinoRibbonMenu.tsx";
import {
  OfficeHomeCheckbox,
  OfficeHomeCommand,
  OfficeHomeCommandRow,
  OfficeHomeCommandStack,
  OfficeHomeMenuTrigger,
  OfficeHomeSplitLayout,
} from "@mrnote/core/web/components/office-ribbon/OfficeHomeRibbon.tsx";
import { OfficeArtisticEffectsMenu } from "@mrnote/core/web/components/office-ribbon/OfficeArtisticEffectsMenu.tsx";
import { OfficeColorPalettePopover } from "@mrnote/core/web/components/office-ribbon/OfficeColorPalettePopover.tsx";
import {
  OfficePictureColorMenu,
  OfficePictureCorrectionsMenu,
  OfficePictureTransparencyMenu,
} from "@mrnote/core/web/components/office-ribbon/OfficePictureAdjustMenus.tsx";
import { OfficePictureCropAspectRatioMenu } from "@mrnote/core/web/components/office-ribbon/OfficePictureCropAspectRatioMenu.tsx";
import { OfficePictureStyleGallery } from "@mrnote/core/web/components/office-ribbon/OfficePictureStyleGallery.tsx";
import { OfficeRibbonNumberField } from "@mrnote/core/web/components/office-ribbon/OfficeRibbonNumberField.tsx";
import { OfficeShapeCatalogPanel, officeShapeCatalogStyles } from "@mrnote/core/web/components/office-ribbon/OfficeShapeCatalogPanel.tsx";
import { OfficePictureMenuIcon } from "@mrnote/core/web/components/office-ribbon/icons/OfficePictureMenuIcon.tsx";
import { OfficeSemanticIcon } from "@mrnote/core/web/components/office-ribbon/icons/OfficeSemanticIcon.tsx";
import { PositionGlyph, POSITION_KEYS } from "@mrnote/core/web/components/console/editor/native-document/DocumentChartPlacementControls.tsx";
import { DocumentPictureDrawIcon } from "@mrnote/docs/web/components/console/editor/native-document/DocumentPictureDrawDialog.tsx";
import {
  DropdownMenu,
  DropdownMenuCheckboxItem,
  DropdownMenuItem,
  DropdownMenuLabel,
  DropdownMenuRadioGroup,
  DropdownMenuRadioItem,
  DropdownMenuSeparator,
  DropdownMenuSub,
  DropdownMenuSubTrigger,
  DropdownMenuTrigger,
} from "@mrnote/core/web/components/ui/dropdown-menu.tsx";
import {
  documentLengthUnit,
  documentPxPerLengthUnit,
  parseDocumentLengthPx,
} from "@mrnote/core/web/lib/document-length-fields.ts";
import type { OfficeArtisticEffectId } from "@mrnote/core/web/lib/office-artistic-effects.ts";
import {
  OFFICE_RECOLOR_GALLERY_ROWS,
  officeAdjustFilterSteps,
  officeRecolorLabelKey,
  type OfficeRecolorPresetId,
} from "@mrnote/core/web/lib/office-recolor-presets.ts";
import { officeShapeDefinition, officeShapePresetFromDrawingMl } from "@mrnote/core/web/lib/office-shape-catalog.ts";
import type { OfficePictureStyleId } from "@mrnote/core/web/lib/office-picture-styles/catalog.ts";
import {
  PICTURE_CROP_ASPECT_RATIOS,
  type PictureCropAspectRatioId,
} from "@mrnote/core/web/lib/presentation-crop-group.ts";
import type { PdfPictureBorder, PdfPictureBorderDash, PdfPictureFormat } from "../../../../lib/pdf/picture-format";

import styles from "./PdfReaderPictureFormatRibbon.module.css";
import { PdfSurfaceRibbonCommand, PdfSurfaceRibbonGroup } from "./PdfSurfaceRibbon";
import {
  PDF_PICTURE_POSITION_PRESETS,
  type PdfPictureAlign,
  type PdfPicturePositionPreset,
} from "./viewer/readerPictureFormatGeometry";

export interface PdfReaderPictureRibbonModel {
  /** A write of the picture is on its way. */
  readonly busy: boolean;
  readonly writable: boolean;
  readonly disabledReason?: string;
  /** The picture's frame on the page, in points. */
  readonly widthPoints: number;
  readonly heightPoints: number;
  /** Picture Transparency, percent (0..100). */
  readonly transparency: number;
  readonly format: PdfPictureFormat;
  readonly canReset: boolean;
  /** Why the picture cannot be cropped now (a skewed paint), or null. */
  readonly cropDisabledReason: string | null;
  /** Part of the picture is cropped away: Fill, Fit and Reset Crop apply. */
  readonly cropped: boolean;
  /** The Format Picture / Alt Text panes are open. */
  readonly formatPaneOpen: boolean;
  readonly altPaneOpen: boolean;
  /** The picture has an alternate description. */
  readonly hasAlt: boolean;
}

export interface PdfReaderPictureRibbonCommands {
  readonly onStyle: (id: OfficePictureStyleId) => void;
  /** A style hovered (null: none now): shown as its press would draw it. */
  readonly onStylePreview?: (id: OfficePictureStyleId | null) => void;
  readonly onBorder: (border: PdfPictureBorder | "none") => void;
  readonly onPosition: (preset: PdfPicturePositionPreset) => void;
  readonly onArrange: (direction: "front" | "back" | "forward" | "backward") => void;
  readonly onAlign: (align: PdfPictureAlign) => void;
  readonly onRotate: (direction: "left" | "right") => void;
  readonly onFlip: (axis: "horizontal" | "vertical") => void;
  readonly onRotateReset: () => void;
  readonly onCrop: () => void;
  readonly onCropAspectRatio: (width: number, height: number) => void;
  readonly onCropToShape: (preset: string) => void;
  /** Crop > Fill / Fit / Reset Crop. */
  readonly onCropMode: (mode: "fill" | "fit" | "reset") => void;
  /** Word's Format Picture pane and Alt Text pane. */
  readonly onFormatPane: () => void;
  readonly onAltText: () => void;
  readonly onResize: (widthPoints: number, heightPoints: number) => void;
  readonly onArtisticEffect: (id: OfficeArtisticEffectId) => void;
  readonly onCorrections: (brightness: number, contrast: number) => void;
  readonly onSharpness: (amount: number) => void;
  readonly onSharpen: () => void;
  readonly onSaturation: (percent: number) => void;
  readonly onTone: (kelvin: number) => void;
  readonly onRecolor: (preset: OfficeRecolorPresetId | null) => void;
  readonly onTransparency: (percent: number) => void;
  readonly onCompress: () => void;
  readonly onChangePicture: () => void;
  readonly onReset: (withSize: boolean) => void;
  readonly onDraw: () => void;
  /** WPS's 添加图片 (PdfEditAddImage), first in its picture tab. */
  readonly onInsertPicture?: () => void;
  /** WPS's 提取图片 (pdfImageExtract_PicEdit): the picture saved as a file. */
  readonly onExtractPicture: () => void;
  /** WPS's 图片转文字 (PictureEdit_EditTextInPicture): the picture's words read. */
  readonly onExtractText: () => void;
  /** WPS's 图片翻译 (PdfEditImgTabTranslate): the picture's words read and translated. */
  readonly onTranslatePicture: () => void;
  /** WPS's 删除 (PdfEditDeleteImageMenu). */
  readonly onDelete: () => void;
  readonly onExit: () => void;
}

const BORDER_WEIGHTS = [0.5, 1, 1.5, 2.25, 3, 4.5, 6] as const;
const BORDER_DASHES: readonly PdfPictureBorderDash[] = ["solid", "dashed", "dotted"];
const DEFAULT_BORDER_COLOR = "#000000";
const PT_PER_PX = 72 / 96;

/** A recolour tile: the sample through the preset's own filter steps. */
function RecolorSample({ id }: { id: OfficeRecolorPresetId | "none" }) {
  const steps = id === "none" ? [] : officeAdjustFilterSteps({ recolor: id });
  const filterId = `pdf-recolor-${id}`;
  return (
    <svg width="44" height="30" viewBox="0 0 44 30" aria-hidden="true">
      {steps.length ? (
        <filter id={filterId} colorInterpolationFilters="sRGB">
          {steps.map((step, index) =>
            step.type === "matrix" ? (
              <feColorMatrix key={index} type="matrix" values={step.values.join(" ")} />
            ) : (
              <feComponentTransfer key={index}>
                {(["R", "G", "B"] as const).map((channel) => {
                  const Func = channel === "R" ? "feFuncR" : channel === "G" ? "feFuncG" : "feFuncB";
                  return step.type === "discrete" ? (
                    <Func key={channel} type="discrete" tableValues={step.table.join(" ")} />
                  ) : (
                    <Func key={channel} type="linear" slope={step.slope} intercept={step.intercept} />
                  );
                })}
              </feComponentTransfer>
            ),
          )}
        </filter>
      ) : null}
      <g filter={steps.length ? `url(#${filterId})` : undefined}>
        <rect width="44" height="30" fill="#8bc9ed" />
        <path d="M0 19 C8 11 14 12 19 16 C25 8 31 10 36 15 C40 13 42 14 44 16 V30 H0Z" fill="#fff" />
        <path d="M0 27 L12 21 Q17 18 22 22 Q28 26 33 22 Q38 17 44 19 V30 H0Z" fill="#9fdcb0" />
      </g>
    </svg>
  );
}

export function PdfReaderPictureFormatRibbon({
  model,
  commands,
  pointerTools,
  locale,
}: {
  model: PdfReaderPictureRibbonModel;
  commands: PdfReaderPictureRibbonCommands;
  pointerTools?: ReactNode;
  locale: string;
}) {
  const t = useTranslations("console-notebooks");
  const [lockAspect, setLockAspect] = useState(true);
  const [borderPaletteOpen, setBorderPaletteOpen] = useState(false);
  const disabled = model.busy || !model.writable;
  const disabledReason = model.writable ? undefined : model.disabledReason;
  const format = model.format;
  const border = format.border && format.border !== "none" ? format.border : null;
  const arrangeLabel = (key: string) => t(`spreadsheet.chartFormatCopy.arrange.${key}` as Parameters<typeof t>[0]);
  const chevron = <OfficeSemanticIcon semanticId="menu.chevron.down" size={12} />;

  // Height and Width in the locale's length unit, as Docs' (cm in Chinese).
  const unit = documentLengthUnit(locale);
  const pxPerUnit = documentPxPerLengthUnit(unit);
  const unitLabel = unit === "cm" ? t("toolbar.indentUnitCm") : "in";
  const decimals = unit === "cm" ? 2 : 3;
  const shown = (points: number) => String(Number(((points / PT_PER_PX) / pxPerUnit).toFixed(decimals)));
  const commitSize = (axis: "width" | "height", raw: string) => {
    const px = parseDocumentLengthPx(raw, unit);
    if (px === null || !(px > 0)) return;
    const points = px * PT_PER_PX;
    const width = axis === "width" ? points : model.widthPoints;
    const height = axis === "height" ? points : model.heightPoints;
    if (!lockAspect) {
      commands.onResize(width, height);
      return;
    }
    const ratio = model.heightPoints / Math.max(1e-6, model.widthPoints);
    commands.onResize(axis === "width" ? width : height / ratio, axis === "width" ? width * ratio : height);
  };
  const setBorder = (next: Partial<PdfPictureBorder>) =>
    commands.onBorder({
      color: next.color ?? border?.color ?? DEFAULT_BORDER_COLOR,
      widthPt: next.widthPt ?? border?.widthPt ?? 1,
      dash: next.dash ?? border?.dash ?? "solid",
    });

  const largeTrigger = (props: {
    icon: Parameters<typeof OfficeSemanticIcon>[0]["semanticId"];
    label: string;
    testId: string;
    active?: boolean;
    menu?: boolean;
  }) => (
    <OfficeHomeMenuTrigger
      active={props.active}
      disabled={disabled}
      disabledReason={disabledReason}
      hasMenuIndicator={props.menu !== false}
      icon={<OfficeSemanticIcon semanticId={props.icon} size={24} />}
      label={props.label}
      testId={props.testId}
      variant="large"
    />
  );

  return (
    <>
      <span className={styles.surface} data-testid="pdf-picture-format-context" data-picture-ribbon-layout="docs" hidden />
      {pointerTools}

      {/* 图片样式: the shared gallery and Docs' Picture Border split. */}
      {commands.onInsertPicture ? (
        // WPS CT_PictureEdit opens its Edit group with 添加图片.
        <PdfSurfaceRibbonGroup className={styles.group} label={t("toolbar.ribbonInsert")} tab="picture-format" testId="pdf-picture-insert-group">
          <PdfSurfaceRibbonCommand
            controlId="PDFCTL.CONTEXT.IMAGE.INSERT"
            disabled={model.busy}
            icon={<OfficeSemanticIcon semanticId="image.insert" size={24} />}
            label={t("pdfContextRibbon.picture.insertPicture")}
            onClick={commands.onInsertPicture}
            testId="pdf-picture-insert"
          />
        </PdfSurfaceRibbonGroup>
      ) : null}
      <PdfSurfaceRibbonGroup className={styles.group} label={t("toolbar.pictureStyles")} tab="picture-format" testId="pdf-picture-styles-group">
        <OfficeHomeCommandRow className={styles.commandRow}>
          <OfficePictureStyleGallery
            application="docs"
            value={format.style}
            labels={{ ariaLabel: t("toolbar.pictureStyles"), more: t("toolbar.morePictureStyles") }}
            disabled={disabled}
            disabledReason={disabledReason}
            onPreviewChange={(id) => commands.onStylePreview?.(id)}
            onCommit={commands.onStyle}
            testIdPrefix="pdf-picture-style"
          />
          <div className={styles.borderHost}>
            <DropdownMenu modal={false}>
              <OfficeHomeSplitLayout
                testId="pdf-picture-border-split"
                primary={
                  <OfficeHomeMenuTrigger
                    active={border !== null}
                    disabled={disabled}
                    disabledReason={disabledReason}
                    icon={<OfficeSemanticIcon semanticId="image.border" size={24} />}
                    label={t("toolbar.paragraphBorders")}
                    aria-expanded={borderPaletteOpen}
                    aria-haspopup="dialog"
                    onClick={() => setBorderPaletteOpen((open) => !open)}
                    testId="pdf-picture-border"
                    variant="large"
                  />
                }
                trigger={
                  <DropdownMenuTrigger asChild>
                    <OfficeHomeMenuTrigger
                      className="note-editor-office-toolbar__split-chevron"
                      disabled={disabled}
                      icon={chevron}
                      label={t("toolbar.paragraphBorders")}
                      onPointerDown={() => setBorderPaletteOpen(false)}
                      testId="pdf-picture-border-menu"
                      variant="icon"
                    />
                  </DropdownMenuTrigger>
                }
              />
              <DropdownMenuContent align="start" className="note-editor-toolbar-menu" data-testid="pdf-picture-border-menu-content">
                <DropdownMenuItem data-testid="pdf-picture-border-none" onSelect={() => commands.onBorder("none")}>
                  <OfficePictureMenuIcon icon={{ kind: "no-outline" }} />
                  {t("toolbar.noBorder")}
                </DropdownMenuItem>
                <DropdownMenuSeparator />
                <DropdownMenuLabel>{t("toolbar.weight")}</DropdownMenuLabel>
                <DropdownMenuRadioGroup value={border ? String(border.widthPt) : ""}>
                  {BORDER_WEIGHTS.map((width) => (
                    <DropdownMenuRadioItem key={width} value={String(width)} data-testid={`pdf-picture-border-width-${width}`}
                      onSelect={() => setBorder({ widthPt: width })}>
                      <OfficePictureMenuIcon icon={{ kind: "line-weight", weightPt: width }} size={20} />
                      {width} pt
                    </DropdownMenuRadioItem>
                  ))}
                </DropdownMenuRadioGroup>
                <DropdownMenuSeparator />
                <DropdownMenuLabel>{t("toolbar.dash")}</DropdownMenuLabel>
                <DropdownMenuRadioGroup value={border?.dash ?? ""}>
                  {BORDER_DASHES.map((dash) => (
                    <DropdownMenuRadioItem key={dash} value={dash} data-testid={`pdf-picture-border-style-${dash}`}
                      onSelect={() => setBorder({ dash })}>
                      <OfficePictureMenuIcon icon={{ kind: "line-dash", dash }} size={20} />
                      {dash === "solid" ? t("toolbar.solid") : dash === "dashed" ? t("toolbar.dashed") : t("toolbar.dotted")}
                    </DropdownMenuRadioItem>
                  ))}
                </DropdownMenuRadioGroup>
              </DropdownMenuContent>
            </DropdownMenu>
            {borderPaletteOpen ? (
              <OfficeColorPalettePopover
                value={border?.color ?? undefined}
                allowNoFill
                noFillSelected={border === null}
                labels={{
                  themeColors: t("toolbar.themeColors"),
                  standardColors: t("toolbar.standardColors"),
                  recentColors: t("spreadsheet.recentColors"),
                  moreColors: t("presentation.imageBorderColor"),
                  noFill: t("presentation.home.noOutline"),
                }}
                onPick={(color) => {
                  if (color) setBorder({ color });
                  else commands.onBorder("none");
                  setBorderPaletteOpen(false);
                }}
                onDismiss={() => setBorderPaletteOpen(false)}
                testIdPrefix="pdf-picture-border"
              />
            ) : null}
          </div>
        </OfficeHomeCommandRow>
      </PdfSurfaceRibbonGroup>

      {/* 排列: Position, Bring Forward / Send Backward, Align, Rotate. */}
      <PdfSurfaceRibbonGroup className={styles.group} label={t("toolbar.arrange")} tab="picture-format" testId="pdf-picture-arrange-group">
        <OfficeHomeCommandRow className={styles.commandRow}>
          <DropdownMenu modal={false}>
            <DropdownMenuTrigger asChild>
              {largeTrigger({ icon: "object.align", label: arrangeLabel("position"), testId: "pdf-picture-position" })}
            </DropdownMenuTrigger>
            <DropdownMenuContent align="start" className="note-editor-toolbar-menu">
              <DropdownMenuLabel>{arrangeLabel("withTextWrapping")}</DropdownMenuLabel>
              <div role="group" className={styles.positionGrid}>
                {PDF_PICTURE_POSITION_PRESETS.map((preset) => {
                  const name = arrangeLabel(POSITION_KEYS[preset]);
                  return (
                    <DropdownMenuItem key={preset} aria-label={name} title={name} className={styles.positionCell}
                      data-testid={`pdf-picture-position-${preset}`} onSelect={() => commands.onPosition(preset)}>
                      <PositionGlyph preset={preset} />
                    </DropdownMenuItem>
                  );
                })}
              </div>
            </DropdownMenuContent>
          </DropdownMenu>
          <OfficeHomeCommandStack>
            <OfficeHomeCommandRow>
              <DropdownMenu modal={false}>
                <DropdownMenuTrigger asChild>
                  <OfficeHomeMenuTrigger disabled={disabled} disabledReason={disabledReason} hasMenuIndicator
                    icon={<OfficeSemanticIcon semanticId="object.bringForward" size={18} />}
                    label={arrangeLabel("bringForward")} testId="pdf-picture-bring-forward" variant="large" />
                </DropdownMenuTrigger>
                <DropdownMenuContent align="start" className="note-editor-toolbar-menu">
                  <DropdownMenuItem data-testid="pdf-picture-bring-forward-step" onSelect={() => commands.onArrange("forward")}>
                    {arrangeLabel("bringForward")}
                  </DropdownMenuItem>
                  <DropdownMenuItem data-testid="pdf-picture-bring-to-front" onSelect={() => commands.onArrange("front")}>
                    {arrangeLabel("bringToFront")}
                  </DropdownMenuItem>
                </DropdownMenuContent>
              </DropdownMenu>
              <DropdownMenu modal={false}>
                <DropdownMenuTrigger asChild>
                  <OfficeHomeMenuTrigger disabled={disabled} disabledReason={disabledReason} hasMenuIndicator
                    icon={<OfficeSemanticIcon semanticId="object.sendBackward" size={18} />}
                    label={arrangeLabel("sendBackward")} testId="pdf-picture-send-backward" variant="large" />
                </DropdownMenuTrigger>
                <DropdownMenuContent align="start" className="note-editor-toolbar-menu">
                  <DropdownMenuItem data-testid="pdf-picture-send-backward-step" onSelect={() => commands.onArrange("backward")}>
                    {arrangeLabel("sendBackward")}
                  </DropdownMenuItem>
                  <DropdownMenuItem data-testid="pdf-picture-send-to-back" onSelect={() => commands.onArrange("back")}>
                    {arrangeLabel("sendToBack")}
                  </DropdownMenuItem>
                </DropdownMenuContent>
              </DropdownMenu>
            </OfficeHomeCommandRow>
          </OfficeHomeCommandStack>
          <DropdownMenu modal={false}>
            <DropdownMenuTrigger asChild>
              {largeTrigger({ icon: "object.align", label: t("toolbar.objectAlign"), testId: "pdf-picture-align" })}
            </DropdownMenuTrigger>
            <DropdownMenuContent align="start" className="note-editor-toolbar-menu">
              <DropdownMenuLabel>{t("toolbar.objectAlign")}</DropdownMenuLabel>
              {(["left", "center", "right"] as const).map((align) => (
                <DropdownMenuItem key={align} data-testid={`pdf-picture-align-${align}`} onSelect={() => commands.onAlign(align)}>
                  <OfficePictureMenuIcon icon={{ kind: `align-${align}` }} />
                  {align === "left" ? t("toolbar.imageAlignLeft") : align === "center" ? t("toolbar.imageAlignCenter") : t("toolbar.imageAlignRight")}
                </DropdownMenuItem>
              ))}
            </DropdownMenuContent>
          </DropdownMenu>
          <DropdownMenu modal={false}>
            <DropdownMenuTrigger asChild>
              {largeTrigger({ icon: "object.rotate", label: t("toolbar.objectRotate"), testId: "pdf-picture-rotate" })}
            </DropdownMenuTrigger>
            <DropdownMenuContent align="start" className="note-editor-toolbar-menu">
              <DropdownMenuItem data-testid="pdf-picture-rotate-right" onSelect={() => commands.onRotate("right")}>
                <OfficePictureMenuIcon icon={{ kind: "rotate-right-90" }} />
                {t("presentation.rotate.right")}
              </DropdownMenuItem>
              <DropdownMenuItem data-testid="pdf-picture-rotate-left" onSelect={() => commands.onRotate("left")}>
                <OfficePictureMenuIcon icon={{ kind: "rotate-left-90" }} />
                {t("presentation.rotate.left")}
              </DropdownMenuItem>
              <DropdownMenuSeparator />
              <DropdownMenuCheckboxItem checked={false} data-testid="pdf-picture-flip-vertical" onSelect={() => commands.onFlip("vertical")}>
                <OfficeSemanticIcon semanticId="object.flipVertical" size={16} />
                {t("presentation.flip.vertical")}
              </DropdownMenuCheckboxItem>
              <DropdownMenuCheckboxItem checked={false} data-testid="pdf-picture-flip-horizontal" onSelect={() => commands.onFlip("horizontal")}>
                <OfficeSemanticIcon semanticId="object.flipHorizontal" size={16} />
                {t("presentation.flip.horizontal")}
              </DropdownMenuCheckboxItem>
              <DropdownMenuSeparator />
              <DropdownMenuItem data-testid="pdf-picture-rotate-reset" onSelect={commands.onRotateReset}>
                <OfficePictureMenuIcon icon={{ kind: "picture-reset" }} />
                {t("presentation.image.reset")}
              </DropdownMenuItem>
            </DropdownMenuContent>
          </DropdownMenu>
        </OfficeHomeCommandRow>
      </PdfSurfaceRibbonGroup>

      {/* 大小: the Crop split, then Lock Aspect Ratio, Height and Width. */}
      <PdfSurfaceRibbonGroup className={styles.group} label={t("toolbar.size")} tab="picture-format" testId="pdf-picture-size-group">
        <OfficeHomeCommandRow className={styles.commandRow}>
          <OfficeHomeSplitLayout
            testId="pdf-picture-crop-split"
            primary={
              <OfficeHomeCommand
                activation="click"
                disabled={disabled || model.cropDisabledReason !== null}
                disabledReason={disabledReason ?? model.cropDisabledReason ?? undefined}
                icon={<OfficeSemanticIcon semanticId="image.crop" size={24} />}
                label={t("presentation.crop.toggle")}
                onClick={commands.onCrop}
                testId="pdf-picture-crop"
                variant="large"
              />
            }
            trigger={
              <DropdownMenu modal={false}>
                <DropdownMenuTrigger asChild>
                  <OfficeHomeMenuTrigger className="note-editor-office-toolbar__split-chevron" disabled={disabled}
                    icon={chevron} label={t("presentation.crop.toggle")} testId="pdf-picture-crop-menu" variant="icon" />
                </DropdownMenuTrigger>
                <DropdownMenuContent align="start" className="note-editor-toolbar-menu">
                  <DropdownMenuItem data-testid="pdf-picture-crop-menu-crop" disabled={model.cropDisabledReason !== null}
                    title={model.cropDisabledReason ?? undefined} onSelect={commands.onCrop}>
                    <OfficeSemanticIcon semanticId="image.crop" size={16} />
                    {t("presentation.crop.toggle")}
                  </DropdownMenuItem>
                  <DropdownMenuSub>
                    <DropdownMenuSubTrigger data-testid="pdf-picture-crop-to-shape">
                      <OfficeSemanticIcon semanticId="shape.change" size={16} />
                      {t("presentation.cameo.cropToShape")}
                    </DropdownMenuSubTrigger>
                    <DropdownMenuSubContent className={`note-editor-toolbar-menu ${officeShapeCatalogStyles.menu}`}
                      alignOffset={-12} collisionPadding={16} sticky="always" data-testid="pdf-picture-crop-to-shape-gallery">
                      <OfficeShapeCatalogPanel
                        application="document"
                        ariaLabel={t("presentation.cameo.cropToShape")}
                        excludeCategories={["lines"]}
                        onSelect={(preset) => commands.onCropToShape(officeShapeDefinition(preset).drawingPreset)}
                        recentStorageKey="mrai.pdf.pictureCropToShape.recent"
                        selectedPreset={format.shape ? officeShapePresetFromDrawingMl(format.shape) : null}
                        testIdPrefix="pdf-picture-crop-shape"
                      />
                    </DropdownMenuSubContent>
                  </DropdownMenuSub>
                  <OfficePictureCropAspectRatioMenu
                    icon={<OfficeSemanticIcon semanticId="image.crop" size={16} />}
                    testIdPrefix="pdf-picture-crop-aspect-ratio"
                    disabled={model.cropDisabledReason !== null}
                    labels={{
                      aspectRatio: t("presentation.crop.aspectRatioLabel"),
                      square: t("presentation.crop.aspectRatio.square"),
                      portrait: t("presentation.crop.aspectRatio.portrait"),
                      landscape: t("presentation.crop.aspectRatio.landscape"),
                    }}
                    onSelect={(ratioId: PictureCropAspectRatioId) => {
                      const ratio = PICTURE_CROP_ASPECT_RATIOS.find((entry) => entry.id === ratioId);
                      if (ratio) commands.onCropAspectRatio(ratio.width, ratio.height);
                    }}
                  />
                  {/* Word's Fill and Fit act on a cropped picture's area;
                      Reset Crop takes the crop away. */}
                  <DropdownMenuSeparator />
                  <DropdownMenuItem data-testid="pdf-picture-crop-fill"
                    disabled={!model.cropped || model.cropDisabledReason !== null}
                    onSelect={() => commands.onCropMode("fill")}>
                    <OfficePictureMenuIcon icon={{ kind: "crop-fill" }} />
                    {t("presentation.crop.fill")}
                  </DropdownMenuItem>
                  <DropdownMenuItem data-testid="pdf-picture-crop-fit"
                    disabled={!model.cropped || model.cropDisabledReason !== null}
                    onSelect={() => commands.onCropMode("fit")}>
                    <OfficePictureMenuIcon icon={{ kind: "crop-fit" }} />
                    {t("presentation.crop.fit")}
                  </DropdownMenuItem>
                  <DropdownMenuSeparator />
                  <DropdownMenuItem data-testid="pdf-picture-crop-reset" disabled={!model.cropped}
                    onSelect={() => commands.onCropMode("reset")}>
                    <OfficePictureMenuIcon icon={{ kind: "crop-reset" }} />
                    {t("presentation.crop.reset")}
                  </DropdownMenuItem>
                </DropdownMenuContent>
              </DropdownMenu>
            }
          />
          <OfficeHomeCommandStack className={styles.sizeStack}>
            <OfficeHomeCommandRow className={styles.sizeRow} testId="pdf-picture-size-aspect-row">
              <OfficeHomeCheckbox
                checked={lockAspect}
                className={styles.aspectTarget}
                disabled={disabled}
                label={t("toolbar.lockAspectRatio")}
                onCheckedChange={setLockAspect}
                testId="pdf-picture-lock-ratio"
              />
            </OfficeHomeCommandRow>
            <OfficeHomeCommandRow className={styles.sizeRow} testId="pdf-picture-size-height-row">
              <OfficeRibbonNumberField
                key={`h:${model.heightPoints.toFixed(2)}`}
                icon={<OfficeSemanticIcon semanticId="object.height" size={22} />}
                label={t("toolbar.pageSetupHeight")}
                testId="pdf-picture-height"
                acceptsUnits
                value={shown(model.heightPoints)}
                disabled={disabled}
                disabledReason={disabledReason}
                min={0.1}
                max={unit === "cm" ? 508 : 200}
                step={0.1}
                unit={unitLabel}
                inputWidthPx={61}
                classNames={{ field: styles.sizeField, input: styles.sizeInput, icon: styles.sizeIcon }}
                onCommit={(raw) => commitSize("height", raw)}
              />
            </OfficeHomeCommandRow>
            <OfficeHomeCommandRow className={styles.sizeRow} testId="pdf-picture-size-width-row">
              <OfficeRibbonNumberField
                key={`w:${model.widthPoints.toFixed(2)}`}
                icon={<OfficeSemanticIcon semanticId="object.width" size={22} />}
                label={t("toolbar.pageSetupWidth")}
                testId="pdf-picture-width"
                acceptsUnits
                value={shown(model.widthPoints)}
                disabled={disabled}
                disabledReason={disabledReason}
                min={0.1}
                max={unit === "cm" ? 508 : 200}
                step={0.1}
                unit={unitLabel}
                inputWidthPx={61}
                classNames={{ field: styles.sizeField, input: styles.sizeInput, icon: styles.sizeIcon }}
                onCommit={(raw) => commitSize("width", raw)}
              />
            </OfficeHomeCommandRow>
          </OfficeHomeCommandStack>
        </OfficeHomeCommandRow>
      </PdfSurfaceRibbonGroup>

      {/* 调整: Artistic Effects, Corrections, Color, Transparency, Compress
          Pictures, Change Picture, Reset Picture. */}
      <PdfSurfaceRibbonGroup className={styles.group} label={t("toolbar.adjust")} tab="picture-format" testId="pdf-picture-adjust-group">
        <OfficeHomeCommandRow className={styles.commandRow}>
          <OfficeHomeCommand
            active={model.formatPaneOpen}
            ariaExpanded={model.formatPaneOpen}
            icon={<OfficeSemanticIcon semanticId="image.format" size={24} />}
            label={t("presentation.image.formatPicture")}
            onClick={commands.onFormatPane}
            testId="pdf-picture-format-pane"
            variant="large"
          />
          <OfficeArtisticEffectsMenu
            value={format.artisticEffect ?? "none"}
            disabled={disabled}
            disabledReason={disabledReason}
            onCommit={commands.onArtisticEffect}
            testIdPrefix="pdf-picture-artistic-effects"
            trigger={largeTrigger({ icon: "image.corrections", label: t("toolbar.artisticEffects"), testId: "pdf-picture-artistic-effects", active: format.artisticEffect !== null })}
          />
          <OfficePictureCorrectionsMenu
            trigger={largeTrigger({ icon: "image.corrections", label: t("toolbar.corrections"), testId: "pdf-picture-corrections", active: format.brightness !== 0 || format.contrast !== 0 || format.sharpness !== 0 })}
            disabled={disabled}
            disabledReason={disabledReason}
            testIdPrefix="pdf-picture-corrections"
            brightnessDelta={format.brightness}
            contrastDelta={format.contrast}
            recolor="none"
            sharpnessUnavailableReason=""
            sharpness={{ value: format.sharpness, onCommit: commands.onSharpness }}
            onCommit={commands.onCorrections}
            extraItems={
              <>
                <DropdownMenuSeparator />
                <DropdownMenuItem data-testid="pdf-picture-corrections-sharpen" onSelect={commands.onSharpen}>
                  <OfficePictureMenuIcon icon={{ kind: "sharpness" }} />
                  {t("pdf.viewer.picture.sharpen")}
                </DropdownMenuItem>
              </>
            }
          />
          <OfficePictureColorMenu
            trigger={largeTrigger({ icon: "image.color", label: t("toolbar.color"), testId: "pdf-picture-color", active: format.saturation !== 100 || format.recolor !== null || format.tone !== 6500 })}
            disabled={disabled}
            disabledReason={disabledReason}
            testIdPrefix="pdf-picture-color"
            recolor="none"
            unavailableReason=""
            onCommitRecolor={() => undefined}
            omitUnavailable
            saturation={{ value: format.saturation, onCommit: commands.onSaturation }}
            tone={{ value: format.tone, onCommit: commands.onTone }}
            recolorGallery={{
              rows: OFFICE_RECOLOR_GALLERY_ROWS,
              value: format.recolor ?? "none",
              label: (id) => t(officeRecolorLabelKey(id as OfficeRecolorPresetId | "none") as Parameters<typeof t>[0]),
              renderSample: (id) => <RecolorSample id={id as OfficeRecolorPresetId | "none"} />,
              onCommit: (id) => commands.onRecolor(id === "none" ? null : (id as OfficeRecolorPresetId)),
            }}
          />
          <OfficePictureTransparencyMenu
            trigger={largeTrigger({ icon: "image.transparency", label: t("toolbar.transparency"), testId: "pdf-picture-transparency", active: model.transparency > 0 })}
            disabled={disabled}
            disabledReason={disabledReason}
            testIdPrefix="pdf-picture-transparency"
            transparency={model.transparency}
            onCommit={commands.onTransparency}
          />
          <OfficeHomeCommand
            activation="click"
            controlId="PDFCTL.CONTEXT.PICTURE_FORMAT.COMPRESS"
            disabled={disabled}
            disabledReason={disabledReason}
            icon={<OfficeSemanticIcon semanticId="image.compress" size={24} />}
            label={t("toolbar.compressPictures")}
            onClick={commands.onCompress}
            testId="pdf-picture-compress"
            variant="large"
          />
          <DropdownMenu modal={false}>
            <DropdownMenuTrigger asChild>
              {largeTrigger({ icon: "image.change", label: t("toolbar.changePicture"), testId: "pdf-picture-change" })}
            </DropdownMenuTrigger>
            <DropdownMenuContent align="start" className="note-editor-toolbar-menu" data-testid="pdf-picture-change-menu">
              <DropdownMenuItem data-testid="pdf-picture-change-from-device" onSelect={commands.onChangePicture}>
                <OfficeSemanticIcon semanticId="image.change" size={16} />
                {t("toolbar.pictureFromDevice")}
              </DropdownMenuItem>
            </DropdownMenuContent>
          </DropdownMenu>
          <DropdownMenu modal={false}>
            <OfficeHomeSplitLayout
              className={styles.resetSplit}
              testId="pdf-picture-reset-split"
              primary={
                <OfficeHomeCommand
                  disabled={disabled || !model.canReset}
                  disabledReason={disabledReason ?? (model.canReset ? undefined : t("pdfContextRibbon.picture.noEditsToReset"))}
                  icon={<OfficeSemanticIcon semanticId="image.reset" size={24} />}
                  label={t("toolbar.resetPicture")}
                  onClick={() => commands.onReset(false)}
                  testId="pdf-picture-reset"
                  variant="large"
                />
              }
              trigger={
                <DropdownMenuTrigger asChild>
                  <OfficeHomeMenuTrigger className="note-editor-office-toolbar__split-chevron" disabled={disabled || !model.canReset}
                    icon={chevron} label={t("toolbar.resetPicture")} testId="pdf-picture-reset-menu" variant="icon" />
                </DropdownMenuTrigger>
              }
            />
            <DropdownMenuContent align="start" className="note-editor-toolbar-menu" data-testid="pdf-picture-reset-menu-content">
              <DropdownMenuItem data-testid="pdf-picture-reset-menu-picture" onSelect={() => commands.onReset(false)}>
                <OfficeSemanticIcon semanticId="image.reset" size={16} />
                {t("toolbar.resetPicture")}
              </DropdownMenuItem>
              <DropdownMenuItem data-testid="pdf-picture-reset-menu-picture-and-size" onSelect={() => commands.onReset(true)}>
                <OfficeSemanticIcon semanticId="image.reset" size={16} />
                {t("presentation.image.resetAndSize")}
              </DropdownMenuItem>
            </DropdownMenuContent>
          </DropdownMenu>
        </OfficeHomeCommandRow>
      </PdfSurfaceRibbonGroup>

      {/* 绘图: draw on the picture (Docs' Draw dialog), written into its pixels. */}
      <PdfSurfaceRibbonGroup className={styles.group} label={t("toolbar.draw")} tab="picture-format" testId="pdf-picture-draw-group">
        <OfficeHomeCommandRow className={styles.commandRow}>
          <OfficeHomeCommand
            activation="click"
            disabled={disabled}
            disabledReason={disabledReason}
            icon={<DocumentPictureDrawIcon size={24} />}
            label={t("toolbar.draw")}
            onClick={commands.onDraw}
            testId="pdf-picture-draw"
            variant="large"
          />
        </OfficeHomeCommandRow>
      </PdfSurfaceRibbonGroup>

      {/* 辅助功能: Word's Alt Text, written to the picture's /Figure. */}
      <PdfSurfaceRibbonGroup className={styles.group} label={t("toolbar.accessibility")} tab="picture-format" testId="pdf-picture-alt-group">
        <OfficeHomeCommandRow className={styles.commandRow}>
          <OfficeHomeCommand
            active={model.altPaneOpen || model.hasAlt}
            icon={<OfficeSemanticIcon semanticId="object.altText" size={24} />}
            label={t("toolbar.imageAlt")}
            onClick={commands.onAltText}
            testId="pdf-picture-alt-text"
            variant="large"
          />
        </OfficeHomeCommandRow>
      </PdfSurfaceRibbonGroup>

      {/* PDF: every contextual tab ends with 退出编辑; WPS's picture tab
          also extracts and deletes the picture (CT_PictureEdit). */}
      <PdfSurfaceRibbonGroup label={t("pdfContextRibbon.picture.actions")} tab="picture-format" testId="pdf-picture-context-actions-group">
        <PdfSurfaceRibbonCommand
          controlId="PDFCTL.CONTEXT.IMAGE.EXTRACT"
          disabled={model.busy}
          icon={<OfficeSemanticIcon semanticId="image.download" size={24} />}
          label={t("pdfContextRibbon.picture.extractPicture")}
          onClick={commands.onExtractPicture}
          testId="pdf-picture-extract"
        />
        <PdfSurfaceRibbonCommand
          controlId="PDFCTL.CONTEXT.IMAGE.EXTRACT_TEXT"
          disabled={model.busy}
          icon={<OfficeSemanticIcon semanticId="scan.line" size={24} />}
          label={t("pdf.viewer.picture.extractText")}
          onClick={commands.onExtractText}
          testId="pdf-picture-extract-text"
        />
        {null}
        <PdfSurfaceRibbonCommand
          controlId="PDFCTL.CONTEXT.IMAGE.DELETE"
          disabled={model.busy}
          icon={<OfficeSemanticIcon semanticId="edit.delete" size={24} />}
          label={t("pdfContextRibbon.picture.wpsDelete")}
          onClick={commands.onDelete}
          testId="pdf-picture-delete"
        />
        <PdfSurfaceRibbonCommand
          controlId="PDFCTL.CONTEXT.IMAGE.EXIT"
          disabled={model.busy}
          icon={<OfficeSemanticIcon semanticId="edit.exit" size={24} />}
          label={t("pdfContextRibbon.picture.exitEditing")}
          onClick={commands.onExit}
          testId="pdf-context-image-exit"
        />
      </PdfSurfaceRibbonGroup>
    </>
  );
}
