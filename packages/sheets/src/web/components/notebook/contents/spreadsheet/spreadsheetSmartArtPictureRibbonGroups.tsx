"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Excel's Picture Format tab for a SmartArt picture slot holding a picture:
 * Adjust > Reset Picture, Picture Styles, and Size > Crop (Fill / Fit) --
 * Microsoft Support, "Crop a picture to fit in a shape" and "Add a border or
 * another style to a picture"; cropping is enabled for pictures in SmartArt
 * graphics since Office 2010. The slot keeps the box its layout gave it; only
 * its picture is cropped and styled (`lib/office-smartart/slot-picture-format`).
 */

import { OfficePictureStyleGallery } from "@mrnote/core/web/components/office-ribbon/OfficePictureStyleGallery.tsx";
import { OfficeSemanticIcon } from "@mrnote/core/web/components/office-ribbon/icons/OfficeSemanticIcon.tsx";
import type { OfficePictureStyleId } from "@mrnote/core/web/lib/office-picture-styles/catalog.ts";
import type { SmartArtSlotPictureCommand, SmartArtSlotPictureFormat } from "@mrnote/core/web/lib/office-smartart/slot-picture-format.ts";
import type { SpreadsheetRibbonGroup } from "../../../../lib/spreadsheet/SpreadsheetRibbon";
import { SpreadsheetRibbonAction } from "./SpreadsheetRibbonControls";

export interface SpreadsheetSmartArtPictureRibbonGroupsInputs {
  /** The selected slot's current Picture Format; `null` when no filled slot is selected. */
  format: SmartArtSlotPictureFormat | null;
  disabledReason?: string;
  onCommand: (command: SmartArtSlotPictureCommand) => void;
  onStylePreview?: (style: OfficePictureStyleId | null) => void;
  t: (key: string) => string;
}

export function spreadsheetSmartArtPictureRibbonGroups({
  format,
  disabledReason,
  onCommand,
  onStylePreview,
  t,
}: SpreadsheetSmartArtPictureRibbonGroupsInputs): SpreadsheetRibbonGroup[] {
  if (!format) return [];
  const disabled = Boolean(disabledReason);
  return [
    {
      id: "smartart-picture-adjust",
      label: t("presentation.image.adjustGroup"),
      content: (
        <SpreadsheetRibbonAction
          controlId="XLSCTL.CONTEXT.SHAPE.FILL_LINE.MRAI.APPLY"
          disabled={disabled || (!format.crop && !format.style)}
          disabledReason={disabledReason}
          homeVariant="large"
          icon={<OfficeSemanticIcon semanticId="image.corrections" size={24} />}
          label={t("presentation.pictureRibbon.resetPicture")}
          onClick={() => onCommand({ kind: "reset" })}
          testId="spreadsheet-ribbon-smartart-picture-reset"
        />
      ),
    },
    {
      id: "smartart-picture-styles",
      label: t("presentation.image.stylesGroup"),
      content: (
        <span className="relative flex min-h-[64px] items-center px-1">
          <OfficePictureStyleGallery
            application="spreadsheet"
            value={format.style ?? null}
            disabled={disabled}
            disabledReason={disabledReason}
            labels={{
              ariaLabel: t("presentation.image.stylesGroup"),
              more: t("spreadsheet.uiCopy.moreStyles"),
            }}
            visibleItemCount={6}
            onPreviewChange={(id) => onStylePreview?.(id && !disabled ? id : null)}
            onCommit={(id) => onCommand({ kind: "style", style: id })}
            testIdPrefix="spreadsheet-smartart-picture-style-gallery"
          />
        </span>
      ),
    },
    {
      id: "smartart-picture-size",
      label: t("presentation.smartartContext.sizeGroup"),
      content: (
        <div className="flex min-h-[72px] items-center gap-2">
          <OfficeSemanticIcon semanticId="image.crop" size={24} />
          <div className="flex flex-col gap-1">
            {(["fill", "fit"] as const).map((mode) => (
              <SpreadsheetRibbonAction
                key={mode}
                controlId="XLSCTL.CONTEXT.SHAPE.FILL_LINE.MRAI.APPLY"
                disabled={disabled}
                disabledReason={disabledReason}
                homeVariant="compact"
                icon={<OfficeSemanticIcon semanticId="image.crop" size={16} />}
                label={t(`presentation.crop.${mode}`)}
                onClick={() => onCommand({ kind: "crop", mode })}
                testId={`spreadsheet-ribbon-smartart-picture-crop-${mode}`}
              />
            ))}
          </div>
        </div>
      ),
    },
  ];
}
