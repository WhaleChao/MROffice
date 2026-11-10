"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Excel's Picture Format > Adjust > Corrections, Color and Transparency on a
 * sheet picture: the shared galleries (`OfficePictureAdjustMenus`), each tile
 * drawn through the filter the grid draws the picture with
 * (`office-recolor-presets` in the workbook's theme). What Excel saves on the
 * blip is offered -- the 25 brightness / contrast tiles, the 21 Recolor
 * tiles, the Transparency presets; Sharpen / Soften, Saturation, Tone and Set
 * Transparent Color are not, as the sheet keeps none of them.
 */

import { useId, type ReactNode } from "react";
import { useTranslations } from "@mrnote/core/host/next-intl.tsx";

import { OfficeAdjustSvgFilter } from "@mrnote/core/web/components/office-ribbon/OfficeAdjustSvgFilter.tsx";
import {
  OfficePictureAdjustSample,
  OfficePictureColorMenu,
  OfficePictureCorrectionsMenu,
  OfficePictureTransparencyMenu,
} from "@mrnote/core/web/components/office-ribbon/OfficePictureAdjustMenus.tsx";
import {
  isOfficeRecolorPreset,
  officeAdjustFilterSteps,
  officeRecolorLabelKey,
  OFFICE_RECOLOR_GALLERY_ROWS,
  type OfficeRecolorPresetId,
} from "@mrnote/core/web/lib/office-recolor-presets.ts";
import type { OfficeThemeGridColors } from "@mrnote/core/web/lib/office-theme-color-grid.ts";
import type { SpreadsheetObject } from "./spreadsheetSheetModel";

function AdjustTile({ steps }: { steps: Parameters<typeof OfficeAdjustSvgFilter>[0]["steps"] }) {
  const id = `sheet-picture-adjust-tile-${useId().replace(/[^A-Za-z0-9_-]/g, "")}`;
  return (
    <>
      <OfficeAdjustSvgFilter id={id} steps={steps} />
      <OfficePictureAdjustSample filter={steps.length ? `url(#${id})` : undefined} />
    </>
  );
}

function corrections(object: SpreadsheetObject): { brightness: number; contrast: number } {
  const value = object.pictureCorrections;
  return {
    brightness: typeof value?.brightness === "number" ? value.brightness : 0,
    contrast: typeof value?.contrast === "number" ? value.contrast : 0,
  };
}

function recolor(object: SpreadsheetObject): OfficeRecolorPresetId | null {
  return isOfficeRecolorPreset(object.pictureRecolor) ? object.pictureRecolor : null;
}

interface MenuCommon {
  readonly object: SpreadsheetObject;
  readonly theme: OfficeThemeGridColors | null;
  readonly trigger: ReactNode;
  readonly disabled: boolean;
  readonly disabledReason?: string;
  readonly onOptions: () => void;
}

export function SpreadsheetPictureCorrectionsMenu({
  object,
  theme,
  trigger,
  disabled,
  disabledReason,
  onOptions,
  onCommit,
}: MenuCommon & { readonly onCommit: (brightness: number, contrast: number) => void }) {
  const t = useTranslations("console-notebooks");
  const current = corrections(object);
  const preset = recolor(object);
  return (
    <OfficePictureCorrectionsMenu
      trigger={trigger}
      brightnessDelta={current.brightness}
      contrastDelta={current.contrast}
      recolor="none"
      disabled={disabled}
      disabledReason={disabledReason}
      sharpnessUnavailableReason=""
      showSharpness={false}
      renderTileSample={(brightness, contrast) => (
        <AdjustTile
          steps={officeAdjustFilterSteps(
            { recolor: preset === "washout" ? null : preset, brightness, contrast },
            theme,
          )}
        />
      )}
      onCommit={onCommit}
      optionsLabel={t("presentation.pictureRibbon.correctionsOptions")}
      onOptions={onOptions}
      testIdPrefix="spreadsheet-ribbon-picture-corrections"
    />
  );
}

export function SpreadsheetPictureColorMenu({
  object,
  theme,
  trigger,
  disabled,
  disabledReason,
  onOptions,
  onCommit,
}: MenuCommon & { readonly onCommit: (preset: OfficeRecolorPresetId | null) => void }) {
  const t = useTranslations("console-notebooks");
  const preset = recolor(object);
  const current = preset === "washout" ? { brightness: 0, contrast: 0 } : corrections(object);
  return (
    <OfficePictureColorMenu
      trigger={trigger}
      recolor="none"
      disabled={disabled}
      disabledReason={disabledReason}
      unavailableReason=""
      onCommitRecolor={() => undefined}
      showSaturationTone={false}
      showTransparentColor={false}
      recolorGallery={{
        rows: OFFICE_RECOLOR_GALLERY_ROWS,
        value: preset ?? "none",
        label: (id) => t(officeRecolorLabelKey(id as OfficeRecolorPresetId)),
        renderSample: (id) => (
          <AdjustTile
            steps={officeAdjustFilterSteps(
              { recolor: isOfficeRecolorPreset(id) ? id : null, ...(id === "washout" ? {} : current) },
              theme,
            )}
          />
        ),
        onCommit: (id) => onCommit(isOfficeRecolorPreset(id) ? id : null),
      }}
      optionsLabel={t("presentation.pictureRibbon.colorOptions")}
      onOptions={onOptions}
      testIdPrefix="spreadsheet-ribbon-picture-color"
    />
  );
}

export function SpreadsheetPictureTransparencyMenu({
  object,
  trigger,
  disabled,
  disabledReason,
  onOptions,
  onCommit,
}: Omit<MenuCommon, "theme"> & { readonly onCommit: (transparency: number) => void }) {
  const t = useTranslations("console-notebooks");
  return (
    <OfficePictureTransparencyMenu
      trigger={trigger}
      transparency={typeof object.pictureTransparency === "number" ? object.pictureTransparency : 0}
      disabled={disabled}
      disabledReason={disabledReason}
      onCommit={onCommit}
      optionsLabel={t("presentation.pictureRibbon.transparencyOptions")}
      onOptions={onOptions}
      testIdPrefix="spreadsheet-ribbon-picture-transparency"
    />
  );
}
