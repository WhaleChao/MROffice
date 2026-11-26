"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * PowerPoint's question when a deck is saved as pictures: "Which slides do
 * you want to export?" -- All Slides (each slide a separate picture) or Just
 * This One (the selected slide); Microsoft Support, "Save a slide as an image
 * or as a separate presentation file". The resolution is PowerPoint's 96 dpi
 * unless a higher one is chosen (`lib/presentation-picture-export`).
 */

import { useState, type RefObject } from "react";
import { useTranslations } from "@mrnote/core/host/next-intl.tsx";

import { OfficeDialogShell } from "@mrnote/core/web/components/office-ribbon/OfficeDialogShell.tsx";
import { officeDialogParts } from "@mrnote/core/web/components/office-ribbon/officeDialogChrome.tsx";
import {
  SLIDE_PICTURE_DEFAULT_DPI,
  SLIDE_PICTURE_RESOLUTIONS,
  slidePicturePixels,
  type SlidePictureResolution,
} from "../../../../lib/presentation-picture-export";

export interface PresentationExportPicturesDialogProps {
  readonly open: boolean;
  /** The picture format's name, as the menu shows it (PNG, JPEG, ...). */
  readonly formatLabel: string;
  readonly slideWidthEmu: number | null | undefined;
  readonly slideHeightEmu: number | null | undefined;
  readonly onAllSlides: (dpi: SlidePictureResolution) => void;
  readonly onJustThisOne: (dpi: SlidePictureResolution) => void;
  readonly onCancel: () => void;
  readonly returnFocusRef?: RefObject<HTMLElement | null>;
}

export function PresentationExportPicturesDialog({
  open,
  formatLabel,
  slideWidthEmu,
  slideHeightEmu,
  onAllSlides,
  onJustThisOne,
  onCancel,
  returnFocusRef,
}: PresentationExportPicturesDialogProps) {
  const t = useTranslations("console-notebooks");
  const [dpi, setDpi] = useState<SlidePictureResolution>(SLIDE_PICTURE_DEFAULT_DPI);
  return (
    <OfficeDialogShell
      open={open}
      title={t("presentation.exportPictures.title", { format: formatLabel })}
      description={t("presentation.exportPictures.question")}
      defaultActionLabel={t("presentation.exportPictures.allSlides")}
      cancelActionLabel={t("toolbar.cancel")}
      onDefaultAction={() => onAllSlides(dpi)}
      onCancel={onCancel}
      footerBeforeDefault={
        <button
          type="button"
          className={officeDialogParts.button}
          onClick={() => onJustThisOne(dpi)}
          data-testid="presentation-export-pictures-this-one"
        >
          {t("presentation.exportPictures.justThisOne")}
        </button>
      }
      returnFocusRef={returnFocusRef}
      testIdPrefix="presentation-export-pictures"
      defaultActionTestId="presentation-export-pictures-all"
      className="sm:max-w-[420px]"
    >
      <label className={officeDialogParts.field}>
        <span className={officeDialogParts.label}>{t("presentation.exportPictures.resolution")}</span>
        <select
          value={dpi}
          onChange={(event) => setDpi(Number(event.target.value) as SlidePictureResolution)}
          data-testid="presentation-export-pictures-resolution"
        >
          {SLIDE_PICTURE_RESOLUTIONS.map((value) => {
            const size = slidePicturePixels(value, slideWidthEmu, slideHeightEmu);
            return (
              <option key={value} value={value}>
                {t("presentation.exportPictures.resolutionOption", { dpi: value, width: size.width, height: size.height })}
              </option>
            );
          })}
        </select>
      </label>
    </OfficeDialogShell>
  );
}
