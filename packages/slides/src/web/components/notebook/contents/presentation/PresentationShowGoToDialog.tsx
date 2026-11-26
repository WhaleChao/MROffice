"use client";












/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { useId, useRef, useState } from "react";
import { useTranslations } from "@mrnote/core/host/next-intl.tsx";

import { OfficeDialogShell } from "@mrnote/core/web/components/office-ribbon/OfficeDialogShell.tsx";
import { officeDialogParts } from "@mrnote/core/web/components/office-ribbon/officeDialogChrome.tsx";
import type { PresentationBlock } from "@mrnote/core/web/components/console/editor/native-document/documentCommands.ts";
import { presentationOutlineProjection } from "../../../../lib/presentation-outline-view";

type GoToSlide = {
  readonly id: string;
  readonly title?: string;
  readonly blocks?: readonly PresentationBlock[];
  /** A slide read from a file: where it came from (its title is then the file's). */
  readonly source?: unknown;
};

export interface PresentationShowGoToDialogProps<S extends GoToSlide = GoToSlide> {
  readonly slides: readonly S[];
  readonly currentIndex: number;
  /** The slide shown before this one (Last slide viewed), or null. */
  readonly lastViewedIndex: number | null;
  readonly isHidden: (slide: S) => boolean;
  readonly onGoTo: (index: number) => void;
  readonly onCancel: () => void;
  /** Where the dialog is drawn: inside the show (it may be full screen). */
  readonly container?: HTMLElement | null;
  /**
   * What takes the keyboard back as the dialog closes: the show (PowerPoint
   * puts its window in front again, truth show-see-all-v2), not whatever had
   * it before the dialog opened.
   */
  readonly returnFocusTarget?: () => HTMLElement | null;
}

export function PresentationShowGoToDialog<S extends GoToSlide>(props: PresentationShowGoToDialogProps<S>) {
  const t = useTranslations("console-notebooks");
  const id = useId();
  const listRef = useRef<HTMLSelectElement | null>(null);
  const [picked, setPicked] = useState(props.currentIndex);
  // A slide's name here is its title placeholder's words, as PowerPoint
  // lists it; a slide read from a file with none is "Slide N" in the
  // interface's language (PowerPoint's own word for a slide: Folie,
  // Diapositive, 幻灯片 ...), never the words of some other box on it, which
  // the import keeps as the slide's stored title (truth show-see-all 13: a
  // number typed in a text box listed as "1. Slide 1"). A slide made here
  // goes by the title its author gave it.
  const name = (index: number) => {
    const slide = props.slides[index];
    const projection = slide
      ? presentationOutlineProjection({ title: "", blocks: slide.blocks ?? [] })
      : null;
    const placeholderTitle = projection?.titleBlockId ? projection.title.replace(/\s+/g, " ").trim() : "";
    const authoredTitle = slide && slide.source == null && typeof slide.title === "string" ? slide.title.replace(/\s+/g, " ").trim() : "";
    return placeholderTitle || authoredTitle || t("presentation.slideThumbnailName", { index: index + 1 });
  };
  const line = (index: number) =>
    props.slides[index] && props.isHidden(props.slides[index]) ? `(${index + 1}.) ${name(index)}` : `${index + 1}. ${name(index)}`;
  return (
    <OfficeDialogShell
      open
      title={t("presentation.slideshow.allSlidesDialog.title")}
      description={<span className="sr-only">{t("presentation.slideshow.allSlidesDialog.slideTitles")}</span>}
      defaultActionLabel={t("presentation.slideshow.allSlidesDialog.goTo")}
      cancelActionLabel={t("toolbar.cancel")}
      onDefaultAction={() => props.onGoTo(picked)}
      onCancel={props.onCancel}
      initialFocusRef={listRef}
      container={props.container ?? undefined}
      returnFocusTarget={props.returnFocusTarget}
      className="presentation-show-goto"
      testIdPrefix="presentation-show-goto"
    >
      <div className="grid gap-2">
        <label className={officeDialogParts.field} htmlFor={`${id}-show`}>
          <span className={officeDialogParts.label}>{t("presentation.slideshow.allSlidesDialog.show")}</span>
        </label>
        {/* Custom shows are not offered here yet: All Slides, as PowerPoint
            shows it for a deck with none. */}
        <select id={`${id}-show`} disabled value="all" data-testid="presentation-show-goto-show">
          <option value="all">{t("presentation.slideshow.allSlidesDialog.title")}</option>
        </select>
        <label className={officeDialogParts.label} htmlFor={`${id}-titles`}>
          {t("presentation.slideshow.allSlidesDialog.slideTitles")}
        </label>
        <select
          id={`${id}-titles`}
          ref={listRef}
          size={Math.min(10, Math.max(4, props.slides.length))}
          value={String(picked)}
          onChange={(event) => setPicked(Number(event.target.value))}
          onDoubleClick={() => props.onGoTo(picked)}
          data-testid="presentation-show-goto-titles"
        >
          {props.slides.map((slide, index) => (
            <option key={slide.id} value={String(index)} data-testid={`presentation-show-goto-title-${index + 1}`}>
              {line(index)}
            </option>
          ))}
        </select>
        <fieldset className={officeDialogParts.group}>
          <legend className={officeDialogParts.sectionTitle}>{t("presentation.slideshow.allSlidesDialog.lastViewed")}</legend>
          <span data-testid="presentation-show-goto-last">
            {props.lastViewedIndex !== null ? `${props.lastViewedIndex + 1}. ${name(props.lastViewedIndex)}` : ""}
          </span>
        </fieldset>
      </div>
    </OfficeDialogShell>
  );
}
