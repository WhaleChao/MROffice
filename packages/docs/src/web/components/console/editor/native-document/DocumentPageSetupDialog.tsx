"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { EMU_PER_TWIP } from "@mrnote/core/web/lib/office-units.ts";
import { useMemo, useRef, useState } from "react";
import { useLocale, useTranslations } from "@mrnote/core/host/next-intl.tsx";
import {
  formatPresentationMeasurement,
  parsePresentationMeasurement,
  presentationMeasurementUnit,
  type PresentationMeasurementUnit,
} from "@mrnote/core/web/lib/presentation-measurement.ts";

import { OfficeDialogShell } from "@mrnote/core/web/components/office-ribbon/OfficeDialogShell.tsx";
import { OfficeSemanticIcon } from "@mrnote/core/web/components/office-ribbon/icons/OfficeSemanticIcon.tsx";

import { documentValueInputClassName } from "@mrnote/core/web/components/console/editor/documentValueInput.ts";
import styles from "@mrnote/core/web/components/console/editor/native-document/DocumentLayoutDialogs.module.css";
import {
  sectionApplyToOptions,
  defaultSectionApplyTo,
  type SectionApplyTo,
} from "@mrnote/core/web/components/console/editor/native-document/sectionApplyTo.ts";
import type { SectionBreakType } from "@mrnote/core/web/components/console/editor/extensions/SectionBreak.ts";
import type { SectionLayout, SectionVerticalAlignment } from "@mrnote/core/web/components/console/editor/native-document/sectionLayouts.ts";
import {
  resolveWordPageBox,
  type WordDocumentSettings,
  type WordMultiplePages,
} from "@mrnote/core/web/components/console/editor/native-document/wordDocumentSettings.ts";

/*
 * Word's Page Setup dialog (Layout > Margins > Custom Margins… opens it on
 * Margins, Layout > Size > More Paper Sizes… on Paper). Semantics reproduced,
 * from "Change margins", "Set gutter margins", "Set mirror margins for facing
 * pages", "Page margin basics" and "Change the paper size" (Microsoft
 * Support) and ECMA-376 17.6.11 pgMar / 17.6.13 pgSz / 17.15.1.78 settings:
 *
 * Margins tab
 * - Top, Bottom, Left, Right, Gutter: w:pgMar top/bottom/left/right/gutter.
 *   Mirror margins and Book fold rename Left/Right to Inside/Outside; 2 pages
 *   per sheet renames the two margins across its fold the same way.
 * - Gutter position Left/Top: settings w:gutterAtTop. Unavailable (decided by
 *   Word) for Mirror margins, 2 pages per sheet and Book fold.
 * - Orientation Portrait/Landscape: w:pgSz w:orient; switching swaps the
 *   paper's width and height.
 * - Pages > Multiple pages: Normal / Mirror margins / 2 pages per sheet /
 *   Book fold (settings w:mirrorMargins / w:printTwoOnOne /
 *   w:bookFoldPrinting). Book fold switches the sheet to landscape.
 * Paper tab
 * - Paper size: Word's paper list, or Custom size when the width and height
 *   match none; Width/Height are the page as oriented.
 * Both tabs
 * - Apply to: Whole document / This section (several sections only) / This
 *   point forward (inserts a Next Page section break at the insertion point).
 * - OK applies both tabs at once, in one undo step; Cancel discards both.
 *
 * Measurements are in the Office unit of the UI language (Microsoft Support,
 * "Change the measurement units from inches to centimeters"; the shared
 * officeLocaleUsesInches rule): centimetres in Chinese Word, inches in US
 * English. A typed value may name its own unit.
 */

const TWIPS_PER_INCH = 1_440;
const mmToTwips = (mm: number) => Math.round((mm * TWIPS_PER_INCH) / 25.4);

/** Word's paper list, portrait width x height in twips. */
export const WORD_PAPER_SIZES = [
  { id: "letter", widthTwips: 12_240, heightTwips: 15_840 },
  { id: "legal", widthTwips: 12_240, heightTwips: 20_160 },
  { id: "executive", widthTwips: 10_440, heightTwips: 15_120 },
  { id: "a3", widthTwips: 16_838, heightTwips: 23_811 },
  { id: "a4", widthTwips: 11_906, heightTwips: 16_838 },
  { id: "a5", widthTwips: 8_391, heightTwips: 11_906 },
  { id: "b4", widthTwips: 14_570, heightTwips: 20_636 },
  { id: "b5", widthTwips: 10_318, heightTwips: 14_570 },
  { id: "tabloid", widthTwips: 15_840, heightTwips: 24_480 },
  { id: "envelope10", widthTwips: 5_940, heightTwips: 13_680 },
  { id: "envelopeDL", widthTwips: mmToTwips(110), heightTwips: mmToTwips(220) },
  { id: "envelopeC5", widthTwips: mmToTwips(162), heightTwips: mmToTwips(229) },
  { id: "envelopeB5", widthTwips: mmToTwips(176), heightTwips: mmToTwips(250) },
  { id: "envelopeMonarch", widthTwips: 5_580, heightTwips: 10_800 },
  {
    id: "japanesePostcard",
    widthTwips: mmToTwips(100),
    heightTwips: mmToTwips(148),
  },
] as const;
export type WordPaperSizeId = (typeof WORD_PAPER_SIZES)[number]["id"];

/** Word's paper width/height bounds: 0.1 in to 22 in. */
const MIN_PAPER_TWIPS = 144;
const MAX_PAPER_TWIPS = 31_680;
/** Word keeps at least this much text width/height between the margins. */
const MIN_TEXT_TWIPS = 144;

export interface PageSetupDialogValue {
  readonly pageSize: {
    widthTwips: number;
    heightTwips: number;
    orientation: "portrait" | "landscape";
  };
  readonly pageMargins: {
    topTwips: number;
    rightTwips: number;
    bottomTwips: number;
    leftTwips: number;
    gutterTwips: number;
  };
  readonly multiplePages: WordMultiplePages | null;
  readonly gutterAtTop: boolean;
  readonly applyTo: SectionApplyTo;
  /** The Document Grid tab's grid (w:docGrid), when it was changed. */
  readonly docGrid?: { type: DocumentGridType; linePitch?: number; charSpace?: number } | null;
  /** The Layout tab, when anything on it was changed. */
  readonly layoutTab?: PageSetupLayoutTabValue;
}

/**
 * Word's Page Setup > Layout tab (Microsoft Support, "Change the layout of
 * a section"; "Vertically align text"; "Different headers and footers on
 * odd and even pages"): Section start (the break that begins the section),
 * Headers and footers -- Different odd and even (the whole document),
 * Different first page (the section), the header's and footer's distance
 * from the edge -- and the page's vertical alignment; Line Numbers... and
 * Borders... open their own dialogs.
 */
export interface PageSetupLayoutTabValue {
  /** Absent for the first section, which no break begins. */
  readonly sectionStart?: SectionBreakType;
  readonly differentOddEven: boolean;
  readonly differentFirstPage: boolean;
  readonly headerTwips: number;
  readonly footerTwips: number;
  readonly verticalAlignment: SectionVerticalAlignment;
}

export interface PageSetupLayoutTabSource {
  /** The break that begins the selected section; null in the first one. */
  readonly sectionStart: SectionBreakType | null;
  readonly differentOddEven: boolean;
  readonly differentFirstPage: boolean;
  readonly onLineNumbers?: () => void;
  readonly onBorders?: () => void;
}

const SECTION_STARTS: readonly SectionBreakType[] = ["continuous", "nextColumn", "nextPage", "evenPage", "oddPage"];
const VERTICAL_ALIGNMENTS: readonly SectionVerticalAlignment[] = ["top", "center", "both", "bottom"];

/** Word's Document Grid choices: No grid, Specify line grid only, Specify
 *  line and character grid, Text snaps to character grid. */
export type DocumentGridType = "default" | "lines" | "linesAndChars" | "snapToChars";

/**
 * Word's character and line grid from its counts (Page Setup > Document
 * Grid; measured in Word 16.86): the character pitch is the text width over
 * the characters per line, and w:charSpace (4096ths of a point) is what it
 * adds to the default font size (floored, as Word writes it); the line pitch
 * is the text height over the lines per page, in whole twips. 30 characters
 * on a 415 pt line at an 11 pt default wrote charSpace 11605, 39 wrote -1471;
 * 40 lines on 13958 twips wrote 348.
 */
export function documentGridFromCounts(
  textWidthTwips: number,
  textHeightTwips: number,
  charactersPerLine: number,
  linesPerPage: number,
  defaultFontPt: number,
): { linePitch: number; charSpace: number } {
  const charPitchPt = textWidthTwips / 20 / Math.max(1, charactersPerLine);
  return {
    linePitch: Math.max(20, Math.floor(textHeightTwips / Math.max(1, linesPerPage))),
    charSpace: Math.floor((charPitchPt - defaultFontPt) * 4096 + 1e-6),
  };
}

export interface DocumentPageSetupDialogProps {
  readonly initialTab: "margins" | "paper" | "layout" | "grid";
  /** The Layout tab's other sources: the section's break and header
   *  options. Without it the tab is not shown. */
  readonly layoutSource?: PageSetupLayoutTabSource;
  /** The document's default (Normal) font size, the character grid's base. */
  readonly gridDefaultFontPt?: number;
  readonly layout: SectionLayout;
  readonly settings: WordDocumentSettings | null;
  readonly sectionCount: number;
  /** This point forward needs a collapsed caret (it inserts a break there). */
  readonly pointForwardAvailable: boolean;
  readonly canEdit?: boolean;
  readonly onApply: (value: PageSetupDialogValue) => boolean;
  readonly onCancel: () => void;
  readonly returnFocusTarget?: () => HTMLElement | null;
}

/** EMU per twip (914400 EMU per inch / 1440 twips per inch). */

/** A length in the field's unit, at most two decimals ("2.54", "1"). */
function formatPageLength(twips: number, unit: PresentationMeasurementUnit): string {
  return formatPresentationMeasurement(twips * EMU_PER_TWIP, unit, 2);
}

/**
 * A typed length in twips: the field's unit unless the text names its own
 * ("1in", "2.5 厘米", "20mm"); a leading minus is kept for the margins Word
 * lets go negative, and the range checks below refuse it elsewhere.
 */
function parsePageLength(draft: string, unit: PresentationMeasurementUnit): number | null {
  const trimmed = draft.trim();
  const negative = trimmed.startsWith("-");
  const emu = parsePresentationMeasurement(negative ? trimmed.slice(1) : trimmed, unit);
  if (emu === null) return null;
  const twips = Math.round(emu / EMU_PER_TWIP);
  return negative ? -twips : twips;
}

export function wordPaperSizeFor(
  widthTwips: number,
  heightTwips: number,
): WordPaperSizeId | null {
  const shortSide = Math.min(widthTwips, heightTwips);
  const longSide = Math.max(widthTwips, heightTwips);
  return (
    WORD_PAPER_SIZES.find(
      (paper) =>
        Math.abs(paper.widthTwips - shortSide) <= 15 &&
        Math.abs(paper.heightTwips - longSide) <= 15,
    )?.id ?? null
  );
}

type MarginKey = "top" | "bottom" | "left" | "right" | "gutter";

export function DocumentPageSetupDialog({
  initialTab,
  layout,
  settings,
  sectionCount,
  pointForwardAvailable,
  canEdit = true,
  onApply,
  onCancel,
  returnFocusTarget,
  gridDefaultFontPt = 10.5,
  layoutSource,
}: DocumentPageSetupDialogProps) {
  const t = useTranslations("console-notebooks");
  const lengthUnit = presentationMeasurementUnit(useLocale());
  const shownLength = (twips: number) => formatPageLength(twips, lengthUnit);
  const parseLength = (draft: string) => parsePageLength(draft, lengthUnit);
  const firstFieldRef = useRef<HTMLInputElement | null>(null);
  const [tab, setTab] = useState<"margins" | "paper" | "layout" | "grid">(initialTab);
  const sourceGrid = layout.docGrid;
  const [gridType, setGridType] = useState<DocumentGridType>(
    (sourceGrid?.type as DocumentGridType | undefined) ?? "default",
  );
  const [gridTouched, setGridTouched] = useState(false);
  const [layoutTouched, setLayoutTouched] = useState(false);
  const [sectionStart, setSectionStart] = useState<SectionBreakType | null>(layoutSource?.sectionStart ?? null);
  const [oddEven, setOddEven] = useState(layoutSource?.differentOddEven === true);
  const [firstPage, setFirstPage] = useState(layoutSource?.differentFirstPage === true);
  const [verticalAlignment, setVerticalAlignment] = useState<SectionVerticalAlignment>(
    layout.verticalAlignment ?? "top",
  );
  const [headerDraft, setHeaderDraft] = useState(() => formatPageLength(layout.pageMargins.headerTwips, lengthUnit));
  const [footerDraft, setFooterDraft] = useState(() => formatPageLength(layout.pageMargins.footerTwips, lengthUnit));
  const touchLayout = () => setLayoutTouched(true);
  const [margins, setMargins] = useState<Record<MarginKey, string>>(() => ({
    top: shownLength(layout.pageMargins.topTwips),
    bottom: shownLength(layout.pageMargins.bottomTwips),
    left: shownLength(layout.pageMargins.leftTwips),
    right: shownLength(layout.pageMargins.rightTwips),
    gutter: shownLength(layout.pageMargins.gutterTwips),
  }));
  const [gutterAtTop, setGutterAtTop] = useState(settings?.gutterAtTop === true);
  const [orientation, setOrientation] = useState(layout.pageSize.orientation);
  const [multiplePages, setMultiplePages] = useState<WordMultiplePages | null>(
    settings?.multiplePages ?? null,
  );
  const [width, setWidth] = useState(shownLength(layout.pageSize.widthTwips));
  const [height, setHeight] = useState(shownLength(layout.pageSize.heightTwips));
  // The exact twips behind each shown value: a field the user has not
  // changed keeps its source value rather than the 0.01 in rounding shown
  // (A5 is 8391 twips; 5.83 in would write 8395).
  const [exact, setExact] = useState<Record<string, { draft: string; twips: number }>>(
    () => ({
      top: { draft: shownLength(layout.pageMargins.topTwips), twips: layout.pageMargins.topTwips },
      bottom: { draft: shownLength(layout.pageMargins.bottomTwips), twips: layout.pageMargins.bottomTwips },
      left: { draft: shownLength(layout.pageMargins.leftTwips), twips: layout.pageMargins.leftTwips },
      right: { draft: shownLength(layout.pageMargins.rightTwips), twips: layout.pageMargins.rightTwips },
      gutter: { draft: shownLength(layout.pageMargins.gutterTwips), twips: layout.pageMargins.gutterTwips },
      width: { draft: shownLength(layout.pageSize.widthTwips), twips: layout.pageSize.widthTwips },
      header: { draft: shownLength(layout.pageMargins.headerTwips), twips: layout.pageMargins.headerTwips },
      footer: { draft: shownLength(layout.pageMargins.footerTwips), twips: layout.pageMargins.footerTwips },
      height: { draft: shownLength(layout.pageSize.heightTwips), twips: layout.pageSize.heightTwips },
    }),
  );
  const twipsOf = (key: string, draft: string): number | null =>
    exact[key]?.draft === draft ? exact[key].twips : parseLength(draft);
  const applyToOptions = sectionApplyToOptions(sectionCount);
  const [applyTo, setApplyTo] = useState<SectionApplyTo>(() =>
    defaultSectionApplyTo(sectionCount),
  );

  const mirrored = multiplePages === "mirrorMargins" || multiplePages === "bookFold";
  const twoOnOne = multiplePages === "printTwoOnOne";
  const foldAcrossHeight = twoOnOne && orientation === "portrait";
  const label = (key: MarginKey): string => {
    if (key === "gutter") return t("toolbar.pageSetupGutter");
    if ((mirrored || (twoOnOne && !foldAcrossHeight)) && key === "left") {
      return t("toolbar.pageSetupMarginInside");
    }
    if ((mirrored || (twoOnOne && !foldAcrossHeight)) && key === "right") {
      return t("toolbar.pageSetupMarginOutside");
    }
    if (foldAcrossHeight && key === "top") return t("toolbar.pageSetupMarginOutside");
    if (foldAcrossHeight && key === "bottom") return t("toolbar.pageSetupMarginInside");
    return t(
      key === "top"
        ? "toolbar.pageSetupMarginTop"
        : key === "bottom"
          ? "toolbar.pageSetupMarginBottom"
          : key === "left"
            ? "toolbar.pageSetupMarginLeft"
            : "toolbar.pageSetupMarginRight",
    );
  };

  const candidate = useMemo((): PageSetupDialogValue | "invalid" | "tooLarge" => {
    const top = twipsOf("top", margins.top);
    const bottom = twipsOf("bottom", margins.bottom);
    const left = twipsOf("left", margins.left);
    const right = twipsOf("right", margins.right);
    const gutter = twipsOf("gutter", margins.gutter);
    const widthTwips = twipsOf("width", width);
    const heightTwips = twipsOf("height", height);
    if (
      [top, bottom, left, right, gutter].some(
        (value) => value === null || value < 0 || value > MAX_PAPER_TWIPS,
      ) ||
      widthTwips === null ||
      heightTwips === null ||
      widthTwips < MIN_PAPER_TWIPS ||
      heightTwips < MIN_PAPER_TWIPS ||
      widthTwips > MAX_PAPER_TWIPS ||
      heightTwips > MAX_PAPER_TWIPS
    ) {
      return "invalid";
    }
    const value: PageSetupDialogValue = {
      pageSize: { widthTwips, heightTwips, orientation },
      pageMargins: {
        topTwips: top!,
        bottomTwips: bottom!,
        leftTwips: left!,
        rightTwips: right!,
        gutterTwips: gutter!,
      },
      multiplePages,
      gutterAtTop: multiplePages === null && gutterAtTop,
      applyTo,
    };
    // The text area of the page as it will be drawn, on both a right-hand
    // and a left-hand page, must keep room for text.
    const nextSettings: WordDocumentSettings = {
      ...(multiplePages ? { multiplePages } : {}),
      ...(value.gutterAtTop ? { gutterAtTop: true as const } : {}),
    };
    for (const pageIndex of [0, 1]) {
      const box = resolveWordPageBox(
        { pageSize: value.pageSize, pageMargins: value.pageMargins },
        nextSettings,
        pageIndex,
      );
      if (
        box.widthTwips - box.leftTwips - box.rightTwips < MIN_TEXT_TWIPS ||
        box.heightTwips - box.topTwips - box.bottomTwips < MIN_TEXT_TWIPS
      ) {
        return "tooLarge";
      }
    }
    return value;
    // twipsOf reads `exact`, which only changes alongside the drafts.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [applyTo, exact, gutterAtTop, height, margins, multiplePages, orientation, width]);
  // The text area the grid divides, from the page as it now reads.
  const textBox = (() => {
    const source = typeof candidate === "object" ? candidate : layout;
    const margins = source.pageMargins;
    const gutter = "gutterAtTop" in source && source.gutterAtTop ? 0 : margins.gutterTwips;
    return {
      width: Math.max(MIN_TEXT_TWIPS, source.pageSize.widthTwips - margins.leftTwips - margins.rightTwips - gutter),
      height: Math.max(MIN_TEXT_TWIPS, source.pageSize.heightTwips - margins.topTwips - margins.bottomTwips),
    };
  })();
  const [charsDraft, setCharsDraft] = useState(() => {
    const pitch = gridDefaultFontPt + (sourceGrid?.charSpace ?? 0) / 4096;
    return String(Math.max(1, Math.floor(textBox.width / 20 / pitch + 1e-6)));
  });
  const [linesDraft, setLinesDraft] = useState(() =>
    String(Math.max(1, Math.floor(textBox.height / (sourceGrid?.linePitch ?? 312) + 1e-6))),
  );
  const charsPerLine = Number.parseInt(charsDraft, 10);
  const linesPerPage = Number.parseInt(linesDraft, 10);
  const gridCountsValid =
    Number.isFinite(charsPerLine) && charsPerLine >= 1 && charsPerLine <= 200 &&
    Number.isFinite(linesPerPage) && linesPerPage >= 1 && linesPerPage <= 300;
  const gridValue = gridCountsValid
    ? documentGridFromCounts(textBox.width, textBox.height, charsPerLine, linesPerPage, gridDefaultFontPt)
    : null;
  const headerTwips = twipsOf("header", headerDraft);
  const footerTwips = twipsOf("footer", footerDraft);
  const layoutValid =
    !layoutTouched ||
    (headerTwips !== null && footerTwips !== null && headerTwips >= 0 && footerTwips >= 0 &&
      headerTwips <= MAX_PAPER_TWIPS && footerTwips <= MAX_PAPER_TWIPS);
  const valid =
    typeof candidate === "object" && layoutValid && (!gridTouched || gridType === "default" || gridValue !== null);
  const withLayout = (value: PageSetupDialogValue): PageSetupDialogValue =>
    !layoutTouched || headerTwips === null || footerTwips === null
      ? value
      : {
          ...value,
          layoutTab: {
            ...(sectionStart ? { sectionStart } : {}),
            differentOddEven: oddEven,
            differentFirstPage: firstPage,
            headerTwips,
            footerTwips,
            verticalAlignment,
          },
        };
  const applied = (value: PageSetupDialogValue): PageSetupDialogValue =>
    !gridTouched
      ? withLayout(value)
      : withLayout({
          ...value,
          docGrid:
            gridType === "default"
              ? { type: "default", ...(sourceGrid?.linePitch ? { linePitch: sourceGrid.linePitch } : {}) }
              : {
                  type: gridType,
                  linePitch: gridValue!.linePitch,
                  ...(gridType === "lines" ? {} : { charSpace: gridValue!.charSpace }),
                },
        });

  const paperId = (() => {
    const widthTwips = twipsOf("width", width);
    const heightTwips = twipsOf("height", height);
    return widthTwips !== null && heightTwips !== null
      ? wordPaperSizeFor(widthTwips, heightTwips)
      : null;
  })();

  const changeOrientation = (next: "portrait" | "landscape") => {
    if (next === orientation) return;
    setOrientation(next);
    // Word turns the sheet: the width and height trade places.
    setWidth(height);
    setHeight(width);
    setExact((current) => ({
      ...current,
      ...(current.height ? { width: current.height } : {}),
      ...(current.width ? { height: current.width } : {}),
    }));
  };

  const changeMultiplePages = (next: WordMultiplePages | null) => {
    setMultiplePages(next);
    // Word prints a book fold on a landscape sheet and turns it for you.
    if (next === "bookFold") changeOrientation("landscape");
  };

  const choosePaper = (id: string) => {
    const paper = WORD_PAPER_SIZES.find((entry) => entry.id === id);
    if (!paper) return;
    const [nextWidth, nextHeight] =
      orientation === "landscape"
        ? [paper.heightTwips, paper.widthTwips]
        : [paper.widthTwips, paper.heightTwips];
    setWidth(shownLength(nextWidth));
    setHeight(shownLength(nextHeight));
    setExact((current) => ({
      ...current,
      width: { draft: shownLength(nextWidth), twips: nextWidth },
      height: { draft: shownLength(nextHeight), twips: nextHeight },
    }));
  };

  const unit = t(
    lengthUnit === "in" ? "toolbar.pageSetupUnitInches" : "toolbar.indentUnitCm",
  );
  const valueField = (
    key: string,
    caption: string,
    value: string,
    onChange: (value: string) => void,
    testId: string,
    options: { disabled?: boolean; focus?: boolean } = {},
  ) => (
    <label className="office-dialog__field" key={key}>
      <span className="office-dialog__label">{caption}</span>
      <span className={styles.valueRow}>
        <input
          ref={options.focus ? firstFieldRef : undefined}
          type="text"
          inputMode="decimal"
          className={documentValueInputClassName("note-editor-office-toolbar__input")}
          data-testid={testId}
          aria-label={caption}
          value={value}
          disabled={!canEdit || options.disabled}
          onChange={(event) => onChange(event.currentTarget.value)}
        />
        <span className={styles.unit}>{unit}</span>
      </span>
    </label>
  );

  const applyToField = (panel: "margins" | "paper" | "layout" | "grid") => (
    <label className="office-dialog__row">
      <span className="office-dialog__label">{t("toolbar.pageSetupApplyTo")}</span>
      <select
        data-testid={`document-page-setup-apply-to-${panel}`}
        value={applyTo}
        disabled={!canEdit}
        onChange={(event) => setApplyTo(event.currentTarget.value as SectionApplyTo)}
      >
        {applyToOptions.map((option) => (
          <option
            key={option}
            value={option}
            disabled={option === "thisPointForward" && !pointForwardAvailable}
          >
            {t(`toolbar.sectionApplyTo_${option}` as "toolbar.sectionApplyTo_wholeDocument")}
          </option>
        ))}
      </select>
    </label>
  );

  // A small page drawing, as Word's Preview: the text area inside the
  // margins, with the gutter hatched on its binding side.
  const preview = (() => {
    if (!valid) return null;
    const value = candidate as PageSetupDialogValue;
    const nextSettings: WordDocumentSettings = {
      ...(multiplePages ? { multiplePages } : {}),
      ...(value.gutterAtTop ? { gutterAtTop: true as const } : {}),
    };
    const pages = multiplePages ? [0, 1] : [0];
    return (
      <div className={`office-dialog__well ${styles.previewWell}`} aria-hidden>
        {pages.map((pageIndex) => {
          const box = resolveWordPageBox(
            { pageSize: value.pageSize, pageMargins: value.pageMargins },
            nextSettings,
            pageIndex,
          );
          const scale = 80 / Math.max(box.widthTwips, box.heightTwips);
          const w = box.widthTwips * scale;
          const h = box.heightTwips * scale;
          return (
            <svg
              key={pageIndex}
              viewBox={`-1 -1 ${w + 2} ${h + 2}`}
              data-testid={`document-page-setup-preview-${pageIndex}`}
            >
              <rect x={0} y={0} width={w} height={h} fill="#fff" stroke="#9ca3af" />
              <rect
                x={box.leftTwips * scale}
                y={box.topTwips * scale}
                width={Math.max(0, w - (box.leftTwips + box.rightTwips) * scale)}
                height={Math.max(0, h - (box.topTwips + box.bottomTwips) * scale)}
                fill="#e5e7eb"
                stroke="#6b7280"
                strokeDasharray="2 2"
              />
            </svg>
          );
        })}
      </div>
    );
  })();

  return (
    <OfficeDialogShell
      open
      testIdPrefix={initialTab === "paper" ? "document-page-size" : "document-page-margins"}
      title={t("toolbar.pageSetupDialogTitle")}
      description=""
      defaultActionLabel={t("toolbar.dialogOk")}
      defaultActionTestId={
        initialTab === "paper" ? "document-page-size-apply" : "document-page-margins-apply"
      }
      defaultActionDisabled={!canEdit || !valid}
      cancelActionLabel={t("toolbar.pageSetupCancel")}
      onDefaultAction={() => {
        if (!canEdit || typeof candidate !== "object" || !valid) return;
        onApply(applied(candidate));
      }}
      onCancel={onCancel}
      initialFocusRef={firstFieldRef}
      returnFocusTarget={returnFocusTarget}
      className="sm:max-w-[520px]"
      tabStrip={
        <div className="office-dialog__tabs" role="tablist" aria-label={t("toolbar.pageSetupDialogTitle")}>
          {(layoutSource ? (["margins", "paper", "layout", "grid"] as const) : (["margins", "paper", "grid"] as const)).map((value) => (
            <button
              key={value}
              type="button"
              role="tab"
              id={`document-page-setup-tab-${value}`}
              aria-selected={tab === value}
              aria-controls={`document-page-setup-page-${value}`}
              className="office-dialog__tab"
              data-testid={`document-page-setup-tab-${value}`}
              onClick={() => setTab(value)}
            >
              {t(
                value === "margins"
                  ? "toolbar.pageSetupTabMargins"
                  : value === "paper"
                    ? "toolbar.pageSetupTabPaper"
                    : value === "layout"
                      ? "toolbar.pageSetupTabLayout"
                      : "toolbar.pageSetupTabGrid",
              )}
            </button>
          ))}
        </div>
      }
    >
      <div
        id="document-page-setup-page-margins"
        role="tabpanel"
        aria-labelledby="document-page-setup-tab-margins"
        hidden={tab !== "margins"}
        className="office-dialog__body"
      >
        <section className="office-dialog__group">
          <h3 className="office-dialog__section-title">{t("toolbar.pageSetupMarginsGroup")}</h3>
          <div className={styles.twoColumns}>
            {(["top", "bottom", "left", "right"] as const).map((key, index) =>
              valueField(
                key,
                label(key),
                margins[key],
                (value) => setMargins((current) => ({ ...current, [key]: value })),
                `document-page-margins-${key}`,
                { focus: initialTab === "margins" && index === 0 },
              ),
            )}
            {valueField(
              "gutter",
              label("gutter"),
              margins.gutter,
              (value) => setMargins((current) => ({ ...current, gutter: value })),
              "document-page-margins-gutter",
            )}
            <label className="office-dialog__field">
              <span className="office-dialog__label">{t("toolbar.pageSetupGutterPosition")}</span>
              <select
                data-testid="document-page-margins-gutter-position"
                value={multiplePages === null && gutterAtTop ? "top" : "left"}
                disabled={!canEdit || multiplePages !== null}
                onChange={(event) => setGutterAtTop(event.currentTarget.value === "top")}
              >
                <option value="left">{t("toolbar.pageSetupGutterLeft")}</option>
                <option value="top">{t("toolbar.pageSetupGutterTop")}</option>
              </select>
            </label>
          </div>
        </section>
        <fieldset className="office-dialog__group" disabled={!canEdit}>
          <legend>{t("toolbar.pageSetupOrientation")}</legend>
          <div className={styles.choiceRow} role="radiogroup">
            {(["portrait", "landscape"] as const).map((value) => (
              <label className="office-dialog__check" key={value}>
                <input
                  type="radio"
                  name="document-page-setup-orientation"
                  value={value}
                  checked={orientation === value}
                  data-testid={`document-page-margins-orientation-${value}`}
                  onChange={() => changeOrientation(value)}
                />
                <OfficeSemanticIcon
                  semanticId={`page.orientation.${value}`}
                  size={16}
                  className="office-dialog__icon"
                />
                {t(value === "portrait" ? "toolbar.pageSetupPortrait" : "toolbar.pageSetupLandscape")}
              </label>
            ))}
          </div>
        </fieldset>
        <section className="office-dialog__group">
          <h3 className="office-dialog__section-title">{t("toolbar.pageSetupPages")}</h3>
          <label className="office-dialog__row">
            <span className="office-dialog__label">{t("toolbar.pageSetupMultiplePages")}</span>
            <select
              data-testid="document-page-margins-multiple-pages"
              value={multiplePages ?? "normal"}
              disabled={!canEdit}
              onChange={(event) => {
                const value = event.currentTarget.value;
                changeMultiplePages(value === "normal" ? null : (value as WordMultiplePages));
              }}
            >
              <option value="normal">{t("toolbar.pageSetupPagesNormal")}</option>
              <option value="mirrorMargins">{t("toolbar.pageSetupPagesMirror")}</option>
              <option value="printTwoOnOne">{t("toolbar.pageSetupPagesTwoOnOne")}</option>
              <option value="bookFold">{t("toolbar.pageSetupPagesBookFold")}</option>
            </select>
          </label>
        </section>
        <section className="office-dialog__group">
          <h3 className="office-dialog__section-title">{t("toolbar.pageSetupPreview")}</h3>
          {preview}
          {applyToField("margins")}
        </section>
      </div>
      <div
        id="document-page-setup-page-paper"
        role="tabpanel"
        aria-labelledby="document-page-setup-tab-paper"
        hidden={tab !== "paper"}
        className="office-dialog__body"
      >
        <label className="office-dialog__field">
          <span className="office-dialog__label">{t("toolbar.pageSetupPaperSize")}</span>
          <select
            data-testid="document-page-size-paper"
            value={paperId ?? "custom"}
            disabled={!canEdit}
            onChange={(event) => choosePaper(event.currentTarget.value)}
          >
            {WORD_PAPER_SIZES.map((paper) => (
              <option key={paper.id} value={paper.id}>
                {t(`toolbar.pageSize_${paper.id}` as "toolbar.pageSize_letter")}
              </option>
            ))}
            <option value="custom" disabled>
              {t("toolbar.pageSetupCustomPaper")}
            </option>
          </select>
        </label>
        <div className={styles.twoColumns}>
          {valueField("width", t("toolbar.pageSetupWidth"), width, setWidth, "document-page-size-width", {
            focus: initialTab === "paper",
          })}
          {valueField("height", t("toolbar.pageSetupHeight"), height, setHeight, "document-page-size-height")}
        </div>
        {preview}
        {applyToField("paper")}
      </div>
      {layoutSource ? (
        <div
          id="document-page-setup-page-layout"
          role="tabpanel"
          aria-labelledby="document-page-setup-tab-layout"
          hidden={tab !== "layout"}
          className="office-dialog__body"
        >
          <section className="office-dialog__group">
            <h3 className="office-dialog__section-title">{t("toolbar.pageSetupLayoutSection")}</h3>
            <label className="office-dialog__row">
              <span className="office-dialog__label">{t("toolbar.pageSetupSectionStart")}</span>
              <select
                data-testid="document-page-setup-section-start"
                value={sectionStart ?? "nextPage"}
                disabled={!canEdit || sectionStart === null}
                onChange={(event) => {
                  setSectionStart(event.currentTarget.value as SectionBreakType);
                  touchLayout();
                }}
              >
                {SECTION_STARTS.map((value) => (
                  <option key={value} value={value}>
                    {t(`toolbar.pageSetupSectionStart_${value}` as "toolbar.pageSetupSectionStart_nextPage")}
                  </option>
                ))}
              </select>
            </label>
          </section>
          <section className="office-dialog__group">
            <h3 className="office-dialog__section-title">{t("toolbar.pageSetupHeadersFooters")}</h3>
            <label className="office-dialog__check">
              <input
                type="checkbox"
                checked={oddEven}
                disabled={!canEdit}
                data-testid="document-page-setup-odd-even"
                onChange={(event) => {
                  setOddEven(event.currentTarget.checked);
                  touchLayout();
                }}
              />
              {t("toolbar.pageSetupDifferentOddEven")}
            </label>
            <label className="office-dialog__check">
              <input
                type="checkbox"
                checked={firstPage}
                disabled={!canEdit}
                data-testid="document-page-setup-first-page"
                onChange={(event) => {
                  setFirstPage(event.currentTarget.checked);
                  touchLayout();
                }}
              />
              {t("toolbar.pageSetupDifferentFirstPage")}
            </label>
            <div className={styles.twoColumns}>
              {valueField("header", t("toolbar.pageSetupFromEdgeHeader"), headerDraft, (next) => {
                setHeaderDraft(next);
                touchLayout();
              }, "document-page-setup-header-distance")}
              {valueField("footer", t("toolbar.pageSetupFromEdgeFooter"), footerDraft, (next) => {
                setFooterDraft(next);
                touchLayout();
              }, "document-page-setup-footer-distance")}
            </div>
          </section>
          <section className="office-dialog__group">
            <h3 className="office-dialog__section-title">{t("toolbar.pageSetupLayoutPage")}</h3>
            <label className="office-dialog__row">
              <span className="office-dialog__label">{t("toolbar.pageSetupVerticalAlignment")}</span>
              <select
                data-testid="document-page-setup-vertical-alignment"
                value={verticalAlignment}
                disabled={!canEdit}
                onChange={(event) => {
                  setVerticalAlignment(event.currentTarget.value as SectionVerticalAlignment);
                  touchLayout();
                }}
              >
                {VERTICAL_ALIGNMENTS.map((value) => (
                  <option key={value} value={value}>
                    {t(`toolbar.pageSetupVerticalAlignment_${value}` as "toolbar.pageSetupVerticalAlignment_top")}
                  </option>
                ))}
              </select>
            </label>
          </section>
          <div className="office-dialog__row">
            {layoutSource.onLineNumbers ? (
              <button
                type="button"
                className="office-dialog__button"
                data-testid="document-page-setup-line-numbers"
                disabled={!canEdit}
                onClick={layoutSource.onLineNumbers}
              >
                {t("toolbar.pageSetupLineNumbersButton")}
              </button>
            ) : null}
            {layoutSource.onBorders ? (
              <button
                type="button"
                className="office-dialog__button"
                data-testid="document-page-setup-borders"
                disabled={!canEdit}
                onClick={layoutSource.onBorders}
              >
                {t("toolbar.pageSetupBordersButton")}
              </button>
            ) : null}
          </div>
          {applyToField("layout")}
        </div>
      ) : null}
      <div
        id="document-page-setup-page-grid"
        role="tabpanel"
        aria-labelledby="document-page-setup-tab-grid"
        hidden={tab !== "grid"}
        className="office-dialog__body"
      >
        <fieldset className="office-dialog__group" role="radiogroup">
          <legend className="office-dialog__section-title">{t("toolbar.pageSetupGrid")}</legend>
          {(["default", "lines", "linesAndChars", "snapToChars"] as const).map((type) => (
            <label key={type} className="office-dialog__check">
              <input
                type="radio"
                name="document-page-setup-grid"
                data-testid={`document-page-setup-grid-${type}`}
                checked={gridType === type}
                disabled={!canEdit}
                onChange={() => {
                  setGridType(type);
                  setGridTouched(true);
                }}
              />
              {t(`toolbar.pageSetupGrid_${type}` as "toolbar.pageSetupGrid_default")}
            </label>
          ))}
        </fieldset>
        <div className={styles.twoColumns}>
          <label className="office-dialog__field">
            <span className="office-dialog__label">{t("toolbar.pageSetupGridCharsPerLine")}</span>
            <input
              type="number"
              min={1}
              max={200}
              data-testid="document-page-setup-grid-chars"
              value={charsDraft}
              disabled={!canEdit || (gridType !== "linesAndChars" && gridType !== "snapToChars")}
              onChange={(event) => {
                setCharsDraft(event.currentTarget.value);
                setGridTouched(true);
              }}
            />
            <span className="office-dialog__hint" data-testid="document-page-setup-grid-char-pitch">
              {gridValue
                ? t("toolbar.pageSetupGridPitch", {
                    pitch: (gridDefaultFontPt + gridValue.charSpace / 4096).toFixed(2),
                  })
                : null}
            </span>
          </label>
          <label className="office-dialog__field">
            <span className="office-dialog__label">{t("toolbar.pageSetupGridLinesPerPage")}</span>
            <input
              type="number"
              min={1}
              max={300}
              data-testid="document-page-setup-grid-lines"
              value={linesDraft}
              disabled={!canEdit || gridType === "default"}
              onChange={(event) => {
                setLinesDraft(event.currentTarget.value);
                setGridTouched(true);
              }}
            />
            <span className="office-dialog__hint" data-testid="document-page-setup-grid-line-pitch">
              {gridValue
                ? t("toolbar.pageSetupGridPitch", { pitch: (gridValue.linePitch / 20).toFixed(2) })
                : null}
            </span>
          </label>
        </div>
        {applyToField("grid")}
      </div>
      {candidate === "invalid" || candidate === "tooLarge" ? (
        <p className="office-dialog__message" role="alert" data-testid="document-page-setup-error">
          {t(
            candidate === "invalid"
              ? "toolbar.pageSetupInvalidValue"
              : "toolbar.pageSetupMarginsTooLarge",
          )}
        </p>
      ) : null}
    </OfficeDialogShell>
  );
}
