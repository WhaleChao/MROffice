"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * PowerPoint's Font dialog on a slide's text (Home > Font launcher, and
 * Character Spacing > More Spacing... which opens it on its Character Spacing
 * page). The Font page is the shared Office page Docs and a chart's Font
 * dialog draw (`DocumentFontDialogFontPage`); the Character Spacing page is
 * PowerPoint's Spacing / By pair. What OK writes is decided by
 * `lib/presentation-font-dialog`; the host commits it the way the Home
 * buttons do.
 */

import { useRef, useState } from "react";
import { useTranslations } from "@mrnote/core/host/next-intl.tsx";

import DocumentFontDialogFontPage, {
  type DocumentFontPageDraft,
  type FontDialogStyle,
  type FontDialogUnderline,
} from "@mrnote/core/web/components/console/editor/native-document/DocumentFontDialogFontPage.tsx";
import { OfficeDialogShell } from "@mrnote/core/web/components/office-ribbon/OfficeDialogShell.tsx";
import { officeFontSizeOptions } from "@mrnote/core/web/components/office-ribbon/officeFontSizeCatalog.ts";
import {
  OFFICE_FONT_FAMILY_VALUES,
  officeFontFamilyLabels,
} from "@mrnote/core/web/components/office-ribbon/officeTypographyCatalog.ts";
import { ribbonFontFamilyLabel } from "@mrnote/core/web/components/office-ribbon/RibbonFontFamilyPicker.tsx";
import {
  changePresentationFontDialogDraft,
  PRESENTATION_MAX_CHARACTER_SPACING_PT,
  presentationFontDialogDraft,
  presentationFontDialogResult,
  type PresentationFontDialogDraft,
  type PresentationFontDialogField,
  type PresentationFontDialogProblem,
  type PresentationFontDialogResult,
  type PresentationFontDialogSource,
  type PresentationFontDialogSpacingKind,
} from "../presentation-font-dialog";

export type PresentationFontDialogTab = "font" | "spacing";

export interface PresentationFontDialogProps {
  readonly open: boolean;
  readonly initialTab: PresentationFontDialogTab;
  readonly source: PresentationFontDialogSource;
  /** False where the target keeps no box-wide spacing (a table cell, a shape group). */
  readonly spacingWritable: boolean;
  /** The inherited (theme) face's option label and CSS stack. */
  readonly inheritedFamilyLabel: string;
  readonly inheritedFamilyCss: string;
  readonly automaticColor?: string;
  readonly previewText?: string;
  /** Commits OK's result; false keeps the dialog open. */
  readonly onApply: (result: PresentationFontDialogResult) => boolean;
  readonly onCancel: () => void;
}

export function PresentationFontDialog(props: PresentationFontDialogProps): React.JSX.Element | null {
  if (!props.open) return null;
  return <PresentationFontDialogBody {...props} />;
}

const TEST_ID = "presentation-font-dialog";

function PresentationFontDialogBody({
  initialTab,
  source,
  spacingWritable,
  inheritedFamilyLabel,
  inheritedFamilyCss,
  automaticColor = "#000000",
  previewText = "",
  onApply,
  onCancel,
}: PresentationFontDialogProps): React.JSX.Element {
  const t = useTranslations("console-notebooks");
  const word = (key: string, values?: Record<string, string | number>) =>
    t(key as Parameters<typeof t>[0], values as never);
  const [draft, setDraft] = useState<PresentationFontDialogDraft>(() => presentationFontDialogDraft(source));
  const draftRef = useRef(draft);
  const [touched, setTouched] = useState<ReadonlySet<PresentationFontDialogField>>(() => new Set());
  const [tab, setTab] = useState<PresentationFontDialogTab>(spacingWritable ? initialTab : "font");
  const [problem, setProblem] = useState<PresentationFontDialogProblem | null>(null);
  const initialFocusRef = useRef<HTMLInputElement | null>(null);
  const labels = officeFontFamilyLabels(word);
  const unavailable = word("presentation.fontDialog.optionUnavailable");

  const change = <K extends PresentationFontDialogField>(field: K, value: PresentationFontDialogDraft[K]) => {
    setProblem(null);
    const result = changePresentationFontDialogDraft(draftRef.current, field, value);
    draftRef.current = result.draft;
    setDraft(result.draft);
    setTouched((previous) => new Set([...previous, ...result.touched]));
  };

  const familyOptions: { value: string; label: string }[] = OFFICE_FONT_FAMILY_VALUES.map((stack) => ({
    value: stack,
    label: ribbonFontFamilyLabel(stack, labels),
  }));
  if (draft.fontFamily && !(OFFICE_FONT_FAMILY_VALUES as readonly string[]).includes(draft.fontFamily)) {
    familyOptions.unshift({ value: draft.fontFamily, label: ribbonFontFamilyLabel(draft.fontFamily, labels) });
  }

  const pageDraft: DocumentFontPageDraft = {
    family: draft.fontFamily === null ? "" : draft.fontFamily || null,
    cjkFamily: null,
    style: (draft.bold && draft.italic
      ? "boldItalic"
      : draft.bold
        ? "bold"
        : draft.italic
          ? "italic"
          : "regular") as FontDialogStyle,
    sizePt: draft.sizePt,
    color: draft.color === null ? "" : draft.color || null,
    underline: (draft.underline ?? "") as FontDialogUnderline | "",
    underlineColor: null,
    superscript: draft.superscript,
    subscript: draft.subscript,
  };

  const problemMessage = problem
    ? problem === "size"
      ? word("office.chartFontDialog.invalidSize")
      : word("presentation.fontDialog.invalidSpacing", { maximum: PRESENTATION_MAX_CHARACTER_SPACING_PT })
    : null;

  const apply = () => {
    const result = presentationFontDialogResult(draft, touched);
    if ("problem" in result) {
      setProblem(result.problem);
      setTab(result.problem === "spacing" ? "spacing" : "font");
      return;
    }
    if (touched.size === 0 || onApply(result)) onCancel();
  };

  const tabs = (
    [
      ["font", word("toolbar.fontTabFont")],
      ["spacing", word("toolbar.fontSectionCharacterSpacing")],
    ] as const
  ).filter(([id]) => id === "font" || spacingWritable);

  return (
    <OfficeDialogShell
      open
      testIdPrefix={TEST_ID}
      title={word("toolbar.fontDialogTitle")}
      description=""
      defaultActionLabel={word("toolbar.dialogOk")}
      defaultActionTestId={`${TEST_ID}-apply`}
      cancelActionLabel={word("toolbar.pageSetupCancel")}
      cancelActionTestId={`${TEST_ID}-cancel`}
      overlayTestId={`${TEST_ID}-layer`}
      onDefaultAction={apply}
      onCancel={onCancel}
      initialFocusRef={initialFocusRef}
      className="note-editor-font-dialog sm:max-w-[680px]"
      bodyClassName="note-editor-dialog-scroll"
      footerStart={
        problemMessage ? (
          <span role="alert" className="office-dialog__error" data-testid={`${TEST_ID}-problem`}>
            {problemMessage}
          </span>
        ) : undefined
      }
      tabStrip={
        <div className="office-dialog__tabs" role="tablist" aria-label={word("toolbar.fontDialogTitle")}>
          {tabs.map(([id, label]) => (
            <button
              key={id}
              id={`${TEST_ID}-tab-${id}`}
              type="button"
              role="tab"
              aria-selected={tab === id}
              aria-controls={`${TEST_ID}-page-${id}`}
              className="office-dialog__tab"
              data-testid={`${TEST_ID}-tab-${id}`}
              onClick={() => setTab(id)}
            >
              {label}
            </button>
          ))}
        </div>
      }
    >
      <div
        id={`${TEST_ID}-page-font`}
        className="note-editor-font-dialog__page"
        role="tabpanel"
        aria-labelledby={`${TEST_ID}-tab-font`}
        hidden={tab !== "font"}
      >
        <DocumentFontDialogFontPage
          draft={pageDraft}
          onDraftChange={(field, value) => {
            switch (field) {
              case "family":
                change("fontFamily", (value as string | null) || "");
                break;
              case "style": {
                const style = value as FontDialogStyle;
                change("bold", style === "bold" || style === "boldItalic");
                change("italic", style === "italic" || style === "boldItalic");
                break;
              }
              case "sizePt":
                change("sizePt", String(value ?? ""));
                break;
              case "color":
                change("color", (value as string | null) || "");
                break;
              case "underline":
                change("underline", (value === "double" ? "double" : value && value !== "none" ? "single" : "none"));
                break;
              case "superscript":
              case "subscript":
                change(field, value === true);
                break;
            }
          }}
          effects={{
            strike: draft.strike,
            doubleStrike: false,
            smallCaps: draft.caps === "small",
            allCaps: draft.caps === "all",
            hidden: false,
          }}
          onEffectChange={(effect, checked) => {
            if (effect === "strike") change("strike", checked);
            if (effect === "smallCaps") change("caps", checked ? "small" : "none");
            if (effect === "allCaps") change("caps", checked ? "all" : "none");
          }}
          drawingEffects={{
            offset: "",
            offsetLabel: word("office.chartFontDialog.offset"),
            onOffsetChange: () => {},
            equalizeHeight: false,
            equalizeHeightLabel: word("office.chartFontDialog.equalizeHeight"),
            onEqualizeHeightChange: () => {},
          }}
          unsupported={{
            reason: unavailable,
            asianFamily: true,
            underlineColor: true,
            underlineStyles: ["none", "single", "double"],
            effectNames: ["doubleStrike", "offset", "equalizeHeight"],
          }}
          latinFamilies={familyOptions}
          asianFamilies={[]}
          documentDefaultFamily={inheritedFamilyLabel}
          documentDefaultFamilyCss={inheritedFamilyCss}
          sizeOptions={officeFontSizeOptions("powerpoint")}
          normalizeSizePt={(raw) => {
            const value = Number(raw.trim().replace(/pt$/i, ""));
            return Number.isFinite(value) && value >= 1 && value <= 400 ? String(Math.round(value * 10) / 10) : null;
          }}
          previewText={previewText}
          automaticColor={automaticColor}
          initialFocusRef={tab === "font" ? initialFocusRef : undefined}
          labels={{
            sectionFont: word("toolbar.fontDialogSectionFont"),
            latinFamily: word("toolbar.fontDialogLatinFamily"),
            asianFamily: word("toolbar.fontDialogAsianFamily"),
            style: word("toolbar.fontDialogStyle"),
            size: word("toolbar.fontSize"),
            sectionColor: word("toolbar.fontDialogSectionColor"),
            color: word("toolbar.textColor"),
            underlineStyle: word("toolbar.underlineStyleLabel"),
            underlineColor: word("toolbar.underlineColorLabel"),
            sectionEffects: word("toolbar.fontSectionEffects"),
            strike: word("toolbar.strikethrough"),
            doubleStrike: word("toolbar.doubleStrikethrough"),
            superscript: word("toolbar.superscript"),
            subscript: word("toolbar.subscript"),
            smallCaps: word("toolbar.smallCaps"),
            allCaps: word("toolbar.allCaps"),
            hidden: "",
            preview: word("toolbar.fontDialogPreview"),
            previewSample: word("toolbar.fontDialogPreviewSample"),
            setAsDefault: "",
            setAsDefaultHint: "",
            mixed: word("toolbar.fontPickerMixed"),
            documentDefault: inheritedFamilyLabel,
            followLatin: word("toolbar.fontDialogFollowLatin"),
            automatic: word("toolbar.textColorAutomatic"),
            styleValue: Object.fromEntries(
              (["regular", "italic", "bold", "boldItalic"] as const).map((style) => [
                style,
                word(`toolbar.fontDialogStyleValue.${style}`),
              ]),
            ) as Record<FontDialogStyle, string>,
            underlineValue: Object.fromEntries(
              (["none", "single", "double"] as const).map((style) => [
                style,
                word(`toolbar.underlineStyleValue.${style}`),
              ]),
            ) as Record<FontDialogUnderline, string>,
            themeColors: word("toolbar.themeColors"),
            standardColors: word("toolbar.standardColors"),
            recentColors: word("spreadsheet.recentColors"),
            moreColors: word("toolbar.textColorCustom"),
          }}
        />
      </div>
      {spacingWritable ? (
        <div
          id={`${TEST_ID}-page-spacing`}
          className="note-editor-font-dialog__page"
          role="tabpanel"
          aria-labelledby={`${TEST_ID}-tab-spacing`}
          hidden={tab !== "spacing"}
        >
          <section className="office-dialog__group">
            <h3>{word("toolbar.fontSectionCharacterSpacing")}</h3>
            <div className="note-editor-font-dialog__grid">
              <label className="office-dialog__field">
                {word("office.chartFontDialog.spacing")}
                <select
                  data-testid={`${TEST_ID}-spacing`}
                  value={draft.spacing}
                  onChange={(event) => change("spacing", event.target.value as PresentationFontDialogSpacingKind)}
                >
                  <option value="normal">{word("office.chartFontDialog.spacingNormal")}</option>
                  <option value="expanded">{word("office.chartFontDialog.spacingExpanded")}</option>
                  <option value="condensed">{word("office.chartFontDialog.spacingCondensed")}</option>
                </select>
              </label>
              <label className="office-dialog__field">
                {word("office.chartFontDialog.by")}
                <input
                  type="number"
                  min={0}
                  max={PRESENTATION_MAX_CHARACTER_SPACING_PT}
                  step="0.1"
                  data-testid={`${TEST_ID}-spacing-by`}
                  disabled={draft.spacing === "normal"}
                  value={draft.spacingBy}
                  placeholder="0"
                  onChange={(event) => change("spacingBy", event.target.value)}
                />
              </label>
            </div>
          </section>
        </div>
      ) : null}
    </OfficeDialogShell>
  );
}
