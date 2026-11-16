"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * PowerPoint's Paragraph dialog (Home > Paragraph launcher, Line Spacing >
 * Line Spacing Options...): the Indents and Spacing page, and the Asian
 * Typography page's Text alignment (`a:pPr@fontAlgn`: Auto, Top, Centered,
 * Baseline, Bottom; the box Docs' Paragraph dialog shares). What OK writes is decided
 * by `lib/presentation-paragraph-dialog`; the host commits it through the Home
 * paragraph writer. The labels are the Docs Paragraph dialog's own strings
 * (Word and PowerPoint share them) plus PowerPoint's Before text / First line
 * / Hanging.
 */

import { useRef, useState } from "react";
import { useLocale, useTranslations } from "@mrnote/core/host/next-intl.tsx";

import { OfficeDialogShell } from "@mrnote/core/web/components/office-ribbon/OfficeDialogShell.tsx";
import { officeDialogParts } from "@mrnote/core/web/components/office-ribbon/officeDialogChrome.tsx";
import {
  OFFICE_TEXT_ALIGNMENT_CHOICES,
  OfficeTextAlignmentField,
} from "@mrnote/core/web/components/office-ribbon/OfficeTextAlignmentField.tsx";
import {
  formatPresentationMeasurement,
  presentationMeasurementUnit,
} from "@mrnote/core/web/lib/presentation-measurement.ts";
import {
  changePresentationParagraphDialogDraft,
  PRESENTATION_PARAGRAPH_DIALOG_BOUNDS,
  presentationParagraphDialogDraft,
  presentationParagraphDialogResult,
  type PresentationLineRule,
  type PresentationParagraphDialogDraft,
  type PresentationParagraphDialogField,
  type PresentationParagraphDialogProblem,
  type PresentationParagraphDialogResult,
  type PresentationParagraphDialogSource,
  type PresentationSpecialIndent,
} from "@mrnote/core/web/lib/presentation-paragraph-dialog.ts";
import {
  PARAGRAPH_INDENT_EMU_LIMIT,
  type ParagraphAlign,
  type PresentationFontAlign,
} from "@mrnote/core/web/lib/presentation-paragraphs.ts";

export interface PresentationParagraphDialogProps {
  readonly open: boolean;
  readonly source: PresentationParagraphDialogSource;
  /** False where the target keeps no row-level Exactly height (a table cell, objects). */
  readonly exactWritable: boolean;
  /** False where only alignment is writable (a table cell). */
  readonly spacingWritable: boolean;
  /** False where the target keeps no row indentation (a table cell, objects). */
  readonly indentWritable?: boolean;
  /** False where the target keeps no rows to state a Text alignment (objects). */
  readonly fontAlignWritable?: boolean;
  readonly onApply: (result: PresentationParagraphDialogResult) => boolean;
  readonly onCancel: () => void;
}

const TEST_ID = "presentation-paragraph-dialog";

/** ST_TextFontAlignType and the shared box's choices, both ways. */
const FONT_ALIGN_CHOICE = { auto: "auto", t: "top", ctr: "center", base: "baseline", b: "bottom" } as const;
type FontAlignChoice = (typeof OFFICE_TEXT_ALIGNMENT_CHOICES)[number];
const CHOICE_FONT_ALIGN = Object.fromEntries(
  Object.entries(FONT_ALIGN_CHOICE).map(([value, choice]) => [choice, value]),
) as Record<FontAlignChoice, PresentationFontAlign>;

export function PresentationParagraphDialog(
  props: PresentationParagraphDialogProps,
): React.JSX.Element | null {
  if (!props.open) return null;
  return <PresentationParagraphDialogBody {...props} />;
}

function PresentationParagraphDialogBody({
  source,
  exactWritable,
  spacingWritable,
  indentWritable = false,
  fontAlignWritable = false,
  onApply,
  onCancel,
}: PresentationParagraphDialogProps): React.JSX.Element {
  const [page, setPage] = useState<"indents" | "asian">("indents");
  const t = useTranslations("console-notebooks");
  const word = (key: string, values?: Record<string, string | number>) =>
    t(key as Parameters<typeof t>[0], values as never);
  const unit = presentationMeasurementUnit(useLocale());
  const [draft, setDraft] = useState<PresentationParagraphDialogDraft>(() =>
    presentationParagraphDialogDraft(source, unit),
  );
  const draftRef = useRef(draft);
  const [touched, setTouched] = useState<ReadonlySet<PresentationParagraphDialogField>>(
    () => new Set(),
  );
  const [problem, setProblem] = useState<PresentationParagraphDialogProblem | null>(null);
  const initialFocusRef = useRef<HTMLSelectElement | null>(null);

  const change = <K extends PresentationParagraphDialogField>(
    field: K,
    value: PresentationParagraphDialogDraft[K],
  ) => {
    setProblem(null);
    const result = changePresentationParagraphDialogDraft(draftRef.current, field, value);
    draftRef.current = result.draft;
    setDraft(result.draft);
    setTouched((previous) => new Set([...previous, ...result.touched]));
  };

  const apply = () => {
    const result = presentationParagraphDialogResult(draft, touched);
    if ("problem" in result) {
      setProblem(result.problem);
      return;
    }
    if (touched.size === 0 || onApply(result)) onCancel();
  };

  const problemMessage = (() => {
    if (!problem) return null;
    if (problem === "beforeText" || problem === "specialBy") {
      return word("presentation.paragraphDialog.invalidValue", {
        field:
          problem === "beforeText"
            ? word("presentation.paragraphDialog.beforeText")
            : word("toolbar.paragraphByValue"),
        minimum: 0,
        maximum: formatPresentationMeasurement(PARAGRAPH_INDENT_EMU_LIMIT, unit),
      });
    }
    const bounds =
      problem === "lineAt"
        ? PRESENTATION_PARAGRAPH_DIALOG_BOUNDS[draft.lineRule === "exactly" ? "exactly" : "multiple"]
        : PRESENTATION_PARAGRAPH_DIALOG_BOUNDS.spacing;
    const field = {
      spaceBefore: word("toolbar.paragraphBefore"),
      spaceAfter: word("toolbar.paragraphAfter"),
      lineAt: word("toolbar.paragraphAtValue"),
    }[problem];
    return word("presentation.paragraphDialog.invalidValue", {
      field,
      minimum: bounds[0],
      maximum: bounds[1],
    });
  })();

  const unavailable = word("presentation.fontDialog.optionUnavailable");
  const unitLabel = unit === "in" ? '"' : word("presentation.unit.cm");

  const alignOptions: readonly (readonly [ParagraphAlign, string])[] = [
    ["left", word("toolbar.alignLeft")],
    ["center", word("toolbar.alignCenter")],
    ["right", word("toolbar.alignRight")],
    ["justify", word("toolbar.alignJustify")],
    ["distribute", word("presentation.align.distribute")],
  ];
  const lineRules: readonly (readonly [PresentationLineRule, string])[] = [
    ["single", word("toolbar.paragraphLineSingle")],
    ["onePointFive", word("toolbar.paragraphLineOneHalf")],
    ["double", word("toolbar.paragraphLineDouble")],
    ["exactly", word("toolbar.paragraphLineExactly")],
    ["multiple", word("toolbar.paragraphLineMultiple")],
  ];
  const atEnabled = draft.lineRule === "exactly" || draft.lineRule === "multiple";

  return (
    <OfficeDialogShell
      open
      testIdPrefix={TEST_ID}
      title={word("toolbar.paragraphDialogTitle")}
      description=""
      defaultActionLabel={word("toolbar.dialogOk")}
      defaultActionTestId={`${TEST_ID}-apply`}
      cancelActionLabel={word("toolbar.pageSetupCancel")}
      cancelActionTestId={`${TEST_ID}-cancel`}
      overlayTestId={`${TEST_ID}-layer`}
      onDefaultAction={apply}
      onCancel={onCancel}
      initialFocusRef={initialFocusRef}
      className="sm:max-w-[560px]"
      footerStart={
        problemMessage ? (
          <span role="alert" className="office-dialog__error" data-testid={`${TEST_ID}-problem`}>
            {problemMessage}
          </span>
        ) : undefined
      }
      tabStrip={
        <div className="office-dialog__tabs" role="tablist" aria-label={word("toolbar.paragraphDialogTitle")}>
          <button
            type="button"
            role="tab"
            aria-selected={page === "indents"}
            className="office-dialog__tab"
            data-testid={`${TEST_ID}-tab-indents`}
            onClick={() => setPage("indents")}
          >
            {word("toolbar.paragraphTabIndents")}
          </button>
          <button
            type="button"
            role="tab"
            aria-selected={page === "asian"}
            className="office-dialog__tab"
            data-testid={`${TEST_ID}-tab-asian`}
            onClick={() => setPage("asian")}
          >
            {word("toolbar.paragraphTabAsian")}
          </button>
        </div>
      }
    >
      <div className="grid gap-3" hidden={page !== "asian"} data-testid={`${TEST_ID}-asian`}>
        <section className="office-dialog__group" title={fontAlignWritable ? undefined : unavailable}>
          <h3>{word("toolbar.paragraphSectionGeneral")}</h3>
          <OfficeTextAlignmentField
            testId={`${TEST_ID}-font-align`}
            value={draft.fontAlign === "" ? "" : FONT_ALIGN_CHOICE[draft.fontAlign]}
            choices={OFFICE_TEXT_ALIGNMENT_CHOICES}
            disabled={!fontAlignWritable}
            placeholder={word("toolbar.fontPickerMixed")}
            onChange={(choice) => change("fontAlign", CHOICE_FONT_ALIGN[choice])}
          />
        </section>
      </div>
      <div className="grid gap-3" hidden={page !== "indents"}>
        <section className="office-dialog__group">
          <h3>{word("toolbar.paragraphSectionGeneral")}</h3>
          <label className={officeDialogParts.field}>
            <span>{word("toolbar.paragraphAlignment")}</span>
            <select
              ref={initialFocusRef}
              data-testid={`${TEST_ID}-align`}
              value={draft.align}
              onChange={(event) => change("align", event.target.value as ParagraphAlign)}
            >
              {draft.align === "" ? <option value="">{word("toolbar.fontPickerMixed")}</option> : null}
              {alignOptions.map(([value, label]) => (
                <option key={value} value={value}>
                  {label}
                </option>
              ))}
            </select>
          </label>
        </section>
        <section className="office-dialog__group" title={indentWritable ? undefined : unavailable}>
          <h3>{word("toolbar.paragraphSectionIndentation")}</h3>
          <div className="grid grid-cols-3 gap-3">
            <label className={officeDialogParts.field}>
              <span>{word("presentation.paragraphDialog.beforeText")}</span>
              <span className={officeDialogParts.measure}>
                <input
                  inputMode="decimal"
                  disabled={!indentWritable}
                  data-testid={`${TEST_ID}-before-text`}
                  value={draft.beforeText}
                  onChange={(event) => change("beforeText", event.target.value)}
                />
                <span aria-hidden="true">{unitLabel}</span>
              </span>
            </label>
            <label className={officeDialogParts.field}>
              <span>{word("toolbar.paragraphSpecialIndent")}</span>
              <select
                disabled={!indentWritable}
                data-testid={`${TEST_ID}-special`}
                value={draft.special}
                onChange={(event) => change("special", event.target.value as PresentationSpecialIndent)}
              >
                <option value="none">{word("toolbar.paragraphSpecialNone")}</option>
                <option value="firstLine">{word("presentation.paragraphDialog.firstLine")}</option>
                <option value="hanging">{word("presentation.paragraphDialog.hanging")}</option>
              </select>
            </label>
            <label className={officeDialogParts.field}>
              <span>{word("toolbar.paragraphByValue")}</span>
              <span className={officeDialogParts.measure}>
                <input
                  inputMode="decimal"
                  disabled={!indentWritable || draft.special === "none"}
                  data-testid={`${TEST_ID}-special-by`}
                  value={draft.specialBy}
                  onChange={(event) => change("specialBy", event.target.value)}
                />
                <span aria-hidden="true">{unitLabel}</span>
              </span>
            </label>
          </div>
        </section>
        <section className="office-dialog__group" title={spacingWritable ? undefined : unavailable}>
          <h3>{word("toolbar.paragraphSectionSpacing")}</h3>
          <div className="grid grid-cols-2 gap-3">
            <label className={officeDialogParts.field}>
              <span>{word("toolbar.paragraphBefore")}</span>
              <input
                inputMode="decimal"
                disabled={!spacingWritable}
                data-testid={`${TEST_ID}-space-before`}
                value={draft.spaceBefore}
                onChange={(event) => change("spaceBefore", event.target.value)}
              />
            </label>
            <label className={officeDialogParts.field}>
              <span>{word("toolbar.paragraphAfter")}</span>
              <input
                inputMode="decimal"
                disabled={!spacingWritable}
                data-testid={`${TEST_ID}-space-after`}
                value={draft.spaceAfter}
                onChange={(event) => change("spaceAfter", event.target.value)}
              />
            </label>
            <label className={officeDialogParts.field}>
              <span>{word("toolbar.paragraphLineSpacing")}</span>
              <select
                disabled={!spacingWritable}
                data-testid={`${TEST_ID}-line-rule`}
                value={draft.lineRule}
                onChange={(event) => change("lineRule", event.target.value as PresentationLineRule)}
              >
                {draft.lineRule === "" ? <option value="">{word("toolbar.fontPickerMixed")}</option> : null}
                {lineRules.map(([value, label]) => (
                  <option
                    key={value}
                    value={value}
                    disabled={value === "exactly" && !exactWritable}
                    title={value === "exactly" && !exactWritable ? unavailable : undefined}
                  >
                    {label}
                  </option>
                ))}
              </select>
            </label>
            <label className={officeDialogParts.field}>
              <span>{word("toolbar.paragraphAtValue")}</span>
              <input
                inputMode="decimal"
                disabled={!spacingWritable || !atEnabled}
                data-testid={`${TEST_ID}-line-at`}
                value={draft.lineAt}
                onChange={(event) => change("lineAt", event.target.value)}
              />
            </label>
          </div>
        </section>
      </div>
    </OfficeDialogShell>
  );
}
