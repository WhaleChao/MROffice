"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { useCallback, useMemo, useRef, useState, type RefObject } from "react";
import { useLocale, useTranslations } from "@mrnote/core/host/next-intl.tsx";

import { OfficeDialogShell } from "@mrnote/core/web/components/office-ribbon/OfficeDialogShell.tsx";
import {
  documentLengthUnit,
  formatDocumentLengthTwips,
  parseDocumentLengthTwips,
  type DocumentLengthUnit,
} from "@mrnote/core/web/lib/document-length-fields.ts";
import { officeDialogParts as parts } from "@mrnote/core/web/components/office-ribbon/officeDialogChrome.tsx";
import type { OfficeSemanticIconId } from "@mrnote/core/web/components/office-ribbon/icons/officeIconRegistry.ts";
import { DocumentDialogIcon } from "@mrnote/core/web/components/console/editor/native-document/DocumentDialogIcon.tsx";
import { documentValueInputClassName } from "@mrnote/core/web/components/console/editor/documentValueInput.ts";
import type { LineNumberRestart } from "@mrnote/core/web/components/console/editor/native-document/pageSetup.ts";

/**
 * Word's Line Numbers dialog (Layout > Line Numbers > Line Numbering Options…
 * > Line Numbers…): Add line numbering, Start at, From text (Auto), Count by
 * and the three Numbering modes. Word opens it on the section at the caret;
 * with numbering off every field but the checkbox is unavailable and shows
 * Word's defaults (1, Auto, 1, Restart each page).
 */
export interface DocumentLineNumbersValue {
  readonly enabled: boolean;
  readonly start: number;
  readonly countBy: number;
  /** Points from the text; 0 is Word's Auto (`w:distance` omitted). */
  readonly distancePt: number;
  readonly restart: LineNumberRestart;
}

export interface DocumentLineNumbersDialogProps {
  readonly open: boolean;
  readonly canEdit?: boolean;
  readonly locale: string;
  readonly initial: DocumentLineNumbersValue;
  /** False where the document model keeps no start value (legacy pages). */
  readonly startSupported?: boolean;
  readonly returnFocusRef: RefObject<HTMLElement | null>;
  readonly onSubmit: (value: DocumentLineNumbersValue) => boolean;
  readonly onClose: () => void;
}

interface LineNumbersCopy {
  readonly title: string;
  readonly description: string;
  readonly add: string;
  readonly start: string;
  readonly fromText: string;
  readonly auto: string;
  readonly countBy: string;
  readonly numbering: string;
  readonly restart: Readonly<Record<LineNumberRestart, string>>;
  readonly ok: string;
  readonly cancel: string;
  readonly invalidStart: string;
  readonly invalidCountBy: string;
  readonly invalidDistance: string;
  readonly stale: string;
}

const RESTART_ORDER: readonly LineNumberRestart[] = [
  "newPage",
  "newSection",
  "continuous",
];

const LINE_NUMBER_RESTART_ICON = {
  newPage: "file.document",
  newSection: "layout.separator.horizontal",
  continuous: "list.numbered",
} as const satisfies Record<LineNumberRestart, OfficeSemanticIconId>;

function lineNumbersCopy(t: (key: string) => string): LineNumbersCopy {
  return {
    title: t("toolbar.lineNumbersDialog.title"),
    description: t("toolbar.lineNumbersDialog.description"),
    add: t("toolbar.lineNumbersDialog.add"),
    start: t("toolbar.lineNumbersDialog.start"),
    fromText: t("toolbar.lineNumbersDialog.fromText"),
    auto: t("toolbar.lineNumbersDialog.auto"),
    countBy: t("toolbar.lineNumbersDialog.countBy"),
    numbering: t("toolbar.lineNumbersDialog.numbering"),
    restart: {
      newPage: t("toolbar.lineNumbersDialog.restartNewPage"),
      newSection: t("toolbar.lineNumbersDialog.restartNewSection"),
      continuous: t("toolbar.lineNumbersDialog.restartContinuous"),
    },
    ok: t("toolbar.lineNumbersDialog.ok"),
    cancel: t("toolbar.lineNumbersDialog.cancel"),
    invalidStart: t("toolbar.lineNumbersDialog.invalidStart"),
    invalidCountBy: t("toolbar.lineNumbersDialog.invalidCountBy"),
    invalidDistance: t("toolbar.lineNumbersDialog.invalidDistance"),
    stale: t("toolbar.lineNumbersDialog.stale"),
  };
}

function wholeNumber(raw: string, min: number, max: number): number | null {
  const value = Number(raw.trim());
  return Number.isInteger(value) && value >= min && value <= max ? value : null;
}

/** "Auto" (or a blank box) is Word's automatic distance; otherwise a length
 * in Word's unit for the locale, or the unit it names ("0.5 cm", "18 pt"),
 * as Word's From text box takes it. Returns points. */
function parseDistance(
  raw: string,
  auto: string,
  unit: DocumentLengthUnit,
): number | null {
  const trimmed = raw.trim();
  if (!trimmed || trimmed.toLowerCase() === auto.toLowerCase()) return 0;
  const twips = parseDocumentLengthTwips(trimmed, unit);
  if (twips === null) return null;
  const points = twips / 20;
  return points >= 0 && points <= 72 ? points : null;
}

/** A distance in points as Word shows it: "0.64 厘米" in Chinese Word,
 * `0.25"` in US English. */
function formatDistance(points: number, unit: DocumentLengthUnit, cmLabel: string): string {
  const shown = formatDocumentLengthTwips(points * 20, unit).replace(/\.?0+$/, "");
  return unit === "cm" ? `${shown} ${cmLabel}` : `${shown}"`;
}

export function DocumentLineNumbersDialog({
  open,
  canEdit = true,
  initial,
  startSupported = true,
  returnFocusRef,
  onSubmit,
  onClose,
}: DocumentLineNumbersDialogProps): React.JSX.Element {
  const t = useTranslations("console-notebooks");
  const copy = useMemo(() => lineNumbersCopy(t), [t]);
  const addRef = useRef<HTMLInputElement | null>(null);
  const [enabled, setEnabled] = useState(initial.enabled);
  const [start, setStart] = useState(String(initial.start));
  const [countBy, setCountBy] = useState(String(initial.countBy));
  const lengthUnit = documentLengthUnit(useLocale());
  // The distance as first shown: left as it is, it keeps its exact points
  // (18 pt shows as 0.64 cm, which reads back as 18.15 pt).
  const [shownDistance] = useState(() =>
    initial.distancePt > 0
      ? formatDistance(initial.distancePt, lengthUnit, t("toolbar.indentUnitCm"))
      : copy.auto,
  );
  const [distance, setDistance] = useState(shownDistance);
  const [restart, setRestart] = useState<LineNumberRestart>(initial.restart);
  const [error, setError] = useState<string | null>(null);
  const fieldsDisabled = !canEdit || !enabled;

  const submit = useCallback(() => {
    if (!canEdit) return;
    let value: DocumentLineNumbersValue = { ...initial, enabled: false };
    if (enabled) {
      const parsedStart = startSupported
        ? wholeNumber(start, 1, 32767)
        : initial.start;
      if (parsedStart === null) {
        setError(copy.invalidStart);
        return;
      }
      const parsedCountBy = wholeNumber(countBy, 1, 100);
      if (parsedCountBy === null) {
        setError(copy.invalidCountBy);
        return;
      }
      const parsedDistance =
        distance === shownDistance && initial.distancePt > 0
          ? initial.distancePt
          : parseDistance(distance, copy.auto, lengthUnit);
      if (parsedDistance === null) {
        setError(copy.invalidDistance);
        return;
      }
      value = {
        enabled: true,
        start: parsedStart,
        countBy: parsedCountBy,
        distancePt: parsedDistance,
        restart,
      };
    }
    if (!onSubmit(value)) {
      setError(copy.stale);
      return;
    }
    onClose();
  }, [
    canEdit,
    copy,
    countBy,
    distance,
    lengthUnit,
    shownDistance,
    enabled,
    initial,
    onClose,
    onSubmit,
    restart,
    start,
    startSupported,
  ]);

  const numberField = (
    label: string,
    testId: string,
    value: string,
    onChange: (next: string) => void,
    props: { min?: number; max?: number; type?: "number" | "text" } = {},
  ) => (
    <label className={parts.field}>
      <span className={parts.label}>{label}</span>
      <input
        disabled={fieldsDisabled}
        type={props.type ?? "number"}
        inputMode={props.type === "text" ? "decimal" : "numeric"}
        min={props.min}
        max={props.max}
        step={1}
        value={value}
        className={documentValueInputClassName("min-w-0")}
        data-testid={testId}
        aria-invalid={error ? "true" : undefined}
        onChange={(event) => {
          onChange(event.target.value);
          setError(null);
        }}
      />
    </label>
  );

  return (
    <OfficeDialogShell
      open={open}
      title={copy.title}
      description={copy.description}
      defaultActionLabel={copy.ok}
      cancelActionLabel={copy.cancel}
      defaultActionDisabled={!canEdit}
      onDefaultAction={submit}
      onCancel={onClose}
      initialFocusRef={addRef}
      returnFocusRef={returnFocusRef}
      testIdPrefix="document-line-numbers"
      defaultActionTestId="document-line-numbers-apply"
    >
      <label className={parts.check}>
        <input
          ref={addRef}
          disabled={!canEdit}
          type="checkbox"
          checked={enabled}
          data-testid="document-line-numbers-add"
          onChange={(event) => {
            setEnabled(event.target.checked);
            setError(null);
          }}
        />
        <span>{copy.add}</span>
      </label>
      <div className="grid grid-cols-3 gap-3">
        {startSupported
          ? numberField(copy.start, "document-line-numbers-start", start, setStart, {
              min: 1,
              max: 32767,
            })
          : null}
        {numberField(
          copy.fromText,
          "document-line-numbers-distance",
          distance,
          setDistance,
          { type: "text" },
        )}
        {numberField(
          copy.countBy,
          "document-line-numbers-count-by",
          countBy,
          setCountBy,
          { min: 1, max: 100 },
        )}
      </div>
      <fieldset className={parts.group} disabled={fieldsDisabled}>
        <legend className={parts.sectionTitle}>{copy.numbering}</legend>
        {RESTART_ORDER.map((mode) => (
          <label key={mode} className={parts.check}>
            <input
              type="radio"
              name="document-line-numbers-restart"
              value={mode}
              checked={restart === mode}
              data-testid={`document-line-numbers-restart-${mode}`}
              onChange={() => {
                setRestart(mode);
                setError(null);
              }}
            />
            <DocumentDialogIcon icon={LINE_NUMBER_RESTART_ICON[mode]} />
            <span>{copy.restart[mode]}</span>
          </label>
        ))}
      </fieldset>
      {error ? (
        <p
          role="alert"
          className={parts.error}
          data-testid="document-line-numbers-error"
        >
          {error}
        </p>
      ) : null}
    </OfficeDialogShell>
  );
}
