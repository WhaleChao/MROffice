"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Advanced Filter dialog (数据 ▸ 高级筛选).
 *
 * Excel's dialog asks four things — the mode, the list range, the criteria
 * range, and (when copying) where to put the result — plus the unique-records
 * checkbox. This component collects exactly that and hands it to the host; the
 * matching rules live in `advanced-filter.ts` and the sheet mutation lives in
 * the window, so this file stays a form.
 *
 * Display copy comes from the caller's active message catalogue.
 */

import { useEffect, useRef, useState } from "react";

import {
  officeDialogParts as dialogParts,
  spreadsheetDialogClassName,
} from "@mrnote/core/web/components/office-ribbon/officeDialogChrome.tsx";
import { SpreadsheetCollapseDialogButton } from "../../components/notebook/contents/spreadsheet/SpreadsheetRangePicker";
import type { SpreadsheetRangePick } from "../../components/notebook/contents/spreadsheet/useSpreadsheetRangePick";

export type AdvancedFilterMode = "in-place" | "copy";

export interface AdvancedFilterRequest {
  mode: AdvancedFilterMode;
  list: string;
  criteria: string;
  copyTo: string;
  uniqueOnly: boolean;
}

export interface AdvancedFilterModeAvailability {
  readonly enabled: boolean;
  readonly reason?: string;
}

export interface AdvancedFilterDialogProps {
  open: boolean;
  translate: (key: string, values?: Record<string, string | number>) => string;
  /** Pre-fill the list range with the current selection. */
  defaultList?: string;
  /** The sheet's Criteria range from its last Advanced Filter, which Excel
   * fills in again; empty keeps what was typed last. */
  defaultCriteria?: string;
  /** Live writer availability for each confirmed mode. */
  modeAvailability?: Readonly<
    Record<AdvancedFilterMode, AdvancedFilterModeAvailability>
  >;
  /** Runs the filter; the returned message is shown verbatim. A promise
   * (criteria formulas evaluated by the calculation session, C2) keeps
   * Apply disabled until it settles. */
  onApply: (request: AdvancedFilterRequest) =>
    | { ok: boolean; message: string }
    | Promise<{ ok: boolean; message: string }>;
  onClose: () => void;
  /** The window's RefEdit: each range box gets a Collapse Dialog button. */
  rangePick?: SpreadsheetRangePick;
}

export function AdvancedFilterDialog({
  open,
  translate,
  defaultList = "",
  defaultCriteria = "",
  modeAvailability,
  onApply,
  onClose,
  rangePick,
}: AdvancedFilterDialogProps) {
  const text = {
    title: translate("spreadsheet.helperCopy.advancedFilter.title"),
    inPlace: translate("spreadsheet.helperCopy.advancedFilter.inPlace"),
    copy: translate("spreadsheet.helperCopy.advancedFilter.copy"),
    list: translate("spreadsheet.helperCopy.advancedFilter.list"),
    criteria: translate("spreadsheet.helperCopy.advancedFilter.criteria"),
    copyTo: translate("spreadsheet.helperCopy.advancedFilter.copyTo"),
    unique: translate("spreadsheet.helperCopy.advancedFilter.unique"),
    apply: translate("spreadsheet.helperCopy.advancedFilter.apply"),
    close: translate("spreadsheet.helperCopy.advancedFilter.close"),
    hint: translate("spreadsheet.helperCopy.advancedFilter.hint"),
    collapse: translate("spreadsheet.cfDialog.collapse"),
  };
  const inPlaceAvailability = modeAvailability?.["in-place"] ?? {
    enabled: true,
  };
  const copyAvailability = modeAvailability?.copy ?? { enabled: true };
  const [mode, setMode] = useState<AdvancedFilterMode>("in-place");
  const [list, setList] = useState(defaultList);
  const [criteria, setCriteria] = useState("");
  const [copyTo, setCopyTo] = useState("");
  const [uniqueOnly, setUniqueOnly] = useState(false);
  const [status, setStatus] = useState<{ ok: boolean; message: string } | null>(
    null,
  );
  const [busy, setBusy] = useState(false);
  // An answer for an earlier Apply (or an earlier opening) is dropped.
  const applyGenerationRef = useRef(0);
  const apply = () => {
    const outcome = onApply({ mode, list, criteria, copyTo, uniqueOnly });
    if (!(outcome instanceof Promise)) { setStatus(outcome); return; }
    const generation = ++applyGenerationRef.current;
    setBusy(true);
    void outcome.then((result) => {
      if (applyGenerationRef.current === generation) setStatus(result);
    }, () => undefined).finally(() => {
      if (applyGenerationRef.current === generation) setBusy(false);
    });
  };

  useEffect(() => {
    if (!open) return;
    setMode(
      inPlaceAvailability.enabled
        ? "in-place"
        : copyAvailability.enabled
          ? "copy"
          : "in-place",
    );
    setList(defaultList);
    if (defaultCriteria) setCriteria(defaultCriteria);
    setStatus(null);
    applyGenerationRef.current += 1;
    setBusy(false);
  }, [
    copyAvailability.enabled,
    defaultCriteria,
    defaultList,
    inPlaceAvailability.enabled,
    open,
  ]);

  if (!open) return null;
  // Collapsed to the RefEdit bar while one of its ranges is picked.
  if (rangePick?.picking("spreadsheet-advanced-filter-")) return null;

  // Excel writes a picked range here as $A$1:$C$20.
  const pickButton = (
    id: "list" | "criteria" | "copy-to",
    label: string,
    value: string,
    write: (text: string) => void,
  ) =>
    rangePick ? (
      <SpreadsheetCollapseDialogButton
        label={text.collapse}
        onClick={() =>
          rangePick.start({
            id: `spreadsheet-advanced-filter-${id}`,
            dialogLabel: text.title,
            inputLabel: label,
            style: "absolute",
            value,
            onChange: (next) => {
              write(next);
              setStatus(null);
            },
          })
        }
        testId={`spreadsheet-advanced-filter-${id}-pick`}
      />
    ) : null;

  const selectedModeAvailability =
    mode === "in-place" ? inPlaceAvailability : copyAvailability;

  return (
    <div
      className={spreadsheetDialogClassName("md")}
      role="dialog"
      aria-modal="true"
      aria-label={text.title}
      data-testid="spreadsheet-advanced-filter-dialog"
      onKeyDown={(event) => {
        if (event.key === "Escape") {
          event.preventDefault();
          onClose();
        }
      }}
    >
      <span className={dialogParts.title}>{text.title}</span>
      <label className={dialogParts.check}>
        <input
          type="radio"
          name="advanced-filter-mode"
          checked={mode === "in-place"}
          disabled={!inPlaceAvailability.enabled}
          aria-label={
            inPlaceAvailability.reason
              ? `${text.inPlace}: ${inPlaceAvailability.reason}`
              : text.inPlace
          }
          title={inPlaceAvailability.reason}
          onChange={() => {
            setMode("in-place");
            setStatus(null);
          }}
          data-testid="spreadsheet-advanced-filter-mode-in-place"
        />
        <span className={dialogParts.label}>{text.inPlace}</span>
      </label>
      {!inPlaceAvailability.enabled && inPlaceAvailability.reason ? (
        <span
          data-testid="spreadsheet-advanced-filter-mode-in-place-reason"
          className={dialogParts.hint}
        >
          {inPlaceAvailability.reason}
        </span>
      ) : null}
      <label className={dialogParts.check}>
        <input
          type="radio"
          name="advanced-filter-mode"
          checked={mode === "copy"}
          disabled={!copyAvailability.enabled}
          aria-label={
            copyAvailability.reason
              ? `${text.copy}: ${copyAvailability.reason}`
              : text.copy
          }
          title={copyAvailability.reason}
          onChange={() => {
            setMode("copy");
            setStatus(null);
          }}
          data-testid="spreadsheet-advanced-filter-mode-copy"
        />
        <span className={dialogParts.label}>{text.copy}</span>
      </label>
      {!copyAvailability.enabled && copyAvailability.reason ? (
        <span
          data-testid="spreadsheet-advanced-filter-mode-copy-reason"
          className={dialogParts.hint}
        >
          {copyAvailability.reason}
        </span>
      ) : null}
      <label className={dialogParts.field}>
        <span className={dialogParts.label}>{text.list}</span>
        <span style={{ display: "flex", gap: 4 }}>
          <input
            value={list}
            onChange={(event) => setList(event.target.value)}
            placeholder="A1:C20"
            data-testid="spreadsheet-advanced-filter-list"
            style={{ flex: 1, minWidth: 0 }}
          />
          {pickButton("list", text.list, list, setList)}
        </span>
      </label>
      <label className={dialogParts.field}>
        <span className={dialogParts.label}>{text.criteria}</span>
        <span style={{ display: "flex", gap: 4 }}>
          <input
            value={criteria}
            onChange={(event) => setCriteria(event.target.value)}
            placeholder="E1:F3"
            data-testid="spreadsheet-advanced-filter-criteria"
            style={{ flex: 1, minWidth: 0 }}
          />
          {pickButton("criteria", text.criteria, criteria, setCriteria)}
        </span>
      </label>
      {mode === "copy" ? (
        <label className={dialogParts.field}>
          <span className={dialogParts.label}>{text.copyTo}</span>
          <span style={{ display: "flex", gap: 4 }}>
            <input
              value={copyTo}
              onChange={(event) => setCopyTo(event.target.value)}
              placeholder="H1"
              data-testid="spreadsheet-advanced-filter-copy-to"
              style={{ flex: 1, minWidth: 0 }}
            />
            {pickButton("copy-to", text.copyTo, copyTo, setCopyTo)}
          </span>
        </label>
      ) : null}
      <label className={dialogParts.check}>
        <input
          type="checkbox"
          checked={uniqueOnly}
          onChange={(event) => setUniqueOnly(event.target.checked)}
          data-testid="spreadsheet-advanced-filter-unique"
        />
        <span className={dialogParts.label}>{text.unique}</span>
      </label>
      <span className={dialogParts.hint}>{text.hint}</span>
      {status ? (
        <span
          data-testid="spreadsheet-advanced-filter-status"
          className={dialogParts.message}
          data-tone={status.ok ? undefined : "error"}
        >
          {status.message}
        </span>
      ) : null}
      <div className={dialogParts.footer}>
        <button
          type="button"
          data-testid="spreadsheet-advanced-filter-close"
          onClick={onClose}
          className={dialogParts.button}
        >
          {text.close}
        </button>
        <button
          type="button"
          data-testid="spreadsheet-advanced-filter-apply"
          disabled={!selectedModeAvailability.enabled || busy}
          aria-busy={busy || undefined}
          title={selectedModeAvailability.reason}
          onClick={apply}
          className={dialogParts.primaryButton}
        >
          {text.apply}
        </button>
      </div>
    </div>
  );
}

export default AdvancedFilterDialog;
