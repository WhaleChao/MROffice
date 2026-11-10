"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { EditorLoadingState } from "@mrnote/core/web/components/ui/editor-loading-state.tsx";
import { useEffect, useState, type RefObject } from "react";
import { ArrowRight } from "@mrnote/core/web/components/office-ribbon/icons/houseIcons.tsx";
import { useTranslations } from "@mrnote/core/host/next-intl.tsx";

import UserAvatar from "@mrnote/core/web/components/console/UserAvatar.tsx";
import {
  OfficeTaskPane,
  officeTaskPaneParts,
} from "@mrnote/core/web/components/office-ribbon/OfficeTaskPane.tsx";
import {
  changeCellLabel,
  changeLocationLabel,
  changeRangeLabel,
  type SpreadsheetChangeCell,
  type SpreadsheetChangeEntry,
} from "@mrnote/core/web/lib/spreadsheet/change-log.ts";
import { avatarFromWire } from "@mrnote/core/web/lib/user-avatar.ts";

import styles from "./SpreadsheetChangesPane.module.css";

export type SpreadsheetChangesScope = "workbook" | "sheet" | "range";

export interface SpreadsheetChangesPaneProps {
  /** Merged, filtered and newest first. */
  entries: readonly SpreadsheetChangeEntry[];
  status: "loading" | "ready" | "error";
  hasMore: boolean;
  loadingMore: boolean;
  retentionDays: number;
  scope: SpreadsheetChangesScope;
  activeSheetName: string;
  rangeText: string;
  rangeInvalid: boolean;
  locale: string;
  /** Epoch ms the relative times count from (tests pin it). */
  now?: number;
  paneRef?: RefObject<HTMLElement | null>;
  onScopeChange: (scope: SpreadsheetChangesScope) => void;
  onRangeTextChange: (text: string) => void;
  onApplyRange: () => void;
  onLoadMore: () => void;
  onSelect: (entry: SpreadsheetChangeEntry, cell?: SpreadsheetChangeCell) => void;
  onClose: () => void;
}

const PREVIEW_CHARS = 60;

function preview(value: string): string {
  const flat = value.replace(/\s+/g, " ");
  return flat.length > PREVIEW_CHARS ? `${flat.slice(0, PREVIEW_CHARS - 1)}…` : flat;
}

function relativeTime(at: number, now: number, locale: string): string {
  const seconds = Math.round((at - now) / 1000);
  const abs = Math.abs(seconds);
  try {
    if (abs >= 7 * 86_400) {
      return new Intl.DateTimeFormat(locale, {
        day: "numeric",
        month: "short",
        year: "numeric",
      }).format(new Date(at));
    }
    const rtf = new Intl.RelativeTimeFormat(locale, { numeric: "auto" });
    if (abs < 45) return rtf.format(0, "second");
    if (abs < 3_600) return rtf.format(Math.round(seconds / 60), "minute");
    if (abs < 86_400) return rtf.format(Math.round(seconds / 3_600), "hour");
    return rtf.format(Math.round(seconds / 86_400), "day");
  } catch {
    return new Date(at).toLocaleString();
  }
}

function absoluteTime(at: number, locale: string): string {
  try {
    return new Intl.DateTimeFormat(locale, {
      dateStyle: "medium",
      timeStyle: "short",
    }).format(new Date(at));
  } catch {
    return new Date(at).toLocaleString();
  }
}

/**
 * Excel's Changes pane (Review > Show Changes): one card per edit, newest
 * first, with who made it, when, where and what changed. A bulk edit lists its
 * cells under "See changes". Clicking a card or a cell selects it on its
 * sheet. The filter narrows the list to the active sheet or to a range typed
 * in the box and confirmed with the arrow, as in Excel.
 */
export function SpreadsheetChangesPane({
  entries,
  status,
  hasMore,
  loadingMore,
  retentionDays,
  scope,
  activeSheetName,
  rangeText,
  rangeInvalid,
  locale,
  now,
  paneRef,
  onScopeChange,
  onRangeTextChange,
  onApplyRange,
  onLoadMore,
  onSelect,
  onClose,
}: SpreadsheetChangesPaneProps) {
  const t = useTranslations("console-notebooks");
  const [expanded, setExpanded] = useState<ReadonlySet<string>>(() => new Set());
  // "2 minutes ago" moves on while the pane stays open.
  const [tick, setTick] = useState(() => Date.now());
  useEffect(() => {
    if (now !== undefined) return;
    const timer = setInterval(() => setTick(Date.now()), 30_000);
    return () => clearInterval(timer);
  }, [now]);
  const clock = now ?? tick;

  const cellSentence = (cell: SpreadsheetChangeCell) => {
    const values = {
      cell: changeCellLabel(cell),
      before: preview(cell.before),
      after: preview(cell.after),
    };
    if (cell.before === "") return t("spreadsheet.changes.entered", values);
    if (cell.after === "") return t("spreadsheet.changes.cleared", values);
    return t("spreadsheet.changes.changed", values);
  };

  const summary = (entry: SpreadsheetChangeEntry) => {
    const range = changeRangeLabel(entry);
    switch (entry.kind) {
      case "insert_rows":
        return t("spreadsheet.changes.insertedRows", { range });
      case "delete_rows":
        return t("spreadsheet.changes.deletedRows", { range });
      case "insert_columns":
        return t("spreadsheet.changes.insertedColumns", { range });
      case "delete_columns":
        return t("spreadsheet.changes.deletedColumns", { range });
      default: {
        const only = entry.cells[0];
        if (entry.cellCount === 1 && only && !entry.truncated) return cellSentence(only);
        return t("spreadsheet.changes.bulk", { count: entry.cellCount, range });
      }
    }
  };

  const toggle = (id: string) =>
    setExpanded((current) => {
      const next = new Set(current);
      if (next.has(id)) next.delete(id);
      else next.add(id);
      return next;
    });

  return (
    <OfficeTaskPane
      title={t("spreadsheet.changes.title")}
      closeLabel={t("common.close")}
      onClose={onClose}
      testIdPrefix="spreadsheet-changes"
      testId="spreadsheet-changes-pane"
      closeTestId="spreadsheet-changes-close"
      paneRef={paneRef}
      dockRight={false}
      className={styles.pane}
    >
      <div className={styles.filter}>
        <label className={officeTaskPaneParts.field}>
          <span className={styles.filterLabel}>{t("spreadsheet.changes.filterLabel")}</span>
          <select
            value={scope}
            onChange={(event) => onScopeChange(event.target.value as SpreadsheetChangesScope)}
            data-testid="spreadsheet-changes-scope"
          >
            <option value="workbook">{t("spreadsheet.changes.scopeWorkbook")}</option>
            <option value="sheet">
              {t("spreadsheet.changes.scopeSheet", { name: activeSheetName })}
            </option>
            <option value="range">{t("spreadsheet.changes.scopeRange")}</option>
          </select>
        </label>
        {scope === "range" ? (
          <form
            className={styles.rangeRow}
            onSubmit={(event) => {
              event.preventDefault();
              onApplyRange();
            }}
          >
            <input
              aria-label={t("spreadsheet.changes.rangeLabel")}
              aria-invalid={rangeInvalid || undefined}
              className={styles.rangeInput}
              value={rangeText}
              placeholder="A1:C10"
              spellCheck={false}
              onChange={(event) => onRangeTextChange(event.target.value)}
              data-testid="spreadsheet-changes-range"
            />
            <button
              type="submit"
              className={styles.rangeApply}
              aria-label={t("spreadsheet.changes.rangeApply")}
              title={t("spreadsheet.changes.rangeApply")}
              data-testid="spreadsheet-changes-range-apply"
            >
              <ArrowRight size={14} aria-hidden />
            </button>
          </form>
        ) : null}
        {scope === "range" && rangeInvalid ? (
          <p role="alert" className={styles.error}>
            {t("spreadsheet.changes.rangeInvalid")}
          </p>
        ) : null}
      </div>
      {status === "error" ? (
        <p className={officeTaskPaneParts.hint} data-testid="spreadsheet-changes-error">
          {t("spreadsheet.changes.loadFailed")}
        </p>
      ) : null}
      {entries.length === 0 && status !== "loading" ? (
        <p className={officeTaskPaneParts.hint} data-testid="spreadsheet-changes-empty">
          {scope === "workbook"
            ? t("spreadsheet.changes.empty")
            : t("spreadsheet.changes.emptyFiltered")}
        </p>
      ) : null}
      {status === "loading" && entries.length === 0 ? (
        <EditorLoadingState
          kind="spreadsheet"
          minimal
          label={t("spreadsheet.changes.loading")}
        />
      ) : null}
      <ol className={styles.list} data-testid="spreadsheet-changes-list">
        {entries.map((entry) => {
          const name = entry.actor.name || t("spreadsheet.changes.unknownAuthor");
          const bulk = entry.cellCount > 1 || entry.truncated;
          const open = expanded.has(entry.clientChangeId);
          return (
            <li
              key={entry.clientChangeId}
              className={styles.card}
              data-testid={`spreadsheet-changes-item-${entry.clientChangeId}`}
              data-pending={entry.pending ? "true" : undefined}
            >
              <button
                type="button"
                className={styles.cardMain}
                onClick={() => onSelect(entry)}
                data-testid={`spreadsheet-changes-select-${entry.clientChangeId}`}
              >
                <UserAvatar
                  name={name}
                  accentColor={entry.actor.accentColor ?? undefined}
                  avatar={avatarFromWire({
                    avatar_kind: entry.actor.avatarKind,
                    avatar_doodle: entry.actor.avatarDoodle,
                    avatar_url_small: entry.actor.avatarUrlSmall,
                  })}
                  size={28}
                />
                <span className={styles.cardText}>
                  <span className={styles.cardHead}>
                    <span className={styles.author}>{name}</span>
                    <time
                      className={styles.time}
                      dateTime={new Date(entry.createdAt).toISOString()}
                      title={absoluteTime(entry.createdAt, locale)}
                    >
                      {entry.pending
                        ? t("spreadsheet.changes.saving")
                        : relativeTime(entry.createdAt, clock, locale)}
                    </time>
                  </span>
                  <span className={styles.summary}>
                    {entry.kind === "undo" || entry.kind === "redo" ? (
                      <span className={styles.badge}>
                        {t(entry.kind === "undo" ? "spreadsheet.changes.undo" : "spreadsheet.changes.redo")}
                      </span>
                    ) : null}
                    {summary(entry)}
                  </span>
                  <span className={styles.location}>{changeLocationLabel(entry)}</span>
                </span>
              </button>
              {bulk && entry.cells.length > 0 ? (
                <button
                  type="button"
                  className={styles.seeChanges}
                  aria-expanded={open}
                  onClick={() => toggle(entry.clientChangeId)}
                  data-testid={`spreadsheet-changes-see-${entry.clientChangeId}`}
                >
                  {t(open ? "spreadsheet.changes.hideChanges" : "spreadsheet.changes.seeChanges")}
                </button>
              ) : null}
              {bulk && open ? (
                <ul className={styles.cells}>
                  {entry.cells.map((cell) => (
                    <li key={`${cell.r}:${cell.c}`}>
                      <button
                        type="button"
                        className={styles.cellLine}
                        onClick={() => onSelect(entry, cell)}
                        data-testid={`spreadsheet-changes-cell-${entry.clientChangeId}-${cell.r}:${cell.c}`}
                      >
                        {cellSentence(cell)}
                      </button>
                    </li>
                  ))}
                  {entry.truncated && entry.cellCount > entry.cells.length ? (
                    <li className={styles.more}>
                      {t("spreadsheet.changes.moreCells", {
                        count: entry.cellCount - entry.cells.length,
                      })}
                    </li>
                  ) : null}
                </ul>
              ) : null}
            </li>
          );
        })}
      </ol>
      {hasMore ? (
        <button
          type="button"
          className={styles.loadMore}
          disabled={loadingMore}
          onClick={onLoadMore}
          data-testid="spreadsheet-changes-load-more"
        >
          {t("spreadsheet.changes.loadMore")}
        </button>
      ) : null}
      <p className={styles.footnote}>
        {t("spreadsheet.changes.retention", { days: retentionDays })}
      </p>
    </OfficeTaskPane>
  );
}
