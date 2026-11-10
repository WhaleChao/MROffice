"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { useEffect, useMemo, useRef, useState } from "react";
import { useTranslations } from "../../../../host/next-intl";
import { containDialogTab } from "../../../lib/dialog-tab-containment";
import {
  officeDialogClassName,
  officeDialogParts as parts,
  useOfficeDialogChrome,
} from "../../office-ribbon/officeDialogChrome";
import { Database, X } from "../../office-ribbon/icons/houseIcons";

import { DocumentDialogIcon } from "./native-document/DocumentDialogIcon";
import {
  MERGE_CONDITION_OPS,
  type MergeConditionOp,
} from "./native-document/mailMerge";

/** Symbol/label per op (symbols are language-neutral; word ops are i18n'd). */
const OP_SYMBOLS: Record<MergeConditionOp, string> = {
  eq: "=",
  ne: "≠",
  gt: ">",
  lt: "<",
  ge: "≥",
  le: "≤",
  contains: "",
  empty: "",
  notEmpty: "",
};

/**
 * Column picker for "Insert merge field" (Word/WPS "邮件合并" parity).
 *
 * Lists the recipient data source's columns (from the live `mailMergeData` store)
 * so the user can drop a «FieldName» placeholder for one. When no data source is
 * attached yet, it offers a free-text fallback so a field can still be authored
 * ahead of the data (Word lets you insert a field name that does not yet exist).
 *
 * Purely presentational + local-state: it never touches the editor directly (so it
 * stays testable in isolation); it only reports the chosen column upward. The name
 * is sanitized again at the command edge (see mergeFieldExtension.tsx), so this
 * form does no security-relevant scrubbing of its own.
 */

interface MergeFieldPickerProps {
  /** The data source's column names (may be empty before recipients are attached). */
  columns: string[];
  /** Live write authority; searching and closing remain available when false. */
  canInsert?: boolean;
  onPick: (field: string) => void;
  /** Insert a conditional (IF) field. The shape is re-validated at the command edge. */
  onInsertCondition: (condition: {
    column: string;
    op: MergeConditionOp;
    value: string;
    trueText: string;
    falseText: string;
  }) => void;
  onClose: () => void;
}

export default function MergeFieldPicker({
  columns,
  canInsert = true,
  onPick,
  onInsertCondition,
  onClose,
}: MergeFieldPickerProps) {
  const t = useTranslations("console-notebooks");
  const chromeTone = useOfficeDialogChrome();
  const [query, setQuery] = useState("");
  // Conditional (IF) builder local state.
  const [condColumn, setCondColumn] = useState("");
  const [condOp, setCondOp] = useState<MergeConditionOp>("eq");
  const [condValue, setCondValue] = useState("");
  const [condTrue, setCondTrue] = useState("");
  const [condFalse, setCondFalse] = useState("");
  const searchInputRef = useRef<HTMLInputElement | null>(null);
  useEffect(() => {
    let focusFrame = 0;
    const settleFrame = window.requestAnimationFrame(() => {
      focusFrame = window.requestAnimationFrame(() => {
        searchInputRef.current?.focus({ preventScroll: true });
      });
    });
    return () => {
      window.cancelAnimationFrame(settleFrame);
      window.cancelAnimationFrame(focusFrame);
    };
  }, []);
  const condNeedsValue = condOp !== "empty" && condOp !== "notEmpty";

  const filtered = useMemo(() => {
    const q = query.trim().toLowerCase();
    if (!q) return columns;
    return columns.filter((c) => c.toLowerCase().includes(q));
  }, [columns, query]);

  const custom = query.trim();

  return (
    <div
      className={`merge-field-picker ${officeDialogClassName({ tone: chromeTone ?? "document" })}`}
      role="dialog"
      aria-modal="true"
      aria-label={t("mailMerge.fieldPickerTitle")}
      data-testid="merge-field-picker"
      onKeyDown={(event) => {
        // aria-modal: Tab and Shift+Tab wrap inside, never reach the page.
        if (containDialogTab(event, event.currentTarget)) return;
        if (event.key !== "Escape") return;
        event.preventDefault();
        event.stopPropagation();
        onClose();
      }}
    >
      <div className={`merge-field-picker__head ${parts.header}`}>
        <span className={`merge-field-picker__heading ${parts.title}`}>
          <Database size={14} /> {t("mailMerge.fieldPickerTitle")}
        </span>
        <button
          type="button"
          onClick={onClose}
          className={`merge-field-picker__close ${parts.close}`}
          aria-label={t("mailMerge.close")}
          data-testid="merge-field-picker-close"
        >
          <X size={14} />
        </button>
      </div>

      <input
        ref={searchInputRef}
        type="text"
        value={query}
        onChange={(e) => setQuery(e.target.value)}
        placeholder={t("mailMerge.fieldSearchPlaceholder")}
        className="merge-field-picker__search"
        data-testid="merge-field-picker-search"
      />

      <ul className={`merge-field-picker__results ${parts.list}`}>
        {filtered.map((column) => (
          <li key={column}>
            <button
              type="button"
              className={`merge-field-picker__item ${parts.listItem}`}
              disabled={!canInsert}
              onClick={() => {
                if (canInsert) onPick(column);
              }}
              data-testid="merge-field-picker-item"
            >
              «{column}»
            </button>
          </li>
        ))}
        {filtered.length === 0 && columns.length > 0 && (
          <li
            className={`merge-field-picker__empty ${parts.hint}`}
            data-testid="merge-field-picker-empty"
          >
            {t("mailMerge.noColumnsMatch")}
          </li>
        )}
        {columns.length === 0 && (
          <li
            className={`merge-field-picker__empty ${parts.hint}`}
            data-testid="merge-field-picker-no-data"
          >
            {t("mailMerge.noRecipients")}
          </li>
        )}
      </ul>

      {/* Free-text fallback: author a field name even before the data is attached
          (or for a column not in the current source). Enabled once non-blank. */}
      {custom && !columns.includes(custom) && (
        <button
          type="button"
          className={`merge-field-picker__custom-btn ${parts.button}`}
          disabled={!canInsert}
          onClick={() => {
            if (canInsert) onPick(custom);
          }}
          data-testid="merge-field-picker-custom"
        >
          {t("mailMerge.insertCustomField", { name: custom })}
        </button>
      )}

      {/* Conditional (IF) field builder: emit one of two texts based on a column
          test for the merged record (Word's IF field). The shape is re-validated
          at the command edge, so this form does no security scrubbing of its own. */}
      <div
        className="merge-field-picker__condition"
        data-testid="merge-field-condition"
      >
        <div className={`merge-field-picker__heading ${parts.sectionTitle}`} style={{ marginTop: 10 }}>
          {t("mailMerge.conditionTitle")}
        </div>
        <input
          type="text"
          value={condColumn}
          onChange={(e) => setCondColumn(e.target.value)}
          placeholder={t("mailMerge.conditionColumn")}
          list="merge-condition-columns"
          className="merge-field-picker__search"
          data-testid="merge-condition-column"
          disabled={!canInsert}
        />
        <datalist id="merge-condition-columns">
          {columns.map((c) => (
            <option key={c} value={c} />
          ))}
        </datalist>
        <select
          value={condOp}
          onChange={(e) => setCondOp(e.target.value as MergeConditionOp)}
          className="merge-field-picker__search"
          data-testid="merge-condition-op"
          aria-label={t("mailMerge.conditionTitle")}
          disabled={!canInsert}
        >
          {MERGE_CONDITION_OPS.map((op) => (
            <option key={op} value={op}>
              {OP_SYMBOLS[op] ||
                t(`mailMerge.op_${op}` as "mailMerge.op_contains")}
            </option>
          ))}
        </select>
        {condNeedsValue ? (
          <input
            type="text"
            value={condValue}
            onChange={(e) => setCondValue(e.target.value)}
            placeholder={t("mailMerge.conditionValue")}
            className="merge-field-picker__search"
            data-testid="merge-condition-value"
            disabled={!canInsert}
          />
        ) : null}
        <input
          type="text"
          value={condTrue}
          onChange={(e) => setCondTrue(e.target.value)}
          placeholder={t("mailMerge.conditionThen")}
          className="merge-field-picker__search"
          data-testid="merge-condition-true"
          disabled={!canInsert}
        />
        <input
          type="text"
          value={condFalse}
          onChange={(e) => setCondFalse(e.target.value)}
          placeholder={t("mailMerge.conditionElse")}
          className="merge-field-picker__search"
          data-testid="merge-condition-false"
          disabled={!canInsert}
        />
        <button
          type="button"
          className={`merge-field-picker__custom-btn ${parts.button}`}
          disabled={!canInsert || !condColumn.trim()}
          onClick={() => {
            if (!canInsert) return;
            onInsertCondition({
              column: condColumn,
              op: condOp,
              value: condValue,
              trueText: condTrue,
              falseText: condFalse,
            });
          }}
          data-testid="merge-condition-insert"
        >
          <DocumentDialogIcon icon="code.braces" />
          {t("mailMerge.insertCondition")}
        </button>
      </div>
    </div>
  );
}
