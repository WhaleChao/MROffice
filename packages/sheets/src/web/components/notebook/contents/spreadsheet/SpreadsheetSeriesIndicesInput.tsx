"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { useRef, useState, type InputHTMLAttributes } from "react";
import { isSpreadsheetImeKeydown } from "../../../../lib/spreadsheet/ime-composition";

type Props = Omit<InputHTMLAttributes<HTMLInputElement>, "value" | "defaultValue" | "onChange" | "onBlur" | "onKeyDown"> & {
  indices: readonly number[];
  onCommit: (indices: number[]) => void;
};

/** Keep an in-progress comma list separate from the committed chart indices. */
export function SpreadsheetSeriesIndicesInput({ indices, onCommit, disabled, ...props }: Props) {
  const source = indices.map((index) => index + 1).join(", ");
  const [draft, setDraft] = useState({ source, text: source });
  const skipBlur = useRef(false);
  const value = draft.source === source ? draft.text : source;
  const reset = () => setDraft({ source, text: source });
  const commit = () => {
    if (disabled || value === source) return;
    const text = value.trim();
    if (text && !/^\d+(?:\s*,\s*\d+)*$/.test(text)) return;
    const numbers = text ? text.split(",").map((item) => Number(item.trim())) : [];
    if (numbers.some((number) => !Number.isSafeInteger(number) || number < 1)) return;
    onCommit([...new Set(numbers.map((number) => number - 1))]);
  };
  return <input {...props} disabled={disabled} value={value}
    onChange={(event) => { if (!disabled) setDraft({ source, text: event.currentTarget.value }); }}
    onBlur={() => { if (!skipBlur.current) commit(); skipBlur.current = false; reset(); }}
    onKeyDown={(event) => {
      if (isSpreadsheetImeKeydown(event.nativeEvent)) return;
      if (event.key === "Enter") { event.preventDefault(); skipBlur.current = true; commit(); event.currentTarget.blur(); }
      else if (event.key === "Escape") { event.preventDefault(); event.stopPropagation(); reset(); skipBlur.current = true; event.currentTarget.blur(); }
    }}/>
}
