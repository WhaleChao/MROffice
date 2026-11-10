"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { useMemo, useRef, useState, type RefObject } from "react";

import { OfficeDialogShell } from "@mrnote/core/web/components/office-ribbon/OfficeDialogShell.tsx";
import { officeDialogParts } from "@mrnote/core/web/components/office-ribbon/officeDialogChrome.tsx";
import {
  MAX_TABLE_CELL_MARGIN,
  tableCellMarginPresetPatch,
} from "@mrnote/core/web/lib/presentation-table-style.ts";
import {
  TABLE_FRAME_EMU_PER_CSS_PX,
  tableAxisCmToPx,
  tableAxisPxToCm,
} from "@mrnote/core/web/lib/presentation-table-size.ts";
import {
  formatPresentationMeasurement,
  parsePresentationMeasurementCm,
  presentationMeasurementStep,
  type PresentationMeasurementUnit,
} from "@mrnote/core/web/lib/presentation-measurement.ts";

export const TABLE_CELL_MARGIN_FIELDS = [
  "marginTop",
  "marginBottom",
  "marginLeft",
  "marginRight",
] as const;

export type TableCellMarginField = (typeof TABLE_CELL_MARGIN_FIELDS)[number];

export type TableCellMarginDialogSelection = Readonly<
  Record<
    TableCellMarginField,
    Readonly<{
      mixed: boolean;
      value: number | null;
    }>
  >
>;

export type TableCellMarginDialogPatch = Partial<
  Record<TableCellMarginField, number>
>;

export type TableCellMarginDialogLabels = Readonly<{
  apply: string;
  cancel: string;
  description: string;
  mixed: string;
  sides: Readonly<Record<TableCellMarginField, string>>;
  title: string;
  /** The unit word after each field: "厘米" / "cm", or `"` for inches. */
  unit: string;
}>;

const normalTableCellMargins = tableCellMarginPresetPatch("normal");
const DEFAULT_TABLE_CELL_MARGINS = {
  marginTop:
    typeof normalTableCellMargins.marginTop === "number"
      ? normalTableCellMargins.marginTop
      : 0,
  marginBottom:
    typeof normalTableCellMargins.marginBottom === "number"
      ? normalTableCellMargins.marginBottom
      : 0,
  marginLeft:
    typeof normalTableCellMargins.marginLeft === "number"
      ? normalTableCellMargins.marginLeft
      : 0,
  marginRight:
    typeof normalTableCellMargins.marginRight === "number"
      ? normalTableCellMargins.marginRight
      : 0,
} satisfies Record<TableCellMarginField, number>;

/** A margin stored in px as the field shows it: "0.25" cm, or "0.1" inch. */
function marginDisplay(px: number, unit: PresentationMeasurementUnit): string {
  return unit === "in"
    ? formatPresentationMeasurement(px * TABLE_FRAME_EMU_PER_CSS_PX, "in")
    : String(tableAxisPxToCm(px));
}

function initialDrafts(
  selection: TableCellMarginDialogSelection,
  unit: PresentationMeasurementUnit,
): Record<TableCellMarginField, string> {
  return Object.fromEntries(
    TABLE_CELL_MARGIN_FIELDS.map((field) => [
      field,
      selection[field].mixed
        ? ""
        : marginDisplay(
            selection[field].value ?? DEFAULT_TABLE_CELL_MARGINS[field],
            unit,
          ),
    ]),
  ) as Record<TableCellMarginField, string>;
}

/** The margin a draft names (in the field's unit, or its own), in the model's px. */
function parseTableCellMargin(
  value: string,
  unit: PresentationMeasurementUnit,
): number | null {
  const centimetres = parsePresentationMeasurementCm(value, unit);
  if (centimetres === null) return null;
  const parsed = tableAxisCmToPx(centimetres);
  if (parsed < 0 || parsed > MAX_TABLE_CELL_MARGIN) return null;
  return parsed;
}

export function TableMarginsDialog({
  labels,
  measurementUnit = "cm",
  onApply,
  onCancel,
  open,
  returnFocusRef,
  selection,
}: {
  labels: TableCellMarginDialogLabels;
  /** PowerPoint's system measurement unit (en: inches). */
  measurementUnit?: PresentationMeasurementUnit;
  onApply: (patch: TableCellMarginDialogPatch) => void;
  onCancel: () => void;
  open: boolean;
  returnFocusRef: RefObject<HTMLButtonElement | null>;
  selection: TableCellMarginDialogSelection;
}) {
  const [initial] = useState(() => initialDrafts(selection, measurementUnit));
  const [drafts, setDrafts] = useState(initial);
  const [dirtyFields, setDirtyFields] = useState<
    ReadonlySet<TableCellMarginField>
  >(() => new Set());
  const firstInputRef = useRef<HTMLInputElement | null>(null);
  const parsedDirtyValues = useMemo(() => {
    const parsed: TableCellMarginDialogPatch = {};
    let valid = true;
    for (const field of dirtyFields) {
      // A field typed back to the centimetres it showed is unchanged; the
      // cm -> px round trip must not nudge the stored px (9.6 -> 9.4).
      if (!selection[field].mixed && drafts[field].trim() === initial[field]) continue;
      const value = parseTableCellMargin(drafts[field], measurementUnit);
      if (value === null) {
        valid = false;
        continue;
      }
      const initialValue =
        selection[field].value ?? DEFAULT_TABLE_CELL_MARGINS[field];
      if (!selection[field].mixed && Object.is(value, initialValue)) continue;
      parsed[field] = value;
    }
    return { parsed, valid };
  }, [dirtyFields, drafts, initial, measurementUnit, selection]);

  return (
    <OfficeDialogShell
      open={open}
      title={labels.title}
      description={labels.description}
      defaultActionLabel={labels.apply}
      cancelActionLabel={labels.cancel}
      onDefaultAction={() => {
        if (!parsedDirtyValues.valid) return;
        onApply(parsedDirtyValues.parsed);
      }}
      onCancel={onCancel}
      defaultActionDisabled={!parsedDirtyValues.valid}
      initialFocusRef={firstInputRef}
      returnFocusRef={returnFocusRef}
      testIdPrefix="presentation-ribbon-table-cell-margins-custom"
    >
      <div className="grid grid-cols-2 gap-3">
        {TABLE_CELL_MARGIN_FIELDS.map((field, index) => {
          const dirty = dirtyFields.has(field);
          const invalid =
            dirty && parseTableCellMargin(drafts[field], measurementUnit) === null;
          const mixed = selection[field].mixed && !dirty;
          const label = labels.sides[field];
          return (
            <label key={field} className={officeDialogParts.field}>
              <span className={officeDialogParts.label}>{label}</span>
              <input
                ref={index === 0 ? firstInputRef : undefined}
                type="number"
                inputMode="decimal"
                min={0}
                max={Number(marginDisplay(MAX_TABLE_CELL_MARGIN, measurementUnit))}
                step={presentationMeasurementStep(measurementUnit) / 10}
                value={drafts[field]}
                placeholder={mixed ? labels.mixed : undefined}
                aria-label={
                  mixed
                    ? `${label} (${labels.unit}): ${labels.mixed}`
                    : `${label} (${labels.unit})`
                }
                aria-invalid={invalid || undefined}
                data-mixed={mixed || undefined}
                data-testid={`presentation-ribbon-table-cell-margins-custom-${field}`}
                onChange={(event) => {
                  const value = event.target.value;
                  setDrafts((current) => ({ ...current, [field]: value }));
                  setDirtyFields((current) => {
                    const next = new Set(current);
                    next.add(field);
                    return next;
                  });
                }}
              />
              <span aria-hidden>{labels.unit}</span>
            </label>
          );
        })}
      </div>
    </OfficeDialogShell>
  );
}
