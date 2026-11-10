"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { useRef, useState } from "react";
import { useTranslations } from "@mrnote/core/host/next-intl.tsx";

import { OfficeDialogShell } from "@mrnote/core/web/components/office-ribbon/OfficeDialogShell.tsx";
import { officeDialogParts } from "@mrnote/core/web/components/office-ribbon/officeDialogChrome.tsx";
import {
  MAX_CONTROL_ITEM_LENGTH,
  MAX_CONTROL_LABEL_LENGTH,
  MAX_CONTROL_LIST_ITEMS,
  MAX_RADIO_GROUP_LENGTH,
  RADIO_CHECKED_GLYPH,
  RADIO_UNCHECKED_GLYPH,
  type ContentControlCheckboxGlyph,
  type ContentControlListItem,
  type ContentControlMeta,
} from "@mrnote/core/web/components/console/editor/native-document/contentControls.ts";

/**
 * Word's Content Control Properties (Developer > Properties; Microsoft
 * Support, "Content controls in Word"): Title and Tag; Remove content control
 * when contents are edited; the two locks (cannot be deleted, contents cannot
 * be edited); and the type's own options -- a plain text control's line
 * breaks, a check box's two symbols, a list's entries, a date's format.
 */

const CHECKED_SYMBOLS: readonly ContentControlCheckboxGlyph[] = [
  { val: "2612", font: "MS Gothic" },
  { val: "2611", font: "Segoe UI Symbol" },
  { val: "2713", font: "Segoe UI Symbol" },
  { val: "2714", font: "Segoe UI Symbol" },
  { val: "25A0", font: "Segoe UI Symbol" },
];
const UNCHECKED_SYMBOLS: readonly ContentControlCheckboxGlyph[] = [
  { val: "2610", font: "MS Gothic" },
  { val: "25A1", font: "Segoe UI Symbol" },
];
const DATE_FORMATS = [
  "yyyy/M/d",
  "yyyy年M月d日",
  "yyyy-MM-dd",
  "M/d/yyyy",
  "d MMMM yyyy",
  "dddd, MMMM d, yyyy",
] as const;

const glyphText = (glyph: ContentControlCheckboxGlyph) =>
  String.fromCodePoint(Number.parseInt(glyph.val, 16));

export interface DocumentContentControlPropertiesDialogProps {
  readonly meta: ContentControlMeta;
  readonly canEdit: boolean;
  readonly onApply: (patch: Partial<ContentControlMeta>) => boolean;
  readonly onCancel: () => void;
}

export function DocumentContentControlPropertiesDialog({
  meta,
  canEdit,
  onApply,
  onCancel,
}: DocumentContentControlPropertiesDialogProps) {
  const t = useTranslations("console-notebooks");
  const titleRef = useRef<HTMLInputElement | null>(null);
  const [alias, setAlias] = useState(meta.alias);
  const [tag, setTag] = useState(meta.tag);
  const [temporary, setTemporary] = useState(meta.temporary);
  const [noDelete, setNoDelete] = useState(meta.lock === "sdtLocked" || meta.lock === "sdtContentLocked");
  const [noEdit, setNoEdit] = useState(meta.lock === "contentLocked" || meta.lock === "sdtContentLocked");
  const [multiline, setMultiline] = useState(meta.multiline);
  const [checked, setChecked] = useState<ContentControlCheckboxGlyph>(meta.checkedState ?? CHECKED_SYMBOLS[0]!);
  const [unchecked, setUnchecked] = useState<ContentControlCheckboxGlyph>(meta.uncheckedState ?? UNCHECKED_SYMBOLS[0]!);
  const [items, setItems] = useState<ContentControlListItem[]>(meta.listItems);
  const [selected, setSelected] = useState(0);
  const [dateFormat, setDateFormat] = useState(meta.dateFormat || DATE_FORMATS[0]);
  const [required, setRequired] = useState(meta.required);
  const [radioGroup, setRadioGroup] = useState(meta.radioGroup);
  const type = meta.controlType;
  const list = type === "dropDownList" || type === "comboBox";

  const apply = () => {
    const lock = noDelete && noEdit ? "sdtContentLocked" : noDelete ? "sdtLocked" : noEdit ? "contentLocked" : null;
    const patch: Partial<ContentControlMeta> = {
      alias: alias.trim().slice(0, MAX_CONTROL_LABEL_LENGTH),
      tag: tag.trim().slice(0, MAX_CONTROL_LABEL_LENGTH),
      temporary,
      lock,
      required,
      ...(type === "text" ? { multiline } : {}),
      ...(type === "checkbox"
        ? radioGroup.trim()
          ? {
              // An option button keeps Word's ◉/○ symbols.
              radioGroup: radioGroup.trim().slice(0, MAX_RADIO_GROUP_LENGTH),
              checkedState: meta.radioGroup ? checked : RADIO_CHECKED_GLYPH,
              uncheckedState: meta.radioGroup ? unchecked : RADIO_UNCHECKED_GLYPH,
            }
          : { radioGroup: "", checkedState: checked, uncheckedState: unchecked }
        : {}),
      ...(list
        ? {
            listItems: items
              .map((item) => ({
                displayText: item.displayText.trim().slice(0, MAX_CONTROL_ITEM_LENGTH),
                value: item.value.trim().slice(0, MAX_CONTROL_ITEM_LENGTH),
              }))
              .filter((item) => item.displayText),
          }
        : {}),
      ...(type === "date" ? { dateFormat: dateFormat.trim().slice(0, MAX_CONTROL_ITEM_LENGTH) } : {}),
    };
    if (onApply(patch)) onCancel();
  };

  const move = (by: -1 | 1) => {
    const target = selected + by;
    if (target < 0 || target >= items.length) return;
    const next = [...items];
    [next[selected], next[target]] = [next[target]!, next[selected]!];
    setItems(next);
    setSelected(target);
  };

  return (
    <OfficeDialogShell
      open
      title={t("contentControl.properties.title")}
      description=""
      defaultActionLabel={t("contentControl.properties.ok")}
      cancelActionLabel={t("contentControl.properties.cancel")}
      defaultActionDisabled={!canEdit}
      onDefaultAction={apply}
      onCancel={onCancel}
      initialFocusRef={titleRef}
      testIdPrefix="document-content-control-properties"
    >
      <section className={officeDialogParts.group}>
        <h3 className={officeDialogParts.sectionTitle}>{t("contentControl.properties.general")}</h3>
        <label className={officeDialogParts.field}>
          <span className={officeDialogParts.label}>{t("contentControl.properties.titleField")}</span>
          <input
            ref={titleRef}
            type="text"
            value={alias}
            maxLength={MAX_CONTROL_LABEL_LENGTH}
            data-testid="document-content-control-properties-alias"
            onChange={(event) => setAlias(event.target.value)}
          />
        </label>
        <label className={officeDialogParts.field}>
          <span className={officeDialogParts.label}>{t("contentControl.properties.tag")}</span>
          <input
            type="text"
            value={tag}
            maxLength={MAX_CONTROL_LABEL_LENGTH}
            data-testid="document-content-control-properties-tag"
            onChange={(event) => setTag(event.target.value)}
          />
        </label>
        <label className={officeDialogParts.check}>
          <input type="checkbox" checked={temporary} onChange={(event) => setTemporary(event.target.checked)} />
          {t("contentControl.properties.temporary")}
        </label>
      </section>
      {type !== "picture" ? (
        <section className={officeDialogParts.group}>
          <h3 className={officeDialogParts.sectionTitle}>{t("contentControl.properties.form")}</h3>
          <label className={officeDialogParts.check}>
            <input
              type="checkbox"
              checked={required}
              data-testid="document-content-control-properties-required"
              onChange={(event) => setRequired(event.target.checked)}
            />
            {t("contentControl.properties.required")}
          </label>
          {type === "checkbox" ? (
            <label className={officeDialogParts.field}>
              <span className={officeDialogParts.label}>{t("contentControl.properties.radioGroup")}</span>
              <input
                type="text"
                value={radioGroup}
                maxLength={MAX_RADIO_GROUP_LENGTH}
                data-testid="document-content-control-properties-radio-group"
                onChange={(event) => setRadioGroup(event.target.value)}
              />
            </label>
          ) : null}
        </section>
      ) : null}
      <section className={officeDialogParts.group}>
        <h3 className={officeDialogParts.sectionTitle}>{t("contentControl.properties.locking")}</h3>
        <label className={officeDialogParts.check}>
          <input
            type="checkbox"
            checked={noDelete}
            data-testid="document-content-control-properties-no-delete"
            onChange={(event) => setNoDelete(event.target.checked)}
          />
          {t("contentControl.properties.cannotDelete")}
        </label>
        <label className={officeDialogParts.check}>
          <input
            type="checkbox"
            checked={noEdit}
            data-testid="document-content-control-properties-no-edit"
            onChange={(event) => setNoEdit(event.target.checked)}
          />
          {t("contentControl.properties.cannotEdit")}
        </label>
      </section>
      {type === "text" ? (
        <section className={officeDialogParts.group}>
          <h3 className={officeDialogParts.sectionTitle}>{t("contentControl.properties.plainText")}</h3>
          <label className={officeDialogParts.check}>
            <input type="checkbox" checked={multiline} onChange={(event) => setMultiline(event.target.checked)} />
            {t("contentControl.properties.multiline")}
          </label>
        </section>
      ) : null}
      {type === "checkbox" ? (
        <section className={officeDialogParts.group}>
          <h3 className={officeDialogParts.sectionTitle}>{t("contentControl.properties.checkBox")}</h3>
          {(
            [
              ["checkedSymbol", CHECKED_SYMBOLS, checked, setChecked],
              ["uncheckedSymbol", UNCHECKED_SYMBOLS, unchecked, setUnchecked],
            ] as const
          ).map(([key, symbols, value, set]) => (
            <label key={key} className={officeDialogParts.row}>
              <span className={officeDialogParts.label}>{t(`contentControl.properties.${key}`)}</span>
              <select
                value={value.val}
                data-testid={`document-content-control-properties-${key}`}
                onChange={(event) => {
                  const glyph = symbols.find((symbol) => symbol.val === event.target.value);
                  if (glyph) set(glyph);
                }}
              >
                {symbols.map((symbol) => (
                  <option key={symbol.val} value={symbol.val}>
                    {glyphText(symbol)}
                  </option>
                ))}
              </select>
            </label>
          ))}
        </section>
      ) : null}
      {list ? (
        <section className={officeDialogParts.group}>
          <h3 className={officeDialogParts.sectionTitle}>{t("contentControl.properties.listItems")}</h3>
          <ul className={officeDialogParts.list} role="listbox" data-testid="document-content-control-properties-items">
            {items.map((item, index) => (
              <li key={index} role="option" aria-selected={index === selected} onClick={() => setSelected(index)}>
                <input
                  type="text"
                  value={item.displayText}
                  aria-label={t("contentControl.properties.displayName")}
                  maxLength={MAX_CONTROL_ITEM_LENGTH}
                  onFocus={() => setSelected(index)}
                  onChange={(event) =>
                    setItems(items.map((entry, at) => (at === index ? { ...entry, displayText: event.target.value } : entry)))
                  }
                />
                <input
                  type="text"
                  value={item.value}
                  aria-label={t("contentControl.properties.value")}
                  maxLength={MAX_CONTROL_ITEM_LENGTH}
                  onFocus={() => setSelected(index)}
                  onChange={(event) =>
                    setItems(items.map((entry, at) => (at === index ? { ...entry, value: event.target.value } : entry)))
                  }
                />
              </li>
            ))}
          </ul>
          <div className={officeDialogParts.row}>
            <button
              type="button"
              className={officeDialogParts.button}
              data-testid="document-content-control-properties-add"
              disabled={items.length >= MAX_CONTROL_LIST_ITEMS}
              onClick={() => {
                setItems([...items, { displayText: "", value: "" }]);
                setSelected(items.length);
              }}
            >
              {t("contentControl.properties.add")}
            </button>
            <button
              type="button"
              className={officeDialogParts.button}
              disabled={!items.length}
              onClick={() => {
                setItems(items.filter((_, at) => at !== selected));
                setSelected(Math.max(0, selected - 1));
              }}
            >
              {t("contentControl.properties.remove")}
            </button>
            <button type="button" className={officeDialogParts.button} disabled={selected <= 0} onClick={() => move(-1)}>
              {t("contentControl.properties.moveUp")}
            </button>
            <button
              type="button"
              className={officeDialogParts.button}
              disabled={selected >= items.length - 1}
              onClick={() => move(1)}
            >
              {t("contentControl.properties.moveDown")}
            </button>
          </div>
        </section>
      ) : null}
      {type === "date" ? (
        <section className={officeDialogParts.group}>
          <h3 className={officeDialogParts.sectionTitle}>{t("contentControl.properties.datePicker")}</h3>
          <label className={officeDialogParts.field}>
            <span className={officeDialogParts.label}>{t("contentControl.properties.dateFormat")}</span>
            <input
              type="text"
              list="document-content-control-date-formats"
              value={dateFormat}
              maxLength={MAX_CONTROL_ITEM_LENGTH}
              data-testid="document-content-control-properties-date-format"
              onChange={(event) => setDateFormat(event.target.value)}
            />
            <datalist id="document-content-control-date-formats">
              {DATE_FORMATS.map((format) => (
                <option key={format} value={format} />
              ))}
            </datalist>
          </label>
        </section>
      ) : null}
    </OfficeDialogShell>
  );
}
