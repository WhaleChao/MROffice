"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Word's Layout dialog (More Layout Options) for one floating drawing --
 * a chart, a picture, a shape or a text box: the Position tab (each axis by
 * Alignment, Book layout, Absolute or Relative position, against the origin
 * Word offers for it; Move object with text, Lock anchor, Allow overlap,
 * Layout in table cell) and the Text Wrapping tab (the style, the sides the
 * text may use, the distance from the text). OK writes one anchor
 * (`floatAnchorFromLayoutForm`); switching an origin keeps the drawing where
 * it is, as Word restates the number. The Size tab sets the height and the
 * width, the other side following when Lock aspect ratio is on.
 */

import { useEffect, useState, type ReactElement, type RefObject } from "react";

import { OfficeDialogShell } from "./OfficeDialogShell";
import { officeDialogParts as dialogParts } from "./officeDialogChrome";
import type { FloatAnchor, FloatWrapText } from "../../lib/office-float-anchor";
import {
  FLOAT_LAYOUT_ORIGINS_H,
  FLOAT_LAYOUT_ORIGINS_V,
  FLOAT_WRAP_CHOICES,
  floatAnchorFromLayoutForm,
  floatLayoutFormOf,
  floatPaintedOffsetEmu,
  type FloatAxisMode,
  type FloatLayoutForm,
  type FloatPaint,
  type FloatWrapChoice,
} from "../../lib/office-float-anchor-edit";

const EMU_PER_CM = 360000;

export interface OfficeFloatLayoutDialogLabels {
  readonly title: string;
  readonly ok: string;
  readonly cancel: string;
  readonly positionTab: string;
  readonly wrappingTab: string;
  readonly sizeTab: string;
  readonly height: string;
  readonly width: string;
  readonly lockAspectRatio: string;
  readonly horizontal: string;
  readonly vertical: string;
  readonly alignment: string;
  readonly bookLayout: string;
  readonly absolutePosition: string;
  readonly relativePosition: string;
  readonly relativeTo: string;
  readonly toTheRightOf: string;
  readonly below: string;
  readonly of: string;
  readonly options: string;
  readonly moveWithText: string;
  readonly lockAnchor: string;
  readonly allowOverlap: string;
  readonly layoutInCell: string;
  readonly wrappingStyle: string;
  readonly wrapText: string;
  readonly distanceFromText: string;
  readonly top: string;
  readonly bottom: string;
  readonly left: string;
  readonly right: string;
  readonly cm: string;
  /** Every option name: alignments, origins, wraps, wrap sides. */
  readonly option: (key: string) => string;
}

function cm(emu: number): string {
  return String(Math.round((emu / EMU_PER_CM) * 100) / 100);
}

function emuOf(text: string, fallback: number): number {
  const value = Number(text.trim().replace(",", "."));
  return Number.isFinite(value) ? Math.round(value * EMU_PER_CM) : fallback;
}

/** The drawing's size as the Size tab reads and writes it. */
export interface OfficeFloatLayoutSize {
  readonly widthEmu: number;
  readonly heightEmu: number;
  readonly lockAspectRatio: boolean;
}

export function OfficeFloatLayoutDialog({
  open,
  anchor,
  paint,
  size = null,
  labels,
  onApply,
  onCancel,
  returnFocusRef,
  testIdPrefix,
}: {
  readonly open: boolean;
  readonly anchor: FloatAnchor | null;
  readonly paint: FloatPaint | null;
  /** Absent: no Size tab. */
  readonly size?: OfficeFloatLayoutSize | null;
  readonly labels: OfficeFloatLayoutDialogLabels;
  /** `size` is the Size tab's, when it has one. */
  readonly onApply: (anchor: FloatAnchor, size: OfficeFloatLayoutSize | null) => void;
  readonly onCancel: () => void;
  readonly returnFocusRef?: RefObject<HTMLElement | null>;
  readonly testIdPrefix: string;
}): ReactElement | null {
  const [tab, setTab] = useState<"position" | "wrapping" | "size">("position");
  const [form, setForm] = useState<FloatLayoutForm | null>(null);
  const [extent, setExtent] = useState<OfficeFloatLayoutSize | null>(null);
  useEffect(() => {
    if (open && anchor) setForm(floatLayoutFormOf(anchor, paint));
    if (open) setExtent(size);
    if (open) setTab("position");
  }, [open, anchor, paint, size]);
  if (!anchor || !form) return null;

  const id = (suffix: string) => `${testIdPrefix}-${suffix}`;
  const sided = form.wrap === "square" || form.wrap === "tight" || form.wrap === "through";
  const overlay = form.wrap === "behindText" || form.wrap === "inFrontOfText";
  const moveWithText = form.vertical.relativeFrom === "paragraph" || form.vertical.relativeFrom === "line";

  const axisSection = (which: "horizontal" | "vertical") => {
    const value = form[which];
    const origins = which === "horizontal" ? FLOAT_LAYOUT_ORIGINS_H : FLOAT_LAYOUT_ORIGINS_V;
    const alignments =
      which === "horizontal"
        ? (["left", "center", "right"] as const)
        : (["top", "center", "bottom"] as const);
    const setAxis = (patch: Partial<typeof value>) =>
      setForm((current) => (current ? { ...current, [which]: { ...current[which], ...patch } } : current));
    const setMode = (mode: FloatAxisMode) => {
      const list = origins[mode];
      const relativeFrom = (list as readonly string[]).includes(value.relativeFrom) ? value.relativeFrom : list[0];
      setAxis({
        mode,
        relativeFrom: relativeFrom as never,
        align: (mode === "book"
          ? "inside"
          : value.align === "inside" || value.align === "outside"
            ? alignments[0]
            : value.align) as never,
        offsetEmu: floatPaintedOffsetEmu(paint, which, relativeFrom) ?? value.offsetEmu,
      });
    };
    const setOrigin = (relativeFrom: string) =>
      // Word restates the number on another origin so the drawing stays put.
      setAxis({
        relativeFrom: relativeFrom as never,
        offsetEmu: value.mode === "absolute" ? (floatPaintedOffsetEmu(paint, which, relativeFrom) ?? value.offsetEmu) : value.offsetEmu,
      });
    const originSelect = (mode: FloatAxisMode) => (
      <select
        value={value.relativeFrom}
        disabled={value.mode !== mode}
        onChange={(event) => setOrigin(event.target.value)}
        data-testid={id(`${which}-${mode}-origin`)}
      >
        {origins[mode].map((origin) => (
          <option key={origin} value={origin}>
            {labels.option(`origin.${which === "horizontal" ? "h" : "v"}.${origin}`)}
          </option>
        ))}
      </select>
    );
    const radio = (mode: FloatAxisMode, label: string) => (
      <label className={dialogParts.check}>
        <input
          type="radio"
          name={id(`${which}-mode`)}
          checked={value.mode === mode}
          onChange={() => setMode(mode)}
          data-testid={id(`${which}-${mode}`)}
        />
        {label}
      </label>
    );
    const row = { display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap" as const };
    return (
      <fieldset style={{ border: "none", margin: 0, padding: 0, display: "grid", gap: 6 }}>
        <legend style={{ fontWeight: 600, marginBottom: 4 }}>{labels[which]}</legend>
        <div style={row}>
          {radio("align", labels.alignment)}
          <select
            value={value.mode === "align" ? value.align : alignments[0]}
            disabled={value.mode !== "align"}
            onChange={(event) => setAxis({ align: event.target.value as never })}
            data-testid={id(`${which}-align-value`)}
          >
            {alignments.map((align) => (
              <option key={align} value={align}>
                {labels.option(`align.${which === "horizontal" ? "h" : "v"}.${align}`)}
              </option>
            ))}
          </select>
          {labels.relativeTo}
          {originSelect("align")}
        </div>
        <div style={row}>
          {radio("book", labels.bookLayout)}
          <select
            value={value.mode === "book" ? value.align : "inside"}
            disabled={value.mode !== "book"}
            onChange={(event) => setAxis({ align: event.target.value as never })}
            data-testid={id(`${which}-book-value`)}
          >
            {(["inside", "outside"] as const).map((align) => (
              <option key={align} value={align}>
                {labels.option(`align.book.${align}`)}
              </option>
            ))}
          </select>
          {labels.of}
          {originSelect("book")}
        </div>
        <div style={row}>
          {radio("absolute", labels.absolutePosition)}
          <input
            type="number"
            step={0.01}
            style={{ width: 72 }}
            value={cm(value.offsetEmu)}
            disabled={value.mode !== "absolute"}
            onChange={(event) => setAxis({ offsetEmu: emuOf(event.target.value, value.offsetEmu) })}
            data-testid={id(`${which}-absolute-value`)}
          />
          {labels.cm} {which === "horizontal" ? labels.toTheRightOf : labels.below}
          {originSelect("absolute")}
        </div>
        <div style={row}>
          {radio("relative", labels.relativePosition)}
          <input
            type="number"
            step={1}
            style={{ width: 72 }}
            value={String(Math.round(value.percent * 10) / 10)}
            disabled={value.mode !== "relative"}
            onChange={(event) => {
              const percent = Number(event.target.value);
              if (Number.isFinite(percent)) setAxis({ percent });
            }}
            data-testid={id(`${which}-relative-value`)}
          />
          % {labels.relativeTo}
          {originSelect("relative")}
        </div>
      </fieldset>
    );
  };

  const check = (key: "locked" | "allowOverlap" | "layoutInCell", label: string) => (
    <label className={dialogParts.check}>
      <input
        type="checkbox"
        checked={form[key]}
        onChange={(event) => setForm({ ...form, [key]: event.target.checked })}
        data-testid={id(key)}
      />
      {label}
    </label>
  );
  const distance = (key: "distTEmu" | "distBEmu" | "distLEmu" | "distREmu", label: string, enabled: boolean) => (
    <label className={dialogParts.field} style={{ display: "flex", alignItems: "center", gap: 6 }}>
      {label}
      <input
        type="number"
        step={0.01}
        min={0}
        style={{ width: 72 }}
        value={cm(form[key])}
        disabled={!enabled}
        onChange={(event) => setForm({ ...form, [key]: Math.max(0, emuOf(event.target.value, form[key])) })}
        data-testid={id(key)}
      />
      {labels.cm}
    </label>
  );

  return (
    <OfficeDialogShell
      open={open}
      title={labels.title}
      description={null}
      defaultActionLabel={labels.ok}
      cancelActionLabel={labels.cancel}
      onDefaultAction={() => onApply(floatAnchorFromLayoutForm(anchor, form), extent)}
      onCancel={onCancel}
      returnFocusRef={returnFocusRef}
      testIdPrefix={testIdPrefix}
      className="sm:max-w-[560px]"
    >
      <div role="tablist" style={{ display: "flex", gap: 4, marginBottom: 8 }}>
        {(extent ? (["position", "wrapping", "size"] as const) : (["position", "wrapping"] as const)).map((name) => (
          <button
            key={name}
            type="button"
            role="tab"
            aria-selected={tab === name}
            onClick={() => setTab(name)}
            data-testid={id(`tab-${name}`)}
            style={{
              padding: "4px 10px",
              borderRadius: 6,
              border: "1px solid var(--office-dialog-border, #d0d5dd)",
              background: tab === name ? "var(--office-dialog-selected, #e8f0fe)" : "transparent",
            }}
          >
            {name === "position" ? labels.positionTab : name === "wrapping" ? labels.wrappingTab : labels.sizeTab}
          </button>
        ))}
      </div>
      {tab === "size" && extent ? (
        <div style={{ display: "grid", gap: 8 }}>
          {(["heightEmu", "widthEmu"] as const).map((side) => (
            <label key={side} className={dialogParts.field} style={{ display: "flex", alignItems: "center", gap: 6 }}>
              {side === "heightEmu" ? labels.height : labels.width}
              <input
                type="number"
                step={0.01}
                min={0.01}
                style={{ width: 80 }}
                value={cm(extent[side])}
                onChange={(event) => {
                  const next = emuOf(event.target.value, extent[side]);
                  if (!(next > 0)) return;
                  // Locked, the other side keeps the proportions (Word).
                  const other = side === "heightEmu" ? "widthEmu" : "heightEmu";
                  const scaled = extent.lockAspectRatio && extent[side] > 0
                    ? Math.max(1, Math.round((extent[other] * next) / extent[side]))
                    : extent[other];
                  setExtent({ ...extent, [side]: next, [other]: scaled });
                }}
                data-testid={id(`size-${side === "heightEmu" ? "height" : "width"}`)}
              />
              {labels.cm}
            </label>
          ))}
          <label className={dialogParts.check}>
            <input
              type="checkbox"
              checked={extent.lockAspectRatio}
              onChange={(event) => setExtent({ ...extent, lockAspectRatio: event.target.checked })}
              data-testid={id("size-lock-aspect")}
            />
            {labels.lockAspectRatio}
          </label>
        </div>
      ) : tab === "position" ? (
        <div style={{ display: "grid", gap: 12 }}>
          {axisSection("horizontal")}
          {axisSection("vertical")}
          <fieldset style={{ border: "none", margin: 0, padding: 0, display: "grid", gap: 4 }}>
            <legend style={{ fontWeight: 600, marginBottom: 4 }}>{labels.options}</legend>
            <label className={dialogParts.check}>
              <input
                type="checkbox"
                checked={moveWithText}
                disabled={form.vertical.mode !== "absolute"}
                onChange={(event) => {
                  // Word: the vertical origin becomes the paragraph, or the
                  // page, and the number is restated so the drawing stays put.
                  const relativeFrom = event.target.checked ? "paragraph" : "page";
                  setForm({
                    ...form,
                    vertical: {
                      ...form.vertical,
                      relativeFrom,
                      offsetEmu: floatPaintedOffsetEmu(paint, "vertical", relativeFrom) ?? form.vertical.offsetEmu,
                    },
                  });
                }}
                data-testid={id("move-with-text")}
              />
              {labels.moveWithText}
            </label>
            {check("locked", labels.lockAnchor)}
            {check("allowOverlap", labels.allowOverlap)}
            {check("layoutInCell", labels.layoutInCell)}
          </fieldset>
        </div>
      ) : (
        <div style={{ display: "grid", gap: 12 }}>
          <fieldset style={{ border: "none", margin: 0, padding: 0, display: "grid", gap: 4 }}>
            <legend style={{ fontWeight: 600, marginBottom: 4 }}>{labels.wrappingStyle}</legend>
            {FLOAT_WRAP_CHOICES.map((wrap: FloatWrapChoice) => (
              <label key={wrap} className={dialogParts.check}>
                <input
                  type="radio"
                  name={id("wrap")}
                  checked={form.wrap === wrap}
                  onChange={() => setForm({ ...form, wrap })}
                  data-testid={id(`wrap-${wrap}`)}
                />
                {labels.option(`wrap.${wrap}`)}
              </label>
            ))}
          </fieldset>
          <fieldset style={{ border: "none", margin: 0, padding: 0, display: "flex", gap: 12, flexWrap: "wrap" }}>
            <legend style={{ fontWeight: 600, marginBottom: 4 }}>{labels.wrapText}</legend>
            {(["bothSides", "left", "right", "largest"] as const).map((side: FloatWrapText) => (
              <label key={side} className={dialogParts.check}>
                <input
                  type="radio"
                  name={id("wrap-text")}
                  checked={form.wrapText === side}
                  disabled={!sided}
                  onChange={() => setForm({ ...form, wrapText: side })}
                  data-testid={id(`wrap-text-${side}`)}
                />
                {labels.option(`wrapText.${side}`)}
              </label>
            ))}
          </fieldset>
          <fieldset style={{ border: "none", margin: 0, padding: 0, display: "grid", gridTemplateColumns: "1fr 1fr", gap: 6 }}>
            <legend style={{ fontWeight: 600, marginBottom: 4 }}>{labels.distanceFromText}</legend>
            {distance("distTEmu", labels.top, !overlay && (form.wrap === "square" || form.wrap === "topAndBottom"))}
            {distance("distLEmu", labels.left, sided)}
            {distance("distBEmu", labels.bottom, !overlay && (form.wrap === "square" || form.wrap === "topAndBottom"))}
            {distance("distREmu", labels.right, sided)}
          </fieldset>
        </div>
      )}
    </OfficeDialogShell>
  );
}
