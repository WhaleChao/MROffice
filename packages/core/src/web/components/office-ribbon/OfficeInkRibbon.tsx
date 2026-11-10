"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The Draw tab's pens as Office draws them, for any editor that inks: the
 * gallery of pens (the picked one in a raised well wearing its drop arrow),
 * each pen's menu (sample line, six thickness dots, sixteen colours, eight
 * effects for a pen, More Colors, Delete), the Add Pen menu (Pen,
 * Highlighter, Pencil), the Eraser's kinds and the Draw with Trackpad and
 * Draw with Touch switches. The editor owns the pens and what a pick does; these draw them.
 *
 * Word's and Excel's Draw tabs share these menus ("Draw and write with ink in
 * Office", Microsoft Support); the three editors draw their Draw tabs with
 * these, so a pen is the same pen everywhere.
 */

import { ptToPx } from "../../lib/office-units";
import { type ReactNode, useId, useState } from "react";
import { Palette, Trash2 } from "./icons/houseIcons";
import { useTranslations } from "../../../host/next-intl";

import DocumentInkPenGlyph from "../console/editor/native-document/DocumentInkPenGlyph";
import { IconDrawWithTouch, IconInkHideRibbon } from "../notebook/contents/PresentationRibbonIcons";
import { RibbonDropdownMenuContent as DropdownMenuContent } from "./CupertinoRibbonMenu";
import {
  DropdownMenu,
  DropdownMenuItem,
  DropdownMenuLabel,
  DropdownMenuRadioGroup,
  DropdownMenuRadioItem,
  DropdownMenuSeparator,
  DropdownMenuTrigger,
} from "../ui/dropdown-menu";
import { OfficeSemanticIcon } from "./icons/OfficeSemanticIcon";
import {
  DOCUMENT_INK_COLORS,
  DOCUMENT_INK_ERASERS,
  DOCUMENT_INK_THICKNESSES,
  DOCUMENT_INK_TOOLS,
  documentInkEffectPaint,
  documentInkPenDefaults,
  type DocumentInkEraser,
  type DocumentInkPreset,
  type DocumentInkTool,
} from "../../lib/document-ink";

import { saveDrawWithTouch, useDrawWithTouch, useDeviceHasTouch } from "../../lib/draw-with-touch";
import { lockInkTrackpadPointer } from "../../lib/ink-trackpad-lock";
import {
  OfficeHomeCommand,
  OfficeHomeLargeSplitLayout,
  OfficeHomeMenuTrigger,
} from "./OfficeHomeRibbon";

import { InkEffectStroke } from "./InkEffectStroke";
import { OfficeInkEffectSection, officeInkEffectLabel } from "./OfficeInkEffectSection";

import type { InkPenPatch } from "../../lib/office-ink-pen-dynamics";
import {
  OfficeInkPenGesturesPage,
  OfficeInkPenLevels,
  OfficeInkPenSettings,
  OfficeInkPenTypeRow,
} from "./OfficeInkPenFeel";
import styles from "./OfficeInkRibbon.module.css";

/** A pen menu's change to one pen (colour, width, effect, Pressure, feel). */
export type OfficeInkPenPatch = InkPenPatch;

export interface OfficeInkPenGalleryProps {
  pens: readonly DocumentInkPreset[];
  selectedId: string;
  /** Draw is armed: the picked pen reads as pressed. */
  drawing: boolean;
  disabled?: boolean;
  disabledReason?: string;
  /** `${testIdPrefix}-${pen.id}` names each pen, and `-menu…` its menu. */
  testIdPrefix: string;
  /** The editor's own control name for a pen, when it has one. */
  controlIdFor?: (pen: DocumentInkPreset) => string | undefined;
  onChoose: (pen: DocumentInkPreset) => void;
  /** The picked pen's menu (thickness, colours, effects, More Colors,
   * Delete). An editor whose pens are fixed presets leaves these out and
   * the pen wears no drop arrow. */
  onUpdate?: (id: string, patch: OfficeInkPenPatch) => void;
  onRemove?: (id: string) => void;
  onMoreColors?: (id: string) => void;
  /** Where each pen's face is, for a colour palette anchored to it. */
  penRef?: (id: string, element: HTMLButtonElement | null) => void;
  /** An editor whose ink cannot paint the eight effects leaves them out of
   * the pen menu (default: shown). */
  effects?: boolean;
  /** An editor whose ink always follows the pen's pressure leaves the
   * pressure choice out of the pen menu (default: shown). */
  pressure?: boolean;
  /** The pen's name as the editor's Office says it (PowerPoint: "Pen: Black,
   * 0.5 mm"); default tool · colour · points. */
  penLabel?: (pen: DocumentInkPreset) => string;
}

export function OfficeInkPenGallery({
  pens,
  selectedId,
  drawing,
  disabled = false,
  disabledReason,
  testIdPrefix,
  controlIdFor,
  onChoose,
  onUpdate,
  onRemove,
  onMoreColors,
  penRef,
  effects = true,
  pressure = true,
  penLabel: nameOf,
}: OfficeInkPenGalleryProps) {
  const t = useTranslations("console-notebooks");
  const groupLabel = t("toolbar.ribbonDrawPens");
  const hasMenu = Boolean(onUpdate && onRemove);
  return (
    <div
      className={styles.inkGallery}
      role="group"
      aria-label={groupLabel}
      data-testid={`${testIdPrefix}-gallery`}
    >
      {pens.map((pen) => {
        // A pen reads as the picked one only while Draw is armed: with Draw
        // off nothing is drawing, so no pen is highlighted and none wears
        // its menu arrow (Word: choosing a pen turns Draw on and lights it).
        const selected = drawing && selectedId === pen.id;
        const toolLabel = t(`toolbar.drawTool_${pen.tool}`);
        const penLabel = nameOf
          ? nameOf(pen)
          : pen.effect
            ? `${toolLabel} · ${officeInkEffectLabel(t, pen.effect)} · ${pen.widthPt}pt`
            : `${toolLabel} · ${pen.color} · ${pen.widthPt}pt`;
        return (
          <div
            key={pen.id}
            className={styles.inkPen}
            data-selected={selected || undefined}
            data-testid={`${testIdPrefix}-${pen.id}-split`}
          >
            <button
              type="button"
              className={`office-ribbon-control ${styles.inkPenFace}`}
              aria-label={penLabel}
              aria-pressed={selected}
              title={disabledReason ? `${penLabel} (${disabledReason})` : penLabel}
              disabled={disabled}
              data-control-id={controlIdFor?.(pen)}
              data-ink-tool={pen.tool}
              data-testid={`${testIdPrefix}-${pen.id}`}
              ref={(element) => penRef?.(pen.id, element)}
              onClick={() => onChoose(pen)}
            >
              <DocumentInkPenGlyph preset={pen} />
            </button>
            {selected && hasMenu ? (
              <DropdownMenu modal={false}>
                <DropdownMenuTrigger asChild>
                  <button
                    type="button"
                    className={`office-ribbon-control ${styles.inkPenMenuTrigger}`}
                    aria-label={`${toolLabel}: ${t("toolbar.drawPenOptions")}`}
                    title={t("toolbar.drawPenOptions")}
                    disabled={disabled}
                    data-testid={`${testIdPrefix}-${pen.id}-menu`}
                  >
                    <OfficeSemanticIcon semanticId="menu.chevron.down" size={10} />
                  </button>
                </DropdownMenuTrigger>
                <OfficeInkPenMenuContent
                  pen={pen}
                  penCount={pens.length}
                  testIdPrefix={testIdPrefix}
                  effects={effects}
                  pressure={pressure}
                  onUpdate={onUpdate}
                  onRemove={onRemove}
                  onMoreColors={onMoreColors}
                />
              </DropdownMenu>
            ) : null}
          </div>
        );
      })}
    </div>
  );
}

export interface OfficeInkPenMenuContentProps {
  pen: DocumentInkPreset;
  /** How many pens there are: the last one cannot be deleted. */
  penCount: number;
  /** `${testIdPrefix}-${pen.id}-…` names the menu and its items. */
  testIdPrefix: string;
  effects?: boolean;
  /** Show the pressure choice for a pen or pencil (default: shown). */
  pressure?: boolean;
  onUpdate?: (id: string, patch: OfficeInkPenPatch) => void;
  onRemove?: (id: string) => void;
  onMoreColors?: (id: string) => void;
  /** Where the menu opens from its pen (default: below, Radix's). */
  side?: "top" | "right" | "bottom" | "left";
  align?: "start" | "center" | "end";
}

/**
 * A pen's menu (sample line, six thickness dots, sixteen colours, eight
 * effects for a pen, More Colors, Delete): the drop-down under the picked pen
 * in the gallery, and what a long press on a pen in the pen palette
 * (OfficePenPalette) opens.
 */
export function OfficeInkPenMenuContent({
  pen,
  penCount,
  testIdPrefix,
  effects = true,
  pressure = true,
  onUpdate,
  onRemove,
  onMoreColors,
  side,
  align = "start",
}: OfficeInkPenMenuContentProps) {
  const t = useTranslations("console-notebooks");
  // Pen Gestures is a page of the menu, with a way back to the pen.
  const [page, setPage] = useState<"pen" | "gestures">("pen");
  return (
    <DropdownMenuContent
      align={align}
      side={side}
      sideOffset={side ? 10 : undefined}
      collisionPadding={side ? 8 : undefined}
      className={styles.inkPenMenu}
      data-testid={`${testIdPrefix}-${pen.id}-menu-content`}
      onCloseAutoFocus={(event) => event.preventDefault()}
      // A custom effect's colour palette (OfficeInkEffectSection) opens beside
      // the menu: a press or focus in it is not one outside the menu.
      onInteractOutside={(event) => {
        const target = event.target;
        if (target instanceof Element && target.closest(".office-color-palette")) event.preventDefault();
      }}
    >
      {page === "gestures" ? (
        <OfficeInkPenGesturesPage testIdPrefix={`${testIdPrefix}-${pen.id}`} onBack={() => setPage("pen")} />
      ) : (
      <>
      <div className={styles.inkPenMenuTitle}>
        {pen.tool === "pen" && pen.penType ? t(`toolbar.drawPenType_${pen.penType}`) : t(`toolbar.drawTool_${pen.tool}`)}
      </div>
      <OfficeInkPenSample pen={pen} pressure={pressure} />
      <OfficeInkPenTypeRow pen={pen} testIdPrefix={`${testIdPrefix}-${pen.id}`} onUpdate={onUpdate} />
      <DropdownMenuSeparator />
      <div
        className={styles.inkThicknessRow}
        role="radiogroup"
        aria-label={t("toolbar.drawInkWidth")}
      >
        {DOCUMENT_INK_THICKNESSES.map((thickness, index) => (
          <button
            key={thickness.mm}
            type="button"
            role="radio"
            aria-checked={pen.widthPt === thickness.pt}
            aria-label={`${thickness.mm} mm`}
            title={`${thickness.mm} mm`}
            className={styles.inkThicknessDot}
            data-selected={pen.widthPt === thickness.pt || undefined}
            data-testid={`${testIdPrefix}-${pen.id}-thickness-${thickness.mm}`}
            onClick={() => onUpdate?.(pen.id, { widthPt: thickness.pt })}
          >
            <span style={{ width: 5 + index * 4, height: 5 + index * 4 }} />
          </button>
        ))}
      </div>
      <DropdownMenuSeparator />
      <DropdownMenuLabel className={styles.inkMenuLabel}>{t("toolbar.drawInkColor")}</DropdownMenuLabel>
      <div
        className={styles.inkColorGrid}
        role="group"
        aria-label={t("toolbar.drawInkColor")}
      >
        {DOCUMENT_INK_COLORS.map((color) => (
          <button
            key={color}
            type="button"
            className={styles.inkColorSwatch}
            data-selected={(!pen.effect && pen.color === color) || undefined}
            data-testid={`${testIdPrefix}-${pen.id}-color-${color.slice(1)}`}
            aria-label={color}
            aria-pressed={!pen.effect && pen.color === color}
            title={color}
            style={{ background: color }}
            onClick={() => onUpdate?.(pen.id, { color })}
          />
        ))}
      </div>
      {pen.tool === "pen" && effects ? (
        <OfficeInkEffectSection
          pen={pen}
          testIdPrefix={`${testIdPrefix}-${pen.id}`}
          onPick={(effect) => onUpdate?.(pen.id, { effect })}
        />
      ) : null}
      <OfficeInkPenLevels
        pen={pen}
        pressure={pressure}
        testIdPrefix={`${testIdPrefix}-${pen.id}`}
        onUpdate={onUpdate}
      />
      <OfficeInkPenSettings testIdPrefix={`${testIdPrefix}-${pen.id}`} onOpenGestures={() => setPage("gestures")} />
      <DropdownMenuSeparator />
      {onMoreColors ? (
        <DropdownMenuItem
          data-testid={`${testIdPrefix}-${pen.id}-more-colors`}
          onSelect={() => onMoreColors(pen.id)}
        >
          <Palette size={16} aria-hidden />
          {t("toolbar.drawMoreColors")}
        </DropdownMenuItem>
      ) : null}
      <DropdownMenuItem
        disabled={penCount <= 1}
        data-testid={`${testIdPrefix}-${pen.id}-delete`}
        onSelect={() => onRemove?.(pen.id)}
      >
        <Trash2 size={16} aria-hidden />
        {t("toolbar.drawDeletePen")}
      </DropdownMenuItem>
      </>
      )}
    </DropdownMenuContent>
  );
}

/** Add Pen's menu: a Pen, a Highlighter or a Pencil, each in its first colours. */
export function OfficeInkAddPenMenuItems({
  testIdPrefix,
  onAdd,
  tools = DOCUMENT_INK_TOOLS,
}: {
  testIdPrefix: string;
  onAdd: (tool: DocumentInkTool) => void;
  /** The tools the editor's ink can draw (default: all three). */
  tools?: readonly DocumentInkTool[];
}) {
  const t = useTranslations("console-notebooks");
  return (
    <>
      {tools.map((tool) => (
        <DropdownMenuItem
          key={tool}
          data-testid={`${testIdPrefix}-${tool}`}
          onSelect={() => onAdd(tool)}
        >
          <span aria-hidden="true" className={styles.inkAddPenGlyph}>
            <DocumentInkPenGlyph preset={{ id: tool, ...documentInkPenDefaults(tool) }} size={0.8} />
          </span>
          {t(`toolbar.drawTool_${tool}`)}
        </DropdownMenuItem>
      ))}
    </>
  );
}

/** The Eraser's kinds: Stroke Eraser, Small Eraser, Medium Eraser. */
export function OfficeInkEraserMenuItems({
  value,
  testIdPrefix,
  onChoose,
}: {
  value: DocumentInkEraser;
  testIdPrefix: string;
  onChoose: (kind: DocumentInkEraser) => void;
}) {
  const t = useTranslations("console-notebooks");
  return (
    <DropdownMenuRadioGroup value={value}>
      {DOCUMENT_INK_ERASERS.map((kind) => (
        <DropdownMenuRadioItem
          key={kind}
          value={kind}
          data-testid={`${testIdPrefix}-${kind}`}
          onSelect={() => onChoose(kind)}
        >
          <span aria-hidden="true" className={styles.inkEraserSample} data-eraser={kind} />
          {t(`toolbar.drawEraser_${kind}`)}
        </DropdownMenuRadioItem>
      ))}
    </DropdownMenuRadioGroup>
  );
}

/** Draw with Trackpad: a switch showing its state, the command's name below. */
export function OfficeInkTrackpadSwitch(props: {
  on: boolean;
  disabled?: boolean;
  disabledReason?: string;
  controlId?: string;
  testId: string;
  /** Defaults to Office's "Draw with Trackpad". */
  label?: ReactNode;
  onToggle: () => void;
}) {
  const t = useTranslations("console-notebooks");
  // A touch screen draws with a finger instead (OfficeInkTouchSwitch).
  const hasTouch = useDeviceHasTouch();
  if (hasTouch) return null;
  const { on, onToggle } = props;
  return (
    <OfficeInkSwitch
      {...props}
      name={t("toolbar.drawTrackpad")}
      onToggle={() => {
        // The pointer locks in this click, the only place a browser allows
        // it (lib/ink-trackpad-lock); the mode's frame takes it over.
        if (!on && !props.disabled) lockInkTrackpadPointer();
        onToggle();
      }}
    />
  );
}

/**
 * Draw with Touch (Office's Draw tab): on (the default on a touch device,
 * as Office ships it), a finger inks with the picked pen as the pen does;
 * off, the pen alone inks and a finger scrolls and pinches ("Draw and write
 * with ink in Office", Microsoft Support). The state is the
 * device's own (lib/draw-with-touch): every editor that puts this switch on
 * its Draw tab shows and sets the same one, and its ink surface asks
 * `inkAcceptsPointer` at each press.
 */
export function OfficeInkTouchSwitch({
  disabled,
  disabledReason,
  controlId,
  testId,
}: {
  disabled?: boolean;
  disabledReason?: string;
  controlId?: string;
  testId: string;
}) {
  const t = useTranslations("console-notebooks");
  const on = useDrawWithTouch();
  // Nothing to switch without a touch screen (a Mac draws with its trackpad).
  const hasTouch = useDeviceHasTouch();
  if (!hasTouch) return null;
  return (
    <OfficeInkSwitch
      on={on}
      name={t("toolbar.drawTouch")}
      disabled={disabled}
      disabledReason={disabledReason}
      controlId={controlId}
      testId={testId}
      onToggle={() => saveDrawWithTouch(!on)}
    />
  );
}










export function useInkDrawTools(inkDrawing: boolean): {
  /** The device has a touch screen: the Draw with Touch tool is offered. */
  readonly touchTool: boolean;
  /** Draw (the pen alone) is the lit tool. */
  readonly penDrawActive: boolean;
  /** Draw with Touch is the lit tool. */
  readonly touchDrawActive: boolean;
  /** Choose Draw: on a touch screen the finger stops inking. */
  readonly choosePenDraw: () => void;
  /** Choose Draw with Touch: the finger inks. */
  readonly chooseTouchDraw: () => void;
} {
  const on = useDrawWithTouch();
  const touchTool = useDeviceHasTouch();
  return {
    touchTool,
    penDrawActive: inkDrawing && !(touchTool && on),
    touchDrawActive: touchTool && inkDrawing && on,
    choosePenDraw: () => {
      if (touchTool) saveDrawWithTouch(false);
    },
    chooseTouchDraw: () => saveDrawWithTouch(true),
  };
}

/** The Draw with Touch tool of a Draw tab's tools group (`useInkDrawTools`). */
export function OfficeInkTouchDrawCommand({
  inkDrawing,
  disabled,
  disabledReason,
  controlId,
  testId,
  onArm,
  onExit,
}: {
  /** The editor's draw mode is on (with either drawing tool). */
  inkDrawing: boolean;
  disabled?: boolean;
  disabledReason?: string;
  controlId?: string;
  testId: string;
  /** Turn the editor's draw mode on. */
  onArm: () => void;
  /** Back to Select. */
  onExit: () => void;
}) {
  const t = useTranslations("console-notebooks");
  const tools = useInkDrawTools(inkDrawing);
  if (!tools.touchTool) return null;
  return (
    <OfficeHomeCommand
      active={tools.touchDrawActive}
      controlId={controlId}
      variant="large"
      disabled={disabled}
      disabledReason={disabledReason}
      icon={<IconDrawWithTouch home size={24} />}
      label={t("toolbar.drawTouch")}
      onClick={() => {
        if (tools.touchDrawActive) {
          onExit();
          return;
        }
        tools.chooseTouchDraw();
        if (!inkDrawing) onArm();
      }}
      testId={testId}
    />
  );
}








export function OfficeInkHideRibbonCommand({
  controlId,
  testId,
  onHide,
}: {
  controlId?: string;
  testId: string;
  onHide: () => void;
}) {
  const t = useTranslations("console-notebooks");
  return (
    <OfficeHomeCommand
      controlId={controlId}
      variant="large"
      icon={<IconInkHideRibbon home size={24} />}
      label={t("toolbar.penPaletteHideRibbon")}
      onClick={onHide}
      testId={testId}
    />
  );
}

/** A Draw-tab switch: its state beside the track, the command's name below. */
function OfficeInkSwitch({
  on,
  name,
  disabled = false,
  disabledReason,
  controlId,
  testId,
  label,
  onToggle,
}: {
  on: boolean;
  name: string;
  disabled?: boolean;
  disabledReason?: string;
  controlId?: string;
  testId: string;
  label?: ReactNode;
  onToggle: () => void;
}) {
  const t = useTranslations("console-notebooks");
  return (
    <span className={styles.inkSwitchHost}>
      <button
        type="button"
        role="switch"
        aria-checked={on}
        aria-label={name}
        className={`office-ribbon-control ${styles.inkSwitch}`}
        data-control-id={controlId}
        data-testid={testId}
        disabled={disabled}
        title={disabledReason ?? name}
        onClick={onToggle}
      >
        <span className={styles.inkSwitchRow}>
          <span className={styles.inkSwitchTrack} data-on={on || undefined}>
            <span className={styles.inkSwitchKnob} />
          </span>
          <span className={styles.inkSwitchState}>
            {on ? t("toolbar.drawTrackpadOn") : t("toolbar.drawTrackpadOff")}
          </span>
        </span>
        <span className={styles.inkSwitchLabel}>{label ?? name}</span>
      </button>
    </span>
  );
}


/** The pen menu's sample: one large wave across a 244x96 box, dipping
 *  then sweeping up into a wide arc and down again. */
const PEN_SAMPLE_CURVE: ReadonlyArray<readonly [number, number]> = Array.from({ length: 97 }, (_, index) => {
  const t = index / 96;
  return [16 + t * 212, 46 + Math.sin(t * Math.PI * 1.9 + 0.25) * 28] as const;
});
const PEN_SAMPLE_LINE = `M${PEN_SAMPLE_CURVE.map(([x, y]) => `${x.toFixed(2)},${y.toFixed(2)}`).join(" L")}`;

/** The curve as a pressed pen draws it: thin where it lands and lifts. */
function penSampleOutline(widthPx: number): string {
  const left: string[] = [];
  const right: string[] = [];
  PEN_SAMPLE_CURVE.forEach(([x, y], index) => {
    const [px, py] = PEN_SAMPLE_CURVE[Math.max(0, index - 1)];
    const [nx, ny] = PEN_SAMPLE_CURVE[Math.min(PEN_SAMPLE_CURVE.length - 1, index + 1)];
    const length = Math.hypot(nx - px, ny - py) || 1;
    // Pressure builds after the pen lands and eases off before it lifts.
    const w = (widthPx / 2) * (0.12 + 0.88 * Math.pow(Math.sin(Math.PI * (index / (PEN_SAMPLE_CURVE.length - 1))), 0.85));
    left.push(`${(x - ((ny - py) / length) * w).toFixed(2)},${(y + ((nx - px) / length) * w).toFixed(2)}`);
    right.push(`${(x + ((ny - py) / length) * w).toFixed(2)},${(y - ((nx - px) / length) * w).toFixed(2)}`);
  });
  return `M${left.join(" L")} L${right.reverse().join(" L")} Z`;
}

/** The pen's line as it will draw: its colour or effect, width and pressure. */
function OfficeInkPenSample({ pen, pressure }: { pen: DocumentInkPreset; pressure: boolean }) {
  const gradientId = `ink-pen-sample-${useId().replace(/:/g, "")}`;
  const highlighter = pen.tool === "highlighter";
  // Drawn larger than life, so every step of thickness reads at a glance.
  const widthPx = Math.max(2.5, Math.min(highlighter ? 20 : 16, ptToPx(pen.widthPt) * 2.4));
  const paint = pen.effect ? `url(#${gradientId})` : pen.color;
  // With pressure the line swells and tapers; without, one width throughout.
  const swells = pressure && !highlighter && pen.steady !== true;
  if (pen.tool === "pen" && pen.effect) {
    // The effect's own layers (a halo, a sheen, glints), as a stroke lands.
    const count = PEN_SAMPLE_CURVE.length - 1;
    const points = PEN_SAMPLE_CURVE.map(([x, y], index) => ({
      x,
      y,
      ...(swells ? { p: 0.12 + 0.88 * Math.pow(Math.sin((Math.PI * index) / count), 0.85) } : {}),
    }));
    return (
      <div className={styles.inkPenSample} aria-hidden="true">
        <svg viewBox="0 0 244 96" width="244" height="96" focusable="false">
          <InkEffectStroke
            effect={pen.effect}
            gradientId={gradientId}
            geometry={{ points, widthPx, pressured: swells }}
          />
        </svg>
      </div>
    );
  }
  return (
    <div className={styles.inkPenSample} aria-hidden="true">
      <svg viewBox="0 0 244 96" width="244" height="96" focusable="false">
        {pen.effect ? (
          <defs>
            <linearGradient id={gradientId} x1="0" x2="1" y1="0" y2="0">
              {documentInkEffectPaint(pen.effect).stops.map((stop, index, stops) => (
                <stop key={index} offset={stops.length > 1 ? index / (stops.length - 1) : 0} stopColor={stop} />
              ))}
            </linearGradient>
          </defs>
        ) : null}
        {swells ? (
          <path d={penSampleOutline(widthPx * 2.1)} fill={paint} opacity={pen.tool === "pencil" ? 0.82 : 1} />
        ) : (
          <path
            d={PEN_SAMPLE_LINE}
            fill="none"
            stroke={paint}
            // Inline: the menu's icon rule sets every svg line's width.
            style={{ strokeWidth: widthPx }}
            strokeLinecap={highlighter ? "butt" : "round"}
            strokeLinejoin="round"
            opacity={highlighter ? 0.45 : pen.tool === "pencil" ? 0.82 : 1}
          />
        )}
      </svg>
    </div>
  );
}

/** A squiggle's centre line, for the pressure tiles (a 44x24 box). */
const PRESSURE_SQUIGGLE: ReadonlyArray<readonly [number, number]> = Array.from({ length: 49 }, (_, index) => {
  const t = index / 48;
  return [5 + t * 34, 12 + Math.sin(t * Math.PI * 2.5) * 5.5] as const;
});

/**
 * The squiggle as a pressed pen draws it: a filled outline swelling from a
 * hairline at the ends to 6 px where the pen presses hardest, with a second,
 * lighter swell, so it reads at a glance against the even line.
 */
const PRESSURE_SQUIGGLE_OUTLINE: string = (() => {
  const halfWidth = (t: number) => 0.35 + 2.75 * Math.pow(Math.sin(Math.PI * t), 1.4) * (0.75 + 0.25 * Math.cos(t * Math.PI * 4));
  const left: string[] = [];
  const right: string[] = [];
  PRESSURE_SQUIGGLE.forEach(([x, y], index) => {
    const [px, py] = PRESSURE_SQUIGGLE[Math.max(0, index - 1)];
    const [nx, ny] = PRESSURE_SQUIGGLE[Math.min(PRESSURE_SQUIGGLE.length - 1, index + 1)];
    const dx = nx - px;
    const dy = ny - py;
    const length = Math.hypot(dx, dy) || 1;
    const w = halfWidth(index / (PRESSURE_SQUIGGLE.length - 1));
    left.push(`${(x - (dy / length) * w).toFixed(2)},${(y + (dx / length) * w).toFixed(2)}`);
    right.push(`${(x + (dy / length) * w).toFixed(2)},${(y - (dx / length) * w).toFixed(2)}`);
  });
  return `M${left.join(" L")} L${right.reverse().join(" L")} Z`;
})();

/**
 * The pen's pressure setting, drawn as two samples of the same squiggle
 * (as a note-taking app's pen offers): with pressure, the line swells where
 * the pen presses; without, one width throughout (a ballpoint's line).
 * Shared by every editor's pen menu (Docs, Slides and Sheets).
 */
export function OfficeInkPressureChoice({
  steady,
  color,
  testIdPrefix,
  onChange,
}: {
  /** True when the pen draws one width regardless of pressure. */
  steady: boolean;
  /** The pen's colour for the samples; the current text colour when absent. */
  color?: string;
  testIdPrefix: string;
  onChange: (steady: boolean) => void;
}) {
  const t = useTranslations("console-notebooks");
  const stroke = color ?? "currentColor";
  const choices = [
    { steady: false, label: t("toolbar.drawPenPressureOn"), testId: `${testIdPrefix}-pressure` },
    { steady: true, label: t("toolbar.drawPenPressureOff"), testId: `${testIdPrefix}-pressure-off` },
  ];
  return (
    <div className={styles.inkPressureRow} role="radiogroup" aria-label={t("toolbar.drawPenPressure")}>
      {choices.map((choice) => (
        <button
          key={String(choice.steady)}
          type="button"
          role="radio"
          aria-checked={steady === choice.steady}
          aria-label={choice.label}
          title={choice.label}
          className={styles.inkPressureTile}
          data-selected={steady === choice.steady || undefined}
          data-testid={choice.testId}
          onClick={() => onChange(choice.steady)}
        >
          <svg viewBox="0 0 44 24" width="44" height="24" aria-hidden="true" focusable="false">
            {choice.steady ? (
              <polyline
                points={PRESSURE_SQUIGGLE.map(([x, y]) => `${x.toFixed(2)},${y.toFixed(2)}`).join(" ")}
                fill="none"
                stroke={stroke}
                strokeWidth={1.6}
                strokeLinecap="round"
                strokeLinejoin="round"
              />
            ) : (
              <path d={PRESSURE_SQUIGGLE_OUTLINE} fill={stroke} stroke={stroke} strokeWidth={0.4} strokeLinejoin="round" />
            )}
          </svg>
        </button>
      ))}
    </div>
  );
}
