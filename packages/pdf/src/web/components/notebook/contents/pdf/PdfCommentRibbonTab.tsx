"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { Type } from "@mrnote/core/web/components/office-ribbon/icons/houseIcons.tsx";
import { type ReactNode, useRef, useState } from "react";
import { Hexagon, Ruler, SquareDashed } from "lucide-react";
import { useTranslations } from "@mrnote/core/host/next-intl.tsx";

import { OfficeInkTouchSwitch } from "@mrnote/core/web/components/office-ribbon/OfficeInkRibbon.tsx";
import { useDeviceHasTouch } from "@mrnote/core/web/lib/draw-with-touch.ts";
import { colorHueName } from "@mrnote/sheets/web/lib/spreadsheet/color-hue-name.ts";

import { RibbonDropdownMenuContent } from "@mrnote/core/web/components/office-ribbon/CupertinoRibbonMenu.tsx";
import { PdfRibbonChoiceMenu } from "./PdfRibbonChoiceMenu";
import {
  OfficeHomeCheckbox,
  OfficeHomeCommand,
  OfficeHomeMenuTrigger,
  OfficeHomeSplitLayout,
} from "@mrnote/core/web/components/office-ribbon/OfficeHomeRibbon.tsx";
import { OfficeRibbonNumberField } from "@mrnote/core/web/components/office-ribbon/OfficeRibbonNumberField.tsx";
import { OfficeColorPalettePopover } from "@mrnote/core/web/components/office-ribbon/OfficeColorPalettePopover.tsx";
import { OfficeSemanticIcon } from "@mrnote/core/web/components/office-ribbon/icons/OfficeSemanticIcon.tsx";
import {
  DropdownMenu,
  DropdownMenuItem,
  DropdownMenuLabel,
  DropdownMenuRadioGroup,
  DropdownMenuRadioItem,
  DropdownMenuSeparator,
  DropdownMenuTrigger,
} from "@mrnote/core/web/components/ui/dropdown-menu.tsx";
import {
  SHAPE_LINE_ENDINGS,
  SHAPE_LINE_STYLES,
  type ShapeLineEnding,
  type ShapeLineStyle,
} from "@mrnote/core/web/lib/pdf/shapes.ts";

import anatomyStyles from "./PdfInsertRibbonTab.module.css";
import { PdfSurfaceRibbonGroup } from "./PdfSurfaceRibbon";
import PdfCommentRibbonIcon, {
  type PdfCommentRibbonIconName,
} from "./PdfCommentRibbonIcons";
import { OfficeHomeCommandStack } from "@mrnote/core/web/components/office-ribbon/OfficeHomeRibbon.tsx";
import PdfRibbonButton from "./PdfRibbonButton";
import ribbonLayout from "./pdfRibbonLayout.module.css";

export type { PdfCommentRibbonIconName };

export interface PdfCommentRibbonCommand {
  active?: boolean;
  disabledAccessibleLabel?: string;
  disabled?: boolean;
  disabledReason?: string;
  displayLabel?: string;
  icon: PdfCommentRibbonIconName | "textComment";
  label: string;
  onInvoke?: (trigger?: HTMLButtonElement | null) => void;
  testId: string;
}

interface PaintMenu {
  activeColor: string;
  activeWidth: number;
  colorLabel: string;
  colors: readonly string[];
  onColorChange: (color: string) => void;
  onWidthChange: (width: number) => void;
  undo?: PdfCommentRibbonCommand;
  widthLabel: string;
  widths: readonly number[];
}

interface ColorMenu {
  activeColor: string;
  colorLabel: string;
  colors: readonly string[];
  onColorChange: (color: string) => void;
}

interface AreaHighlightMenu extends ColorMenu {
  activeOpacity: number;
  onOpacityChange: (opacity: number) => void;
  opacities: readonly number[];
  opacityLabel: string;
}

interface ShapeMenu {
  activeColor: string;
  activeWidth: number;
  colorLabel: string;
  colors: readonly string[];
  kinds: readonly PdfCommentRibbonCommand[];
  onColorChange: (color: string) => void;
  onWidthChange: (width: number) => void;
  widthLabel: string;
  /** The widths' unit, shown beside each ("磅"): points, as Shape editing. */
  widthUnit?: string;
  widths: readonly number[];
}

interface PdfCommentRibbonTabProps {
  comments: {
    commentMode: PdfCommentRibbonCommand;
    hide: PdfCommentRibbonCommand;
    manage: PdfCommentRibbonCommand;
  };
  drawing: {
    draw: PdfCommentRibbonCommand;
    drawMenu: PaintMenu;
    note: PdfCommentRibbonCommand;
    noteMenu: ColorMenu;
    shape: PdfCommentRibbonCommand;
    shapeMenu: ShapeMenu;
    /** WPS's 图章 among the drawing tools (CT_Postil TMpxa3). */
    stamp?: PdfCommentRibbonCommand;
    textComment: PdfCommentRibbonCommand;
  };
  /** WPS's 测量 group (CT_Postil MeasureTool): the tool, on show. */
  measure?: { label: string; command: PdfCommentRibbonCommand };
  drawingContext?: {
    disabledReason?: string;
    fillColor: string | null;
    fillDisabledReason?: string;
    fillLabel: string;
    fillSupported: boolean;
    kind: "ink" | "measurement" | "shape";
    label: string;
    lineWidth: number;
    /** The field's range in its own unit (points for a shape). */
    lineWidthMin?: number;
    lineWidthMax?: number;
    lineWidthLabel: string;
    lineStyle: ShapeLineStyle;
    lineStyleLabel: string;
    lineStyleLabels: Record<ShapeLineStyle, string>;
    lineEndingsSupported: boolean;
    lineStart: ShapeLineEnding;
    lineStartLabel: string;
    lineEnd: ShapeLineEnding;
    lineEndLabel: string;
    lineEndingLabels: Record<ShapeLineEnding, string>;
    onFillColorChange: (color: string | null) => void;
    onLineWidthChange: (width: number) => void;
    onLineStyleChange: (style: ShapeLineStyle) => void;
    onLineStartChange: (ending: ShapeLineEnding) => void;
    onLineEndChange: (ending: ShapeLineEnding) => void;
    onOutlineColorChange: (color: string) => void;
    outlineColor: string;
    outlineColorLabel: string;
    styleEnabled: boolean;
    /** Caption of the settings group (形状样式 / 样式), beside the tool group. */
    styleLabel: string;
    





    measurement?: {
      kind: string;
      kindLabel: string;
      kindOptions: ReadonlyArray<{ value: string; label: string }>;
      onKindChange: (kind: string) => void;
      unit: string;
      unitLabel: string;
      units: ReadonlyArray<string>;
      onUnitChange: (unit: string) => void;
      ratio: number;
      ratioLabel: string;
      onRatioChange: (ratio: number) => void;
      /** Keep the measurement drawn and remove the last one kept, in the
       * tab the tool works in (user audit PDF 0010). */
      keep?: PdfCommentRibbonCommand;
      removeLast?: PdfCommentRibbonCommand;
      /** WPS's snapping (CT_MeasureTool PDFAdsorptionEndpoint, ...Path,
       * ...Midpoint, ...Intersection): what a click snaps to. */
      snap?: {
        label: string;
        options: ReadonlyArray<{ kind: string; label: string; checked: boolean; onToggle: () => void }>;
      };
      /** WPS's 结果面板 (CT_MeasureTool Measure_ResultPanel): every
       * measurement kept, each opening its page. */
      results?: {
        label: string;
        emptyLabel: string;
        items: ReadonlyArray<{ id: string; text: string; page: number }>;
        onOpen: (page: number) => void;
      };
    };
    /** The ink tool's Eraser (WPS CT_Graffiti RB_Graffiti_Eraser). */
    eraser?: PdfCommentRibbonCommand;
    /** The ink tool's line type: freehand or straight (WPS CT_Graffiti
     * RB_GraffitiLineType: Annot_GraffitiDraw, Annot_GraffitiSline). */
    lineType?: {
      label: string;
      freehand: PdfCommentRibbonCommand;
      straight: PdfCommentRibbonCommand;
    };
  };
  groupLabels: {
    comments: string;
    drawing: string;
    markup: string;
    reading: string;
    /** 涂黑: the mark toggle, the pending review and 应用. */
    redaction?: string;
  };
  /** The redaction tools (the host's own markup), shown as a group of
   * their own after the drawing tools; none when the file takes none. */
  redaction?: ReactNode;
  markup: {
    areaHighlight: PdfCommentRibbonCommand;
    areaHighlightMenu: AreaHighlightMenu;
    highlight: PdfCommentRibbonCommand;
    highlightMenu: ColorMenu;
    insert: PdfCommentRibbonCommand;
    insertMenu: ColorMenu;
    replace: PdfCommentRibbonCommand;
    replaceMenu: ColorMenu;
    strikeout: PdfCommentRibbonCommand;
    strikeoutMenu: ColorMenu;
    underline: PdfCommentRibbonCommand;
    underlineMenu: ColorMenu;
  };
  pointerTools: ReactNode;
  reading: {
    readAloud: PdfCommentRibbonCommand;
  };
  status?: ReactNode;
  statusTestId?: string;
  /** The squiggly, stamp, attachment, measurement, XFDF and read-aloud
   *  commands the retired "more" menu held. */
  extraTools?: { label: string; commands: readonly PdfCommentRibbonCommand[] };
  /** 退出: shown while a comment or drawing tool is armed, as WPS's annotation
   * and graffiti scenes end with an exit (audit A5). */
  exitTool?: { groupLabel: string; label: string; onClick: () => void };
}

function CommandIcon({
  name,
  size = 24,
}: {
  name: PdfCommentRibbonCommand["icon"];
  size?: number;
}) {
  if (name === "textComment") {
    return <OfficeSemanticIcon semanticId="textBox.insert" size={size} />;
  }
  return <PdfCommentRibbonIcon name={name} size={size} />;
}

function commandDisabled(command: PdfCommentRibbonCommand): boolean {
  return Boolean(command.disabled || !command.onInvoke);
}

function commandTitle(command: PdfCommentRibbonCommand): string {
  if (commandDisabled(command) && command.disabledAccessibleLabel) {
    return command.disabledAccessibleLabel;
  }
  return commandDisabled(command) && command.disabledReason
    ? `${command.label} — ${command.disabledReason}`
    : command.label;
}

function LargeCommand({ command }: { command: PdfCommentRibbonCommand }) {
  const triggerRef = useRef<HTMLButtonElement | null>(null);
  return (
    <PdfRibbonButton
      active={command.active}
      buttonRef={triggerRef}
      data-disabled-reason={
        commandDisabled(command) ? command.disabledReason : undefined
      }
      disabled={commandDisabled(command)}
      displayLabel={command.displayLabel ?? command.label}
      label={commandTitle(command)}
      onClick={() => command.onInvoke?.(triggerRef.current)}
      testId={command.testId}
    >
      <CommandIcon name={command.icon} />
    </PdfRibbonButton>
  );
}

function CommandMenuItem({ command }: { command: PdfCommentRibbonCommand }) {
  return (
    <DropdownMenuItem
      aria-label={commandTitle(command)}
      aria-current={command.active ? "true" : undefined}
      data-active={command.active ? "true" : undefined}
      disabled={commandDisabled(command)}
      title={commandTitle(command)}
      data-testid={command.testId}
      onSelect={() => command.onInvoke?.()}
    >
      <CommandIcon name={command.icon} size={16} />
      <span>{command.displayLabel ?? command.label}</span>
    </DropdownMenuItem>
  );
}

function CommentGroup({
  children,
  controlId,
  label,
  tabs = "annotate comment-tool",
  testId,
}: {
  children: ReactNode;
  controlId?: string;
  label: string;
  tabs?: string;
  testId: string;
}) {
  return (
    <PdfSurfaceRibbonGroup controlId={controlId} label={label} tab={tabs} testId={testId}>
      {children}
    </PdfSurfaceRibbonGroup>
  );
}

function SplitCommandMenu({
  command,
  children,
  menuLabel,
}: {
  command: PdfCommentRibbonCommand;
  children: ReactNode;
  menuLabel: string;
}) {
  const primaryTriggerRef = useRef<HTMLButtonElement | null>(null);
  const menuTriggerRef = useRef<HTMLButtonElement | null>(null);
  const menuTitle = commandTitle(command);
  return (
    <DropdownMenu modal={false}>
      <OfficeHomeSplitLayout
        primary={
          <OfficeHomeCommand
            active={command.active}
            buttonRef={primaryTriggerRef}
            disabled={commandDisabled(command)}
            disabledReason={
              commandDisabled(command) ? command.disabledReason : undefined
            }
            displayLabel={command.displayLabel ?? command.label}
            icon={<CommandIcon name={command.icon} />}
            label={command.label}
            onClick={() => command.onInvoke?.(primaryTriggerRef.current)}
            testId={command.testId}
            variant="large"
          />
        }
        testId={`${command.testId}-split`}
        trigger={
          <DropdownMenuTrigger asChild>
            <OfficeHomeMenuTrigger
              ref={menuTriggerRef}
              className="note-editor-office-toolbar__split-chevron"
              data-disabled-reason={
                command.disabled ? command.disabledReason : undefined
              }
              disabled={command.disabled}
              icon={
                <OfficeSemanticIcon semanticId="menu.chevron.down" size={12} />
              }
              label={menuTitle}
              testId={`${command.testId}-menu`}
              variant="icon"
            />
          </DropdownMenuTrigger>
        }
      />
      <RibbonDropdownMenuContent
        align="start"
        sideOffset={5}
        className={anatomyStyles.menuContent}
        aria-label={menuLabel}
        data-testid={`${command.testId}-menu-content`}
      >
        {children}
      </RibbonDropdownMenuContent>
    </DropdownMenu>
  );
}

/** A swatch's name in a colour menu: the hue word the sheets' colour tips use,
 * not the raw value (RGBA(252, 211, 77, 0.45)). */
function ColorName({ color }: { color: string }) {
  const t = useTranslations("console-notebooks");
  const rgb = /^rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)/iu.exec(color);
  const hex = rgb
    ? `#${rgb.slice(1, 4).map((part) => Number(part).toString(16).padStart(2, "0")).join("")}`
    : color;
  const hue = colorHueName(hex);
  return <span>{hue ? t(`spreadsheet.colorName.${hue}`) : color.toUpperCase()}</span>;
}

function PaintSettings({ menu }: { menu: PaintMenu }) {
  return (
    <>
      <DropdownMenuLabel>{menu.colorLabel}</DropdownMenuLabel>
      <DropdownMenuRadioGroup
        value={menu.activeColor}
        onValueChange={menu.onColorChange}
      >
        {menu.colors.map((color) => (
          <DropdownMenuRadioItem key={color} value={color}>
            <span
              aria-hidden
              style={{
                background: color,
                border:
                  "1px solid color-mix(in srgb, currentColor 18%, transparent)",
                borderRadius: 999,
                blockSize: 14,
                inlineSize: 14,
              }}
            />
            <ColorName color={color} />
          </DropdownMenuRadioItem>
        ))}
      </DropdownMenuRadioGroup>
      <DropdownMenuSeparator />
      <DropdownMenuLabel>{menu.widthLabel}</DropdownMenuLabel>
      <DropdownMenuRadioGroup
        value={String(menu.activeWidth)}
        onValueChange={(value) => menu.onWidthChange(Number(value))}
      >
        {menu.widths.map((width) => (
          <DropdownMenuRadioItem key={width} value={String(width)}>
            <span
              aria-hidden
              style={{
                background: "currentColor",
                blockSize: Math.max(1, Math.min(6, width)),
                inlineSize: 28,
              }}
            />
            <span>{width}</span>
          </DropdownMenuRadioItem>
        ))}
      </DropdownMenuRadioGroup>
      {menu.undo ? (
        <>
          <DropdownMenuSeparator />
          <CommandMenuItem command={menu.undo} />
        </>
      ) : null}
    </>
  );
}

function ColorSettings({ menu }: { menu: ColorMenu }) {
  return (
    <>
      <DropdownMenuLabel>{menu.colorLabel}</DropdownMenuLabel>
      <DropdownMenuRadioGroup
        value={menu.activeColor}
        onValueChange={menu.onColorChange}
      >
        {menu.colors.map((color) => (
          <DropdownMenuRadioItem key={color} value={color}>
            <span
              aria-hidden
              style={{
                background: color,
                border:
                  "1px solid color-mix(in srgb, currentColor 18%, transparent)",
                borderRadius: 999,
                blockSize: 14,
                inlineSize: 14,
              }}
            />
            <ColorName color={color} />
          </DropdownMenuRadioItem>
        ))}
      </DropdownMenuRadioGroup>
    </>
  );
}

function AreaHighlightSettings({ menu }: { menu: AreaHighlightMenu }) {
  return (
    <>
      <ColorSettings menu={menu} />
      <DropdownMenuSeparator />
      <DropdownMenuLabel>{menu.opacityLabel}</DropdownMenuLabel>
      <DropdownMenuRadioGroup
        value={String(menu.activeOpacity)}
        onValueChange={(value) => menu.onOpacityChange(Number(value))}
      >
        {menu.opacities.map((opacity) => (
          <DropdownMenuRadioItem key={opacity} value={String(opacity)}>
            {Math.round(opacity * 100)}%
          </DropdownMenuRadioItem>
        ))}
      </DropdownMenuRadioGroup>
    </>
  );
}

function ExtraToolCommand({ command }: { command: PdfCommentRibbonCommand }) {
  const triggerRef = useRef<HTMLButtonElement | null>(null);
  return (
    <PdfRibbonButton
      active={command.active}
      big={false}
      buttonRef={triggerRef}
      captioned
      data-disabled-reason={
        commandDisabled(command) ? command.disabledReason : undefined
      }
      disabled={commandDisabled(command)}
      displayLabel={command.displayLabel ?? command.label}
      label={commandTitle(command)}
      onClick={() => command.onInvoke?.(triggerRef.current)}
      testId={`${command.testId}-overflow-item`}
    >
      <CommandIcon name={command.icon} size={20} />
    </PdfRibbonButton>
  );
}

export function PdfCommentToolsOverflow({
  commands,
  label,
}: {
  commands: readonly PdfCommentRibbonCommand[];
  label: string;
}) {
  return (
    <span
      className="contents"
      role="group"
      aria-label={label}
      data-testid="pdf-global-comment-tools"
    >
      {Array.from({ length: Math.ceil(commands.length / 3) }, (_, lane) => (
        <OfficeHomeCommandStack key={lane} rows={3} className={ribbonLayout.keepsWidth}>
          {commands.slice(lane * 3, lane * 3 + 3).map((command) => (
            <ExtraToolCommand key={command.testId} command={command} />
          ))}
        </OfficeHomeCommandStack>
      ))}
    </span>
  );
}

export type PdfTextCommentPointerTool = "hand" | "select" | "text-comment";

export function PdfTextCommentToolContextRibbon({
  fontFamilies,
  labels,
  model,
  onFontColorChange,
  onFontFamilyChange,
  onFontSizeChange,
  onToolChange,
  pointerTools,
}: {
  fontFamilies: readonly Readonly<{ label: string; value: string }>[];
  labels: {
    font: string;
    fontColor: string;
    fontFamily: string;
    fontSize: string;
    textComment: string;
  };
  model: {
    activeTool: PdfTextCommentPointerTool;
    disabledReason?: string;
    fontColor: string;
    fontFamily: string;
    fontSizePoints: number;
    writable: boolean;
  };
  onFontColorChange: (color: string) => void;
  onFontFamilyChange: (fontFamily: string) => void;
  onFontSizeChange: (fontSizePoints: number) => void;
  onToolChange: (tool: PdfTextCommentPointerTool) => void;
  pointerTools: ReactNode;
}) {
  const disabledReason = model.writable ? undefined : model.disabledReason;
  const fieldLabel = (label: string) => disabledReason ?? label;

  // The groups are the track's direct children (the band's `track > group`
  // geometry and caption rules), and the fields share the Drawing Tools
  // anatomy. Leaving the tool is Esc, the tool button or another tab.
  return (
    <>
      {pointerTools}
      <CommentGroup
        controlId="PDFCTL.CONTEXT.COMMENT_TOOL.TEXT_COMMENT"
        label={labels.textComment}
        tabs="comment-tool"
        testId="pdf-text-comment-command-group"
      >
        <LargeCommand
          command={{
            active: model.activeTool === "text-comment",
            disabled: !model.writable,
            disabledAccessibleLabel: disabledReason,
            disabledReason,
            icon: "textComment",
            label: labels.textComment,
            onInvoke: model.writable
              ? () => onToolChange("text-comment")
              : undefined,
            testId: "pdf-text-comment-tool",
          }}
        />
      </CommentGroup>
      <CommentGroup
        controlId="PDFCTL.CONTEXT.COMMENT_TOOL.FONT"
        label={labels.font}
        tabs="comment-tool"
        testId="pdf-text-comment-font-group"
      >
        <div className={anatomyStyles.fieldStack}>
          <label
            className={anatomyStyles.fieldRow}
            title={fieldLabel(labels.fontFamily)}
          >
            <span className="sr-only">{labels.fontFamily}</span>
            {/* The shared ribbon drop-down, not a native select (audit A5). */}
            <PdfRibbonChoiceMenu
              controlId="PDFCTL.COMMENT.TEXT.FONT_FAMILY"
              disabled={!model.writable}
              disabledReason={disabledReason}
              icon={<OfficeSemanticIcon semanticId="font.browse" size={16} />}
              label={fieldLabel(labels.fontFamily)}
              options={fontFamilies}
              testId="pdf-text-comment-font-family"
              value={model.fontFamily}
              onPick={onFontFamilyChange}
            />
          </label>
          <span className={anatomyStyles.fieldRow}>
            <OfficeRibbonNumberField
              disabled={!model.writable}
              disabledReason={disabledReason}
              icon={<Type aria-hidden size={18} />}
              label={fieldLabel(labels.fontSize)}
              max={1000}
              min={1}
              onCommit={(raw) => {
                const next = Number.parseFloat(raw);
                if (
                  !model.writable ||
                  !Number.isFinite(next) ||
                  next <= 0 ||
                  Math.abs(next - model.fontSizePoints) < 0.05
                ) {
                  return;
                }
                onFontSizeChange(next);
              }}
              step={1}
              testId="pdf-text-comment-font-size"
              value={String(model.fontSizePoints)}
            />
            <DrawingSwatch
              disabled={!model.writable}
              disabledReason={disabledReason}
              onChange={onFontColorChange}
              testId="pdf-text-comment-font-color"
              title={fieldLabel(labels.fontColor)}
              value={model.fontColor}
            />
          </span>
        </div>
      </CommentGroup>
    </>
  );
}

/**
 * The Drawing Tools context (ink / shape / measurement) uses Shape Editing's
 * anatomy: captioned fields in two compact rows per column, Office-skinned
 * swatches and selects, never bare form controls. Test ids and the change
 * semantics are the retired inline form's, so the drawing walks still hold.
 */
function DrawingField({
  children,
  label,
  title,
}: {
  children: ReactNode;
  label: string;
  title: string;
}) {
  return (
    <label className={anatomyStyles.fieldRow} title={title}>
      <span className={anatomyStyles.fieldLabel}>{label}</span>
      {children}
    </label>
  );
}

/** A choice among a few named values (line style, arrow heads, measuring
 * type and unit): the shared ribbon drop-down -- a compact Office menu
 * trigger naming the current value and a radio list -- not a bare select.
 * Each option is `${testId}-option-${value}`. */
function DrawingSelect<Value extends string>({
  disabledReason,
  disabled,
  onChange,
  options,
  testId,
  title,
  value,
}: {
  disabledReason?: string;
  disabled?: boolean;
  onChange: (value: Value) => void;
  options: ReadonlyArray<{ value: Value; label: string }>;
  testId: string;
  title: string;
  value: Value;
}) {
  const current = options.find((option) => option.value === value);
  return (
    <DropdownMenu modal={false}>
      <DropdownMenuTrigger asChild>
        <OfficeHomeMenuTrigger
          className={anatomyStyles.styleSelect}
          disabled={disabled}
          disabledReason={disabled ? disabledReason : undefined}
          displayLabel={current?.label ?? value}
          hasMenuIndicator
          icon={null}
          label={title}
          testId={testId}
          variant="compact"
        />
      </DropdownMenuTrigger>
      <RibbonDropdownMenuContent
        align="start"
        sideOffset={5}
        className={anatomyStyles.menuContent}
        aria-label={title}
        data-testid={`${testId}-menu`}
      >
        <DropdownMenuRadioGroup
          value={value}
          onValueChange={(next) => onChange(next as Value)}
        >
          {options.map((option) => (
            <DropdownMenuRadioItem
              key={option.value}
              value={option.value}
              data-testid={`${testId}-option-${option.value}`}
            >
              {option.label}
            </DropdownMenuRadioItem>
          ))}
        </DropdownMenuRadioGroup>
      </RibbonDropdownMenuContent>
    </DropdownMenu>
  );
}

/** A colour field: a swatch button opening the shared Office colour
 * palette (theme, standard and recent colours, More Colors), as the
 * annotation Format tab's fill and outline do. The palette's cells are
 * `${testId}-palette-…`. */
function DrawingSwatch({
  disabledReason,
  disabled,
  onChange,
  testId,
  title,
  value,
}: {
  disabledReason?: string;
  disabled?: boolean;
  onChange: (color: string) => void;
  testId: string;
  title: string;
  value: string;
}) {
  const t = useTranslations("console-notebooks");
  const triggerRef = useRef<HTMLButtonElement | null>(null);
  const [open, setOpen] = useState(false);
  return (
    <span className={anatomyStyles.swatchHost}>
      <OfficeHomeMenuTrigger
        ref={triggerRef}
        aria-expanded={open && !disabled}
        aria-haspopup="dialog"
        className={anatomyStyles.swatch}
        disabled={disabled}
        disabledReason={disabled ? disabledReason : undefined}
        icon={
          <span
            className={anatomyStyles.swatchChip}
            style={{ background: value }}
          />
        }
        label={title}
        onClick={() => setOpen((current) => !current)}
        testId={testId}
        variant="icon"
      />
      {open && !disabled ? (
        <OfficeColorPalettePopover
          anchorRef={triggerRef}
          labels={{
            themeColors: t("pdf.annotationFormat.themeColors"),
            standardColors: t("pdf.annotationFormat.standardColors"),
            recentColors: t("pdf.annotationFormat.recentColors"),
            moreColors: t("pdf.annotationFormat.moreColors"),
          }}
          testIdPrefix={`${testId}-palette`}
          value={value}
          onDismiss={() => setOpen(false)}
          onPick={(color) => {
            setOpen(false);
            if (color) onChange(color);
          }}
        />
      ) : null}
    </span>
  );
}

const drawingFieldClassNames = {
  field: anatomyStyles.captionField,
  icon: anatomyStyles.noIcon,
  label: anatomyStyles.fieldLabel,
} as const;

function PdfDrawingToolsContext({
  context,
  exitTool,
  pointerTools,
  status,
  statusTestId,
}: {
  context: NonNullable<PdfCommentRibbonTabProps["drawingContext"]>;
  exitTool?: PdfCommentRibbonTabProps["exitTool"];
  pointerTools: ReactNode;
  status?: ReactNode;
  statusTestId?: string;
}) {
  const styleReason = context.styleEnabled ? undefined : context.disabledReason;
  const fillReason = context.fillSupported
    ? styleReason
    : context.fillDisabledReason;
  const fillColorReason =
    styleReason ??
    (context.fillColor === null ? context.fillDisabledReason : undefined);
  const lineWidthField = (
    <OfficeRibbonNumberField
      classNames={drawingFieldClassNames}
      disabled={!context.styleEnabled}
      disabledReason={styleReason}
      icon={null}
      inputWidthPx={48}
      label={styleReason ?? context.lineWidthLabel}
      max={context.lineWidthMax ?? 12}
      min={context.lineWidthMin ?? 0.1}
      // Written once, on Enter or leaving the field (as Shape editing's
      // width), not on every keystroke of a half-typed number.
      onCommit={(raw) => {
        const width = Number.parseFloat(raw);
        if (
          !context.styleEnabled ||
          !Number.isFinite(width) ||
          width === context.lineWidth
        ) {
          return;
        }
        context.onLineWidthChange(width);
      }}
      step={0.1}
      testId="pdf-drawing-line-width"
      value={String(context.lineWidth)}
      visibleLabel={context.lineWidthLabel}
    />
  );
  const outlineField = (
    <DrawingField
      label={context.outlineColorLabel}
      title={styleReason ?? context.outlineColorLabel}
    >
      <DrawingSwatch
        disabled={!context.styleEnabled}
        disabledReason={styleReason}
        onChange={context.onOutlineColorChange}
        testId="pdf-drawing-outline-color"
        title={styleReason ?? context.outlineColorLabel}
        value={context.outlineColor}
      />
    </DrawingField>
  );
  // The groups are the track's direct children, as every other PDF group
  // is: the band's 100px group geometry and its caption are `track > group`
  // rules, which a wrapping element (even `display: contents`) defeats. The
  // style group carries the context's test id.
  // Leaving the tool is also Esc, the tool's own button or another tab; the
  // band ends with 退出编辑 as WPS's graffiti, shape and measure tabs do
  // (CT_Graffiti, CT_ShapeStyle, CT_MeasureTool: Exit).
  const measurement = context.kind === "measurement" ? context.measurement : undefined;
  return (
    <>
      {pointerTools}
      {measurement ? (
        // WPS CT_MeasureTool RB_MeasureTool: Distance, Perimeter, Area, each
        // its own command, the one in use pressed.
        <CommentGroup
          controlId="PDFCTL.CONTEXT.MEASURE.KIND"
          label={measurement.kindLabel}
          tabs="comment-tool"
          testId="pdf-drawing-measure-kind-group"
        >
          {measurement.kindOptions.map((option) => (
            <OfficeHomeCommand
              key={option.value}
              active={measurement.kind === option.value}
              activation="click"
              displayLabel={option.label}
              icon={<MeasureKindIcon kind={option.value} />}
              label={option.label}
              onClick={() => measurement.onKindChange(option.value)}
              testId={`pdf-drawing-measure-kind-${option.value}`}
              variant="large"
            />
          ))}
        </CommentGroup>
      ) : null}
      {context.kind === "ink" && context.lineType ? (
        <CommentGroup
          controlId="PDFCTL.CONTEXT.GRAFFITI.LINE_TYPE"
          label={context.lineType.label}
          tabs="comment-tool"
          testId="pdf-drawing-ink-line-type-group"
        >
          <LargeCommand command={context.lineType.freehand} />
          <LargeCommand command={context.lineType.straight} />
        </CommentGroup>
      ) : null}
      <CommentGroup
        controlId="PDFCTL.CONTEXT.DRAWING_TOOLS.STYLE"
        label={context.styleLabel}
        tabs="comment-tool"
        testId="pdf-drawing-tools-context"
      >
        {context.kind === "shape" ? (
          <>
            <div className={anatomyStyles.fieldStack}>
              <span
                className={anatomyStyles.fieldRow}
                title={fillReason ?? context.fillLabel}
              >
                <OfficeHomeCheckbox
                  aria-label={fillReason ?? context.fillLabel}
                  checked={context.fillColor !== null}
                  data-disabled-reason={fillReason}
                  disabled={!context.fillSupported || !context.styleEnabled}
                  label={context.fillLabel}
                  onCheckedChange={(checked) =>
                    context.onFillColorChange(
                      checked
                        ? (context.fillColor ?? context.outlineColor)
                        : null,
                    )
                  }
                  testId="pdf-drawing-fill-enabled"
                  title={fillReason ?? context.fillLabel}
                />
                {context.fillSupported ? (
                  <DrawingSwatch
                    disabled={!context.styleEnabled || context.fillColor === null}
                    disabledReason={fillColorReason}
                    onChange={context.onFillColorChange}
                    testId="pdf-drawing-fill-color"
                    title={fillColorReason ?? context.fillLabel}
                    value={context.fillColor ?? context.outlineColor}
                  />
                ) : null}
              </span>
              {outlineField}
            </div>
            <div className={anatomyStyles.fieldStack}>
              {lineWidthField}
              <DrawingField
                label={context.lineStyleLabel}
                title={styleReason ?? context.lineStyleLabel}
              >
                <DrawingSelect
                  disabled={!context.styleEnabled}
                  disabledReason={styleReason}
                  onChange={context.onLineStyleChange}
                  options={SHAPE_LINE_STYLES.map((style) => ({
                    value: style,
                    label: context.lineStyleLabels[style],
                  }))}
                  testId="pdf-drawing-line-style"
                  title={styleReason ?? context.lineStyleLabel}
                  value={context.lineStyle}
                />
              </DrawingField>
            </div>
            {context.lineEndingsSupported ? (
              <div className={anatomyStyles.fieldStack}>
                {(
                  [
                    [
                      "start",
                      context.lineStartLabel,
                      context.lineStart,
                      context.onLineStartChange,
                    ],
                    [
                      "end",
                      context.lineEndLabel,
                      context.lineEnd,
                      context.onLineEndChange,
                    ],
                  ] as const
                ).map(([edge, label, value, onChange]) => (
                  <DrawingField
                    key={edge}
                    label={label}
                    title={styleReason ?? label}
                  >
                    <DrawingSelect
                      disabled={!context.styleEnabled}
                      disabledReason={styleReason}
                      onChange={onChange}
                      options={SHAPE_LINE_ENDINGS.map((ending) => ({
                        value: ending,
                        label: context.lineEndingLabels[ending],
                      }))}
                      testId={`pdf-drawing-line-${edge}`}
                      title={styleReason ?? label}
                      value={value}
                    />
                  </DrawingField>
                ))}
              </div>
            ) : null}
          </>
        ) : null}
        {context.kind === "ink" ? (
          <div className={anatomyStyles.fieldStack}>
            {outlineField}
            {lineWidthField}
          </div>
        ) : null}
        {context.kind === "ink" && context.eraser ? (
          <LargeCommand command={context.eraser} />
        ) : null}
        {context.kind === "measurement" && context.measurement ? (
          <>
            <div className={anatomyStyles.fieldStack}>
              <DrawingField
                label={context.measurement.unitLabel}
                title={context.measurement.unitLabel}
              >
                <DrawingSelect
                  onChange={context.measurement.onUnitChange}
                  options={context.measurement.units.map((unit) => ({
                    value: unit,
                    label: unit,
                  }))}
                  testId="pdf-drawing-measure-unit"
                  title={context.measurement.unitLabel}
                  value={context.measurement.unit}
                />
              </DrawingField>
            </div>
            <div className={anatomyStyles.fieldStack}>
              <OfficeRibbonNumberField
                classNames={drawingFieldClassNames}
                icon={null}
                inputWidthPx={56}
                label={context.measurement.ratioLabel}
                max={1000000}
                min={0.0001}
                // Written when the value is committed (Enter, leaving the box,
                // a step), not per keystroke (audit A5).
                onCommit={(raw) => {
                  const ratio = Number(raw);
                  context.measurement?.onRatioChange(
                    Number.isFinite(ratio) && ratio > 0 ? ratio : 1,
                  );
                }}
                step={0.01}
                testId="pdf-drawing-measure-ratio"
                value={String(context.measurement.ratio)}
                visibleLabel={context.measurement.ratioLabel}
              />
            </div>
            {context.measurement.keep ? (
              <LargeCommand command={context.measurement.keep} />
            ) : null}
            {context.measurement.removeLast ? (
              <LargeCommand command={context.measurement.removeLast} />
            ) : null}
            {context.measurement.results ? (
              <MeasureResultsMenu results={context.measurement.results} />
            ) : null}
            {context.measurement.snap ? (
              <span
                className={anatomyStyles.fieldRow}
                role="group"
                aria-label={context.measurement.snap.label}
                data-testid="pdf-drawing-measure-snap"
              >
                {[context.measurement.snap.options.slice(0, 2), context.measurement.snap.options.slice(2)].map(
                  (column, index) => (
                    <div key={index} className={anatomyStyles.fieldStack}>
                      {column.map((option) => (
                        <OfficeHomeCheckbox
                          key={option.kind}
                          checked={option.checked}
                          label={option.label}
                          onCheckedChange={() => option.onToggle()}
                          testId={`pdf-drawing-measure-snap-${option.kind}`}
                        />
                      ))}
                    </div>
                  ),
                )}
              </span>
            ) : null}
          </>
        ) : null}
        {status ? (
          <span role="status" data-testid={statusTestId}>
            {status}
          </span>
        ) : null}
      </CommentGroup>
      {context.kind === "ink" ? <DrawWithTouchGroup /> : null}
      {exitTool ? <ExitToolGroup exitTool={exitTool} /> : null}
    </>
  );
}

function MeasureResultsMenu({
  results,
}: {
  results: NonNullable<
    NonNullable<NonNullable<PdfCommentRibbonTabProps["drawingContext"]>["measurement"]>["results"]
  >;
}) {
  return (
    <DropdownMenu modal={false}>
      <DropdownMenuTrigger asChild>
        <OfficeHomeMenuTrigger
          displayLabel={results.label}
          hasMenuIndicator
          icon={<OfficeSemanticIcon semanticId="view.ruler" size={24} />}
          label={results.label}
          testId="pdf-drawing-measure-results"
          variant="large"
        />
      </DropdownMenuTrigger>
      <RibbonDropdownMenuContent
        align="start"
        aria-label={results.label}
        data-testid="pdf-drawing-measure-results-menu"
      >
        <DropdownMenuLabel>{results.label}</DropdownMenuLabel>
        {results.items.length === 0 ? (
          <DropdownMenuItem disabled data-testid="pdf-drawing-measure-results-empty">
            {results.emptyLabel}
          </DropdownMenuItem>
        ) : (
          results.items.map((item) => (
            <DropdownMenuItem
              key={item.id}
              data-testid={`pdf-drawing-measure-result-${item.id}`}
              onSelect={() => results.onOpen(item.page)}
            >
              {item.text}
            </DropdownMenuItem>
          ))
        )}
      </RibbonDropdownMenuContent>
    </DropdownMenu>
  );
}

function MeasureKindIcon({ kind }: { kind: string }) {
  const Icon = kind === "area" ? SquareDashed : kind === "perimeter" ? Hexagon : Ruler;
  return <Icon aria-hidden size={24} strokeWidth={1.6} />;
}

/** 退出编辑, the last group of every tool context (WPS: Exit). */
function ExitToolGroup({
  exitTool,
}: {
  exitTool: NonNullable<PdfCommentRibbonTabProps["exitTool"]>;
}) {
  return (
    <CommentGroup label={exitTool.groupLabel} testId="pdf-comment-exit-group">
      <OfficeHomeCommand
        activation="click"
        controlId="PDFCTL.COMMENT.EXIT_TOOL"
        displayLabel={exitTool.label}
        icon={<OfficeSemanticIcon semanticId="edit.exit" size={24} />}
        label={exitTool.label}
        onClick={exitTool.onClick}
        testId="pdf-comment-exit-tool"
        variant="large"
      />
    </CommentGroup>
  );
}






function DrawWithTouchGroup() {
  const t = useTranslations("console-notebooks");
  const hasTouch = useDeviceHasTouch();
  if (!hasTouch) return null;
  return (
    <CommentGroup
      controlId="PDFCTL.CONTEXT.DRAWING_TOOLS.DRAW_WITH_TOUCH"
      label={t("toolbar.drawTouch")}
      tabs="comment-tool"
      testId="pdf-drawing-touch-group"
    >
      <OfficeInkTouchSwitch testId="pdf-drawing-draw-touch" />
    </CommentGroup>
  );
}

/** PDF-owned adapter over the shared 100px Office Comments surface. */
export default function PdfCommentRibbonTab({
  comments,
  drawing,
  drawingContext,
  groupLabels,
  markup,
  pointerTools,
  reading,
  status,
  statusTestId,
  extraTools,
  measure,
  exitTool,
  redaction,
}: PdfCommentRibbonTabProps) {
  if (drawingContext) {
    return (
      <PdfDrawingToolsContext
        context={drawingContext}
        exitTool={exitTool}
        pointerTools={pointerTools}
        status={status}
        statusTestId={statusTestId}
      />
    );
  }
  return (
    <>
      {pointerTools}

      <CommentGroup
        label={groupLabels.drawing}
        testId="pdf-comment-primary-group"
      >
        <SplitCommandMenu
          command={markup.highlight}
          menuLabel={markup.highlight.label}
        >
          <ColorSettings menu={markup.highlightMenu} />
        </SplitCommandMenu>
        <SplitCommandMenu
          command={markup.areaHighlight}
          menuLabel={markup.areaHighlight.label}
        >
          <AreaHighlightSettings menu={markup.areaHighlightMenu} />
        </SplitCommandMenu>
        <SplitCommandMenu command={drawing.draw} menuLabel={drawing.draw.label}>
          <PaintSettings menu={drawing.drawMenu} />
        </SplitCommandMenu>
        <LargeCommand command={drawing.textComment} />
        <SplitCommandMenu
          command={drawing.shape}
          menuLabel={drawing.shape.label}
        >
          <DropdownMenuLabel>{drawing.shape.label}</DropdownMenuLabel>
          {drawing.shapeMenu.kinds.map((command) => (
            <CommandMenuItem key={command.testId} command={command} />
          ))}
          <DropdownMenuSeparator />
          <DropdownMenuLabel>{drawing.shapeMenu.colorLabel}</DropdownMenuLabel>
          <DropdownMenuRadioGroup
            value={drawing.shapeMenu.activeColor}
            onValueChange={drawing.shapeMenu.onColorChange}
          >
            {drawing.shapeMenu.colors.map((color) => (
              <DropdownMenuRadioItem key={color} value={color}>
                <span
                  aria-hidden
                  style={{
                    background: color,
                    border:
                      "1px solid color-mix(in srgb, currentColor 18%, transparent)",
                    borderRadius: 999,
                    blockSize: 14,
                    inlineSize: 14,
                  }}
                />
                <ColorName color={color} />
              </DropdownMenuRadioItem>
            ))}
          </DropdownMenuRadioGroup>
          <DropdownMenuSeparator />
          <DropdownMenuLabel>{drawing.shapeMenu.widthLabel}</DropdownMenuLabel>
          <DropdownMenuRadioGroup
            value={String(drawing.shapeMenu.activeWidth)}
            onValueChange={(value) =>
              drawing.shapeMenu.onWidthChange(Number(value))
            }
          >
            {drawing.shapeMenu.widths.map((width) => (
              <DropdownMenuRadioItem key={width} value={String(width)}>
                <span
                  aria-hidden
                  style={{
                    background: "currentColor",
                    blockSize: Math.max(1, Math.min(6, width)),
                    inlineSize: 28,
                  }}
                />
                <span>
                  {drawing.shapeMenu.widthUnit
                    ? `${width} ${drawing.shapeMenu.widthUnit}`
                    : width}
                </span>
              </DropdownMenuRadioItem>
            ))}
          </DropdownMenuRadioGroup>
        </SplitCommandMenu>
        {drawing.stamp ? <LargeCommand command={drawing.stamp} /> : null}
        <SplitCommandMenu command={drawing.note} menuLabel={drawing.note.label}>
          <ColorSettings menu={drawing.noteMenu} />
        </SplitCommandMenu>
      </CommentGroup>

      <CommentGroup
        label={groupLabels.markup}
        testId="pdf-comment-markup-group"
      >
        <SplitCommandMenu
          command={markup.underline}
          menuLabel={markup.underline.label}
        >
          <ColorSettings menu={markup.underlineMenu} />
        </SplitCommandMenu>
        <SplitCommandMenu
          command={markup.strikeout}
          menuLabel={markup.strikeout.label}
        >
          <ColorSettings menu={markup.strikeoutMenu} />
        </SplitCommandMenu>
        <SplitCommandMenu
          command={markup.replace}
          menuLabel={markup.replace.label}
        >
          <ColorSettings menu={markup.replaceMenu} />
        </SplitCommandMenu>
        <SplitCommandMenu
          command={markup.insert}
          menuLabel={markup.insert.label}
        >
          <ColorSettings menu={markup.insertMenu} />
        </SplitCommandMenu>
      </CommentGroup>

      {measure ? (
        <CommentGroup label={measure.label} testId="pdf-comment-measure-group">
          <LargeCommand command={measure.command} />
        </CommentGroup>
      ) : null}

      <CommentGroup
        label={groupLabels.comments}
        testId="pdf-comment-management-group"
      >
        <LargeCommand command={comments.commentMode} />
        <LargeCommand command={comments.hide} />
        <LargeCommand command={comments.manage} />
      </CommentGroup>

      <CommentGroup
        label={groupLabels.reading}
        testId="pdf-comment-reading-group"
      >
        <LargeCommand command={reading.readAloud} />
      </CommentGroup>
      {redaction ? (
        <CommentGroup
          label={groupLabels.redaction ?? ""}
          testId="pdf-comment-redaction-group"
        >
          {redaction}
        </CommentGroup>
      ) : null}

      {extraTools ? (
        <CommentGroup
          label={extraTools.label}
          testId="pdf-comment-extra-tools-group"
        >
          <PdfCommentToolsOverflow
            commands={extraTools.commands}
            label={extraTools.label}
          />
        </CommentGroup>
      ) : null}

      {exitTool ? <ExitToolGroup exitTool={exitTool} /> : null}

      {status ? (
        <span
          className="pdf-viewer__annotation-clipboard-error"
          role="status"
          data-testid={statusTestId}
        >
          {status}
        </span>
      ) : null}
    </>
  );
}
