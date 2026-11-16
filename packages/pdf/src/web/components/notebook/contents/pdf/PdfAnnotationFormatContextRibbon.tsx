"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { PdfNativeVectorMergeOperation } from "../../../../lib/pdf/native-vector-selection-edit";
import {
  ClipboardPaste,
  Copy,
  CopyPlus,
  Download,
  Eraser,
  Maximize2,
  Paperclip,
  PenTool,
  RotateCcw,
  RotateCw,
  ShieldAlert,
  Trash2,
} from "@mrnote/core/web/components/office-ribbon/icons/houseIcons.tsx";
import { type ReactNode, useEffect, useRef, useState } from "react";

import {
  RibbonDropdownMenuContent,
  RibbonDropdownMenuSubContent,
} from "@mrnote/core/web/components/office-ribbon/CupertinoRibbonMenu.tsx";
import { OfficeColorPalettePopover } from "@mrnote/core/web/components/office-ribbon/OfficeColorPalettePopover.tsx";
import {
  OfficeHomeCommand,
  OfficeHomeCommandRow,
  OfficeHomeCommandStack,
  OfficeHomeMenuTrigger,
  OfficeHomeRibbonGroup,
  OfficeHomeSplitLayout,
} from "@mrnote/core/web/components/office-ribbon/OfficeHomeRibbon.tsx";
import {
  DropdownMenu,
  DropdownMenuItem,
  DropdownMenuLabel,
  DropdownMenuRadioGroup,
  DropdownMenuRadioItem,
  DropdownMenuSeparator,
  DropdownMenuSub,
  DropdownMenuSubTrigger,
  DropdownMenuTrigger,
} from "@mrnote/core/web/components/ui/dropdown-menu.tsx";
import { OfficeRibbonNumberField } from "@mrnote/core/web/components/office-ribbon/OfficeRibbonNumberField.tsx";
import { PdfRibbonChoiceMenu } from "./PdfRibbonChoiceMenu";
import { OfficeSemanticIcon } from "@mrnote/core/web/components/office-ribbon/icons/OfficeSemanticIcon.tsx";
import { OfficeShapeStyleGallery } from "@mrnote/core/web/components/office-ribbon/OfficeShapeStyleGallery.tsx";
import {
  OFFICE_WEB_SHAPE_ACCENTS,
  officeShapeStyleRecipes,
  type OfficeShapeStyleRecipe,
} from "@mrnote/core/web/lib/office-shape-style-catalog.ts";
import {
  SHAPE_LINE_ENDINGS,
  SHAPE_LINE_STYLES,
  type ShapeKind,
  type ShapeGeometryTransform,
  type ShapeLineEnding,
  type ShapeLineStyle,
} from "@mrnote/core/web/lib/pdf/shapes.ts";
import {
  PDF_STANDARD_BLEND_MODES,
  type PdfStandardBlendMode,
} from "../../../../lib/pdf/blend-modes";

import styles from "./PdfAnnotationFormatContextRibbon.module.css";
import type { AlignEdge } from "@mrnote/core/web/lib/presentation-align.ts";
import { pdfShapeAlignAvailable, type PdfShapeAlignCommand } from "../../../../lib/pdf/shape-align";
import type { PdfRibbonTranslator } from "./pdfRibbonTranslator";
import ribbonLayout from "./pdfRibbonLayout.module.css";

export type PdfAnnotationContextKind =
  | "shape"
  | "signature"
  | "stamp"
  | "measurement"
  | "attachment"
  | "note"
  | "ink"
  | "redaction";

export interface PdfAnnotationFormatContextRibbonModel {
  kind: PdfAnnotationContextKind;
  nativeSource: boolean;
  /** A drawing of the page's own content (Edit Content), not an annotation:
   * its delete is 删除图形, never 删除批注 (user audit PDF 0053). */
  contentObject?: boolean;
  page: number;
  width: number;
  height: number;
  widthPoints?: number;
  heightPoints?: number;
  sizeEnabled?: boolean;
  sizeDisabledReason?: string;
  /** Changes when a size (or any geometry) write for this object was turned
   *  down, so the size fields show the object's real size again. */
  sizeResyncKey?: number;
  pointAnchor?: boolean;
  busy: boolean;
  copyEnabled: boolean;
  copyDisabledReason?: string;
  pasteEnabled: boolean;
  pasteDisabledReason?: string;
  duplicateEnabled: boolean;
  duplicateDisabledReason?: string;
  arrangeEnabled?: boolean;
  arrangeDisabledReason?: string;
  /** 对齐 / 分布 for the selected source vectors: on the box around two or
   * more, on the page for one; 分布 needs three. */
  alignEnabled?: boolean;
  selectionCount?: number;
  /** 合并形状 (Shape Format > Merge Shapes) on two selected closed shapes. */
  mergeEnabled?: boolean;
  mergeDisabledReason?: string;
  /** Native z-order (bring forward / send backward) for source vectors. */
  reorderEnabled?: boolean;
  reorderDisabledReason?: string;
  /** The object is already drawn last (first) among its stream's vectors, so
   * the moves in that direction grey out and say so, as Office does. */
  reorderAtFront?: boolean;
  reorderAtBack?: boolean;
  /** Cover the source vector with an appearance-only visual mask. */
  maskEnabled?: boolean;
  maskDisabledReason?: string;
  /** Vertex editing of a source vector (a line or a subpath). */
  vertexEditing?: boolean;
  vertexEditEnabled?: boolean;
  vertexEditDisabledReason?: string;
  deleteEnabled: boolean;
  deleteDisabledReason?: string;
  styleEnabled?: boolean;
  styleDisabledReason?: string;
  /** The style takes a change while the object's last one is still being
   * written (queued after it): `busy` does not grey the style controls. */
  styleQueues?: boolean;
  strokeSupported?: boolean;
  outlineColor?: string;
  lineWidth?: number;
  shapeKind?: ShapeKind;
  lineStyle?: ShapeLineStyle;
  lineStart?: ShapeLineEnding;
  lineEnd?: ShapeLineEnding;
  lineEndingsSupported?: boolean;
  fillSupported?: boolean;
  fillToggleEnabled?: boolean;
  fillToggleDisabledReason?: string;
  fillColor?: string | null;
  fillOpacitySupported?: boolean;
  fillOpacity?: number;
  opacitySupported?: boolean;
  opacity?: number;
  /** Why the transparency cannot be set as one number (a drawing whose
   * fill and outline differ): said on the control, which then shows no
   * value instead of a 0 that is not the drawing's (user audit PDF 0052). */
  opacityDisabledReason?: string;
  blendModeSupported?: boolean;
  blendMode?: PdfStandardBlendMode;
  cornerSupported?: boolean;
  cornerRadius?: number;
  /** Line caps and corner joins of the outline (PDF `J` / `j`). */
  capJoinSupported?: boolean;
  capJoinDisabledReason?: string;
  lineCap?: PdfShapeLineCap;
  lineJoin?: PdfShapeLineJoin;
  /** Any angle (a source vector); otherwise quarter turns only. */
  freeRotation?: boolean;
  /** The angle the object was turned to (degrees, clockwise), when it is
   * known: the angle field reads it, as WPS's 旋转 reads 30 after a 30 degree
   * turn, and a typed angle turns the object to it. Without it the field is
   * a command that always reads 0 and turns by what is typed. */
  rotationDegrees?: number;
  /** The Format Shape pane is open for this shape. */
  formatPaneOpen?: boolean;
  attachment?: {
    name: string;
    mimeType: string;
    sizeBytes: number;
    saveEnabled: boolean;
    saveDisabledReason?: string;
  };
}

export type PdfShapeLineCap = "butt" | "round" | "square";
export type PdfShapeLineJoin = "miter" | "round" | "bevel";
export const PDF_SHAPE_LINE_CAPS: readonly PdfShapeLineCap[] = [
  "butt",
  "round",
  "square",
];
export const PDF_SHAPE_LINE_JOINS: readonly PdfShapeLineJoin[] = [
  "miter",
  "round",
  "bevel",
];

export type PdfAnnotationStyleChange = Readonly<{
  outlineColor?: string;
  lineWidth?: number;
  lineStyle?: ShapeLineStyle;
  lineStart?: ShapeLineEnding;
  lineEnd?: ShapeLineEnding;
  fillColor?: string | null;
  fillOpacity?: number;
  opacity?: number;
  blendMode?: PdfStandardBlendMode;
  cornerRadius?: number;
  lineCap?: PdfShapeLineCap;
  lineJoin?: PdfShapeLineJoin;
}>;

export interface PdfAnnotationFormatContextRibbonLabels {
  annotation: string;
  style: string;
  inkStyle: string;
  clipboard: string;
  arrange: string;
  size: string;
  width: string;
  height: string;
  points: string;
  actions: string;
  shape: string;
  signature: string;
  stamp: string;
  measurement: string;
  attachment: string;
  attachmentFile: string;
  attachmentName: string;
  attachmentType: string;
  attachmentSize: string;
  saveAttachment: string;
  note: string;
  ink: string;
  redaction: string;
  managed: string;
  nativeSource: string;
  copy: string;
  paste: string;
  duplicate: string;
  rotateLeft: string;
  rotateRight: string;
  flipHorizontal: string;
  flipVertical: string;
  align: string;
  alignToPage: string;
  alignSelected: string;
  alignEdges: Readonly<Record<AlignEdge, string>>;
  distributeHorizontally: string;
  distributeVertically: string;
  mergeShapes: string;
  mergeOperations: Readonly<Record<PdfNativeVectorMergeOperation, string>>;
  bringForward: string;
  sendBackward: string;
  bringToFront: string;
  sendToBack: string;
  alreadyAtFront: string;
  alreadyAtBack: string;
  /** Absent where the ribbon has no Selection pane (other editors' use). */
  selectionPane?: string;
  editVertices: string;
  exitVertexEdit: string;
  mask: string;
  delete: string;
  /** Delete for a drawing of the page's own content. */
  deleteShape: string;
  outlineColor: string;
  inkColor: string;
  lineWidth: string;
  lineStyle: string;
  lineStart: string;
  lineEnd: string;
  lineStyleLabels: Record<ShapeLineStyle, string>;
  lineEndingLabels: Record<ShapeLineEnding, string>;
  fill: string;
  fillColor: string;
  fillOpacity: string;
  opacity: string;
  blendMode: string;
  blendModeLabels: Record<PdfStandardBlendMode, string>;
  cornerRadius: string;
  pointAnchor: string;
  backToEditing: string;
  pageNumber: (page: number) => string;
  insert: string;
  outlineMenu: string;
  transparency: string;
  lineCap: string;
  lineJoin: string;
  lineCapLabels: Record<PdfShapeLineCap, string>;
  lineJoinLabels: Record<PdfShapeLineJoin, string>;
  lockAspectRatio: string;
  rotate: string;
  rotateQuarterTurnsOnly: string;
  more: string;
  exitEditing: string;
  noFill: string;
  moreColors: string;
  themeColors: string;
  standardColors: string;
  recentColors: string;
  /** Quick Styles gallery: its "more" menu and why a filled style is off
   * for an annotation that has no fill (the gallery names each style from
   * the catalogue itself). */
  moreShapeStyles: string;
  outlineStylesOnly: string;
  /** Format Shape (the pane the Shape Styles launcher opens). */
  formatShape: string;
}

/**
 * The Quick Styles an annotation can take exactly: Office's flat rows
 * (colored outline, colored fill, solid fill) on the default theme accents.
 * The gradient and shadow rows are left out, since a PDF annotation's
 * appearance paints one fill and one stroke. "Colored Outline" is outline
 * only here (no fill), so the style never hides the page under it.
 */
export const PDF_ANNOTATION_STYLE_RECIPES: readonly OfficeShapeStyleRecipe[] =
  officeShapeStyleRecipes(OFFICE_WEB_SHAPE_ACCENTS)
    .filter(
      (recipe) =>
        recipe.row === "outline" ||
        recipe.row === "outlined-fill" ||
        recipe.row === "solid",
    )
    .map((recipe) =>
      recipe.row === "outline" ? { ...recipe, fill: "none" } : recipe,
    );

const sameColor = (left: string | null | undefined, right: string) =>
  (left ?? "").toLowerCase() === right.toLowerCase();

/** The Quick Style the annotation wears now, if its paint is one of them. */
export function pdfAnnotationStyleRecipeId(
  model: Pick<
    PdfAnnotationFormatContextRibbonModel,
    "outlineColor" | "lineWidth" | "fillColor" | "fillSupported"
  >,
): string | null {
  const match = PDF_ANNOTATION_STYLE_RECIPES.find(
    (recipe) =>
      sameColor(model.outlineColor, recipe.stroke) &&
      Math.abs((model.lineWidth ?? 0) - recipe.strokeWidth) < 0.01 &&
      (!model.fillSupported ||
        (recipe.row === "outline"
          ? model.fillColor == null
          : sameColor(model.fillColor, recipe.fill))),
  );
  return match?.id ?? null;
}

export function createPdfAnnotationFormatContextRibbonLabels(
  t: PdfRibbonTranslator,
): PdfAnnotationFormatContextRibbonLabels {
  return {
    annotation: t("pdf.annotationFormat.annotation"),
    style: t("pdf.annotationFormat.style"),
    inkStyle: t("pdf.annotationFormat.inkStyle"),
    clipboard: t("pdf.annotationFormat.clipboard"),
    arrange: t("pdf.annotationFormat.arrange"),
    size: t("pdf.annotationFormat.size"),
    width: t("pdf.annotationFormat.width"),
    height: t("pdf.annotationFormat.height"),
    points: t("pdf.annotationFormat.points"),
    actions: t("pdf.annotationFormat.actions"),
    shape: t("pdf.annotationFormat.shape"),
    signature: t("pdf.annotationFormat.signature"),
    stamp: t("pdf.annotationFormat.stamp"),
    measurement: t("pdf.annotationFormat.measurement"),
    attachment: t("pdf.annotationFormat.attachment"),
    attachmentFile: t("pdf.annotationFormat.attachmentFile"),
    attachmentName: t("pdf.annotationFormat.attachmentName"),
    attachmentType: t("pdf.annotationFormat.attachmentType"),
    attachmentSize: t("pdf.annotationFormat.attachmentSize"),
    saveAttachment: t("pdf.annotationFormat.saveAttachment"),
    note: t("pdf.annotationFormat.note"),
    ink: t("pdf.annotationFormat.ink"),
    redaction: t("pdf.annotationFormat.redaction"),
    managed: t("pdf.annotationFormat.managed"),
    nativeSource: t("pdf.annotationFormat.nativeSource"),
    copy: t("pdf.annotationFormat.copy"),
    paste: t("pdf.annotationFormat.paste"),
    duplicate: t("pdf.annotationFormat.duplicate"),
    rotateLeft: t("pdf.annotationFormat.rotateLeft"),
    rotateRight: t("pdf.annotationFormat.rotateRight"),
    flipHorizontal: t("pdf.annotationFormat.flipHorizontal"),
    flipVertical: t("pdf.annotationFormat.flipVertical"),
    align: t("pdf.annotationFormat.align"),
    alignToPage: t("pdf.annotationFormat.alignToPage"),
    // The Office words the deck editor's Align menu uses.
    alignSelected: t("presentation.arrange.alignSelectedObjects"),
    alignEdges: {
      left: t("presentation.arrange.alignLeft"),
      hcenter: t("presentation.arrange.alignCenter"),
      right: t("presentation.arrange.alignRight"),
      top: t("presentation.arrange.alignTop"),
      vcenter: t("presentation.arrange.alignMiddle"),
      bottom: t("presentation.arrange.alignBottom"),
    },
    distributeHorizontally: t("presentation.arrange.distributeHorizontally"),
    distributeVertically: t("presentation.arrange.distributeVertically"),
    mergeShapes: t("presentation.contextMenu.mergeShapes"),
    mergeOperations: {
      union: t("presentation.mergeShapes.union"),
      combine: t("presentation.mergeShapes.combine"),
      fragment: t("presentation.mergeShapes.fragment"),
      intersect: t("presentation.mergeShapes.intersect"),
      subtract: t("presentation.mergeShapes.subtract"),
    },
    bringForward: t("pdf.annotationFormat.bringForward"),
    sendBackward: t("pdf.annotationFormat.sendBackward"),
    bringToFront: t("pdf.annotationFormat.bringToFront"),
    sendToBack: t("pdf.annotationFormat.sendToBack"),
    alreadyAtFront: t("pdf.annotationFormat.alreadyAtFront"),
    alreadyAtBack: t("pdf.annotationFormat.alreadyAtBack"),
    selectionPane: t("toolbar.selectionPane"),
    editVertices: t("pdf.annotationFormat.editVertices"),
    exitVertexEdit: t("pdf.annotationFormat.exitVertexEdit"),
    mask: t("pdf.annotationFormat.mask"),
    delete: t("pdf.annotationFormat.delete"),
    deleteShape: t("pdfShape.delete"),
    outlineColor: t("pdf.annotationFormat.outlineColor"),
    inkColor: t("pdf.annotationFormat.inkColor"),
    lineWidth: t("pdf.annotationFormat.lineWidth"),
    lineStyle: t("pdf.annotationFormat.lineStyle"),
    lineStart: t("pdf.annotationFormat.lineStart"),
    lineEnd: t("pdf.annotationFormat.lineEnd"),
    lineStyleLabels: {
      solid: t("pdf.annotationFormat.lineStyleSolid"),
      dashed: t("pdf.annotationFormat.lineStyleDashed"),
      dotted: t("pdf.annotationFormat.lineStyleDotted"),
    },
    lineEndingLabels: {
      none: t("pdf.annotationFormat.lineEndingNone"),
      openArrow: t("pdf.annotationFormat.lineEndingOpenArrow"),
      closedArrow: t("pdf.annotationFormat.lineEndingClosedArrow"),
      circle: t("pdf.annotationFormat.lineEndingCircle"),
      square: t("pdf.annotationFormat.lineEndingSquare"),
    },
    fill: t("pdf.annotationFormat.fill"),
    fillColor: t("pdf.annotationFormat.fillColor"),
    fillOpacity: t("pdf.annotationFormat.fillOpacity"),
    opacity: t("pdf.annotationFormat.opacity"),
    blendMode: t("pdf.annotationFormat.blendMode"),
    blendModeLabels: {
      normal: t("pdf.annotationFormat.blendModeNormal"),
      multiply: t("pdf.annotationFormat.blendModeMultiply"),
      screen: t("pdf.annotationFormat.blendModeScreen"),
      overlay: t("pdf.annotationFormat.blendModeOverlay"),
      darken: t("pdf.annotationFormat.blendModeDarken"),
      lighten: t("pdf.annotationFormat.blendModeLighten"),
      "color-dodge": t("pdf.annotationFormat.blendModeColorDodge"),
      "color-burn": t("pdf.annotationFormat.blendModeColorBurn"),
      "hard-light": t("pdf.annotationFormat.blendModeHardLight"),
      "soft-light": t("pdf.annotationFormat.blendModeSoftLight"),
      difference: t("pdf.annotationFormat.blendModeDifference"),
      exclusion: t("pdf.annotationFormat.blendModeExclusion"),
      hue: t("pdf.annotationFormat.blendModeHue"),
      saturation: t("pdf.annotationFormat.blendModeSaturation"),
      color: t("pdf.annotationFormat.blendModeColor"),
      luminosity: t("pdf.annotationFormat.blendModeLuminosity"),
    },
    cornerRadius: t("pdf.annotationFormat.cornerRadius"),
    pointAnchor: t("pdf.annotationFormat.pointAnchor"),
    backToEditing: t("pdf.annotationFormat.backToEditing"),
    pageNumber: (page) => t("pdf.annotationFormat.pageNumber", { page }),
    insert: t("pdf.annotationFormat.insert"),
    outlineMenu: t("pdf.annotationFormat.outlineMenu"),
    transparency: t("pdf.annotationFormat.transparency"),
    lineCap: t("pdf.annotationFormat.lineCap"),
    lineJoin: t("pdf.annotationFormat.lineJoin"),
    lineCapLabels: {
      butt: t("pdf.annotationFormat.lineCapButt"),
      round: t("pdf.annotationFormat.lineCapRound"),
      square: t("pdf.annotationFormat.lineCapSquare"),
    },
    lineJoinLabels: {
      miter: t("pdf.annotationFormat.lineJoinMiter"),
      round: t("pdf.annotationFormat.lineJoinRound"),
      bevel: t("pdf.annotationFormat.lineJoinBevel"),
    },
    lockAspectRatio: t("pdf.annotationFormat.lockAspectRatio"),
    rotate: t("pdf.annotationFormat.rotate"),
    rotateQuarterTurnsOnly: t("pdf.annotationFormat.rotateQuarterTurnsOnly"),
    more: t("pdf.annotationFormat.more"),
    exitEditing: t("pdf.annotationFormat.exitEditing"),
    noFill: t("pdf.annotationFormat.noFill"),
    moreColors: t("pdf.annotationFormat.moreColors"),
    moreShapeStyles: t("pdf.annotationFormat.moreShapeStyles"),
    outlineStylesOnly: t("pdf.annotationFormat.outlineStylesOnly"),
    formatShape: t("document.shapeFormatPane.formatShape"),
    themeColors: t("pdf.annotationFormat.themeColors"),
    standardColors: t("pdf.annotationFormat.standardColors"),
    recentColors: t("pdf.annotationFormat.recentColors"),
  };
}

export function formatPdfAttachmentSize(sizeBytes: number): string {
  if (!Number.isFinite(sizeBytes) || sizeBytes < 0) return "—";
  if (sizeBytes < 1024) return `${Math.round(sizeBytes)} B`;
  if (sizeBytes < 1024 * 1024) {
    return `${Math.round((sizeBytes / 1024) * 10) / 10} KB`;
  }
  return `${Math.round((sizeBytes / (1024 * 1024)) * 10) / 10} MB`;
}

function formatPointSize(value: number): string {
  return (Math.round(value * 10) / 10).toFixed(1);
}

/** Stroke widths read to the hundredth: a 0.25 pt hairline is not 0.3. */
function formatLineWidthPoints(value: number): string {
  return String(Math.round(value * 100) / 100);
}

function LineCapGlyph({ cap }: { cap: PdfShapeLineCap }) {
  // A thick stroke ending at the dashed guide: flat stops at it, round and
  // square carry past it by half the width.
  return (
    <svg aria-hidden width={20} height={20} viewBox="0 0 20 20">
      <line x1={13} y1={3} x2={13} y2={17} stroke="currentColor" strokeWidth={1} strokeDasharray="2 2" />
      <line
        x1={3}
        y1={10}
        x2={13}
        y2={10}
        stroke="currentColor"
        strokeWidth={6}
        strokeLinecap={cap}
        opacity={0.8}
      />
    </svg>
  );
}

function LineJoinGlyph({ join }: { join: PdfShapeLineJoin }) {
  return (
    <svg aria-hidden width={20} height={20} viewBox="0 0 20 20">
      <polyline
        points="4,16 4,4 16,4"
        fill="none"
        stroke="currentColor"
        strokeWidth={5}
        strokeLinejoin={join}
        strokeLinecap="butt"
        opacity={0.8}
      />
    </svg>
  );
}

/**
 * A range that previews while dragged and commits once, on release. Each
 * commit is a native PDF write; committing every input event wrote the file
 * dozens of times per drag.
 */
function CommitOnReleaseRange({
  value,
  min,
  max,
  step,
  disabled,
  label,
  disabledReason,
  testId,
  className,
  onCommit,
}: {
  value: number;
  min: number;
  max: number;
  step: number;
  disabled: boolean;
  label: string;
  disabledReason?: string;
  testId: string;
  className?: string;
  onCommit: (value: number) => void;
}) {
  const [draft, setDraft] = useState<number | null>(null);
  const inputRef = useRef<HTMLInputElement | null>(null);
  const valueRef = useRef(value);
  valueRef.current = value;
  const commitRef = useRef(onCommit);
  commitRef.current = onCommit;
  useEffect(() => {
    const input = inputRef.current;
    if (!input) return;
    const commit = () => {
      if (input.disabled) return;
      const next = Number(input.value);
      setDraft(null);
      if (Number.isFinite(next) && next !== valueRef.current) {
        commitRef.current(next);
      }
    };
    input.addEventListener("change", commit);
    return () => input.removeEventListener("change", commit);
  }, []);
  return (
    <input
      ref={inputRef}
      type="range"
      className={className}
      min={min}
      max={max}
      step={step}
      value={draft ?? value}
      disabled={disabled}
      aria-label={disabled ? (disabledReason ?? label) : label}
      data-disabled-reason={disabled ? disabledReason : undefined}
      data-testid={testId}
      title={disabled ? (disabledReason ?? label) : label}
      onChange={(event) => setDraft(Number(event.currentTarget.value))}
    />
  );
}

/**
 * Shape Editing, in the order of the WPS PDF editor's own contextual tab:
 * pointer tools, Insert Shape, Fill and Outline color, transparency, line
 * style and weight, caps and corners, height and width with a ratio lock,
 * rotation with the two flips, and Exit. Commands the product has beyond that
 * tab (z-order, clipboard, duplicate, blend mode, delete) stay reachable from
 * one More menu instead of widening the band.
 */
/** 大小: the selection's width and height in points, written when a value is
 * committed (Enter, leaving the box, a step), with the ratio locked on request,
 * for every placed object whose size is known in points (shapes, signatures,
 * stamps); a point anchor (a note) or a size not known in points shows its
 * page-relative readout instead. One group for the shape tab and the others. */
function PdfAnnotationSizeGroup({
  labels,
  model,
  onSizeChange,
  showPage = false,
}: {
  labels: PdfAnnotationFormatContextRibbonLabels;
  model: PdfAnnotationFormatContextRibbonModel;
  onSizeChange?: (
    axis: "width" | "height",
    valuePoints: number,
    lockAspectRatio?: boolean,
  ) => void;
  showPage?: boolean;
}) {
  const [lockAspectRatio, setLockAspectRatio] = useState(false);
  const sizeDisabled = !model.sizeEnabled || !onSizeChange;
  return (
    <OfficeHomeRibbonGroup
      label={labels.size}
      ribbon="annotation-format"
      controlId="PDFCTL.CONTEXT.ANNOTATION.SIZE"
      className={styles.group}
      testId="pdf-annotation-context-size-group"
    >
      {model.widthPoints !== undefined && model.heightPoints !== undefined ? (
        <div className={styles.shapeBand}>
          <div className={styles.fieldStack}>
            {(
              [
                ["height", labels.height, model.heightPoints],
                ["width", labels.width, model.widthPoints],
              ] as const
            ).map(([axis, axisLabel, value]) => (
              <OfficeRibbonNumberField
                key={axis}
                classNames={{
              field: styles.captionField,
              icon: styles.noIcon,
              label: styles.fieldLabel,
            }}
                controlId={`PDFCTL.CONTEXT.SHAPE.${axis.toUpperCase()}`}
                disabled={sizeDisabled}
                disabledReason={model.sizeDisabledReason}
                icon={null}
                label={sizeDisabled ? (model.sizeDisabledReason ?? axisLabel) : axisLabel}
                max={100_000}
                min={0.1}
                step={0.1}
                inputWidthPx={56}
                testId={`pdf-annotation-size-${axis}`}
                unit={labels.points}
                value={formatPointSize(value)}
                visibleLabel={axisLabel}
                resyncKey={model.sizeResyncKey}
                onCommit={(raw) => {
                  const next = Number.parseFloat(raw);
                  if (
                    !Number.isFinite(next) ||
                    next <= 0 ||
                    Math.abs(next - value) < 0.05
                  ) {
                    return;
                  }
                  onSizeChange?.(axis, next, lockAspectRatio);
                }}
              />
            ))}
          </div>
          <OfficeHomeCommand
            activation="click"
            active={lockAspectRatio}
            controlId="PDFCTL.CONTEXT.SHAPE.LOCK_ASPECT_RATIO"
            disabled={sizeDisabled}
            disabledReason={model.sizeDisabledReason}
            icon={<OfficeSemanticIcon semanticId="object.lockAspectRatio" size={18} />}
            label={labels.lockAspectRatio}
            testId="pdf-annotation-size-lock"
            variant="icon"
            onClick={() => setLockAspectRatio((locked) => !locked)}
          />
        </div>
      ) : (
        <OfficeHomeCommandRow testId="pdf-annotation-context-size-readout">
          <Maximize2 aria-hidden size={16} />
          <span className={styles.measurement}>
            {model.pointAnchor
              ? labels.pointAnchor
              : `${model.width.toFixed(1)} × ${model.height.toFixed(1)}%`}
          </span>
          {showPage ? (
            <span className={styles.page}>{labels.pageNumber(model.page)}</span>
          ) : null}
        </OfficeHomeCommandRow>
      )}
    </OfficeHomeRibbonGroup>
  );
}

/** The dash pattern a line style draws with, as the menu's preview. */
function LineStylePreview({ style }: { style: ShapeLineStyle }) {
  const dash = style === "dashed" ? "8 4" : style === "dotted" ? "2 3" : undefined;
  return (
    <svg aria-hidden width={40} height={12} viewBox="0 0 40 12">
      <line x1={2} y1={6} x2={38} y2={6} stroke="currentColor" strokeWidth={2} strokeDasharray={dash} strokeLinecap={style === "dotted" ? "round" : "butt"} />
    </svg>
  );
}

/** A line ending drawn at the line's start or end, as the menu's preview. */
function LineEndingPreview({ ending, atStart }: { ending: ShapeLineEnding; atStart: boolean }) {
  const head = (() => {
    switch (ending) {
      case "openArrow":
        return <path d="M12 2 L4 6 L12 10" fill="none" stroke="currentColor" strokeWidth={2} />;
      case "closedArrow":
        return <path d="M12 1.5 L3 6 L12 10.5 Z" fill="currentColor" />;
      case "circle":
        return <circle cx={7} cy={6} r={4} fill="currentColor" />;
      case "square":
        return <rect x={3} y={2} width={8} height={8} fill="currentColor" />;
      default:
        return null;
    }
  })();
  return (
    <svg aria-hidden width={40} height={12} viewBox="0 0 40 12" style={atStart ? undefined : { transform: "scaleX(-1)" }}>
      <line x1={ending === "none" ? 2 : 8} y1={6} x2={38} y2={6} stroke="currentColor" strokeWidth={2} />
      {head}
    </svg>
  );
}

function PdfShapeEditingGroups({
  labels,
  model,
  status,
  pointerTools,
  insertShapes,
  onCopy,
  onPaste,
  onDuplicate,
  onArrangeShape,
  onAlignShapes,
  onMergeShapes,
  onReorder,
  onToggleVertexEdit,
  onSelectionPane,
  selectionPaneOpen,
  onMaskShape,
  onRotateBy,
  onFormatPane,
  onStyleChange,
  onSizeChange,
  onDelete,
  onExit,
}: {
  labels: PdfAnnotationFormatContextRibbonLabels;
  model: PdfAnnotationFormatContextRibbonModel;
  status?: ReactNode;
  pointerTools?: ReactNode;
  insertShapes?: ReactNode;
  onCopy: () => void;
  onPaste: () => void;
  onDuplicate: () => void;
  onArrangeShape: (action: ShapeGeometryTransform) => void;
  onAlignShapes?: (command: PdfShapeAlignCommand) => void;
  /** 合并形状 on the selected source vectors (the first selected is kept). */
  onMergeShapes?: (operation: PdfNativeVectorMergeOperation) => void;
  onReorder?: (direction: "forward" | "backward" | "front" | "back") => void;
  onToggleVertexEdit?: () => void;
  /** Selection Pane (Office: Arrange > Selection Pane, Alt+F10). */
  onSelectionPane?: () => void;
  selectionPaneOpen?: boolean;
  onMaskShape?: () => void;
  onRotateBy?: (degrees: number) => void;
  /** Format Shape (the Shape Styles group's launcher): the pane beside. */
  onFormatPane?: () => void;
  onStyleChange: (change: PdfAnnotationStyleChange) => void;
  onSizeChange?: (
    axis: "width" | "height",
    valuePoints: number,
    lockAspectRatio?: boolean,
  ) => void;
  onDelete: () => void;
  onExit: () => void;
}) {
  const [palette, setPalette] = useState<"fill" | "outline" | null>(null);
  const fillAnchorRef = useRef<HTMLButtonElement | null>(null);
  const outlineAnchorRef = useRef<HTMLButtonElement | null>(null);
  // The angle field reads the object's angle when it is known (a source
  // vector turned in this session) and is otherwise a command that reads 0
  // and rotates by what is typed. Remounting the field after each commit
  // clears what was typed (a controlled draft would equal the value by the
  // time the field commits, so it would never commit).
  const [rotationCommits, setRotationCommits] = useState(0);
  const styleDisabled = (model.busy && !model.styleQueues) || !model.styleEnabled;
  const styleReason = styleDisabled
    ? (model.styleDisabledReason ?? labels.style)
    : undefined;
  const reasonOr = (label: string) => styleReason ?? label;
  const strokeDisabled = styleDisabled || model.strokeSupported === false;
  const fillDisabled =
    styleDisabled ||
    !model.fillSupported ||
    model.fillToggleEnabled === false;
  const fillReason =
    styleReason ??
    (!model.fillSupported || model.fillToggleEnabled === false
      ? (model.fillToggleDisabledReason ?? labels.fill)
      : undefined);
  const capJoinDisabled =
    strokeDisabled || model.capJoinSupported !== true;
  const capJoinReason =
    styleReason ??
    (model.capJoinSupported !== true
      ? (model.capJoinDisabledReason ?? labels.lineCap)
      : undefined);
  const opacityDisabled = styleDisabled || model.opacitySupported === false;
  const opacityUnknown = model.opacitySupported === false;
  const opacityReason =
    styleReason ?? (opacityUnknown ? model.opacityDisabledReason : undefined);
  const transparencyPercent = Math.round((1 - (model.opacity ?? 1)) * 100);
  const commitTransparency = (percent: number) => {
    const bounded = Math.min(100, Math.max(0, percent));
    onStyleChange({ opacity: Math.round((1 - bounded / 100) * 100) / 100 });
  };
  const arrangeDisabled = model.busy || !model.arrangeEnabled;
  const paletteLabels = {
    themeColors: labels.themeColors,
    standardColors: labels.standardColors,
    recentColors: labels.recentColors,
    moreColors: labels.moreColors,
    noFill: labels.noFill,
  };
  const shownRotation =
    model.rotationDegrees === undefined
      ? 0
      : Math.round((((model.rotationDegrees % 360) + 360) % 360) * 100) / 100;
  const commitRotation = (raw: string) => {
    setRotationCommits((count) => count + 1);
    const typed = Number.parseFloat(raw);
    if (!Number.isFinite(typed) || !onRotateBy) return;
    // A known angle: turn by what takes the object to the angle typed.
    const degrees = model.rotationDegrees === undefined ? typed : typed - shownRotation;
    if (degrees === 0) return;
    const bounded = Math.max(-360, Math.min(360, degrees));
    onRotateBy(
      model.freeRotation ? bounded : Math.round(bounded / 90) * 90,
    );
  };
  return (
    <>
      {pointerTools}

      {insertShapes ? (
        <OfficeHomeRibbonGroup
          label={labels.insert}
          ribbon="annotation-format"
          controlId="PDFCTL.CONTEXT.SHAPE.INSERT"
          className={styles.group}
          testId="pdf-annotation-context-insert-group"
        >
          {insertShapes}
        </OfficeHomeRibbonGroup>
      ) : null}

      <OfficeHomeRibbonGroup
        label={labels.style}
        ribbon="annotation-format"
        controlId="PDFCTL.CONTEXT.ANNOTATION.STYLE"
        className={styles.group}
        testId="pdf-annotation-context-style-group"
        launcher={onFormatPane ? {
          label: labels.formatShape,
          onClick: onFormatPane,
          active: Boolean(model.formatPaneOpen),
          testId: "pdf-annotation-format-pane-launcher",
        } : undefined}
      >
        <div className={styles.shapeBand}>
          {/* Office's Shape Styles gallery: one pick sets fill, outline and
              width together (the shared gallery the other editors use). */}
          <OfficeShapeStyleGallery
            application="pdf"
            recipes={PDF_ANNOTATION_STYLE_RECIPES}
            quickRow="outline"
            quickCount={3}
            label={labels.style}
            moreLabel={labels.moreShapeStyles}
            value={pdfAnnotationStyleRecipeId(model)}
            disabled={strokeDisabled}
            getDisabledReason={(recipe) =>
              recipe.row !== "outline" && fillDisabled
                ? (fillReason ?? labels.outlineStylesOnly)
                : undefined
            }
            onCommit={(recipe) =>
              onStyleChange({
                outlineColor: recipe.stroke,
                lineWidth: recipe.strokeWidth,
                ...(fillDisabled
                  ? {}
                  : { fillColor: recipe.row === "outline" ? null : recipe.fill }),
              })
            }
            testIdPrefix="pdf-annotation-style-gallery"
          />
          <span className={styles.paletteHost}>
            <OfficeHomeMenuTrigger
              ref={fillAnchorRef}
              variant="large"
              hasMenuIndicator
              active={palette === "fill"}
              aria-expanded={palette === "fill"}
              aria-haspopup="dialog"
              controlId="PDFCTL.CONTEXT.SHAPE.FILL"
              disabled={fillDisabled}
              disabledReason={fillReason}
              displayLabel={labels.fill}
              icon={
                <span className="note-editor-office-toolbar__color-glyph">
                  <OfficeSemanticIcon semanticId="shape.fill" size={24} />
                  <span
                    className="note-editor-office-toolbar__color-underline"
                    style={{
                      backgroundColor: model.fillColor ?? "transparent",
                    }}
                  />
                </span>
              }
              label={labels.fill}
              onClick={() =>
                setPalette((current) => (current === "fill" ? null : "fill"))
              }
              testId="pdf-annotation-style-fill"
            />
            {palette === "fill" ? (
              <OfficeColorPalettePopover
                anchorRef={fillAnchorRef}
                allowNoFill
                labels={paletteLabels}
                noFillSelected={model.fillColor == null}
                testIdPrefix="pdf-annotation-style-fill-palette"
                value={model.fillColor ?? undefined}
                onDismiss={() => setPalette(null)}
                onPick={(color) => {
                  setPalette(null);
                  onStyleChange({ fillColor: color });
                }}
              />
            ) : null}
          </span>
          <span className={styles.paletteHost}>
            <OfficeHomeMenuTrigger
              ref={outlineAnchorRef}
              variant="large"
              hasMenuIndicator
              active={palette === "outline"}
              aria-expanded={palette === "outline"}
              aria-haspopup="dialog"
              controlId="PDFCTL.CONTEXT.SHAPE.OUTLINE"
              disabled={strokeDisabled}
              disabledReason={styleReason}
              displayLabel={labels.outlineMenu}
              icon={
                <span className="note-editor-office-toolbar__color-glyph">
                  <OfficeSemanticIcon semanticId="shape.outline" size={24} />
                  <span
                    className="note-editor-office-toolbar__color-underline"
                    style={{
                      backgroundColor: model.outlineColor ?? "transparent",
                    }}
                  />
                </span>
              }
              label={labels.outlineMenu}
              onClick={() =>
                setPalette((current) =>
                  current === "outline" ? null : "outline",
                )
              }
              testId="pdf-annotation-style-outline"
            />
            {palette === "outline" ? (
              <OfficeColorPalettePopover
                anchorRef={outlineAnchorRef}
                labels={paletteLabels}
                testIdPrefix="pdf-annotation-style-outline-palette"
                value={model.outlineColor ?? undefined}
                onDismiss={() => setPalette(null)}
                onPick={(color) => {
                  setPalette(null);
                  if (color) onStyleChange({ outlineColor: color });
                }}
              />
            ) : null}
          </span>

          <div className={styles.fieldStack}>
            <OfficeRibbonNumberField
              classNames={{
                field: styles.captionField,
                icon: styles.noIcon,
                label: styles.fieldLabel,
              }}
              controlId="PDFCTL.CONTEXT.SHAPE.TRANSPARENCY"
              disabled={opacityDisabled}
              disabledReason={opacityReason}
              icon={null}
              label={opacityReason ?? labels.transparency}
              max={100}
              min={0}
              step={1}
              inputWidthPx={44}
              testId="pdf-annotation-style-transparency"
              unit="%"
              value={opacityUnknown ? "" : String(transparencyPercent)}
              visibleLabel={labels.transparency}
              onCommit={(raw) => {
                const next = Number.parseFloat(raw);
                if (Number.isFinite(next) && next !== transparencyPercent) {
                  commitTransparency(next);
                }
              }}
            />
            <CommitOnReleaseRange
              className={styles.slider}
              disabled={opacityDisabled}
              disabledReason={opacityReason}
              label={labels.transparency}
              max={100}
              min={0}
              step={1}
              testId="pdf-annotation-style-opacity"
              value={transparencyPercent}
              onCommit={commitTransparency}
            />
          </div>

          <div className={styles.fieldStack}>
            <label className={styles.fieldRow} title={reasonOr(labels.lineStyle)}>
              <span className={styles.fieldLabel}>{labels.lineStyle}</span>
              <PdfRibbonChoiceMenu
                controlId="PDFCTL.CONTEXT.SHAPE.LINE_STYLE"
                disabled={strokeDisabled}
                disabledReason={strokeDisabled ? styleReason : undefined}
                label={labels.lineStyle}
                options={SHAPE_LINE_STYLES.map((style) => ({
                  value: style,
                  label: labels.lineStyleLabels[style],
                  preview: <LineStylePreview style={style} />,
                }))}
                testId="pdf-annotation-style-line-style"
                value={model.lineStyle ?? "solid"}
                onPick={(style) => onStyleChange({ lineStyle: style as ShapeLineStyle })}
              />
            </label>
            <OfficeRibbonNumberField
              classNames={{
                field: styles.captionField,
                icon: styles.noIcon,
                label: styles.fieldLabel,
              }}
              controlId="PDFCTL.CONTEXT.SHAPE.LINE_WIDTH"
              disabled={strokeDisabled}
              disabledReason={styleReason}
              icon={null}
              label={reasonOr(labels.lineWidth)}
              max={72}
              min={0.1}
              step={0.25}
              inputWidthPx={48}
              testId="pdf-annotation-style-width"
              unit={labels.points}
              value={formatLineWidthPoints(model.lineWidth ?? 0.6)}
              visibleLabel={labels.lineWidth}
              onCommit={(raw) => {
                const next = Number.parseFloat(raw);
                if (
                  Number.isFinite(next) &&
                  next > 0 &&
                  Math.abs(next - (model.lineWidth ?? 0.6)) >= 0.005
                ) {
                  onStyleChange({ lineWidth: next });
                }
              }}
            />
          </div>

          <div className={styles.fieldStack}>
            {(
              [
                ["cap", labels.lineCap, PDF_SHAPE_LINE_CAPS, model.lineCap ?? "butt"],
                ["join", labels.lineJoin, PDF_SHAPE_LINE_JOINS, model.lineJoin ?? "miter"],
              ] as const
            ).map(([family, familyLabel, options, current]) => (
              <div
                className={styles.fieldRow}
                key={family}
                role="radiogroup"
                aria-label={capJoinReason ?? familyLabel}
              >
                <span className={styles.fieldLabel}>{familyLabel}</span>
                <span className={styles.segmented}>
                  {options.map((option) => {
                    const optionLabel =
                      family === "cap"
                        ? labels.lineCapLabels[option as PdfShapeLineCap]
                        : labels.lineJoinLabels[option as PdfShapeLineJoin];
                    return (
                      <OfficeHomeCommand
                        key={option}
                        activation="click"
                        active={current === option}
                        controlId={`PDFCTL.CONTEXT.SHAPE.${family === "cap" ? "LINE_CAP" : "LINE_JOIN"}`}
                        disabled={capJoinDisabled}
                        disabledReason={capJoinReason}
                        icon={
                          family === "cap" ? (
                            <LineCapGlyph cap={option as PdfShapeLineCap} />
                          ) : (
                            <LineJoinGlyph join={option as PdfShapeLineJoin} />
                          )
                        }
                        label={`${familyLabel}: ${optionLabel}`}
                        testId={`pdf-annotation-style-${family}-${option}`}
                        variant="icon"
                        onClick={() =>
                          onStyleChange(
                            family === "cap"
                              ? { lineCap: option as PdfShapeLineCap }
                              : { lineJoin: option as PdfShapeLineJoin },
                          )
                        }
                      />
                    );
                  })}
                </span>
              </div>
            ))}
          </div>

          {model.cornerSupported || model.fillOpacitySupported ? (
            <div className={styles.fieldStack}>
              {model.cornerSupported ? (
                <label className={styles.fieldRow}>
                  <span className={styles.fieldLabel}>{labels.cornerRadius}</span>
                  <CommitOnReleaseRange
                    className={styles.slider}
                    disabled={styleDisabled}
                    disabledReason={styleReason}
                    label={labels.cornerRadius}
                    max={50}
                    min={0}
                    step={1}
                    testId="pdf-annotation-style-corner-radius"
                    value={model.cornerRadius ?? 0}
                    onCommit={(next) => onStyleChange({ cornerRadius: next })}
                  />
                </label>
              ) : null}
              {model.fillOpacitySupported ? (
                <label className={styles.fieldRow}>
                  <span className={styles.fieldLabel}>{labels.fillOpacity}</span>
                  <CommitOnReleaseRange
                    className={styles.slider}
                    disabled={styleDisabled || model.fillColor == null}
                    disabledReason={styleReason ?? labels.fill}
                    label={labels.fillOpacity}
                    max={1}
                    min={0.05}
                    step={0.05}
                    testId="pdf-annotation-style-fill-opacity"
                    value={model.fillOpacity ?? 0.25}
                    onCommit={(next) => onStyleChange({ fillOpacity: next })}
                  />
                </label>
              ) : null}
            </div>
          ) : null}

          {model.strokeSupported !== false && model.lineEndingsSupported ? (
            <div className={styles.fieldStack}>
              {(
                [
                  ["start", labels.lineStart, model.lineStart ?? "none"],
                  ["end", labels.lineEnd, model.lineEnd ?? "none"],
                ] as const
              ).map(([edge, edgeLabel, value]) => (
                <label className={styles.fieldRow} key={edge}>
                  <span className={styles.fieldLabel}>{edgeLabel}</span>
                  <PdfRibbonChoiceMenu
                    controlId={`PDFCTL.CONTEXT.SHAPE.LINE_${edge.toUpperCase()}`}
                    disabled={styleDisabled}
                    disabledReason={styleDisabled ? styleReason : undefined}
                    label={edgeLabel}
                    options={SHAPE_LINE_ENDINGS.map((ending) => ({
                      value: ending,
                      label: labels.lineEndingLabels[ending],
                      preview: <LineEndingPreview ending={ending} atStart={edge === "start"} />,
                    }))}
                    testId={`pdf-annotation-style-line-${edge}`}
                    value={value}
                    onPick={(ending) =>
                      onStyleChange({
                        [edge === "start" ? "lineStart" : "lineEnd"]: ending as ShapeLineEnding,
                      })
                    }
                  />
                </label>
              ))}
            </div>
          ) : null}
        </div>
      </OfficeHomeRibbonGroup>

      <PdfAnnotationSizeGroup labels={labels} model={model} onSizeChange={onSizeChange} />

      <OfficeHomeRibbonGroup
        label={labels.arrange}
        ribbon="annotation-format"
        controlId="PDFCTL.CONTEXT.ANNOTATION.ARRANGE"
        className={styles.group}
        testId="pdf-annotation-context-arrange-group"
      >
        <span className={styles.arrangeBand}>
        <div className={styles.fieldStack}>
          <OfficeRibbonNumberField
            classNames={{
                field: styles.captionField,
                icon: styles.noIcon,
                label: styles.fieldLabel,
              }}
            controlId="PDFCTL.CONTEXT.SHAPE.ROTATE"
            disabled={arrangeDisabled || !onRotateBy}
            disabledReason={model.arrangeDisabledReason}
            icon={null}
            label={
              arrangeDisabled
                ? (model.arrangeDisabledReason ?? labels.rotate)
                : model.freeRotation
                  ? labels.rotate
                  : `${labels.rotate} — ${labels.rotateQuarterTurnsOnly}`
            }
            max={360}
            min={-360}
            step={model.freeRotation ? 15 : 90}
            inputWidthPx={44}
            key={`${rotationCommits}:${shownRotation}`}
            testId="pdf-annotation-rotate"
            unit="°"
            value={String(shownRotation)}
            visibleLabel={labels.rotate}
            onCommit={commitRotation}
          />
          <span className={styles.fieldRow}>
            {(
              [
                ["flip-horizontal", labels.flipHorizontal, "object.flipHorizontal"],
                ["flip-vertical", labels.flipVertical, "object.flipVertical"],
              ] as const
            ).map(([action, actionLabel, semanticId]) => (
              <OfficeHomeCommand
                key={action}
                activation="click"
                controlId={`PDFCTL.CONTEXT.SHAPE.${action.toUpperCase().replace("-", "_")}`}
                disabled={arrangeDisabled}
                disabledReason={model.arrangeDisabledReason}
                icon={<OfficeSemanticIcon semanticId={semanticId} size={18} />}
                label={actionLabel}
                testId={`pdf-annotation-shape-${action}`}
                variant="icon"
                onClick={() => onArrangeShape(action)}
              />
            ))}
            {onAlignShapes ? (
              <DropdownMenu modal={false}>
                <DropdownMenuTrigger asChild>
                  <OfficeHomeMenuTrigger
                    controlId="PDFCTL.CONTEXT.SHAPE.ALIGN"
                    disabled={model.busy || !model.alignEnabled}
                    disabledReason={model.alignEnabled ? undefined : model.arrangeDisabledReason}
                    hasMenuIndicator
                    icon={<OfficeSemanticIcon semanticId="object.align" size={18} />}
                    label={labels.align}
                    testId="pdf-annotation-shape-align"
                    variant="icon"
                  />
                </DropdownMenuTrigger>
                <RibbonDropdownMenuContent
                  align="start"
                  aria-label={labels.align}
                  data-testid="pdf-annotation-shape-align-menu"
                >
                  <DropdownMenuLabel>
                    {(model.selectionCount ?? 1) > 1 ? labels.alignSelected : labels.alignToPage}
                  </DropdownMenuLabel>
                  {(
                    [
                      ["left", "object.align.left"],
                      ["hcenter", "object.align.center"],
                      ["right", "object.align.right"],
                      ["top", "object.align.top"],
                      ["vcenter", "object.align.middle"],
                      ["bottom", "object.align.bottom"],
                    ] as const
                  ).map(([edge, semanticId]) => (
                    <DropdownMenuItem
                      key={edge}
                      data-testid={`pdf-annotation-shape-align-${edge}`}
                      onSelect={() => onAlignShapes({ kind: "align", edge })}
                    >
                      <OfficeSemanticIcon semanticId={semanticId} size={16} />
                      <span>{labels.alignEdges[edge]}</span>
                    </DropdownMenuItem>
                  ))}
                  <DropdownMenuSeparator />
                  {(
                    [
                      ["horizontal", labels.distributeHorizontally, "object.distribute.horizontal"],
                      ["vertical", labels.distributeVertically, "object.distribute.vertical"],
                    ] as const
                  ).map(([axis, axisLabel, semanticId]) => (
                    <DropdownMenuItem
                      key={axis}
                      data-testid={`pdf-annotation-shape-distribute-${axis}`}
                      disabled={!pdfShapeAlignAvailable({ kind: "distribute", axis }, model.selectionCount ?? 1)}
                      onSelect={() => onAlignShapes({ kind: "distribute", axis })}
                    >
                      <OfficeSemanticIcon semanticId={semanticId} size={16} />
                      <span>{axisLabel}</span>
                    </DropdownMenuItem>
                  ))}
                </RibbonDropdownMenuContent>
              </DropdownMenu>
            ) : null}
            {onMergeShapes ? (
              <DropdownMenu modal={false}>
                <DropdownMenuTrigger asChild>
                  <OfficeHomeMenuTrigger
                    controlId="PDFCTL.CONTEXT.SHAPE.MERGE"
                    disabled={model.busy || !model.mergeEnabled}
                    disabledReason={model.mergeEnabled ? undefined : model.mergeDisabledReason}
                    hasMenuIndicator
                    icon={<OfficeSemanticIcon semanticId="shape.change" size={18} />}
                    label={labels.mergeShapes}
                    testId="pdf-annotation-shape-merge"
                    variant="icon"
                  />
                </DropdownMenuTrigger>
                <RibbonDropdownMenuContent
                  align="start"
                  aria-label={labels.mergeShapes}
                  data-testid="pdf-annotation-shape-merge-menu"
                >
                  {/* PowerPoint's five, in its order (Microsoft Support,
                      "Combine shapes"). */}
                  {(["union", "combine", "fragment", "intersect", "subtract"] as const).map((operation) => (
                    <DropdownMenuItem
                      key={operation}
                      data-testid={`pdf-annotation-shape-merge-${operation}`}
                      onSelect={() => onMergeShapes(operation)}
                    >
                      {labels.mergeOperations[operation]}
                    </DropdownMenuItem>
                  ))}
                </RibbonDropdownMenuContent>
              </DropdownMenu>
            ) : null}
            {model.nativeSource && onToggleVertexEdit ? (
              <OfficeHomeCommand
                activation="click"
                active={Boolean(model.vertexEditing)}
                controlId="PDFCTL.CONTEXT.SHAPE.EDIT_VERTICES"
                disabled={model.busy || !model.vertexEditEnabled}
                disabledReason={
                  model.vertexEditEnabled ? undefined : model.vertexEditDisabledReason
                }
                icon={<PenTool aria-hidden size={18} />}
                label={model.vertexEditing ? labels.exitVertexEdit : labels.editVertices}
                testId="pdf-annotation-shape-edit-vertices"
                variant="icon"
                onClick={onToggleVertexEdit}
              />
            ) : null}
          </span>
        </div>
        {model.nativeSource && onReorder ? (
          // Office's Arrange group: "Bring Forward" and "Send Backward" split
          // buttons whose menus hold the one-step and the all-the-way move.
          // A source vector's layer is its place in the page content stream.
          <OfficeHomeCommandStack className={`${styles.layerStack} ${ribbonLayout.keepsWidth}`}>
            {(
              [
                ["layer", "forward", "front", labels.bringForward, labels.bringToFront, "object.bringForward"],
                ["send", "backward", "back", labels.sendBackward, labels.sendToBack, "object.sendBackward"],
              ] as const
            ).map(([key, stepDirection, endDirection, stepLabel, endLabel, semanticId]) => {
              const atEnd =
                key === "layer" ? model.reorderAtFront === true : model.reorderAtBack === true;
              const atEndReason = key === "layer" ? labels.alreadyAtFront : labels.alreadyAtBack;
              const disabled = model.busy || !model.reorderEnabled || atEnd;
              const reason = model.busy
                ? undefined
                : !model.reorderEnabled
                  ? model.reorderDisabledReason
                  : atEnd
                    ? atEndReason
                    : undefined;
              const stepId = stepDirection === "forward" ? "bring-forward" : "send-backward";
              const endId = endDirection === "front" ? "bring-to-front" : "send-to-back";
              const controlId = `PDFCTL.CONTEXT.SHAPE.${stepId.toUpperCase().replace(/-/g, "_")}`;
              return (
                <OfficeHomeSplitLayout
                  key={key}
                  className={styles.layerSplit}
                  testId={`pdf-annotation-shape-${key}-split`}
                  primary={
                    <OfficeHomeCommand
                      activation="click"
                      controlId={controlId}
                      disabled={disabled}
                      disabledReason={reason}
                      displayLabel={stepLabel}
                      icon={<OfficeSemanticIcon semanticId={semanticId} size={18} />}
                      label={stepLabel}
                      testId={`pdf-annotation-shape-${stepId}`}
                      variant="compact"
                      onClick={() => onReorder(stepDirection)}
                    />
                  }
                  trigger={
                    <DropdownMenu modal={false}>
                      <DropdownMenuTrigger asChild>
                        <OfficeHomeMenuTrigger
                          controlId={controlId}
                          disabled={disabled}
                          disabledReason={reason}
                          icon={<OfficeSemanticIcon semanticId="menu.chevron.down" size={12} />}
                          label={stepLabel}
                          testId={`pdf-annotation-shape-${key}-menu`}
                          variant="icon"
                        />
                      </DropdownMenuTrigger>
                      <RibbonDropdownMenuContent
                        align="start"
                        aria-label={stepLabel}
                        data-testid={`pdf-annotation-shape-${key}-menu-content`}
                      >
                        <DropdownMenuItem
                          data-testid={`pdf-annotation-shape-${stepId}-item`}
                          onSelect={() => onReorder(stepDirection)}
                        >
                          <OfficeSemanticIcon semanticId={semanticId} size={16} />
                          <span>{stepLabel}</span>
                        </DropdownMenuItem>
                        <DropdownMenuItem
                          data-testid={`pdf-annotation-shape-${endId}`}
                          onSelect={() => onReorder(endDirection)}
                        >
                          <OfficeSemanticIcon
                            semanticId={endDirection === "front" ? "object.bringToFront" : "object.sendToBack"}
                            size={16}
                          />
                          <span>{endLabel}</span>
                        </DropdownMenuItem>
                      </RibbonDropdownMenuContent>
                    </DropdownMenu>
                  }
                />
              );
            })}
          </OfficeHomeCommandStack>
        ) : null}
        {model.nativeSource && onSelectionPane && labels.selectionPane ? (
          <OfficeHomeCommand
            activation="click"
            controlId="PDFCTL.CONTEXT.SHAPE.SELECTION_PANE"
            displayLabel={labels.selectionPane}
            icon={<OfficeSemanticIcon semanticId="object.selectionPane" size={18} />}
            label={labels.selectionPane}
            active={selectionPaneOpen === true}
            testId="pdf-annotation-shape-selection-pane"
            variant="compact"
            onClick={onSelectionPane}
          />
        ) : null}
        </span>
      </OfficeHomeRibbonGroup>

      <OfficeHomeRibbonGroup
        label={labels.actions}
        ribbon="annotation-format"
        controlId="PDFCTL.CONTEXT.ANNOTATION.ACTIONS"
        className={styles.group}
        testId="pdf-annotation-context-actions-group"
      >
        <div className={styles.shapeBand}>
        {/* 删除 in the open, beside 更多 (which keeps it among copy/paste/
            duplicate): a selected shape's delete should not need a menu
            (audit P3). */}
        <OfficeHomeCommand
          activation="click"
          controlId="PDFCTL.CONTEXT.SHAPE.DELETE"
          disabled={model.busy || !model.deleteEnabled}
          disabledReason={model.deleteEnabled ? undefined : model.deleteDisabledReason}
          displayLabel={labels.deleteShape}
          icon={<OfficeSemanticIcon semanticId="edit.delete" size={24} />}
          label={model.deleteDisabledReason ?? labels.deleteShape}
          onClick={onDelete}
          testId="pdf-annotation-shape-delete"
          variant="large"
        />
        <DropdownMenu modal={false}>
          <DropdownMenuTrigger asChild>
            <OfficeHomeMenuTrigger
              variant="large"
              hasMenuIndicator
              controlId="PDFCTL.CONTEXT.SHAPE.MORE"
              disabled={model.busy}
              displayLabel={labels.more}
              icon={<OfficeSemanticIcon semanticId="overflow.more" size={24} />}
              label={labels.more}
              testId="pdf-annotation-context-more"
            />
          </DropdownMenuTrigger>
          <RibbonDropdownMenuContent
            align="end"
            aria-label={labels.more}
            data-testid="pdf-annotation-context-more-menu"
          >
            {(
              [
                ["rotate-left", labels.rotateLeft, RotateCcw],
                ["rotate-right", labels.rotateRight, RotateCw],
              ] as const
            ).map(([action, actionLabel, Icon]) => (
              <DropdownMenuItem
                key={action}
                data-disabled-reason={arrangeDisabled ? model.arrangeDisabledReason : undefined}
                data-testid={`pdf-annotation-shape-${action}`}
                disabled={arrangeDisabled}
                title={arrangeDisabled ? model.arrangeDisabledReason : undefined}
                onSelect={() => onArrangeShape(action)}
              >
                <Icon aria-hidden size={16} />
                <span>{actionLabel}</span>
              </DropdownMenuItem>
            ))}
            {model.nativeSource && onMaskShape ? (
              <DropdownMenuItem
                data-disabled-reason={!model.maskEnabled ? model.maskDisabledReason : undefined}
                data-testid="pdf-annotation-shape-mask"
                disabled={model.busy || !model.maskEnabled}
                title={!model.maskEnabled ? model.maskDisabledReason : undefined}
                onSelect={onMaskShape}
              >
                <Eraser aria-hidden size={16} />
                <span>{labels.mask}</span>
              </DropdownMenuItem>
            ) : null}
            <DropdownMenuSeparator />
            <DropdownMenuItem
              data-disabled-reason={!model.copyEnabled ? model.copyDisabledReason : undefined}
              data-testid="pdf-annotation-copy"
              disabled={model.busy || !model.copyEnabled}
              title={!model.copyEnabled ? model.copyDisabledReason : undefined}
              onSelect={onCopy}
            >
              <Copy aria-hidden size={16} />
              <span>{labels.copy}</span>
            </DropdownMenuItem>
            <DropdownMenuItem
              data-disabled-reason={!model.pasteEnabled ? model.pasteDisabledReason : undefined}
              data-testid="pdf-annotation-paste"
              disabled={model.busy || !model.pasteEnabled}
              title={!model.pasteEnabled ? model.pasteDisabledReason : undefined}
              onSelect={onPaste}
            >
              <ClipboardPaste aria-hidden size={16} />
              <span>{labels.paste}</span>
            </DropdownMenuItem>
            <DropdownMenuItem
              data-disabled-reason={!model.duplicateEnabled ? model.duplicateDisabledReason : undefined}
              data-testid="pdf-annotation-duplicate"
              disabled={model.busy || !model.duplicateEnabled}
              title={!model.duplicateEnabled ? model.duplicateDisabledReason : undefined}
              onSelect={onDuplicate}
            >
              <CopyPlus aria-hidden size={16} />
              <span>{labels.duplicate}</span>
            </DropdownMenuItem>
            {model.blendModeSupported ? (
              <>
                <DropdownMenuSeparator />
                <DropdownMenuSub>
                  <DropdownMenuSubTrigger
                    data-testid="pdf-annotation-style-blend-mode"
                    disabled={styleDisabled}
                    title={styleReason}
                  >
                    <span>{labels.blendMode}</span>
                    <span className={styles.menuValue}>
                      {labels.blendModeLabels[model.blendMode ?? "normal"]}
                    </span>
                  </DropdownMenuSubTrigger>
                  <RibbonDropdownMenuSubContent>
                    <DropdownMenuRadioGroup value={model.blendMode ?? "normal"}>
                      {PDF_STANDARD_BLEND_MODES.map((mode) => (
                        <DropdownMenuRadioItem
                          key={mode}
                          data-testid={`pdf-annotation-style-blend-mode-${mode}`}
                          value={mode}
                          onSelect={() => onStyleChange({ blendMode: mode })}
                        >
                          {labels.blendModeLabels[mode]}
                        </DropdownMenuRadioItem>
                      ))}
                    </DropdownMenuRadioGroup>
                  </RibbonDropdownMenuSubContent>
                </DropdownMenuSub>
              </>
            ) : null}
            <DropdownMenuSeparator />
            <DropdownMenuItem
              data-disabled-reason={!model.deleteEnabled ? model.deleteDisabledReason : undefined}
              data-testid="pdf-annotation-delete-selected"
              disabled={model.busy || !model.deleteEnabled}
              title={!model.deleteEnabled ? model.deleteDisabledReason : undefined}
              onSelect={onDelete}
            >
              {model.deleteEnabled ? (
                <Trash2 aria-hidden size={16} />
              ) : (
                <ShieldAlert aria-hidden size={16} />
              )}
              <span>{model.contentObject ? labels.deleteShape : labels.delete}</span>
            </DropdownMenuItem>
          </RibbonDropdownMenuContent>
        </DropdownMenu>
        <OfficeHomeCommand
          activation="click"
          controlId="PDFCTL.CONTEXT.ANNOTATION.EXIT"
          disabled={model.busy}
          displayLabel={labels.exitEditing}
          icon={<OfficeSemanticIcon semanticId="edit.exit" size={24} />}
          label={labels.exitEditing}
          testId="pdf-annotation-context-exit"
          variant="large"
          onClick={onExit}
        />
        </div>
        {status}
      </OfficeHomeRibbonGroup>
    </>
  );
}

export function PdfAnnotationFormatContextRibbon({
  labels,
  model,
  status,
  onCopy,
  onPaste,
  onDuplicate,
  onArrangeShape,
  onAlignShapes,
  onMergeShapes,
  onReorder,
  onToggleVertexEdit,
  onSelectionPane,
  selectionPaneOpen,
  onMaskShape,
  onStyleChange,
  onSizeChange,
  onDelete,
  onSaveAttachment,
  onExit,
  pointerTools,
  insertShapes,
  onRotateBy,
  onFormatPane,
}: {
  labels: PdfAnnotationFormatContextRibbonLabels;
  model: PdfAnnotationFormatContextRibbonModel;
  status?: ReactNode;
  onCopy: () => void;
  onPaste: () => void;
  onDuplicate: () => void;
  onArrangeShape: (action: ShapeGeometryTransform) => void;
  onAlignShapes?: (command: PdfShapeAlignCommand) => void;
  /** 合并形状 on the selected source vectors (the first selected is kept). */
  onMergeShapes?: (operation: PdfNativeVectorMergeOperation) => void;
  onReorder?: (direction: "forward" | "backward" | "front" | "back") => void;
  onToggleVertexEdit?: () => void;
  /** Selection Pane (Office: Arrange > Selection Pane, Alt+F10). */
  onSelectionPane?: () => void;
  selectionPaneOpen?: boolean;
  onMaskShape?: () => void;
  onStyleChange: (change: PdfAnnotationStyleChange) => void;
  onSizeChange?: (
    axis: "width" | "height",
    valuePoints: number,
    lockAspectRatio?: boolean,
  ) => void;
  onDelete: () => void;
  onSaveAttachment?: () => void;
  onExit: () => void;
  /** Shape Editing leads with the viewer's own pointer tools. */
  pointerTools?: ReactNode;
  /** The Edit tab's Insert Shape split, re-hosted for Shape Editing. */
  insertShapes?: ReactNode;
  onRotateBy?: (degrees: number) => void;
  /** Format Shape: opens or closes the shape's format pane. */
  onFormatPane?: () => void;
}) {
  if (model.kind === "shape") {
    return (
      <PdfShapeEditingGroups
        labels={labels}
        model={model}
        status={status}
        pointerTools={pointerTools}
        insertShapes={insertShapes}
        onCopy={onCopy}
        onPaste={onPaste}
        onDuplicate={onDuplicate}
        onArrangeShape={onArrangeShape}
        onAlignShapes={onAlignShapes}
        onMergeShapes={onMergeShapes}
        onReorder={onReorder}
        onToggleVertexEdit={onToggleVertexEdit}
        onSelectionPane={onSelectionPane}
        selectionPaneOpen={selectionPaneOpen}
        onMaskShape={onMaskShape}
        onRotateBy={onRotateBy}
        onFormatPane={onFormatPane}
        onStyleChange={onStyleChange}
        onSizeChange={onSizeChange}
        onDelete={onDelete}
        onExit={onExit}
      />
    );
  }
  const styleDisabled = (model.busy && !model.styleQueues) || !model.styleEnabled;
  const styleReason = styleDisabled
    ? (model.styleDisabledReason ?? labels.style)
    : undefined;
  const styleAccessibleLabel = (label: string): string => styleReason ?? label;
  const fillControlReason =
    styleReason ?? (model.fillColor == null ? labels.fill : undefined);
  return (
    <>
      {/* No identity panel: WPS's contextual tabs carry none; the tab's own
          label names what is selected (audit A4, WPS CT_*.kuip). */}

      {model.kind === "ink" ? (
        <OfficeHomeRibbonGroup
          label={labels.inkStyle}
          ribbon="annotation-format"
          controlId="PDFCTL.CONTEXT.ANNOTATION.STYLE"
          className={styles.group}
          testId="pdf-annotation-context-style-group"
        >
          <div className={styles.styleControls}>
            {model.strokeSupported !== false ? (
              <>
                <label
                  title={styleAccessibleLabel(
                    model.kind === "ink"
                      ? labels.inkColor
                      : labels.outlineColor,
                  )}
                >
                  <span>
                    {model.kind === "ink"
                      ? labels.inkColor
                      : labels.outlineColor}
                  </span>
                  <input
                    type="color"
                    value={model.outlineColor ?? "#000000"}
                    disabled={styleDisabled}
                    aria-label={styleAccessibleLabel(
                      model.kind === "ink"
                        ? labels.inkColor
                        : labels.outlineColor,
                    )}
                    data-disabled-reason={styleReason}
                    data-testid="pdf-annotation-style-outline"
                    title={styleAccessibleLabel(
                      model.kind === "ink"
                        ? labels.inkColor
                        : labels.outlineColor,
                    )}
                    onChange={(event) =>
                      onStyleChange({ outlineColor: event.target.value })
                    }
                  />
                </label>
                <label title={styleAccessibleLabel(labels.lineWidth)}>
                  <span>{labels.lineWidth}</span>
                  <input
                    type="range"
                    min={0.1}
                    max={12}
                    step={0.1}
                    value={model.lineWidth ?? 0.6}
                    disabled={styleDisabled}
                    aria-label={styleAccessibleLabel(labels.lineWidth)}
                    data-disabled-reason={styleReason}
                    data-testid="pdf-annotation-style-width"
                    title={styleAccessibleLabel(labels.lineWidth)}
                    onChange={(event) =>
                      onStyleChange({ lineWidth: Number(event.target.value) })
                    }
                  />
                </label>
              </>
            ) : null}
            {model.fillSupported ? (
              <>
                <label
                  className={styles.fillToggleLabel}
                  title={styleAccessibleLabel(labels.fill)}
                >
                  <input
                    type="checkbox"
                    className={styles.fillToggleInput}
                    checked={model.fillColor != null}
                    disabled={
                      styleDisabled || model.fillToggleEnabled === false
                    }
                    aria-label={
                      styleReason ??
                      model.fillToggleDisabledReason ??
                      labels.fill
                    }
                    data-disabled-reason={
                      styleReason ?? model.fillToggleDisabledReason
                    }
                    data-testid="pdf-annotation-style-fill-enabled"
                    title={
                      styleReason ??
                      model.fillToggleDisabledReason ??
                      labels.fill
                    }
                    onChange={(event) =>
                      onStyleChange({
                        fillColor: event.target.checked
                          ? (model.fillColor ?? model.outlineColor ?? "#000000")
                          : null,
                      })
                    }
                  />
                  <span
                    aria-hidden
                    className={styles.fillToggleMark}
                  />
                  <span>{labels.fill}</span>
                </label>
                <label title={fillControlReason ?? labels.fillColor}>
                  <span>{labels.fillColor}</span>
                  <input
                    type="color"
                    value={model.fillColor ?? model.outlineColor ?? "#000000"}
                    disabled={styleDisabled || model.fillColor == null}
                    aria-label={fillControlReason ?? labels.fillColor}
                    data-disabled-reason={fillControlReason}
                    data-testid="pdf-annotation-style-fill"
                    title={fillControlReason ?? labels.fillColor}
                    onChange={(event) =>
                      onStyleChange({ fillColor: event.target.value })
                    }
                  />
                </label>
                {model.fillOpacitySupported !== false ? (
                  <label title={fillControlReason ?? labels.fillOpacity}>
                    <span>{labels.fillOpacity}</span>
                    <input
                      type="range"
                      min={0.05}
                      max={1}
                      step={0.05}
                      value={model.fillOpacity ?? 0.25}
                      disabled={styleDisabled || model.fillColor == null}
                      aria-label={fillControlReason ?? labels.fillOpacity}
                      data-disabled-reason={fillControlReason}
                      data-testid="pdf-annotation-style-fill-opacity"
                      title={fillControlReason ?? labels.fillOpacity}
                      onChange={(event) =>
                        onStyleChange({
                          fillOpacity: Number(event.target.value),
                        })
                      }
                    />
                  </label>
                ) : null}
              </>
            ) : null}
          </div>
        </OfficeHomeRibbonGroup>
      ) : null}

      {model.kind === "attachment" && model.attachment ? (
        <OfficeHomeRibbonGroup
          label={labels.attachmentFile}
          ribbon="annotation-format"
          controlId="PDFCTL.CONTEXT.ANNOTATION.ATTACHMENT"
          className={styles.group}
          testId="pdf-attachment-context-file-group"
        >
          <div className={styles.attachmentDetails}>
            <Paperclip aria-hidden size={18} />
            <span className={styles.attachmentName} title={model.attachment.name}>
              {model.attachment.name}
            </span>
            <span className={styles.attachmentMeta}>
              {model.attachment.mimeType} ·{" "}
              {formatPdfAttachmentSize(model.attachment.sizeBytes)}
            </span>
          </div>
          <OfficeHomeCommand
            activation="click"
            controlId="PDFCTL.CONTEXT.ANNOTATION.ATTACHMENT.SAVE"
            disabled={
              model.busy ||
              !model.attachment.saveEnabled ||
              !onSaveAttachment
            }
            icon={<Download size={18} />}
            label={
              model.attachment.saveDisabledReason ?? labels.saveAttachment
            }
            displayLabel={labels.saveAttachment}
            onClick={() => onSaveAttachment?.()}
            testId="pdf-attachment-context-save"
          />
        </OfficeHomeRibbonGroup>
      ) : null}

      <OfficeHomeRibbonGroup
        label={labels.clipboard}
        ribbon="annotation-format"
        controlId="PDFCTL.CONTEXT.ANNOTATION.CLIPBOARD"
        className={styles.group}
        testId="pdf-annotation-context-clipboard-group"
      >
        <OfficeHomeCommand
          activation="click"
          controlId="PDFCTL.CONTEXT.ANNOTATION.COPY"
          disabled={model.busy || !model.copyEnabled}
          icon={<Copy size={18} />}
          label={model.copyDisabledReason ?? labels.copy}
          displayLabel={labels.copy}
          onClick={onCopy}
          testId="pdf-annotation-copy"
        />
        <OfficeHomeCommand
          activation="click"
          controlId="PDFCTL.CONTEXT.ANNOTATION.PASTE"
          disabled={model.busy || !model.pasteEnabled}
          icon={<ClipboardPaste size={18} />}
          label={model.pasteDisabledReason ?? labels.paste}
          displayLabel={labels.paste}
          onClick={onPaste}
          testId="pdf-annotation-paste"
        />
      </OfficeHomeRibbonGroup>

      <PdfAnnotationSizeGroup labels={labels} model={model} onSizeChange={onSizeChange} showPage />

      {/* 操作: duplicate, delete, 退出编辑. A 排列 group holding only 创建副本
          said more than it held (audit A4). */}
      <OfficeHomeRibbonGroup
        label={labels.actions}
        ribbon="annotation-format"
        controlId="PDFCTL.CONTEXT.ANNOTATION.ACTIONS"
        className={styles.group}
        testId="pdf-annotation-context-actions-group"
      >
        <OfficeHomeCommand
          activation="click"
          controlId="PDFCTL.CONTEXT.ANNOTATION.DUPLICATE"
          disabled={model.busy || !model.duplicateEnabled}
          icon={<CopyPlus size={18} />}
          label={model.duplicateDisabledReason ?? labels.duplicate}
          displayLabel={labels.duplicate}
          onClick={onDuplicate}
          testId="pdf-annotation-duplicate"
        />
        <OfficeHomeCommand
          activation="click"
          controlId="PDFCTL.CONTEXT.ANNOTATION.DELETE"
          disabled={model.busy || !model.deleteEnabled}
          icon={
            model.deleteEnabled ? (
              <Trash2 size={18} />
            ) : (
              <ShieldAlert size={18} />
            )
          }
          label={model.deleteDisabledReason ?? (model.contentObject ? labels.deleteShape : labels.delete)}
          displayLabel={model.contentObject ? labels.deleteShape : labels.delete}
          onClick={onDelete}
          testId="pdf-annotation-delete-selected"
        />
        <OfficeHomeCommand
          activation="click"
          controlId="PDFCTL.CONTEXT.ANNOTATION.EXIT"
          disabled={model.busy}
          displayLabel={labels.backToEditing}
          icon={<OfficeSemanticIcon semanticId="edit.exit" size={24} />}
          label={labels.backToEditing}
          onClick={onExit}
          testId="pdf-annotation-context-exit"
          variant="large"
        />
        {status}
      </OfficeHomeRibbonGroup>
    </>
  );
}
