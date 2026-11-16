"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { PdfArmedTextKind } from "@mrnote/core/web/lib/pdf/text-box-create.ts";
import type { PdfInsertRibbonIconName } from "../PdfInsertRibbonIcons";
import type { PdfShapeAlignCommand } from "../../../../../lib/pdf/shape-align";
import { OfficeHomeCommandStack } from "@mrnote/core/web/components/office-ribbon/OfficeHomeRibbon.tsx";
import PdfRibbonButton from "../PdfRibbonButton";
import PdfFileRibbonTab from "../PdfFileRibbonTab";
import PdfInsertRibbonTab from "../PdfInsertRibbonTab";
import PdfEditPointerTools from "../PdfEditPointerTools";
import PdfEditRibbonTab from "../PdfEditRibbonTab";
import PdfCommentRibbonTab from "../PdfCommentRibbonTab";
import { PdfDrawRibbonTab, type PdfDrawRibbonModel } from "../PdfDrawRibbonTab";
import WindowNotice from "@mrnote/core/web/components/notebook/WindowNotice.tsx";
import PdfHomeCompactRibbon, {
  HomeGroup,
  type PdfHomeCompactCommand,
} from "../PdfHomeCompactRibbon";
import type { SnapshotShape } from "../../../../../lib/pdf/snapshot-region";
import PdfRedactionPendingReview from "../PdfRedactionPendingReview";
import PdfProtectRibbonTab from "../PdfProtectRibbonTab";
import PdfSignatureProofStatus from "../PdfSignatureProofStatus";
import { useState } from "react";
import type {
  ReactNode,
  Dispatch,
  SetStateAction,
  RefObject,
  JSX,
} from "react";
import { OfficeTaskPaneDockSlot } from "@mrnote/core/web/components/office-ribbon/OfficeTaskPaneDockSlot.tsx";
import type { OfficeTaskPaneDockRegistry } from "@mrnote/core/web/components/office-ribbon/officeTaskPaneDock.ts";
import { PdfShapeFormatPane } from "../PdfShapeFormatPane";
import {
  PdfOfficeRibbon,
  type PdfRibbonSaveStatus,
  type PdfOfficeRibbonTab,
  type PdfOfficeActionRail,
} from "../PdfOfficeRibbon";
import {
  OfficeRailButton,
} from "@mrnote/core/web/components/office-ribbon/OfficeEditorActionRail.tsx";
import {
  Printer,
  CopyPlus,
  Hash,
  Moon,
  ChevronLeft,
  Sparkles,
  SquareDashed,
  FileStack,
  ShieldCheck,
  Square,
  FileDown,
  Paperclip,
  Pause,
  Volume2,
  Combine,
  FileInput,
  Download,
  Search,
} from "@mrnote/core/web/components/office-ribbon/icons/houseIcons.tsx";
import {
  DropdownMenu,
  DropdownMenuItem,
  DropdownMenuLabel,
  DropdownMenuTrigger,
} from "@mrnote/core/web/components/ui/dropdown-menu.tsx";
import { RibbonDropdownMenuContent, RibbonMenuRichRow } from "@mrnote/core/web/components/office-ribbon/CupertinoRibbonMenu.tsx";
import {
  PDF_READER_PRINT_COMMAND,
  PDF_READER_DOWNLOAD_CURRENT_COMMAND,
  PDF_READER_DOWNLOAD_ORIGINAL_COMMAND,
  PDF_READER_APPLY_REDACTIONS_COMMAND,
  PDF_READER_DOWNLOAD_ARCHIVAL_COMMAND,
  PDF_READER_DOWNLOAD_ANNOTATED_COMMAND,
  PDF_READER_DOWNLOAD_FILLABLE_COMMAND,
  PDF_READER_DOWNLOAD_EDITABLE_ANNOTATIONS_COMMAND,
  PDF_READER_MERGE_COMMAND,
} from "./readerCommands";
import {
  PdfFeaturedRibbonTab,
  type PdfFeaturedCommand,
} from "../PdfFeaturedRibbonTab";
import {
  type PdfCommentRibbonCommand,
} from "../PdfCommentRibbonTab";
import {
  type PdfInsertRibbonCommand,
  type PdfPageRibbonCommand,
} from "../PdfInsertRibbonTab";
import { type PdfReadMode, isReadModeActive } from "../../../../../lib/pdf/invert-filter";
import { resolvePdfCommandDisabledReason } from "./commandDisabledReasons";
import { PDF_BOOKMARK_MAX, type UserBookmark } from "../../../../../lib/pdf/user-bookmarks";
import { type PdfNativeListTextControlId } from "../../../../../lib/pdf/engine/native/list-text-command-lifecycle";
import {
  kindSupportsFill,
  isShapeColor,
  PDF_SHAPE_COLORS,
  PDF_SHAPE_MIN_WIDTH,
  PDF_SHAPE_MAX_WIDTH,
  shapeWidthFromPoints,
  shapeWidthToPoints,
  type ShapeKind,
  type ShapeGeometryTransform,
} from "@mrnote/core/web/lib/pdf/shapes.ts";
import {
  PDF_INK_MAX_WIDTH,
  PDF_INK_MIN_WIDTH,
  isInkColor,
  PDF_INK_COLORS,
  DEFAULT_INK_WIDTH,
  type InkAnnotation,
} from "@mrnote/core/web/lib/pdf/ink.ts";
import {
  MEASURE_MODES,
  asMeasureMode,
  MEASURE_UNITS,
  asMeasureUnit,
  type MeasureMode,
  type MeasureUnit,
  type MeasurePoint,
} from "../../../../../lib/pdf/measure";
import {
  PDF_HIGHLIGHT_COLOR_OPTIONS,
  PDF_AREA_HIGHLIGHT_OPACITIES,
} from "./highlightPayload";
import { PDF_NOTE_COLORS } from "@mrnote/core/web/lib/pdf/sticky-notes.ts";
import { isSafeHexColor } from "@mrnote/core/web/lib/pdf/safe-color.ts";
import { isExactZoomPreset } from "../../../../../lib/pdf/view-modes";
import {
  PDF_COMPRESS_LOSSLESS_PRESET,
  type PdfCompressServerPreset,
  type PdfCompressionEvidenceSource,
} from "../../../../../lib/pdf/compress";
import {
  PDF_ERASE_BRUSH_WIDTHS,
  PDF_ERASE_DEFAULT_BRUSH_WIDTH,
  type PdfEraseBrushWidth,
  type PdfEraseShape,
} from "../../../../../lib/pdf/page-erase";
import { Loader2 } from "lucide-react";
import {
  PdfFormFieldContextRibbon,
  type PdfFormFieldContextKind,
  type PdfFormFieldContextRibbonLabels,
} from "../PdfFormFieldContextRibbon";
import {
  PdfLinkContextRibbon,
  type PdfLinkContextRibbonLabels,
} from "../PdfLinkContextRibbon";
import {
  PdfAnnotationFormatContextRibbon,
  type PdfAnnotationContextKind,
  type PdfAnnotationFormatContextRibbonModel,
  type PdfShapeLineCap,
  type PdfShapeLineJoin,
  type PdfAnnotationFormatContextRibbonLabels,
  type PdfAnnotationStyleChange,
} from "../PdfAnnotationFormatContextRibbon";
import {
  PdfEditRibbonSplitCommand,
  type PdfEditRibbonCommand,
} from "../PdfEditRibbonTab";
import { PdfPageLabelsControl } from "./PdfPageLabelsControl";
import { PdfLegacyDocumentTools } from "./PdfLegacyDocumentTools";
import { PdfExportImageControl } from "./PdfExportImageControl";
import { PdfLayersControl } from "./PdfLayersControl";
import { PdfSplitControl } from "./PdfSplitControl";
import { PdfLegacyDownloadMenu } from "./PdfLegacyDownloadMenu";
import type { PdfPageOrganizerOperation } from "../../../../../lib/pdf/page-organizer";
import type {
  ManagedNativeSignatureAnnotation,
} from "../../../../../lib/pdf/native-signature-roundtrip";
import type { PdfManagedStampCommandOrigin } from "./readerStamp";
import type { PdfLinkOverlay } from "@mrnote/core/web/lib/pdf/links.ts";
import type { PageLabelRange } from "@mrnote/core/web/lib/pdf/page-labels.ts";
import type {
  PdfNativeListTextSelectionEvidence,
  PdfNativeListTextSelectionProof,
} from "../../../../../lib/pdf/engine/native/list-text-fidelity";
import type {
  PdfNativeTextEditActivationOutcome,
} from "./nativeTextActivation";
import type { PdfSourceRiskSnapshot } from "../../../../../lib/pdf/save-policy";
import type { PageRecord } from "./pageRecord";
import type { PdfRibbonTab } from "./ribbonTabs";
import type { HighlightMarkupStyle } from "@mrnote/core/web/lib/pdf/highlight-markup.ts";
import type { OcgLayer, OcgOrderEntry } from "../../../../../lib/pdf/optional-content";
import type { SignatureComposerMode } from "../SignatureComposer";
import type { RedactionAnnotation } from "@mrnote/core/web/lib/pdf/redaction.ts";
import type {
  RedactionHydrationStatus,
  VisualMaskHydrationStatus,
  VisualMaskOutputAuthoritySnapshot,
} from "./outputAuthority";
import type { PdfPlacedAnnotationSelection } from "./annotationModels";
import type { PdfViewerFeedbackTone } from "./viewerFeedback";
import type {
  PdfDecorationKind,
  PdfDocumentDecorations,
} from "../../../../../lib/pdf/document-decorations";
import type { OutlineNode } from "@mrnote/core/web/lib/pdf/outline.ts";
import type { PdfFormDataFormat } from "@/lib/pdf/form-export";
import type {
  ImageExportScope,
  ImageExportFormat,
  ImageExportRenderSize,
} from "../../../../../lib/pdf/canvas-export";
import type {
  PdfTranslationPhase,
  PdfTranslationTarget,
  PdfTranslationScope,
} from "../PdfTranslationPreflight";
import type { PdfPageImportSourceKind } from "./pageImport";
import type {
  PdfNativeVectorSelectionRoute,
} from "../../../../../lib/pdf/semantic-appearance-inspection";
import type { PdfOcrReviewAnnotation } from "../../../../../lib/pdf/ocr-acceptance";
import type { PdfRotation } from "@mrnote/core/web/lib/pdf/page-ops.ts";
import type {
  RenderedFormField,
  SelectedFormFieldContext,
  DetectedFieldSourceIdentity,
} from "./formFieldSessions";
import type { PdfCompressPreset } from "../PdfCompressSettingsDialog";
import type { CreateFormFieldCandidate } from "@/lib/pdf/form-field-create";
import type { PercentBox } from "../../../../../lib/pdf/annotation-transform";
import type {
  NativePdfFormWidgetAppearanceStyle,
} from "../../../../../lib/pdf/form-field-authoring";
import type {
  PdfCommandInvocation,
  ResolvedPdfFeatureCommand,
} from "../commandRegistry";
import type { PdfRuntimeFeatureDecision } from "../capabilityAdapter";
import type { SpeechStatus } from "../../../../../lib/pdf/speech";
import type { PdfEditableProjectionTarget } from "@mrnote/core/web/lib/page-export.ts";
import type { PdfNavigationPanel } from "../PdfNavigationRail";
import type {
  PdfNativeVectorLayerMove,
} from "../../../../../lib/pdf/native-vector-selection-edit";
import type { StampAnnotation } from "../../../../../lib/pdf/stamps-annotation";
import type {
  ManagedPdfProofingSync,
  ManagedPdfProofingKind,
} from "@mrnote/core/web/lib/pdf/native-text-markup-writeback.ts";
import type {
  PdfTextBoxCommandTargetLike,
} from "../../../../../lib/pdf/text-box-command-target";
import type {
  PdfNativeTextInputLayout,
} from "../../../../../lib/pdf/native-text-input-layout";
import type { PdfSignatureProofPresentation } from "../../../../../lib/pdf/signature-proof";
import type { PDFDocumentProxy } from "pdfjs-dist";
import type {
  FeatureCommandResult,
} from "@mrnote/core/web/components/office-ribbon/featureCommandRegistry.ts";
import type { useTranslations } from "@mrnote/core/host/next-intl.tsx";
import { PdfRibbonToolButton } from "../PdfRibbonButton";
import { PdfSurfaceRibbonGroup } from "../PdfSurfaceRibbon";
import surfaceStyles from "../PdfSurfaceRibbon.module.css";

export interface PdfViewerRibbonModel {
  /** The 图片编辑 tab's command run while a picture is selected in place. */
  pictureContextRibbon?: ReactNode;
  /** The 文字编辑 tab's command run while a text box added on the page is
   * selected (the host's text ribbon, as in the editing surface). */
  textBoxContextRibbon?: ReactNode;
  /** 表格布局: the selected table's commands (PdfTableContextRibbon). */
  tableContextRibbon?: ReactNode;
  /** 图表设计 / 格式: the selected chart's ribbon, for the active one of them. */
  chartContextRibbon?: ReactNode;
  /** The reader's right-edge dock, where a selected shape's Format Shape
   * pane opens (pdfTaskPaneDock.tsx). */
  formatPaneDock?: OfficeTaskPaneDockRegistry | null;
  filename: string | undefined;
  downloadUrl: string | undefined;
  ribbonQuickActions: ReactNode;
  ribbonActionRail: PdfOfficeActionRail | undefined;
  ribbonQuickActionTail: ReactNode;
  ribbonQuickOverflowOnlyTail: ReactNode;
  annotationContextWritable: boolean;
  onPageOrganizerApply: ((operations: PdfPageOrganizerOperation[]) => void | Promise<void>) | undefined;
  onDownload: (() => boolean | void | Promise<boolean | void>) | undefined;
  preferAnnotatedDownload: boolean | undefined;
  annotatedDownloadEnabled: boolean;
  pageId: string | undefined;
  onRequestPageSummary: ((payload: { pageNumber: number; text: string; }) => void) | undefined;
  onManagedSignatureUpsert: ((signature: ManagedNativeSignatureAnnotation, authorityGeneration: string) => void | Promise<void>) | undefined;
  onManagedStampUpsert: ((stamp: StampAnnotation, origin: PdfManagedStampCommandOrigin) => void | Promise<void>) | undefined;
  onManagedLinkUpsert: ((link: PdfLinkOverlay, authorityGeneration: string) => void | Promise<void>) | undefined;
  onManagedProofingUpsert: ((proofing: ManagedPdfProofingSync) => void | Promise<void>) | undefined;
  dataItemId: string | undefined;
  lossyCompressionSource: PdfCompressionEvidenceSource | null | undefined;
  lossyCompressionWritable: boolean;
  nativePageLabelRanges: readonly PageLabelRange[] | null | undefined;
  onNativePageLabelRangesApply: ((next: PageLabelRange[]) => void | Promise<void>) | undefined;
  onNativeFileAttachmentUpsert: ((input: { file: File; page: number; x: number; y: number; }) => Promise<void>) | undefined;
  nativeFileAttachmentBusy: boolean;
  nativeFileAttachmentDisabledReason: string | undefined;
  onNativePageImport: ((input: { file: File; sourcePageIndices: number[]; targetInsertionIndex: number; }) => Promise<void>) | undefined;
  nativePageImportBusy: boolean;
  nativePageImportDisabledReason: string | undefined;
  onEditContent: ((trigger: HTMLElement | null, controlId: PdfNativeListTextControlId, selection?: PdfNativeListTextSelectionEvidence | undefined, activation?: Readonly<{ caretOffset: number | null; caretPoint?: Readonly<{ x: number; y: number; }> | undefined; copySourceBytes?: (() => Uint8Array<ArrayBufferLike> | null) | undefined; editingSurface?: Readonly<{ key: string; isCurrent: () => boolean; }> | undefined; renderNativeEditingSurface?: ((proof: PdfNativeListTextSelectionProof, signal: AbortSignal) => Promise<HTMLCanvasElement | null>) | undefined; renderProvisionalEditingSurface?: ((signal: AbortSignal) => Promise<HTMLCanvasElement | null>) | undefined; renderManagedEditingSurface?: ((annotationId: string, signal: AbortSignal) => Promise<HTMLCanvasElement | null>) | undefined; layout?: Readonly<{ box: PdfTextBoxCommandTargetLike; pageWidthPx: number; pageHeightPx: number; pageHeightPoints?: number | undefined; pdfToCssTransform?: readonly number[] | undefined; sourceInputLayout?: PdfNativeTextInputLayout | null | undefined; sourceInputKey?: string | undefined; }> | null | undefined; }> | undefined) => boolean | void | PdfNativeTextEditActivationOutcome | Promise<boolean | void | PdfNativeTextEditActivationOutcome>) | undefined;
  editContentDisabledReason: string | undefined;
  editContentBusy: boolean;
  /** Edit Content mode (WPS 编辑内容 toggle) is on. */
  contentEditModeActive: boolean;
  contentEditModeDisabledReason: string | undefined;
  onToggleContentEditMode: () => void;
  onEnterContentEditMode: () => void;
  onExitContentEditMode: () => void;
  onInsertText: ((origin: "insert" | "edit" | "annotate") => void) | undefined;
  onCropPage: ((trigger: HTMLElement | null, resolveTrigger?: (() => HTMLButtonElement | null) | undefined) => void) | undefined;
  eraseRibbon: {
    armed: boolean;
    color: string | null;
    shape: PdfEraseShape;
    brushWidth: PdfEraseBrushWidth;
    disabledReason: string | undefined;
    onToggle: () => void;
    onColorChange: (color: string | null) => void;
    onShapeChange: (shape: PdfEraseShape) => void;
    onBrushWidthChange: (width: PdfEraseBrushWidth) => void;
  } | null;
  /** Enter (or leave) WPS 分割页面 mode; undefined when the host cannot split. */
  onSplitPage: (() => void) | undefined;
  splitPageActive: boolean;
  splitPageDisabledReason: string | undefined;
  onOpenPageBackground: ((trigger: HTMLButtonElement | null) => void) | undefined;
  pageBackgroundOpen: boolean;
  pageBackgroundBusy: boolean;
  pageBackgroundDisabledReason: string | undefined;
  onCreateEditableVersion: (() => void) | undefined;
  onConvertToSpreadsheet: (() => void) | undefined;
  onConvertToPresentation: (() => void) | undefined;
  /** The convert menu opened (the host may let the server start reading). */
  onOfficeConversionPrepare: (() => void) | undefined;
  officeProjectionDisabledReason: string | undefined;
  isCreatingEditableVersion: boolean | undefined;
  t: ReturnType<typeof useTranslations<"console-notebooks">>;
  pdfHomeDisplayLabels: { hand: string; convert: string; convertGroup: string; outputImages: string; outputGroup: string; splitMerge: string; splitMergeGroup: string; play: string; rotateDocument: string; editContent: string; editGroup: string; fullTranslation: string; selectionTranslation: string; screenshot: string; compress: string; documentEncrypt: string; };
  savePolicyV1: boolean;
  signatureProofLoading: boolean;
  matchedSourceRiskSnapshot: PdfSourceRiskSnapshot | null;
  signatureProofPresentation: PdfSignatureProofPresentation;
  showSignatureProof: boolean;
  pageOrganizerLabel: string;
  pages: PageRecord[];
  loading: boolean;
  setPageOrganizerOpen: Dispatch<SetStateAction<boolean>>;
  pageOrganizerDialogReturnFocusRef: RefObject<HTMLButtonElement | null>;
  scale: number;
  activeRibbonTab: PdfRibbonTab;
  ribbonInstanceId: string;
  currentPage: number;
  currentVisibleOrdinal: number;
  pageInput: string;
  setPageInput: Dispatch<SetStateAction<string>>;
  pendingRibbonMarkupStyle: HighlightMarkupStyle | null;
  setPendingRibbonMarkupStyle: Dispatch<SetStateAction<HighlightMarkupStyle | null>>;
  pendingProofingKind: ManagedPdfProofingKind | null;
  proofingBusy: boolean;
  proofingError: string | null;
  ribbonMarkupError: string | null;
  ribbonMarkupColor: string;
  noteMode: boolean;
  /** The Comment tab ink tool's Eraser (WPS's graffiti tab). */
  commentInkEraser: { active: boolean; available: boolean; toggle: () => void };
  /** The Comment tab ink tool's 直线 (WPS's graffiti line type). */
  commentInkStraight: { on: boolean; set: (on: boolean) => void };
  /** The measure tab's snapping (WPS PDFAdsorption*). */
  measureSnap: {
    kinds: ReadonlySet<"endpoint" | "path" | "midpoint" | "intersection">;
    toggle: (kind: "endpoint" | "path" | "midpoint" | "intersection") => void;
  };
  /** The measure tab's 结果面板 (WPS Measure_ResultPanel). */
  measureResults: {
    items: ReadonlyArray<{ id: string; mode: string; label: string; page: number }>;
    open: (page: number) => void;
  };
  /** The command that armed the text box, while it is armed. */
  armedTextOrigin: "insert" | "edit" | "annotate" | null;
  /** What the armed text tool places: a text box, or one set as a formula. */
  armedTextKind: PdfArmedTextKind;
  noteColor: string;
  commentModeActive: boolean;
  commentsHidden: boolean;
  attachMode: boolean;
  attachmentError: string | null;
  placedAttachments: { id: string; page: number; x: number; y: number; name: string; mimeType: string; size: number; sha256: string; attachmentDataItemId: string; icon: "Paperclip" | "PushPin" | "Graph" | "Tag"; }[];
  ocgLayers: OcgLayer[];
  /** The author's layer list (/Order), when the document has one. */
  ocgOrder: OcgOrderEntry[] | null;
  setLayersPanelOpen: Dispatch<SetStateAction<boolean>>;
  layersPanelOpen: boolean;
  handleToggleLayer: (id: string, visible: boolean) => void;
  inks: InkAnnotation[];
  drawMode: boolean;
  /** Draw armed from the Comment tab's freehand tool, not from the Draw tab: only then do the freehand buttons light. */
  commentDrawMode?: boolean;
  setInkColor: Dispatch<SetStateAction<string>>;
  inkColor: string;
  inkWidth: number;
  setInkWidth: Dispatch<SetStateAction<number>>;
  inkSaving: boolean;
  signMode: boolean;
  signSaving: boolean;
  signatureComposerMode: SignatureComposerMode;
  stampPickerOpen: boolean;
  armedStamp: { kind: "preset"; preset: string; origin: PdfManagedStampCommandOrigin; } | { kind: "image"; image: string; aspectRatio: number; origin: PdfManagedStampCommandOrigin; } | null;
  redactions: RedactionAnnotation[];
  redactMode: boolean;
  redactSaving: boolean;
  redactionHydrationStatusForCurrentDocument: RedactionHydrationStatus;
  redactionProductCopy: { pending: string; reviewSummary: string; reviewMark: string; reviewPage: (page: number) => string; reviewMapped: (outputCount: number) => string; reviewCopyOnly: (outputCount: number) => string; reviewUnmapped: string; reviewRemove: string; removed: string; undoRemove: string; reviewBlocked: string; apply: string; applying: string; applyTitle: string; applyDescription: string; applyConfirm: string; blockedOutput: string; hydrationPending: string; hydrationFailed: string; success: string; failure: string; };
  shapesHydrationStatusForCurrentDocument: VisualMaskHydrationStatus;
  shapeMode: boolean;
  setShapeMode: Dispatch<SetStateAction<boolean>>;
  shapeToolPageContentRef: RefObject<boolean>;
  shapeKind: ShapeKind;
  setShapeKind: Dispatch<SetStateAction<ShapeKind>>;
  setShapeColor: Dispatch<SetStateAction<string>>;
  shapeColor: string;
  shapeWidth: number;
  setShapeWidth: Dispatch<SetStateAction<number>>;
  shapeLineStyle: "solid" | "dashed" | "dotted";
  setShapeLineStyle: Dispatch<SetStateAction<"solid" | "dashed" | "dotted">>;
  setShapeLineStarts: Dispatch<SetStateAction<Record<"line" | "arrow" | "polyline", "none" | "openArrow" | "closedArrow" | "circle" | "square">>>;
  setShapeLineEnds: Dispatch<SetStateAction<Record<"line" | "arrow" | "polyline", "none" | "openArrow" | "closedArrow" | "circle" | "square">>>;
  activeShapeLineKind: "line" | "arrow" | "polyline" | null;
  shapeLineStart: "none" | "openArrow" | "closedArrow" | "circle" | "square";
  shapeLineEnd: "none" | "openArrow" | "closedArrow" | "circle" | "square";
  shapeFillColor: string | null;
  setShapeFillColor: Dispatch<SetStateAction<string | null>>;
  areaHighlightActive: boolean;
  setAreaHighlightActive: Dispatch<SetStateAction<boolean>>;
  areaHighlightColor: string;
  areaHighlightOpacity: number;
  visualMaskActive: boolean;
  setVisualMaskActive: Dispatch<SetStateAction<boolean>>;
  visualMaskWarning: string;
  shapeSaving: boolean;
  selectedPlaced: PdfPlacedAnnotationSelection | null;
  selectedPlacedRef: RefObject<PdfPlacedAnnotationSelection | null>;
  placedAnnotationClipboardError: boolean;
  linkMode: boolean;
  linkSaving: boolean;
  linkPlacementDestinationKindRef: RefObject<"url" | "page">;
  snapshotMode: boolean;
  setSnapshotMode: Dispatch<SetStateAction<boolean>>;
  snapshotIntentRef: RefObject<"copy" | "latex" | "text">;
  snapshotShape: SnapshotShape;
  setSnapshotShape: (shape: SnapshotShape) => void;
  captureScreenSnapshot: () => Promise<void>;
  setDrawingSnapshot: Dispatch<SetStateAction<{ page: number; start: { x: number; y: number; }; current: { x: number; y: number; }; } | null>>;
  snapshotDragRef: RefObject<{ page: number; pointerId: number; outputAuthorityIsCurrent: () => boolean; } | null>;
  snapshotTextReturnFocusRef: RefObject<HTMLButtonElement | null>;
  setSnapshotToast: (message: string | null, tone?: PdfViewerFeedbackTone) => void;
  setSnapshotLatex: Dispatch<SetStateAction<{ page: number; start: { x: number; y: number; }; current: { x: number; y: number; }; mode: "latex" | "text" | "offer"; imageB64: string | null; outputAuthority: VisualMaskOutputAuthoritySnapshot; textAuthority?: Readonly<{ dataItemId: string; sourceSha256: string; pageIndex: number; region: readonly [number, number, number, number]; authorityIsCurrent: () => boolean; }> | undefined; } | null>>;
  measureMode: boolean;
  measureKind: MeasureMode;
  setMeasureKind: Dispatch<SetStateAction<MeasureMode>>;
  measureUnit: MeasureUnit;
  setMeasureUnit: Dispatch<SetStateAction<MeasureUnit>>;
  measureRatio: number;
  setMeasureRatio: Dispatch<SetStateAction<number>>;
  setMeasurePts: Dispatch<SetStateAction<{ page: number; points: MeasurePoint[]; } | null>>;
  pageLabelsOpen: boolean;
  pageLabelsDialogReturnFocusRef: RefObject<HTMLButtonElement | null>;
  pageLabelRanges: PageLabelRange[];
  pageLabelReplacementStarted: boolean;
  setPageLabelReplacementStarted: Dispatch<SetStateAction<boolean>>;
  pageLabelsApplying: boolean;
  pageLabelsError: string | null;
  nativePageLabelsControlled: boolean;
  nativePageLabelsSourceOwned: boolean;
  updatePageLabelRanges: (update: (current: PageLabelRange[]) => PageLabelRange[]) => void;
  printing: boolean;
  printPreflightTriggerRef: RefObject<HTMLButtonElement | null>;
  batchPrintOpen: boolean;
  batchPrintPreparing: boolean;
  decorationDialogKind: PdfDecorationKind | null;
  effectiveDocumentDecorations: PdfDocumentDecorations;
  openDecorationDialog: (kind: PdfDecorationKind, trigger: HTMLElement | null) => void;
  removeDocumentDecoration: (kind: PdfDecorationKind) => Promise<void>;
  annotationsOpen: boolean;
  outline: OutlineNode[];
  setPropertiesOpen: Dispatch<SetStateAction<boolean>>;
  propertiesDialogReturnFocusRef: RefObject<HTMLButtonElement | null>;
  userBookmarks: UserBookmark[];
  bookmarkSaving: boolean;
  flattening: boolean;
  formDataExportFormat: PdfFormDataFormat;
  setFormDataExportFormat: Dispatch<SetStateAction<PdfFormDataFormat>>;
  compressPreset: "lossless-structural-v1" | PdfCompressServerPreset;
  setCompressPreset: Dispatch<SetStateAction<"lossless-structural-v1" | PdfCompressServerPreset>>;
  openCompressSettings: (trigger: HTMLButtonElement | null) => void;
  mergeBusy: boolean;
  splitBusy: boolean;
  featuredImagesToPdfBusy: boolean;
  featuredMarkdownToPdfBusy: boolean;
  featuredCadToPdfBusy: boolean;
  /** Markdown is converted by the server for a saved PDF page only. */
  featuredMarkdownToPdfAvailable: boolean;
  setSplitOpen: Dispatch<SetStateAction<boolean>>;
  splitOpen: boolean;
  splitDialogReturnFocusRef: RefObject<HTMLButtonElement | null>;
  featuredLegacyToolsOpen: boolean;
  setFeaturedLegacyToolsOpen: Dispatch<SetStateAction<boolean>>;
  setExportImageOpen: Dispatch<SetStateAction<boolean>>;
  exportImageOpen: boolean;
  exportImageDialogReturnFocusRef: RefObject<HTMLButtonElement | null>;
  exportImageScope: ImageExportScope;
  setExportImageScope: Dispatch<SetStateAction<ImageExportScope>>;
  exportImageRange: string;
  setExportImageRange: Dispatch<SetStateAction<string>>;
  exportImageCombinePages: boolean;
  setExportImageCombinePages: Dispatch<SetStateAction<boolean>>;
  exportImageResolution: 72 | 96 | 150 | 300 | 600;
  setExportImageResolution: Dispatch<SetStateAction<72 | 96 | 150 | 300 | 600>>;
  exportImageBusy: boolean;
  textExportBusy: boolean;
  imageExtractionBusy: boolean;
  translationDialog: { phase: PdfTranslationPhase; target: PdfTranslationTarget; scope: PdfTranslationScope; selectionText: string | null; selectionAuthorityIsCurrent: (() => boolean) | null; translatedText: string; error: string | null; } | null;
  translationWorking: boolean;
  cadExportBusy: boolean;
  screenRecordingState: "error" | "idle" | "requesting" | "recording" | "saving" | "saved";
  screenRecordingSupported: boolean;
  documentEncryptBusy: boolean;
  requestSignatureProofPanel: () => void;
  dangerousActionSafeCopyBusy: boolean;
  verifiedPdfaExportBusy: boolean;
  verifiedPdfaExportTriggerRef: RefObject<HTMLButtonElement | null>;
  imageOnlyExportBusy: boolean;
  xfdfImportNotice: string | null;
  rasterOutputAuthorityBlocked: boolean;
  splitMode: "ranges" | "combine" | "everyN" | "bookmarks";
  setSplitMode: Dispatch<SetStateAction<"ranges" | "combine" | "everyN" | "bookmarks">>;
  splitRangeSpec: string;
  /** Why the typed split range cannot run (a page the file lacks). */
  splitRangeError: string | null;
  setSplitRangeSpec: Dispatch<SetStateAction<string>>;
  splitEveryNValue: string;
  setSplitEveryNValue: Dispatch<SetStateAction<string>>;
  featuredImagesToPdfInputRef: RefObject<HTMLInputElement | null>;
  featuredMarkdownToPdfInputRef: RefObject<HTMLInputElement | null>;
  featuredCadToPdfInputRef: RefObject<HTMLInputElement | null>;
  featuredImagesToPdfReturnFocusRef: RefObject<HTMLButtonElement | null>;
  setInsertOpen: Dispatch<SetStateAction<boolean>>;
  insertOpen: boolean;
  selectInsertSourceKind: (kind: PdfPageImportSourceKind) => void;
  insertDialogReturnFocusRef: RefObject<HTMLButtonElement | null>;
  insertBusy: boolean;
  editCropTriggerRef: RefObject<HTMLButtonElement | null>;
  layersDialogReturnFocusRef: RefObject<HTMLButtonElement | null>;
  formImportInputRef: RefObject<HTMLInputElement | null>;
  formImportReturnFocusRef: RefObject<HTMLButtonElement | null>;
  viewMode: "continuous" | "single" | "spread";
  setViewMode: Dispatch<SetStateAction<"continuous" | "single" | "spread">>;
  readMode: PdfReadMode;
  cycleReadMode: () => void;
  lineNumbersVisible: boolean;
  toggleLineNumbers: () => void;
  fitMode: "page" | "width" | null;
  panMode: boolean;
  presentationActive: boolean;
  compareInputRef: RefObject<HTMLInputElement | null>;
  compareFile: File | null;
  viewerRootRef: RefObject<HTMLDivElement | null>;
  setSelectedNativeVectorSelection: Dispatch<SetStateAction<PdfNativeVectorSelectionRoute | null>>;
  /** Bumped when the user leaves a source vector (see PDFViewer). */
  nativeVectorUserSelectionEpochRef: RefObject<number>;
  setExtraNativeVectorSelections: Dispatch<SetStateAction<readonly PdfNativeVectorSelectionRoute[]>>;
  ocrReviewByPage: ReadonlyMap<number, PdfOcrReviewAnnotation>;
  ocrAnnotationHydrationReady: boolean;
  ocrCandidatePages: ReadonlySet<number>;
  ocrPendingPages: ReadonlySet<number>;
  ocrPreflightLoadingPage: number | null;
  ocrPreflightTriggerRef: RefObject<HTMLElement | null>;
  pdfOcrServerSourceAuthority: Readonly<{ document: PDFDocumentProxy; dataItemId: string; url: string; byteGenerationKey: string; documentRevision: string | number; sourceSha256: string; derived?: true | undefined; }> | null;
  pageOrganizerPages: { page: number; width: number; height: number; rotation: PdfRotation; }[];
  visiblePageCount: number;
  previousVisibleOrdinal: number | null;
  nextVisibleOrdinal: number | null;
  rotateDocumentDisabled: boolean;
  goToVisibleOrdinal: (ordinal: number, options?: { scrollTop?: number | undefined; recordHistory?: boolean | undefined; restoringView?: boolean | undefined; } | undefined) => void;
  handleSummarizeCurrentPage: () => Promise<void>;
  pageOpsEnabled: boolean;
  handleRotatePage: (page: number) => void;
  handleRotateCcwPage: (page: number) => void;
  handleRotateDocumentClockwise: () => void;
  handlePageInputCommit: () => void;
  selectionEnabled: boolean;
  clearPlacedAnnotationSelection: () => void;
  requestOcrPreflight: (pageNumber: number) => Promise<void>;
  openOcrReview: (pageNumber: number) => void;
  handleInkUndoLast: () => void;
  pendingRedactionReviewItems: { id: string; page: number; outputPageCount: number; canNavigate: boolean; order: number; }[];
  handleRedactionReviewNavigate: (annotationId: string) => void;
  handleRedactionReviewDelete: (annotationId: string) => void;
  handleSelectedManagedShapeStyle: (change: { outlineColor?: string | undefined; lineWidth?: number | undefined; lineStyle?: "solid" | "dashed" | "dotted" | undefined; lineStart?: "none" | "openArrow" | "closedArrow" | "circle" | "square" | undefined; lineEnd?: "none" | "openArrow" | "closedArrow" | "circle" | "square" | undefined; fillColor?: string | null | undefined; fillOpacity?: number | undefined; opacity?: number | undefined; cornerRadius?: number | undefined; }) => Promise<void>;
  handleSelectedManagedShapeArrange: (action: ShapeGeometryTransform) => Promise<void>;
  handleSelectedManagedInkStyle: (change: { outlineColor?: string | undefined; lineWidth?: number | undefined; }) => Promise<void>;
  handleSelectedShapeSizeChange: (axis: "width" | "height", valuePoints: number, lockAspectRatio?: boolean) => void;
  handleCopyPlacedAnnotation: () => boolean;
  handlePastePlacedAnnotation: () => Promise<boolean>;
  handleDuplicatePlacedAnnotation: () => Promise<boolean>;
  handleDeleteSelectedPlacedAnnotation: () => boolean;
  handleInsertBookmark: (trigger: HTMLButtonElement | null) => Promise<boolean>;
  clearSelectedLinkContext: (restoreFocus?: boolean) => void;
  openSelectedLinkContext: () => void;
  editSelectedLinkContext: () => void;
  deleteSelectedLinkContext: () => Promise<void>;
  allFormFields: RenderedFormField[];
  formFieldsUnavailable: boolean;
  formFieldsUnavailableReason: string;
  nativeFormMode: boolean;
  formMutationBlocked: boolean;
  formMutationBlockedReason: string;
  nativeFormDataMutationBlocked: boolean;
  nativeFormDataMutationBlockedReason: string;
  hasNativeFormEditsExcludedFromSource: boolean;
  hasFlattenableAnnotations: boolean;
  hasEffectivePdfOutputEdits: boolean;
  hasEditableXfdfAnnotations: boolean;
  handleCompressPdf: (requestedPreset?: PdfCompressPreset) => Promise<void>;
  handleExportFormData: (requestedFormat?: PdfFormDataFormat) => Promise<void>;
  handleResetForm: () => Promise<void>;
  fieldCandidates: CreateFormFieldCandidate[] | null;
  detectBusy: boolean;
  detectedFieldsCommandRef: RefObject<HTMLButtonElement | null>;
  existingTextFieldEditMode: boolean;
  selectedFormFieldContextRef: RefObject<SelectedFormFieldContext | null>;
  existingTextFieldEditBusy: boolean;
  existingFormFieldDeleteMode: boolean;
  existingFormFieldDeleteBusy: boolean;
  currentDetectedFieldSourceIdentityRef: RefObject<DetectedFieldSourceIdentity | null>;
  handleDetectFormFields: () => Promise<void>;
  acceptedCandidates: CreateFormFieldCandidate[];
  clearSelectedFormFieldContext: () => void;
  openSelectedFormFieldProperties: () => void;
  openSelectedFormFieldOptions: () => void;
  hasVisibleFormFieldForPropertyEdit: boolean;
  hasVisibleFormFieldForDelete: boolean;
  handleCancelDetectedFields: () => void;
  handleSelectedFormFieldTransform: (box: PercentBox) => Promise<void>;
  applySelectedFormWidgetAppearance: (change: Partial<NativePdfFormWidgetAppearanceStyle>) => Promise<void>;
  applySelectedFormFieldQuickProperty: (property: "required" | "readOnly", value: boolean) => Promise<void>;
  handleCreateDetectedFields: () => Promise<void>;
  handleExportPagesAsImages: (format: ImageExportFormat) => Promise<void>;
  exportImageOverLimit: { count: number; batches: string[]; } | null;
  exportImageCurrentPageSize: ImageExportRenderSize | null;
  openBatchPrint: (trigger: HTMLButtonElement | null) => Promise<void>;
  handleOpenImageExtractionReview: (trigger: HTMLButtonElement | null) => Promise<void>;
  handleSplitDownload: (override?: { mode: "ranges" | "combine" | "everyN" | "bookmarks"; rangeSpec: string; everyNValue?: string | undefined; } | undefined) => Promise<boolean>;
  handleOpenMergePicker: () => boolean;
  executeReaderCommand: (invocation: PdfCommandInvocation, decision: PdfRuntimeFeatureDecision) => Promise<FeatureCommandResult>;
  executeEditContentActivation: (trigger: HTMLElement | null, controlId: PdfNativeListTextControlId, selection?: PdfNativeListTextSelectionEvidence | undefined, activation?: Readonly<{ caretOffset: number | null; caretPoint?: Readonly<{ x: number; y: number; }> | undefined; copySourceBytes?: (() => Uint8Array<ArrayBufferLike> | null) | undefined; editingSurface?: Readonly<{ key: string; isCurrent: () => boolean; }> | undefined; renderNativeEditingSurface?: ((proof: PdfNativeListTextSelectionProof, signal: AbortSignal) => Promise<HTMLCanvasElement | null>) | undefined; renderProvisionalEditingSurface?: ((signal: AbortSignal) => Promise<HTMLCanvasElement | null>) | undefined; renderManagedEditingSurface?: ((annotationId: string, signal: AbortSignal) => Promise<HTMLCanvasElement | null>) | undefined; layout?: Readonly<{ box: PdfTextBoxCommandTargetLike; pageWidthPx: number; pageHeightPx: number; pageHeightPoints?: number | undefined; pdfToCssTransform?: readonly number[] | undefined; sourceInputLayout?: PdfNativeTextInputLayout | null | undefined; sourceInputKey?: string | undefined; }> | null | undefined; }> | undefined) => Promise<{ command: FeatureCommandResult; outcome: PdfNativeTextEditActivationOutcome | null; }>;
  originalDownloadRuntimeDecision: PdfRuntimeFeatureDecision;
  currentDownloadRuntimeDecision: PdfRuntimeFeatureDecision;
  annotatedDownloadRuntimeDecision: PdfRuntimeFeatureDecision;
  fillableDownloadRuntimeDecision: PdfRuntimeFeatureDecision;
  editableAnnotationsDownloadRuntimeDecision: PdfRuntimeFeatureDecision;
  archivalDownloadRuntimeDecision: PdfRuntimeFeatureDecision;
  printRuntimeDecision: PdfRuntimeFeatureDecision;
  applyRedactionsRuntimeDecision: PdfRuntimeFeatureDecision;
  originalDownloadCommand: ResolvedPdfFeatureCommand;
  handleSignatureProofOriginalDownload: () => void;
  currentDownloadCommand: ResolvedPdfFeatureCommand;
  annotatedDownloadCommand: ResolvedPdfFeatureCommand;
  fillableDownloadCommand: ResolvedPdfFeatureCommand;
  editableAnnotationsDownloadCommand: ResolvedPdfFeatureCommand;
  archivalDownloadCommand: ResolvedPdfFeatureCommand;
  printCommand: ResolvedPdfFeatureCommand;
  applyRedactionsCommand: ResolvedPdfFeatureCommand;
  activeSplitCommand: { readonly featureId: "PDF-PAG-STRUCT-03"; readonly action: "pdf.page.extract"; } | { readonly featureId: "PDF-TOO-FILE-01"; readonly action: "pdf.file.split"; };
  zoomReset: () => void;
  fitWidth: () => void;
  fitPage: () => void;
  disarmPageTools: (except?: "draw" | "shape" | "note" | "stamp" | "form" | "link" | "sign" | "redact" | "attach" | "snapshot" | "measure" | "proofing" | "pan" | undefined) => void;
  toggleRedactionMarkMode: () => void;
  openRedactPages: () => void;
  openHiddenInformation?: () => void;
  handleToggleCommentMode: () => void;
  handleToggleCommentsHidden: () => void;
  handleRibbonMarkupColorChange: (color: string) => void;
  handleNoteColorChange: (color: string) => void;
  handleAreaHighlightColorChange: (color: string) => void;
  handleAreaHighlightOpacityChange: (opacity: number) => void;
  closePageLabelsDialog: () => void;
  togglePageLabelsDialog: (trigger?: HTMLButtonElement | null) => void;
  applyPageLabelRanges: () => Promise<void>;
  toggleVisualMaskTool: () => void;
  toggleVisualSignature: (trigger?: HTMLButtonElement | null) => void;
  armLinkMode: (destinationKind: "url" | "page", trigger?: HTMLButtonElement | null) => void;
  deleteAllLinks: () => Promise<void>;
  documentLinkCount: number;
  deleteAllLinksAvailable: boolean;
  toggleLinkMode: (trigger?: HTMLButtonElement | null) => void;
  selectRibbonTab: (nextTab: PdfRibbonTab) => void;
  handleHandToolToggle: () => void;
  handleSelectTool: () => void;
  handleToggleExistingTextFieldEditMode: (trigger?: HTMLButtonElement | null | undefined) => void;
  handleToggleExistingFormFieldDeleteMode: (trigger?: HTMLButtonElement | null | undefined) => void;
  selectToolActive: boolean;
  executeMarkupActivation: (markupStyle: HighlightMarkupStyle) => Promise<FeatureCommandResult>;
  executeProofingActivation: (kind: ManagedPdfProofingKind) => Promise<FeatureCommandResult>;
  executeFreeTextActivation: (
    origin: "insert" | "edit" | "annotate",
    kind?: PdfArmedTextKind,
  ) => Promise<FeatureCommandResult>;
   
  readerTextBoxes: boolean;
  /** 表格: tables can be placed on the page (the notebook's native writer). */
  readerTables: boolean;
  /** 图表: charts can be placed on the page. */
  readerCharts: boolean;
  executeImageActivation: (trigger?: HTMLButtonElement | null) => Promise<FeatureCommandResult>;
  executeNoteActivation: () => Promise<FeatureCommandResult>;
  executeInkActivation: () => Promise<FeatureCommandResult>;
  executeShapeActivation: (origin?: "comment" | "page-content") => Promise<FeatureCommandResult>;
  executeAreaHighlightActivation: () => Promise<FeatureCommandResult>;
  executeStampActivation: (trigger?: HTMLButtonElement | null) => Promise<FeatureCommandResult>;
  executeLinkActivation: (destinationKind: "url" | "page", trigger?: HTMLButtonElement | null) => Promise<FeatureCommandResult>;
  executeVisualSignatureActivation: (mode: SignatureComposerMode, trigger?: HTMLButtonElement | null) => Promise<FeatureCommandResult>;
  executeFileAttachmentActivation: (trigger?: HTMLButtonElement | null) => Promise<FeatureCommandResult>;
  handlePresentationToggle: () => void;
  deskewStatus: string | null;
  speechStatus: SpeechStatus;
  speechSupported: boolean;
  handleToggleReadAloud: () => void;
  handleStopReadAloud: () => void;
  handleSaveSelectedAttachment: () => void;
  handleExportDocumentText: (trigger?: HTMLButtonElement | null) => Promise<void>;
  openFullTranslation: (trigger: HTMLButtonElement) => void;
  openSelectionTranslation: (trigger: HTMLButtonElement | null) => void;
  openOfficeProjection: (target: PdfEditableProjectionTarget, trigger: HTMLButtonElement | null) => void;
  openScreenRecording: (trigger: HTMLButtonElement | null) => void;
  stopScreenRecording: () => void;
  openCadExport: (trigger: HTMLButtonElement | null) => void;
  dangerousActionSafeCopyDisabledReason: string;
  dangerousActionCount: number;
  dangerousActionCompactStatusLabel: string;
  dangerousActionSafeCopyCommand: ResolvedPdfFeatureCommand;
  executeDangerousActionSafeCopyReview: (trigger: HTMLButtonElement) => Promise<void>;
  openDocumentEncrypt: (trigger: HTMLButtonElement | null) => void;
  openImageOnlyExport: (trigger: HTMLButtonElement | null) => void;
  toggleNavigationPanel: (panel: PdfNavigationPanel) => void;
  docSearchOpen: boolean;
  effectiveRibbonSaveStatus: PdfRibbonSaveStatus | undefined;
  effectiveRibbonSaveStatusLabel: string | undefined;
  currentNativeVectorSelection: PdfNativeVectorSelectionRoute | null;
  maskCurrentNativeVector: () => Promise<void>;
  toggleNativeVectorVertexEditing: () => void;
  /** Home > Arrange > Selection Pane (Alt+F10): the page's objects, front first. */
  toggleNativeSelectionPane?: () => void;
  nativeSelectionPaneOpen?: boolean;
  /** Copy and paste of source vectors (the copy lands where the pointer
   * was over a page, as WPS pastes; from the ribbon it is nudged instead). */
  nativeVectorClipboardCommands: Readonly<{ copy: () => void; paste: () => void }>;
  runOrQueueNativeVectorCommand: (run: (handlers: { rotate: (degrees: number) => Promise<void>; duplicate: () => Promise<void>; arrange: (action: ShapeGeometryTransform) => Promise<void>; align: (command: PdfShapeAlignCommand) => Promise<void>; reorder: (direction: PdfNativeVectorLayerMove) => Promise<void>; resize: (axis: "width" | "height", valuePoints: number, lockAspectRatio?: boolean | undefined) => Promise<void>; remove: () => Promise<void>; merge: (operation: import("../../../../../lib/pdf/native-vector-selection-edit").PdfNativeVectorMergeOperation) => Promise<void>; style: (change: Readonly<{ outlineColor?: string | undefined; lineWidth?: number | undefined; lineStyle?: "solid" | "dashed" | "dotted" | undefined; lineStart?: "none" | "openArrow" | "closedArrow" | "circle" | "square" | undefined; lineEnd?: "none" | "openArrow" | "closedArrow" | "circle" | "square" | undefined; fillColor?: string | null | undefined; fillOpacity?: number | undefined; opacity?: number | undefined; blendMode?: "screen" | "normal" | "color" | "multiply" | "overlay" | "darken" | "lighten" | "color-dodge" | "color-burn" | "hard-light" | "soft-light" | "difference" | "exclusion" | "hue" | "saturation" | "luminosity" | undefined; cornerRadius?: number | undefined; lineCap?: PdfShapeLineCap | undefined; lineJoin?: PdfShapeLineJoin | undefined; }>) => void; }) => unknown) => void;
  annotationContextModel: { deleteDisabledReason?: string | undefined; deleteEnabled: boolean; duplicateDisabledReason?: string | undefined; duplicateEnabled: boolean; pasteDisabledReason?: string | undefined; pasteEnabled: boolean; copyDisabledReason?: string | undefined; copyEnabled: boolean; arrangeDisabledReason?: string | undefined; arrangeEnabled: boolean; styleDisabledReason?: string | undefined; styleEnabled: boolean; sizeDisabledReason?: string | undefined; busy: boolean; sizeEnabled: boolean; kind: PdfAnnotationContextKind; nativeSource: boolean; page: number; width: number; height: number; widthPoints?: number | undefined; heightPoints?: number | undefined; pointAnchor?: boolean | undefined; outlineColor?: string | undefined; lineWidth?: number | undefined; fillSupported?: boolean | undefined; fillColor?: string | null | undefined; fillOpacity?: number | undefined; attachment?: { name: string; mimeType: string; sizeBytes: number; saveEnabled: boolean; saveDisabledReason?: string | undefined; } | undefined; } | PdfAnnotationFormatContextRibbonModel | null;
  effectiveAnnotationContextLabels: PdfAnnotationFormatContextRibbonLabels;
  formFieldContextLabels: PdfFormFieldContextRibbonLabels;
  formFieldContextModel: { name: string; kind: PdfFormFieldContextKind; required: boolean; readOnly: boolean; width: number; height: number; busy: boolean; geometryEnabled: boolean; geometryDisabledReason: string | undefined; appearanceStyle: { backgroundColor: string; borderColor: string; borderWidth: number; } | undefined; appearanceStyleEnabled: boolean; appearanceStyleDisabledReason: string | undefined; quickPropertiesEnabled: boolean; quickPropertiesDisabledReason: string | undefined; propertiesEnabled: boolean; propertiesDisabledReason: string | undefined; optionsEnabled: boolean; optionsDisabledReason: string | undefined; actionsEnabled: boolean; actionsDisabledReason: string; } | null;
  linkContextLabels: PdfLinkContextRibbonLabels;
  linkContextModel: { deleteDisabledReason?: string | undefined; deleteEnabled: boolean; editDisabledReason?: string | undefined; kind: "url" | "page"; source: "native" | "managed"; page: number; destination: string; width: number; height: number; busy: boolean; openEnabled: boolean; editEnabled: boolean; } | null;
  drawingToolContextActive: boolean;
  drawingContextKind: "shape" | "measurement" | "ink";
  drawingStyleEnabled: boolean;
  drawingStyleDisabledReason: string | undefined;
  commentToolContextLabel: string;
  ribbonTabs: PdfOfficeRibbonTab<PdfRibbonTab>[];
  /** The Draw tab (绘图): Word's Draw tab over the reader's ink. */
  readerDrawModel: PdfDrawRibbonModel;
  /** Pen-first drawing: the fold control after the tabs (lib/pen-first-draw). */
  readerDrawCollapse?: { onCollapse: () => void };
  /** The ribbon is folded into the pen palette. */
  readerPenCompact?: boolean;
  documentEncryptDisabled: boolean;
  documentEncryptDisabledReason: string;
  protectLabels: { encryption: string; redaction: string; validation: string; safety: string; encryptCopy: string; markRedactions: string; applyRedactions: string; safeCopy: string; archivalCopy: string; signatureStatus: string; actionStatus: string; };
  totalPagesLabel: string;
  /** The current page's document label when it is not its number. */
  currentPageDocumentLabel?: string | null;
  insertPagesFromFileCommands: readonly PdfInsertRibbonCommand[];
  insertPageMenuBlankCommand: { displayLabel: string; menu?: ReactNode; menuLabel?: string | undefined; active?: boolean | undefined; busy?: boolean | undefined; controlId?: string | undefined; disabled?: boolean | undefined; disabledReason?: string | undefined; expanded?: boolean | undefined; icon: PdfInsertRibbonIconName; label: string; menuControlId?: string | undefined; onInvoke?: ((trigger: HTMLButtonElement | null, resolveTrigger?: (() => HTMLButtonElement | null) | undefined) => void) | undefined; popup?: "dialog" | "menu" | undefined; testId: string; triggerRef?: RefObject<HTMLButtonElement | null> | undefined; } | undefined;
  homeRotateLeftCommand: PdfInsertRibbonCommand | undefined;
  homeRotateRightCommand: PdfInsertRibbonCommand | undefined;
  pagePersistentStructureCommands: PdfPageRibbonCommand[];
  pageFileCommands: PdfPageRibbonCommand[];
  pageRotateCommands: PdfPageRibbonCommand[];
  loadingDisabledReason: string;
  outputAuthorityDisabledReason: string;
  sourceAuthorityDisabledReason: string;
  sourceAuthorityUnavailable: boolean;
  printCommandDisabledReason: string | undefined;
  commentMarkupDisabled: boolean;
  commentMarkupDisabledReason: string | undefined;
  areaHighlightDisabled: boolean;
  areaHighlightDisabledReason: string | undefined;
  pageNavigation: JSX.Element;
  selectionTranslationDisabledReason: string | undefined;
  signatureCommandDisabledReason: string | undefined;
  legacyToolsHostRequired: boolean;
  featuredScreenRecordingStatus: string | null;
  featuredStatus: string | null;
  globalCommentToolsCommands: PdfCommentRibbonCommand[];
  globalPdfToolsCommands: { fillForm: { active: boolean; disabled: boolean; disabledReason: string | undefined; icon: "formFields"; label: string; onInvoke: () => void; testId: string; }; fillFormAs?: { command: PdfFeaturedCommand; items: readonly PdfFeaturedCommand[] }; addFormField: { command: { disabled: boolean; disabledReason: string | undefined; icon: "formFields"; label: string; testId: string; }; items: readonly PdfFeaturedCommand[]; }; addFormAction: { command: { disabled: boolean; disabledReason: string | undefined; icon: "printButton"; label: string; testId: string; }; items: readonly PdfFeaturedCommand[]; }; editFormField: { active: boolean; buttonRef: RefObject<HTMLButtonElement | null>; disabled: boolean; disabledReason: string | undefined; icon: "formFields"; label: string; onInvoke: (trigger: HTMLButtonElement | null) => void; testId: string; }; deleteFormField: { active: boolean; buttonRef: RefObject<HTMLButtonElement | null>; disabled: boolean; disabledReason: string | undefined; icon: "deleteField"; label: string; onInvoke: (trigger: HTMLButtonElement | null) => void; testId: string; }; detectFormFields: { buttonRef: RefObject<HTMLButtonElement | null>; disabled: boolean; disabledReason: string | undefined; displayLabel: string; icon: "formFields"; label: string; onInvoke: () => undefined; testId: string; } | { buttonRef: RefObject<HTMLButtonElement | null>; disabled: boolean; disabledReason: string | undefined; icon: "formFields"; label: string; onInvoke: () => undefined; testId: string; displayLabel?: undefined; }; cancelDetectedFields: { disabled: boolean; icon: "cancel"; label: string; onInvoke: () => void; testId: string; } | undefined; exportFormData: { disabled: boolean; disabledReason: string | undefined; icon: "exportData"; label: string; onInvoke: (trigger: HTMLButtonElement | null) => void; testId: string; }; importFormData: { disabled: boolean; disabledReason: string | undefined; icon: "importData"; label: string; onInvoke: (trigger: HTMLButtonElement | null) => void; testId: string; }; resetForm: { disabled: boolean; disabledReason: string | undefined; icon: "resetData"; label: string; onInvoke: () => undefined; testId: string; }; versions: { disabled: boolean; disabledReason: string | undefined; hidden: boolean; icon: "versions"; label: string; onInvoke: ((trigger: HTMLButtonElement | null) => void) | undefined; testId: string; }; properties: { disabled: boolean; icon: "properties"; label: string; onInvoke: (trigger: HTMLButtonElement | null) => void; testId: string; }; deskew: { disabled: boolean; disabledReason: string | undefined; displayLabel: string; icon: "deskew"; label: string; onInvoke: () => undefined; testId: string; }; moreTools: { icon: "more"; label: string; onInvoke: () => void; testId: string; }; };
  globalLinkDisabled: boolean;
  globalLinkDisabledReason: string | undefined;
  globalLinkLabel: string;
  editShapesRibbonCommand: (testId: string) => PdfEditRibbonCommand;
  editShapeRibbonItems: (testIdPrefix: string) => PdfEditRibbonCommand[];
}

/** The Shape menu's stroke weights, in points (Office's line weights). */
const SHAPE_MENU_WIDTH_POINTS = [0.5, 0.75, 1, 1.5, 2.25, 3, 4.5, 6] as const;

/** The measure tool's keep / remove-last commands, from the Comment tab's
 * list, as the Drawing Tools tab shows them. */
/** The comment tools WPS shows on the 批注 tab itself, not under More. */
const COMMENT_TOOLS_ON_SHOW = new Set([
  "pdf-comment-stamp",
  "pdf-comment-measure",
  "pdf-comment-measure-save",
  "pdf-comment-measure-remove-last",
]);

function drawingMeasurementCommands(
  commands: readonly PdfCommentRibbonCommand[],
): { keep?: PdfCommentRibbonCommand; removeLast?: PdfCommentRibbonCommand } {
  const keep = commands.find((command) => command.testId === "pdf-comment-measure-save");
  const removeLast = commands.find(
    (command) => command.testId === "pdf-comment-measure-remove-last",
  );
  return {
    ...(keep ? { keep: { ...keep, testId: "pdf-drawing-measure-save" } } : {}),
    ...(removeLast
      ? { removeLast: { ...removeLast, testId: "pdf-drawing-measure-remove-last" } }
      : {}),
  };
}

export function PdfViewerRibbon({ ribbon }: { ribbon: PdfViewerRibbonModel }) {
  const {
    filename,
    downloadUrl,
    ribbonQuickActions,
    ribbonActionRail,
    ribbonQuickActionTail,
    ribbonQuickOverflowOnlyTail,
    annotationContextWritable,
    onPageOrganizerApply,
    onDownload,
    preferAnnotatedDownload,
    annotatedDownloadEnabled,
    pageId,
    onRequestPageSummary,
    onManagedSignatureUpsert,
    onManagedStampUpsert,
    onManagedLinkUpsert,
    onManagedProofingUpsert,
    dataItemId,
    lossyCompressionSource,
    lossyCompressionWritable,
    nativePageLabelRanges,
    onNativePageLabelRangesApply,
    onNativeFileAttachmentUpsert,
    nativeFileAttachmentBusy,
    nativeFileAttachmentDisabledReason,
    onNativePageImport,
    nativePageImportBusy,
    nativePageImportDisabledReason,
    onEditContent,
    contentEditModeActive,
    contentEditModeDisabledReason,
    onToggleContentEditMode,
    onEnterContentEditMode,
    onExitContentEditMode,
    onInsertText,
    onCropPage,
    eraseRibbon,
    onSplitPage,
    splitPageActive,
    splitPageDisabledReason,
    onOpenPageBackground,
    pageBackgroundOpen,
    pageBackgroundBusy,
    pageBackgroundDisabledReason,
    onCreateEditableVersion,
    onConvertToSpreadsheet,
    onConvertToPresentation,
    onOfficeConversionPrepare,
    officeProjectionDisabledReason,
    isCreatingEditableVersion,
    t,
    pdfHomeDisplayLabels,
    savePolicyV1,
    signatureProofLoading,
    matchedSourceRiskSnapshot,
    signatureProofPresentation,
    showSignatureProof,
    pageOrganizerLabel,
    pages,
    loading,
    setPageOrganizerOpen,
    pageOrganizerDialogReturnFocusRef,
    scale,
    activeRibbonTab,
    ribbonInstanceId,
    currentPage,
    currentVisibleOrdinal,
    pageInput,
    setPageInput,
    pendingRibbonMarkupStyle,
    setPendingRibbonMarkupStyle,
    pendingProofingKind,
    proofingBusy,
    proofingError,
    ribbonMarkupError,
    ribbonMarkupColor,
    noteMode,
    commentInkEraser,
    commentInkStraight,
    measureResults,
    measureSnap,
    armedTextOrigin,
    armedTextKind,
    noteColor,
    commentModeActive,
    commentsHidden,
    attachMode,
    attachmentError,
    placedAttachments,
    ocgLayers,
    ocgOrder,
    setLayersPanelOpen,
    layersPanelOpen,
    handleToggleLayer,
    inks,
    drawMode,
    commentDrawMode = drawMode,
    setInkColor,
    inkColor,
    inkWidth,
    setInkWidth,
    inkSaving,
    signMode,
    signSaving,
    signatureComposerMode,
    stampPickerOpen,
    armedStamp,
    redactions,
    redactMode,
    redactSaving,
    redactionHydrationStatusForCurrentDocument,
    redactionProductCopy,
    shapesHydrationStatusForCurrentDocument,
    shapeMode,
    setShapeMode,
    shapeToolPageContentRef,
    shapeKind,
    setShapeKind,
    setShapeColor,
    shapeColor,
    shapeWidth,
    setShapeWidth,
    shapeLineStyle,
    setShapeLineStyle,
    setShapeLineStarts,
    setShapeLineEnds,
    activeShapeLineKind,
    shapeLineStart,
    shapeLineEnd,
    shapeFillColor,
    setShapeFillColor,
    areaHighlightActive,
    setAreaHighlightActive,
    areaHighlightColor,
    areaHighlightOpacity,
    visualMaskActive,
    setVisualMaskActive,
    visualMaskWarning,
    shapeSaving,
    selectedPlaced,
    selectedPlacedRef,
    placedAnnotationClipboardError,
    linkMode,
    linkSaving,
    linkPlacementDestinationKindRef,
    snapshotMode,
    setSnapshotMode,
    snapshotIntentRef,
    snapshotShape,
    setSnapshotShape,
    captureScreenSnapshot,
    setDrawingSnapshot,
    snapshotDragRef,
    snapshotTextReturnFocusRef,
    setSnapshotToast,
    setSnapshotLatex,
    measureMode,
    measureKind,
    setMeasureKind,
    measureUnit,
    setMeasureUnit,
    measureRatio,
    setMeasureRatio,
    setMeasurePts,
    pageLabelsOpen,
    pageLabelsDialogReturnFocusRef,
    pageLabelRanges,
    pageLabelReplacementStarted,
    setPageLabelReplacementStarted,
    pageLabelsApplying,
    pageLabelsError,
    nativePageLabelsControlled,
    nativePageLabelsSourceOwned,
    updatePageLabelRanges,
    printing,
    printPreflightTriggerRef,
    batchPrintOpen,
    batchPrintPreparing,
    decorationDialogKind,
    effectiveDocumentDecorations,
    openDecorationDialog,
    removeDocumentDecoration,
    annotationsOpen,
    outline,
    setPropertiesOpen,
    propertiesDialogReturnFocusRef,
    userBookmarks,
    bookmarkSaving,
    flattening,
    formDataExportFormat,
    setFormDataExportFormat,
    compressPreset,
    setCompressPreset,
    openCompressSettings,
    mergeBusy,
    splitBusy,
    featuredImagesToPdfBusy,
    featuredMarkdownToPdfBusy,
    featuredCadToPdfBusy,
    featuredMarkdownToPdfAvailable,
    setSplitOpen,
    splitOpen,
    splitDialogReturnFocusRef,
    featuredLegacyToolsOpen,
    setFeaturedLegacyToolsOpen,
    setExportImageOpen,
    exportImageOpen,
    exportImageDialogReturnFocusRef,
    exportImageScope,
    setExportImageScope,
    exportImageRange,
    setExportImageRange,
    exportImageCombinePages,
    setExportImageCombinePages,
    exportImageResolution,
    setExportImageResolution,
    exportImageBusy,
    textExportBusy,
    imageExtractionBusy,
    translationDialog,
    translationWorking,
    cadExportBusy,
    screenRecordingState,
    screenRecordingSupported,
    documentEncryptBusy,
    requestSignatureProofPanel,
    dangerousActionSafeCopyBusy,
    verifiedPdfaExportBusy,
    verifiedPdfaExportTriggerRef,
    imageOnlyExportBusy,
    xfdfImportNotice,
    rasterOutputAuthorityBlocked,
    splitMode,
    setSplitMode,
    splitRangeSpec,
    splitRangeError,
    setSplitRangeSpec,
    splitEveryNValue,
    setSplitEveryNValue,
    featuredImagesToPdfInputRef,
    featuredMarkdownToPdfInputRef,
    featuredCadToPdfInputRef,
    featuredImagesToPdfReturnFocusRef,
    setInsertOpen,
    insertOpen,
    selectInsertSourceKind,
    insertDialogReturnFocusRef,
    insertBusy,
    editCropTriggerRef,
    layersDialogReturnFocusRef,
    formImportInputRef,
    formImportReturnFocusRef,
    viewMode,
    setViewMode,
    fitMode,
    panMode,
    presentationActive,
    compareInputRef,
    compareFile,
    viewerRootRef,
    ocrReviewByPage,
    ocrAnnotationHydrationReady,
    ocrCandidatePages,
    ocrPendingPages,
    ocrPreflightLoadingPage,
    ocrPreflightTriggerRef,
    pdfOcrServerSourceAuthority,
    pageOrganizerPages,
    visiblePageCount,
    previousVisibleOrdinal,
    nextVisibleOrdinal,
    rotateDocumentDisabled,
    goToVisibleOrdinal,
    handleSummarizeCurrentPage,
    pageOpsEnabled,
    handleRotatePage,
    handleRotateCcwPage,
    handleRotateDocumentClockwise,
    handlePageInputCommit,
    selectionEnabled,
    clearPlacedAnnotationSelection,
    requestOcrPreflight,
    openOcrReview,
    handleInkUndoLast,
    pendingRedactionReviewItems,
    handleRedactionReviewNavigate,
    handleRedactionReviewDelete,
    handleSelectedManagedShapeStyle,
    handleSelectedManagedShapeArrange,
    handleSelectedManagedInkStyle,
    handleSelectedShapeSizeChange,
    handleCopyPlacedAnnotation,
    handlePastePlacedAnnotation,
    handleDuplicatePlacedAnnotation,
    handleDeleteSelectedPlacedAnnotation,
    handleInsertBookmark,
    clearSelectedLinkContext,
    openSelectedLinkContext,
    editSelectedLinkContext,
    deleteSelectedLinkContext,
    allFormFields,
    formFieldsUnavailable,
    formFieldsUnavailableReason,
    nativeFormMode,
    formMutationBlocked,
    formMutationBlockedReason,
    nativeFormDataMutationBlocked,
    nativeFormDataMutationBlockedReason,
    hasNativeFormEditsExcludedFromSource,
    hasFlattenableAnnotations,
    hasEffectivePdfOutputEdits,
    hasEditableXfdfAnnotations,
    handleCompressPdf,
    handleExportFormData,
    handleResetForm,
    fieldCandidates,
    detectBusy,
    detectedFieldsCommandRef,
    existingTextFieldEditMode,
    selectedFormFieldContextRef,
    existingTextFieldEditBusy,
    existingFormFieldDeleteMode,
    existingFormFieldDeleteBusy,
    currentDetectedFieldSourceIdentityRef,
    handleDetectFormFields,
    acceptedCandidates,
    clearSelectedFormFieldContext,
    openSelectedFormFieldProperties,
    openSelectedFormFieldOptions,
    hasVisibleFormFieldForPropertyEdit,
    hasVisibleFormFieldForDelete,
    handleCancelDetectedFields,
    handleSelectedFormFieldTransform,
    applySelectedFormWidgetAppearance,
    applySelectedFormFieldQuickProperty,
    handleCreateDetectedFields,
    handleExportPagesAsImages,
    exportImageOverLimit,
    exportImageCurrentPageSize,
    openBatchPrint,
    handleOpenImageExtractionReview,
    handleSplitDownload,
    handleOpenMergePicker,
    executeReaderCommand,
    originalDownloadRuntimeDecision,
    currentDownloadRuntimeDecision,
    annotatedDownloadRuntimeDecision,
    fillableDownloadRuntimeDecision,
    editableAnnotationsDownloadRuntimeDecision,
    archivalDownloadRuntimeDecision,
    printRuntimeDecision,
    applyRedactionsRuntimeDecision,
    originalDownloadCommand,
    handleSignatureProofOriginalDownload,
    currentDownloadCommand,
    annotatedDownloadCommand,
    fillableDownloadCommand,
    editableAnnotationsDownloadCommand,
    archivalDownloadCommand,
    printCommand,
    applyRedactionsCommand,
    activeSplitCommand,
    zoomReset,
    fitWidth,
    fitPage,
    disarmPageTools,
    toggleRedactionMarkMode,
    openRedactPages,
    openHiddenInformation,
    handleToggleCommentMode,
    handleToggleCommentsHidden,
    handleRibbonMarkupColorChange,
    handleNoteColorChange,
    handleAreaHighlightColorChange,
    handleAreaHighlightOpacityChange,
    closePageLabelsDialog,
    togglePageLabelsDialog,
    applyPageLabelRanges,
    toggleVisualMaskTool,
    toggleVisualSignature,
    armLinkMode,
    deleteAllLinks,
    documentLinkCount,
    deleteAllLinksAvailable,
    toggleLinkMode,
    selectRibbonTab,
    handleHandToolToggle,
    handleSelectTool,
    handleToggleExistingTextFieldEditMode,
    handleToggleExistingFormFieldDeleteMode,
    selectToolActive,
    executeMarkupActivation,
    executeProofingActivation,
    executeFreeTextActivation,
    readerTextBoxes,
    readerTables,
    readerCharts,
    executeImageActivation,
    executeNoteActivation,
    executeInkActivation,
    executeShapeActivation,
    executeAreaHighlightActivation,
    executeStampActivation,
    executeLinkActivation,
    executeVisualSignatureActivation,
    executeFileAttachmentActivation,
    handlePresentationToggle,
    deskewStatus,
    speechStatus,
    speechSupported,
    handleToggleReadAloud,
    handleStopReadAloud,
    handleSaveSelectedAttachment,
    handleExportDocumentText,
    openFullTranslation,
    openSelectionTranslation,
    openOfficeProjection,
    openScreenRecording,
    stopScreenRecording,
    openCadExport,
    dangerousActionSafeCopyDisabledReason,
    dangerousActionCount,
    dangerousActionCompactStatusLabel,
    dangerousActionSafeCopyCommand,
    executeDangerousActionSafeCopyReview,
    openDocumentEncrypt,
    openImageOnlyExport,
    toggleNavigationPanel,
    docSearchOpen,
    effectiveRibbonSaveStatus,
    effectiveRibbonSaveStatusLabel,
    currentNativeVectorSelection,
    maskCurrentNativeVector,
    toggleNativeVectorVertexEditing,
    toggleNativeSelectionPane,
    nativeSelectionPaneOpen,
    nativeVectorClipboardCommands,
    runOrQueueNativeVectorCommand,
    annotationContextModel,
    effectiveAnnotationContextLabels,
    formFieldContextLabels,
    formFieldContextModel,
    linkContextLabels,
    linkContextModel,
    drawingToolContextActive,
    drawingContextKind,
    drawingStyleEnabled,
    drawingStyleDisabledReason,
    commentToolContextLabel,
    readMode,
    cycleReadMode,
    lineNumbersVisible,
    toggleLineNumbers,
    ribbonTabs,
    readerDrawModel,
    readerDrawCollapse,
    readerPenCompact = false,
    documentEncryptDisabled,
    documentEncryptDisabledReason,
    protectLabels,
    totalPagesLabel,
    currentPageDocumentLabel = null,
    insertPagesFromFileCommands,
    insertPageMenuBlankCommand,
    homeRotateLeftCommand,
    homeRotateRightCommand,
    pagePersistentStructureCommands,
    pageFileCommands,
    pageRotateCommands,
    loadingDisabledReason,
    outputAuthorityDisabledReason,
    sourceAuthorityDisabledReason,
    sourceAuthorityUnavailable,
    printCommandDisabledReason,
    commentMarkupDisabled,
    commentMarkupDisabledReason,
    areaHighlightDisabled,
    areaHighlightDisabledReason,
    pageNavigation,
    selectionTranslationDisabledReason,
    signatureCommandDisabledReason,
    legacyToolsHostRequired,
    featuredScreenRecordingStatus,
    featuredStatus,
    globalCommentToolsCommands,
    globalPdfToolsCommands,
    editShapesRibbonCommand,
    editShapeRibbonItems,
    formatPaneDock,
  } = ribbon;
  // A shape's stroke is kept as a share of the page's shorter side; the
  // Drawing Tools field and the Shape menu show and take points, as Shape
  
  // Letter, which Shape editing then showed as 12.24 (user audit PDF 0015).
  const shapeWidthReferencePage =
    pages.find((page) => page.pageNumber === currentPage) ?? pages[0];
  const shapeWidthPageMinPoints = shapeWidthReferencePage
    ? Math.min(shapeWidthReferencePage.width, shapeWidthReferencePage.height)
    : 612;
  const shapeWidthPoints = (width: number) =>
    Math.round(shapeWidthToPoints(width, shapeWidthPageMinPoints) * 100) / 100;
  const shapeWidthFromPointsOnPage = (points: number) =>
    shapeWidthFromPoints(points, shapeWidthPageMinPoints);
  const shapeMenuWidthPoints = Array.from(
    new Set([
      ...SHAPE_MENU_WIDTH_POINTS.filter(
        (points) =>
          points >= shapeWidthPoints(PDF_SHAPE_MIN_WIDTH) &&
          points <= shapeWidthPoints(PDF_SHAPE_MAX_WIDTH),
      ),
      shapeWidthPoints(shapeWidth),
    ]),
  ).sort((left, right) => left - right);

  // 识别文本 (Special Features) and WPS' 扫描件编辑 (Edit) are one command:
  // OCR the scanned page in view so its words can be edited.
  const recognizeTextCommand = {
    disabled:
      loading ||
      matchedSourceRiskSnapshot === null ||
      pdfOcrServerSourceAuthority === null ||
      pdfOcrServerSourceAuthority.derived === true ||
      !ocrAnnotationHydrationReady ||
      ocrPendingPages.has(currentPage) ||
      ocrPreflightLoadingPage === currentPage ||
      (!ocrCandidatePages.has(currentPage) &&
        !ocrReviewByPage.has(currentPage)),
    // Why, in the order the checks run: the page has to be a
    // scanned page of the version the server binds OCR to (the
    // upload, or an edited page's current saved version).
    disabledReason: resolvePdfCommandDisabledReason(
      true,
      loading ? loadingDisabledReason : undefined,
      matchedSourceRiskSnapshot === null || pdfOcrServerSourceAuthority === null
        ? t("pdfFeatured.ocrSourcePending")
        : undefined,
      pdfOcrServerSourceAuthority?.derived === true
        ? t("pdfFeatured.ocrEditedSource")
        : undefined,
      !ocrAnnotationHydrationReady ||
        ocrPendingPages.has(currentPage) ||
        ocrPreflightLoadingPage === currentPage
        ? t("pdfFeatured.ocrPageBusy")
        : undefined,
      t("pdfFeatured.ocrPageHasText"),
    ),
    label: t("pdfFeatured.recognizeText"),
    onInvoke: (trigger: HTMLButtonElement | null) => {
      ocrPreflightTriggerRef.current = trigger;
      if (
        ocrReviewByPage.get(currentPage)?.status === "pending_review"
      ) {
        openOcrReview(currentPage);
      } else {
        void requestOcrPreflight(currentPage);
      }
    },
  };

  // WPS' screenshot menu regions: choosing one arms the tool with it, and the
  // split button's face then takes that shape until another is chosen.
  const snapshotShapeCommand = (
    shape: SnapshotShape,
    labelKey: string,
  ): PdfHomeCompactCommand => ({
    active: snapshotMode && snapshotShape === shape,
    disabled: loading || rasterOutputAuthorityBlocked,
    label: t(labelKey),
    onInvoke: () => {
      snapshotIntentRef.current = "copy";
      setSnapshotShape(shape);
      disarmPageTools("snapshot");
      snapshotDragRef.current = null;
      setDrawingSnapshot(null);
      setSnapshotToast(null);
      setSnapshotLatex(null);
      setSnapshotMode(true);
    },
    testId: `pdf-home-screenshot-shape-${shape}`,
  });

  // Shape Editing's style, size and turn: one set of commands for the tab and
  // its Format Shape pane.
  const [shapeFormatPaneOpen, setShapeFormatPaneOpen] = useState(false);
  const annotationStyleChange = (change: PdfAnnotationStyleChange) => {
    if (currentNativeVectorSelection) {
      runOrQueueNativeVectorCommand((run) => run.style(change));
    } else if (
      selectedPlacedRef.current?.kind === "ink" ||
      selectedPlacedRef.current?.kind === "nativeInk"
    ) {
      void handleSelectedManagedInkStyle(change);
    } else {
      void handleSelectedManagedShapeStyle(change);
    }
  };
  const annotationSizeChange = currentNativeVectorSelection
    ? (axis: "width" | "height", valuePoints: number, lockAspectRatio?: boolean) =>
        runOrQueueNativeVectorCommand((run) =>
          run.resize(axis, valuePoints, lockAspectRatio),
        )
    : handleSelectedShapeSizeChange;
  const annotationRotateBy = (degrees: number) => {
    if (currentNativeVectorSelection) {
      runOrQueueNativeVectorCommand((run) => run.rotate(degrees));
      return;
    }
    // A placed shape turns in quarter turns only.
    const quarterTurns = ((Math.round(degrees / 90) % 4) + 4) % 4;
    if (quarterTurns === 0) return;
    void (async () => {
      if (quarterTurns === 3) {
        await handleSelectedManagedShapeArrange("rotate-left");
        return;
      }
      for (let turn = 0; turn < quarterTurns; turn += 1) {
        await handleSelectedManagedShapeArrange("rotate-right");
      }
    })();
  };
  const shapeFormatPaneShown = shapeFormatPaneOpen && annotationContextModel?.kind === "shape" && Boolean(formatPaneDock);

  return (
    <PdfOfficeRibbon
      activeTab={activeRibbonTab}
      ariaLabel={t("pdf.viewMode")}
      className="pdf-viewer__office-ribbon"
      onActiveTabChange={selectRibbonTab}
      actionRail={{
        ...ribbonActionRail,
        // Print in the title row, as Docs, the spreadsheet and the
        // presentation have it: the same Print command (preflight,
        // final-byte check, disabled reason) as Tools > Print.
        leading: (
          <>
            <OfficeRailButton
              label={t("pdfViewer.print")}
              icon={<Printer />}
              busy={printing}
              disabled={!printCommand.enabled}
              title={printCommandDisabledReason ?? t("pdfViewer.print")}
              onClick={(event) => {
                printPreflightTriggerRef.current = event.currentTarget;
                void executeReaderCommand(
                  PDF_READER_PRINT_COMMAND,
                  printRuntimeDecision,
                );
              }}
              testId="pdf-rail-print"
            />
            {/* Download in the title row beside Print, as the spreadsheet
                has its Export: the PDF as it now stands (its annotations
                and edits) or as it was uploaded, by the same reader
                commands (and their reasons) as Tools > Download. */}
            <DropdownMenu modal={false}>
              <DropdownMenuTrigger asChild>
                <OfficeRailButton
                  menu
                  label={t("references.download")}
                  icon={<Download />}
                  disabled={!currentDownloadCommand.enabled && !originalDownloadCommand.enabled}
                  testId="pdf-rail-download"
                />
              </DropdownMenuTrigger>
              <RibbonDropdownMenuContent align="end" className="docs-cupertino-menu--described note-editor-export-menu">
                <DropdownMenuLabel className="docs-cupertino-menu__heading">
                  {t("references.download")}
                </DropdownMenuLabel>
                <DropdownMenuItem
                  className="docs-cupertino-menu__rich-item"
                  disabled={!currentDownloadCommand.enabled}
                  onSelect={() =>
                    void executeReaderCommand(PDF_READER_DOWNLOAD_CURRENT_COMMAND, currentDownloadRuntimeDecision)}
                  data-testid="pdf-rail-download-current"
                >
                  <RibbonMenuRichRow
                    fileType="pdf"
                    hoverScene
                    title={t("pdfViewer.downloadAnnotated")}
                    detail={currentDownloadCommand.enabled ? undefined : currentDownloadCommand.reasonLabel}
                    blocked={!currentDownloadCommand.enabled}
                    badge="PDF"
                  />
                </DropdownMenuItem>
                <DropdownMenuItem
                  className="docs-cupertino-menu__rich-item"
                  disabled={!originalDownloadCommand.enabled}
                  onSelect={() =>
                    void executeReaderCommand(PDF_READER_DOWNLOAD_ORIGINAL_COMMAND, originalDownloadRuntimeDecision)}
                  data-testid="pdf-rail-download-original"
                >
                  <RibbonMenuRichRow
                    fileType="pdf"
                    hoverScene
                    title={t("pdf.viewer.downloadOriginal")}
                    detail={originalDownloadCommand.enabled ? undefined
                      : originalDownloadCommand.reasonLabel ?? t("pdf.viewer.documentNoDownloadableOriginalPdf")}
                    blocked={!originalDownloadCommand.enabled}
                    badge="PDF"
                  />
                </DropdownMenuItem>
              </RibbonDropdownMenuContent>
            </DropdownMenu>
            {ribbonActionRail?.leading}
          </>
        ),
        // One Comments button, as every editor has: the window's (its
        // comments pane on the right) where the reader stands in a window,
        // else the reader's own annotations list. The window's pane also
        // showed an edge tab of its own beside it: two buttons for comments.
        primary: ribbonActionRail?.primary
          ? { ...ribbonActionRail.primary, label: t("pdfViewer.annotationsPanel") }
          : {
            kind: "comments",
            label: t("pdfViewer.annotationsPanel"),
            onClick: () => toggleNavigationPanel("annotations"),
            pressed: annotationsOpen,
            disabled: loading,
            testId: "pdf-action-comments",
          },
      }}
      quickOverflowLabel={t("toolbar.quickOverflow")}
      saveStatus={effectiveRibbonSaveStatus}
      saveStatusLabel={effectiveRibbonSaveStatusLabel}
      scrollLeftLabel={t("toolbar.ribbonScrollLeft")}
      scrollRightLabel={t("toolbar.ribbonScrollRight")}
      tabs={ribbonTabs}
      testId="pdf-ribbon"
      instanceId={ribbonInstanceId}
      collapse={readerDrawCollapse}
      penCompact={readerPenCompact}
    >
      {/* 绘图: its groups show while the Draw tab is the ribbon's tab. */}
      <PdfDrawRibbonTab t={t} model={readerDrawModel} testIdPrefix="pdf-reader-draw" />
      <PdfFileRibbonTab
        visible={activeRibbonTab === "tools" && featuredLegacyToolsOpen}
        ribbon="tools"
        error={null}
        outputGroupLabel={t("pdf.viewer.print")}
        documentGroupLabel={t("pdf.viewer.document")}
        print={{
          busy: printing,
          disabled: !printCommand.enabled,
          disabledReason: printCommandDisabledReason,
          label: t("pdfViewer.print"),
          onInvoke: (trigger) => {
            printPreflightTriggerRef.current = trigger;
            void executeReaderCommand(
              PDF_READER_PRINT_COMMAND,
              printRuntimeDecision,
            );
          },
          testId: "pdf-file-print",
        }}
        batchPrint={{
          busy: batchPrintPreparing,
          disabled:
            loading ||
            printing ||
            flattening ||
            batchPrintOpen ||
            hasEffectivePdfOutputEdits ||
            matchedSourceRiskSnapshot === null ||
            !dataItemId,
          disabledReason: hasEffectivePdfOutputEdits
            ? t("pdf.viewer.finishSaveCurrentPdfEdits2")
            : matchedSourceRiskSnapshot === null || !dataItemId
              ? t("pdf.viewer.currentPdfSourceMustVerified")
              : undefined,
          label: t("pdf.viewer.batchPrint"),
          onInvoke: openBatchPrint,
          testId: "pdf-file-batch-print",
        }}
        downloadOriginal={{
          disabled: !originalDownloadCommand.enabled,
          disabledReason: resolvePdfCommandDisabledReason(
            !originalDownloadCommand.enabled,
            originalDownloadCommand.reasonLabel,
            originalDownloadRuntimeDecision.reasonLabel,
            loading ? loadingDisabledReason : undefined,
            t("pdf.viewer.documentNoDownloadableOriginalPdf"),
          ),
          label: t("pdf.viewer.downloadOriginal"),
          onInvoke: () =>
            void executeReaderCommand(
              PDF_READER_DOWNLOAD_ORIGINAL_COMMAND,
              originalDownloadRuntimeDecision,
            ),
          testId: "pdf-file-download-original",
        }}
        properties={{
          disabled: loading || flattening,
          label: t("pdfViewer.properties"),
          onInvoke: (trigger) => {
            propertiesDialogReturnFocusRef.current = trigger;
            setPropertiesOpen(true);
          },
          testId: "pdf-file-properties",
        }}
      />

      <PdfInsertRibbonTab
        pointerTools={
          <PdfEditPointerTools
            controlScope={activeRibbonTab === "page" ? "PAGE" : "INSERT"}
            groupLabel={t("pdf.pointerTools")}
            groupTestId="pdf-insert-pointer-tools"
            handActive={panMode}
            handActiveLabel={t("pdfViewer.handToolActive")}
            handDisplayLabel={pdfHomeDisplayLabels.hand}
            handLabel={t("pdfViewer.handTool")}
            handTestId="pdf-insert-hand-tool"
            loading={loading}
            loadingLabel={t("pdfViewer.loading")}
            onHandActivate={handleHandToolToggle}
            onSelectActivate={handleSelectTool}
            selectActive={selectToolActive}
            selectLabel={t("toolbar.select")}
            selectTestId="pdf-insert-select-tool"
            tabs="insert page"
          />
        }
        pages={{
          bookmark: {
            busy: bookmarkSaving,
            disabled:
              loading ||
              !dataItemId ||
              bookmarkSaving ||
              userBookmarks.length >= PDF_BOOKMARK_MAX,
            disabledReason: !dataItemId
              ? t("pdf.viewer.documentNoWritablePdfData")
              : userBookmarks.length >= PDF_BOOKMARK_MAX
                ? t("pdf.viewer.bookmarkLimitBeenReached")
                : loading
                  ? t("pdfViewer.loading")
                  : undefined,
            icon: "bookmark",
            label: t("pdfViewer.bookmarkAdd"),
            onInvoke: (trigger) => void handleInsertBookmark(trigger),
            testId: "pdf-insert-bookmark",
          },
          organize: onPageOrganizerApply
            ? {
                disabled: loading || pageOrganizerPages.length === 0,
                icon: "organize",
                label: pageOrganizerLabel,
                onInvoke: (trigger) => {
                  pageOrganizerDialogReturnFocusRef.current = trigger;
                  setPageOrganizerOpen(true);
                },
                popup: "dialog",
                testId: "pdf-page-organizer-open",
              }
            : undefined,
          insertBlank: insertPageMenuBlankCommand,
          secondary: insertPagesFromFileCommands,
        }}
        pageTab={{
          files: pageFileCommands,
          navigation: pageNavigation,
          rotate: pageRotateCommands,
          structure: pagePersistentStructureCommands,
          
          // tools are the page view's own -- a page copied and pasted from
          // its menu, moved by a drag (or cut and pasted, Ctrl/⌘+↑↓), turned
          // by 向左/向右旋转 on the pages picked.
        }}
        decoration={{
          background: {
            active: pageBackgroundOpen,
            busy: pageBackgroundBusy,
            disabled:
              loading ||
              flattening ||
              pageBackgroundBusy ||
              !onOpenPageBackground,
            disabledReason:
              pageBackgroundDisabledReason ??
              (onOpenPageBackground
                ? undefined
                : t("pdf.viewer.pageBackgroundOutputUnavailableHost")),
            icon: "background",
            label: t("pdfViewer.background"),
            onInvoke: onOpenPageBackground,
            popup: "dialog",
            testId: "pdf-insert-background",
          },
          watermark: {
            active: decorationDialogKind === "watermark",
            disabled: loading || flattening || !dataItemId,
            icon: "watermark",
            label: t("pdfStamp.watermarkText"),
            displayLabel: t("pdf.viewer.watermark"),
            onInvoke: (trigger) =>
              openDecorationDialog("watermark", trigger),
            popup: "dialog",
            testId: "pdf-stamp-toggle",
          },
          watermarkItems: [
            {
              disabled: loading || flattening || !dataItemId,
              icon: "watermark",
              label: t("pdf.viewer.addEditWatermark"),
              onInvoke: (trigger) =>
                openDecorationDialog("watermark", trigger),
              popup: "dialog",
              testId: "pdf-insert-watermark-configure",
            },
            {
              disabled:
                loading ||
                flattening ||
                !dataItemId ||
                !effectiveDocumentDecorations.watermark,
              icon: "watermark",
              label: t("pdf.viewer.removeWatermark"),
              onInvoke: () => void removeDocumentDecoration("watermark"),
              testId: "pdf-insert-watermark-remove",
            },
          ],
          pageNumber: {
            active: decorationDialogKind === "page-number",
            disabled: loading || flattening || !dataItemId,
            icon: "pageNumber",
            label: t("pdfStamp.modePageNumber"),
            displayLabel: t("pdf.viewer.addPageNumber"),
            onInvoke: (trigger) =>
              openDecorationDialog("page-number", trigger),
            popup: "dialog",
            testId: "pdf-insert-page-number",
          },
          // WPS CT_Insertion Insert_pageNumber: add, update, delete.
          pageNumberItems: [
            {
              disabled: loading || flattening || !dataItemId,
              icon: "pageNumber",
              label: t("pdf.decoration.titlePageNumber"),
              onInvoke: (trigger) =>
                openDecorationDialog("page-number", trigger),
              popup: "dialog",
              testId: "pdf-insert-page-number-configure",
            },
            {
              disabled:
                loading ||
                flattening ||
                !dataItemId ||
                !effectiveDocumentDecorations.pageNumber,
              disabledReason: !effectiveDocumentDecorations.pageNumber
                ? t("pdf.viewer.documentNoPageNumber")
                : undefined,
              icon: "pageNumber",
              label: t("pdf.viewer.removePageNumber"),
              onInvoke: () => void removeDocumentDecoration("page-number"),
              testId: "pdf-insert-page-number-remove",
            },
          ],
          headerFooter: {
            active: decorationDialogKind === "header-footer",
            controlId: "PDFCTL.RIBBON.INSERT.HEADER_FOOTER.PRIMARY",
            disabled: loading || flattening || !dataItemId,
            icon: "headerFooter",
            label: t("pdfStamp.modeCustomText"),
            displayLabel: t("pdf.viewer.headerFooter"),
            menuControlId: "PDFCTL.RIBBON.INSERT.HEADER_FOOTER.MENU",
            onInvoke: (trigger) =>
              openDecorationDialog("header-footer", trigger),
            popup: "dialog",
            testId: "pdf-insert-header-footer",
          },
          headerFooterItems: [
            {
              disabled: loading || flattening || !dataItemId,
              icon: "headerFooter",
              label: t("pdf.viewer.addEditHeaderFooter"),
              onInvoke: (trigger) =>
                openDecorationDialog("header-footer", trigger),
              popup: "dialog",
              testId: "pdf-insert-header-footer-configure",
            },
            {
              disabled:
                loading ||
                flattening ||
                !dataItemId ||
                !effectiveDocumentDecorations.headerFooter,
              icon: "headerFooter",
              label: t("pdf.viewer.removeHeaderFooter"),
              onInvoke: () => void removeDocumentDecoration("header-footer"),
              testId: "pdf-insert-header-footer-remove",
            },
          ],
        }}
        content={{
          // WPS CT_Insertion RB_InsertPageMarks: text, image, signature, stamp.
          stamp: dataItemId
            ? {
                active: stampPickerOpen || armedStamp != null,
                disabled: loading || !onManagedStampUpsert,
                disabledReason: onManagedStampUpsert
                  ? undefined
                  : t("document.capability.unavailable"),
                icon: "stamp",
                label: t("pdfStampTool.modeToggle"),
                onInvoke: (trigger) => void executeStampActivation(trigger ?? null),
                testId: "pdf-insert-stamp",
              }
            : undefined,
          text: onInsertText
            ? {
                active: armedTextOrigin === "insert" && armedTextKind === "text",
                disabled: loading,
                displayLabel: t("pdf.viewer.text"),
                icon: "textBox",
                label: t("pdf.addText"),
                onInvoke: () => void executeFreeTextActivation("insert"),
                testId: "pdf-insert-text",
              }
            : undefined,
          equation: onInsertText && readerTextBoxes
            ? {
                active: armedTextOrigin === "insert" && armedTextKind === "formula",
                disabled: loading,
                displayLabel: t("toolbar.equation"),
                icon: "equation",
                label: t("toolbar.equation"),
                onInvoke: () => void executeFreeTextActivation("insert", "formula"),
                testId: "pdf-insert-equation",
              }
            : undefined,
          table: onInsertText && readerTables
            ? {
                active: armedTextOrigin === "insert" && armedTextKind.startsWith("table:"),
                disabled: loading,
                displayLabel: t("pdf.table.insert"),
                label: t("pdf.table.insert"),
                testId: "pdf-insert-table",
                onInsertTable: (rows: number, columns: number) => {
                  void executeFreeTextActivation("insert", `table:${rows}x${columns}`);
                  return true;
                },
              }
            : undefined,
          chart: onInsertText && readerCharts
            ? {
                disabled: loading,
                onPick: (family: string, subtype: string) => void executeFreeTextActivation("insert", `chart:${family}:${subtype}`),
              }
            : undefined,
          
          smartArt: onInsertText && readerCharts
            ? {
                active: armedTextOrigin === "insert" && armedTextKind.startsWith("smartart:"),
                disabled: loading,
                onPick: (layoutId: string) => void executeFreeTextActivation("insert", `smartart:${layoutId}`),
              }
            : undefined,
          wordArt: onInsertText && readerTextBoxes
            ? {
                active: armedTextOrigin === "insert" && armedTextKind.startsWith("wordart:"),
                disabled: loading,
                displayLabel: t("toolbar.wordArt"),
                icon: "wordArt",
                label: t("toolbar.wordArt"),
                popup: "menu",
                onPick: (recipeId: string) => void executeFreeTextActivation("insert", `wordart:${recipeId}`),
                testId: "pdf-insert-wordart",
              }
            : undefined,
          image: {
            disabled:
              loading || !dataItemId || !onManagedStampUpsert,
            disabledReason:
              dataItemId && onManagedStampUpsert
                ? undefined
                : t("pdf.editNotReady"),
            displayLabel: t("pdf.viewer.image"),
            icon: "image",
            label: t("pdf.insertImage"),
            onInvoke: (trigger) =>
              void executeImageActivation(trigger ?? null),
            testId: "pdf-insert-image",
          },
          link: {
            active: linkMode,
            disabled:
              loading ||
              linkSaving ||
              !dataItemId ||
              !onManagedLinkUpsert,
            disabledReason: !dataItemId || !onManagedLinkUpsert
              ? t("pdf.viewer.nativeLinkWritebackUnavailable")
              : linkSaving
                ? t("pdf.viewer.linkBeingSaved")
                : loading
                  ? t("pdfViewer.loading")
                  : undefined,
            icon: "link",
            label: t("pdfViewer.addLink"),
            onInvoke: (trigger) => {
              if (linkMode) toggleLinkMode(trigger);
              else void executeLinkActivation("url", trigger);
            },
            testId: "pdf-insert-link",
          },
          visualMask: {
            active: shapeMode && visualMaskActive,
            disabled: loading || shapeSaving || !dataItemId,
            disabledReason: !dataItemId
              ? t("pdf.viewer.documentNoWritablePdfData")
              : shapeSaving
                ? t("pdf.viewer.visualCoverBeingSaved")
                : loading
                  ? t("pdfViewer.loading")
                  : undefined,
            displayLabel: t("pdfShape.whiteout"),
            icon: "visualMask",
            label: `${t("pdfShape.whiteout")} — ${visualMaskWarning}`,
            onInvoke: () => toggleVisualMaskTool(),
            testId: "pdf-insert-visual-mask",
          },
        }}
        signatureDraw={{
          signature: {
            active: signMode,
            disabled:
              loading ||
              signSaving ||
              !dataItemId ||
              !onManagedSignatureUpsert,
            disabledReason: signatureCommandDisabledReason,
            displayLabel: t("pdf.viewer.signature"),
            icon: "signature",
            label: `${t("pdfSign.modeToggle")} — ${t("pdfSign.note")}`,
            onInvoke: (trigger) => {
              if (signMode) toggleVisualSignature(trigger);
              else void executeVisualSignatureActivation("draw", trigger);
            },
            testId: "pdf-insert-signature",
          },
          signatureItems: [
            {
              active: signMode && signatureComposerMode === "draw",
              disabled:
                loading ||
                signSaving ||
                !dataItemId ||
                !onManagedSignatureUpsert,
              disabledReason: signatureCommandDisabledReason,
              displayLabel: t("pdfSign.draw"),
              icon: "draw",
              label: `${t("pdfSign.draw")} — ${t("pdfSign.note")}`,
              onInvoke: (trigger) =>
                void executeVisualSignatureActivation("draw", trigger),
              testId: "pdf-insert-signature-draw",
            },
            {
              active: signMode && signatureComposerMode === "type",
              disabled:
                loading ||
                signSaving ||
                !dataItemId ||
                !onManagedSignatureUpsert,
              disabledReason: signatureCommandDisabledReason,
              displayLabel: t("pdfSign.type"),
              icon: "textBox",
              label: `${t("pdfSign.type")} — ${t("pdfSign.note")}`,
              onInvoke: (trigger) =>
                void executeVisualSignatureActivation("type", trigger),
              testId: "pdf-insert-signature-type",
            },
            {
              active: signMode && signatureComposerMode === "upload",
              disabled:
                loading ||
                signSaving ||
                !dataItemId ||
                !onManagedSignatureUpsert,
              disabledReason: signatureCommandDisabledReason,
              displayLabel: t("pdfSign.upload"),
              icon: "image",
              label: `${t("pdfSign.upload")} — ${t("pdfSign.note")}`,
              onInvoke: (trigger) =>
                void executeVisualSignatureActivation("upload", trigger),
              testId: "pdf-insert-signature-upload",
            },
          ],
          draw: dataItemId
            ? {
                active: commentDrawMode,
                disabled: loading,
                displayLabel: t("pdf.viewer.freeDraw"),
                icon: "draw",
                label: t("pdfInk.modeToggle"),
                onInvoke: () => void executeInkActivation(),
                testId: "pdf-insert-draw",
              }
            : undefined,
        }}
        comments={{
          // WPS CT_Insertion Insert_TextComment is its typewriter text
          // comment, the one the Comment tab arms (not a sticky note).
          textComment: onInsertText
            ? {
                active: armedTextOrigin === "annotate",
                disabled: loading,
                icon: "comment",
                label: t("pdfComment.textComment"),
                onInvoke: () => void executeFreeTextActivation("annotate"),
                testId: "pdf-insert-text-comment",
              }
            : undefined,
          more: {
            disabled: loading,
            icon: "moreComments",
            label: t("pdf.viewer.moreComments"),
            onInvoke: () => selectRibbonTab("annotate"),
            testId: "pdf-insert-more-comments",
          },
          moreItems: dataItemId
            ? [
                {
                  active: shapeMode,
                  disabled: loading,
                  icon: "shape" as const,
                  label: t("pdfShape.modeToggle"),
                  onInvoke: () => void executeShapeActivation(),
                  testId: "pdf-insert-more-shape",
                },
              ]
            : [],
        }}
        groupLabels={{
          comments: t("toolbar.ribbonInsertComment"),
          content: t("pdf.viewer.insertObjectGroup"),
          decoration: t("pdfStamp.title"),
          pages: t("pdf.viewer.pages"),
          pageFiles: pdfHomeDisplayLabels.splitMergeGroup,
          pageNavigation: t("toolbar.ribbonInsertPages"),
          pageRotate: t("toolbar.ribbonPageView"),
        }}
      />
      <PdfEditRibbonTab
        pointerTools={
          <PdfEditPointerTools
            groupLabel={t("pdf.pointerTools")}
            handActive={panMode}
            handActiveLabel={t("pdfViewer.handToolActive")}
            handDisplayLabel={pdfHomeDisplayLabels.hand}
            handLabel={t("pdfViewer.handTool")}
            loading={loading}
            loadingLabel={t("pdfViewer.loading")}
            onHandActivate={handleHandToolToggle}
            onSelectActivate={handleSelectTool}
            selectActive={selectToolActive}
            selectLabel={t("toolbar.select")}
          />
        }
        commands={{
          // WPS: 编辑内容 is the mode toggle, pressed while the mode is on.
          editContent: {
            active: contentEditModeActive,
            disabled:
              loading ||
              !onEditContent ||
              Boolean(contentEditModeDisabledReason),
            disabledReason:
              contentEditModeDisabledReason ??
              (onEditContent
                ? undefined
                : t("pdf.viewer.editingUnavailableDocument")),
            icon: "editContent",
            label: t("pdf.editContent"),
            onInvoke: onEditContent ? () => onToggleContentEditMode() : undefined,
            testId: "pdf-edit-content",
          },
          insertText: {
            active: armedTextOrigin === "edit",
            disabled: loading || !onInsertText,
            disabledReason: onInsertText
              ? undefined
              : t("pdf.viewer.textInsertionUnavailable"),
            icon: "insertText",
            label: t("pdf.editInsertText"),
            onInvoke: onInsertText
              ? () => void executeFreeTextActivation("edit")
              : undefined,
            testId: "pdf-edit-insert-text",
          },
          insertImage: {
            disabled:
              loading || !dataItemId || !onManagedStampUpsert,
            disabledReason:
              dataItemId && onManagedStampUpsert
                ? undefined
                : t("pdf.viewer.imageInsertionUnavailable"),
            icon: "insertImage",
            label: t("pdf.insertImage"),
            onInvoke: (trigger) =>
              void executeImageActivation(trigger ?? null),
            testId: "pdf-edit-insert-image",
          },
          links: {
            active: linkMode,
            disabled: loading || linkSaving || !dataItemId,
            disabledReason: !dataItemId
              ? t("pdf.viewer.documentNoWritablePdfData")
              : linkSaving
                ? t("pdf.viewer.linkBeingSaved")
                : loading
                  ? t("pdfViewer.loading")
                  : undefined,
            icon: "link",
            label: t("pdfViewer.addLink"),
            onInvoke: (trigger) => {
              if (linkMode) toggleLinkMode(trigger);
              else void executeLinkActivation("url", trigger);
            },
            testId: "pdf-edit-links",
          },
          scanEdit: {
            ...recognizeTextCommand,
            hidden: true,
            icon: "scanEdit",
            label: t("pdf.edit.scanEdit"),
            displayLabel: t("pdf.edit.scanEdit"),
            testId: "pdf-edit-scan-edit",
          },
          stamp: dataItemId
            ? {
                active: stampPickerOpen || armedStamp != null,
                disabled: loading || !onManagedStampUpsert,
                disabledReason: onManagedStampUpsert
                  ? undefined
                  : t("document.capability.unavailable"),
                icon: "stamp",
                label: t("pdfStampTool.modeToggle"),
                onInvoke: (trigger) => void executeStampActivation(trigger ?? null),
                testId: "pdf-edit-stamp",
              }
            : undefined,
          linkItems: [
            {
              active:
                linkMode && linkPlacementDestinationKindRef.current === "url",
              disabled: loading || linkSaving || !dataItemId,
              disabledReason: !dataItemId
                ? t("pdf.viewer.documentNoWritablePdfData")
                : linkSaving
                  ? t("pdf.viewer.linkBeingSaved")
                  : undefined,
              icon: "link",
              label: t("pdfViewer.linkUrlLabel"),
              onInvoke: (trigger) => armLinkMode("url", trigger),
              testId: "pdf-edit-links-external",
            },
            {
              active:
                linkMode &&
                linkPlacementDestinationKindRef.current === "page",
              disabled: loading || linkSaving || !dataItemId,
              disabledReason: !dataItemId
                ? t("pdf.viewer.documentNoWritablePdfData")
                : linkSaving
                  ? t("pdf.viewer.linkBeingSaved")
                  : undefined,
              icon: "link",
              label: t("pdfViewer.linkGoToPageLabel"),
              onInvoke: (trigger) => armLinkMode("page", trigger),
              testId: "pdf-edit-links-internal",
            },
            ...(deleteAllLinksAvailable
              ? [
                  {
                    disabled:
                      loading || linkSaving || documentLinkCount === 0,
                    disabledReason:
                      documentLinkCount === 0
                        ? t("pdf.viewer.documentHasNoLinks")
                        : linkSaving
                          ? t("pdf.viewer.linkBeingSaved")
                          : undefined,
                    icon: "remove" as const,
                    label: t("pdf.viewer.deleteAllLinks"),
                    onInvoke: () => {
                      void deleteAllLinks();
                    },
                    testId: "pdf-edit-links-delete-all",
                  },
                ]
              : []),
          ],
          erase: {
            active: eraseRibbon?.armed ?? false,
            brushWidth:
              eraseRibbon?.brushWidth ?? PDF_ERASE_DEFAULT_BRUSH_WIDTH,
            brushWidths: PDF_ERASE_BRUSH_WIDTHS,
            color: eraseRibbon?.color ?? null,
            controlId: "PDFCTL.RIBBON.EDIT.ERASE.PRIMARY",
            menuControlId: "PDFCTL.RIBBON.EDIT.ERASE.MENU",
            disabled:
              loading || !eraseRibbon || Boolean(eraseRibbon.disabledReason),
            disabledReason: !eraseRibbon
              ? t("pdf.erase.unavailableHost")
              : eraseRibbon.disabledReason,
            displayLabel: t("pdf.erase.button"),
            icon: "erase",
            // The name says what it is: a cover, not a redaction.
            label: t("pdf.erase.tooltip"),
            labels: {
              menu: t("pdf.erase.menu"),
              automatic: t("pdf.erase.automatic"),
              themeColors: t("pdf.annotationFormat.themeColors"),
              standardColors: t("pdf.annotationFormat.standardColors"),
              recentColors: t("pdf.annotationFormat.recentColors"),
              moreColors: t("pdf.annotationFormat.moreColors"),
              shape: t("pdf.erase.shape"),
              rectangle: t("pdf.erase.rectangle"),
              brush: t("pdf.erase.brush"),
              brushWidth: t("pdf.erase.brushWidth"),
              brushWidthValue: (width) =>
                t("pdf.erase.brushWidthValue", { width }),
              coverNote: t("pdf.erase.coverNote"),
            },
            onBrushWidthChange: (width) =>
              eraseRibbon?.onBrushWidthChange(width),
            onColorChange: (color) => eraseRibbon?.onColorChange(color),
            onInvoke: eraseRibbon ? () => eraseRibbon.onToggle() : undefined,
            onShapeChange: (shape) => eraseRibbon?.onShapeChange(shape),
            shape: eraseRibbon?.shape ?? "rect",
            testId: "pdf-edit-erase",
          },
          shapes: editShapesRibbonCommand("pdf-edit-shapes"),
          shapeItems: editShapeRibbonItems("pdf-edit-shapes"),
          splitPage: {
            active: splitPageActive,
            controlId: "PDFCTL.RIBBON.EDIT.SPLIT_PAGE",
            disabled:
              loading ||
              !onSplitPage ||
              (!splitPageActive && Boolean(splitPageDisabledReason)),
            disabledReason:
              splitPageActive
                ? undefined
                : splitPageDisabledReason ??
                  (onSplitPage ? undefined : t("pdf.splitPage.unavailableHost")),
            icon: "splitPage",
            label: t("pdf.splitPage.button"),
            onInvoke: onSplitPage ? () => onSplitPage() : undefined,
            testId: "pdf-edit-split-page",
          },
          cropPage: {
            disabled: loading || !onCropPage,
            disabledReason: onCropPage
              ? undefined
              : t("pdf.viewer.pageCropUnavailableHost"),
            icon: "cropPage",
            label: t("pdf.crop.button"),
            onInvoke: onCropPage
              ? (trigger, resolveTrigger) =>
                  onCropPage(trigger, resolveTrigger)
              : undefined,
            popup: "dialog",
            testId: "pdf-edit-crop-page",
            triggerRef: editCropTriggerRef,
          },
          watermark: {
            active: decorationDialogKind === "watermark",
            disabled: loading || flattening || !dataItemId,
            disabledReason: !dataItemId
              ? t("document.capability.unavailable")
              : flattening
                ? t("pdf.viewer.pdfOutputProgress")
                : undefined,
            icon: "watermark",
            label: t("toolbar.watermark"),
            onInvoke: (trigger) =>
              openDecorationDialog("watermark", trigger),
            popup: "dialog",
            testId: "pdf-edit-watermark",
          },
          watermarkItems: [
            {
              disabled: loading || flattening || !dataItemId,
              disabledReason: !dataItemId
                ? t("document.capability.unavailable")
                : flattening
                  ? t("pdf.viewer.pdfOutputProgress")
                  : undefined,
              icon: "watermark",
              label: t("pdf.viewer.addEditWatermark"),
              onInvoke: (trigger) =>
                openDecorationDialog("watermark", trigger),
              popup: "dialog",
              testId: "pdf-edit-watermark-configure",
            },
            {
              disabled:
                loading ||
                flattening ||
                !dataItemId ||
                !effectiveDocumentDecorations.watermark,
              disabledReason: !dataItemId
                ? t("document.capability.unavailable")
                : !effectiveDocumentDecorations.watermark
                  ? t("pdf.viewer.documentNoWatermark")
                  : flattening
                    ? t("pdf.viewer.pdfOutputProgress")
                    : undefined,
              icon: "remove",
              label: t("pdf.viewer.removeWatermark"),
              onInvoke: () => void removeDocumentDecoration("watermark"),
              testId: "pdf-edit-watermark-remove",
            },
          ],
          editFormField: {
            active: existingTextFieldEditMode,
            disabled:
              loading ||
              formMutationBlocked ||
              existingTextFieldEditBusy ||
              existingFormFieldDeleteBusy ||
              currentDetectedFieldSourceIdentityRef.current === null ||
              !hasVisibleFormFieldForPropertyEdit,
            disabledReason: formMutationBlocked
              ? formMutationBlockedReason
              : currentDetectedFieldSourceIdentityRef.current === null
                ? t("document.capability.unavailable")
                : !hasVisibleFormFieldForPropertyEdit
                  ? t("pdf.viewer.pdfNoVisibleFormField")
                  : undefined,
            icon: "formFields",
            label: t("pdf.viewer.editFieldProperties"),
            onInvoke: (trigger) =>
              handleToggleExistingTextFieldEditMode(trigger),
            testId: "pdf-edit-form-field",
          },
          deleteFormField: {
            active: existingFormFieldDeleteMode,
            disabled:
              loading ||
              formMutationBlocked ||
              existingTextFieldEditBusy ||
              existingFormFieldDeleteBusy ||
              currentDetectedFieldSourceIdentityRef.current === null ||
              !hasVisibleFormFieldForDelete,
            disabledReason: formMutationBlocked
              ? formMutationBlockedReason
              : currentDetectedFieldSourceIdentityRef.current === null
                ? t("document.capability.unavailable")
                : !hasVisibleFormFieldForDelete
                  ? t("pdf.viewer.pdfNoVisibleFormFields")
                  : undefined,
            icon: "deleteField",
            label: t("pdf.viewer.deleteFormField"),
            onInvoke: (trigger) =>
              handleToggleExistingFormFieldDeleteMode(trigger),
            testId: "pdf-edit-delete-form-field",
          },
          // Merge and split are not grey while the file shown is inspected
          // again after a save: the press waits for it, busy.
          mergeDocument: {
            busy: mergeBusy,
            disabled:
              loading ||
              mergeBusy,
            icon: "mergeDocument",
            label: t("pdf.editMergeDocument"),
            onInvoke: () => {
              handleOpenMergePicker();
            },
            testId: "pdf-edit-merge-document",
          },
          splitDocument: {
            busy: splitBusy,
            disabled:
              loading ||
              splitBusy,
            icon: "splitDocument",
            label: t("pdf.editSplitDocument"),
            onInvoke: (trigger) => {
              splitDialogReturnFocusRef.current = trigger;
              setExportImageOpen(false);
              setSplitOpen(true);
            },
            testId: "pdf-edit-split-document",
          },
        }}
        groupLabels={{
          document: pdfHomeDisplayLabels.splitMergeGroup,
          editContent: t("pdf.editContent"),
          forms: t("pdf.viewer.formFields"),
          page: t("pdf.viewer.editPageGroup"),
          shapes: t("toolbar.ribbonDraw"),
          watermark: t("pdfStamp.watermarkText"),
        }}
      />
      <PdfCommentRibbonTab
        // WPS's newer 批注 tab shows 图章 among the drawing tools and 测量 in
        // a group of its own (CT_Postil TMpxa3, MeasureTool); they leave More.
        extraTools={{
          label: t("pdfFeatured.moreGroup"),
          commands: globalCommentToolsCommands.filter(
            (command) => !COMMENT_TOOLS_ON_SHOW.has(command.testId),
          ),
        }}
        measure={(() => {
          const command = globalCommentToolsCommands.find(
            (entry) => entry.testId === "pdf-comment-measure",
          );
          return command ? { label: t("pdfViewer.measureTool"), command } : undefined;
        })()}
        // 涂黑: the mark toggle, the pending review and 应用 lived only in the
        
        redaction={(
          <>
            {dataItemId ? (
              <PdfRibbonButton
                active={redactMode}
                onClick={toggleRedactionMarkMode}
                disabled={loading}
                label={t("pdfRedact.modeToggle")}
                testId="pdf-redact-mode-toggle"
              >
                <SquareDashed size={20} />
              </PdfRibbonButton>
            ) : null}
            {

}
            <PdfRibbonButton
              active={docSearchOpen}
              onClick={() => toggleNavigationPanel("search")}
              disabled={loading}
              label={t("pdfViewer.findBarLabel")}
              testId="pdf-redact-search"
            >
              <Search size={20} />
            </PdfRibbonButton>
            {dataItemId ? (
              <PdfRibbonButton
                onClick={openRedactPages}
                disabled={loading || redactSaving}
                label={t("pdfRedact.pages")}
                testId="pdf-redact-pages"
              >
                <FileStack size={20} />
              </PdfRibbonButton>
            ) : null}
            {dataItemId && openHiddenInformation ? (
              <PdfRibbonButton
                onClick={openHiddenInformation}
                disabled={loading}
                label={t("pdfHidden.title")}
                testId="pdf-remove-hidden-information"
              >
                <ShieldCheck size={20} />
              </PdfRibbonButton>
            ) : null}

            {dataItemId && (redactMode || redactions.length > 0) ? (
              <div
                className="pdf-viewer__redact-controls"
                data-testid="pdf-redact-controls"
              >
                {redactMode ? (
                  <span
                    className="pdf-viewer__redact-hint"
                    role="status"
                    aria-label={t("pdfRedact.hint")}
                  >
                    {t("pdfRedact.hint")}
                  </span>
                ) : null}
                {redactions.length > 0 ? (
                  <>
                    <span
                      className="pdf-viewer__redaction-pending-status"
                      role="status"
                      aria-live="polite"
                      data-testid="pdf-redaction-pending-status"
                    >
                      {redactionProductCopy.pending}
                    </span>
                    <PdfRedactionPendingReview
                      items={pendingRedactionReviewItems}
                      busy={loading || flattening || redactSaving}
                      copy={{
                        summary: redactionProductCopy.reviewSummary,
                        markLabel: redactionProductCopy.reviewMark,
                        pageLabel: redactionProductCopy.reviewPage,
                        mappedLabel: redactionProductCopy.reviewMapped,
                        copyOnlyLabel: redactionProductCopy.reviewCopyOnly,
                        unmappedLabel: redactionProductCopy.reviewUnmapped,
                        removeLabel: redactionProductCopy.reviewRemove,
                      }}
                      onNavigate={handleRedactionReviewNavigate}
                      onDelete={handleRedactionReviewDelete}
                    />
                    <PdfRibbonButton
                      big={false}
                      className="pdf-viewer__apply-redactions"
                      onClick={() =>
                        void executeReaderCommand(
                          PDF_READER_APPLY_REDACTIONS_COMMAND,
                          applyRedactionsRuntimeDecision,
                        )
                      }
                      disabled={!applyRedactionsCommand.enabled}
                      label={
                        flattening
                          ? redactionProductCopy.applying
                          : redactionProductCopy.apply
                      }
                      title={
                        applyRedactionsCommand.enabled
                          ? redactionProductCopy.apply
                          : (applyRedactionsCommand.reasonLabel ??
                            redactionProductCopy.failure)
                      }
                      testId="pdf-apply-redactions"
                      data-capability-id={
                        PDF_READER_APPLY_REDACTIONS_COMMAND.featureId
                      }
                      data-output-disposition={
                        applyRedactionsCommand.disposition
                      }
                      data-disabled-reason={
                        applyRedactionsCommand.reasonKey ?? undefined
                      }
                    >
                      {flattening ? (
                        <Loader2 size={15} className="animate-spin" />
                      ) : (
                        <ShieldCheck size={15} />
                      )}
                    </PdfRibbonButton>
                  </>
                ) : null}
              </div>
            ) : null}
          </>
        )}
        pointerTools={
          <PdfEditPointerTools
            controlScope="COMMENT"
            groupLabel={t("pdf.pointerTools")}
            handActive={panMode}
            handActiveLabel={t("pdfViewer.handToolActive")}
            handDisplayLabel={pdfHomeDisplayLabels.hand}
            handLabel={t("pdfViewer.handTool")}
            loading={loading}
            loadingLabel={t("pdfViewer.loading")}
            onHandActivate={handleHandToolToggle}
            onSelectActivate={handleSelectTool}
            selectActive={selectToolActive}
            selectLabel={t("toolbar.select")}
            tabs="annotate comment-tool"
            testIdPrefix="pdf-comment"
          />
        }
        drawingContext={
          drawingToolContextActive
            ? {
                disabledReason: drawingStyleDisabledReason,
                fillColor: shapeFillColor,
                fillDisabledReason: t("pdf.viewer.shapeNoInteriorFill"),
                fillLabel: t("pdf.viewer.fill"),
                fillSupported:
                  drawingContextKind === "shape" &&
                  kindSupportsFill(shapeKind),
                kind: drawingContextKind,
                label: commentToolContextLabel,
                lineWidth:
                  drawingContextKind === "ink"
                    ? inkWidth
                    : shapeWidthPoints(shapeWidth),
                ...(drawingContextKind === "ink"
                  ? {}
                  : {
                      lineWidthMin: shapeWidthPoints(PDF_SHAPE_MIN_WIDTH),
                      lineWidthMax: shapeWidthPoints(PDF_SHAPE_MAX_WIDTH),
                    }),
                lineWidthLabel:
                  drawingContextKind === "ink"
                    ? t("pdf.viewer.lineWidth")
                    : `${t("pdf.viewer.lineWidth")} (${t("presentation.pointsShort")})`,
                lineStyle: shapeLineStyle,
                lineStyleLabel: t("pdf.viewer.lineStyle"),
                lineStyleLabels: {
                  solid: t("pdf.annotationFormat.lineStyleSolid"),
                  dashed: t("pdf.annotationFormat.lineStyleDashed"),
                  dotted: t("pdf.annotationFormat.lineStyleDotted"),
                },
                lineEndingsSupported: activeShapeLineKind !== null,
                lineStart: shapeLineStart,
                lineStartLabel: t("pdf.viewer.start"),
                lineEnd: shapeLineEnd,
                lineEndLabel: t("pdf.viewer.end"),
                lineEndingLabels: {
                  none: t("pdf.viewer.lineEndingNone"),
                  openArrow: t("pdf.viewer.lineEndingOpenArrow"),
                  closedArrow: t("pdf.viewer.lineEndingClosedArrow"),
                  circle: t("pdf.viewer.lineEndingCircle"),
                  square: t("pdf.viewer.lineEndingSquare"),
                },
                onFillColorChange: (color) => {
                  if (color === null || isShapeColor(color)) {
                    setShapeFillColor(color);
                    setVisualMaskActive(false);
                  }
                },
                onLineWidthChange: (width) => {
                  if (drawingContextKind === "ink") {
                    setInkWidth(
                      Math.min(
                        PDF_INK_MAX_WIDTH,
                        Math.max(PDF_INK_MIN_WIDTH, width),
                      ),
                    );
                    return;
                  }
                  setShapeWidth(shapeWidthFromPointsOnPage(width));
                  setVisualMaskActive(false);
                },
                onLineStyleChange: (style) => {
                  setShapeLineStyle(style);
                  setVisualMaskActive(false);
                },
                onLineStartChange: (ending) => {
                  if (!activeShapeLineKind) return;
                  setShapeLineStarts((current) => ({
                    ...current,
                    [activeShapeLineKind]: ending,
                  }));
                  setVisualMaskActive(false);
                },
                onLineEndChange: (ending) => {
                  if (!activeShapeLineKind) return;
                  setShapeLineEnds((current) => ({
                    ...current,
                    [activeShapeLineKind]: ending,
                  }));
                  setVisualMaskActive(false);
                },
                onOutlineColorChange: (color) => {
                  if (drawingContextKind === "ink") {
                    if (isInkColor(color)) setInkColor(color);
                    return;
                  }
                  if (isShapeColor(color)) {
                    setShapeColor(color);
                    setVisualMaskActive(false);
                  }
                },
                outlineColor:
                  drawingContextKind === "ink" ? inkColor : shapeColor,
                outlineColorLabel: t("pdf.viewer.outlineColor"),
                ...(drawingContextKind === "ink"
                  ? {
                      lineType: {
                        label: t("pdf.viewer.graffitiLineType"),
                        freehand: {
                          active: !commentInkEraser.active && !commentInkStraight.on,
                          icon: "draw" as const,
                          label: t("pdf.ribbon.wps.freehand"),
                          onInvoke: () => {
                            commentInkStraight.set(false);
                            if (commentInkEraser.active) commentInkEraser.toggle();
                          },
                          testId: "pdf-drawing-ink-freehand",
                        },
                        straight: {
                          active: !commentInkEraser.active && commentInkStraight.on,
                          icon: "line" as const,
                          label: t("presentation.shape.line"),
                          onInvoke: () => {
                            commentInkStraight.set(true);
                            if (commentInkEraser.active) commentInkEraser.toggle();
                          },
                          testId: "pdf-drawing-ink-straight",
                        },
                      },
                      eraser: {
                        active: commentInkEraser.active,
                        disabled: !commentInkEraser.available,
                        disabledReason: commentInkEraser.available
                          ? undefined
                          : t("toolbar.drawNoInk"),
                        icon: "eraser" as const,
                        label: t("toolbar.drawEraser"),
                        onInvoke: commentInkEraser.toggle,
                        testId: "pdf-drawing-ink-eraser",
                      },
                    }
                  : {}),
                styleEnabled: drawingStyleEnabled,
                styleLabel:
                  drawingContextKind === "shape"
                    ? t("pdf.annotationFormat.style")
                    : drawingContextKind === "measurement"
                      ? t("pdfViewer.measureSettings")
                      : t("toolbar.ribbonStyles"),
                ...(drawingContextKind === "measurement"
                  ? {
                      measurement: {
                        kind: measureKind,
                        kindLabel: t("pdfViewer.measureKind"),
                        kindOptions: MEASURE_MODES.map((mode) => ({
                          value: mode,
                          label:
                            mode === "distance"
                              ? t("pdfViewer.measureDistance")
                              : mode === "perimeter"
                                ? t("pdfViewer.measurePerimeter")
                                : t("pdfViewer.measureArea"),
                        })),
                        onKindChange: (kind: string) => {
                          setMeasureKind(asMeasureMode(kind));
                          setMeasurePts(null);
                        },
                        unit: measureUnit,
                        unitLabel: t("pdfViewer.measureUnit"),
                        units: MEASURE_UNITS,
                        onUnitChange: (unit: string) =>
                          setMeasureUnit(asMeasureUnit(unit)),
                        ratio: measureRatio,
                        ratioLabel: t("pdfViewer.measureRatio"),
                        onRatioChange: setMeasureRatio,
                        // The same commands the Comment tab lists, in the
                        // tab the tool works in (user audit PDF 0010).
                        ...drawingMeasurementCommands(globalCommentToolsCommands),
                        snap: {
                          label: t("pdfViewer.measureSnap"),
                          options: (["endpoint", "path", "midpoint", "intersection"] as const).map((kind) => ({
                            kind,
                            label: t(`pdfViewer.measureSnap_${kind}` as "pdfViewer.measureSnap_endpoint"),
                            checked: measureSnap.kinds.has(kind),
                            onToggle: () => measureSnap.toggle(kind),
                          })),
                        },
                        results: {
                          label: t("pdfViewer.measureResults"),
                          emptyLabel: t("pdfViewer.measureResultsEmpty"),
                          items: measureResults.items.map((item, index) => ({
                            id: item.id,
                            text: `${index + 1}. ${
                              item.mode === "area"
                                ? t("pdfViewer.measureArea")
                                : item.mode === "perimeter"
                                  ? t("pdfViewer.measurePerimeter")
                                  : t("pdfViewer.measureDistance")
                            } · ${item.label} · ${t("pdf.excerptPage.pageHeading", {
                              page: item.page,
                            })}`,
                            page: item.page,
                          })),
                          onOpen: measureResults.open,
                        },
                      },
                    }
                  : {}),
              }
            : undefined
        }
        markup={{
          highlight: {
            active: pendingRibbonMarkupStyle === "highlight",
            disabled: commentMarkupDisabled,
            disabledReason: commentMarkupDisabledReason,
            icon: "highlight",
            label: t("pdfSelection.highlight"),
            displayLabel: t("pdf.ribbon.wps.highlightText"),
            onInvoke: dataItemId
              ? () => void executeMarkupActivation("highlight")
              : undefined,
            testId: "pdf-comment-highlight",
          },
          highlightMenu: {
            activeColor: ribbonMarkupColor,
            colorLabel: t("highlightEdit.colorLabel"),
            colors: PDF_HIGHLIGHT_COLOR_OPTIONS,
            onColorChange: handleRibbonMarkupColorChange,
          },
          areaHighlight: {
            active: shapeMode && areaHighlightActive,
            disabled: areaHighlightDisabled,
            disabledReason: areaHighlightDisabledReason,
            icon: "areaHighlight",
            label: t("pdfShape.areaHighlight"),
            onInvoke: dataItemId
              ? () => void executeAreaHighlightActivation()
              : undefined,
            testId: "pdf-comment-area-highlight",
          },
          areaHighlightMenu: {
            activeColor: areaHighlightColor,
            activeOpacity: areaHighlightOpacity,
            colorLabel: t("pdfShape.fillColorLabel"),
            colors: PDF_SHAPE_COLORS,
            onColorChange: handleAreaHighlightColorChange,
            onOpacityChange: handleAreaHighlightOpacityChange,
            opacities: PDF_AREA_HIGHLIGHT_OPACITIES,
            opacityLabel: t("pdfShape.fillOpacityLabel"),
          },
          underline: {
            active: pendingRibbonMarkupStyle === "underline",
            disabled: commentMarkupDisabled,
            disabledReason: commentMarkupDisabledReason,
            icon: "underline",
            label: t("pdfSelection.underline"),
            onInvoke: dataItemId
              ? () => void executeMarkupActivation("underline")
              : undefined,
            testId: "pdf-comment-underline",
          },
          underlineMenu: {
            activeColor: ribbonMarkupColor,
            colorLabel: t("highlightEdit.colorLabel"),
            colors: PDF_HIGHLIGHT_COLOR_OPTIONS,
            onColorChange: handleRibbonMarkupColorChange,
          },
          strikeout: {
            active: pendingRibbonMarkupStyle === "strikethrough",
            disabled: commentMarkupDisabled,
            disabledReason: commentMarkupDisabledReason,
            icon: "strikeout",
            label: t("pdfSelection.strikethrough"),
            onInvoke: dataItemId
              ? () => void executeMarkupActivation("strikethrough")
              : undefined,
            testId: "pdf-comment-strikeout",
          },
          strikeoutMenu: {
            activeColor: ribbonMarkupColor,
            colorLabel: t("highlightEdit.colorLabel"),
            colors: PDF_HIGHLIGHT_COLOR_OPTIONS,
            onColorChange: handleRibbonMarkupColorChange,
          },
          replace: {
            active: pendingProofingKind === "replace",
            disabled:
              loading ||
              !dataItemId ||
              !selectionEnabled ||
              !annotationContextWritable ||
              !onManagedProofingUpsert ||
              proofingBusy,
            disabledReason:
              dataItemId &&
              selectionEnabled &&
              annotationContextWritable &&
              onManagedProofingUpsert
                ? undefined
                : t("pdfComment.proofingUnavailable"),
            icon: "replaceMark",
            label: t("pdfComment.replaceMark"),
            onInvoke: () => void executeProofingActivation("replace"),
            testId: "pdf-comment-replace-mark",
          },
          replaceMenu: {
            activeColor: ribbonMarkupColor,
            colorLabel: t("highlightEdit.colorLabel"),
            colors: PDF_HIGHLIGHT_COLOR_OPTIONS,
            onColorChange: handleRibbonMarkupColorChange,
          },
          insert: {
            active: pendingProofingKind === "insert",
            disabled:
              loading ||
              !dataItemId ||
              !annotationContextWritable ||
              !onManagedProofingUpsert ||
              proofingBusy,
            disabledReason:
              dataItemId &&
              annotationContextWritable &&
              onManagedProofingUpsert
                ? undefined
                : t("pdfComment.proofingUnavailable"),
            icon: "insertCaret",
            label: t("pdfComment.insertCaret"),
            onInvoke: () => void executeProofingActivation("insert"),
            testId: "pdf-comment-insert-caret",
          },
          insertMenu: {
            activeColor: ribbonMarkupColor,
            colorLabel: t("highlightEdit.colorLabel"),
            colors: PDF_HIGHLIGHT_COLOR_OPTIONS,
            onColorChange: handleRibbonMarkupColorChange,
          },
        }}
        drawing={{
          stamp: globalCommentToolsCommands.find(
            (command) => command.testId === "pdf-comment-stamp",
          ),
          draw: {
            active: commentDrawMode,
            disabled: loading || !dataItemId,
            icon: "draw",
            label: t("pdfInk.modeToggle"),
            displayLabel: t("pdf.ribbon.wps.freehand"),
            onInvoke: dataItemId
              ? () => void executeInkActivation()
              : undefined,
            testId: "pdf-comment-draw",
          },
          drawMenu: {
            activeColor: inkColor,
            activeWidth: inkWidth,
            colorLabel: t("pdfInk.colorLabel"),
            colors: PDF_INK_COLORS,
            onColorChange: (color) => {
              if (isInkColor(color)) setInkColor(color);
            },
            onWidthChange: setInkWidth,
            undo: {
              disabled: inks.length === 0 || inkSaving,
              icon: "delete",
              label: t("pdfInk.undoLast"),
              onInvoke: handleInkUndoLast,
              testId: "pdf-comment-undo-ink",
            },
            widthLabel: t("pdfInk.widthLabel"),
            widths: Array.from(
              new Set([
                PDF_INK_MIN_WIDTH,
                DEFAULT_INK_WIDTH,
                Math.min(PDF_INK_MAX_WIDTH, 2),
                PDF_INK_MAX_WIDTH,
              ]),
            ),
          },
          textComment: {
            active: armedTextOrigin === "annotate",
            disabled: loading || !onInsertText,
            disabledReason: onInsertText
              ? undefined
              : t("pdf.editNotReady"),
            icon: "textComment",
            label: t("pdfComment.textComment"),
            onInvoke: onInsertText
              ? () => void executeFreeTextActivation("annotate")
              : undefined,
            testId: "pdf-comment-text-comment",
          },
          note: {
            active: noteMode,
            disabled: loading || !dataItemId,
            icon: "note",
            label: t("pdfNote.modeToggle"),
            displayLabel: t("pdf.ribbon.wps.noteAnnotation"),
            onInvoke: dataItemId
              ? () => {
                  setPendingRibbonMarkupStyle(null);
                  void executeNoteActivation();
                }
              : undefined,
            testId: "pdf-comment-note",
          },
          noteMenu: {
            activeColor: noteColor,
            colorLabel: t("pdfNote.colorLabel"),
            colors: PDF_NOTE_COLORS,
            onColorChange: handleNoteColorChange,
          },
          shape: {
            active: shapeMode && !areaHighlightActive,
            disabled: loading || !dataItemId,
            icon: "shape",
            label: t("pdfShape.modeToggle"),
            displayLabel: t("pdf.ribbon.wps.shapeComment"),
            onInvoke: dataItemId
              ? () => void executeShapeActivation()
              : undefined,
            testId: "pdf-comment-shape",
          },
          shapeMenu: {
            activeColor: shapeColor,
            activeWidth: shapeWidthPoints(shapeWidth),
            colorLabel: t("pdfShape.colorLabel"),
            colors: PDF_SHAPE_COLORS,
            kinds: (
              [
                ["line", "line", t("pdfShape.kindLine")],
                ["arrow", "arrow", t("pdfShape.kindArrow")],
                ["rect", "rectangle", t("pdfShape.kindRect")],
                ["ellipse", "ellipse", t("pdfShape.kindEllipse")],
                ["cloud", "shape", t("pdfShape.kindCloud")],
                ["polygon", "shape", t("pdfShape.kindPolygon")],
              ] as const
            ).map(([kind, icon, label]) => ({
              active: shapeKind === kind && shapeMode,
              disabled: loading || !dataItemId,
              icon,
              label,
              onInvoke: dataItemId
                ? () => {
                    setPendingRibbonMarkupStyle(null);
                    disarmPageTools("shape");
                    shapeToolPageContentRef.current = false;
                    setShapeMode(true);
                    setShapeKind(kind);
                    setAreaHighlightActive(false);
                    setVisualMaskActive(false);
                  }
                : undefined,
              testId: `pdf-comment-shape-${kind}`,
            })),
            onColorChange: (color) => {
              if (isSafeHexColor(color)) {
                setShapeColor(color);
                setVisualMaskActive(false);
              }
            },
            onWidthChange: (points) => {
              setShapeWidth(shapeWidthFromPointsOnPage(points));
              setVisualMaskActive(false);
            },
            widthLabel: t("pdfShape.widthLabel"),
            widthUnit: t("presentation.pointsShort"),
            widths: shapeMenuWidthPoints,
          },
        }}
        comments={{
          commentMode: {
            active: commentModeActive,
            disabled: loading,
            icon: "commentMode",
            label: t("pdfComment.commentMode"),
            onInvoke: handleToggleCommentMode,
            testId: "pdf-comment-mode",
          },
          hide: {
            active: commentsHidden,
            disabled: loading,
            icon: "hideComments",
            label: t("pdfComment.hideComments"),
            onInvoke: handleToggleCommentsHidden,
            testId: "pdf-comment-hide",
          },
          manage: {
            active: annotationsOpen,
            disabled: loading,
            icon: "manageComments",
            label: t("pdfComment.manageComments"),
            onInvoke: () => toggleNavigationPanel("annotations"),
            testId: "pdf-comment-manage",
          },
        }}
        reading={{
          readAloud: {
            active: speechStatus === "speaking",
            disabled: loading || !speechSupported,
            disabledReason: speechSupported
              ? undefined
              : t("document.capability.unavailable"),
            icon: speechStatus === "speaking" ? "pause" : "readAloud",
            label:
              speechStatus === "speaking"
                ? t("pdfViewer.readAloudPause")
                : speechStatus === "paused"
                  ? t("pdfViewer.readAloudResume")
                  : t("pdfViewer.readAloud"),
            onInvoke: speechSupported ? handleToggleReadAloud : undefined,
            testId: "pdf-comment-read-aloud",
          },
        }}
        groupLabels={{
          comments: t("pdfComment.commentsGroup"),
          redaction: t("pdfRedact.groupLabel"),
          drawing: t("pdfComment.drawingGroup"),
          markup: t("pdfComment.markupGroup"),
          reading: t("pdfComment.readingGroup"),
        }}
        status={
          xfdfImportNotice
            ? xfdfImportNotice
            : ribbonMarkupError
            ? ribbonMarkupError
            : proofingError
              ? proofingError
            : pendingProofingKind
              ? t("pdfComment.markupArmed", {
                  tool:
                    pendingProofingKind === "replace"
                      ? t("pdfComment.replaceMark")
                      : t("pdfComment.insertCaret"),
                })
            : pendingRibbonMarkupStyle
              ? t("pdfComment.markupArmed", {
                  tool:
                    pendingRibbonMarkupStyle === "highlight"
                      ? t("pdfSelection.highlight")
                      : pendingRibbonMarkupStyle === "underline"
                        ? t("pdfSelection.underline")
                        : pendingRibbonMarkupStyle === "strikethrough"
                          ? t("pdfSelection.strikethrough")
                          : t("pdfSelection.squiggly"),
                })
              : undefined
        }
        statusTestId={
          xfdfImportNotice ? "pdf-import-xfdf-notice" : undefined
        }
        // 退出编辑 while a tool is armed: WPS's annotation and graffiti scenes
        // end with an exit; ours puts the tools away (the Select tool).
        exitTool={
          noteMode || drawMode || signMode || armedStamp != null || shapeMode ||
          measureMode || attachMode || snapshotMode || linkMode || redactMode ||
          commentInkEraser.active
            ? {
                groupLabel: t("pdfContextRibbon.text.actions"),
                label: t("pdfContextRibbon.text.exit"),
                onClick: handleSelectTool,
              }
            : undefined
        }
      />
      {attachmentError && selectedPlaced?.kind !== "fileAttachment" ? (
        <WindowNotice title={t("windowNotices.title")}>
          <span
            className="pdf-viewer__annotation-clipboard-error"
            role="alert"
            data-testid="pdf-attachment-error"
          >
            {attachmentError}
          </span>
        </WindowNotice>
      ) : null}
      <PdfHomeCompactRibbon
        visible={activeRibbonTab === "home"}
        onConvertMenuOpen={onOfficeConversionPrepare}
        // The view utilities and the workspace actions the retired "more"
        
        // the side-by-side compare beside the reader's own view group, the
        // workspace's own actions (generate a deck, text editing mode,
        // replace / download / export Markdown) in a group of their own.
        
        
        // download / export Markdown, and the side-by-side compare, in a
        // group of their own after the tab's; night reading and line numbers
        // in a view group before it.
        trailing={
          <>
            <HomeGroup label={t("toolbar.ribbonView")} testId="pdf-home-view-utilities">
              <OfficeHomeCommandStack>
                <span
                  className="contents"
                  role="group"
                  aria-label={t("toolbar.ribbonView")}
                  data-testid="pdf-global-view-tools"
                >
                  <PdfRibbonButton
                    active={isReadModeActive(readMode)}
                    big={false}
                    captioned
                    data-read-mode={readMode}
                    disabled={loading}
                    displayLabel={t(`pdfViewer.readModeShort.${readMode}`)}
                    label={t(`pdfViewer.readMode.${readMode}`)}
                    onClick={cycleReadMode}
                    testId="pdf-read-mode-toggle"
                    toggle
                  >
                    <Moon size={20} aria-hidden />
                  </PdfRibbonButton>
                  <PdfRibbonButton
                    active={lineNumbersVisible}
                    big={false}
                    captioned
                    disabled={loading}
                    label={t("toolbar.lineNumbersEnable")}
                    onClick={toggleLineNumbers}
                    testId="pdf-line-numbers-toggle"
                    toggle
                  >
                    <Hash size={20} aria-hidden />
                  </PdfRibbonButton>
                </span>
              </OfficeHomeCommandStack>
            </HomeGroup>
            <HomeGroup label={t("pdf.viewer.document")} testId="pdf-home-workspace-actions">
            {onRequestPageSummary ? (
              null
            ) : null}
            {ribbonQuickActions}
                <PdfRibbonButton
                  active={compareFile !== null}
                  disabled={loading}
                  label={t("pdfViewer.compareSideBySide")}
                  onClick={() => compareInputRef.current?.click()}
                  testId="pdf-compare-open"
                >
                  <CopyPlus aria-hidden size={24} />
                </PdfRibbonButton>
            {ribbonQuickOverflowOnlyTail}
            {/* The host's tail (the workspace-authority retry) lived in the
                retired "more" menu; the chassis never draws it beside an
                action rail. The window mounts its issues indicator itself. */}
            {ribbonQuickActionTail}
            </HomeGroup>
          </>
        }
        pointerTools={
          <PdfEditPointerTools
            controlScope="HOME"
            groupLabel={t("pdf.pointerTools")}
            groupTestId="pdf-home-pointer-tools"
            handActive={panMode}
            handActiveLabel={t("pdfViewer.handToolActive")}
            handDisplayLabel={pdfHomeDisplayLabels.hand}
            handLabel={t("pdfViewer.handTool")}
            handTestId="pdf-hand-tool-toggle"
            loading={loading}
            loadingLabel={t("pdfViewer.loading")}
            onHandActivate={handleHandToolToggle}
            onSelectActivate={handleSelectTool}
            selectActive={selectToolActive}
            selectLabel={t("toolbar.select")}
            selectTestId="pdf-select-tool"
            tabs="home"
            testIdPrefix="pdf-home"
          />
        }
        groupLabels={{
          pointer: t("pdf.pointerTools"),
          file: t("pdf.viewer.fileTools"),
          view: t("pdf.viewer.readingView"),
          editing: t("pdf.viewer.editTools"),
          security: t("pdf.viewer.documentSecurity"),
          convert: pdfHomeDisplayLabels.convertGroup,
          convertMenuHeading: t("pdfHome.convertMenu.heading"),
          output: pdfHomeDisplayLabels.outputGroup,
          splitMerge: pdfHomeDisplayLabels.splitMergeGroup,
          reading: t("toolbar.ribbonView"),
          pageNavigation: t("toolbar.ribbonInsertPages"),
          pageView: t("toolbar.ribbonPageView"),
          edit: pdfHomeDisplayLabels.editGroup,
          translation: t("toolbar.translate"),
          utilities: t("toolbar.ribbonMenus"),
        }}
        commands={{
          convertPdf: {
            disabled: loading || isCreatingEditableVersion || !onCreateEditableVersion,
            disabledReason: onCreateEditableVersion
              ? undefined
              : (officeProjectionDisabledReason ??
                t("document.capability.unavailable")),
            label: t("references.createEditable"),
            displayLabel: pdfHomeDisplayLabels.convert,
            loading: isCreatingEditableVersion,
            loadingLabel: t("references.createEditableInflight"),
            onInvoke: onCreateEditableVersion
              ? (trigger) => openOfficeProjection("document", trigger)
              : undefined,
            menuTitle: t("pdfHome.convertMenu.wordTitle"),
            menuDetail: t("pdfHome.convertMenu.wordDetail"),
            testId: "pdf-convert-to-editable",
          },
          convertExcel: {
            disabled: loading || isCreatingEditableVersion || !onConvertToSpreadsheet,
            disabledReason: onConvertToSpreadsheet
              ? undefined
              : (officeProjectionDisabledReason ??
                t("document.capability.unavailable")),
            // A host that says why the conversion is closed keeps it listed,
            // greyed with that reason, the way WPS always lists it.
            hidden: !onConvertToSpreadsheet && !officeProjectionDisabledReason,
            label: t("pages.spreadsheet"),
            displayLabel: "Excel",
            loading: isCreatingEditableVersion,
            onInvoke: onConvertToSpreadsheet
              ? (trigger) => openOfficeProjection("spreadsheet", trigger)
              : undefined,
            menuTitle: t("pdfHome.convertMenu.excelTitle"),
            menuDetail: t("pdfHome.convertMenu.excelDetail"),
            testId: "pdf-convert-to-spreadsheet",
          },
          convertPpt: {
            disabled: loading || isCreatingEditableVersion || !onConvertToPresentation,
            disabledReason: onConvertToPresentation
              ? undefined
              : (officeProjectionDisabledReason ??
                t("document.capability.unavailable")),
            hidden:
              !onConvertToPresentation && !officeProjectionDisabledReason,
            label: t("pages.presentation"),
            displayLabel: "PPT",
            loading: isCreatingEditableVersion,
            onInvoke: onConvertToPresentation
              ? (trigger) => openOfficeProjection("presentation", trigger)
              : undefined,
            menuTitle: t("pdfHome.convertMenu.pptTitle"),
            menuDetail: t("pdfHome.convertMenu.pptDetail"),
            testId: "pdf-convert-to-presentation",
          },
          convertImageOnly: {
            disabled:
              loading ||
              imageOnlyExportBusy ||
              rasterOutputAuthorityBlocked ||
              hasEffectivePdfOutputEdits ||
              matchedSourceRiskSnapshot === null ||
              !savePolicyV1,
            disabledReason: sourceAuthorityUnavailable
              ? sourceAuthorityDisabledReason
              : t("pdf.viewer.chooseResolutionImageEncoding"),
            hidden: sourceAuthorityUnavailable,
            label: t("pdfHome.convertMenu.imageOnlyTitle"),
            loading: imageOnlyExportBusy,
            onInvoke: openImageOnlyExport,
            menuTitle: t("pdfHome.convertMenu.imageOnlyTitle"),
            menuDetail: t("pdfHome.convertMenu.imageOnlyDetail"),
            testId: "pdf-convert-to-image-only",
          },
          convertTxt: {
            disabled: loading || textExportBusy || rasterOutputAuthorityBlocked,
            label: t("toolbar.exportFormat.text"),
            displayLabel: "TXT",
            loading: textExportBusy,
            onInvoke: (trigger) => void handleExportDocumentText(trigger),
            menuTitle: t("pdfHome.convertMenu.textTitle"),
            menuDetail: t("pdfHome.convertMenu.textDetail"),
            testId: "pdf-convert-to-text",
          },
          convertCad: {
            disabled:
              loading ||
              cadExportBusy ||
              rasterOutputAuthorityBlocked ||
              hasEffectivePdfOutputEdits ||
              hasNativeFormEditsExcludedFromSource ||
              matchedSourceRiskSnapshot === null ||
              !dataItemId,
            disabledReason: sourceAuthorityUnavailable
              ? sourceAuthorityDisabledReason
              : t("pdfFeatured.cadVectorOnly"),
            hidden: sourceAuthorityUnavailable,
            label: t("pdfHome.convertMenu.cadTitle"),
            loading: cadExportBusy,
            onInvoke: openCadExport,
            menuTitle: t("pdfHome.convertMenu.cadTitle"),
            menuDetail: t("pdfHome.convertMenu.cadDetail"),
            testId: "pdf-convert-to-cad",
          },
          outputImages: {
            disabled: loading || exportImageBusy || rasterOutputAuthorityBlocked,
            disabledReason: resolvePdfCommandDisabledReason(
              loading || exportImageBusy || rasterOutputAuthorityBlocked,
              loading ? loadingDisabledReason : undefined,
              exportImageBusy
                ? t("pdf.viewer.pageImagesBeingPrepared")
                : undefined,
              rasterOutputAuthorityBlocked
                ? outputAuthorityDisabledReason
                : undefined,
            ),
            label: t("pdfViewer.exportImage"),
            displayLabel: pdfHomeDisplayLabels.outputImages,
            loading: exportImageBusy,
            onInvoke: (trigger) => {
              exportImageDialogReturnFocusRef.current = trigger;
              setExportImageOpen(true);
            },
            testId: "pdf-convert-to-image",
          },
          splitMerge: {
            disabled:
              loading ||
              splitBusy ||
              mergeBusy,
            disabledReason: sourceAuthorityUnavailable
              ? sourceAuthorityDisabledReason
              : undefined,
            label: `${t("pdfViewer.splitFile")} / ${t("pdfViewer.mergeFiles")}`,
            displayLabel: pdfHomeDisplayLabels.splitMerge,
            loading: splitBusy || mergeBusy,
            onInvoke: (trigger) => {
              splitDialogReturnFocusRef.current = trigger;
              setExportImageOpen(false);
              setSplitOpen(true);
            },
            testId: "pdf-home-split-merge",
          },
          mergeFiles: {
            disabled:
              loading ||
              mergeBusy ||
              (savePolicyV1 && matchedSourceRiskSnapshot === null),
            disabledReason: sourceAuthorityUnavailable
              ? sourceAuthorityDisabledReason
              : undefined,
            label: t("pdfViewer.mergeFiles"),
            loading: mergeBusy,
            onInvoke: handleOpenMergePicker,
            testId: "pdf-home-merge-files",
          },
          play: {
            active: presentationActive,
            disabled: loading,
            label: presentationActive
              ? t("pdfViewer.presentationExit")
              : t("pdfViewer.presentationMode"),
            displayLabel: pdfHomeDisplayLabels.play,
            onInvoke: handlePresentationToggle,
            testId: "pdf-presentation-toggle",
          },
          actualSize: {
            active: isExactZoomPreset(scale, 100),
            disabled: loading,
            label: t("pdfViewer.actualSize"),
            onInvoke: zoomReset,
            testId: "pdf-home-actual-size",
          },
          fitPage: {
            active: fitMode === "page",
            disabled: loading,
            label: t("pdfViewer.fitPage"),
            onInvoke: fitPage,
            testId: "pdf-home-fit-page",
          },
          fitWidth: {
            active: fitMode === "width",
            disabled: loading,
            label: t("pdfViewer.fitWidth"),
            onInvoke: fitWidth,
            testId: "pdf-home-fit-width",
          },
          rotateLeft: {
            disabled:
              loading ||
              (homeRotateLeftCommand
                ? Boolean(
                    homeRotateLeftCommand.disabled ||
                      homeRotateLeftCommand.busy ||
                      !homeRotateLeftCommand.onInvoke,
                  )
                : !pageOpsEnabled),
            disabledReason:
              homeRotateLeftCommand?.disabledReason ??
              (!pageOpsEnabled && !homeRotateLeftCommand?.onInvoke
                ? t("pdf.viewer.pageRotationUnavailableHost")
                : undefined),
            label: t("pdfPageOps.rotateCcw"),
            loading: homeRotateLeftCommand?.busy,
            onInvoke: homeRotateLeftCommand?.onInvoke
              ? (trigger) => homeRotateLeftCommand.onInvoke?.(trigger)
              : pageOpsEnabled
                ? () => handleRotateCcwPage(currentPage)
                : undefined,
            testId: "pdf-home-rotate-left",
          },
          rotateRight: {
            disabled:
              loading ||
              (homeRotateRightCommand
                ? Boolean(
                    homeRotateRightCommand.disabled ||
                      homeRotateRightCommand.busy ||
                      !homeRotateRightCommand.onInvoke,
                  )
                : !pageOpsEnabled),
            disabledReason:
              homeRotateRightCommand?.disabledReason ??
              (!pageOpsEnabled && !homeRotateRightCommand?.onInvoke
                ? t("pdf.viewer.pageRotationUnavailableHost")
                : undefined),
            label: t("pdfPageOps.rotate"),
            loading: homeRotateRightCommand?.busy,
            onInvoke: homeRotateRightCommand?.onInvoke
              ? (trigger) => homeRotateRightCommand.onInvoke?.(trigger)
              : pageOpsEnabled
                ? () => handleRotatePage(currentPage)
                : undefined,
            testId: "pdf-home-rotate-right",
          },
          pageFirst: {
            disabled: loading || currentVisibleOrdinal <= 1,
            disabledReason: resolvePdfCommandDisabledReason(
              loading || currentVisibleOrdinal <= 1,
              loading ? loadingDisabledReason : undefined,
              currentVisibleOrdinal <= 1
                ? t("pdf.viewer.alreadyFirstPage")
                : undefined,
            ),
            label: t("pdf.viewer.firstPage"),
            onInvoke: () => goToVisibleOrdinal(1),
            testId: "pdf-home-page-first",
          },
          pagePrevious: {
            disabled: loading || previousVisibleOrdinal == null,
            disabledReason: resolvePdfCommandDisabledReason(
              loading || previousVisibleOrdinal == null,
              loading ? loadingDisabledReason : undefined,
              previousVisibleOrdinal == null
                ? t("pdf.viewer.thereNoPreviousPage")
                : undefined,
            ),
            label: t("pdfViewer.prev"),
            onInvoke: () => {
              if (previousVisibleOrdinal != null) {
                goToVisibleOrdinal(previousVisibleOrdinal);
              }
            },
            testId: "pdf-home-page-previous",
          },
          pageNext: {
            disabled: loading || nextVisibleOrdinal == null,
            label: t("pdfViewer.next"),
            onInvoke: () => {
              if (nextVisibleOrdinal != null) {
                goToVisibleOrdinal(nextVisibleOrdinal);
              }
            },
            testId: "pdf-home-page-next",
          },
          pageLast: {
            disabled:
              loading || currentVisibleOrdinal >= visiblePageCount,
            label: t("pdf.viewer.lastPage"),
            onInvoke: () => goToVisibleOrdinal(visiblePageCount),
            testId: "pdf-home-page-last",
          },
          rotateDocument: {
            disabled: rotateDocumentDisabled,
            label: t("pdfViewer.rotateAllCw"),
            displayLabel: pdfHomeDisplayLabels.rotateDocument,
            onInvoke: handleRotateDocumentClockwise,
            testId: "pdf-rotate-all-cw",
          },
          singlePage: {
            active: viewMode === "single",
            disabled: loading,
            label: t("pdfViewer.viewModeSingle"),
            onInvoke: () => setViewMode("single"),
            testId: "pdf-home-single-page",
          },
          doublePage: {
            active: viewMode === "spread",
            disabled: loading,
            label: t("pdfViewer.viewModeSpread"),
            displayLabel: t("pdf.ribbon.wps.doublePage"),
            onInvoke: () => setViewMode("spread"),
            testId: "pdf-home-double-page",
          },
          continuousReading: {
            active: viewMode === "continuous",
            disabled: loading,
            label: t("pdfViewer.viewModeContinuous"),
            displayLabel: t("pdf.ribbon.wps.continuousReading"),
            onInvoke: () => setViewMode("continuous"),
            testId: "pdf-home-continuous-reading",
          },
          // WPS: Home's 编辑内容 opens the Edit tab with the mode on.
          editContent: {
            active: contentEditModeActive,
            disabled:
              loading ||
              !onEditContent ||
              Boolean(contentEditModeDisabledReason),
            disabledReason: resolvePdfCommandDisabledReason(
              loading ||
                !onEditContent ||
                Boolean(contentEditModeDisabledReason),
              contentEditModeDisabledReason,
              loading ? loadingDisabledReason : undefined,
              !onEditContent
                ? t("pdf.viewer.currentHostDoesNotProvide2")
                : undefined,
            ),
            label: t("pdf.editContent"),
            displayLabel: pdfHomeDisplayLabels.editContent,
            onInvoke: onEditContent ? () => onEnterContentEditMode() : undefined,
            testId: "pdf-home-edit-content",
          },
          fullTranslation: {
            disabled:
              loading ||
              rasterOutputAuthorityBlocked ||
              matchedSourceRiskSnapshot === null ||
              pdfOcrServerSourceAuthority === null ||
              !ocrAnnotationHydrationReady ||
              !pageId ||
              translationWorking,
            disabledReason: resolvePdfCommandDisabledReason(
              loading ||
                rasterOutputAuthorityBlocked ||
                matchedSourceRiskSnapshot === null ||
                pdfOcrServerSourceAuthority === null ||
                !ocrAnnotationHydrationReady ||
                !pageId ||
                translationWorking,
              loading ? loadingDisabledReason : undefined,
              rasterOutputAuthorityBlocked
                ? outputAuthorityDisabledReason
                : undefined,
              matchedSourceRiskSnapshot === null
                ? sourceAuthorityDisabledReason
                : undefined,
              pdfOcrServerSourceAuthority === null
                ? t("pdf.viewer.ocrSourceNotBeenVerified")
                : undefined,
              !ocrAnnotationHydrationReady
                ? t("pdf.viewer.acceptedOcrTextStillLoading")
                : undefined,
              !pageId
                ? t("pdf.viewer.currentPageNotReady")
                : undefined,
              translationWorking
                ? t("pdf.viewer.fullDocumentTranslationAlreadyProgress")
                : undefined,
            ),
            label: pdfHomeDisplayLabels.fullTranslation,
            displayLabel: pdfHomeDisplayLabels.fullTranslation,
            loading: translationWorking && translationDialog?.scope === "document",
            onInvoke: (trigger) => {
              if (trigger) openFullTranslation(trigger);
            },
            testId: "pdf-home-full-translation",
          },
          selectionTranslation: {
            disabled: selectionTranslationDisabledReason !== undefined,
            disabledReason: selectionTranslationDisabledReason,
            label: pdfHomeDisplayLabels.selectionTranslation,
            displayLabel: pdfHomeDisplayLabels.selectionTranslation,
            loading: translationWorking && translationDialog?.scope === "selection",
            onInvoke: openSelectionTranslation,
            testId: "pdf-home-selection-translation",
          },
          screenshot: {
            active: snapshotMode,
            disabled: loading || rasterOutputAuthorityBlocked,
            disabledReason: resolvePdfCommandDisabledReason(
              loading || rasterOutputAuthorityBlocked,
              loading ? loadingDisabledReason : undefined,
              rasterOutputAuthorityBlocked
                ? outputAuthorityDisabledReason
                : undefined,
            ),
            label: t("pdfViewer.snapshotTool"),
            displayLabel: pdfHomeDisplayLabels.screenshot,
            onInvoke: () => {
              snapshotIntentRef.current = "copy";
              setSnapshotMode((on) => {
                const next = !on;
                if (next) {
                  disarmPageTools("snapshot");
                } else {
                  snapshotDragRef.current = null;
                  setDrawingSnapshot(null);
                }
                setSnapshotToast(null);
                setSnapshotLatex(null);
                return next;
              });
            },
            testId: "pdf-snapshot-mode-toggle",
          },
          screenshotRect: snapshotShapeCommand("rect", "pdfViewer.snapshotRect"),
          screenshotEllipse: snapshotShapeCommand("ellipse", "pdfViewer.snapshotEllipse"),
          screenshotRounded: snapshotShapeCommand("rounded", "pdfViewer.snapshotRounded"),
          screenshotFreeform: snapshotShapeCommand("freeform", "pdfViewer.snapshotFreeform"),
          screenshotScreen: {
            label: t("pdfViewer.screenSnapshot"),
            onInvoke: () => {
              void captureScreenSnapshot();
            },
            testId: "pdf-home-screenshot-screen",
          },
          snapshotLatex: {
            hidden: true,
            disabled: loading || rasterOutputAuthorityBlocked,
            label: t("pdf.snapshot.latexAction"),
            onInvoke: () => {
              snapshotIntentRef.current = "latex";
              if (process.env.NODE_ENV !== "production") {
                console.debug("[pdf-snapshot] latex armed");
              }
              disarmPageTools("snapshot");
              setSnapshotToast(null);
              setSnapshotLatex(null);
              setSnapshotMode(true);
            },
            testId: "pdf-home-snapshot-latex",
          },
          compress: {
            disabled:
              loading ||
              flattening ||
              rasterOutputAuthorityBlocked ||
              (savePolicyV1 && matchedSourceRiskSnapshot === null) ||
              (savePolicyV1 &&
                compressPreset !== PDF_COMPRESS_LOSSLESS_PRESET &&
                (!lossyCompressionSource || !lossyCompressionWritable)),
            disabledReason: rasterOutputAuthorityBlocked
              ? outputAuthorityDisabledReason
              : savePolicyV1 && matchedSourceRiskSnapshot === null
                ? sourceAuthorityDisabledReason
                : savePolicyV1 &&
                    compressPreset !== PDF_COMPRESS_LOSSLESS_PRESET &&
                    (!lossyCompressionSource || !lossyCompressionWritable)
                  ? t("pdfViewer.compressLossyUnavailable")
                  : undefined,
            label: t("pdfViewer.compress"),
            displayLabel: pdfHomeDisplayLabels.compress,
            loading: flattening,
            onInvoke: openCompressSettings,
            testId: "pdf-compress",
          },
          documentEncrypt: {
            disabled: documentEncryptDisabled,
            disabledReason: documentEncryptDisabledReason,
            label: pdfHomeDisplayLabels.documentEncrypt,
            displayLabel: pdfHomeDisplayLabels.documentEncrypt,
            loading: documentEncryptBusy,
            onInvoke: openDocumentEncrypt,
            testId: "pdf-home-document-encrypt",
          },
          verifySignatures: {
            disabled: !showSignatureProof,
            disabledReason: showSignatureProof
              ? undefined
              : t("pdf.viewer.documentNoSignatures"),
            label: t("pdf.viewer.verifySignatures"),
            onInvoke: requestSignatureProofPanel,
            testId: "pdf-home-verify-signatures",
          },
        }}
        pageNavigation={{
          countLabel: totalPagesLabel,
          label: currentPageDocumentLabel,
          inputLabel: t("pdfViewer.pageOf", {
            current: currentVisibleOrdinal,
            total: visiblePageCount,
          }),
          total: visiblePageCount,
          value: pageInput,
          onCancel: () => setPageInput(String(currentVisibleOrdinal)),
          onValueChange: setPageInput,
          onCommit: handlePageInputCommit,
        }}
        layoutMenuLabel={t("pdfViewer.viewModeSpread")}
      />
      {/* The Home tab's output results (text, DXF, image-only PDF) are
          status chips in the reader's status bar (pdfStatusBar below). */}
      {formFieldContextModel ? (
        <PdfFormFieldContextRibbon
          labels={formFieldContextLabels}
          model={formFieldContextModel}
          onOpenProperties={openSelectedFormFieldProperties}
          onOpenOptions={openSelectedFormFieldOptions}
          onRequiredChange={(required) => {
            void applySelectedFormFieldQuickProperty("required", required);
          }}
          onReadOnlyChange={(readOnly) => {
            void applySelectedFormFieldQuickProperty("readOnly", readOnly);
          }}
          onGeometrySizeChange={(axis, value) => {
            const selected = selectedFormFieldContextRef.current;
            if (!selected) return;
            void handleSelectedFormFieldTransform({
              ...selected.field.percent,
              [axis]: value,
            });
          }}
          onAppearanceStyleChange={(change) => {
            void applySelectedFormWidgetAppearance(change);
          }}
          onExit={clearSelectedFormFieldContext}
        />
      ) : null}

      {linkContextModel ? (
        <PdfLinkContextRibbon
          labels={linkContextLabels}
          model={linkContextModel}
          onOpen={openSelectedLinkContext}
          onEdit={editSelectedLinkContext}
          onDelete={() => {
            void deleteSelectedLinkContext();
          }}
          onExit={clearSelectedLinkContext}
        />
      ) : null}

      <PdfProtectRibbonTab
        visible={activeRibbonTab === "tools" && featuredLegacyToolsOpen}
        ribbon="tools"
        labels={protectLabels}
        commands={{
          encrypt: {
            busy: documentEncryptBusy,
            disabled: documentEncryptDisabled,
            disabledReason: documentEncryptDisabledReason,
            label: protectLabels.encryptCopy,
            onInvoke: openDocumentEncrypt,
            testId: "pdf-protect-encrypt",
          },
          mark: {
            active: redactMode,
            disabled: loading || !dataItemId,
            disabledReason: !dataItemId
              ? t("pdf.viewer.documentNoWritablePdfData")
              : undefined,
            label: protectLabels.markRedactions,
            onInvoke: toggleRedactionMarkMode,
            testId: "pdf-protect-redaction-mark",
          },
          apply: {
            busy: flattening,
            disabled: !applyRedactionsCommand.enabled,
            disabledReason: resolvePdfCommandDisabledReason(
              !applyRedactionsCommand.enabled,
              loading ? loadingDisabledReason : undefined,
              redactions.length === 0
                ? t("pdf.viewer.addRedactionMarkFirst")
                : undefined,
              applyRedactionsCommand.reasonLabel,
              applyRedactionsRuntimeDecision.reasonLabel,
              t("pdf.viewer.redactionWriteAuthorityUnavailablePdf"),
            ),
            label: protectLabels.applyRedactions,
            onInvoke: () =>
              void executeReaderCommand(
                PDF_READER_APPLY_REDACTIONS_COMMAND,
                applyRedactionsRuntimeDecision,
              ),
            testId: "pdf-protect-redaction-apply",
          },
          safeCopy: {
            busy: dangerousActionSafeCopyBusy,
            disabled: !dangerousActionSafeCopyCommand.enabled,
            disabledReason: dangerousActionSafeCopyDisabledReason,
            label: protectLabels.safeCopy,
            onInvoke: (trigger) => {
              if (trigger) {
                void executeDangerousActionSafeCopyReview(trigger);
              }
            },
            testId: "pdf-protect-safe-copy",
          },
          archival: {
            busy: verifiedPdfaExportBusy,
            disabled: !archivalDownloadCommand.enabled,
            disabledReason: archivalDownloadCommand.reasonLabel ?? undefined,
            label: protectLabels.archivalCopy,
            onInvoke: (trigger) => {
              verifiedPdfaExportTriggerRef.current = trigger;
              void executeReaderCommand(
                PDF_READER_DOWNLOAD_ARCHIVAL_COMMAND,
                archivalDownloadRuntimeDecision,
              );
            },
            testId: "pdf-protect-pdfa",
          },
        }}
        status={{
          signature: signatureProofLoading
            ? signatureProofPresentation.loadingHeadline
            : signatureProofPresentation.headline,
          actions: dangerousActionCompactStatusLabel,
          actionRisk: dangerousActionCount > 0,
        }}
      />

      <div
        className="note-editor-office-toolbar__group pdf-viewer__toolbar-group pdf-viewer__toolbar-group--docs"
        data-pdf-ribbon-tab="text-format"
        data-pdf-ribbon-label={t("pdf.viewer.textEditing")}
        data-ribbon-label={t("pdf.viewer.textEditing")}
        data-pdf-native-list-text-ribbon-host="true"
        data-testid="pdf-native-list-text-ribbon-host"
      />

      {ribbon.pictureContextRibbon ?? null}

      {activeRibbonTab === "text-format" ? ribbon.textBoxContextRibbon ?? null : null}

      {activeRibbonTab === "table-format" ? ribbon.tableContextRibbon ?? null : null}

      {activeRibbonTab === "chart-design" || activeRibbonTab === "chart-format" ? ribbon.chartContextRibbon ?? null : null}

      {/* Only on its own contextual tab: a source-vector selection lives on
          the Edit surface, and without this gate its groups were appended
          to the Edit command run. */}
      {/* Format Shape: the selected shape's pane at the reader's right
          edge, open across the tabs while the shape stays selected. */}
      {shapeFormatPaneShown && annotationContextModel && formatPaneDock ? (
        <OfficeTaskPaneDockSlot dock={formatPaneDock} id="format-shape">
          <PdfShapeFormatPane
            labels={effectiveAnnotationContextLabels}
            model={annotationContextModel}
            onStyleChange={annotationStyleChange}
            onSizeChange={annotationSizeChange}
            onRotateBy={annotationRotateBy}
            onClose={() => setShapeFormatPaneOpen(false)}
          />
        </OfficeTaskPaneDockSlot>
      ) : null}
      {annotationContextModel && activeRibbonTab === "annotation-format" ? (
        <PdfAnnotationFormatContextRibbon
          labels={effectiveAnnotationContextLabels}
          model={{ ...annotationContextModel, formatPaneOpen: shapeFormatPaneShown }}
          pointerTools={
            <PdfEditPointerTools
              groupLabel={t("pdf.pointerTools")}
              handActive={panMode}
              handActiveLabel={t("pdfViewer.handToolActive")}
              handDisplayLabel={pdfHomeDisplayLabels.hand}
              handLabel={t("pdfViewer.handTool")}
              loading={loading}
              loadingLabel={t("pdfViewer.loading")}
              onHandActivate={handleHandToolToggle}
              onSelectActivate={handleSelectTool}
              selectActive={selectToolActive}
              selectLabel={t("toolbar.select")}
              tabs="annotation-format"
              testIdPrefix="pdf-shape-edit"
            />
          }
          insertShapes={
            <PdfEditRibbonSplitCommand
              command={{
                ...editShapesRibbonCommand("pdf-shape-edit-insert-shapes"),
                displayLabel: t("pdf.viewer.insertShape"),
                label: t("pdf.viewer.insertShape"),
              }}
              menuItems={editShapeRibbonItems("pdf-shape-edit-insert-shapes")}
              order="02.01"
            />
          }
          onRotateBy={annotationRotateBy}
          onFormatPane={formatPaneDock ? () => setShapeFormatPaneOpen((open) => !open) : undefined}
          onCopy={() => {
            if (currentNativeVectorSelection) nativeVectorClipboardCommands.copy();
            else handleCopyPlacedAnnotation();
          }}
          onPaste={() => {
            if (currentNativeVectorSelection) {
              nativeVectorClipboardCommands.paste();
            } else {
              void handlePastePlacedAnnotation();
            }
          }}
          onDuplicate={() => {
            if (currentNativeVectorSelection) {
              runOrQueueNativeVectorCommand((run) => run.duplicate());
            } else {
              void handleDuplicatePlacedAnnotation();
            }
          }}
          onArrangeShape={(action) => {
            if (currentNativeVectorSelection) {
              runOrQueueNativeVectorCommand((run) => run.arrange(action));
            } else {
              void handleSelectedManagedShapeArrange(action);
            }
          }}
          onAlignShapes={
            currentNativeVectorSelection
              ? (command) => runOrQueueNativeVectorCommand((run) => run.align(command))
              : undefined
          }
          onMergeShapes={
            currentNativeVectorSelection
              ? (operation) => runOrQueueNativeVectorCommand((run) => run.merge(operation))
              : undefined
          }
          onReorder={
            currentNativeVectorSelection
              ? (direction) =>
                  runOrQueueNativeVectorCommand((run) => run.reorder(direction))
              : undefined
          }
          onToggleVertexEdit={
            currentNativeVectorSelection
              ? toggleNativeVectorVertexEditing
              : undefined
          }
          onSelectionPane={currentNativeVectorSelection ? toggleNativeSelectionPane : undefined}
          selectionPaneOpen={nativeSelectionPaneOpen}
          onMaskShape={
            currentNativeVectorSelection
              ? () => void maskCurrentNativeVector()
              : undefined
          }
          onStyleChange={annotationStyleChange}
          onSizeChange={annotationSizeChange}
          onDelete={() => {
            if (currentNativeVectorSelection) {
              runOrQueueNativeVectorCommand((run) => run.remove());
            } else {
              handleDeleteSelectedPlacedAnnotation();
            }
          }}
          onSaveAttachment={handleSaveSelectedAttachment}
          onExit={() => {
            // 退出编辑 ends Edit Content (WPS) when the shape was picked in
            // it; a comment selected elsewhere is only let go.
            if (currentNativeVectorSelection || contentEditModeActive) {
              onExitContentEditMode();
            } else {
              clearPlacedAnnotationSelection();
            }
          }}
          status={
            selectedPlaced?.kind === "fileAttachment" && attachmentError ? (
              <span
                className="pdf-viewer__annotation-clipboard-error"
                role="alert"
                data-testid="pdf-attachment-context-error"
              >
                {attachmentError}
              </span>
            ) : placedAnnotationClipboardError ? (
              <span
                className="pdf-viewer__annotation-clipboard-error"
                role="alert"
                data-testid="pdf-annotation-clipboard-error"
              >
                {t("pages.saveFailed")}
              </span>
            ) : null
          }
        />
      ) : null}

      <PdfFeaturedRibbonTab
        visible={activeRibbonTab === "tools" && !featuredLegacyToolsOpen}
        globalTools={{ label: t("pdfFeatured.moreTools"), commands: globalPdfToolsCommands }}
        pointerTools={
          <PdfEditPointerTools
            controlScope="TOOLS"
            groupLabel={t("pdf.pointerTools")}
            groupTestId="pdf-featured-pointer-tools"
            handActive={panMode}
            handActiveLabel={t("pdfViewer.handToolActive")}
            handDisplayLabel={pdfHomeDisplayLabels.hand}
            handLabel={t("pdfViewer.handTool")}
            handTestId="pdf-featured-hand"
            loading={loading}
            loadingLabel={t("pdfViewer.loading")}
            onHandActivate={handleHandToolToggle}
            onSelectActivate={handleSelectTool}
            selectActive={selectToolActive}
            selectLabel={t("toolbar.select")}
            selectTestId="pdf-featured-select"
            tabs="tools"
            testIdPrefix="pdf-featured"
          />
        }
        status={featuredStatus}
        statusTestId={
          featuredScreenRecordingStatus === null && deskewStatus
            ? "pdf-deskew-status"
            : undefined
        }
        groupLabels={{
          convert: t("pdfFeatured.convertGroup"),
          extract: t("pdfFeatured.extractGroup"),
          service: t("pdfFeatured.serviceGroup"),
        }}
        commands={{
          word: {
            disabled: loading || isCreatingEditableVersion || !onCreateEditableVersion,
            disabledReason: onCreateEditableVersion
              ? undefined
              : (officeProjectionDisabledReason ??
                t("pdfFeatured.serviceUnavailable")),
            hidden: !onCreateEditableVersion && !officeProjectionDisabledReason,
            icon: "convertEditable",
            label: t("pdfFeatured.toWord"),
            onInvoke: onCreateEditableVersion
              ? (trigger) => openOfficeProjection("document", trigger)
              : undefined,
            displayLabel: t("pdfFeatured.caption.toWord"),
            testId: "pdf-featured-to-word",
          },
          excel: {
            disabled: loading || isCreatingEditableVersion || !onConvertToSpreadsheet,
            disabledReason: onConvertToSpreadsheet
              ? undefined
              : (officeProjectionDisabledReason ??
                t("pdfFeatured.serviceUnavailable")),
            hidden: !onConvertToSpreadsheet && !officeProjectionDisabledReason,
            icon: "convertSpreadsheet",
            label: t("pdfFeatured.toExcel"),
            onInvoke: onConvertToSpreadsheet
              ? (trigger) => openOfficeProjection("spreadsheet", trigger)
              : undefined,
            displayLabel: t("pdfFeatured.caption.toExcel"),
            testId: "pdf-featured-to-excel",
          },
          ppt: {
            disabled: loading || isCreatingEditableVersion || !onConvertToPresentation,
            disabledReason: onConvertToPresentation
              ? undefined
              : (officeProjectionDisabledReason ??
                t("pdfFeatured.serviceUnavailable")),
            hidden: !onConvertToPresentation && !officeProjectionDisabledReason,
            icon: "convertPresentation",
            label: t("pdfFeatured.toPpt"),
            onInvoke: onConvertToPresentation
              ? (trigger) => openOfficeProjection("presentation", trigger)
              : undefined,
            displayLabel: t("pdfFeatured.caption.toPpt"),
            testId: "pdf-featured-to-ppt",
          },
          txt: {
            disabled: loading || textExportBusy || rasterOutputAuthorityBlocked,
            icon: "convertText",
            label: t("pdfFeatured.toTxt"),
            onInvoke: (trigger) => void handleExportDocumentText(trigger),
            displayLabel: t("pdfFeatured.caption.toTxt"),
            testId: "pdf-featured-to-txt",
          },
          cad: {
            disabled: loading,
            hidden: sourceAuthorityUnavailable,
            icon: "convertCad",
            label: t("pdfFeatured.toCad"),
            onInvoke: () => void import("@mrnote/core/host/desktop-app.ts").then((module) => module.noteDesktopFeature("PDF to CAD")),
            displayLabel: t("pdfFeatured.caption.toCad"),
            testId: "pdf-featured-to-cad",
          },
          imageOnly: {
            disabled: loading,
            hidden: sourceAuthorityUnavailable,
            icon: "convertImageOnly",
            label: t("pdfFeatured.imageOnlyPdf"),
            onInvoke: () => void import("@mrnote/core/host/desktop-app.ts").then((module) => module.noteDesktopFeature("Image-only PDF")),
            displayLabel: t("pdfFeatured.caption.imageOnly"),
            testId: "pdf-featured-image-only",
          },
          exportImages: {
            disabled: loading || exportImageBusy || rasterOutputAuthorityBlocked,
            disabledReason: resolvePdfCommandDisabledReason(
              loading || exportImageBusy || rasterOutputAuthorityBlocked,
              loading ? loadingDisabledReason : undefined,
              exportImageBusy
                ? t("pdf.viewer.pageImagesBeingPrepared")
                : undefined,
              rasterOutputAuthorityBlocked
                ? outputAuthorityDisabledReason
                : undefined,
            ),
            icon: "convertImage",
            label: t("pdfFeatured.exportImages"),
            onInvoke: (trigger) => {
              exportImageDialogReturnFocusRef.current = trigger;
              setExportImageOpen(true);
            },
            displayLabel: t("pdfFeatured.caption.exportImages"),
            testId: "pdf-featured-export-images",
          },
          imagesToPdf: {
            disabled: loading || featuredImagesToPdfBusy,
            icon: "imagesToPdf",
            label: featuredImagesToPdfBusy
              ? t("pdfFeatured.imagesToPdfBusy")
              : t("pages.createPdfFromImages"),
            onInvoke: (trigger) => {
              featuredImagesToPdfReturnFocusRef.current = trigger;
              featuredImagesToPdfInputRef.current?.click();
            },
            // The busy state names its work instead of the caption.
            displayLabel: featuredImagesToPdfBusy
              ? undefined
              : t("pdfFeatured.caption.imagesToPdf"),
            testId: "pdf-featured-images-to-pdf",
          },
          markdownToPdf: {
            disabled: loading || featuredMarkdownToPdfBusy,
            hidden: !featuredMarkdownToPdfAvailable,
            icon: "markdownToPdf",
            label: featuredMarkdownToPdfBusy
              ? t("pdfFeatured.imagesToPdfBusy")
              : t("pdfFeatured.markdownToPdf"),
            onInvoke: (trigger) => {
              featuredImagesToPdfReturnFocusRef.current = trigger;
              featuredMarkdownToPdfInputRef.current?.click();
            },
            displayLabel: featuredMarkdownToPdfBusy
              ? undefined
              : t("pdfFeatured.caption.markdownToPdf"),
            testId: "pdf-featured-markdown-to-pdf",
          },
          cadToPdf: {
            disabled: loading || featuredCadToPdfBusy,
            hidden: !featuredMarkdownToPdfAvailable,
            icon: "cadToPdf",
            label: featuredCadToPdfBusy
              ? t("pdfFeatured.imagesToPdfBusy")
              : t("pdfFeatured.cadToPdf"),
            onInvoke: (trigger) => {
              featuredImagesToPdfReturnFocusRef.current = trigger;
              featuredCadToPdfInputRef.current?.click();
            },
            displayLabel: featuredCadToPdfBusy
              ? undefined
              : t("pdfFeatured.caption.cadToPdf"),
            testId: "pdf-featured-cad-to-pdf",
          },
          splitFile: {
            disabled:
              loading ||
              splitBusy ||
              (savePolicyV1 && matchedSourceRiskSnapshot === null),
            disabledReason: sourceAuthorityUnavailable
              ? sourceAuthorityDisabledReason
              : undefined,
            hidden: sourceAuthorityUnavailable,
            icon: "splitPages",
            label: t("pdfFeatured.splitFile"),
            onInvoke: (trigger) => {
              splitDialogReturnFocusRef.current = trigger;
              setExportImageOpen(false);
              setSplitOpen(true);
            },
            displayLabel: t("pdfFeatured.caption.splitFile"),
            testId: "pdf-featured-split-file",
          },
          merge: {
            disabled:
              loading ||
              mergeBusy ||
              (savePolicyV1 && matchedSourceRiskSnapshot === null),
            disabledReason: sourceAuthorityUnavailable
              ? sourceAuthorityDisabledReason
              : undefined,
            hidden: sourceAuthorityUnavailable,
            icon: "mergePages",
            label: t("pdfFeatured.mergeFile"),
            onInvoke: handleOpenMergePicker,
            displayLabel: t("pdfFeatured.caption.mergeFile"),
            testId: "pdf-featured-merge",
          },
          extractText: {
            disabled: loading || textExportBusy || rasterOutputAuthorityBlocked,
            disabledReason: resolvePdfCommandDisabledReason(
              loading || textExportBusy || rasterOutputAuthorityBlocked,
              loading ? loadingDisabledReason : undefined,
              textExportBusy
                ? t("pdf.viewer.pdfTextExtractionProgress")
                : undefined,
              rasterOutputAuthorityBlocked
                ? outputAuthorityDisabledReason
                : undefined,
            ),
            icon: "convertText",
            label: t("pdfFeatured.extractText"),
            onInvoke: (trigger) => void handleExportDocumentText(trigger),
            displayLabel: t("pdfFeatured.caption.extractText"),
            testId: "pdf-featured-extract-text",
          },
          extractImages: {
            disabled:
              loading ||
              imageExtractionBusy ||
              rasterOutputAuthorityBlocked ||
              matchedSourceRiskSnapshot === null ||
              !dataItemId,
            disabledReason: resolvePdfCommandDisabledReason(
              loading ||
                imageExtractionBusy ||
                rasterOutputAuthorityBlocked ||
                matchedSourceRiskSnapshot === null ||
                !dataItemId,
              loading ? loadingDisabledReason : undefined,
              imageExtractionBusy
                ? t("pdf.viewer.pdfImageExtractionProgress")
                : undefined,
              rasterOutputAuthorityBlocked
                ? outputAuthorityDisabledReason
                : undefined,
              matchedSourceRiskSnapshot === null
                ? sourceAuthorityDisabledReason
                : undefined,
              !dataItemId
                ? t("pdf.viewer.documentNoAuthenticatedOriginalPdf")
                : undefined,
            ),
            hidden: sourceAuthorityUnavailable,
            icon: "convertImage",
            label: t("pdfFeatured.extractImages"),
            onInvoke: (trigger) =>
              void handleOpenImageExtractionReview(trigger),
            displayLabel: t("pdfFeatured.caption.extractImages"),
            testId: "pdf-featured-extract-images",
          },
          recognizeText: {
            ...recognizeTextCommand,
            hidden: true,
            icon: "ocr",
            displayLabel: t("pdfFeatured.caption.recognizeText"),
            testId: "pdf-featured-recognize-text",
          },
          screenshotText: {
            hidden: true,
            disabled:
              loading ||
              rasterOutputAuthorityBlocked ||
              matchedSourceRiskSnapshot === null ||
              pdfOcrServerSourceAuthority === null ||
              pdfOcrServerSourceAuthority.derived === true ||
              !dataItemId,
            disabledReason: resolvePdfCommandDisabledReason(
              loading ||
                rasterOutputAuthorityBlocked ||
                matchedSourceRiskSnapshot === null ||
                pdfOcrServerSourceAuthority === null ||
                pdfOcrServerSourceAuthority.derived === true ||
                !dataItemId,
              loading ? loadingDisabledReason : undefined,
              rasterOutputAuthorityBlocked ? outputAuthorityDisabledReason : undefined,
              !dataItemId ? t("pdf.viewer.documentNoAuthenticatedOriginalPdf") : undefined,
              pdfOcrServerSourceAuthority?.derived === true
                ? t("pdfFeatured.ocrEditedSource")
                : t("pdfFeatured.ocrSourcePending"),
            ),
            icon: "ocr",
            label: t("pdfFeatured.screenshotText"),
            onInvoke: (trigger) => {
              snapshotTextReturnFocusRef.current = trigger;
              snapshotIntentRef.current = "text";
              disarmPageTools("snapshot");
              setSnapshotToast(null);
              setSnapshotLatex(null);
              setSnapshotMode(true);
            },
            displayLabel: t("pdfFeatured.caption.screenshotText"),
            testId: "pdf-featured-screenshot-text",
          },
          compress: {
            disabled:
              loading ||
              flattening ||
              rasterOutputAuthorityBlocked ||
              (savePolicyV1 && matchedSourceRiskSnapshot === null),
            disabledReason: rasterOutputAuthorityBlocked
              ? outputAuthorityDisabledReason
              : savePolicyV1 && matchedSourceRiskSnapshot === null
                ? sourceAuthorityDisabledReason
                : undefined,
            hidden: sourceAuthorityUnavailable,
            icon: "compress",
            label: t("pdfViewer.compress"),
            onInvoke: openCompressSettings,
            displayLabel: t("pdfFeatured.caption.compress"),
            testId: "pdf-featured-compress",
          },
          translate: {
            hidden: true,
            disabled:
              loading ||
              rasterOutputAuthorityBlocked ||
              matchedSourceRiskSnapshot === null ||
              pdfOcrServerSourceAuthority === null ||
              !ocrAnnotationHydrationReady ||
              !pageId ||
              translationWorking,
            disabledReason: resolvePdfCommandDisabledReason(
              loading ||
                rasterOutputAuthorityBlocked ||
                matchedSourceRiskSnapshot === null ||
                pdfOcrServerSourceAuthority === null ||
                !ocrAnnotationHydrationReady ||
                !pageId ||
                translationWorking,
              loading ? loadingDisabledReason : undefined,
              rasterOutputAuthorityBlocked
                ? outputAuthorityDisabledReason
                : undefined,
              matchedSourceRiskSnapshot === null
                ? sourceAuthorityDisabledReason
                : undefined,
              pdfOcrServerSourceAuthority === null
                ? t("pdf.viewer.ocrSourceNotBeenVerified")
                : undefined,
              !ocrAnnotationHydrationReady
                ? t("pdf.viewer.acceptedOcrTextStillLoading")
                : undefined,
              !pageId
                ? t("pdf.viewer.currentPageNotReady")
                : undefined,
              translationWorking
                ? t("pdf.viewer.fullDocumentTranslationAlreadyProgress")
                : undefined,
            ),
            icon: "translation",
            label: t("pdfFeatured.fullTranslation"),
            onInvoke: (trigger) => {
              if (trigger) openFullTranslation(trigger);
            },
            displayLabel: t("pdfFeatured.caption.translate"),
            testId: "pdf-featured-full-translation",
          },
          readAloud: {
            active: speechStatus === "speaking",
            disabled: loading || !speechSupported,
            disabledReason: speechSupported
              ? undefined
              : t("document.capability.unavailable"),
            icon: speechStatus === "speaking" ? "stop" : "continuousReading",
            label:
              speechStatus === "speaking"
                ? t("pdfViewer.readAloudPause")
                : speechStatus === "paused"
                  ? t("pdfViewer.readAloudResume")
                  : t("pdfViewer.readAloud"),
            displayLabel:
              speechStatus === "speaking" || speechStatus === "paused"
                ? undefined
                : t("pdfFeatured.caption.readAloud"),
            onInvoke: handleToggleReadAloud,
            testId: "pdf-featured-read-aloud",
          },
          stopReadAloud:
            speechStatus === "speaking" || speechStatus === "paused"
              ? {
                  icon: "stop",
                  label: t("pdfViewer.readAloudStop"),
                  onInvoke: handleStopReadAloud,
                  testId: "pdf-featured-read-aloud-stop",
              }
            : undefined,
          screenRecording: {
            disabled:
              !screenRecordingSupported ||
              screenRecordingState === "requesting" ||
              screenRecordingState === "recording" ||
              screenRecordingState === "saving",
            disabledReason: screenRecordingSupported
              ? undefined
              : t("document.capability.unavailable"),
            icon: "screenRecording",
            label: t("pdfFeatured.screenRecord"),
            onInvoke: openScreenRecording,
            displayLabel: t("pdfFeatured.caption.screenRecording"),
            testId: "pdf-featured-screen-recording",
          },
          stopScreenRecording:
            screenRecordingState === "recording" ||
            screenRecordingState === "saving"
              ? {
                  disabled: screenRecordingState === "saving",
                  icon: "stop",
                  label:
                    screenRecordingState === "saving"
                      ? t("pdf.viewer.saving")
                      : t("pdf.viewer.stopRecording"),
                  onInvoke: stopScreenRecording,
                  testId: "pdf-featured-screen-recording-stop",
                }
              : undefined,
        }}
      />

      {/* Eager: its export, split and download dialogs open from Home and
          Featured too, before this tab was ever shown. */}
      {legacyToolsHostRequired ? (
        <PdfSurfaceRibbonGroup eager label={t("toolbar.ribbonMenus")} tab="tools">
        <OfficeHomeCommandStack className={surfaceStyles.toolsLanes} rows={3}>
        {featuredLegacyToolsOpen ? (
          <PdfRibbonToolButton
            icon={<ChevronLeft size={14} />}
            onClick={() => setFeaturedLegacyToolsOpen(false)}
            data-testid="pdf-featured-back"
          >
            {t("pdfFeatured.back")}
          </PdfRibbonToolButton>
        ) : null}
        {/* Page labels (Acrobat "Number Pages" — /PageLabels). An export-time
            choice: define logical-numbering ranges baked into the exported
            PDF's number tree so a viewer shows e.g. roman front matter. */}
        <PdfPageLabelsControl
          annotatedDownloadEnabled={annotatedDownloadEnabled}
          dataItemId={dataItemId}
          nativePageLabelRanges={nativePageLabelRanges}
          onNativePageLabelRangesApply={onNativePageLabelRangesApply}
          t={t}
          pages={pages}
          loading={loading}
          pageLabelsOpen={pageLabelsOpen}
          pageLabelsDialogReturnFocusRef={pageLabelsDialogReturnFocusRef}
          pageLabelRanges={pageLabelRanges}
          pageLabelReplacementStarted={pageLabelReplacementStarted}
          setPageLabelReplacementStarted={setPageLabelReplacementStarted}
          pageLabelsApplying={pageLabelsApplying}
          pageLabelsError={pageLabelsError}
          nativePageLabelsControlled={nativePageLabelsControlled}
          nativePageLabelsSourceOwned={nativePageLabelsSourceOwned}
          updatePageLabelRanges={updatePageLabelRanges}
          flattening={flattening}
          closePageLabelsDialog={closePageLabelsDialog}
          togglePageLabelsDialog={togglePageLabelsDialog}
          applyPageLabelRanges={applyPageLabelRanges}
        />

        {/* Print: always available once loaded (prints the flattened snapshot
            — annotations / stamps / links included — via a hidden iframe). */}
        {!loading ? (
          <PdfRibbonToolButton
            icon={printing ? (
              <Loader2 size={14} className="animate-spin" />
            ) : (
              <Printer size={14} />
            )}
            ref={printPreflightTriggerRef}
            data-testid="pdf-print"
            onClick={() =>
              void executeReaderCommand(
                PDF_READER_PRINT_COMMAND,
                printRuntimeDecision,
              )
            }
            disabled={!printCommand.enabled}
            title={
              printCommand.enabled
                ? t("pdfViewer.print")
                : (printCommandDisabledReason ?? t("pdfViewer.print"))
            }
            aria-label={
              printCommand.enabled || !printCommandDisabledReason
                ? t("pdfViewer.print")
                : `${t("pdfViewer.print")} — ${printCommandDisabledReason}`
            }
            data-capability-id={PDF_READER_PRINT_COMMAND.featureId}
            data-output-disposition={printCommand.disposition}
            data-disabled-reason={printCommand.reasonKey ?? undefined}
          >
            {printing
                  ? t("pdfViewer.printInflight")
                  : t("pdfViewer.print")}
          </PdfRibbonToolButton>
        ) : null}

        {/* Render only once the document finished loading: a visible-but-
            disabled flatten button silently swallows real clicks for the
            whole (book-length) load sweep and reads as "broken". */}
          {annotatedDownloadEnabled &&
          hasFlattenableAnnotations &&
          !loading ? (
          <PdfRibbonToolButton
            icon={flattening ? (
              <Loader2 size={14} className="animate-spin" />
            ) : (
              <FileDown size={14} />
            )}
            data-testid="pdf-download-annotated"
            onClick={() =>
              void executeReaderCommand(
                PDF_READER_DOWNLOAD_ANNOTATED_COMMAND,
                annotatedDownloadRuntimeDecision,
              )
            }
            disabled={!annotatedDownloadCommand.enabled}
            title={
              annotatedDownloadCommand.enabled
                ? t("pdfViewer.downloadAnnotated")
                : (annotatedDownloadCommand.reasonLabel ??
                  t("pdfViewer.downloadAnnotated"))
            }
            aria-label={t("pdfViewer.downloadAnnotated")}
            data-capability-id={
              PDF_READER_DOWNLOAD_ANNOTATED_COMMAND.featureId
            }
            data-output-disposition={annotatedDownloadCommand.disposition}
            data-disabled-reason={
              annotatedDownloadCommand.reasonKey ?? undefined
            }
          >
            {flattening
                ? t("pdfViewer.downloadAnnotatedInflight")
                : t("pdfViewer.downloadAnnotated")}
          </PdfRibbonToolButton>
        ) : null}

        {/* Secondary download mode, shown only for documents that actually
            have AcroForm fields: same bake, but the form stays fillable
            (flattenForm: false) — values set + appearances refreshed. */}
        {annotatedDownloadEnabled &&
        hasFlattenableAnnotations &&
        !loading &&
        allFormFields.length > 0 ? (
          <PdfRibbonToolButton
            icon={flattening ? (
              <Loader2 size={14} className="animate-spin" />
            ) : (
              <FileDown size={14} />
            )}
            data-testid="pdf-download-fillable"
            onClick={() =>
              void executeReaderCommand(
                PDF_READER_DOWNLOAD_FILLABLE_COMMAND,
                fillableDownloadRuntimeDecision,
              )
            }
            disabled={!fillableDownloadCommand.enabled}
            title={
              fillableDownloadCommand.enabled
                ? t("pdfViewer.downloadFillable")
                : (fillableDownloadCommand.reasonLabel ??
                  t("pdfViewer.downloadFillable"))
            }
            aria-label={t("pdfViewer.downloadFillable")}
            data-capability-id={
              PDF_READER_DOWNLOAD_FILLABLE_COMMAND.featureId
            }
            data-output-disposition={fillableDownloadCommand.disposition}
            data-disabled-reason={
              fillableDownloadCommand.reasonKey ?? undefined
            }
          >
            {t("pdfViewer.downloadFillable")}
          </PdfRibbonToolButton>
        ) : null}

        {/* "Download with editable annotations": highlights and sticky notes
            leave as real /Annots, so the file opens in Acrobat / WPS with the
            markup still movable, recolourable and commentable — and with the
            comments and reply threads a baked export drops entirely. */}
        {hasEditableXfdfAnnotations && !loading ? (
          <PdfRibbonToolButton
            icon={flattening ? (
              <Loader2 size={14} className="animate-spin" />
            ) : (
              <FileDown size={14} />
            )}
            data-testid="pdf-download-editable-annots"
            onClick={() =>
              void executeReaderCommand(
                PDF_READER_DOWNLOAD_EDITABLE_ANNOTATIONS_COMMAND,
                editableAnnotationsDownloadRuntimeDecision,
              )
            }
            disabled={!editableAnnotationsDownloadCommand.enabled}
            title={
              editableAnnotationsDownloadCommand.enabled
                ? t("pdfViewer.downloadEditableAnnotsHint")
                : (editableAnnotationsDownloadCommand.reasonLabel ??
                  t("pdfViewer.downloadEditableAnnotsHint"))
            }
            aria-label={t("pdfViewer.downloadEditableAnnots")}
            data-capability-id={
              PDF_READER_DOWNLOAD_EDITABLE_ANNOTATIONS_COMMAND.featureId
            }
            data-output-disposition={
              editableAnnotationsDownloadCommand.disposition
            }
            data-disabled-reason={
              editableAnnotationsDownloadCommand.reasonKey ?? undefined
            }
          >
            {t("pdfViewer.downloadEditableAnnots")}
          </PdfRibbonToolButton>
        ) : null}

        {/* Pin a file to the page: arm the mode, click a spot, pick a file.
            The record is written as a real /FileAttachment on export. */}
        {!loading ? (
          <PdfRibbonToolButton
            active={attachMode}
            icon={<Paperclip size={14} />}
            data-testid="pdf-attach-file-toggle"
            onClick={() => void executeFileAttachmentActivation()}
            disabled={nativeFileAttachmentBusy ||
              Boolean(nativeFileAttachmentDisabledReason) ||
              !onNativeFileAttachmentUpsert}
            aria-pressed={attachMode}
            title={nativeFileAttachmentDisabledReason ||
              t("pdfViewer.attachFileHint")}
            aria-label={t("pdfViewer.attachFile")}
          >
            {placedAttachments.length > 0
                ? t("pdfViewer.attachFileCount", {
                    count: placedAttachments.length,
                  })
                : t("pdfViewer.attachFile")}
          </PdfRibbonToolButton>
        ) : null}
        {/* Read aloud: the browser's own voice, page by page. The pair of
            controls collapses to one toggle plus a stop that only shows
            while something is being read. */}
        {speechSupported && !loading ? (
          <>
            <PdfRibbonToolButton
              icon={speechStatus === "speaking" ? (
                <Pause size={14} />
              ) : (
                <Volume2 size={14} />
              )}
              data-testid="pdf-read-aloud-toggle"
              onClick={handleToggleReadAloud}
              title={t("pdfViewer.readAloudHint")}
              aria-label={
                speechStatus === "speaking"
                  ? t("pdfViewer.readAloudPause")
                  : t("pdfViewer.readAloud")
              }
              aria-pressed={speechStatus === "speaking"}
            >
              {speechStatus === "speaking"
                  ? t("pdfViewer.readAloudPause")
                  : speechStatus === "paused"
                    ? t("pdfViewer.readAloudResume")
                    : t("pdfViewer.readAloud")}
            </PdfRibbonToolButton>
            {speechStatus === "speaking" || speechStatus === "paused" ? (
              <PdfRibbonToolButton
                icon={<Square size={14} />}
                data-testid="pdf-read-aloud-stop"
                onClick={handleStopReadAloud}
                aria-label={t("pdfViewer.readAloudStop")}
              >
                {t("pdfViewer.readAloudStop")}
              </PdfRibbonToolButton>
            ) : null}
          </>
        ) : null}

        {/* PDF/A is a source-bound server conversion followed by independent
            veraPDF validation. It never falls back to the reader's local
            annotation bake or a metadata-only PDF/A claim. */}
        {annotatedDownloadEnabled && !loading ? (
          <PdfRibbonToolButton
            icon={flattening ? (
              <Loader2 size={14} className="animate-spin" />
            ) : (
              <FileDown size={14} />
            )}
            data-testid="pdf-download-pdfa"
            onClick={(event) => {
              verifiedPdfaExportTriggerRef.current = event.currentTarget;
              void executeReaderCommand(
                PDF_READER_DOWNLOAD_ARCHIVAL_COMMAND,
                archivalDownloadRuntimeDecision,
              );
            }}
            disabled={!archivalDownloadCommand.enabled}
            title={
              archivalDownloadCommand.enabled
                ? t("pdfViewer.downloadArchivalHint")
                : (archivalDownloadCommand.reasonLabel ??
                  t("pdfViewer.downloadArchivalHint"))
            }
            aria-label={t("pdfViewer.downloadArchival")}
            data-capability-id={
              PDF_READER_DOWNLOAD_ARCHIVAL_COMMAND.featureId
            }
            data-output-disposition={archivalDownloadCommand.disposition}
            data-disabled-reason={
              archivalDownloadCommand.reasonKey ?? undefined
            }
          >
            {t("pdfViewer.downloadArchival")}
          </PdfRibbonToolButton>
        ) : null}

        {!loading && !fieldCandidates ? (
          <PdfRibbonToolButton
            icon={detectBusy ? (
              <Loader2 size={14} className="animate-spin" />
            ) : (
              <FileDown size={14} />
            )}
            ref={detectedFieldsCommandRef}
            data-testid="pdf-detect-fields"
            onClick={() => void handleDetectFormFields()}
            disabled={detectBusy}
            title={t("pdfViewer.detectFieldsHint")}
            aria-label={t("pdfViewer.detectFields")}
          >
            {t("pdfViewer.detectFields")}
          </PdfRibbonToolButton>
        ) : null}

        {fieldCandidates && fieldCandidates.length > 0 ? (
          <>
            <span
              className="pdf-viewer__snapshot-toast office-pop-in"
              role="status"
              data-testid="pdf-detect-count"
            >
              {t("pdfViewer.detectFieldsSelected", {
                selected: acceptedCandidates.length,
                total: fieldCandidates.length,
              })}
            </span>
            <PdfRibbonToolButton
              icon={detectBusy ? (
                <Loader2 size={14} className="animate-spin" />
              ) : (
                <FileDown size={14} />
              )}
              ref={detectedFieldsCommandRef}
              data-testid="pdf-detect-create"
              onClick={() => void handleCreateDetectedFields()}
              disabled={detectBusy || acceptedCandidates.length === 0}
              title={t("pdfViewer.detectFieldsCreate")}
              aria-label={t("pdfViewer.detectFieldsCreate")}
            >
              {t("pdfViewer.detectFieldsCreate")}
            </PdfRibbonToolButton>
            <PdfRibbonToolButton
              data-testid="pdf-detect-cancel"
              onClick={handleCancelDetectedFields}
              disabled={detectBusy}
              title={t("pdfViewer.detectFieldsCancel")}
              aria-label={t("pdfViewer.detectFieldsCancel")}
            >
              {t("pdfViewer.detectFieldsCancel")}
            </PdfRibbonToolButton>
          </>
        ) : null}

        {/* Local PDF tools — available for any loaded PDF (not just annotated
            ones): document properties, structural compress + export AcroForm
            data to CSV. */}
        <PdfLegacyDocumentTools
          lossyCompressionSource={lossyCompressionSource}
          lossyCompressionWritable={lossyCompressionWritable}
          t={t}
          savePolicyV1={savePolicyV1}
          matchedSourceRiskSnapshot={matchedSourceRiskSnapshot}
          loading={loading}
          flattening={flattening}
          formDataExportFormat={formDataExportFormat}
          setFormDataExportFormat={setFormDataExportFormat}
          compressPreset={compressPreset}
          setCompressPreset={setCompressPreset}
          formImportInputRef={formImportInputRef}
          formImportReturnFocusRef={formImportReturnFocusRef}
          formFieldsUnavailable={formFieldsUnavailable}
          formFieldsUnavailableReason={formFieldsUnavailableReason}
          nativeFormMode={nativeFormMode}
          nativeFormDataMutationBlocked={nativeFormDataMutationBlocked}
          nativeFormDataMutationBlockedReason={nativeFormDataMutationBlockedReason}
          handleCompressPdf={handleCompressPdf}
          handleExportFormData={handleExportFormData}
          handleResetForm={handleResetForm}
        />

        {showSignatureProof ? (
          <PdfSignatureProofStatus
            loading={signatureProofLoading}
            presentation={signatureProofPresentation}
            downloadOriginalEnabled={originalDownloadCommand.enabled}
            downloadOriginalReason={
              originalDownloadCommand.reasonLabel ?? undefined
            }
            downloadOriginalReasonKey={
              originalDownloadCommand.reasonKey ?? undefined
            }
            downloadOriginalDisposition={originalDownloadCommand.disposition}
            downloadOriginalCapabilityId={
              PDF_READER_DOWNLOAD_ORIGINAL_COMMAND.featureId
            }
            onDownloadOriginal={handleSignatureProofOriginalDownload}
          />
        ) : null}

        {/* File-level: merge other PDFs into this one. Rendered once loaded so
            the disabled-button-swallows-clicks pitfall above is avoided. */}
        {!loading ? (
          <>
            <PdfRibbonToolButton
              icon={mergeBusy ? (
                <Loader2 size={14} className="animate-spin" />
              ) : (
                <Combine size={14} />
              )}
              data-testid="pdf-merge-files"
              onClick={handleOpenMergePicker}
              disabled={
                mergeBusy ||
                (savePolicyV1 && matchedSourceRiskSnapshot === null)
              }
              title={t("pdfViewer.mergeFiles")}
              aria-label={t("pdfViewer.mergeFiles")}
              data-capability-id={PDF_READER_MERGE_COMMAND.featureId}
              data-output-disposition="flatten-with-consent"
              data-disabled-reason={
                savePolicyV1 && matchedSourceRiskSnapshot === null
                  ? "pdf.savePolicy.sourceIdentityRequired"
                  : undefined
              }
            >
              {mergeBusy
                  ? t("pdfViewer.mergeInflight")
                  : t("pdfViewer.mergeFiles")}
            </PdfRibbonToolButton>
          </>
        ) : null}

        {/* File-level: insert pages FROM another PDF at a chosen position +
            optional page-subset (Acrobat "Insert Pages"). A popover collects
            the 1-based position + range; a hidden file input supplies bytes. */}
        {!loading ? (
          <div className="pdf-viewer__split pdf-viewer__insert-pages">
            <PdfRibbonToolButton
              icon={insertBusy || nativePageImportBusy ? (
                <Loader2 size={14} className="animate-spin" />
              ) : (
                <FileInput size={14} />
              )}
              data-testid="pdf-insert-toggle"
              onClick={(event) => {
                insertDialogReturnFocusRef.current = event.currentTarget;
                selectInsertSourceKind("pdf");
                setInsertOpen((open) => !open);
              }}
              disabled={
                insertBusy ||
                nativePageImportBusy ||
                !onNativePageImport ||
                Boolean(nativePageImportDisabledReason)
              }
              aria-expanded={insertOpen}
              aria-haspopup="dialog"
              title={
                nativePageImportDisabledReason ??
                (!onNativePageImport
                  ? t("pdf.viewer.nativePageImportUnavailableHost")
                  : t("pdfViewer.insertPages"))
              }
              aria-label={t("pdfViewer.insertPages")}
            >
              {t("pdfViewer.insertPages")}
            </PdfRibbonToolButton>
          </div>
        ) : null}

        {/* Export page(s) as a PNG / JPG image (Acrobat/WPS). A scope chooser
            selects the current page, ALL pages, or an Acrobat-style range;
            each format button rasterizes the selected pages and downloads one
            image per page. The "current" scope grabs the on-screen canvas;
            all / range render each page offscreen with pdf.js. */}
        <PdfExportImageControl
          t={t}
          loading={loading}
          featuredLegacyToolsOpen={featuredLegacyToolsOpen}
          setExportImageOpen={setExportImageOpen}
          exportImageOpen={exportImageOpen}
          exportImageDialogReturnFocusRef={exportImageDialogReturnFocusRef}
          exportImageScope={exportImageScope}
          setExportImageScope={setExportImageScope}
          exportImageRange={exportImageRange}
          setExportImageRange={setExportImageRange}
          exportImageCombinePages={exportImageCombinePages}
          setExportImageCombinePages={setExportImageCombinePages}
          exportImageResolution={exportImageResolution}
          setExportImageResolution={setExportImageResolution}
          exportImageBusy={exportImageBusy}
          rasterOutputAuthorityBlocked={rasterOutputAuthorityBlocked}
          handleExportPagesAsImages={handleExportPagesAsImages}
          exportImageOverLimit={exportImageOverLimit}
          exportImageCurrentPageSize={exportImageCurrentPageSize}
        />

        {/* Optional Content Groups ("layers", Acrobat 图层): toggle the
            visibility of the document's layers. Shown only when the PDF has
            any. Reuses the split-popover chrome for a consistent look. */}
        <PdfLayersControl
          t={t}
          loading={loading}
          ocgLayers={ocgLayers}
          ocgOrder={ocgOrder}
          setLayersPanelOpen={setLayersPanelOpen}
          layersPanelOpen={layersPanelOpen}
          handleToggleLayer={handleToggleLayer}
          layersDialogReturnFocusRef={layersDialogReturnFocusRef}
        />
        {/* File-level: split this PDF into multiple downloadable files. */}
        <PdfSplitControl
          t={t}
          savePolicyV1={savePolicyV1}
          matchedSourceRiskSnapshot={matchedSourceRiskSnapshot}
          loading={loading}
          outline={outline}
          splitBusy={splitBusy}
          setSplitOpen={setSplitOpen}
          splitOpen={splitOpen}
          splitDialogReturnFocusRef={splitDialogReturnFocusRef}
          splitMode={splitMode}
          setSplitMode={setSplitMode}
          splitRangeSpec={splitRangeSpec}
          splitRangeError={splitRangeError}
          setSplitRangeSpec={setSplitRangeSpec}
          splitEveryNValue={splitEveryNValue}
          setSplitEveryNValue={setSplitEveryNValue}
          handleSplitDownload={handleSplitDownload}
          activeSplitCommand={activeSplitCommand}
        />

        <PdfLegacyDownloadMenu

          filename={filename}

          downloadUrl={downloadUrl}

          onDownload={onDownload}

          preferAnnotatedDownload={preferAnnotatedDownload}

          annotatedDownloadEnabled={annotatedDownloadEnabled}

          t={t}

          redactionHydrationStatusForCurrentDocument={redactionHydrationStatusForCurrentDocument}

          shapesHydrationStatusForCurrentDocument={shapesHydrationStatusForCurrentDocument}

          executeReaderCommand={executeReaderCommand}

          originalDownloadRuntimeDecision={originalDownloadRuntimeDecision}

          currentDownloadRuntimeDecision={currentDownloadRuntimeDecision}

          originalDownloadCommand={originalDownloadCommand}

          currentDownloadCommand={currentDownloadCommand}

        />
        </OfficeHomeCommandStack>
        </PdfSurfaceRibbonGroup>
      ) : null}
    </PdfOfficeRibbon>
  );
}
