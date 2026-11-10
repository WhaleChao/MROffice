"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { ptToPx } from "@mrnote/core/web/lib/office-units.ts";
import type { Editor } from "@tiptap/core";
import { useMemo, type CSSProperties, type ReactNode } from "react";

import { OfficeHomeCheckbox } from "@mrnote/core/web/components/office-ribbon/OfficeHomeRibbon.tsx";
import {
  OfficeTableStyleGallery,
  OfficeTableStylePreviewTable,
  officeTableStyleGallerySlot,
  type OfficeTableStyleGalleryDocumentItem,
  type OfficeTableStyleGalleryItemId,
} from "@mrnote/core/web/components/office-ribbon/OfficeTableStyleGallery.tsx";
import {
  documentTableDesignStyleCatalogFromDoc,
  resolveDocumentTableStyleCellVisual,
  setDocumentTableDesignState,
  type DocumentTableDesignTarget,
  type DocumentTableLook,
  type DocumentTableStyleVisualAuthority,
  type DocumentTableStyleVisualBorder,
  type DocumentTableStyleVisualLayer,
} from "@mrnote/core/web/lib/document-table-design-authority.ts";
import {
  officeWordTableStyleByName,
  officeWordTableStyles,
} from "@mrnote/core/web/lib/office-table-style-word.ts";
import { useOfficeData } from "@mrnote/core/web/lib/office-data/use-office-data.ts";

import styles from "./DocumentTableStyleGallery.module.css";
import {
  documentOfficeTableStyleForId,
  documentOfficeTableStyleIdForSlot,
  documentTableStyleVisualForId,
  findSelectedDocumentTable,
  isSourceBoundDocxTable,
  localOfficeTableLookForApplication,
} from "@mrnote/core/web/components/console/editor/native-document/documentOfficeTableStyles.ts";
import {
  setDocumentLocalTableStylePreview,
  setDocumentTableStylePreview,
} from "@mrnote/core/web/components/console/editor/native-document/tableBorders.ts";
import { commitDocumentTableStyleSlot } from "./tableButtons";

const PREVIEW_SIZE = 5;
const CURRENT_COLOR = "#64748b";

type PhysicalEdge = "top" | "right" | "bottom" | "left";

export interface DocumentTableStyleGalleryLabels {
  readonly ariaLabel: string;
  readonly more: string;
  readonly clear: string;
  readonly currentStyle: string;
  readonly documentStyles: string;
  readonly disabledReason: string;
  readonly currentStyleUnavailableReason: string;
  readonly options: string;
  readonly tableColor: string;
  readonly shadingAndBorders: string;
  readonly headerRow: string;
  readonly totalRow: string;
  readonly firstColumn: string;
  readonly lastColumn: string;
  readonly bandedRows: string;
  readonly bandedColumns: string;
}

interface DocumentTableStyleGalleryBaseProps {
  readonly editor: Editor;
  readonly labels: DocumentTableStyleGalleryLabels;
  readonly look: DocumentTableLook;
  readonly onLookChange: (key: keyof DocumentTableLook) => void;
  readonly disabledReason?: string;
  readonly sourceLimitation?: string;
  readonly borders: ReactNode;
  readonly shading: ReactNode;
  readonly tableColor: ReactNode;
}

export type DocumentTableStyleGalleryProps =
  | (DocumentTableStyleGalleryBaseProps & {
      readonly mode: "native";
      readonly target: DocumentTableDesignTarget;
    })
  | (DocumentTableStyleGalleryBaseProps & {
      readonly mode: "local";
      /** `tableStyleTheme`: a gallery style id or a legacy recipe id. */
      readonly currentStyleId: string | null;
      /** Display name of a legacy recipe the table still wears. */
      readonly legacyStyleLabel?: string | null;
    });

type BorderCandidate =
  | { readonly kind: "cleared" }
  | {
      readonly kind: "line";
      readonly line: DocumentTableStyleVisualBorder;
    };

function cellBorderCandidate(
  layer: DocumentTableStyleVisualLayer | null,
  edge: PhysicalEdge,
): BorderCandidate | null {
  if (!layer) return null;
  if (layer.clearedCellBorders.includes(edge)) return { kind: "cleared" };
  const line = layer.cellBorders[edge];
  return line ? { kind: "line", line } : null;
}

function tableBorderCandidate(
  layer: DocumentTableStyleVisualLayer | null,
  edge: PhysicalEdge,
  row: number,
  column: number,
): BorderCandidate | null {
  if (!layer) return null;
  const last = PREVIEW_SIZE - 1;
  const tableEdge =
    edge === "top"
      ? row === 0
        ? "top"
        : "insideH"
      : edge === "bottom"
        ? row === last
          ? "bottom"
          : "insideH"
        : edge === "left"
          ? column === 0
            ? "left"
            : "insideV"
          : column === last
            ? "right"
            : "insideV";
  if (layer.clearedTableBorders.includes(tableEdge)) {
    return { kind: "cleared" };
  }
  const line = layer.tableBorders[tableEdge];
  return line ? { kind: "line", line } : null;
}

/** The line Word draws on one cell edge: own cell border, neighbour, table. */
function resolvedBoundaryLine(
  layers: ReadonlyMap<string, DocumentTableStyleVisualLayer | null>,
  row: number,
  column: number,
  edge: PhysicalEdge,
): DocumentTableStyleVisualBorder | null {
  const ownLayer = layers.get(`${row}-${column}`) ?? null;
  const ownCell = cellBorderCandidate(ownLayer, edge);
  if (ownCell) return ownCell.kind === "line" ? ownCell.line : null;
  const opposite = {
    top: "bottom",
    right: "left",
    bottom: "top",
    left: "right",
  }[edge] as PhysicalEdge;
  const [neighborRow, neighborColumn] =
    edge === "top"
      ? [row - 1, column]
      : edge === "bottom"
        ? [row + 1, column]
        : edge === "left"
          ? [row, column - 1]
          : [row, column + 1];
  const neighborCell = cellBorderCandidate(
    layers.get(`${neighborRow}-${neighborColumn}`) ?? null,
    opposite,
  );
  if (neighborCell) {
    return neighborCell.kind === "line" ? neighborCell.line : null;
  }
  const table = tableBorderCandidate(ownLayer, edge, row, column);
  return table?.kind === "line" ? table.line : null;
}

function borderCss(line: DocumentTableStyleVisualBorder | null): string {
  if (!line || line.style === "none" || line.style === "nil") return "none";
  // Thumbnails draw every line at half its page width, as the Presentation
  // gallery does, so a 1 pt rule stays a hairline in a 60 x 40 sample.
  const width = Math.max(0.5, ptToPx((line.sizePt ?? 0.5)) / 2);
  const style =
    line.style === "double"
      ? "double"
      : line.style === "dotted"
        ? "dotted"
        : line.style.startsWith("dash") ||
            line.style === "dotDash" ||
            line.style === "dotDotDash"
          ? "dashed"
          : "solid";
  const color =
    !line.color || line.color === "currentColor" ? CURRENT_COLOR : line.color;
  return `${Number(width.toFixed(3))}px ${style} ${color}`;
}

/** A 5 x 5 sample resolved through the same Word cascade as the page. */
function DocumentTableStylePreview({
  visual,
  look,
}: {
  readonly visual: DocumentTableStyleVisualAuthority;
  readonly look: DocumentTableLook;
}) {
  const layers = new Map<string, DocumentTableStyleVisualLayer | null>();
  for (let row = 0; row < PREVIEW_SIZE; row += 1) {
    for (let column = 0; column < PREVIEW_SIZE; column += 1) {
      layers.set(
        `${row}-${column}`,
        resolveDocumentTableStyleCellVisual(visual, look, {
          rowStart: row,
          rowEnd: row,
          columnStart: column,
          columnEnd: column,
          rows: PREVIEW_SIZE,
          columns: PREVIEW_SIZE,
        }),
      );
    }
  }
  return (
    <OfficeTableStylePreviewTable
      cellStyle={(row, column) => {
        const layer = layers.get(`${row}-${column}`) ?? null;
        const css: CSSProperties = {
          borderTop: borderCss(resolvedBoundaryLine(layers, row, column, "top")),
          borderLeft: borderCss(
            resolvedBoundaryLine(layers, row, column, "left"),
          ),
          borderBottom: borderCss(
            resolvedBoundaryLine(layers, row, column, "bottom"),
          ),
          borderRight: borderCss(
            resolvedBoundaryLine(layers, row, column, "right"),
          ),
        };
        if (layer?.fill) css.background = layer.fill;
        if (layer?.color) css.color = layer.color;
        if (layer?.bold) css.fontWeight = 700;
        return css;
      }}
    />
  );
}

function PlainTablePreview() {
  return (
    <OfficeTableStylePreviewTable
      cellStyle={() => ({ border: "0.5px solid #cbd5e1" })}
    />
  );
}

export function DocumentTableStyleGallery(
  props: DocumentTableStyleGalleryProps,
) {
  const {
    editor,
    labels,
    look,
    onLookChange,
    disabledReason,
    sourceLimitation,
    borders,
    shading,
    tableColor,
  } = props;
  const nativeTarget = props.mode === "native" ? props.target : null;
  const doc = nativeTarget ? nativeTarget.doc : editor.state.doc;
  const localTable =
    props.mode === "local" ? findSelectedDocumentTable(editor) : null;
  const sourceBound = Boolean(
    localTable && isSourceBoundDocxTable(localTable.node),
  );
  const editable =
    (nativeTarget ? nativeTarget.authority.receipt.editable : !sourceBound) &&
    !disabledReason;
  const currentStyleId = nativeTarget
    ? nativeTarget.authority.current.styleId
    : props.mode === "local"
      ? props.currentStyleId
      : null;
  // Thumbnails show each style with the options this table will wear: its
  // own look, or Word's default look for a table that never chose one.
  const previewLook =
    nativeTarget || !localTable
      ? look
      : localOfficeTableLookForApplication(localTable.node);
  const previewLookKey = JSON.stringify(previewLook);
  const legacyStyleLabel =
    props.mode === "local" ? (props.legacyStyleLabel ?? null) : null;
  // The thumbnails depend on the document's style graph, catalogue and
  // theme, not on its text: typing in a table must not redraw 74 samples.
  const {
    styleDefinitions,
    tableStyleVisualCatalog,
    documentThemeColors,
    designColorSet,
  } = doc.attrs as Record<string, unknown>;
  const nativeMode = props.mode === "native";
  // The 74 looks are served (a signed-in page asks once); until they land,
  // and on a guest page, the thumbnails draw MRAI's recipe of each family
  // (`office-table-style-fallback.ts`), and draw again when they land.
  const servedCatalog = useOfficeData("word-table-styles");

  const model = useMemo(() => {
    const slotPreviews = new Map<number, ReactNode>();
    for (const style of officeWordTableStyles()) {
      const styleId = documentOfficeTableStyleIdForSlot(doc, style.slot);
      const visual = styleId ? documentTableStyleVisualForId(doc, styleId) : null;
      slotPreviews.set(
        style.slot,
        visual ? (
          <DocumentTableStylePreview visual={visual} look={previewLook} />
        ) : (
          <PlainTablePreview />
        ),
      );
    }
    const documentItems: OfficeTableStyleGalleryDocumentItem[] = [];
    const documentStyleIds = new Map<string, string>();
    let value: OfficeTableStyleGalleryItemId | null = null;
    const currentGallery = documentOfficeTableStyleForId(doc, currentStyleId);
    if (currentGallery) value = `web:${currentGallery.slot}`;
    if (nativeMode) {
      // The document's own Word table styles follow the gallery, as the
      // Presentation gallery lists a deck's own styles: applicable built-in
      // styles only, never a second tile for a gallery style it defines.
      for (const entry of documentTableDesignStyleCatalogFromDoc(doc)) {
        if (officeWordTableStyleByName(entry.name)) continue;
        const id = `deck:${entry.styleId}` as const;
        documentStyleIds.set(id, entry.styleId);
        documentItems.push({
          id,
          label: entry.name,
          preview: (
            <DocumentTableStylePreview visual={entry.visual} look={previewLook} />
          ),
        });
        if (entry.styleId === currentStyleId) value = id;
      }
    }
    if (currentStyleId && value === null) {
      const label =
        legacyStyleLabel ?? `${labels.currentStyle} (${currentStyleId})`;
      const visual = documentTableStyleVisualForId(doc, currentStyleId);
      const id = `current:${currentStyleId}` as const;
      value = id;
      documentItems.push({
        id,
        label,
        disabled: true,
        disabledReason: `${labels.currentStyleUnavailableReason}: ${label}`,
        preview: visual ? (
          <DocumentTableStylePreview visual={visual} look={previewLook} />
        ) : (
          <PlainTablePreview />
        ),
      });
    }
    return { slotPreviews, documentItems, documentStyleIds, value };
    // `previewLookKey` stands for `previewLook`'s content and the three style
    // attrs for `doc`: its text never changes a thumbnail.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [
    currentStyleId,
    designColorSet,
    documentThemeColors,
    labels.currentStyle,
    labels.currentStyleUnavailableReason,
    legacyStyleLabel,
    nativeMode,
    previewLookKey,
    servedCatalog,
    styleDefinitions,
    tableStyleVisualCatalog,
  ]);
  const galleryLabels = useMemo(
    () => ({
      ariaLabel: labels.ariaLabel,
      more: labels.more,
      clear: labels.clear,
      document: labels.documentStyles,
    }),
    [labels.ariaLabel, labels.clear, labels.documentStyles, labels.more],
  );

  const effectiveDisabledReason = disabledReason ?? labels.disabledReason;
  const ariaLabel = !editable
    ? `${labels.ariaLabel}: ${effectiveDisabledReason}`
    : labels.ariaLabel;
  const preserveOnlyTitle = !editable
    ? `${currentStyleId ?? labels.currentStyle} — ${effectiveDisabledReason}`
    : undefined;
  const surfaceTitle = [preserveOnlyTitle, sourceLimitation]
    .filter((value): value is string => Boolean(value))
    .join(" · ");
  const testIdPrefix =
    props.mode === "native"
      ? "document-toolbar-native-table-styles"
      : "document-toolbar-table-styles";

  /** The native table-design state a gallery item stands for. */
  const nativeStateFor = (id: OfficeTableStyleGalleryItemId) => {
    if (!nativeTarget) return null;
    const slot = officeTableStyleGallerySlot(id);
    const styleId =
      id === "clear"
        ? null
        : slot !== null
          ? documentOfficeTableStyleIdForSlot(editor.state.doc, slot)
          : (model.documentStyleIds.get(id) ?? undefined);
    if (styleId === undefined || (id !== "clear" && styleId === null)) {
      return null;
    }
    return { ...nativeTarget.authority.current, styleId };
  };

  const gallery = (
    <OfficeTableStyleGallery
      application="docs"
      labels={galleryLabels}
      className={styles.gallery}
      disabled={!editable}
      disabledReason={effectiveDisabledReason}
      documentItems={model.documentItems}
      renderSlotPreview={(slot) => model.slotPreviews.get(slot) ?? null}
      value={model.value}
      testIdPrefix={testIdPrefix}
      onPreviewChange={(id) => {
        if (nativeTarget) {
          const state = id === null ? null : nativeStateFor(id);
          if (id !== null && !state) return;
          setDocumentTableStylePreview(editor, nativeTarget, state);
          return;
        }
        if (id === null) {
          setDocumentLocalTableStylePreview(editor, undefined);
          return;
        }
        const slot = officeTableStyleGallerySlot(id);
        if (id === "clear" || slot !== null) {
          setDocumentLocalTableStylePreview(editor, slot);
        }
      }}
      onCommit={(id) => {
        const slot = officeTableStyleGallerySlot(id);
        if (nativeTarget) {
          if (id === "clear" || slot !== null) {
            commitDocumentTableStyleSlot(editor, slot, nativeTarget);
            return;
          }
          const state = nativeStateFor(id);
          if (state) setDocumentTableDesignState(editor, nativeTarget, state);
          return;
        }
        if (id !== "clear" && slot === null) return;
        setDocumentLocalTableStylePreview(editor, undefined);
        if (commitDocumentTableStyleSlot(editor, slot, null)) {
          editor.chain().focus().run();
        }
      }}
    />
  );
  const lookOptions: ReadonlyArray<{
    key: keyof DocumentTableLook;
    label: string;
    testId: string;
  }> = [
    {
      key: "firstRow",
      label: labels.headerRow,
      testId: "document-toolbar-toggle-header-row",
    },
    {
      key: "lastRow",
      label: labels.totalRow,
      testId: "document-toolbar-table-style-total-row",
    },
    {
      key: "bandedRows",
      label: labels.bandedRows,
      testId: "document-toolbar-table-style-banded-rows",
    },
    {
      key: "firstColumn",
      label: labels.firstColumn,
      testId: "document-toolbar-table-style-first-column",
    },
    {
      key: "lastColumn",
      label: labels.lastColumn,
      testId: "document-toolbar-table-style-last-column",
    },
    {
      key: "bandedColumns",
      label: labels.bandedColumns,
      testId: "document-toolbar-table-style-banded-columns",
    },
  ];

  return (
    <div
      className={styles.surface}
      aria-label={
        surfaceTitle ? `${labels.ariaLabel}: ${surfaceTitle}` : labels.ariaLabel
      }
      data-mode={props.mode}
      data-testid={
        props.mode === "native"
          ? "document-toolbar-native-table-design-surface"
          : "document-toolbar-table-design-surface"
      }
      title={surfaceTitle || undefined}
    >
      <section
        className={`${styles.section} ${styles.stylesFamily}`}
        aria-label={labels.ariaLabel}
        data-office-table-design-group="styles"
      >
        <div className={styles.stylesFamilyContent}>
          <div
            className={styles.galleryPart}
            role="group"
            aria-label={labels.ariaLabel}
            data-office-table-design-section="styles"
          >
            {gallery}
          </div>
          <div
            className={styles.tableColorPart}
            role="group"
            aria-label={labels.tableColor}
            data-office-table-design-section="table-color"
          >
            {tableColor}
          </div>
          <div
            className={styles.options}
            role="group"
            aria-label={labels.options}
            data-office-table-design-section="options"
          >
            {lookOptions.map((option) => {
              const accessibleLabel = !editable
                ? `${option.label}: ${effectiveDisabledReason}`
                : option.label;
              return (
                <OfficeHomeCheckbox
                  key={option.key}
                  checked={look[option.key]}
                  disabled={!editable}
                  label={option.label}
                  aria-label={accessibleLabel}
                  title={!editable ? effectiveDisabledReason : option.label}
                  onCheckedChange={() => onLookChange(option.key)}
                  testId={option.testId}
                />
              );
            })}
          </div>
        </div>
        <span className={styles.caption}>{labels.ariaLabel}</span>
      </section>
      <section
        className={`${styles.section} ${styles.bordersShadingFamily}`}
        aria-label={labels.shadingAndBorders}
        data-office-table-design-section="borders-shading"
        data-office-table-design-group="borders-shading"
      >
        <div className={styles.commands}>
          {borders}
          {shading}
        </div>
        <span className={styles.caption}>{labels.shadingAndBorders}</span>
      </section>
    </div>
  );
}
