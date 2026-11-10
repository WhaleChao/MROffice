"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import {
  useCallback,
  useEffect,
  useId,
  useMemo,
  useRef,
  useState,
  type KeyboardEvent as ReactKeyboardEvent,
  type MouseEvent as ReactMouseEvent,
} from "react";
import {
  ChartNoAxesCombined,
  ChevronDown,
  ChevronRight,
  Grid2X2,
  Grid3x3,
  Eye,
  EyeOff,
  ImageIcon,
  MoreHorizontal,
  Pencil,
  Search,
  Sheet as SheetIcon,
  Shapes,
  Table2,
  Trash2,
  X,
  type HouseIcon,
} from "@mrnote/core/web/components/office-ribbon/icons/houseIcons.tsx";

import {
  DropdownMenu,
  DropdownMenuContent,
  DropdownMenuItem,
  DropdownMenuTrigger,
} from "@mrnote/core/web/components/ui/dropdown-menu.tsx";

import {
  officeDialogParts as dialogParts,
  spreadsheetPaneClassName,
} from "@mrnote/core/web/components/office-ribbon/officeDialogChrome.tsx";
import { OfficeTaskPaneDockedTitle } from "@mrnote/core/web/components/office-ribbon/OfficeTaskPane.tsx";
import { useOfficeTaskPaneDockSlot } from "@mrnote/core/web/components/office-ribbon/officeTaskPaneDock.ts";
import { OfficeTaskPaneOptions } from "@mrnote/core/web/components/office-ribbon/officeTaskPaneLayout.tsx";

import styles from "./SpreadsheetNavigationPane.module.css";

export type SpreadsheetNavigationElementKind =
  "range" | "table" | "chart" | "pivot" | "image" | "shape";

/**
 * A render-only workbook object. Keep this snapshot free of editor/runtime
 * instances so SpreadsheetWindow remains the sole selection authority.
 */
export interface SpreadsheetNavigationElementSnapshot {
  readonly id: string;
  readonly name: string;
  readonly detail?: string | null;
  readonly hidden?: boolean;
  readonly actions?: readonly SpreadsheetNavigationElementActionSnapshot[];
}

export type SpreadsheetNavigationElementActionId =
  | "rename"
  | "delete"
  | "hide"
  | "show";

export interface SpreadsheetNavigationElementActionSnapshot {
  readonly id: SpreadsheetNavigationElementActionId;
  readonly label: string;
  readonly destructive?: boolean;
  readonly disabledReason?: string | null;
}

/** A render-only projection of the navigable objects on one worksheet. */
export interface SpreadsheetNavigationSheetSnapshot {
  readonly id: string;
  readonly name: string;
  readonly namedRanges?: readonly SpreadsheetNavigationElementSnapshot[];
  readonly tables?: readonly SpreadsheetNavigationElementSnapshot[];
  readonly charts?: readonly SpreadsheetNavigationElementSnapshot[];
  readonly pivotTables?: readonly SpreadsheetNavigationElementSnapshot[];
  readonly images?: readonly SpreadsheetNavigationElementSnapshot[];
  readonly shapes?: readonly SpreadsheetNavigationElementSnapshot[];
}

export interface SpreadsheetNavigationElementSelection {
  readonly sheetId: string;
  readonly kind: SpreadsheetNavigationElementKind;
  readonly id: string;
}

/**
 * All user-facing copy is supplied by the caller. Disclosure labels may use a
 * `{sheet}` token, and emptySheet may use the same token.
 */
type NavigationDynamicLabel = string | ((values: Readonly<Record<string, string>>) => string);

export interface SpreadsheetNavigationPaneLabels {
  readonly title: string;
  readonly close: string;
  readonly expandSheet: NavigationDynamicLabel;
  readonly collapseSheet: NavigationDynamicLabel;
  readonly sheets: string;
  readonly search: string;
  readonly clearSearch: string;
  readonly namedRanges: string;
  readonly tables: string;
  readonly charts: string;
  readonly pivotTables: string;
  readonly images: string;
  readonly shapes: string;
  readonly emptyWorkbook: string;
  readonly noMatches: string;
  readonly emptySheet: NavigationDynamicLabel;
  readonly elementActions: NavigationDynamicLabel;
}

export interface SpreadsheetNavigationPaneProps {
  readonly sheets: readonly SpreadsheetNavigationSheetSnapshot[];
  readonly activeSheetId: string | null;
  readonly selectedElement?: SpreadsheetNavigationElementSelection | null;
  readonly labels: SpreadsheetNavigationPaneLabels;
  readonly className?: string;
  /** Focuses the active sheet row when the newly mounted pane opens. */
  readonly autoFocus?: boolean;
  readonly onClose: () => void;
  readonly onSelectSheet: (sheetId: string) => void;
  readonly onSelectElement: (
    selection: SpreadsheetNavigationElementSelection,
  ) => void;
  readonly onElementAction?: (
    selection: SpreadsheetNavigationElementSelection,
    action: SpreadsheetNavigationElementActionId,
  ) => void;
}

interface NavigationElementGroup {
  readonly kind: SpreadsheetNavigationElementKind;
  readonly label: string;
  readonly elements: readonly SpreadsheetNavigationElementSnapshot[];
  readonly Icon: HouseIcon;
}

interface NavigationSheetGroup {
  readonly sheet: SpreadsheetNavigationSheetSnapshot;
  readonly elementGroups: readonly NavigationElementGroup[];
  readonly elementCount: number;
}

type VisibleNavigationRow =
  | {
      readonly key: string;
      readonly type: "sheet";
      readonly sheetId: string;
      readonly hasChildren: boolean;
      readonly collapsed: boolean;
    }
  | {
      readonly key: string;
      readonly type: "element";
      readonly sheetId: string;
      readonly selection: SpreadsheetNavigationElementSelection;
    };

function formatLabel(
  template: NavigationDynamicLabel,
  values: Readonly<Record<string, string>>,
): string {
  if (typeof template === "function") return template(values);
  return Object.entries(values).reduce(
    (label, [token, value]) => label.replaceAll(`{${token}}`, () => value),
    template,
  );
}

function sheetRowKey(sheetId: string): string {
  return JSON.stringify(["sheet", sheetId]);
}

function elementRowKey(
  selection: SpreadsheetNavigationElementSelection,
): string {
  return JSON.stringify([
    "element",
    selection.sheetId,
    selection.kind,
    selection.id,
  ]);
}

function collectElementGroups(
  sheet: SpreadsheetNavigationSheetSnapshot,
  labels: SpreadsheetNavigationPaneLabels,
): readonly NavigationElementGroup[] {
  const groups: readonly NavigationElementGroup[] = [
    {
      kind: "range",
      label: labels.namedRanges,
      elements: sheet.namedRanges ?? [],
      Icon: Grid2X2,
    },
    {
      kind: "table",
      label: labels.tables,
      elements: sheet.tables ?? [],
      Icon: Table2,
    },
    {
      kind: "chart",
      label: labels.charts,
      elements: sheet.charts ?? [],
      Icon: ChartNoAxesCombined,
    },
    {
      kind: "pivot",
      label: labels.pivotTables,
      elements: sheet.pivotTables ?? [],
      Icon: Grid3x3,
    },
    {
      kind: "image",
      label: labels.images,
      elements: sheet.images ?? [],
      Icon: ImageIcon,
    },
    {
      kind: "shape",
      label: labels.shapes,
      elements: sheet.shapes ?? [],
      Icon: Shapes,
    },
  ];
  return groups.filter((group) => group.elements.length > 0);
}

function isSameSelection(
  left: SpreadsheetNavigationElementSelection | null | undefined,
  right: SpreadsheetNavigationElementSelection,
): boolean {
  return (
    left?.sheetId === right.sheetId &&
    left.kind === right.kind &&
    left.id === right.id
  );
}

/**
 * Excel-Web-style, read-only workbook Navigation pane. The pane owns only
 * disclosure and roving-focus state; every workbook selection is delegated to
 * the callbacks supplied by SpreadsheetWindow.
 */
export function SpreadsheetNavigationPane({
  sheets,
  activeSheetId,
  selectedElement = null,
  labels,
  className,
  autoFocus = true,
  onClose,
  onSelectSheet,
  onSelectElement,
  onElementAction,
}: SpreadsheetNavigationPaneProps) {
  const paneId = useId();
  // Docked in the workbook's task pane column (OfficeDockedPane): the title
  // row is the dock's tab strip. Escape leaves the pane open, as Excel's
  // task panes do; its X or Task Pane Options > Close closes it.
  const dockSlot = useOfficeTaskPaneDockSlot(labels.title);
  const paneRef = useRef<HTMLElement>(null);
  const closeButtonRef = useRef<HTMLButtonElement>(null);
  const rowRefs = useRef(new Map<string, HTMLButtonElement>());
  const initialFocusAppliedRef = useRef(false);
  const [collapsedSheetIds, setCollapsedSheetIds] = useState<
    ReadonlySet<string>
  >(() => new Set());
  const [focusedRowKey, setFocusedRowKey] = useState<string | null>(null);
  const [actionMenuKey, setActionMenuKey] = useState<string | null>(null);
  const [searchQuery, setSearchQuery] = useState("");
  const normalizedSearch = searchQuery.trim().toLocaleLowerCase();

  const sheetGroups = useMemo<readonly NavigationSheetGroup[]>(
    () =>
      sheets.flatMap((sheet) => {
        const allElementGroups = collectElementGroups(sheet, labels);
        const sheetMatches =
          normalizedSearch !== "" &&
          sheet.name.toLocaleLowerCase().includes(normalizedSearch);
        const elementGroups =
          normalizedSearch === "" || sheetMatches
            ? allElementGroups
            : allElementGroups
                .map((group) => ({
                  ...group,
                  elements: group.elements.filter((element) =>
                    [element.name, element.detail]
                      .filter((value): value is string =>
                        typeof value === "string",
                      )
                      .some((value) =>
                        value.toLocaleLowerCase().includes(normalizedSearch),
                      ),
                  ),
                }))
                .filter((group) => group.elements.length > 0);
        if (
          normalizedSearch !== "" &&
          !sheetMatches &&
          elementGroups.length === 0
        )
          return [];
        return {
          sheet,
          elementGroups,
          elementCount: elementGroups.reduce(
            (count, group) => count + group.elements.length,
            0,
          ),
        };
      }),
    [labels, normalizedSearch, sheets],
  );

  const visibleRows = useMemo<readonly VisibleNavigationRow[]>(() => {
    const rows: VisibleNavigationRow[] = [];
    sheetGroups.forEach(({ sheet, elementGroups, elementCount }) => {
      const collapsed =
        normalizedSearch === "" &&
        elementCount > 0 &&
        collapsedSheetIds.has(sheet.id);
      rows.push({
        key: sheetRowKey(sheet.id),
        type: "sheet",
        sheetId: sheet.id,
        hasChildren: elementCount > 0,
        collapsed,
      });
      if (collapsed) return;
      elementGroups.forEach((group) => {
        group.elements.forEach((element) => {
          const selection = {
            sheetId: sheet.id,
            kind: group.kind,
            id: element.id,
          } as const;
          rows.push({
            key: elementRowKey(selection),
            type: "element",
            sheetId: sheet.id,
            selection,
          });
        });
      });
    });
    return rows;
  }, [collapsedSheetIds, normalizedSearch, sheetGroups]);

  const rowIndexByKey = useMemo(
    () => new Map(visibleRows.map((row, index) => [row.key, index])),
    [visibleRows],
  );
  const activeSheetRowKey = activeSheetId ? sheetRowKey(activeSheetId) : null;
  const fallbackRowKey =
    (activeSheetRowKey && rowIndexByKey.has(activeSheetRowKey)
      ? activeSheetRowKey
      : null) ??
    visibleRows[0]?.key ??
    null;
  const rovingRowKey =
    focusedRowKey && rowIndexByKey.has(focusedRowKey)
      ? focusedRowKey
      : fallbackRowKey;

  const setRowRef = useCallback(
    (key: string, node: HTMLButtonElement | null) => {
      if (node) rowRefs.current.set(key, node);
      else rowRefs.current.delete(key);
    },
    [],
  );

  const focusRow = useCallback((key: string) => {
    setFocusedRowKey(key);
    rowRefs.current.get(key)?.focus({ preventScroll: true });
  }, []);

  const focusRowAt = useCallback(
    (index: number) => {
      if (visibleRows.length === 0) return;
      const boundedIndex = Math.max(0, Math.min(index, visibleRows.length - 1));
      const row = visibleRows[boundedIndex];
      if (row) focusRow(row.key);
    },
    [focusRow, visibleRows],
  );

  const toggleSheet = useCallback((sheetId: string) => {
    setCollapsedSheetIds((previous) => {
      const next = new Set(previous);
      if (next.has(sheetId)) next.delete(sheetId);
      else next.add(sheetId);
      return next;
    });
  }, []);

  const handleRowKeyDown = useCallback(
    (event: ReactKeyboardEvent<HTMLButtonElement>, rowIndex: number) => {
      const row = visibleRows[rowIndex];
      if (!row) return;

      switch (event.key) {
        case "ArrowDown":
          event.preventDefault();
          focusRowAt(rowIndex + 1);
          return;
        case "ArrowUp":
          event.preventDefault();
          focusRowAt(rowIndex - 1);
          return;
        case "Home":
          event.preventDefault();
          focusRowAt(0);
          return;
        case "End":
          event.preventDefault();
          focusRowAt(visibleRows.length - 1);
          return;
        case "ArrowRight":
          if (row.type !== "sheet" || !row.hasChildren) return;
          event.preventDefault();
          if (row.collapsed) {
            toggleSheet(row.sheetId);
            return;
          }
          if (visibleRows[rowIndex + 1]?.sheetId === row.sheetId) {
            focusRowAt(rowIndex + 1);
          }
          return;
        case "ArrowLeft":
          event.preventDefault();
          if (row.type === "sheet") {
            if (row.hasChildren && !row.collapsed) toggleSheet(row.sheetId);
            return;
          }
          focusRow(sheetRowKey(row.sheetId));
          return;
        default:
          return;
      }
    },
    [focusRow, focusRowAt, toggleSheet, visibleRows],
  );

  useEffect(() => {
    if (!autoFocus || initialFocusAppliedRef.current) return;
    const target =
      (fallbackRowKey ? rowRefs.current.get(fallbackRowKey) : null) ??
      closeButtonRef.current;
    if (!target) return;
    initialFocusAppliedRef.current = true;
    if (fallbackRowKey) setFocusedRowKey(fallbackRowKey);
    target.focus({ preventScroll: true });
  }, [autoFocus, fallbackRowKey]);

  // The dialog chrome's task pane: the same card, title row, close button and
  // search field as a dialog; this module keeps the tree's own rows.
  const paneClassName = [
    styles.pane,
    dockSlot ? styles.paneDocked : null,
    spreadsheetPaneClassName(),
  ]
    .filter(Boolean)
    .join(" ");
  const rootClassName = className ? `${paneClassName} ${className}` : paneClassName;

  return (
    <aside
      ref={paneRef}
      // F6 and Task Pane Options > Move / Size hand the keyboard to the pane.
      tabIndex={-1}
      role="complementary"
      aria-label={labels.title}
      className={rootClassName}
      data-testid="spreadsheet-navigation-pane"
      data-navigation-pane="workbook-elements"
    >
      <header className={dialogParts.header}>
        {dockSlot ? (
          <span className={dialogParts.title} dir="auto">
            <OfficeTaskPaneDockedTitle title={labels.title} onClose={onClose} />
          </span>
        ) : (
          <h2 className={dialogParts.title} dir="auto">
            {labels.title}
          </h2>
        )}
        {dockSlot ? (
          <OfficeTaskPaneOptions
            paneId={dockSlot.id}
            onClose={onClose}
            className={dialogParts.iconButton}
            testIdPrefix="spreadsheet-navigation"
          />
        ) : null}
        <button
          ref={closeButtonRef}
          type="button"
          className={dialogParts.close}
          aria-label={labels.close}
          title={labels.close}
          data-testid="spreadsheet-navigation-pane-close"
          onClick={onClose}
        >
          <X size={14} aria-hidden="true" />
        </button>
      </header>

      <div
        className={styles.sheetsHeader}
        data-testid="spreadsheet-navigation-pane-sheets"
      >
        <span dir="auto">{labels.sheets}</span>
        <span className={styles.sheetsCount} aria-hidden="true">
          {sheetGroups.length}
        </span>
      </div>

      <div className={`${dialogParts.inputGroup} ${styles.searchArea}`} role="search">
        <Search size={14} strokeWidth={1.7} aria-hidden="true" />
        <input
          type="search"
          value={searchQuery}
          aria-label={labels.search}
          placeholder={labels.search}
          className={styles.searchInput}
          data-testid="spreadsheet-navigation-pane-search"
          onChange={(event) => setSearchQuery(event.currentTarget.value)}
          onKeyDown={(event) => {
            if (event.key !== "Escape" || searchQuery === "") return;
            event.preventDefault();
            event.stopPropagation();
            setSearchQuery("");
          }}
        />
        {searchQuery ? (
          <button
            type="button"
            aria-label={labels.clearSearch}
            title={labels.clearSearch}
            className={styles.searchClearButton}
            data-testid="spreadsheet-navigation-pane-search-clear"
            onClick={() => setSearchQuery("")}
          >
            <X size={14} strokeWidth={1.8} aria-hidden="true" />
          </button>
        ) : null}
      </div>

      {sheetGroups.length === 0 ? (
        <p
          role="status"
          className={styles.emptyWorkbook}
          data-testid="spreadsheet-navigation-pane-empty"
        >
          {normalizedSearch ? labels.noMatches : labels.emptyWorkbook}
        </p>
      ) : (
        <div
          role="tree"
          aria-label={labels.sheets}
          className={`${styles.tree} ${dialogParts.paneBody}`}
          data-testid="spreadsheet-navigation-pane-tree"
        >
          {sheetGroups.map(
            ({ sheet, elementGroups, elementCount }, sheetIndex) => {
              const rowKey = sheetRowKey(sheet.id);
              const rowIndex = rowIndexByKey.get(rowKey) ?? 0;
              const active = sheet.id === activeSheetId;
              const hasChildren = elementCount > 0;
              const collapsed =
                normalizedSearch === "" &&
                hasChildren &&
                collapsedSheetIds.has(sheet.id);
              const groupId = `${paneId}-sheet-${sheetIndex}-group`;
              const disclosureLabel = formatLabel(
                collapsed ? labels.expandSheet : labels.collapseSheet,
                { sheet: sheet.name },
              );

              return (
                <section
                  key={sheet.id}
                  role="none"
                  className={styles.sheetSection}
                  data-sheet-section={sheet.id}
                >
                  <button
                    ref={(node) => setRowRef(rowKey, node)}
                    type="button"
                    role="treeitem"
                    aria-level={1}
                    aria-posinset={sheetIndex + 1}
                    aria-setsize={sheetGroups.length}
                    aria-current={active ? "page" : undefined}
                    aria-expanded={hasChildren ? !collapsed : undefined}
                    aria-controls={hasChildren ? groupId : undefined}
                    tabIndex={rovingRowKey === rowKey ? 0 : -1}
                    className={styles.sheetRow}
                    title={sheet.name}
                    data-testid={`spreadsheet-navigation-pane-sheet-${sheet.id}`}
                    data-sheet-id={sheet.id}
                    data-active={active ? "true" : "false"}
                    data-expanded={
                      hasChildren ? (collapsed ? "false" : "true") : undefined
                    }
                    onFocus={() => setFocusedRowKey(rowKey)}
                    onKeyDown={(event) => handleRowKeyDown(event, rowIndex)}
                    onClick={(event: ReactMouseEvent<HTMLButtonElement>) => {
                      focusRow(rowKey);
                      const disclosureClicked =
                        hasChildren &&
                        event.target instanceof Element &&
                        event.target.closest(
                          '[data-navigation-disclosure="true"]',
                        ) !== null;
                      if (disclosureClicked) {
                        toggleSheet(sheet.id);
                        return;
                      }
                      onSelectSheet(sheet.id);
                    }}
                  >
                    <span
                      className={styles.disclosureSlot}
                      data-navigation-disclosure={
                        hasChildren ? "true" : undefined
                      }
                      title={hasChildren ? disclosureLabel : undefined}
                      aria-hidden="true"
                    >
                      {hasChildren ? (
                        collapsed ? (
                          <ChevronRight
                            className={styles.collapsedChevron}
                            size={15}
                            strokeWidth={1.8}
                          />
                        ) : (
                          <ChevronDown size={15} strokeWidth={1.8} />
                        )
                      ) : null}
                    </span>
                    <span className={styles.sheetIcon} aria-hidden="true">
                      <SheetIcon size={16} strokeWidth={1.7} />
                    </span>
                    <span className={styles.sheetName} dir="auto">
                      {sheet.name}
                    </span>
                    {elementCount > 0 ? (
                      <span className={styles.countBadge} aria-hidden="true">
                        {elementCount}
                      </span>
                    ) : null}
                  </button>

                  {!collapsed ? (
                    <div
                      id={groupId}
                      role="group"
                      aria-label={sheet.name}
                      className={styles.sheetContents}
                    >
                      {elementCount === 0 ? (
                        <p
                          role="status"
                          className={styles.emptySheet}
                          data-testid="spreadsheet-navigation-pane-sheet-empty"
                          data-sheet-id={sheet.id}
                        >
                          {formatLabel(labels.emptySheet, {
                            sheet: sheet.name,
                          })}
                        </p>
                      ) : (
                        elementGroups.map((group, groupIndex) => {
                          const categoryId = `${paneId}-sheet-${sheetIndex}-category-${groupIndex}`;
                          return (
                            <div
                              key={group.kind}
                              role="none"
                              className={styles.category}
                              data-testid="spreadsheet-navigation-pane-category"
                              data-kind={group.kind}
                            >
                              <div
                                id={categoryId}
                                role="presentation"
                                className={styles.categoryHeader}
                              >
                                <span dir="auto">{group.label}</span>
                                <span
                                  className={styles.categoryCount}
                                  aria-hidden="true"
                                >
                                  {group.elements.length}
                                </span>
                              </div>
                              <div
                                role="group"
                                aria-labelledby={categoryId}
                                className={styles.elementList}
                              >
                                {group.elements.map((element, elementIndex) => {
                                  const selection = {
                                    sheetId: sheet.id,
                                    kind: group.kind,
                                    id: element.id,
                                  } as const;
                                  const elementKey = elementRowKey(selection);
                                  const elementRowIndex =
                                    rowIndexByKey.get(elementKey) ?? 0;
                                  const selected = isSameSelection(
                                    selectedElement,
                                    selection,
                                  );
                                  const Icon = group.Icon;
                                  const visibilityAction = element.actions?.find(
                                    (action) =>
                                      action.id === "hide" ||
                                      action.id === "show",
                                  );

                                  return (
                                    <div
                                      key={elementKey}
                                      role="none"
                                      className={styles.elementRowShell}
                                      data-has-visibility={
                                        visibilityAction ? "true" : "false"
                                      }
                                    >
                                      <button
                                        ref={(node) =>
                                          setRowRef(elementKey, node)
                                        }
                                        type="button"
                                        role="treeitem"
                                        aria-level={2}
                                        aria-posinset={elementIndex + 1}
                                        aria-setsize={group.elements.length}
                                        aria-selected={selected}
                                        tabIndex={
                                          rovingRowKey === elementKey ? 0 : -1
                                        }
                                        className={styles.elementRow}
                                        data-testid={`spreadsheet-navigation-pane-${group.kind}-${element.id}`}
                                        data-sheet-id={sheet.id}
                                        data-kind={group.kind}
                                        data-element-id={element.id}
                                        data-selected={
                                          selected ? "true" : "false"
                                        }
                                        data-hidden={
                                          element.hidden ? "true" : "false"
                                        }
                                        onFocus={() =>
                                          setFocusedRowKey(elementKey)
                                        }
                                        onContextMenu={(event) => {
                                          if (!element.actions?.length) return;
                                          event.preventDefault();
                                          setFocusedRowKey(elementKey);
                                          setActionMenuKey(elementKey);
                                        }}
                                        onKeyDown={(event) => {
                                          if (
                                            element.actions?.length &&
                                            (event.key === "ContextMenu" ||
                                              (event.key === "F10" &&
                                                event.shiftKey))
                                          ) {
                                            event.preventDefault();
                                            setActionMenuKey(elementKey);
                                            return;
                                          }
                                          handleRowKeyDown(
                                            event,
                                            elementRowIndex,
                                          );
                                        }}
                                        onClick={() => {
                                          focusRow(elementKey);
                                          onSelectElement(selection);
                                        }}
                                      >
                                        <span
                                          className={styles.elementIcon}
                                          aria-hidden="true"
                                        >
                                          <Icon size={15} strokeWidth={1.7} />
                                        </span>
                                        <span className={styles.elementText}>
                                          <span
                                            className={styles.elementName}
                                            dir="auto"
                                          >
                                            {element.name}
                                          </span>
                                          {element.detail ? (
                                            <span
                                              className={styles.elementDetail}
                                              dir="auto"
                                            >
                                              {element.detail}
                                            </span>
                                          ) : null}
                                        </span>
                                      </button>
                                      {visibilityAction ? (
                                        <button
                                          type="button"
                                          aria-label={visibilityAction.label}
                                          aria-pressed={!element.hidden}
                                          title={
                                            visibilityAction.disabledReason ??
                                            visibilityAction.label
                                          }
                                          disabled={Boolean(
                                            visibilityAction.disabledReason,
                                          )}
                                          className={
                                            styles.elementVisibilityButton
                                          }
                                          data-testid={`spreadsheet-navigation-pane-visibility-${group.kind}-${element.id}`}
                                          onClick={() =>
                                            onElementAction?.(
                                              selection,
                                              visibilityAction.id,
                                            )
                                          }
                                        >
                                          {element.hidden ? (
                                            <EyeOff
                                              size={16}
                                              strokeWidth={1.7}
                                              aria-hidden="true"
                                            />
                                          ) : (
                                            <Eye
                                              size={16}
                                              strokeWidth={1.7}
                                              aria-hidden="true"
                                            />
                                          )}
                                        </button>
                                      ) : null}
                                      {element.actions?.length ? (
                                        <DropdownMenu
                                          modal={false}
                                          open={actionMenuKey === elementKey}
                                          onOpenChange={(open) =>
                                            setActionMenuKey(
                                              open ? elementKey : null,
                                            )
                                          }
                                        >
                                          <DropdownMenuTrigger asChild>
                                            <button
                                              type="button"
                                              aria-label={formatLabel(
                                                labels.elementActions,
                                                { element: element.name },
                                              )}
                                              title={formatLabel(
                                                labels.elementActions,
                                                { element: element.name },
                                              )}
                                              className={
                                                styles.elementActionButton
                                              }
                                              data-testid={`spreadsheet-navigation-pane-more-${group.kind}-${element.id}`}
                                            >
                                              <MoreHorizontal
                                                size={17}
                                                aria-hidden="true"
                                              />
                                            </button>
                                          </DropdownMenuTrigger>
                                          <DropdownMenuContent
                                            side="left"
                                            align="start"
                                            // A row's right-click menu: the
                                            // one right-click size.
                                            className="docs-cupertino-menu office-context-menu"
                                            onCloseAutoFocus={(event) => {
                                              event.preventDefault();
                                              rowRefs.current
                                                .get(elementKey)
                                                ?.focus({ preventScroll: true });
                                            }}
                                          >
                                            {element.actions.map((action) => (
                                              <DropdownMenuItem
                                                key={action.id}
                                                disabled={Boolean(
                                                  action.disabledReason,
                                                )}
                                                aria-label={
                                                  action.disabledReason
                                                    ? `${action.label}: ${action.disabledReason}`
                                                    : action.label
                                                }
                                                title={
                                                  action.disabledReason ??
                                                  undefined
                                                }
                                                className={
                                                  action.destructive
                                                    ? "text-red-700 focus:text-red-700"
                                                    : undefined
                                                }
                                                data-testid={`spreadsheet-navigation-pane-action-${group.kind}-${element.id}-${action.id}`}
                                                onSelect={() =>
                                                  onElementAction?.(
                                                    selection,
                                                    action.id,
                                                  )
                                                }
                                              >
                                                {action.id === "rename" ? (
                                                  <Pencil aria-hidden="true" />
                                                ) : action.id === "delete" ? (
                                                  <Trash2 aria-hidden="true" />
                                                ) : action.id === "hide" ? (
                                                  <EyeOff aria-hidden="true" />
                                                ) : (
                                                  <Eye aria-hidden="true" />
                                                )}
                                                <span>{action.label}</span>
                                              </DropdownMenuItem>
                                            ))}
                                          </DropdownMenuContent>
                                        </DropdownMenu>
                                      ) : null}
                                    </div>
                                  );
                                })}
                              </div>
                            </div>
                          );
                        })
                      )}
                    </div>
                  ) : null}
                </section>
              );
            },
          )}
        </div>
      )}
    </aside>
  );
}
