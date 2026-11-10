"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { isDoubleTap } from "../../lib/double-tap";
import type {
  CSSProperties,
  KeyboardEvent as ReactKeyboardEvent,
  MouseEvent as ReactMouseEvent,
  ReactNode,
} from "react";
import { ChevronRight, Eye, EyeOff } from "./icons/houseIcons";

import styles from "./OfficeSelectionPaneList.module.css";
import { ButtonUnlessRenaming } from "../ui/ButtonUnlessRenaming";

/**
 * The Selection Pane's object list (Word / PowerPoint / Excel: Home >
 * Select > Selection Pane), shared by the editors: a list of rows, each a
 * type icon, the object's name (renamed in place), its stacking ordinal, an
 * eye that shows or hides it and, for a reorderable object, the Bring /
 * Send rail. The pane's frame is the shared OfficeTaskPane; what the rows
 * list and what their commands do stay with each editor.
 */

/** The list's class names, for a caller's own parts beside the rows. */
export const officeSelectionPaneListClassNames = styles;

export function OfficeSelectionPaneList({
  isEmpty,
  emptyLabel,
  emptyTestId,
  listLabel,
  children,
}: {
  readonly isEmpty: boolean;
  readonly emptyLabel: ReactNode;
  readonly emptyTestId?: string;
  /** The list's accessible name. */
  readonly listLabel: string;
  /** The OfficeSelectionPaneRow items. */
  readonly children?: ReactNode;
}): React.JSX.Element {
  return (
    <div className={styles.body}>
      {isEmpty ? (
        <p className={styles.empty} data-testid={emptyTestId}>
          {emptyLabel}
        </p>
      ) : (
        // The explicit role keeps the list a list for VoiceOver, which drops
        // list semantics from a `list-style: none` list.
        <ul className={styles.list} role="list" aria-label={listLabel}>
          {children}
        </ul>
      )}
    </div>
  );
}

export interface OfficeSelectionPaneRename {
  readonly draft: string;
  readonly inputLabel: string;
  readonly testId?: string;
  readonly maxLength?: number;
  readonly onChange: (draft: string) => void;
  /** Enter or leaving the field. */
  readonly onCommit: () => void;
  /** Escape. */
  readonly onCancel: () => void;
}

export interface OfficeSelectionPaneVisibility {
  readonly hidden: boolean;
  readonly disabled?: boolean;
  readonly title: string;
  readonly ariaLabel: string;
  readonly testId?: string;
  readonly onToggle: () => void;
}

export interface OfficeSelectionPaneReorderAction {
  readonly key: string;
  readonly label: string;
  readonly icon: ReactNode;
  readonly disabled?: boolean;
  readonly testId?: string;
  readonly onClick: () => void;
}

/** A group row's disclosure: PowerPoint's pane expands a group (Right arrow)
 * to list its objects and collapses it (Left arrow). */
export interface OfficeSelectionPaneExpansion {
  readonly expanded: boolean;
  readonly label: string;
  readonly testId?: string;
  readonly onToggle: () => void;
}

export interface OfficeSelectionPaneRowProps {
  /** The row button, for the pane's roving focus. */
  readonly rowRef?: (element: HTMLButtonElement | null) => void;
  readonly icon: ReactNode;
  readonly label: string;
  /** The row's hover text, when it says more than the name. */
  readonly title?: string;
  /** The name's own test id (a click on it selects, a double-click renames). */
  readonly labelTestId?: string;
  /** The row's place in the list, for a list the editor pages or groups. */
  readonly position?: { readonly index: number; readonly setSize: number };
  /** A screen-reader description of the row (its order, why it is locked). */
  readonly description?: { readonly id: string; readonly text: string } | null;
  readonly ordinal?: ReactNode;
  readonly selected: boolean;
  /** The primary (last clicked) selection: aria-current. */
  readonly current?: boolean;
  readonly hidden?: boolean;
  /** Widens the row's grid for the reorder rail column. */
  readonly reorderable?: boolean;
  readonly tabIndex?: number;
  readonly testId?: string;
  /** Extra data-* attributes on the row button (data-object-type). */
  readonly rowData?: Readonly<Record<`data-${string}`, string | undefined>>;
  readonly onFocus?: () => void;
  readonly onClick: (event: ReactMouseEvent<HTMLButtonElement>) => void;
  readonly onKeyDown?: (event: ReactKeyboardEvent<HTMLButtonElement>) => void;
  /** Double-click on the name (start renaming). */
  readonly onLabelDoubleClick?: (event: ReactMouseEvent<HTMLSpanElement>) => void;
  /** While set, the name is an input. */
  readonly rename?: OfficeSelectionPaneRename | null;
  /** The eye; null keeps its column empty (an object that cannot hide). */
  readonly visibility?: OfficeSelectionPaneVisibility | null;
  /** The Bring / Send rail; null or empty omits it. */
  readonly reorder?: {
    readonly label: string;
    readonly actions: readonly OfficeSelectionPaneReorderAction[];
  } | null;
  /** Nesting depth in a pane that lists groups' objects (0 = the slide's own);
   * a pane that sets it on every row keeps room for the disclosure arrow. */
  readonly level?: number;
  /** The group row's expand / collapse arrow. */
  readonly expansion?: OfficeSelectionPaneExpansion | null;
}

export function OfficeSelectionPaneRow({
  rowRef,
  icon,
  label,
  title,
  labelTestId,
  position,
  description,
  ordinal,
  selected,
  current,
  hidden,
  reorderable,
  tabIndex,
  testId,
  rowData,
  onFocus,
  onClick,
  onKeyDown,
  onLabelDoubleClick,
  rename,
  visibility,
  reorder,
  level,
  expansion,
}: OfficeSelectionPaneRowProps): React.JSX.Element {
  return (
    <li
      className={styles.item}
      aria-posinset={position?.index}
      aria-setsize={position?.setSize}
      data-hidden={hidden || undefined}
      data-reorderable={reorderable ? "true" : undefined}
    >
      {description ? (
        <span id={description.id} className="sr-only">
          {description.text}
        </span>
      ) : null}
      <ButtonUnlessRenaming
        renaming={Boolean(rename)}
        ref={rowRef}
        type="button"
        className={styles.row}
        data-selected={selected || undefined}
        {...rowData}
        data-testid={testId}
        aria-current={current ? "true" : undefined}
        aria-pressed={selected}
        aria-expanded={expansion ? expansion.expanded : undefined}
        aria-describedby={description?.id}
        data-level={level === undefined ? undefined : String(level)}
        style={level === undefined ? undefined : ({ "--pane-indent": `${level * 18}px` } as CSSProperties)}
        tabIndex={tabIndex}
        title={title ?? label}
        onFocus={onFocus}
        onClick={onClick}
        onKeyDown={(event) => {
          if (expansion && !rename) {
            if (event.key === "ArrowRight" && !expansion.expanded) {
              event.preventDefault();
              event.stopPropagation();
              expansion.onToggle();
              return;
            }
            if (event.key === "ArrowLeft" && expansion.expanded) {
              event.preventDefault();
              event.stopPropagation();
              expansion.onToggle();
              return;
            }
          }
          onKeyDown?.(event);
        }}
      >
        {expansion ? (
          // Not a second button inside the row: the arrow keys reach it from
          // the keyboard, and a click on the arrow only toggles.
          <span
            className={styles.disclosure}
            data-expanded={expansion.expanded ? "true" : "false"}
            data-testid={expansion.testId}
            title={expansion.label}
            aria-hidden
            onClick={(event) => {
              event.stopPropagation();
              expansion.onToggle();
            }}
          >
            <ChevronRight size={12} aria-hidden />
          </span>
        ) : null}
        <span className={styles.icon}>{icon}</span>
        {rename ? (
          <input
            className={styles.label}
            autoFocus
            value={rename.draft}
            maxLength={rename.maxLength ?? 255}
            aria-label={rename.inputLabel}
            data-testid={rename.testId}
            onClick={(event) => event.stopPropagation()}
            onChange={(event) => rename.onChange(event.target.value)}
            onBlur={rename.onCommit}
            onKeyDown={(event) => {
              event.stopPropagation();
              if (event.key === "Enter") {
                event.preventDefault();
                rename.onCommit();
              } else if (event.key === "Escape") {
                event.preventDefault();
                rename.onCancel();
              }
            }}
          />
        ) : (
          <span
            className={styles.label}
            data-testid={labelTestId}
            onDoubleClick={onLabelDoubleClick}
            // A finger's double tap renames as a double-click does (an iPad
            // does not dependably send dblclick).
            onClick={
              onLabelDoubleClick
                ? (event) => {
                    if (isDoubleTap(event.nativeEvent)) onLabelDoubleClick(event);
                  }
                : undefined
            }
          >
            {label}
          </span>
        )}
        <span className={styles.ordinal}>{ordinal}</span>
      </ButtonUnlessRenaming>
      {visibility ? (
        <button
          type="button"
          className={styles.visibilityButton}
          disabled={visibility.disabled}
          aria-pressed={!visibility.hidden}
          title={visibility.title}
          aria-label={visibility.ariaLabel}
          data-testid={visibility.testId}
          data-hidden={visibility.hidden ? "true" : "false"}
          onClick={visibility.onToggle}
        >
          {visibility.hidden ? (
            <EyeOff size={14} aria-hidden />
          ) : (
            <Eye size={14} aria-hidden />
          )}
        </button>
      ) : (
        <span aria-hidden />
      )}
      {reorder && reorder.actions.length > 0 ? (
        <span className={styles.reorderRail} aria-label={reorder.label}>
          {reorder.actions.map((action) => (
            <button
              key={action.key}
              type="button"
              className={styles.reorderButton}
              disabled={action.disabled}
              title={action.label}
              aria-label={action.label}
              data-testid={action.testId}
              onClick={action.onClick}
            >
              {action.icon}
            </button>
          ))}
        </span>
      ) : null}
    </li>
  );
}
