"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Design > Document Formatting's style-set gallery, as Word draws it: one
 * thumbnail per style set (the document's own first, then Word's built-in
 * sets), each a small page with its Title and Heading 1 in the set's styles
 * under the document's CURRENT theme, so the thumbnails follow Colors and
 * Fonts. The strip sits in the ribbon; its More menu lists every set with
 * Reset to the Default Style Set.
 */

import { useState, type CSSProperties, type MouseEvent } from "react";

import { OfficeHomeMenuTrigger } from "@mrnote/core/web/components/office-ribbon/OfficeHomeRibbon.tsx";
import { OfficeRibbonGallery } from "@mrnote/core/web/components/office-ribbon/OfficeRibbonGallery.tsx";
import {
  OfficeGalleryMenuContent,
  OfficeGalleryMenuFooter,
  OfficeGalleryMenuGrid,
  OfficeGalleryMenuSection,
  officeGalleryMenuStyles,
} from "@mrnote/core/web/components/office-ribbon/OfficeGalleryMenu.tsx";
import {
  DropdownMenu,
  DropdownMenuItem,
  DropdownMenuTrigger,
} from "@mrnote/core/web/components/ui/dropdown-menu.tsx";
import { PointerContextMenu } from "@mrnote/core/web/components/ui/PointerContextMenu.tsx";
import type { DocumentDesign } from "@mrnote/core/web/lib/document-design.ts";
import {
  customStyleSetMatches,
  documentDesignWithCustomStyleSet,
  type DocumentCustomStyleSet,
} from "@mrnote/core/web/lib/document-custom-style-sets.ts";
import { documentDesignStylePreviewCss } from "@mrnote/core/web/lib/document-design-preview.ts";
import { useOfficeData } from "@mrnote/core/web/lib/office-data/use-office-data.ts";
import { officeThemeLabel } from "@mrnote/core/web/lib/office-theme-catalog.ts";
import {
  WORD_DEFAULT_STYLE_SET_ID,
  WORD_STYLE_SET_IDS,
  wordStyleSets,
  type WordStyleSet,
} from "@mrnote/core/web/lib/office-word-style-sets.ts";

import { DocsHomeIcon } from "@mrnote/core/web/components/console/editor/DocumentHomeRibbonIcons.tsx";
import styles from "./DocumentStyleSetGallery.module.css";

/** The gallery's sets: Word lists its built-in sets, the default being "This Document"'s reset. */
function builtInSets(sets: readonly WordStyleSet[]): readonly WordStyleSet[] {
  return sets.filter((set) => set.id !== WORD_DEFAULT_STYLE_SET_ID);
}

/** Tiles drawn while Word's style sets (served Office data) are on their way. */
const PLACEHOLDER_IDS = WORD_STYLE_SET_IDS.filter((id) => id !== WORD_DEFAULT_STYLE_SET_ID);

export interface DocumentStyleSetGalleryProps {
  readonly design: DocumentDesign;
  readonly disabled: boolean;
  readonly locale: string;
  readonly labels: {
    readonly gallery: string;
    readonly thisDocument: string;
    readonly builtIn: string;
    readonly reset: string;
    readonly title: string;
    readonly heading1: string;
    /** A sentence of body text for the thumbnails' paragraph. */
    readonly sample: string;
  };
  readonly onRememberSelection: () => void;
  readonly onSelect: (styleSetId: string) => void;
  readonly onReset: () => void;
  /**
   * The user's saved sets (Save as a New Style Set), listed under Custom
   * between This Document and the built-in sets, as Word lists them;
   * right-click > Delete removes one.
   */
  readonly custom?: {
    readonly sets: readonly DocumentCustomStyleSet[];
    readonly label: string;
    readonly deleteLabel: string;
    readonly onSelect: (entry: DocumentCustomStyleSet) => void;
    readonly onDelete: (entry: DocumentCustomStyleSet) => void;
  };
  /** The More menu's Save as a New Style Set... */
  readonly saveAs?: { readonly label: string; readonly disabled?: boolean; readonly onSelect: () => void };
}

function StyleSetThumbnail({
  design,
  labels,
}: {
  design: DocumentDesign;
  labels: DocumentStyleSetGalleryProps["labels"];
}) {
  // Word draws each set as a small page: the Title, a Heading 1 and a
  // paragraph of body text, each in the set's own look. The page is laid out
  // at a readable size and scaled down whole (transform, never zoom: Safari's
  // minimum font size would clamp the tiny body text).
  const title = documentDesignStylePreviewCss(design, "title", { minPx: 20, maxPx: 34 });
  const heading = documentDesignStylePreviewCss(design, "heading1", { minPx: 13, maxPx: 19 });
  const body = documentDesignStylePreviewCss(design, "normal", { minPx: 10, maxPx: 12 });
  if (!title || !heading || !body) return <span className={styles.page} data-placeholder="true" aria-hidden="true" />;
  return (
    <span className={styles.page} aria-hidden="true">
      <span className={styles.sheet}>
        <span className={styles.title} style={title}>
          {labels.title}
        </span>
        <span className={styles.heading} style={heading}>
          {labels.heading1}
        </span>
        <span className={styles.body} style={body as CSSProperties}>
          {labels.sample}
        </span>
      </span>
    </span>
  );
}

export function DocumentStyleSetGallery({
  design,
  disabled,
  locale,
  labels,
  onRememberSelection,
  onSelect,
  onReset,
  custom,
  saveAs,
}: DocumentStyleSetGalleryProps) {
  const [customMenu, setCustomMenu] = useState<{ x: number; y: number; entry: DocumentCustomStyleSet } | null>(null);
  const customSets = custom?.sets ?? [];
  const currentCustom = customSets.find((entry) => customStyleSetMatches(design, entry)) ?? null;
  // Word's style sets are served Office data: asked for here, drawn when here.
  useOfficeData("word-style-sets");
  const sets = wordStyleSets();
  const label = (set: WordStyleSet) => officeThemeLabel(set.label, locale);
  // Word marks the set last applied; until one is, "This Document".
  const current = currentCustom || design.styleSet === "default" ? null : design.styleSet;
  const tile = (set: WordStyleSet | null, variant: "strip" | "menu") => {
    const id = set?.id ?? "document";
    const name = set ? label(set) : labels.thisDocument;
    const selected = set ? current === set.id : current === null;
    const preview = set ? { ...design, styleSet: set.id } : design;
    const testId =
      variant === "strip"
        ? `document-toolbar-ribbon-design-style-set-${id}`
        : `document-toolbar-ribbon-design-style-set-menu-${id}`;
    const body = (
      <>
        <StyleSetThumbnail design={preview} labels={labels} />
        {variant === "menu" ? <span className={officeGalleryMenuStyles.name}>{name}</span> : null}
      </>
    );
    if (variant === "menu") {
      return (
        <DropdownMenuItem
          asChild
          key={id}
          disabled={disabled}
          onSelect={() => (set ? onSelect(set.id) : undefined)}
        >
          <button
            type="button"
            className={officeGalleryMenuStyles.tile}
            data-selected={selected ? "true" : undefined}
            data-testid={testId}
            aria-pressed={selected}
            title={name}
          >
            {body}
          </button>
        </DropdownMenuItem>
      );
    }
    return (
      <button
        key={id}
        type="button"
        className={styles.tile}
        data-selected={selected ? "true" : undefined}
        data-testid={testId}
        aria-pressed={selected}
        disabled={disabled}
        title={name}
        aria-label={name}
        onPointerDown={onRememberSelection}
        onMouseDown={(event) => event.preventDefault()}
        onClick={() => (set ? onSelect(set.id) : undefined)}
      >
        {body}
      </button>
    );
  };
  const customTile = (entry: DocumentCustomStyleSet, variant: "strip" | "menu") => {
    const selected = currentCustom?.id === entry.id;
    const testId =
      variant === "strip"
        ? `document-toolbar-ribbon-design-style-set-custom-${entry.id}`
        : `document-toolbar-ribbon-design-style-set-menu-custom-${entry.id}`;
    const openMenu = (event: MouseEvent) => {
      if (!custom || disabled) return;
      event.preventDefault();
      setCustomMenu({ x: event.clientX, y: event.clientY, entry });
    };
    const body = (
      <>
        <StyleSetThumbnail design={documentDesignWithCustomStyleSet(design, entry)} labels={labels} />
        {variant === "menu" ? <span className={officeGalleryMenuStyles.name}>{entry.name}</span> : null}
      </>
    );
    if (variant === "menu") {
      return (
        <DropdownMenuItem asChild key={entry.id} disabled={disabled} onSelect={() => custom?.onSelect(entry)}>
          <button
            type="button"
            className={officeGalleryMenuStyles.tile}
            data-selected={selected ? "true" : undefined}
            data-testid={testId}
            aria-pressed={selected}
            title={entry.name}
            onContextMenu={openMenu}
          >
            {body}
          </button>
        </DropdownMenuItem>
      );
    }
    return (
      <button
        key={entry.id}
        type="button"
        className={styles.tile}
        data-selected={selected ? "true" : undefined}
        data-testid={testId}
        aria-pressed={selected}
        disabled={disabled}
        title={entry.name}
        aria-label={entry.name}
        onPointerDown={onRememberSelection}
        onMouseDown={(event) => event.preventDefault()}
        onClick={() => custom?.onSelect(entry)}
        onContextMenu={openMenu}
      >
        {body}
      </button>
    );
  };
  const contextMenu =
    customMenu && custom ? (
      <PointerContextMenu
        x={customMenu.x}
        y={customMenu.y}
        label={customMenu.entry.name}
        entries={[
          {
            id: "delete",
            label: custom.deleteLabel,
            testId: "document-toolbar-ribbon-design-style-set-custom-delete",
            onSelect: () => custom.onDelete(customMenu.entry),
          },
        ]}
        onClose={() => setCustomMenu(null)}
        testId="document-toolbar-ribbon-design-style-set-custom-menu"
      />
    ) : null;
  return (
    <>
    {contextMenu}
    <OfficeRibbonGallery
      ariaLabel={labels.gallery}
      className={styles.gallery}
      testId="document-toolbar-ribbon-design-style-sets"
      viewport={
        <div className={styles.strip} data-testid="document-toolbar-ribbon-design-style-set-strip">
          {tile(null, "strip")}
          {sets ? customSets.map((entry) => customTile(entry, "strip")) : null}
          {sets
            ? builtInSets(sets).map((set) => tile(set, "strip"))
            : PLACEHOLDER_IDS.map((id) => (
                <button
                  key={id}
                  type="button"
                  className={styles.tile}
                  disabled
                  aria-label={labels.gallery}
                  data-testid={`document-toolbar-ribbon-design-style-set-${id}`}
                >
                  <span className={styles.page} data-placeholder="true" aria-hidden="true" />
                </button>
              ))}
        </div>
      }
      more={
        <DropdownMenu>
          <DropdownMenuTrigger asChild>
            <OfficeHomeMenuTrigger
              geometry="gallery"
              icon={<DocsHomeIcon name="galleryMore" size={13} />}
              label={labels.gallery}
              disabled={disabled}
              onPointerDown={onRememberSelection}
              testId="document-toolbar-ribbon-design-style-set-more"
              variant="icon"
            />
          </DropdownMenuTrigger>
          <OfficeGalleryMenuContent data-testid="document-toolbar-ribbon-design-style-set-menu">
            <OfficeGalleryMenuSection>{labels.thisDocument}</OfficeGalleryMenuSection>
            <OfficeGalleryMenuGrid columns={4} tileWidth={92} className={styles.menuGrid}>
              {tile(null, "menu")}
            </OfficeGalleryMenuGrid>
            {custom && sets && customSets.length ? (
              <>
                <OfficeGalleryMenuSection>{custom.label}</OfficeGalleryMenuSection>
                <OfficeGalleryMenuGrid columns={4} tileWidth={92} className={styles.menuGrid}>
                  {customSets.map((entry) => customTile(entry, "menu"))}
                </OfficeGalleryMenuGrid>
              </>
            ) : null}
            <OfficeGalleryMenuSection>{labels.builtIn}</OfficeGalleryMenuSection>
            <OfficeGalleryMenuGrid columns={4} tileWidth={92} className={styles.menuGrid}>
              {sets ? builtInSets(sets).map((set) => tile(set, "menu")) : null}
            </OfficeGalleryMenuGrid>
            <OfficeGalleryMenuFooter>
              <DropdownMenuItem
                disabled={disabled}
                data-testid="document-toolbar-ribbon-design-style-set-reset"
                onSelect={onReset}
              >
                {labels.reset}
              </DropdownMenuItem>
              {saveAs ? (
                <DropdownMenuItem
                  disabled={disabled || saveAs.disabled}
                  data-testid="document-toolbar-ribbon-design-style-set-save-as"
                  onSelect={saveAs.onSelect}
                >
                  {saveAs.label}
                </DropdownMenuItem>
              ) : null}
            </OfficeGalleryMenuFooter>
          </OfficeGalleryMenuContent>
        </DropdownMenu>
      }
    />
    </>
  );
}
