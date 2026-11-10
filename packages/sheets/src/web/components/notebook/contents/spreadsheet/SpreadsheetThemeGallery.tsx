"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { useId } from "react";

import { OfficeSemanticIcon } from "@mrnote/core/web/components/office-ribbon/icons/OfficeSemanticIcon.tsx";
import {
  DropdownMenuItem,
  DropdownMenuRadioGroup,
  DropdownMenuRadioItem,
} from "@mrnote/core/web/components/ui/dropdown-menu.tsx";
import {
  officeThemeLabel,
  type OfficeEffectScheme,
} from "@mrnote/core/web/lib/office-theme-catalog.ts";
import { paletteById, tintColor } from "@mrnote/core/web/lib/spreadsheet/color-palette.ts";
import {
  SPREADSHEET_THEME_EFFECTS,
  SPREADSHEET_WORKBOOK_THEMES,
  type SpreadsheetWorkbookThemePreset,
} from "@mrnote/core/web/lib/spreadsheet/workbook-theme.ts";

import { SpreadsheetPageLayoutIcon } from "./SpreadsheetPageLayoutRibbonIcons";
import styles from "./SpreadsheetThemeGallery.module.css";
import {
  officeGalleryMenuGridStyle,
  officeGalleryMenuStyles,
} from "@mrnote/core/web/components/office-ribbon/OfficeGalleryMenu.tsx";

/** Office's Themes / Effects sheets: four 76 px tiles a row (the shared gallery menu). */
const THEME_GRID = officeGalleryMenuGridStyle(4, 76);

export interface SpreadsheetThemeGalleryProps {
  readonly disabled?: boolean;
  readonly locale?: string;
  readonly translate: (key: string) => string;
  readonly onBrowse: () => void;
  readonly onSaveCurrent: () => void;
  readonly onValueChange: (themeId: string) => void;
  readonly themes?: readonly SpreadsheetWorkbookThemePreset[];
  readonly value: string;
  /**
   * Word's Themes menu also offers Reset to Theme from Template, above
   * Browse and Save; a host that has it names it here.
   */
  readonly reset?: { readonly label: string; readonly onSelect: () => void };
  /**
   * Word's Custom section above Office's themes: the themes the user saved
   * (Save Current Theme). A host that keeps such a list names it here.
   */
  readonly custom?: {
    readonly label: string;
    readonly items: readonly {
      readonly id: string;
      readonly name: string;
      /** The theme's light background and its six accents, for the tile. */
      readonly background: string;
      readonly accents: readonly string[];
    }[];
    readonly onSelect: (id: string) => void;
    /** Right-click on a Custom theme (Word offers Delete there). */
    readonly onContextMenu?: (id: string, x: number, y: number) => void;
  };
  /** The gallery is Office's one Themes list; each host names its own test ids. */
  readonly testId?: string;
  readonly itemTestIdPrefix?: string;
}

/**
 * A theme's tile as Office draws one: a folded page on the theme's ground,
 * the six accents along its foot and the theme's decoration at the edge. It
 * is drawn from the theme's own colors, not from Office's pictures.
 */
function ThemeTile({ theme }: { theme: SpreadsheetWorkbookThemePreset }) {
  const palette = paletteById(theme.selection.colorId);
  const { background, decoration } = theme.tile;
  return (
    <svg
      aria-hidden="true"
      className={styles.preview}
      data-testid={`spreadsheet-theme-card-preview-${theme.id}`}
      viewBox="0 0 64 48"
    >
      <rect x="0.5" y="0.5" width="63" height="47" fill={background} stroke="#8f96a0" />
      {decoration ? (
        <path d="M50 0.5 L63.5 0.5 L63.5 47.5 L56 47.5 Z" fill={decoration} />
      ) : null}
      <path
        d="M5.5 5.5 H40.5 L47.5 12.5 V42.5 H5.5 Z"
        fill="#ffffff"
        stroke="#9aa1aa"
      />
      <path d="M40.5 5.5 V12.5 H47.5" fill="#eef0f2" stroke="#9aa1aa" />
      {palette.accents.map((color, index) => (
        <rect
          key={`${color}-${index}`}
          data-theme-preview-color={index}
          x={9 + index * 6}
          y={33}
          width={5}
          height={5}
          fill={color}
        />
      ))}
    </svg>
  );
}

export function SpreadsheetThemeGallery({
  disabled = false,
  locale,
  translate,
  onBrowse,
  onSaveCurrent,
  onValueChange,
  themes = SPREADSHEET_WORKBOOK_THEMES,
  value,
  reset,
  custom,
  testId = "spreadsheet-theme-gallery",
  itemTestIdPrefix = "spreadsheet-ribbon-page-layout-theme-",
}: SpreadsheetThemeGalleryProps) {
  return (
    <div data-testid={testId}>
      {custom && custom.items.length ? (
        <>
          <div className={`${officeGalleryMenuStyles.section} px-2 py-1.5 font-semibold`} data-testid={`${testId}-custom-heading`}>
            {custom.label}
          </div>
          <div className={officeGalleryMenuStyles.grid}
          style={THEME_GRID} role="group" aria-label={custom.label}>
            {custom.items.map((item) => (
              <DropdownMenuItem
                key={item.id}
                aria-label={item.name}
                className={officeGalleryMenuStyles.tile}
                data-testid={`${testId}-custom-${item.id}`}
                disabled={disabled}
                title={item.name}
                onSelect={() => custom.onSelect(item.id)}
                onContextMenu={(event) => {
                  if (!custom.onContextMenu || disabled) return;
                  event.preventDefault();
                  custom.onContextMenu(item.id, event.clientX, event.clientY);
                }}
              >
                <svg aria-hidden="true" className={styles.preview} viewBox="0 0 64 48">
                  <rect x="0.5" y="0.5" width="63" height="47" fill={item.background} stroke="#8f96a0" />
                  {item.accents.slice(0, 6).map((accent, index) => (
                    <rect key={index} x={6 + index * 9} y={36} width={8} height={7} fill={accent} />
                  ))}
                </svg>
                <span className={officeGalleryMenuStyles.name}>{item.name}</span>
              </DropdownMenuItem>
            ))}
          </div>
        </>
      ) : null}
      <div
        className={`${officeGalleryMenuStyles.section} px-2 py-1.5 font-semibold`}
        data-testid="spreadsheet-theme-gallery-heading"
      >
        Office
      </div>
      <div
        data-testid="spreadsheet-theme-gallery-scroll"
      >
        <DropdownMenuRadioGroup
          className={officeGalleryMenuStyles.grid}
          style={THEME_GRID}
          value={value}
          onValueChange={onValueChange}
          data-testid="spreadsheet-theme-gallery-grid"
        >
          {themes.map((theme) => {
            const label = officeThemeLabel(theme.label, locale);
            return (
              <DropdownMenuRadioItem
                key={theme.id}
                value={theme.id}
                aria-label={label}
                className={officeGalleryMenuStyles.tile}
                data-selected={value === theme.id ? "true" : undefined}
                data-testid={`${itemTestIdPrefix}${theme.id}`}
                disabled={disabled}
                title={label}
              >
                <ThemeTile theme={theme} />
                <span className={officeGalleryMenuStyles.name}>{label}</span>
              </DropdownMenuRadioItem>
            );
          })}
        </DropdownMenuRadioGroup>
      </div>
      <div className={styles.footer}>
        {reset ? (
          <DropdownMenuItem
            className={styles.footerAction}
            data-testid={`${testId}-reset`}
            disabled={disabled}
            onSelect={reset.onSelect}
          >
            <SpreadsheetPageLayoutIcon name="themes" size={18} />
            <span>{reset.label}</span>
          </DropdownMenuItem>
        ) : null}
        <DropdownMenuItem
          className={styles.footerAction}
          data-testid="spreadsheet-theme-browse"
          disabled={disabled}
          onSelect={onBrowse}
        >
          <SpreadsheetPageLayoutIcon name="themes" size={18} />
          <span>{translate("spreadsheet.themeGallery.browse")}</span>
        </DropdownMenuItem>
        <DropdownMenuItem
          className={styles.footerAction}
          data-testid="spreadsheet-theme-save-current"
          disabled={disabled}
          onSelect={onSaveCurrent}
        >
          <OfficeSemanticIcon semanticId="edit.saveStatus" size={18} />
          <span>{translate("spreadsheet.themeGallery.save")}</span>
        </DropdownMenuItem>
      </div>
    </div>
  );
}

export interface SpreadsheetThemeEffectGalleryProps {
  readonly disabled?: boolean;
  readonly locale?: string;
  /** The workbook's palette id; the tiles draw in its accents, as Office's do. */
  readonly paletteId: string;
  readonly onValueChange: (effectId: string) => void;
  readonly value: string;
  /** Test ids of another host's copy (PowerPoint's Design > Variants > Effects); Excel's by default. */
  readonly testId?: string;
  readonly itemTestIdPrefix?: string;
}

const EMU_PER_TILE_UNIT = 9525;

/**
 * An effect scheme's tile: three shapes in the workbook's first accents, each
 * with the scheme's matching fill, line and effect style.
 */
function EffectTile({
  effect,
  paletteId,
}: {
  effect: OfficeEffectScheme;
  paletteId: string;
}) {
  const id = useId().replace(/:/g, "");
  const accents = paletteById(paletteId).accents;
  const shapes = [
    { d: "M20 6 A11 11 0 1 1 19.99 6 Z" },
    { d: "M36 8 L49 30 L23 30 Z" },
    { d: "M8 26 H34 Q37 26 37 29 V40 Q37 43 34 43 H8 Q5 43 5 40 V29 Q5 26 8 26 Z" },
  ];
  return (
    <svg aria-hidden="true" className={styles.preview} viewBox="0 0 64 48">
      <defs>
        {shapes.map((_, index) => {
          const color = accents[index] ?? accents[0];
          const shadow = effect.effects[index]?.shadow;
          return (
            <g key={index}>
              <linearGradient id={`${id}-fill-${index}`} x1="0" y1="0" x2="0" y2="1">
                <stop offset="0" stopColor={tintColor(color, 0.35)} />
                <stop offset="1" stopColor={tintColor(color, -0.2)} />
              </linearGradient>
              {shadow ? (
                <filter id={`${id}-shadow-${index}`} x="-30%" y="-30%" width="160%" height="160%">
                  <feDropShadow
                    dx={
                      (Math.cos(((shadow.direction / 60000) * Math.PI) / 180) * shadow.distance) /
                      EMU_PER_TILE_UNIT
                    }
                    dy={
                      (Math.sin(((shadow.direction / 60000) * Math.PI) / 180) * shadow.distance) /
                      EMU_PER_TILE_UNIT
                    }
                    stdDeviation={Math.max(0.4, shadow.blur / EMU_PER_TILE_UNIT / 2)}
                    floodColor="#000000"
                    floodOpacity={Math.min(0.6, (shadow.alpha / 100000) * 0.9)}
                  />
                </filter>
              ) : null}
            </g>
          );
        })}
      </defs>
      <rect x="0.5" y="0.5" width="63" height="47" fill="#ffffff" stroke="#c6cad0" />
      {shapes.map((shape, index) => {
        const color = accents[index] ?? accents[0];
        const fill = effect.fills[index] ?? "solid";
        const style = effect.effects[index];
        const width = Math.max(0.5, (effect.lines[index] ?? 9525) / EMU_PER_TILE_UNIT / 1.5);
        return (
          <g
            key={index}
            filter={style?.shadow ? `url(#${id}-shadow-${index})` : undefined}
            data-effect-fill={fill}
          >
            <path
              d={shape.d}
              fill={
                fill === "none"
                  ? "none"
                  : fill === "solid"
                    ? color
                    : `url(#${id}-fill-${index})`
              }
              stroke={tintColor(color, -0.5)}
              strokeWidth={width}
            />
            {style?.bevel ? (
              <path d={shape.d} fill="none" stroke="#ffffff" strokeOpacity={0.55} strokeWidth={1} transform="translate(-0.6 -0.6)" />
            ) : null}
          </g>
        );
      })}
    </svg>
  );
}

export function SpreadsheetThemeEffectGallery({
  disabled = false,
  locale,
  paletteId,
  onValueChange,
  value,
  testId = "spreadsheet-theme-effect-gallery",
  itemTestIdPrefix = "spreadsheet-ribbon-page-layout-effect-",
}: SpreadsheetThemeEffectGalleryProps) {
  return (
    <div data-testid={testId}>
      <div className={`${officeGalleryMenuStyles.section} px-2 py-1.5 font-semibold`}>Office</div>
      <div
        data-testid="spreadsheet-theme-effect-gallery-scroll"
      >
        <DropdownMenuRadioGroup
          className={officeGalleryMenuStyles.grid}
          style={THEME_GRID}
          value={value}
          onValueChange={onValueChange}
        >
          {SPREADSHEET_THEME_EFFECTS.map((effect) => {
            const label = officeThemeLabel(effect.label, locale);
            return (
              <DropdownMenuRadioItem
                key={effect.id}
                value={effect.id}
                aria-label={label}
                className={officeGalleryMenuStyles.tile}
                data-selected={value === effect.id ? "true" : undefined}
                data-testid={`${itemTestIdPrefix}${effect.id}`}
                disabled={disabled}
                title={label}
              >
                <EffectTile effect={effect} paletteId={paletteId} />
                <span className={officeGalleryMenuStyles.name}>{label}</span>
              </DropdownMenuRadioItem>
            );
          })}
        </DropdownMenuRadioGroup>
      </div>
    </div>
  );
}
