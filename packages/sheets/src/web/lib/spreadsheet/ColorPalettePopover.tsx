/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import {
  OfficeColorPalettePanel,
  type OfficeColorPalettePanelProps,
} from "@mrnote/core/web/components/office-ribbon/OfficeColorPalettePanel.tsx";
import {
  OfficeColorPalettePopover,
  type OfficeColorPaletteEntry,
  type OfficeColorPalettePopoverLabels as ColorPalettePopoverLabels,
  type OfficeColorPalettePopoverProps,
} from "@mrnote/core/web/components/office-ribbon/OfficeColorPalettePopover.tsx";
import { paletteById } from "@mrnote/core/web/lib/spreadsheet/color-palette.ts";
import {
  themeColorGrid,
  themeColorFromEntry,
  themeTintRow,
  type SpreadsheetThemeColor,
} from "@mrnote/core/web/lib/spreadsheet/theme-colors.ts";

export type { ColorPalettePopoverLabels };

interface SpreadsheetPaletteBinding {
  paletteId?: string;
  onPickTheme?: (color: string, reference: SpreadsheetThemeColor) => void;
  selectedThemeColor?: SpreadsheetThemeColor;
  onPick: (color: string | null) => void;
}

/**
 * The spreadsheet's view of the shared Office palette: Excel's theme grid for
 * the workbook's palette, and a pick that reports the theme binding when the
 * swatch is a theme colour. Shared by the popup opened from a button and the
 * panel embedded in a dialog page, so both bind theme colours identically.
 */
function spreadsheetPalettePanelProps({
  paletteId = "office",
  onPickTheme,
  selectedThemeColor,
  onPick,
}: SpreadsheetPaletteBinding): Pick<
  OfficeColorPalettePanelProps,
  "paletteId" | "themeColumns" | "onPickEntry" | "selectedThemeOccurrenceId"
> {
  const palette = paletteById(paletteId);
  // Excel: light1/dark1/light2/dark2, then the six theme accents.
  const grid = themeColorGrid(palette.id);
  const pick = (color: string | null, entry?: OfficeColorPaletteEntry) => {
    const ref = color ? themeColorFromEntry(color, entry, palette.id) : undefined;
    if (color && ref && onPickTheme) onPickTheme(color, ref);
    else onPick(color);
  };
  const selectedRow = selectedThemeColor
    ? themeTintRow(selectedThemeColor, palette.id)
    : -1;
  return {
    paletteId: palette.id,
    themeColumns: grid,
    onPickEntry: pick,
    selectedThemeOccurrenceId:
      selectedThemeColor && selectedRow >= 0
        ? `theme-${selectedThemeColor.theme}-${selectedRow}`
        : undefined,
  };
}

export interface ColorPalettePopoverProps extends Omit<
  OfficeColorPalettePopoverProps,
  "onPickEntry"
> {
  onPickTheme?: (color: string, reference: SpreadsheetThemeColor) => void;
  selectedThemeColor?: SpreadsheetThemeColor;
}

/** The palette as a popup: portalled to the viewport and placed by its anchor. */
export function ColorPalettePopover({
  paletteId,
  onPickTheme,
  selectedThemeColor,
  ...props
}: ColorPalettePopoverProps) {
  return (
    <OfficeColorPalettePopover
      {...props}
      {...spreadsheetPalettePanelProps({
        paletteId,
        onPickTheme,
        selectedThemeColor,
        onPick: props.onPick,
      })}
    />
  );
}

export interface ColorPalettePanelProps extends Omit<
  OfficeColorPalettePanelProps,
  "onPickEntry"
> {
  onPickTheme?: (color: string, reference: SpreadsheetThemeColor) => void;
  selectedThemeColor?: SpreadsheetThemeColor;
}

/**
 * The same palette rendered in place, where it IS the page rather than a popup
 * from a button (Format Cells > Fill). Nothing is portalled, so it lives inside
 * whatever dialog or submenu hosts it.
 */
export function ColorPalettePanel({
  paletteId,
  onPickTheme,
  selectedThemeColor,
  ...props
}: ColorPalettePanelProps) {
  return (
    <OfficeColorPalettePanel
      {...props}
      {...spreadsheetPalettePanelProps({
        paletteId,
        onPickTheme,
        selectedThemeColor,
        onPick: props.onPick,
      })}
    />
  );
}
