/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Design > Themes > Custom: Word lists the themes the user saved with Save
 * Current Theme (its Themes folder) above Office's, and applying one is
 * applying that theme file. MRAI keeps each saved theme for the user's
 * browser (as Word's folder is the computer's): its three schemes, by
 * Office id or as the document's own, under the name it was saved with.
 */

import { useEffect, useState } from "react";

import type { DocumentDesign } from "./document-design";
import { documentDesignColors, documentDesignThemeId } from "./document-design";
import type { OfficeThemeFileTheme } from "./office-theme-file";
import {
  normalizeSpreadsheetNativeThemeColors,
  normalizeSpreadsheetNativeThemeEffects,
  normalizeSpreadsheetNativeThemeFonts,
} from "./spreadsheet/workbook-theme";

export interface DocumentCustomTheme {
  readonly id: string;
  readonly name: string;
  readonly theme: OfficeThemeFileTheme;
}

const STORAGE_KEY = "mrai.docs.customThemes";
const CHANGE_EVENT = "mrai:docs-custom-themes";
const MAX_THEMES = 50;

function axis<T>(value: unknown, readNative: (raw: unknown) => T | null): { id: string } | { native: T } | null {
  if (!value || typeof value !== "object") return null;
  const record = value as Record<string, unknown>;
  if (typeof record.id === "string" && record.id) return { id: record.id };
  const native = readNative(record.native);
  return native ? { native } : null;
}

function normalizeEntry(value: unknown): DocumentCustomTheme | null {
  if (!value || typeof value !== "object" || Array.isArray(value)) return null;
  const record = value as Record<string, unknown>;
  const id = typeof record.id === "string" && /^[a-z0-9-]{1,64}$/.test(record.id) ? record.id : null;
  const name = typeof record.name === "string" ? record.name.trim().slice(0, 120) : "";
  const theme = record.theme as Record<string, unknown> | undefined;
  if (!id || !name || !theme || typeof theme !== "object") return null;
  const colors = axis(theme.colors, normalizeSpreadsheetNativeThemeColors);
  const fonts = axis(theme.fonts, normalizeSpreadsheetNativeThemeFonts);
  const effects = axis(theme.effects, normalizeSpreadsheetNativeThemeEffects);
  if (!colors || !fonts || !effects) return null;
  return Object.freeze({
    id,
    name,
    theme: Object.freeze({
      name,
      colors,
      fonts,
      effects,
      theme: typeof theme.theme === "string" ? theme.theme : null,
    }) as OfficeThemeFileTheme,
  });
}

export function readCustomThemes(): DocumentCustomTheme[] {
  if (typeof window === "undefined") return [];
  try {
    const raw = JSON.parse(window.localStorage.getItem(STORAGE_KEY) ?? "[]");
    if (!Array.isArray(raw)) return [];
    return raw
      .map(normalizeEntry)
      .filter((entry): entry is DocumentCustomTheme => entry !== null)
      .slice(0, MAX_THEMES)
      .sort((a, b) => a.name.localeCompare(b.name));
  } catch {
    return [];
  }
}

function write(entries: readonly DocumentCustomTheme[]): boolean {
  try {
    window.localStorage.setItem(STORAGE_KEY, JSON.stringify(entries.slice(0, MAX_THEMES)));
    window.dispatchEvent(new Event(CHANGE_EVENT));
    return true;
  } catch {
    return false;
  }
}

/** Save Current Theme: the design's theme under `name` (a theme of that name is replaced). */
export function saveCustomTheme(rawName: string, design: DocumentDesign): DocumentCustomTheme | null {
  if (typeof window === "undefined") return null;
  const name = rawName.trim().slice(0, 120);
  if (!name) return null;
  const others = readCustomThemes().filter((entry) => entry.name.toLowerCase() !== name.toLowerCase());
  if (others.length >= MAX_THEMES) return null;
  const entry: DocumentCustomTheme = {
    id: `theme-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`,
    name,
    theme: {
      name,
      colors: design.nativeColors ? { native: design.nativeColors } : { id: design.colors },
      fonts: design.nativeFonts ? { native: design.nativeFonts } : { id: design.fonts },
      effects: design.nativeEffects ? { native: design.nativeEffects } : { id: design.effects },
      theme: documentDesignThemeId(design),
    },
  };
  return write([...others, entry]) ? entry : null;
}

export function deleteCustomTheme(id: string): boolean {
  if (typeof window === "undefined") return false;
  const entries = readCustomThemes();
  const next = entries.filter((entry) => entry.id !== id);
  return next.length !== entries.length && write(next);
}

/** The tile's colours: the theme's light background and its six accents. */
export function customThemeTileColors(
  entry: DocumentCustomTheme,
  design: DocumentDesign,
): { background: string; accents: string[] } {
  const colors =
    "native" in entry.theme.colors
      ? entry.theme.colors.native.colors
      : documentDesignColors({ ...design, colors: entry.theme.colors.id, nativeColors: undefined });
  return {
    background: colors.lt1,
    accents: [colors.accent1, colors.accent2, colors.accent3, colors.accent4, colors.accent5, colors.accent6],
  };
}

export function useCustomThemes(): DocumentCustomTheme[] {
  const [entries, setEntries] = useState<DocumentCustomTheme[]>([]);
  useEffect(() => {
    const read = () => setEntries(readCustomThemes());
    read();
    const onStorage = (event: StorageEvent) => {
      if (event.key === STORAGE_KEY) read();
    };
    window.addEventListener(CHANGE_EVENT, read);
    window.addEventListener("storage", onStorage);
    return () => {
      window.removeEventListener(CHANGE_EVENT, read);
      window.removeEventListener("storage", onStorage);
    };
  }, []);
  return entries;
}
