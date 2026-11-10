"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Design tab theme helpers: the theme font-size input parser and the small
 * theme preview swatch.
 *
 * Moved verbatim out of `components/notebook/contents/PresentationWindow.tsx`
 * (structural split, no behaviour change).
 */

import type { CSSProperties } from "react";
import { resolveSlideBackgroundFill, slideBackgroundCss } from "@mrnote/core/web/lib/presentation-slide-bg.ts";
import {
  THEME_FONT_FAMILY_CSS,
  themeFontSchemeFamily,
  type DeckTheme,
} from "@mrnote/core/web/lib/presentation-theme.ts";

/**
 * Parse a theme font-size number input into a value for `updateThemeField` /
 * `updateMasterThemeField`. An empty (or whitespace-only) field clears the
 * override (→ `undefined`, the field is deleted and the size falls back to the
 * built-in default). A non-empty value is passed through as a number; the
 * serialize edge (`normalizeTheme` via `normalizeThemeFontSize`) rounds + clamps
 * it to the allowed range and silently drops a non-finite / out-of-range value,
 * so the control never persists a degenerate size.
 */
export function parseThemeSizeInput(raw: string): number | undefined {
  if (raw.trim() === "") return undefined;
  return Number(raw);
}

export function presentationThemePreview(theme: DeckTheme): {
  kind: string;
  style: CSSProperties;
} {
  const fill = resolveSlideBackgroundFill({}, theme, () => false);
  return {
    kind: fill.kind === "none" && theme.background ? "solid" : fill.kind,
    style:
      fill.kind === "none"
        ? { background: theme.background ?? "#ffffff" }
        : slideBackgroundCss(fill),
  };
}

export function PresentationThemePreview({
  theme,
  className,
  expanded = false,
}: {
  theme: DeckTheme;
  className: string;
  expanded?: boolean;
}) {
  const preview = presentationThemePreview(theme);
  return (
    <span
      aria-hidden
      className={className}
      data-theme-background-kind={preview.kind}
      style={{
        ...preview.style,
        alignItems: "flex-start",
        display: "flex",
        flexDirection: "column",
        justifyContent: "space-between",
        padding: expanded ? 8 : 3,
      }}
    >
      <span
        data-theme-preview-aa
        style={{
          color: theme.titleColor,
          fontFamily:
            themeFontSchemeFamily(theme, true) ??
            THEME_FONT_FAMILY_CSS[theme.fontFamily ?? "system"],
          fontSize: expanded ? 26 : 15,
          fontWeight: theme.titleBold ? 700 : 400,
          lineHeight: 1,
        }}
      >
        Aa
      </span>
      <span data-theme-preview-accents style={{ alignSelf: "flex-start", display: "flex", gap: 2 }}>
        {/* A theme may repeat an accent (High Contrast's #999999 twice): keyed by place. */}
        {(theme.accentColors ?? []).slice(0, 6).map((accent, index) => (
          <span
            key={`${index}-${accent}`}
            style={{ background: accent, height: 4, width: expanded ? 10 : 4 }}
          />
        ))}
      </span>
    </span>
  );
}
