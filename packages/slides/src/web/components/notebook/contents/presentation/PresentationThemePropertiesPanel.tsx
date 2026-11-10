"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The theme properties panel (preset, fonts, title and body sizes and colors,
 * background, accents).
 *
 * Extracted verbatim from the render of
 * `components/notebook/contents/PresentationWindow.tsx`: the JSX is unchanged;
 * the window passes the bindings it used to close over as `ctl`, and the render
 * condition and key stay at the call site. No hooks run here, so the window's
 * hook order is unchanged.
 */

import { presentationOfficeThemeLabel, type PresentationOfficeTheme } from "@mrnote/core/web/lib/presentation-office-themes.ts";
import {
  PRESET_THEMES,
  normalizeThemeColor,
  normalizeThemeFontFamily,
  DEFAULT_THEME_FONT_FAMILY,
  type ThemeFontFamily,
  THEME_FONT_FAMILIES,
  MIN_THEME_FONT_SIZE,
  MAX_THEME_FONT_SIZE,
  resolveThemeAccentColors,
  MAX_THEME_ACCENT_COLORS,
  type DeckTheme,
} from "@mrnote/core/web/lib/presentation-theme.ts";
import { presentationPresetChoiceId } from "../../../../lib/presentation-design-buttons";
import { parseThemeSizeInput } from "./PresentationThemePreview";
import { OfficeSemanticIcon } from "@mrnote/core/web/components/office-ribbon/icons/OfficeSemanticIcon.tsx";
import type { useTranslations } from "@mrnote/core/host/next-intl.tsx";
import type { RefObject } from "react";

/** The window bindings this tab reads, passed through unchanged. */
export interface PresentationThemePropertiesPanelCtl {
  t: ReturnType<typeof useTranslations<"console-notebooks">>;
  locale: string;
  themePropertiesRef: RefObject<HTMLDivElement | null>;
  theme: DeckTheme | undefined;
  activeThemePresetId: string;
  activeOfficeTheme: PresentationOfficeTheme | undefined;
  themeSelectionDisabled: boolean;
  themeSelectionUnavailableReason: string | null;
  applyDeckTheme: (nextTheme: DeckTheme | undefined) => void;
  applyThemePreset: (presetId: string) => void;
  updateThemeField: (patch: Partial<DeckTheme>) => void;
  updateThemeAccentColor: (index: number, color: string) => void;
  removeThemeAccentColor: (index: number) => void;
  addThemeAccentColor: () => void;
}

export function PresentationThemePropertiesPanel({ ctl }: { ctl: PresentationThemePropertiesPanelCtl }) {
  const {
    t,
    locale,
    themePropertiesRef,
    theme,
    activeThemePresetId,
    activeOfficeTheme,
    themeSelectionDisabled,
    themeSelectionUnavailableReason,
    applyDeckTheme,
    applyThemePreset,
    updateThemeField,
    updateThemeAccentColor,
    removeThemeAccentColor,
    addThemeAccentColor,
  } = ctl;
  return (
          <div
            ref={themePropertiesRef}
            className="presentation-window__theme-props"
            data-testid="presentation-theme-props"
            role="group"
            aria-label={t("presentation.theme.toolbarLabel")}
          >
            <label
              style={{
                display: "inline-flex",
                alignItems: "center",
                gap: 4,
                fontSize: 12,
              }}
            >
              {t("presentation.theme.label")}
              <select
                value={activeThemePresetId}
                onChange={(event) => applyThemePreset(event.target.value)}
                data-testid="presentation-theme-preset"
                aria-label={t("presentation.themePresetLabel")}
                disabled={themeSelectionDisabled}
                title={themeSelectionUnavailableReason ?? undefined}
              >
                {activeThemePresetId === "custom" ? (
                  <option value="custom" disabled>
                    {t("presentation.theme.preset.custom")}
                  </option>
                ) : null}
                {activeOfficeTheme ? (
                  <option value={activeThemePresetId} disabled>
                    {presentationOfficeThemeLabel(activeOfficeTheme, locale)}
                  </option>
                ) : null}
                {PRESET_THEMES.map((preset) => (
                  <option
                    key={preset.id}
                    value={presentationPresetChoiceId(preset.id)}
                  >
                    {t(
                      `presentation.theme.preset.${preset.labelKey}` as "presentation.theme.preset.none",
                    )}
                  </option>
                ))}
              </select>
            </label>
            <label
              style={{
                display: "inline-flex",
                alignItems: "center",
                gap: 4,
                fontSize: 12,
              }}
            >
              {t("presentation.theme.background")}
              <input
                type="color"
                value={normalizeThemeColor(theme?.background) ?? "#ffffff"}
                onChange={(event) =>
                  updateThemeField({ background: event.target.value })
                }
                data-testid="presentation-theme-bg"
                aria-label={t("presentation.theme.background")}
              />
            </label>
            <label
              style={{
                display: "inline-flex",
                alignItems: "center",
                gap: 4,
                fontSize: 12,
              }}
            >
              {t("presentation.themeFontLabel")}
              <select
                value={
                  normalizeThemeFontFamily(theme?.fontFamily) ??
                  DEFAULT_THEME_FONT_FAMILY
                }
                onChange={(event) =>
                  updateThemeField({
                    fontFamily: event.target.value as ThemeFontFamily,
                  })
                }
                data-testid="presentation-theme-font"
                aria-label={t("presentation.themeFontLabel")}
              >
                {THEME_FONT_FAMILIES.map((font) => (
                  <option key={font} value={font}>
                    {t(
                      `presentation.theme.font.${font}` as "presentation.theme.font.system",
                    )}
                  </option>
                ))}
              </select>
            </label>
            <label
              style={{
                display: "inline-flex",
                alignItems: "center",
                gap: 4,
                fontSize: 12,
              }}
            >
              {t("presentation.theme.titleColor")}
              <input
                type="color"
                value={normalizeThemeColor(theme?.titleColor) ?? "#1e293b"}
                onChange={(event) =>
                  updateThemeField({ titleColor: event.target.value })
                }
                data-testid="presentation-theme-title-color"
                aria-label={t("presentation.theme.titleColor")}
              />
            </label>
            <label
              style={{
                display: "inline-flex",
                alignItems: "center",
                gap: 4,
                fontSize: 12,
              }}
            >
              {t("presentation.theme.bodyColor")}
              <input
                type="color"
                value={normalizeThemeColor(theme?.bodyTextColor) ?? "#334155"}
                onChange={(event) =>
                  updateThemeField({ bodyTextColor: event.target.value })
                }
                data-testid="presentation-theme-body-color"
                aria-label={t("presentation.theme.bodyColor")}
              />
            </label>
            {/* Slide-master TITLE / BODY text styles: default font size (px) +
                default bold. Empty size = inherit (the renderer's built-in size).
                A per-block style.fontSize / style.bold still overrides these. */}
            <label
              style={{
                display: "inline-flex",
                alignItems: "center",
                gap: 4,
                fontSize: 12,
              }}
            >
              {t("presentation.theme.titleSizeLabel")}
              <input
                type="number"
                min={MIN_THEME_FONT_SIZE}
                max={MAX_THEME_FONT_SIZE}
                value={theme?.titleFontSize ?? ""}
                onChange={(event) =>
                  updateThemeField({
                    titleFontSize: parseThemeSizeInput(event.target.value),
                  })
                }
                data-testid="presentation-theme-title-size"
                aria-label={t("presentation.theme.titleSizeLabel")}
                placeholder={t("presentation.theme.sizeInherit")}
                style={{ fontSize: 12, width: 64 }}
              />
            </label>
            <label
              style={{
                display: "inline-flex",
                alignItems: "center",
                gap: 4,
                fontSize: 12,
              }}
            >
              {t("presentation.theme.bodySizeLabel")}
              <input
                type="number"
                min={MIN_THEME_FONT_SIZE}
                max={MAX_THEME_FONT_SIZE}
                value={theme?.bodyFontSize ?? ""}
                onChange={(event) =>
                  updateThemeField({
                    bodyFontSize: parseThemeSizeInput(event.target.value),
                  })
                }
                data-testid="presentation-theme-body-size"
                aria-label={t("presentation.theme.bodySizeLabel")}
                placeholder={t("presentation.theme.sizeInherit")}
                style={{ fontSize: 12, width: 64 }}
              />
            </label>
            <label
              style={{
                display: "inline-flex",
                alignItems: "center",
                gap: 4,
                fontSize: 12,
              }}
            >
              <input
                type="checkbox"
                checked={theme?.titleBold === true}
                onChange={(event) =>
                  updateThemeField({
                    titleBold: event.target.checked ? true : undefined,
                  })
                }
                data-testid="presentation-theme-title-bold"
                aria-label={t("presentation.theme.titleBoldLabel")}
              />
              {t("presentation.theme.titleBoldLabel")}
            </label>
            <label
              style={{
                display: "inline-flex",
                alignItems: "center",
                gap: 4,
                fontSize: 12,
              }}
            >
              <input
                type="checkbox"
                checked={theme?.bodyBold === true}
                onChange={(event) =>
                  updateThemeField({
                    bodyBold: event.target.checked ? true : undefined,
                  })
                }
                data-testid="presentation-theme-body-bold"
                aria-label={t("presentation.theme.bodyBoldLabel")}
              />
              {t("presentation.theme.bodyBoldLabel")}
            </label>
            {/* Accent palette editor: the theme's quick-pick swatches, shown
                additively in the shape-fill / text / run colour pickers. Each
                swatch is an editable colour input with a remove control; "+ adds
                one up to the cap. An empty palette drops the field (byte-clean). */}
            <div
              className="presentation-window__theme-accents"
              data-testid="presentation-theme-accents"
              role="group"
              aria-label={t("presentation.theme.accentsLabel")}
              style={{
                display: "inline-flex",
                alignItems: "center",
                gap: 4,
                flexWrap: "wrap",
              }}
            >
              <span style={{ fontSize: 12 }}>
                {t("presentation.theme.accentsLabel")}
              </span>
              {resolveThemeAccentColors(theme).map((c, index) => (
                <span
                  key={`accent-edit-${index}-${c}`}
                  style={{
                    display: "inline-flex",
                    alignItems: "center",
                    gap: 2,
                  }}
                >
                  <input
                    type="color"
                    value={c}
                    onChange={(event) =>
                      updateThemeAccentColor(index, event.target.value)
                    }
                    data-testid="presentation-theme-accent"
                    data-color={c}
                    aria-label={`${t("presentation.theme.accentsLabel")} ${index + 1}`}
                    style={{ width: 24, height: 20, padding: 0 }}
                  />
                  <button
                    type="button"
                    onClick={() => removeThemeAccentColor(index)}
                    data-testid="presentation-theme-accent-remove"
                    aria-label={`${t("presentation.theme.accentRemoveLabel")} ${index + 1}`}
                    title={t("presentation.theme.accentRemoveLabel")}
                    style={{
                      fontSize: 11,
                      lineHeight: 1,
                      cursor: "pointer",
                      padding: "0 2px",
                    }}
                  >
                    <OfficeSemanticIcon semanticId="edit.close" size={12} />
                  </button>
                </span>
              ))}
              {resolveThemeAccentColors(theme).length <
              MAX_THEME_ACCENT_COLORS ? (
                <button
                  type="button"
                  onClick={addThemeAccentColor}
                  data-testid="presentation-theme-accent-add"
                  aria-label={t("presentation.theme.accentAddLabel")}
                  title={t("presentation.theme.accentAddLabel")}
                  style={{ fontSize: 11, cursor: "pointer", padding: "0 4px" }}
                >
                  {"+"}
                </button>
              ) : null}
            </div>
            {theme ? (
              <button
                type="button"
                onClick={() => applyDeckTheme(undefined)}
                data-testid="presentation-theme-clear"
                style={{ fontSize: 11, cursor: "pointer" }}
              >
                {t("presentation.theme.clear")}
              </button>
            ) : null}
          </div>
  );
}
