"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The Design tab of the presentation Ribbon (Themes, Variants, Customize, Designer), drawn from the window bindings it reads.
 *
 * Extracted verbatim from the render of
 * `components/notebook/contents/PresentationWindow.tsx`: the JSX is unchanged;
 * the window passes the bindings it used to close over as `ctl`, and the tab
 * condition and key stay at the call site. No hooks run here, so the window's
 * hook order is unchanged.
 */

import {
  OfficeHomeRibbonGroup,
  OfficeHomeMenuTrigger,
  OfficeHomeCommandRow,
  OfficeHomeCommand,
} from "@mrnote/core/web/components/office-ribbon/OfficeHomeRibbon.tsx";
import { OfficeRibbonGallery } from "@mrnote/core/web/components/office-ribbon/OfficeRibbonGallery.tsx";
import presentationChrome from "@mrnote/core/web/lib/presentation/presentation-window-shell.module.css";
import { presentationThemePreview, PresentationThemePreview } from "./PresentationThemePreview";
import { PresentationOfficeThemeThumbnail } from "./PresentationOfficeThemeThumbnail";
import {
  PRESET_THEMES,
  resolveCatalogThemeDesign,
  type DeckTheme,
  type PRESENTATION_THEME_VARIANTS,
  type PresentationThemeVariantId,
} from "@mrnote/core/web/lib/presentation-theme.ts";
import { presentationPresetChoiceId } from "../../../../lib/presentation-design-buttons";
import {
  DropdownMenu,
  DropdownMenuTrigger,
  DropdownMenuLabel,
  DropdownMenuItem,
  DropdownMenuSeparator,
  DropdownMenuSub,
  DropdownMenuSubTrigger,
} from "@mrnote/core/web/components/ui/dropdown-menu.tsx";
import { OfficeSemanticIcon } from "@mrnote/core/web/components/office-ribbon/icons/OfficeSemanticIcon.tsx";
import { RibbonDropdownMenuSubContent } from "@mrnote/core/web/components/office-ribbon/CupertinoRibbonMenu.tsx";
import {
  OfficeGalleryMenuContent,
  officeGalleryMenuGridStyle,
  officeGalleryMenuStyles,
} from "@mrnote/core/web/components/office-ribbon/OfficeGalleryMenu.tsx";
import spreadsheetThemeAxisStyles from "@mrnote/sheets/web/components/notebook/contents/spreadsheet/SpreadsheetThemeAxisGallery.module.css";
import { SpreadsheetThemeColorGallery, SpreadsheetThemeFontGallery, type SpreadsheetThemeCustomFontSet } from "@mrnote/sheets/web/components/notebook/contents/spreadsheet/SpreadsheetThemeAxisGallery.tsx";
import { IconDesignBackground, IconDesignerCopilot } from "@mrnote/core/web/components/notebook/contents/PresentationRibbonIcons.tsx";
import type { useTranslations } from "@mrnote/core/host/next-intl.tsx";
import type { RefObject } from "react";
import type { PresentationOfficeTheme } from "@mrnote/core/web/lib/presentation-office-themes.ts";
import type { JSX } from "react/jsx-runtime";
import type { ResolvedPresentationFeatureCommand } from "@mrnote/core/web/components/notebook/contents/ribbon/commandRegistry.ts";
import type { OfficeBackgroundStyle } from "@mrnote/core/web/lib/presentation-office-background-styles.ts";

/** The window bindings this tab reads, passed through unchanged. */
export interface PresentationDesignRibbonTabCtl {
  t: ReturnType<typeof useTranslations<"console-notebooks">>;
  locale: string;
  backgroundPaneOpen: boolean;
  backgroundPaneTriggerRef: RefObject<HTMLButtonElement | null>;
  openBackgroundPane: () => void;
  theme: DeckTheme | undefined;
  activeThemePresetId: string;
  activeOfficeTheme: PresentationOfficeTheme | undefined;
  activeThemeVariantId: string | null;
  themeSelectionDisabled: boolean;
  themeSelectionUnavailableReason: string | null;
  availableThemeVariants: { label: string; id: PresentationThemeVariantId; labelKey?: typeof PRESENTATION_THEME_VARIANTS[number]["labelKey"]; officeIndex?: number; theme: DeckTheme; }[];
  officeThemeTiles: { id: string; design: DeckTheme & Required<Pick<DeckTheme, "background" | "titleColor" | "bodyTextColor" | "accentColors" | "colorSlots" | "fontScheme" | "fontFamily">>; label: string; }[];
  currentThemeTile: { label: string; design: DeckTheme; };
  applyThemePreset: (presetId: string) => void;
  applyThemeVariant: (variantId: PresentationThemeVariantId) => void;
  applyThemeColorScheme: (schemeId: string) => void;
  applyThemeFontScheme: (fontSchemeId: string) => void;
  /** The Fonts gallery's Custom section and its Customize Fonts... command. */
  themeFontsGallery: {
    customFonts: readonly SpreadsheetThemeCustomFontSet[];
    value: string;
    openCustomize: () => void;
  };
  designIdeasOpen: boolean;
  designerDisabledReason: string | null;
  presentationEffectGallery: (host: "design" | "slide-master") => JSX.Element;
  featureDecision: (featureId: string) => ResolvedPresentationFeatureCommand;
  executePresentationFeature: (featureId: string) => undefined;
  officeBackgroundStyleCells: OfficeBackgroundStyle[];
  presentationBackgroundStyleItems: (prefix: string) => JSX.Element | null;
  slideSizeMenu: JSX.Element;
}

export function PresentationDesignRibbonTab({ ctl }: { ctl: PresentationDesignRibbonTabCtl }) {
  const {
    t,
    locale,
    backgroundPaneOpen,
    backgroundPaneTriggerRef,
    openBackgroundPane,
    theme,
    activeThemePresetId,
    activeOfficeTheme,
    activeThemeVariantId,
    themeSelectionDisabled,
    themeSelectionUnavailableReason,
    availableThemeVariants,
    officeThemeTiles,
    currentThemeTile,
    applyThemePreset,
    applyThemeVariant,
    applyThemeColorScheme,
    applyThemeFontScheme,
    themeFontsGallery,
    designIdeasOpen,
    designerDisabledReason,
    presentationEffectGallery,
    featureDecision,
    executePresentationFeature,
    officeBackgroundStyleCells,
    presentationBackgroundStyleItems,
    slideSizeMenu,
  } = ctl;
  return (
                  <>
                    <OfficeHomeRibbonGroup
                      label={t("presentation.theme.label")}
                      ribbon="design"
                      size="gallery"
                      testId="presentation-ribbon-design-themes-group"
                    >
                      <OfficeRibbonGallery
                        ariaLabel={t("presentation.theme.label")}
                        className={presentationChrome.secondaryGallery}
                        testId="presentation-theme-strip"
                        viewport={<div className="note-style-strip">
                        <div
                          role="group"
                          aria-label={t("presentation.theme.group.current")}
                          data-testid="presentation-theme-group-current"
                          style={{ display: "contents" }}
                        >
                          <button
                            type="button"
                            disabled={themeSelectionDisabled}
                            className={`note-style-strip__card ${presentationChrome.secondaryGalleryCard}`}
                            data-testid="presentation-theme-current"
                            data-theme-background-kind={
                              presentationThemePreview(currentThemeTile.design).kind
                            }
                            aria-pressed
                            data-selected="true"
                            title={
                              themeSelectionUnavailableReason ??
                              `${t("presentation.theme.group.current")}: ${currentThemeTile.label}`
                            }
                            onMouseDown={(event) => event.preventDefault()}
                          >
                            {activeOfficeTheme ? (
                              <span
                                className="note-style-strip__preview"
                                style={{ height: 32, padding: 0, width: 57 }}
                              >
                                <PresentationOfficeThemeThumbnail
                                  theme={currentThemeTile.design}
                                  themeId={activeOfficeTheme.id}
                                  variantId={activeThemeVariantId}
                                />
                              </span>
                            ) : (
                              <PresentationThemePreview
                                className="note-style-strip__preview"
                                theme={currentThemeTile.design}
                              />
                            )}
                            <span className="note-style-strip__name">
                              {currentThemeTile.label}
                            </span>
                          </button>
                        </div>
                        <div
                          role="group"
                          aria-label={t("presentation.theme.group.office")}
                          data-testid="presentation-theme-group-office"
                          style={{ display: "contents" }}
                        >
                          {officeThemeTiles.map((tile) => (
                            <button
                              key={`office-${tile.id}`}
                              type="button"
                              disabled={themeSelectionDisabled}
                              className={`note-style-strip__card ${presentationChrome.secondaryGalleryCard}`}
                              data-testid={`presentation-theme-quick-${tile.id}`}
                              data-theme-background-kind="solid"
                              data-office-theme-id={tile.id}
                              aria-pressed={activeOfficeTheme?.id === tile.id}
                              data-selected={
                                activeOfficeTheme?.id === tile.id ? "true" : "false"
                              }
                              title={themeSelectionUnavailableReason ?? tile.label}
                              onMouseDown={(event) => event.preventDefault()}
                              onClick={() => applyThemePreset(tile.id)}
                            >
                              <span
                                className="note-style-strip__preview"
                                style={{ height: 32, padding: 0, width: 57 }}
                              >
                                <PresentationOfficeThemeThumbnail
                                  theme={tile.design}
                                  themeId={tile.id}
                                />
                              </span>
                              <span className="note-style-strip__name">
                                {tile.label}
                              </span>
                            </button>
                          ))}
                        </div>
                        <div
                          role="group"
                          aria-label={t("presentation.theme.group.custom")}
                          data-testid="presentation-theme-group-custom"
                          style={{ display: "contents" }}
                        >
                        {PRESET_THEMES.filter((preset) => preset.id !== "none")
                          .map((preset) => {
                            const choiceId = presentationPresetChoiceId(preset.id);
                            const design =
                              resolveCatalogThemeDesign(preset.id, "base") ??
                              preset.theme;
                            const backgroundKind =
                              presentationThemePreview(design).kind;
                            return (
                              <button
                                key={preset.id}
                                type="button"
                                disabled={themeSelectionDisabled}
                                className={`note-style-strip__card ${presentationChrome.secondaryGalleryCard}`}
                                data-testid={`presentation-theme-quick-${choiceId}`}
                                data-theme-background-kind={backgroundKind}
                                aria-pressed={activeThemePresetId === choiceId}
                                data-selected={
                                  activeThemePresetId === choiceId
                                    ? "true"
                                    : "false"
                                }
                                title={
                                  themeSelectionUnavailableReason ??
                                  t(
                                    `presentation.theme.preset.${preset.labelKey}` as "presentation.theme.preset.classic",
                                  )
                                }
                                onMouseDown={(event) => event.preventDefault()}
                                onClick={() => applyThemePreset(choiceId)}
                              >
                                <PresentationThemePreview
                                  className="note-style-strip__preview"
                                  theme={design}
                                />
                                <span className="note-style-strip__name">
                                  {t(
                                    `presentation.theme.preset.${preset.labelKey}` as "presentation.theme.preset.classic",
                                  )}
                                </span>
                              </button>
                            );
                          })}
                        </div>
                        </div>}
                        more={<DropdownMenu>
                          <DropdownMenuTrigger asChild>
                            <OfficeHomeMenuTrigger
                              geometry="gallery"
                              disabled={themeSelectionDisabled}
                              displayLabel={t("presentation.theme.label")}
                              icon={<OfficeSemanticIcon semanticId="menu.chevron.down" size={12} />}
                              label={
                                themeSelectionUnavailableReason ??
                                t("presentation.theme.label")
                              }
                              testId="presentation-theme-gallery-open"
                              variant="icon"
                            />
                          </DropdownMenuTrigger>
                          <OfficeGalleryMenuContent
                            data-testid="presentation-theme-gallery-menu"
                          >
                            <DropdownMenuLabel className={officeGalleryMenuStyles.section}>
                              {t("presentation.theme.group.current")}
                            </DropdownMenuLabel>
                            <div
                              className={officeGalleryMenuStyles.grid}
                              style={officeGalleryMenuGridStyle(6, 124)}
                              role="grid"
                              aria-label={t("presentation.theme.group.current")}
                              data-testid="presentation-theme-gallery-group-current"
                            >
                              <DropdownMenuItem asChild key="theme-menu-current">
                                <button
                                  type="button"
                                  className={officeGalleryMenuStyles.tile}
                                  aria-pressed
                                  data-selected="true"
                                  data-testid="presentation-theme-gallery-current"
                                >
                                  {activeOfficeTheme ? (
                                    <span
                                      className={
                                        presentationChrome.secondaryExpandedGalleryPreview
                                      }
                                      style={{ padding: 0 }}
                                    >
                                      <PresentationOfficeThemeThumbnail
                                        theme={currentThemeTile.design}
                                        themeId={activeOfficeTheme.id}
                                        variantId={activeThemeVariantId}
                                      />
                                    </span>
                                  ) : (
                                    <PresentationThemePreview
                                      className={
                                        presentationChrome.secondaryExpandedGalleryPreview
                                      }
                                      expanded
                                      theme={currentThemeTile.design}
                                    />
                                  )}
                                  <span
                                    className={officeGalleryMenuStyles.name}
                                  >
                                    {currentThemeTile.label}
                                  </span>
                                </button>
                              </DropdownMenuItem>
                            </div>
                            <DropdownMenuLabel className={officeGalleryMenuStyles.section}>
                              {t("presentation.theme.group.office")}
                            </DropdownMenuLabel>
                            <div
                              className={officeGalleryMenuStyles.grid}
                              style={officeGalleryMenuGridStyle(6, 124)}
                              role="grid"
                              aria-label={t("presentation.theme.group.office")}
                              data-testid="presentation-theme-gallery-group-office"
                            >
                              {officeThemeTiles.map((tile) => (
                                <DropdownMenuItem
                                  asChild
                                  key={`theme-menu-office-${tile.id}`}
                                  onSelect={() => applyThemePreset(tile.id)}
                                >
                                  <button
                                    type="button"
                                    className={officeGalleryMenuStyles.tile}
                                    aria-pressed={activeOfficeTheme?.id === tile.id}
                                    data-selected={
                                      activeOfficeTheme?.id === tile.id
                                        ? "true"
                                        : "false"
                                    }
                                    data-testid={`presentation-theme-gallery-${tile.id}`}
                                  >
                                    <span
                                      className={
                                        presentationChrome.secondaryExpandedGalleryPreview
                                      }
                                      style={{ padding: 0 }}
                                    >
                                      <PresentationOfficeThemeThumbnail
                                        theme={tile.design}
                                        themeId={tile.id}
                                      />
                                    </span>
                                    <span
                                      className={officeGalleryMenuStyles.name}
                                    >
                                      {tile.label}
                                    </span>
                                  </button>
                                </DropdownMenuItem>
                              ))}
                            </div>
                            <DropdownMenuLabel className={officeGalleryMenuStyles.section}>
                              {t("presentation.theme.group.custom")}
                            </DropdownMenuLabel>
                            <div
                              className={officeGalleryMenuStyles.grid}
                              style={officeGalleryMenuGridStyle(6, 124)}
                              role="grid"
                              aria-label={t("presentation.theme.group.custom")}
                              data-testid="presentation-theme-gallery-group-custom"
                            >
                              {PRESET_THEMES.filter(
                                (preset) => preset.id !== "none",
                              ).map((preset) => {
                                const choiceId = presentationPresetChoiceId(preset.id);
                                return (
                                <DropdownMenuItem
                                  asChild
                                  key={`theme-menu-${preset.id}`}
                                  onSelect={() => applyThemePreset(choiceId)}
                                >
                                  <button
                                    type="button"
                                    className={officeGalleryMenuStyles.tile}
                                    aria-pressed={
                                      activeThemePresetId === choiceId
                                    }
                                    data-selected={
                                      activeThemePresetId === choiceId
                                        ? "true"
                                        : "false"
                                    }
                                    data-testid={`presentation-theme-gallery-${choiceId}`}
                                  >
                                    <PresentationThemePreview
                                      className={
                                        presentationChrome.secondaryExpandedGalleryPreview
                                      }
                                      expanded
                                      theme={
                                        resolveCatalogThemeDesign(
                                          preset.id,
                                          "base",
                                        ) ?? preset.theme
                                      }
                                    />
                                    <span
                                      className={officeGalleryMenuStyles.name}
                                    >
                                      {t(
                                        `presentation.theme.preset.${preset.labelKey}` as "presentation.theme.preset.classic",
                                      )}
                                    </span>
                                  </button>
                                </DropdownMenuItem>
                                );
                              })}
                            </div>
                          </OfficeGalleryMenuContent>
                        </DropdownMenu>}
                      />
                    </OfficeHomeRibbonGroup>

                    <OfficeHomeRibbonGroup
                      label={t("presentation.themePresetLabel")}
                      ribbon="design"
                      size="gallery"
                      testId="presentation-ribbon-design-variants-group"
                    >
                      <OfficeRibbonGallery
                        ariaLabel={t("presentation.themePresetLabel")}
                        className={`${presentationChrome.secondaryGallery} ${presentationChrome.variantGallery}`}
                        testId="presentation-theme-variant-strip"
                        viewport={<div className="note-style-strip">
                        {availableThemeVariants.map((variant) => (
                          <button
                            key={variant.id}
                            type="button"
                            disabled={themeSelectionDisabled}
                            className={`note-style-strip__card ${presentationChrome.secondaryGalleryCard}`}
                            data-testid={`presentation-theme-variant-${variant.id}`}
                            data-theme-background-kind={
                              presentationThemePreview(variant.theme).kind
                            }
                            aria-pressed={activeThemeVariantId === variant.id}
                            data-selected={
                              activeThemeVariantId === variant.id
                                ? "true"
                                : "false"
                            }
                            title={
                              themeSelectionUnavailableReason ??
                              variant.label
                            }
                            aria-label={variant.label}
                            onMouseDown={(event) => event.preventDefault()}
                            onClick={() => applyThemeVariant(variant.id)}
                          >
                            {activeOfficeTheme ? (
                              // An Office variant is its picture alone, as in
                              // PowerPoint: its name ("Office 主题 变型 1")
                              // is the tooltip, and drawn under the picture it
                              // was cut to "Office 主题…" in every tile.
                              <span
                                className="note-style-strip__preview"
                                style={{ height: 37, padding: 0, width: 66 }}
                              >
                                <PresentationOfficeThemeThumbnail
                                  theme={variant.theme}
                                  themeId={activeOfficeTheme.id}
                                  variantId={variant.id}
                                />
                              </span>
                            ) : (
                              <PresentationThemePreview
                                className="note-style-strip__preview"
                                theme={variant.theme}
                              />
                            )}
                            {activeOfficeTheme ? null : (
                              <span className="note-style-strip__name">
                                {variant.label}
                              </span>
                            )}
                          </button>
                        ))}
                        </div>}
                        more={<DropdownMenu>
                          <DropdownMenuTrigger asChild>
                            <OfficeHomeMenuTrigger
                              geometry="gallery"
                              disabled={themeSelectionDisabled}
                              displayLabel={t("presentation.themePresetLabel")}
                              icon={<OfficeSemanticIcon semanticId="menu.chevron.down" size={12} />}
                              label={
                                themeSelectionUnavailableReason ??
                                t("presentation.themePresetLabel")
                              }
                              testId="presentation-theme-variant-gallery-open"
                              variant="icon"
                            />
                          </DropdownMenuTrigger>
                          <OfficeGalleryMenuContent
                            data-testid="presentation-theme-variant-gallery-menu"
                          >
                            <DropdownMenuLabel className={officeGalleryMenuStyles.section}>
                              {t("presentation.themePresetLabel")}
                            </DropdownMenuLabel>
                            <div
                              className={officeGalleryMenuStyles.grid}
                              role="grid"
                              aria-label={t("presentation.themePresetLabel")}
                              style={officeGalleryMenuGridStyle(
                                Math.max(1, Math.min(4, availableThemeVariants.length)),
                                128,
                              )}
                            >
                              {availableThemeVariants.map((variant) => (
                                <DropdownMenuItem
                                  asChild
                                  key={`variant-menu-${variant.id}`}
                                  onSelect={() => applyThemeVariant(variant.id)}
                                >
                                  <button
                                    type="button"
                                    className={officeGalleryMenuStyles.tile}
                                    aria-pressed={
                                      activeThemeVariantId === variant.id
                                    }
                                    data-selected={
                                      activeThemeVariantId === variant.id
                                        ? "true"
                                        : "false"
                                    }
                                    data-testid={`presentation-theme-variant-gallery-${variant.id}`}
                                    title={variant.label}
                                    aria-label={variant.label}
                                  >
                                    {activeOfficeTheme ? (
                                      <span
                                        className={
                                          presentationChrome.secondaryExpandedGalleryPreview
                                        }
                                        style={{ padding: 0 }}
                                      >
                                        <PresentationOfficeThemeThumbnail
                                          theme={variant.theme}
                                          themeId={activeOfficeTheme.id}
                                          variantId={variant.id}
                                        />
                                      </span>
                                    ) : (
                                      <PresentationThemePreview
                                        className={
                                          presentationChrome.secondaryExpandedGalleryPreview
                                        }
                                        expanded
                                        theme={variant.theme}
                                      />
                                    )}
                                    {activeOfficeTheme ? null : (
                                      <span
                                        className={officeGalleryMenuStyles.name}
                                      >
                                        {variant.label}
                                      </span>
                                    )}
                                  </button>
                                </DropdownMenuItem>
                              ))}
                            </div>
                            <DropdownMenuSeparator />
                            {/* PowerPoint's Variants menu: Colors, Fonts,
                                Effects, Background Styles (PPT-A073). Colors
                                is Office's one Colors gallery, shared with
                                Excel's Page Layout > Colors. */}
                            <DropdownMenuSub>
                              <DropdownMenuSubTrigger
                                disabled={themeSelectionDisabled}
                                data-testid="presentation-theme-colors-open"
                              >
                                {t("toolbar.designColors")}
                              </DropdownMenuSubTrigger>
                              <RibbonDropdownMenuSubContent
                                className={spreadsheetThemeAxisStyles.colorMenu}
                              >
                                <SpreadsheetThemeColorGallery
                                  disabled={themeSelectionDisabled}
                                  locale={locale}
                                  value={theme?.colorSchemeId ?? ""}
                                  onValueChange={applyThemeColorScheme}
                                  testId="presentation-theme-color-gallery"
                                  itemTestIdPrefix="presentation-theme-color-"
                                />
                              </RibbonDropdownMenuSubContent>
                            </DropdownMenuSub>
                            <DropdownMenuSub>
                              <DropdownMenuSubTrigger
                                disabled={themeSelectionDisabled}
                                data-testid="presentation-theme-fonts-open"
                              >
                                {t("toolbar.designFonts")}
                              </DropdownMenuSubTrigger>
                              <RibbonDropdownMenuSubContent
                                className={spreadsheetThemeAxisStyles.fontMenu}
                              >
                                <SpreadsheetThemeFontGallery
                                  disabled={themeSelectionDisabled}
                                  locale={locale}
                                  showEastAsianFaces={locale.startsWith("zh")}
                                  value={themeFontsGallery.value}
                                  customFonts={themeFontsGallery.customFonts}
                                  customHeading={t("officeThemeFonts.custom")}
                                  onCustomize={themeFontsGallery.openCustomize}
                                  customizeLabel={t("officeThemeFonts.customize")}
                                  onValueChange={applyThemeFontScheme}
                                  testId="presentation-theme-font-gallery"
                                  itemTestIdPrefix="presentation-theme-font-"
                                />
                                {theme?.identity?.source ? (
                                  // Imported text carries the face its theme
                                  // reference resolved to at import; the page
                                  // cannot tell it from a face typed by hand,
                                  // so the canvas keeps it while the saved
                                  // package follows the new scheme (§75.1).
                                  <p
                                    className="px-3 pb-2 text-xs text-muted-foreground"
                                    style={{ maxWidth: 320 }}
                                    data-testid="presentation-theme-fonts-imported-note"
                                  >
                                    {t("presentation.theme.fontsImportedNote")}
                                  </p>
                                ) : null}
                              </RibbonDropdownMenuSubContent>
                            </DropdownMenuSub>
                            {/* Effects: Office's Effects gallery, shared
                                with Excel's Page Layout > Effects (PPT-A073). */}
                            <DropdownMenuSub>
                              <DropdownMenuSubTrigger
                                disabled={themeSelectionDisabled}
                                data-testid="presentation-theme-effects-open"
                              >
                                {t("toolbar.designEffects")}
                              </DropdownMenuSubTrigger>
                              <RibbonDropdownMenuSubContent className={officeGalleryMenuStyles.menu}>
                                {presentationEffectGallery("design")}
                              </RibbonDropdownMenuSubContent>
                            </DropdownMenuSub>
                            <DropdownMenuSub>
                              <DropdownMenuSubTrigger
                                disabled={themeSelectionDisabled || officeBackgroundStyleCells.length === 0}
                                data-testid="presentation-theme-background-styles-open"
                              >
                                {t("presentation.slideMaster.backgroundStyles")}
                              </DropdownMenuSubTrigger>
                              <RibbonDropdownMenuSubContent>
                                {presentationBackgroundStyleItems("design")}
                              </RibbonDropdownMenuSubContent>
                            </DropdownMenuSub>
                          </OfficeGalleryMenuContent>
                        </DropdownMenu>}
                      />
                    </OfficeHomeRibbonGroup>

                    <OfficeHomeRibbonGroup
                      label={t("ai.mode.customize")}
                      ribbon="design"
                      testId="presentation-ribbon-design-customize-group"
                    >
                      <OfficeHomeCommandRow>
                        {slideSizeMenu}
                        {/* Design > Customize > Format Background opens the
                            pane directly, as PowerPoint's button does
                            (PPT-A074); the fills, Apply to All and Reset
                            Background are in the pane. */}
                        <OfficeHomeCommand
                          buttonRef={backgroundPaneTriggerRef}
                          active={backgroundPaneOpen}
                          icon={<IconDesignBackground home size={24} />}
                          label={t("presentation.backgroundPane.title")}
                          onClick={openBackgroundPane}
                          testId="presentation-ribbon-deck-settings"
                          variant="large"
                        />
                      </OfficeHomeCommandRow>
                    </OfficeHomeRibbonGroup>

                    <OfficeHomeRibbonGroup
                      label={t("toolbar.ribbonEdit")}
                      ribbon="design"
                      testId="presentation-ribbon-design-designer-group"
                    >
                      {null}
                    </OfficeHomeRibbonGroup>
                  </>
  );
}
