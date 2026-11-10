"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The Slide Master tab of the presentation Ribbon in Slide Master view, drawn
 * from the window bindings it reads.
 *
 * Extracted verbatim from the render of
 * `components/notebook/contents/PresentationWindow.tsx`: the JSX is unchanged;
 * the window passes the bindings it used to close over as `ctl`, and the tab
 * condition and key stay at the call site. No hooks run here, so the window's
 * hook order is unchanged.
 */

import { PresentationSlideMasterRibbon } from "./PresentationSlideMasterRibbon";
import {
  officeLayoutHasTitle,
  officeLayoutHasFooters,
  themeWithInsertedOfficePlaceholder,
  deckWithOfficeLayoutTitle,
  themeWithOfficeLayoutFooters,
  themeWithDeletedOfficeLayout,
  type OfficeDeckEdit,
} from "../../../../lib/presentation-office-layout-edit";
import {
  officeLayoutHidesBackgroundGraphics,
  themeWithOfficeLayoutBackgroundGraphics,
  themeWithInsertedOfficeLayout,
  themeWithInsertedOfficeSlideMaster,
  deckWithDeletedOfficeSlideMaster,
} from "../../../../lib/presentation-office-master-structure";
import { OFFICE_PRIMARY_MASTER_KEY, type OfficePlaceholderKind } from "@mrnote/core/web/lib/presentation-office-master-model.ts";
import {
  DropdownMenu,
  DropdownMenuTrigger,
  DropdownMenuItem,
  DropdownMenuSeparator,
} from "@mrnote/core/web/components/ui/dropdown-menu.tsx";
import { OfficeHomeMenuTrigger } from "@mrnote/core/web/components/office-ribbon/OfficeHomeRibbon.tsx";
import { Palette, Type as TypeIcon, Sparkles } from "@mrnote/core/web/components/office-ribbon/icons/houseIcons.tsx";
import { RibbonDropdownMenuContent } from "@mrnote/core/web/components/office-ribbon/CupertinoRibbonMenu.tsx";
import {
  OfficeGalleryMenuContent,
  officeGalleryMenuGridStyle,
  officeGalleryMenuStyles,
} from "@mrnote/core/web/components/office-ribbon/OfficeGalleryMenu.tsx";
import presentationChrome from "@mrnote/core/web/lib/presentation/presentation-window-shell.module.css";
import { PresentationOfficeThemeThumbnail } from "./PresentationOfficeThemeThumbnail";
import spreadsheetThemeAxisStyles from "@mrnote/sheets/web/components/notebook/contents/spreadsheet/SpreadsheetThemeAxisGallery.module.css";
import { SpreadsheetThemeColorGallery, SpreadsheetThemeFontGallery, type SpreadsheetThemeCustomFontSet } from "@mrnote/sheets/web/components/notebook/contents/spreadsheet/SpreadsheetThemeAxisGallery.tsx";
import { IconDesignBackground } from "@mrnote/core/web/components/notebook/contents/PresentationRibbonIcons.tsx";
import {
  themeWithoutOfficeMasterBackground,
  officeMasterPreserved,
  themeWithOfficeMasterPreserved,
  type OfficeMasterViewEntry,
} from "../../../../lib/presentation-office-master";
import type { useTranslations } from "@mrnote/core/host/next-intl.tsx";
import type { DeckTheme } from "@mrnote/core/web/lib/presentation-theme.ts";
import type { Dispatch, RefObject, SetStateAction } from "react";
import type { SlideModel } from "@mrnote/core/web/lib/presentation/deck-model.ts";
import type { embeddedPresentationAutosave } from "../../../../lib/presentation-embedded-session";
import type { PresentationOfficeTheme } from "@mrnote/core/web/lib/presentation-office-themes.ts";
import type { PresentationLiveView } from "../../../../lib/presentation/view-zoom";
import type { JSX } from "react/jsx-runtime";

/** The window bindings this tab reads, passed through unchanged. */
export interface PresentationSlideMasterRibbonTabCtl {
  t: ReturnType<typeof useTranslations<"console-notebooks">>;
  locale: string;
  theme: DeckTheme | undefined;
  themeRef: RefObject<DeckTheme | undefined>;
  slidesRef: RefObject<SlideModel[]>;
  autosave: ReturnType<typeof embeddedPresentationAutosave>;
  activeOfficeTheme: PresentationOfficeTheme | undefined;
  themeSelectionDisabled: boolean;
  officeThemeTiles: { id: string; design: DeckTheme & Required<Pick<DeckTheme, "background" | "titleColor" | "bodyTextColor" | "accentColors" | "colorSlots" | "fontScheme" | "fontFamily">>; label: string; }[];
  applyDeckTheme: (nextTheme: DeckTheme | undefined) => void;
  applyThemePreset: (presetId: string) => void;
  applyThemeColorScheme: (schemeId: string) => void;
  applyThemeFontScheme: (fontSchemeId: string) => void;
  /** The Fonts gallery's Custom section and its Customize Fonts... command. */
  themeFontsGallery: {
    customFonts: readonly SpreadsheetThemeCustomFontSet[];
    value: string;
    openCustomize: () => void;
  };
  setOfficeMasterKey: Dispatch<SetStateAction<string>>;
  setMasterBackgroundPaneOpen: Dispatch<SetStateAction<boolean>>;
  setMasterRenameOpen: Dispatch<SetStateAction<boolean>>;
  setMasterLayoutDialogOpen: Dispatch<SetStateAction<boolean>>;
  dispatchPresentationView: (next: PresentationLiveView, options?: { trigger?: HTMLElement | null; restoreReadingFocus?: boolean; }) => boolean;
  officeMasterEntries: OfficeMasterViewEntry[];
  officeMasterSelectedEntry: OfficeMasterViewEntry;
  officeMasterSelectedLayoutId: "title" | "title-content" | "two-content" | "comparison" | "title-only" | "content-caption" | "picture-caption" | "section-header" | "title-vertical-text" | "vertical-title-text" | "blank" | null;
  officeMasterLayoutWritable: boolean;
  officeMasterSelectedMasterKey: string;
  officeMasterSelectedIsInserted: boolean;
  slideMasterThemeDisabled: boolean;
  slideMasterThemeReason: string | undefined;
  officeMasterDeleteReason: string | undefined;
  officePlaceholderKindLabels: Record<"chart" | "table" | "smartart" | "media" | "online-image" | "content" | "content-vertical" | "text" | "text-vertical" | "picture", string>;
  commitOfficeDeckEdit: (edit: OfficeDeckEdit<SlideModel> | null) => void;
  presentationEffectGallery: (host: "design" | "slide-master") => JSX.Element;
  presentationBackgroundStyleItems: (prefix: string) => JSX.Element | null;
  slideSizeMenu: JSX.Element;
}

export function PresentationSlideMasterRibbonTab({ ctl }: { ctl: PresentationSlideMasterRibbonTabCtl }) {
  const {
    t,
    locale,
    theme,
    themeRef,
    slidesRef,
    autosave,
    activeOfficeTheme,
    themeSelectionDisabled,
    officeThemeTiles,
    applyDeckTheme,
    applyThemePreset,
    applyThemeColorScheme,
    applyThemeFontScheme,
    themeFontsGallery,
    setOfficeMasterKey,
    setMasterBackgroundPaneOpen,
    setMasterRenameOpen,
    setMasterLayoutDialogOpen,
    dispatchPresentationView,
    officeMasterEntries,
    officeMasterSelectedEntry,
    officeMasterSelectedLayoutId,
    officeMasterLayoutWritable,
    officeMasterSelectedMasterKey,
    officeMasterSelectedIsInserted,
    slideMasterThemeDisabled,
    slideMasterThemeReason,
    officeMasterDeleteReason,
    officePlaceholderKindLabels,
    commitOfficeDeckEdit,
    presentationEffectGallery,
    presentationBackgroundStyleItems,
    slideSizeMenu,
  } = ctl;
  return (
                  <PresentationSlideMasterRibbon
                    labels={{
                      editMaster: t("presentation.slideMaster.editMaster"),
                      insertSlideMaster: t("presentation.slideMaster.insertSlideMaster"),
                      insertLayout: t("presentation.slideMaster.insertLayout"),
                      delete: t("presentation.slideMaster.delete"),
                      rename: t("presentation.slideMaster.rename"),
                      preserve: t("presentation.slideMaster.preserve"),
                      masterLayoutGroup: t("presentation.slideMaster.masterLayout"),
                      masterLayout: t("presentation.slideMaster.masterLayout"),
                      insertPlaceholder: t("presentation.slideMaster.insertPlaceholder"),
                      title: t("presentation.slideMaster.title"),
                      footers: t("presentation.slideMaster.footers"),
                      editTheme: t("presentation.slideMaster.editTheme"),
                      background: t("presentation.slideMaster.background"),
                      hideBackgroundGraphics: t("presentation.slideMaster.hideBackgroundGraphics"),
                      size: t("presentation.slideMaster.size"),
                      close: t("presentation.slideMaster.close"),
                      closeMasterView: t("presentation.view.closeMaster"),
                    }}
                    structureUnavailableReason={t("presentation.slideMaster.structureUnavailable")}
                    layoutHasTitle={
                      officeMasterSelectedLayoutId
                        ? officeLayoutHasTitle(theme, officeMasterSelectedLayoutId)
                        : true
                    }
                    layoutHasFooters={
                      officeMasterSelectedLayoutId
                        ? officeLayoutHasFooters(theme, officeMasterSelectedLayoutId)
                        : true
                    }
                    layoutHidesBackgroundGraphics={
                      officeMasterSelectedEntry?.kind === "layout"
                        ? officeLayoutHidesBackgroundGraphics(theme, officeMasterSelectedEntry.layoutId)
                        : false
                    }
                    // Background > Hide Background Graphics on the selected
                    // layout and Master Layout on the master (PPT-B030).
                    {...(officeMasterLayoutWritable && officeMasterSelectedEntry?.kind === "layout"
                      ? {
                          onHideBackgroundGraphicsChange: (hide: boolean) => {
                            const next = themeWithOfficeLayoutBackgroundGraphics(
                              themeRef.current,
                              officeMasterSelectedEntry.layoutId,
                              hide,
                            );
                            if (next) applyDeckTheme(next);
                          },
                        }
                      : { hideBackgroundGraphicsUnavailableReason: t("presentation.slideMaster.selectLayoutFirst") })}
                    {...(officeMasterLayoutWritable && officeMasterSelectedEntry?.kind === "master"
                      ? { onMasterLayout: () => setMasterLayoutDialogOpen(true) }
                      : { masterLayoutUnavailableReason: t("presentation.slideMaster.selectMasterFirst") })}
                    {...(officeMasterLayoutWritable
                      ? {
                          // Edit Master > Insert Layout: a Custom Layout right
                          // after the selected layout, selected (PPT-B030).
                          onInsertLayout: () => {
                            const inserted = themeWithInsertedOfficeLayout(
                              themeRef.current,
                              t("presentation.layout.custom"),
                              officeMasterSelectedEntry?.kind === "layout"
                                ? officeMasterSelectedEntry.layoutId
                                : null,
                              officeMasterSelectedMasterKey,
                            );
                            if (!inserted) return;
                            applyDeckTheme(inserted.theme);
                            setOfficeMasterKey(inserted.layoutId);
                          },
                          // Edit Master > Insert Slide Master: Office's default
                          // theme as a master below the others, selected (PPT-B030).
                          onInsertSlideMaster: () => {
                            const inserted = themeWithInsertedOfficeSlideMaster(
                              themeRef.current,
                              t("presentation.slideMaster.customDesign"),
                            );
                            if (!inserted) return;
                            applyDeckTheme(inserted.theme);
                            setOfficeMasterKey(inserted.masterKey);
                          },
                        }
                      : {})}
                    {...(officeMasterLayoutWritable &&
                    officeMasterSelectedEntry?.kind === "master" &&
                    officeMasterDeleteReason === undefined
                      ? {
                          // Edit Master > Delete on a master no slide uses,
                          // while another remains (PPT-B030).
                          onDelete: () => {
                            const masterKey = officeMasterSelectedEntry.masterKey;
                            const index = officeMasterEntries.findIndex((entry) => entry.key === officeMasterSelectedEntry.key);
                            const edit = deckWithDeletedOfficeSlideMaster(themeRef.current, slidesRef.current, masterKey);
                            if (!edit) return;
                            commitOfficeDeckEdit(edit);
                            // The master above takes the selection; deleting master one selects the new first.
                            const above = officeMasterEntries
                              .slice(0, Math.max(0, index))
                              .reverse()
                              .find((entry) => entry.kind === "master");
                            setOfficeMasterKey(above?.key ?? OFFICE_PRIMARY_MASTER_KEY);
                          },
                        }
                      : {})}
                    layoutCommandsUnavailableReason={
                      officeMasterEntries.length > 0 && !officeMasterSelectedLayoutId
                        ? t("presentation.slideMaster.selectLayoutFirst")
                        : undefined
                    }
                    deleteUnavailableReason={officeMasterDeleteReason}
                    placeholderKindLabels={officePlaceholderKindLabels}
                    {...(officeMasterLayoutWritable && officeMasterSelectedLayoutId
                      ? {
                          onInsertPlaceholder: (kind: OfficePlaceholderKind) => {
                            const next = themeWithInsertedOfficePlaceholder(
                              themeRef.current,
                              officeMasterSelectedLayoutId,
                              kind,
                            );
                            if (next) applyDeckTheme(next);
                          },
                          onTitleChange: (on: boolean) =>
                            commitOfficeDeckEdit(
                              deckWithOfficeLayoutTitle(
                                themeRef.current,
                                slidesRef.current,
                                officeMasterSelectedLayoutId,
                                on,
                              ),
                            ),
                          onFootersChange: (on: boolean) => {
                            const next = themeWithOfficeLayoutFooters(
                              themeRef.current,
                              officeMasterSelectedLayoutId,
                              on,
                            );
                            if (next) applyDeckTheme(next);
                          },
                          ...(officeMasterDeleteReason === undefined
                            ? {
                                onDelete: () => {
                                  const index = officeMasterEntries.findIndex(
                                    (entry) => entry.key === officeMasterSelectedLayoutId,
                                  );
                                  const next = themeWithDeletedOfficeLayout(
                                    themeRef.current,
                                    slidesRef.current,
                                    officeMasterSelectedLayoutId,
                                  );
                                  if (!next) return;
                                  applyDeckTheme(next);
                                  // PowerPoint selects the thumbnail above the deleted one.
                                  setOfficeMasterKey(officeMasterEntries[Math.max(0, index - 1)]?.key ?? "master");
                                },
                              }
                            : {}),
                        }
                      : {})}
                    themesMenu={
                      <DropdownMenu>
                        <DropdownMenuTrigger asChild>
                          <OfficeHomeMenuTrigger
                            disabled={slideMasterThemeDisabled}
                            disabledReason={slideMasterThemeReason}
                            hasMenuIndicator
                            icon={<Palette size={24} />}
                            label={t("presentation.theme.label")}
                            testId="presentation-slide-master-themes"
                            variant="large"
                          />
                        </DropdownMenuTrigger>
                        <OfficeGalleryMenuContent
                          data-testid="presentation-slide-master-themes-menu"
                        >
                          <div
                            className={officeGalleryMenuStyles.grid}
                              style={officeGalleryMenuGridStyle(6, 124)}
                            role="grid"
                            aria-label={t("presentation.theme.label")}
                          >
                            {officeThemeTiles.map((tile) => (
                              <DropdownMenuItem
                                asChild
                                key={`slide-master-theme-${tile.id}`}
                                onSelect={() => applyThemePreset(tile.id)}
                              >
                                <button
                                  type="button"
                                  className={officeGalleryMenuStyles.tile}
                                  aria-pressed={activeOfficeTheme?.id === tile.id}
                                  data-testid={`presentation-slide-master-theme-${tile.id}`}
                                >
                                  <span
                                    className={presentationChrome.secondaryExpandedGalleryPreview}
                                    style={{ padding: 0 }}
                                  >
                                    <PresentationOfficeThemeThumbnail theme={tile.design} themeId={tile.id} />
                                  </span>
                                  <span className={officeGalleryMenuStyles.name}>
                                    {tile.label}
                                  </span>
                                </button>
                              </DropdownMenuItem>
                            ))}
                          </div>
                        </OfficeGalleryMenuContent>
                      </DropdownMenu>
                    }
                    colorsMenu={
                      <DropdownMenu>
                        <DropdownMenuTrigger asChild>
                          <OfficeHomeMenuTrigger
                            disabled={slideMasterThemeDisabled}
                            disabledReason={slideMasterThemeReason}
                            hasMenuIndicator
                            icon={<Palette size={16} />}
                            label={t("toolbar.designColors")}
                            testId="presentation-slide-master-colors"
                            variant="compact"
                          />
                        </DropdownMenuTrigger>
                        <RibbonDropdownMenuContent align="start" className={spreadsheetThemeAxisStyles.colorMenu}>
                          <SpreadsheetThemeColorGallery
                            disabled={themeSelectionDisabled}
                            locale={locale}
                            value={theme?.colorSchemeId ?? ""}
                            onValueChange={applyThemeColorScheme}
                            testId="presentation-slide-master-color-gallery"
                            itemTestIdPrefix="presentation-slide-master-color-"
                          />
                        </RibbonDropdownMenuContent>
                      </DropdownMenu>
                    }
                    fontsMenu={
                      <DropdownMenu>
                        <DropdownMenuTrigger asChild>
                          <OfficeHomeMenuTrigger
                            disabled={slideMasterThemeDisabled}
                            disabledReason={slideMasterThemeReason}
                            hasMenuIndicator
                            icon={<TypeIcon size={16} />}
                            label={t("toolbar.designFonts")}
                            testId="presentation-slide-master-fonts"
                            variant="compact"
                          />
                        </DropdownMenuTrigger>
                        <RibbonDropdownMenuContent align="start" className={spreadsheetThemeAxisStyles.fontMenu}>
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
                            testId="presentation-slide-master-font-gallery"
                            itemTestIdPrefix="presentation-slide-master-font-"
                          />
                        </RibbonDropdownMenuContent>
                      </DropdownMenu>
                    }
                    effectsMenu={
                      <DropdownMenu>
                        <DropdownMenuTrigger asChild>
                          <OfficeHomeMenuTrigger
                            disabled={slideMasterThemeDisabled}
                            disabledReason={slideMasterThemeReason}
                            hasMenuIndicator
                            icon={<Sparkles size={16} />}
                            label={t("toolbar.designEffects")}
                            testId="presentation-slide-master-effects"
                            variant="compact"
                          />
                        </DropdownMenuTrigger>
                        <RibbonDropdownMenuContent align="start" className={officeGalleryMenuStyles.menu}>
                          {presentationEffectGallery("slide-master")}
                        </RibbonDropdownMenuContent>
                      </DropdownMenu>
                    }
                    backgroundStylesMenu={
                      <DropdownMenu>
                        <DropdownMenuTrigger asChild>
                          <OfficeHomeMenuTrigger
                            disabled={officeMasterSelectedIsInserted}
                            disabledReason={slideMasterThemeReason}
                            hasMenuIndicator
                            icon={<IconDesignBackground home size={16} />}
                            label={t("presentation.slideMaster.backgroundStyles")}
                            testId="presentation-slide-master-background-styles"
                            variant="compact"
                          />
                        </DropdownMenuTrigger>
                        <RibbonDropdownMenuContent align="start" data-testid="presentation-slide-master-background-styles-menu">
                          {presentationBackgroundStyleItems("slide-master")}
                          <DropdownMenuSeparator />
                          <DropdownMenuItem
                            disabled={
                              officeMasterEntries.length === 0 || autosave.readOnly || officeMasterSelectedIsInserted
                            }
                            onSelect={() => setMasterBackgroundPaneOpen(true)}
                            data-testid="presentation-slide-master-format-background"
                          >
                            {t("presentation.backgroundPane.title")}...
                          </DropdownMenuItem>
                          <DropdownMenuItem
                            disabled={
                              officeMasterEntries.length === 0 ||
                              autosave.readOnly ||
                              officeMasterSelectedIsInserted ||
                              theme?.officeMaster?.background !== true
                            }
                            onSelect={() => {
                              const next = themeWithoutOfficeMasterBackground(themeRef.current);
                              if (next) applyDeckTheme(next);
                            }}
                            data-testid="presentation-slide-master-reset-background"
                          >
                            {t("presentation.backgroundMenu.reset")}
                          </DropdownMenuItem>
                        </RibbonDropdownMenuContent>
                      </DropdownMenu>
                    }
                    slideSizeMenu={slideSizeMenu}
                    onClose={() => dispatchPresentationView("normal")}
                    {...(officeMasterEntries.length > 0 && !autosave.readOnly && !themeSelectionDisabled
                      ? {
                          onRename: () => setMasterRenameOpen(true),
                          // Preserve on the selected card's master (PPT-B030).
                          preserve: {
                            pressed: officeMasterPreserved(theme, officeMasterSelectedMasterKey),
                            onToggle: () => {
                              const next = themeWithOfficeMasterPreserved(
                                themeRef.current,
                                !officeMasterPreserved(themeRef.current, officeMasterSelectedMasterKey),
                                officeMasterSelectedMasterKey,
                              );
                              if (next) applyDeckTheme(next);
                            },
                          },
                        }
                      : {})}
                  />
  );
}
