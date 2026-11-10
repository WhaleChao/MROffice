"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The deck's own master pane (assign, create, delete and rename a master),
 * shown in Slide Master view when the deck has no Office master entries.
 *
 * Extracted verbatim from the render of
 * `components/notebook/contents/PresentationWindow.tsx`: the JSX is unchanged;
 * the window passes the bindings it used to close over as `ctl`, and the render
 * condition and key stay at the call site. No hooks run here, so the window's
 * hook order is unchanged.
 */

import presentationChrome from "@mrnote/core/web/lib/presentation/presentation-window-shell.module.css";
import { PRESENTATION_PANE_CHROME } from "../../../../lib/presentation/window-constants";
import { officeDialogParts as dialogParts } from "@mrnote/core/web/components/office-ribbon/officeDialogChrome.tsx";
import { X } from "@mrnote/core/web/components/office-ribbon/icons/houseIcons.tsx";
import { MAX_MASTERS, type SlideMaster, type SlideMasterLayout } from "@mrnote/core/web/lib/presentation-master.ts";
import { normalizeLayoutId, LAYOUT_DEFINITIONS, type LayoutId } from "@mrnote/core/web/lib/presentation-layouts.ts";
import { MasterLayoutThumbnail, masterLayoutPreviewSlide } from "./PresentationSlideThumbnails";
import {
  normalizeThemeColor,
  normalizeThemeFontFamily,
  type ThemeFontFamily,
  THEME_FONT_FAMILIES,
  MIN_THEME_FONT_SIZE,
  MAX_THEME_FONT_SIZE,
  type DeckTheme,
} from "@mrnote/core/web/lib/presentation-theme.ts";
import { parseThemeSizeInput } from "./PresentationThemePreview";
import type { useTranslations } from "@mrnote/core/host/next-intl.tsx";
import type { Dispatch, RefObject, SetStateAction } from "react";
import type { SlideModel } from "@mrnote/core/web/lib/presentation/deck-model.ts";

/** The window bindings this tab reads, passed through unchanged. */
export interface PresentationSlideMasterPaneCtl {
  t: ReturnType<typeof useTranslations<"console-notebooks">>;
  masterPropertiesRef: RefObject<HTMLElement | null>;
  theme: DeckTheme | undefined;
  masters: SlideMaster[];
  /** `masters` with the layout cards drawn so far, for the thumbnails only
   *  (lib/presentation-layout-cards). */
  previewMasters?: SlideMaster[];
  activeSlide: SlideModel;
  activeMaster: SlideMaster | undefined;
  activeNativeLayout: SlideMasterLayout | undefined;
  activeMasterText: string;
  createMaster: () => string | undefined;
  renameMaster: (masterId: string, name: string) => void;
  updateMasterLayout: (masterId: string, layoutId: LayoutId | undefined) => void;
  updateMasterThemeField: (masterId: string, patch: Partial<DeckTheme>) => void;
  updateMasterText: (masterId: string, text: string) => void;
  deleteMaster: (masterId: string) => void;
  assignActiveSlideMaster: (masterId: string | undefined) => void;
  assignActiveSlideMasterLayout: (masterId: string, layoutId: string | undefined) => void;
  setMasterPaneOpen: Dispatch<SetStateAction<boolean>>;
}

export function PresentationSlideMasterPane({ ctl }: { ctl: PresentationSlideMasterPaneCtl }) {
  const {
    t,
    masterPropertiesRef,
    theme,
    masters,
    previewMasters,
    activeSlide,
    activeMaster,
    activeNativeLayout,
    activeMasterText,
    createMaster,
    renameMaster,
    updateMasterLayout,
    updateMasterThemeField,
    updateMasterText,
    deleteMaster,
    assignActiveSlideMaster,
    assignActiveSlideMasterLayout,
    setMasterPaneOpen,
  } = ctl;
  return (
            <aside
              ref={masterPropertiesRef}
              className={`${presentationChrome.taskPane} ${PRESENTATION_PANE_CHROME}`}
              role="dialog"
              aria-modal="false"
              aria-labelledby="presentation-master-pane-title"
              data-presentation-view-layer="master-pane"
              data-testid="presentation-master-props"
            >
              <header className={`${presentationChrome.taskPaneHeader} ${dialogParts.header}`}>
                <span>
                  <strong id="presentation-master-pane-title">
                    {t("presentation.master.toolbarLabel")}
                  </strong>
                  <small>{t("presentation.master.paneDescription")}</small>
                </span>
                <button
                  type="button"
                  className={`${presentationChrome.taskPaneClose} ${dialogParts.close}`}
                  onClick={() => setMasterPaneOpen(false)}
                  aria-label={t("common.close")}
                  title={t("common.close")}
                  data-testid="presentation-master-pane-close"
                >
                  <X size={16} aria-hidden="true" />
                </button>
              </header>
              <section className={presentationChrome.taskPaneSection}>
                <span className={presentationChrome.taskPaneSectionLabel}>
                  {t("presentation.master.label")}
                </span>
                <label className={presentationChrome.taskPaneField}>
                  <span>{t("presentation.master.assign")}</span>
                  <select
                    // Assign / clear the ACTIVE slide's master. The empty value
                    // clears the reference so the slide reverts to the theme-only
                    // chain. Reflects the slide's persisted (and still-valid)
                    // masterId.
                    value={activeMaster?.id ?? ""}
                    onChange={(event) =>
                      assignActiveSlideMaster(event.target.value || undefined)
                    }
                    data-testid="presentation-master-assign"
                    aria-label={t("presentation.master.assign")}
                  >
                    <option value="">{t("presentation.master.none")}</option>
                    {masters.map((master) => (
                      <option key={master.id} value={master.id}>
                        {master.name}
                      </option>
                    ))}
                  </select>
                </label>
                <div className={presentationChrome.taskPaneRow}>
                  <button
                    type="button"
                    className={presentationChrome.taskPaneButton}
                    onClick={() => {
                      const id = createMaster();
                      if (id) assignActiveSlideMaster(id);
                    }}
                    disabled={masters.length >= MAX_MASTERS}
                    data-testid="presentation-master-create"
                  >
                    {t("presentation.master.create")}
                  </button>
                  {activeMaster ? (
                    <button
                      type="button"
                      className={presentationChrome.taskPaneButton}
                      onClick={() => deleteMaster(activeMaster.id)}
                      data-testid="presentation-master-delete"
                    >
                      {t("presentation.master.delete")}
                    </button>
                  ) : null}
                </div>
              </section>
              {activeMaster ? (
                <>
                  <section className={presentationChrome.taskPaneSection}>
                    <span className={presentationChrome.taskPaneSectionLabel}>
                      {t("presentation.master.sectionEdit")}
                    </span>
                    <label className={presentationChrome.taskPaneField}>
                      <span>{t("presentation.master.name")}</span>
                      <input
                        type="text"
                        value={activeMaster.name}
                        onChange={(event) =>
                          renameMaster(activeMaster.id, event.target.value)
                        }
                        data-testid="presentation-master-name"
                        aria-label={t("presentation.master.name")}
                        maxLength={200}
                      />
                    </label>
                    {/* Master PLACEHOLDER TEXT: a single inherited text block
                        shown on every slide using this master (unless the slide
                        overrides it). Typing writes the block; clearing removes
                        it. Cascades live via `updateMasterText`. */}
                    <label className={presentationChrome.taskPaneField}>
                      <span>{t("presentation.master.placeholderText")}</span>
                      <input
                        type="text"
                        value={activeMasterText}
                        onChange={(event) =>
                          updateMasterText(activeMaster.id, event.target.value)
                        }
                        data-testid="presentation-master-text"
                        aria-label={t("presentation.master.placeholderText")}
                        placeholder={t("presentation.master.placeholderTextHint")}
                        maxLength={2000}
                      />
                    </label>
                  </section>
                  <section className={presentationChrome.taskPaneSection}>
                    <span className={presentationChrome.taskPaneSectionLabel}>
                      {t("presentation.master.sectionLayout")}
                    </span>
                    <label className={presentationChrome.taskPaneField}>
                      <span>{t("presentation.master.layout")}</span>
                      <select
                        // The master's inherited layout id. The empty value
                        // clears it so the master contributes no geometry
                        // (slides fall through to their own layout / the
                        // default).
                        value={activeMaster.layoutId ?? ""}
                        onChange={(event) =>
                          updateMasterLayout(
                            activeMaster.id,
                            event.target.value
                              ? normalizeLayoutId(event.target.value)
                              : undefined,
                          )
                        }
                        data-testid="presentation-master-layout"
                        aria-label={t("presentation.master.layout")}
                      >
                        <option value="">
                          {t("presentation.master.layoutNone")}
                        </option>
                        {LAYOUT_DEFINITIONS.map((layout) => (
                          <option key={layout.id} value={layout.id}>
                            {t(
                              `presentation.layout.${layout.nameKey}` as "presentation.layout.blank",
                            )}
                          </option>
                        ))}
                      </select>
                    </label>
                    {activeMaster.layouts && activeMaster.layouts.length > 0 ? (
                      <div
                        data-testid="presentation-master-layout-gallery"
                        role="listbox"
                        aria-label={t("presentation.master.layout")}
                        className={presentationChrome.taskPaneGallery}
                      >
                        {activeMaster.layouts.map((layout) => {
                          const selected = activeNativeLayout?.id === layout.id;
                          return (
                            <button
                              key={layout.id}
                              type="button"
                              role="option"
                              aria-selected={selected}
                              className={presentationChrome.taskPaneGalleryOption}
                              onClick={() =>
                                assignActiveSlideMasterLayout(
                                  activeMaster.id,
                                  layout.id,
                                )
                              }
                              title={layout.sourcePart ?? layout.name}
                              data-testid="presentation-master-layout-option"
                              data-layout-id={layout.id}
                              data-layout-type={
                                layout.nativeType ??
                                layout.builtInLayoutId ??
                                "custom"
                              }
                            >
                              <span style={{ display: "block", width: 72 }}>
                                {(() => {
                                  // A layout no slide uses draws its card once the server has drawn it.
                                  const previewMaster = previewMasters?.find((master) => master.id === activeMaster.id) ?? activeMaster;
                                  const previewLayout = previewMaster.layouts?.find((item) => item.id === layout.id) ?? layout;
                                  return (
                                    <MasterLayoutThumbnail
                                      layout={previewLayout}
                                      slide={masterLayoutPreviewSlide(
                                        previewMaster,
                                        previewLayout,
                                        activeSlide,
                                      )}
                                      theme={theme}
                                      master={previewMaster}
                                      masters={previewMasters ?? masters}
                                    />
                                  );
                                })()}
                              </span>
                              <span style={{ minWidth: 0, display: "grid", gap: 3 }}>
                                <span
                                  style={{
                                    overflow: "hidden",
                                    textOverflow: "ellipsis",
                                    whiteSpace: "nowrap",
                                    fontSize: 11,
                                    fontWeight: selected ? 700 : 550,
                                  }}
                                >
                                  {layout.name}
                                </span>
                                {layout.sourcePart ? (
                                  <span
                                    style={{
                                      fontSize: 9,
                                      letterSpacing: "0.08em",
                                      color: "var(--ppt-accent)",
                                    }}
                                  >
                                    OOXML
                                  </span>
                                ) : null}
                              </span>
                            </button>
                          );
                        })}
                      </div>
                    ) : null}
                  </section>
                  <section className={presentationChrome.taskPaneSection}>
                    <span className={presentationChrome.taskPaneSectionLabel}>
                      {t("presentation.master.sectionBackground")}
                    </span>
                    <label className={presentationChrome.taskPaneCheck}>
                      <input
                        type="color"
                        value={
                          normalizeThemeColor(activeMaster.theme?.background) ??
                          "#ffffff"
                        }
                        onChange={(event) =>
                          updateMasterThemeField(activeMaster.id, {
                            background: event.target.value,
                          })
                        }
                        data-testid="presentation-master-bg"
                        aria-label={t("presentation.master.background")}
                      />
                      {t("presentation.master.background")}
                    </label>
                  </section>
                  <section className={presentationChrome.taskPaneSection}>
                    <span className={presentationChrome.taskPaneSectionLabel}>
                      {t("presentation.master.sectionText")}
                    </span>
                    <label className={presentationChrome.taskPaneField}>
                      <span>{t("presentation.master.font")}</span>
                      <select
                        value={
                          normalizeThemeFontFamily(
                            activeMaster.theme?.fontFamily,
                          ) ?? ""
                        }
                        onChange={(event) =>
                          updateMasterThemeField(activeMaster.id, {
                            fontFamily: event.target.value
                              ? (event.target.value as ThemeFontFamily)
                              : undefined,
                          })
                        }
                        data-testid="presentation-master-font"
                        aria-label={t("presentation.master.font")}
                      >
                        <option value="">
                          {t("presentation.master.fontInherit")}
                        </option>
                        {THEME_FONT_FAMILIES.map((font) => (
                          <option key={font} value={font}>
                            {t(
                              `presentation.theme.font.${font}` as "presentation.theme.font.system",
                            )}
                          </option>
                        ))}
                      </select>
                    </label>
                    {/* Master TITLE / BODY text styles (default size px + bold
                        + colour), layered on top of the deck theme. Empty size
                        = inherit. A per-block style still overrides; a field
                        absent here falls through to the deck theme then the
                        built-in default. */}
                    <div className={presentationChrome.taskPaneRow}>
                      <label className={presentationChrome.taskPaneField}>
                        <span>{t("presentation.master.titleSizeLabel")}</span>
                        <input
                          type="number"
                          min={MIN_THEME_FONT_SIZE}
                          max={MAX_THEME_FONT_SIZE}
                          value={activeMaster.theme?.titleFontSize ?? ""}
                          onChange={(event) =>
                            updateMasterThemeField(activeMaster.id, {
                              titleFontSize: parseThemeSizeInput(event.target.value),
                            })
                          }
                          data-testid="presentation-master-title-size"
                          aria-label={t("presentation.master.titleSizeLabel")}
                          placeholder={t("presentation.theme.sizeInherit")}
                        />
                      </label>
                      <label className={presentationChrome.taskPaneField}>
                        <span>{t("presentation.master.bodySizeLabel")}</span>
                        <input
                          type="number"
                          min={MIN_THEME_FONT_SIZE}
                          max={MAX_THEME_FONT_SIZE}
                          value={activeMaster.theme?.bodyFontSize ?? ""}
                          onChange={(event) =>
                            updateMasterThemeField(activeMaster.id, {
                              bodyFontSize: parseThemeSizeInput(event.target.value),
                            })
                          }
                          data-testid="presentation-master-body-size"
                          aria-label={t("presentation.master.bodySizeLabel")}
                          placeholder={t("presentation.theme.sizeInherit")}
                        />
                      </label>
                    </div>
                    <div className={presentationChrome.taskPaneRow}>
                      <label className={presentationChrome.taskPaneCheck}>
                        <input
                          type="checkbox"
                          checked={activeMaster.theme?.titleBold === true}
                          onChange={(event) =>
                            updateMasterThemeField(activeMaster.id, {
                              titleBold: event.target.checked ? true : undefined,
                            })
                          }
                          data-testid="presentation-master-title-bold"
                          aria-label={t("presentation.master.titleBoldLabel")}
                        />
                        {t("presentation.master.titleBoldLabel")}
                      </label>
                      <label className={presentationChrome.taskPaneCheck}>
                        <input
                          type="checkbox"
                          checked={activeMaster.theme?.bodyBold === true}
                          onChange={(event) =>
                            updateMasterThemeField(activeMaster.id, {
                              bodyBold: event.target.checked ? true : undefined,
                            })
                          }
                          data-testid="presentation-master-body-bold"
                          aria-label={t("presentation.master.bodyBoldLabel")}
                        />
                        {t("presentation.master.bodyBoldLabel")}
                      </label>
                    </div>
                    <div className={presentationChrome.taskPaneRow}>
                      {(["title", "body"] as const).map((style) => {
                        const key = style === "title" ? "titleItalic" : "bodyItalic";
                        const label = t(`presentation.master.${key}Label`);
                        return (
                          <label key={key} className={presentationChrome.taskPaneCheck}>
                            <input
                              type="checkbox"
                              checked={activeMaster.theme?.[key] === true}
                              onChange={(event) =>
                                updateMasterThemeField(activeMaster.id, {
                                  [key]: event.target.checked ? true : undefined,
                                })
                              }
                              data-testid={`presentation-master-${style}-italic`}
                              aria-label={label}
                            />
                            {label}
                          </label>
                        );
                      })}
                    </div>
                    <div className={presentationChrome.taskPaneRow}>
                      <label className={presentationChrome.taskPaneCheck}>
                        <input
                          type="color"
                          value={
                            normalizeThemeColor(activeMaster.theme?.titleColor) ??
                            "#1e293b"
                          }
                          onChange={(event) =>
                            updateMasterThemeField(activeMaster.id, {
                              titleColor: event.target.value,
                            })
                          }
                          data-testid="presentation-master-title-color"
                          aria-label={t("presentation.master.titleColor")}
                        />
                        {t("presentation.master.titleColor")}
                      </label>
                      <label className={presentationChrome.taskPaneCheck}>
                        <input
                          type="color"
                          value={
                            normalizeThemeColor(activeMaster.theme?.bodyTextColor) ??
                            "#334155"
                          }
                          onChange={(event) =>
                            updateMasterThemeField(activeMaster.id, {
                              bodyTextColor: event.target.value,
                            })
                          }
                          data-testid="presentation-master-body-color"
                          aria-label={t("presentation.master.bodyColor")}
                        />
                        {t("presentation.master.bodyColor")}
                      </label>
                    </div>
                  </section>
                </>
              ) : null}
            </aside>
  );
}
