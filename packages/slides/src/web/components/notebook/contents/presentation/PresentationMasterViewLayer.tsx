"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The Slide Master view layer of the stage: the Office master editor with its
 * cards (including masters Reuse Slides brought in) or, for a deck with its
 * own masters, their card grid.
 *
 * Extracted verbatim from the render of
 * `components/notebook/contents/PresentationWindow.tsx`: the JSX is unchanged;
 * the window passes the bindings it used to close over as `ctl`, and the render
 * condition and key stay at the call site. No hooks run here, so the window's
 * hook order is unchanged.
 */

import { PRESENTATION_RIBBON_LABELED_BUTTON_STYLE } from "../../../../lib/presentation/window-chrome-styles";
import presentationChrome from "@mrnote/core/web/lib/presentation/presentation-window-shell.module.css";
import { PresentationOfficeMasterTextStyleBar } from "./PresentationOfficeMasterTextStyleBar";
import {
  officeMasterEffectiveTextValues,
  officeMasterTextValues,
  themeWithOfficeMasterTextStyle,
  officeMasterGivenName,
  officeMasterCardFrames,
  themeWithOfficeMasterFrame,
  type OfficeMasterTextLevel,
  type OfficeMasterViewEntry,
} from "../../../../lib/presentation-office-master";
import { PresentationSlideMasterView } from "./PresentationSlideMasterView";
import { PresentationOfficeMasterPreview, MasterLayoutThumbnail } from "./PresentationSlideThumbnails";
import { officeMasterTheme } from "../../../../lib/presentation-office-master-routing";
import { PresentationMasterFrameEditor } from "./PresentationMasterFrameEditor";
import { themeWithOfficeMasterPlaceholder } from "../../../../lib/presentation-office-master-structure";
import { deckWithDeletedOfficePlaceholder, type OfficeDeckEdit } from "../../../../lib/presentation-office-layout-edit";
import { RibbonValueCombobox } from "@mrnote/core/web/components/office-ribbon/RibbonValueCombobox.tsx";
import { FONT_SIZE_VALUES, TEXT_COLOR_VALUES } from "@mrnote/core/web/components/console/editor/extensions/TextStyling.ts";
import { parseRunFontSizeInput } from "@mrnote/core/web/lib/presentation-point-runs.ts";
import { DropdownMenu, DropdownMenuTrigger, DropdownMenuSeparator } from "@mrnote/core/web/components/ui/dropdown-menu.tsx";
import { OfficeHomeMenuTrigger } from "@mrnote/core/web/components/office-ribbon/OfficeHomeRibbon.tsx";
import { DocsHomeIcon } from "@mrnote/core/web/components/console/editor/DocumentHomeRibbonIcons.tsx";
import { RibbonDropdownMenuContent as DropdownMenuContent } from "@mrnote/core/web/components/office-ribbon/CupertinoRibbonMenu.tsx";
import { PresentationSwatchGrid, PresentationCustomColorItem } from "@mrnote/core/web/components/notebook/contents/ribbon/ColorMenu.tsx";
import { LayoutTemplate, X } from "@mrnote/core/web/components/office-ribbon/icons/houseIcons.tsx";
import type { useTranslations } from "@mrnote/core/host/next-intl.tsx";
import type { SlideModel } from "@mrnote/core/web/lib/presentation/deck-model.ts";
import type { DeckTheme } from "@mrnote/core/web/lib/presentation-theme.ts";
import type { Dispatch, RefObject, SetStateAction } from "react";
import type { SlideMaster, SlideMasterLayout } from "@mrnote/core/web/lib/presentation-master.ts";
import type { embeddedPresentationAutosave } from "../../../../lib/presentation-embedded-session";
import type { PresentationLiveView } from "../../../../lib/presentation/view-zoom";

/** The window bindings this tab reads, passed through unchanged. */
export interface PresentationMasterViewLayerCtl {
  t: ReturnType<typeof useTranslations<"console-notebooks">>;
  slides: SlideModel[];
  theme: DeckTheme | undefined;
  themeRef: RefObject<DeckTheme | undefined>;
  masters: SlideMaster[];
  masterFontSizeDraft: { key: string; value: string; } | null;
  setMasterFontSizeDraft: Dispatch<SetStateAction<{ key: string; value: string; } | null>>;
  slidesRef: RefObject<SlideModel[]>;
  autosave: ReturnType<typeof embeddedPresentationAutosave>;
  themeSelectionDisabled: boolean;
  applyDeckTheme: (nextTheme: DeckTheme | undefined) => void;
  assignActiveSlideMasterLayout: (masterId: string, layoutId: string | undefined) => void;
  updateMasterTextStyle: (masterId: string, patch: Partial<DeckTheme>) => void;
  masterPaneOpen: boolean;
  setMasterPaneOpen: Dispatch<SetStateAction<boolean>>;
  officeMasterKey: string;
  setOfficeMasterKey: Dispatch<SetStateAction<string>>;
  officeMasterTextTarget: OfficeMasterTextLevel;
  setOfficeMasterTextTarget: Dispatch<SetStateAction<OfficeMasterTextLevel>>;
  dispatchPresentationView: (next: PresentationLiveView, options?: { trigger?: HTMLElement | null; restoreReadingFocus?: boolean; }) => boolean;
  masterViewLayouts: { masterId: string; masterName: string; master: SlideMaster; layout: SlideMasterLayout; slide: SlideModel; usage: number; selected: boolean; }[];
  masterViewMasters: { master: SlideMaster; slide: SlideModel; layoutCount: number; usage: number; selected: boolean; }[];
  officeMasterPrompts: { title: string; subtitle: string; textLevels: string[]; kinds: { [k: string]: string; }; };
  officeMasterEntries: OfficeMasterViewEntry[];
  officeMasterReusedSelected: boolean;
  officeMasterLayoutWritable: boolean;
  officeMasterSelectedMasterKey: string;
  officePlaceholderKindLabels: Record<"chart" | "table" | "smartart" | "media" | "online-image" | "content" | "content-vertical" | "text" | "text-vertical" | "picture", string>;
  commitOfficeDeckEdit: (edit: OfficeDeckEdit<SlideModel> | null) => void;
  officeSlideMasterDisplayName: (masterKey: string) => string;
  officeMasterUsageLabel: (entry: OfficeMasterViewEntry, deckSlides: readonly SlideModel[]) => string;
  masterTextStyleTarget: SlideMaster | undefined;
  masterTextStyleLevels: { key: string; label: string; fontSize: number | undefined; color: string | undefined; withFontSize: (titleFontSize: number) => Partial<DeckTheme>; withColor: (titleColor: string) => Partial<DeckTheme>; }[];
  masterViewNativeReferences: { masterId: string; layoutId: string; usage: number; }[];
}

export function PresentationMasterViewLayer({ ctl }: { ctl: PresentationMasterViewLayerCtl }) {
  const {
    t,
    slides,
    theme,
    themeRef,
    masters,
    masterFontSizeDraft,
    setMasterFontSizeDraft,
    slidesRef,
    autosave,
    themeSelectionDisabled,
    applyDeckTheme,
    assignActiveSlideMasterLayout,
    updateMasterTextStyle,
    masterPaneOpen,
    setMasterPaneOpen,
    officeMasterKey,
    setOfficeMasterKey,
    officeMasterTextTarget,
    setOfficeMasterTextTarget,
    dispatchPresentationView,
    masterViewLayouts,
    masterViewMasters,
    officeMasterPrompts,
    officeMasterEntries,
    officeMasterReusedSelected,
    officeMasterLayoutWritable,
    officeMasterSelectedMasterKey,
    officePlaceholderKindLabels,
    commitOfficeDeckEdit,
    officeSlideMasterDisplayName,
    officeMasterUsageLabel,
    masterTextStyleTarget,
    masterTextStyleLevels,
    masterViewNativeReferences,
  } = ctl;
  return (
            <div
              data-presentation-view-layer="master"
              className={presentationChrome.stageViewLayer}
              data-testid="presentation-master-view"
            >
              {officeMasterEntries.length > 0 ? (
                /* An Office theme deck's one master, as PowerPoint's Slide
                   Master view shows it (PPT-B030): its text styles above the
                   thumbnail pane and the selected card at stage size. */
                <>
                  <div
                    className={presentationChrome.masterBar}
                    data-testid="presentation-master-view-bar"
                  >
                    <PresentationOfficeMasterTextStyleBar
                      labels={{
                        level: t("presentation.master.sectionText"),
                        title: t("presentation.master.levelTitle"),
                        bodyLevels: [
                          t("presentation.master.levelBody"),
                          ...officeMasterPrompts.textLevels.slice(1, 5),
                        ],
                        fontSize: t("toolbar.fontSize"),
                        bold: t("toolbar.bold"),
                        italic: t("toolbar.italic"),
                        color: t("toolbar.textColor"),
                        customColor: t("toolbar.textColorCustom"),
                        reset: t("presentation.master.resetLevel"),
                      }}
                      target={officeMasterTextTarget}
                      onTarget={setOfficeMasterTextTarget}
                      // The selected card's master's text styles (PPT-B030).
                      values={officeMasterEffectiveTextValues(
                        theme,
                        officeMasterTextTarget,
                        officeMasterSelectedMasterKey,
                      )}
                      overridden={
                        Object.keys(
                          officeMasterTextValues(theme, officeMasterTextTarget, officeMasterSelectedMasterKey),
                        ).length > 0
                      }
                      disabled={autosave.readOnly || themeSelectionDisabled || officeMasterReusedSelected}
                      onPatch={(patch) => {
                        const next = themeWithOfficeMasterTextStyle(
                          themeRef.current,
                          officeMasterTextTarget,
                          patch,
                          officeMasterSelectedMasterKey,
                        );
                        if (next) applyDeckTheme(next);
                      }}
                      onReset={() => {
                        const next = themeWithOfficeMasterTextStyle(
                          themeRef.current,
                          officeMasterTextTarget,
                          { fontSize: null, bold: null, italic: null, color: null },
                          officeMasterSelectedMasterKey,
                        );
                        if (next) applyDeckTheme(next);
                      }}
                    />
                  </div>
                  <PresentationSlideMasterView
                    treeLabel={t("presentation.master.layoutsGrid")}
                    selectedKey={officeMasterKey}
                    onSelect={setOfficeMasterKey}
                    cards={officeMasterEntries.map((entry) => ({
                      key: entry.key,
                      kind: entry.kind,
                      label:
                        entry.kind === "master"
                          ? t("presentation.master.officeMasterName", {
                              theme: officeSlideMasterDisplayName(entry.masterKey),
                            })
                          : t("presentation.master.layoutName", {
                              layout:
                                officeMasterGivenName(theme, entry.key) ??
                                t(
                                  `presentation.layout.${entry.nameKey ?? "blank"}` as "presentation.layout.blank",
                                ),
                            }),
                      usage: officeMasterUsageLabel(entry, slides),
                      // An inserted master's cards draw under its own theme (PPT-B030).
                      preview: (
                        <PresentationOfficeMasterPreview
                          slide={entry.slide}
                          theme={officeMasterTheme(theme, entry.masterKey) ?? theme}
                        />
                      ),
                    })).concat(
                      // Masters Reuse Slides brought in, each above its
                      // layouts, after the deck's own (kept as they are).
                      masterViewMasters.flatMap((entry) => [
                        {
                          key: `reused-master:${entry.master.id}`,
                          kind: "master" as const,
                          label: entry.master.name,
                          usage: t("presentation.master.cardUsage", {
                            layouts: entry.layoutCount,
                            slides: entry.usage,
                          }),
                          preview: (
                            <MasterLayoutThumbnail
                              layout={{ id: entry.master.id, name: entry.master.name }}
                              slide={entry.slide}
                              theme={theme}
                              master={entry.master}
                              masters={masters}
                            />
                          ),
                        },
                        ...masterViewLayouts
                          .filter((layoutEntry) => layoutEntry.masterId === entry.master.id)
                          .map((layoutEntry) => ({
                            key: `reused-layout:${layoutEntry.masterId}:${layoutEntry.layout.id}`,
                            kind: "layout" as const,
                            label: layoutEntry.layout.name,
                            usage:
                              layoutEntry.usage > 0
                                ? t("presentation.master.layoutUsage", { count: layoutEntry.usage })
                                : t("presentation.master.layoutUnused"),
                            preview: (
                              <MasterLayoutThumbnail
                                layout={layoutEntry.layout}
                                slide={layoutEntry.slide}
                                theme={theme}
                                master={layoutEntry.master}
                                masters={masters}
                              />
                            ),
                          })),
                      ]),
                    )}
                    stage={(() => {
                      if (officeMasterReusedSelected) {
                        const [, masterId, layoutId] = officeMasterKey.split(":");
                        const masterEntry = masterViewMasters.find((entry) => entry.master.id === masterId);
                        const layoutEntry = masterViewLayouts.find(
                          (entry) => entry.masterId === masterId && entry.layout.id === layoutId,
                        );
                        if (!masterEntry) return null;
                        return (
                          <div data-testid="presentation-master-view-reused">
                            <MasterLayoutThumbnail
                              layout={layoutEntry?.layout ?? { id: masterEntry.master.id, name: masterEntry.master.name }}
                              slide={layoutEntry?.slide ?? masterEntry.slide}
                              theme={theme}
                              master={masterEntry.master}
                              masters={masters}
                            />
                            <p
                              className={presentationChrome.backgroundCapabilityNotice}
                              data-testid="presentation-master-view-reused-notice"
                            >
                              {t("presentation.master.reusedKept")}
                            </p>
                          </div>
                        );
                      }
                      const selected =
                        officeMasterEntries.find((entry) => entry.key === officeMasterKey) ??
                        officeMasterEntries[0];
                      if (!selected) return null;
                      // The card's placeholder frames over it, movable and
                      // resizable (PPT-B030): the master's title and body,
                      // or the layout's own placeholders.
                      const cardFrames = officeMasterCardFrames(theme, selected);
                      return (
                        <div style={{ position: "relative" }}>
                          <PresentationOfficeMasterPreview
                            slide={selected.slide}
                            theme={officeMasterTheme(theme, selected.masterKey) ?? theme}
                          />
                          <PresentationMasterFrameEditor
                            disabled={autosave.readOnly || themeSelectionDisabled}
                            frames={cardFrames.map((frame, index) => {
                              // The master's date, footer and slide number
                              // (PPT-B030): shown so Delete can take them off.
                              const footerLabel =
                                frame.target.kind === "master"
                                  ? {
                                      date: t("presentation.slideMaster.masterLayoutDate"),
                                      footer: t("presentation.slideMaster.masterLayoutFooter"),
                                      slideNumber: t("presentation.slideMaster.masterLayoutSlideNumber"),
                                      title: null,
                                      body: null,
                                    }[frame.target.family]
                                  : null;
                              return {
                                key: String(index),
                                label:
                                  footerLabel ??
                                  (frame.kind
                                    ? officePlaceholderKindLabels[frame.kind]
                                    : frame.role === "title"
                                      ? t("presentation.master.levelTitle")
                                      : t("presentation.master.levelBody")),
                                left: frame.left,
                                top: frame.top,
                                width: frame.width,
                                height: frame.height,
                                ...(frame.fixed ? { fixed: true } : {}),
                                ...(footerLabel
                                  ? { text: frame.role === "slideNumber" ? "‹#›" : footerLabel }
                                  : {}),
                              };
                            })}
                            onCommit={(key, rect) => {
                              const frame = cardFrames[Number(key)];
                              if (!frame) return;
                              const next = themeWithOfficeMasterFrame(themeRef.current, frame.target, rect);
                              if (next) applyDeckTheme(next);
                            }}
                            {...(selected.kind === "master" && officeMasterLayoutWritable
                              ? {
                                  // Delete on the master's title, text, date,
                                  // footer or slide number placeholder takes it
                                  // off the master; Master Layout puts it back
                                  // (PPT-B030).
                                  onDelete: (key: string) => {
                                    const target = cardFrames[Number(key)]?.target;
                                    if (target?.kind !== "master") return;
                                    const next = themeWithOfficeMasterPlaceholder(
                                      themeRef.current,
                                      target.family,
                                      false,
                                      target.masterKey,
                                    );
                                    if (next) applyDeckTheme(next);
                                  },
                                }
                              : {})}
                            {...(selected.kind === "layout" && officeMasterLayoutWritable
                              ? {
                                  // Delete on a selected placeholder takes it off the layout (PPT-B030).
                                  onDelete: (key: string) =>
                                    commitOfficeDeckEdit(
                                      deckWithDeletedOfficePlaceholder(
                                        themeRef.current,
                                        slidesRef.current,
                                        selected.layoutId,
                                        Number(key),
                                      ),
                                    ),
                                }
                              : {})}
                          />
                        </div>
                      );
                    })()}
                  />
                </>
              ) : (
              <>
              {/* PowerPoint's Slide Master tab carries one command that leaves
              the view — "Close Master View" (Microsoft Support, "Customize a
              slide master"). We mount no Slide Master ribbon tab, so the view
              carries its own exit, which is also where entering the view puts
              focus. */}
              <div
                className={presentationChrome.masterBar}
                data-testid="presentation-master-view-bar"
              >
                {/* PowerPoint edits master text from the HOME tab while Slide
                Master view is open ("Change the font size", Microsoft
                Support). The same shared controls are mounted here rather than
                a second font group: the size box is the ribbon's own
                `RibbonValueCombobox`, the colours its own swatch grid. They
                address the master the ACTIVE slide inherits from, which is the
                master PowerPoint would have selected in the thumbnail pane. */}
                {masterTextStyleTarget ? (
                  <div
                    className={presentationChrome.masterStyleGroup}
                    data-testid="presentation-master-text-styles"
                    data-master-id={masterTextStyleTarget.id}
                  >
                    {masterTextStyleLevels.map((level) => (
                      <span
                        key={level.key}
                        className={presentationChrome.masterStyleLevel}
                      >
                        <span className={presentationChrome.masterCardUsage}>
                          {level.label}
                        </span>
                        <RibbonValueCombobox
                          testId={`presentation-master-${level.key}-font-size`}
                          label={level.label}
                          placeholder={t("toolbar.fontSize")}
                          inputMode="decimal"
                          className="note-editor-ribbon-value-combobox--font-size"
                          selectOnFocus
                          value={
                            (masterFontSizeDraft?.key === level.key
                              ? masterFontSizeDraft.value
                              : undefined) ?? String(level.fontSize ?? "")
                          }
                          options={FONT_SIZE_VALUES.map((value) => ({
                            value: String(Number.parseFloat(value)),
                            label: String(Number.parseFloat(value)),
                          }))}
                          onDraftChange={(value) =>
                            setMasterFontSizeDraft({ key: level.key, value })
                          }
                          onFocusChange={(focused) => {
                            if (!focused) setMasterFontSizeDraft(null);
                          }}
                          onRememberSelection={() => {}}
                          onCommit={(next) => {
                            const size = parseRunFontSizeInput(next);
                            if (size === undefined) return false;
                            setMasterFontSizeDraft(null);
                            updateMasterTextStyle(
                              masterTextStyleTarget.id,
                              level.withFontSize(size),
                            );
                            return true;
                          }}
                        />
                        <DropdownMenu>
                          <DropdownMenuTrigger asChild>
                            <OfficeHomeMenuTrigger
                              variant="icon"
                              hasMenuIndicator
                              icon={
                                <span
                                  className="note-editor-office-toolbar__color-glyph"
                                  aria-hidden="true"
                                >
                                  <DocsHomeIcon name="fontColor" size={17} />
                                  <span
                                    className="note-editor-office-toolbar__color-underline"
                                    style={{
                                      backgroundColor: level.color ?? "#000000",
                                    }}
                                  />
                                </span>
                              }
                              label={t("toolbar.textColor")}
                              testId={`presentation-master-${level.key}-color-menu`}
                            />
                          </DropdownMenuTrigger>
                          <DropdownMenuContent
                            align="start"
                            className="note-editor-color-menu"
                          >
                            <PresentationSwatchGrid
                              colors={TEXT_COLOR_VALUES}
                              current={level.color}
                              label={t("toolbar.textColor")}
                              onPick={(color) =>
                                updateMasterTextStyle(
                                  masterTextStyleTarget.id,
                                  level.withColor(color),
                                )
                              }
                              testPrefix={`presentation-master-${level.key}-color`}
                              variant="text"
                            />
                            <DropdownMenuSeparator />
                            <PresentationCustomColorItem
                              label={t("toolbar.textColorCustom")}
                              value={level.color ?? "#000000"}
                              onPick={(color) =>
                                updateMasterTextStyle(
                                  masterTextStyleTarget.id,
                                  level.withColor(color),
                                )
                              }
                              testId={`presentation-master-${level.key}-color-custom`}
                            />
                          </DropdownMenuContent>
                        </DropdownMenu>
                      </span>
                    ))}
                  </div>
                ) : null}
                {!masterPaneOpen ? (
                  <button
                    type="button"
                    className="office-ribbon-control note-editor-office-toolbar__button"
                    data-office-control-part="button"
                    data-testid="presentation-master-pane-open"
                    style={PRESENTATION_RIBBON_LABELED_BUTTON_STYLE}
                    onClick={() => setMasterPaneOpen(true)}
                  >
                    <LayoutTemplate size={14} aria-hidden="true" />
                    {t("presentation.master.toolbarLabel")}
                  </button>
                ) : null}
                <button
                  type="button"
                  className="office-ribbon-control note-editor-office-toolbar__button"
                  data-office-control-part="button"
                  data-control-id="PPTCTL.RIBBON.SLIDE_MASTER.CLOSE"
                  data-testid="presentation-master-view-close"
                  style={PRESENTATION_RIBBON_LABELED_BUTTON_STYLE}
                  onClick={() => dispatchPresentationView("normal")}
                >
                  <X size={14} aria-hidden="true" />
                  {t("presentation.view.closeMaster")}
                </button>
              </div>
              {masterViewLayouts.length === 0 && masterViewMasters.length === 0 ? (
                <p
                  className={presentationChrome.masterEmpty}
                  data-testid="presentation-master-view-empty"
                >
                  {masterViewNativeReferences.length === 0 ? (
                    <>{t("presentation.master.viewEmpty")}</>
                  ) : (
                    <span data-testid="presentation-master-view-native-only">
                      {t("presentation.master.viewNativeOnly")}
                      <span
                        className={presentationChrome.masterCardUsage}
                        data-testid="presentation-master-view-native-parts"
                      >
                        {masterViewNativeReferences
                          .map(
                            (entry) =>
                              `${entry.layoutId || entry.masterId} × ${entry.usage}`,
                          )
                          .join(" · ")}
                      </span>
                    </span>
                  )}
                </p>
              ) : (
                <div
                  className={presentationChrome.masterGrid}
                  role="listbox"
                  aria-label={t("presentation.master.layoutsGrid")}
                  data-testid="presentation-master-view-grid"
                >
                  {masterViewMasters.map((entry) => (
                    <div
                      key={`master:${entry.master.id}`}
                      role="presentation"
                      className={`${presentationChrome.masterCard} office-page-enter`}
                      data-testid="presentation-master-view-master"
                      data-master-id={entry.master.id}
                      data-master-selected={entry.selected ? "true" : "false"}
                      title={entry.master.sourcePart ?? entry.master.name}
                      style={{ gridColumn: "1 / -1", cursor: "default" }}
                    >
                      <span className={presentationChrome.masterCardPreview}>
                        <MasterLayoutThumbnail
                          layout={{ id: entry.master.id, name: entry.master.name }}
                          slide={entry.slide}
                          theme={theme}
                          master={entry.master}
                          masters={masters}
                        />
                      </span>
                      <span className={presentationChrome.masterCardLabel}>
                        <span>{entry.master.name}</span>
                        <span className={presentationChrome.masterCardUsage}>
                          {t("presentation.master.cardUsage", {
                            layouts: entry.layoutCount,
                            slides: entry.usage,
                          })}
                        </span>
                      </span>
                      <span className={presentationChrome.masterCardUsage}>
                        {t("presentation.view.slideMaster")}
                      </span>
                    </div>
                  ))}
                  {masterViewLayouts.map((entry) => (
                    <button
                      key={`${entry.masterId}:${entry.layout.id}`}
                      type="button"
                      role="option"
                      aria-selected={entry.selected}
                      className={`${presentationChrome.masterCard} office-page-enter`}
                      data-testid="presentation-master-view-layout"
                      data-layout-id={entry.layout.id}
                      data-master-id={entry.masterId}
                      onClick={() =>
                        assignActiveSlideMasterLayout(
                          entry.masterId,
                          entry.layout.id,
                        )
                      }
                      title={entry.layout.sourcePart ?? entry.layout.name}
                    >
                      {/* The rail's own miniature renderer draws the layout's
                      real blocks; the silhouette only when it has none. */}
                      <span className={presentationChrome.masterCardPreview}>
                        <MasterLayoutThumbnail
                          layout={entry.layout}
                          slide={entry.slide}
                          theme={theme}
                          master={entry.master}
                          masters={masters}
                        />
                      </span>
                      <span className={presentationChrome.masterCardLabel}>
                        <span>{entry.layout.name}</span>
                        <span className={presentationChrome.masterCardUsage}>
                          {entry.usage > 0
                            ? t("presentation.master.layoutUsage", { count: entry.usage })
                            : t("presentation.master.layoutUnused")}
                        </span>
                      </span>
                      <span className={presentationChrome.masterCardUsage}>
                        {entry.masterName}
                      </span>
                    </button>
                  ))}
                </div>
              )}
              </>
              )}
            </div>
  );
}
