"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { useId, useMemo, useRef, useState, type RefObject } from "react";

import { OfficeDialogShell } from "@mrnote/core/web/components/office-ribbon/OfficeDialogShell.tsx";
import { officeDialogParts } from "@mrnote/core/web/components/office-ribbon/officeDialogChrome.tsx";
import {
  buildSlideLinkHref,
  normalizePresentationClickAction,
  parseSlideLinkTarget,
  sanitizeBlockHref,
  type PresentationActionButtonShapeKind,
  type PresentationMediaClickActionKind,
  type PresentationRelativeClickActionTarget,
} from "@mrnote/core/web/lib/presentation-block-fx.ts";
import {
  NO_PRESENTATION_ACTION_SETTINGS,
  type PresentationActionTarget,
  type PresentationHoverTarget,
  type PresentationObjectActionSettings,
} from "../presentation-action-settings";

/** The click target an inserted Action Button is created with. */
export type PresentationActionSettingsValue = PresentationActionTarget;

export interface PresentationActionSettingsSlide {
  readonly id: string;
  readonly label: string;
}

export interface PresentationActionSettingsMedia {
  readonly id: string;
  readonly kind: PresentationMediaClickActionKind;
  readonly label: string;
}

interface PresentationActionSettingsDialogProps {
  readonly open: boolean;
  /**
   * The Action Button being inserted, or null for Action Settings… on an
   * object already on the slide (Insert > Action, right-click > Action
   * Settings…; PPT-C003, PPT-B012): OK applies rather than inserts.
   */
  readonly shape: PresentationActionButtonShapeKind | null;
  /** The object's Action Settings now, for an existing object. */
  readonly initial?: PresentationObjectActionSettings;
  readonly slides: readonly PresentationActionSettingsSlide[];
  readonly media: readonly PresentationActionSettingsMedia[];
  readonly returnFocusRef: RefObject<HTMLElement | null>;
  readonly t: (
    key: string,
    values?: Record<string, string | number>,
  ) => string;
  readonly onApply: (value: PresentationObjectActionSettings) => boolean;
  readonly onClose: () => void;
}

type ActionPage = "click" | "hover";
type ActionKind = "none" | "hyperlink" | "media";
/** PowerPoint's Hyperlink to: list, in its order. */
const HYPERLINK_CHOICES = [
  "next",
  "previous",
  "first",
  "last",
  "last-viewed",
  "end-show",
  "custom-show",
  "slide",
  "url",
  "other-presentation",
  "other-file",
] as const;
type HyperlinkChoice = (typeof HYPERLINK_CHOICES)[number];
/** Choices a browser editor cannot carry out: listed, as PowerPoint lists
 * them, and disabled. */
const UNAVAILABLE_HYPERLINK_CHOICES: ReadonlySet<HyperlinkChoice> = new Set([
  "custom-show",
  "other-presentation",
  "other-file",
]);
const HYPERLINK_CHOICE_LABEL: Record<HyperlinkChoice, string> = {
  next: "presentation.actionSettings.relative.next",
  previous: "presentation.actionSettings.relative.previous",
  first: "presentation.actionSettings.relative.first",
  last: "presentation.actionSettings.relative.last",
  "last-viewed": "presentation.actionSettings.relative.last-viewed",
  "end-show": "presentation.actionSettings.relative.end-show",
  "custom-show": "presentation.actionSettings.target.customShow",
  slide: "presentation.actionSettings.target.slide",
  url: "presentation.actionSettings.target.url",
  "other-presentation": "presentation.actionSettings.target.otherPresentation",
  "other-file": "presentation.actionSettings.target.otherFile",
};

interface PageDraft {
  readonly kind: ActionKind;
  readonly choice: HyperlinkChoice;
  readonly slideId: string;
  readonly url: string;
  readonly mediaId: string;
  readonly highlight: boolean;
}

function requiredMediaKind(
  shape: PresentationActionButtonShapeKind | null,
): PresentationMediaClickActionKind | null {
  if (shape === "action-button-movie") return "video";
  if (shape === "action-button-sound") return "audio";
  return null;
}

function pageDraft(
  target: PresentationActionTarget | PresentationHoverTarget | null,
  highlight: boolean,
  firstSlideId: string,
  defaultKind: ActionKind,
): PageDraft {
  const base: PageDraft = {
    kind: defaultKind,
    choice: "next",
    slideId: firstSlideId,
    url: "",
    mediaId: "",
    highlight,
  };
  if (!target) return base;
  if ("href" in target) {
    const slideId = parseSlideLinkTarget(target.href);
    return slideId
      ? { ...base, kind: "hyperlink", choice: "slide", slideId }
      : { ...base, kind: "hyperlink", choice: "url", url: target.href };
  }
  const action = "action" in target ? target.action : target.clickAction;
  if (action.kind === "play-media") {
    return { ...base, kind: "media", mediaId: action.targetBlockId };
  }
  return { ...base, kind: "hyperlink", choice: action.target };
}

function draftTarget(
  draft: PageDraft,
  slides: readonly PresentationActionSettingsSlide[],
  media: readonly PresentationActionSettingsMedia[],
): PresentationActionTarget | null | undefined {
  if (draft.kind === "none") return null;
  if (draft.kind === "media") {
    const selected = media.find((item) => item.id === draft.mediaId);
    const clickAction = selected
      ? normalizePresentationClickAction({
          kind: "play-media",
          mediaKind: selected.kind,
          targetBlockId: selected.id,
        })
      : null;
    return clickAction ? { clickAction } : undefined;
  }
  if (draft.choice === "slide") {
    const href = buildSlideLinkHref(draft.slideId);
    return href && slides.some((slide) => slide.id === draft.slideId)
      ? { href }
      : undefined;
  }
  if (draft.choice === "url") {
    const href = sanitizeBlockHref(draft.url.trim());
    return href && !href.startsWith("slide://") ? { href } : undefined;
  }
  if (UNAVAILABLE_HYPERLINK_CHOICES.has(draft.choice)) return undefined;
  const clickAction = normalizePresentationClickAction({
    kind: "relative-slide",
    target: draft.choice as PresentationRelativeClickActionTarget,
  });
  return clickAction ? { clickAction } : undefined;
}

/**
 * PowerPoint's Action Settings: the Mouse Click and Mouse Over pages, each
 * None, Hyperlink to:, Run program:, Run macro:, Object action: and Play
 * sound: with Highlight click / Highlight when mouse over (Mac 16.86's
 * strings). A browser runs no program, macro or object verb and this editor
 * adds no sound yet: those rows are shown disabled with the reason.
 */
export function PresentationActionSettingsDialog({
  open,
  shape,
  initial = NO_PRESENTATION_ACTION_SETTINGS,
  slides,
  media,
  returnFocusRef,
  t,
  onApply,
  onClose,
}: PresentationActionSettingsDialogProps): React.JSX.Element {
  const id = useId();
  const mediaKind = requiredMediaKind(shape);
  const firstSlideId = slides[0]?.id ?? "";
  const [page, setPage] = useState<ActionPage>("click");
  const [drafts, setDrafts] = useState<Record<ActionPage, PageDraft>>(() => ({
    click: pageDraft(
      initial.click,
      initial.clickHighlight,
      firstSlideId,
      mediaKind ? "media" : shape ? "hyperlink" : "none",
    ),
    hover: pageDraft(initial.hover, initial.hoverHighlight, firstSlideId, "none"),
  }));
  const [commandError, setCommandError] = useState(false);
  const urlInputRef = useRef<HTMLInputElement | null>(null);
  const availableMedia = useMemo(
    () => (mediaKind ? media.filter((item) => item.kind === mediaKind) : []),
    [media, mediaKind],
  );
  const draft = drafts[page];
  const update = (patch: Partial<PageDraft>) => {
    setCommandError(false);
    setDrafts((current) => ({ ...current, [page]: { ...current[page], ...patch } }));
  };

  const value = useMemo<PresentationObjectActionSettings | null>(() => {
    const click = draftTarget(drafts.click, slides, availableMedia);
    const hoverTarget = draftTarget(drafts.hover, slides, []);
    if (click === undefined || hoverTarget === undefined) return null;
    // An Action Button is inserted with what it does on a click.
    if (shape && !click) return null;
    const hover: PresentationHoverTarget | null = !hoverTarget
      ? null
      : "href" in hoverTarget
        ? { href: hoverTarget.href }
        : hoverTarget.clickAction.kind === "relative-slide"
          ? { action: hoverTarget.clickAction }
          : null;
    return {
      click,
      clickHighlight: click !== null && drafts.click.highlight,
      hover,
      hoverHighlight: hover !== null && drafts.hover.highlight,
    };
  }, [availableMedia, drafts, shape, slides]);

  const submit = () => {
    if (!value) return;
    if (!onApply(value)) {
      setCommandError(true);
      return;
    }
    onClose();
  };

  const pagePrefix = `${id}-${page}`;
  const testPrefix = `presentation-action-settings-${page}`;
  const radio = (kind: ActionKind, label: string, disabled = false) => (
    <label className={officeDialogParts.check}>
      <input
        type="radio"
        name={`${pagePrefix}-kind`}
        checked={draft.kind === kind}
        disabled={disabled}
        onChange={() => {
          update({ kind });
          if (kind === "hyperlink" && draft.choice === "url") {
            requestAnimationFrame(() => urlInputRef.current?.focus());
          }
        }}
        data-testid={`${testPrefix}-${kind}`}
      />
      <span>{t(label)}</span>
    </label>
  );
  const unavailable = (key: string, testId: string) => (
    <label className={officeDialogParts.check} title={t("presentation.actionSettings.programUnavailable")}>
      <input type="radio" name={`${pagePrefix}-kind`} checked={false} disabled data-testid={`${testPrefix}-${testId}`} onChange={() => undefined} />
      <span>{t(key)}</span>
    </label>
  );

  return (
    <OfficeDialogShell
      open={open}
      title={t("presentation.actionSettings.title")}
      description={shape ? t(`presentation.shape.${shape}`) : undefined}
      defaultActionLabel={t("presentation.actionSettings.ok")}
      cancelActionLabel={t("common.cancel")}
      onDefaultAction={submit}
      onCancel={onClose}
      defaultActionDisabled={!value}
      initialFocusRef={draft.kind === "hyperlink" && draft.choice === "url" ? urlInputRef : undefined}
      returnFocusRef={returnFocusRef}
      testIdPrefix="presentation-action-settings"
      className="sm:max-w-[520px]"
    >
      <div className="grid gap-4">
        <div
          className={officeDialogParts.tabs}
          role="tablist"
          aria-label={t("presentation.actionSettings.title")}
        >
          {(["click", "hover"] as const).map((item) => (
            <button
              key={item}
              type="button"
              role="tab"
              id={`${id}-${item}-tab`}
              aria-controls={`${id}-${item}-panel`}
              aria-selected={page === item}
              tabIndex={page === item ? 0 : -1}
              className={officeDialogParts.tab}
              data-testid={`presentation-action-settings-tab-${item}`}
              onClick={() => setPage(item)}
            >
              {t(`presentation.actionSettings.tab.${item}`)}
            </button>
          ))}
        </div>
        <section
          role="tabpanel"
          id={`${id}-${page}-panel`}
          aria-labelledby={`${id}-${page}-tab`}
          className={officeDialogParts.group}
          data-testid={`presentation-action-settings-${page}-panel`}
        >
          <span className={officeDialogParts.sectionTitle}>
            {t(`presentation.actionSettings.group.${page}`)}
          </span>
          {radio("none", "presentation.actionSettings.none", Boolean(shape) && page === "click")}
          {radio("hyperlink", "presentation.actionSettings.hyperlinkTo", page === "click" && mediaKind !== null)}
          <div className="grid gap-2 pl-6">
            <select
              value={draft.choice}
              disabled={draft.kind !== "hyperlink"}
              aria-label={t("presentation.actionSettings.hyperlinkTo")}
              data-testid={`${testPrefix}-target`}
              onChange={(event) => {
                const choice = event.target.value as HyperlinkChoice;
                update({ choice });
                if (choice === "url") requestAnimationFrame(() => urlInputRef.current?.focus());
              }}
            >
              {HYPERLINK_CHOICES.map((choice) => (
                <option
                  key={choice}
                  value={choice}
                  disabled={UNAVAILABLE_HYPERLINK_CHOICES.has(choice)}
                >
                  {t(HYPERLINK_CHOICE_LABEL[choice])}
                </option>
              ))}
            </select>
            {draft.kind === "hyperlink" && draft.choice === "slide" ? (
              <select
                value={draft.slideId}
                aria-label={t("presentation.actionSettings.target.slide")}
                data-testid={`${testPrefix}-slide`}
                onChange={(event) => update({ slideId: event.target.value })}
              >
                {slides.map((slide) => (
                  <option key={slide.id} value={slide.id}>
                    {slide.label}
                  </option>
                ))}
              </select>
            ) : null}
            {draft.kind === "hyperlink" && draft.choice === "url" ? (
              <input
                ref={urlInputRef}
                type="url"
                inputMode="url"
                autoComplete="url"
                value={draft.url}
                placeholder="https://"
                aria-label={t("presentation.actionSettings.target.url")}
                data-testid={`${testPrefix}-url`}
                onChange={(event) => update({ url: event.target.value })}
              />
            ) : null}
          </div>
          {unavailable("presentation.actionSettings.runProgram", "program")}
          {unavailable("presentation.actionSettings.runMacro", "macro")}
          {page === "click" && mediaKind ? (
            <>
              {radio("media", `presentation.actionSettings.media.${mediaKind}`)}
              <div className="grid gap-1 pl-6">
                <select
                  value={draft.mediaId}
                  disabled={draft.kind !== "media"}
                  data-testid="presentation-action-settings-media"
                  onChange={(event) => update({ mediaId: event.target.value })}
                >
                  <option value="">{t("presentation.actionSettings.media.choose")}</option>
                  {availableMedia.map((item) => (
                    <option key={item.id} value={item.id}>
                      {item.label}
                    </option>
                  ))}
                </select>
                {availableMedia.length === 0 ? (
                  <span className={officeDialogParts.hint}>
                    {t("presentation.actionSettings.media.empty")}
                  </span>
                ) : null}
              </div>
            </>
          ) : (
            unavailable("presentation.actionSettings.objectAction", "object")
          )}
          <span className={officeDialogParts.hint}>
            {t("presentation.actionSettings.programUnavailable")}
          </span>
          <label className={officeDialogParts.check} title={t("presentation.actionSettings.soundUnavailable")}>
            <input type="checkbox" checked={false} disabled data-testid={`${testPrefix}-sound`} onChange={() => undefined} />
            <span>{t("presentation.actionSettings.playSound")}</span>
          </label>
          <div className="grid gap-1 pl-6">
            <select disabled aria-label={t("presentation.actionSettings.playSound")} value="none" onChange={() => undefined}>
              <option value="none">{t("presentation.actionSettings.noSound")}</option>
            </select>
            <span className={officeDialogParts.hint}>
              {t("presentation.actionSettings.soundUnavailable")}
            </span>
          </div>
          <label className={officeDialogParts.check}>
            <input
              type="checkbox"
              checked={draft.kind !== "none" && draft.highlight}
              disabled={draft.kind === "none"}
              onChange={(event) => update({ highlight: event.target.checked })}
              data-testid={`${testPrefix}-highlight`}
            />
            <span>
              {t(
                page === "click"
                  ? "presentation.actionSettings.highlightClick"
                  : "presentation.actionSettings.highlightHover",
              )}
            </span>
          </label>
        </section>
        {commandError ? (
          <p className={officeDialogParts.message} data-tone="error" role="alert">
            {t("presentation.actionSettings.failed")}
          </p>
        ) : null}
      </div>
    </OfficeDialogShell>
  );
}
