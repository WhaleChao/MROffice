"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The Playback contextual tab of the presentation Ribbon for video and audio
 * (Video Styles, Bookmarks, Editing, Captions, Options), drawn from the window
 * bindings it reads.
 *
 * Extracted verbatim from the render of
 * `components/notebook/contents/PresentationWindow.tsx`: the JSX is unchanged;
 * the window passes the bindings it used to close over as `ctl`, and the tab
 * condition and key stay at the call site. No hooks run in the tab itself, so
 * the window's hook order is unchanged; the Preview button keeps its own.
 */

import { useEffect, useRef, useState } from "react";
import {
  OfficeHomeRibbonGroup,
  OfficeHomeCommandRow,
  OfficeHomeCommand,
  OfficeHomeMenuTrigger,
  OfficeHomeCommandStack,
} from "@mrnote/core/web/components/office-ribbon/OfficeHomeRibbon.tsx";
import presentationChrome from "@mrnote/core/web/lib/presentation/presentation-window-shell.module.css";
import {
  IconMediaPreview,
  IconMediaTrim,
  IconMediaVolume,
  IconMediaStart,
  IconVideoOutline,
  IconMediaCaptions,
  IconAudioNoStyle,
  IconAudioBackground,
} from "@mrnote/core/web/components/notebook/contents/PresentationRibbonIcons.tsx";
import { BookmarkPlus, Bookmark, Trash2 } from "@mrnote/core/web/components/office-ribbon/icons/houseIcons.tsx";
import { presentationCanvasMediaSource } from "./PresentationMediaTrimDialog";
import { PresentationMediaFadeFields } from "./PresentationMediaFadeFields";
import {
  DropdownMenu,
  DropdownMenuTrigger,
  DropdownMenuCheckboxItem,
  DropdownMenuRadioGroup,
  DropdownMenuRadioItem,
  DropdownMenuItem,
} from "@mrnote/core/web/components/ui/dropdown-menu.tsx";
import { RibbonDropdownMenuContent } from "@mrnote/core/web/components/office-ribbon/CupertinoRibbonMenu.tsx";
import { PRESENTATION_MEDIA_VOLUMES, type PersistedPresentationMediaPlaybackSettings, type PresentationMediaPlaybackSettings } from "@mrnote/core/web/lib/presentation-embed.ts";
import { PresentationMediaPlaybackOptions } from "./PresentationMediaPlaybackOptions";
import {
  MAX_PRESENTATION_MEDIA_CAPTION_TRACKS,
  type NativePresentationCaptionAuthority,
  type PresentationMediaCaptionTrack,
  type nativePresentationCaptionAuthority,
} from "@mrnote/core/web/lib/presentation-media-captions.ts";
import type { useTranslations } from "@mrnote/core/host/next-intl.tsx";
import type { Dispatch, RefObject, SetStateAction } from "react";
import type { SlideModel } from "@mrnote/core/web/lib/presentation/deck-model.ts";
import type { PresentationEmbedBlock, PresentationImageBlock } from "@mrnote/core/web/components/console/editor/native-document/documentCommands.ts";
import type { JSX } from "react/jsx-runtime";
import type { ResolvedPresentationFeatureCommand } from "@mrnote/core/web/components/notebook/contents/ribbon/commandRegistry.ts";

/** The window bindings this tab reads, passed through unchanged. */
export interface PresentationMediaPlaybackRibbonTabCtl {
  t: ReturnType<typeof useTranslations<"console-notebooks">>;
  captionMenuTriggerRef: RefObject<HTMLButtonElement | null>;
  captionInsertBusy: boolean;
  setMediaTrimDialog: Dispatch<SetStateAction<{ slideId: string; blockId: string; kind: "video" | "audio"; src: string | null; durationMs: number; trimStartMs: number; trimEndMs: number; } | null>>;
  slideCanvasRef: RefObject<HTMLDivElement | null>;
  activeSlide: SlideModel;
  contextualMediaBlock: PresentationEmbedBlock;
  contextualMediaKind: "video" | "audio" | "iframe" | null;
  contextualMediaPlaybackSettings: PresentationMediaPlaybackSettings | null;
  contextualMediaPlaybackIsWritable: boolean;
  contextualMediaTrimDurationMs: number | undefined;
  contextualMediaTrimIsWritable: boolean;
  contextualMediaCaptionAuthority: NativePresentationCaptionAuthority | null;
  contextualMediaCaptionTracks: readonly PresentationMediaCaptionTrack[];
  contextualMediaCaptionsIsWritable: boolean;
  mediaPlaybackCommandLabel: (label: string) => string;
  contextualMediaBackgroundActive: boolean;
  contextualMediaBackgroundResidue: boolean;
  mediaPlaybackStartLabels: { click: string; automatic: string; sequence: string; };
  mediaPlayInBackgroundLabel: string;
  mediaAudioStylesLabel: string;
  mediaNoAudioStyleLabel: string;
  contextualMediaCanPreview: boolean;
  /** The selected clip's player on the slide, as its control bar finds it. */
  getMediaElement: (blockId: string) => HTMLMediaElement | null;
  commitMediaCaptionTracks: (target: { slideId: string; blockId: string; nativeAuthority: ReturnType<typeof nativePresentationCaptionAuthority>; }, rawTracks: unknown) => boolean;
  updateMediaPlaybackSettings: (blockId: string, patch: PersistedPresentationMediaPlaybackSettings) => void;
  selectedMediaBookmark: { blockId: string; name: string; } | null;
  addMediaBookmark: (blockId: string) => void;
  removeMediaBookmark: (blockId: string) => void;
  beginMediaCaptionInsert: (trigger: HTMLElement | null) => boolean;
  renderVideoShapeMenu: (block: PresentationEmbedBlock) => JSX.Element;
  renderVideoStylesGallery: (block: PresentationEmbedBlock) => JSX.Element;
  renderVideoAdjustGroup: (block: PresentationEmbedBlock) => JSX.Element;
  renderVideoBorderMenu: (block: PresentationEmbedBlock) => JSX.Element;
  renderPictureEffectsMenu: (block: PresentationImageBlock | PresentationEmbedBlock, options: { label: string; testId: string; controlId: string; }) => JSX.Element;
  featureDecision: (featureId: string) => ResolvedPresentationFeatureCommand;
  executePresentationFeature: (featureId: string) => undefined;
  activeRibbonTab: "media-format" | "media-playback";
}

export function PresentationMediaPlaybackRibbonTab({ ctl }: { ctl: PresentationMediaPlaybackRibbonTabCtl }) {
  const {
    t,
    captionMenuTriggerRef,
    captionInsertBusy,
    setMediaTrimDialog,
    slideCanvasRef,
    activeSlide,
    contextualMediaBlock,
    contextualMediaKind,
    contextualMediaPlaybackSettings,
    contextualMediaPlaybackIsWritable,
    contextualMediaTrimDurationMs,
    contextualMediaTrimIsWritable,
    contextualMediaCaptionAuthority,
    contextualMediaCaptionTracks,
    contextualMediaCaptionsIsWritable,
    mediaPlaybackCommandLabel,
    contextualMediaBackgroundActive,
    contextualMediaBackgroundResidue,
    mediaPlaybackStartLabels,
    mediaPlayInBackgroundLabel,
    mediaAudioStylesLabel,
    mediaNoAudioStyleLabel,
    contextualMediaCanPreview,
    getMediaElement,
    commitMediaCaptionTracks,
    updateMediaPlaybackSettings,
    selectedMediaBookmark,
    addMediaBookmark,
    removeMediaBookmark,
    beginMediaCaptionInsert,
    renderVideoShapeMenu,
    renderVideoStylesGallery,
    renderVideoAdjustGroup,
    renderVideoBorderMenu,
    renderPictureEffectsMenu,
    featureDecision,
    executePresentationFeature,
    activeRibbonTab,
  } = ctl;
  return (
                  <>
                    <OfficeHomeRibbonGroup
                      className={`${presentationChrome.mediaPlaybackPreviewGroup} ${
                        contextualMediaKind === "audio"
                          ? presentationChrome.mediaAudioOrderPreview
                          : activeRibbonTab === "media-format"
                            ? presentationChrome.mediaVideoOrderPreview
                            : ""
                      }`}
                      label={t("presentation.media.previewGroup")}
                      size="standard"
                      testId="presentation-ribbon-media-preview-group"
                    >
                      <OfficeHomeCommandRow>
                        <PresentationMediaPreviewCommand
                          blockId={contextualMediaBlock?.id ?? null}
                          canPreview={contextualMediaCanPreview}
                          getMediaElement={getMediaElement}
                          showFromCurrent={
                            featureDecision("slideshow.fromCurrent").enabled
                              ? () => executePresentationFeature("slideshow.fromCurrent")
                              : null
                          }
                          labels={{
                            play: t("presentation.media.preview"),
                            pause: t("presentation.media.controlPause"),
                            unsupported: t("presentation.embedUnsupported"),
                          }}
                        />
                      </OfficeHomeCommandRow>
                    </OfficeHomeRibbonGroup>
                    {activeRibbonTab === "media-playback" ? (
                      /* PowerPoint: Playback ▸ Bookmarks ▸ Add Bookmark /
                         Remove Bookmark (PPT-A157): Add marks the clip's
                         current position, Remove the bookmark selected on
                         the clip's control bar. */
                      <OfficeHomeRibbonGroup
                        className={
                          contextualMediaKind === "audio"
                            ? presentationChrome.mediaAudioOrderBookmarks
                            : undefined
                        }
                        label={t("presentation.media.bookmarksGroup")}
                        size="standard"
                        testId="presentation-ribbon-media-bookmarks-group"
                      >
                        <OfficeHomeCommandRow>
                          <OfficeHomeCommand
                            variant="large"
                            icon={<BookmarkPlus size={24} aria-hidden="true" />}
                            label={t("presentation.media.addBookmark")}
                            disabled={!contextualMediaPlaybackIsWritable}
                            onClick={() => {
                              if (contextualMediaBlock) {
                                addMediaBookmark(contextualMediaBlock.id);
                              }
                            }}
                            testId="presentation-ribbon-media-add-bookmark"
                          />
                          <OfficeHomeCommand
                            variant="large"
                            icon={<Bookmark size={24} aria-hidden="true" />}
                            label={t("presentation.media.removeBookmark")}
                            disabled={
                              !contextualMediaPlaybackIsWritable ||
                              !contextualMediaBlock ||
                              selectedMediaBookmark?.blockId !==
                                contextualMediaBlock.id ||
                              !contextualMediaPlaybackSettings?.bookmarks.some(
                                (bookmark) =>
                                  bookmark.name === selectedMediaBookmark?.name,
                              )
                            }
                            onClick={() => {
                              if (contextualMediaBlock) {
                                removeMediaBookmark(contextualMediaBlock.id);
                              }
                            }}
                            testId="presentation-ribbon-media-remove-bookmark"
                          />
                        </OfficeHomeCommandRow>
                      </OfficeHomeRibbonGroup>
                    ) : null}
                    {activeRibbonTab === "media-playback" ? (
                      /* PowerPoint: Playback ▸ Editing ▸ Trim Audio / Trim Video
                         ("Trim a video, music clip, or sound clip in
                         PowerPoint", Microsoft Support). Same dialog and
                         writability rule as Trim Video; the audio branch of
                         `contextualMediaTrimIsWritable` existed without any
                         command reaching it (PPT-DIALOG-AUDIT §10.5). */
                      <OfficeHomeRibbonGroup
                        className={`${presentationChrome.mediaPlaybackEditingGroup} ${
                          contextualMediaKind === "audio"
                            ? presentationChrome.mediaAudioOrderEditing
                            : ""
                        }`}
                        label={t("presentation.media.editingGroup")}
                        size="standard"
                        testId="presentation-ribbon-media-editing-group"
                      >
                        <OfficeHomeCommandRow>
                          <OfficeHomeCommand
                            variant="large"
                            icon={<IconMediaTrim home size={24} />}
                            label={t(
                              contextualMediaKind === "video"
                                ? "presentation.media.trimVideo"
                                : "presentation.media.trimAudio",
                            )}
                            disabled={!contextualMediaTrimIsWritable}
                            disabledReason={
                              contextualMediaTrimIsWritable
                                ? undefined
                                : t("presentation.media.preciseTimingUnavailable")
                            }
                            onClick={() => {
                              if (
                                !contextualMediaBlock ||
                                !contextualMediaPlaybackSettings ||
                                !contextualMediaTrimDurationMs ||
                                !contextualMediaTrimIsWritable
                              )
                                return;
                              setMediaTrimDialog({
                                slideId: activeSlide.id,
                                blockId: contextualMediaBlock.id,
                                kind:
                                  contextualMediaKind === "video"
                                    ? "video"
                                    : "audio",
                                // The canvas player already holds the clip
                                // (an asset ticket or the inline data), so the
                                // dialog previews that same source.
                                src: presentationCanvasMediaSource(
                                  slideCanvasRef.current,
                                  contextualMediaBlock.id,
                                ),
                                durationMs: contextualMediaTrimDurationMs,
                                trimStartMs:
                                  contextualMediaPlaybackSettings.trimStartMs,
                                trimEndMs:
                                  contextualMediaPlaybackSettings.trimEndMs ??
                                  contextualMediaTrimDurationMs,
                              });
                            }}
                            testId="presentation-ribbon-media-trim"
                          />
                          <PresentationMediaFadeFields
                            fadeInMs={contextualMediaPlaybackSettings?.fadeInMs ?? 0}
                            fadeOutMs={contextualMediaPlaybackSettings?.fadeOutMs ?? 0}
                            disabled={!contextualMediaPlaybackIsWritable}
                            labels={{
                              group: t("presentation.media.fade"),
                              fadeIn: t("presentation.media.fadeIn"),
                              fadeOut: t("presentation.media.fadeOut"),
                            }}
                            onCommit={(patch) => {
                              if (!contextualMediaBlock) return;
                              updateMediaPlaybackSettings(
                                contextualMediaBlock.id,
                                patch,
                              );
                            }}
                          />
                        </OfficeHomeCommandRow>
                      </OfficeHomeRibbonGroup>
                    ) : null}
                    {activeRibbonTab === "media-playback" ? <OfficeHomeRibbonGroup
                      className={`${presentationChrome.mediaPlaybackOptionsGroup} ${
                        contextualMediaKind === "audio"
                          ? presentationChrome.mediaAudioOrderOptions
                          : ""
                      }`}
                      label={t("presentation.media.optionsGroup")}
                      size="standard"
                      testId="presentation-ribbon-media-options-group"
                    >
                      <OfficeHomeCommandRow
                        className={presentationChrome.mediaPlaybackOptionsRow}
                        testId="presentation-ribbon-media-options-row"
                      >
                        <DropdownMenu>
                          <DropdownMenuTrigger asChild>
                            <OfficeHomeMenuTrigger
                              variant="large"
                              displayLabel={t("presentation.media.volume.label")}
                              hasMenuIndicator
                              icon={<IconMediaVolume home size={24} />}
                              label={mediaPlaybackCommandLabel(
                                t("presentation.media.volume.label"),
                              )}
                              disabled={!contextualMediaPlaybackIsWritable}
                              controlId="PPTCTL.CONTEXT.MEDIA.PLAYBACK.VOLUME"
                              testId="presentation-ribbon-media-volume"
                            />
                          </DropdownMenuTrigger>
                          <RibbonDropdownMenuContent
                            align="start"
                            data-testid="presentation-ribbon-media-volume-menu"
                          >
                            {PRESENTATION_MEDIA_VOLUMES.map((volume) => (
                              <DropdownMenuCheckboxItem
                                key={volume}
                                checked={
                                  contextualMediaPlaybackSettings?.volume ===
                                  volume
                                }
                                onCheckedChange={(checked) => {
                                  if (
                                    checked !== true ||
                                    !contextualMediaBlock
                                  ) {
                                    return;
                                  }
                                  updateMediaPlaybackSettings(
                                    contextualMediaBlock.id,
                                    { volume },
                                  );
                                }}
                                data-testid={`presentation-ribbon-media-volume-${volume}`}
                              >
                                {t(
                                  `presentation.media.volume.${volume}` as "presentation.media.volume.high",
                                )}
                              </DropdownMenuCheckboxItem>
                            ))}
                          </RibbonDropdownMenuContent>
                        </DropdownMenu>
                        <DropdownMenu>
                          <DropdownMenuTrigger asChild>
                            <OfficeHomeMenuTrigger
                              variant="compact"
                              displayLabel={
                                mediaPlaybackStartLabels[
                                  contextualMediaPlaybackSettings?.start ??
                                    "click"
                                ]
                              }
                              hasMenuIndicator
                              icon={<IconMediaStart home size={24} />}
                              label={mediaPlaybackCommandLabel(
                                `${t("presentation.media.start")}: ${
                                  mediaPlaybackStartLabels[
                                    contextualMediaPlaybackSettings?.start ??
                                      "click"
                                  ]
                                }`,
                              )}
                              disabled={!contextualMediaPlaybackIsWritable}
                              controlId="PPTCTL.CONTEXT.MEDIA.PLAYBACK.START"
                              testId="presentation-ribbon-media-start"
                            />
                          </DropdownMenuTrigger>
                          <RibbonDropdownMenuContent
                            align="start"
                            data-testid="presentation-ribbon-media-start-menu"
                          >
                            <DropdownMenuRadioGroup
                              value={
                                contextualMediaPlaybackSettings?.start ??
                                "click"
                              }
                              onValueChange={(value) => {
                                if (
                                  !contextualMediaBlock ||
                                  (value !== "click" &&
                                    value !== "automatic")
                                ) {
                                  return;
                                }
                                updateMediaPlaybackSettings(
                                  contextualMediaBlock.id,
                                  { start: value },
                                );
                              }}
                            >
                              {(["click", "automatic"] as const).map(
                                (start) => (
                                  <DropdownMenuRadioItem
                                    key={start}
                                    value={start}
                                    data-testid={`presentation-ribbon-media-start-${start}`}
                                  >
                                    {mediaPlaybackStartLabels[start]}
                                  </DropdownMenuRadioItem>
                                ),
                              )}
                              {/* PowerPoint for the web starts audio Automatically or
                                  When Clicked On only; "In Click Sequence" is a desktop
                                  value, so it appears here only when the file already
                                  carries it — the current choice stays visible and stays
                                  preserved, and nobody is offered a value that cannot be set. */}
                              {contextualMediaPlaybackSettings?.start === "sequence" ? (
                                <DropdownMenuRadioItem
                                  value="sequence"
                                  disabled
                                  title={t(
                                    "presentation.media.startSequenceUnavailable",
                                  )}
                                  data-testid="presentation-ribbon-media-start-sequence"
                                >
                                  {mediaPlaybackStartLabels.sequence}
                                </DropdownMenuRadioItem>
                              ) : null}
                            </DropdownMenuRadioGroup>
                          </RibbonDropdownMenuContent>
                        </DropdownMenu>
                          {contextualMediaKind === "video" ||
                          contextualMediaKind === "audio" ? (
                            <PresentationMediaPlaybackOptions
                              kind={contextualMediaKind}
                              settings={contextualMediaPlaybackSettings}
                              disabled={!contextualMediaPlaybackIsWritable}
                              describe={mediaPlaybackCommandLabel}
                              labels={{
                                playFullScreen: t("presentation.media.playFullScreen"),
                                playAcrossSlides: t("presentation.media.playAcrossSlides"),
                                hideDuringShow: t(
                                  contextualMediaKind === "video"
                                    ? "presentation.media.hideWhileNotPlaying"
                                    : "presentation.media.hideDuringShow",
                                ),
                                loop: t("presentation.media.loop"),
                                rewindAfterPlaying: t("presentation.media.rewind"),
                              }}
                              onPatch={(patch) => {
                                if (!contextualMediaBlock) return;
                                updateMediaPlaybackSettings(
                                  contextualMediaBlock.id,
                                  patch,
                                );
                              }}
                            />
                          ) : null}
                      </OfficeHomeCommandRow>
                    </OfficeHomeRibbonGroup> : <OfficeHomeRibbonGroup className={`${presentationChrome.mediaPlaybackVideoStylesGroup} ${presentationChrome.mediaVideoOrderStyles}`} label={t("presentation.media.videoStylesGroup")} size="standard" testId="presentation-ribbon-media-video-styles-group">
                      {/* Video Styles as PowerPoint lays it out: the gallery,
                          then Video Shape / Video Border / Video Effects in
                          one column of compact commands. Placed as four
                          siblings the last ones wrapped under the gallery,
                          half outside the ribbon. */}
                      <OfficeHomeCommandRow className={presentationChrome.galleryWithCommandStack}>
                        {contextualMediaBlock.kind === "video" ? renderVideoStylesGallery(contextualMediaBlock) : null}
                        <OfficeHomeCommandStack className={presentationChrome.galleryCommandStack} rows={3}>
                          {contextualMediaBlock.kind === "video" ? (
                            <>
                              <OfficeHomeCommandRow>{renderVideoShapeMenu(contextualMediaBlock)}</OfficeHomeCommandRow>
                              <OfficeHomeCommandRow>{renderVideoBorderMenu(contextualMediaBlock)}</OfficeHomeCommandRow>
                              <OfficeHomeCommandRow>
                                {renderPictureEffectsMenu(contextualMediaBlock, { label: t("presentation.media.videoEffects"), testId: "presentation-ribbon-media-video-effects", controlId: "PPTCTL.CONTEXT.VIDEO.STYLE.EFFECTS" })}
                              </OfficeHomeCommandRow>
                            </>
                          ) : (
                            <OfficeHomeCommandRow>
                              <OfficeHomeCommand variant="compact" icon={<IconVideoOutline home size={16} />} label={t("presentation.media.videoOutline")} disabled disabledReason={t("presentation.media.videoOutlineUnavailable")} onClick={() => undefined} testId="presentation-ribbon-media-video-outline" />
                            </OfficeHomeCommandRow>
                          )}
                        </OfficeHomeCommandStack>
                      </OfficeHomeCommandRow>
                    </OfficeHomeRibbonGroup>}
                    {activeRibbonTab === "media-format" && contextualMediaKind === "video" && contextualMediaBlock.kind === "video"
                      ? renderVideoAdjustGroup(contextualMediaBlock)
                      : null}
                    {activeRibbonTab === "media-playback" &&
                    contextualMediaKind === "video" ? (
                      <OfficeHomeRibbonGroup
                        className={presentationChrome.mediaPlaybackEditingGroup}
                        label={t("presentation.media.captionsGroup")}
                        size="standard"
                        testId="presentation-ribbon-media-captions-group"
                      >
                        <OfficeHomeCommandRow>
                        <DropdownMenu>
                          <DropdownMenuTrigger asChild>
                            <OfficeHomeMenuTrigger
                              ref={captionMenuTriggerRef}
                              variant="large"
                              displayLabel={t("presentation.media.captionsEdit")}
                              hasMenuIndicator
                              icon={<IconMediaCaptions home size={24} />}
                              label={
                                contextualMediaCaptionsIsWritable
                                  ? `${t("presentation.media.captions")} (${contextualMediaCaptionTracks.length})`
                                  : `${t("presentation.media.captions")} — ${t(
                                      "presentation.media.captionsAssetUnavailable",
                                    )}`
                              }
                              disabled={!contextualMediaCaptionsIsWritable}
                              aria-busy={captionInsertBusy}
                              controlId="PPTCTL.CONTEXT.MEDIA.PLAYBACK.CAPTIONS"
                              testId="presentation-ribbon-media-captions"
                            />
                          </DropdownMenuTrigger>
                          <RibbonDropdownMenuContent
                            align="start"
                            data-testid="presentation-ribbon-media-captions-menu"
                          >
                            <DropdownMenuItem
                              disabled={
                                captionInsertBusy ||
                                contextualMediaCaptionTracks.length >=
                                  MAX_PRESENTATION_MEDIA_CAPTION_TRACKS
                              }
                              onSelect={() =>
                                beginMediaCaptionInsert(
                                  captionMenuTriggerRef.current,
                                )
                              }
                              data-testid="presentation-ribbon-media-captions-insert"
                            >
                              <IconMediaCaptions home size={14} />
                              {t("presentation.media.captionsInsert")}
                            </DropdownMenuItem>
                            <DropdownMenuItem
                              disabled={
                                contextualMediaCaptionTracks.length === 0
                              }
                              onSelect={() => {
                                if (!contextualMediaBlock) return;
                                commitMediaCaptionTracks(
                                  {
                                    slideId: activeSlide.id,
                                    blockId: contextualMediaBlock.id,
                                    nativeAuthority:
                                      contextualMediaCaptionAuthority,
                                  },
                                  [],
                                );
                              }}
                              data-testid="presentation-ribbon-media-captions-remove-all"
                            >
                              <Trash2 size={14} />
                              {t("presentation.media.captionsRemoveAll")}
                            </DropdownMenuItem>
                          </RibbonDropdownMenuContent>
                        </DropdownMenu>
                        </OfficeHomeCommandRow>
                      </OfficeHomeRibbonGroup>
                    ) : null}
                    {activeRibbonTab === "media-playback" &&
                    contextualMediaKind === "audio" ? (
                      <OfficeHomeRibbonGroup
                        className={
                          `${presentationChrome.mediaPlaybackAudioStylesGroup} ${presentationChrome.mediaAudioOrderAudioStyles}`
                        }
                        label={mediaAudioStylesLabel}
                        size="standard"
                        testId="presentation-ribbon-media-audio-styles-group"
                      >
                        <OfficeHomeCommandRow
                          className={
                            presentationChrome.mediaPlaybackAudioStylesRow
                          }
                          testId="presentation-ribbon-media-audio-styles-row"
                        >
                          <OfficeHomeCommand
                            variant="large"
                            active={!contextualMediaBackgroundResidue}
                            icon={<IconAudioNoStyle home size={24} />}
                            displayLabel={mediaNoAudioStyleLabel}
                            label={mediaPlaybackCommandLabel(
                              mediaNoAudioStyleLabel,
                            )}
                            disabled={!contextualMediaPlaybackIsWritable}
                            onClick={() => {
                              if (
                                !contextualMediaBlock ||
                                !contextualMediaBackgroundResidue
                              ) {
                                return;
                              }
                              updateMediaPlaybackSettings(
                                contextualMediaBlock.id,
                                {
                                  start: "click",
                                  playAcrossSlides: false,
                                  loop: false,
                                  hideDuringShow: false,
                                },
                              );
                            }}
                            testId="presentation-ribbon-media-no-audio-style"
                          />
                          <OfficeHomeCommand
                            variant="large"
                            active={contextualMediaBackgroundActive}
                            icon={<IconAudioBackground home size={24} />}
                            displayLabel={mediaPlayInBackgroundLabel}
                            label={mediaPlaybackCommandLabel(
                              mediaPlayInBackgroundLabel,
                            )}
                            disabled={!contextualMediaPlaybackIsWritable}
                            onClick={() => {
                              if (
                                !contextualMediaBlock ||
                                contextualMediaBackgroundActive
                              ) {
                                return;
                              }
                              updateMediaPlaybackSettings(
                                contextualMediaBlock.id,
                                {
                                  // PowerPoint's Play in Background: start
                                  // automatically, play across slides, loop
                                  // until stopped, hide during the show.
                                  start: "automatic",
                                  playAcrossSlides: true,
                                  loop: true,
                                  hideDuringShow: true,
                                },
                              );
                            }}
                            testId="presentation-ribbon-media-background"
                          />
                        </OfficeHomeCommandRow>
                      </OfficeHomeRibbonGroup>
                    ) : null}
                  </>
  );
}








function PresentationMediaPreviewCommand({
  blockId,
  canPreview,
  getMediaElement,
  showFromCurrent,
  labels,
}: {
  blockId: string | null;
  canPreview: boolean;
  getMediaElement: (blockId: string) => HTMLMediaElement | null;
  showFromCurrent: (() => void) | null;
  labels: { play: string; pause: string; unsupported: string };
}) {
  const [playing, setPlaying] = useState(false);
  const watchedRef = useRef<{ media: HTMLMediaElement; sync: () => void } | null>(null);
  const unwatch = () => {
    const watched = watchedRef.current;
    if (!watched) return;
    for (const name of ["play", "pause", "ended", "emptied"]) {
      watched.media.removeEventListener(name, watched.sync);
    }
    watchedRef.current = null;
  };
  const watch = (media: HTMLMediaElement) => {
    if (watchedRef.current?.media === media) return;
    unwatch();
    const sync = () => setPlaying(!media.paused && !media.ended);
    for (const name of ["play", "pause", "ended", "emptied"]) {
      media.addEventListener(name, sync);
    }
    watchedRef.current = { media, sync };
    sync();
  };
  useEffect(() => {
    const media = blockId ? getMediaElement(blockId) : null;
    if (media) watch(media);
    else setPlaying(false);
    return unwatch;
    // The player is looked up once per clip; `getMediaElement` is stable.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [blockId]);
  const label = playing ? labels.pause : labels.play;
  return (
    <OfficeHomeCommand
      variant="large"
      icon={<IconMediaPreview home size={24} />}
      displayLabel={label}
      label={canPreview ? label : `${label} — ${labels.unsupported}`}
      active={playing}
      disabled={!canPreview || !blockId}
      onClick={() => {
        if (!blockId) return;
        const media = getMediaElement(blockId);
        if (!media) {
          showFromCurrent?.();
          return;
        }
        watch(media);
        if (media.paused || media.ended) {
          void Promise.resolve(media.play?.()).catch(() => {});
        } else {
          media.pause?.();
        }
      }}
      testId="presentation-ribbon-media-preview"
    />
  );
}
