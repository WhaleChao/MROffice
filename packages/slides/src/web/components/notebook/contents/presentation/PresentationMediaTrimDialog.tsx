"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { ChevronLeft, ChevronRight, Pause, Play } from "lucide-react";
import {
  useCallback,
  useEffect,
  useRef,
  useState,
  type ReactElement,
} from "react";

import { OfficeDialogShell } from "@mrnote/core/web/components/office-ribbon/OfficeDialogShell.tsx";
import { officeDialogParts as dialogParts } from "@mrnote/core/web/components/office-ribbon/officeDialogChrome.tsx";

/**
 * Playback ▸ Trim Video / Trim Audio (PPT-B021).
 *
 * Microsoft Support, "Trim a video, music clip, or sound clip in PowerPoint":
 * the dialog plays the clip, has a start marker (left) and an end marker
 * (right) that are dragged to the new start and end, Previous/Next Frame
 * (Mac: Nudge Backward/Forward) to fine-tune the marker being moved, Start
 * Time / End Time fields and the trimmed Duration (`00:04.445`). Trimming is
 * not destructive: the result is `p14:trim@st|@end` in milliseconds, which the
 * playback model already reads, saves and writes (`trimStartMs` /
 * `trimEndMs`). Each marker and the play position is a range input, so the
 * timeline is keyboard operable as well as draggable.
 */

/** One nudge. PowerPoint moves by one frame; a browser does not expose the
 * frame rate of a clip, so a tenth of a second is used for both kinds. */
export const PRESENTATION_MEDIA_TRIM_NUDGE_MS = 100;

/** `mm:ss.fff`, the form PowerPoint shows in Start Time / End Time / Duration. */
export function formatPresentationMediaTrimTime(ms: number): string {
  const safe = Number.isFinite(ms) && ms > 0 ? Math.round(ms) : 0;
  const minutes = Math.floor(safe / 60_000);
  const seconds = Math.floor((safe % 60_000) / 1_000);
  const millis = safe % 1_000;
  return `${String(minutes).padStart(2, "0")}:${String(seconds).padStart(2, "0")}.${String(millis).padStart(3, "0")}`;
}

/**
 * Read a typed time: `mm:ss.fff`, `m:ss`, or plain seconds (`7.5`). A comma
 * decimal separator is accepted. Anything else is null, never a guess.
 */
export function parsePresentationMediaTrimTime(raw: string): number | null {
  const text = raw.trim().replace(",", ".");
  const match = /^(?:(\d+):)?(\d+(?:\.\d*)?|\.\d+)$/u.exec(text);
  if (!match) return null;
  const minutes = match[1] === undefined ? 0 : Number(match[1]);
  const seconds = Number(match[2]);
  if (match[1] !== undefined && seconds >= 60) return null;
  const ms = Math.round((minutes * 60 + seconds) * 1_000);
  return Number.isSafeInteger(ms) ? ms : null;
}

/**
 * The source the editor canvas has already loaded for one media block, so
 * the dialog previews the same clip without fetching a second ticket. Null
 * when the block has no player yet (jsdom, an unresolved asset).
 */
export function presentationCanvasMediaSource(
  root: ParentNode | null,
  blockId: string,
): string | null {
  if (!root) return null;
  // jsdom has no CSS.escape; a block id only needs its quotes escaped here.
  const id =
    typeof CSS !== "undefined" && typeof CSS.escape === "function"
      ? CSS.escape(blockId)
      : blockId.replace(/["\\]/gu, "\\$&");
  const media = root.querySelector<HTMLMediaElement>(
    `[data-block-id="${id}"] video, [data-block-id="${id}"] audio, ` +
      `[data-presentation-block-id="${id}"] video, [data-presentation-block-id="${id}"] audio`,
  );
  const src = media?.currentSrc || media?.getAttribute("src") || "";
  return /^(?:https?:|blob:|data:(?:audio|video)\/)/iu.test(src) ? src : null;
}

type TrimHandle = "start" | "end" | "position";

export function PresentationMediaTrimDialog({
  kind,
  src,
  durationMs,
  trimStartMs,
  trimEndMs,
  onApply,
  onClose,
  t,
}: {
  kind: "video" | "audio";
  /** The clip the canvas is already playing, when it has one loaded. */
  src: string | null;
  durationMs: number;
  trimStartMs: number;
  trimEndMs: number;
  onApply: (startMs: number, endMs: number) => void;
  onClose: () => void;
  t: (key: string, values?: Record<string, string | number>) => string;
}): ReactElement {
  const [startDraft, setStartDraft] = useState(() =>
    formatPresentationMediaTrimTime(trimStartMs),
  );
  const [endDraft, setEndDraft] = useState(() =>
    formatPresentationMediaTrimTime(trimEndMs),
  );
  const [positionMs, setPositionMs] = useState(trimStartMs);
  const [playing, setPlaying] = useState(false);
  const [activeHandle, setActiveHandle] = useState<TrimHandle>("position");
  const mediaRef = useRef<HTMLVideoElement & HTMLAudioElement>(null);
  const startFieldRef = useRef<HTMLInputElement | null>(null);

  const parsedStart = parsePresentationMediaTrimTime(startDraft);
  const parsedEnd = parsePresentationMediaTrimTime(endDraft);
  const startMs = parsedStart ?? trimStartMs;
  const endMs = parsedEnd ?? trimEndMs;
  const valid =
    parsedStart !== null &&
    parsedEnd !== null &&
    parsedStart >= 0 &&
    parsedEnd > parsedStart &&
    parsedEnd <= durationMs;

  const seek = useCallback((ms: number) => {
    setPositionMs(ms);
    const media = mediaRef.current;
    if (media && Number.isFinite(media.duration)) {
      media.currentTime = ms / 1_000;
    }
  }, []);

  const setStart = useCallback(
    (ms: number) => {
      const next = Math.max(0, Math.min(ms, endMs - 1));
      setStartDraft(formatPresentationMediaTrimTime(next));
      seek(next);
    },
    [endMs, seek],
  );
  const setEnd = useCallback(
    (ms: number) => {
      const next = Math.min(durationMs, Math.max(ms, startMs + 1));
      setEndDraft(formatPresentationMediaTrimTime(next));
      seek(next);
    },
    [durationMs, seek, startMs],
  );

  const nudge = (direction: -1 | 1) => {
    const delta = direction * PRESENTATION_MEDIA_TRIM_NUDGE_MS;
    if (activeHandle === "start") setStart(startMs + delta);
    else if (activeHandle === "end") setEnd(endMs + delta);
    else seek(Math.max(startMs, Math.min(endMs, positionMs + delta)));
  };

  const togglePlay = () => {
    const media = mediaRef.current;
    if (!media) return;
    if (!media.paused) {
      media.pause();
      return;
    }
    // Plays the kept part: from the play position when it is inside the
    // markers, otherwise from the start marker.
    const from =
      positionMs >= startMs && positionMs < endMs ? positionMs : startMs;
    media.currentTime = from / 1_000;
    void media.play().catch(() => setPlaying(false));
  };

  // Stop at the end marker, as the dialog's preview does.
  useEffect(() => {
    const media = mediaRef.current;
    if (!media) return;
    const onTime = () => {
      const ms = Math.round(media.currentTime * 1_000);
      if (!media.paused && ms >= endMs) {
        media.pause();
        setPositionMs(endMs);
        return;
      }
      setPositionMs(ms);
    };
    media.addEventListener("timeupdate", onTime);
    return () => media.removeEventListener("timeupdate", onTime);
  }, [endMs]);

  const keptDurationMs = Math.max(0, endMs - startMs);
  const percent = (ms: number) =>
    `${durationMs > 0 ? (Math.min(Math.max(ms, 0), durationMs) / durationMs) * 100 : 0}%`;
  const title = t(
    kind === "video" ? "presentation.media.trimVideo" : "presentation.media.trimAudio",
  );

  return (
    <OfficeDialogShell
      open
      title={title}
      description=""
      defaultActionLabel={t("presentation.media.trimApply")}
      cancelActionLabel={t("presentation.media.trimCancel")}
      defaultActionDisabled={!valid}
      onDefaultAction={() => {
        if (parsedStart === null || parsedEnd === null) return;
        onApply(parsedStart, parsedEnd);
      }}
      onCancel={onClose}
      initialFocusRef={startFieldRef}
      testIdPrefix="presentation-media-trim"
      defaultActionTestId="presentation-media-trim-apply"
    >
      {src && kind === "video" ? (
        <video
          ref={mediaRef}
          src={src}
          preload="metadata"
          playsInline
          onLoadedMetadata={(event) => {
            event.currentTarget.currentTime = positionMs / 1_000;
          }}
          onPlay={() => setPlaying(true)}
          onPause={() => setPlaying(false)}
          style={{ width: "100%", maxHeight: 240, background: "#000" }}
          data-testid="presentation-media-trim-preview"
        />
      ) : src ? (
        <audio
          ref={mediaRef}
          src={src}
          preload="metadata"
          onLoadedMetadata={(event) => {
            event.currentTarget.currentTime = positionMs / 1_000;
          }}
          onPlay={() => setPlaying(true)}
          onPause={() => setPlaying(false)}
          hidden
          data-testid="presentation-media-trim-preview"
        />
      ) : null}
      <div
        className={dialogParts.row}
        style={{ justifyContent: "space-between", alignItems: "baseline" }}
      >
        <span data-testid="presentation-media-trim-duration">
          {t("presentation.media.trimDurationLabel")}{" "}
          <output>{formatPresentationMediaTrimTime(keptDurationMs)}</output>
        </span>
        <output
          aria-label={t("presentation.media.trimPosition")}
          data-testid="presentation-media-trim-position-time"
        >
          {formatPresentationMediaTrimTime(positionMs)}
        </output>
      </div>
      {/* The timeline: the kept span is shaded between the start (green) and
          end (red) markers; the play position is the thin line. */}
      <div
        style={{ position: "relative", height: 56 }}
        data-testid="presentation-media-trim-timeline"
      >
        <div
          aria-hidden="true"
          style={{
            position: "absolute",
            top: 22,
            left: 0,
            right: 0,
            height: 12,
            borderRadius: 3,
            background: "var(--office-dialog-field-border, #c8c8c8)",
          }}
        >
          <div
            style={{
              position: "absolute",
              top: 0,
              bottom: 0,
              left: percent(startMs),
              width: `calc(${percent(endMs)} - ${percent(startMs)})`,
              background: "var(--office-accent, #2b579a)",
              opacity: 0.45,
              borderRadius: 3,
            }}
          />
          <div
            style={{
              position: "absolute",
              top: -6,
              bottom: -6,
              left: percent(positionMs),
              width: 2,
              background: "currentColor",
            }}
          />
        </div>
        <input
          type="range"
          min={0}
          max={durationMs}
          step={1}
          value={positionMs}
          aria-label={t("presentation.media.trimPosition")}
          aria-valuetext={formatPresentationMediaTrimTime(positionMs)}
          onFocus={() => setActiveHandle("position")}
          onChange={(event) => {
            setActiveHandle("position");
            seek(Number(event.currentTarget.value));
          }}
          data-testid="presentation-media-trim-position"
          style={{ position: "absolute", top: 0, left: 0, width: "100%", margin: 0 }}
        />
        <input
          type="range"
          min={0}
          max={durationMs}
          step={1}
          value={startMs}
          aria-label={t("presentation.media.trimStartMarker")}
          aria-valuetext={formatPresentationMediaTrimTime(startMs)}
          onFocus={() => setActiveHandle("start")}
          onChange={(event) => {
            setActiveHandle("start");
            setStart(Number(event.currentTarget.value));
          }}
          data-testid="presentation-media-trim-start-marker"
          style={{
            position: "absolute",
            top: 36,
            left: 0,
            width: "100%",
            margin: 0,
            accentColor: "#2e9e44",
          }}
        />
        <input
          type="range"
          min={0}
          max={durationMs}
          step={1}
          value={endMs}
          aria-label={t("presentation.media.trimEndMarker")}
          aria-valuetext={formatPresentationMediaTrimTime(endMs)}
          onFocus={() => setActiveHandle("end")}
          onChange={(event) => {
            setActiveHandle("end");
            setEnd(Number(event.currentTarget.value));
          }}
          data-testid="presentation-media-trim-end-marker"
          style={{
            position: "absolute",
            top: 14,
            left: 0,
            width: "100%",
            margin: 0,
            accentColor: "#c42b1c",
          }}
        />
      </div>
      <div
        className={dialogParts.row}
        style={{ justifyContent: "center", gap: 8 }}
      >
        <button
          type="button"
          className={dialogParts.compactButton}
          aria-label={t("presentation.media.trimNudgeBackward")}
          title={t("presentation.media.trimNudgeBackward")}
          onClick={() => nudge(-1)}
          data-testid="presentation-media-trim-nudge-backward"
        >
          <ChevronLeft size={16} aria-hidden="true" />
        </button>
        <button
          type="button"
          className={dialogParts.compactButton}
          aria-label={t(
            playing ? "presentation.media.trimPause" : "presentation.media.trimPlay",
          )}
          title={src ? undefined : t("presentation.media.trimPreviewUnavailable")}
          disabled={!src}
          onClick={togglePlay}
          data-testid="presentation-media-trim-play"
        >
          {playing ? (
            <Pause size={16} aria-hidden="true" />
          ) : (
            <Play size={16} aria-hidden="true" />
          )}
        </button>
        <button
          type="button"
          className={dialogParts.compactButton}
          aria-label={t("presentation.media.trimNudgeForward")}
          title={t("presentation.media.trimNudgeForward")}
          onClick={() => nudge(1)}
          data-testid="presentation-media-trim-nudge-forward"
        >
          <ChevronRight size={16} aria-hidden="true" />
        </button>
      </div>
      <div className={dialogParts.row}>
        <label className={dialogParts.field} style={{ flex: 1 }}>
          <span className={dialogParts.label}>
            {t("presentation.media.trimStartTime")}
          </span>
          <input
            ref={startFieldRef}
            type="text"
            inputMode="decimal"
            value={startDraft}
            aria-invalid={parsedStart === null || !valid ? true : undefined}
            onFocus={() => setActiveHandle("start")}
            onChange={(event) => setStartDraft(event.currentTarget.value)}
            onBlur={() => {
              if (parsedStart !== null) {
                setStartDraft(formatPresentationMediaTrimTime(parsedStart));
              }
            }}
            data-testid="presentation-media-trim-start"
          />
        </label>
        <label className={dialogParts.field} style={{ flex: 1 }}>
          <span className={dialogParts.label}>
            {t("presentation.media.trimEndTime")}
          </span>
          <input
            type="text"
            inputMode="decimal"
            value={endDraft}
            aria-invalid={parsedEnd === null || !valid ? true : undefined}
            onFocus={() => setActiveHandle("end")}
            onChange={(event) => setEndDraft(event.currentTarget.value)}
            onBlur={() => {
              if (parsedEnd !== null) {
                setEndDraft(formatPresentationMediaTrimTime(parsedEnd));
              }
            }}
            data-testid="presentation-media-trim-end"
          />
        </label>
      </div>
    </OfficeDialogShell>
  );
}
