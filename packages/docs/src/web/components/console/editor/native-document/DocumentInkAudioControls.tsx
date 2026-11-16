"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Draw ▸ Record and its playback (lib/ink-audio-recording): Record starts
 * the microphone and the strokes written meanwhile keep their moment in the
 * recording; Stop stores it on the page. Recordings lists the page's
 * recordings; one played dims the writing still to come and a tapped stroke
 * plays from where it was written.
 */

import { useCallback, useEffect, useRef, useState, type MutableRefObject } from "react";
import { useTranslations } from "@mrnote/core/host/next-intl.tsx";

import { OfficeHomeCommand, OfficeHomeMenuTrigger } from "@mrnote/core/web/components/office-ribbon/OfficeHomeRibbon.tsx";
import { OfficeSemanticIcon } from "@mrnote/core/web/components/office-ribbon/icons/OfficeSemanticIcon.tsx";
import {
  DropdownMenu,
  DropdownMenuContent,
  DropdownMenuItem,
  DropdownMenuLabel,
  DropdownMenuTrigger,
} from "@mrnote/core/web/components/ui/dropdown-menu.tsx";
import {
  inkAudioSeekSeconds,
  inkRecordingUrl,
  loadInkRecordings,
  showInkAudioMoment,
  startInkRecording,
  uploadInkRecording,
  type InkRecording,
  type InkRecordingSession,
} from "@mrnote/ink/web/lib/ink-audio-recording.ts";
import { pageInterfaceLocale } from "@mrnote/core/web/lib/i18n/content-locale.ts";

function clock(seconds: number): string {
  const whole = Math.max(0, Math.floor(seconds));
  return `${Math.floor(whole / 60)}:${String(whole % 60).padStart(2, "0")}`;
}

export function DocumentInkAudioControls({
  pageId,
  canRecord,
  sessionRef,
  strokesRoot,
}: {
  pageId: string;
  canRecord: boolean;
  /** The running recording, read by the strokes as they land (their `inkAudio`). */
  sessionRef: MutableRefObject<InkRecordingSession | null>;
  /** Where the page's strokes are (the editor's root). */
  strokesRoot: () => HTMLElement | null;
}) {
  const t = useTranslations("console-notebooks");
  const [recording, setRecording] = useState<{ startedAt: number } | null>(null);
  const [elapsed, setElapsed] = useState(0);
  const [busy, setBusy] = useState(false);
  const [failed, setFailed] = useState<string | null>(null);
  const [recordings, setRecordings] = useState<InkRecording[] | null>(null);
  const [playing, setPlaying] = useState<{ recording: InkRecording; paused: boolean; at: number; duration: number } | null>(null);
  const audioRef = useRef<HTMLAudioElement | null>(null);

  useEffect(() => {
    if (!recording) return undefined;
    const timer = window.setInterval(() => setElapsed((performance.now() - recording.startedAt) / 1000), 500);
    return () => window.clearInterval(timer);
  }, [recording]);

  const record = useCallback(async () => {
    setFailed(null);
    if (sessionRef.current) {
      const session = sessionRef.current;
      sessionRef.current = null;
      setRecording(null);
      setBusy(true);
      try {
        const audio = await session.stop();
        if (audio.size > 0) await uploadInkRecording(pageId, session.id, audio);
        setRecordings(null);
      } catch {
        setFailed(t("toolbar.drawRecordSaveFailed"));
      } finally {
        setBusy(false);
      }
      return;
    }
    try {
      const session = await startInkRecording();
      sessionRef.current = session;
      setElapsed(0);
      setRecording({ startedAt: session.startedAt });
    } catch {
      setFailed(t("toolbar.drawRecordNoMicrophone"));
    }
  }, [pageId, sessionRef, t]);

  // Leaving the page stops the microphone and keeps what was recorded.
  useEffect(() => () => {
    const session = sessionRef.current;
    if (!session) return;
    sessionRef.current = null;
    void session.stop().then((audio) => (audio.size > 0 ? uploadInkRecording(pageId, session.id, audio) : undefined)).catch(() => {});
  }, [pageId, sessionRef]);

  const stopPlaying = useCallback(() => {
    const audio = audioRef.current;
    audioRef.current = null;
    if (audio) {
      audio.pause();
      audio.removeAttribute("src");
    }
    const root = strokesRoot();
    if (root) showInkAudioMoment(root, null, 0);
    setPlaying(null);
  }, [strokesRoot]);
  useEffect(() => stopPlaying, [stopPlaying]);

  const play = useCallback(async (item: InkRecording) => {
    stopPlaying();
    try {
      const audio = new Audio(await inkRecordingUrl(item.attachmentId));
      audioRef.current = audio;
      setPlaying({ recording: item, paused: false, at: 0, duration: 0 });
      await audio.play();
    } catch {
      setFailed(t("toolbar.drawRecordPlayFailed"));
      stopPlaying();
    }
  }, [stopPlaying, t]);

  // While a recording plays: the writing still to come is dimmed, frame by
  // frame, and a stroke tapped plays from where it was written.
  const playingId = playing?.recording.recordingId ?? null;
  useEffect(() => {
    const audio = audioRef.current;
    if (!audio || !playingId) return undefined;
    let frame = 0;
    const tick = () => {
      const root = strokesRoot();
      if (root) showInkAudioMoment(root, playingId, audio.currentTime * 1000);
      setPlaying((current) => current && { ...current, at: audio.currentTime, duration: Number.isFinite(audio.duration) ? audio.duration : current.duration, paused: audio.paused });
      frame = window.requestAnimationFrame(tick);
    };
    frame = window.requestAnimationFrame(tick);
    const onEnded = () => setPlaying((current) => current && { ...current, paused: true });
    audio.addEventListener("ended", onEnded);
    const root = strokesRoot();
    const onPress = (event: PointerEvent) => {
      const seconds = inkAudioSeekSeconds(event.target, playingId);
      if (seconds === null) return;
      audio.currentTime = seconds;
      if (audio.paused) void audio.play().catch(() => {});
    };
    root?.addEventListener("pointerdown", onPress, true);
    return () => {
      window.cancelAnimationFrame(frame);
      audio.removeEventListener("ended", onEnded);
      root?.removeEventListener("pointerdown", onPress, true);
    };
  }, [playingId, strokesRoot]);

  const recordLabel = recording
    ? `${t("toolbar.drawRecordStop")} ${clock(elapsed)}`
    : t("toolbar.drawRecord");
  return (
    <>
      <OfficeHomeCommand
        controlId="draw.record"
        variant="large"
        active={recording !== null}
        disabled={(!canRecord && !recording) || busy}
        icon={<OfficeSemanticIcon semanticId={recording ? "media.stop" : "media.record"} size={24} />}
        label={recordLabel}
        disabledReason={failed ?? undefined}
        onClick={() => void record()}
        testId="document-toolbar-ribbon-draw-record"
      />
      <DropdownMenu modal={false} onOpenChange={(open) => {
        if (open && recordings === null) void loadInkRecordings(pageId).then(setRecordings).catch(() => setRecordings([]));
      }}>
        <DropdownMenuTrigger asChild>
          <OfficeHomeMenuTrigger
            controlId="draw.recordings"
            variant="large"
            hasMenuIndicator
            active={playing !== null}
            icon={<OfficeSemanticIcon semanticId="media.mic" size={24} />}
            label={playing ? `${clock(playing.at)} / ${clock(playing.duration)}` : t("toolbar.drawRecordings")}
            title={failed ?? t("toolbar.drawRecordings")}
            testId="document-toolbar-ribbon-draw-recordings"
          />
        </DropdownMenuTrigger>
        <DropdownMenuContent align="start" className="note-editor-toolbar-menu" data-testid="document-toolbar-ribbon-draw-recordings-menu">
          {playing ? (
            <>
              <DropdownMenuItem
                data-testid="document-toolbar-ribbon-draw-recording-toggle"
                onSelect={() => {
                  const audio = audioRef.current;
                  if (!audio) return;
                  if (audio.paused) void audio.play().catch(() => {});
                  else audio.pause();
                }}
              >
                <OfficeSemanticIcon semanticId={playing.paused ? "media.play" : "media.pause"} size={16} />
                {playing.paused ? t("toolbar.drawRecordingPlay") : t("toolbar.drawRecordingPause")}
              </DropdownMenuItem>
              <DropdownMenuItem data-testid="document-toolbar-ribbon-draw-recording-close" onSelect={stopPlaying}>
                <OfficeSemanticIcon semanticId="media.stop" size={16} />
                {t("toolbar.drawRecordingClose")}
              </DropdownMenuItem>
            </>
          ) : null}
          <DropdownMenuLabel>{t("toolbar.drawRecordings")}</DropdownMenuLabel>
          {recordings === null ? (
            <DropdownMenuItem disabled>…</DropdownMenuItem>
          ) : recordings.length === 0 ? (
            <DropdownMenuItem disabled>{t("toolbar.drawRecordingNone")}</DropdownMenuItem>
          ) : (
            [...recordings].reverse().map((item) => (
              <DropdownMenuItem
                key={item.recordingId}
                data-testid={`document-toolbar-ribbon-draw-recording-${item.recordingId}`}
                onSelect={() => void play(item)}
              >
                <OfficeSemanticIcon semanticId="media.play" size={16} />
                {item.createdAt ? new Date(item.createdAt).toLocaleString(pageInterfaceLocale()) : item.recordingId.slice(0, 8)}
              </DropdownMenuItem>
            ))
          )}
        </DropdownMenuContent>
      </DropdownMenu>
    </>
  );
}
