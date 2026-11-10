/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Draw ▸ Record: audio recorded beside the writing, as a note-taking app
 * records a lecture with its notes (Goodnotes' Audio Recording). While it
 * runs, every stroke that lands keeps the recording's id and the moment it
 * was written (`inkAudio`, "<id> <ms>"); the recording is stored as an audio
 * attachment of the page named by that id (`ink-audio-<id>.<ext>`, listed by
 * `GET /pages/{id}/ink-recordings`). Played back, the writing still to come
 * is dimmed as the recording goes on, and a stroke tapped plays from where
 * it was written.
 */

import { apiGet, apiPostFormData } from "@mrnote/core/web/lib/api.ts";
import { inkAudioPlace } from "@mrnote/core/web/lib/document-shapes.ts";

/** A stroke is stamped this much before it landed: it lands after the pause that follows it. */
export const INK_AUDIO_LANDING_LEAD_MS = 900;
/** A stroke tapped plays from a moment before it was written. */
export const INK_AUDIO_SEEK_LEAD_MS = 1_500;

export interface InkRecordingSession {
  readonly id: string;
  readonly startedAt: number;
  /** The moment `at` (a `performance.now()` time) as a stroke's `inkAudio`. */
  stamp(at?: number): string;
  /** Stops the microphone; the recording's audio. */
  stop(): Promise<Blob>;
}

function newId(): string {
  if (typeof crypto !== "undefined" && typeof crypto.randomUUID === "function") return crypto.randomUUID();
  const hex = Array.from({ length: 32 }, () => Math.floor(Math.random() * 16).toString(16)).join("");
  return `${hex.slice(0, 8)}-${hex.slice(8, 12)}-4${hex.slice(13, 16)}-a${hex.slice(17, 20)}-${hex.slice(20, 32)}`;
}

/** Starts recording from the microphone (the browser asks the first time). */
export async function startInkRecording(): Promise<InkRecordingSession> {
  const stream = await navigator.mediaDevices.getUserMedia({ audio: true });
  const preferred = ["audio/webm;codecs=opus", "audio/webm", "audio/mp4"].find((mime) =>
    typeof MediaRecorder !== "undefined" && MediaRecorder.isTypeSupported(mime),
  );
  const recorder = preferred ? new MediaRecorder(stream, { mimeType: preferred }) : new MediaRecorder(stream);
  const chunks: Blob[] = [];
  recorder.addEventListener("dataavailable", (event) => {
    if (event.data.size > 0) chunks.push(event.data);
  });
  recorder.start(1_000);
  const id = newId();
  const startedAt = performance.now();
  return {
    id,
    startedAt,
    stamp: (at = performance.now()) => `${id} ${Math.max(0, Math.round(at - startedAt - INK_AUDIO_LANDING_LEAD_MS))}`,
    stop: () =>
      new Promise<Blob>((resolve) => {
        recorder.addEventListener(
          "stop",
          () => {
            for (const track of stream.getTracks()) track.stop();
            resolve(new Blob(chunks, { type: recorder.mimeType || preferred || "audio/webm" }));
          },
          { once: true },
        );
        if (recorder.state !== "inactive") recorder.stop();
        else {
          for (const track of stream.getTracks()) track.stop();
          resolve(new Blob(chunks, { type: recorder.mimeType || "audio/webm" }));
        }
      }),
  };
}

export function inkRecordingFileName(id: string, type: string): string {
  const ext = type.includes("mp4") ? "m4a" : type.includes("ogg") ? "ogg" : "webm";
  return `ink-audio-${id}.${ext}`;
}

/** Stores the recording as an audio attachment of the page. */
export async function uploadInkRecording(pageId: string, id: string, audio: Blob): Promise<void> {
  const name = inkRecordingFileName(id, audio.type);
  const form = new FormData();
  form.append("file", new File([audio], name, { type: audio.type.split(";")[0] || "audio/webm" }));
  form.append("title", name);
  await apiPostFormData(`/api/v1/pages/${encodeURIComponent(pageId)}/attachments/upload`, form);
}

export interface InkRecording {
  readonly recordingId: string;
  readonly attachmentId: string;
  readonly createdAt: string | null;
}

export async function loadInkRecordings(pageId: string): Promise<InkRecording[]> {
  const answer = await apiGet<{ recordings?: unknown }>(`/api/v1/pages/${encodeURIComponent(pageId)}/ink-recordings`);
  const rows = Array.isArray(answer?.recordings) ? answer.recordings : [];
  return rows.flatMap((row) => {
    const value = row as Record<string, unknown>;
    return typeof value.recording_id === "string" && typeof value.attachment_id === "string"
      ? [{
          recordingId: value.recording_id,
          attachmentId: value.attachment_id,
          createdAt: typeof value.created_at === "string" ? value.created_at : null,
        }]
      : [];
  });
}

/** Where a recording's audio is fetched from (the attachment's signed URL). */
export async function inkRecordingUrl(attachmentId: string): Promise<string> {
  const answer = await apiGet<{ url?: string }>(`/api/v1/attachments/${encodeURIComponent(attachmentId)}/url`);
  if (!answer?.url) throw new Error("ink_recording_unavailable");
  return answer.url;
}

/**
 * Shows the writing as it was at `atMs` into `recordingId`: the strokes
 * written later wear `data-ink-audio-ahead`, the rest none. `null` clears it
 * (playback over). Strokes are the editor's elements carrying
 * `data-ink-audio` (the shape node view marks them).
 */
export function showInkAudioMoment(root: ParentNode, recordingId: string | null, atMs: number): void {
  for (const element of root.querySelectorAll<HTMLElement>("[data-ink-audio]")) {
    const place = inkAudioPlace(element.dataset.inkAudio);
    const ahead = recordingId !== null && place?.recordingId === recordingId && place.atMs > atMs;
    if (ahead) element.dataset.inkAudioAhead = "true";
    else if (element.dataset.inkAudioAhead) delete element.dataset.inkAudioAhead;
  }
}

/** The moment a tapped stroke plays from, in seconds; null when it was not written in `recordingId`. */
export function inkAudioSeekSeconds(target: EventTarget | null, recordingId: string): number | null {
  const element = target instanceof Element ? target.closest<HTMLElement>("[data-ink-audio]") : null;
  const place = element ? inkAudioPlace(element.dataset.inkAudio) : null;
  if (!place || place.recordingId !== recordingId) return null;
  return Math.max(0, place.atMs - INK_AUDIO_SEEK_LEAD_MS) / 1000;
}
