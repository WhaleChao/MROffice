/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
export type PdfScreenRecordingOutcome = "saved" | "discarded" | "failed";

export interface PdfScreenRecordingSession {
  readonly mimeType: string;
  readonly done: Promise<PdfScreenRecordingOutcome>;
  stop(): void;
  discard(): void;
}

export interface PdfScreenRecordingStartOptions {
  readonly includeSharedAudio: boolean;
  readonly getDisplayMedia?: (
    constraints: DisplayMediaStreamOptions,
  ) => Promise<MediaStream>;
  readonly createRecorder?: (
    stream: MediaStream,
    options?: MediaRecorderOptions,
  ) => MediaRecorder;
  readonly onOutput: (blob: Blob) => boolean | Promise<boolean>;
}

const WEBM_CANDIDATES = [
  "video/webm;codecs=vp9,opus",
  "video/webm;codecs=vp8,opus",
  "video/webm",
] as const;

export function preferredPdfScreenRecordingMimeType(
  supports: (mimeType: string) => boolean,
): string | null {
  return WEBM_CANDIDATES.find((candidate) => supports(candidate)) ?? null;
}

export function pdfScreenRecordingExtension(mimeType: string): "mp4" | "webm" {
  return mimeType.toLowerCase().startsWith("video/mp4") ? "mp4" : "webm";
}

export async function startPdfScreenRecording(
  options: PdfScreenRecordingStartOptions,
): Promise<PdfScreenRecordingSession> {
  const getDisplayMedia =
    options.getDisplayMedia ??
    navigator.mediaDevices?.getDisplayMedia?.bind(navigator.mediaDevices);
  if (!getDisplayMedia || typeof MediaRecorder === "undefined") {
    throw new Error("Screen recording is unavailable");
  }
  const stream = await getDisplayMedia({
    video: true,
    audio: options.includeSharedAudio,
  });
  const videoTracks = stream.getVideoTracks();
  if (videoTracks.length !== 1 || videoTracks[0].readyState === "ended") {
    stream.getTracks().forEach((track) => track.stop());
    throw new Error("Screen recording source is unavailable");
  }

  const mimeType = preferredPdfScreenRecordingMimeType((candidate) =>
    MediaRecorder.isTypeSupported(candidate),
  );
  const createRecorder =
    options.createRecorder ??
    ((recordingStream, recorderOptions) =>
      new MediaRecorder(recordingStream, recorderOptions));
  let recorder: MediaRecorder;
  try {
    recorder = createRecorder(
      stream,
      mimeType ? { mimeType, videoBitsPerSecond: 4_000_000 } : undefined,
    );
  } catch (error) {
    stream.getTracks().forEach((track) => track.stop());
    throw error;
  }

  let discarded = false;
  let settled = false;
  const chunks: Blob[] = [];
  let resolveDone!: (outcome: PdfScreenRecordingOutcome) => void;
  const done = new Promise<PdfScreenRecordingOutcome>((resolve) => {
    resolveDone = resolve;
  });
  const settle = (outcome: PdfScreenRecordingOutcome) => {
    if (settled) return;
    settled = true;
    stream.getTracks().forEach((track) => track.stop());
    resolveDone(outcome);
  };

  recorder.addEventListener("dataavailable", (event) => {
    if (!discarded && event.data.size > 0) chunks.push(event.data);
  });
  recorder.addEventListener("error", () => settle("failed"), { once: true });
  recorder.addEventListener(
    "stop",
    () => {
      if (discarded) {
        settle("discarded");
        return;
      }
      const blob = new Blob(chunks, {
        type: recorder.mimeType || mimeType || "video/webm",
      });
      if (blob.size < 1) {
        settle("failed");
        return;
      }
      void Promise.resolve(options.onOutput(blob))
        .then((saved) => settle(saved ? "saved" : "failed"))
        .catch(() => settle("failed"));
    },
    { once: true },
  );
  videoTracks[0].addEventListener(
    "ended",
    () => {
      if (recorder.state !== "inactive") recorder.stop();
    },
    { once: true },
  );

  try {
    recorder.start(1_000);
  } catch (error) {
    discarded = true;
    settle("failed");
    throw error;
  }

  return Object.freeze({
    mimeType: recorder.mimeType || mimeType || "video/webm",
    done,
    stop: () => {
      if (recorder.state !== "inactive") recorder.stop();
    },
    discard: () => {
      discarded = true;
      chunks.length = 0;
      if (recorder.state !== "inactive") recorder.stop();
      else settle("discarded");
    },
  });
}
