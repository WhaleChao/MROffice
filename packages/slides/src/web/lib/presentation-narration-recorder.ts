/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/** Browser-only microphone formats that the native PPTX narration writer accepts. */
export const PRESENTATION_NARRATION_RECORDER_MIME_TYPES = [
  "audio/mp4;codecs=mp4a.40.2",
  "audio/mp4",
  "audio/ogg;codecs=opus",
  "audio/ogg",
] as const;











const PRESENTATION_PAGE_STRING_MAX_LENGTH = 1_048_576;
const LONGEST_NARRATION_DATA_URL_PREFIX = "data:audio/mp4;codecs=mp4a.40.2;base64,".length;
export const MAX_PRESENTATION_NARRATION_RECORDING_BYTES =
  Math.floor((PRESENTATION_PAGE_STRING_MAX_LENGTH - LONGEST_NARRATION_DATA_URL_PREFIX) / 4) * 3;

type MediaRecorderSupport = Pick<typeof MediaRecorder, "isTypeSupported">;

export function preferredPresentationNarrationRecorderMime(
  recorder: MediaRecorderSupport | undefined = typeof MediaRecorder ===
  "undefined"
    ? undefined
    : MediaRecorder,
): (typeof PRESENTATION_NARRATION_RECORDER_MIME_TYPES)[number] | null {
  if (!recorder || typeof recorder.isTypeSupported !== "function") return null;
  return (
    PRESENTATION_NARRATION_RECORDER_MIME_TYPES.find((mime) =>
      recorder.isTypeSupported(mime),
    ) ?? null
  );
}

export function presentationNarrationBlobDataUrl(
  blob: Blob,
): Promise<string | null> {
  if (
    blob.size <= 0 ||
    blob.size > MAX_PRESENTATION_NARRATION_RECORDING_BYTES ||
    !PRESENTATION_NARRATION_RECORDER_MIME_TYPES.some(
      (mime) => mime === blob.type,
    )
  ) {
    return Promise.resolve(null);
  }
  return new Promise((resolve) => {
    const reader = new FileReader();
    reader.onerror = () => resolve(null);
    reader.onload = () =>
      resolve(typeof reader.result === "string" ? reader.result : null);
    reader.readAsDataURL(blob);
  });
}
