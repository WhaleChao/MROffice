/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/** The picture already sits at the end the move would put it. */
export const MRAI_NATIVE_EMBEDDED_IMAGE_ALREADY_FRONT =
  "PDF image is already in front of its page content." as const;

export const MRAI_NATIVE_EMBEDDED_IMAGE_ALREADY_BACK =
  "PDF image is already behind its page content." as const;

/** The picture is drawn under no clip a move could take it out from under. */
export const MRAI_NATIVE_EMBEDDED_IMAGE_NOT_CLIPPED =
  "PDF image is drawn under no clip it can leave." as const;

export function pdfNativeEmbeddedImageIsAlreadyArranged(
  error: unknown,
): "front" | "back" | null {
  const message = error instanceof Error ? error.message : "";
  if (message === MRAI_NATIVE_EMBEDDED_IMAGE_ALREADY_FRONT) return "front";
  if (message === MRAI_NATIVE_EMBEDDED_IMAGE_ALREADY_BACK) return "back";
  return null;
}
