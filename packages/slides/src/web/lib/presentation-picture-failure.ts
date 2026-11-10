/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Why a picture was not inserted into (or did not replace one on) a slide,
 * as a sentence naming the file. The failure codes of the preparation
 * (`presentation-image-file`), of the insert itself and of the upload all
 * end here, so no toast shows an internal name such as "bad-magic".
 */

import { formatFileSize, uploadFailureReason } from "@mrnote/core/web/lib/errors/upload-failure.ts";
import {
  MAX_PRESENTATION_IMAGE_DIMENSION,
  MAX_PRESENTATION_IMAGE_FILE_BYTES,
  MAX_PRESENTATION_IMAGE_PIXELS,
  PresentationImageFileError,
} from "@mrnote/core/web/lib/presentation-image-file.ts";

type Translate = (key: string, values?: Record<string, string | number>) => string;

function codeOf(error: unknown): string | null {
  let current: unknown = error;
  for (let depth = 0; depth < 4 && current; depth += 1) {
    if (current instanceof PresentationImageFileError) return `file:${current.code}`;
    const candidate = current as { name?: unknown; code?: unknown; cause?: unknown };
    if (
      typeof candidate.code === "string" &&
      (candidate.name === "PresentationImageInsertError" ||
        candidate.name === "PresentationPictureReplaceError")
    ) {
      if (candidate.cause) {
        const inner = codeOf(candidate.cause);
        if (inner) return inner;
      }
      return `insert:${candidate.code}`;
    }
    current = candidate.cause;
  }
  return null;
}

export function presentationPictureFailureDetail(
  error: unknown,
  file: { readonly name: string; readonly size: number },
  t: Translate,
  maxBlocksPerSlide: number,
): string {
  const name = file.name || t("presentation.image.failure.unnamed");
  switch (codeOf(error)) {
    case "file:unsupported-type":
      return t("presentation.image.failure.unsupportedType", { file: name });
    case "file:too-large":
      return t("presentation.image.failure.tooLarge", {
        file: name,
        size: formatFileSize(file.size),
        limit: formatFileSize(MAX_PRESENTATION_IMAGE_FILE_BYTES),
      });
    case "file:bad-dimensions":
      return t("presentation.image.failure.tooManyPixels", {
        file: name,
        edge: MAX_PRESENTATION_IMAGE_DIMENSION.toLocaleString(),
        megapixels: Math.round(MAX_PRESENTATION_IMAGE_PIXELS / 1_000_000),
      });
    case "file:empty":
    case "file:bad-magic":
    case "file:decode-failed":
      return t("presentation.image.failure.unreadable", { file: name });
    case "file:svg-not-drawable":
      return t("presentation.image.failure.svgNotDrawable", { file: name });
    case "file:cache-failed":
      return t("presentation.image.failure.cacheFailed", { file: name });
    case "insert:slide-full":
      return t("presentation.image.failure.slideFull", { limit: maxBlocksPerSlide });
    case "insert:stale-target":
    case "insert:no-target":
      return t("presentation.image.failure.slideChanged");
    case "insert:busy":
      return t("presentation.image.failure.busy");
    default: {
      const upload = uploadFailureReason(error, file);
      if (upload) return t("presentation.image.failure.upload", { file: name, reason: upload });
      return t("presentation.image.failure.generic", { file: name });
    }
  }
}
