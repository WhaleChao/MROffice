/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Insert > Online Pictures: URL validation and the bounded byte read of a
 * fetched picture.
 *
 * Moved verbatim out of `components/notebook/contents/PresentationWindow.tsx`
 * (structural split, no behaviour change).
 */

import { MAX_PRESENTATION_IMAGE_FILE_BYTES } from "@mrnote/core/web/lib/presentation-image-file.ts";

const ONLINE_PICTURE_URL_MAX_CHARS = 2_048;

export const ONLINE_PICTURE_FETCH_TIMEOUT_MS = 15_000;

const ONLINE_PICTURE_FORBIDDEN_HOST_SUFFIX =
  /\.(?:localhost|local|internal|home|lan|test)$/iu;

/**
 * Online pictures are fetched directly by the browser. Reject local, private,
 * documentation-only, and credential-bearing targets so this user-facing
 * command cannot become an intranet probe. Redirects are checked again after
 * the fetch resolves.
 */
export function parseOnlinePictureUrl(value: string): URL | null {
  const candidate = value.trim();
  if (
    !candidate ||
    candidate.length > ONLINE_PICTURE_URL_MAX_CHARS ||
    /[\s\u007f-\u009f]/u.test(candidate)
  ) {
    return null;
  }
  let parsed: URL;
  try {
    parsed = new URL(candidate);
  } catch {
    return null;
  }
  const hostname = parsed.hostname
    .replace(/^\[|\]$/gu, "")
    .replace(/\.$/u, "")
    .toLowerCase();
  if (
    parsed.protocol !== "https:" ||
    !hostname ||
    parsed.username ||
    parsed.password ||
    hostname === "localhost" ||
    ONLINE_PICTURE_FORBIDDEN_HOST_SUFFIX.test(hostname)
  ) {
    return null;
  }

  const ipv4 = hostname.split(".").map(Number);
  if (
    ipv4.length === 4 &&
    ipv4.every(
      (part) => Number.isInteger(part) && part >= 0 && part <= 255,
    )
  ) {
    const [a, b, c] = ipv4;
    if (
      a === 0 ||
      a === 10 ||
      a === 127 ||
      (a === 100 && b >= 64 && b <= 127) ||
      (a === 169 && b === 254) ||
      (a === 172 && b >= 16 && b <= 31) ||
      (a === 192 && b === 0 && (c === 0 || c === 2)) ||
      (a === 192 && b === 168) ||
      (a === 198 && (b === 18 || b === 19)) ||
      (a === 198 && b === 51 && c === 100) ||
      (a === 203 && b === 0 && c === 113) ||
      a >= 224
    ) {
      return null;
    }
  } else if (hostname.includes(":")) {
    if (
      hostname === "::" ||
      hostname === "::1" ||
      hostname.startsWith("fc") ||
      hostname.startsWith("fd") ||
      /^fe[89ab]/u.test(hostname) ||
      hostname.startsWith("::ffff:") ||
      hostname.startsWith("2001:db8:")
    ) {
      return null;
    }
  } else if (!hostname.includes(".")) {
    return null;
  }
  return parsed;
}

export async function readBoundedOnlinePictureBytes(
  response: Response,
): Promise<Uint8Array> {
  if (!response.body) {
    const bytes = new Uint8Array(await response.arrayBuffer());
    if (
      bytes.byteLength <= 0 ||
      bytes.byteLength > MAX_PRESENTATION_IMAGE_FILE_BYTES
    ) {
      throw new Error("online-picture-size");
    }
    return bytes;
  }

  // The reader is not released: the body is read to its end or cancelled,
  // and a stream is not used again either way. Releasing it rejects the
  // reader's settled `closed` promise with a fresh error, whose stack Node
  // formats at once -- in a test runner mapping the editor's source that
  // overflowed the stack and turned a fetched picture into a failure.
  const reader = response.body.getReader();
  const chunks: Uint8Array[] = [];
  let byteLength = 0;
  for (;;) {
    const chunk = await reader.read();
    if (chunk.done) break;
    byteLength += chunk.value.byteLength;
    if (byteLength > MAX_PRESENTATION_IMAGE_FILE_BYTES) {
      await reader.cancel();
      throw new Error("online-picture-size");
    }
    chunks.push(chunk.value);
  }
  if (byteLength <= 0) throw new Error("online-picture-size");
  const bytes = new Uint8Array(byteLength);
  let offset = 0;
  for (const chunk of chunks) {
    bytes.set(chunk, offset);
    offset += chunk.byteLength;
  }
  return bytes;
}
