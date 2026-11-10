/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import {
  ALLOWED_PERSISTED_RASTER_IMAGE_MIMES,
  MAX_PERSISTED_ATTACHMENT_CHARS,
} from "./output-persistence-budgets";

type AttachmentMap = Readonly<
  Record<string, Readonly<Record<string, unknown>>>
>;

const MAX_ATTACHMENT_PARTS = 10_000;

function ownValue<T>(
  record: Readonly<Record<string, T>>,
  key: string,
): T | undefined {
  try {
    const descriptor = Object.getOwnPropertyDescriptor(record, key);
    return descriptor && "value" in descriptor ? descriptor.value : undefined;
  } catch {
    return undefined;
  }
}

function boundedAttachmentData(value: unknown): string | null {
  if (typeof value === "string") {
    return value.length <= MAX_PERSISTED_ATTACHMENT_CHARS ? value : null;
  }
  let isArray = false;
  try {
    isArray = Array.isArray(value);
  } catch {
    return null;
  }
  if (!isArray) return null;
  let lengthDescriptor: PropertyDescriptor | undefined;
  try {
    lengthDescriptor = Object.getOwnPropertyDescriptor(value, "length");
  } catch {
    return null;
  }
  if (
    !lengthDescriptor ||
    !("value" in lengthDescriptor) ||
    !Number.isSafeInteger(lengthDescriptor.value) ||
    lengthDescriptor.value < 0 ||
    lengthDescriptor.value > MAX_ATTACHMENT_PARTS
  ) {
    return null;
  }
  const parts: string[] = [];
  let chars = 0;
  for (let index = 0; index < lengthDescriptor.value; index += 1) {
    let descriptor: PropertyDescriptor | undefined;
    try {
      descriptor = Object.getOwnPropertyDescriptor(value, String(index));
    } catch {
      return null;
    }
    if (!descriptor || !("value" in descriptor)) return null;
    const part = descriptor.value;
    if (typeof part !== "string") return null;
    chars += part.length;
    if (chars > MAX_PERSISTED_ATTACHMENT_CHARS) return null;
    parts.push(part);
  }
  return parts.join("");
}

function bytesStartWith(bytes: Uint8Array, signature: readonly number[]) {
  return (
    bytes.length >= signature.length &&
    signature.every((value, index) => bytes[index] === value)
  );
}

function mimeMatchesMagic(mime: string, bytes: Uint8Array): boolean {
  if (mime === "image/png") {
    return bytesStartWith(
      bytes,
      [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a],
    );
  }
  if (mime === "image/jpeg") return bytesStartWith(bytes, [0xff, 0xd8, 0xff]);
  if (mime === "image/gif") {
    return (
      bytesStartWith(bytes, [0x47, 0x49, 0x46, 0x38, 0x37, 0x61]) ||
      bytesStartWith(bytes, [0x47, 0x49, 0x46, 0x38, 0x39, 0x61])
    );
  }
  if (mime === "image/webp") {
    return (
      bytesStartWith(bytes, [0x52, 0x49, 0x46, 0x46]) &&
      bytes.length >= 12 &&
      bytes[8] === 0x57 &&
      bytes[9] === 0x45 &&
      bytes[10] === 0x42 &&
      bytes[11] === 0x50
    );
  }
  return mime === "image/bmp" && bytesStartWith(bytes, [0x42, 0x4d]);
}

function validRasterData(mime: string, data: string): boolean {
  if (
    data.length === 0 ||
    data.length > MAX_PERSISTED_ATTACHMENT_CHARS ||
    data.length % 4 !== 0 ||
    !/^[A-Za-z0-9+/]+={0,2}$/.test(data)
  ) {
    return false;
  }
  try {
    // Twelve decoded bytes cover every supported magic signature. Decode only
    // a short prefix so preview validation never allocates a second full image.
    const decoded = atob(data.slice(0, 24));
    const bytes = Uint8Array.from(decoded, (character) =>
      character.charCodeAt(0),
    );
    return mimeMatchesMagic(mime, bytes);
  } catch {
    return false;
  }
}

/**
 * Resolve a Jupyter markdown-cell attachment into a validated raster data URL.
 * `undefined` means the URL is not an attachment and should use the caller's
 * normal URL policy; `null` means it was an attachment but must be rejected.
 */
export function resolveOutputAttachmentImageUrl(
  attachments: AttachmentMap | undefined,
  url: string,
): string | null | undefined {
  if (!url.toLowerCase().startsWith("attachment:")) return undefined;
  if (!attachments) return null;
  const encodedName = url.slice("attachment:".length);
  let decodedName = encodedName;
  try {
    decodedName = decodeURIComponent(encodedName);
  } catch {
    // An invalid escape cannot name a valid imported attachment, but checking
    // the exact literal below keeps legacy names deterministic.
  }
  const representations =
    ownValue(attachments, decodedName) ?? ownValue(attachments, encodedName);
  if (!representations) return null;
  for (const mime of ALLOWED_PERSISTED_RASTER_IMAGE_MIMES) {
    const raw = ownValue(representations, mime);
    const data = boundedAttachmentData(raw);
    if (data !== null && validRasterData(mime, data)) {
      return `data:${mime};base64,${data}`;
    }
  }
  return null;
}
