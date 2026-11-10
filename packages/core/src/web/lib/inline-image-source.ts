/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
const MAX_INLINE_IMAGE_BYTES = 2_000_000;
const DATA_IMAGE_URL_RE = /^data:(image\/[a-z0-9.+-]+);base64,(.+)$/is;

const SAFE_INLINE_SVG_ELEMENTS = new Set([
  "svg",
  "g",
  "defs",
  "title",
  "desc",
  "path",
  "rect",
  "circle",
  "ellipse",
  "line",
  "polyline",
  "polygon",
  "lineargradient",
  "radialgradient",
  "stop",
  "clippath",
  "mask",
  // Words, reused symbols, patterns and markers: the server's Office Graphic
  // profile (`office_graphic._SAFE_ELEMENTS`) states the same, and its SVG
  // sanitizer writes an uploaded SVG into exactly this.
  "text",
  "tspan",
  "symbol",
  "use",
  "pattern",
  "marker",
]);

/**
 * The one stylesheet the profile admits: what PowerPoint writes into an SVG
 * graphic for Graphics Fill / Outline (PPT-A161), a root `<style>` whose only
 * rules are Office's colour classes (`.MsftOfcThm_Accent2_Fill_v2 {
 * fill:#E97132; }`). The server's Office Graphic profile states the same
 * (`office_graphic.OFFICE_CLASS_RULE`).
 */
const OFFICE_CLASS_SHEET_RE =
  /^\s*(?:\.MsftOfc(?:Thm|Responsive)_[A-Za-z0-9_]{1,120}\s*\{\s*(?:fill|stroke)\s*:\s*#[0-9a-f]{6}\s*;?\s*\}\s*)+$/i;
const STYLE_ELEMENT_RE = /<style\b([^>]*)>([\s\S]*?)<\/style\s*>/gi;

/** The payload without its Office class sheets, or null for any other stylesheet. */
function withoutOfficeClassSheets(payload: string): string | null {
  let refused = false;
  const rest = payload.replace(STYLE_ELEMENT_RE, (_match, attributes: string, body: string) => {
    if (attributes.trim() || !OFFICE_CLASS_SHEET_RE.test(body)) refused = true;
    return "";
  });
  return refused ? null : rest;
}

/**
 * A deliberately small, image-only SVG profile.
 *
 * Office Graphics are rendered through `<img>`, never mounted as live DOM, but
 * the source still crosses import/export boundaries. Reject every active,
 * externally-referencing, CSS-bearing or extensible node up front so a later
 * renderer cannot accidentally broaden the trust boundary.
 */
export function isSafeInlineSvgPayload(payload: string): boolean {
  if (
    !payload ||
    payload.length > MAX_INLINE_IMAGE_BYTES ||
    payload.includes("\0")
  ) {
    return false;
  }
  const normalized = payload.trim();
  if (!/^<\?xml\b[^>]*>\s*<svg\b|^<svg\b/i.test(normalized)) return false;
  const unstyled = withoutOfficeClassSheets(normalized);
  if (unstyled === null) return false;
  if (
    /<!DOCTYPE|<!ENTITY|<\/?(?:script|foreignObject|iframe|object|embed|audio|video|canvas|style|link|meta|image)\b/i.test(
      unstyled,
    ) ||
    /\son[a-z0-9_-]+\s*=/i.test(normalized) ||
    /\sstyle\s*=/i.test(normalized) ||
    /(?:href|xlink:href)\s*=\s*["'](?!#)/i.test(normalized) ||
    /url\(\s*(?!#[A-Za-z_][A-Za-z0-9_.:-]*\s*\))/i.test(normalized)
  ) {
    return false;
  }

  if (typeof DOMParser === "undefined") {
    const elementNames = Array.from(
      unstyled.matchAll(/<\s*\/?\s*([A-Za-z][A-Za-z0-9:._-]*)\b/g),
      (match) => match[1].split(":").pop()?.toLowerCase() ?? "",
    );
    return (
      elementNames.length > 0 &&
      elementNames.every((name) => SAFE_INLINE_SVG_ELEMENTS.has(name))
    );
  }
  const document = new DOMParser().parseFromString(normalized, "image/svg+xml");
  if (
    document.getElementsByTagName("parsererror").length > 0 ||
    document.documentElement.localName.toLowerCase() !== "svg"
  ) {
    return false;
  }
  for (const element of Array.from(document.querySelectorAll("*"))) {
    if (element.localName.toLowerCase() === "style") {
      // Only the root's own Office class sheet (checked above).
      if (
        element.parentElement !== document.documentElement ||
        element.attributes.length > 0 ||
        element.children.length > 0 ||
        !OFFICE_CLASS_SHEET_RE.test(element.textContent ?? "")
      ) {
        return false;
      }
      continue;
    }
    if (!SAFE_INLINE_SVG_ELEMENTS.has(element.localName.toLowerCase())) {
      return false;
    }
    for (const attribute of Array.from(element.attributes)) {
      const name = attribute.name.toLowerCase();
      const value = attribute.value.trim();
      if (name.startsWith("on") || name === "style") return false;
      if (
        (name === "href" || name === "xlink:href") &&
        !value.startsWith("#")
      ) {
        return false;
      }
      if (
        /url\(/i.test(value) &&
        !/^url\(\s*#[A-Za-z_][A-Za-z0-9_.:-]*\s*\)$/i.test(value)
      ) {
        return false;
      }
    }
  }
  return true;
}

function imagePayloadMatchesMime(mimeType: string, payload: string): boolean {
  if (mimeType === "image/png") {
    return payload.startsWith("\x89PNG\r\n\x1a\n");
  }
  if (mimeType === "image/jpeg" || mimeType === "image/jpg") {
    return payload.startsWith("\xff\xd8\xff");
  }
  if (mimeType === "image/gif") {
    return payload.startsWith("GIF87a") || payload.startsWith("GIF89a");
  }
  if (mimeType === "image/webp") {
    return (
      payload.length >= 12 &&
      payload.startsWith("RIFF") &&
      payload.slice(8, 12) === "WEBP"
    );
  }
  if (mimeType === "image/bmp") {
    return payload.startsWith("BM");
  }
  if (mimeType === "image/svg+xml") {
    return isSafeInlineSvgPayload(payload);
  }
  return false;
}

/** Base64 characters enough to hold any raster's signature (`imagePayloadMatchesMime`). */
const RASTER_SIGNATURE_BASE64_CHARS = 32;










export function isDrawableImageSource(src: string): boolean {
  const trimmed = src.trim();
  if (!trimmed.toLowerCase().startsWith("data:image/")) return true;
  const comma = trimmed.indexOf(",");
  const head = comma > 0 ? trimmed.slice(0, comma) : "";
  const mime = /^data:(image\/[a-z0-9.+-]+);base64$/i.exec(head)?.[1].toLowerCase();
  if (!mime || comma + 1 >= trimmed.length) return false;
  if (mime === "image/svg+xml") return isValidInlineImageSource(trimmed);
  const decode = globalThis.atob;
  if (typeof decode !== "function") return false;
  try {
    return imagePayloadMatchesMime(mime, decode(trimmed.slice(comma + 1, comma + 1 + RASTER_SIGNATURE_BASE64_CHARS)));
  } catch {
    return false;
  }
}

export function isValidInlineImageSource(src: string): boolean {
  const trimmed = src.trim();
  if (!trimmed.toLowerCase().startsWith("data:image/")) return true;
  const match = DATA_IMAGE_URL_RE.exec(trimmed);
  if (!match) return false;
  const decode = globalThis.atob;
  if (typeof decode !== "function") return false;
  let payload = "";
  try {
    payload = decode(match[2]);
  } catch {
    return false;
  }
  if (!payload || payload.length > MAX_INLINE_IMAGE_BYTES) return false;
  return imagePayloadMatchesMime(match[1].toLowerCase(), payload);
}

export function recordHasInvalidInlineImageSrc(value: unknown): boolean {
  if (!value || typeof value !== "object" || Array.isArray(value)) return false;
  const src = (value as { src?: unknown }).src;
  return typeof src === "string" && !isValidInlineImageSource(src);
}
