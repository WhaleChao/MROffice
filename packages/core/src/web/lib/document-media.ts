/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import {
  EMBED_IFRAME_REFERRER_POLICY,
  EMBED_IFRAME_SANDBOX,
  sanitizeEmbedUrl,
} from "./presentation-embed";
import {
  onlineVideoEmbedCodeSrc,
  vimeoVideoId,
  youTubeVideoId,
} from "./online-video";

export const DOCUMENT_MEDIA_NODE_NAME = "documentMedia";
export const DOCUMENT_MEDIA_SCHEMA_VERSION = 1;
export const DOCUMENT_MEDIA_DEFAULT_WIDTH = 640;
export const DOCUMENT_MEDIA_DEFAULT_HEIGHT = 360;
export const DOCUMENT_MEDIA_MIN_WIDTH = 160;
export const DOCUMENT_MEDIA_MAX_WIDTH = 1_200;
export const DOCUMENT_MEDIA_MIN_HEIGHT = 90;
export const DOCUMENT_MEDIA_MAX_HEIGHT = 900;
export const DOCUMENT_MEDIA_MAX_TITLE_CHARS = 240;
export const DOCUMENT_MEDIA_MAX_URL_CHARS = 4_096;

export const DOCUMENT_MEDIA_KINDS = ["iframe", "video"] as const;
export type DocumentMediaKind = (typeof DOCUMENT_MEDIA_KINDS)[number];

/**
 * Word's documented online-video providers. Some providers expose only
 * tenant/embed URLs; accepting the host is not enough, so the path gates below
 * still require an embeddable shape before any iframe is mounted.
 */
export const DOCUMENT_MEDIA_IFRAME_HOSTS = [
  "youtube.com",
  "youtube-nocookie.com",
  "youtu.be",
  "vimeo.com",
  "player.vimeo.com",
  "ted.com",
  "embed.ted.com",
  "slideshare.net",
  "microsoftstream.com",
  "stream.office.com",
  "sharepoint.com",
  "onedrive.live.com",
] as const;

export const DOCUMENT_MEDIA_IFRAME_SANDBOX = EMBED_IFRAME_SANDBOX;
export const DOCUMENT_MEDIA_IFRAME_REFERRER_POLICY =
  EMBED_IFRAME_REFERRER_POLICY;

const DOCUMENT_MEDIA_KIND_SET = new Set<string>(DOCUMENT_MEDIA_KINDS);
const DOCUMENT_MEDIA_CONTROL_RE = /[\u0000-\u001f\u007f-\u009f]/;
const DOCUMENT_MEDIA_URL_FORBIDDEN_RE = /[\s\u007f-\u009f]/;

/**
 * Word places an online video in line with text, so the paragraph that holds
 * it aligns it (Home > Align Left/Center/Right). A video cannot be justified;
 * unset reads as centred, how the block has always been drawn.
 */
export const DOCUMENT_MEDIA_ALIGNMENTS = ["left", "center", "right"] as const;
export type DocumentMediaAlignment = (typeof DOCUMENT_MEDIA_ALIGNMENTS)[number];
const DOCUMENT_MEDIA_ALIGNMENT_SET = new Set<string>(DOCUMENT_MEDIA_ALIGNMENTS);

export function normalizeDocumentMediaAlignment(
  value: unknown,
): DocumentMediaAlignment | null {
  return typeof value === "string" && DOCUMENT_MEDIA_ALIGNMENT_SET.has(value)
    ? (value as DocumentMediaAlignment)
    : null;
}

export interface DocumentMediaAttrs {
  schemaVersion: 1;
  kind: DocumentMediaKind;
  url: string;
  title: string;
  poster: string | null;
  width: number;
  height: number;
  textAlign?: DocumentMediaAlignment | null;
}

export interface NormalizeDocumentMediaOptions {
  /** Accept a pasted `<iframe src="…">` and retain only its validated src. */
  allowEmbedCode?: boolean;
}

function hasUnpairedSurrogate(value: string): boolean {
  for (const character of value) {
    if (character.length !== 1) continue;
    const codeUnit = character.charCodeAt(0);
    if (codeUnit >= 0xd800 && codeUnit <= 0xdfff) return true;
  }
  return false;
}

function replaceUnpairedSurrogates(value: string): string {
  return [...value]
    .map((character) =>
      character.length === 1 &&
      character.charCodeAt(0) >= 0xd800 &&
      character.charCodeAt(0) <= 0xdfff
        ? "\ufffd"
        : character,
    )
    .join("");
}

function truncateUtf16WithoutSplittingPair(
  value: string,
  maximum: number,
): string {
  const truncated = value.slice(0, maximum);
  if (!truncated) return truncated;
  const lastCodeUnit = truncated.charCodeAt(truncated.length - 1);
  return lastCodeUnit >= 0xd800 && lastCodeUnit <= 0xdbff
    ? truncated.slice(0, -1)
    : truncated;
}

export const DOCUMENT_MEDIA_ALLOWED_ATTRS = new Set([
  "schemaVersion",
  "kind",
  "url",
  "title",
  "poster",
  "width",
  "height",
  "textAlign",
  // Shared native-document identity/provenance attributes.
  "block_id",
  "source",
]);

export function normalizeDocumentMediaKind(
  value: unknown,
): DocumentMediaKind | null {
  if (typeof value !== "string") return null;
  const normalized = value.trim().toLowerCase();
  if (DOCUMENT_MEDIA_KIND_SET.has(normalized)) {
    return normalized as DocumentMediaKind;
  }
  if (
    normalized === "embed" ||
    normalized === "online" ||
    normalized === "youtube" ||
    normalized === "vimeo"
  ) {
    return "iframe";
  }
  if (normalized === "movie" || normalized === "mp4") return "video";
  return null;
}

// The pasted-embed-code reader and the YouTube / Vimeo id rules are the
// online-video module's (shared with the presentation editor).
const extractIframeSrc = onlineVideoEmbedCodeSrc;

function hostMatches(host: string, allowed: string): boolean {
  return host === allowed || host.endsWith(`.${allowed}`);
}

function supportedIframeHost(host: string): boolean {
  const normalized = host.toLowerCase();
  return DOCUMENT_MEDIA_IFRAME_HOSTS.some((allowed) =>
    hostMatches(normalized, allowed),
  );
}

function pathSegment(pathname: string, index: number): string {
  return pathname.split("/").filter(Boolean)[index] ?? "";
}

function normalizeYouTubeUrl(url: URL): string | null {
  const id = youTubeVideoId(url);
  return id ? `https://www.youtube-nocookie.com/embed/${id}` : null;
}

function normalizeVimeoUrl(url: URL): string | null {
  const id = vimeoVideoId(url);
  return id ? `https://player.vimeo.com/video/${id}` : null;
}

function normalizeTedUrl(url: URL): string | null {
  const host = url.hostname.toLowerCase().replace(/^www\./, "");
  if (!hostMatches(host, "ted.com")) return null;
  const first = pathSegment(url.pathname, 0).toLowerCase();
  const slug = pathSegment(url.pathname, 1);
  return first === "talks" && /^[A-Za-z0-9_-]{1,180}$/.test(slug)
    ? `https://embed.ted.com/talks/${slug}`
    : null;
}

function isProviderEmbedUrl(url: URL): boolean {
  const host = url.hostname.toLowerCase();
  const path = url.pathname.toLowerCase();
  if (hostMatches(host, "slideshare.net")) {
    return path.includes("/slideshow/embed_code/");
  }
  if (
    hostMatches(host, "microsoftstream.com") ||
    hostMatches(host, "stream.office.com")
  ) {
    return path.includes("/embed/");
  }
  if (hostMatches(host, "sharepoint.com")) {
    return (
      path.endsWith("/embed.aspx") ||
      url.searchParams.get("action")?.toLowerCase() === "embedview"
    );
  }
  if (hostMatches(host, "onedrive.live.com")) {
    return path.endsWith("/embed") && url.searchParams.has("resid");
  }
  return false;
}

function parseDocumentMediaUrl(value: string): URL | null {
  if (
    !value ||
    value.length > DOCUMENT_MEDIA_MAX_URL_CHARS ||
    DOCUMENT_MEDIA_URL_FORBIDDEN_RE.test(value) ||
    hasUnpairedSurrogate(value)
  ) {
    return null;
  }
  try {
    const parsed = new URL(value);
    return parsed.protocol === "https:" &&
      parsed.hostname &&
      !parsed.username &&
      !parsed.password
      ? parsed
      : null;
  } catch {
    return null;
  }
}

export function normalizeDocumentMediaUrl(
  kind: DocumentMediaKind,
  value: unknown,
  options: NormalizeDocumentMediaOptions = {},
): string | null {
  if (typeof value !== "string") return null;
  const extracted = options.allowEmbedCode
    ? (extractIframeSrc(value) ?? value.trim())
    : value.trim();
  const parsed = parseDocumentMediaUrl(extracted);
  if (!parsed) return null;

  if (kind === "video") {
    return sanitizeEmbedUrl(parsed.toString(), { kind: "video" });
  }
  if (!supportedIframeHost(parsed.hostname)) return null;

  const host = parsed.hostname.toLowerCase();
  const providerUrl =
    hostMatches(host, "youtube.com") ||
    hostMatches(host, "youtube-nocookie.com") ||
    hostMatches(host, "youtu.be")
      ? normalizeYouTubeUrl(parsed)
      : hostMatches(host, "vimeo.com")
        ? normalizeVimeoUrl(parsed)
        : hostMatches(host, "ted.com")
          ? normalizeTedUrl(parsed)
          : isProviderEmbedUrl(parsed)
            ? parsed.toString()
            : null;
  return providerUrl
    ? sanitizeEmbedUrl(providerUrl, {
        kind: "iframe",
        // The document-level provider/path allowlist above is stricter than
        // presentation's generic-https opt-in; this call owns scheme/control
        // normalization and never widens the accepted host universe.
        allowGenericIframe: true,
      })
    : null;
}

export function normalizeDocumentMediaTitle(value: unknown): string {
  if (typeof value !== "string") return "";
  const normalized = replaceUnpairedSurrogates(value)
    .replace(DOCUMENT_MEDIA_CONTROL_RE, " ")
    .replace(/\s+/g, " ")
    .trim();
  return truncateUtf16WithoutSplittingPair(
    normalized,
    DOCUMENT_MEDIA_MAX_TITLE_CHARS,
  );
}

function normalizeDimension(
  value: unknown,
  fallback: number,
  minimum: number,
  maximum: number,
): number {
  const numeric = Number(value);
  if (!Number.isFinite(numeric)) return fallback;
  return Math.min(Math.max(Math.round(numeric), minimum), maximum);
}

export function parseDocumentMediaDimension(
  value: unknown,
  fallback: number,
  minimum: number,
  maximum: number,
): number {
  if (value == null || value === "") return fallback;
  return normalizeDimension(value, fallback, minimum, maximum);
}

export function normalizeDocumentMediaAttrs(
  value: unknown,
  options: NormalizeDocumentMediaOptions = {},
): DocumentMediaAttrs | null {
  if (!value || typeof value !== "object" || Array.isArray(value)) return null;
  const attrs = value as Record<string, unknown>;
  const kind = normalizeDocumentMediaKind(attrs.kind);
  if (!kind) return null;
  const url = normalizeDocumentMediaUrl(kind, attrs.url, options);
  if (!url) return null;
  const poster =
    kind === "video" && typeof attrs.poster === "string" && attrs.poster.trim()
      ? normalizeDocumentMediaUrl("video", attrs.poster)
      : null;
  if (
    kind === "video" &&
    attrs.poster != null &&
    attrs.poster !== "" &&
    !poster
  ) {
    return null;
  }
  return {
    schemaVersion: DOCUMENT_MEDIA_SCHEMA_VERSION,
    kind,
    url,
    title: normalizeDocumentMediaTitle(attrs.title),
    poster,
    width: parseDocumentMediaDimension(
      attrs.width,
      DOCUMENT_MEDIA_DEFAULT_WIDTH,
      DOCUMENT_MEDIA_MIN_WIDTH,
      DOCUMENT_MEDIA_MAX_WIDTH,
    ),
    height: parseDocumentMediaDimension(
      attrs.height,
      DOCUMENT_MEDIA_DEFAULT_HEIGHT,
      DOCUMENT_MEDIA_MIN_HEIGHT,
      DOCUMENT_MEDIA_MAX_HEIGHT,
    ),
    textAlign: normalizeDocumentMediaAlignment(attrs.textAlign),
  };
}

function persistedDimensionIsValid(
  value: unknown,
  minimum: number,
  maximum: number,
): boolean {
  return (
    typeof value === "number" &&
    Number.isInteger(value) &&
    value >= minimum &&
    value <= maximum
  );
}

/** Independent save/API-side contract; unlike the renderer, never repairs. */
export function isPersistableDocumentMediaAttrs(value: unknown): boolean {
  if (!value || typeof value !== "object" || Array.isArray(value)) return false;
  const attrs = value as Record<string, unknown>;
  if (
    Object.keys(attrs).some((key) => !DOCUMENT_MEDIA_ALLOWED_ATTRS.has(key))
  ) {
    return false;
  }
  if (
    attrs.schemaVersion !== DOCUMENT_MEDIA_SCHEMA_VERSION ||
    !DOCUMENT_MEDIA_KIND_SET.has(String(attrs.kind)) ||
    typeof attrs.url !== "string" ||
    normalizeDocumentMediaUrl(attrs.kind as DocumentMediaKind, attrs.url) ===
      null
  ) {
    return false;
  }
  if (
    typeof attrs.title !== "string" ||
    attrs.title.length > DOCUMENT_MEDIA_MAX_TITLE_CHARS ||
    DOCUMENT_MEDIA_CONTROL_RE.test(attrs.title) ||
    hasUnpairedSurrogate(attrs.title)
  ) {
    return false;
  }
  if (
    !persistedDimensionIsValid(
      attrs.width,
      DOCUMENT_MEDIA_MIN_WIDTH,
      DOCUMENT_MEDIA_MAX_WIDTH,
    ) ||
    !persistedDimensionIsValid(
      attrs.height,
      DOCUMENT_MEDIA_MIN_HEIGHT,
      DOCUMENT_MEDIA_MAX_HEIGHT,
    )
  ) {
    return false;
  }
  if (attrs.kind === "iframe" && attrs.poster != null && attrs.poster !== "") {
    return false;
  }
  if (
    attrs.textAlign != null &&
    normalizeDocumentMediaAlignment(attrs.textAlign) === null
  ) {
    return false;
  }
  if (
    attrs.kind === "video" &&
    attrs.poster != null &&
    (typeof attrs.poster !== "string" ||
      normalizeDocumentMediaUrl("video", attrs.poster) === null)
  ) {
    return false;
  }
  if (
    attrs.block_id != null &&
    (typeof attrs.block_id !== "string" ||
      !attrs.block_id ||
      attrs.block_id.length > 128 ||
      DOCUMENT_MEDIA_CONTROL_RE.test(attrs.block_id) ||
      hasUnpairedSurrogate(attrs.block_id))
  ) {
    return false;
  }
  return true;
}

export function documentMediaDisplayLabel(attrs: DocumentMediaAttrs): string {
  if (attrs.title) return attrs.title;
  try {
    return new URL(attrs.url).hostname;
  } catch {
    return "Online video";
  }
}
