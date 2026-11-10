/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Presentation media / embed blocks (PowerPoint parity subset).
 *
 * A slide can embed THREE kinds of remote media:
 *  - `iframe`  — an embeddable web player (YouTube / Vimeo) or a generic https
 *                page, rendered inside a strictly-sandboxed `<iframe>`.
 *  - `video`   — a direct video file URL, rendered with native `<video controls>`.
 *  - `audio`   — a direct audio file URL, rendered with native `<audio controls>`.
 *
 * Pure, side-effect-free helpers shared by the presentation schema, the editor,
 * and the present-mode overlay so the allowed kinds, the URL security policy,
 * and the default geometry live in exactly one place. Keep this module
 * dependency-free (no React / no DOM) so it can be unit tested in isolation and
 * imported from the zod schema. Mirrors the conventions of
 * `presentation-shapes.ts` and `presentation-block-fx.ts`: a frozen allowed
 * list, a type guard, a lenient normalizer that collapses unknown/stale values
 * to a safe default, and a single URL sanitizer that is the ONE security gate.
 *
 * SECURITY: the URL is NEVER injected as innerHTML — callers set it as the
 * `src` attribute only (the renderer enforces this). `sanitizeEmbedUrl` is the
 * single remote gate: it accepts https URLs only, rejects `javascript:` /
 * `file:` / `blob:` / protocol-relative / scheme-less references, strips
 * embedded control characters (a classic `java\tscript:` bypass), and — for
 * `iframe` embeds — additionally restricts the host to a small allowlist of
 * known embed providers unless the caller opts into a generic-https iframe.
 * The only data URL exception is a separately validated, 1.5 MB bounded local
 * audio payload for a scratch-deck native media export.
 *
 * Geometry is expressed in EMUs (English Metric Units), matching the existing
 * presentation geometry (`left`/`top`/`width`/`height` on every block).
 */

import { parseOnlineVideo, type OnlineVideo } from "./online-video";

/** A 4 : 3-ish default slide canvas, mirroring PresentationWindow defaults. */
const DEFAULT_SLIDE_WIDTH_EMU = 9_144_000;
const DEFAULT_SLIDE_HEIGHT_EMU = 5_143_500;
export const MAX_PRESENTATION_INLINE_AUDIO_BYTES = 1_500_000;
export const MAX_PRESENTATION_INLINE_AUDIO_DATA_URL_CHARS =
  Math.ceil((MAX_PRESENTATION_INLINE_AUDIO_BYTES * 4) / 3) + 128;

/** Strict schema/render guard for inline audio owned by a scratch deck. */
export function isPersistablePresentationAudioDataUrl(
  value: unknown,
): value is string {
  if (
    typeof value !== "string" ||
    value.length > MAX_PRESENTATION_INLINE_AUDIO_DATA_URL_CHARS
  ) {
    return false;
  }
  const match = value.match(
    /^data:(audio\/(?:aac|mp4|mpeg|ogg|wav|x-wav));base64,([A-Za-z0-9+/]+={0,2})$/,
  );
  if (!match) return false;
  try {
    const decoded = atob(match[2]);
    if (
      decoded.length <= 0 ||
      decoded.length > MAX_PRESENTATION_INLINE_AUDIO_BYTES
    ) {
      return false;
    }
    const bytes = Uint8Array.from(
      decoded.slice(0, 16),
      (character) => character.charCodeAt(0),
    );
    const mime = match[1];
    if (mime === "audio/wav" || mime === "audio/x-wav") {
      return (
        String.fromCharCode(...bytes.slice(0, 4)) === "RIFF" &&
        String.fromCharCode(...bytes.slice(8, 12)) === "WAVE"
      );
    }
    if (mime === "audio/mp4") {
      return String.fromCharCode(...bytes.slice(4, 8)) === "ftyp";
    }
    if (mime === "audio/mpeg") {
      return (
        String.fromCharCode(...bytes.slice(0, 3)) === "ID3" ||
        (bytes[0] === 0xff && (bytes[1] ?? 0) >= 0xe0)
      );
    }
    if (mime === "audio/ogg") {
      return String.fromCharCode(...bytes.slice(0, 4)) === "OggS";
    }
    return bytes[0] === 0xff && ((bytes[1] ?? 0) & 0xf6) === 0xf0;
  } catch {
    return false;
  }
}

/** The supported embed kinds. `iframe` is the implicit default. */
export const EMBED_KINDS = ["iframe", "video", "audio"] as const;

export type EmbedKind = (typeof EMBED_KINDS)[number];

/** The kind used when a block has no explicit (or invalid) value. */
export const DEFAULT_EMBED_KIND: EmbedKind = "iframe";

export const PRESENTATION_MEDIA_PLAYBACK_STARTS = [
  "click",
  "automatic",
  "sequence",
] as const;
export type PresentationMediaPlaybackStart =
  (typeof PRESENTATION_MEDIA_PLAYBACK_STARTS)[number];

export const PRESENTATION_MEDIA_VOLUMES = [
  "mute",
  "low",
  "medium",
  "high",
] as const;
export type PresentationMediaVolume =
  (typeof PRESENTATION_MEDIA_VOLUMES)[number];

/** Playback > Add Bookmark: a name unique on its clip and a time (ms) from
 *  the clip's start (`p14:bmk`, PPT-A157). */
export interface PresentationMediaBookmark {
  readonly name: string;
  readonly timeMs: number;
}

/** The API's bounds (`_PRESENTATION_MEDIA_BOOKMARKS_MAX`, name length). */
export const MAX_PRESENTATION_MEDIA_BOOKMARKS = 200;
export const MAX_PRESENTATION_MEDIA_BOOKMARK_NAME_CHARS = 255;

export interface PresentationMediaPlaybackSettings {
  start: PresentationMediaPlaybackStart;
  volume: PresentationMediaVolume;
  loop: boolean;
  rewindAfterPlaying: boolean;
  playAcrossSlides: boolean;
  hideDuringShow: boolean;
  playFullScreen: boolean;
  trimStartMs: number;
  trimEndMs: number | null;
  fadeInMs: number;
  fadeOutMs: number;
  /** In time order; empty when the clip has none. */
  bookmarks: readonly PresentationMediaBookmark[];
}

export type PersistedPresentationMediaPlaybackSettings = Partial<
  PresentationMediaPlaybackSettings
>;

export const DEFAULT_PRESENTATION_MEDIA_PLAYBACK_SETTINGS = Object.freeze({
  start: "click" as const,
  volume: "high" as const,
  loop: false,
  rewindAfterPlaying: false,
  playAcrossSlides: false,
  hideDuringShow: false,
  playFullScreen: false,
  trimStartMs: 0,
  trimEndMs: null,
  fadeInMs: 0,
  fadeOutMs: 0,
  bookmarks: Object.freeze([]) as readonly PresentationMediaBookmark[],
} satisfies PresentationMediaPlaybackSettings);

/** One day is a generous deterministic ceiling for authored clip timing. */
export const MAX_PRESENTATION_MEDIA_TIME_MS = 24 * 60 * 60 * 1_000;
/** Fade duration is deliberately tighter than clip duration. */
export const MAX_PRESENTATION_MEDIA_FADE_MS = 10 * 60 * 1_000;

const PLAYBACK_START_SET: ReadonlySet<string> = new Set(
  PRESENTATION_MEDIA_PLAYBACK_STARTS,
);
const MEDIA_VOLUME_SET: ReadonlySet<string> = new Set(
  PRESENTATION_MEDIA_VOLUMES,
);

function playbackRecord(value: unknown): Record<string, unknown> {
  return value && typeof value === "object" && !Array.isArray(value)
    ? (value as Record<string, unknown>)
    : {};
}

/** The exact bookmark list, or [] when the value is not one the file can hold. */
function resolvePresentationMediaBookmarks(
  value: unknown,
): readonly PresentationMediaBookmark[] {
  if (!Array.isArray(value) || value.length > MAX_PRESENTATION_MEDIA_BOOKMARKS) {
    return DEFAULT_PRESENTATION_MEDIA_PLAYBACK_SETTINGS.bookmarks;
  }
  const names = new Set<string>();
  const bookmarks: PresentationMediaBookmark[] = [];
  for (const raw of value) {
    const record = playbackRecord(raw);
    const name = record.name;
    const timeMs = record.timeMs;
    if (
      Object.keys(record).length !== 2 ||
      typeof name !== "string" ||
      !name.trim() ||
      name.length > MAX_PRESENTATION_MEDIA_BOOKMARK_NAME_CHARS ||
      names.has(name) ||
      !Number.isInteger(timeMs) ||
      (timeMs as number) < 0 ||
      (timeMs as number) > MAX_PRESENTATION_MEDIA_TIME_MS
    ) {
      return DEFAULT_PRESENTATION_MEDIA_PLAYBACK_SETTINGS.bookmarks;
    }
    names.add(name);
    bookmarks.push(Object.freeze({ name, timeMs: timeMs as number }));
  }
  return Object.freeze(bookmarks);
}

function boundedPlaybackMs(
  value: unknown,
  maximum: number,
  fallback: number,
): number {
  return Number.isInteger(value) &&
    (value as number) >= 0 &&
    (value as number) <= maximum
    ? (value as number)
    : fallback;
}

/**
 * Resolve a persisted media playback atom into a total, bounded runtime value.
 * Iframes intentionally receive defaults: online-player behavior stays under
 * the provider's own controls and is never synthesized by the deck model.
 */
export function resolvePresentationMediaPlaybackSettings(
  value: unknown,
  kind: EmbedKind,
): PresentationMediaPlaybackSettings {
  if (kind === "iframe") return DEFAULT_PRESENTATION_MEDIA_PLAYBACK_SETTINGS;
  const record = playbackRecord(value);
  const start =
    typeof record.start === "string" && PLAYBACK_START_SET.has(record.start)
      ? (record.start as PresentationMediaPlaybackStart)
      : DEFAULT_PRESENTATION_MEDIA_PLAYBACK_SETTINGS.start;
  const volume =
    typeof record.volume === "string" && MEDIA_VOLUME_SET.has(record.volume)
      ? (record.volume as PresentationMediaVolume)
      : DEFAULT_PRESENTATION_MEDIA_PLAYBACK_SETTINGS.volume;
  const trimStartMs = boundedPlaybackMs(
    record.trimStartMs,
    MAX_PRESENTATION_MEDIA_TIME_MS,
    0,
  );
  const rawTrimEnd = boundedPlaybackMs(
    record.trimEndMs,
    MAX_PRESENTATION_MEDIA_TIME_MS,
    -1,
  );
  return Object.freeze({
    start,
    volume,
    loop: record.loop === true,
    rewindAfterPlaying: record.rewindAfterPlaying === true,
    playAcrossSlides: record.playAcrossSlides === true,
    hideDuringShow: record.hideDuringShow === true,
    playFullScreen: kind === "video" && record.playFullScreen === true,
    trimStartMs,
    trimEndMs: rawTrimEnd > trimStartMs ? rawTrimEnd : null,
    fadeInMs: boundedPlaybackMs(
      record.fadeInMs,
      MAX_PRESENTATION_MEDIA_FADE_MS,
      0,
    ),
    fadeOutMs: boundedPlaybackMs(
      record.fadeOutMs,
      MAX_PRESENTATION_MEDIA_FADE_MS,
      0,
    ),
    bookmarks: resolvePresentationMediaBookmarks(record.bookmarks),
  });
}

/**
 * Canonical persistence form: omit every default and reject provider-owned
 * iframe settings. `null` means the block needs no playback atom at all.
 */
export function normalizePresentationMediaPlaybackSettings(
  value: unknown,
  kind: EmbedKind,
): PersistedPresentationMediaPlaybackSettings | null {
  if (kind === "iframe") return null;
  const resolved = resolvePresentationMediaPlaybackSettings(value, kind);
  const defaults = DEFAULT_PRESENTATION_MEDIA_PLAYBACK_SETTINGS;
  const persisted: PersistedPresentationMediaPlaybackSettings = {};
  for (const key of Object.keys(defaults) as Array<
    keyof PresentationMediaPlaybackSettings
  >) {
    if (key === "bookmarks") {
      // A clip with no bookmarks states none (the save planner removes the
      // file's list when an earlier state had one).
      if (resolved.bookmarks.length > 0) {
        persisted.bookmarks = resolved.bookmarks.map((bookmark) => ({ ...bookmark }));
      }
      continue;
    }
    if (!Object.is(resolved[key], defaults[key])) {
      (persisted as Record<string, unknown>)[key] = resolved[key];
    }
  }
  return Object.keys(persisted).length > 0 ? Object.freeze(persisted) : null;
}

/** One bounded writer edge for future ribbon/inspector commands. */
export function patchPresentationMediaPlaybackSettings(
  current: unknown,
  patch: PersistedPresentationMediaPlaybackSettings,
  kind: EmbedKind,
): PersistedPresentationMediaPlaybackSettings | null {
  if (kind === "iframe") return null;
  return normalizePresentationMediaPlaybackSettings(
    { ...resolvePresentationMediaPlaybackSettings(current, kind), ...patch },
    kind,
  );
}

/** PowerPoint for the web's documented local-video ceiling (256 MB). */
export const MAX_PRESENTATION_VIDEO_FILE_BYTES = 256 * 1024 * 1024;

export interface AssetBackedPresentationVideoSource {
  readonly sourceKind: "asset";
  readonly assetId: string;
  readonly mimeType: "video/mp4";
  readonly sizeBytes: number;
}

export function assetBackedPresentationVideoSource(
  input: Record<string, unknown>,
): AssetBackedPresentationVideoSource | null {
  if (
    input.kind !== "video" ||
    input.sourceKind !== "asset" ||
    typeof input.assetId !== "string" ||
    !input.assetId ||
    input.assetId !== input.assetId.trim() ||
    input.assetId.length > 200 ||
    input.mimeType !== "video/mp4" ||
    !Number.isInteger(input.sizeBytes) ||
    (input.sizeBytes as number) <= 0 ||
    (input.sizeBytes as number) > MAX_PRESENTATION_VIDEO_FILE_BYTES ||
    (typeof input.url === "string" && input.url !== "")
  ) {
    return null;
  }
  return Object.freeze({
    sourceKind: "asset" as const,
    assetId: input.assetId,
    mimeType: "video/mp4" as const,
    sizeBytes: input.sizeBytes as number,
  });
}

/** The audio formats a stored clip may be (presentation-audio-file.ts). */
export const PRESENTATION_AUDIO_ASSET_MIME_TYPES = [
  "audio/aac",
  "audio/mp4",
  "audio/mpeg",
  "audio/ogg",
  "audio/wav",
] as const;
export type PresentationAudioAssetMimeType =
  (typeof PRESENTATION_AUDIO_ASSET_MIME_TYPES)[number];

/**
 * An audio clip stored beside the deck (Insert > Audio > Audio on My PC):
 * the DataItem is the clip, as a local video's is. It has no size of its own
 * past the deck's media total (lib/document-size-limits); a clip written
 * before as an inline data: URL stays one.
 */
export interface AssetBackedPresentationAudioSource {
  readonly sourceKind: "asset";
  readonly assetId: string;
  readonly mimeType: PresentationAudioAssetMimeType;
  readonly sizeBytes: number;
}

function isAssetId(value: unknown): value is string {
  return (
    typeof value === "string" &&
    value.length > 0 &&
    value === value.trim() &&
    value.length <= 200 &&
    !CONTROL_CHAR_RE.test(value)
  );
}

export function assetBackedPresentationAudioSource(
  input: Record<string, unknown>,
): AssetBackedPresentationAudioSource | null {
  if (
    input.kind !== "audio" ||
    input.sourceKind !== "asset" ||
    !isAssetId(input.assetId) ||
    !(PRESENTATION_AUDIO_ASSET_MIME_TYPES as readonly unknown[]).includes(
      input.mimeType,
    ) ||
    !Number.isSafeInteger(input.sizeBytes) ||
    (input.sizeBytes as number) <= 0 ||
    (typeof input.url === "string" && input.url !== "")
  ) {
    return null;
  }
  return Object.freeze({
    sourceKind: "asset" as const,
    assetId: input.assetId,
    mimeType: input.mimeType as PresentationAudioAssetMimeType,
    sizeBytes: input.sizeBytes as number,
  });
}

/** A local video or audio clip stored beside the deck, whichever it is. */
export type AssetBackedPresentationMediaSource =
  | (AssetBackedPresentationVideoSource & { readonly kind: "video" })
  | (AssetBackedPresentationAudioSource & { readonly kind: "audio" });

export function assetBackedPresentationMediaSource(
  input: Record<string, unknown>,
): AssetBackedPresentationMediaSource | null {
  const video = assetBackedPresentationVideoSource(input);
  if (video) return { ...video, kind: "video" };
  const audio = assetBackedPresentationAudioSource(input);
  return audio ? { ...audio, kind: "audio" } : null;
}

/**
 * What the deck's stored media already weigh: every distinct local video,
 * audio clip and transition sound asset, as the deck's media total counts
 * them (lib/document-size-limits, the API's document_size_limits).
 */
export function presentationStoredMediaBytes(
  slides: readonly {
    readonly blocks?: readonly unknown[];
    readonly transitionSound?: unknown;
  }[],
  options: { readonly audioOnly?: boolean } = {},
): number {
  const sizes = new Map<string, number>();
  for (const slide of slides) {
    const sound = slide.transitionSound;
    if (sound && typeof sound === "object" && !Array.isArray(sound)) {
      const record = sound as Record<string, unknown>;
      if (
        record.sourceKind === "asset" &&
        isAssetId(record.assetId) &&
        Number.isSafeInteger(record.sizeBytes)
      ) {
        sizes.set(record.assetId, record.sizeBytes as number);
      }
    }
    for (const block of slide.blocks ?? []) {
      if (!block || typeof block !== "object") continue;
      const media = assetBackedPresentationMediaSource(
        block as Record<string, unknown>,
      );
      if (media && (!options.audioOnly || media.kind === "audio") && !sizes.has(media.assetId)) {
        sizes.set(media.assetId, media.sizeBytes);
      }
    }
  }
  let total = 0;
  for (const size of sizes.values()) total += size;
  return total;
}

const EMBED_KIND_SET: ReadonlySet<string> = new Set(EMBED_KINDS);

/**
 * The host allowlist for `iframe` embeds. A host matches when it equals one of
 * these OR is a sub-domain of one (so `www.youtube.com` and
 * `player.vimeo.com` are covered). Lower-cased, no port. These are the known
 * privacy-respecting / embed-friendly players; anything else must go through
 * the generic-https path (opt-in) so an author can't silently frame an
 * arbitrary origin that may set cookies / run trackers under our roof.
 */
export const EMBED_IFRAME_ALLOWED_HOSTS = [
  "youtube.com",
  "youtube-nocookie.com",
  "youtu.be",
  "vimeo.com",
  "player.vimeo.com",
] as const;

const EMBED_IFRAME_ALLOWED_HOST_SET: ReadonlySet<string> = new Set(
  EMBED_IFRAME_ALLOWED_HOSTS,
);

/**
 * The `sandbox` token set applied to EVERY embed iframe. Deliberately minimal:
 *  - `allow-scripts`       — players need JS to run.
 *  - `allow-same-origin`   — players read their own origin's storage/config.
 *  - `allow-presentation`  — lets the framed player request fullscreen / cast.
 * Notably absent: `allow-popups`, `allow-top-navigation`, `allow-forms`,
 * `allow-modals` — so a framed page can neither navigate our top window, spawn
 * pop-unders, nor submit forms on our behalf. Exported so the renderer and its
 * test assert the SAME string (one source of truth).
 */
export const EMBED_IFRAME_SANDBOX =
  "allow-scripts allow-same-origin allow-presentation";

/**
 * The `referrerPolicy` applied to every embed iframe — send only the origin
 * (not the full deck URL / query) to the framed player.
 */
export const EMBED_IFRAME_REFERRER_POLICY = "strict-origin-when-cross-origin";

/**
 * Matches any ASCII control character (U+0000–U+001F or U+007F). Browsers strip
 * embedded control chars (newline/tab/etc.) when parsing a URL, which would let
 * `java\tscript:…` slip past a naive scheme check; rejecting them up front
 * closes that bypass. Mirrors the guard in `presentation-block-fx.ts`.
 */
const CONTROL_CHAR_RE = /[\u0000-\u001f\u007f]/;

/**
 * Type guard: is `value` exactly one of the allowed embed identifiers?
 * Exact, case-sensitive — use {@link normalizeEmbedKind} for lenient input.
 */
export function isEmbedKind(value: unknown): value is EmbedKind {
  return typeof value === "string" && EMBED_KIND_SET.has(value);
}

/**
 * Normalize an arbitrary persisted value into an allowed embed kind. Unknown,
 * empty, or non-string values collapse to {@link DEFAULT_EMBED_KIND} so a stale
 * or hand-edited deck can never inject an unsupported kind. Common synonyms
 * (`youtube`/`vimeo`/`embed` → `iframe`, `movie`/`mp4` → `video`, `sound`/`mp3`
 * → `audio`) are accepted. Case-insensitive; trims surrounding whitespace.
 */
export function normalizeEmbedKind(value: unknown): EmbedKind {
  if (typeof value !== "string") return DEFAULT_EMBED_KIND;
  const candidate = value.trim().toLowerCase();
  if (
    candidate === "youtube" ||
    candidate === "vimeo" ||
    candidate === "embed" ||
    candidate === "frame"
  ) {
    return "iframe";
  }
  if (candidate === "movie" || candidate === "mp4" || candidate === "webm") {
    return "video";
  }
  if (candidate === "sound" || candidate === "mp3" || candidate === "wav") {
    return "audio";
  }
  return isEmbedKind(candidate) ? candidate : DEFAULT_EMBED_KIND;
}

/**
 * Parse a URL string with the WHATWG `URL` parser. Returns the parsed `URL` or
 * `null` for anything that fails to parse (the parser itself rejects most
 * garbage and resolves the scheme/host for us). `URL` is available in the
 * browser, in Node, and in the test (jsdom) runtime, so this stays dependency-
 * free. A scheme-less / relative value (no base) throws and yields `null`.
 */
function parseAbsoluteUrl(raw: string): URL | null {
  try {
    return new URL(raw);
  } catch {
    return null;
  }
}

/**
 * Lower-cased registrable-ish host check: does `host` equal an allowlisted host
 * or end with `.<allowlisted host>` (a sub-domain)? `host` is already the
 * parser's `hostname` (no port, no userinfo), so this can't be fooled by an
 * `@`-embedded authority.
 */
function isAllowlistedIframeHost(host: string): boolean {
  const lower = host.toLowerCase();
  if (EMBED_IFRAME_ALLOWED_HOST_SET.has(lower)) return true;
  for (const allowed of EMBED_IFRAME_ALLOWED_HOSTS) {
    if (lower.endsWith(`.${allowed}`)) return true;
  }
  return false;
}

/**
 * Whether a video's address is a known player page (YouTube, Vimeo) rather
 * than a clip a `<video>` can play: PowerPoint's Online Videos.
 */
export function isOnlinePlayerPage(url: unknown): boolean {
  if (typeof url !== "string") return false;
  try {
    const parsed = new URL(url);
    return parsed.protocol === "https:" && isAllowlistedIframeHost(parsed.hostname);
  } catch {
    return false;
  }
}

/**
 * The online video (YouTube, Vimeo) a media block plays, or null: an
 * Insert > Video > Online Video block (`iframe`) or an imported PowerPoint
 * online video (`video` linked to the player's address). Such a block plays
 * in the provider's privacy-enhanced player frame (lib/online-video), never
 * in a `<video>` element, and its address is what a .pptx links to.
 */
export function presentationOnlineVideo(block: {
  readonly kind?: unknown;
  readonly url?: unknown;
  readonly sourceKind?: unknown;
}): OnlineVideo | null {
  if (block.sourceKind === "asset") return null;
  const kind = normalizeEmbedKind(block.kind);
  return kind === "audio" ? null : parseOnlineVideo(block.url);
}

export interface SanitizeEmbedUrlOptions {
  /** The embed kind the URL is for — gates the host-allowlist rule. */
  kind: EmbedKind;
  /**
   * When `true`, an `iframe` URL on ANY https host is allowed (still strictly
   * sandboxed). When `false`/omitted, an `iframe` URL must be on an allowlisted
   * provider host (see {@link EMBED_IFRAME_ALLOWED_HOSTS}). Has no effect on
   * `video`/`audio` (which are always any-https direct media files).
   */
  allowGenericIframe?: boolean;
}

/**
 * Validate and normalize an embed URL against the security policy. Returns the
 * normalized https URL when it is safe to render, or `null` when it is missing,
 * malformed, or violates the policy. This is the SINGLE security gate; the
 * schema, the editor, and the renderer all funnel through it so an unsafe value
 * can never reach an `<iframe>`/`<video>`/`<audio>` `src`.
 *
 * Rules (in order):
 *  - Non-string / empty / whitespace-only → `null`.
 *  - Any embedded ASCII control character → `null` (anti-`java\nscript:`).
 *  - Must parse as an absolute URL (scheme-less / protocol-relative `//host`
 *    fails to parse here and is rejected — present mode must never resolve a
 *    media src against an unexpected base).
 *  - Scheme MUST be exactly `https:` — this alone rejects `javascript:`,
 *    `data:`, `file:`, `blob:`, `vbscript:`, and plain `http:`.
 *  - For `kind: "iframe"` WITHOUT `allowGenericIframe`, the host must be on the
 *    {@link EMBED_IFRAME_ALLOWED_HOSTS} allowlist. `video`/`audio` (and a
 *    generic-iframe opt-in) accept any https host.
 *
 * The returned value is `url.toString()` (parser-normalized) so a persisted URL
 * is canonical and stable.
 */
export function sanitizeEmbedUrl(
  value: unknown,
  options: SanitizeEmbedUrlOptions,
): string | null {
  if (typeof value !== "string") return null;
  const trimmed = value.trim();
  if (!trimmed) return null;
  if (CONTROL_CHAR_RE.test(trimmed)) return null;
  const parsed = parseAbsoluteUrl(trimmed);
  if (!parsed) return null;
  // https ONLY — the loudest, simplest gate. Rejects javascript:/data:/file:/
  // blob:/http: in one check.
  if (parsed.protocol !== "https:") return null;
  if (options.kind === "iframe" && !options.allowGenericIframe) {
    if (!isAllowlistedIframeHost(parsed.hostname)) return null;
  }
  return parsed.toString();
}

/**
 * Convenience: is `value` an acceptable embed URL for `options`? Thin wrapper
 * over {@link sanitizeEmbedUrl} for places that only need a boolean (e.g. the
 * zod refine).
 */
export function isSafeEmbedUrl(
  value: unknown,
  options: SanitizeEmbedUrlOptions,
): boolean {
  return sanitizeEmbedUrl(value, options) !== null;
}

/**
 * Schema-side URL guard: accept a URL that is safe for AT LEAST ONE embed kind.
 * The persistence boundary does not know the block's kind ordering guarantees,
 * so it admits a value that is a valid https media URL (video/audio) OR a valid
 * iframe URL (allowlisted host, OR generic https — the renderer always sandboxes
 * generic iframes, so a generic https iframe URL is structurally acceptable at
 * rest). The render-time resolver ({@link resolveEmbed}) re-checks against the
 * block's ACTUAL kind and drops a URL that doesn't fit, so this looser gate can
 * never cause an unsafe render — it only avoids rejecting a legitimately-typed
 * deck whose kind/url pairing is re-validated downstream. It still rejects every
 * dangerous scheme (only https can pass either branch).
 */
export function isPersistableEmbedUrl(value: unknown): boolean {
  return (
    isSafeEmbedUrl(value, { kind: "video" }) ||
    isSafeEmbedUrl(value, { kind: "iframe", allowGenericIframe: true })
  );
}

/**
 * Schema-side URL guard that ALSO admits an empty/whitespace-only string — a
 * DRAFT embed the author just inserted but has not yet filled in. An empty URL
 * is not a security risk: the render-time resolver ({@link resolveEmbed})
 * returns `url: null` for it, so the renderer shows a non-interactive fallback
 * and never mounts an element. Any NON-empty value must still pass {@link
 * isPersistableEmbedUrl} (https-only; dangerous schemes rejected). This lets a
 * freshly-inserted, not-yet-configured embed persist + round-trip without
 * failing whole-deck validation, mirroring how an image block carries a
 * placeholder src at insert time.
 */
export function isPersistableEmbedUrlOrEmpty(value: unknown): boolean {
  if (typeof value === "string" && value.trim() === "") return true;
  return isPersistableEmbedUrl(value);
}

export interface ResolvedEmbed {
  kind: EmbedKind;
  /** A sanitized https URL or strict bounded local-audio data URL. */
  url: string | null;
  /** Sandbox token string for an iframe embed (only meaningful when kind === "iframe"). */
  sandbox: string;
  /** referrerPolicy for an iframe embed. */
  referrerPolicy: string;
}

/**
 * Resolve an embed block's persisted kind + url into a normalized, render-ready
 * descriptor. The kind is normalized first; the URL is then re-sanitized
 * against THAT kind (generic https iframes are permitted at render time because
 * the renderer always applies the strict sandbox). A `null` `url` means the
 * persisted value was unsafe/missing and the renderer should show a fallback
 * rather than a broken/dangerous element. Side-effect free; safe in both the
 * editor and the read-only present overlay.
 */
export function resolveEmbed(input: {
  kind?: unknown;
  url?: unknown;
  src?: unknown;
}): ResolvedEmbed {
  const kind = normalizeEmbedKind(input.kind);
  // Accept either `url` (canonical) or `src` (image-block-style alias) so the
  // normalizer can read whichever a hand-edited / imported deck supplied.
  const rawUrl =
    typeof input.url === "string"
      ? input.url
      : typeof input.src === "string"
        ? input.src
        : undefined;
  const url =
    kind === "audio" && isPersistablePresentationAudioDataUrl(rawUrl)
      ? rawUrl
      : sanitizeEmbedUrl(rawUrl, {
          kind,
          // The renderer always sandboxes iframes, so a generic https iframe is safe
          // to render; the host allowlist is an editor-time nicety, not a render-time
          // invariant. Keeping render permissive (but still https-only + sandboxed)
          // means an allowlisted-host typo doesn't blank a legitimately https embed.
          allowGenericIframe: kind === "iframe",
        });
  return {
    kind,
    url,
    sandbox: EMBED_IFRAME_SANDBOX,
    referrerPolicy: EMBED_IFRAME_REFERRER_POLICY,
  };
}

export interface EmbedGeometry {
  left: number;
  top: number;
  width: number;
  height: number;
}

/**
 * A sensible default rect for a newly inserted embed block, centred-ish on the
 * canvas. Iframe/video get a 16:9-ish landscape box; audio gets a short wide
 * strip (a player bar). Geometry is clamped to the provided slide bounds
 * (falling back to the standard canvas) so an inserted block always lands fully
 * on-slide regardless of the deck's dimensions. Mirrors
 * `defaultShapeGeometry` in `presentation-shapes.ts`.
 */
export function defaultEmbedGeometry(
  kind: EmbedKind,
  bounds?: { width?: number | null; height?: number | null },
): EmbedGeometry {
  const slideWidth =
    typeof bounds?.width === "number" && bounds.width > 0
      ? bounds.width
      : DEFAULT_SLIDE_WIDTH_EMU;
  const slideHeight =
    typeof bounds?.height === "number" && bounds.height > 0
      ? bounds.height
      : DEFAULT_SLIDE_HEIGHT_EMU;
  const isAudio = kind === "audio";
  const width = Math.round(slideWidth * (isAudio ? 0.5 : 0.5));
  const height = Math.round(slideHeight * (isAudio ? 0.08 : 0.46));
  return {
    left: Math.round((slideWidth - width) / 2),
    top: Math.round((slideHeight - height) / 2),
    width,
    height,
  };
}
