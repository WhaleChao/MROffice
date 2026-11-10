/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
export type OnlineVideoProvider = "youtube" | "vimeo";

/**
 * The frame origins an online video plays from. The app CSP's `frame-src`
 * for workspace pages is extended by exactly this list (proxy.ts) and the
 * slide frames nothing else, so the two can never drift apart.
 */
export const ONLINE_VIDEO_FRAME_ORIGINS = [
  "https://www.youtube-nocookie.com",
  "https://player.vimeo.com",
] as const;

/**
 * Sandbox tokens for an online video's player frame. The player needs its
 * scripts and its own origin's storage; `allow-presentation` lets it cast;
 * `allow-popups` + `allow-popups-to-escape-sandbox` let a viewer's click on
 * the player's own "Watch on YouTube" / Vimeo logo (and the provider's
 * "this video can't be played here — watch it on the site" notice) open the
 * provider's page in a normal tab. Never `allow-top-navigation*`,
 * `allow-forms` or `allow-modals`: the player cannot navigate the app,
 * submit forms for it or block it with dialogs. Generic web pages are never
 * framed at all.
 */
export const ONLINE_VIDEO_IFRAME_SANDBOX =
  "allow-scripts allow-same-origin allow-presentation allow-popups allow-popups-to-escape-sandbox";

/** Send only our origin, never the deck's address, to the provider. */
export const ONLINE_VIDEO_IFRAME_REFERRER_POLICY = "strict-origin-when-cross-origin";

/** Permissions delegated to the player (autoplay only after a click to play). */
export function onlineVideoIframeAllow(autoplay: boolean): string {
  return autoplay
    ? "autoplay; encrypted-media; fullscreen; picture-in-picture"
    : "encrypted-media; fullscreen; picture-in-picture";
}

export interface OnlineVideo {
  readonly provider: OnlineVideoProvider;
  /** The provider's video id (YouTube 11-char id, Vimeo number). */
  readonly id: string;
  /** YouTube: where the video starts, in whole seconds (`t=` / `start=`). */
  readonly startSeconds: number | null;
  /** Vimeo: the unlisted video's privacy hash (`vimeo.com/<id>/<hash>`, `h=`). */
  readonly vimeoHash: string | null;
  /** The player address the deck keeps and the .pptx links to. */
  readonly linkUrl: string;
  /** The video's own page on the provider (open in a new tab). */
  readonly watchUrl: string;
}

const MAX_URL_CHARS = 4_096;
const MAX_EMBED_CODE_CHARS = 8_192;
const MAX_START_SECONDS = 24 * 60 * 60;
const URL_FORBIDDEN_RE = /[\s\u0000-\u001f\u007f-\u009f]/;
const YOUTUBE_ID_RE = /^[A-Za-z0-9_-]{6,64}$/;
const VIMEO_ID_RE = /^\d{1,20}$/;
const VIMEO_HASH_RE = /^[0-9A-Za-z]{6,40}$/;

function hostMatches(host: string, allowed: string): boolean {
  return host === allowed || host.endsWith(`.${allowed}`);
}

function pathSegment(pathname: string, index: number): string {
  return pathname.split("/").filter(Boolean)[index] ?? "";
}

/**
 * The `src` of a pasted `<iframe …>` embed code (Share > Embed on YouTube and
 * Vimeo), or null when the value is not one.
 */
export function onlineVideoEmbedCodeSrc(value: string): string | null {
  if (value.length > MAX_EMBED_CODE_CHARS || !/^\s*<iframe\b/i.test(value)) {
    return null;
  }
  const match = value.match(/\bsrc\s*=\s*(["'])([^"'<>]+)\1/i);
  if (!match) return null;
  return match[2].replace(/&amp;/gi, "&");
}

/** An https URL with a host and no credentials, or null. */
export function parseHttpsUrl(value: string): URL | null {
  if (!value || value.length > MAX_URL_CHARS || URL_FORBIDDEN_RE.test(value)) {
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

/** `90`, `90s`, `1m30s`, `1h2m3s` → seconds; null when not a time. */
function parseYouTubeTime(value: string | null): number | null {
  if (!value) return null;
  const trimmed = value.trim().toLowerCase();
  let seconds: number;
  if (/^\d{1,6}s?$/.test(trimmed)) {
    seconds = Number.parseInt(trimmed, 10);
  } else {
    const match = trimmed.match(/^(?:(\d{1,3})h)?(?:(\d{1,4})m)?(?:(\d{1,6})s?)?$/);
    if (!match || (!match[1] && !match[2] && !match[3])) return null;
    seconds =
      Number(match[1] ?? 0) * 3_600 +
      Number(match[2] ?? 0) * 60 +
      Number(match[3] ?? 0);
  }
  return Number.isSafeInteger(seconds) && seconds > 0 && seconds <= MAX_START_SECONDS
    ? seconds
    : null;
}

/**
 * The YouTube video id of a youtube.com / youtu.be / youtube-nocookie.com
 * address, or null. The address shapes Word's Online Video takes (and the
 * Docs editor's document-media relies on this exact set): `watch?v=`,
 * `embed/`, `shorts/`, `live/`, `youtu.be/<id>`, `youtube-nocookie.com/embed/`.
 */
export function youTubeVideoId(url: URL): string | null {
  const host = url.hostname.toLowerCase().replace(/^www\./, "");
  let id = "";
  if (host === "youtu.be") {
    id = pathSegment(url.pathname, 0);
  } else if (hostMatches(host, "youtube.com")) {
    const first = pathSegment(url.pathname, 0).toLowerCase();
    id =
      first === "watch"
        ? (url.searchParams.get("v") ?? "")
        : first === "embed" || first === "shorts" || first === "live"
          ? pathSegment(url.pathname, 1)
          : "";
  } else if (hostMatches(host, "youtube-nocookie.com")) {
    id =
      pathSegment(url.pathname, 0).toLowerCase() === "embed"
        ? pathSegment(url.pathname, 1)
        : "";
  }
  return YOUTUBE_ID_RE.test(id) ? id : null;
}

/**
 * The Vimeo video id of a vimeo.com / player.vimeo.com address, or null:
 * `vimeo.com/<id>[/<hash>]`, `player.vimeo.com/video/<id>`.
 */
export function vimeoVideoId(url: URL): string | null {
  const host = url.hostname.toLowerCase().replace(/^www\./, "");
  const first = pathSegment(url.pathname, 0);
  const id =
    host === "player.vimeo.com" && first.toLowerCase() === "video"
      ? pathSegment(url.pathname, 1)
      : hostMatches(host, "vimeo.com")
        ? first
        : "";
  return VIMEO_ID_RE.test(id) ? id : null;
}

function youTube(url: URL, id: string): OnlineVideo {
  const startSeconds =
    parseYouTubeTime(url.searchParams.get("start")) ??
    parseYouTubeTime(url.searchParams.get("t"));
  return Object.freeze({
    provider: "youtube" as const,
    id,
    startSeconds,
    vimeoHash: null,
    linkUrl: `https://www.youtube.com/embed/${id}${startSeconds ? `?start=${startSeconds}` : ""}`,
    watchUrl: `https://www.youtube.com/watch?v=${id}${startSeconds ? `&t=${startSeconds}s` : ""}`,
  });
}

function vimeo(url: URL, id: string): OnlineVideo {
  const host = url.hostname.toLowerCase().replace(/^www\./, "");
  const rawHash =
    host === "player.vimeo.com"
      ? url.searchParams.get("h")
      : pathSegment(url.pathname, 1) || url.searchParams.get("h");
  // (`player.vimeo.com/video/<id>?h=`; `vimeo.com/<id>/<hash>` as Share
  // gives an unlisted video.)
  const vimeoHash = rawHash && VIMEO_HASH_RE.test(rawHash) ? rawHash : null;
  return Object.freeze({
    provider: "vimeo" as const,
    id,
    startSeconds: null,
    vimeoHash,
    linkUrl: `https://player.vimeo.com/video/${id}${vimeoHash ? `?h=${vimeoHash}` : ""}`,
    watchUrl: `https://vimeo.com/${id}${vimeoHash ? `/${vimeoHash}` : ""}`,
  });
}

export interface ParseOnlineVideoOptions {
  /** Accept a pasted `<iframe src="…">` embed code (the insert dialog). */
  allowEmbedCode?: boolean;
}

/**
 * Recognise a YouTube or Vimeo video from its page address, player address
 * or (opt-in) embed code. Null for anything else: another site, a channel or
 * playlist page, a non-https or credentialed URL, a malformed id.
 */
export function parseOnlineVideo(
  value: unknown,
  options: ParseOnlineVideoOptions = {},
): OnlineVideo | null {
  if (typeof value !== "string") return null;
  const trimmed = value.trim();
  const candidate = options.allowEmbedCode
    ? (onlineVideoEmbedCodeSrc(trimmed) ?? trimmed)
    : trimmed;
  const parsed = parseHttpsUrl(candidate);
  if (!parsed) return null;
  const host = parsed.hostname.toLowerCase();
  if (
    hostMatches(host, "youtube.com") ||
    hostMatches(host, "youtube-nocookie.com") ||
    hostMatches(host, "youtu.be")
  ) {
    const id = youTubeVideoId(parsed);
    return id ? youTube(parsed, id) : null;
  }
  if (hostMatches(host, "vimeo.com")) {
    const id = vimeoVideoId(parsed);
    return id ? vimeo(parsed, id) : null;
  }
  return null;
}

/**
 * Whole seconds an online video starts at, from what a person or the AI says:
 * `45`, `"0:45"`, `"1:30"`, `"1:02:03"`, `"90s"`, `"1m30s"`. `0`, `null` and
 * `""` are null (from the beginning); anything else (or more than 24 hours)
 * is undefined. Mirrors `online_video_start_seconds` in the API's
 * presentation_online_video.py.
 */
export function onlineVideoStartSeconds(value: unknown): number | null | undefined {
  if (value === null || value === undefined) return null;
  let total: number;
  if (typeof value === "number") {
    if (!Number.isFinite(value) || value < 0) return undefined;
    total = Math.floor(value);
  } else if (typeof value === "string") {
    const text = value.trim().toLowerCase();
    if (!text) return null;
    const clock = text.match(/^(?:(\d{1,2}):)?(\d{1,4}):(\d{2})$/);
    if (clock) {
      const hours = Number(clock[1] ?? 0);
      const minutes = Number(clock[2]);
      const seconds = Number(clock[3]);
      if (seconds >= 60 || (clock[1] !== undefined && minutes >= 60)) return undefined;
      total = hours * 3_600 + minutes * 60 + seconds;
    } else if (/^\d+$/.test(text)) {
      total = Number.parseInt(text, 10);
    } else {
      const parsed = parseYouTubeTime(text);
      if (parsed === null) return undefined;
      total = parsed;
    }
  } else {
    return undefined;
  }
  if (!Number.isSafeInteger(total) || total > MAX_START_SECONDS) return undefined;
  return total || null;
}

/**
 * The same YouTube video starting at `startSeconds` (null: from the
 * beginning), any start it had replaced. An online video cannot be trimmed
 * (PowerPoint's Trim, Fade and Bookmarks apply to video files); where it
 * starts is its player address's `start` (YouTube IFrame Player parameters),
 * which PowerPoint and the slide's player both play from. Null for a Vimeo
 * video, whose player link takes no start time here.
 */
export function onlineVideoAt(video: OnlineVideo, startSeconds: number | null): OnlineVideo | null {
  if (video.provider !== "youtube") return null;
  return youTube(new URL(`https://www.youtube.com/embed/${video.id}${startSeconds ? `?start=${startSeconds}` : ""}`), video.id);
}

export interface OnlineVideoEmbedOptions {
  /** The embedding page's origin: lets the YouTube player post its state. */
  origin?: string | null;
  /** Start playing at once (only after the viewer's click to play). */
  autoplay?: boolean;
}

/**
 * The privacy-enhanced player address a slide frames. YouTube: the
 * youtube-nocookie.com player (no cookies until the viewer plays), related
 * videos from the same channel only, inline on phones, and its JS API on so
 * the slide can tell when the owner refuses embedding. Vimeo: Do Not Track
 * (`dnt=1`, no session cookies or analytics) and the unlisted hash.
 */
export function onlineVideoEmbedUrl(
  video: OnlineVideo,
  options: OnlineVideoEmbedOptions = {},
): string {
  const params = new URLSearchParams();
  if (video.provider === "youtube") {
    params.set("rel", "0");
    params.set("playsinline", "1");
    if (video.startSeconds) params.set("start", String(video.startSeconds));
    params.set("enablejsapi", "1");
    const origin = options.origin ? parseHttpsOrLoopbackOrigin(options.origin) : null;
    if (origin) params.set("origin", origin);
    if (options.autoplay) params.set("autoplay", "1");
    return `https://www.youtube-nocookie.com/embed/${video.id}?${params.toString()}`;
  }
  if (video.vimeoHash) params.set("h", video.vimeoHash);
  params.set("dnt", "1");
  if (options.autoplay) params.set("autoplay", "1");
  return `https://player.vimeo.com/video/${video.id}?${params.toString()}`;
}

/** The frame origin a provider's player is served from. */
export function onlineVideoFrameOrigin(provider: OnlineVideoProvider): string {
  return provider === "youtube"
    ? ONLINE_VIDEO_FRAME_ORIGINS[0]
    : ONLINE_VIDEO_FRAME_ORIGINS[1];
}

/** The provider's display name. */
export function onlineVideoProviderName(provider: OnlineVideoProvider): string {
  return provider === "youtube" ? "YouTube" : "Vimeo";
}

function parseHttpsOrLoopbackOrigin(value: string): string | null {
  try {
    const parsed = new URL(value);
    if (parsed.protocol === "https:") return parsed.origin;
    if (
      parsed.protocol === "http:" &&
      ["localhost", "127.0.0.1", "[::1]"].includes(parsed.hostname)
    ) {
      return parsed.origin;
    }
    return null;
  } catch {
    return null;
  }
}

/**
 * What a provider's player reported through `postMessage`: `refused` when the
 * video cannot be played embedded (YouTube IFrame API `onError` 100 removed /
 * private, 101 and 150 embedding disallowed; 2 and 5 are bad-parameter /
 * HTML5 errors the player shows itself), Vimeo player.js `error`; `ready`
 * when the player answered at all; null for anything else. The caller has
 * already checked the message's origin and source window.
 */
export type OnlineVideoPlayerSignal = "ready" | "refused" | null;

export function onlineVideoPlayerSignal(
  provider: OnlineVideoProvider,
  data: unknown,
): OnlineVideoPlayerSignal {
  let record: unknown = data;
  if (typeof data === "string") {
    if (data.length > 65_536) return null;
    try {
      record = JSON.parse(data);
    } catch {
      return null;
    }
  }
  if (!record || typeof record !== "object" || Array.isArray(record)) return null;
  const event = (record as { event?: unknown }).event;
  if (provider === "youtube") {
    if (event === "onError") {
      const info = (record as { info?: unknown }).info;
      return info === 100 || info === 101 || info === 150 ? "refused" : null;
    }
    return event === "initialDelivery" || event === "onReady" || event === "infoDelivery"
      ? "ready"
      : null;
  }
  if (event === "error") return "refused";
  return event === "ready" ? "ready" : null;
}

/**
 * The message a slide sends a provider's player once its frame has loaded so
 * the player reports its state back (YouTube's widget `listening` handshake;
 * Vimeo player.js `addEventListener` for `error`).
 */
export function onlineVideoListenMessage(
  provider: OnlineVideoProvider,
  channelId: string,
): string {
  return provider === "youtube"
    ? JSON.stringify({ event: "listening", id: channelId, channel: "widget" })
    : JSON.stringify({ method: "addEventListener", value: "error" });
}
