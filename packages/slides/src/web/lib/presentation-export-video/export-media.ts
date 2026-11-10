/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { getApiHttpBaseUrl } from "@mrnote/core/web/lib/env.ts";

const proto = HTMLMediaElement.prototype;

type ClipState = {
  readonly media: HTMLMediaElement;
  playing: boolean;
  /** The recording's time (ms) and the clip's position (s) when it last started. */
  startedAt: number;
  startOffset: number;
  /** The frame painted for the picture (videos). */
  canvas: HTMLCanvasElement | null;
  /** The clip's frames cannot be read (another site's, no CORS): its poster is drawn. */
  unreadable: boolean;
  /** A seek this clock made (not the show's). */
  ownSeek: boolean;
  /** Whether it ever showed a frame of its own. */
  shown: boolean;
  /** The stretch it is playing now, as heard: from when (ms), from where (s). */
  heard: { at: number; offset: number } | null;
};

/** One stretch of a clip's sound in the recording. */
export type ExportHeardClip = {
  /** The clip's address, as the page plays it. */
  readonly src: string;
  /** Where in the recording it starts and ends (ms). */
  readonly startMs: number;
  readonly endMs: number;
  /** Where in the clip it starts (s). */
  readonly offsetS: number;
  readonly rate: number;
  /** 0..1; 0 when the clip is muted. */
  readonly volume: number;
  /** The clip plays round and round (its own Loop until Stopped). */
  readonly loop: boolean;
  /** The clip's own trim end (s), where a loop goes round. */
  readonly endS?: number;
};

export type ExportMediaClock = {
  /** The recording's time, in ms. */
  readonly now: number;
  /** Take in the clips under the host (called as the show mounts them). */
  adopt(): void;
  /** Move the recording's clock on. */
  advance(ms: number): void;
  /** Seek every clip to the clock and paint each video's frame. */
  sync(): Promise<void>;
  /** Whether a clip is playing that the slide waits for (not a loop, not one playing across slides). */
  blocking(): boolean;
  /** Whether a playing clip has frames to draw (a video, not a sound). */
  playingVideo(): boolean;
  /** How long (ms) the clips the slide waits for still play, at most. */
  remainingMs(): number;
  /** Whether `node` is a clip the picture leaves out (a video: its canvas stands for it). */
  leftOut(node: Node): boolean;
  /** Every stretch of sound the clips played, those still playing ended now. */
  heard(): ExportHeardClip[];
  dispose(): void;
};

const SEEK_WAIT_MS = 4_000;
const LOAD_WAIT_MS = 8_000;
/** Positions closer than this are the same frame. */
const SAME_POSITION_S = 0.0005;

function apiOrigin(): string | null {
  try {
    // The API's own media are read with the viewer's sign-in.
    const base = getApiHttpBaseUrl();
    return base ? new URL(base, location.href).origin : null;
  } catch {
    return null;
  }
}

function crossOriginFor(src: string): "use-credentials" | "anonymous" | null {
  if (!src || src.startsWith("data:") || src.startsWith("blob:")) return null;
  try {
    const origin = new URL(src, location.href).origin;
    if (origin === location.origin) return null;
    const api = apiOrigin();
    if ((api && origin === api) || /\/api\/v1\//.test(src)) return "use-credentials";
    return "anonymous";
  } catch {
    return null;
  }
}

function waitFor(media: HTMLMediaElement, events: string[], ms: number): Promise<void> {
  return new Promise<void>((resolve) => {
    const done = () => {
      clearTimeout(timer);
      for (const name of events) media.removeEventListener(name, done);
      resolve();
    };
    const timer = setTimeout(done, ms);
    for (const name of events) media.addEventListener(name, done, { once: true });
  });
}

function endOf(media: HTMLMediaElement): number {
  const duration = Number.isFinite(media.duration) ? media.duration : NaN;
  const trim = Number(media.dataset.playbackTrimEnd);
  const trimEnd = Number.isFinite(trim) && trim > 0 ? trim / 1000 : NaN;
  if (Number.isFinite(trimEnd) && Number.isFinite(duration)) return Math.min(trimEnd, duration);
  return Number.isFinite(trimEnd) ? trimEnd : duration;
}

const posterPictures = new Map<string, Promise<HTMLImageElement | null>>();

function posterPicture(url: string): Promise<HTMLImageElement | null> {
  let picture = posterPictures.get(url);
  if (!picture) {
    picture = new Promise<HTMLImageElement | null>((resolve) => {
      const image = new Image();
      const mode = crossOriginFor(url);
      if (mode) image.crossOrigin = mode;
      image.onload = () => resolve(image);
      image.onerror = () => resolve(null);
      image.src = url;
    });
    posterPictures.set(url, picture);
  }
  return picture;
}

const COPIED_STYLE = [
  "object-fit",
  "object-position",
  "transform",
  "transform-origin",
  "border-radius",
  "clip-path",
  "filter",
  "opacity",
  "visibility",
  "z-index",
  "mix-blend-mode",
] as const;

/** How loud the show plays a clip (its Volume, `data-playback-volume`; the element's own volume). */
function heardVolume(media: HTMLMediaElement): number {
  if (media.dataset.playbackVolume === "mute") return 0;
  const volume = Number.isFinite(media.volume) ? media.volume : 1;
  return Math.max(0, Math.min(1, volume));
}

function heardSource(media: HTMLMediaElement): string {
  return media.currentSrc || media.src || media.getAttribute("src") || "";
}

export function createExportMediaClock(host: HTMLElement): ExportMediaClock {
  let now = 0;
  const clips = new Map<HTMLMediaElement, ClipState>();
  const heard: ExportHeardClip[] = [];

  /** The stretch a clip is playing ends here (it paused, ended, left the slide). */
  const stopHearing = (state: ClipState) => {
    const open = state.heard;
    state.heard = null;
    if (!open || now <= open.at) return;
    const media = state.media;
    const src = heardSource(media);
    if (!src) return;
    const end = endOf(media);
    heard.push({
      src,
      startMs: open.at,
      endMs: now,
      offsetS: open.offset,
      rate: media.playbackRate > 0 ? media.playbackRate : 1,
      volume: heardVolume(media),
      loop: media.loop,
      ...(Number.isFinite(end) && end > 0 ? { endS: end } : {}),
    });
  };

  const start = (state: ClipState) => {
    stopHearing(state);
    state.playing = true;
    state.startedAt = now;
    state.startOffset = state.media.currentTime;
    state.heard = { at: now, offset: state.startOffset };
  };

  const adoptOne = (media: HTMLMediaElement) => {
    if (clips.has(media)) return;
    const state: ClipState = {
      media,
      playing: false,
      startedAt: 0,
      startOffset: 0,
      canvas: null,
      unreadable: false,
      ownSeek: false,
      shown: false,
      heard: null,
    };
    clips.set(media, state);
    media.muted = true;
    const wasAutoplay = media.autoplay;
    media.autoplay = false;
    // The show's play() and pause() run on this clock, and `paused` says so.
    Object.defineProperty(media, "play", {
      configurable: true,
      value: () => {
        if (!state.playing) {
          start(state);
          media.dispatchEvent(new Event("play"));
          media.dispatchEvent(new Event("playing"));
        } else {
          start(state);
        }
        return Promise.resolve();
      },
    });
    Object.defineProperty(media, "pause", {
      configurable: true,
      value: () => {
        if (!state.playing) return;
        state.playing = false;
        stopHearing(state);
        media.dispatchEvent(new Event("pause"));
      },
    });
    Object.defineProperty(media, "paused", { configurable: true, get: () => !state.playing });
    // A seek the show makes while the clip plays moves where it plays from.
    media.addEventListener("seeking", () => {
      if (!state.ownSeek && state.playing) start(state);
    });
    // Should the browser play it after all (an attribute set later), it
    // plays on this clock instead.
    media.addEventListener("playing", (event) => {
      if (!event.isTrusted) return;
      proto.pause.call(media);
      if (!state.playing) start(state);
    });
    const mode = crossOriginFor(media.currentSrc || media.src || media.getAttribute("src") || "");
    if (mode && media.crossOrigin !== mode) {
      media.crossOrigin = mode;
      media.preload = "auto";
      media.addEventListener(
        "error",
        () => {
          // Not readable with CORS: its poster is drawn instead.
          state.unreadable = true;
        },
        { once: true },
      );
      try {
        media.load();
      } catch {
        state.unreadable = true;
      }
    } else if (media.preload !== "auto") {
      media.preload = "auto";
    }
    if (wasAutoplay) start(state);
  };

  const adopt = () => {
    for (const media of Array.from(host.querySelectorAll<HTMLMediaElement>("video, audio"))) adoptOne(media);
    for (const [media, state] of clips) {
      if (!media.isConnected) {
        // A clip that leaves the show stops sounding (unless it plays across
        // slides: that one is the overlay's and stays).
        stopHearing(state);
        state.canvas?.remove();
        clips.delete(media);
      }
    }
  };

  const seek = async (state: ClipState, to: number) => {
    const media = state.media;
    if (media.readyState < HTMLMediaElement.HAVE_METADATA && !state.unreadable) {
      await waitFor(media, ["loadedmetadata", "error"], LOAD_WAIT_MS);
    }
    if (Math.abs(media.currentTime - to) < SAME_POSITION_S && media.readyState >= HTMLMediaElement.HAVE_CURRENT_DATA) return;
    state.ownSeek = true;
    try {
      const sought = waitFor(media, ["seeked", "error"], SEEK_WAIT_MS);
      media.currentTime = to;
      await sought;
    } catch {
      // Not seekable: stays where it is.
    } finally {
      state.ownSeek = false;
    }
    if (media instanceof HTMLVideoElement && media.readyState < HTMLMediaElement.HAVE_CURRENT_DATA && !state.unreadable) {
      await waitFor(media, ["loadeddata", "canplay", "error"], LOAD_WAIT_MS);
    }
  };

  const paint = async (state: ClipState) => {
    const video = state.media;
    if (!(video instanceof HTMLVideoElement)) return;
    let canvas = state.canvas;
    if (!canvas || !canvas.isConnected) {
      canvas = video.ownerDocument.createElement("canvas");
      canvas.setAttribute("data-export-video-frame", "true");
      canvas.setAttribute("aria-hidden", "true");
      video.after(canvas);
      state.canvas = canvas;
    }
    // Where the video is, as the video is: the same box in the same
    // containing block, under the same transforms (its ancestors' too).
    const style = getComputedStyle(video);
    canvas.style.position = "absolute";
    canvas.style.left = `${video.offsetLeft}px`;
    canvas.style.top = `${video.offsetTop}px`;
    canvas.style.width = `${video.offsetWidth}px`;
    canvas.style.height = `${video.offsetHeight}px`;
    canvas.style.display = style.display === "none" || video.hidden ? "none" : "block";
    canvas.style.pointerEvents = "none";
    for (const property of COPIED_STYLE) canvas.style.setProperty(property, style.getPropertyValue(property));
    const context = canvas.getContext("2d");
    if (!context) return;
    const frameReady =
      !state.unreadable && video.readyState >= HTMLMediaElement.HAVE_CURRENT_DATA && video.videoWidth > 0;
    // Before it first plays a clip shows its poster, as in the show.
    const usePoster = !state.shown && !state.playing && Boolean(video.poster);
    if (frameReady && !usePoster) {
      if (canvas.width !== video.videoWidth || canvas.height !== video.videoHeight) {
        canvas.width = video.videoWidth;
        canvas.height = video.videoHeight;
      }
      try {
        context.drawImage(video, 0, 0, canvas.width, canvas.height);
        // A frame that cannot be read back would spoil the whole picture.
        context.getImageData(0, 0, 1, 1);
        state.shown = true;
        return;
      } catch {
        state.unreadable = true;
        canvas.width = 0;
        canvas.height = 0;
      }
    }
    const poster = video.poster ? await posterPicture(video.poster) : null;
    if (poster) {
      if (canvas.width !== poster.naturalWidth || canvas.height !== poster.naturalHeight) {
        canvas.width = poster.naturalWidth;
        canvas.height = poster.naturalHeight;
      }
      try {
        context.drawImage(poster, 0, 0, canvas.width, canvas.height);
        context.getImageData(0, 0, 1, 1);
        return;
      } catch {
        // fall through: nothing drawn
      }
    }
    canvas.width = 1;
    canvas.height = 1;
    context.clearRect(0, 0, 1, 1);
  };

  const syncOne = async (state: ClipState) => {
    const media = state.media;
    if (state.playing) {
      const rate = media.playbackRate > 0 ? media.playbackRate : 1;
      let to = state.startOffset + ((now - state.startedAt) / 1000) * rate;
      const end = Number.isFinite(media.duration) ? media.duration : Number.POSITIVE_INFINITY;
      let ended = false;
      if (to >= end) {
        if (media.loop && end > 0) {
          to = to % end;
        } else {
          to = end;
          ended = true;
        }
      }
      await seek(state, to);
      if (ended && state.playing) {
        state.playing = false;
        stopHearing(state);
        media.dispatchEvent(new Event("pause"));
        media.dispatchEvent(new Event("ended"));
      }
    } else if (media instanceof HTMLVideoElement && media.readyState < HTMLMediaElement.HAVE_CURRENT_DATA && !state.unreadable && !media.poster) {
      // A clip with no poster shows its first frame (or its trim start).
      await seek(state, media.currentTime || 0);
    }
    await paint(state);
  };

  return {
    get now() {
      return now;
    },
    adopt,
    advance(ms: number) {
      now += Math.max(0, ms);
    },
    async sync() {
      adopt();
      await Promise.all(Array.from(clips.values()).map((state) => syncOne(state)));
    },
    blocking() {
      for (const state of clips.values()) {
        if (!state.playing) continue;
        const media = state.media;
        if (media.loop || media.dataset.playbackLoop === "true" || media.dataset.playbackAcross === "true") continue;
        return true;
      }
      return false;
    },
    playingVideo() {
      for (const state of clips.values()) {
        if (state.playing && state.media instanceof HTMLVideoElement) return true;
      }
      return false;
    },
    remainingMs() {
      let longest = 0;
      for (const state of clips.values()) {
        const media = state.media;
        if (!state.playing || media.loop || media.dataset.playbackLoop === "true" || media.dataset.playbackAcross === "true") continue;
        const rate = media.playbackRate > 0 ? media.playbackRate : 1;
        const at = state.startOffset + ((now - state.startedAt) / 1000) * rate;
        const end = endOf(media);
        if (!Number.isFinite(end)) continue;
        longest = Math.max(longest, ((end - at) / rate) * 1000);
      }
      return Math.max(0, longest);
    },
    leftOut(node: Node) {
      // Every video: the ones taken in are drawn by their canvas.
      return node instanceof HTMLVideoElement;
    },
    heard() {
      for (const state of clips.values()) {
        if (!state.heard) continue;
        // Still playing at the end: heard up to now, and on from here.
        const at = now;
        stopHearing(state);
        state.heard = { at, offset: state.startOffset + ((at - state.startedAt) / 1000) * (state.media.playbackRate > 0 ? state.media.playbackRate : 1) };
      }
      return heard.slice();
    },
    dispose() {
      for (const state of clips.values()) state.canvas?.remove();
      clips.clear();
    },
  };
}
