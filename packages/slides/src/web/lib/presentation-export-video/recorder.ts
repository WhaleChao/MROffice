/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * File > Export > Create a Video / Create an Animated GIF: the frames.
 *
 * The show is drawn by its own renderer (the present overlay, in a hidden
 * frame the size of the video); everything that moves in it is a CSS
 * animation or transition, so time is the recorder's: every animation that
 * started is paused and set to the frame's time, and the frame is drawn to a
 * canvas (`modern-screenshot`, the computed styles of that instant). Frames
 * go to an MP4 (WebCodecs, H.264 as PowerPoint writes it, else VP9) or to a
 * GIF.
 */

import type { EncodedExportAudio } from "./export-audio";
import { createByteSpool } from "./byte-spool";
import { Mp4Muxer, type Mp4VideoCodec } from "./mp4-muxer";
import { GifEncoder } from "./gif-encoder";

export type VideoExportFormat = "mp4" | "gif";

/** PowerPoint's video qualities (Create a Video) and GIF sizes, by height. */
export const VIDEO_EXPORT_QUALITIES = [
  { id: "uhd", height: 2160 },
  { id: "fhd", height: 1080 },
  { id: "hd", height: 720 },
  { id: "sd", height: 480 },
] as const;
export const GIF_EXPORT_QUALITIES = [
  { id: "extraLarge", height: 1080 },
  { id: "large", height: 720 },
  { id: "medium", height: 480 },
  { id: "small", height: 240 },
] as const;

/** Frames per second: PowerPoint's video is 30; a GIF is kept to 15. */
export const VIDEO_FPS = 30;
export const GIF_FPS = 15;

/** Animations that started and will end: what a segment waits for. */
export function finiteRunningAnimations(root: Element): Animation[] {
  return root
    .getAnimations({ subtree: true })
    .filter((animation) => animation.playState !== "finished" && Number.isFinite(endTimeOf(animation)));
}

/** Animations that never end (an emphasis that repeats, Animate as Background). */
export function endlessAnimations(root: Element): Animation[] {
  return root
    .getAnimations({ subtree: true })
    .filter((animation) => animation.playState !== "finished" && !Number.isFinite(endTimeOf(animation)));
}

export function endTimeOf(animation: Animation): number {
  const timing = animation.effect?.getComputedTiming();
  const end = timing?.endTime;
  return typeof end === "number" ? end : Number(end ?? 0);
}

/** How long a segment runs: until its last finite animation ends. */
export function segmentLengthMs(animations: readonly Animation[]): number {
  let longest = 0;
  for (const animation of animations) longest = Math.max(longest, endTimeOf(animation));
  return longest;
}









export function endlessFirstPassMs(animations: readonly Animation[]): number {
  let longest = 0;
  for (const animation of animations) {
    try {
      const timing = animation.effect?.getComputedTiming();
      if (!timing) continue;
      const rate = animation.playbackRate > 0 ? animation.playbackRate : 1;
      const pass = Number(timing.delay ?? 0) + Number(timing.duration ?? 0);
      if (Number.isFinite(pass)) longest = Math.max(longest, pass / rate);
    } catch {
      // Cancelled meanwhile.
    }
  }
  return longest;
}

/**
 * The recording's clock for effects that never end, each on its own time:
 * an effect is at 0 on the frame it was first seen and runs on from there
 * across the groups and holds of its slide (the clock of the whole
 * recording put one first seen 13 s in at 13 s: past its delay, mid-pass).
 */
export function createEndlessClock(): { seek(animations: readonly Animation[], nowMs: number): void } {
  const origins = new WeakMap<Animation, number>();
  return {
    seek(animations, nowMs) {
      for (const animation of animations) {
        let origin = origins.get(animation);
        if (origin === undefined) {
          origin = nowMs;
          origins.set(animation, origin);
        }
        try {
          animation.currentTime = nowMs - origin;
        } catch {
          // Cancelled meanwhile.
        }
      }
    },
  };
}

/** Stop the clock on these animations; each frame sets their time. */
export function pauseAnimations(animations: readonly Animation[]): void {
  for (const animation of animations) {
    try {
      animation.pause();
    } catch {
      // A transition cancelled under us: nothing to hold.
    }
  }
}

/**
 * A keyframe-less animation is a clock (a timing tree's media call or
 * effect sound waits on it): it draws nothing, and what it starts happens
 * when it finishes.
 */
function isClock(animation: Animation): boolean {
  const effect = animation.effect as KeyframeEffect | null;
  try {
    return typeof effect?.getKeyframes === "function" && effect.getKeyframes().length === 0;
  } catch {
    return false;
  }
}

export function seekAnimations(animations: readonly Animation[], timeMs: number): void {
  for (const animation of animations) {
    try {
      animation.currentTime = timeMs;
      // A paused animation never finishes by being sought past its end: a
      // clock the frame has reached is finished, so what waits on it (a
      // clip's play, an effect's sound) starts on this frame, not at the
      // segment's end.
      if (timeMs >= endTimeOf(animation) && animation.playState !== "finished" && isClock(animation)) animation.finish();
    } catch {
      // Cancelled meanwhile.
    }
  }
}










export function stillUntilMs(animations: readonly Animation[], timeMs: number): number {
  let until = Number.POSITIVE_INFINITY;
  for (const animation of animations) {
    if (animation.playState === "finished") continue;
    const rate = animation.playbackRate;
    if (!(rate > 0)) return timeMs;
    let timing: ComputedEffectTiming | undefined;
    try {
      timing = animation.effect?.getComputedTiming();
    } catch {
      return timeMs;
    }
    if (!timing) continue;
    const delay = Number(timing.delay ?? 0) / rate;
    const active = Number(timing.activeDuration ?? 0) / rate;
    const end = delay + active;
    if (!Number.isFinite(end)) return timeMs;
    if (isClock(animation)) {
      if (timeMs < end) until = Math.min(until, end);
      continue;
    }
    if (timeMs < delay) until = Math.min(until, delay);
    else if (timeMs < end) return timeMs;
  }
  return until;
}

/** Whether a clock reached its end at this frame (`seekAnimations` finished it just now). */
export function clockEndedWithin(animations: readonly Animation[], fromMs: number, toMs: number): boolean {
  return animations.some((animation) => {
    if (!isClock(animation)) return false;
    const end = endTimeOf(animation) / (animation.playbackRate > 0 ? animation.playbackRate : 1);
    return end > fromMs && end <= toMs;
  });
}

/** Land every finite animation on its end, as the show leaves it. */
export function finishAnimations(animations: readonly Animation[]): void {
  for (const animation of animations) {
    try {
      animation.finish();
    } catch {
      // An endless one cannot finish; it stays where the clock left it.
    }
  }
}

export interface FrameSink {
  /** Add the canvas as it stands, shown for `durationMs`. */
  addFrame(canvas: HTMLCanvasElement, durationMs: number): Promise<void>;
  /** The file, with the show's sound as its track where the format has one (MP4). */
  finish(audio?: EncodedExportAudio | null): Promise<Blob>;
  abort(): void;
}

// ---------------------------------------------------------------------------
// MP4 through WebCodecs.
// ---------------------------------------------------------------------------

type CodecChoice = { codec: string; kind: Mp4VideoCodec; vp9Level?: number };

function avcLevel(width: number, height: number): string {
  const pixels = width * height;
  // H.264 levels by frame size (Table A-1): 3.0 ≤ 414,720, 3.1 ≤ 921,600,
  // 4.0 ≤ 2,097,152, 5.1 ≤ 9,437,184 luma samples.
  if (pixels <= 414_720) return "1e";
  if (pixels <= 921_600) return "1f";
  if (pixels <= 2_097_152) return "28";
  return "33";
}

function vp9Level(width: number, height: number): number {
  const pixels = width * height;
  if (pixels <= 409_920) return 30;
  if (pixels <= 921_600) return 31;
  if (pixels <= 2_228_224) return 40;
  return 51;
}

function codecChoices(width: number, height: number): CodecChoice[] {
  const level = avcLevel(width, height);
  const vp9 = vp9Level(width, height);
  return [
    { codec: `avc1.6400${level}`, kind: "avc" },
    { codec: `avc1.4d00${level}`, kind: "avc" },
    { codec: `avc1.4200${level}`, kind: "avc" },
    { codec: `vp09.00.${vp9}.08`, kind: "vp9", vp9Level: vp9 },
  ];
}

export function videoExportSupported(): boolean {
  return typeof VideoEncoder !== "undefined" && typeof VideoFrame !== "undefined";
}

function bitrateFor(width: number, height: number, fps: number): number {
  return Math.min(40_000_000, Math.max(1_000_000, Math.round(width * height * fps * 0.12)));
}

export async function createMp4Sink(width: number, height: number, fps: number): Promise<FrameSink> {
  if (!videoExportSupported()) throw new Error("video-encoder-unavailable");
  const bitrate = bitrateFor(width, height, fps);
  let chosen: CodecChoice | null = null;
  let config: VideoEncoderConfig | null = null;
  for (const choice of codecChoices(width, height)) {
    const candidate: VideoEncoderConfig = {
      codec: choice.codec,
      width,
      height,
      bitrate,
      framerate: fps,
      // No reordered (B) frames: the file states no composition offsets.
      latencyMode: "realtime",
      ...(choice.kind === "avc" ? { avc: { format: "avc" } } : {}),
    };
    try {
      const support = await VideoEncoder.isConfigSupported(candidate);
      if (support.supported) {
        chosen = choice;
        config = candidate;
        break;
      }
    } catch {
      // Not this one.
    }
  }
  if (!chosen || !config) throw new Error("video-codec-unavailable");
  const choice = chosen;
  // The encoded frames go out of the page as they come (`byte-spool`).
  const spool = await createByteSpool("frames");
  let muxer: Mp4Muxer | null = null;
  const durations = new Map<number, number>();
  let failure: Error | null = null;
  const encoder = new VideoEncoder({
    output(chunk, metadata) {
      if (!muxer) {
        const description = metadata?.decoderConfig?.description;
        muxer = new Mp4Muxer({
          codec: choice.kind,
          width,
          height,
          ...(choice.vp9Level !== undefined ? { vp9Level: choice.vp9Level } : {}),
          ...(description
            ? {
                description: ArrayBuffer.isView(description)
                  ? new Uint8Array(description.buffer.slice(description.byteOffset, description.byteOffset + description.byteLength))
                  : new Uint8Array((description as ArrayBuffer).slice(0)),
              }
            : {}),
        }, spool);
      }
      const data = new Uint8Array(chunk.byteLength);
      chunk.copyTo(data);
      const durationUs = durations.get(chunk.timestamp) ?? chunk.duration ?? Math.round(1_000_000 / fps);
      durations.delete(chunk.timestamp);
      try {
        muxer.addFrame(data, durationUs, chunk.type === "key");
      } catch (error) {
        // The disk refused (full): the export fails rather than drop frames.
        failure = error instanceof Error ? error : new Error(String(error));
      }
    },
    error(error) {
      failure = error instanceof Error ? error : new Error(String(error));
    },
  });
  encoder.configure(config);
  let timestampUs = 0;
  let frames = 0;
  let sinceKeyUs = Infinity;
  return {
    async addFrame(canvas, durationMs) {
      if (failure) throw failure;
      // A long still is cut into one-second frames: a player seeks into it.
      let remainingUs = Math.max(1, Math.round(durationMs * 1000));
      let first = true;
      while (remainingUs > 0) {
        const piece = Math.min(remainingUs, 1_000_000);
        const frame = new VideoFrame(canvas, { timestamp: timestampUs, duration: piece });
        durations.set(timestampUs, piece);
        // A key frame every two seconds of a moving picture; a still's later
        // seconds are the same picture again (a key frame there was coded
        
        const key = first && sinceKeyUs >= 2_000_000;
        first = false;
        encoder.encode(frame, { keyFrame: key });
        frame.close();
        sinceKeyUs = key ? piece : sinceKeyUs + piece;
        timestampUs += piece;
        remainingUs -= piece;
        frames += 1;
        if (encoder.encodeQueueSize > 8) {
          await new Promise<void>((resolve) => encoder.addEventListener("dequeue", () => resolve(), { once: true }));
        }
      }
      if (frames % 60 === 0) await encoder.flush();
      // The recording waits for the disk rather than hold frames for it.
      await spool.drained();
    },
    async finish(audio) {
      await encoder.flush();
      encoder.close();
      if (failure) throw failure;
      if (!muxer) throw new Error("no-frames");
      return (muxer as Mp4Muxer).finalize(audio);
    },
    abort() {
      try {
        encoder.close();
      } catch {
        // Already closed.
      }
      void spool.discard();
    },
  };
}

// ---------------------------------------------------------------------------
// GIF.
// ---------------------------------------------------------------------------

export function createGifSink(width: number, height: number, transparent: boolean): FrameSink {
  const encoder = new GifEncoder({ width, height, transparent });
  return {
    async addFrame(canvas, durationMs) {
      const context = canvas.getContext("2d", { willReadFrequently: true });
      if (!context) throw new Error("canvas-unavailable");
      const pixels = context.getImageData(0, 0, width, height);
      encoder.addFrame({ rgba: pixels.data, durationMs });
    },
    async finish() {
      return new Blob([encoder.finish() as BlobPart], { type: "image/gif" });
    },
    abort() {},
  };
}
