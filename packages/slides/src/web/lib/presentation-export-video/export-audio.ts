/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { createBlobSpool } from "./byte-spool";
import type { ExportHeardClip } from "./export-media";

export const EXPORT_AUDIO_SAMPLE_RATE = 48_000;
export const EXPORT_AUDIO_CHANNELS = 2;

/** A clip larger than this is not read for its sound (the video keeps its pictures). */
const MAX_CLIP_BYTES = 400 * 1024 * 1024;

export type ExportAudioCodec = "aac" | "opus";

export type EncodedExportAudio = {
  readonly codec: ExportAudioCodec;
  readonly sampleRate: number;
  readonly channels: number;
   
  readonly description?: Uint8Array;
  /** Each encoded chunk's size and length, in order; their bytes, one after another, are `payload`. */
  readonly chunks: ReadonlyArray<{ readonly size: number; readonly durationUs: number }>;
  readonly payload: Blob;
  /** Samples the encoder put before the sound (AAC's priming): the file skips them. */
  readonly primingSamples?: number;
  /** How many samples of sound there are (the rest is the encoder's padding). */
  readonly playedSamples?: number;
};

/** A stretch laid on the track: when it starts and how long it lasts, in the clip's own seconds. */
export type ExportAudioPlacement = {
  readonly when: number;
  readonly offset: number;
  readonly duration: number;
  readonly loop: boolean;
  readonly loopEnd?: number;
};

/**
 * Where a stretch of a clip sits on the track (seconds): from its start in
 * the recording, for as long as it played, never past the clip's end (or its
 * trim end) unless it loops.
 */
export function placeHeardClip(clip: ExportHeardClip, clipSeconds: number): ExportAudioPlacement | null {
  const when = Math.max(0, clip.startMs / 1000);
  const heardS = Math.max(0, (clip.endMs - clip.startMs) / 1000) * clip.rate;
  const end = Math.min(clipSeconds, clip.endS ?? clipSeconds);
  if (!(end > 0) || heardS <= 0) return null;
  if (clip.loop) {
    const offset = clip.offsetS >= end ? 0 : Math.max(0, clip.offsetS);
    return { when, offset, duration: heardS, loop: true, loopEnd: end };
  }
  const offset = Math.max(0, clip.offsetS);
  const duration = Math.min(heardS, end - offset);
  return duration > 0 ? { when, offset, duration, loop: false } : null;
}

function crossOriginCredentials(src: string): RequestCredentials {
  try {
    const url = new URL(src, location.href);
    if (url.origin === location.origin) return "same-origin";
    // The API's own media are read with the viewer's sign-in (as the show reads them).
    return /\/api\/v1\//.test(url.pathname) ? "include" : "omit";
  } catch {
    return "omit";
  }
}

async function readClip(src: string, context: BaseAudioContext): Promise<AudioBuffer | null> {
  try {
    const response = await fetch(src, { credentials: crossOriginCredentials(src) });
    if (!response.ok) return null;
    const length = Number(response.headers.get("content-length"));
    if (Number.isFinite(length) && length > MAX_CLIP_BYTES) return null;
    const bytes = await response.arrayBuffer();
    if (bytes.byteLength > MAX_CLIP_BYTES) return null;
    return await context.decodeAudioData(bytes);
  } catch {
    // Not readable here, or no sound in it (a silent video): nothing heard.
    return null;
  }
}

/** How much of the track one offline graph renders: a minute. */
export const EXPORT_AUDIO_WINDOW_FRAMES = 60 * EXPORT_AUDIO_SAMPLE_RATE;

type LaidClip = { readonly buffer: AudioBuffer; readonly place: ExportAudioPlacement; readonly rate: number; readonly volume: number };

/**
 * Where a laid stretch plays inside the window that starts at `windowStart`
 * (seconds) and lasts `windowLength`: when in the window it starts, from
 * where in the clip, for how much of the clip, and when it stops; null when
 * it is not heard in that window. One long render and its windows, one
 * after another, lay the same samples.
 */
export function placeInWindow(
  place: ExportAudioPlacement,
  rate: number,
  windowStart: number,
  windowLength: number,
): { when: number; offset: number; duration: number; stop: number } | null {
  const playRate = rate > 0 ? rate : 1;
  const trackEnd = place.when + place.duration / playRate;
  const windowEnd = windowStart + windowLength;
  if (!(trackEnd > windowStart) || !(place.when < windowEnd)) return null;
  const when = Math.max(0, place.when - windowStart);
  // The clip's seconds already played before this window began.
  const skipped = Math.max(0, windowStart - place.when) * playRate;
  let offset = place.offset + skipped;
  const duration = place.duration - skipped;
  if (!(duration > 0)) return null;
  if (place.loop) {
    const loopEnd = place.loopEnd ?? 0;
    if (loopEnd > 0 && offset >= loopEnd) offset %= loopEnd;
  }
  return { when, offset, duration, stop: trackEnd - windowStart };
}

/** Render one window of the track: its frames, every laid stretch heard in it. */
async function renderWindow(laid: readonly LaidClip[], windowStartFrame: number, frames: number): Promise<AudioBuffer | null> {
  const windowStart = windowStartFrame / EXPORT_AUDIO_SAMPLE_RATE;
  const windowLength = frames / EXPORT_AUDIO_SAMPLE_RATE;
  const context = new OfflineAudioContext(EXPORT_AUDIO_CHANNELS, frames, EXPORT_AUDIO_SAMPLE_RATE);
  let heard = 0;
  for (const clip of laid) {
    const here = placeInWindow(clip.place, clip.rate, windowStart, windowLength);
    if (!here) continue;
    const source = context.createBufferSource();
    source.buffer = clip.buffer;
    source.playbackRate.value = clip.rate;
    const gain = context.createGain();
    gain.gain.value = clip.volume;
    source.connect(gain).connect(context.destination);
    if (clip.place.loop) {
      source.loop = true;
      source.loopStart = 0;
      source.loopEnd = clip.place.loopEnd ?? clip.buffer.duration;
      source.start(here.when, here.offset);
      source.stop(here.stop);
    } else {
      source.start(here.when, here.offset, here.duration);
    }
    heard += 1;
  }
  return heard > 0 ? context.startRendering() : null;
}

/** The clips heard, read once each and laid where they play; empty when nothing is heard. */
async function layHeardClips(heard: readonly ExportHeardClip[], lengthMs: number): Promise<LaidClip[]> {
  const audible = heard.filter((clip) => clip.volume > 0 && clip.endMs > clip.startMs && clip.startMs < lengthMs);
  if (audible.length === 0 || !(lengthMs > 0) || typeof OfflineAudioContext === "undefined") return [];
  // Clips are decoded at the track's rate (a context of one frame serves).
  const decoder = new OfflineAudioContext(EXPORT_AUDIO_CHANNELS, 1, EXPORT_AUDIO_SAMPLE_RATE);
  const buffers = new Map<string, Promise<AudioBuffer | null>>();
  const laid: LaidClip[] = [];
  for (const clip of audible) {
    let buffer = buffers.get(clip.src);
    if (!buffer) {
      buffer = readClip(clip.src, decoder);
      buffers.set(clip.src, buffer);
    }
    const decoded = await buffer;
    if (!decoded) continue;
    const place = placeHeardClip({ ...clip, endMs: Math.min(clip.endMs, lengthMs) }, decoded.duration);
    if (!place) continue;
    laid.push({ buffer: decoded, place, rate: clip.rate, volume: clip.volume });
  }
  return laid;
}

function audioEncoderAvailable(): boolean {
  return typeof AudioEncoder !== "undefined" && typeof AudioData !== "undefined";
}

async function pickCodec(): Promise<{ codec: ExportAudioCodec; config: AudioEncoderConfig } | null> {
  if (!audioEncoderAvailable()) return null;
  const choices: Array<{ codec: ExportAudioCodec; config: AudioEncoderConfig }> = [
    {
      codec: "aac",
      config: { codec: "mp4a.40.2", sampleRate: EXPORT_AUDIO_SAMPLE_RATE, numberOfChannels: EXPORT_AUDIO_CHANNELS, bitrate: 192_000 },
    },
    {
      codec: "opus",
      config: { codec: "opus", sampleRate: EXPORT_AUDIO_SAMPLE_RATE, numberOfChannels: EXPORT_AUDIO_CHANNELS, bitrate: 160_000 },
    },
  ];
  for (const choice of choices) {
    try {
      const support = await AudioEncoder.isConfigSupported(choice.config);
      if (support.supported) return choice;
    } catch {
      // Not this one.
    }
  }
  return null;
}

function copyDescription(description: AllowSharedBufferSource | undefined): Uint8Array | undefined {
  if (!description) return undefined;
  if (ArrayBuffer.isView(description)) {
    return new Uint8Array(description.buffer.slice(description.byteOffset, description.byteOffset + description.byteLength));
  }
  return new Uint8Array((description as ArrayBuffer).slice(0));
}

type Encoded = { chunks: Array<{ data: Uint8Array; durationUs: number }>; description?: Uint8Array };

/** One encoder fed planar float sound window after window, never flushed mid-track. */
function trackEncoder(config: AudioEncoderConfig, onChunk: (data: Uint8Array, durationUs: number) => void) {
  let description: Uint8Array | undefined;
  let failure: Error | null = null;
  let written = 0;
  const encoder = new AudioEncoder({
    output(chunk, metadata) {
      if (!description) description = copyDescription(metadata?.decoderConfig?.description);
      const data = new Uint8Array(chunk.byteLength);
      chunk.copyTo(data);
      onChunk(data, chunk.duration ?? 0);
    },
    error(error) {
      failure = error instanceof Error ? error : new Error(String(error));
    },
  });
  encoder.configure(config);
  const piece = 4096;
  return {
    /** Encode `length` frames (`channel(c)` gives channel c), after those encoded before. */
    async encode(length: number, channel: (index: number) => Float32Array): Promise<void> {
      for (let at = 0; at < length; at += piece) {
        if (failure) throw failure;
        const frames = Math.min(piece, length - at);
        // Planar float samples, channel after channel.
        const planes = new Float32Array(frames * EXPORT_AUDIO_CHANNELS);
        for (let index = 0; index < EXPORT_AUDIO_CHANNELS; index += 1) {
          planes.set(channel(index).subarray(at, at + frames), index * frames);
        }
        const data = new AudioData({
          format: "f32-planar",
          sampleRate: EXPORT_AUDIO_SAMPLE_RATE,
          numberOfFrames: frames,
          numberOfChannels: EXPORT_AUDIO_CHANNELS,
          timestamp: Math.round(((written + at) / EXPORT_AUDIO_SAMPLE_RATE) * 1_000_000),
          data: planes,
        });
        encoder.encode(data);
        data.close();
        // Wait for the queue to drain, never flush mid-track: a flush ends the
        // stream there and the encoder starts again with a fresh priming gap.
        while (encoder.encodeQueueSize > 16) {
          await new Promise<void>((resolve) => encoder.addEventListener("dequeue", () => resolve(), { once: true }));
        }
      }
      written += length;
    },
    async finish(): Promise<Uint8Array | undefined> {
      await encoder.flush();
      encoder.close();
      if (failure) throw failure;
      return description;
    },
    abort() {
      try {
        encoder.close();
      } catch {
        // Already closed.
      }
    },
  };
}

/** Encode `length` frames of planar float sound at once (a short probe). */
async function encodePlanes(
  config: AudioEncoderConfig,
  length: number,
  channel: (index: number) => Float32Array,
): Promise<Encoded> {
  const chunks: Array<{ data: Uint8Array; durationUs: number }> = [];
  const encoder = trackEncoder(config, (data, durationUs) => chunks.push({ data, durationUs }));
  await encoder.encode(length, channel);
  const description = await encoder.finish();
  return { chunks, ...(description ? { description } : {}) };
}

/** AAC's usual encoder delay (Apple's AudioToolbox, Chrome on a Mac): 2,112 samples. */
const AAC_USUAL_PRIMING = 2112;

/**
 * How many samples this browser's AAC encoder puts before the sound
 * (its priming): a click encoded and decoded back. The file's edit list
 * skips them, so the sound lands on its frames (PowerPoint's does).
 */
async function measureAacPriming(config: AudioEncoderConfig): Promise<number> {
  if (typeof AudioDecoder === "undefined") return AAC_USUAL_PRIMING;
  try {
    const length = 9600;
    const at = 4800;
    const click = new Float32Array(length);
    for (let index = 0; index < 96; index += 1) click[at + index] = Math.sin((2 * Math.PI * 1000 * index) / EXPORT_AUDIO_SAMPLE_RATE) * 0.8;
    const encoded = await encodePlanes(config, length, () => click);
    const decoded: Float32Array[] = [];
    let failed = false;
    const decoder = new AudioDecoder({
      output(data) {
        const plane = new Float32Array(data.numberOfFrames);
        data.copyTo(plane, { planeIndex: 0, format: "f32-planar" });
        decoded.push(plane);
        data.close();
      },
      error() {
        failed = true;
      },
    });
    decoder.configure({
      codec: config.codec,
      sampleRate: EXPORT_AUDIO_SAMPLE_RATE,
      numberOfChannels: EXPORT_AUDIO_CHANNELS,
      ...(encoded.description ? { description: encoded.description } : {}),
    });
    let timestamp = 0;
    for (const chunk of encoded.chunks) {
      decoder.decode(new EncodedAudioChunk({ type: "key", timestamp, duration: chunk.durationUs, data: chunk.data }));
      timestamp += chunk.durationUs;
    }
    await decoder.flush();
    decoder.close();
    if (failed) return AAC_USUAL_PRIMING;
    let index = 0;
    for (const plane of decoded) {
      for (const sample of plane) {
        if (Math.abs(sample) > 0.1) {
          const delay = index - at;
          return delay >= 0 && delay <= 8192 ? delay : AAC_USUAL_PRIMING;
        }
        index += 1;
      }
    }
    return AAC_USUAL_PRIMING;
  } catch {
    return AAC_USUAL_PRIMING;
  }
}

/**
 * The show's sound track, mixed a window at a time and encoded as it is
 * mixed; null when nothing was heard or this browser can neither mix
 * offline nor encode it.
 */
export async function makeExportAudio(
  heard: readonly ExportHeardClip[],
  lengthMs: number,
  { cancelled = () => false }: { cancelled?: () => boolean } = {},
): Promise<EncodedExportAudio | null> {
  const laid = await layHeardClips(heard, lengthMs);
  if (laid.length === 0) return null;
  const choice = await pickCodec();
  if (!choice) return null;
  const total = Math.max(1, Math.ceil((lengthMs / 1000) * EXPORT_AUDIO_SAMPLE_RATE));
  const spool = createBlobSpool();
  const chunks: Array<{ size: number; durationUs: number }> = [];
  const encoder = trackEncoder(choice.config, (data, durationUs) => {
    spool.append(data);
    chunks.push({ size: data.length, durationUs });
  });
  let silence: Float32Array | null = null;
  try {
    for (let start = 0; start < total; start += EXPORT_AUDIO_WINDOW_FRAMES) {
      if (cancelled()) throw new Error("cancelled");
      const frames = Math.min(EXPORT_AUDIO_WINDOW_FRAMES, total - start);
      const rendered = await renderWindow(laid, start, frames);
      if (rendered) {
        const channels = rendered.numberOfChannels;
        await encoder.encode(frames, (index) => rendered.getChannelData(Math.min(index, channels - 1)));
      } else {
        // A window where nothing plays is silence, still on the track's clock.
        if (!silence || silence.length < frames) silence = new Float32Array(frames);
        const quiet = silence.subarray(0, frames);
        await encoder.encode(frames, () => quiet);
      }
    }
  } catch (error) {
    encoder.abort();
    throw error;
  }
  const description = await encoder.finish();
  if (chunks.length === 0) return null;
  
  const primingSamples = choice.codec === "aac" ? await measureAacPriming(choice.config) : 0;
  // A chunk the encoder states no length for is a frame of its codec's own size.
  const frameUs = choice.codec === "aac" ? (1024 / EXPORT_AUDIO_SAMPLE_RATE) * 1_000_000 : 20_000;
  return {
    codec: choice.codec,
    sampleRate: EXPORT_AUDIO_SAMPLE_RATE,
    channels: EXPORT_AUDIO_CHANNELS,
    ...(description ? { description } : {}),
    primingSamples,
    playedSamples: total,
    chunks: chunks.map((chunk) => ({ size: chunk.size, durationUs: chunk.durationUs > 0 ? chunk.durationUs : frameUs })),
    payload: await spool.seal(),
  };
}
