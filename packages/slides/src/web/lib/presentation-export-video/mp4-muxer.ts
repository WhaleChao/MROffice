/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { EncodedExportAudio } from "./export-audio";
import { createBlobSpool, type ByteSpool } from "./byte-spool";

export type Mp4VideoCodec = "avc" | "vp9";

export type Mp4MuxerOptions = {
  readonly codec: Mp4VideoCodec;
  readonly width: number;
  readonly height: number;
  /** The encoder's `decoderConfig.description` (the `avcC` record) for H.264. */
  readonly description?: Uint8Array;
  /** VP9's level as its codec string states it (`vp09.00.<level>.08`), 10 = 1.0. */
  readonly vp9Level?: number;
};

type Sample = { readonly size: number; readonly durationUs: number; readonly key: boolean };

/** Media time units per second: microseconds stated in 90 kHz ticks. */
const TIMESCALE = 90_000;
const MOVIE_TIMESCALE = 1_000;

const encoder = new TextEncoder();

function u8(value: number): Uint8Array {
  return Uint8Array.of(value & 0xff);
}

function u16(value: number): Uint8Array {
  return Uint8Array.of((value >>> 8) & 0xff, value & 0xff);
}

function u32(value: number): Uint8Array {
  return Uint8Array.of((value >>> 24) & 0xff, (value >>> 16) & 0xff, (value >>> 8) & 0xff, value & 0xff);
}

/** Big-endian 32-bit values, one after another (a table's rows). */
function u32s(values: readonly number[]): Uint8Array {
  const out = new Uint8Array(values.length * 4);
  const view = new DataView(out.buffer);
  values.forEach((value, index) => view.setUint32(index * 4, value >>> 0));
  return out;
}

function ascii(text: string): Uint8Array {
  return encoder.encode(text);
}

function concat(parts: readonly Uint8Array[]): Uint8Array {
  let length = 0;
  for (const part of parts) length += part.length;
  const out = new Uint8Array(length);
  let at = 0;
  for (const part of parts) {
    out.set(part, at);
    at += part.length;
  }
  return out;
}

function box(type: string, parts: readonly Uint8Array[]): Uint8Array {
  const body = concat(parts);
  return concat([u32(body.length + 8), ascii(type), body]);
}

/** A full box: version and flags before the body. */
function fullBox(type: string, version: number, flags: number, parts: readonly Uint8Array[]): Uint8Array {
  return box(type, [u8(version), u8(flags >>> 16), u16(flags & 0xffff), ...parts]);
}

const zeros = (count: number) => new Uint8Array(count);

/** The unity matrix of a header, 16.16 / 2.30 fixed point. */
const MATRIX = concat([u32(0x00010000), u32(0), u32(0), u32(0), u32(0x00010000), u32(0), u32(0), u32(0), u32(0x40000000)]);

export class Mp4Muxer {
  private readonly options: Mp4MuxerOptions;
  private readonly spool: ByteSpool;
  private readonly samples: Sample[] = [];

  constructor(options: Mp4MuxerOptions, spool: ByteSpool = createBlobSpool()) {
    if (options.codec === "avc" && !options.description) {
      throw new Error("an H.264 track needs its avcC record");
    }
    this.options = options;
    this.spool = spool;
  }

  /** One encoded frame, in decode order, lasting `durationUs`. */
  addFrame(data: Uint8Array, durationUs: number, key: boolean): void {
    if (this.samples.length === 0 && !key) throw new Error("the first frame must be a key frame");
    this.spool.append(data);
    this.samples.push({ size: data.length, durationUs: Math.max(1, Math.round(durationUs)), key });
  }

  get frameCount(): number {
    return this.samples.length;
  }

  /** Resolves once the frames so far are (nearly) all written out. */
  drained(): Promise<void> {
    return this.spool.drained();
  }

  /** The whole file, with the show's sound when it has any. */
  async finalize(audio?: EncodedExportAudio | null): Promise<Blob> {
    if (this.samples.length === 0) throw new Error("no frames");
    const sound = audio && audio.chunks.length > 0 ? audio : null;
    const ftyp = box("ftyp", [ascii("isom"), u32(0x200), ascii("isom"), ascii("iso2"), ascii(this.options.codec === "avc" ? "avc1" : "vp09"), ascii("mp41")]);
    const payload = await this.spool.seal();
    const soundPayload = sound ? sound.payload : null;
    const payloadBytes = this.spool.size;
    const soundBytes = sound ? sound.chunks.reduce((sum, chunk) => sum + chunk.size, 0) : 0;
    const mdatBody = payloadBytes + soundBytes;
    // Past 4 GiB: a 64-bit size (size 1, then the size after the type).
    const large = mdatBody + 8 > 0xffffffff;
    const mdatHeader = large
      ? concat([u32(1), ascii("mdat"), u64(mdatBody + 16)])
      : concat([u32(mdatBody + 8), ascii("mdat")]);
    // The chunk offsets depend on moov's size, which depends on the offsets
    // only through whether they need 64 bits: lay moov out with the widths
    // the true offsets need, then for real.
    const wide = ftyp.length + mdatHeader.length + mdatBody + 1024 * 1024 > 0xffffffff;
    const probe = this.moov(0, 0, sound, wide);
    const first = ftyp.length + probe.length + mdatHeader.length;
    const moov = this.moov(first, first + payloadBytes, sound, wide);
    const parts: BlobPart[] = [ftyp as BlobPart, moov as BlobPart, mdatHeader as BlobPart, payload];
    if (soundPayload) parts.push(soundPayload);
    return new Blob(parts, { type: "video/mp4" });
  }

  /** Drop the frames written so far (an export cancelled or failed). */
  discard(): Promise<void> {
    return this.spool.discard();
  }

  private moov(chunkOffset: number, soundOffset: number, sound: EncodedExportAudio | null, wide: boolean): Uint8Array {
    const { width, height } = this.options;
    const mediaDuration = this.samples.reduce((sum, sample) => sum + ticks(sample.durationUs), 0);
    const videoDuration = Math.round((mediaDuration / TIMESCALE) * MOVIE_TIMESCALE);
    const soundTrak = sound ? soundTrack(sound, soundOffset, wide) : null;
    const movieDuration = Math.max(videoDuration, soundTrak?.movieDuration ?? 0);
    const mvhd = fullBox("mvhd", 0, 0, [
      u32(0), u32(0), u32(MOVIE_TIMESCALE), u32(movieDuration),
      u32(0x00010000), u16(0x0100), zeros(10), MATRIX, zeros(24), u32(soundTrak ? 3 : 2),
    ]);
    const tkhd = fullBox("tkhd", 0, 3, [
      u32(0), u32(0), u32(1), u32(0), u32(videoDuration), zeros(8),
      u16(0), u16(0), u16(0), u16(0), MATRIX, u32(width << 16), u32(height << 16),
    ]);
    const mdhd = fullBox("mdhd", 0, 0, [u32(0), u32(0), u32(TIMESCALE), u32(mediaDuration), u16(0x55c4), u16(0)]);
    const hdlr = fullBox("hdlr", 0, 0, [u32(0), ascii("vide"), zeros(12), ascii("VideoHandler"), u8(0)]);
    const vmhd = fullBox("vmhd", 0, 1, [zeros(8)]);
    const dinf = box("dinf", [fullBox("dref", 0, 0, [u32(1), fullBox("url ", 0, 1, [])])]);
    const stbl = box("stbl", [
      fullBox("stsd", 0, 0, [u32(1), this.sampleEntry()]),
      this.stts(),
      this.stss(),
      fullBox("stsc", 0, 0, [u32(1), u32(1), u32(this.samples.length), u32(1)]),
      fullBox("stsz", 0, 0, [u32(0), u32(this.samples.length), u32s(this.samples.map((sample) => sample.size))]),
      chunkOffsets(chunkOffset, wide),
    ]);
    const minf = box("minf", [vmhd, dinf, stbl]);
    const mdia = box("mdia", [mdhd, hdlr, minf]);
    return box("moov", [mvhd, box("trak", [tkhd, mdia]), ...(soundTrak ? [soundTrak.trak] : [])]);
  }

  private sampleEntry(): Uint8Array {
    const { width, height, codec, description } = this.options;
    const visual = [
      zeros(6), u16(1), zeros(16), u16(width), u16(height),
      u32(0x00480000), u32(0x00480000), u32(0), u16(1), zeros(32), u16(0x0018), u16(0xffff),
    ];
    if (codec === "avc") return box("avc1", [...visual, box("avcC", [description!])]);
    // vpcC v1: profile 0, the encoder's level, 8-bit 4:2:0 colocated,
    // BT.709 primaries / transfer / matrix, limited range, no init data.
    const level = this.options.vp9Level ?? 10;
    const vpcC = fullBox("vpcC", 1, 0, [u8(0), u8(level), u8((8 << 4) | (1 << 1) | 0), u8(1), u8(1), u8(1), u16(0)]);
    return box("vp09", [...visual, vpcC]);
  }

  /** Decoding times as runs of equal durations. */
  private stts(): Uint8Array {
    const runs: Array<[count: number, delta: number]> = [];
    for (const sample of this.samples) {
      const delta = ticks(sample.durationUs);
      const last = runs[runs.length - 1];
      if (last && last[1] === delta) last[0] += 1;
      else runs.push([1, delta]);
    }
    return fullBox("stts", 0, 0, [u32(runs.length), u32s(runs.flat())]);
  }

  /** The key frames, numbered from 1. */
  private stss(): Uint8Array {
    const keys: number[] = [];
    this.samples.forEach((sample, index) => {
      if (sample.key) keys.push(index + 1);
    });
    return fullBox("stss", 0, 0, [u32(keys.length), u32s(keys)]);
  }
}

/** A track's one chunk offset: `stco`, or `co64` in a file past 4 GiB. */
function chunkOffsets(offset: number, wide: boolean): Uint8Array {
  return wide ? fullBox("co64", 0, 0, [u32(1), u64(offset)]) : fullBox("stco", 0, 0, [u32(1), u32(offset)]);
}

function u64(value: number): Uint8Array {
  const high = Math.floor(value / 0x1_0000_0000);
  return concat([u32(high), u32(value - high * 0x1_0000_0000)]);
}

function ticks(durationUs: number): number {
  return Math.max(1, Math.round((durationUs * TIMESCALE) / 1_000_000));
}

/** AAC-LC's AudioSpecificConfig for a rate and channel count (ISO/IEC 14496-3 1.6.2.1). */
export function aacAudioSpecificConfig(sampleRate: number, channels: number): Uint8Array {
  const rates = [96000, 88200, 64000, 48000, 44100, 32000, 24000, 22050, 16000, 12000, 11025, 8000, 7350];
  const index = Math.max(0, rates.indexOf(sampleRate));
  const objectType = 2;
  return Uint8Array.of((objectType << 3) | (index >> 1), ((index & 1) << 7) | ((channels & 0x0f) << 3));
}

/** A descriptor of ISO/IEC 14496-1: tag, size, body. */
function descriptor(tag: number, parts: readonly Uint8Array[]): Uint8Array {
  const body = concat(parts);
  if (body.length > 0x0fffffff) throw new Error("descriptor too large");
  const size = Uint8Array.of(
    0x80 | ((body.length >>> 21) & 0x7f),
    0x80 | ((body.length >>> 14) & 0x7f),
    0x80 | ((body.length >>> 7) & 0x7f),
    body.length & 0x7f,
  );
  return concat([u8(tag), size, body]);
}

/** `dOps` from the encoder's OpusHead (little-endian fields there, big-endian here). */
function opusSpecificBox(sound: EncodedExportAudio): Uint8Array {
  const head = sound.description;
  const isHead = head && head.length >= 19 && new TextDecoder().decode(head.subarray(0, 8)) === "OpusHead";
  const view = isHead ? new DataView(head!.buffer, head!.byteOffset, head!.byteLength) : null;
  const channels = view ? view.getUint8(9) : sound.channels;
  const preSkip = view ? view.getUint16(10, true) : 312;
  const inputRate = view ? view.getUint32(12, true) : sound.sampleRate;
  const gain = view ? view.getInt16(16, true) : 0;
  return box("dOps", [u8(0), u8(channels), u16(preSkip), u32(inputRate), u16(gain & 0xffff), u8(0)]);
}

function soundTrack(sound: EncodedExportAudio, chunkOffset: number, wide: boolean): { trak: Uint8Array; movieDuration: number } {
  const rate = sound.sampleRate;
  const deltas = sound.chunks.map((chunk) => Math.max(1, Math.round((chunk.durationUs * rate) / 1_000_000)));
  const mediaDuration = deltas.reduce((sum, delta) => sum + delta, 0);
  // The encoder's priming is skipped and its padding cut by the edit list
  // (ISO/IEC 14496-12 8.6.6): the sound starts on the file's first frame.
  const priming = Math.max(0, Math.min(mediaDuration - 1, Math.round(sound.primingSamples ?? 0)));
  const played = Math.max(1, Math.min(mediaDuration - priming, Math.round(sound.playedSamples ?? mediaDuration - priming)));
  const movieDuration = Math.round((played / rate) * MOVIE_TIMESCALE);
  const edts = priming > 0 || played < mediaDuration
    ? box("edts", [fullBox("elst", 0, 0, [u32(1), u32(movieDuration), u32(priming), u16(1), u16(0)])])
    : null;
  const tkhd = fullBox("tkhd", 0, 3, [
    u32(0), u32(0), u32(2), u32(0), u32(movieDuration), zeros(8),
    u16(0), u16(1), u16(0x0100), u16(0), MATRIX, u32(0), u32(0),
  ]);
  const mdhd = fullBox("mdhd", 0, 0, [u32(0), u32(0), u32(rate), u32(mediaDuration), u16(0x55c4), u16(0)]);
  const hdlr = fullBox("hdlr", 0, 0, [u32(0), ascii("soun"), zeros(12), ascii("SoundHandler"), u8(0)]);
  const smhd = fullBox("smhd", 0, 0, [u16(0), u16(0)]);
  const dinf = box("dinf", [fullBox("dref", 0, 0, [u32(1), fullBox("url ", 0, 1, [])])]);
  const audioEntry = [zeros(6), u16(1), zeros(8), u16(sound.channels), u16(16), u16(0), u16(0), u32(rate << 16 >>> 0)];
  let entry: Uint8Array;
  if (sound.codec === "aac") {
    const config = sound.description && sound.description.length > 0 ? sound.description : aacAudioSpecificConfig(rate, sound.channels);
    const average = Math.round(sound.chunks.reduce((sum, chunk) => sum + chunk.size, 0) * 8 / Math.max(1e-6, mediaDuration / rate));
    const esds = fullBox("esds", 0, 0, [
      descriptor(0x03, [
        u16(2), u8(0),
        descriptor(0x04, [u8(0x40), u8(0x15), Uint8Array.of(0, 0, 0), u32(average), u32(average), descriptor(0x05, [config])]),
        descriptor(0x06, [u8(0x02)]),
      ]),
    ]);
    entry = box("mp4a", [...audioEntry, esds]);
  } else {
    entry = box("Opus", [...audioEntry, opusSpecificBox(sound)]);
  }
  const runs: Array<[number, number]> = [];
  for (const delta of deltas) {
    const last = runs[runs.length - 1];
    if (last && last[1] === delta) last[0] += 1;
    else runs.push([1, delta]);
  }
  const stbl = box("stbl", [
    fullBox("stsd", 0, 0, [u32(1), entry]),
    fullBox("stts", 0, 0, [u32(runs.length), u32s(runs.flat())]),
    fullBox("stsc", 0, 0, [u32(1), u32(1), u32(sound.chunks.length), u32(1)]),
    fullBox("stsz", 0, 0, [u32(0), u32(sound.chunks.length), u32s(sound.chunks.map((chunk) => chunk.size))]),
    chunkOffsets(chunkOffset, wide),
  ]);
  const minf = box("minf", [smhd, dinf, stbl]);
  const mdia = box("mdia", [mdhd, hdlr, minf]);
  return { trak: box("trak", edts ? [tkhd, edts, mdia] : [tkhd, mdia]), movieDuration };
}
