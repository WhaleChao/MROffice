/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Transitions > Sound's built-in list (E037): desktop PowerPoint offers a set
 * of named sound effects above "Other Sound...", each embedded in the file as
 * a WAV the moment it is chosen. The names are PowerPoint's; the audio is our
 * own, synthesized here from a seeded generator, so a sound is the same bytes
 * on every machine and nothing is fetched. A chosen sound is an ordinary
 * embedded-WAV transition sound (`PresentationTransitionSound`), named with
 * its English name so PowerPoint and our importer read the same name back.
 */

import type { InlinePresentationTransitionSound } from "./presentation-transitions";

export const PRESENTATION_BUILTIN_TRANSITION_SOUNDS = [
  { id: "applause", name: "Applause" },
  { id: "arrow", name: "Arrow" },
  { id: "bomb", name: "Bomb" },
  { id: "breeze", name: "Breeze" },
  { id: "camera", name: "Camera" },
  { id: "cash-register", name: "Cash Register" },
  { id: "chime", name: "Chime" },
  { id: "click", name: "Click" },
  { id: "coin", name: "Coin" },
  { id: "drum-roll", name: "Drum Roll" },
  { id: "explosion", name: "Explosion" },
  { id: "hammer", name: "Hammer" },
  { id: "kiss", name: "Kiss" },
  { id: "laser", name: "Laser" },
  { id: "push", name: "Push" },
  { id: "suction", name: "Suction" },
  { id: "typewriter", name: "Typewriter" },
  { id: "voltage", name: "Voltage" },
  { id: "whoosh", name: "Whoosh" },
  { id: "wind", name: "Wind" },
] as const;

/**
 * Sounds beyond PowerPoint's list, offered under "More Sounds": synthesized
 * the same way and embedded the same way (a named WAV), so PowerPoint plays
 * them like any other embedded transition sound. Their names are none of
 * PowerPoint's, so a file's own built-in sound is never read as one of these.
 */
export const PRESENTATION_EXTRA_TRANSITION_SOUNDS = [
  { id: "doorbell", name: "Doorbell" },
  { id: "gong", name: "Gong" },
  { id: "bell-toll", name: "Bell Toll" },
  { id: "fanfare", name: "Fanfare" },
  { id: "magic", name: "Magic" },
  { id: "bubbles", name: "Bubbles" },
  { id: "water-drop", name: "Water Drop" },
  { id: "page-turn", name: "Page Turn" },
  { id: "heartbeat", name: "Heartbeat" },
  { id: "clock-tick", name: "Clock Tick" },
  { id: "thunder", name: "Thunder" },
  { id: "rain", name: "Rain" },
  { id: "boing", name: "Boing" },
  { id: "whistle", name: "Whistle" },
  { id: "success", name: "Success" },
  { id: "countdown", name: "Countdown" },
  { id: "harp", name: "Harp" },
  { id: "xylophone", name: "Xylophone" },
  { id: "sleigh-bells", name: "Sleigh Bells" },
  { id: "riser", name: "Riser" },
] as const;

const ALL_TRANSITION_SOUNDS = [
  ...PRESENTATION_BUILTIN_TRANSITION_SOUNDS,
  ...PRESENTATION_EXTRA_TRANSITION_SOUNDS,
] as const;

export type PresentationBuiltinTransitionSoundId =
  (typeof ALL_TRANSITION_SOUNDS)[number]["id"];

const SAMPLE_RATE = 22_050;

/**
 * The built-in sound a stored name stands for: our English name, or the
 * file name PowerPoint writes for it ("applause.wav", "Cash Register.wav").
 */
export function builtinTransitionSoundIdForName(
  name: string | null | undefined,
): PresentationBuiltinTransitionSoundId | null {
  if (!name) return null;
  const key = name.trim().replace(/\.wav$/i, "").replace(/[\s_-]+/g, "").toLowerCase();
  const match = ALL_TRANSITION_SOUNDS.find(
    (sound) => sound.name.replace(/\s+/g, "").toLowerCase() === key,
  );
  return match?.id ?? null;
}

function mulberry32(seed: number): () => number {
  let state = seed >>> 0;
  return () => {
    state = (state + 0x6d2b79f5) >>> 0;
    let t = state;
    t = Math.imul(t ^ (t >>> 15), t | 1);
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

/** A two-pole band-pass whose centre may move with time (RBJ cookbook). */
function bandPass(
  input: Float32Array,
  centreAt: (seconds: number) => number,
  q: number,
): Float32Array {
  const out = new Float32Array(input.length);
  let x1 = 0;
  let x2 = 0;
  let y1 = 0;
  let y2 = 0;
  let b0 = 0;
  let b2 = 0;
  let a1 = 0;
  let a2 = 0;
  for (let i = 0; i < input.length; i += 1) {
    if (i % 32 === 0) {
      const w = (2 * Math.PI * Math.min(centreAt(i / SAMPLE_RATE), SAMPLE_RATE * 0.45)) / SAMPLE_RATE;
      const alpha = Math.sin(w) / (2 * q);
      const a0 = 1 + alpha;
      b0 = alpha / a0;
      b2 = -alpha / a0;
      a1 = (-2 * Math.cos(w)) / a0;
      a2 = (1 - alpha) / a0;
    }
    const x = input[i];
    const y = b0 * x + b2 * x2 - a1 * y1 - a2 * y2;
    x2 = x1;
    x1 = x;
    y2 = y1;
    y1 = y;
    out[i] = y;
  }
  return out;
}

function lowPass(input: Float32Array, cutoffAt: (seconds: number) => number): Float32Array {
  const out = new Float32Array(input.length);
  let y = 0;
  for (let i = 0; i < input.length; i += 1) {
    const rc = 1 / (2 * Math.PI * cutoffAt(i / SAMPLE_RATE));
    const alpha = 1 / (1 + rc * SAMPLE_RATE);
    y += alpha * (input[i] - y);
    out[i] = y;
  }
  return out;
}

class Track {
  readonly data: Float32Array;
  readonly random: () => number;

  constructor(seconds: number, seed: number) {
    this.data = new Float32Array(Math.round(seconds * SAMPLE_RATE));
    this.random = mulberry32(seed);
  }

  noise(seconds: number): Float32Array {
    const out = new Float32Array(Math.max(1, Math.round(seconds * SAMPLE_RATE)));
    for (let i = 0; i < out.length; i += 1) out[i] = this.random() * 2 - 1;
    return out;
  }

  /** Mix `source` in at `at` seconds, shaped by `envelope(t, length)`. */
  mix(
    source: Float32Array,
    at: number,
    gain: number,
    envelope: (t: number, length: number) => number = () => 1,
  ): void {
    const start = Math.round(at * SAMPLE_RATE);
    const length = source.length / SAMPLE_RATE;
    for (let i = 0; i < source.length && start + i < this.data.length; i += 1) {
      if (start + i < 0) continue;
      this.data[start + i] += source[i] * gain * envelope(i / SAMPLE_RATE, length);
    }
  }

  /** A tone whose frequency follows `frequencyAt`, summed with its partials. */
  tone(
    seconds: number,
    frequencyAt: (t: number) => number,
    partials: readonly (readonly [ratio: number, gain: number, decay: number])[] = [[1, 1, 0]],
    shape: "sine" | "square" | "saw" = "sine",
  ): Float32Array {
    const out = new Float32Array(Math.max(1, Math.round(seconds * SAMPLE_RATE)));
    const phases = partials.map(() => 0);
    for (let i = 0; i < out.length; i += 1) {
      const t = i / SAMPLE_RATE;
      const base = frequencyAt(t);
      let value = 0;
      partials.forEach(([ratio, gain, decay], index) => {
        phases[index] += (2 * Math.PI * base * ratio) / SAMPLE_RATE;
        const phase = phases[index];
        const wave =
          shape === "sine"
            ? Math.sin(phase)
            : shape === "square"
              ? Math.sign(Math.sin(phase)) * 0.6
              : ((phase / Math.PI) % 2) - 1;
        value += wave * gain * Math.exp(-decay * t);
      });
      out[i] = value;
    }
    return out;
  }
}

const decayEnvelope =
  (attack: number, decay: number) =>
  (t: number): number =>
    (t < attack ? t / attack : 1) * Math.exp(-decay * Math.max(0, t - attack));

const swellEnvelope = (t: number, length: number): number => Math.sin((Math.PI * t) / length) ** 1.5;

const fadeEnvelope =
  (fadeIn: number, fadeOut: number) =>
  (t: number, length: number): number =>
    Math.min(1, t / fadeIn, Math.max(0, (length - t) / fadeOut));

/** A bell strike: inharmonic partials that die away at their own rates. */
function bell(track: Track, at: number, frequency: number, gain: number, seconds = 1.2): void {
  track.mix(
    track.tone(seconds, () => frequency, [
      [1, 1, 3],
      [2.76, 0.45, 5],
      [5.4, 0.25, 8],
      [8.93, 0.12, 12],
    ]),
    at,
    gain,
    decayEnvelope(0.002, 0),
  );
}

function click(track: Track, at: number, gain: number, centre = 3000, seconds = 0.012): void {
  track.mix(bandPass(track.noise(seconds), () => centre, 1.2), at, gain, decayEnvelope(0.0005, 6 / seconds));
}

function noiseBurst(
  track: Track,
  at: number,
  seconds: number,
  gain: number,
  cutoffAt: (t: number) => number,
  envelope: (t: number, length: number) => number,
): void {
  track.mix(lowPass(track.noise(seconds), cutoffAt), at, gain, envelope);
}

const SYNTHESIZERS: Record<PresentationBuiltinTransitionSoundId, () => Track> = {
  applause: () => {
    const track = new Track(2.2, 11);
    for (let at = 0; at < 2.05; at += 0.012 + track.random() * 0.03) {
      const centre = 1100 + track.random() * 1600;
      track.mix(
        bandPass(track.noise(0.03), () => centre, 1.4),
        at,
        0.5 + track.random() * 0.5,
        decayEnvelope(0.001, 140),
      );
    }
    return track;
  },
  arrow: () => {
    const track = new Track(0.75, 12);
    track.mix(
      bandPass(track.noise(0.4), (t) => 3200 - t * 5000, 2.5),
      0,
      1.4,
      swellEnvelope,
    );
    track.mix(track.tone(0.3, (t) => 140 - t * 120), 0.4, 0.9, decayEnvelope(0.002, 18));
    click(track, 0.4, 0.8, 1800, 0.02);
    return track;
  },
  bomb: () => {
    const track = new Track(2.1, 13);
    track.mix(
      track.tone(1.2, (t) => 1700 * Math.exp(-t * 1.2)),
      0,
      0.35,
      fadeEnvelope(0.15, 0.05),
    );
    noiseBurst(track, 1.2, 0.9, 1.2, (t) => 1800 * Math.exp(-t * 3) + 120, decayEnvelope(0.004, 4.5));
    track.mix(track.tone(0.9, (t) => 55 - t * 20), 1.2, 0.8, decayEnvelope(0.004, 4));
    return track;
  },
  breeze: () => {
    const track = new Track(2.2, 14);
    const wobble = track.random() * Math.PI;
    noiseBurst(
      track,
      0,
      2.2,
      2.4,
      (t) => 650 + 300 * Math.sin(t * 2.3 + wobble),
      (t, length) => swellEnvelope(t, length) * (0.75 + 0.25 * Math.sin(t * 5.1)),
    );
    return track;
  },
  camera: () => {
    const track = new Track(0.4, 15);
    click(track, 0, 1, 4200, 0.015);
    track.mix(track.tone(0.06, () => 900), 0.01, 0.15, decayEnvelope(0.001, 60));
    click(track, 0.13, 0.85, 3000, 0.02);
    return track;
  },
  "cash-register": () => {
    const track = new Track(1.4, 16);
    for (const at of [0, 0.07, 0.14]) click(track, at, 0.7, 2200, 0.018);
    noiseBurst(track, 0.18, 0.35, 0.8, () => 500, fadeEnvelope(0.02, 0.15));
    bell(track, 0.22, 2093, 0.55, 1.15);
    bell(track, 0.22, 2637, 0.3, 1.15);
    return track;
  },
  chime: () => {
    const track = new Track(2, 17);
    bell(track, 0, 1319, 0.5, 1.6);
    bell(track, 0.22, 1047, 0.45, 1.5);
    bell(track, 0.44, 784, 0.5, 1.5);
    return track;
  },
  click: () => {
    const track = new Track(0.08, 18);
    click(track, 0, 1, 2600, 0.01);
    return track;
  },
  coin: () => {
    const track = new Track(0.55, 19);
    track.mix(track.tone(0.08, () => 988, [[1, 1, 0]], "square"), 0, 0.5);
    track.mix(track.tone(0.45, () => 1319, [[1, 1, 0]], "square"), 0.08, 0.5, decayEnvelope(0.002, 7));
    return track;
  },
  "drum-roll": () => {
    const track = new Track(2, 20);
    let hit = 0;
    for (let at = 0; at < 1.6; at += 0.034, hit += 1) {
      const level = 0.25 + 0.55 * (at / 1.6) + (hit % 2 ? 0.06 : 0);
      track.mix(bandPass(track.noise(0.05), () => 2200, 0.9), at, level, decayEnvelope(0.001, 70));
      track.mix(track.tone(0.05, () => 190), at, level * 0.4, decayEnvelope(0.001, 60));
    }
    track.mix(bandPass(track.noise(0.35), () => 1800, 0.8), 1.62, 1.1, decayEnvelope(0.001, 12));
    track.mix(track.tone(0.35, () => 70), 1.62, 0.9, decayEnvelope(0.001, 10));
    return track;
  },
  explosion: () => {
    const track = new Track(1.8, 21);
    noiseBurst(track, 0, 1.8, 1.6, (t) => 2200 * Math.exp(-t * 2.5) + 90, decayEnvelope(0.003, 2.6));
    track.mix(track.tone(1.6, (t) => 48 - t * 14), 0, 0.9, decayEnvelope(0.005, 2.4));
    return track;
  },
  hammer: () => {
    const track = new Track(0.75, 22);
    for (const at of [0, 0.32]) {
      track.mix(track.tone(0.25, () => 95), at, 0.8, decayEnvelope(0.001, 22));
      track.mix(
        track.tone(0.3, () => 1210, [
          [1, 0.6, 18],
          [1.93, 0.35, 24],
          [3.07, 0.2, 30],
        ]),
        at,
        0.6,
      );
      click(track, at, 0.6, 3500, 0.01);
    }
    return track;
  },
  kiss: () => {
    const track = new Track(0.3, 23);
    track.mix(bandPass(track.noise(0.08), (t) => 900 + t * 9000, 3), 0, 0.5, swellEnvelope);
    track.mix(bandPass(track.noise(0.025), () => 2400, 2), 0.08, 1.2, decayEnvelope(0.001, 160));
    track.mix(track.tone(0.03, (t) => 1500 - t * 20000), 0.08, 0.3, decayEnvelope(0.001, 120));
    return track;
  },
  laser: () => {
    const track = new Track(0.75, 24);
    for (const at of [0, 0.33]) {
      track.mix(
        track.tone(0.32, (t) => 2600 * Math.exp(-t * 6) + 220 + 40 * Math.sin(t * 160), [[1, 1, 0], [2, 0.3, 0]], "square"),
        at,
        0.45,
        decayEnvelope(0.002, 6),
      );
    }
    return track;
  },
  push: () => {
    const track = new Track(0.6, 25);
    track.mix(bandPass(track.noise(0.42), (t) => 500 + t * 2600, 1.8), 0, 1.2, swellEnvelope);
    track.mix(track.tone(0.18, () => 110), 0.36, 0.6, decayEnvelope(0.002, 20));
    return track;
  },
  suction: () => {
    const track = new Track(0.9, 26);
    track.mix(bandPass(track.noise(0.75), (t) => 300 + t * 3200, 4), 0, 1.6, fadeEnvelope(0.08, 0.04));
    track.mix(track.tone(0.06, (t) => 600 + t * 9000), 0.76, 0.5, decayEnvelope(0.001, 70));
    click(track, 0.78, 0.5, 1500, 0.015);
    return track;
  },
  typewriter: () => {
    const track = new Track(1.7, 27);
    for (let at = 0; at < 1.15; at += 0.07 + track.random() * 0.07) {
      click(track, at, 0.6 + track.random() * 0.4, 1800 + track.random() * 1500, 0.014);
      track.mix(track.tone(0.03, () => 420), at, 0.12, decayEnvelope(0.001, 90));
    }
    bell(track, 1.22, 2794, 0.45, 0.45);
    return track;
  },
  voltage: () => {
    const track = new Track(1.3, 28);
    track.mix(
      track.tone(1.3, () => 120, [[1, 1, 0], [3, 0.3, 0]], "saw"),
      0,
      0.3,
      (t, length) => fadeEnvelope(0.03, 0.15)(t, length) * (0.6 + 0.4 * Math.abs(Math.sin(t * 37))),
    );
    for (let at = 0; at < 1.2; at += 0.03 + track.random() * 0.12) {
      click(track, at, 0.3 + track.random() * 0.5, 4500, 0.006);
    }
    return track;
  },
  whoosh: () => {
    const track = new Track(0.8, 29);
    track.mix(
      bandPass(track.noise(0.8), (t) => 400 + 2400 * Math.sin((Math.PI * t) / 0.8), 1.6),
      0,
      1.8,
      swellEnvelope,
    );
    return track;
  },
  wind: () => {
    const track = new Track(2.4, 30);
    const howl = bandPass(track.noise(2.4), (t) => 520 + 180 * Math.sin(t * 1.7) + 90 * Math.sin(t * 4.3), 7);
    track.mix(howl, 0, 3.2, swellEnvelope);
    noiseBurst(track, 0, 2.4, 1.2, (t) => 400 + 200 * Math.sin(t * 2.9), swellEnvelope);
    return track;
  },
  // ---- More Sounds (beyond PowerPoint's list) ----
  doorbell: () => {
    const track = new Track(1.9, 31);
    const strike = (at: number, frequency: number) =>
      track.mix(
        track.tone(1.4, () => frequency, [
          [1, 1, 1.6],
          [2, 0.22, 2.8],
          [3, 0.08, 4],
        ]),
        at,
        0.7,
        decayEnvelope(0.003, 0),
      );
    strike(0, 659);
    strike(0.5, 523);
    return track;
  },
  gong: () => {
    const track = new Track(2.6, 32);
    track.mix(
      track.tone(2.6, (t) => 98 * (1 + 0.004 * Math.sin(t * 5.5)), [
        [1, 1, 0.9],
        [1.48, 0.7, 1.1],
        [2.12, 0.55, 1.4],
        [2.69, 0.4, 1.8],
        [3.43, 0.3, 2.2],
        [4.17, 0.22, 2.7],
        [5.02, 0.14, 3.2],
      ]),
      0,
      0.6,
      decayEnvelope(0.012, 0),
    );
    noiseBurst(track, 0, 0.5, 0.5, (t) => 1400 * Math.exp(-t * 6) + 200, decayEnvelope(0.002, 9));
    return track;
  },
  "bell-toll": () => {
    const track = new Track(2.6, 33);
    track.mix(
      track.tone(2.6, () => 392, [
        [0.5, 0.45, 0.7],
        [1, 1, 1.1],
        [1.19, 0.55, 1.4],
        [1.5, 0.4, 1.8],
        [2, 0.32, 2.3],
        [2.52, 0.18, 2.9],
        [3.01, 0.12, 3.4],
      ]),
      0,
      0.6,
      decayEnvelope(0.002, 0),
    );
    click(track, 0, 0.35, 2600, 0.02);
    return track;
  },
  fanfare: () => {
    const track = new Track(1.6, 34);
    const notes: readonly (readonly [at: number, frequency: number, seconds: number])[] = [
      [0, 523, 0.14],
      [0.15, 659, 0.14],
      [0.3, 784, 0.14],
      [0.45, 1047, 0.95],
    ];
    for (const [at, frequency, seconds] of notes) {
      const brass = lowPass(
        track.tone(seconds, (t) => frequency * (1 + (seconds > 0.5 ? 0.006 * Math.sin(t * 34) : 0)), [
          [1, 1, 0],
          [2, 0.5, 0],
          [3, 0.3, 0],
        ], "saw"),
        () => 2600,
      );
      track.mix(brass, at, 0.55, fadeEnvelope(0.025, seconds > 0.5 ? 0.35 : 0.04));
    }
    return track;
  },
  magic: () => {
    const track = new Track(1.8, 35);
    const scale = [1047, 1175, 1319, 1568, 1760, 2093, 2349, 2637, 3136, 3520, 4186, 4699];
    scale.forEach((frequency, index) => bell(track, index * 0.055, frequency, 0.32 - index * 0.012, 0.9));
    track.mix(
      bandPass(track.noise(1.2), (t) => 5000 + t * 3000, 3),
      0.05,
      0.5,
      (t, length) => swellEnvelope(t, length) * (0.6 + 0.4 * Math.sin(t * 60)),
    );
    return track;
  },
  bubbles: () => {
    const track = new Track(0.75, 36);
    const pops: readonly (readonly [at: number, base: number])[] = [
      [0, 420],
      [0.13, 560],
      [0.22, 380],
      [0.36, 700],
    ];
    for (const [at, base] of pops) {
      track.mix(
        track.tone(0.09, (t) => base + t * 14000),
        at,
        0.8,
        decayEnvelope(0.002, 45),
      );
    }
    return track;
  },
  "water-drop": () => {
    const track = new Track(0.8, 37);
    for (const [at, gain] of [[0, 0.9], [0.18, 0.35], [0.34, 0.15]] as const) {
      track.mix(
        track.tone(0.25, (t) => 900 + 1600 * (1 - Math.exp(-t * 60))),
        at,
        gain,
        decayEnvelope(0.001, 22),
      );
    }
    return track;
  },
  "page-turn": () => {
    const track = new Track(0.7, 38);
    track.mix(
      bandPass(track.noise(0.5), (t) => 4200 - t * 3600, 1.1),
      0,
      1.2,
      (t, length) => swellEnvelope(t, length) * (0.55 + 0.45 * Math.abs(Math.sin(t * 95))),
    );
    click(track, 0.47, 0.45, 2200, 0.03);
    return track;
  },
  heartbeat: () => {
    const track = new Track(1.6, 39);
    for (const at of [0, 0.82]) {
      for (const [offset, gain] of [[0, 1], [0.19, 0.7]] as const) {
        track.mix(track.tone(0.22, (t) => 62 - t * 40), at + offset, gain, decayEnvelope(0.006, 18));
        noiseBurst(track, at + offset, 0.08, 0.35 * gain, () => 180, decayEnvelope(0.002, 40));
      }
    }
    return track;
  },
  "clock-tick": () => {
    const track = new Track(2, 40);
    [0, 0.5, 1, 1.5].forEach((at, index) => {
      const pitch = index % 2 ? 2400 : 3400;
      click(track, at, 0.9, pitch, 0.012);
      track.mix(track.tone(0.05, () => pitch / 2), at, 0.25, decayEnvelope(0.001, 80));
    });
    return track;
  },
  thunder: () => {
    const track = new Track(2.6, 41);
    noiseBurst(track, 0, 0.35, 1, (t) => 5000 * Math.exp(-t * 14) + 400, decayEnvelope(0.002, 9));
    noiseBurst(
      track,
      0.05,
      2.5,
      2.4,
      (t) => 260 * Math.exp(-t * 0.8) + 70,
      (t, length) => decayEnvelope(0.08, 1.4)(t) * (0.65 + 0.35 * Math.sin(t * 7.3) * Math.sin(t * 2.1)) * Math.min(1, (length - t) / 0.3),
    );
    return track;
  },
  rain: () => {
    const track = new Track(2.3, 42);
    noiseBurst(track, 0, 2.3, 0.9, () => 2600, swellEnvelope);
    for (let at = 0.02; at < 2.2; at += 0.004 + track.random() * 0.03) {
      const level = Math.sin((Math.PI * at) / 2.3) ** 1.2;
      click(track, at, (0.2 + track.random() * 0.4) * level, 3500 + track.random() * 4500, 0.006);
    }
    return track;
  },
  boing: () => {
    const track = new Track(0.9, 43);
    track.mix(
      track.tone(0.85, (t) => 260 + 180 * Math.sin(2 * Math.PI * 11 * t) * Math.exp(-t * 3.5) + t * 120, [
        [1, 1, 0],
        [2, 0.25, 0],
      ]),
      0,
      0.8,
      decayEnvelope(0.004, 4.2),
    );
    return track;
  },
  whistle: () => {
    const track = new Track(0.85, 44);
    const trill = (t: number) => 2450 + 140 * Math.sin(2 * Math.PI * 27 * t);
    track.mix(track.tone(0.7, trill, [[1, 1, 0], [2, 0.12, 0]]), 0, 0.55, fadeEnvelope(0.02, 0.08));
    track.mix(bandPass(track.noise(0.7), () => 2500, 5), 0, 0.5, fadeEnvelope(0.02, 0.08));
    return track;
  },
  success: () => {
    const track = new Track(1.1, 45);
    [1047, 1319, 1568, 2093].forEach((frequency, index) => {
      track.mix(
        track.tone(0.7, () => frequency, [
          [1, 1, 6],
          [3.9, 0.22, 16],
          [9.2, 0.06, 28],
        ]),
        index * 0.075,
        index === 3 ? 0.65 : 0.5,
        decayEnvelope(0.002, 0),
      );
    });
    return track;
  },
  countdown: () => {
    const track = new Track(2.2, 46);
    for (const at of [0, 0.5, 1]) {
      track.mix(track.tone(0.13, () => 880, [[1, 1, 0], [3, 0.15, 0]]), at, 0.6, fadeEnvelope(0.005, 0.02));
    }
    track.mix(track.tone(0.5, () => 1760, [[1, 1, 0], [3, 0.15, 0]]), 1.5, 0.6, fadeEnvelope(0.005, 0.12));
    return track;
  },
  harp: () => {
    const track = new Track(2.2, 47);
    const scale = [262, 294, 330, 349, 392, 440, 494, 523, 587, 659, 698, 784, 880, 988, 1047];
    scale.forEach((frequency, index) => {
      track.mix(
        track.tone(1.1, () => frequency, [
          [1, 1, 2.6],
          [2, 0.45, 4.2],
          [3, 0.2, 6],
          [4, 0.08, 8],
        ]),
        index * 0.045,
        0.32,
        decayEnvelope(0.003, 0),
      );
    });
    return track;
  },
  xylophone: () => {
    const track = new Track(1.1, 48);
    ([[0, 784], [0.16, 659], [0.32, 1047]] as const).forEach(([at, frequency]) => {
      track.mix(
        track.tone(0.6, () => frequency, [
          [1, 1, 9],
          [3.93, 0.32, 20],
          [9.24, 0.1, 32],
        ]),
        at,
        0.7,
        decayEnvelope(0.001, 0),
      );
      click(track, at, 0.25, 4500, 0.006);
    });
    return track;
  },
  "sleigh-bells": () => {
    const track = new Track(1.4, 49);
    for (let shake = 0; shake < 8; shake += 1) {
      const at = shake * 0.14 + (shake % 2 ? 0.02 : 0);
      for (let jingle = 0; jingle < 5; jingle += 1) {
        const frequency = 3800 + track.random() * 2600;
        track.mix(
          track.tone(0.18, () => frequency, [[1, 1, 18], [1.6, 0.4, 24]]),
          at + track.random() * 0.03,
          0.18,
          decayEnvelope(0.001, 0),
        );
      }
      track.mix(bandPass(track.noise(0.06), () => 7000, 1.5), at, 0.35, decayEnvelope(0.001, 50));
    }
    return track;
  },
  riser: () => {
    const track = new Track(1.25, 50);
    track.mix(
      bandPass(track.noise(1.15), (t) => 300 * Math.exp(t * 2.6), 2.2),
      0,
      1.4,
      (t, length) => (t / length) ** 1.6 * Math.min(1, (length - t) / 0.02),
    );
    track.mix(
      track.tone(1.15, (t) => 180 * Math.exp(t * 1.6), [[1, 1, 0], [2, 0.3, 0]]),
      0,
      0.4,
      (t, length) => (t / length) ** 2 * Math.min(1, (length - t) / 0.02),
    );
    return track;
  },
};

function wavBytes(samples: Float32Array): Uint8Array {
  let peak = 0;
  for (const value of samples) peak = Math.max(peak, Math.abs(value));
  const scale = peak > 0 ? 0.8 / peak : 0;
  const bytes = new Uint8Array(44 + samples.length * 2);
  const view = new DataView(bytes.buffer);
  const ascii = (offset: number, text: string) => {
    for (let i = 0; i < text.length; i += 1) bytes[offset + i] = text.charCodeAt(i);
  };
  ascii(0, "RIFF");
  view.setUint32(4, 36 + samples.length * 2, true);
  ascii(8, "WAVE");
  ascii(12, "fmt ");
  view.setUint32(16, 16, true);
  view.setUint16(20, 1, true);
  view.setUint16(22, 1, true);
  view.setUint32(24, SAMPLE_RATE, true);
  view.setUint32(28, SAMPLE_RATE * 2, true);
  view.setUint16(32, 2, true);
  view.setUint16(34, 16, true);
  ascii(36, "data");
  view.setUint32(40, samples.length * 2, true);
  samples.forEach((value, index) => {
    view.setInt16(44 + index * 2, Math.round(Math.max(-1, Math.min(1, value * scale)) * 32767), true);
  });
  return bytes;
}

function base64(bytes: Uint8Array): string {
  let binary = "";
  for (let i = 0; i < bytes.length; i += 0x8000) {
    binary += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
  }
  return globalThis.btoa(binary);
}

const cache = new Map<PresentationBuiltinTransitionSoundId, InlinePresentationTransitionSound>();

/** The embedded-WAV transition sound for one built-in entry. */
export function builtinTransitionSound(
  id: PresentationBuiltinTransitionSoundId,
): InlinePresentationTransitionSound {
  const cached = cache.get(id);
  if (cached) return cached;
  const entry = ALL_TRANSITION_SOUNDS.find((sound) => sound.id === id)!;
  const bytes = wavBytes(SYNTHESIZERS[id]().data);
  const sound: InlinePresentationTransitionSound = Object.freeze({
    dataUrl: `data:audio/wav;base64,${base64(bytes)}`,
    mimeType: "audio/wav",
    name: entry.name,
    sizeBytes: bytes.length,
  });
  cache.set(id, sound);
  return sound;
}
