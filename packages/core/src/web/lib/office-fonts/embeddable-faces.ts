/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { officeFontSubstitution } from "./substitutes";
import { loadOfficeDataParts } from "../office-data/store";

/** One pinned WOFF2 slice of a face. */
export interface EmbeddableFontSlice {
  /** CSS unicode-range the slice covers; null for the whole face. */
  readonly unicodeRange: string | null;
  /** Path under /fonts/office/ (and under the server's slices directory). */
  readonly out: string;
  readonly sha256: string;
  readonly size: number;
}

/** One face of a substitute, as the lock serves it. */
export interface EmbeddableFontFace {
  readonly substitute: string;
  readonly weight: number;
  readonly italic: boolean;
  readonly slices: readonly EmbeddableFontSlice[];
}

/** A family as the font list names it, and what stands in for it. */
export interface EmbeddableFontFamily {
  /** The Office family the name belongs to (aliases name their family). */
  readonly office: string;
  /** The substitutes in stack order: a character is set in the first one
   * that has it (a composite family lists one per script). */
  readonly substitutes: readonly string[];
  /** The substitute weight this family's regular style is drawn at. */
  readonly regularWeight: number;
}

/** A face chosen for a style, and what the face itself cannot draw. */
export interface EmbeddableFontFaceChoice {
  readonly face: EmbeddableFontFace;
  /** Bold asked of a face that has no bold weight: drawn stroked. */
  readonly syntheticBold: boolean;
  /** Italic asked of a face that has no italic: drawn leaned. */
  readonly syntheticItalic: boolean;
}

export type EmbeddableFontErrorCode = "unknown-family" | "slice-unavailable" | "slice-mismatch";

export class EmbeddableFontError extends Error {
  constructor(readonly code: EmbeddableFontErrorCode, message?: string) {
    super(message ?? `Embeddable font refused: ${code}.`);
    this.name = "EmbeddableFontError";
  }
}

type ManifestSlice = readonly [number, string, string, number];
type ManifestFamily = readonly [readonly (readonly [number, 0 | 1, number])[], readonly (readonly ManifestSlice[])[]];

export interface EmbeddableFontManifest {
  readonly version: 1;
  /** By index: the whole manifest's array, or the served parts' sparse record. */
  readonly ranges: Readonly<Record<number, string>>;
  readonly families: Readonly<Record<string, ManifestFamily>>;
}

/** Where the browser fetches a slice (the served copy of the lock's `out`). */
export const EMBEDDABLE_FONT_PUBLIC_PREFIX = "/fonts/office/";

/**
 * The manifest holding (at least) these substitutes' faces, asked for once
 * per family; a host that holds the manifest answers it whole. Rejects with
 * {@link EmbeddableFontError} `slice-unavailable` when this page cannot have
 * it (a guest, no connection).
 */
export async function loadEmbeddableFontManifest(
  substitutes: readonly string[] = [],
): Promise<EmbeddableFontManifest> {
  const served = await loadOfficeDataParts("font-embed-manifest", substitutes);
  if (!served || !served.families || typeof served.families !== "object") {
    throw new EmbeddableFontError("slice-unavailable", "The embeddable font list is unavailable.");
  }
  return served as unknown as EmbeddableFontManifest;
}

/**
 * The family a font-list name stands for, or null when MRAI serves nothing
 * for it (a font found on the reader's own machine, an unknown name).
 */
export function embeddableFontFamily(name: string | null | undefined): EmbeddableFontFamily | null {
  const substitution = officeFontSubstitution(name);
  if (!substitution || substitution.substitutes.length === 0) return null;
  return Object.freeze({
    office: substitution.office,
    substitutes: substitution.substitutes,
    regularWeight: substitution.weight ?? 400,
  });
}

/** The faces of one substitute, as the lock serves them. */
export function embeddableFontFaces(manifest: EmbeddableFontManifest, substitute: string): readonly EmbeddableFontFace[] {
  const entry = manifest.families[substitute];
  if (!entry) return [];
  const [faces, sets] = entry;
  return faces.map(([weight, italic, set]) => Object.freeze({
    substitute,
    weight,
    italic: italic === 1,
    slices: Object.freeze((sets[set] ?? []).map(([range, out, sha256, size]) => Object.freeze({
      unicodeRange: range < 0 ? null : manifest.ranges[range] ?? null,
      out,
      sha256,
      size,
    }))),
  }));
}

/**
 * CSS font matching (CSS Fonts 4, 5.2) over the faces of one substitute:
 * the italic faces when italic is asked and there are any, then the weight
 * nearest the one asked in the direction CSS prefers. A bold the chosen face
 * lacks, or an italic, is drawn by the writer as it draws any face's
 * (stroked, leaned) -- what a browser does for the same request.
 */
export function selectEmbeddableFontFace(
  manifest: EmbeddableFontManifest,
  family: EmbeddableFontFamily,
  substitute: string,
  style: Readonly<{ bold: boolean; italic: boolean }>,
): EmbeddableFontFaceChoice | null {
  const faces = embeddableFontFaces(manifest, substitute);
  if (faces.length === 0) return null;
  const italics = faces.filter((face) => face.italic === style.italic);
  const pool = italics.length ? italics : faces;
  const desired = style.bold ? (family.regularWeight >= 600 ? 900 : 700) : family.regularWeight;
  const weights = [...new Set(pool.map((face) => face.weight))].sort((a, b) => a - b);
  const lighter = weights.filter((weight) => weight < desired).reverse();
  const heavier = weights.filter((weight) => weight > desired);
  let order: number[];
  if (weights.includes(desired)) order = [desired];
  else if (desired >= 400 && desired <= 500) {
    const upTo500 = heavier.filter((weight) => weight <= 500);
    order = [...upTo500, ...lighter, ...heavier.filter((weight) => weight > 500)];
  } else if (desired < 400) order = [...lighter, ...heavier];
  else order = [...heavier, ...lighter];
  const weight = order[0];
  const face = pool.find((candidate) => candidate.weight === weight)!;
  return Object.freeze({
    face,
    syntheticBold: style.bold && face.weight < 600,
    syntheticItalic: style.italic && !face.italic,
  });
}

const rangeCache = new Map<string, readonly (readonly [number, number])[]>();

function parseUnicodeRange(range: string): readonly (readonly [number, number])[] {
  let parsed = rangeCache.get(range);
  if (parsed) return parsed;
  parsed = Object.freeze(range.split(",").map((part) => {
    const text = part.trim().replace(/^U\+/iu, "");
    if (text.includes("?")) {
      return Object.freeze([Number.parseInt(text.replace(/\?/gu, "0"), 16), Number.parseInt(text.replace(/\?/gu, "F"), 16)] as const);
    }
    const [start, end] = text.split("-");
    return Object.freeze([Number.parseInt(start, 16), Number.parseInt(end ?? start, 16)] as const);
  }).filter(([start, end]) => Number.isFinite(start) && Number.isFinite(end)));
  rangeCache.set(range, parsed);
  return parsed;
}

/** The slices of a face whose range covers a code point, in the lock's
 * order. A slice's range is a promise of coverage only; the caller still
 * checks the slice's own character map. */
export function embeddableSlicesFor(face: EmbeddableFontFace, codePoint: number): readonly EmbeddableFontSlice[] {
  return face.slices.filter((slice) =>
    slice.unicodeRange === null ||
    parseUnicodeRange(slice.unicodeRange).some(([start, end]) => codePoint >= start && codePoint <= end));
}

async function sha256Hex(bytes: Uint8Array): Promise<string> {
  if (!globalThis.crypto?.subtle) throw new EmbeddableFontError("slice-mismatch", "SHA-256 is unavailable.");
  const owned = new Uint8Array(bytes.byteLength);
  owned.set(bytes);
  const digest = await globalThis.crypto.subtle.digest("SHA-256", owned.buffer);
  return [...new Uint8Array(digest)].map((value) => value.toString(16).padStart(2, "0")).join("");
}

/** The bytes when they are exactly the pinned slice, else null. */
export async function verifyEmbeddableFontSlice(
  slice: EmbeddableFontSlice,
  value: Uint8Array | ArrayBuffer,
): Promise<Uint8Array | null> {
  const bytes = value instanceof Uint8Array ? Uint8Array.from(value) : new Uint8Array(value.slice(0));
  if (bytes.byteLength !== slice.size) return null;
  return (await sha256Hex(bytes)) === slice.sha256 ? bytes : null;
}

const sliceCache = new Map<string, Promise<Uint8Array>>();
const SLICE_FETCH_ATTEMPTS = 3;
const SLICE_FETCH_TIMEOUT_MS = 15_000;
const SLICE_FETCH_RETRY_PAUSE_MS = 1_000;

/** Fetch one slice, verified against the lock (kept for the session: a
 * document edited line by line asks for the same few slices again). */
export function fetchEmbeddableFontSlice(
  slice: EmbeddableFontSlice,
  options?: Readonly<{ fetch?: typeof globalThis.fetch; signal?: AbortSignal }>,
): Promise<Uint8Array> {
  const cached = sliceCache.get(slice.sha256);
  if (cached) return cached;
  const fetcher = options?.fetch ?? globalThis.fetch;
  const loading = (async () => {
    if (typeof fetcher !== "function") throw new EmbeddableFontError("slice-unavailable");
    const unavailable = () => new EmbeddableFontError("slice-unavailable", `The font slice ${slice.out} could not be loaded.`);
    // Each try is given SLICE_FETCH_TIMEOUT_MS and the slice asked again
    // after a pause when the server does not answer or answers 5xx: a fetch
    // left hanging held a box's format, and the words typed with it, for
    
    let response: Response | null = null;
    for (let attempt = 0; attempt < SLICE_FETCH_ATTEMPTS && !response; attempt += 1) {
      if (attempt > 0) await new Promise((resolve) => setTimeout(resolve, SLICE_FETCH_RETRY_PAUSE_MS * attempt));
      if (options?.signal?.aborted) throw unavailable();
      const timeout = new AbortController();
      const timer = setTimeout(() => timeout.abort(), SLICE_FETCH_TIMEOUT_MS);
      const abort = () => timeout.abort();
      options?.signal?.addEventListener("abort", abort, { once: true });
      try {
        const answer = await fetcher(`${EMBEDDABLE_FONT_PUBLIC_PREFIX}${slice.out}`, { signal: timeout.signal });
        if (answer.ok) response = answer;
        else if (answer.status < 500) throw unavailable();
      } catch (error) {
        if (error instanceof EmbeddableFontError || options?.signal?.aborted) throw unavailable();
      } finally {
        clearTimeout(timer);
        options?.signal?.removeEventListener("abort", abort);
      }
    }
    if (!response) throw unavailable();
    const verified = await verifyEmbeddableFontSlice(slice, await response.arrayBuffer());
    if (!verified) throw new EmbeddableFontError("slice-mismatch", `The font slice ${slice.out} is not the pinned one.`);
    return verified;
  })();
  sliceCache.set(slice.sha256, loading);
  loading.catch(() => sliceCache.delete(slice.sha256));
  return loading.then((bytes) => Uint8Array.from(bytes));
}
