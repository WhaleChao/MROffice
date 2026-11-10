/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { randomUUID } from "./random-uuid";

/** Strict local WebVTT tracks attached to one PowerPoint video or audio block. */

export const MAX_PRESENTATION_MEDIA_CAPTION_TRACKS = 8;
export const MAX_PRESENTATION_MEDIA_CAPTION_BYTES = 1_000_000;
export const MAX_PRESENTATION_MEDIA_CAPTION_LABEL_CHARS = 200;
export const PRESENTATION_NATIVE_CAPTION_AUTHORITY_SCHEMA_VERSION =
  "pptx-native-media-captions-v1" as const;
export const PRESENTATION_ABSENT_CAPTION_FINGERPRINT =
  "cc852463ec93732271a7ecbd0b7b5879de84d57a95eec87fd61399e4d8c4c579" as const;

const CAPTION_TRACK_ID =
  /^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i;
const CAPTION_TRACK_LANGUAGE = /^[A-Za-z]{2,8}(?:-[A-Za-z0-9]{1,8})*$/;
// WebVTT timestamps: hours are two OR MORE digits (000:00:01.000 is a
// legal zero hour, PPT-B397); minutes and seconds are exactly two.
const WEBVTT_CUE_TIMING =
  /(?:^|\n)(?:(?:[^\n]+)\n)?(?:\d{2,}:)?\d{2}:\d{2}\.\d{3}[ \t]+-->[ \t]+(?:\d{2,}:)?\d{2}:\d{2}\.\d{3}(?:[ \t]|\n|$)/;
const CAPTION_FINGERPRINT = /^[0-9a-f]{64}$/;
const CAPTION_SLIDE_ID = /^sld[1-9][0-9]*$/;
const CAPTION_SLIDE_PART = /^ppt\/slides\/slide[1-9][0-9]*\.xml$/;
const CAPTION_SHAPE_ID = /^shp[1-9][0-9]*$/;
const CAPTION_RELATIONSHIP_ID = /^rId[1-9][0-9]*$/;
const CAPTION_TRACK_PART =
  /^ppt\/media\/[A-Za-z0-9][A-Za-z0-9._-]{0,199}\.vtt$/;

export interface PresentationMediaCaptionTrack {
  readonly id: string;
  readonly label: string;
  readonly language?: string;
  readonly vtt: string;
}

export interface NativePresentationCaptionTrackAuthority {
  readonly id: string;
  readonly label: string;
  readonly language?: string;
  readonly relationshipId: string;
  readonly part: string;
  readonly partFingerprint: string;
  readonly sizeBytes: number;
}

export interface NativePresentationCaptionAuthority {
  readonly schemaVersion: typeof PRESENTATION_NATIVE_CAPTION_AUTHORITY_SCHEMA_VERSION;
  readonly disposition: "absent" | "native";
  readonly slideId: string;
  readonly slidePart: string;
  readonly relationshipsPart: string;
  readonly shapeId: string;
  readonly captionFingerprint: string;
  readonly tracks: readonly NativePresentationCaptionTrackAuthority[];
}

export type PresentationMediaCaptionFileErrorCode =
  | "unsupported-file"
  | "file-size"
  | "invalid-utf8"
  | "invalid-webvtt";

export class PresentationMediaCaptionFileError extends Error {
  readonly code: PresentationMediaCaptionFileErrorCode;

  constructor(code: PresentationMediaCaptionFileErrorCode) {
    super(`presentation_media_caption_${code}`);
    this.name = "PresentationMediaCaptionFileError";
    this.code = code;
  }
}

function captionRecord(value: unknown): Record<string, unknown> | null {
  return value && typeof value === "object" && !Array.isArray(value)
    ? (value as Record<string, unknown>)
    : null;
}

function exactKeys(
  record: Record<string, unknown>,
  required: readonly string[],
  optional: readonly string[] = [],
): boolean {
  const keys = Object.keys(record);
  return (
    required.every((key) => Object.hasOwn(record, key)) &&
    keys.every((key) => required.includes(key) || optional.includes(key))
  );
}

function captionRelationshipsPart(slidePart: string): string {
  const slash = slidePart.lastIndexOf("/");
  const directory = slidePart.slice(0, slash);
  const filename = slidePart.slice(slash + 1);
  return `${directory}/_rels/${filename}.rels`;
}

/** Canonical UTF-8 WebVTT text or null. No SRT/HTML/remote URL coercion. */
export function normalizePresentationWebVtt(value: unknown): string | null {
  if (typeof value !== "string") return null;
  const normalized = value.replace(/^\uFEFF/, "").replace(/\r\n?/g, "\n");
  if (
    !/^WEBVTT(?:[ \t][^\n]*)?(?:\n|$)/.test(normalized) ||
    /[\u0000\u000B\u000C\u000E-\u001F\u007F]/.test(normalized) ||
    new TextEncoder().encode(normalized).length >
      MAX_PRESENTATION_MEDIA_CAPTION_BYTES ||
    !WEBVTT_CUE_TIMING.test(normalized)
  ) {
    return null;
  }
  return normalized;
}

export function normalizePresentationMediaCaptionTrack(
  value: unknown,
): PresentationMediaCaptionTrack | null {
  const record = captionRecord(value);
  if (
    !record ||
    Object.keys(record).some(
      (key) => !["id", "label", "language", "vtt"].includes(key),
    )
  ) {
    return null;
  }
  const id = typeof record.id === "string" ? record.id : "";
  const label = typeof record.label === "string" ? record.label.trim() : "";
  const language =
    typeof record.language === "string" ? record.language.trim() : undefined;
  const vtt = normalizePresentationWebVtt(record.vtt);
  if (
    !CAPTION_TRACK_ID.test(id) ||
    !label ||
    label.length > MAX_PRESENTATION_MEDIA_CAPTION_LABEL_CHARS ||
    (language !== undefined && !CAPTION_TRACK_LANGUAGE.test(language)) ||
    vtt === null
  ) {
    return null;
  }
  return Object.freeze({
    id: id.toLowerCase(),
    label,
    ...(language ? { language } : {}),
    vtt,
  });
}

/** Strict, bounded and identity-unique track list. Invalid input fails closed. */
export function normalizePresentationMediaCaptionTracks(
  value: unknown,
): readonly PresentationMediaCaptionTrack[] | null {
  if (
    !Array.isArray(value) ||
    value.length > MAX_PRESENTATION_MEDIA_CAPTION_TRACKS
  ) {
    return null;
  }
  const tracks: PresentationMediaCaptionTrack[] = [];
  const ids = new Set<string>();
  for (const raw of value) {
    const track = normalizePresentationMediaCaptionTrack(raw);
    if (!track || ids.has(track.id)) return null;
    ids.add(track.id);
    tracks.push(track);
  }
  return Object.freeze(tracks);
}

/** Strict caption-only native authority. Unknown or stale-looking shapes fail closed. */
export function parseNativePresentationCaptionAuthority(
  value: unknown,
): NativePresentationCaptionAuthority | null {
  const record = captionRecord(value);
  if (
    !record ||
    !exactKeys(record, [
      "schemaVersion",
      "disposition",
      "slideId",
      "slidePart",
      "relationshipsPart",
      "shapeId",
      "captionFingerprint",
      "tracks",
    ]) ||
    record.schemaVersion !==
      PRESENTATION_NATIVE_CAPTION_AUTHORITY_SCHEMA_VERSION ||
    (record.disposition !== "absent" && record.disposition !== "native") ||
    typeof record.slideId !== "string" ||
    !CAPTION_SLIDE_ID.test(record.slideId) ||
    typeof record.slidePart !== "string" ||
    !CAPTION_SLIDE_PART.test(record.slidePart) ||
    typeof record.relationshipsPart !== "string" ||
    record.relationshipsPart !== captionRelationshipsPart(record.slidePart) ||
    typeof record.shapeId !== "string" ||
    !CAPTION_SHAPE_ID.test(record.shapeId) ||
    typeof record.captionFingerprint !== "string" ||
    !CAPTION_FINGERPRINT.test(record.captionFingerprint) ||
    !Array.isArray(record.tracks)
  ) {
    return null;
  }
  if (record.disposition === "absent") {
    if (
      record.tracks.length !== 0 ||
      record.captionFingerprint !== PRESENTATION_ABSENT_CAPTION_FINGERPRINT
    ) {
      return null;
    }
    return Object.freeze({
      schemaVersion: PRESENTATION_NATIVE_CAPTION_AUTHORITY_SCHEMA_VERSION,
      disposition: "absent",
      slideId: record.slideId,
      slidePart: record.slidePart,
      relationshipsPart: record.relationshipsPart,
      shapeId: record.shapeId,
      captionFingerprint: record.captionFingerprint,
      tracks: Object.freeze([]),
    });
  }
  if (
    record.tracks.length < 1 ||
    record.tracks.length > MAX_PRESENTATION_MEDIA_CAPTION_TRACKS
  ) {
    return null;
  }
  const ids = new Set<string>();
  const relationshipIds = new Set<string>();
  const parts = new Set<string>();
  const tracks: NativePresentationCaptionTrackAuthority[] = [];
  for (const value of record.tracks) {
    const track = captionRecord(value);
    if (
      !track ||
      !exactKeys(
        track,
        [
          "id",
          "label",
          "relationshipId",
          "part",
          "partFingerprint",
          "sizeBytes",
        ],
        ["language"],
      ) ||
      typeof track.id !== "string" ||
      !CAPTION_TRACK_ID.test(track.id) ||
      typeof track.label !== "string" ||
      !track.label.trim() ||
      track.label.trim().length > MAX_PRESENTATION_MEDIA_CAPTION_LABEL_CHARS ||
      (track.language !== undefined &&
        (typeof track.language !== "string" ||
          !CAPTION_TRACK_LANGUAGE.test(track.language.trim()))) ||
      typeof track.relationshipId !== "string" ||
      !CAPTION_RELATIONSHIP_ID.test(track.relationshipId) ||
      typeof track.part !== "string" ||
      !CAPTION_TRACK_PART.test(track.part) ||
      typeof track.partFingerprint !== "string" ||
      !CAPTION_FINGERPRINT.test(track.partFingerprint) ||
      !Number.isInteger(track.sizeBytes) ||
      (track.sizeBytes as number) < 1 ||
      (track.sizeBytes as number) > MAX_PRESENTATION_MEDIA_CAPTION_BYTES
    ) {
      return null;
    }
    const id = track.id.toLowerCase();
    if (
      ids.has(id) ||
      relationshipIds.has(track.relationshipId) ||
      parts.has(track.part)
    ) {
      return null;
    }
    ids.add(id);
    relationshipIds.add(track.relationshipId);
    parts.add(track.part);
    tracks.push(
      Object.freeze({
        id,
        label: track.label.trim(),
        ...(typeof track.language === "string"
          ? { language: track.language.trim() }
          : {}),
        relationshipId: track.relationshipId,
        part: track.part,
        partFingerprint: track.partFingerprint,
        sizeBytes: track.sizeBytes as number,
      }),
    );
  }
  return Object.freeze({
    schemaVersion: PRESENTATION_NATIVE_CAPTION_AUTHORITY_SCHEMA_VERSION,
    disposition: "native",
    slideId: record.slideId,
    slidePart: record.slidePart,
    relationshipsPart: record.relationshipsPart,
    shapeId: record.shapeId,
    captionFingerprint: record.captionFingerprint,
    tracks: Object.freeze(tracks),
  });
}

export function nativePresentationCaptionAuthority(
  block: unknown,
): NativePresentationCaptionAuthority | null {
  const record = captionRecord(block);
  const source = captionRecord(record?.source);
  if (
    record?.type !== "embed" ||
    (record.kind !== "video" && record.kind !== "audio") ||
    !source
  ) {
    return null;
  }
  const authority = parseNativePresentationCaptionAuthority(
    source.nativeCaptionAuthority,
  );
  return authority &&
    source.slideId === authority.slideId &&
    source.shapeId === authority.shapeId
    ? authority
    : null;
}

/** Local, bounded WebVTT file preparation. No SRT conversion or remote URL. */
export async function preparePresentationMediaCaptionFile(
  file: File,
): Promise<PresentationMediaCaptionTrack> {
  if (
    !file.name.toLowerCase().endsWith(".vtt") ||
    (file.type !== "" && file.type.toLowerCase() !== "text/vtt")
  ) {
    throw new PresentationMediaCaptionFileError("unsupported-file");
  }
  if (
    !Number.isInteger(file.size) ||
    file.size < 1 ||
    file.size > MAX_PRESENTATION_MEDIA_CAPTION_BYTES
  ) {
    throw new PresentationMediaCaptionFileError("file-size");
  }
  let text: string;
  try {
    text = new TextDecoder("utf-8", { fatal: true }).decode(
      await file.arrayBuffer(),
    );
  } catch {
    throw new PresentationMediaCaptionFileError("invalid-utf8");
  }
  const vtt = normalizePresentationWebVtt(text);
  if (vtt === null) {
    throw new PresentationMediaCaptionFileError("invalid-webvtt");
  }
  const withoutExtension = file.name.slice(0, -4).trim();
  const segments = withoutExtension.split(".");
  const possibleLanguage = segments.at(-1)?.trim();
  // A language tag is the dotted suffix AFTER a name ("English.en.vtt"); a
  // one-segment name ("English.vtt", "French.vtt") is the name itself, never
  // a language to strip, else every such file is labelled "Captions" (PPT-B376).
  const language =
    segments.length >= 2 &&
    possibleLanguage &&
    CAPTION_TRACK_LANGUAGE.test(possibleLanguage)
      ? possibleLanguage
      : undefined;
  const labelSource = language
    ? segments.slice(0, -1).join(".")
    : withoutExtension;
  const label = (labelSource.trim() || "Captions").slice(
    0,
    MAX_PRESENTATION_MEDIA_CAPTION_LABEL_CHARS,
  );
  return Object.freeze({
    id: randomUUID().toLowerCase(),
    label,
    ...(language ? { language } : {}),
    vtt,
  });
}

export function presentationMediaCaptionDataUrl(
  track: PresentationMediaCaptionTrack,
): string {
  return `data:text/vtt;charset=utf-8,${encodeURIComponent(track.vtt)}`;
}
