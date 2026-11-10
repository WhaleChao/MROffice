/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { PresentationAudioAssetMimeType } from "@mrnote/core/web/lib/presentation-embed.ts";

/**
 * Insert > Audio > Audio on My PC and Transitions > Sound > Other Sound.
 *
 * The file is checked here (a format the deck plays and PowerPoint embeds,
 * whose first bytes are that container) and then uploaded whole as a
 * DataItem of the deck's notebook, which the deck names by `assetId`
 * (presentation-embed's stored audio). It has no size of its own: the
 * deck's media total bounds it (lib/document-size-limits), as PowerPoint
 * embeds a clip or a sound of any size.
 */

const AUDIO_FORMATS = {
  aac: { mimeType: "audio/aac", aliases: ["audio/aac", "audio/x-aac"] },
  m4a: {
    mimeType: "audio/mp4",
    aliases: ["audio/mp4", "audio/m4a", "audio/x-m4a"],
  },
  mp3: {
    mimeType: "audio/mpeg",
    aliases: ["audio/mpeg", "audio/mp3"],
  },
  oga: { mimeType: "audio/ogg", aliases: ["audio/ogg"] },
  ogg: { mimeType: "audio/ogg", aliases: ["audio/ogg"] },
  wav: {
    mimeType: "audio/wav",
    aliases: ["audio/wav", "audio/x-wav", "audio/vnd.wave"],
  },
} as const;

type PresentationAudioExtension = keyof typeof AUDIO_FORMATS;

export type PresentationAudioFileFailureCode =
  | "bad-magic"
  | "empty"
  | "read-failed"
  | "unsupported-type";

export class PresentationAudioFileError extends Error {
  readonly code: PresentationAudioFileFailureCode;

  constructor(code: PresentationAudioFileFailureCode) {
    super(`presentation_audio_file_${code}`);
    this.name = "PresentationAudioFileError";
    this.code = code;
  }
}

export interface PreparedPresentationAudioFile {
  /** The file as it is uploaded: the same bytes, typed as the deck states. */
  readonly file: File;
  readonly mimeType: PresentationAudioAssetMimeType;
  readonly sizeBytes: number;
  readonly title: string;
}

function extensionForFile(file: File): PresentationAudioExtension | null {
  const extension = file.name.toLowerCase().match(/\.([a-z0-9]+)$/)?.[1];
  return extension && extension in AUDIO_FORMATS
    ? (extension as PresentationAudioExtension)
    : null;
}

function ascii(bytes: Uint8Array, start: number, length: number): string {
  return String.fromCharCode(...bytes.slice(start, start + length));
}

function audioSignatureMatches(
  extension: PresentationAudioExtension,
  bytes: Uint8Array,
): boolean {
  if (extension === "wav") {
    return ascii(bytes, 0, 4) === "RIFF" && ascii(bytes, 8, 4) === "WAVE";
  }
  if (extension === "m4a") return ascii(bytes, 4, 4) === "ftyp";
  if (extension === "mp3") {
    return (
      ascii(bytes, 0, 3) === "ID3" ||
      (bytes[0] === 0xff && (bytes[1] ?? 0) >= 0xe0)
    );
  }
  if (extension === "oga" || extension === "ogg") {
    return ascii(bytes, 0, 4) === "OggS";
  }
  return bytes[0] === 0xff && ((bytes[1] ?? 0) & 0xf6) === 0xf0;
}

async function readHead(file: File): Promise<Uint8Array> {
  try {
    return new Uint8Array(await file.slice(0, 16).arrayBuffer());
  } catch {
    throw new PresentationAudioFileError("read-failed");
  }
}

/** Check a local audio file and type it for upload; its bytes are not read whole. */
export async function preparePresentationAudioFile(
  file: File,
): Promise<PreparedPresentationAudioFile> {
  if (file.size <= 0) throw new PresentationAudioFileError("empty");
  const extension = extensionForFile(file);
  if (!extension) throw new PresentationAudioFileError("unsupported-type");
  const format = AUDIO_FORMATS[extension];
  const declaredType = file.type.trim().toLowerCase();
  if (declaredType && !(format.aliases as readonly string[]).includes(declaredType)) {
    throw new PresentationAudioFileError("unsupported-type");
  }
  const head = await readHead(file);
  if (head.length < 4 || !audioSignatureMatches(extension, head)) {
    throw new PresentationAudioFileError("bad-magic");
  }
  // A Blob of the same File: no copy of the bytes, only the canonical type
  // the upload declares and the API stores (it refuses audio/x-m4a & co.).
  const typed =
    file.type === format.mimeType
      ? file
      : new File([file], file.name, {
          type: format.mimeType,
          lastModified: file.lastModified,
        });
  return {
    file: typed,
    mimeType: format.mimeType,
    sizeBytes: file.size,
    title: file.name.replace(/\.[^.]+$/, "").trim().slice(0, 160) || "Audio",
  };
}

/** Transition sounds use the Open XML embedded-WAV contract, not media clips. */
export async function preparePresentationTransitionSoundFile(
  file: File,
): Promise<PreparedPresentationAudioFile & { readonly mimeType: "audio/wav" }> {
  const prepared = await preparePresentationAudioFile(file);
  if (prepared.mimeType !== "audio/wav" || prepared.sizeBytes < 12) {
    throw new PresentationAudioFileError("unsupported-type");
  }
  return prepared as PreparedPresentationAudioFile & { readonly mimeType: "audio/wav" };
}
