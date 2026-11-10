/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import {
  CODE_LANGUAGE_OPTIONS,
  extensionForLanguage,
  languageForExtension,
} from "@mrnote/core/web/lib/code-runner/languages.ts";
import {
  convertEol,
  type EolStyle,
} from "./editor-commands";

/**
 * Local code-file imports stay deliberately bounded. Two MiB is large enough
 * for substantial source files while matching the existing workspace text
 * import ceiling and keeping a single browser read/decode predictable.
 */
export const MAX_CODE_IMPORT_BYTES = 2 * 1024 * 1024;

/**
 * The most a code page can HOLD, mirroring MAX_CODE_CONTENT_CHARS in
 * apps/api/app/routers/notebooks.py. The server rejects a longer `content` with
 * 400 invalid_page_content, so importing past this produces a buffer that can
 * never be saved; the importer refuses it instead, on the file picker where the
 * user can still do something about it.
 */
export const MAX_CODE_SOURCE_CHARS = 1_000_000;

export const MAX_CODE_SOURCE_FILENAME_CHARS = 255;
const WINDOWS_RESERVED_FILENAME =
  /^(?:con|prn|aux|nul|com[1-9]|lpt[1-9])(?:\.|$)/i;

/** Native file-picker filter generated from the same language registry used by
 * page routing, so the picker and the validator cannot drift. */
export const CODE_FILE_ACCEPT = Array.from(
  new Set(CODE_LANGUAGE_OPTIONS.flatMap((option) => option.extensions)),
).join(",");

export type CodeFileImportErrorCode =
  | "unsupported_type"
  | "too_large"
  | "not_text"
  | "read_failed";
export type CodeSourceEncoding =
  | "utf8"
  | "utf8bom"
  | "utf16le"
  | "utf16be";

export const CODE_SOURCE_ENCODINGS: readonly CodeSourceEncoding[] = [
  "utf8",
  "utf8bom",
  "utf16le",
  "utf16be",
];

export class CodeFileImportError extends Error {
  constructor(
    readonly code: CodeFileImportErrorCode,
    message: string,
  ) {
    super(message);
    this.name = "CodeFileImportError";
  }
}

export interface ImportedCodeFile {
  name: string;
  language: string;
  content: string;
  lineEnding: EolStyle;
  encoding: CodeSourceEncoding;
}

export interface CodeFileLike {
  name: string;
  size: number;
  text?: () => Promise<string>;
  arrayBuffer?: () => Promise<ArrayBuffer>;
}

function codeFileArrayBufferReader(
  file: CodeFileLike,
): (() => Promise<ArrayBuffer>) | null {
  const arrayBuffer = file.arrayBuffer;
  if (typeof arrayBuffer === "function") {
    return () => arrayBuffer.call(file);
  }
  // Modern File/Blob implementations expose arrayBuffer(), while older
  // WebViews (and the browser-compatible File shim used by this app's import
  // tests) expose the same bytes only through FileReader. Keep this fallback
  // here so every Code import entry uses the identical BOM/encoding boundary.
  if (
    typeof FileReader !== "undefined" &&
    typeof Blob !== "undefined" &&
    file instanceof Blob
  ) {
    return () =>
      new Promise<ArrayBuffer>((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = () => {
          if (reader.result instanceof ArrayBuffer) resolve(reader.result);
          else reject(new Error("Code file read did not return bytes."));
        };
        reader.onerror = () =>
          reject(reader.error ?? new Error("Code file could not be read."));
        reader.readAsArrayBuffer(file);
      });
  }
  return null;
}

export function normalizeCodeSourceEncoding(
  value: unknown,
): CodeSourceEncoding {
  return value === "utf8bom" ||
    value === "utf16le" ||
    value === "utf16be"
    ? value
    : "utf8";
}

export function codeSourceEncodingLabel(
  encoding: CodeSourceEncoding,
): string {
  if (encoding === "utf8bom") return "UTF-8 BOM";
  if (encoding === "utf16le") return "UTF-16 LE";
  if (encoding === "utf16be") return "UTF-16 BE";
  return "UTF-8";
}

export function nextCodeSourceEncoding(
  encoding: CodeSourceEncoding,
): CodeSourceEncoding {
  if (encoding === "utf8") return "utf8bom";
  if (encoding === "utf8bom") return "utf8";
  if (encoding === "utf16le") return "utf16be";
  return "utf8";
}

/**
 * Reduce an untrusted imported filename to one bounded local basename.
 *
 * Browsers normally expose `File.name` without a path, but drag/drop adapters
 * and persisted legacy payloads are not guaranteed to. Path separators and
 * control characters never reach a later download filename.
 */
export function normalizeCodeSourceFilename(value: unknown): string | null {
  if (typeof value !== "string") return null;
  const basename = value.trim().split(/[\\/]/).at(-1)?.trim() ?? "";
  const clean = basename
    .replace(
      /[\u0000-\u001f\u007f-\u009f\u200e\u200f\u2028\u2029\u202a-\u202e\u2066-\u2069]/g,
      "",
    )
    .replace(/[<>:"|?*]/g, "-")
    .slice(0, MAX_CODE_SOURCE_FILENAME_CHARS)
    .trim()
    .replace(/[ .]+$/g, "");
  if (!clean || clean === "." || clean === "..") return null;
  return WINDOWS_RESERVED_FILENAME.test(clean.replace(/ +(?=\.)/, ""))
    ? `_${clean.slice(0, MAX_CODE_SOURCE_FILENAME_CHARS - 1)}`
    : clean;
}

/** Return the final filename suffix, preserving the C++-specific uppercase .C. */
export function extensionFromCodeFilename(filename: string): string | null {
  const basename = normalizeCodeSourceFilename(filename) ?? "";
  const dot = basename.lastIndexOf(".");
  if (dot < 0 || dot === basename.length - 1) return null;
  const extension = basename.slice(dot);
  return extension === ".C" ? extension : extension.toLowerCase();
}

/**
 * Choose the current source download name.
 *
 * An imported basename survives edits and reopen. Registered extension aliases
 * also survive while they still describe the active language (`.tsx`, `.jsx`,
 * `.hpp`, `.yml`, …); a later language change replaces only a mismatched final
 * extension so source and filename cannot drift. Pages without imported
 * provenance keep the historical title-based fallback.
 */
export function codeExportFilename(
  sourceFilename: unknown,
  fallbackTitle: unknown,
  language: string,
): string {
  const source = normalizeCodeSourceFilename(sourceFilename);
  // The fallback is a page title, not a path: a "/" in "Timer (25/5)" is part
  // of the name, never a directory to strip.
  const fallback = normalizeCodeSourceFilename(
    typeof fallbackTitle === "string" ? fallbackTitle.replace(/[\\/]+/g, "-") : fallbackTitle,
  ) ?? "code";
  const normalizedExtension = source
    ? extensionFromCodeFilename(source)
    : null;
  const extensionStart = source?.lastIndexOf(".") ?? -1;
  const sourceExtension =
    source && extensionStart >= 0 && extensionStart < source.length - 1
      ? source.slice(extensionStart)
      : null;
  const extension =
    normalizedExtension &&
    sourceExtension &&
    languageForExtension(normalizedExtension) === language
      ? sourceExtension
      : extensionForLanguage(language);
  const rawStem = source
    ? sourceExtension
      ? source.slice(0, -sourceExtension.length)
      : source
    : fallback;
  const maxStemLength = Math.max(
    1,
    MAX_CODE_SOURCE_FILENAME_CHARS - extension.length,
  );
  const stem = rawStem.slice(0, maxStemLength) || "code";
  return `${stem}${extension}`;
}

/** Re-encode the editor's LF-normalized text with the source file's EOL. */
export function codeContentForDownload(
  content: string,
  lineEnding: EolStyle,
  encoding: "utf8" | "utf8bom" = "utf8",
): string {
  const withoutBom = content.startsWith("\uFEFF") ? content.slice(1) : content;
  const encodedContent = convertEol(withoutBom, lineEnding);
  return encoding === "utf8bom" ? `\uFEFF${encodedContent}` : encodedContent;
}

/**
 * Encode the editor's LF-normalized source into its original local-file byte
 * format. UTF-16 exports always include their endian BOM so a later import can
 * identify the encoding without heuristic content sniffing.
 */
export function codeBytesForDownload(
  content: string,
  lineEnding: EolStyle,
  encoding: CodeSourceEncoding,
): Uint8Array {
  if (encoding === "utf8" || encoding === "utf8bom") {
    return new TextEncoder().encode(
      codeContentForDownload(content, lineEnding, encoding),
    );
  }

  const withoutBom = content.startsWith("\uFEFF") ? content.slice(1) : content;
  const encodedContent = convertEol(withoutBom, lineEnding);
  const bytes = new Uint8Array(2 + encodedContent.length * 2);
  const littleEndian = encoding === "utf16le";
  bytes.set(littleEndian ? [0xff, 0xfe] : [0xfe, 0xff]);
  const view = new DataView(bytes.buffer);
  for (let index = 0; index < encodedContent.length; index += 1) {
    view.setUint16(
      2 + index * 2,
      encodedContent.charCodeAt(index),
      littleEndian,
    );
  }
  return bytes;
}

/**
 * Read one user-selected source file without executing or interpreting it.
 *
 * The browser-provided File size is checked BEFORE reading, unsupported
 * extensions never reach the decoder, UTF-8 BOM is removed, and embedded NUL
 * bytes are rejected as a binary-file signal. CodeMirror normalizes its Text
 * model to LF, so the buffer follows suit while the detected source EOL is
 * returned separately for status, persistence, reopen, and download.
 *
 * `unknownExtensionLanguage` admits a file whose extension names no language
 * (a LaTeX project's refs.bib or class file uploaded to the Code workspace)
 * as that language; without it such a file is refused as before.
 */
export async function readImportedCodeFile(
  file: CodeFileLike,
  options: { readonly unknownExtensionLanguage?: string } = {},
): Promise<ImportedCodeFile> {
  const name = normalizeCodeSourceFilename(file.name);
  const extension = name ? extensionFromCodeFilename(name) : null;
  const language =
    (extension ? languageForExtension(extension) : null) ??
    (name ? (options.unknownExtensionLanguage ?? null) : null);
  if (!language) {
    throw new CodeFileImportError(
      "unsupported_type",
      "Unsupported code file type.",
    );
  }
  if (
    !Number.isFinite(file.size) ||
    file.size < 0 ||
    file.size > MAX_CODE_IMPORT_BYTES
  ) {
    throw new CodeFileImportError(
      "too_large",
      "Code file exceeds the import limit.",
    );
  }

  let content: string;
  let encoding: CodeSourceEncoding;
  const readArrayBuffer = codeFileArrayBufferReader(file);
  if (readArrayBuffer) {
    let buffer: ArrayBuffer;
    try {
      buffer = await readArrayBuffer();
    } catch {
      throw new CodeFileImportError(
        "read_failed",
        "Code file could not be read.",
      );
    }
    if (buffer.byteLength > MAX_CODE_IMPORT_BYTES) {
      throw new CodeFileImportError(
        "too_large",
        "Code file exceeds the import limit.",
      );
    }
    const bytes = new Uint8Array(buffer);
    const hasUtf8Bom =
      bytes.length >= 3 &&
      bytes[0] === 0xef &&
      bytes[1] === 0xbb &&
      bytes[2] === 0xbf;
    const hasUtf16LeBom =
      bytes.length >= 2 && bytes[0] === 0xff && bytes[1] === 0xfe;
    const hasUtf16BeBom =
      bytes.length >= 2 && bytes[0] === 0xfe && bytes[1] === 0xff;
    encoding = hasUtf8Bom
      ? "utf8bom"
      : hasUtf16LeBom
        ? "utf16le"
        : hasUtf16BeBom
          ? "utf16be"
          : "utf8";
    const decoderEncoding =
      encoding === "utf16le"
        ? "utf-16le"
        : encoding === "utf16be"
          ? "utf-16be"
          : "utf-8";
    const bomLength = hasUtf8Bom ? 3 : hasUtf16LeBom || hasUtf16BeBom ? 2 : 0;
    try {
      content = new TextDecoder(decoderEncoding, { fatal: true }).decode(
        bytes.subarray(bomLength),
      );
    } catch {
      throw new CodeFileImportError(
        "not_text",
        "Code file is not valid UTF-8 text.",
      );
    }
  } else if (typeof file.text === "function") {
    try {
      content = await file.text();
    } catch {
      throw new CodeFileImportError(
        "read_failed",
        "Code file could not be read.",
      );
    }
    encoding = content.startsWith("\uFEFF") ? "utf8bom" : "utf8";
  } else {
    throw new CodeFileImportError(
      "read_failed",
      "Code file could not be read.",
    );
  }
  // Two ceilings, because they bound different things. The byte check above is
  // about reading a file predictably; this one is about being able to KEEP what
  // was read. A file between the two used to import successfully, replace the
  // buffer, and then have every autosave rejected 400 invalid_page_content —
  // the user's editor holding content the product structurally cannot store,
  // with the only signal arriving later as a generic save error naming the
  // wrong cause. Refusing here puts the failure on the file picker, where it is
  // actionable. (The two are not redundant: 2 MiB of CJK is well under a
  // million characters and persists fine.)
  if (content.length > MAX_CODE_SOURCE_CHARS) {
    throw new CodeFileImportError(
      "too_large",
      "Code file exceeds the import limit.",
    );
  }
  if (content.includes("\0")) {
    throw new CodeFileImportError(
      "not_text",
      "Binary content cannot be imported into the code editor.",
    );
  }

  const withoutBom = content.startsWith("\uFEFF") ? content.slice(1) : content;
  const lineEnding: EolStyle = withoutBom.includes("\r\n") ? "CRLF" : "LF";
  return {
    name: name!,
    language,
    content: withoutBom.replace(/\r\n?|\u2028|\u2029/g, "\n"),
    lineEnding,
    encoding,
  };
}
