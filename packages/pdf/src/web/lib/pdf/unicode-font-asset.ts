/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
export const PDF_UNICODE_FALLBACK_FONT_URL =
  "/fonts/NotoSansSC-Regular.ttf" as const;
export const PDF_UNICODE_FALLBACK_FONT_BYTE_LENGTH = 10_595_916;
export const PDF_UNICODE_FALLBACK_FONT_SHA256 =
  "4b2be8791bb2f58f7d9c1a2977ed82b5bda0d475dc480f1a6ad7782c07858280" as const;

type FontFetch = (
  input: RequestInfo | URL,
  init?: RequestInit,
) => Promise<Response>;

function digestHex(bytes: Uint8Array): Promise<string> {
  if (!globalThis.crypto?.subtle) {
    return Promise.reject(new Error("pdf_unicode_font_digest_unavailable"));
  }
  const owned = new Uint8Array(bytes.byteLength);
  owned.set(bytes);
  return globalThis.crypto.subtle
    .digest("SHA-256", owned.buffer)
    .then((digest) =>
      [...new Uint8Array(digest)]
        .map((value) => value.toString(16).padStart(2, "0"))
        .join(""),
    );
}

export async function verifyPdfUnicodeFallbackFontBytes(
  value: Uint8Array | ArrayBuffer,
): Promise<Uint8Array<ArrayBuffer> | null> {
  const source = value instanceof Uint8Array ? value : new Uint8Array(value);
  const bytes = new Uint8Array(source.byteLength);
  bytes.set(source);
  if (bytes.byteLength !== PDF_UNICODE_FALLBACK_FONT_BYTE_LENGTH) return null;
  if ((await digestHex(bytes)) !== PDF_UNICODE_FALLBACK_FONT_SHA256) {
    return null;
  }
  return bytes;
}

async function loadVerifiedPdfUnicodeFallbackFont(fetcher: FontFetch, signal?: AbortSignal): Promise<ArrayBuffer | null> {
  try {
    const response = await fetcher(PDF_UNICODE_FALLBACK_FONT_URL, signal ? { signal } : undefined);
    if (!response.ok) return null;
    const bytes = await verifyPdfUnicodeFallbackFontBytes(
      await response.arrayBuffer(),
    );
    return bytes?.buffer ?? null;
  } catch {
    return null;
  }
}





let verifiedFallbackFont: Promise<ArrayBuffer | null> | null = null;

/** Load the canonical unicode fallback only when deployed bytes match Git.
 * Each call answers its own copy (the callers wipe theirs once used). */
export async function fetchVerifiedPdfUnicodeFallbackFont(options?: {
  readonly fetch?: FontFetch;
  readonly signal?: AbortSignal;
}): Promise<ArrayBuffer | null> {
  if (options?.fetch) return loadVerifiedPdfUnicodeFallbackFont(options.fetch, options.signal);
  const fetcher = globalThis.fetch;
  if (typeof fetcher !== "function") return null;
  if (!verifiedFallbackFont) {
    const loading = loadVerifiedPdfUnicodeFallbackFont(fetcher.bind(globalThis));
    verifiedFallbackFont = loading;
    void loading.then((buffer) => { if (!buffer && verifiedFallbackFont === loading) verifiedFallbackFont = null; });
  }
  const signal = options?.signal;
  if (signal?.aborted) return null;
  const buffer = await (signal ? Promise.race([verifiedFallbackFont, new Promise<null>((resolve) =>
    signal.addEventListener("abort", () => resolve(null), { once: true }))]) : verifiedFallbackFont);
  return buffer ? buffer.slice(0) : null;
}

/**
 * The packaged faces of the complex scripts (Arabic, Hebrew, Thai,
 * Devanagari, Bengali): the same files, by size and digest, the API shapes
 * them with (`app.services.agent_runs.document_packaged_fonts`,
 * `app.services.pdf_text_shaping`) and admits as bundled programs of a font
 * package (`native_font_extension_contract.bundled_program_serves`). A
 * shaped piece names its face by the API's source id.
 */
export const PDF_COMPLEX_SCRIPT_FONTS: Readonly<Record<string, Readonly<{ url: string; byteLength: number; sha256: string }>>> =
  Object.freeze({
    "mrnote.noto-sans-arabic.v1": Object.freeze({ url: "/fonts/NotoSansArabic-Regular.ttf", byteLength: 234_892,
      sha256: "bdff3e5659d67e67def05b33f749683b9376ae819d65d3dd62ac4640b3aaef48" }),
    "mrnote.noto-sans-hebrew.v1": Object.freeze({ url: "/fonts/NotoSansHebrew-Regular.ttf", byteLength: 26_860,
      sha256: "cdefaf8efd47045f6820928eba84db5bed7557539328952b5f828315485e02ee" }),
    "mrnote.noto-sans-thai.v1": Object.freeze({ url: "/fonts/NotoSansThai-Regular.ttf", byteLength: 37_780,
      sha256: "61cf814eec46b294d6ea4401ac295d0cecd5207bd2331dcc5a15e7301d30ee44" }),
    "mrnote.noto-sans-devanagari.v1": Object.freeze({ url: "/fonts/NotoSansDevanagari-Regular.ttf", byteLength: 243_520,
      sha256: "4e3c66638958c3e2ab5d37f47a8deb89fffeb7be9985c665a519bbc7ba762313" }),
    "mrnote.noto-sans-bengali.v1": Object.freeze({ url: "/fonts/NotoSansBengali-Regular.ttf", byteLength: 143_072,
      sha256: "b55c62ee531e3214da6c0701daecea89a52ba42db7d8206b92e6b51f397a3193" }),
  });

/** Whether `sha256` is one of the packaged complex-script faces. */
export function isPdfComplexScriptFontSha256(sha256: string): boolean {
  return Object.values(PDF_COMPLEX_SCRIPT_FONTS).some((face) => face.sha256 === sha256);
}

const verifiedComplexScriptFonts = new Map<string, Promise<ArrayBuffer | null>>();

/** One packaged complex-script face, only when the deployed bytes are the
 * pinned ones; each call answers its own copy. */
export async function fetchVerifiedPdfComplexScriptFont(sourceId: string, options?: {
  readonly fetch?: FontFetch;
  readonly signal?: AbortSignal;
}): Promise<ArrayBuffer | null> {
  const face = PDF_COMPLEX_SCRIPT_FONTS[sourceId];
  const fetcher = options?.fetch ?? (typeof globalThis.fetch === "function" ? globalThis.fetch.bind(globalThis) : null);
  if (!face || !fetcher) return null;
  const load = async (): Promise<ArrayBuffer | null> => {
    try {
      const response = await fetcher(face.url);
      if (!response.ok) return null;
      const bytes = new Uint8Array(await response.arrayBuffer());
      if (bytes.byteLength !== face.byteLength || (await digestHex(bytes)) !== face.sha256) return null;
      return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;
    } catch {
      return null;
    }
  };
  let pending = options?.fetch ? null : verifiedComplexScriptFonts.get(sourceId) ?? null;
  if (!pending) {
    pending = load();
    if (!options?.fetch) {
      verifiedComplexScriptFonts.set(sourceId, pending);
      void pending.then((buffer) => { if (!buffer) verifiedComplexScriptFonts.delete(sourceId); });
    }
  }
  if (options?.signal?.aborted) return null;
  const buffer = await pending;
  return buffer ? buffer.slice(0) : null;
}
