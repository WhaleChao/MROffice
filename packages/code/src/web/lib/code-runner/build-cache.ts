/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The intermediate files a LaTeX compile returns for the next one (.aux, .bbl,
 * .toc and latexmk's own record of what it ran), so a recompile reruns only
 * the passes its edit needs instead of every pass from nothing, and the
 * generated files that spare a pass work (minted's highlighted code, TikZ
 * externalized figures), as Overleaf keeps a project's compile directory.
 *
 * Mirrors the API's BuildCache and the Worker's internal/protocol/
 * build_cache.go: the same names, order and bounds. A receipt carries them
 * for the compile's downloadable output files; the next compile is started
 * from them by the server, which keeps them per page.
 */
export const MAX_BUILD_CACHE_BYTES = 1_000_000;
export const MAX_BUILD_CACHE_FILES = 128;

const BUILD_CACHE_LANGUAGES: ReadonlySet<string> = new Set(["latex"]);
const INTERMEDIATE_NAME =
  /^(main\.(aux|bbl|bcf|blg|toc|lof|lot|loa|lol|out|fdb_latexmk|run\.xml|nav|snm|vrb|idx|ind|ilg|glo|gls|glg|acn|acr|alg|ist)|[A-Za-z0-9][A-Za-z0-9_-]{0,63}\.bib)$/;
// minted's highlighted code, named by a hash of the code and its options.
const MINTED_CACHE_NAME =
  /^_minted(-[A-Za-z0-9][A-Za-z0-9._+-]{0,63})?\/[A-Za-z0-9][A-Za-z0-9._+-]{0,127}$/;
// A TikZ externalized figure: name.pdf, name.dpth and the name.md5 of its code.
const EXTERNAL_FIGURE_NAME =
  /^([A-Za-z0-9][A-Za-z0-9._+-]{0,127}\/){0,3}[A-Za-z0-9][A-Za-z0-9._+-]{0,127}\.(md5|dpth|pdf)$/;

/** Never the source, the document or a path out of the workspace. */
function isBuildCacheName(name: string): boolean {
  if (new TextEncoder().encode(name).length > 255 || name === "main.pdf") return false;
  return (
    INTERMEDIATE_NAME.test(name) ||
    MINTED_CACHE_NAME.test(name) ||
    EXTERNAL_FIGURE_NAME.test(name)
  );
}
const BASE64 = /^(?:[A-Za-z0-9+/]{4})*(?:[A-Za-z0-9+/]{2}==|[A-Za-z0-9+/]{3}=)?$/;

export interface LanguageExecutionBuildCacheFile {
  readonly name: string;
  readonly contentBase64: string;
}

export interface LanguageExecutionBuildCache {
  readonly files: readonly LanguageExecutionBuildCacheFile[];
}

export function carriesBuildCache(language: string): boolean {
  return BUILD_CACHE_LANGUAGES.has(language);
}

function plainRecord(value: unknown, keys: readonly string[]): Record<string, unknown> | null {
  if (!value || typeof value !== "object" || Array.isArray(value)) return null;
  const prototype = Object.getPrototypeOf(value);
  if (prototype !== Object.prototype && prototype !== null) return null;
  const own = Object.keys(value);
  if (own.length !== keys.length || !keys.every((key) => own.includes(key))) {
    return null;
  }
  return value as Record<string, unknown>;
}

function decodedLength(base64: string): number | null {
  if (!BASE64.test(base64)) return null;
  try {
    // A canonical encoding re-encodes to itself (no stray padding bits).
    if (btoa(atob(base64)) !== base64) return null;
  } catch {
    return null;
  }
  const padding = base64.endsWith("==") ? 2 : base64.endsWith("=") ? 1 : 0;
  return (base64.length / 4) * 3 - padding;
}

/**
 * Whether `value` is a well-formed cache of `language`: allowed names in
 * strictly ascending order, canonical base64, within every bound.
 */
export function isLanguageExecutionBuildCache(
  value: unknown,
  language: string,
): value is LanguageExecutionBuildCache {
  if (!carriesBuildCache(language)) return false;
  const cache = plainRecord(value, ["files"]);
  if (!cache || !Array.isArray(cache.files)) return false;
  const files = cache.files as unknown[];
  if (files.length < 1 || files.length > MAX_BUILD_CACHE_FILES) return false;
  let previous = "";
  let total = 0;
  for (const entry of files) {
    const file = plainRecord(entry, ["name", "contentBase64"]);
    if (
      !file ||
      typeof file.name !== "string" ||
      typeof file.contentBase64 !== "string" ||
      !isBuildCacheName(file.name) ||
      file.name <= previous
    ) {
      return false;
    }
    previous = file.name;
    const length = decodedLength(file.contentBase64);
    if (length === null) return false;
    total += length;
    if (total > MAX_BUILD_CACHE_BYTES) return false;
  }
  return true;
}

/** The wire form (snake_case) of a validated cache. */
export function buildCacheToWire(cache: LanguageExecutionBuildCache): {
  files: { name: string; content_base64: string }[];
} {
  return {
    files: cache.files.map((file) => ({
      name: file.name,
      content_base64: file.contentBase64,
    })),
  };
}

/** A wire-form cache from a receipt, validated, or null. */
export function buildCacheFromWire(
  value: unknown,
  language: string,
): LanguageExecutionBuildCache | null {
  const wire = plainRecord(value, ["files"]);
  if (!wire || !Array.isArray(wire.files)) return null;
  const files: LanguageExecutionBuildCacheFile[] = [];
  for (const entry of wire.files as unknown[]) {
    const file = plainRecord(entry, ["name", "content_base64"]);
    if (
      !file ||
      typeof file.name !== "string" ||
      typeof file.content_base64 !== "string"
    ) {
      return null;
    }
    files.push({ name: file.name, contentBase64: file.content_base64 });
  }
  const cache = { files };
  return isLanguageExecutionBuildCache(cache, language) ? cache : null;
}

/**
 * The database a browser tab once kept its pages' caches in, sealed to that
 * tab and login. The server keeps them now (the API's copy and the
 * executor's compile directory), so a compile uploads none; logout drops
 * what an older page load left.
 */
const RETIRED_DATABASE = "mrnote-latex-build-cache-v1";

export function forgetRetiredBuildCacheStore(): void {
  try {
    globalThis.indexedDB?.deleteDatabase(RETIRED_DATABASE);
  } catch {
    // Best effort: a sealed record nobody can read back.
  }
}
