/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Pure logic for "Format Document" (VSCode parity), powered by Prettier's
 * BROWSER standalone build.
 *
 * `formatCode(text, language)` pretty-prints a source buffer using the Prettier
 * parser that matches the editor language — or the dedicated `sql-formatter`
 * engine for SQL (Prettier has no SQL parser). It is intentionally TOTAL:
 *   - a language with no available built-in or provider formatter returns the text
 *     UNCHANGED — the action no-ops rather than corrupting the buffer;
 *   - invalid syntax (the formatter throws while parsing) returns the ORIGINAL text;
 *   - any unexpected runtime error returns the ORIGINAL text.
 * It NEVER throws, so callers can wire it straight to a buffer-replace without a
 * try/catch and trust that the worst case is "nothing changed".
 *
 * Formatting is a TEXT transform: it returns a new string and the editor applies
 * it through its normal change path, so the browser's native undo stack records
 * the format as one reversible step (Ctrl/Cmd+Z reverts it).
 *
 * Prettier and its plugins are imported lazily (dynamic `import()`) so the
 * editor bundle only pulls the formatter in when the user actually formats, and
 * only the plugin needed for the current language is loaded.
 */

/**
 * Prettier parser ids keyed by code language id (see CODE_LANGUAGE_OPTIONS).
 * Only languages with a Prettier parser available in the installed plugin set
 * (babel, estree, typescript, postcss, html, markdown, yaml) appear here; every
 * other language id has no parser and `formatCode` no-ops for it.
 *
 * - JavaScript family → "babel"      (babel + estree plugins)
 * - TypeScript family → "typescript" (typescript + estree plugins)
 * - JSON              → "json"       (babel + estree plugins)
 * - CSS               → "css"        (postcss plugin)
 * - HTML / XML        → "html"       (html plugin)
 * - Markdown          → "markdown"   (markdown plugin)
 * - YAML              → "yaml"       (yaml plugin)
 */
export const PARSER_BY_LANGUAGE: Readonly<Record<string, string>> = {
  javascript: "babel",
  typescript: "typescript",
  json: "json",
  css: "css",
  // CODE_LANGUAGE_OPTIONS labels this "HTML / XML" under the highlight.js grammar
  // id "xml"; Prettier's html parser handles both.
  xml: "html",
  html: "html",
  markdown: "markdown",
  yaml: "yaml",
};

/** The editor language id formatted by the dedicated SQL engine (not Prettier). */
export const SQL_LANGUAGE = "sql";

/** The editor language id formatted via the Pyodide worker (autopep8). */
export const PYTHON_LANGUAGE = "python";

/**
 * Go is formatted by its registered, version-pinned WASM provider. Rust still
 * uses the bounded structural fallback until a real rustfmt provider exists.
 * Keeping the ids here preserves the public language constants without
 * pretending Go has a built-in text reindenter.
 */
export const GO_LANGUAGE = "go";
export const RUST_LANGUAGE = "rust";

/**
 * The Prettier parser id for a code language, or `null` when the language has no
 * available parser (so the Format action should be disabled / no-op).
 */
export function parserForLanguage(language: string): string | null {
  return PARSER_BY_LANGUAGE[language] ?? null;
}

/**
 * Whether the Format action can run for this language — true when a Prettier
 * parser exists, OR the language is SQL (handled by `sql-formatter`), OR the
 * language is Python (handled via autopep8 in the Pyodide worker), OR the
 * language is Rust (handled by the best-effort structural formatter). Go is
 * intentionally absent: its real gofmt arrives through `providerAvailable`.
 */
export function canFormatLanguage(language: string): boolean {
  return (
    language === SQL_LANGUAGE ||
    language === PYTHON_LANGUAGE ||
    language === RUST_LANGUAGE ||
    parserForLanguage(language) !== null
  );
}

/**
 * Decision logic for "format-on-save" (VSCode `editor.formatOnSave` parity):
 * should an explicit save gesture (Ctrl/Cmd+S) pretty-print the buffer FIRST?
 *
 * Returns true only when BOTH hold:
 *   - the user has the "Format on save" toggle ON, and
 *   - the current language actually has a formatter ({@link canFormatLanguage} —
 *     the Prettier set plus SQL / Python / Rust, or a live provider formatter).
 *     A language with no formatter (Java, Bash, plaintext, …) saves its raw
 *     text exactly like the disabled Format button no-ops.
 *
 * A false result means "persist the buffer unchanged"; the save itself still
 * happens (formatting is never allowed to block a save). Pure + total so it can
 * be unit-tested in isolation and reused by every editor surface.
 */
export function shouldFormatOnSave(options: {
  language: string;
  formatOnSave: boolean;
  providerAvailable?: boolean;
}): boolean {
  return (
    options.formatOnSave &&
    (canFormatLanguage(options.language) || options.providerAvailable === true)
  );
}

/**
 * Truthful outcome from one formatter attempt. `ok: true` includes a valid,
 * byte-identical result (the source was already formatted); callers compare the
 * returned text with their invocation snapshot to distinguish that from a real
 * edit. `ok: false` always carries the original source so a failed formatter can
 * never corrupt the buffer.
 */
export interface FormatCodeReceipt {
  ok: boolean;
  formatted: string;
}

export interface FormatCodeOptions {
  /** Cancels the pending worker receipt; non-worker engines are retired by caller authority. */
  signal?: AbortSignal;
}

function formatSuccess(formatted: string): FormatCodeReceipt {
  return { ok: true, formatted };
}

function formatFailure(original: string): FormatCodeReceipt {
  return { ok: false, formatted: original };
}

function validateFormattedText(
  original: string,
  formatted: unknown,
): FormatCodeReceipt {
  if (
    typeof formatted !== "string" ||
    (original.trim().length > 0 && formatted.trim().length === 0)
  ) {
    return formatFailure(original);
  }
  return formatSuccess(formatted);
}

/**
 * Format SQL via the lazily-loaded `sql-formatter` (Prettier has no SQL parser).
 * Uses the generic standard-SQL dialect — the editor exposes a single `sql`
 * language id, not per-dialect ones. TOTAL: returns the original text on a parse
 * error or any failure, matching `formatCode`'s never-throw contract.
 */
async function formatSqlWithReceipt(
  text: string,
  signal?: AbortSignal,
): Promise<FormatCodeReceipt> {
  try {
    const mod = (await import("sql-formatter")) as {
      format?: (sql: string, opts?: Record<string, unknown>) => string;
      default?: { format?: (sql: string, opts?: Record<string, unknown>) => string };
    };
    const format = mod.format ?? mod.default?.format;
    if (typeof format !== "function" || signal?.aborted) {
      return formatFailure(text);
    }
    const out = format(text, { language: "sql" });
    return signal?.aborted ? formatFailure(text) : validateFormattedText(text, out);
  } catch {
    return formatFailure(text);
  }
}

/**
 * Lazily load the Prettier standalone formatter plus exactly the plugins needed
 * for `parser`. Returns `null` if the dynamic import fails for any reason (so
 * the caller falls back to the original text). Kept separate from `formatCode`
 * so it can be stubbed in tests if needed; in production it pulls
 * `prettier/standalone` + `prettier/plugins/*`.
 */
async function loadPrettier(
  parser: string,
): Promise<{
  format: (source: string, options: Record<string, unknown>) => Promise<string>;
  plugins: unknown[];
} | null> {
  try {
    const standalone = await import("prettier/standalone");
    const plugins: unknown[] = [];
    // estree is the shared printer for the JS/TS/JSON ASTs.
    if (parser === "babel" || parser === "json") {
      const [babel, estree] = await Promise.all([
        import("prettier/plugins/babel"),
        import("prettier/plugins/estree"),
      ]);
      plugins.push(babel.default ?? babel, estree.default ?? estree);
    } else if (parser === "typescript") {
      const [ts, estree] = await Promise.all([
        import("prettier/plugins/typescript"),
        import("prettier/plugins/estree"),
      ]);
      plugins.push(ts.default ?? ts, estree.default ?? estree);
    } else if (parser === "css") {
      const postcss = await import("prettier/plugins/postcss");
      plugins.push(postcss.default ?? postcss);
    } else if (parser === "html") {
      const html = await import("prettier/plugins/html");
      plugins.push(html.default ?? html);
    } else if (parser === "markdown") {
      const md = await import("prettier/plugins/markdown");
      plugins.push(md.default ?? md);
    } else if (parser === "yaml") {
      const yaml = await import("prettier/plugins/yaml");
      plugins.push(yaml.default ?? yaml);
    } else {
      return null;
    }
    return { format: standalone.format, plugins };
  } catch {
    return null;
  }
}

/**
 * Pretty-print `text` for the given code `language`.
 *
 * Returns the formatted source on success, or the ORIGINAL `text` unchanged when
 * the language has no formatter, the source has a syntax error the formatter
 * rejects, or anything else goes wrong. Never throws.
 *
 * Language routing:
 *  - SQL     → `sql-formatter` (Prettier has no SQL parser)
 *  - Python  → autopep8 via the Pyodide worker (lazy-loaded via micropip)
 *  - Go      → registered provider backed by the pinned Go WASM toolchain
 *  - Others  → Prettier (lazy dynamic import)
 *
 * @param text     Source buffer to format.
 * @param language Code language id (see CODE_LANGUAGE_OPTIONS).
 * @returns A promise resolving to the formatted (or original) source.
 */
export type ProviderFormatter = (
  text: string,
  signal?: AbortSignal,
) => Promise<string | null>;

export async function formatCodeWithReceipt(
  text: string,
  language: string,
  providerFormat?: ProviderFormatter,
  options: FormatCodeOptions = {},
): Promise<FormatCodeReceipt> {
  if (options.signal?.aborted) return formatFailure(text);
  // SQL has no Prettier parser — route it to the dedicated sql-formatter engine.
  if (language === SQL_LANGUAGE) {
    return formatSqlWithReceipt(text, options.signal);
  }
  // Python is formatted via autopep8 in the Pyodide worker (lazy micropip load).
  if (language === PYTHON_LANGUAGE) {
    return formatPythonWithReceipt(text, options);
  }
  // Rust still uses a best-effort structural formatter. Go deliberately falls
  // through to its registered provider so the editor never labels reindentation
  // as gofmt.
  if (language === RUST_LANGUAGE) {
    return formatRustWithReceipt(text, options.signal);
  }
  const parser = parserForLanguage(language);
  // No built-in formatter for this language: a registered provider may still
  // have one. Tried LAST, so registering a provider never changes how a
  // language that already formatted here is formatted.
  if (!parser) {
    if (!providerFormat) return formatFailure(text);
    try {
      const formatted = await providerFormat(text, options.signal);
      if (options.signal?.aborted) return formatFailure(text);
      // A provider that declines, or hands back something that is not text,
      // leaves the buffer alone — `formatCode` never corrupts a document.
      return validateFormattedText(text, formatted);
    } catch {
      return formatFailure(text);
    }
  }

  const loaded = await loadPrettier(parser);
  if (!loaded || options.signal?.aborted) return formatFailure(text);

  try {
    const formatted = await loaded.format(text, {
      parser,
      plugins: loaded.plugins,
      // XML is printed by Prettier's HTML printer, which by default keeps the
      // whitespace around every element it does not know as block-level and
      // so wrapped tags as `<Root\n  ><Item …`. XML has no inline flow: lay
      // its elements out one per line, as an XML formatter does.
      ...(language === "xml" ? { htmlWhitespaceSensitivity: "ignore" } : {}),
    });
    // Defensive: never hand back a non-string (some edge plugins could).
    if (options.signal?.aborted) return formatFailure(text);
    return validateFormattedText(text, formatted);
  } catch {
    // Invalid syntax (Prettier parse error) or any other failure: leave the
    // buffer exactly as it was.
    return formatFailure(text);
  }
}

/**
 * Backward-compatible string transform used by the rollback textarea editor.
 * New visible surfaces should consume {@link formatCodeWithReceipt} so a real
 * formatter failure is not misreported as an idempotent success.
 */
export async function formatCode(
  text: string,
  language: string,
  providerFormat?: ProviderFormatter,
  options: FormatCodeOptions = {},
): Promise<string> {
  const receipt = await formatCodeWithReceipt(
    text,
    language,
    providerFormat,
    options,
  );
  return receipt.formatted;
}

// ---------------------------------------------------------------------------
// Python formatter (Pyodide worker — autopep8 via micropip).
// ---------------------------------------------------------------------------

/**
 * Format Python source via autopep8 in the Pyodide worker. TOTAL: returns the
 * original `text` on any failure (no worker, micropip unavailable, autopep8
 * error) so the buffer is never corrupted. Matches the never-throw contract of
 * `formatCode`.
 *
 * Kept internal to this module so the worker dependency stays lazy: the
 * `formatPython` function in python-runner.ts is only imported when the user
 * actually formats a Python buffer.
 */
async function formatPythonWithReceipt(
  text: string,
  options: FormatCodeOptions,
): Promise<FormatCodeReceipt> {
  try {
    const { formatPythonWithReceipt: workerFormat } = await import(
      "@mrnote/core/web/lib/code-runner/python-runner.ts"
    );
    const result = await workerFormat(text, { signal: options.signal });
    return result.ok
      ? validateFormattedText(text, result.formatted)
      : formatFailure(text);
  } catch {
    return formatFailure(text);
  }
}

// ---------------------------------------------------------------------------
// Rust formatter (best-effort structural reindenter).
// ---------------------------------------------------------------------------

/**
 * Format Rust via the best-effort structural reindenter in
 * go-rust-formatter.ts. The module is imported lazily so it only loads when
 * the user actually formats a Rust buffer. TOTAL: returns the original
 * `text` on any failure, matching the never-throw contract of `formatCode`.
 *
 * Note: this is NOT rustfmt. Go no longer comes through this fallback: its
 * provider runs the real standard-library formatter in the pinned WASM worker.
 */
async function formatRustWithReceipt(
  text: string,
  signal?: AbortSignal,
): Promise<FormatCodeReceipt> {
  try {
    const mod = await import("./go-rust-formatter");
    if (signal?.aborted) return formatFailure(text);
    const result = mod.formatRust(text);
    return signal?.aborted
      ? formatFailure(text)
      : validateFormattedText(text, result);
  } catch {
    return formatFailure(text);
  }
}
