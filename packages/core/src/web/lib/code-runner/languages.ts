/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Code editor languages.
 *
 * Each option's `id` is a highlight.js / lowlight grammar name (so syntax
 * highlighting works) and `label` is the display name. `extensions` drives
 * routing a dropped/imported code file into a native code page; entries with
 * no extensions are selectable but never auto-detected. Plain prose formats
 * (.txt, .csv, .md, .log) are intentionally absent so they keep importing as
 * document / spreadsheet pages.
 */

export interface CodeLanguageOption {
  /** lowlight/highlight.js grammar name. */
  id: string;
  /** Human-facing label for the language selector. */
  label: string;
  /** File extensions (with dot) that import into a code page as this language. */
  extensions: string[];
  /**
   * Whether this selector entry represents executable source. Data, markup and
   * prose formats deliberately remain non-executable: supporting JSON or CSS
   * means editing/validation/preview, never pretending that a Run command has a
   * program to invoke.
   */
  executionKind: "executable" | "non-executable";
}

export const CODE_LANGUAGE_OPTIONS: CodeLanguageOption[] = [
  { id: "javascript", label: "JavaScript", extensions: [".js", ".jsx", ".mjs", ".cjs"], executionKind: "executable" },
  { id: "typescript", label: "TypeScript", extensions: [".ts", ".tsx", ".mts", ".cts"], executionKind: "executable" },
  { id: "python", label: "Python", extensions: [".py", ".pyw"], executionKind: "executable" },
  { id: "json", label: "JSON", extensions: [".json"], executionKind: "non-executable" },
  { id: "sql", label: "SQL", extensions: [".sql"], executionKind: "executable" },
  { id: "bash", label: "Shell", extensions: [".sh", ".bash"], executionKind: "executable" },
  { id: "go", label: "Go", extensions: [".go"], executionKind: "executable" },
  { id: "rust", label: "Rust", extensions: [".rs"], executionKind: "executable" },
  { id: "java", label: "Java", extensions: [".java"], executionKind: "executable" },
  { id: "c", label: "C", extensions: [".c", ".h"], executionKind: "executable" },
  { id: "cpp", label: "C++", extensions: [".cpp", ".cc", ".cxx", ".hpp", ".hh", ".hxx", ".C"], executionKind: "executable" },
  { id: "csharp", label: "C#", extensions: [".cs"], executionKind: "executable" },
  { id: "ruby", label: "Ruby", extensions: [".rb"], executionKind: "executable" },
  { id: "php", label: "PHP", extensions: [".php"], executionKind: "executable" },
  { id: "swift", label: "Swift", extensions: [".swift"], executionKind: "executable" },
  { id: "r", label: "R", extensions: [".r"], executionKind: "executable" },
  { id: "perl", label: "Perl", extensions: [".pl", ".pm"], executionKind: "executable" },
  { id: "lua", label: "Lua", extensions: [".lua"], executionKind: "executable" },
  // Compile-only: Run typesets the document on the server runtime and shows
  // the PDF it produced instead of program output.
  { id: "latex", label: "LaTeX", extensions: [".tex"], executionKind: "executable" },
  { id: "yaml", label: "YAML", extensions: [".yaml", ".yml"], executionKind: "non-executable" },
  // Vibe windows and imported pages carry "html" as their language; it is
  // its own selector so a saved id is never rewritten to "xml" by the editor.
  { id: "html", label: "HTML", extensions: [".html", ".htm"], executionKind: "non-executable" },
  { id: "xml", label: "XML", extensions: [".xml", ".xhtml"], executionKind: "non-executable" },
  { id: "css", label: "CSS", extensions: [".css"], executionKind: "non-executable" },
  { id: "ini", label: "INI / TOML", extensions: [".ini", ".toml"], executionKind: "non-executable" },
  { id: "markdown", label: "Markdown", extensions: [], executionKind: "non-executable" },
  { id: "plaintext", label: "Plain text", extensions: [], executionKind: "non-executable" },
];

export const CODE_LANGUAGE_IDS: string[] = CODE_LANGUAGE_OPTIONS.map(
  (option) => option.id,
);
const CODE_LANGUAGE_ID_SET: ReadonlySet<string> = new Set(CODE_LANGUAGE_IDS);

/** Exact selector identity for persisted/command metadata, or null when foreign. */
export function codeLanguageId(value: unknown): string | null {
  return typeof value === "string" && CODE_LANGUAGE_ID_SET.has(value)
    ? value
    : null;
}

/** Resolve partial shared-draft metadata without resetting an omitted language. */
export function codeDraftLanguage(
  value: unknown,
  currentLanguage: string,
): string | null {
  return value === undefined
    ? (codeLanguageId(currentLanguage) ?? "plaintext")
    : codeLanguageId(value);
}

/**
 * Executable languages whose Run is held back for now. PHP and R can only be
 * published by an owner-signed approval that has not been given yet, so the
 * editor keeps them editable but groups them with the non-runnable entries and
 * registers no provider for them. Remove an id here to offer Run again.
 */
export const HELD_EXECUTION_LANGUAGE_IDS: ReadonlySet<string> = new Set([
  "php",
  "r",
]);

/** Every selector-backed language that can acquire a real execution provider. */
export const EXECUTABLE_LANGUAGE_IDS: ReadonlySet<string> = new Set(
  CODE_LANGUAGE_OPTIONS.filter(
    (option) =>
      option.executionKind === "executable" &&
      !HELD_EXECUTION_LANGUAGE_IDS.has(option.id),
  ).map((option) => option.id),
);

/** Data/markup/prose entries that intentionally never expose Run. */
export const NON_EXECUTABLE_LANGUAGE_IDS: ReadonlySet<string> = new Set(
  CODE_LANGUAGE_OPTIONS.filter(
    (option) => option.executionKind === "non-executable",
  ).map((option) => option.id),
);

/**
 * Languages that actually execute in the browser runtime (JS sandbox iframe,
 * Pyodide worker, sucrase→JS for TS, sql.js for SQL, the pinned Lua 5.5 WASM
 * worker, the pinned .NET 10 / Roslyn 5.6 C# worker, the pinned Ruby 4.0 WASI
 * worker, and the pinned Bashkit 0.16 WASM worker). Every other language in
 * CODE_LANGUAGE_OPTIONS is editable / highlightable but cannot run here.
 * Single source of truth so the run dispatch and the Run-button gating never
 * drift.
 */
export const RUNNABLE_LANGUAGE_IDS: ReadonlySet<string> = new Set([
  "javascript",
  "typescript",
  "python",
  "sql",
  "lua",
  "csharp",
  "ruby",
  "bash",
]);

/**
 * Rollback-compatible static answer for the built-in browser providers.
 * New Run surfaces must use the provider registry's capability receipt instead
 * of expanding this set before an authenticated provider exists.
 */
export function isRunnableLanguage(languageId: string): boolean {
  return RUNNABLE_LANGUAGE_IDS.has(languageId);
}

/** Whether a selector entry represents executable source at all. */
export function isExecutableLanguage(languageId: string): boolean {
  return EXECUTABLE_LANGUAGE_IDS.has(languageId);
}

/**
 * Markup whose Run command renders the buffer as a page instead of executing
 * it. These stay non-executable (no execution provider is ever registered for
 * them); the Code window runs the page in the AI window's sandbox frame
 * (scripts allowed, opaque origin, no network).
 */
export const PREVIEW_LANGUAGE_IDS: ReadonlySet<string> = new Set(["html"]);

/** Whether Run previews this language as a sandboxed page. */
export function isPreviewLanguage(languageId: string): boolean {
  return PREVIEW_LANGUAGE_IDS.has(languageId);
}

/**
 * Where a selector entry can run at all — the question a person is really
 * answering when they pick a language:
 *
 *   browser — runs here, in this tab, with no account and no server.
 *   server  — executable, but only through the attested remote runtime, which
 *             a deployment has to activate (docs/ops/code-executor-c-activation.md).
 *   preview — markup with no execution provider; Run shows it as a sandboxed page.
 *   none    — data / markup / prose that deliberately never exposes Run.
 *
 * This is the STATIC shape of a language, not today's availability: whether the
 * remote runtime is reachable right now is a capability receipt, and the Run
 * command reads that. Keeping the two apart is what lets the selector say
 * "needs the server runtime" without promising that it is up.
 */
export type CodeLanguageRunKind = "browser" | "server" | "preview" | "none";

export const CODE_LANGUAGE_RUN_KINDS: readonly CodeLanguageRunKind[] = [
  "browser",
  "server",
  "preview",
  "none",
];

export function codeLanguageRunKind(languageId: string): CodeLanguageRunKind {
  if (RUNNABLE_LANGUAGE_IDS.has(languageId)) return "browser";
  if (EXECUTABLE_LANGUAGE_IDS.has(languageId)) return "server";
  return PREVIEW_LANGUAGE_IDS.has(languageId) ? "preview" : "none";
}

/**
 * The selector's entries grouped by that answer, each group keeping the
 * curated order of CODE_LANGUAGE_OPTIONS. Every entry lands in exactly one
 * group, so a language can never fall out of the list by being grouped.
 */
export const CODE_LANGUAGE_RUN_GROUPS: readonly {
  readonly kind: CodeLanguageRunKind;
  readonly options: readonly CodeLanguageOption[];
}[] = CODE_LANGUAGE_RUN_KINDS.map((kind) => ({
  kind,
  options: CODE_LANGUAGE_OPTIONS.filter(
    (option) => codeLanguageRunKind(option.id) === kind,
  ),
}));

const LANGUAGE_BY_EXTENSION: Map<string, string> = (() => {
  const map = new Map<string, string>();
  for (const option of CODE_LANGUAGE_OPTIONS) {
    for (const extension of option.extensions) {
      map.set(extension === ".C" ? extension : extension.toLowerCase(), option.id);
    }
  }
  return map;
})();

/**
 * Map a file extension (with leading dot, any case) to a code language id, or
 * null when the extension is not a recognised code/structured-source file.
 */
export function languageForExtension(extension: string): string | null {
  // Uppercase .C is the long-standing Unix convention for C++ source. It is
  // the one suffix where blindly lowercasing changes the language.
  const key = extension === ".C" ? extension : extension.toLowerCase();
  return LANGUAGE_BY_EXTENSION.get(key) ?? null;
}

/**
 * Text files a LaTeX project keeps beside its .tex sources that no editor
 * language claims: bibliographies and their styles, document classes and
 * packages, font and language definitions, figure sources and data. The Code
 * workspace creates and renames them as plain text under the name given.
 */
const LATEX_PROJECT_TEXT_EXTENSIONS: ReadonlySet<string> = new Set([
  ".bib", ".bst", ".bbx", ".cbx", ".lbx", ".dbx", ".sty", ".cls", ".clo",
  ".def", ".cfg", ".fd", ".ldf", ".dtx", ".ins", ".ltx", ".tikz", ".pgf",
  ".mp", ".asy", ".gp", ".txt", ".csv", ".tsv", ".dat",
]);

export function isLatexProjectTextExtension(extension: string): boolean {
  return LATEX_PROJECT_TEXT_EXTENSIONS.has(extension.toLowerCase());
}

/** The conventional file extension for a code language id (defaults to .txt). */
export function extensionForLanguage(languageId: string): string {
  const option = CODE_LANGUAGE_OPTIONS.find((entry) => entry.id === languageId);
  return option?.extensions[0] ?? ".txt";
}
