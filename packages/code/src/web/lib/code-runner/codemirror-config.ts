/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Pure, headless CodeMirror 6 configuration for the MRNote code editor.
 *
 * Everything in this module is framework-agnostic: it returns CM6
 * `Extension`s (or plain values) and never touches React or the DOM beyond
 * what CM6 itself does. That keeps it unit-testable in node — an
 * `EditorState` can be constructed from these extensions without a layout
 * engine — and keeps `CodeMirrorEditor.tsx` thin.
 *
 * Stage 1 of the textarea -> CodeMirror migration. The existing
 * `CodeEditor.tsx` (textarea overlay) is untouched; this config only powers
 * the new, flag-gated `CodeMirrorEditor.tsx`.
 */
import {
  language as cmLanguageFacet,
  syntaxTree,
} from "@codemirror/language";
import { linter, type Diagnostic } from "@codemirror/lint";
import { oneDark } from "@codemirror/theme-one-dark";
import { EditorView, type ViewUpdate } from "@codemirror/view";
import { StateEffect, type EditorState, type Extension } from "@codemirror/state";

import {
  isPythonLspLanguage,
  isTsLspLanguage,
  parsePythonLintResult,
  pyflakesWarningsToCm,
  tsLintSource,
} from "./lsp-diagnostics";
import {
  bracketDiagnostics,
  supportsBracketDiagnostics,
  type BracketDiagnostic,
} from "./bracket-diagnostics";
import {
  normalizeLanguage,
  type SupportedLanguage,
} from "./language-metadata";

export { normalizeLanguage };
export type { SupportedLanguage };

/**
 * The languages whose CM6 grammar is a `@codemirror/legacy-modes` STREAM parser
 * (no Lezer tree): go / rust / java / c / cpp / csharp / ruby / swift / r /
 * perl / lua / yaml / toml / bash. (`shell`/`sh`/`zsh` normalize to `bash`.)
 * These render a fold gutter but, lacking a syntax tree, CM6's default
 * tree-based folding never folds them — so CodeMirrorEditor mounts the indent +
 * `#region` `foldService` (lib/code-runner/editor-cm-extensions.ts
 * `streamFoldExtension`) ONLY for these. The tree-backed languages
 * (js/ts/jsx/tsx/python/css/html/json/sql/markdown/php) already fold via their
 * grammar and are intentionally excluded.
 */
const STREAM_MODE_LANGUAGES: ReadonlySet<SupportedLanguage> = new Set([
  "bash",
  "csharp",
  "ruby",
  "swift",
  "r",
  "perl",
  "lua",
  "latex",
  "toml",
]);

/**
 * Whether `language` is highlighted by a stream parser (see
 * {@link STREAM_MODE_LANGUAGES}) and therefore needs the explicit fold service
 * for its gutter to actually fold. Accepts any raw `language` prop string (it
 * normalizes first), so an alias like `"sh"` or `"golang"` is recognised.
 */
export function isStreamModeLanguage(language: string | undefined | null): boolean {
  return STREAM_MODE_LANGUAGES.has(normalizeLanguage(language));
}

/**
 * The TREE-BACKED languages whose CM6 grammar IS a Lezer parser (js / ts / jsx /
 * tsx / python / css / html / json / php / sql / markdown). They already fold via
 * their grammar's `foldNodeProp` (syntax / indentation folding), but CM6's
 * default tree folding does NOT understand the `// #region` / `# region` marker
 * convention. CodeMirrorEditor mounts the region-ONLY `foldService`
 * (lib/code-runner/editor-cm-extensions.ts `regionFoldExtension`) for these so
 * `#region`/`#endregion` markers fold too — composed with, not replacing, the
 * tree fold (the region service returns null off a marker line, falling through
 * to the grammar's own folding). The stream-mode languages (see
 * {@link STREAM_MODE_LANGUAGES}) instead get `streamFoldExtension`, which bundles
 * the SAME region service PLUS indentation folding (they have no tree). `unknown`
 * (plain text) is intentionally excluded — it has neither tree nor a meaningful
 * comment syntax to anchor markers.
 */
const TREE_FOLD_LANGUAGES: ReadonlySet<SupportedLanguage> = new Set([
  "javascript",
  "typescript",
  "jsx",
  "tsx",
  "python",
  // Real Lezer grammars, so their own syntax folding must be left to run: an
  // indentation fold service answers first and would shadow it.
  "c",
  "cpp",
  "go",
  "java",
  "rust",
  "yaml",
  "css",
  "html",
  "json",
  "php",
  "sql",
  "markdown",
]);

/**
 * Whether `language` is highlighted by a Lezer-tree grammar (see
 * {@link TREE_FOLD_LANGUAGES}) and therefore folds via its grammar already — so
 * it needs ONLY the supplementary `#region` marker fold service, not the full
 * stream-mode indent+region one. Accepts any raw `language` prop string (it
 * normalizes first), so an alias like `"py"` or `"md"` is recognised.
 */
export function isTreeFoldLanguage(language: string | undefined | null): boolean {
  return TREE_FOLD_LANGUAGES.has(normalizeLanguage(language));
}

/**
 * Parse-error position derived from a `JSON.parse` failure message. Browsers
 * report the offending position differently:
 *  - V8 (Chrome/Node): "... at position 12" or "... at line 2 column 3 ..."
 *  - SpiderMonkey/WebKit: "... at line 2 column 3 of the JSON data"
 *
 * We extract a 0-based character offset where possible; callers clamp it to
 * the document length.
 */
export function jsonErrorOffset(message: string, source: string): number {
  const posMatch = /at position (\d+)/i.exec(message);
  if (posMatch) {
    return Number(posMatch[1]);
  }
  const lineColMatch = /line (\d+) column (\d+)/i.exec(message);
  if (lineColMatch) {
    const line = Number(lineColMatch[1]);
    const column = Number(lineColMatch[2]);
    const lines = source.split("\n");
    let offset = 0;
    for (let i = 0; i < line - 1 && i < lines.length; i += 1) {
      offset += lines[i].length + 1; // +1 for the consumed "\n"
    }
    return offset + (column - 1);
  }
  return 0;
}

/**
 * Run `JSON.parse` over the document text and surface any syntax error as a
 * single CM6 diagnostic anchored at (or near) the failing character. Empty /
 * whitespace-only documents are treated as clean (no diagnostic) so a fresh
 * editor isn't immediately "errored".
 *
 * Exported standalone (not just inside an extension) so it can be unit-tested
 * without constructing an EditorView.
 */
export function jsonDiagnostics(text: string): Diagnostic[] {
  if (text.trim().length === 0) return [];
  try {
    JSON.parse(text);
    return [];
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    const raw = jsonErrorOffset(message, text);
    // CodeMirror anchors a diagnostic at a character position, so `from` must
    // stay strictly inside [0, text.length): a value of text.length points past
    // the document end and breaks lint-gutter rendering. text is non-empty here
    // (empty/whitespace returned above), so clamp to the last addressable char.
    const from = Math.max(0, Math.min(raw, text.length - 1));
    const to = Math.min(text.length, from + 1);
    return [
      {
        from,
        to: to > from ? to : from,
        severity: "error",
        message,
      },
    ];
  }
}

/** CM6 `linter` extension that flags invalid JSON via {@link jsonDiagnostics}. */
export function jsonLinter(): Extension {
  return linter((view) => jsonDiagnostics(view.state.doc.toString()));
}

/**
 * Lightweight JavaScript/TypeScript syntax-error linter. We do NOT execute the
 * code — `new Function(text)` only *parses* the body (the function is never
 * called), so a syntax error throws a `SyntaxError` at construction time while
 * runtime behaviour (timers, network, infinite loops) never runs.
 *
 * This catches the common, unambiguous mistakes (unbalanced brackets, stray
 * tokens, bad string literals). It is deliberately coarse: it cannot validate
 * TypeScript type annotations (the JS parser would reject `x: number`), so for
 * TS we strip nothing and simply skip reporting when the only failure looks like
 * a type annotation — we surface a diagnostic only when the message is a generic
 * JS syntax error. Empty/whitespace docs are treated as clean.
 *
 * Exported standalone so it can be unit-tested without an EditorView.
 */
export function jsSyntaxDiagnostics(text: string): Diagnostic[] {
  if (text.trim().length === 0) return [];
  try {
    // Parse-only: constructing the Function compiles the body but never invokes
    // it. ES modules use import/export which `new Function` rejects, so bail to
    // "clean" when the source clearly uses module syntax (can't be parsed this
    // way without false positives).
    if (/^\s*(import|export)\b/m.test(text)) return [];
    new Function(text);
    return [];
  } catch (error) {
    if (!(error instanceof SyntaxError)) return [];
    const message = error.message;
    return [
      {
        // We can't recover a reliable offset from a SyntaxError across engines,
        // so anchor at the document start; the message names the problem.
        from: 0,
        to: Math.min(text.length, 1) > 0 ? 1 : 0,
        severity: "error",
        message,
      },
    ];
  }
}

/**
 * Synchronous TypeScript fallback linter — a deliberate no-op.
 *
 * The JS `new Function` parser rejects type annotations as syntax errors (a
 * false positive), so this synchronous placeholder reports nothing. The REAL
 * TypeScript diagnostics (syntactic + semantic/type errors) now come from the
 * async {@link lintExtensions} path, which lazy-loads the installed TypeScript
 * compiler (see lib/code-runner/lsp-diagnostics.ts). This export and
 * {@link linterForLanguage} remain the synchronous, value-only selector used by
 * unit tests and any caller that wants a pure `text -> Diagnostic[]` function.
 */
export const tsSyntaxDiagnostics: LanguageLinter = () => [];

/**
 * Pluggable per-language lint interface. A `LanguageLinter` receives the
 * current document text and returns CM6 diagnostics. Languages without a real
 * linter use {@link noopLinter}, which returns no diagnostics — the lint
 * gutter still renders (so the UI is consistent) but never flags anything.
 */
export type LanguageLinter = (text: string) => Diagnostic[];

/** A linter that never reports problems. Placeholder for unlinted languages. */
export const noopLinter: LanguageLinter = () => [];

/**
 * Tree-backed grammars that can surface structural parse errors without
 * executing user code or loading a runtime. JavaScript/TypeScript and Python
 * use their richer compiler/runtime lint paths; JSON keeps its precise
 * JSON.parse diagnostic.
 */
const TREE_DIAGNOSTIC_LANGUAGES: ReadonlySet<SupportedLanguage> = new Set([
  "css",
  "html",
  "sql",
  // Real Lezer grammars, so these report where the parse actually broke rather
  // than only that something never closed.
  "c",
  "go",
  "yaml",
  // Not Java, Rust, PHP or C++: those grammars predate syntax people now
  // write (records and switch arrows, let-else, readonly promoted
  // properties, templated lambdas), so valid code read as broken. They get
  // the structural check instead (see bracket-diagnostics).
]);

const TREE_DIAGNOSTIC_LABELS: Partial<Record<SupportedLanguage, string>> = {
  css: "CSS",
  html: "HTML/XML",
  php: "PHP",
  sql: "SQL",
  c: "C",
  cpp: "C++",
  go: "Go",
  java: "Java",
  rust: "Rust",
  yaml: "YAML",
};

/** Keep malformed generated files from creating an unbounded Problems list. */
export const TREE_SYNTAX_DIAGNOSTIC_LIMIT = 200;
/** Bound the tree walk itself when a very large, mostly-valid file is open. */
const TREE_SYNTAX_NODE_BUDGET = 100_000;

/**
 * Project Lezer error nodes from a live CodeMirror state into clickable
 * diagnostics. The grammar is already maintained by CodeMirror, so this adds no
 * parser dependency and never executes the document.
 *
 * Error nodes are often zero-width (for example a missing `}` at EOF). Those
 * are anchored to the nearest real character so the lint gutter and Problems
 * jump remain operable. Duplicate nested error nodes are collapsed, and both
 * the result count and traversal work are bounded for large malformed files.
 */
export function treeSyntaxDiagnostics(
  state: EditorState,
  language: string,
  maxDiagnostics = TREE_SYNTAX_DIAGNOSTIC_LIMIT,
  messages: DiagnosticMessages = DEFAULT_DIAGNOSTIC_MESSAGES,
): Diagnostic[] {
  const normalized = normalizeLanguage(language);
  if (
    state.doc.length === 0 ||
    maxDiagnostics <= 0 ||
    !TREE_DIAGNOSTIC_LANGUAGES.has(normalized)
  ) {
    return [];
  }

  const limit = Math.min(
    TREE_SYNTAX_DIAGNOSTIC_LIMIT,
    Math.max(1, Math.floor(maxDiagnostics)),
  );
  const label = TREE_DIAGNOSTIC_LABELS[normalized] ?? normalized;
  const message = messages("syntax.invalid", { language: label });
  const diagnostics: Diagnostic[] = [];
  const seen = new Set<string>();
  const cursor = syntaxTree(state).cursor();
  let visited = 0;

  do {
    visited += 1;
    if (!cursor.type.isError) continue;

    const docLength = state.doc.length;
    let from = Math.max(0, Math.min(cursor.from, docLength));
    let to = Math.max(from, Math.min(cursor.to, docLength));
    if (from === to) {
      if (from < docLength) {
        to = from + 1;
      } else {
        from = Math.max(0, docLength - 1);
        to = docLength;
      }
    }
    const key = `${from}:${to}`;
    if (seen.has(key)) continue;
    seen.add(key);
    diagnostics.push({ from, to, severity: "error", message });
  } while (
    diagnostics.length < limit &&
    visited < TREE_SYNTAX_NODE_BUDGET &&
    cursor.next()
  );

  return diagnostics;
}

/**
 * Resolve the {@link LanguageLinter} for a language.
 *
 * REAL linters (produce diagnostics):
 *  - `json`        -> {@link jsonDiagnostics} (JSON.parse).
 *  - `javascript`  -> {@link jsSyntaxDiagnostics} (parse-only `new Function`).
 *  - `jsx`         -> {@link jsSyntaxDiagnostics} (plain JS body parse; JSX in a
 *                     bare body would not parse, so this is best-effort and only
 *                     flags non-JSX syntax errors).
 *
 * PURE fallback no-op linters (gutter renders, nothing is flagged):
 *  - `typescript` / `tsx` -> {@link tsSyntaxDiagnostics} (a real TS check needs
 *                     the TS compiler; the JS parser would false-positive on type
 *                     annotations, so we intentionally report nothing).
 *  - `python` / `sql` / `css` / `html` / `markdown` / `unknown` -> {@link noopLinter}.
 *    SQL/CSS/HTML/PHP get real state-aware grammar diagnostics through
 *    {@link lintExtensions}; this value-only selector has no syntax tree.
 *
 * Adding a new real linter is a one-line case here plus the implementation.
 */
export function linterForLanguage(language: string): LanguageLinter {
  const lang = normalizeLanguage(language);
  switch (lang) {
    case "json":
      return jsonDiagnostics;
    case "javascript":
    case "jsx":
      return jsSyntaxDiagnostics;
    case "typescript":
    case "tsx":
      return tsSyntaxDiagnostics;
    default:
      return noopLinter;
  }
}

/**
 * Async CM6 lint source for the TypeScript-compiler languages (js/jsx/ts/tsx):
 * defers to {@link tsLintSource}, which lazy-loads the compiler and returns real
 * syntactic + semantic diagnostics. Non-throwing — degrades to `[]`.
 */
export type ProjectDiagnosticSource = (
  source: string,
) => Promise<readonly Diagnostic[] | null>;

/**
 * Something outside this document changed what its linter would report: a
 * peer window's exported type, a workspace file the project compiles against.
 * CodeMirror's lint plugin only re-runs for a document edit, a lint
 * reconfiguration or a `needsRefresh` transaction, and `forceLinting` runs
 * nothing unless one of those is pending; dispatching this effect first makes
 * the force count. (Rebuilding the lint extensions used to do this by
 * accident whenever the host re-rendered with a new translator.)
 */
export const refreshLintEffect = StateEffect.define<null>();

/** A lint `needsRefresh` that answers to {@link refreshLintEffect}. */
export function lintRefreshRequested(update: ViewUpdate): boolean {
  return update.transactions.some((transaction) =>
    transaction.effects.some((effect) => effect.is(refreshLintEffect)),
  );
}

function tsLinterExtension(
  language: string,
  projectDiagnostics?: ProjectDiagnosticSource,
): Extension {
  return linter(
    async (view) => {
      const source = view.state.doc.toString();
      if (projectDiagnostics) {
        const diagnostics = await projectDiagnostics(source);
        if (diagnostics) return diagnostics;
      }
      return tsLintSource(source, language);
    },
    { needsRefresh: lintRefreshRequested },
  );
}

/**
 * Async CM6 lint source for Python: parse-checks the document via Pyodide
 * ({@link lintPython}) and maps the first SyntaxError to a diagnostic, PLUS runs
 * pyflakes (AST-only, never executes code) for undefined-name, unused-import, and
 * unused/shadowed-variable warnings. Resolves to `[]` when the runtime is
 * unavailable so a fresh editor never squiggles.
 *
 * Imported lazily so Pyodide orchestration stays out of the synchronous config
 * path and the linter only touches the runtime when a Python doc is edited.
 */
function pythonLinterExtension(): Extension {
  return linter(async (view) => {
    const doc = view.state.doc.toString();
    if (doc.trim().length === 0) return [];
    try {
      const { lintPython } = await import("@mrnote/core/web/lib/code-runner/python-runner.ts");
      const { lint, warnings } = await lintPython(doc);
      // Syntax errors take priority — surface them and skip semantic warnings
      // (pyflakes is also skipped in the worker when there's a syntax error).
      if (!lint.ok) {
        return parsePythonLintResult(JSON.stringify(lint), doc);
      }
      // Clean parse: surface any pyflakes semantic warnings.
      return pyflakesWarningsToCm(warnings, doc);
    } catch {
      return [];
    }
  });
}

/**
 * Lint-related extensions for a language: the language's linter(s) wired through
 * CM6's `linter()`. Always returns an array (possibly with a no-op linter) so
 * the caller can unconditionally include a `lintGutter()`.
 *
 * REAL diagnostics:
 *  - json                 -> {@link jsonDiagnostics} (synchronous JSON.parse).
 *  - js/jsx/ts/tsx        -> {@link tsLintSource} (TypeScript compiler:
 *                            syntactic + best-effort semantic/type errors).
 *  - python               -> Pyodide parse-check (first SyntaxError).
 *
 *  - css/html/php/sql     -> bounded CodeMirror grammar error nodes.
 *
 * Other languages fall back to the synchronous {@link linterForLanguage} (a
 * no-op for stream/plain-text languages), so the gutter still renders
 * consistently.
 */
/**
 * Resolves this module's OWN diagnostic copy. Compiler and browser messages
 * arrive in their own language and are passed through untouched; everything
 * this file authors goes through here so it can be localized.
 */
export type DiagnosticMessages = (
  key: string,
  params?: Record<string, string>,
) => string;

const DEFAULT_DIAGNOSTIC_MESSAGES: DiagnosticMessages = (key, params) => {
  switch (key) {
    case "syntax.invalid":
      return `Invalid ${params?.language ?? ""} syntax`.replace("  ", " ");
    case "bracket.unclosed":
      return `Unclosed "${params?.token ?? ""}"`;
    case "bracket.unmatched":
      return `Missing matching delimiter for "${params?.token ?? ""}"`;
    case "bracket.unterminatedString":
      return "Unterminated string";
    case "bracket.unterminatedComment":
      return "Unterminated block comment";
    default:
      return key;
  }
};

/**
 * A registered language provider's diagnostics, already resolved to this
 * buffer. Supplied by the editor; absent everywhere the registry is not wired.
 */
export type ProviderDiagnosticSource = (
  text: string,
) => Promise<
  Array<{
    from: number;
    to: number;
    message: string;
    severity: "error" | "warning" | "info";
  }>
>;

export function lintExtensions(
  language: string,
  projectDiagnostics?: ProjectDiagnosticSource,
  messages: DiagnosticMessages = DEFAULT_DIAGNOSTIC_MESSAGES,
  providerDiagnostics?: ProviderDiagnosticSource,
): Extension {
  if (isTsLspLanguage(language)) {
    return tsLinterExtension(language, projectDiagnostics);
  }
  if (isPythonLspLanguage(language)) {
    return pythonLinterExtension();
  }
  // A registered provider, if one claims this language. It REPLACES the
  // structural fallbacks below rather than adding to them: those exist because
  // nothing better was available, and reporting "this bracket never closes"
  // alongside a real parser's account of the same line is worse than either.
  if (providerDiagnostics) {
    return linter(async (view) => {
      const text = view.state.doc.toString();
      let reported;
      try {
        reported = await providerDiagnostics(text);
      } catch {
        // A provider that throws reports nothing, never a broken panel.
        return [];
      }
      const limit = view.state.doc.length;
      return reported
        .filter(
          (diagnostic) =>
            Number.isFinite(diagnostic.from) &&
            Number.isFinite(diagnostic.to) &&
            diagnostic.from >= 0 &&
            diagnostic.to >= diagnostic.from &&
            // The buffer may have changed under a slow provider; a marker past
            // the end would throw inside CodeMirror.
            diagnostic.to <= limit &&
            typeof diagnostic.message === "string" &&
            diagnostic.message.length > 0,
        )
        .map((diagnostic) => ({
          from: diagnostic.from,
          to: diagnostic.to,
          severity: diagnostic.severity,
          message: diagnostic.message,
        }));
    });
  }
  if (TREE_DIAGNOSTIC_LANGUAGES.has(normalizeLanguage(language))) {
    return linter(
      (view) =>
        treeSyntaxDiagnostics(view.state, language, undefined, messages),
      {
        // Lazy grammar installation is a Compartment reconfiguration, not a
        // document edit. CodeMirror's lint plugin otherwise keeps the empty
        // result produced during the plain-text gap until the user types.
        needsRefresh: (update) =>
          update.startState.facet(cmLanguageFacet) !==
          update.state.facet(cmLanguageFacet),
      },
    );
  }
  if (supportsBracketDiagnostics(normalizeLanguage(language))) {
    // No language server and no syntax tree, so the file's structure is the
    // only thing that can be checked here — and "this never closes" is exactly
    // the error a parser would have reported first.
    return linter((view) =>
      bracketDiagnostics(
        view.state.doc.toString(),
        normalizeLanguage(language),
      ).map((diagnostic) => ({
        from: diagnostic.from,
        to: diagnostic.to,
        severity: diagnostic.severity,
        message: bracketDiagnosticMessage(diagnostic, messages),
      })),
    );
  }
  const fn = linterForLanguage(language);
  return linter((view) => fn(view.state.doc.toString()));
}

/**
 * English message for one structural diagnostic. The editor's diagnostics are
 * plain strings (CodeMirror renders them directly), and this module has no
 * translator; the codes stay machine-readable for a caller that wants to
 * localize them.
 */
function bracketDiagnosticMessage(
  diagnostic: BracketDiagnostic,
  messages: DiagnosticMessages,
): string {
  switch (diagnostic.code) {
    case "unclosed_bracket":
      return messages("bracket.unclosed", { token: diagnostic.token });
    case "unmatched_bracket":
      return messages("bracket.unmatched", { token: diagnostic.token });
    case "unterminated_string":
      return messages("bracket.unterminatedString");
    case "unterminated_comment":
      return messages("bracket.unterminatedComment");
  }
}

/**
 * Resolve a 1-based line number to a character offset range covering that whole
 * line in `text`, clamped to [1, lineCount]. Used by the go-to-line bar to build
 * an EditorView selection/scroll without depending on a live view's geometry.
 *
 * Returns the line's start offset (`from`) and end-of-line offset (`to`, before
 * the newline). Out-of-range requests clamp to the first / last line so a stray
 * "9999" jumps to the document end rather than erroring.
 */
export function lineToOffsetRange(
  text: string,
  line: number,
): { line: number; from: number; to: number } {
  const lines = text.split("\n");
  const target = Math.max(1, Math.min(Math.floor(line) || 1, lines.length));
  let from = 0;
  for (let i = 0; i < target - 1; i += 1) {
    from += lines[i].length + 1; // +1 for the consumed "\n"
  }
  const to = from + lines[target - 1].length;
  return { line: target, from, to };
}

/**
 * Accessible-name extension for the editor's editing surface.
 *
 * CM6 renders the editable region as `<div contenteditable role="textbox">`
 * (`.cm-content`); a bare textbox with no accessible name fails axe's
 * `aria-input-field-name`. This adds an `aria-label` to that node via
 * `EditorView.contentAttributes`, giving screen-reader users a name for the
 * editor without altering any editing behaviour.
 *
 * The label is supplied by the caller so it can be localized (the calling
 * React component threads a translated string). A blank/whitespace label is
 * treated as "no label" and returns an empty (inert) extension so we never set
 * an empty `aria-label`.
 */
export function contentAttributesExtension(label: string): Extension {
  if (!label || !label.trim()) return [];
  return EditorView.contentAttributes.of({ "aria-label": label });
}

/**
 * Bracket-pair colorization palette (VSCode parity): six depth-cycling foreground
 * colors keyed by the `cm-bracket-depth-N` classes that
 * {@link bracketColorizationExtension} (lib/code-runner/editor-cm-extensions.ts)
 * stamps on each structural bracket, plus the render-whitespace glyph color. Two
 * palettes so brackets stay legible on the dark and light surfaces. Built as a
 * theme so it composes with the base theme below.
 */
const BRACKET_DEPTH_PALETTE: { dark: string[]; light: string[]; wsDark: string; wsLight: string } = {
  // Tuned to read on oneDark; mirrors VSCode's default bracket colors.
  dark: ["#ffd700", "#da70d6", "#179fff", "#5cb85c", "#ff8c42", "#f06292"],
  // GitHub-light-ish hues with enough separation on a white surface.
  light: ["#b58900", "#a626a4", "#0550ae", "#116329", "#bc5215", "#cf222e"],
  wsDark: "rgba(255,255,255,0.28)",
  wsLight: "rgba(15,23,42,0.30)",
};

/** Build the bracket-depth + whitespace-glyph CSS rules for a surface. */
function bracketAndWhitespaceTheme(isDark: boolean): Record<string, Record<string, string>> {
  const colors = isDark ? BRACKET_DEPTH_PALETTE.dark : BRACKET_DEPTH_PALETTE.light;
  const rules: Record<string, Record<string, string>> = {};
  colors.forEach((color, depth) => {
    rules[`.cm-bracket-depth-${depth}`] = { color, fontWeight: "600" };
  });
  // `highlightWhitespace` renders dots/arrows in `.cm-highlightSpace` /
  // `.cm-highlightTab` pseudo-elements; tint them faintly so they don't shout.
  const ws = isDark ? BRACKET_DEPTH_PALETTE.wsDark : BRACKET_DEPTH_PALETTE.wsLight;
  rules[".cm-highlightSpace:before"] = { color: ws };
  rules[".cm-highlightTab"] = { backgroundColor: "transparent" };
  rules[".cm-highlightTab:before"] = { borderColor: ws };
  // Debugger current statement: amber fill plus a solid inset rail, so the
  // paused line remains identifiable without relying on colour alone.
  rules[".cm-line.cm-debug-paused-line"] = {
    backgroundColor: isDark
      ? "rgba(245, 158, 11, 0.18)"
      : "rgba(245, 158, 11, 0.14)",
    boxShadow: "inset 3px 0 #f59e0b",
  };
  return rules;
}

/**
 * The autocomplete popup materializes with the Docs floating-surface motion
 * (`office-pop-in`, styles/office-motion.css, whose keyframes are global),
 * only when the user has not asked for reduced motion.
 */
const floatingSurfaceMotionTheme = EditorView.theme({
  "@media (prefers-reduced-motion: no-preference)": {
    ".cm-tooltip.cm-tooltip-autocomplete": {
      animation: "office-pop-in 140ms ease-out both",
    },
  },
});

/**
 * Theme extension. Dark uses CM6's `oneDark`; light uses a small theme that
 * matches the existing textarea editor's white surface so flipping the flag
 * doesn't visually jar. Returning an array for both keeps the type uniform and
 * lets us append the bracket-colorization / whitespace-glyph palette (which
 * applies only when those toggles add their extensions).
 */
export function themeExtension(isDark: boolean): Extension {
  const overlay = EditorView.theme(bracketAndWhitespaceTheme(isDark), { dark: isDark });
  if (isDark) return [oneDark, overlay, floatingSurfaceMotionTheme];
  return [
    floatingSurfaceMotionTheme,
    EditorView.theme(
      {
        "&": {
          backgroundColor: "#ffffff",
          color: "#24292f",
          "--cm-code-lens-color": "#475569",
          "--cm-inlay-hint-color": "#475569",
          "--cm-inlay-hint-bg": "rgba(71, 85, 105, 0.12)",
        },
        ".cm-gutters": {
          backgroundColor: "#f8fafa",
          color: "#64748b",
          border: "none",
          borderRight: "1px solid #eef2f1",
        },
        ".cm-activeLine": { backgroundColor: "rgba(0, 0, 0, 0.03)" },
        ".cm-activeLineGutter": { backgroundColor: "rgba(0, 0, 0, 0.04)" },
        ".cm-selectionMatch": { backgroundColor: "rgba(59, 130, 246, 0.18)" },
        ".cm-content": { caretColor: "#24292f" },
      },
      { dark: false },
    ),
    overlay,
  ];
}
