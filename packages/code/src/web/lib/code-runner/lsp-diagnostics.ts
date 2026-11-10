/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Full-language inline lint diagnostics for the MRNote CodeMirror 6 editor.
 *
 * This raises the code domain toward VSCode-level inline squiggles WITHOUT any
 * new dependency:
 *  - TS/JS/JSX/TSX: the already-installed TypeScript compiler API (`typescript`)
 *    produces real SYNTACTIC and (best-effort) SEMANTIC diagnostics — unbalanced
 *    brackets, stray tokens, AND type errors like "Type 'string' is not
 *    assignable to type 'number'". The compiler is a big module, so it is loaded
 *    lazily via `import("typescript")` inside the linter and stays OUT of the
 *    initial editor chunk.
 *  - Python: the already-loaded Pyodide runtime is asked to `compile()` the
 *    source (parse-only, never executed) and any `SyntaxError` (lineno/offset/
 *    msg) is mapped to a CM diagnostic.
 *
 * The *pure* mapping logic lives here so it is exhaustively unit-testable in
 * node (where `typescript` is installed): `tsDiagnosticsToCm`,
 * `pythonSyntaxToCm`, and the language → linter selector. Everything is bounded
 * and non-throwing: any internal failure degrades to `[]` (no diagnostics)
 * rather than breaking the editor.
 *
 * Live rendering of the squiggles inside CodeMirror and the Pyodide-worker
 * round-trip are browser-gated; see the unit tests for the mapping coverage.
 */
import type { Diagnostic } from "@codemirror/lint";
import type * as TsModule from "typescript";
import { localizeTypeScript } from "./typescript-locale";
import { TYPESCRIPT_AMBIENT_LIB } from "./typescript-ambient-lib";

import {
  normalizeLanguage,
  type SupportedLanguage,
} from "./language-metadata";

/**
 * A structural subset of `ts.Diagnostic` carrying exactly what the CM mapper
 * needs. Using a local shape (instead of importing the heavy `ts` types at the
 * value level) keeps {@link tsDiagnosticsToCm} a pure function that tests can
 * feed hand-built objects to, while real `ts.Diagnostic`s also satisfy it.
 */
export interface TsLikeDiagnostic {
  /** 0-based character offset of the problem in the source, if known. */
  start?: number;
  /** Length in characters of the flagged span, if known. */
  length?: number;
  /** Flattened, human-readable message text. */
  messageText: string;
  /** ts.DiagnosticCategory: 0=Warning, 1=Error, 2=Suggestion, 3=Message. */
  category: number;
  /** ts.Diagnostic error code (e.g. 2322). Used to filter lib noise. */
  code?: number;
}

/** CodeMirror severities; mirrors `@codemirror/lint`'s `Severity` union. */
type CmSeverity = Diagnostic["severity"];

/**
 * Map a `ts.DiagnosticCategory` to a CodeMirror severity.
 *  - Error (1)            -> "error"
 *  - Warning (0)          -> "warning"
 *  - Suggestion (2)       -> "hint"
 *  - Message (3) / other  -> "info"
 */
export function tsCategoryToSeverity(category: number): CmSeverity {
  switch (category) {
    case 1:
      return "error";
    case 0:
      return "warning";
    case 2:
      return "hint";
    case 3:
    default:
      return "info";
  }
}

/**
 * Diagnostic codes that, under our minimal synthetic lib (see
 * {@link buildCompilerHost}), are pure "missing standard-library type" noise
 * rather than real user mistakes. We do not ship the full DOM/ESNext `.d.ts`
 * (no fs in the browser, and bundling it is out of scope), so an uncommon
 * global the synthetic lib doesn't declare would otherwise false-positive as
 * "Cannot find name". Dropping these keeps the high-value assignability /
 * argument / syntax errors while avoiding spurious squiggles on valid code.
 *
 *  - 2304 Cannot find name 'X'.
 *  - 2552 Cannot find name 'X'. Did you mean 'Y'?
 *  - 2580/2581/2582/2583 Cannot find name (require/Buffer/describe/Map…) — env.
 *  - 2584 Cannot find name 'X'. Do you need to change your target library?
 *  - 2339 Property 'X' does not exist on type 'Y'. (members on un-libbed types)
 *  - 2693 'X' only refers to a type but is used as a value (un-libbed ctor).
 *  - 2585 'X' only refers to a type … change your target library (same).
 *  - 2318 Cannot find global type 'X' (a lib type the compiler asks for).
 */
const LIB_NOISE_CODES: ReadonlySet<number> = new Set([
  2304, 2552, 2580, 2581, 2582, 2583, 2584, 2339, 2693, 2585, 2318,
]);

/**
 * Map a list of TypeScript-style diagnostics to CodeMirror {@link Diagnostic}s,
 * clamping every span into `[0, doc.length]` so the lint gutter renders.
 *
 * Pure and total: any diagnostic with an out-of-range / missing `start` is
 * anchored at the document start; lib-noise codes are dropped (see
 * {@link LIB_NOISE_CODES}). Exported for exhaustive unit testing without the
 * real compiler.
 */
export function tsDiagnosticsToCm(
  diags: readonly TsLikeDiagnostic[],
  doc: string,
): Diagnostic[] {
  const docLen = doc.length;
  const out: Diagnostic[] = [];
  for (const d of diags) {
    if (typeof d.code === "number" && LIB_NOISE_CODES.has(d.code)) continue;
    const rawStart = typeof d.start === "number" && d.start >= 0 ? d.start : 0;
    const from = Math.max(0, Math.min(rawStart, docLen));
    const rawLen = typeof d.length === "number" && d.length > 0 ? d.length : 1;
    // A diagnostic must cover at least one character to render; clamp the end
    // into the document but keep `to > from` when there is room.
    let to = Math.min(docLen, from + rawLen);
    if (to <= from) to = Math.min(docLen, from + 1);
    const message =
      typeof d.messageText === "string" && d.messageText.length > 0
        ? d.messageText
        : "Problem";
    out.push({
      from,
      to,
      severity: tsCategoryToSeverity(d.category),
      message,
      source: "typescript",
    });
  }
  return out;
}

/**
 * ts.ScriptKind values for the four flavours we lint, kept as plain numbers so
 * the selector doesn't pull in the `ts` value module. JS=1, JSX=2, TS=3, TSX=4.
 */
const SCRIPT_KIND: Record<"javascript" | "jsx" | "typescript" | "tsx", number> = {
  javascript: 1,
  jsx: 2,
  typescript: 3,
  tsx: 4,
};

/**
 * Ambient built-ins so common ones (console, Array/String methods, Promise,
 * for...of, Map/Set …) don't false-positive under `noLib: true`, while real
 * assignability / argument-count errors stay; shared with hover/completion and
 * the open-window project.
 */
const SYNTHETIC_LIB = TYPESCRIPT_AMBIENT_LIB;

const LINT_FILE_BASENAMES: Record<number, string> = {
  1: "__mrnote_lint__.js",
  2: "__mrnote_lint__.jsx",
  3: "__mrnote_lint__.ts",
  4: "__mrnote_lint__.tsx",
};

const LIB_FILE_NAME = "__mrnote_lib__.d.ts";

/**
 * Build an in-memory {@link TsModule.CompilerHost} that serves exactly two
 * files — the snippet under lint and our {@link SYNTHETIC_LIB} — and returns
 * `undefined` for everything else. With `noLib: true` this never touches the
 * filesystem, so it works identically in node and the browser.
 */
function buildCompilerHost(
  ts: typeof TsModule,
  sourceFile: TsModule.SourceFile,
  libFile: TsModule.SourceFile,
  fileName: string,
): TsModule.CompilerHost {
  return {
    getSourceFile: (name) =>
      name === fileName ? sourceFile : name === LIB_FILE_NAME ? libFile : undefined,
    writeFile: () => {
      /* noEmit: never called */
    },
    getDefaultLibFileName: () => LIB_FILE_NAME,
    getCurrentDirectory: () => "",
    getDirectories: () => [],
    fileExists: (name) => name === fileName || name === LIB_FILE_NAME,
    readFile: () => undefined,
    getCanonicalFileName: (name) => name,
    useCaseSensitiveFileNames: () => true,
    getNewLine: () => "\n",
  };
}

/**
 * Run the TypeScript compiler over `source` and return CodeMirror diagnostics.
 *
 * `ts` is passed in (lazily `import()`ed by the caller) so this module never
 * pins the heavy compiler into the initial bundle. Empty/whitespace-only docs
 * are clean. Non-throwing: any internal compiler error degrades to `[]`.
 *
 * Combines syntactic diagnostics (always) with semantic diagnostics (type
 * checking) computed against the {@link SYNTHETIC_LIB}; lib-resolution noise is
 * filtered in {@link tsDiagnosticsToCm}.
 */
export function tsDiagnostics(
  ts: typeof TsModule,
  source: string,
  language: string,
): Diagnostic[] {
  if (source.trim().length === 0) return [];
  const lang = normalizeLanguage(language);
  const scriptKind =
    lang === "javascript" || lang === "jsx" || lang === "typescript" || lang === "tsx"
      ? SCRIPT_KIND[lang]
      : undefined;
  if (scriptKind === undefined) return [];
  const fileName = LINT_FILE_BASENAMES[scriptKind];
  try {
    const compilerOptions: TsModule.CompilerOptions = {
      noEmit: true,
      noLib: true,
      allowJs: true,
      checkJs: false,
      jsx: ts.JsxEmit.React,
      target: ts.ScriptTarget.Latest,
      module: ts.ModuleKind.ESNext,
      moduleResolution: ts.ModuleResolutionKind.Bundler,
      skipLibCheck: true,
      skipDefaultLibCheck: true,
      strict: false,
      types: [],
      isolatedModules: false,
    };
    const sourceFile = ts.createSourceFile(
      fileName,
      source,
      ts.ScriptTarget.Latest,
      /* setParentNodes */ true,
      scriptKind as TsModule.ScriptKind,
    );
    const libFile = ts.createSourceFile(
      LIB_FILE_NAME,
      SYNTHETIC_LIB,
      ts.ScriptTarget.Latest,
      true,
      ts.ScriptKind.TS,
    );
    const host = buildCompilerHost(ts, sourceFile, libFile, fileName);
    const program = ts.createProgram(
      [LIB_FILE_NAME, fileName],
      compilerOptions,
      host,
    );
    const syntactic = program.getSyntacticDiagnostics(sourceFile);
    // Semantic diagnostics are the type-checking layer (assignability, argument
    // counts, …). They can be more expensive and can occasionally throw on
    // pathological partial ASTs, so guard them separately: a syntactic-only
    // result is still useful.
    let semantic: readonly TsModule.Diagnostic[] = [];
    try {
      semantic = program.getSemanticDiagnostics(sourceFile);
    } catch {
      semantic = [];
    }
    const flat: TsLikeDiagnostic[] = [];
    for (const d of [...syntactic, ...semantic]) {
      if (d.file && d.file.fileName !== fileName) continue;
      flat.push({
        start: d.start,
        length: d.length,
        messageText: ts.flattenDiagnosticMessageText(d.messageText, "\n"),
        category: d.category,
        code: d.code,
      });
    }
    return tsDiagnosticsToCm(flat, source);
  } catch {
    return [];
  }
}

// ---------------------------------------------------------------------------
// Python (Pyodide) syntax diagnostics.
// ---------------------------------------------------------------------------

/**
 * A Python `SyntaxError` reduced to the fields the CM mapper needs. The Pyodide
 * worker reports these from `compile(src, "<lint>", "exec")` (parse-only, never
 * executed): `lineno` (1-based), `offset` (1-based column, may be null), and the
 * message. `endLineno`/`endOffset` are populated on Python 3.10+ so we can
 * highlight a real span instead of a single caret.
 */
export interface PythonSyntaxErrorLike {
  lineno?: number | null;
  offset?: number | null;
  endLineno?: number | null;
  endOffset?: number | null;
  msg?: string | null;
}

/**
 * Convert a 1-based (line, column) Python position into a 0-based character
 * offset in `doc`, clamped to `[0, doc.length]`. A missing/zero column anchors
 * at the start of the line.
 */
export function pythonPosToOffset(doc: string, line: number, column: number): number {
  const lines = doc.split("\n");
  const targetLine = Math.max(1, Math.min(Math.floor(line) || 1, lines.length));
  let offset = 0;
  for (let i = 0; i < targetLine - 1; i += 1) {
    offset += lines[i].length + 1; // +1 for the consumed "\n"
  }
  // Python `offset` is 1-based and points at the column AFTER the problem on
  // some versions; we treat it as 1-based column → 0-based by subtracting 1 and
  // clamp into the line.
  const col = Math.max(0, (Math.floor(column) || 1) - 1);
  const lineLen = lines[targetLine - 1]?.length ?? 0;
  offset += Math.min(col, lineLen);
  return Math.max(0, Math.min(offset, doc.length));
}

/**
 * Map a Python {@link PythonSyntaxErrorLike} to a single CodeMirror diagnostic
 * anchored at the reported line/column. Pure and total: an error with no usable
 * position anchors at the document start. Returns `[]` only for a null/empty
 * error so a clean compile reports nothing.
 */
export function pythonSyntaxToCm(
  err: PythonSyntaxErrorLike | null | undefined,
  doc: string,
): Diagnostic[] {
  if (!err) return [];
  const docLen = doc.length;
  const line = typeof err.lineno === "number" && err.lineno > 0 ? err.lineno : 1;
  const column = typeof err.offset === "number" && err.offset > 0 ? err.offset : 1;
  const from = pythonPosToOffset(doc, line, column);
  // Prefer a real span when the runtime gives an end position (3.10+); else
  // highlight from the caret to end-of-line (at least one char) so the squiggle
  // is visible.
  let to: number;
  if (typeof err.endLineno === "number" && typeof err.endOffset === "number") {
    to = pythonPosToOffset(doc, err.endLineno, err.endOffset);
  } else {
    // Extend to end of the offending line.
    const lineEnd = doc.indexOf("\n", from);
    to = lineEnd === -1 ? docLen : lineEnd;
  }
  if (to <= from) to = Math.min(docLen, from + 1);
  if (to <= from) to = from; // empty doc edge case
  const message =
    typeof err.msg === "string" && err.msg.length > 0
      ? `SyntaxError: ${err.msg}`
      : "SyntaxError";
  return [
    {
      from,
      to,
      severity: "error",
      message,
      source: "python",
    },
  ];
}

/**
 * Python snippet that parse-checks the user's source and reports the first
 * SyntaxError as JSON, or `{"ok": true}` when it compiles. Parse-only:
 * `compile(..., "exec")` never executes the code, so timers / loops / network
 * never run — this is purely a syntax gate. Designed to be sent to the Pyodide
 * worker (or run in-process) with the user source bound to `_mrnote_lint_src`.
 *
 * Exported so the worker round-trip and any in-process path share one snippet.
 */
export const PYTHON_LINT_SNIPPET = [
  "import json as _json",
  "def _mrnote_lint(_src):",
  "    try:",
  "        compile(_src, '<lint>', 'exec')",
  "        return _json.dumps({'ok': True})",
  "    except SyntaxError as _e:",
  "        return _json.dumps({",
  "            'ok': False,",
  "            'lineno': getattr(_e, 'lineno', None),",
  "            'offset': getattr(_e, 'offset', None),",
  "            'endLineno': getattr(_e, 'end_lineno', None),",
  "            'endOffset': getattr(_e, 'end_offset', None),",
  "            'msg': getattr(_e, 'msg', None) or str(_e),",
  "        })",
  "    except Exception as _e:",
  "        return _json.dumps({'ok': True})",
  "_mrnote_lint(_mrnote_lint_src)",
].join("\n");

/**
 * Parse the JSON string returned by {@link PYTHON_LINT_SNIPPET} into CM
 * diagnostics. Total + non-throwing: malformed JSON or `{ok:true}` → `[]`.
 */
export function parsePythonLintResult(raw: string, doc: string): Diagnostic[] {
  try {
    const parsed = JSON.parse(raw) as PythonSyntaxErrorLike & { ok?: boolean };
    if (parsed.ok) return [];
    return pythonSyntaxToCm(parsed, doc);
  } catch {
    return [];
  }
}

// ---------------------------------------------------------------------------
// Python semantic lint diagnostics (pyflakes).
// ---------------------------------------------------------------------------

/**
 * One pyflakes warning as reported by the worker's `lint` op. `row` and `col`
 * are 1-based. `message` is the human-readable warning text; `type` is the
 * pyflakes message class name (e.g. "UnusedImport", "UndefinedName").
 */
export interface PyflakesWarningLike {
  row?: number | null;
  col?: number | null;
  message?: string | null;
  type?: string | null;
}

/**
 * Map a list of pyflakes warnings to CodeMirror {@link Diagnostic}s with
 * `"warning"` severity. Pure and total: a warning with a missing/invalid position
 * anchors at the document start; an empty or non-array input returns `[]`.
 *
 * Each diagnostic spans from the column position to the end of the symbol (best
 * effort: we try to extend to the next whitespace/punctuation to give the squiggle
 * visible width, falling back to a single character). The `source` is "pyflakes".
 */
export function pyflakesWarningsToCm(
  warnings: readonly PyflakesWarningLike[],
  doc: string,
): Diagnostic[] {
  if (!Array.isArray(warnings) || warnings.length === 0) return [];
  const docLen = doc.length;
  const out: Diagnostic[] = [];
  for (const w of warnings) {
    const line = typeof w.row === "number" && w.row > 0 ? w.row : 1;
    const col = typeof w.col === "number" && w.col > 0 ? w.col : 1;
    const from = pythonPosToOffset(doc, line, col);
    // Extend `to` to the end of the current token (next non-word char), giving
    // the squiggle visible width. Fall back to a single character.
    // All paths are clamped to [0, docLen] so diagnostics never exceed bounds.
    let to: number;
    if (from < docLen) {
      const rest = doc.slice(from);
      const nonWord = rest.search(/[\s,;()[\]{}.]/);
      if (nonWord > 0) {
        to = Math.min(docLen, from + nonWord);
      } else {
        to = Math.min(docLen, from + 1);
      }
    } else {
      // `from` is already at or past the end of the document.
      to = from;
    }
    if (to <= from && from < docLen) to = Math.min(docLen, from + 1);
    const message =
      typeof w.message === "string" && w.message.length > 0 ? w.message : "Warning";
    out.push({
      from,
      to,
      severity: "warning",
      message,
      source: "pyflakes",
    });
  }
  return out;
}

/**
 * Parse the combined lint+warnings payload from the worker's `lint` op into a
 * flat list of CM diagnostics: syntax errors (severity "error") followed by
 * pyflakes semantic warnings (severity "warning"). Total + non-throwing: any
 * malformed input degrades to `[]`.
 *
 * @param lintRaw  JSON string from the syntax-check pass (PYTHON_LINT_SNIPPET).
 * @param warnings Array of pyflakes warning objects from the worker.
 * @param doc      The source document (for offset mapping).
 */
export function parsePythonDiagnostics(
  lintRaw: string,
  warnings: readonly PyflakesWarningLike[],
  doc: string,
): Diagnostic[] {
  const syntaxDiags = parsePythonLintResult(lintRaw, doc);
  // Don't surface semantic warnings when there's a syntax error — the pyflakes
  // pass is skipped in the worker when ok===false anyway, but guard here too.
  if (syntaxDiags.length > 0) return syntaxDiags;
  return pyflakesWarningsToCm(warnings, doc);
}

// ---------------------------------------------------------------------------
// Language → async lint source selector (wired into codemirror-config).
// ---------------------------------------------------------------------------

/**
 * Languages that get a full TypeScript-compiler lint source. JSON keeps its
 * existing `JSON.parse` linter; everything else (sql/css/html/markdown/unknown)
 * has no LSP-grade linter here.
 */
const TS_LSP_LANGUAGES: ReadonlySet<SupportedLanguage> = new Set([
  "javascript",
  "jsx",
  "typescript",
  "tsx",
]);

/** Whether a language is linted by the TypeScript compiler path. */
export function isTsLspLanguage(language: string): boolean {
  return TS_LSP_LANGUAGES.has(normalizeLanguage(language));
}

/** Whether a language is linted by the Pyodide syntax path. */
export function isPythonLspLanguage(language: string): boolean {
  return normalizeLanguage(language) === "python";
}

/**
 * Lazily-loaded TypeScript module promise, memoised so the (large) compiler is
 * fetched at most once per session and only when a JS/TS doc is actually
 * linted. Kept out of the initial editor chunk via dynamic `import()`.
 */
let tsModulePromise: Promise<typeof TsModule> | null = null;

function loadTs(): Promise<typeof TsModule> {
  if (!tsModulePromise) {
    tsModulePromise = import("typescript")
      .then((m) => (m as { default?: typeof TsModule }).default ?? (m as typeof TsModule))
      .then(async (ts) => {
        await localizeTypeScript(ts);
        return ts;
      })
      .catch((err) => {
        tsModulePromise = null; // allow a later retry
        throw err;
      });
  }
  return tsModulePromise;
}

/**
 * Async CM lint source for JS/TS/JSX/TSX: lazy-loads the compiler then runs
 * {@link tsDiagnostics}. Non-throwing — a failed compiler load yields `[]` so
 * the editor never breaks. Returns a `Promise` because CM6's `LintSource`
 * accepts `Promise<readonly Diagnostic[]>`.
 */
export async function tsLintSource(
  source: string,
  language: string,
): Promise<Diagnostic[]> {
  if (source.trim().length === 0) return [];
  if (!isTsLspLanguage(language)) return [];
  try {
    const ts = await loadTs();
    return tsDiagnostics(ts, source, language);
  } catch {
    return [];
  }
}
