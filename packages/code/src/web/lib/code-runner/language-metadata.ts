/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Grammar-free language metadata shared by the editor and lazy loader.
 *
 * Keep this module free of `@codemirror/lang-*` and legacy-mode imports. It is
 * safe to load before a language chunk and supplies only the small pieces that
 * must not disappear during that gap: canonical ids, comment tokens, and a
 * conservative indentation fallback.
 */

import { EditorState, type Extension } from "@codemirror/state";
import {
  getIndentUnit,
  indentService,
  type IndentContext,
} from "@codemirror/language";

export type SupportedLanguage =
  | "javascript"
  | "typescript"
  | "jsx"
  | "tsx"
  | "python"
  | "css"
  | "html"
  | "json"
  | "sql"
  | "markdown"
  | "bash"
  | "go"
  | "rust"
  | "java"
  | "c"
  | "cpp"
  | "csharp"
  | "ruby"
  | "php"
  | "swift"
  | "r"
  | "perl"
  | "lua"
  | "latex"
  | "yaml"
  | "toml"
  | "unknown";

/** Normalize catalog ids and common aliases without importing a parser. */
export function normalizeLanguage(
  language: string | undefined | null,
): SupportedLanguage {
  const key = (language ?? "").trim().toLowerCase();
  switch (key) {
    case "javascript":
    case "js":
    case "node":
      return "javascript";
    case "typescript":
    case "ts":
      return "typescript";
    case "jsx":
      return "jsx";
    case "tsx":
      return "tsx";
    case "python":
    case "py":
      return "python";
    case "css":
      return "css";
    case "html":
    case "htm":
    case "xml":
    case "xhtml":
      return "html";
    case "json":
      return "json";
    case "sql":
      return "sql";
    case "markdown":
    case "md":
      return "markdown";
    case "bash":
    case "sh":
    case "zsh":
    case "shell":
      return "bash";
    case "go":
    case "golang":
      return "go";
    case "rust":
    case "rs":
      return "rust";
    case "java":
      return "java";
    case "c":
      return "c";
    case "cpp":
    case "c++":
      return "cpp";
    case "csharp":
    case "cs":
    case "c#":
      return "csharp";
    case "ruby":
    case "rb":
      return "ruby";
    case "php":
      return "php";
    case "swift":
      return "swift";
    case "r":
      return "r";
    case "perl":
    case "pl":
      return "perl";
    case "lua":
      return "lua";
    case "latex":
    case "tex":
      return "latex";
    case "yaml":
    case "yml":
      return "yaml";
    case "ini":
    case "toml":
      return "toml";
    default:
      return "unknown";
  }
}

interface CommentTokens {
  line?: string;
  block?: { open: string; close: string };
}

const C_STYLE_COMMENTS: CommentTokens = {
  line: "//",
  block: { open: "/*", close: "*/" },
};
const HASH_COMMENTS: CommentTokens = { line: "#" };
const COMMENT_TOKENS: Partial<Record<SupportedLanguage, CommentTokens>> = {
  javascript: C_STYLE_COMMENTS,
  typescript: C_STYLE_COMMENTS,
  jsx: C_STYLE_COMMENTS,
  tsx: C_STYLE_COMMENTS,
  python: HASH_COMMENTS,
  css: { block: { open: "/*", close: "*/" } },
  html: { block: { open: "<!--", close: "-->" } },
  sql: { line: "--", block: { open: "/*", close: "*/" } },
  markdown: { block: { open: "<!--", close: "-->" } },
  bash: HASH_COMMENTS,
  go: C_STYLE_COMMENTS,
  rust: C_STYLE_COMMENTS,
  java: C_STYLE_COMMENTS,
  c: C_STYLE_COMMENTS,
  cpp: C_STYLE_COMMENTS,
  csharp: C_STYLE_COMMENTS,
  ruby: HASH_COMMENTS,
  php: C_STYLE_COMMENTS,
  swift: C_STYLE_COMMENTS,
  r: HASH_COMMENTS,
  perl: HASH_COMMENTS,
  lua: { line: "--", block: { open: "--[[", close: "]]--" } },
  latex: { line: "%" },
  yaml: HASH_COMMENTS,
  toml: HASH_COMMENTS,
};

/**
 * Whether Toggle Comment has a real language-data receipt for this buffer.
 *
 * CodeMirror's command returns `false` when no `commentTokens` facet exists.
 * Hosts use the same pure lookup to disable their ribbon/context affordances
 * before dispatch, while the command itself still rechecks the live CM state.
 */
export function languageSupportsComments(language: string): boolean {
  return COMMENT_TOKENS[normalizeLanguage(language)] !== undefined;
}

const BRACE_INDENT_LANGUAGES: ReadonlySet<SupportedLanguage> = new Set([
  "javascript",
  "typescript",
  "jsx",
  "tsx",
  "css",
  "json",
  "go",
  "rust",
  "java",
  "c",
  "cpp",
  "csharp",
  "php",
  "swift",
]);

function gapIndentation(
  context: IndentContext,
  pos: number,
  language: SupportedLanguage,
): number | null {
  const line = context.lineAt(pos, -1);
  const before = line.text.slice(0, Math.max(0, pos - line.from)).trimEnd();
  const base = context.lineIndent(pos, -1);
  const opensBlock =
    (language === "python" && before.endsWith(":")) ||
    (language === "yaml" && (before.endsWith(":") || before.endsWith("-"))) ||
    (BRACE_INDENT_LANGUAGES.has(language) && /[({[]$/.test(before));
  return opensBlock ? base + getIndentUnit(context.state) : null;
}

/**
 * Metadata used only while the real grammar is absent.
 *
 * Returning `null` from the fallback indent service tells CodeMirror to inherit
 * the current line's indentation. The few unambiguous block openers add one
 * configured indent unit. Once the parser arrives the whole extension is
 * replaced, so normal grammar indentation and nested-language metadata win.
 */
export function grammarGapLanguageExtension(language: string): Extension {
  const normalized = normalizeLanguage(language);
  const commentTokens = COMMENT_TOKENS[normalized];
  return [
    commentTokens ? EditorState.languageData.of(() => [{ commentTokens }]) : [],
    indentService.of((context, pos) =>
      gapIndentation(context, pos, normalized),
    ),
  ];
}
