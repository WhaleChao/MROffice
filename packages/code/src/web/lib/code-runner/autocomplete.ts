/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Pure autocomplete suggestions for the code editor: language keywords plus
 * identifiers already present in the document, filtered by the word at the
 * caret. Framework-free + unit-tested; CodeEditor renders the popup.
 *
 * Built-in code SNIPPETS (see {@link ./code-snippets}) are surfaced here too, as
 * an ADDITIVE, separate field on {@link CompletionContext} (`snippets`): the
 * `items: string[]` identifier/keyword list is unchanged, so every existing
 * caller and test keeps working byte-for-byte. The popup shows snippets
 * alongside the identifier completions; accepting one inserts its body via
 * {@link applySnippet} instead of {@link applyCompletion}.
 */
import {
  expandSnippet,
  lineIndentAt,
  matchSnippets,
  type Snippet,
} from "./code-snippets";

const KEYWORDS: Record<string, string[]> = {
  javascript: [
    "const", "let", "var", "function", "return", "if", "else", "for", "while",
    "do", "switch", "case", "break", "continue", "class", "extends", "new",
    "this", "super", "import", "export", "default", "from", "async", "await",
    "try", "catch", "finally", "throw", "typeof", "instanceof", "yield", "null",
    "undefined", "true", "false", "console", "Promise", "Object", "Array",
    "Math", "JSON", "Map", "Set", "Number", "String", "Boolean",
  ],
  python: [
    "def", "return", "if", "elif", "else", "for", "while", "break", "continue",
    "class", "import", "from", "as", "with", "try", "except", "finally", "raise",
    "lambda", "yield", "global", "nonlocal", "pass", "assert", "del", "in", "is",
    "not", "and", "or", "None", "True", "False", "print", "len", "range",
    "list", "dict", "set", "tuple", "int", "str", "float", "bool", "enumerate",
    "zip", "map", "filter", "open", "self",
  ],
  sql: [
    "SELECT", "FROM", "WHERE", "INSERT", "INTO", "VALUES", "UPDATE", "SET",
    "DELETE", "CREATE", "TABLE", "DROP", "ALTER", "JOIN", "LEFT", "RIGHT",
    "INNER", "OUTER", "ON", "GROUP", "BY", "ORDER", "HAVING", "LIMIT", "DISTINCT",
    "AS", "AND", "OR", "NOT", "NULL", "PRIMARY", "KEY", "FOREIGN", "INDEX",
    "COUNT", "SUM", "AVG", "MIN", "MAX", "BETWEEN", "LIKE", "IN", "UNION",
  ],
  go: [
    "package", "import", "func", "return", "if", "else", "for", "range",
    "switch", "case", "default", "break", "continue", "type", "struct",
    "interface", "map", "chan", "go", "select", "defer", "var", "const",
    "nil", "true", "false", "make", "new", "len", "cap", "append", "copy",
    "delete", "panic", "recover", "error", "string", "int", "int64", "float64",
    "bool", "byte", "rune", "fmt", "Println", "Printf", "Sprintf",
  ],
  rust: [
    "fn", "let", "mut", "const", "static", "if", "else", "match", "loop",
    "while", "for", "in", "break", "continue", "return", "struct", "enum",
    "impl", "trait", "type", "where", "pub", "use", "mod", "crate", "self",
    "Self", "super", "as", "ref", "move", "dyn", "async", "await", "unsafe",
    "true", "false", "Some", "None", "Ok", "Err", "Option", "Result", "String",
    "Vec", "Box", "println!", "format!", "vec!", "i32", "u32", "i64", "u64",
    "f64", "usize", "bool", "str",
  ],
  java: [
    "public", "private", "protected", "static", "final", "abstract", "class",
    "interface", "enum", "extends", "implements", "import", "package", "new",
    "this", "super", "return", "if", "else", "for", "while", "do", "switch",
    "case", "default", "break", "continue", "try", "catch", "finally", "throw",
    "throws", "void", "int", "long", "double", "float", "boolean", "char",
    "byte", "short", "String", "Object", "List", "Map", "Set", "ArrayList",
    "HashMap", "System", "println", "true", "false", "null", "instanceof",
    "synchronized", "volatile", "record", "var",
  ],
  c: [
    "int", "long", "short", "char", "float", "double", "unsigned", "signed",
    "void", "struct", "union", "enum", "typedef", "const", "static", "extern",
    "register", "volatile", "inline", "sizeof", "return", "if", "else", "for",
    "while", "do", "switch", "case", "default", "break", "continue", "goto",
    "NULL", "size_t", "printf", "scanf", "malloc", "calloc", "realloc", "free",
    "memcpy", "memset", "strlen", "strcmp", "FILE", "fopen", "fclose",
  ],
  cpp: [
    "int", "long", "char", "float", "double", "bool", "void", "auto", "const",
    "constexpr", "static", "extern", "inline", "struct", "class", "union",
    "enum", "typedef", "using", "namespace", "template", "typename", "public",
    "private", "protected", "virtual", "override", "final", "new", "delete",
    "this", "nullptr", "true", "false", "return", "if", "else", "for", "while",
    "do", "switch", "case", "default", "break", "continue", "try", "catch",
    "throw", "std", "string", "vector", "map", "set", "cout", "cin", "endl",
    "unique_ptr", "shared_ptr", "size_t", "sizeof", "operator", "noexcept",
  ],
  csharp: [
    "public", "private", "protected", "internal", "static", "readonly", "const",
    "class", "struct", "interface", "enum", "record", "namespace", "using",
    "new", "this", "base", "return", "if", "else", "for", "foreach", "while",
    "do", "switch", "case", "default", "break", "continue", "try", "catch",
    "finally", "throw", "void", "int", "long", "double", "float", "decimal",
    "bool", "char", "string", "object", "var", "dynamic", "async", "await",
    "true", "false", "null", "is", "as", "in", "out", "ref", "get", "set",
    "List", "Dictionary", "Console", "WriteLine", "Task",
  ],
  ruby: [
    "def", "end", "class", "module", "if", "elsif", "else", "unless", "case",
    "when", "while", "until", "for", "in", "do", "begin", "rescue", "ensure",
    "raise", "return", "yield", "lambda", "proc", "require",
    "require_relative", "include", "extend", "attr_accessor", "attr_reader",
    "attr_writer", "initialize", "self", "super", "nil", "true", "false",
    "puts", "print", "each", "map", "select", "reject", "reduce", "new",
  ],
  php: [
    "function", "return", "if", "else", "elseif", "for", "foreach", "while",
    "do", "switch", "case", "default", "break", "continue", "class", "interface",
    "trait", "extends", "implements", "new", "public", "private", "protected",
    "static", "const", "abstract", "final", "namespace", "use", "require",
    "include", "echo", "print", "isset", "empty", "unset", "array", "true",
    "false", "null", "this", "self", "parent", "try", "catch", "finally",
    "throw", "match", "fn", "strlen", "count", "implode", "explode", "var_dump",
  ],
  swift: [
    "func", "let", "var", "if", "else", "guard", "switch", "case", "default",
    "for", "in", "while", "repeat", "break", "continue", "return", "class",
    "struct", "enum", "protocol", "extension", "init", "deinit", "self", "super",
    "import", "public", "private", "internal", "fileprivate", "open", "static",
    "final", "lazy", "weak", "unowned", "mutating", "throws", "throw", "try",
    "catch", "defer", "async", "await", "some", "any", "nil", "true", "false",
    "String", "Int", "Double", "Bool", "Array", "Dictionary", "Optional",
    "print", "map", "filter", "reduce",
  ],
  r: [
    "function", "if", "else", "for", "in", "while", "repeat", "break", "next",
    "return", "library", "require", "TRUE", "FALSE", "NULL", "NA", "NaN", "Inf",
    "c", "vector", "list", "matrix", "data.frame", "factor", "length", "names",
    "print", "cat", "paste", "paste0", "sprintf", "apply", "sapply", "lapply",
    "mean", "median", "sum", "min", "max", "sd", "var", "summary", "str",
    "head", "tail", "read.csv", "write.csv", "plot",
  ],
  perl: [
    "my", "our", "local", "sub", "return", "if", "elsif", "else", "unless",
    "for", "foreach", "while", "until", "do", "last", "next", "redo", "use",
    "require", "package", "bless", "ref", "defined", "undef", "exists",
    "delete", "scalar", "wantarray", "print", "printf", "say", "chomp", "chop",
    "split", "join", "map", "grep", "sort", "reverse", "keys", "values", "each",
    "push", "pop", "shift", "unshift", "die", "warn", "eval", "qw",
  ],
  lua: [
    "function", "local", "return", "if", "then", "elseif", "else", "end",
    "for", "in", "while", "repeat", "until", "do", "break", "goto", "and",
    "or", "not", "nil", "true", "false", "self", "pairs", "ipairs", "next",
    "type", "tostring", "tonumber", "print", "require", "pcall", "xpcall",
    "error", "assert", "setmetatable", "getmetatable", "rawget", "rawset",
    "table", "insert", "remove", "concat", "string", "format", "math",
  ],
  bash: [
    "if", "then", "elif", "else", "fi", "for", "in", "do", "done", "while",
    "until", "case", "esac", "function", "return", "break", "continue", "local",
    "declare", "export", "readonly", "unset", "shift", "exit", "trap", "source",
    "echo", "printf", "read", "cd", "pwd", "grep", "sed", "awk", "cut",
    "sort", "uniq", "head", "tail", "cat", "find", "xargs", "test", "true",
    "false", "eval", "exec", "set",
  ],
};

// Unicode-aware identifier grammar (ES2018 property escapes): a legal JS/TS (or
// Python) identifier can contain non-ASCII letters — café, naïve, 中文Var, etc.
// ASCII-only `\w` would split those and break completion/scraping on them.
const IDENTIFIER = /[\p{ID_Start}_$][\p{ID_Continue}$]*/gu;
const WORD_CHAR = /[\p{ID_Continue}$\u200c\u200d]/u;

function previousCodePointStart(value: string, index: number): number {
  const previous = index - 1;
  if (previous <= 0) return Math.max(0, previous);
  const trailing = value.charCodeAt(previous);
  const leading = value.charCodeAt(previous - 1);
  return trailing >= 0xdc00 &&
    trailing <= 0xdfff &&
    leading >= 0xd800 &&
    leading <= 0xdbff
    ? previous - 1
    : previous;
}

export interface CompletionContext {
  /** The word being typed (text from the word start up to the caret). */
  prefix: string;
  /** Index where the word starts (where an accepted item replaces from). */
  start: number;
  /** Ranked, de-duplicated identifier/keyword suggestions (unchanged contract). */
  items: string[];
  /**
   * Built-in code snippets whose prefix matches `prefix`, ranked. ADDITIVE: a
   * separate list from `items` so existing callers that only read `items` are
   * unaffected. Empty array when no snippet matches.
   */
  snippets: Snippet[];
}

function keywordsFor(language?: string): string[] {
  const lang = (language || "").toLowerCase();
  switch (lang) {
    case "typescript":
    case "jsx":
    case "tsx":
      return KEYWORDS.javascript;
    case "sh":
    case "zsh":
    case "shell":
      return KEYWORDS.bash;
    case "c++":
      return KEYWORDS.cpp;
    case "c#":
    case "cs":
      return KEYWORDS.csharp;
    case "golang":
      return KEYWORDS.go;
    default:
      return KEYWORDS[lang] ?? KEYWORDS.javascript;
  }
}

/**
 * Compute completions for the caret position. Returns null when there is no
 * word to complete (so the caller closes the popup).
 *
 * `extraIdentifiers` contributes identifiers from OUTSIDE this buffer — e.g. the
 * other cells of a shared-scope JS/TS multi-buffer host, where a symbol declared in one
 * cell is in scope in the next. They join the same pool as the document's own
 * identifiers, so the existing prefix filter / ranking / dedup apply uniformly.
 */
export function getCompletions(
  value: string,
  caret: number,
  language?: string,
  limit = 12,
  extraIdentifiers?: Iterable<string>,
): CompletionContext | null {
  // Walk back over identifier characters to find the current word.
  const safeCaret = Number.isFinite(caret)
    ? Math.max(0, Math.min(Math.trunc(caret), value.length))
    : 0;
  let start = safeCaret;
  while (start > 0) {
    const previous = previousCodePointStart(value, start);
    if (!WORD_CHAR.test(value.slice(previous, start))) break;
    start = previous;
  }
  const prefix = value.slice(start, safeCaret);
  if (prefix.length === 0 || /^\d/.test(prefix)) return null; // need a real word

  const lower = prefix.toLowerCase();
  const pool = new Set<string>(keywordsFor(language));
  // Identifiers elsewhere in the document.
  for (const match of value.matchAll(IDENTIFIER)) {
    const token = match[0];
    // Skip the token currently being typed at the caret.
    if (match.index === start) continue;
    pool.add(token);
  }
  // Identifiers contributed from outside this buffer (e.g. sibling multi-buffer host
  // cells). The pool is a Set, so any overlap with document identifiers dedupes.
  if (extraIdentifiers) {
    for (const token of extraIdentifiers) {
      if (token) pool.add(token);
    }
  }

  const items = Array.from(pool)
    .filter((w) => w.length > prefix.length && w.toLowerCase().startsWith(lower))
    .sort((a, b) => {
      // Case-sensitive prefix matches first, then shorter, then alpha.
      const aExact = a.startsWith(prefix) ? 0 : 1;
      const bExact = b.startsWith(prefix) ? 0 : 1;
      if (aExact !== bExact) return aExact - bExact;
      if (a.length !== b.length) return a.length - b.length;
      return a.localeCompare(b);
    })
    .slice(0, limit);

  // Built-in snippets matching the same prefix (cheap O(#snippets) scan). These
  // are ADDITIVE: when only snippets match (e.g. "trycatch"), the popup still
  // opens so the user can insert one — but `items` keeps its exact prior value.
  const snippets = matchSnippets(prefix, language);

  if (items.length === 0 && snippets.length === 0) return null;
  return { prefix, start, items, snippets };
}

/** Apply an accepted completion, returning the new value + caret. */
export function applyCompletion(
  value: string,
  start: number,
  caret: number,
  item: string,
): { value: string; caret: number } {
  return {
    value: value.slice(0, start) + item + value.slice(caret),
    caret: start + item.length,
  };
}

/**
 * Apply an accepted SNIPPET: replace the typed prefix `[start, caret)` with the
 * snippet body, and return the new value plus the caret at the snippet's `$0`
 * marker (or the end of the body when it has none). Continuation lines of a
 * multi-line body are re-indented to the prefix's column so the snippet stays
 * aligned. Mirrors {@link applyCompletion}'s replacement window, so the editors
 * can reuse the same accept plumbing.
 */
export function applySnippet(
  value: string,
  start: number,
  caret: number,
  snippet: Snippet,
): { value: string; caret: number } {
  const indent = lineIndentAt(value, start);
  const { text, cursorOffset } = expandSnippet(snippet.body, indent);
  return {
    value: value.slice(0, start) + text + value.slice(caret),
    caret: start + cursorOffset,
  };
}
