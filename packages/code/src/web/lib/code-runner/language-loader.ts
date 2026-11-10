/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Load one language's grammar on demand, instead of all of them on every page.
 *
 * Before this split, `codemirror-config.ts` reached every grammar through static
 * imports. That implementation now lives in `language-static.ts` solely as the
 * deferred rollback path. Measured on the installed packages: the twelve
 * `@codemirror/lang-*` grammars with their `@lezer` parser tables come to about
 * 1.77 MB of unminified JavaScript, and the eight `legacy-modes` stream parsers
 * add ~99 KB. Opening a JavaScript file should not pay for Python, C++, Rust,
 * PHP, Java, SQL, Markdown, YAML, HTML, CSS, JSON and eight stream modes.
 *
 * Each grammar here sits behind its own `import()`, so the bundler gives each
 * one its own chunk and a page fetches exactly the one it needs.
 *
 * A CACHE HIT RESOLVES SYNCHRONOUSLY, through {@link peekLanguageExtension}.
 * This is the part that decides whether the change is felt: within a session a
 * user opens several files of the same language, and only the first of them can
 * possibly show unhighlighted text. Callers are expected to check the cache
 * first and only fall back to the promise, so the gap exists once per language
 * per session rather than once per page.
 *
 * REFUSALS AND FAILURES BOTH RESOLVE TO A VALID EXTENSION. An unknown language
 * resolves to `[]`, which is what the synchronous path already returns for one,
 * and a chunk that fails to load resolves to `[]` as well rather than rejecting.
 * A page whose grammar could not be fetched is a page without highlighting; it
 * is not a page whose editor failed to mount. The failure is recorded so a
 * caller can say so rather than leaving the user to wonder — see
 * {@link languageLoadFailed}.
 *
 * This module owns loading only. What the editor does during the gap, and what
 * it does when the grammar arrives, is the caller's contract.
 */

import type { Extension } from "@codemirror/state";

import { ENABLE_LAZY_LANGUAGE_GRAMMARS } from "@mrnote/core/web/lib/code-runner/editor-flags.ts";
import {
  normalizeLanguage,
  type SupportedLanguage,
} from "./language-metadata";

/** Resolved grammars, by normalized language. Populated once per session. */
const _loaded = new Map<SupportedLanguage, Extension>();
/** In-flight loads, so two editors opening at once fetch one chunk. */
const _inFlight = new Map<SupportedLanguage, Promise<Extension>>();
/** Languages whose chunk failed, so a caller can report it instead of guessing. */
const _failed = new Set<SupportedLanguage>();

/**
 * Languages that share one grammar, folded onto a single cache key.
 *
 * `lang-cpp` parses C as the subset it is, so a page of each would otherwise
 * build the grammar twice and hold two copies of the same parser table.
 */
const CACHE_ALIASES: Partial<Record<SupportedLanguage, SupportedLanguage>> = {
  c: "cpp",
};

/** The cache identity of a language: itself, unless it shares another's grammar. */
function cacheKey(language: string): SupportedLanguage {
  const lang = normalizeLanguage(language);
  return CACHE_ALIASES[lang] ?? lang;
}

/**
 * The already-loaded grammar for `language`, or null if it has not been fetched
 * yet. Never triggers a load — callers use this to skip the gap entirely.
 */
export function peekLanguageExtension(language: string): Extension | null {
  return _loaded.get(cacheKey(language)) ?? null;
}

/** Whether this language's chunk has been tried and failed. */
export function languageLoadFailed(language: string): boolean {
  return _failed.has(cacheKey(language));
}

/**
 * Fetch (or reuse) the grammar for `language`.
 *
 * Always resolves — to the grammar, or to `[]` for an unknown language or a
 * chunk that could not be fetched.
 */
export function loadLanguageExtension(language: string): Promise<Extension> {
  const lang = cacheKey(language);
  const ready = _loaded.get(lang);
  if (ready) return Promise.resolve(ready);
  const pending = _inFlight.get(lang);
  if (pending) return pending;

  const load = importLanguage(lang)
    .then((extension) => {
      _loaded.set(lang, extension);
      _failed.delete(lang);
      return extension;
    })
    .catch(() => {
      // Offline, a stale deployment whose hashed chunk is gone, a proxy refusal.
      // Plain text is a degraded editor; a rejected promise here would be a
      // broken one, and this module is not the right place to decide that.
      _failed.add(lang);
      return [] as Extension;
    })
    .finally(() => {
      _inFlight.delete(lang);
    });

  _inFlight.set(lang, load);
  return load;
}

/**
 * Load the configured grammar strategy.
 *
 * The ON path fetches one language. The OFF path dynamically fetches the former
 * all-language implementation, preserving a reliable runtime rollback without
 * making that implementation part of CodeWindow's initial dependency graph.
 */
export function loadConfiguredLanguageExtension(
  language: string,
): Promise<Extension> {
  if (ENABLE_LAZY_LANGUAGE_GRAMMARS) {
    return loadLanguageExtension(language);
  }

  const lang = cacheKey(language);
  const ready = _loaded.get(lang);
  if (ready) return Promise.resolve(ready);
  const pending = _inFlight.get(lang);
  if (pending) return pending;

  const load = import("./language-static")
    .then(({ languageExtension }) => {
      const extension = languageExtension(lang);
      _loaded.set(lang, extension);
      _failed.delete(lang);
      return extension;
    })
    .catch(() => {
      _failed.add(lang);
      return [] as Extension;
    })
    .finally(() => {
      _inFlight.delete(lang);
    });
  _inFlight.set(lang, load);
  return load;
}

/**
 * One dynamic import per language, deliberately written out rather than built
 * from a template: a bundler can only split what it can see statically, and
 * `import(\`@codemirror/lang-${name}\`)` is not something it can see.
 */
async function importLanguage(lang: SupportedLanguage): Promise<Extension> {
  switch (lang) {
    case "javascript": {
      const { javascript } = await import("@codemirror/lang-javascript");
      return javascript();
    }
    case "typescript": {
      const { javascript } = await import("@codemirror/lang-javascript");
      return javascript({ typescript: true });
    }
    case "jsx": {
      // `jsxAutoCloseTags` closes the matching `</tag>` on `>`; a no-op outside JSX.
      const { javascript, autoCloseTags } = await import(
        "@codemirror/lang-javascript"
      );
      return [javascript({ jsx: true }), autoCloseTags];
    }
    case "tsx": {
      const { javascript, autoCloseTags } = await import(
        "@codemirror/lang-javascript"
      );
      return [javascript({ jsx: true, typescript: true }), autoCloseTags];
    }
    case "python": {
      const { python } = await import("@codemirror/lang-python");
      return python();
    }
    case "css": {
      const { css } = await import("@codemirror/lang-css");
      return css();
    }
    case "html": {
      const { html, autoCloseTags } = await import("@codemirror/lang-html");
      // XML pages share this grammar, and in XML `<Item/>` closes the element:
      // without it every self-closing element read as left open, so valid XML
      // showed "syntax is invalid" at the parent's end tag.
      return [html({ selfClosingTags: true }), autoCloseTags];
    }
    case "json": {
      const { json } = await import("@codemirror/lang-json");
      return json();
    }
    case "sql": {
      const { sql, SQLite } = await import("@codemirror/lang-sql");
      // SQL pages run on SQLite (sql.js), so the editor reads them as SQLite
      // does: `name` and "name" are identifiers, not syntax errors.
      return sql({ dialect: SQLite });
    }
    case "markdown": {
      const { markdown } = await import("@codemirror/lang-markdown");
      return markdown();
    }
    case "php": {
      const { php } = await import("@codemirror/lang-php");
      return php();
    }
    case "go": {
      const { go } = await import("@codemirror/lang-go");
      return go();
    }
    case "rust": {
      const { rust } = await import("@codemirror/lang-rust");
      return rust();
    }
    case "java": {
      const { java } = await import("@codemirror/lang-java");
      return java();
    }
    // One grammar covers both: lang-cpp parses C as the subset it is.
    case "c":
    case "cpp": {
      const { cpp } = await import("@codemirror/lang-cpp");
      return cpp();
    }
    case "yaml": {
      const { yaml } = await import("@codemirror/lang-yaml");
      return yaml();
    }
    case "bash":
      return streamMode("shell");
    case "csharp":
      return streamMode("csharp");
    case "ruby":
      return streamMode("ruby");
    case "swift":
      return streamMode("swift");
    case "r":
      return streamMode("r");
    case "perl":
      return streamMode("perl");
    case "lua":
      return streamMode("lua");
    case "latex":
      return streamMode("stex");
    case "toml":
      return streamMode("toml");
    case "unknown":
      return [];
    default: {
      // Exhaustiveness: a language added to the catalog without a case here
      // reaches this line, and gets the same plain text an unknown one does
      // rather than a crash.
      const _never: never = lang;
      void _never;
      return [];
    }
  }
}

/** The `@codemirror/legacy-modes` stream parsers, which have no Lezer tree. */
type StreamModeName =
  | "shell"
  | "csharp"
  | "ruby"
  | "swift"
  | "r"
  | "perl"
  | "lua"
  | "stex"
  | "toml";

async function streamMode(name: StreamModeName): Promise<Extension> {
  const { StreamLanguage } = await import("@codemirror/language");
  switch (name) {
    case "shell": {
      const { shell } = await import("@codemirror/legacy-modes/mode/shell");
      return StreamLanguage.define(shell);
    }
    case "csharp": {
      const { csharp } = await import("@codemirror/legacy-modes/mode/clike");
      return StreamLanguage.define(csharp);
    }
    case "ruby": {
      const { ruby } = await import("@codemirror/legacy-modes/mode/ruby");
      return StreamLanguage.define(ruby);
    }
    case "swift": {
      const { swift } = await import("@codemirror/legacy-modes/mode/swift");
      return StreamLanguage.define(swift);
    }
    case "r": {
      const { r } = await import("@codemirror/legacy-modes/mode/r");
      return StreamLanguage.define(r);
    }
    case "perl": {
      const { perl } = await import("@codemirror/legacy-modes/mode/perl");
      return StreamLanguage.define(perl);
    }
    case "lua": {
      const { lua } = await import("@codemirror/legacy-modes/mode/lua");
      return StreamLanguage.define(lua);
    }
    case "stex": {
      // LaTeX source; the mode's own language data supplies `%` comments.
      const { stex } = await import("@codemirror/legacy-modes/mode/stex");
      return StreamLanguage.define(stex);
    }
    case "toml": {
      const { toml } = await import("@codemirror/legacy-modes/mode/toml");
      return StreamLanguage.define(toml);
    }
  }
}

/**
 * Forget everything loaded. Tests only — a real session never unloads a
 * grammar, and dropping one mid-session would leave live editors configured
 * with an extension this module no longer admits to having.
 * @internal
 */
export function _resetLanguageLoaderForTest(): void {
  _loaded.clear();
  _inFlight.clear();
  _failed.clear();
}
