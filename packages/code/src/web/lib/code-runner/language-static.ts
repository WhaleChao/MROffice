/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Bundled CodeMirror language fallback.
 *
 * Production code must only reach this module through `import()`. It keeps the
 * former all-languages implementation available as a reliable rollback path
 * without pinning every parser table into the first CodeWindow chunk.
 */
import {
  autoCloseTags as jsxAutoCloseTags,
  javascript,
} from "@codemirror/lang-javascript";
import { python } from "@codemirror/lang-python";
import { cpp } from "@codemirror/lang-cpp";
import { go } from "@codemirror/lang-go";
import { java } from "@codemirror/lang-java";
import { rust } from "@codemirror/lang-rust";
import { yaml } from "@codemirror/lang-yaml";
import { css } from "@codemirror/lang-css";
import {
  autoCloseTags as htmlAutoCloseTags,
  html,
} from "@codemirror/lang-html";
import { json } from "@codemirror/lang-json";
import { SQLite, sql } from "@codemirror/lang-sql";
import { markdown } from "@codemirror/lang-markdown";
import { php } from "@codemirror/lang-php";
import { StreamLanguage } from "@codemirror/language";
import { csharp } from "@codemirror/legacy-modes/mode/clike";
import { lua } from "@codemirror/legacy-modes/mode/lua";
import { perl } from "@codemirror/legacy-modes/mode/perl";
import { r } from "@codemirror/legacy-modes/mode/r";
import { ruby } from "@codemirror/legacy-modes/mode/ruby";
import { shell } from "@codemirror/legacy-modes/mode/shell";
import { stex } from "@codemirror/legacy-modes/mode/stex";
import { swift } from "@codemirror/legacy-modes/mode/swift";
import { toml } from "@codemirror/legacy-modes/mode/toml";
import type { Extension } from "@codemirror/state";

import { normalizeLanguage } from "./language-metadata";

/** Return a language extension synchronously after this bundled module loads. */
export function languageExtension(language: string): Extension {
  switch (normalizeLanguage(language)) {
    case "javascript":
      return javascript();
    case "typescript":
      return javascript({ typescript: true });
    case "jsx":
      return [javascript({ jsx: true }), jsxAutoCloseTags];
    case "tsx":
      return [javascript({ jsx: true, typescript: true }), jsxAutoCloseTags];
    case "python":
      return python();
    case "css":
      return css();
    case "html":
      // XML shares this grammar: `<Item/>` closes the element (see language-loader).
      return [html({ selfClosingTags: true }), htmlAutoCloseTags];
    case "json":
      return json();
    case "sql":
      // SQL pages run on SQLite (sql.js); see language-loader.
      return sql({ dialect: SQLite });
    case "markdown":
      return markdown();
    case "php":
      return php();
    case "bash":
      return StreamLanguage.define(shell);
    case "go":
      return go();
    case "rust":
      return rust();
    case "java":
      return java();
    case "c":
    case "cpp":
      return cpp();
    case "csharp":
      return StreamLanguage.define(csharp);
    case "ruby":
      return StreamLanguage.define(ruby);
    case "swift":
      return StreamLanguage.define(swift);
    case "r":
      return StreamLanguage.define(r);
    case "perl":
      return StreamLanguage.define(perl);
    case "lua":
      return StreamLanguage.define(lua);
    case "latex":
      return StreamLanguage.define(stex);
    case "yaml":
      return yaml();
    case "toml":
      return StreamLanguage.define(toml);
    case "unknown":
    default:
      return [];
  }
}
