/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Lightweight document-outline (Go-to-Symbol / Ctrl+Shift+O) extraction for the
 * NON-TS/JS languages, where there is no TypeScript LanguageService to drive
 * {@link import("./code-intelligence").documentSymbols}.
 *
 * For TS/JS the editor uses the real navigation tree; this module is the
 * fallback for python / go / rust / ruby / sql / markdown / … It scans the
 * source line-by-line with a per-language set of header patterns (def / class /
 * function / heading lines) and returns the matched lines as
 * {@link OutlineSymbol}s — name, kind, and the buffer offset of the symbol NAME
 * (so a jump lands on the identifier, mirroring how the TS path selects the
 * symbol span). Indentation is folded into a `depth` so nested defs nest in the
 * picker.
 *
 * Deliberately simple and TOTAL: it never parses, never throws, and a language
 * with no sensible structure (json / yaml / css / plain text) yields `[]` so the
 * picker shows its empty state. The result is in document order (the scan order),
 * which is already ascending by offset.
 */
import { normalizeLanguage } from "./language-metadata";

/**
 * A single symbol detected by the regex fallback. `from`/`to` bracket the symbol
 * NAME within the buffer (a jump selects exactly the identifier); `depth` is the
 * indentation level (0 = outermost) so members nest under their container.
 */
export interface OutlineSymbol {
  name: string;
  kind: string;
  from: number;
  to: number;
  depth: number;
}

/**
 * A header matcher: a regex whose capture group #`nameGroup` (default 1) is the
 * symbol name. A few languages need a different group (e.g. markdown captures the
 * `#` run in group 1 and the heading text in group 2).
 */
interface HeaderPattern {
  re: RegExp;
  kind: string;
  /** 1-based capture-group index of the symbol name (defaults to 1). */
  nameGroup?: number;
  /**
   * The name must be a class or struct the outline has already listed (a C++
   * constructor written inside its class, `Box() {}`, which otherwise reads
   * exactly like a macro call such as `TEST(Suite, Name) {`).
   */
  namesKnownType?: boolean;
}

const IDENTIFIER_START_SOURCE = "[$_\\p{ID_Start}]";
const IDENTIFIER_CONTINUE_SOURCE = "[$\\u200c\\u200d\\p{ID_Continue}]*";
const IDENTIFIER_SOURCE = `${IDENTIFIER_START_SOURCE}${IDENTIFIER_CONTINUE_SOURCE}`;

function unicodePattern(source: string): RegExp {
  return new RegExp(source, "u");
}

/**
 * Per-language header patterns. Each regex must anchor at the start of the
 * (whitespace-trimmed) line and capture the symbol name in group 1. Order
 * matters only when two could match the same line — the FIRST match wins.
 */
/**
 * Words that begin a control statement, so `if (ready) {` is never mistaken for
 * a function definition. Line-anchored matching plus a required opening brace
 * already exclude calls; this excludes the statements that look like calls.
 */
const NOT_A_DEFINITION =
  /^(?:(?:if|for|foreach|while|switch|catch|return|else|do|sizeof|synchronized|using|lock|fixed|checked|unchecked|try|throw|await|yield|case)\b|new\s+[\w$.<>]+\s*\()/;

/**
 * What may follow a definition's parameter list on its line: nothing (the
 * brace on the next line), an opening brace, or the whole body when it closes
 * on the same line (`int main() { return 0; }`), as VS Code's outline lists it.
 */
const DEFINITION_TAIL_SOURCE = "(?:\\{.*\\}|\\{)?\\s*$";

/** A parameter list, allowing one level of nested parentheses (`int (*cmp)(int)`, `@Param("id") long id`). */
const PARAMETER_LIST_SOURCE = "\\((?:[^()]|\\([^()]*\\))*\\)";

/**
 * A C/C++ function DEFINITION on one line: a return type, the name, a parameter
 * list, and an opening brace. Declarations (ending in `;`) are deliberately
 * skipped — an outline of a header's prototypes is noise.
 */
// A C++ member defined outside its class is named `A::f` (or `A::~A`), as
// VS Code's outline shows it; plain C names never carry the qualifier.
const C_QUALIFIED_NAME_SOURCE = `(?:${IDENTIFIER_SOURCE}::)*~?${IDENTIFIER_SOURCE}`;

// Member qualifiers and a trailing return type may sit between the parameter
// list and the body: `int area() const noexcept`, `auto f(int x) -> int`.
const C_DEFINITION_QUALIFIERS_SOURCE =
  "(?:(?:const|volatile|noexcept|override|final|&&|&)\\s*)*(?:->\\s*[^{;]+?\\s*)?";

const C_FUNCTION_DEFINITION = unicodePattern(
  `^(?!.*;\\s*$)[A-Za-z_][\\w\\s*&:<>,]*?[\\s*&](${C_QUALIFIED_NAME_SOURCE})\\s*${PARAMETER_LIST_SOURCE}\\s*${C_DEFINITION_QUALIFIERS_SOURCE}${DEFINITION_TAIL_SOURCE}`,
);

/** An out-of-class C++ constructor or destructor: `A::A(int x) : x_(x) {`. */
const CPP_QUALIFIED_CONSTRUCTOR = unicodePattern(
  `^(?!.*;\\s*$)((?:${IDENTIFIER_SOURCE}::)+~?${IDENTIFIER_SOURCE})\\s*${PARAMETER_LIST_SOURCE}\\s*(?:noexcept\\s*)?(?::[^{;]*?)?${DEFINITION_TAIL_SOURCE}`,
);

// Java annotations (`@Override`, `@GetMapping("/x")`) and C# attributes
// (`[HttpGet]`) written on the member's own line.
const JVM_ANNOTATIONS_SOURCE = `(?:@${IDENTIFIER_SOURCE}(?:\\.${IDENTIFIER_SOURCE})*(?:\\([^)]*\\))?\\s+|\\[[^\\]]*\\]\\s*)*`;

/**
 * A C++ constructor or destructor written inside its class: `Box() {}`,
 * `explicit Box(int w) : w_(w) {`, `~Box() {`. Only accepted for a name the
 * outline already knows as a class or struct (see `namesKnownType`).
 */
const CPP_MEMBER_CONSTRUCTOR = unicodePattern(
  `^(?!.*;\\s*$)(?:(?:explicit|constexpr|inline|virtual)\\s+)*(~?${IDENTIFIER_SOURCE})\\s*${PARAMETER_LIST_SOURCE}\\s*(?:noexcept\\s*)?(?::[^{;]*?)?${DEFINITION_TAIL_SOURCE}`,
);

// What follows a type's name where the type is DEFINED (`struct point {`,
// `class Box : public Shape`, `class Box final`, a forward `class Box;`, or the
// brace on the next line), as opposed to a variable or a return type that
// merely names it (`struct point p = {1, 2};`, `struct point *make(void) {`).
const C_TYPE_HEAD_END_SOURCE = "(?=\\s*(?:[:{;]|final\\b|$))";

/** `typedef struct point {`: the tag is the type being defined. */
const C_TYPEDEF_TAG = unicodePattern(
  `^typedef\\s+(?:struct|union|enum)\\s+(${IDENTIFIER_SOURCE})\\s*\\{`,
);

/** `} Point;` closing a `typedef struct {`: the typedef's name. */
const C_TYPEDEF_CLOSE = unicodePattern(`^\\}\\s*(${IDENTIFIER_SOURCE})\\s*;`);

const JVM_MODIFIER_WORDS =
  "public|private|protected|internal|static|final|abstract|synchronized|override|virtual|async|sealed|new|unsafe|extern|default|readonly";
const JVM_MEMBER_MODIFIERS_SOURCE = `(?:(?:${JVM_MODIFIER_WORDS})\\s+)*`;
// A return type is never a modifier, so `public Box(int w)` stays a constructor.
const JVM_NOT_A_MODIFIER_SOURCE = `(?!(?:${JVM_MODIFIER_WORDS})\\s)`;

// A Java generic method's type parameters precede its return type
// (`public static <T> List<T> of(T x)`); a C# one's follow its name
// (`public T Get<T>(int id) where T : class`).
const JVM_LEADING_TYPE_PARAMETERS_SOURCE = "(?:<[^()]*?>\\s+)?";
const JVM_TRAILING_TYPE_PARAMETERS_SOURCE = "(?:<[^()]*?>)?";

// After the parameter list: a Java `throws` clause or a C# `where` constraint,
// then the body: none yet, an opening brace, a one-line body, or a C#
// expression body (`=> 2 * x;`).
const JVM_DEFINITION_TAIL_SOURCE = `\\s*(?:throws\\s+[$\\u200c\\u200d\\p{ID_Continue},\\s.]+?|where\\s+[^{;=]+?)?(?:=>.*$|${DEFINITION_TAIL_SOURCE})`;

/** A JVM-family method definition: modifiers, a return type, a parameter list. */
const JVM_METHOD_DEFINITION = unicodePattern(
  `^${JVM_ANNOTATIONS_SOURCE}${JVM_MEMBER_MODIFIERS_SOURCE}${JVM_LEADING_TYPE_PARAMETERS_SOURCE}${JVM_NOT_A_MODIFIER_SOURCE}${IDENTIFIER_START_SOURCE}[$\\u200c\\u200d\\p{ID_Continue}<>\\[\\].?,]*\\s+(${IDENTIFIER_SOURCE})${JVM_TRAILING_TYPE_PARAMETERS_SOURCE}\\s*${PARAMETER_LIST_SOURCE}${JVM_DEFINITION_TAIL_SOURCE}`,
);

/** Java/C# constructor definition: modifiers, the type name, a parameter list. */
const JVM_CONSTRUCTOR_DEFINITION = unicodePattern(
  `^${JVM_ANNOTATIONS_SOURCE}${JVM_MEMBER_MODIFIERS_SOURCE}(${IDENTIFIER_SOURCE})\\s*${PARAMETER_LIST_SOURCE}\\s*(?::\\s*(?:base|this)\\s*${PARAMETER_LIST_SOURCE})?${JVM_DEFINITION_TAIL_SOURCE}`,
);

/**
 * A C# property: `public int Count { get; set; }`, `public int Area => w * h;`,
 * or a header whose accessors follow on the next lines (`public string Name`
 * with nothing after it; a field always ends in `;`).
 */
const CSHARP_PROPERTY_DEFINITION = unicodePattern(
  `^${JVM_ANNOTATIONS_SOURCE}(?:(?:public|private|protected|internal)\\s+)${JVM_MEMBER_MODIFIERS_SOURCE}${JVM_NOT_A_MODIFIER_SOURCE}${IDENTIFIER_START_SOURCE}[$\\u200c\\u200d\\p{ID_Continue}<>\\[\\].?,]*\\s+(${IDENTIFIER_SOURCE})\\s*(?:\\{\\s*(?:(?:public|private|protected|internal)\\s+)?(?:get|set|init)\\b.*|=>.*|)$`,
);

const PATTERNS: Partial<Record<string, HeaderPattern[]>> = {
  python: [
    { re: unicodePattern(`^(?:async\\s+)?def\\s+(${IDENTIFIER_SOURCE})`), kind: "function" },
    { re: unicodePattern(`^class\\s+(${IDENTIFIER_SOURCE})`), kind: "class" },
  ],
  ruby: [
    { re: unicodePattern(`^def\\s+(?:self\\.)?(${IDENTIFIER_SOURCE}[!?=]?)`), kind: "method" },
    { re: unicodePattern(`^class\\s+(${IDENTIFIER_SOURCE}(?::${IDENTIFIER_SOURCE})*)`), kind: "class" },
    { re: unicodePattern(`^module\\s+(${IDENTIFIER_SOURCE}(?::${IDENTIFIER_SOURCE})*)`), kind: "module" },
  ],
  go: [
    { re: unicodePattern(`^func\\s+(?:\\([^)]*\\)\\s*)?(${IDENTIFIER_SOURCE})`), kind: "function" },
    { re: unicodePattern(`^type\\s+(${IDENTIFIER_SOURCE})`), kind: "type" },
  ],
  rust: [
    {
      re: unicodePattern(
        `^(?:pub(?:\\([^)]*\\))?\\s+)?(?:const\\s+)?(?:async\\s+)?(?:unsafe\\s+)?fn\\s+(${IDENTIFIER_SOURCE})`,
      ),
      kind: "function",
    },
    { re: unicodePattern(`^(?:pub\\s+)?struct\\s+(${IDENTIFIER_SOURCE})`), kind: "struct" },
    { re: unicodePattern(`^(?:pub\\s+)?enum\\s+(${IDENTIFIER_SOURCE})`), kind: "enum" },
    { re: unicodePattern(`^(?:pub\\s+)?trait\\s+(${IDENTIFIER_SOURCE})`), kind: "trait" },
    { re: unicodePattern(`^impl(?:<[^>]*>)?\\s+(${IDENTIFIER_SOURCE})`), kind: "impl" },
    { re: unicodePattern(`^(?:pub\\s+)?mod\\s+(${IDENTIFIER_SOURCE})`), kind: "module" },
  ],
  java: [
    {
      re: unicodePattern(`^(?:public|private|protected|abstract|final|static|\\s)*class\\s+(${IDENTIFIER_SOURCE})`),
      kind: "class",
    },
    {
      re: unicodePattern(`^(?:public|private|protected|abstract|\\s)*interface\\s+(${IDENTIFIER_SOURCE})`),
      kind: "interface",
    },
    { re: unicodePattern(`^(?:public|private|protected|\\s)*enum\\s+(${IDENTIFIER_SOURCE})`), kind: "enum" },
    { re: unicodePattern(`^(?:public|private|protected|\\s)*record\\s+(${IDENTIFIER_SOURCE})`), kind: "record" },
    { re: JVM_METHOD_DEFINITION, kind: "method", nameGroup: 1 },
    { re: JVM_CONSTRUCTOR_DEFINITION, kind: "constructor", nameGroup: 1 },
  ],
  csharp: [
    {
      re: unicodePattern(`^(?:public|private|protected|internal|abstract|sealed|static|partial|\\s)*class\\s+(${IDENTIFIER_SOURCE})`),
      kind: "class",
    },
    {
      re: unicodePattern(`^(?:public|private|protected|internal|\\s)*interface\\s+(${IDENTIFIER_SOURCE})`),
      kind: "interface",
    },
    {
      re: unicodePattern(`^(?:public|private|protected|internal|\\s)*struct\\s+(${IDENTIFIER_SOURCE})`),
      kind: "struct",
    },
    { re: unicodePattern(`^namespace\\s+(${IDENTIFIER_SOURCE}(?:\\.${IDENTIFIER_SOURCE})*)`), kind: "namespace" },
    { re: unicodePattern(`^(?:public|private|protected|internal|\\s)*enum\\s+(${IDENTIFIER_SOURCE})`), kind: "enum" },
    { re: unicodePattern(`^(?:public|private|protected|internal|\\s)*record\\s+(${IDENTIFIER_SOURCE})`), kind: "record" },
    { re: JVM_METHOD_DEFINITION, kind: "method", nameGroup: 1 },
    { re: JVM_CONSTRUCTOR_DEFINITION, kind: "constructor", nameGroup: 1 },
    { re: CSHARP_PROPERTY_DEFINITION, kind: "property", nameGroup: 1 },
  ],
  cpp: [
    { re: unicodePattern(`^(?:template\\s*<[^>]*>\\s*)?class\\s+(${IDENTIFIER_SOURCE})${C_TYPE_HEAD_END_SOURCE}`), kind: "class" },
    { re: unicodePattern(`^(?:template\\s*<[^>]*>\\s*)?struct\\s+(${IDENTIFIER_SOURCE})${C_TYPE_HEAD_END_SOURCE}`), kind: "struct" },
    { re: unicodePattern(`^namespace\\s+(${IDENTIFIER_SOURCE})`), kind: "namespace" },
    { re: unicodePattern(`^(?:enum\\s+class|enum)\\s+(${IDENTIFIER_SOURCE})${C_TYPE_HEAD_END_SOURCE}`), kind: "enum" },
    { re: C_TYPEDEF_TAG, kind: "struct" },
    { re: C_TYPEDEF_CLOSE, kind: "type" },
    { re: unicodePattern(`^using\\s+(${IDENTIFIER_SOURCE})\\s*=`), kind: "type" },
    // A definition, not a call or a control statement: a return type, a name,
    // a parameter list, and an opening brace on the same line.
    { re: C_FUNCTION_DEFINITION, kind: "function", nameGroup: 1 },
    { re: CPP_QUALIFIED_CONSTRUCTOR, kind: "constructor", nameGroup: 1 },
    { re: CPP_MEMBER_CONSTRUCTOR, kind: "constructor", nameGroup: 1, namesKnownType: true },
  ],
  c: [
    { re: unicodePattern(`^struct\\s+(${IDENTIFIER_SOURCE})${C_TYPE_HEAD_END_SOURCE}`), kind: "struct" },
    { re: unicodePattern(`^union\\s+(${IDENTIFIER_SOURCE})${C_TYPE_HEAD_END_SOURCE}`), kind: "union" },
    { re: unicodePattern(`^enum\\s+(${IDENTIFIER_SOURCE})${C_TYPE_HEAD_END_SOURCE}`), kind: "enum" },
    { re: C_TYPEDEF_TAG, kind: "struct" },
    { re: C_TYPEDEF_CLOSE, kind: "type" },
    { re: unicodePattern(`^typedef\\s+.*?\\b(${IDENTIFIER_SOURCE})\\s*;`), kind: "type" },
    { re: C_FUNCTION_DEFINITION, kind: "function", nameGroup: 1 },
  ],
  php: [
    {
      re: unicodePattern(
        `^namespace\\s+(${IDENTIFIER_SOURCE}(?:\\\\${IDENTIFIER_SOURCE})*)`,
      ),
      kind: "namespace",
    },
    {
      re: unicodePattern(
        `^(?:(?:abstract|final|readonly)\\s+)*class\\s+(${IDENTIFIER_SOURCE})`,
      ),
      kind: "class",
    },
    { re: unicodePattern(`^interface\\s+(${IDENTIFIER_SOURCE})`), kind: "interface" },
    { re: unicodePattern(`^trait\\s+(${IDENTIFIER_SOURCE})`), kind: "trait" },
    {
      re: unicodePattern(
        `^(?:(?:public|protected|private|static|final|abstract|readonly)\\s+)*enum\\s+(${IDENTIFIER_SOURCE})`,
      ),
      kind: "enum",
    },
    {
      re: unicodePattern(
        `^(?:(?:public|protected|private|static|final|abstract)\\s+)*function\\s+&?(${IDENTIFIER_SOURCE})`,
      ),
      kind: "function",
    },
  ],
  swift: [
    {
      re: unicodePattern(
        `^(?:(?:public|private|fileprivate|internal|open|final|static|class|mutating|nonmutating|override|required|convenience|distributed|nonisolated)\\s+)*func\\s+(${IDENTIFIER_SOURCE})`,
      ),
      kind: "function",
    },
    { re: unicodePattern(`^(?:public\\s+|private\\s+|internal\\s+|open\\s+|final\\s+)*class\\s+(${IDENTIFIER_SOURCE})`), kind: "class" },
    { re: unicodePattern(`^(?:public\\s+|private\\s+|internal\\s+|open\\s+|final\\s+)*struct\\s+(${IDENTIFIER_SOURCE})`), kind: "struct" },
    { re: unicodePattern(`^(?:public\\s+|private\\s+|internal\\s+|\\s)*enum\\s+(${IDENTIFIER_SOURCE})`), kind: "enum" },
    { re: unicodePattern(`^(?:public\\s+|private\\s+|internal\\s+|\\s)*protocol\\s+(${IDENTIFIER_SOURCE})`), kind: "protocol" },
    {
      re: unicodePattern(
        `^(?:(?:public|private|fileprivate|internal|open|final|distributed)\\s+)*actor\\s+(${IDENTIFIER_SOURCE})`,
      ),
      kind: "actor",
    },
    {
      re: unicodePattern(
        `^extension\\s+(${IDENTIFIER_SOURCE}(?:\\.${IDENTIFIER_SOURCE})*)`,
      ),
      kind: "extension",
    },
    {
      re: /^(?:(?:public|private|fileprivate|internal|required|convenience|override)\s+)*(init)\s*[?!]?\s*\(/,
      kind: "constructor",
    },
  ],
  lua: [
    { re: /^(?:local\s+)?function\s+([A-Za-z_][\w.:]*)/, kind: "function" },
    { re: /^(?:local\s+)?([A-Za-z_]\w*)\s*=\s*function\b/, kind: "function" },
  ],
  perl: [
    {
      re: unicodePattern(`^package\\s+(${IDENTIFIER_SOURCE}(?:::${IDENTIFIER_SOURCE})*)`),
      kind: "module",
    },
    {
      re: unicodePattern(`^sub\\s+(${IDENTIFIER_SOURCE}(?:::${IDENTIFIER_SOURCE})*)`),
      kind: "function",
    },
  ],
  r: [
    {
      re: unicodePattern(`^((?:\\.(?!\\d)|${IDENTIFIER_START_SOURCE})[$\\u200c\\u200d\\p{ID_Continue}._]*)\\s*(?:<-|=)\\s*function\\b`),
      kind: "function",
    },
  ],
  sql: [
    {
      re: new RegExp(
        `^create\\s+(?:or\\s+replace\\s+)?(?:table|view|function|procedure|trigger|index)\\s+(?:if\\s+not\\s+exists\\s+)?[\`"\\[]?(${IDENTIFIER_SOURCE}(?:\\.${IDENTIFIER_SOURCE})*)`,
        "iu",
      ),
      kind: "definition",
    },
  ],
  bash: [
    { re: /^(?:function\s+)?([A-Za-z_]\w*)\s*\(\s*\)/, kind: "function" },
    { re: /^function\s+([A-Za-z_]\w*)/, kind: "function" },
  ],
  markdown: [{ re: /^(#{1,6})\s+(.+?)\s*#*$/, kind: "heading", nameGroup: 2 }],
  // A config file's keys are its structure, and indentation already drives the
  // depth calculation below — so a YAML outline nests the way the file reads.
  yaml: [
    {
      re: unicodePattern(`^-?\\s*(${IDENTIFIER_START_SOURCE}[$\\u200c\\u200d\\p{ID_Continue}.-]*)\\s*:(?:\\s|$)`),
      kind: "key",
    },
  ],
  // The catalog files .ini and .toml under one id, which normalizes to "toml".
  toml: [
    { re: /^\[+([^\]]+)\]+/, kind: "section" },
    { re: /^([A-Za-z_][\w.-]*)\s*=/, kind: "key" },
  ],
  json: [{ re: /^"([^"]+)"\s*:/, kind: "key" }],
  // As VS Code's CSS outline: at-rules and the selectors of rule blocks,
  // nested by indentation (a rule inside @media, a nested `&:hover`).
  css: [
    { re: /^(@[\w-]+[^{;]*?)\s*\{/, kind: "at-rule" },
    { re: /^([^@{};\s/][^{};]*?)\s*\{/, kind: "selector" },
  ],
};

/** Columns one leading tab counts as, for the indentation→depth mapping. */
const TAB_WIDTH = 2;

/** Leading-whitespace width of a line (tabs = {@link TAB_WIDTH} columns). */
function leadingIndent(line: string): number {
  let width = 0;
  for (let i = 0; i < line.length; i += 1) {
    const c = line[i];
    if (c === " ") width += 1;
    else if (c === "\t") width += TAB_WIDTH;
    else break;
  }
  return width;
}

/**
 * Extract a best-effort document outline for a NON-TS/JS `language` from `text`.
 *
 * Scans each line against the language's header patterns; on a match it locates
 * the captured NAME within the original (untrimmed) line so `from`/`to` are exact
 * buffer offsets, and derives `depth` from the relative indentation of header
 * lines (each distinct deeper indent level increments depth, popped back as
 * indentation decreases — so siblings share a depth and a nested def is +1).
 *
 * Markdown is special-cased: depth comes from the heading LEVEL (`#` count), not
 * indentation. Languages with no patterns (json/yaml/css/html/toml/plain) return
 * `[]`. TOTAL: never throws; an empty/whitespace doc returns `[]`.
 */
/** LaTeX's sectioning commands, outermost first. */
const LATEX_SECTION_LEVELS: Readonly<Record<string, number>> = {
  part: 0,
  chapter: 1,
  section: 2,
  subsection: 3,
  subsubsection: 4,
  paragraph: 5,
  subparagraph: 6,
};
const LATEX_SECTION =
  /^\\(part|chapter|section|subsection|subsubsection|paragraph|subparagraph)\*?\s*(?:\[[^\]]*\])?\s*\{/;

/**
 * A LaTeX document's outline is its sectioning, as Overleaf's outline shows
 * it: `\section{…}` and its relatives (starred or with a short title), nested
 * by level and anchored on the title. A title may hold braces of its own
 * (`\section{Intro to \LaTeX{}}`), so it is read to its balancing brace. The
 * outermost level the file uses is depth 0, so an article's sections start at
 * the top. A commented-out heading (`% \section`) never matches.
 */
function latexOutlineSymbols(text: string): OutlineSymbol[] {
  const found: Array<OutlineSymbol & { level: number }> = [];
  let offset = 0;
  for (const rawLine of text.split("\n")) {
    const leadWs = rawLine.length - rawLine.replace(/^\s+/, "").length;
    const match = LATEX_SECTION.exec(rawLine.slice(leadWs));
    if (match) {
      const open = leadWs + match[0].length;
      let depth = 1;
      let index = open;
      while (index < rawLine.length && depth > 0) {
        const char = rawLine[index];
        if (char === "\\") index += 1;
        else if (char === "{") depth += 1;
        else if (char === "}") depth -= 1;
        index += 1;
      }
      const raw = rawLine.slice(open, depth === 0 ? index - 1 : rawLine.length);
      const name = raw.trim();
      if (name) {
        const from = offset + open + (raw.length - raw.trimStart().length);
        found.push({
          name,
          kind: match[1]!,
          from,
          to: from + name.length,
          depth: 0,
          level: LATEX_SECTION_LEVELS[match[1]!]!,
        });
      }
    }
    offset += rawLine.length + 1;
  }
  const top = Math.min(...found.map((symbol) => symbol.level));
  return found.map(({ level, ...symbol }) => ({ ...symbol, depth: level - top }));
}

export function outlineSymbols(text: string, language: string): OutlineSymbol[] {
  const lang = normalizeLanguage(language);
  if (lang === "latex") return text ? latexOutlineSymbols(text) : [];
  const patterns = PATTERNS[lang];
  if (!patterns || !text) return [];

  const isMarkdown = lang === "markdown";
  const out: OutlineSymbol[] = [];
  // Indent levels seen so far (a stack of indentation widths) → depth = stack
  // size after popping levels deeper than the current line. Markdown ignores it.
  const indentStack: number[] = [];
  const knownTypeNames = new Set<string>();

  let offset = 0; // buffer offset of the current line's first char
  const lines = text.split("\n");
  for (let li = 0; li < lines.length; li += 1) {
    const rawLine = lines[li];
    // Strip a trailing \r (CRLF docs) only for matching; offsets use rawLine.
    const line = rawLine.endsWith("\r") ? rawLine.slice(0, -1) : rawLine;
    const trimmed = line.replace(/^\s+/, "");

    for (const pat of patterns) {
      if (
        (pat.re === C_FUNCTION_DEFINITION ||
          pat.re === CPP_QUALIFIED_CONSTRUCTOR ||
          pat.re === CPP_MEMBER_CONSTRUCTOR ||
          pat.re === JVM_METHOD_DEFINITION ||
          pat.re === JVM_CONSTRUCTOR_DEFINITION ||
          pat.re === CSHARP_PROPERTY_DEFINITION) &&
        NOT_A_DEFINITION.test(trimmed)
      ) {
        continue;
      }
      const m = pat.re.exec(trimmed);
      if (!m) continue;
      const name = m[pat.nameGroup ?? 1];
      if (!name) continue;
      if (pat.namesKnownType && !knownTypeNames.has(name.replace(/^~/, ""))) continue;

      // Locate the matched NAME inside the ORIGINAL line so from/to are exact
      // buffer offsets. Search from the leading-whitespace boundary so a name
      // that also appears in the indentation (impossible, but safe) isn't hit.
      const leadWs = line.length - trimmed.length;
      const nameInLine = line.indexOf(name, leadWs);
      // Defensive: if we somehow can't relocate the name, anchor at the first
      // non-whitespace char so the jump still lands on this line.
      const from = offset + (nameInLine >= 0 ? nameInLine : leadWs);
      const to = from + name.length;

      let depth: number;
      if (isMarkdown) {
        // Heading level (# count) minus 1 → depth (H1 = 0, H2 = 1, …).
        depth = Math.max(0, m[1].length - 1);
      } else {
        const indent = leadingIndent(line);
        // Pop any levels deeper than (or equal to, when this line is a sibling)
        // the current indent, then this line sits one deeper than what remains.
        while (indentStack.length > 0 && indentStack[indentStack.length - 1] >= indent) {
          indentStack.pop();
        }
        depth = indentStack.length;
        indentStack.push(indent);
      }

      out.push({ name, kind: pat.kind, from, to, depth });
      if (pat.kind === "class" || pat.kind === "struct") knownTypeNames.add(name);
      break; // first matching pattern wins for this line
    }

    // Advance past this line plus its newline (rawLine excludes the \n split off).
    offset += rawLine.length + 1;
  }

  return out;
}
