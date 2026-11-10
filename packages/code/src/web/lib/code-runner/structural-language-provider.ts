/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Go to Definition for the languages that have no language service at all.
 *
 * Go, Rust, Java, C, C++, C#, Swift, PHP, Ruby, Lua and Perl get syntax
 * highlighting, folding, structural diagnostics and an outline here — and, until
 * now, nothing whatsoever from F12. Pressing it did not fail; it did nothing,
 * which is the worse of the two.
 *
 * The outline already knows where every function, class and method in the file
 * is declared, down to the exact span of its name. That is most of an answer,
 * so this provider gives the one it can and refuses the rest:
 *
 *  - `definition` and `hover` ONLY. Not references, not rename. This resolves by
 *    NAME inside one file, narrowed by the block the caret sits in; it knows
 *    nothing about imports, overloads, or which of two identically-named
 *    siblings a call meant, and a match that good is fine for moving a caret or
 *    showing a line, and nowhere near good enough to rewrite a buffer with.
 *
 *  - Hover shows the DECLARATION AS WRITTEN — the source line, trimmed. It
 *    invents no type and summarises nothing, so it cannot be wrong about the
 *    code in the way a guessed signature could; it is the line the user would
 *    have scrolled to, brought to them instead.
 *
 *  - A name declared EXACTLY ONCE in the file resolves to that declaration. Two
 *    classes in one Java file can both declare `toString`, and jumping to
 *    whichever came first would be wrong more often than right, so a repeated
 *    name goes to the scope rule below, and anything that rule cannot settle is
 *    answered with silence. Silence is a worse answer than a correct jump and a
 *    much better one than a confident wrong jump; a real language server
 *    registered later takes this over automatically.
 *
 *  - A repeated name is resolved BY SCOPE, the way the language reads: look in
 *    the block the caret is in, then the block around that, out to the file, and
 *    take the first level that declares the name exactly once. A call inside
 *    `class Alpha` finds Alpha's method; a call inside `class Beta` finds
 *    Beta's; a call at file scope finds neither and gets nothing.
 *
 *    This was tried once before and REVERTED, because it was built on the
 *    outline's `depth`, which comes from indentation — so a statement written
 *    after `class Beta` closes looked exactly like one written inside it, and
 *    the rule jumped into Beta from outside. What was missing was where a block
 *    ENDS; {@link import("./block-extents").blockExtents} supplies that now, by
 *    scanning braces while stepping over each language's comments and strings.
 *    It refuses (null) for languages it does not model and for any document
 *    whose braces do not balance — which is most documents mid-keystroke — and
 *    every refusal lands back on the exactly-once rule, never on a guess.
 *
 * Registered last, so any provider that actually understands the language wins.
 */

import {
  blockExtents,
  enclosingBlocks,
  innermostBlock,
  sameBlock,
  type BlockExtent,
} from "./block-extents";
import { identifierRangeAt } from "./code-intelligence";
import { normalizeLanguage } from "./language-metadata";
import { outlineSymbols, type OutlineSymbol } from "./outline-symbols";
import type {
  LanguageServiceProvider,
  LanguageServiceRequest,
  ProviderCapabilities,
  ProviderHover,
  ProviderLocation,
} from "./language-service-provider";

/**
 * The languages this serves: every one the outline can read that has no engine
 * of its own here.
 *
 * Python is absent on purpose — jedi answers it, against the live kernel, and
 * this would be strictly worse. So are the data formats (JSON, YAML, TOML,
 * Markdown): their outline entries are keys and headings, and "go to the
 * definition of a key" is not a question anyone is asking.
 */
const LANGUAGES = [
  "go",
  "rust",
  "java",
  "csharp",
  "cpp",
  "c",
  "php",
  "swift",
  "lua",
  "perl",
  "ruby",
  "r",
  // Both have declarations worth jumping to in one buffer: a `create table` a
  // query below it selects from, and a shell function called further down.
  "sql",
  "bash",
] as const;

/**
 * Languages whose identifiers do not care about case. SQL is the one here, and
 * matching `CUSTOMERS` to `create table customers` is not a convenience — it is
 * what that language means by the same name.
 */
const CASE_INSENSITIVE = new Set<string>(["sql"]);

const CAPABILITIES: ProviderCapabilities = { definition: true, hover: true };
const NOTHING: ProviderCapabilities = {};

/** How much of a declaration line to show. Long enough for a signature. */
const MAX_HOVER_LINE_CHARS = 400;

const STRUCTURAL_IDENTIFIER_START = "[$_\\p{ID_Start}]";
const STRUCTURAL_IDENTIFIER_CONTINUE =
  "[$\\u200c\\u200d\\p{ID_Continue}]*";
const STRUCTURAL_IDENTIFIER =
  `${STRUCTURAL_IDENTIFIER_START}${STRUCTURAL_IDENTIFIER_CONTINUE}`;
const R_IDENTIFIER = new RegExp(
  `(?:\\.(?!\\d)|${STRUCTURAL_IDENTIFIER_START})[$\\u200c\\u200d\\p{ID_Continue}._]*`,
  "gu",
);
const PERL_IDENTIFIER = new RegExp(
  `${STRUCTURAL_IDENTIFIER}(?:::${STRUCTURAL_IDENTIFIER})*`,
  "gu",
);
const SQL_IDENTIFIER = new RegExp(
  `${STRUCTURAL_IDENTIFIER}(?:\\.${STRUCTURAL_IDENTIFIER})*`,
  "gu",
);

function regexRangeAt(
  text: string,
  offset: number,
  matcher: RegExp,
): { from: number; to: number } | null {
  const safeOffset = Math.max(0, Math.min(Math.trunc(offset), text.length));
  matcher.lastIndex = 0;
  let match: RegExpExecArray | null;
  while ((match = matcher.exec(text))) {
    const to = match.index + match[0].length;
    if (safeOffset >= match.index && safeOffset <= to) {
      return { from: match.index, to };
    }
    if (match.index > safeOffset) break;
  }
  return null;
}

/** The span of the identifier surrounding `offset`, or null if the caret is not on one. */
function wordSpanAt(
  text: string,
  offset: number,
  language: string,
): { from: number; to: number } | null {
  if (language === "r") return regexRangeAt(text, offset, R_IDENTIFIER);
  if (language === "perl") {
    return regexRangeAt(text, offset, PERL_IDENTIFIER);
  }
  if (language === "sql") return regexRangeAt(text, offset, SQL_IDENTIFIER);
  return identifierRangeAt(text, offset);
}

/** The identifier surrounding `offset`, or null when the caret is not on one. */
function wordAt(text: string, offset: number, language: string): string | null {
  const span = wordSpanAt(text, offset, language);
  return span ? text.slice(span.from, span.to) : null;
}

/**
 * A C++ member function defined outside its class is listed as `Box::area`,
 * but a call (`box.area()`) names it `area`: answer for the last segment, and
 * point at it. Constructors and destructors (`Box::Box`, `Box::~Box`) keep
 * their qualified name, so `Box` still means the class. (Perl's `Foo::bar`
 * is read whole, as Perl names it.)
 */
function memberName(symbol: OutlineSymbol): OutlineSymbol {
  if (symbol.kind !== "function") return symbol;
  const cut = symbol.name.lastIndexOf("::");
  if (cut < 0) return symbol;
  const start = cut + 2;
  return { ...symbol, name: symbol.name.slice(start), from: symbol.from + start };
}

/** The trimmed source line containing `offset`. */
function lineAt(text: string, offset: number): string {
  const start = text.lastIndexOf("\n", Math.max(0, offset - 1)) + 1;
  const nextNewline = text.indexOf("\n", offset);
  const end = nextNewline === -1 ? text.length : nextNewline;
  return text.slice(start, end).trim();
}

/**
 * The declaration a reader would mean, out of several with the same name.
 *
 * Walks the caret's scopes from the inside out — the block it is in, the block
 * around that, and finally the file — and takes the first level that declares
 * the name exactly once. A level that declares it TWICE stops the walk with
 * nothing: an outer declaration cannot be meant when an inner one is in the way,
 * and picking between two siblings is what a real language server is for.
 *
 * `extents` must come from a scan that succeeded; see block-extents.ts.
 */
function declarationInScope(
  declarations: readonly OutlineSymbol[],
  extents: readonly BlockExtent[],
  caret: number,
): OutlineSymbol | null {
  const containers = new Map<OutlineSymbol, BlockExtent | null>(
    declarations.map((symbol) => [symbol, innermostBlock(extents, symbol.from)]),
  );
  const levels: Array<BlockExtent | null> = [
    ...enclosingBlocks(extents, caret),
    null, // the file itself, always the last place to look
  ];
  for (const level of levels) {
    const here = declarations.filter((symbol) =>
      sameBlock(containers.get(symbol) ?? null, level),
    );
    if (here.length === 1) return here[0]!;
    if (here.length > 1) return null;
  }
  return null;
}

export function createStructuralLanguageProvider(): LanguageServiceProvider {
  const serves = (language: string): boolean =>
    (LANGUAGES as readonly string[]).includes(normalizeLanguage(language));

  // Hover fires on every pause of the pointer, and the brace scan reads the
  // whole document — so remember the last one. Keyed by the exact text, so an
  // edit invalidates it without anyone having to remember to say so.
  let scanned: { text: string; language: string; extents: BlockExtent[] | null } | null =
    null;
  const extentsFor = (text: string, language: string): BlockExtent[] | null => {
    if (scanned && scanned.text === text && scanned.language === language) {
      return scanned.extents;
    }
    const extents = blockExtents(text, language);
    scanned = { text, language, extents };
    return extents;
  };

  /**
   * The declaration of the caret's identifier, or null.
   *
   * One declaration is the answer; several go to the scope rule; anything left
   * unsettled is silence — see the module comment. Both answers this provider
   * gives are built on the same resolution, so they share it.
   */
  const resolveDeclaration = (
    request: LanguageServiceRequest & { language: string },
  ): OutlineSymbol | null => {
    if (!serves(request.language)) return null;
    const language = normalizeLanguage(request.language);
    const name = wordAt(request.text, request.offset, language);
    if (!name) return null;
    const folded = CASE_INSENSITIVE.has(language);
    const wanted = folded ? name.toLowerCase() : name;
    const declarations = outlineSymbols(request.text, language)
      .map((symbol) => (language === "cpp" ? memberName(symbol) : symbol))
      .filter((symbol) => (folded ? symbol.name.toLowerCase() : symbol.name) === wanted);
    if (declarations.length === 1) return declarations[0]!;
    if (declarations.length < 2) return null;
    const extents = extentsFor(request.text, language);
    // No trustworthy extents (unmodelled language, or braces mid-edit): the
    // repeated name stays unanswered, exactly as it was before scopes existed.
    if (!extents) return null;
    return declarationInScope(declarations, extents, request.offset);
  };

  return {
    id: "structural-outline",
    languages: LANGUAGES,

    capabilities(language: string): ProviderCapabilities {
      return serves(language) ? CAPABILITIES : NOTHING;
    },

    async hover(
      request: LanguageServiceRequest & { language: string },
    ): Promise<ProviderHover | null> {
      const declaration = resolveDeclaration(request);
      if (!declaration) return null;
      const span = wordSpanAt(
        request.text,
        request.offset,
        normalizeLanguage(request.language),
      );
      // Standing ON the declaration: repeating the line the caret is already in
      // is noise, not information.
      if (span && span.from === declaration.from) return null;
      const line = lineAt(request.text, declaration.from);
      if (!line) return null;
      return {
        contents:
          line.length > MAX_HOVER_LINE_CHARS
            ? `${line.slice(0, MAX_HOVER_LINE_CHARS)}…`
            : line,
        ...(span ? { span } : {}),
      };
    },

    async definition(
      request: LanguageServiceRequest & { language: string },
    ): Promise<ProviderLocation | null> {
      const declaration = resolveDeclaration(request);
      if (!declaration) return null;
      // Already standing on the declaration: report nothing rather than
      // "navigate" to where the caret already is.
      if (
        request.offset >= declaration.from &&
        request.offset <= declaration.to
      ) {
        return null;
      }
      return {
        bufferId: null,
        from: declaration.from,
        to: declaration.to,
      };
    },
  };
}
