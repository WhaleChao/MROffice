/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Beyond-baseline (VSCode-grade) CODE INTELLIGENCE for the MRNote CodeMirror 6
 * editor — hover quick-info, go-to-definition, and rename-symbol for
 * js / jsx / ts / tsx — without any new dependency.
 *
 * The already-installed TypeScript compiler exposes a full `LanguageService`
 * (the same engine behind tsserver / VSCode). We drive it over a SINGLE virtual
 * file whose contents are the live editor buffer, mirroring the in-memory,
 * no-filesystem approach already used for diagnostics
 * (see lib/code-runner/lsp-diagnostics.ts): `noLib: true` + a tiny synthetic
 * lib so common built-ins (console, Array methods, …) resolve, and an in-memory
 * host that serves exactly the buffer file plus that lib. This works identically
 * in node (where tests run against the REAL compiler) and the browser, and the
 * heavy compiler is pulled in lazily via dynamic `import("typescript")` so it
 * stays out of the initial editor chunk.
 *
 * Everything here is PURE / framework-agnostic and exhaustively unit-tested:
 *  - {@link createTsService}   — wrap the buffer in a LanguageService with an
 *                                incrementing version so the host re-reads the
 *                                latest text on every query.
 *  - {@link quickInfoToHtml}   — turn `getQuickInfoAtPosition` display parts into
 *                                SAFE tooltip content (escaped text; never raw
 *                                HTML injection).
 *  - {@link definitionSpan}    — same-file declaration span for go-to-definition.
 *  - {@link typeDefinitionSpan} — same-file TYPE declaration span for
 *                                go-to-type-definition (the symbol's type, not it).
 *  - {@link implementationSpans} — every same-file concrete-implementation span for
 *                                go-to-implementation (interface/abstract members).
 *  - {@link referenceSpans}    — every SAME-FILE reference location of the symbol
 *                                under the caret (find-all-references).
 *  - {@link renameEdits}       — ordered {from,to,insert} edits renaming EVERY
 *                                occurrence of the symbol under the caret.
 *  - {@link isLegalIdentifier} — validate a proposed new name.
 *
 * Cross-file definitions are N/A here: the service models exactly one buffer, so
 * a definition that resolves into the synthetic lib (or any non-buffer file) is
 * reported as "not in this file" rather than a bogus in-buffer jump.
 *
 * The live tooltip / cursor / inline-rename UI is browser-gated; the wiring in
 * CodeMirrorEditor.tsx consumes these pure functions.
 */
import type * as TsModule from "typescript";

import { normalizeLanguage } from "./codemirror-config";
import { TYPESCRIPT_AMBIENT_LIB } from "./typescript-ambient-lib";

/**
 * The four flavours code intelligence supports. CM6's JS package parses TS, but
 * the LanguageService needs the real ScriptKind to type-check annotations / JSX.
 */
export type IntelligenceLanguage = "javascript" | "jsx" | "typescript" | "tsx";

/** ts.ScriptKind numbers kept as plain values (JS=1, JSX=2, TS=3, TSX=4). */
const SCRIPT_KIND: Record<IntelligenceLanguage, number> = {
  javascript: 1,
  jsx: 2,
  typescript: 3,
  tsx: 4,
};

const FILE_BASENAMES: Record<IntelligenceLanguage, string> = {
  javascript: "__mrnote_intel__.js",
  jsx: "__mrnote_intel__.jsx",
  typescript: "__mrnote_intel__.ts",
  tsx: "__mrnote_intel__.tsx",
};

const LIB_FILE_NAME = "__mrnote_intel_lib__.d.ts";

/**
 * Ambient built-ins under `noLib: true` (otherwise hover/definition would hit
 * "Cannot find name"); the same declarations Problems and the open-window
 * project compile against.
 */
const SYNTHETIC_LIB = TYPESCRIPT_AMBIENT_LIB;

/**
 * Whether code intelligence is available for a language. JS/TS/JSX/TSX get the
 * full LanguageService; every other language (python/sql/css/…) returns
 * `undefined` from the factory so callers cleanly skip wiring it.
 */
export function intelligenceLanguage(language: string): IntelligenceLanguage | null {
  const lang = normalizeLanguage(language);
  if (lang === "javascript" || lang === "jsx" || lang === "typescript" || lang === "tsx") {
    return lang;
  }
  return null;
}

/** A character range in the buffer, 0-based `from` inclusive, `to` exclusive. */
export interface BufferSpan {
  from: number;
  to: number;
}

/** A single text edit, in CM6's `{from,to,insert}` change shape. */
export interface BufferEdit {
  from: number;
  to: number;
  insert: string;
}

/**
 * A single Quick Fix / code action (VSCode's Ctrl+. / Cmd+.): a human-readable
 * `description` (e.g. "Change spelling to 'value'") plus the SAME-FILE
 * {@link BufferEdit}s that apply it, ready to dispatch as ONE CM6 transaction.
 * The edits are ordered + non-overlapping so they can be applied as a single
 * `changes` array (one undo step).
 */
export interface CodeFixAction {
  /** The action label shown in the Quick Fix menu (plain text, not HTML). */
  description: string;
  /** The same-file edits that apply this fix, ordered + non-overlapping. */
  edits: BufferEdit[];
}

/**
 * A single same-file reference of a symbol, ready to drive a References panel:
 * the character range plus its 1-based line / 0-based column and the trimmed
 * source line text for a list preview (VSCode's "Find All References").
 */
/** Maximum source-line preview carried by one reference result. */
export const CODE_REFERENCE_PREVIEW_MAX_CHARS = 500;

export interface ReferenceSpan {
  /** 0-based buffer offset of the reference's first char (inclusive). */
  from: number;
  /** 0-based buffer offset just past the reference's last char (exclusive). */
  to: number;
  /** 1-based line number the reference starts on. */
  line: number;
  /** 0-based column (offset within the line) the reference starts at. */
  column: number;
  /** Trimmed text of the source line the reference lives on (panel preview). */
  preview: string;
}

/**
 * A single symbol in the document outline (VSCode's Go-to-Symbol / Ctrl+Shift+O),
 * flattened from the TS `NavigationTree`: the symbol `name`, its `kind`
 * (e.g. "class", "method", "function", "const"), the buffer span to jump to, and
 * its nesting `depth` (0 for a top-level symbol, 1 for a member of a top-level
 * container, …) so the picker can indent members under their container.
 */
export interface DocumentSymbol {
  /** The symbol's display name (e.g. "greet", "Calculator"). */
  name: string;
  /** The symbol kind from the navigation tree (e.g. "class", "method"). */
  kind: string;
  /** 0-based buffer offset of the symbol's span start (inclusive). */
  from: number;
  /** 0-based buffer offset just past the symbol's span end (exclusive). */
  to: number;
  /** Nesting depth: 0 = top-level, 1 = member of a top-level container, … */
  depth: number;
}

/**
 * A single Code Lens (VSCode's inline "N references" annotation shown ABOVE a
 * symbol's definition line): the symbol `name`, the 1-based `line` its definition
 * starts on (where the lens widget is anchored), the 0-based buffer `nameOffset`
 * of the symbol's NAME (the position a reference query / the references panel
 * resolves the symbol from — NOT the `function`/`class` keyword), and the
 * `refCount` of NON-DECLARATION references (see {@link codeLenses} for the rule).
 */
export interface CodeLens {
  /** The symbol's display name (e.g. "greet", "Calculator"). */
  name: string;
  /** 1-based line number the symbol's definition starts on (the lens anchor). */
  line: number;
  /**
   * 0-based buffer offset of the symbol's NAME identifier — the position the
   * references panel / a reference query resolves the symbol from. This is the
   * navigation-tree `nameSpan.start`, NOT the broader definition span start
   * (which for a `function foo` points at the `function` keyword).
   */
  nameOffset: number;
  /**
   * Number of NON-DECLARATION references to the symbol. By VSCode convention the
   * symbol's own declaration is NOT counted — so a function called twice shows
   * "2 references", and an unused symbol shows "0 references". See {@link codeLenses}.
   */
  refCount: number;
}

/** Sanitised hover content: a primary type signature plus optional doc text. */
export interface QuickInfoHtml {
  /** HTML-escaped one-line type/quick-info signature (may be empty). */
  signature: string;
  /** HTML-escaped documentation block, newlines preserved as <br/> (optional). */
  documentation: string;
  /** Whether there is anything worth showing. */
  hasContent: boolean;
}

/**
 * The pure handle returned by {@link createTsService}. `fileName` is exposed so
 * callers can compare definition `fileName`s against the buffer file (to reject
 * cross-file / lib definitions).
 */
export interface TsIntelligenceService {
  /** The underlying TS LanguageService (queried per keystroke). */
  service: TsModule.LanguageService;
  /** The single in-memory buffer file name. */
  fileName: string;
  /** Bump the version so the host re-reads `getText()` on the next query. */
  markDirty: () => void;
  /** The TS module (handy for callers that need ts enums/helpers). */
  ts: typeof TsModule;
}

/**
 * Build an in-memory single-file {@link TsModule.LanguageService}.
 *
 * `getText` is read lazily on every host `getScriptSnapshot` call, and the
 * script version is a string derived from a monotonically-increasing counter
 * that {@link TsIntelligenceService.markDirty} bumps — so the service always
 * sees the freshest buffer without rebuilding a Program. The caller passes the
 * (lazily-imported) `ts` module; this keeps the function pure and lets tests
 * pass the real compiler directly.
 *
 * Non-config: returns `null` for unsupported languages so callers can skip.
 */
export function createTsService(
  ts: typeof TsModule,
  getText: () => string,
  language: string,
): TsIntelligenceService | null {
  const lang = intelligenceLanguage(language);
  if (!lang) return null;

  const fileName = FILE_BASENAMES[lang];
  const scriptKind = SCRIPT_KIND[lang] as TsModule.ScriptKind;

  // Version is bumped on edit so the LanguageService invalidates its cached
  // snapshot. Start at 0; markDirty -> next query re-reads getText().
  let version = 0;

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
    allowNonTsExtensions: true,
  };

  const host: TsModule.LanguageServiceHost = {
    // BOTH the live buffer AND the synthetic ambient lib are program roots —
    // under `noLib` the default lib is skipped, so the lib must be an explicit
    // root for its built-in types (String/Array/Promise/console/…) to load. That
    // is what powers member-access completion (`arr.` -> map/filter/length) and
    // typed hovers on built-ins; without it the lib snapshot below is never read.
    getScriptFileNames: () => [fileName, LIB_FILE_NAME],
    getScriptVersion: (name) =>
      name === fileName ? String(version) : "0",
    getScriptSnapshot: (name) => {
      if (name === fileName) {
        return ts.ScriptSnapshot.fromString(getText());
      }
      if (name === LIB_FILE_NAME) {
        return ts.ScriptSnapshot.fromString(SYNTHETIC_LIB);
      }
      return undefined;
    },
    getScriptKind: (name) => (name === fileName ? scriptKind : ts.ScriptKind.TS),
    getCurrentDirectory: () => "",
    getCompilationSettings: () => compilerOptions,
    getDefaultLibFileName: () => LIB_FILE_NAME,
    fileExists: (name) => name === fileName || name === LIB_FILE_NAME,
    readFile: (name) =>
      name === fileName
        ? getText()
        : name === LIB_FILE_NAME
          ? SYNTHETIC_LIB
          : undefined,
    getDirectories: () => [],
  };

  const service = ts.createLanguageService(host, ts.createDocumentRegistry());

  return {
    service,
    fileName,
    markDirty: () => {
      version += 1;
    },
    ts,
  };
}

/** HTML-escape so quick-info text can never inject markup into the tooltip. */
function escapeHtml(text: string): string {
  return text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

/**
 * Convert a `ts.QuickInfo` into SANITISED tooltip content. The compiler returns
 * `displayParts` (the type signature, e.g. `const x: number`) and
 * `documentation` (JSDoc text). Both are plain strings we concatenate and
 * HTML-escape — there is NO raw HTML passthrough, so a malicious symbol name /
 * doc comment cannot inject markup.
 *
 * Pure and total: a `null`/empty quick-info yields `hasContent: false`.
 */
export function quickInfoToHtml(
  info: TsModule.QuickInfo | null | undefined,
): QuickInfoHtml {
  const empty: QuickInfoHtml = { signature: "", documentation: "", hasContent: false };
  if (!info) return empty;

  const signatureRaw = (info.displayParts ?? [])
    .map((p) => p.text)
    .join("");
  const docRaw = (info.documentation ?? []).map((p) => p.text).join("");

  const signature = escapeHtml(signatureRaw).trim();
  // Preserve doc line breaks as <br/> AFTER escaping (so the <br/> we add is the
  // only markup that survives — the doc text itself is fully escaped first).
  const documentation = escapeHtml(docRaw)
    .replace(/\r\n|\r|\n/g, "<br/>")
    .trim();

  const hasContent = signature.length > 0 || documentation.length > 0;
  return { signature, documentation, hasContent };
}

/**
 * Quick-info (hover) at a buffer offset, or `null` if there's nothing to show.
 * Non-throwing: any compiler error degrades to `null`.
 */
export function quickInfoAt(
  svc: TsIntelligenceService,
  pos: number,
): TsModule.QuickInfo | null {
  try {
    return svc.service.getQuickInfoAtPosition(svc.fileName, pos) ?? null;
  } catch {
    return null;
  }
}

/** A single semantic completion entry — the engine behind member IntelliSense. */
export interface SemanticCompletionEntry {
  /** The inserted name / label, e.g. "map", "length". */
  name: string;
  /** CM completion type (icon): "method" | "property" | "function" | "variable" | … */
  type: string;
  /** TS sort weight (lexicographically lower sorts first). */
  sortText: string;
}

/** A semantic completion result: where it applies + the entries. */
export interface SemanticCompletionResult {
  /** 0-based offset the completion replaces FROM (the member/word start). */
  from: number;
  /** 0-based offset the completion replaces TO (the caret). */
  to: number;
  /** Whether this is a member-access completion (the entries follow a `.`). */
  isMember: boolean;
  /** Distinct entries, mapped to CM completion types. */
  entries: SemanticCompletionEntry[];
}

/** ScriptElementKind → CM6 completion type (icon class). */
const TS_KIND_TO_CM: Record<string, string> = {
  method: "method",
  property: "property",
  getter: "property",
  setter: "property",
  function: "function",
  "local function": "function",
  constructor: "method",
  var: "variable",
  let: "variable",
  const: "variable",
  "local var": "variable",
  parameter: "variable",
  alias: "variable",
  class: "class",
  "local class": "class",
  interface: "interface",
  type: "type",
  "type parameter": "type",
  enum: "enum",
  "enum member": "enum",
  module: "namespace",
  "external module name": "namespace",
  keyword: "keyword",
};

/**
 * Semantic completions at a buffer offset — the real {@link
 * TsModule.LanguageService} `getCompletionsAtPosition` (the engine behind
 * VSCode's IntelliSense). This is what lifts JS/TS completion past mere
 * identifier scraping: after `arr.` it returns `map`/`filter`/`length` with their
 * kinds, typed to the receiver. Returns `null` when there is nothing to complete.
 * Pure + TOTAL: any compiler error degrades to `null`.
 */
export function completionsAt(
  svc: TsIntelligenceService,
  pos: number,
): SemanticCompletionResult | null {
  let info: TsModule.WithMetadata<TsModule.CompletionInfo> | undefined;
  try {
    info = svc.service.getCompletionsAtPosition(svc.fileName, pos, undefined);
  } catch {
    return null;
  }
  if (!info || info.entries.length === 0) return null;
  // Replacement span: prefer TS's own, else walk back over identifier chars so we
  // replace the partial member the user has typed (never the `.` before it).
  let from = pos;
  const to = pos;
  if (info.optionalReplacementSpan) {
    from = info.optionalReplacementSpan.start;
  } else {
    const source =
      svc.service.getProgram()?.getSourceFile(svc.fileName)?.text ?? "";
    let s = pos;
    while (s > 0 && /[\w$]/.test(source[s - 1])) s -= 1;
    from = s;
  }
  const entries: SemanticCompletionEntry[] = [];
  const seen = new Set<string>();
  for (const e of info.entries) {
    if (seen.has(e.name)) continue;
    seen.add(e.name);
    entries.push({
      name: e.name,
      type: TS_KIND_TO_CM[e.kind] ?? "variable",
      sortText: e.sortText,
    });
  }
  return { from, to, isMember: Boolean(info.isMemberCompletion), entries };
}

/** A completion entry's lazily-fetched detail: its type signature + docs. */
export interface CompletionDetail {
  /** The type signature, e.g. "(method) Array<number>.map<U>(...): U[]". */
  detail: string;
  /** JSDoc / documentation text (may be empty). */
  documentation: string;
}

/**
 * Lazily fetch the full detail (type signature + docs) for ONE completion entry —
 * VSCode's details panel that appears when a completion is highlighted. This is
 * the second, per-entry call (`getCompletionEntryDetails`), done lazily (only for
 * the focused item) so the completion LIST stays cheap. Returns plain text (the
 * caller renders via textContent — no markup). `null` when unavailable.
 */
export function completionEntryDetail(
  svc: TsIntelligenceService,
  pos: number,
  name: string,
): CompletionDetail | null {
  let info: TsModule.CompletionEntryDetails | undefined;
  try {
    info = svc.service.getCompletionEntryDetails(
      svc.fileName,
      pos,
      name,
      undefined,
      undefined,
      undefined,
      undefined,
    );
  } catch {
    return null;
  }
  if (!info) return null;
  const detail = svc.ts.displayPartsToString(info.displayParts ?? []);
  const documentation = svc.ts.displayPartsToString(info.documentation ?? []);
  if (!detail && !documentation) return null;
  return { detail, documentation };
}

/**
 * Sanitised signature help (parameter hints): the active overload's label, its
 * parameter labels, and which one the caret is on. Mirrors VSCode's parameter
 * hints popup; every string is HTML-escaped (no raw markup passthrough).
 */
export interface SignatureHelpInfo {
  /** HTML-escaped one-line signature, e.g. `add(a: number, b: number): number`. */
  label: string;
  /** HTML-escaped per-parameter labels, e.g. `["a: number", "b: number"]`. */
  parameters: string[];
  /** 0-based index of the parameter the caret is currently inside. */
  activeParameter: number;
  /** Whether there is a signature worth showing. */
  hasContent: boolean;
}

/**
 * Signature help (parameter hints) at a buffer offset, or a `hasContent:false`
 * result when the caret is not inside a call. Drives the `(` / `,` tooltip.
 *
 * Uses `getSignatureHelpItems` (the engine behind VSCode's parameter hints):
 * picks the SELECTED overload, joins its prefix/parameter/suffix display parts
 * into a single label, extracts each parameter's label, and reports the active
 * argument index (clamped into range). Everything is HTML-escaped via the same
 * sanitiser as quick-info, so a malicious parameter name can't inject markup.
 *
 * Pure and TOTAL: a missing call site, a `null` result, or any compiler error
 * all degrade to `{ hasContent: false }` (never throws).
 */
export function signatureHelpAt(
  svc: TsIntelligenceService,
  pos: number,
): SignatureHelpInfo {
  const empty: SignatureHelpInfo = {
    label: "",
    parameters: [],
    activeParameter: 0,
    hasContent: false,
  };
  try {
    const help = svc.service.getSignatureHelpItems(svc.fileName, pos, undefined);
    if (!help || !help.items || help.items.length === 0) return empty;

    const selected = Math.max(
      0,
      Math.min(help.selectedItemIndex ?? 0, help.items.length - 1),
    );
    const item = help.items[selected];
    if (!item) return empty;

    const partsToText = (parts: readonly TsModule.SymbolDisplayPart[] | undefined): string =>
      (parts ?? []).map((p) => p.text).join("");

    const prefix = partsToText(item.prefixDisplayParts);
    const suffix = partsToText(item.suffixDisplayParts);
    const separator = partsToText(item.separatorDisplayParts) || ", ";
    const paramTexts = (item.parameters ?? []).map((p) => partsToText(p.displayParts));

    const labelRaw = `${prefix}${paramTexts.join(separator)}${suffix}`;
    const label = escapeHtml(labelRaw).trim();
    const parameters = paramTexts.map((p) => escapeHtml(p).trim());

    // `argumentIndex` is the caret's argument slot; clamp to a real parameter so
    // the highlight never points past the last hint.
    const rawActive = help.argumentIndex ?? 0;
    const activeParameter =
      parameters.length === 0
        ? 0
        : Math.max(0, Math.min(rawActive, parameters.length - 1));

    const hasContent = label.length > 0;
    return { label, parameters, activeParameter, hasContent };
  } catch {
    return empty;
  }
}

/**
 * A single inlay hint (VSCode's inline parameter-name / inferred-type labels):
 * the (already-trimmed) hint `text`, the buffer `offset` it anchors at, and its
 * `kind` so the renderer can style a parameter-name hint (`paramName:`,
 * `kind: "Parameter"`) differently from an inferred-type hint (`: Type`,
 * `kind: "Type"` / `"Enum"`). The text is the RAW compiler label; the renderer
 * is responsible for any escaping (a CM6 widget sets it via `textContent`, so it
 * can never inject markup).
 */
export interface InlayHintItem {
  /** The hint label to show inline, e.g. `name` (parameter) or `: number` (type). */
  text: string;
  /** 0-based buffer offset the hint anchors at (clamped into the doc). */
  offset: number;
  /** The TS hint kind: `"Parameter"`, `"Type"`, or `"Enum"`. */
  kind: string;
}

/**
 * Inlay hints (VSCode's inline parameter-name + inferred-type labels) for the
 * `[fromOffset, toOffset)` buffer range — the SAME engine VSCode renders inline.
 *
 * Drives `provideInlayHints(fileName, span, preferences)` over the VISIBLE range
 * only (the caller passes the viewport span), so the compiler never computes
 * hints for the whole document. The span is normalised + clamped into the live
 * buffer first (start ≤ end, both within `[0, docLength]`) so a stale/oversized
 * viewport offset can never make TS throw. Preferences enable the two hint
 * families this editor surfaces:
 *  - parameter-name hints on ALL arguments (`includeInlayParameterNameHints:
 *    "all"`), suppressed only when the argument text already equals the parameter
 *    name (`…WhenArgumentMatchesName: false`) so `foo(x)` with a param `x` is not
 *    cluttered;
 *  - inferred variable-type hints (`includeInlayVariableTypeHints`), again
 *    suppressed when the type name matches the variable name.
 *
 * Each returned `InlayHint` is mapped to `{text, offset, kind}` with the label
 * TRIMMED (TS pads some labels with leading/trailing spaces it expects the
 * renderer to honour via `whitespaceBefore/After`; we render our own spacing).
 * Hints whose `position` falls OUTSIDE the clamped span (or past the buffer) are
 * dropped, and the list is sorted by `offset` ascending for a deterministic,
 * non-overlapping decoration order.
 *
 * NON-THROWING / TOTAL: returns `[]` for a non-TS service, an empty/garbage
 * range, no hints, or ANY compiler throw — exactly mirroring {@link referenceSpans}.
 * Pure: it computes hints but mutates nothing.
 */
export function inlayHints(
  svc: TsIntelligenceService,
  fromOffset: number,
  toOffset: number,
): InlayHintItem[] {
  try {
    const ts = svc.ts;
    const fileName = svc.fileName;
    // The live buffer length bounds the span so a viewport offset past the doc
    // end (e.g. after a delete) can never ask TS for an out-of-range span.
    const source = svc.service.getProgram()?.getSourceFile(fileName)?.text;
    const docLength = typeof source === "string" ? source.length : 0;
    // Normalise: clamp into [0, docLength] and ensure start ≤ end so a reversed
    // or oversized viewport range still probes a sane span.
    const lo = Math.max(0, Math.min(fromOffset, toOffset));
    const hi = Math.max(0, Math.max(fromOffset, toOffset));
    const start = Math.min(lo, docLength);
    const end = Math.min(hi, docLength);
    // An empty span (e.g. an empty doc) has nothing to annotate.
    if (end <= start) return [];

    const preferences: TsModule.UserPreferences = {
      includeInlayParameterNameHints: "all",
      includeInlayParameterNameHintsWhenArgumentMatchesName: false,
      includeInlayFunctionParameterTypeHints: false,
      includeInlayVariableTypeHints: true,
      includeInlayVariableTypeHintsWhenTypeMatchesName: false,
      includeInlayPropertyDeclarationTypeHints: false,
      includeInlayFunctionLikeReturnTypeHints: false,
      includeInlayEnumMemberValueHints: false,
    };

    const span: TsModule.TextSpan = { start, length: end - start };
    const hints = svc.service.provideInlayHints(fileName, span, preferences);
    if (!hints || hints.length === 0) return [];

    const items: InlayHintItem[] = [];
    for (const hint of hints) {
      const offset = hint.position;
      // Drop a hint whose anchor escaped the clamped span / buffer — its widget
      // offset would be stale or out of range.
      if (typeof offset !== "number" || offset < 0 || offset > docLength) continue;
      // `text` is the empty string when the compiler returns `displayParts`
      // (interactive hints); fall back to joining those parts so the label is
      // never blank.
      const raw =
        hint.text && hint.text.length > 0
          ? hint.text
          : (hint.displayParts ?? []).map((p) => p.text).join("");
      const text = raw.trim();
      if (text.length === 0) continue;
      const kind = String(hint.kind ?? ts.InlayHintKind.Type);
      items.push({ text, offset, kind });
    }
    // Deterministic, non-overlapping order: by anchor offset ascending.
    items.sort((a, b) => a.offset - b.offset);
    return items;
  } catch {
    return [];
  }
}

/**
 * Same-file definition span for go-to-definition at a buffer offset.
 *
 * Uses `getDefinitionAtPosition` and returns the FIRST definition that lives in
 * the buffer file (`textSpan` → `{from,to}`). Definitions that resolve into the
 * synthetic lib or any other file are skipped (cross-file is N/A for a single
 * buffer), so the caller never makes a bogus in-buffer jump. Returns `null` when
 * there is no in-file definition. Non-throwing.
 */
export function definitionSpan(
  svc: TsIntelligenceService,
  pos: number,
): BufferSpan | null {
  try {
    const defs = svc.service.getDefinitionAtPosition(svc.fileName, pos);
    if (!defs || defs.length === 0) return null;
    for (const def of defs) {
      if (def.fileName !== svc.fileName) continue;
      const { start, length } = def.textSpan;
      return { from: start, to: start + length };
    }
    return null;
  } catch {
    return null;
  }
}

/**
 * Same-file TYPE-definition span for "Go to Type Definition" at a buffer offset
 * (VSCode's `editor.action.goToTypeDefinition`). This is DISTINCT from
 * {@link definitionSpan}: for `const u: User = …`, go-to-definition on `u` jumps
 * to the `u` declaration, whereas go-to-TYPE-definition jumps to the `User`
 * TYPE declaration. Drives `getTypeDefinitionAtPosition` and returns the FIRST
 * result that lives in the buffer file (cross-file / synthetic-lib results are
 * skipped, EXACTLY like {@link definitionSpan}, so the caller never makes a bogus
 * in-buffer jump). Returns `null` when there is no in-file type definition.
 * Non-throwing (any compiler error degrades to `null`).
 */
export function typeDefinitionSpan(
  svc: TsIntelligenceService,
  pos: number,
): BufferSpan | null {
  try {
    const defs = svc.service.getTypeDefinitionAtPosition(svc.fileName, pos);
    if (!defs || defs.length === 0) return null;
    for (const def of defs) {
      if (def.fileName !== svc.fileName) continue;
      const { start, length } = def.textSpan;
      return { from: start, to: start + length };
    }
    return null;
  } catch {
    return null;
  }
}

/**
 * Every SAME-FILE implementation span for "Go to Implementation" at a buffer
 * offset (VSCode's `editor.action.goToImplementation`). DISTINCT from
 * {@link definitionSpan}: for an interface method / an abstract member, this
 * resolves the concrete IMPLEMENTATIONS rather than the declaration — and there
 * can be MANY (several classes implementing one interface), so unlike the
 * single-span definition/type-definition helpers this returns an ARRAY.
 *
 * Drives `getImplementationAtPosition`, keeps only locations in the buffer file
 * (rejecting cross-file / synthetic-lib results like {@link referenceSpans}),
 * maps each `textSpan` to a {@link BufferSpan}, de-duplicates identical ranges,
 * and sorts by `from` ascending for a deterministic order. Returns `[]` when
 * there is no symbol at `pos`, no in-file implementation, or the compiler throws
 * (non-throwing contract). Pure: it reads the buffer but mutates nothing.
 */
export function implementationSpans(
  svc: TsIntelligenceService,
  pos: number,
): BufferSpan[] {
  try {
    const impls = svc.service.getImplementationAtPosition(svc.fileName, pos);
    if (!impls || impls.length === 0) return [];
    const spans: BufferSpan[] = [];
    const seen = new Set<string>();
    for (const impl of impls) {
      if (impl.fileName !== svc.fileName) continue;
      const { start, length } = impl.textSpan;
      const from = start;
      const to = start + length;
      const key = `${from}:${to}`;
      if (seen.has(key)) continue;
      seen.add(key);
      spans.push({ from, to });
    }
    spans.sort((a, b) => a.from - b.from || a.to - b.to);
    return spans;
  } catch {
    return [];
  }
}

/**
 * Every SAME-FILE reference of the symbol under the caret (VSCode's
 * "Find All References" / Shift+F12).
 *
 * Uses `getReferencesAtPosition` — the SAME engine `findRenameLocations`
 * (renameEdits) drives — which returns a flat list of `{fileName, textSpan}`
 * occurrences of the symbol at `pos` (declaration + every usage). We keep only
 * the locations in the buffer file (rejecting cross-file / synthetic-lib refs
 * EXACTLY like {@link definitionSpan}), map each `textSpan` to a
 * {@link ReferenceSpan} (with its 1-based line, 0-based column and the trimmed
 * source line for a panel preview), de-duplicate identical ranges, and sort by
 * `from` ascending for a deterministic order.
 *
 * Returns `[]` when there is no symbol at `pos`, no in-file references, or the
 * compiler throws (non-throwing contract — same as its siblings). Pure: it reads
 * the buffer but mutates nothing. The line/column/preview are derived from the
 * live buffer text (`getText()` via the host) so they always match what the
 * editor renders.
 */
export function referenceSpans(
  svc: TsIntelligenceService,
  pos: number,
): ReferenceSpan[] {
  try {
    const refs = svc.service.getReferencesAtPosition(svc.fileName, pos);
    if (!refs || refs.length === 0) return [];
    // Read the buffer once through the language-service host so line/column and
    // preview text are computed against the exact text being analysed.
    const source = svc.service.getProgram()?.getSourceFile(svc.fileName)?.text;
    const text = typeof source === "string" ? source : "";
    const lineStarts = computeLineStarts(text);
    const spans: ReferenceSpan[] = [];
    const seen = new Set<string>();
    for (const ref of refs) {
      if (ref.fileName !== svc.fileName) continue;
      const { start, length } = ref.textSpan;
      const from = start;
      const to = start + length;
      // Collapse exact-duplicate ranges (the engine never normally repeats a
      // span, but a defensive de-dup keeps the panel list clean + deterministic).
      const key = `${from}:${to}`;
      if (seen.has(key)) continue;
      seen.add(key);
      const { line, column } = offsetToLineColumn(lineStarts, from);
      const preview = lineTextAt(text, lineStarts, line)
        .trim()
        .slice(0, CODE_REFERENCE_PREVIEW_MAX_CHARS);
      spans.push({ from, to, line, column, preview });
    }
    // Deterministic order: by start offset ascending (ties broken by `to`).
    spans.sort((a, b) => a.from - b.from || a.to - b.to);
    return spans;
  } catch {
    return [];
  }
}

/**
 * The full document outline (VSCode's Go-to-Symbol / Ctrl+Shift+O): every symbol
 * in the buffer flattened from the TS `NavigationTree` into an ordered list.
 *
 * `getNavigationTree` returns a nested tree whose ROOT is a synthetic node for
 * the whole file (its `text` is the file name or "<global>"); the real symbols
 * are its `childItems`. We recurse the children (NOT the root), turning each
 * node's FIRST span (`spans[0]` — a `TextSpan {start,length}`) into a
 * `{from,to}` range and recording its `depth` (0 for a direct child of the root,
 * 1 for that child's child, …). Nodes without a usable span are skipped but their
 * children are still walked (so a member nested under a span-less node keeps the
 * correct depth relative to its real ancestors). The flat list is sorted by
 * `from` ascending (document order; ties broken by `to`) so the picker lists
 * symbols top-to-bottom exactly as they appear.
 *
 * `fileName` defaults to the service's buffer file; it is exposed only so tests
 * can be explicit. Returns `[]` when there is no tree, no symbols, or the
 * compiler throws (non-throwing contract — same as its siblings). Pure.
 */
export function documentSymbols(
  svc: TsIntelligenceService,
  fileName: string = svc.fileName,
): DocumentSymbol[] {
  try {
    const tree = svc.service.getNavigationTree(fileName);
    if (!tree) return [];
    const out: DocumentSymbol[] = [];
    // Walk the ROOT's children (skip the synthetic file/module root itself). The
    // root carries the whole-file span and a name that is the file basename or
    // "<global>"/"global"; emitting it would add a meaningless top entry.
    const walk = (node: TsModule.NavigationTree, depth: number): void => {
      const span = node.spans?.[0];
      if (span) {
        const from = span.start;
        const to = span.start + span.length;
        out.push({ name: node.text, kind: node.kind, from, to, depth });
      }
      // Children deepen by one ONLY when the parent contributed a level (had a
      // span). A span-less node passes its own depth through so its children keep
      // their real nesting relative to emitted ancestors.
      const childDepth = span ? depth + 1 : depth;
      for (const child of node.childItems ?? []) {
        walk(child, childDepth);
      }
    };
    for (const child of tree.childItems ?? []) {
      walk(child, 0);
    }
    // Document order: by start offset ascending (ties broken by `to`).
    out.sort((a, b) => a.from - b.from || a.to - b.to);
    return out;
  } catch {
    return [];
  }
}

/**
 * Code Lenses (VSCode's inline "N references" annotation rendered ABOVE each
 * symbol's definition line) for the whole buffer.
 *
 * Walks the SAME `getNavigationTree` as {@link documentSymbols} (so the symbol
 * set is identical — top-level functions/classes/consts AND members like methods),
 * but resolves the reference count off each node's `nameSpan` (the identifier
 * itself), NOT its broader definition `spans[0]` (which for `function foo` points
 * at the `function` keyword, where `getReferencesAtPosition` resolves no symbol).
 * For each named node we run {@link referenceSpans} at `nameSpan.start` and count
 * the result.
 *
 * SELF-DECLARATION RULE (VSCode convention): the count EXCLUDES the symbol's own
 * declaration. `getReferencesAtPosition` returns the declaration PLUS every use,
 * so we drop the one reference span that overlaps the symbol's `nameSpan` (its
 * declaration occurrence) and count the rest. Net effect: a function called twice
 * reads "2 references"; an unused symbol reads "0 references" (not "1"). Dropping
 * by overlap (rather than blindly subtracting one) is robust to a symbol whose
 * declaration the engine does not echo back as a reference.
 *
 * One lens per symbol-definition LINE: when several symbols share a line (e.g.
 * `class C { foo() {} bar() {} }` lists `foo`/`bar` on the same line) we keep ONE
 * lens per line — the FIRST symbol in document order on that line — so the block
 * widget the renderer draws above the line is unambiguous (CM6 anchors at most one
 * useful "N references" line above a given line). The list is returned in document
 * order (by `nameOffset` ascending).
 *
 * Returns `[]` when there is no tree, no symbols, or the compiler throws
 * (non-throwing contract — same as its siblings). Pure: it reads the buffer but
 * mutates nothing. The line number is derived from the live buffer text so it
 * always matches what the editor renders.
 */
/**
 * Max number of code-lens reference queries per computation. Each query is a
 * whole-PROGRAM reference scan, and the lens extension computes over the WHOLE
 * document (not the viewport — that is what makes it render headless/in jsdom),
 * so without a cap a large file with hundreds of symbols would fire hundreds of
 * whole-program scans per (debounced) recompute and freeze the editor. We dedupe
 * to one definition per line and cap BEFORE the queries; a file with more than
 * this many definition lines shows lenses for the first MAX only.
 */
const MAX_CODE_LENS_SYMBOLS = 200;

export function codeLenses(
  svc: TsIntelligenceService,
  fileName: string = svc.fileName,
): CodeLens[] {
  try {
    const tree = svc.service.getNavigationTree(fileName);
    if (!tree) return [];
    const source = svc.service.getProgram()?.getSourceFile(fileName)?.text;
    const text = typeof source === "string" ? source : "";
    const lineStarts = computeLineStarts(text);

    // PERF: collect the symbol DEFINITION sites first WITHOUT querying references
    // (each reference query is a whole-program scan). Mirrors documentSymbols so
    // the symbol set is identical; only the metric differs.
    type LensDef = {
      name: string;
      line: number;
      nameOffset: number;
      nameLength: number;
    };
    const defs: LensDef[] = [];
    const walk = (node: TsModule.NavigationTree): void => {
      const nameSpan = node.nameSpan;
      if (nameSpan) {
        const { line } = offsetToLineColumn(lineStarts, nameSpan.start);
        defs.push({
          name: node.text,
          line,
          nameOffset: nameSpan.start,
          nameLength: nameSpan.length,
        });
      }
      for (const child of node.childItems ?? []) {
        walk(child);
      }
    };
    for (const child of tree.childItems ?? []) {
      walk(child);
    }
    // Document order, then ONE definition per LINE (the FIRST owns the line so
    // the block widget above the line is unambiguous), CAPPED — all BEFORE the
    // expensive reference queries, so the query count is O(min(lines, MAX)).
    defs.sort((a, b) => a.nameOffset - b.nameOffset);
    const seenLines = new Set<number>();
    const perLine: LensDef[] = [];
    for (const def of defs) {
      if (seenLines.has(def.line)) continue;
      seenLines.add(def.line);
      perLine.push(def);
      if (perLine.length >= MAX_CODE_LENS_SYMBOLS) break;
    }
    // Now (and only now) the bounded reference queries.
    const lenses: CodeLens[] = [];
    for (const def of perLine) {
      const refs = referenceSpans(svc, def.nameOffset);
      // SELF-DECLARATION RULE: drop the reference span overlapping the symbol's
      // own declaration (its nameSpan) and count the rest (VSCode convention).
      const declFrom = def.nameOffset;
      const declTo = def.nameOffset + def.nameLength;
      let refCount = 0;
      for (const ref of refs) {
        const overlapsDeclaration = ref.from < declTo && ref.to > declFrom;
        if (!overlapsDeclaration) refCount += 1;
      }
      lenses.push({
        name: def.name,
        line: def.line,
        nameOffset: def.nameOffset,
        refCount,
      });
    }
    return lenses;
  } catch {
    return [];
  }
}

/**
 * Offsets where each line begins (line 0 starts at offset 0). Pure scan of the
 * buffer; handles `\n`, `\r\n` and lone `\r` line breaks so line/column math
 * matches the source exactly. Used by {@link referenceSpans}.
 */
function computeLineStarts(text: string): number[] {
  const starts = [0];
  for (let i = 0; i < text.length; i += 1) {
    const ch = text.charCodeAt(i);
    if (ch === 10 /* \n */) {
      starts.push(i + 1);
    } else if (ch === 13 /* \r */) {
      // Treat \r\n as a single break (skip the \n); a lone \r still breaks.
      if (text.charCodeAt(i + 1) === 10) i += 1;
      starts.push(i + 1);
    }
  }
  return starts;
}

/**
 * Map a 0-based buffer offset to a 1-based line and 0-based column using the
 * precomputed {@link computeLineStarts} table (binary search). Total: clamps a
 * negative/oversized offset into the buffer's range.
 */
function offsetToLineColumn(
  lineStarts: number[],
  offset: number,
): { line: number; column: number } {
  const clamped = Math.max(0, offset);
  // Binary search for the greatest line start <= clamped.
  let lo = 0;
  let hi = lineStarts.length - 1;
  while (lo < hi) {
    const mid = (lo + hi + 1) >> 1;
    if (lineStarts[mid] <= clamped) lo = mid;
    else hi = mid - 1;
  }
  return { line: lo + 1, column: clamped - lineStarts[lo] };
}

/**
 * The raw text of a 1-based `line` (without its trailing line break). Reads
 * from the line-start at `line - 1` up to the next line start (or EOF), stripping
 * a trailing `\r`/`\n`. Total: an out-of-range line yields "".
 */
function lineTextAt(text: string, lineStarts: number[], line: number): string {
  const index = line - 1;
  if (index < 0 || index >= lineStarts.length) return "";
  const from = lineStarts[index];
  const to = index + 1 < lineStarts.length ? lineStarts[index + 1] : text.length;
  return text.slice(from, to).replace(/\r?\n$/, "");
}

/**
 * A legal ECMAScript / TypeScript identifier: starts with a letter, `_` or `$`,
 * followed by letters, digits, `_` or `$`, and is not a reserved word.
 *
 * Deliberately conservative (ASCII identifiers only). Reserved words are
 * rejected so a rename to e.g. `class` or `return` can't produce broken code.
 * Pure and total.
 */
const RESERVED_WORDS: ReadonlySet<string> = new Set([
  "break", "case", "catch", "class", "const", "continue", "debugger",
  "default", "delete", "do", "else", "enum", "export", "extends", "false",
  "finally", "for", "function", "if", "import", "in", "instanceof", "new",
  "null", "return", "super", "switch", "this", "throw", "true", "try",
  "typeof", "var", "void", "while", "with",
  // Strict-mode / contextual reserved words worth blocking for safety.
  "let", "static", "yield", "await", "implements", "interface", "package",
  "private", "protected", "public",
]);

const ECMASCRIPT_IDENTIFIER =
  /^[$_\p{ID_Start}][$\u200c\u200d\p{ID_Continue}]*$/u;
const ECMASCRIPT_IDENTIFIER_IN_SOURCE =
  /[$_\p{ID_Start}][$\u200c\u200d\p{ID_Continue}]*/gu;
export const MAX_CODE_IDENTIFIER_CHARS = 256;

export function identifierRangeAt(
  source: string,
  offset: number,
): { from: number; to: number } | null {
  if (typeof source !== "string" || !Number.isFinite(offset)) return null;
  const safeOffset = Math.max(0, Math.min(Math.trunc(offset), source.length));
  ECMASCRIPT_IDENTIFIER_IN_SOURCE.lastIndex = 0;
  let match: RegExpExecArray | null;
  while ((match = ECMASCRIPT_IDENTIFIER_IN_SOURCE.exec(source))) {
    const end = match.index + match[0].length;
    if (safeOffset >= match.index && safeOffset <= end) {
      return { from: match.index, to: end };
    }
    if (match.index > safeOffset) break;
  }
  return null;
}

export function identifierAt(source: string, offset: number): string {
  const range = identifierRangeAt(source, offset);
  return range ? source.slice(range.from, range.to) : "";
}

export function isLegalIdentifier(name: string): boolean {
  if (typeof name !== "string") return false;
  if (name.length === 0 || name.length > MAX_CODE_IDENTIFIER_CHARS) return false;
  if (!ECMASCRIPT_IDENTIFIER.test(name)) return false;
  if (RESERVED_WORDS.has(name)) return false;
  return true;
}

/**
 * Ordered {from,to,insert} edits renaming EVERY occurrence of the symbol under
 * the caret to `newName`, suitable for a single undoable CM6 transaction.
 *
 * Uses `findRenameLocations` (the same engine VSCode's rename uses), keeps only
 * locations within the buffer file, and returns them sorted by `from` so the
 * caller can apply them as one `changes` array. Returns `[]` when:
 *  - the new name is not a legal identifier,
 *  - there is no renameable symbol at `pos`, or
 *  - the compiler throws (non-throwing contract).
 *
 * Pure: it computes edits but does not mutate anything.
 */
export function renameEdits(
  svc: TsIntelligenceService,
  pos: number,
  newName: string,
): BufferEdit[] {
  if (!isLegalIdentifier(newName)) return [];
  try {
    const locations = svc.service.findRenameLocations(
      svc.fileName,
      pos,
      /* findInStrings */ false,
      /* findInComments */ false,
      /* providePrefixAndSuffixTextForRename */ false,
    );
    if (!locations || locations.length === 0) return [];
    const edits: BufferEdit[] = [];
    for (const loc of locations) {
      if (loc.fileName !== svc.fileName) continue;
      const { start, length } = loc.textSpan;
      edits.push({ from: start, to: start + length, insert: newName });
    }
    // Sort by position; CM6 wants non-overlapping changes and renameLocations
    // are already non-overlapping (distinct occurrences of one symbol).
    edits.sort((a, b) => a.from - b.from);
    return edits;
  } catch {
    return [];
  }
}

/**
 * Quick Fixes / code actions (VSCode's Ctrl+. / Cmd+.) for the diagnostics that
 * overlap the `[from, to]` buffer range.
 *
 * Drives the SAME engine VSCode's lightbulb uses, in three steps:
 *  1. Collect the diagnostic CODES whose range overlaps `[from, to]` from BOTH
 *     `getSemanticDiagnostics` and `getSyntacticDiagnostics` (a fix is keyed by
 *     the code of the error it repairs, so we must pass exactly the codes present
 *     in range — e.g. 2552/2551 "Did you mean", 2588 "Cannot assign to const").
 *     With NO overlapping diagnostic there is nothing to fix → `[]`.
 *  2. Call `getCodeFixesAtPosition(fileName, from, to, codes, formatOptions,
 *     preferences)` with the default format settings.
 *  3. Map each returned fix's `changes` into CM6-style {from,to,insert} edits.
 *     A fix may touch MULTIPLE files; we keep ONLY changes targeting the buffer
 *     file and REJECT (drop) the whole action if ANY change targets another file
 *     (a cross-file fix can't be applied to a single-buffer editor without
 *     corrupting it). Within the kept action the edits are sorted ascending and
 *     overlapping/adjacent-conflicting edits are dropped so the result is a
 *     valid single CM6 `changes` array (one undo step).
 *
 * De-dup: identical actions (same description + same edits) are collapsed.
 * Order is deterministic — TS's own fix order, preserved.
 *
 * NON-THROWING / TOTAL: returns `[]` when there are no diagnostics in range, no
 * fixes, an empty/garbage range, a non-TS service, or on ANY throw. (TS's
 * `getCodeFixesAtPosition` is known to throw a "Debug Failure" for some
 * fixes — e.g. the "Add missing 'new'" path under a partial lib — so the
 * try/catch is load-bearing, not merely defensive.)
 *
 * Pure: it computes actions but mutates nothing.
 */
export function codeFixes(
  svc: TsIntelligenceService,
  from: number,
  to: number,
): CodeFixAction[] {
  try {
    const ts = svc.ts;
    const fileName = svc.fileName;
    // Normalize the span: clamp into [0, ∞) and ensure start <= end so a caret
    // (from === to) and a reversed selection both probe a sane range.
    const start = Math.max(0, Math.min(from, to));
    const end = Math.max(0, Math.max(from, to));

    // Collect the codes of every diagnostic whose range OVERLAPS [start, end].
    // A zero-length diagnostic (length 0) is treated as covering its single
    // offset; a caret exactly on it (start === end === d.start) still counts.
    const semantic = svc.service.getSemanticDiagnostics(fileName) ?? [];
    const syntactic = svc.service.getSyntacticDiagnostics(fileName) ?? [];
    const codes = new Set<number>();
    for (const diag of [...semantic, ...syntactic]) {
      if (diag.start == null) continue;
      const dStart = diag.start;
      const dEnd = diag.start + (diag.length ?? 0);
      // Overlap test, inclusive at the touching edges so a caret resting at the
      // start/end of the diagnostic still surfaces its fixes.
      if (dStart <= end && dEnd >= start) {
        codes.add(diag.code);
      }
    }
    // No diagnostic in range → nothing to fix.
    if (codes.size === 0) return [];

    const formatOptions = ts.getDefaultFormatCodeSettings("\n");
    const rawFixes = svc.service.getCodeFixesAtPosition(
      fileName,
      start,
      end,
      Array.from(codes),
      formatOptions,
      {},
    );
    if (!rawFixes || rawFixes.length === 0) return [];

    const actions: CodeFixAction[] = [];
    const seen = new Set<string>();
    for (const fix of rawFixes) {
      // Reject the action outright if ANY of its file changes targets a file
      // other than the buffer — a cross-file fix is not applicable here and
      // applying only its in-buffer half would corrupt the edit.
      const changes = fix.changes ?? [];
      if (changes.some((c) => c.fileName !== fileName)) continue;

      const edits: BufferEdit[] = [];
      for (const change of changes) {
        for (const tc of change.textChanges) {
          const editFrom = tc.span.start;
          const editTo = tc.span.start + tc.span.length;
          edits.push({ from: editFrom, to: editTo, insert: tc.newText });
        }
      }
      // A fix with no concrete edits (e.g. a pure command) can't be applied as a
      // text transaction — skip it so the menu only lists actionable fixes.
      if (edits.length === 0) continue;

      // Sort ascending and drop any edit that overlaps the previous one so the
      // result is a strictly non-overlapping `changes` array CM6 can apply in a
      // single transaction. A zero-length insertion that merely touches the
      // previous edit's end (editFrom === lastTo) is NOT an overlap.
      edits.sort((a, b) => a.from - b.from || a.to - b.to);
      const safeEdits: BufferEdit[] = [];
      let lastTo = -1;
      for (const edit of edits) {
        if (edit.from < lastTo) continue; // overlaps the previous edit — drop it
        safeEdits.push(edit);
        lastTo = edit.to;
      }
      if (safeEdits.length === 0) continue;

      // De-dup identical actions (same label + identical edit set).
      const key = `${fix.description} ${safeEdits
        .map((e) => `${e.from}:${e.to}:${e.insert}`)
        .join("|")}`;
      if (seen.has(key)) continue;
      seen.add(key);

      actions.push({ description: fix.description, edits: safeEdits });
    }
    return actions;
  } catch {
    return [];
  }
}

/**
 * Organize Imports (VSCode "Organize Imports", Shift+Alt+O) for the TS/JS buffer:
 * the TS language service's own organize-imports pass — it removes unused
 * imports, merges duplicate import statements from the same module, and sorts the
 * remaining imports. Returns the resulting buffer edits (in-buffer only; a change
 * targeting another file is rejected, as the editor is single-buffer), sorted +
 * de-overlapped so CM6 can apply them in one transaction. Empty (a no-op) when
 * imports are already organized or the service throws. PURE w.r.t. the editor —
 * it only reads the language service and returns edits the caller dispatches.
 */
export function organizeImports(svc: TsIntelligenceService): BufferEdit[] {
  try {
    const ts = svc.ts;
    const fileName = svc.fileName;
    const formatOptions = ts.getDefaultFormatCodeSettings("\n");
    const changes = svc.service.organizeImports(
      { type: "file", fileName },
      formatOptions,
      {},
    );
    if (!changes || changes.length === 0) return [];
    const edits: BufferEdit[] = [];
    for (const change of changes) {
      // Single-buffer editor: a change to any other file isn't applicable.
      if (change.fileName !== fileName) continue;
      for (const tc of change.textChanges) {
        edits.push({
          from: tc.span.start,
          to: tc.span.start + tc.span.length,
          insert: tc.newText,
        });
      }
    }
    edits.sort((a, b) => a.from - b.from || a.to - b.to);
    const safe: BufferEdit[] = [];
    let lastTo = -1;
    for (const edit of edits) {
      if (edit.from < lastTo) continue; // overlaps the previous edit — drop it
      safe.push(edit);
      lastTo = edit.to;
    }
    return safe;
  } catch {
    return [];
  }
}
