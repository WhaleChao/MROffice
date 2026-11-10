/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * CodeMirror 6 autocomplete for the MRNote code editor.
 *
 * Ports the project's existing, framework-free completion logic
 * ({@link getCompletions} / {@link applyCompletion} in `autocomplete.ts`) into a
 * CM6 {@link CompletionSource}, so the new `CodeMirrorEditor` offers the same
 * language-keyword + document-identifier completions the textarea editor did.
 *
 * The source itself is the only CM-specific glue; the ranking/filtering still
 * lives in `autocomplete.ts` and stays unit-tested there. The builder here is a
 * thin, pure function (CM `CompletionContext` -> `CompletionResult | null`) so it
 * can be exercised in node with a hand-rolled context (no EditorView needed).
 */
import {
  acceptCompletion,
  autocompletion,
  closeCompletion,
  completionKeymap,
  completionStatus,
  snippet as cmSnippet,
  type Completion,
  type CompletionContext,
  type CompletionResult,
  type CompletionSource,
} from "@codemirror/autocomplete";
import { Prec, type Extension } from "@codemirror/state";
import { EditorView, keymap, type KeyBinding } from "@codemirror/view";

import { getCompletions } from "./autocomplete";
import {
  completionEntryDetail,
  completionsAt,
  intelligenceLanguage,
  type TsIntelligenceService,
} from "./code-intelligence";
import {
  ensureFinalTabStop,
  toCmSnippetTemplate,
  type Snippet,
} from "./code-snippets";
import type { LanguageServiceProvider } from "./language-service-provider";
import type {
  WorkspaceCompletionEntry,
  WorkspaceCompletionResult,
} from "./workspace-intelligence";

/**
 * Build the `apply` callback for a snippet completion using CodeMirror 6's native
 * `snippet()` engine, so accepting a snippet starts an INTERACTIVE session with
 * ordered tab-stops (Tab / Shift-Tab to cycle), pre-selected placeholders
 * (`${1:label}`) and mirrors (a repeated index updates together).
 *
 * The body is first normalized: {@link ensureFinalTabStop} appends a final `$0`
 * when the body declares none — preserving the legacy "caret at end" behaviour
 * for plain bodies (CM otherwise leaves the caret at the first field) — and
 * {@link toCmSnippetTemplate} rewrites the body's `$0`/`$1`/`${1:label}` tokens
 * into CM6's braced template form while leaving all other text (incl. literal
 * `{`/`}`) untouched. For a single-`$0` body this inserts byte-identical text and
 * lands the caret in exactly the same place as the old `expandSnippet` path; it
 * only ADDS interactivity for bodies that use numbered stops/placeholders.
 *
 * CM6's snippet apply re-indents continuation lines by one indent unit per
 * LEADING TAB after a newline; our bodies author continuation indentation with
 * literal spaces (no leading tabs), so no extra indentation is injected and the
 * inserted text matches what the editor inserted before.
 */
function snippetApply(
  snippet: Snippet,
): (view: EditorView, completion: Completion, from: number, to: number) => void {
  const template = toCmSnippetTemplate(ensureFinalTabStop(snippet.body));
  // `snippet()` returns an apply fn whose first arg is `{state, dispatch}` — the
  // EditorView satisfies that shape, so the Completion.apply view is passed
  // straight through. It dispatches its own (completion-annotated) transaction
  // and installs the tab-stop fields.
  return cmSnippet(template);
}

/**
 * Map matching snippets to CM {@link Completion} options. `type: "snippet"` so
 * CM tags them (icon/class) and they read as snippets next to identifier/keyword
 * options. ADDITIVE — appended after the identifier/keyword options.
 */
export function snippetCompletions(snippets: Snippet[]): Completion[] {
  return snippets.map((snippet) => ({
    label: snippet.prefix,
    type: "snippet",
    detail: snippet.description,
    apply: snippetApply(snippet),
    // Keep snippets visible even when CM re-filters within the same word; the
    // prefix already matched in getCompletions.
    boost: -1,
  }));
}

const safeCompletionKeymap: readonly KeyBinding[] = [
  {
    key: "Enter",
    run: (view) => {
      if (closeCompletion(view)) return false;
      return false;
    },
  },
  ...completionKeymap.filter((binding) => binding.key !== "Enter"),
  { key: "Tab", run: acceptCompletion },
];

function shouldDismissCompletionForPlainInput(event: KeyboardEvent): boolean {
  if (event.isComposing || event.metaKey || event.ctrlKey || event.altKey) {
    return false;
  }
  if (event.key === "Enter" || event.key === " ") return true;
  return event.key.length === 1 && /^[()[\]{};,]$/.test(event.key);
}

const dismissCompletionOnBoundaryInput = EditorView.domEventHandlers({
  keydown: (event, view) => {
    if (
      completionStatus(view.state) === "active" &&
      shouldDismissCompletionForPlainInput(event)
    ) {
      closeCompletion(view);
    }
    return false;
  },
});

/**
 * Build a CM6 {@link CompletionResult} for a caret position by delegating to the
 * shared {@link getCompletions}. Returns `null` when there is no word to
 * complete (CM closes the popup), matching the textarea editor's behaviour.
 *
 * `explicit` (Ctrl/Cmd-Space) is accepted but not required: completions surface
 * automatically while typing a word, same as before.
 */
export function buildCompletionResult(
  docText: string,
  caret: number,
  language: string,
  extraIdentifiers?: Iterable<string>,
): CompletionResult | null {
  const ctx = getCompletions(docText, caret, language, 12, extraIdentifiers);
  if (!ctx) return null;
  const options: Completion[] = ctx.items.map((label) => ({
    label,
    type: "keyword",
  }));
  // Append matching snippets as additional options (additive: identifier/keyword
  // options are unchanged and still come first).
  options.push(...snippetCompletions(ctx.snippets));
  return {
    // `from` is where the typed word starts; CM replaces [from, caret) with the
    // accepted label — the same replacement window applyCompletion used.
    from: ctx.start,
    to: caret,
    options,
    // Our list is already filtered/ranked by getCompletions; let CM re-filter as
    // the user keeps typing within the same word (validFor) without re-running.
    validFor: /^[\p{ID_Continue}$]*$/u,
  };
}

/**
 * A {@link CompletionSource} bound to a language. Reads the live document + caret
 * from the CM context and produces the same suggestions as the textarea editor.
 *
 * `extraIdentifiers`, when supplied, is invoked per request to gather identifiers
 * from OUTSIDE this buffer (e.g. sibling multi-buffer host cells in a shared scope) so the
 * completion list spans the whole project. It is lazy so a host that has nothing
 * to add costs nothing.
 */
export function projectCompletionSource(
  language: string,
  extraIdentifiers?: () => Iterable<string> | undefined,
): CompletionSource {
  return (context: CompletionContext): CompletionResult | null => {
    const docText = context.state.doc.toString();
    const caret = context.pos;
    const result = buildCompletionResult(
      docText,
      caret,
      language,
      extraIdentifiers?.(),
    );
    // When typing produces no word (e.g. just after a space) bail unless the
    // completion was explicitly requested, in which case still return null so CM
    // shows "no completions" rather than a stale list.
    if (!result && !context.explicit) return null;
    return result;
  };
}

/**
 * Maximum source window scanned for identifier completions while the editor is
 * in automatic large-file mode. Most of the window sits before the caret, where
 * local declarations normally live; a smaller suffix keeps nearby later symbols
 * available without rescanning a multi-megabyte document on every typed letter.
 */
export const LIGHTWEIGHT_COMPLETION_SCAN_CHARACTERS = 64 * 1024;
/** Automatic multi-buffer host completion settle delay; superseded keystrokes abort it. */
export const PROJECT_COMPLETION_SETTLE_MS = 80;
/** Defensive UI ceiling even when a project provider violates its own budget. */
export const PROJECT_COMPLETION_UI_ITEM_LIMIT = 200;

/** Versioned project completion result accepted by the CodeMirror adapter. */
export interface ProjectSemanticCompletionResult
  extends WorkspaceCompletionResult {
  requestVersion: number;
}

/** Async bounded-project completion hook supplied by a multi-buffer host/workspace host. */
export type ProjectSemanticCompletionProvider = (
  offsetInBuffer: number,
  requestVersion: number,
  signal: AbortSignal,
) => Promise<ProjectSemanticCompletionResult | null>;

/**
 * Adapt one versioned project completion provider to CodeMirror.
 *
 * Automatic requests settle briefly so rapid typing coalesces. Any document
 * change aborts both the delay and provider signal; mismatched request versions,
 * invalid replacement ranges, thrown providers, and empty/oversized results
 * fail closed.
 */
export function projectSemanticCompletionSource(
  provider: ProjectSemanticCompletionProvider,
  settleMs = PROJECT_COMPLETION_SETTLE_MS,
): CompletionSource {
  let nextRequestVersion = 0;
  const boundedSettleMs = Number.isFinite(settleMs)
    ? Math.max(0, Math.floor(settleMs))
    : PROJECT_COMPLETION_SETTLE_MS;
  return async (
    context: CompletionContext,
  ): Promise<CompletionResult | null> => {
    const word = context.matchBefore(/[\p{ID_Continue}$]*$/u);
    const fragmentStart = word ? word.from : context.pos;
    const afterDot =
      context.state.doc.sliceString(
        Math.max(0, fragmentStart - 1),
        fragmentStart,
      ) === ".";
    if (
      !context.explicit &&
      !afterDot &&
      (!word || word.from === word.to)
    ) {
      return null;
    }

    const requestVersion = (nextRequestVersion += 1);
    const controller = new AbortController();
    context.addEventListener(
      "abort",
      () => controller.abort(),
      { onDocChange: true },
    );
    if (boundedSettleMs > 0) {
      await new Promise<void>((resolve) => {
        const timer = setTimeout(resolve, boundedSettleMs);
        context.addEventListener(
          "abort",
          () => {
            clearTimeout(timer);
            resolve();
          },
          { onDocChange: true },
        );
      });
    }
    if (context.aborted || controller.signal.aborted) return null;

    let result: ProjectSemanticCompletionResult | null;
    try {
      result = await provider(
        context.pos,
        requestVersion,
        controller.signal,
      );
    } catch {
      return null;
    }
    if (
      context.aborted ||
      controller.signal.aborted ||
      !result ||
      result.requestVersion !== requestVersion ||
      !Number.isSafeInteger(result.from) ||
      !Number.isSafeInteger(result.to) ||
      result.from < 0 ||
      result.to < result.from ||
      result.to !== context.pos ||
      result.to > context.state.doc.length
    ) {
      return null;
    }
    const entries: WorkspaceCompletionEntry[] = result.entries.slice(
      0,
      PROJECT_COMPLETION_UI_ITEM_LIMIT,
    );
    if (entries.length === 0) return null;
    return {
      from: result.from,
      to: result.to,
      options: entries.map((entry) => ({
        label: entry.name,
        type: entry.type,
        boost: 2,
      })),
      validFor: /^[\p{ID_Continue}$]*$/u,
    };
  };
}

/**
 * Local keyword / identifier completion over a bounded window around the caret.
 *
 * This keeps completion available in large files while avoiding the full-source
 * `matchAll` pass used by the rich profile. Returned offsets are translated back
 * into document coordinates before CodeMirror applies a choice.
 */
export function lightweightProjectCompletionSource(
  language: string,
  maxScanCharacters = LIGHTWEIGHT_COMPLETION_SCAN_CHARACTERS,
): CompletionSource {
  const boundedMax = Number.isFinite(maxScanCharacters)
    ? Math.max(1_024, Math.floor(maxScanCharacters))
    : LIGHTWEIGHT_COMPLETION_SCAN_CHARACTERS;
  return (context: CompletionContext): CompletionResult | null => {
    const beforeCharacters = Math.floor((boundedMax * 3) / 4);
    const afterCharacters = boundedMax - beforeCharacters;
    const from = Math.max(0, context.pos - beforeCharacters);
    const to = Math.min(context.state.doc.length, context.pos + afterCharacters);
    const sourceWindow = context.state.doc.sliceString(from, to);
    const localCaret = context.pos - from;
    const result = buildCompletionResult(
      sourceWindow,
      localCaret,
      language,
    );
    if (!result) return null;
    return {
      ...result,
      from: result.from + from,
      to: result.to === undefined ? undefined : result.to + from,
    };
  };
}

/** jedi's completion kinds → CM6 icon/type names. */
function cmTypeForJediKind(kind: string | null): string {
  switch (kind) {
    case "function":
      return "function";
    case "class":
      return "class";
    case "module":
      return "namespace";
    case "keyword":
      return "keyword";
    case "property":
      return "property";
    default:
      return "variable";
  }
}

/**
 * Semantic Python source: jedi over the page's LIVE kernel namespace (via
 * `completePython`). Activates only once a kernel exists for the namespace
 * (i.e. the user ran the page at least once) — never boots Pyodide as a
 * typing side effect. Triggers on an identifier fragment or right after a
 * `.` (member access — the case the static keyword source cannot serve).
 * Resolves null on timeout/unavailability so the sync source still applies.
 */
export function pythonSemanticSource(
  namespace: string,
  /**
   * Optional cross-CELL context: given the caret's 1-based line + 0-based column
   * in THIS cell, returns the concatenated multi-buffer host module text plus the caret's
   * position within it — so jedi resolves completions against the whole shared
   * scope (a name defined in another cell is suggested), exactly the module view
   * hover / go-to-definition already use. Absent (or returning null) → this cell's
   * buffer alone, the unchanged single-file behaviour.
   */
  crossCellComplete?: (
    localLine: number,
    localColumn: number,
  ) => { text: string; line: number; column: number } | null,
): CompletionSource {
  return async (context: CompletionContext): Promise<CompletionResult | null> => {
    const { hasPythonKernel, completePython } = await import(
      "@mrnote/core/web/lib/code-runner/python-runner.ts"
    );
    if (!hasPythonKernel(namespace)) return null;
    const word = context.matchBefore(/[\p{ID_Continue}$]*$/u);
    const fragmentStart = word ? word.from : context.pos;
    const charBefore = context.state.doc.sliceString(
      Math.max(0, fragmentStart - 1),
      fragmentStart,
    );
    const afterDot = charBefore === ".";
    // Same gating as the sync source, PLUS the member-access case: complete
    // bare `obj.` (empty fragment after a dot) where identifier sources bail.
    if (!context.explicit && !afterDot && (!word || word.from === word.to)) {
      return null;
    }
    const line = context.state.doc.lineAt(context.pos);
    const localColumn = context.pos - line.from;
    // Resolve against the whole multi-buffer host's shared scope when a cross-cell context
    // is supplied; otherwise this cell's own buffer (single-file editor).
    const cross = crossCellComplete?.(line.number, localColumn) ?? null;
    const items = await completePython(
      cross ? cross.text : context.state.doc.toString(),
      cross ? cross.line : line.number,
      cross ? cross.column : localColumn,
      { namespace },
    );
    if (context.aborted || items.length === 0) return null;
    return {
      from: fragmentStart,
      options: items.map((item) => ({
        label: item.name,
        type: cmTypeForJediKind(item.kind),
        // Jedi results rank above the static keyword/identifier options.
        boost: 1,
      })),
      validFor: /^[\p{ID_Continue}$]*$/u,
    };
  };
}

/**
 * Semantic source for every language served by a REGISTERED provider rather
 * than by one of the two engines wired in by name above.
 *
 * The trigger rule is deliberately the same one the jedi source uses — an
 * identifier fragment, right after a `.`, or an explicit request — so a new
 * language behaves like the ones already here instead of popping up at moments
 * the user has learned not to expect.
 */
export function providerCompletionSource(
  language: string,
  providerFor: (
    language: string,
    capability: "completions",
  ) => LanguageServiceProvider | null,
): CompletionSource {
  return async (context: CompletionContext): Promise<CompletionResult | null> => {
    const provider = providerFor(language, "completions");
    if (!provider?.completions) return null;
    const word = context.matchBefore(/[\p{ID_Continue}$]*$/u);
    const fragmentStart = word ? word.from : context.pos;
    const charBefore = context.state.doc.sliceString(
      Math.max(0, fragmentStart - 1),
      fragmentStart,
    );
    const afterDot = charBefore === ".";
    if (!context.explicit && !afterDot && (!word || word.from === word.to)) {
      return null;
    }
    let items;
    try {
      items = await provider.completions({
        text: context.state.doc.toString(),
        offset: context.pos,
        language,
      });
    } catch {
      // A contributed provider that throws degrades to no suggestions, never to
      // a broken completion popup.
      return null;
    }
    if (context.aborted || items.length === 0) return null;
    return {
      from: fragmentStart,
      options: items.map((item) => ({
        label: item.label,
        ...(item.kind ? { type: item.kind } : {}),
        ...(item.detail ? { detail: item.detail } : {}),
        // Semantic results rank above the static keyword/identifier options,
        // matching how the jedi and TypeScript sources are ordered.
        boost: 1,
      })),
      validFor: /^[\p{ID_Continue}$]*$/u,
    };
  };
}

/**
 * Semantic completion source (TS/JS) — true IntelliSense, the primary source for
 * the JS/TS family. Asks the real {@link TsModule.LanguageService}
 * (`getCompletionsAtPosition`) for what's in scope: typed in-scope symbols, lib
 * globals, keywords, AND member access (`arr.` -> map/filter/length, typed to the
 * receiver — the thing identifier scraping fundamentally cannot do).
 *
 * It is a STRICT SUPERSET, so it can only ADD completions, never drop any: for a
 * bare identifier (not a `.` member) it also folds in the identifier-scraping
 * source's snippets + cross-cell identifiers + any scraped tokens the compiler
 * didn't surface, deduped by label. When no service is available or the compiler
 * yields nothing, it falls back to pure identifier scraping. (Members never mix
 * in scraped tokens — those aren't members of the receiver.)
 */
export function tsCompletionSource(
  language: string,
  getService: () => Promise<TsIntelligenceService | null>,
  extraIdentifiers?: () => Iterable<string> | undefined,
  projectProvider?: ProjectSemanticCompletionProvider,
  projectSettleMs = PROJECT_COMPLETION_SETTLE_MS,
): CompletionSource {
  const projectSource = projectProvider
    ? projectSemanticCompletionSource(projectProvider, projectSettleMs)
    : null;
  return async (context: CompletionContext): Promise<CompletionResult | null> => {
    if (!intelligenceLanguage(language)) return null;
    const docText = context.state.doc.toString();
    // Whether the typed fragment directly follows a `.` (real member access).
    // This — NOT TS's isMemberCompletion (which is also true for a bare member
    // NAME inside a class/interface/type-literal body) — decides whether to fold
    // in the scraped identifier source: after a dot we want members only; for any
    // bare word (incl. a class body) we still offer in-scope + cross-cell tokens.
    const wordBefore = context.matchBefore(/[\p{ID_Continue}$]*$/u);
    const fragStart = wordBefore ? wordBefore.from : context.pos;
    const afterDot =
      context.state.doc.sliceString(Math.max(0, fragStart - 1), fragStart) === ".";
    if (
      !context.explicit &&
      !afterDot &&
      (!wordBefore || wordBefore.from === wordBefore.to)
    ) {
      return null;
    }
    if (projectSource) {
      const projectResult = await projectSource(context);
      if (context.aborted) return null;
      if (projectResult) {
        const options = [...projectResult.options];
        if (!afterDot) {
          const seen = new Set(options.map((option) => option.label));
          const scraped = buildCompletionResult(
            docText,
            context.pos,
            language,
            extraIdentifiers?.(),
          );
          if (scraped && scraped.from === projectResult.from) {
            for (const option of scraped.options) {
              if (!seen.has(option.label)) {
                seen.add(option.label);
                options.push(option);
              }
            }
          }
        }
        return { ...projectResult, options };
      }
    }
    const svc = await getService();
    if (context.aborted) return null;
    if (svc) {
      svc.markDirty();
      const sem = completionsAt(svc, context.pos);
      if (context.aborted) return null;
      if (sem && sem.entries.length > 0) {
        const seen = new Set(sem.entries.map((e) => e.name));
        // The position the completion was requested at — captured for the lazy
        // detail fetch below (close enough: the list is stable within a word).
        const detailPos = context.pos;
        const options: Completion[] = sem.entries.map((e) => ({
          label: e.name,
          type: e.type,
          // Members/typed symbols rank above stray keyword/snippet overlap.
          boost: 1,
          // VSCode's details panel: when this entry is highlighted, fetch its
          // type signature + docs lazily (the per-entry second call) and render
          // them as plain text (textContent — no markup). Best-effort -> null.
          info: () => {
            const d = completionEntryDetail(svc, detailPos, e.name);
            if (!d || (!d.detail && !d.documentation)) return null;
            const dom = document.createElement("div");
            dom.className = "cm-mrnote-completion-info";
            dom.style.cssText =
              "padding:4px 6px;max-width:420px;max-height:260px;overflow:auto;font-family:var(--font-jetbrains,ui-monospace,SFMono-Regular,Menlo,monospace);font-size:12px;line-height:1.4;white-space:pre-wrap;";
            if (d.detail) {
              const sig = document.createElement("div");
              sig.style.cssText = "font-weight:600;";
              sig.textContent = d.detail;
              dom.appendChild(sig);
            }
            if (d.documentation) {
              const doc = document.createElement("div");
              doc.style.cssText = "margin-top:4px;opacity:0.85;";
              doc.textContent = d.documentation;
              dom.appendChild(doc);
            }
            return dom;
          },
        }));
        // For a BARE word (anything NOT directly after a `.`), fold in the
        // scraping source's snippets + cross-cell identifiers + any token the
        // compiler missed — deduped, and only when the replacement window lines
        // up. Gating on `afterDot` (not sem.isMember) keeps in-scope identifiers
        // available inside class/interface/type-literal bodies, where TS reports
        // isMember=true for a bare member name.
        if (!afterDot) {
          const scraped = buildCompletionResult(
            docText,
            context.pos,
            language,
            extraIdentifiers?.(),
          );
          if (scraped && scraped.from === sem.from) {
            for (const opt of scraped.options) {
              if (!seen.has(opt.label)) {
                seen.add(opt.label);
                options.push(opt);
              }
            }
          }
        }
        return { from: sem.from, options, validFor: /^[\p{ID_Continue}$]*$/u };
      }
    }
    // No service / nothing semantic: pure identifier scraping (unchanged path).
    const scraped = buildCompletionResult(
      docText,
      context.pos,
      language,
      extraIdentifiers?.(),
    );
    if (!scraped && !context.explicit) return null;
    return scraped;
  };
}

/**
 * Wrap a native completion source behind a retryable module loader.
 *
 * CSS and HTML grammar packages contain parser tables that should not become a
 * static dependency of every Code editor. The wrapper is deliberately stable:
 * concurrent queries share one import, a failed import is retried by the next
 * query, and an obsolete completion context can neither invoke nor publish a
 * late native result. Registering `onDocChange` makes source replacement a
 * cancellation boundary even when CodeMirror could otherwise reuse `validFor`.
 */
export function lazyNativeCompletionSource(
  load: () => Promise<CompletionSource>,
): CompletionSource {
  let loaded: CompletionSource | null = null;
  let inFlight: Promise<CompletionSource | null> | null = null;

  const loadSource = (): Promise<CompletionSource | null> => {
    if (loaded) return Promise.resolve(loaded);
    if (!inFlight) {
      inFlight = Promise.resolve()
        .then(load)
        .then((source) => {
          if (typeof source !== "function") return null;
          loaded = source;
          return source;
        })
        .catch(() => null)
        .finally(() => {
          inFlight = null;
        });
    }
    return inFlight;
  };

  return async (context) => {
    let retired = context.aborted;
    context.addEventListener(
      "abort",
      () => {
        retired = true;
      },
      { onDocChange: true },
    );
    if (retired || context.aborted) return null;

    const source = await loadSource();
    if (!source || retired || context.aborted) return null;
    try {
      const result = await source(context);
      return retired || context.aborted ? null : result;
    } catch {
      return null;
    }
  };
}

const lazyCssCompletionSource = lazyNativeCompletionSource(async () => {
  const grammarModule = await import("@codemirror/lang-css");
  return grammarModule.cssCompletionSource;
});

const lazyHtmlCompletionSource = lazyNativeCompletionSource(async () => {
  const grammarModule = await import("@codemirror/lang-html");
  return grammarModule.htmlCompletionSource;
});

/**
 * Grammar-aware sources a language's own CodeMirror package ships. Because
 * {@link completionExtension} takes `override`, these are switched off unless
 * this module puts them back — which is why CSS and HTML would otherwise fall
 * back to scraping words out of the document. The stable wrappers defer the
 * grammar packages until a matching completion query actually runs.
 */
export function nativeLanguageCompletionSources(
  language: string,
): CompletionSource[] {
  const lang = (language || "").trim().toLowerCase();
  if (lang === "css" || lang === "scss" || lang === "less") {
    return [lazyCssCompletionSource];
  }
  if (lang === "html" || lang === "xml" || lang === "htm") {
    return [lazyHtmlCompletionSource];
  }
  return [];
}

/**
 * The `autocompletion()` extension wired with our language-aware source. It
 * owns the completion keymap instead of relying on CodeMirror defaults because
 * the default Enter binding accepts the highlighted suggestion. That is hostile
 * to real typing, paste, and AI-generated multi-line edits: a newline can mutate
 * user text into a semantic completion or snippet. MRNote keeps Enter as plain
 * editing input; accepting a suggestion is an explicit Tab/Ctrl-Space flow.
 */
export function completionExtension(
  language: string,
  pythonNamespace?: string,
  extraIdentifiers?: () => Iterable<string> | undefined,
  tsService?: () => Promise<TsIntelligenceService | null>,
  pythonCrossCellComplete?: (
    localLine: number,
    localColumn: number,
  ) => { text: string; line: number; column: number } | null,
  options?: {
    /**
     * Keep autocomplete, but use the bounded local source and skip semantic
     * whole-document engines while automatic large-file mode is active.
     */
    lightweight?: boolean;
    /** Primary bounded-project JS/TS completion provider (e.g. multi-buffer host). */
    projectCompletion?: ProjectSemanticCompletionProvider;
    /** Override only for deterministic tests; production uses the default. */
    projectCompletionSettleMs?: number;
    /**
     * Lookup into the editor's language-service registry. Supplied, a language
     * that is neither TS/JS nor Python-with-a-kernel gets whatever provider was
     * registered for it — which is the whole of "add a language". Absent, the
     * behaviour is exactly what it was before the registry existed.
     */
    providerFor?: (
      language: string,
      capability: "completions",
    ) => LanguageServiceProvider | null;
  },
): Extension {
  const lightweight = options?.lightweight === true;
  // JS/TS with a live service: ONE unified semantic source (a strict superset of
  // scraping). Otherwise the pure identifier/keyword source.
  const sources: CompletionSource[] =
    !lightweight && tsService && intelligenceLanguage(language)
      ? [
          tsCompletionSource(
            language,
            tsService,
            extraIdentifiers,
            options?.projectCompletion,
            options?.projectCompletionSettleMs,
          ),
        ]
      : lightweight
        ? [lightweightProjectCompletionSource(language)]
        : [projectCompletionSource(language, extraIdentifiers)];
  const lang = (language || "").trim().toLowerCase();
  const pythonWired =
    Boolean(pythonNamespace) && (lang === "python" || lang === "py");
  if (!lightweight && pythonNamespace && pythonWired) {
    sources.push(pythonSemanticSource(pythonNamespace, pythonCrossCellComplete));
  }
  // Every other language: whatever provider was registered for it. TS/JS and
  // wired Python are answered above by engines that do more than the interface
  // carries, so they never reach here and their behaviour is untouched.
  if (
    !lightweight &&
    options?.providerFor &&
    !intelligenceLanguage(language) &&
    !pythonWired
  ) {
    sources.push(providerCompletionSource(language, options.providerFor));
  }
  // Properties and their value sets, tags, attributes, and the nested contexts
  // inside an HTML document. The identifier scraper still contributes the
  // document's own class and id names, so both kinds of suggestion appear
  // together. Skipped in lightweight mode with every other semantic engine.
  if (!lightweight) {
    sources.push(...nativeLanguageCompletionSources(language));
  }
  return [
    autocompletion({
      override: sources,
      // Match the textarea editor: complete identifiers/keywords, not on every key.
      activateOnTyping: true,
      defaultKeymap: false,
      icons: false,
    }),
    dismissCompletionOnBoundaryInput,
    Prec.highest(keymap.of(safeCompletionKeymap)),
  ];
}
