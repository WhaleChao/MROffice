/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Python CALL-SITE PARAMETER-NAME inlay hints (VSCode parity, the
 * `editor.inlayHints` "parameter names" subset) for the MRNote CodeMirror 6
 * editor.
 *
 * VSCode draws a faint `name:` chip before each POSITIONAL argument of a call so
 * you can read `slice(start: 2, end: 5)` without opening the signature popup.
 * This module brings that to Python WITHOUT any new dependency or new worker
 * round-trip: the parameter NAMES come straight from the jedi data the editor
 * already fetches for parameter hints (`signaturePython` → `get_signatures()`
 * `.params`), and the placement is pure, deterministic source math — NOT type
 * inference. We never invent a type or a name jedi did not give us.
 *
 * Split into two layers, the first exhaustively unit-tested without a DOM:
 *  - {@link computePythonParameterInlayHints} — the PURE core: given a call's
 *    argument-list source text and the ordered parameter names, return one
 *    `{ offset, label }` hint per positional argument (keyword args, `*args` /
 *    `**kwargs` spreads and surplus arguments produce none). Framework-free and
 *    total.
 *  - {@link pythonInlayHintsExtension} — the CM6 wiring: a browser-/kernel-gated
 *    ViewPlugin that, for the call sites in the VISIBLE range, resolves parameter
 *    names via the existing jedi signature round-trip and renders the hints as
 *    inline atomic widgets. Inert (no decoration, byte-for-byte the pre-feature
 *    editor) for a non-Python buffer, a buffer with no live kernel, or a buffer
 *    with no calls.
 *
 * HONESTLY-SCOPED — what is intentionally NOT done here (no fakery):
 *  - TYPE inlay hints (inferred variable types, `-> ReturnType` chips). jedi can
 *    sometimes infer these, but only via a separate per-symbol inference round
 *    this editor does not perform; rather than fabricate types we omit them. Only
 *    parameter-NAME hints — which jedi hands us deterministically — are rendered.
 *  - Batching. Each visible call site costs one `signaturePython` query (the same
 *    query the parameter-hint tooltip already issues for the active call). The
 *    extension bounds this by scanning ONLY the viewport and capping the number
 *    of resolved calls per refresh ({@link MAX_INLAY_CALLS_PER_REFRESH}); a single
 *    batched jedi op would need a new worker round-trip not available here.
 */
import {
  Decoration,
  EditorView,
  ViewPlugin,
  WidgetType,
  type DecorationSet,
  type ViewUpdate,
} from "@codemirror/view";
import { StateEffect, StateField, type Extension } from "@codemirror/state";

import { computeSkipMask } from "./editor-decorations";

/**
 * One inlay hint: the buffer-relative `offset` the `label` chip is anchored
 * BEFORE (the first char of the positional argument it names) and the rendered
 * `label` text (the bare parameter name plus a trailing colon, e.g. `start:`).
 */
export interface InlayHint {
  /** 0-based offset (in the supplied text's coordinate space) to anchor before. */
  offset: number;
  /** The chip text, e.g. `start:` — the parameter name with a trailing colon. */
  label: string;
}

/** A bracket flavour for the nesting scan inside an argument list. */
const ARG_OPENERS = new Set(["(", "[", "{"]);
const ARG_CLOSERS = new Set([")", "]", "}"]);

/**
 * Whether a parameter name is one we will render a chip for. jedi's `.params`
 * includes the synthetic `self` / `cls` of a bound method and may surface the
 * bare `*` / `**` markers or `*args` / `**kwargs` collectors — none of which name
 * a positional slot the caller fills by position in a way a chip clarifies, so we
 * skip them. An empty name (jedi failed to read one) is also skipped.
 */
function isNameableParam(name: string): boolean {
  if (!name) return false;
  if (name === "self" || name === "cls") return false;
  if (name.startsWith("*")) return false; // `*`, `*args`, `**kwargs`
  return true;
}

/**
 * Split a call's argument-list text into TOP-LEVEL arguments, returning for each
 * the offset of its first non-whitespace char and that argument's raw text. Commas
 * inside nested brackets or string literals do NOT split (so `f(g(a, b), c)` is two
 * args, not three). Uses {@link computeSkipMask} so a comma inside a string is
 * ignored. A trailing empty argument (a dangling comma, e.g. `f(a, )`) is dropped.
 */
function splitTopLevelArgs(
  argsText: string,
): { start: number; text: string }[] {
  const args: { start: number; text: string }[] = [];
  const mask = computeSkipMask(argsText);
  let depth = 0;
  let segStart = 0;
  const pushSegment = (from: number, to: number) => {
    const raw = argsText.slice(from, to);
    // Anchor offset = first non-whitespace char of the segment (or `from` for an
    // all-whitespace segment, which `isPositional` will then reject).
    let i = from;
    while (i < to && /\s/.test(argsText[i])) i += 1;
    args.push({ start: i, text: raw });
  };
  for (let i = 0; i < argsText.length; i += 1) {
    if (mask[i]) continue; // inside a string/comment — never structural
    const ch = argsText[i];
    if (ARG_OPENERS.has(ch)) depth += 1;
    else if (ARG_CLOSERS.has(ch)) depth = Math.max(0, depth - 1);
    else if (ch === "," && depth === 0) {
      pushSegment(segStart, i);
      segStart = i + 1;
    }
  }
  pushSegment(segStart, argsText.length);
  // Drop a trailing empty segment from a dangling comma; keep a sole empty
  // segment (a zero-arg call `f()` → `[""]`) out too — both are non-positional.
  return args.filter((a) => a.text.trim().length > 0);
}

/**
 * Whether an argument's raw text is a POSITIONAL argument (eligible for a name
 * chip) rather than a keyword argument (`name=value`) or a spread (`*xs` /
 * `**kw`). A keyword argument already shows its name in the source, and a spread
 * fills an unknown number of slots, so neither gets a chip.
 *
 * The `=` test is deliberately careful: only a TOP-LEVEL `=` that is NOT part of
 * a comparison / augmented operator (`==`, `<=`, `>=`, `!=`, `:=`) marks a
 * keyword argument. A nested `=` (inside a lambda default, a dict, a call) is
 * ignored via the same bracket/skip-mask discipline as {@link splitTopLevelArgs}.
 */
function isPositional(rawArg: string): boolean {
  const trimmed = rawArg.trim();
  if (trimmed.length === 0) return false;
  if (trimmed.startsWith("*")) return false; // *args / **kwargs spread
  const mask = computeSkipMask(rawArg);
  let depth = 0;
  for (let i = 0; i < rawArg.length; i += 1) {
    if (mask[i]) continue;
    const ch = rawArg[i];
    if (ARG_OPENERS.has(ch)) depth += 1;
    else if (ARG_CLOSERS.has(ch)) depth = Math.max(0, depth - 1);
    else if (ch === "=" && depth === 0) {
      const prev = rawArg[i - 1];
      const next = rawArg[i + 1];
      // Skip ==, :=, <=, >=, != and the second half of any of those.
      if (next === "=") continue; // `==`
      if (prev === "=" || prev === "!" || prev === "<" || prev === ">" || prev === ":") {
        continue;
      }
      return false; // a real top-level `name=value` keyword argument
    }
  }
  return true;
}

/**
 * PURE core: given the source text of a call's ARGUMENT LIST (the text strictly
 * between the call's `(` and its matching `)`, NOT including the parens) and the
 * ordered parameter names jedi reported for that call, return one
 * {@link InlayHint} per POSITIONAL argument — the parameter name (with a trailing
 * colon) anchored at the argument's first char (offset relative to the start of
 * `argsText`).
 *
 * Rules (matching VSCode's parameter-name hints):
 *  - positional arguments are matched to parameter names left-to-right, each
 *    argument consuming exactly one parameter slot;
 *  - a keyword argument (`name=value`) gets no chip (its name is already visible)
 *    and, because Python forbids a positional after a keyword, we STOP at it;
 *  - a spread (`*xs` / `**kw`) consumes the rest ambiguously, so we stop there;
 *  - a `self` / `cls` parameter consumes its positional argument SILENTLY (no
 *    chip) — this only occurs for an explicit unbound call like
 *    `Klass.method(obj, …)`, since jedi already drops a bound instance method's
 *    `self` from `.params`; the chip would otherwise mislabel the receiver;
 *  - a `*args` / `**kwargs` COLLECTOR parameter swallows the remaining arguments
 *    positionally, so we stop emitting once one is reached;
 *  - surplus arguments beyond the known parameters get no chip;
 *  - a hint whose label echoes the argument text verbatim (e.g. arg `value` for
 *    parameter `value`) is suppressed as redundant noise, mirroring VSCode.
 *
 * Total and side-effect-free. An empty `argsText`, empty `paramNames`, or a call
 * with only keyword/spread arguments yields `[]`.
 */
export function computePythonParameterInlayHints(
  argsText: string,
  paramNames: readonly string[],
): InlayHint[] {
  if (!argsText || paramNames.length === 0) return [];
  const args = splitTopLevelArgs(argsText);
  if (args.length === 0) return [];

  const hints: InlayHint[] = [];
  // Walk arguments and parameters in lockstep: each positional argument consumes
  // one parameter slot, so a skipped `self`/`cls` slot is spent on ITS argument
  // (not carried onto the next), keeping the name↔argument alignment correct.
  let paramIdx = 0;
  for (const arg of args) {
    if (paramIdx >= paramNames.length) break; // no parameter left to name
    if (!isPositional(arg.text)) break; // keyword/spread arg → stop (Python ordering)
    const name = paramNames[paramIdx];
    paramIdx += 1;
    // `self`/`cls`: this argument fills the receiver slot — consume it with no chip.
    if (name === "self" || name === "cls") continue;
    // A collector parameter (`*args`/`**kwargs`/bare `*`) takes the rest; no
    // further positional chip is meaningful.
    if (!isNameableParam(name)) break;
    // Suppress a chip that merely echoes the argument text (VSCode parity).
    if (arg.text.trim() === name) continue;
    hints.push({ offset: arg.start, label: `${name}:` });
  }
  return hints;
}

/* ------------------------------------------------------------------------- *
 *  CM6 wiring (browser-/kernel-gated; not exercised by the pure unit tests). *
 * ------------------------------------------------------------------------- */

/**
 * Upper bound on how many distinct call sites one inlay refresh resolves through
 * jedi, so scrolling a call-dense file never fans out into an unbounded burst of
 * worker round-trips. Visible-range-only scanning already keeps this small; the
 * cap is a hard backstop.
 */
export const MAX_INLAY_CALLS_PER_REFRESH = 40;

/** Debounce (ms) before an edit/scroll triggers a fresh inlay resolve. */
const INLAY_REFRESH_DELAY_MS = 250;

/** Inline atomic widget that paints a single faint `name:` parameter chip. */
class InlayChipWidget extends WidgetType {
  constructor(private readonly text: string) {
    super();
  }
  eq(other: InlayChipWidget): boolean {
    return other.text === this.text;
  }
  toDOM(): HTMLElement {
    const span = document.createElement("span");
    span.className = "cm-inlay-hint";
    span.setAttribute("aria-hidden", "true");
    // textContent (never innerHTML) — the parameter name is plain text and must
    // never inject markup into the editor surface.
    span.textContent = this.text;
    return span;
  }
  ignoreEvent(): boolean {
    return true;
  }
}

/**
 * A located call site in the buffer: the offset just AFTER its `(` (where the
 * argument list starts) and the offset of its matching `)`. The args text is
 * `doc.slice(argsFrom, argsTo)`.
 */
interface CallSite {
  /** Offset just after the opening paren (start of the argument list). */
  argsFrom: number;
  /** Offset of the matching closing paren (end of the argument list, exclusive). */
  argsTo: number;
}

/**
 * Find the CALL sites whose argument list intersects `[from, to]` — an opener
 * `(` that (a) is immediately preceded by an identifier char or a `)` / `]`
 * (i.e. it is a call, not a grouping paren or a `def`/`class` header) and (b) has
 * a matching `)` found by a depth scan. Skips parens inside strings/comments via
 * {@link computeSkipMask}. Returns at most {@link MAX_INLAY_CALLS_PER_REFRESH}
 * sites for the bounded resolve.
 */
function findVisibleCallSites(text: string, from: number, to: number): CallSite[] {
  const mask = computeSkipMask(text);
  const sites: CallSite[] = [];
  const isIdentChar = (c: string | undefined): boolean =>
    c != null && /[A-Za-z0-9_]/.test(c);
  for (let i = 0; i < text.length && sites.length < MAX_INLAY_CALLS_PER_REFRESH; i += 1) {
    if (mask[i] || text[i] !== "(") continue;
    // Look back past spaces for the char that makes this a call.
    let j = i - 1;
    while (j >= 0 && (text[j] === " " || text[j] === "\t")) j -= 1;
    const prev = j >= 0 ? text[j] : undefined;
    const isCall = isIdentChar(prev) || prev === ")" || prev === "]";
    if (!isCall) continue;
    // A `def name(` / `class name(` header is a definition, not a call — its
    // parameters are declarations, not arguments, so it gets no argument chips.
    // Detect by scanning the keyword at the start of the opener's line.
    const lineStart = text.lastIndexOf("\n", i - 1) + 1;
    const lead = text.slice(lineStart, i).trimStart();
    if (lead.startsWith("def ") || lead.startsWith("class ") || lead.startsWith("async def ")) {
      continue;
    }
    // Find the matching close paren by depth (only structural parens count).
    let depth = 0;
    let close = -1;
    for (let k = i; k < text.length; k += 1) {
      if (mask[k]) continue;
      const c = text[k];
      if (c === "(") depth += 1;
      else if (c === ")") {
        depth -= 1;
        if (depth === 0) {
          close = k;
          break;
        }
      }
    }
    if (close < 0) continue; // unterminated call — skip
    // Keep only calls whose argument span intersects the visible range.
    if (close >= from && i + 1 <= to) {
      sites.push({ argsFrom: i + 1, argsTo: close });
    }
  }
  return sites;
}

/** One resolved chip ready to place: an absolute buffer `offset` + chip text. */
interface ResolvedHint {
  offset: number;
  label: string;
}

/**
 * StateEffect carrying a freshly-resolved hint set. The async resolver dispatches
 * it; the {@link inlayHintField} folds it into a DecorationSet. A `null` payload
 * clears the hints (e.g. the visible range now has no call sites).
 */
const setInlayHints = StateEffect.define<ResolvedHint[] | null>();

/**
 * Build the atomic widget DecorationSet for a resolved hint list. Hints are
 * sorted by offset (CM6 requires ascending ranges) and each renders as an inline
 * {@link InlayChipWidget} placed BEFORE its argument's first char (`side: -1`).
 */
function decorationsForHints(hints: ResolvedHint[]): DecorationSet {
  const sorted = [...hints].sort((a, b) => a.offset - b.offset);
  return Decoration.set(
    sorted.map((h) =>
      Decoration.widget({
        widget: new InlayChipWidget(h.label),
        side: -1,
      }).range(h.offset),
    ),
    true,
  );
}

/**
 * The StateField holding the live inlay-hint decorations. Starts EMPTY (so the
 * editor is byte-for-byte unchanged until the first resolve lands), maps its
 * ranges through document changes (a stale chip rides along until the next
 * resolve replaces it rather than jumping), and swaps in a fresh set on a
 * {@link setInlayHints} effect.
 */
const inlayHintField = StateField.define<DecorationSet>({
  create: () => Decoration.none,
  update(deco, tr) {
    let next = deco.map(tr.changes);
    for (const effect of tr.effects) {
      if (effect.is(setInlayHints)) {
        next = effect.value ? decorationsForHints(effect.value) : Decoration.none;
      }
    }
    return next;
  },
  provide: (field) => EditorView.decorations.from(field),
});

/** Faint chip styling; palette themable via CSS vars with neutral fallbacks. */
const inlayTheme = EditorView.baseTheme({
  ".cm-inlay-hint": {
    color: "var(--cm-inlay-hint-color, rgba(127, 127, 127, 0.85))",
    backgroundColor: "var(--cm-inlay-hint-bg, rgba(127, 127, 127, 0.12))",
    borderRadius: "4px",
    padding: "0 3px",
    margin: "0 2px 0 0",
    fontSize: "0.85em",
    // Never selectable — a chip must not enter the user's selection or disturb
    // the monospace caret math of the surrounding real text.
    userSelect: "none",
  },
});

/**
 * The ViewPlugin that drives the async resolve: on construction and on every
 * document/viewport change it (debounced) scans the VISIBLE call sites, resolves
 * each call's parameter names via the jedi `signaturePython` round-trip, maps
 * them to chip offsets with the pure {@link computePythonParameterInlayHints},
 * and dispatches a {@link setInlayHints} effect. A monotonic generation guard
 * drops any resolve superseded by a newer edit/scroll or by teardown, so a slow
 * worker reply can never paint stale chips or dispatch after destroy.
 */
function inlayResolverPlugin(namespace: string): Extension {
  return ViewPlugin.fromClass(
    class {
      private timer: ReturnType<typeof setTimeout> | null = null;
      private generation = 0;

      constructor(view: EditorView) {
        this.schedule(view);
      }

      update(update: ViewUpdate) {
        if (update.docChanged || update.viewportChanged) {
          this.schedule(update.view);
        }
      }

      destroy() {
        if (this.timer != null) clearTimeout(this.timer);
        this.generation += 1; // invalidate any in-flight resolve
      }

      private schedule(view: EditorView) {
        if (this.timer != null) clearTimeout(this.timer);
        this.generation += 1;
        const gen = this.generation;
        this.timer = setTimeout(() => {
          void this.resolve(view, gen);
        }, INLAY_REFRESH_DELAY_MS);
      }

      private async resolve(view: EditorView, gen: number) {
        const { hasPythonKernel, signaturePython } = await import(
          "@mrnote/core/web/lib/code-runner/python-runner.ts"
        );
        // Never boot Pyodide as a side effect of scrolling/typing; without a live
        // kernel there are simply no hints (editor unchanged).
        if (!hasPythonKernel(namespace) || gen !== this.generation) return;

        const doc = view.state.doc;
        const text = doc.toString();
        const from = view.visibleRanges[0]?.from ?? 0;
        const to =
          view.visibleRanges[view.visibleRanges.length - 1]?.to ?? text.length;
        const sites = findVisibleCallSites(text, from, to);

        const collected: ResolvedHint[] = [];
        for (const site of sites) {
          if (gen !== this.generation) return; // superseded mid-resolve
          // Query jedi at a position INSIDE the call's parens (just after `(`),
          // where get_signatures resolves the enclosing call's parameters.
          const line = doc.lineAt(site.argsFrom);
          const info = await signaturePython(
            text,
            line.number,
            site.argsFrom - line.from,
            { namespace },
          );
          if (!info || info.parameters.length === 0) continue;
          const argsText = text.slice(site.argsFrom, site.argsTo);
          for (const h of computePythonParameterInlayHints(argsText, info.parameters)) {
            collected.push({ offset: site.argsFrom + h.offset, label: h.label });
          }
        }
        if (gen !== this.generation) return;
        // Dispatch the fresh set (or a clear). Guard the view being torn down.
        try {
          view.dispatch({ effects: setInlayHints.of(collected) });
        } catch {
          /* view destroyed between the guard and dispatch — ignore. */
        }
      }
    },
  );
}

/**
 * Python call-site parameter-name inlay hints (the `editor.inlayHints`
 * parameter-name subset). Activates ONLY when (a) the language is python and
 * (b) a kernel `namespace` is wired — otherwise it returns `[]` so callers append
 * nothing and the editor is byte-for-byte unchanged. Even when active, the field
 * starts empty and only paints once a live kernel resolves the first call's
 * parameters, so a Python buffer with no kernel / no calls is also unchanged.
 */
export function pythonInlayHintsExtension(
  language: string,
  namespace: string | undefined,
): Extension {
  const lang = (language || "").trim().toLowerCase();
  if ((lang !== "python" && lang !== "py") || !namespace) return [];
  return [inlayTheme, inlayHintField, inlayResolverPlugin(namespace)];
}
