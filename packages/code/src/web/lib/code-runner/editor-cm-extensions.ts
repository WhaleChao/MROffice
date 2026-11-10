/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * CodeMirror 6 VIEW extensions for the editor's VSCode-parity display toggles:
 *  - {@link wordWrapExtension}          — soft line wrapping (Alt+Z / chrome);
 *  - {@link renderWhitespaceExtension}  — middle-dots / arrows for spaces+tabs;
 *  - {@link bracketColorizationExtension} — depth-cycling colors on matched
 *    bracket pairs (the pure depth logic lives in editor-decorations.ts).
 *
 * These are thin CM6 wirings of pure helpers; the heavy/branching logic is the
 * unit-tested `computeBracketDepths` in editor-decorations.ts. Each is its own
 * exported `Extension` so `CodeMirrorEditor` can toggle them independently and
 * tests can assert each builds into a valid `EditorState` without a layout
 * engine.
 */
import {
  Decoration,
  EditorView,
  GutterMarker,
  ViewPlugin,
  WidgetType,
  gutter,
  highlightWhitespace,
  type DecorationSet,
  type ViewUpdate,
} from "@codemirror/view";
import {
  Facet,
  Prec,
  RangeSetBuilder,
  StateEffect,
  StateField,
  type ChangeDesc,
  type EditorState,
  type Extension,
  type Text,
  type Transaction,
} from "@codemirror/state";
import { foldService } from "@codemirror/language";

import {
  bracketDepthClass,
  computeBracketDepths,
  computeBracketPairGuides,
  computeIndentGuides,
} from "./editor-decorations";
import {
  computeIndentFoldRange,
  computeRegionFoldRange,
} from "./stream-fold";
import {
  colorLiteralToHex,
  replaceColorLiteral,
  scanColorLiterals,
} from "./css-color-decorators";
import type { CodeLens, InlayHintItem } from "./code-intelligence";
import {
  PAGE_LOCATED_ATTRIBUTE,
  PAGE_LOCATE_FLASH_MS,
  PAGE_LOCATE_SNIPPET_PREFIX,
} from "@mrnote/core/web/lib/page-locate.ts";
import type {
  DirtyDiffHunk,
  DirtyDiffKind,
} from "./dirty-diff";

/** Soft line wrapping (no horizontal scroll) — CM6's built-in toggle. */
export function wordWrapExtension(): Extension {
  return EditorView.lineWrapping;
}

/** A validated 1-based line plus its document start offset. */
export interface CodeLineTarget {
  line: number;
  from: number;
}

/**
 * Resolve an external line request against CodeMirror's rope-backed document.
 *
 * Debuggers and traceback navigators receive 1-based line numbers from another
 * subsystem. This helper makes that boundary total: non-finite requests fall
 * back to line 1, fractional values are truncated, and stale values clamp into
 * the live document. `Text.line()` uses CodeMirror's balanced tree index, so
 * locating a line stays fast for very large buffers instead of rescanning the
 * source from the beginning.
 */
export function resolveCodeLineTarget(
  doc: Text,
  requestedLine: number,
): CodeLineTarget {
  const finiteLine = Number.isFinite(requestedLine)
    ? Math.trunc(requestedLine)
    : 1;
  const line = Math.max(1, Math.min(finiteLine, doc.lines));
  return { line, from: doc.line(line).from };
}

/**
 * Effect driving the debugger's current-statement line decoration. `null`
 * clears it; a number is resolved and clamped against the transaction's live
 * document before the decoration is installed.
 */
export const setDebugPausedLineEffect = StateEffect.define<number | null>();

const debugPausedLineField = StateField.define<DecorationSet>({
  create: () => Decoration.none,
  update(decorations, transaction) {
    let next = decorations.map(transaction.changes);
    for (const effect of transaction.effects) {
      if (!effect.is(setDebugPausedLineEffect)) continue;
      if (effect.value === null) {
        next = Decoration.none;
      } else {
        const target = resolveCodeLineTarget(
          transaction.state.doc,
          effect.value,
        );
        next = Decoration.set([
          Decoration.line({ class: "cm-debug-paused-line" }).range(target.from),
        ]);
      }
    }
    return next;
  },
  provide: (field) => EditorView.decorations.from(field),
});

/**
 * Mount the inert debugger line-decoration field. It renders nothing until
 * {@link setDebugPausedLineEffect} supplies a paused line.
 */
export function debugPausedLineExtension(): Extension {
  return debugPausedLineField;
}

/**
 * Locate flash: the line a jump lands on (go to line, a symbol, a problem, a
 * reference, a breadcrumb, an AI source's "Locate") carries
 * data-located="true" for PAGE_LOCATE_FLASH_MS, so the shared locate tint of
 * styles/office-motion.css (the Docs TOC-jump flash, reduced-motion safe)
 * marks it. The marker is a CodeMirror line decoration, not a DOM attribute
 * written from outside, because CodeMirror owns its line elements and would
 * reset a foreign attribute on its next redraw.
 *
 * The effect's value is a document position on the line to flash (clamped),
 * or null to clear.
 */
export const flashLocatedLineEffect = StateEffect.define<number | null>();

const locatedLineDecoration = Decoration.line({
  attributes: { [PAGE_LOCATED_ATTRIBUTE]: "true" },
});

const locatedLineField = StateField.define<DecorationSet>({
  create: () => Decoration.none,
  update(decorations, transaction) {
    let next = decorations.map(transaction.changes);
    for (const effect of transaction.effects) {
      if (!effect.is(flashLocatedLineEffect)) continue;
      if (effect.value === null) {
        next = Decoration.none;
        continue;
      }
      const doc = transaction.state.doc;
      const position = Number.isFinite(effect.value)
        ? Math.min(Math.max(Math.trunc(effect.value), 0), doc.length)
        : 0;
      next = Decoration.set([
        locatedLineDecoration.range(doc.lineAt(position).from),
      ]);
    }
    return next;
  },
  provide: (field) => EditorView.decorations.from(field),
});

/** Clears the flash when it has run; a newer flash restarts the clock. */
const locatedLineClock = ViewPlugin.fromClass(
  class {
    private timer: ReturnType<typeof setTimeout> | null = null;

    constructor(private readonly view: EditorView) {}

    update(update: ViewUpdate): void {
      const flashed = update.transactions.some((transaction) =>
        transaction.effects.some(
          (effect) => effect.is(flashLocatedLineEffect) && effect.value !== null,
        ),
      );
      if (!flashed) return;
      if (this.timer !== null) clearTimeout(this.timer);
      this.timer = setTimeout(() => {
        this.timer = null;
        this.view.dispatch({ effects: flashLocatedLineEffect.of(null) });
      }, PAGE_LOCATE_FLASH_MS);
    }

    destroy(): void {
      if (this.timer !== null) clearTimeout(this.timer);
    }
  },
);

/** Mount the locate flash; it renders nothing until a jump flashes a line. */
export function locatedLineExtension(): Extension {
  return [locatedLineField, locatedLineClock];
}

/** The 1-based line the locate flash currently marks, or null. */
export function locatedLineNumber(state: EditorState): number | null {
  const decorations = state.field(locatedLineField, false);
  if (!decorations || decorations.size === 0) return null;
  const cursor = decorations.iter();
  return cursor.value ? state.doc.lineAt(cursor.from).number : null;
}

function normalizeLocateText(value: string): string {
  return value.replace(/\s+/g, " ").trim();
}

/**
 * The 1-based line an AI source's "Locate" request points at in a code page:
 * the line where the snippet's first 40 characters begin (whitespace is
 * collapsed on both sides, so a snippet that spans lines still matches),
 * else `paragraphIndex` read as a line number (clamped to the document).
 * Null when neither resolves.
 */
export function codeLocateLine(
  doc: Text,
  target: {
    readonly snippet?: string | null;
    readonly paragraphIndex?: number | null;
  },
): number | null {
  const needle = normalizeLocateText(target.snippet ?? "")
    .replace(/^[\s"'“”‘’「」『』…]+/, "")
    .slice(0, PAGE_LOCATE_SNIPPET_PREFIX)
    .trim();
  if (needle) {
    // One space-joined haystack of the non-blank lines, with each line's
    // start in it, so a match maps back to the line it begins on.
    let haystack = "";
    const starts: number[] = [];
    const lines: number[] = [];
    for (let number = 1; number <= doc.lines; number += 1) {
      const text = normalizeLocateText(doc.line(number).text);
      if (!text) continue;
      if (haystack) haystack += " ";
      starts.push(haystack.length);
      lines.push(number);
      haystack += text;
    }
    const at = haystack.indexOf(needle);
    if (at >= 0) {
      let low = 0;
      let high = starts.length - 1;
      while (low < high) {
        const middle = (low + high + 1) >> 1;
        if (starts[middle] <= at) low = middle;
        else high = middle - 1;
      }
      return lines[low] ?? null;
    }
  }
  const index = target.paragraphIndex;
  if (typeof index === "number" && Number.isInteger(index) && index > 0) {
    return Math.min(index, doc.lines);
  }
  return null;
}

/** Replace the debugger breakpoint markers without rebuilding the editor. */
export const setDebugBreakpointLinesEffect =
  StateEffect.define<readonly number[]>();

const MAX_DEBUG_BREAKPOINT_GUTTER_LINES = 200;

/**
 * Defensive, bounded breakpoint-line normalization for the editor gutter.
 *
 * The persisted debugger configuration already applies the same 200-entry
 * ceiling. Keeping the view boundary defensive as well means an accidental or
 * hostile prop cannot make every visible-line render scan an unbounded array.
 */
export function normalizeDebugBreakpointLines(
  lines: readonly number[],
): readonly number[] {
  const unique = new Set<number>();
  const scanLimit = Math.min(
    lines.length,
    MAX_DEBUG_BREAKPOINT_GUTTER_LINES * 4,
  );
  for (
    let index = 0;
    index < scanLimit && unique.size < MAX_DEBUG_BREAKPOINT_GUTTER_LINES;
    index += 1
  ) {
    const line = lines[index];
    if (
      typeof line === "number" &&
      Number.isFinite(line) &&
      Number.isInteger(line) &&
      line >= 1
    ) {
      unique.add(line);
    }
  }
  return [...unique].sort((left, right) => left - right);
}

export interface DebugBreakpointLineMove {
  readonly fromLine: number;
  readonly toLine: number;
}

/**
 * Map authored breakpoint lines through one CodeMirror change set.
 *
 * A positive association keeps a breakpoint attached to the source statement
 * after text/newlines are inserted exactly at its start. Breakpoints beyond
 * the current document remain authored-but-hidden instead of collapsing onto
 * the last line.
 */
export function rebaseDebugBreakpointLines(
  startDoc: Text,
  endDoc: Text,
  changes: ChangeDesc,
  lines: readonly number[],
): readonly DebugBreakpointLineMove[] {
  return normalizeDebugBreakpointLines(lines).map((fromLine) => {
    if (fromLine > startDoc.lines) return { fromLine, toLine: fromLine };
    const mappedOffset = changes.mapPos(startDoc.line(fromLine).from, 1);
    const toLine = endDoc.lineAt(
      Math.max(0, Math.min(mappedOffset, endDoc.length)),
    ).number;
    return { fromLine, toLine };
  });
}

export interface DebugBreakpointGutterOptions {
  readonly onToggle: (line: number) => void;
  readonly labelForLine: (line: number, active: boolean) => string;
  readonly onRebase?: (
    moves: readonly DebugBreakpointLineMove[],
    source: string,
  ) => void;
  readonly shouldRebaseTransaction?: (transaction: Transaction) => boolean;
}

function fallbackDebugBreakpointLabel(
  line: number,
  active: boolean,
): string {
  return `${active ? "Remove" : "Add"} breakpoint at line ${line}`;
}

const alwaysRebaseDebugBreakpointTransaction = () => true;
const debugBreakpointRebaseTransactionFacet = Facet.define<
  (transaction: Transaction) => boolean,
  (transaction: Transaction) => boolean
>({
  combine: (predicates) =>
    predicates[0] ?? alwaysRebaseDebugBreakpointTransaction,
});

/**
 * The breakpoints the debugger cannot pause at, so the gutter can draw them the
 * way VS Code draws an unverified breakpoint: hollow. Kept as its own effect and
 * field so a host that knows nothing about binding keeps its old behaviour —
 * every breakpoint solid — instead of having to answer a question it cannot.
 */
export const setDebugUnboundBreakpointLinesEffect =
  StateEffect.define<readonly number[]>();

const debugUnboundBreakpointLinesField = StateField.define<readonly number[]>({
  create: () => [],
  update(lines, transaction) {
    for (const effect of transaction.effects) {
      if (effect.is(setDebugUnboundBreakpointLinesEffect)) {
        return normalizeDebugBreakpointLines(effect.value);
      }
    }
    // Deliberately NOT rebased across edits: whether a line can bind depends on
    // the code itself, so the host recomputes and re-sends it. A rebased guess
    // would keep a stale hollow dot on a line that now binds.
    return lines;
  },
});

const debugBreakpointLinesField = StateField.define<readonly number[]>({
  create: () => [],
  update(lines, transaction) {
    for (const effect of transaction.effects) {
      if (effect.is(setDebugBreakpointLinesEffect)) {
        return normalizeDebugBreakpointLines(effect.value);
      }
    }
    if (
      !transaction.docChanged ||
      !transaction.startState
        .facet(debugBreakpointRebaseTransactionFacet)(transaction)
    ) {
      return lines;
    }
    return normalizeDebugBreakpointLines(
      rebaseDebugBreakpointLines(
        transaction.startState.doc,
        transaction.newDoc,
        transaction.changes,
        lines,
      ).map((move) => move.toLine),
    );
  },
});

class DebugBreakpointGutterMarker extends GutterMarker {
  constructor(
    readonly line: number,
    readonly active: boolean,
    readonly label: string,
    /** False when the debugger cannot pause here — drawn hollow. */
    readonly bound: boolean = true,
  ) {
    super();
  }

  eq(other: GutterMarker): boolean {
    return (
      other instanceof DebugBreakpointGutterMarker &&
      other.line === this.line &&
      other.active === this.active &&
      other.bound === this.bound &&
      other.label === this.label
    );
  }

  toDOM(): Node {
    const marker = document.createElement("span");
    marker.className = "cm-debug-breakpoint-marker";
    marker.dataset.active = String(this.active);
    marker.dataset.bound = String(this.bound);
    marker.dataset.debugBreakpointLine = String(this.line);
    marker.title = this.label;
    marker.setAttribute("aria-hidden", "true");
    return marker;
  }
}

const debugBreakpointGutterTheme = EditorView.baseTheme({
  ".cm-debug-breakpoint-gutter": {
    width: "16px",
  },
  ".cm-debug-breakpoint-gutter .cm-gutterElement": {
    alignItems: "center",
    cursor: "pointer",
    display: "flex",
    justifyContent: "center",
    minWidth: "16px",
  },
  ".cm-debug-breakpoint-marker": {
    backgroundColor: "#e5484d",
    border: "1px solid rgba(255, 255, 255, 0.9)",
    borderRadius: "50%",
    boxShadow: "0 0 0 1px rgba(127, 29, 29, 0.28)",
    boxSizing: "border-box",
    display: "block",
    height: "10px",
    opacity: "0",
    width: "10px",
  },
  ".cm-debug-breakpoint-gutter .cm-gutterElement:hover .cm-debug-breakpoint-marker":
    {
      opacity: "0.35",
      transform: "scale(0.9)",
    },
  '.cm-debug-breakpoint-marker[data-active="true"]': {
    opacity: "1",
    transform: "scale(1)",
  },
  // Unverified: present, but visibly not armed — the same language VS Code
  // uses for a breakpoint that could not be bound.
  '.cm-debug-breakpoint-marker[data-active="true"][data-bound="false"]': {
    backgroundColor: "transparent",
    border: "1.5px solid #9ca3af",
    boxShadow: "none",
  },
});

/** One change marker in the dirty-diff gutter. */
class DirtyDiffGutterMarker extends GutterMarker {
  constructor(
    private readonly kind: DirtyDiffKind,
    private readonly line: number,
    private readonly label: string,
  ) {
    super();
  }

  eq(other: DirtyDiffGutterMarker) {
    return (
      other.kind === this.kind &&
      other.line === this.line &&
      other.label === this.label
    );
  }

  toDOM() {
    const marker = document.createElement("button");
    marker.type = "button";
    marker.className = `cm-dirty-diff cm-dirty-diff-${this.kind}`;
    marker.dataset.testid = "code-dirty-diff-marker";
    marker.dataset.kind = this.kind;
    // The click handler reads the line from here rather than from the event's
    // resolved block: coordinate-based resolution is unreliable, and reverting
    // the wrong hunk is worse than not reverting at all.
    marker.dataset.dirtyDiffLine = String(this.line);
    marker.title = this.label;
    marker.setAttribute("aria-label", this.label);
    return marker;
  }
}

/** Replace the hunks the gutter is drawing. */
export const setDirtyDiffHunksEffect = StateEffect.define<DirtyDiffHunk[]>();

export const dirtyDiffHunksField = StateField.define<DirtyDiffHunk[]>({
  create: () => [],
  update(hunks, transaction) {
    for (const effect of transaction.effects) {
      if (effect.is(setDirtyDiffHunksEffect)) return effect.value;
    }
    // Hunks are line ranges into a specific document. Once that document
    // changes they are stale, and drawing stale markers is worse than drawing
    // none — the host recomputes and publishes a fresh set.
    return transaction.docChanged ? [] : hunks;
  },
});

export interface DirtyDiffGutterOptions {
  /** Human label for a marker, by change kind. */
  labelForKind: (kind: DirtyDiffKind) => string;
  /** Revert one hunk. The host owns the write so it stays one undo step. */
  onRevert: (hunk: DirtyDiffHunk) => void;
}

/**
 * A gutter marking every line changed since the baseline, with one click to put
 * that change back (VS Code's `scm.diffDecorations` and "Revert this change").
 *
 * The hunks come from the host, which owns the baseline and recomputes it; this
 * extension only draws what it is given and reports a click. Markers clear the
 * moment the document changes, so a stale range can never be reverted.
 */
export function dirtyDiffGutterExtension(
  options: DirtyDiffGutterOptions,
): Extension {
  const { labelForKind, onRevert } = options;
  return [
    dirtyDiffHunksField,
    // The markers ship their own look: a thin coloured bar in its own narrow
    // gutter, the way VS Code draws them. Without this they render as
    // zero-sized buttons — present in the DOM, invisible and unclickable.
    EditorView.baseTheme({
      ".cm-dirty-diff-gutter": {
        width: "4px",
        padding: "0",
      },
      ".cm-dirty-diff": {
        display: "block",
        width: "4px",
        height: "100%",
        minHeight: "1em",
        padding: "0",
        border: "none",
        borderRadius: "1px",
        cursor: "pointer",
        background: "transparent",
      },
      ".cm-dirty-diff-added": { background: "#2ea043" },
      ".cm-dirty-diff-modified": { background: "#2f81f7" },
      // A deletion has no lines of its own, so it reads as a marker at the
      // boundary rather than a bar spanning content that is still there.
      ".cm-dirty-diff-deleted": {
        background:
          "linear-gradient(to bottom, #cf222e 0 40%, transparent 40% 100%)",
      },
      ".cm-dirty-diff:hover": { filter: "brightness(1.2)" },
      ".cm-dirty-diff:focus-visible": {
        outline: "2px solid #2f81f7",
        outlineOffset: "1px",
      },
    }),
    gutter({
      class: "cm-dirty-diff-gutter",
      renderEmptyElements: false,
      lineMarker(view, block) {
        const line = view.state.doc.lineAt(block.from).number;
        const hunk = view.state
          .field(dirtyDiffHunksField)
          .find((entry) => line >= entry.fromLine && line <= entry.toLine);
        return hunk
          ? new DirtyDiffGutterMarker(hunk.kind, line, labelForKind(hunk.kind))
          : null;
      },
      lineMarkerChange(update) {
        return (
          update.docChanged ||
          update.transactions.some((transaction) =>
            transaction.effects.some((effect) =>
              effect.is(setDirtyDiffHunksEffect),
            ),
          )
        );
      },
      domEventHandlers: {
        mousedown(view, block, event) {
          if (!(event instanceof MouseEvent) || event.button !== 0) return false;
          const marker =
            event.target instanceof Element
              ? event.target.closest<HTMLElement>("[data-dirty-diff-line]")
              : null;
          const markerLine = Number(marker?.dataset.dirtyDiffLine);
          const line =
            Number.isInteger(markerLine) && markerLine >= 1
              ? markerLine
              : view.state.doc.lineAt(block.from).number;
          const hunk = view.state
            .field(dirtyDiffHunksField)
            .find((entry) => line >= entry.fromLine && line <= entry.toLine);
          if (!hunk) return false;
          event.preventDefault();
          onRevert(hunk);
          return true;
        },
      },
    }),
  ];
}

/**
 * VS Code-style breakpoint gutter. The host owns authored breakpoint state;
 * this view extension only renders the bounded line set and reports primary
 * pointer toggles. Keyboard users retain the equivalent, focusable "break at
 * current line" command in CodeWindow's Debug ribbon.
 *
 * The legacy positional form remains accepted so a development HMR cycle that
 * updates this module before its caller cannot crash an already-mounted editor.
 */
export function debugBreakpointGutterExtension(
  optionsOrToggle:
    | DebugBreakpointGutterOptions
    | ((line: number) => void),
  legacyLabelForLine?: (line: number, active: boolean) => string,
): Extension {
  const options: DebugBreakpointGutterOptions =
    typeof optionsOrToggle === "function"
      ? {
          onToggle: optionsOrToggle,
          labelForLine:
            legacyLabelForLine ?? fallbackDebugBreakpointLabel,
        }
      : optionsOrToggle;
  const {
    onToggle,
    labelForLine = fallbackDebugBreakpointLabel,
    onRebase,
    shouldRebaseTransaction = alwaysRebaseDebugBreakpointTransaction,
  } = options;
  const rebaseReporter = onRebase
    ? ViewPlugin.fromClass(
        class {
          private alive = true;

          update(update: ViewUpdate) {
            if (!update.docChanged) return;
            const initialLines = update.startState.field(
              debugBreakpointLinesField,
            );
            let tracked = initialLines.map((line) => ({
              fromLine: line,
              toLine: line,
            }));
            for (const transaction of update.transactions) {
              if (
                !transaction.docChanged ||
                !shouldRebaseTransaction(transaction)
              ) {
                continue;
              }
              const moves = new Map(
                rebaseDebugBreakpointLines(
                  transaction.startState.doc,
                  transaction.newDoc,
                  transaction.changes,
                  tracked.map((move) => move.toLine),
                ).map((move) => [move.fromLine, move.toLine]),
              );
              tracked = tracked.map((move) => ({
                fromLine: move.fromLine,
                toLine: moves.get(move.toLine) ?? move.toLine,
              }));
            }
            const changed = tracked.filter(
              (move) => move.fromLine !== move.toLine,
            );
            if (changed.length === 0) return;
            const source = update.state.doc.toString();
            queueMicrotask(() => {
              if (this.alive) onRebase(changed, source);
            });
          }

          destroy() {
            this.alive = false;
          }
        },
      )
    : [];
  return [
    debugBreakpointRebaseTransactionFacet.of(shouldRebaseTransaction),
    debugBreakpointLinesField,
    debugUnboundBreakpointLinesField,
    rebaseReporter,
    Prec.highest(
      gutter({
        class: "cm-debug-breakpoint-gutter",
        renderEmptyElements: true,
        lineMarker(view, block) {
          const line = view.state.doc.lineAt(block.from).number;
          const active = view.state
            .field(debugBreakpointLinesField)
            .includes(line);
          const bound =
            !active ||
            !view.state.field(debugUnboundBreakpointLinesField).includes(line);
          return new DebugBreakpointGutterMarker(
            line,
            active,
            labelForLine(line, active),
            bound,
          );
        },
        lineMarkerChange(update) {
          return (
            update.docChanged ||
            update.transactions.some((transaction) =>
              transaction.effects.some(
                (effect) =>
                  effect.is(setDebugBreakpointLinesEffect) ||
                  effect.is(setDebugUnboundBreakpointLinesEffect),
              ),
            )
          );
        },
        domEventHandlers: {
          mousedown(view, block, event) {
            if (!(event instanceof MouseEvent) || event.button !== 0) {
              return false;
            }
            const marker =
              event.target instanceof Element
                ? event.target.closest<HTMLElement>(
                    "[data-debug-breakpoint-line]",
                  )
                : null;
            const markerLine = Number(marker?.dataset.debugBreakpointLine);
            const line =
              Number.isInteger(markerLine) && markerLine >= 1
                ? markerLine
                : view.state.doc.lineAt(block.from).number;
            event.preventDefault();
            onToggle(line);
            view.focus();
            return true;
          },
        },
      }),
    ),
    debugBreakpointGutterTheme,
  ];
}

/**
 * Render whitespace as visible glyphs — middle-dot for spaces, arrow for tabs —
 * using CM6's `highlightWhitespace`. The glyphs themselves are CSS `::before`
 * content the editor theme supplies (so we keep the source bytes untouched).
 */
export function renderWhitespaceExtension(): Extension {
  return highlightWhitespace();
}

/**
 * Bracket-pair colorization: decorate every structural bracket with a
 * depth-cycling CSS class (`cm-bracket-depth-0..5`) so nested pairs read at a
 * glance. Recomputes on document/viewport change off the pure
 * {@link computeBracketDepths} (strings/comments skipped). The decoration is a
 * per-character `mark`; colors come from the editor theme.
 */
export function bracketColorizationExtension(): Extension {
  return ViewPlugin.fromClass(
    class {
      decorations: DecorationSet;

      constructor(view: EditorView) {
        this.decorations = buildBracketDecorations(view);
      }

      update(update: ViewUpdate) {
        if (update.docChanged || update.viewportChanged) {
          this.decorations = buildBracketDecorations(update.view);
        }
      }
    },
    { decorations: (plugin) => plugin.decorations },
  );
}

/**
 * Faint guide color via a CSS var so the theme can override per light/dark.
 * `background-origin: content-box` makes every guide position measure from the
 * TEXT origin (the line's content box), so the rules land exactly on the
 * monospace character grid regardless of CM6's `.cm-line` left padding (which is
 * 6px by default) — no magic pixel offset, padding-agnostic.
 */
const indentGuideTheme = EditorView.baseTheme({
  ".cm-line.cm-indent-guides": {
    backgroundRepeat: "no-repeat",
    backgroundOrigin: "content-box",
  },
});

/**
 * Inline style for one line: a 1px vertical rule per guide column, drawn with a
 * stacked `linear-gradient` per column positioned at `column` character cells
 * from the text origin (content-box origin set in the theme above).
 */
export function indentGuideLineStyle(columns: number[]): string {
  const color = "var(--cm-indent-guide-color, rgba(127, 127, 127, 0.28))";
  const images = columns.map(() => `linear-gradient(${color}, ${color})`).join(", ");
  const positions = columns.map((c) => `${c}ch 0`).join(", ");
  const sizes = columns.map(() => "1px 100%").join(", ");
  return `background-image:${images};background-position:${positions};background-size:${sizes};`;
}

/**
 * Build indent-guide line decorations for the visible ranges. Guide columns are
 * computed over the WHOLE document (so a guide running through a blank line
 * inherits the next non-blank line's indent, like VSCode), but only lines inside
 * the rendered viewport get a decoration — keeping the set small on large files.
 */
function buildIndentGuideDecorations(
  view: EditorView,
  tabSize: number,
): DecorationSet {
  const guides = computeIndentGuides(view.state.doc.toString(), tabSize);
  if (guides.length === 0) return Decoration.none;
  // Group guide columns by 0-based line.
  const byLine = new Map<number, number[]>();
  for (const guide of guides) {
    const existing = byLine.get(guide.line);
    if (existing) existing.push(guide.column);
    else byLine.set(guide.line, [guide.column]);
  }
  const builder = new RangeSetBuilder<Decoration>();
  const doc = view.state.doc;
  for (const range of view.visibleRanges) {
    let pos = range.from;
    while (pos <= range.to) {
      const line = doc.lineAt(pos);
      const columns = byLine.get(line.number - 1); // computeIndentGuides is 0-based
      if (columns && columns.length > 0) {
        builder.add(
          line.from,
          line.from,
          Decoration.line({
            class: "cm-indent-guides",
            attributes: { style: indentGuideLineStyle(columns) },
          }),
        );
      }
      pos = line.to + 1;
    }
  }
  return builder.finish();
}

/**
 * Indentation guides (VSCode parity): a faint vertical rule at each indentation
 * level, recomputed on document / viewport change off the pure, unit-tested
 * {@link computeIndentGuides}. Drawn as per-line `background` gradients so the
 * source bytes are never touched and the guides never interfere with selection
 * or the caret.
 */
export function indentGuidesExtension(tabSize = 2): Extension {
  return [
    indentGuideTheme,
    ViewPlugin.fromClass(
      class {
        decorations: DecorationSet;

        constructor(view: EditorView) {
          this.decorations = buildIndentGuideDecorations(view, tabSize);
        }

        update(update: ViewUpdate) {
          if (update.docChanged || update.viewportChanged) {
            this.decorations = buildIndentGuideDecorations(update.view, tabSize);
          }
        }
      },
      { decorations: (plugin) => plugin.decorations },
    ),
  ];
}

/**
 * Bracket-pair guide theme: the thin connectors share the indent-guide drawing
 * mechanism (content-box-origin background gradients on the monospace grid) but
 * carry their OWN class + color var so a theme can tint the bracket connectors
 * independently of the indent rules. Same `background-origin: content-box` rule
 * so a connector at column `N` lands on the Nth character cell regardless of the
 * line's left padding.
 */
const bracketGuideTheme = EditorView.baseTheme({
  ".cm-line.cm-bracket-pair-guides": {
    backgroundRepeat: "no-repeat",
    backgroundOrigin: "content-box",
  },
});

/**
 * Inline style for one line's bracket-pair connectors: a 1px vertical rule per
 * column, positioned at `column` char cells from the text origin — identical
 * geometry to {@link indentGuideLineStyle} but with the bracket-guide color var
 * so the two guide families are themable apart.
 */
export function bracketPairGuideLineStyle(columns: number[]): string {
  const color = "var(--cm-bracket-guide-color, rgba(127, 127, 127, 0.32))";
  const images = columns.map(() => `linear-gradient(${color}, ${color})`).join(", ");
  const positions = columns.map((c) => `${c}ch 0`).join(", ");
  const sizes = columns.map(() => "1px 100%").join(", ");
  return `background-image:${images};background-position:${positions};background-size:${sizes};`;
}

/**
 * Build bracket-pair-guide line decorations for the visible ranges. Connector
 * cells are computed over the WHOLE document (so a pair whose opener is scrolled
 * out of view still draws its rule through the viewport), but only lines inside
 * the rendered viewport get a decoration — keeping the set small on large files.
 */
function buildBracketGuideDecorations(view: EditorView): DecorationSet {
  const guides = computeBracketPairGuides(view.state.doc.toString());
  if (guides.length === 0) return Decoration.none;
  // Group connector columns by 0-based line.
  const byLine = new Map<number, number[]>();
  for (const guide of guides) {
    const existing = byLine.get(guide.line);
    if (existing) existing.push(guide.column);
    else byLine.set(guide.line, [guide.column]);
  }
  const builder = new RangeSetBuilder<Decoration>();
  const doc = view.state.doc;
  for (const range of view.visibleRanges) {
    let pos = range.from;
    while (pos <= range.to) {
      const line = doc.lineAt(pos);
      const columns = byLine.get(line.number - 1); // computeBracketPairGuides is 0-based
      if (columns && columns.length > 0) {
        builder.add(
          line.from,
          line.from,
          Decoration.line({
            class: "cm-bracket-pair-guides",
            attributes: { style: bracketPairGuideLineStyle(columns) },
          }),
        );
      }
      pos = line.to + 1;
    }
  }
  return builder.finish();
}

/**
 * Bracket-pair guides (VSCode parity): a thin vertical connector tying each
 * matched MULTI-LINE bracket pair down the gutter, recomputed on document /
 * viewport change off the pure, unit-tested {@link computeBracketPairGuides}.
 * Drawn as per-line `background` gradients (the same untouched-bytes mechanism
 * as the indent guides) so the connectors never interfere with selection or the
 * caret. A buffer with no multi-line brackets produces no decoration at all.
 */
export function bracketPairGuidesExtension(): Extension {
  return [
    bracketGuideTheme,
    ViewPlugin.fromClass(
      class {
        decorations: DecorationSet;

        constructor(view: EditorView) {
          this.decorations = buildBracketGuideDecorations(view);
        }

        update(update: ViewUpdate) {
          if (update.docChanged || update.viewportChanged) {
            this.decorations = buildBracketGuideDecorations(update.view);
          }
        }
      },
      { decorations: (plugin) => plugin.decorations },
    ),
  ];
}

/**
 * Build the bracket-colorization decoration set for the view's visible ranges.
 * DEPTHS are computed over the WHOLE document (so nesting is correct even when
 * an opener is scrolled out of view), but only brackets inside the rendered
 * viewport ranges get a decoration — keeping the set small on large files.
 */
function buildBracketDecorations(view: EditorView): DecorationSet {
  const text = view.state.doc.toString();
  const depths = computeBracketDepths(text);
  const builder = new RangeSetBuilder<Decoration>();
  // Decoration mark classes are cached per depth so repeated brackets reuse one
  // Decoration object (CM6 dedupes identical specs, but this avoids allocations).
  const markCache = new Map<number, Decoration>();
  for (const { index, depth } of depths) {
    let inViewport = false;
    for (const range of view.visibleRanges) {
      if (index >= range.from && index < range.to) {
        inViewport = true;
        break;
      }
    }
    if (!inViewport) continue;
    let mark = markCache.get(depth);
    if (!mark) {
      mark = Decoration.mark({ class: bracketDepthClass(depth) });
      markCache.set(depth, mark);
    }
    builder.add(index, index + 1, mark);
  }
  return builder.finish();
}

/**
 * The DOM widget for ONE inlay hint: a small, non-editable inline label rendered
 * BETWEEN document characters (VSCode's grey parameter-name / inferred-type
 * pills). The label text is set via `textContent`, never `innerHTML`, so a
 * symbol name / type string can NEVER inject markup; the widget contributes no
 * characters to the document (it's purely a decoration), so copy/paste, the
 * caret, and `onChange` are unaffected by it.
 *
 * `eq` collapses two widgets with the same text+kind so CM6 reuses the DOM node
 * across recomputes (no flicker, fewer allocations). `ignoreEvent` returns false
 * so clicks/selection inside the widget never reach the editor as edits.
 */
class InlayHintWidget extends WidgetType {
  constructor(
    readonly label: string,
    readonly kind: string,
  ) {
    super();
  }

  eq(other: InlayHintWidget): boolean {
    return other.label === this.label && other.kind === this.kind;
  }

  toDOM(): HTMLElement {
    const span = document.createElement("span");
    // The compiler already FORMATS the label for each family: a parameter-name
    // hint arrives as `name:` (trailing colon included) and an inferred-type hint
    // as `: Type` (leading colon+space included). So we render `this.label`
    // verbatim — appending our own `:` would double it (`name::`). The kind only
    // drives a CSS class (so the theme can style each family) and a
    // data-attribute the tests query.
    const isParameter = this.kind === "Parameter";
    span.className = isParameter
      ? "cm-inlay-hint cm-inlay-hint-parameter"
      : "cm-inlay-hint cm-inlay-hint-type";
    span.setAttribute("data-inlay-kind", this.kind);
    span.setAttribute("aria-hidden", "true");
    // textContent (never innerHTML): the label is plain, injection-proof text.
    span.textContent = this.label;
    return span;
  }

  // The widget is inert chrome: never let a pointer/selection event inside it be
  // treated as an editor interaction (it has no editable content anyway).
  ignoreEvent(): boolean {
    return false;
  }
}

/** Faint inlay-hint pill styling via a CSS var so the theme can override it. */
const inlayHintTheme = EditorView.baseTheme({
  ".cm-inlay-hint": {
    color: "var(--cm-inlay-hint-color, rgba(127, 127, 127, 0.85))",
    backgroundColor: "var(--cm-inlay-hint-bg, rgba(127, 127, 127, 0.12))",
    borderRadius: "4px",
    padding: "0 4px",
    margin: "0 1px",
    fontSize: "0.92em",
    // The widget must never be selectable as text — it is not part of the doc.
    userSelect: "none",
  },
});

/**
 * Build the inlay-hint decoration set for the view's CURRENT viewport only. The
 * caller-supplied `getHints(from, to)` is queried once PER visible range (so the
 * compiler computes hints for the rendered text only, never the whole document),
 * and each returned hint becomes an atomic `Decoration.widget` anchored at its
 * offset. Every offset is CLAMPED into the live document so a stale hint (one
 * computed against an older, longer buffer) can never throw or anchor past EOF.
 *
 * Widgets are added in ascending offset order (RangeSetBuilder requires it). A
 * hint whose clamped offset duplicates the previous one is still added — CM6
 * orders side widgets by `side`; we give them a small positive side so they sit
 * AFTER the character at their offset (VSCode places a parameter hint before the
 * argument, a type hint after the binding — both resolve to "between tokens").
 */
function buildInlayHintDecorations(
  view: EditorView,
  getHints: (fromOffset: number, toOffset: number) => InlayHintItem[],
): DecorationSet {
  const docLength = view.state.doc.length;
  const builder = new RangeSetBuilder<Decoration>();
  // Collect across all visible ranges, then sort by offset so the builder's
  // ascending-order contract holds even if ranges return interleaved offsets.
  const collected: { offset: number; deco: Decoration }[] = [];
  for (const range of view.visibleRanges) {
    const hints = getHints(range.from, range.to);
    for (const hint of hints) {
      // Clamp: a hint offset past the live doc end (computed against an older
      // buffer) is pulled back in-bounds so .add never throws.
      const offset = Math.min(Math.max(hint.offset, 0), docLength);
      const deco = Decoration.widget({
        widget: new InlayHintWidget(hint.text, hint.kind),
        // Positive side → the widget renders just AFTER the character boundary at
        // `offset`; CM6 treats it as atomic so the caret steps over it.
        side: 1,
      });
      collected.push({ offset, deco });
    }
  }
  collected.sort((a, b) => a.offset - b.offset);
  for (const { offset, deco } of collected) {
    builder.add(offset, offset, deco);
  }
  return builder.finish();
}

/**
 * Inlay hints (VSCode parity): inline parameter-name + inferred-type labels for
 * TS/JS, rendered as atomic widget decorations (never document text). The hint
 * data is produced by the caller's async `provideHints(fromOffset, toOffset)`
 * (the editor wires this to its single-buffer TypeScript LanguageService via
 * `inlayHints()` in code-intelligence.ts); this extension owns ONLY the CM6
 * lifecycle — viewport-scoped recompute, debounce, and decoration rendering.
 *
 * Recompute is DEBOUNCED (default 150ms) and fires on docChanged + viewportChange
 * (mirroring the lint linter's delay), so fast typing / scrolling never floods
 * the compiler. Because `provideHints` is async (the TS compiler is dynamically
 * imported), each recompute carries a monotonic generation id; a result that
 * resolves after a newer request started — or after the plugin is destroyed — is
 * dropped, so a stale hint set never lands.
 *
 * SAFETY: the hints are decorations only — they add no characters to the
 * document, so copy/paste, `onChange`, and the serialized buffer are byte-for-
 * byte unaffected. Every widget offset is clamped into the live doc on render
 * (see {@link buildInlayHintDecorations}) so a stale offset can never throw.
 *
 * The caller is responsible for only mounting this when the pref is ON and the
 * language is TS/JS; on every other path the extension is simply absent (so the
 * default editor is byte-for-byte unchanged).
 */
export function inlayHintsExtension(
  provideHints: (
    fromOffset: number,
    toOffset: number,
  ) => Promise<InlayHintItem[]>,
  debounceMs = 150,
): Extension {
  return [
    inlayHintTheme,
    ViewPlugin.fromClass(
      class {
        decorations: DecorationSet = Decoration.none;
        // Most-recent hints keyed by viewport range, recomputed async + debounced.
        private latest: InlayHintItem[] = [];
        private timer: ReturnType<typeof setTimeout> | null = null;
        private generation = 0;
        private destroyed = false;

        constructor(view: EditorView) {
          // First compute on mount (debounced like the rest so a burst of mounts
          // doesn't all fire synchronously).
          this.schedule(view);
        }

        update(update: ViewUpdate) {
          if (update.docChanged || update.viewportChanged) {
            // A source edit retires the visible hint generation immediately.
            // Mapping old labels into new code would make stale inferred values
            // look current throughout the debounce window. Viewport-only moves
            // can safely keep and filter the current generation.
            if (update.docChanged) {
              this.latest = [];
              this.decorations = Decoration.none;
            } else {
              this.decorations = this.render(update.view);
            }
            this.schedule(update.view);
          }
        }

        // Debounced async recompute. Each run claims a generation; a result for a
        // superseded generation (newer edit/scroll) or a destroyed plugin is
        // dropped so a stale hint set never replaces a fresher one.
        private schedule(view: EditorView) {
          if (this.timer != null) clearTimeout(this.timer);
          // Claim the new generation when the work is SCHEDULED, not when the
          // timer fires. An older provider may settle during this debounce gap;
          // it must already be stale at that point.
          const gen = (this.generation += 1);
          this.timer = setTimeout(() => {
            this.timer = null;
            if (this.destroyed) return;
            // Query the union of visible ranges (one provider call covers the
            // rendered span; multiple visible ranges are rare in practice).
            const ranges = view.visibleRanges;
            const from = ranges.length > 0 ? ranges[0].from : 0;
            const to =
              ranges.length > 0 ? ranges[ranges.length - 1].to : view.state.doc.length;
            void provideHints(from, to)
              .then((hints) => {
                // Superseded by a newer request, or the plugin was torn down while
                // the async provider (dynamic TS import) was in flight → discard.
                if (this.destroyed || gen !== this.generation) return;
                this.latest = hints;
                this.decorations = this.render(view);
                // Nudge a re-measure so the new widgets paint (the async resolve is
                // outside CM's update cycle).
                view.requestMeasure();
              })
              .catch(() => {
                // Provider failures degrade to no change (keep the last good set).
              });
          }, debounceMs);
        }

        private render(view: EditorView): DecorationSet {
          // Render from the cached `latest` set, re-clamped to the live viewport.
          return buildInlayHintDecorations(view, (rangeFrom, rangeTo) =>
            this.latest.filter(
              (h) => h.offset >= rangeFrom && h.offset <= rangeTo,
            ),
          );
        }

        destroy() {
          this.destroyed = true;
          if (this.timer != null) {
            clearTimeout(this.timer);
            this.timer = null;
          }
        }
      },
      {
        decorations: (plugin) => plugin.decorations,
        // Make the hint widgets atomic for cursor motion: arrow keys / selection
        // step OVER a hint instead of landing "inside" it (it has no doc content,
        // but this keeps caret motion clean and matches VSCode).
        provide: (plugin) =>
          EditorView.atomicRanges.of(
            (view) => view.plugin(plugin)?.decorations ?? Decoration.none,
          ),
      },
    ),
  ];
}

/**
 * The DOM widget for ONE Code Lens: a small, clickable BLOCK label ("N
 * references") rendered on its own line ABOVE a symbol's definition line
 * (VSCode's grey code-lens annotation). The label text is set via `textContent`,
 * never `innerHTML`, so a symbol name can never inject markup; the widget
 * contributes NO characters to the document (it's a block decoration), so
 * copy/paste, the caret, and `onChange` are unaffected.
 *
 * Clicking the lens invokes the supplied `onClick(nameOffset)` — the editor wires
 * this to open the References panel for the symbol at that offset. `eq` collapses
 * two widgets with the same count+offset so CM6 reuses the DOM node across
 * recomputes (no flicker). `ignoreEvent` returns false so the click reaches the
 * widget's own handler rather than being treated as an editor edit.
 */
class CodeLensWidget extends WidgetType {
  constructor(
    readonly label: string,
    readonly referenceCount: number,
    readonly nameOffset: number,
    readonly onClick: (nameOffset: number) => void,
  ) {
    super();
  }

  eq(other: CodeLensWidget): boolean {
    // The click handler is a stable callback for the editor's lifetime; the
    // VISIBLE state is the label + the offset it acts on, so compare only those.
    return other.label === this.label && other.nameOffset === this.nameOffset;
  }

  toDOM(): HTMLElement {
    const wrap = document.createElement("div");
    wrap.className = "cm-code-lens";
    const button = document.createElement("button");
    button.type = "button";
    button.className = "cm-code-lens-link";
    button.setAttribute("data-code-lens-offset", String(this.nameOffset));
    button.setAttribute("data-reference-count", String(this.referenceCount));
    button.setAttribute("aria-label", this.label);
    // textContent (never innerHTML): the label is plain, injection-proof text.
    button.textContent = this.label;
    button.addEventListener("mousedown", (event) => {
      // Stop the pointer from moving the editor caret / starting a selection.
      // Activation happens on `click`, which native buttons also emit for
      // Enter/Space, so mouse, keyboard and assistive tech share one path.
      event.preventDefault();
      event.stopPropagation();
    });
    button.addEventListener("click", (event) => {
      event.preventDefault();
      event.stopPropagation();
      this.onClick(this.nameOffset);
    });
    wrap.appendChild(button);
    return wrap;
  }

  // The lens has its own click handler; let the event reach it rather than being
  // swallowed/handled as an editor interaction.
  ignoreEvent(): boolean {
    return false;
  }
}

/** Faint code-lens styling via CSS vars so the theme can override per light/dark. */
const codeLensTheme = EditorView.baseTheme({
  ".cm-code-lens": {
    color: "var(--cm-code-lens-color, rgba(127, 127, 127, 0.85))",
    fontSize: "0.85em",
    lineHeight: "1.4",
    // Align the lens with the line below it; CM6 lines have ~6px left padding.
    padding: "0 0 0 6px",
    userSelect: "none",
  },
  ".cm-code-lens-link": {
    cursor: "pointer",
  },
  ".cm-code-lens-link:hover": {
    color: "var(--cm-code-lens-hover-color, #2563eb)",
    textDecoration: "underline",
  },
});

/**
 * Effect carrying a freshly-computed lens set into the {@link codeLensField}. The
 * async provider (dynamic TS import) resolves OUTSIDE CM's update cycle, so the
 * ViewPlugin dispatches this effect to push the result into the StateField, which
 * owns the (block) decorations. Block decorations MUST come from a StateField —
 * CM6 forbids them from a ViewPlugin (`Decoration.replace`/block widgets throw
 * "Block decorations may not be specified via plugins"), which is why the data
 * flows plugin → effect → field rather than the plugin rendering directly.
 */
const setCodeLensesEffect = StateEffect.define<CodeLens[]>();

/**
 * Code Lens (VSCode parity): an inline, clickable "N references" annotation
 * rendered as a BLOCK widget ABOVE each symbol-definition line for TS/JS. The lens
 * data is produced by the caller's async `provideLenses()` (the editor wires this
 * to its single-buffer TypeScript LanguageService via `codeLenses()` in
 * code-intelligence.ts); this extension owns ONLY the CM6 lifecycle — a StateField
 * holding the block decorations, plus a debounced ViewPlugin that recomputes the
 * lens set and pushes it into the field.
 *
 * WHY A STATEFIELD (not a ViewPlugin like inlay hints): block decorations may not
 * be supplied by a ViewPlugin. So the field owns the decorations and a sibling
 * ViewPlugin does the async work, dispatching {@link setCodeLensesEffect} to feed
 * the field. The field also maps its existing decorations through every change set
 * so they don't visibly jump while the async recompute is in flight.
 *
 * WHOLE-DOCUMENT COMPUTE (the headless-render fix): {@link buildCodeLensFromState}
 * builds over the WHOLE doc, never `view.visibleRanges`. Code lenses are few (one
 * per symbol definition) so this is cheap, AND it renders regardless of
 * viewport/layout — in jsdom (and any headless render) the viewport is empty, so a
 * viewport-scoped build would yield ZERO widgets. Whole-doc is what makes the
 * widget actually appear under test.
 *
 * Recompute is DEBOUNCED (default 250ms) and fires on docChanged (mirroring the
 * other intelligence extensions), so fast typing never floods the compiler.
 * Because `provideLenses` is async, each recompute carries a monotonic generation
 * id; a result that resolves after a newer request started — or after the plugin
 * is destroyed — is dropped, so a stale lens set never lands.
 *
 * SAFETY: the lenses are decorations only — they add NO characters to the
 * document, so copy/paste, `onChange`, and the serialized buffer are byte-for-byte
 * unaffected. Every line anchor is clamped/validated on render so a stale offset
 * can never throw.
 *
 * The caller is responsible for only mounting this when the pref is ON and the
 * language is TS/JS; on every other path the extension is simply absent (so the
 * default editor is byte-for-byte unchanged). `onLensClick(nameOffset)` is invoked
 * when a lens is clicked (the editor opens the References panel for that offset).
 */
export function codeLensExtension(
  provideLenses: () => Promise<CodeLens[]>,
  onLensClick: (nameOffset: number) => void,
  formatLabel: (referenceCount: number) => string,
  debounceMs = 250,
): Extension {
  // The StateField owns the block decorations. It maps them through edits (so they
  // track the document while async recompute is pending) and replaces them
  // wholesale when a setCodeLensesEffect arrives.
  const codeLensField = StateField.define<DecorationSet>({
    create: () => Decoration.none,
    update(decorations, tr) {
      // A lens is a clickable query authority for one exact source generation.
      // Clear it synchronously on edits instead of mapping a stale button into
      // the replacement document while the debounced compiler query catches up.
      let next = tr.docChanged ? Decoration.none : decorations;
      // Then apply any incoming fresh lens set (async result landed).
      for (const effect of tr.effects) {
        if (effect.is(setCodeLensesEffect)) {
          // Rebuild from the fresh lens list against the CURRENT doc. We need a
          // view to clamp lines; the field has only state, so we recompute line
          // anchors directly from tr.state.doc here.
          next = buildCodeLensFromState(
            tr.state.doc.lines,
            tr.state,
            effect.value,
            onLensClick,
            formatLabel,
          );
        }
      }
      return next;
    },
    provide: (field) => EditorView.decorations.from(field),
  });

  return [
    codeLensTheme,
    codeLensField,
    ViewPlugin.fromClass(
      class {
        private timer: ReturnType<typeof setTimeout> | null = null;
        private generation = 0;
        private destroyed = false;

        constructor(view: EditorView) {
          // First compute on mount (debounced like the rest).
          this.schedule(view);
        }

        update(update: ViewUpdate) {
          // Only the DOCUMENT changing can change the lens set / counts; viewport
          // scrolling does not (whole-doc compute). Recompute on edits only.
          if (update.docChanged) {
            this.schedule(update.view);
          }
        }

        // Debounced async recompute. Each run claims a generation; a result for a
        // superseded generation (a newer edit) or a destroyed plugin is dropped so
        // a stale lens set never replaces a fresher one.
        private schedule(view: EditorView) {
          if (this.timer != null) clearTimeout(this.timer);
          // Invalidate an in-flight provider as soon as a newer computation is
          // scheduled; waiting until the timer fires leaves a stale-result gap.
          const gen = (this.generation += 1);
          this.timer = setTimeout(() => {
            this.timer = null;
            if (this.destroyed) return;
            void provideLenses()
              .then((lenses) => {
                // Superseded by a newer request, or torn down while the async
                // provider (dynamic TS import) was in flight → discard.
                if (this.destroyed || gen !== this.generation) return;
                view.dispatch({ effects: setCodeLensesEffect.of(lenses) });
              })
              .catch(() => {
                // Provider failures degrade to no change (keep the last good set).
              });
          }, debounceMs);
        }

        destroy() {
          this.destroyed = true;
          if (this.timer != null) {
            clearTimeout(this.timer);
            this.timer = null;
          }
        }
      },
    ),
  ];
}

/**
 * Build code-lens block decorations from raw {@link EditorState} (no view) — used
 * inside the StateField's `update`, where only state is available. Mirrors
 * {@link buildCodeLensDecorations} but reads line geometry from `state.doc`.
 * Whole-document (never a viewport): there is no viewport at state-update time and
 * lenses are few, so building the full set here is both correct and cheap.
 */
function buildCodeLensFromState(
  totalLines: number,
  state: EditorState,
  lenses: CodeLens[],
  onClick: (nameOffset: number) => void,
  formatLabel: (referenceCount: number) => string,
): DecorationSet {
  const builder = new RangeSetBuilder<Decoration>();
  const ordered = [...lenses].sort((a, b) => a.line - b.line);
  let lastLineFrom = -1;
  for (const lens of ordered) {
    if (lens.line < 1 || lens.line > totalLines) continue;
    const line = state.doc.line(lens.line);
    if (line.from === lastLineFrom) continue;
    lastLineFrom = line.from;
    const label = formatLabel(lens.refCount);
    builder.add(
      line.from,
      line.from,
      Decoration.widget({
        widget: new CodeLensWidget(
          label,
          lens.refCount,
          lens.nameOffset,
          onClick,
        ),
        block: true,
        side: -1,
      }),
    );
  }
  return builder.finish();
}

// --- Inline Peek (VSCode "Peek Definition" / "Peek References") --------------
// An INLINE embedded peek: rather than jumping, the editor opens a small block
// widget DIRECTLY BELOW the caret line that shows either (a) the target
// definition's source lines (read-only) or (b) a compact list of every
// reference with a one-line code preview and click-to-jump. The widget is a CM6
// BLOCK decoration (so it adds NO characters to the document — copy/paste,
// `onChange`, and the serialized buffer are byte-for-byte unaffected) owned by a
// StateField fed via {@link setPeekEffect}. Closed by default: until an effect
// carries content the field holds `Decoration.none`, so the editor is unchanged.

/** One reference row inside a Peek References widget. */
export interface PeekReferenceItem {
  /** 1-based source line number the reference starts on (shown in the gutter). */
  line: number;
  /** Trimmed source-line text shown as the one-line preview. */
  preview: string;
  /** Optional cell label for a cross-file (cross-cell) entry. */
  cellLabel?: string;
  /** Index into the editor's reference list (passed back to `onJump`). */
  index: number;
}

/**
 * The fully-formed content the editor hands the peek widget. Built in React
 * (where the data + handlers live) and pushed through {@link setPeekEffect}; the
 * widget renders it imperatively (it lives outside React inside CM6's DOM).
 */
export interface PeekContent {
  /** 1-based line BELOW which the peek renders (the caret line at trigger). */
  anchorLine: number;
  /** Distinguishes the definition view from the references list. */
  mode: "definition" | "references";
  /** Heading text (e.g. "Peek Definition" / "Peek References"). */
  title: string;
  /** Accessible label for the close button. */
  closeLabel: string;
  /** Dark-mode flag so the widget matches the editor theme. */
  isDark: boolean;
  /** Definition mode: the read-only target lines to display. */
  defLines?: string[];
  /** Definition mode: 1-based source line number of the first `defLines` entry. */
  defStartLine?: number;
  /** References mode: the rows to list. */
  refs?: PeekReferenceItem[];
  /** References mode: full result count before the bounded inline projection. */
  refsTotal?: number;
  /** References mode: visible notice for omitted rows. */
  moreLabel?: string;
  /** References mode: jump the editor selection to the row's reference. */
  onJump?: (index: number) => void;
  /** Close the peek (clears the widget + restores editor focus). */
  onClose: () => void;
}

/**
 * The DOM widget for the inline peek. Mirrors {@link CodeLensWidget}'s discipline
 * (textContent never innerHTML; its own event handlers; `ignoreEvent` false so
 * clicks reach the rows rather than being treated as edits). `eq` compares the
 * identity of the content object so CM6 reuses the node across unrelated
 * recomputes but rebuilds when the editor pushes a fresh peek.
 */
class PeekWidget extends WidgetType {
  constructor(readonly content: PeekContent) {
    super();
  }

  eq(other: PeekWidget): boolean {
    return other.content === this.content;
  }

  toDOM(): HTMLElement {
    const c = this.content;
    const wrap = document.createElement("div");
    wrap.className = "cm-peek";
    wrap.setAttribute("data-testid", "code-peek");
    wrap.setAttribute("data-peek-mode", c.mode);
    wrap.setAttribute("role", "dialog");
    wrap.setAttribute("aria-label", c.title);

    const header = document.createElement("div");
    header.className = "cm-peek-header";
    const heading = document.createElement("span");
    heading.className = "cm-peek-title";
    heading.textContent = c.title;
    header.appendChild(heading);
    const close = document.createElement("button");
    close.type = "button";
    close.className = "cm-peek-close";
    close.setAttribute("data-testid", "code-peek-close");
    close.setAttribute("aria-label", c.closeLabel);
    close.title = c.closeLabel;
    close.textContent = "×"; // ×
    close.addEventListener("mousedown", (event) => {
      // Keep pointer focus from moving the editor caret. `click` below is the
      // shared native-button activation path for mouse and Enter/Space.
      event.preventDefault();
      event.stopPropagation();
    });
    close.addEventListener("click", (event) => {
      event.preventDefault();
      event.stopPropagation();
      c.onClose();
    });
    header.appendChild(close);
    wrap.appendChild(header);

    const body = document.createElement("div");
    body.className = "cm-peek-body";
    if (c.mode === "definition") {
      const pre = document.createElement("div");
      pre.className = "cm-peek-def";
      pre.setAttribute("data-testid", "code-peek-definition");
      const start = c.defStartLine ?? 1;
      (c.defLines ?? []).forEach((text, i) => {
        const row = document.createElement("div");
        row.className = "cm-peek-line";
        const gutter = document.createElement("span");
        gutter.className = "cm-peek-gutter";
        gutter.textContent = String(start + i);
        const code = document.createElement("span");
        code.className = "cm-peek-code";
        // A blank line still needs height, so fall back to a NBSP.
        code.textContent = text.length > 0 ? text : " ";
        row.appendChild(gutter);
        row.appendChild(code);
        pre.appendChild(row);
      });
      body.appendChild(pre);
    } else {
      const list = document.createElement("div");
      list.className = "cm-peek-list";
      list.setAttribute("role", "list");
      (c.refs ?? []).forEach((ref) => {
        const item = document.createElement("button");
        item.type = "button";
        item.className = "cm-peek-item";
        item.setAttribute("role", "listitem");
        item.setAttribute("data-testid", "code-peek-reference");
        item.title = ref.preview;
        if (ref.cellLabel) {
          const cell = document.createElement("span");
          cell.className = "cm-peek-cell";
          cell.textContent = ref.cellLabel;
          item.appendChild(cell);
        }
        const gutter = document.createElement("span");
        gutter.className = "cm-peek-gutter";
        gutter.textContent = String(ref.line);
        const code = document.createElement("span");
        code.className = "cm-peek-code";
        code.textContent = ref.preview.length > 0 ? ref.preview : " ";
        item.appendChild(gutter);
        item.appendChild(code);
        item.addEventListener("mousedown", (event) => {
          event.preventDefault();
          event.stopPropagation();
        });
        item.addEventListener("click", (event) => {
          event.preventDefault();
          event.stopPropagation();
          c.onJump?.(ref.index);
        });
        list.appendChild(item);
      });
      if ((c.refsTotal ?? 0) > (c.refs?.length ?? 0) && c.moreLabel) {
        const more = document.createElement("div");
        more.className = "cm-peek-more";
        more.setAttribute("data-testid", "code-peek-reference-more");
        more.setAttribute("role", "status");
        more.textContent = c.moreLabel;
        list.appendChild(more);
      }
      body.appendChild(list);
    }
    wrap.appendChild(body);
    return wrap;
  }

  ignoreEvent(): boolean {
    return false;
  }
}

/** Faint peek styling via CSS vars so the theme can tune per light/dark. */
const peekTheme = EditorView.baseTheme({
  ".cm-peek": {
    margin: "2px 0 2px 6px",
    border: "1px solid var(--cm-peek-border, rgba(127,127,127,0.35))",
    borderRadius: "4px",
    background: "var(--cm-peek-bg, rgba(127,127,127,0.06))",
    fontSize: "0.92em",
    maxHeight: "220px",
    overflow: "auto",
    userSelect: "none",
  },
  ".cm-peek-header": {
    display: "flex",
    alignItems: "center",
    gap: "6px",
    padding: "3px 6px",
    borderBottom: "1px solid var(--cm-peek-border, rgba(127,127,127,0.35))",
    fontSize: "0.85em",
    fontWeight: "600",
  },
  ".cm-peek-title": { flex: "1 1 auto" },
  ".cm-peek-close": {
    cursor: "pointer",
    border: "none",
    background: "transparent",
    color: "inherit",
    fontSize: "1.1em",
    lineHeight: "1",
    padding: "0 4px",
  },
  ".cm-peek-body": { padding: "2px 0" },
  ".cm-peek-line": { display: "flex", gap: "8px", padding: "0 8px", whiteSpace: "pre" },
  ".cm-peek-item": {
    display: "flex",
    gap: "8px",
    width: "100%",
    textAlign: "left",
    border: "none",
    background: "transparent",
    color: "inherit",
    cursor: "pointer",
    padding: "1px 8px",
    whiteSpace: "nowrap",
  },
  ".cm-peek-item:hover": { background: "var(--cm-peek-hover, rgba(37,99,235,0.16))" },
  ".cm-peek-gutter": {
    color: "var(--cm-peek-gutter, #94a3b8)",
    minWidth: "32px",
    textAlign: "right",
    fontVariantNumeric: "tabular-nums",
  },
  ".cm-peek-cell": {
    color: "var(--cm-peek-cell, #6366f1)",
    minWidth: "48px",
    fontWeight: "600",
    overflow: "hidden",
    textOverflow: "ellipsis",
  },
  ".cm-peek-code": { overflow: "hidden", textOverflow: "ellipsis" },
  ".cm-peek-more": {
    padding: "3px 8px",
    color: "var(--cm-peek-gutter, #64748b)",
    fontSize: "0.85em",
  },
});

/** Effect carrying a fresh peek payload (or `null` to close) into the field. */
export const setPeekEffect = StateEffect.define<PeekContent | null>();

/**
 * Inline Peek extension: a StateField owning at most ONE block-widget decoration
 * rendered below the peeked line, fed by {@link setPeekEffect}. Block decorations
 * must come from a StateField (CM6 forbids them from a ViewPlugin), so the editor
 * dispatches the effect and the field renders. ALWAYS mountable — it holds
 * `Decoration.none` until an effect arrives, so the editor stays byte-for-byte
 * unchanged until the user opens a peek. ANY document change clears the peek (its
 * listed offsets/lines may have shifted), mirroring the references panel's
 * stale-offset guard.
 */
export function peekExtension(): Extension {
  const peekField = StateField.define<DecorationSet>({
    create: () => Decoration.none,
    update(decorations, tr) {
      // A document edit invalidates the peeked line/offsets → clear the peek.
      if (tr.docChanged) {
        let cleared = false;
        for (const effect of tr.effects) {
          if (effect.is(setPeekEffect) && effect.value) cleared = true;
        }
        if (!cleared) return Decoration.none;
      }
      let next = decorations.map(tr.changes);
      for (const effect of tr.effects) {
        if (effect.is(setPeekEffect)) {
          const content = effect.value;
          if (!content) {
            next = Decoration.none;
            continue;
          }
          const total = tr.state.doc.lines;
          const lineNo = Math.min(Math.max(content.anchorLine, 1), total);
          const line = tr.state.doc.line(lineNo);
          next = Decoration.set([
            Decoration.widget({
              widget: new PeekWidget(content),
              block: true,
              side: 1,
            }).range(line.to),
          ]);
        }
      }
      return next;
    },
    provide: (field) => EditorView.decorations.from(field),
  });

  return [peekTheme, peekField];
}

/**
 * Tree-free code folding for the stream-mode languages (go / rust / java / c /
 * cpp / csharp / ruby / swift / r / perl / lua / yaml / toml / bash),
 * which are highlighted by `@codemirror/legacy-modes` and produce NO Lezer tree
 * — so CM6's default (tree-based) folding never folds them even though the fold
 * gutter renders. CM6 consults the `foldService` facet BEFORE the syntax tree,
 * and a foldService callback needs only the document text + a line's offsets, so
 * it gives these languages real folding.
 *
 * Two providers are registered, tried in CM6's facet order:
 *  1. EXPLICIT `#region` / `#endregion` markers (VSCode marker folding) via
 *     {@link computeRegionFoldRange} — checked FIRST so a region marker that also
 *     opens an indented block folds the whole region (to its `#endregion`),
 *     not just the indentation.
 *  2. INDENTATION via {@link computeIndentFoldRange} — a header line whose body
 *     is more deeply indented folds that body.
 *
 * Both are pure offset math (the heavy logic lives in stream-fold.ts); this is a
 * thin adapter from CM6's `(state, lineStart, lineEnd)` to those functions. The
 * extension is mounted ONLY for the stream-mode languages (CodeMirrorEditor
 * gates it on the language), so on every tree-backed language (js/ts/python/…),
 * whose own folding already works, it is absent and nothing changes.
 *
 * @param tabSize Columns per tab stop for the indentation measure (default 2,
 *                the editor's default indent width).
 */
/**
 * The document as one string, remembered per Text.
 *
 * CM6 asks the fold services about every visible line each time the fold gutter
 * renders, and each ask used to allocate a fresh copy of the whole document —
 * half a megabyte per query on a large file. `Text` is immutable, so one entry
 * per document is exact, and a WeakMap lets it die with the document rather
 * than pinning old revisions.
 */
const documentTextCache = new WeakMap<Text, string>();

function documentText(doc: Text): string {
  let text = documentTextCache.get(doc);
  if (text === undefined) {
    text = doc.toString();
    documentTextCache.set(doc, text);
  }
  return text;
}

export function streamFoldExtension(tabSize = 2): Extension {
  return [
    // Region markers take priority over indentation.
    regionFoldExtension(),
    foldService.of((state, lineStart) =>
      computeIndentFoldRange(documentText(state.doc), lineStart, tabSize),
    ),
  ];
}

/**
 * Tree-COMPATIBLE `#region` / `#endregion` marker folding (VSCode "marker"
 * folding) for the TREE-BACKED languages (js / ts / jsx / tsx / python / css /
 * html / json / php / sql / markdown). Those already fold via their Lezer
 * grammar's `foldNodeProp`, so they get the INDENTATION / syntax fold for free —
 * but CM6's default tree folding does NOT understand the `// #region` /
 * `# region` / `/* #region *\/` / `<!-- #region -->` marker convention, so a
 * region that spans siblings at the SAME nesting level (e.g. a run of top-level
 * statements between `// #region` and `// #endregion`) never folds there.
 *
 * This mounts ONLY the region-marker {@link computeRegionFoldRange} foldService
 * (NOT the indentation one — that would duplicate / fight the tree's own
 * indentation folding). CM6 consults ALL `foldService` callbacks BEFORE the
 * syntax tree and uses the FIRST non-null result, so:
 *  - on a `#region` start line, this returns the marker range → the whole region
 *    (to its matching `#endregion`) folds, exactly like VSCode;
 *  - on every other line, this returns null → CM6 falls through to the language
 *    grammar's own (tree-based) folding, byte-for-byte unchanged.
 *
 * So composing this with a tree-backed language ADDS marker folding without
 * disturbing the existing syntax folding. (Verified: the JS/TS/Python tree fold
 * still folds braces/indented blocks with this mounted; only the marker case is
 * new.) Mounted by CodeMirrorEditor for the tree-backed languages; the
 * stream-mode languages get the SAME region service via {@link
 * streamFoldExtension} (which also adds indentation folding, since they have no
 * tree), so the marker behaviour is identical across every language.
 */
export function regionFoldExtension(): Extension {
  return foldService.of((state, lineStart) =>
    computeRegionFoldRange(documentText(state.doc), lineStart),
  );
}

/**
 * Inline `background` style for ONE line drawing a 1px vertical rule at each
 * ruler `column` (VSCode `editor.rulers`). Mirrors {@link indentGuideLineStyle}:
 * a stacked `linear-gradient` per column, positioned `column` character cells
 * from the text origin (the theme sets `background-origin: content-box`), so the
 * rules land on the monospace character grid regardless of `.cm-line` padding.
 * Columns are de-duplicated and sorted so a config like `[80, 80, 120]` draws
 * two rules, not three.
 */
export function rulerLineStyle(columns: number[]): string {
  const color = "var(--cm-ruler-color, rgba(127, 127, 127, 0.22))";
  const unique = Array.from(
    new Set(columns.filter((c) => Number.isFinite(c) && c > 0).map((c) => Math.floor(c))),
  ).sort((a, b) => a - b);
  if (unique.length === 0) return "";
  const images = unique.map(() => `linear-gradient(${color}, ${color})`).join(", ");
  const positions = unique.map((c) => `${c}ch 0`).join(", ");
  const sizes = unique.map(() => "1px 100%").join(", ");
  return `background-image:${images};background-position:${positions};background-size:${sizes};`;
}

/** Theme for the ruler rules — same content-box origin trick as indent guides. */
const rulerTheme = EditorView.baseTheme({
  ".cm-line.cm-rulers": {
    backgroundRepeat: "no-repeat",
    backgroundOrigin: "content-box",
  },
});

/** Build ruler line decorations for every line in the visible ranges. */
function buildRulerDecorations(view: EditorView, columns: number[]): DecorationSet {
  const style = rulerLineStyle(columns);
  if (!style) return Decoration.none;
  const builder = new RangeSetBuilder<Decoration>();
  const doc = view.state.doc;
  for (const range of view.visibleRanges) {
    let pos = range.from;
    while (pos <= range.to) {
      const line = doc.lineAt(pos);
      builder.add(
        line.from,
        line.from,
        Decoration.line({ class: "cm-rulers", attributes: { style } }),
      );
      pos = line.to + 1;
    }
  }
  return builder.finish();
}

/**
 * Vertical column rulers (VSCode `editor.rulers`): faint full-height guide lines
 * at the given 1-based columns (e.g. 80 / 120) to mark line-length limits. Drawn
 * as per-line `background` gradients (like the indent guides) so the source
 * bytes are never touched and the rules never interfere with the caret or
 * selection. Recomputed on document / viewport change.
 *
 * Mounted by CodeMirrorEditor ONLY when the ruler-columns setting is non-empty,
 * so an editor with rulers off is byte-for-byte unchanged.
 *
 * @param columns 1-based ruler columns; empty / non-positive entries are ignored
 *                and the extension renders nothing.
 */
export function rulersExtension(columns: number[]): Extension {
  return [
    rulerTheme,
    ViewPlugin.fromClass(
      class {
        decorations: DecorationSet;

        constructor(view: EditorView) {
          this.decorations = buildRulerDecorations(view, columns);
        }

        update(update: ViewUpdate) {
          if (update.docChanged || update.viewportChanged) {
            this.decorations = buildRulerDecorations(update.view, columns);
          }
        }
      },
      { decorations: (plugin) => plugin.decorations },
    ),
  ];
}

// --- Inline CSS colour decorators (VSCode parity) ----------------------------
// A clickable colour SWATCH rendered immediately before every `#hex` / `rgb()` /
// `hsl()` literal in a CSS/SCSS/LESS buffer; clicking it opens the browser's
// native colour input, and the chosen value is spliced back over the literal in
// the SAME syntax family (the pure scan/parse/replace logic lives in the
// unit-tested css-color-decorators.ts). The swatch is a `Decoration.widget` that
// adds NO characters to the document (copy/paste, `onChange`, and the serialized
// buffer are byte-for-byte unaffected) — it is pure inline chrome.

/**
 * The DOM widget for ONE colour swatch: a small square showing the literal's
 * colour, rendered just BEFORE the literal. Mirrors {@link InlayHintWidget}'s
 * discipline — no `innerHTML`, its own click handler, `ignoreEvent` false so the
 * click reaches the handler. `eq` collapses two swatches with the same colour +
 * position so CM6 reuses the node across recomputes (no flicker). Clicking calls
 * `onActivate(from, to, value, swatchEl)` (the plugin opens the native picker).
 */
class ColorSwatchWidget extends WidgetType {
  constructor(
    readonly hex: string,
    readonly from: number,
    readonly to: number,
    readonly value: string,
    readonly onActivate: (
      from: number,
      to: number,
      value: string,
      anchor: HTMLElement,
    ) => void,
  ) {
    super();
  }

  eq(other: ColorSwatchWidget): boolean {
    // The handler is stable for the editor's lifetime; the VISIBLE state is the
    // colour + the span it edits, so compare only those.
    return (
      other.hex === this.hex &&
      other.from === this.from &&
      other.to === this.to &&
      other.value === this.value
    );
  }

  toDOM(): HTMLElement {
    const swatch = document.createElement("span");
    swatch.className = "cm-color-swatch";
    // aria-hidden, and NOT role="button": this is a mouse-only convenience on
    // top of a colour literal the reader already hears verbatim. It had
    // role="button" with a hardcoded English label and no tabindex, so a screen
    // reader in the Chinese UI announced an English "button" on every colour in
    // the file that no keyboard could ever focus or activate — a promise the
    // widget could not keep, in the wrong language.
    //
    // Same call, and the same reason, as CodeLensWidget above: an affordance
    // that duplicates nothing and cannot be reached must not add noise to the
    // reading order of the code itself. A CM6 widget also has no translator in
    // scope, so removing the announcement is the honest fix rather than
    // translating a label for a control that does not work.
    swatch.setAttribute("aria-hidden", "true");
    swatch.title = this.value;
    swatch.setAttribute("data-color-swatch", this.hex);
    swatch.style.backgroundColor = this.hex;
    swatch.addEventListener("mousedown", (event) => {
      // Don't let the click move the editor caret / start a selection; just open
      // the colour picker for this literal.
      event.preventDefault();
      event.stopPropagation();
      this.onActivate(this.from, this.to, this.value, swatch);
    });
    return swatch;
  }

  // The swatch has its own click handler; let the event reach it rather than be
  // swallowed/handled as an editor interaction.
  ignoreEvent(): boolean {
    return false;
  }
}

/** Swatch styling via CSS vars so the theme can tune the border per light/dark. */
const colorSwatchTheme = EditorView.baseTheme({
  ".cm-color-swatch": {
    display: "inline-block",
    boxSizing: "border-box",
    width: "0.8em",
    height: "0.8em",
    marginRight: "0.35em",
    verticalAlign: "middle",
    borderRadius: "3px",
    border: "1px solid var(--cm-color-swatch-border, rgba(127, 127, 127, 0.6))",
    cursor: "pointer",
    // The swatch is not part of the document text — never let it be selected.
    userSelect: "none",
  },
});

/**
 * Build the colour-swatch decoration set for the view's CURRENT viewport only.
 * Literals are scanned PER visible range (so a large file only scans rendered
 * text), and each becomes an atomic `Decoration.widget` anchored at the literal's
 * `from` with `side: -1` so the swatch renders just BEFORE the literal. Every
 * offset is clamped into the live doc so a stale scan can never throw. Widgets are
 * added in ascending offset order (RangeSetBuilder requires it).
 */
function buildColorSwatchDecorations(
  view: EditorView,
  onActivate: (from: number, to: number, value: string, anchor: HTMLElement) => void,
): DecorationSet {
  const doc = view.state.doc;
  const docLength = doc.length;
  const collected: { from: number; deco: Decoration }[] = [];
  for (const range of view.visibleRanges) {
    const text = doc.sliceString(range.from, range.to);
    const literals = scanColorLiterals(text);
    for (const lit of literals) {
      const from = Math.min(Math.max(range.from + lit.from, 0), docLength);
      const to = Math.min(Math.max(range.from + lit.to, from), docLength);
      const hex = colorLiteralToHex(lit.value);
      if (!hex) continue;
      const deco = Decoration.widget({
        widget: new ColorSwatchWidget(hex, from, to, lit.value, onActivate),
        side: -1,
      });
      collected.push({ from, deco });
    }
  }
  collected.sort((a, b) => a.from - b.from);
  const builder = new RangeSetBuilder<Decoration>();
  for (const { from, deco } of collected) {
    builder.add(from, from, deco);
  }
  return builder.finish();
}

/**
 * Inline CSS colour decorators (VSCode parity): a clickable colour swatch before
 * every round-trippable `#hex` / `rgb()` / `hsl()` literal, opening the product's
 * colour picker (through a native colour field) on click and splicing the chosen
 * colour back into the same syntax family. Recomputes on document / viewport
 * change.
 *
 * The caller mounts this ONLY for css/scss/less (CodeMirrorEditor gates it on the
 * language), so on every other language the extension is simply absent and the
 * editor is byte-for-byte unchanged. The swatches add NO document characters
 * (they are decorations), so the serialized buffer is unaffected until the user
 * actually picks a new colour.
 *
 * PICKER SESSION: clicking a swatch creates a hidden `<input type="color">`
 * appended to the editor's DOM, seeded with the literal's colour. The native
 * picker fires `input` continuously (live preview) and `change` on commit; each
 * event splices the new colour over the CURRENT literal span. Because each splice
 * changes the literal's length, the session tracks the live `[from, to]` so the
 * next `input` event replaces the just-written literal — not a stale span. The
 * input is removed when the picker closes (`change`, or `blur`).
 *
 * SAFETY: every replace is a single CM6 transaction (one undo step per event is
 * acceptable — VSCode coalesces, but correctness first); the replacement text is
 * produced by the pure {@link replaceColorLiteral} and validated, so a malformed
 * picker value (never expected from a native input) leaves the buffer untouched.
 */
export function colorDecoratorExtension(): Extension {
  return [
    colorSwatchTheme,
    ViewPlugin.fromClass(
      class {
        decorations: DecorationSet;
        // The CM6 view this plugin instance is attached to (captured in the
        // constructor — ViewPlugin passes it there, not as an auto-set property —
        // so the swatch activation handler can dispatch the replace transaction).
        private readonly view: EditorView;
        // The live input element for the current picker session, if open.
        private input: HTMLInputElement | null = null;
        // The literal span the current session is editing (updated after each
        // splice so successive live `input` events target the fresh literal).
        private session: { from: number; to: number } | null = null;
        private destroyed = false;

        constructor(view: EditorView) {
          this.view = view;
          this.decorations = buildColorSwatchDecorations(view, this.activate);
        }

        update(update: ViewUpdate) {
          if (update.docChanged || update.viewportChanged) {
            this.decorations = buildColorSwatchDecorations(update.view, this.activate);
          }
          // If a doc change came from ELSEWHERE while a picker session is open,
          // map the tracked span through it so the next live splice stays aligned.
          if (update.docChanged && this.session) {
            const from = update.changes.mapPos(this.session.from, 1);
            const to = update.changes.mapPos(this.session.to, -1);
            this.session = { from, to: Math.max(from, to) };
          }
        }

        // Open (or refocus) the native colour input for the activated literal.
        // Bound as an arrow property so it carries `this` when passed to widgets.
        private activate = (
          from: number,
          to: number,
          value: string,
          anchor: HTMLElement,
        ): void => {
          if (this.destroyed) return;
          const view = this.view;
          this.closeInput();
          const seed = colorLiteralToHex(value) ?? "#000000";
          const input = document.createElement("input");
          input.type = "color";
          input.value = seed;
          // Keep it out of layout/AT but still focusable + clickable.
          input.setAttribute("aria-hidden", "true");
          input.style.position = "absolute";
          input.style.width = "1px";
          input.style.height = "1px";
          input.style.opacity = "0";
          input.style.pointerEvents = "none";
          // Near the swatch, for a browser picker that opens at its field.
          const rect = anchor.getBoundingClientRect();
          const hostRect = view.dom.getBoundingClientRect();
          input.style.left = `${rect.left - hostRect.left}px`;
          input.style.top = `${rect.bottom - hostRect.top}px`;
          view.dom.appendChild(input);
          this.input = input;
          this.session = { from, to };

          input.addEventListener("input", () => this.applyPicked(input.value));
          input.addEventListener("change", () => {
            this.applyPicked(input.value);
            this.closeInput();
          });
          // Closing the OS picker without committing blurs the input → clean up.
          input.addEventListener("blur", () => this.closeInput());

          // Activate the field the way a press would: the product's colour
          // picker (NativeFieldPopups → ColorPickerPopover) takes the click and
          // opens against the swatch, the one picker every colour field uses;
          // where that layer stands aside (touch), the browser's own picker
          // opens instead. Not `showPicker()`, which goes straight to the OS
          // picker past that layer. We ARE inside a mousedown handler, so the
          // click carries the user's activation.
          try {
            input.click();
          } catch {
            /* picker unavailable — leave the buffer untouched */
          }
        };

        // Splice the picked colour over the session's CURRENT literal span, then
        // advance the tracked span to cover the just-written text so the next live
        // `input` event edits the fresh literal rather than a stale range.
        private applyPicked(pickedHex: string): void {
          const view = this.view;
          if (!this.session) return;
          const { from, to } = this.session;
          const docLength = view.state.doc.length;
          if (from < 0 || to > docLength || from > to) return;
          const original = view.state.doc.sliceString(from, to);
          const replacement = replaceColorLiteral(original, pickedHex);
          // Null (malformed pick) or a no-op replacement → leave the buffer alone.
          if (replacement === null || replacement === original) return;
          view.dispatch({
            changes: { from, to, insert: replacement },
            // No selection change — keep the caret where the user had it.
          });
          // The literal now occupies [from, from + replacement.length].
          this.session = { from, to: from + replacement.length };
        }

        private closeInput(): void {
          if (this.input) {
            this.input.remove();
            this.input = null;
          }
          this.session = null;
        }

        destroy() {
          this.destroyed = true;
          this.closeInput();
        }
      },
      {
        decorations: (plugin) => plugin.decorations,
        // Atomic: arrow keys / selection step OVER the swatch (it has no doc
        // content) rather than landing "inside" it, matching the other widgets.
        provide: (plugin) =>
          EditorView.atomicRanges.of(
            (view) => view.plugin(plugin)?.decorations ?? Decoration.none,
          ),
      },
    ),
  ];
}
