/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Flag-gated JavaScript breakpoint debugging via lightweight SOURCE
 * INSTRUMENTATION (no new dependency).
 *
 * GOAL (JS/TS only — Python/Pyodide debugging is OUT OF SCOPE; see the note in
 * `runJavaScript`): let the user set line breakpoints, "Run in debug", and pause
 * when execution reaches a breakpoint, showing the current line + a best-effort
 * snapshot of in-scope variables, with Continue / Step.
 *
 * HOW IT WORKS
 * ------------
 * Before running, {@link instrumentForDebug} transforms the source by inserting a
 *   `await __mrnoteDebugHook(<line>, () => ({ ...locals }));`
 * statement at SAFE statement boundaries. The hook, supplied by the worker, is a
 * no-op fast-path unless `<line>` has a breakpoint (or a single-step is pending),
 * in which case it posts a `paused` message to the main thread and AWAITS a
 * resume/step reply before returning. Because the hook is awaited, debug runs
 * execute the user code inside an `async` wrapper (the runner handles this); the
 * NORMAL (non-debug) run path is never instrumented and never wrapped.
 *
 * WHY A HAND-ROLLED LEXER (and its exact, conservative limits)
 * -----------------------------------------------------------
 * Reliable statement-boundary detection of arbitrary JS without a parser is the
 * hard part. We do NOT parse; we LEX. {@link instrumentForDebug} scans the source
 * character-by-character, correctly skipping strings, template literals
 * (including `${…}` interpolations), regex literals, and comments, and tracking
 * bracket nesting with a brace-kind stack (block `{}` vs object/expression `{}`).
 *
 * A line is instrumented ONLY when its first significant character sits at a
 * place where a fresh *statement* may legally begin AND prepending a statement is
 * safe, i.e. ALL of:
 *   - paren depth == 0 and square-bracket depth == 0 (never inside `(...)` /
 *     `[...]`), and
 *   - we are inside a BLOCK brace context (program top level or a `{}` the lexer
 *     classified as a block, never an object literal), and
 *   - the previous significant token is one of `{ } ;` or start-of-program, and
 *   - the line does not start with a token that would be broken by an inserted
 *     statement before it (`case`/`default`/labels/`else`/`catch`/`finally`/
 *     `while` of a do-while/binding `,`/chained `.`?/etc. — see
 *     {@link UNSAFE_LINE_STARTS}).
 *
 * Everything else is LEFT ALONE: it simply gets no breakpoint hook, so a
 * breakpoint on such a line never triggers (degraded — the user can move the
 * breakpoint up to the enclosing statement — but NEVER incorrect). Correctness of
 * the user's program is paramount: if a single instrumentation point is unsafe we
 * skip it, and {@link instrumentForDebug} performs a self-check (see
 * `safeToInstrument`) and FALLS BACK to the original source verbatim if anything
 * looks off, so a debug run can never miscompile the code.
 *
 * The hook captures locals BEST-EFFORT: the instrumenter cannot see runtime
 * scope, so the locals thunk it emits reads a small, lexically-collected set of
 * simple binding names (`let`/`const`/`var`/params) that are in scope at that
 * line, each guarded so an unreachable/TDZ binding never throws. Depth/closure
 * capture is intentionally shallow. This is a snapshot aid, not a full debugger.
 */

/**
 * Tokens a `(` may follow that mean "control-flow header", not "method head".
 * The block after `if (…)`, `for (…)` and friends is an ordinary block that
 * inherits its enclosing async-ness; the block after anything else that looks
 * like a callable head (`m(…)`, `get x(…)`, `constructor(…)`) is a FUNCTION
 * body, where an awaited hook would not compile unless the function is async.
 */
const CONTROL_FLOW_HEADS: ReadonlySet<string> = new Set([
  "if",
  "for",
  "while",
  "switch",
  "catch",
  "with",
]);
const METHOD_HEAD_IDENTIFIER =
  /^[$_\p{ID_Start}][$\u200c\u200d\p{ID_Continue}]*$/u;

/** Whether a `(` following `token` opens a method/function parameter list. */
function isMethodHeadToken(token: string): boolean {
  if (!token || CONTROL_FLOW_HEADS.has(token)) return false;
  // An identifier (or a keyword usable as a method name, e.g. `get`/`delete`).
  return METHOD_HEAD_IDENTIFIER.test(token);
}

/** A line number (1-based) the user marked as a breakpoint. */
export type BreakpointLine = number;

/** Generous cap for a condition expression (one line of code, not a program). */
export const MAX_BREAKPOINT_CONDITION_LENGTH = 500;

/**
 * Pre-validate a breakpoint condition: bounded, single-expression JS that
 * PARSES. The condition is later baked into the instrumented program, so a
 * syntax error here would otherwise break the entire debug run — this is the
 * gate that keeps that impossible. Parsing only; it is never executed here.
 */
export function isValidBreakpointCondition(condition: string): boolean {
  if (typeof condition !== "string") return false;
  const trimmed = condition.trim();
  if (trimmed === "" || trimmed.length > MAX_BREAKPOINT_CONDITION_LENGTH) {
    return false;
  }
  try {
    // Parse-only (the function is never invoked). Wrapping in parens forces a
    // single EXPRESSION — statements / injection attempts like `1); evil((`
    // fail to parse inside the parens.
    new Function(`return (${trimmed});`);
    return true;
  } catch (error) {
    // A Content-Security-Policy that forbids eval makes the Function
    // constructor throw EvalError — the check could not RUN, which is not the
    // same as the condition being bad. Rejecting it there meant that in
    // production, where the app CSP carries no 'unsafe-eval', every condition a
    // user typed was flagged red and silently stripped before the run: a
    // feature that works in dev and nowhere else, and no test could see it
    // because jsdom has no CSP and acceptance runs against a dev server.
    //
    // So: unable to check means unable to condemn. A condition that survives
    // here is still handled safely downstream — the worker evaluates it inside
    // the same isolated realm as the program, and a broken one simply does not
    // match.
    if (error instanceof EvalError) return true;
    return false;
  }
}

/**
 * The payload the worker posts when execution pauses at a breakpoint: the 1-based
 * line, a flat best-effort snapshot of in-scope simple locals (display strings),
 * and the ordinal of this pause within the run (for the UI).
 */
export interface DebugPausePayload {
  readonly line: number;
  /** Best-effort `{ name: displayString }` of simple in-scope locals. */
  readonly locals: Readonly<Record<string, string>>;
  /** 1-based count of how many times we've paused so far this run. */
  readonly step: number;
}

/** The name of the hook the instrumented code calls (must be unique/odd). */
export const DEBUG_HOOK_NAME = "__mrnoteDebugHook";

/**
 * Result of {@link instrumentForDebug}: the (possibly) rewritten source plus the
 * set of line numbers that actually received a hook (so the UI can tell the user
 * which breakpoints are "live"). When `instrumented` is false the `code` is the
 * ORIGINAL source verbatim and `hookedLines` is empty.
 */
export interface InstrumentResult {
  readonly code: string;
  readonly instrumented: boolean;
  readonly hookedLines: readonly number[];
}

// Tokens that must NOT have a statement inserted immediately before them, because
// the inserted `await …;` would syntactically detach them from what they continue
// (a clause head, a label, a continuation of the previous statement, …).
const UNSAFE_LINE_STARTS: readonly string[] = [
  "case",
  "default",
  "else",
  "catch",
  "finally",
  "while", // could be the `while (...)` tail of a do-while
];

// Characters that, when they START a line's significant content, indicate the
// line CONTINUES the previous statement (operator/continuation) rather than
// begins a new one — inserting a statement before them would break the program.
const CONTINUATION_FIRST_CHARS = new Set([
  ".",
  "?",
  ":",
  ",",
  ")",
  "]",
  "}",
  "+",
  "-",
  "*",
  "/",
  "%",
  "<",
  ">",
  "=",
  "&",
  "|",
  "^",
  "`",
]);

/** Reserved/keyword line-leads that begin a statement but bind a name we collect. */
const DECLARATION_KEYWORDS = new Set(["let", "const", "var"]);

/**
 * ECMAScript identifier characters. Unicode property escapes match the
 * IdentifierStart / IdentifierPart sets JavaScript uses, with the language's
 * explicit `$`, `_`, ZWNJ, and ZWJ additions. Reading by code point (rather
 * than UTF-16 code unit) is important for supplementary-plane names such as
 * `𠮷野`.
 */
const IDENTIFIER_START = /^[$_\p{ID_Start}]$/u;
const IDENTIFIER_PART = /^[$\u200C\u200D\p{ID_Continue}]$/u;

function codePointAt(source: string, offset: number): string {
  const value = source.codePointAt(offset);
  return value === undefined ? "" : String.fromCodePoint(value);
}

function isIdentStartAt(source: string, offset: number): boolean {
  return IDENTIFIER_START.test(codePointAt(source, offset));
}

function isIdentPartAt(source: string, offset: number): boolean {
  return IDENTIFIER_PART.test(codePointAt(source, offset));
}

function readIdentifierEnd(source: string, offset: number): number {
  let end = offset;
  while (end < source.length && isIdentPartAt(source, end)) {
    end += codePointAt(source, end).length;
  }
  return end;
}

/** The snapshot and injected thunk are deliberately capped together. */
const MAX_LOCALS_PER_LINE = 24;

/**
 * A brace-kind marker on the lexer's stack. We classify each `{` as either a
 * statement BLOCK (so statements — and our hook — may appear directly inside) or
 * an OBJECT/expression brace (where a bare statement is illegal). The
 * classification uses the previous significant token, the standard heuristic.
 */
type BraceKind = "block" | "object";

/**
 * One open-brace frame. Besides the block/object kind, it records `asyncContext`:
 * whether code DIRECTLY inside this brace runs where `await` is legal. This is the
 * crux of correctness: our hook is `await`-ed, so it may ONLY be inserted where
 * `await` is syntactically valid. The frame inherits the enclosing async-ness for
 * ordinary blocks (if/for/while/bare `{}`) and RESETS to whatever we can prove for
 * a FUNCTION-BODY block: `true` only when we positively detected `async`
 * (`async function`, `async (...) =>`, `async x =>`), else `false`. When unsure
 * we choose `false` (no awaited hook there) — never a miscompile.
 */
interface BraceFrame {
  readonly kind: BraceKind;
  readonly asyncContext: boolean;
}

interface LexState {
  /** Stack of currently-open braces (innermost last). */
  readonly braceStack: BraceFrame[];
  /** Depth of currently-open `(` parentheses. */
  parenDepth: number;
  /** Depth of currently-open `[` brackets. */
  bracketDepth: number;
}

/**
 * Whether the lexer is currently in a context where a fresh STATEMENT may legally
 * begin and our hook can be safely inserted: not inside parens/brackets, and the
 * innermost brace (if any) is a block. Program top level (empty brace stack)
 * qualifies.
 */
function inStatementContext(state: LexState): boolean {
  if (state.parenDepth !== 0 || state.bracketDepth !== 0) return false;
  const top = state.braceStack[state.braceStack.length - 1];
  return top === undefined || top.kind === "block";
}

/**
 * Whether `await` is legal at the current top-of-stack position (i.e. an awaited
 * hook can be inserted). True at program top level (empty stack — the runner wraps
 * the body in an async function) and inside any frame whose `asyncContext` holds.
 */
function inAsyncContext(state: LexState): boolean {
  const top = state.braceStack[state.braceStack.length - 1];
  return top === undefined ? true : top.asyncContext;
}

/**
 * Decide whether a `{` opens a statement block or an object literal, from the
 * previous significant token text (`prevToken`). The SAFE bias is to tag
 * "object" whenever the `{` could be an expression/object position — i.e. when
 * the previous token is one an expression can follow with `{`: `(`, `[`, `,`,
 * `:`, `=`, operators, keywords like `return`/`typeof`, or start-of-input. Only
 * when the previous token clearly ends a statement/clause-head (`)`, `;`, `{`,
 * `}`, `=>`, `else`, `do`, `try`, `finally`) do we tag "block". Tagging a real
 * object as a block is the dangerous direction (could enable an illegal insert),
 * so anything ambiguous stays "object".
 */
function classifyBraceKind(prevToken: string): BraceKind {
  const blockAfter = new Set([")", ";", "{", "}", "=>", "do", "else", "try", "finally"]);
  return blockAfter.has(prevToken) ? "block" : "object";
}

/**
 * Significant-token classifier sufficient for our heuristics. Returns the token
 * TEXT for keywords/identifiers and punctuators we care about; for everything
 * else returns the single punctuator character. Whitespace/comments are skipped
 * by the caller before this is consulted.
 */

interface LineInfo {
  /** 1-based line number. */
  readonly line: number;
  /** Char offset (into the source) of the first significant char on the line. */
  readonly firstSignificantOffset: number;
  /** The previous significant token text seen just before that char. */
  readonly prevToken: string;
  /** Lexer context (statement-able?) at that first significant char. */
  readonly statementContext: boolean;
  /** Whether `await` is legal here (top level or inside a proven-async function). */
  readonly asyncContext: boolean;
  /** The identifier/keyword that the line's significant content STARTS with (if any). */
  readonly leadWord: string;
  /** The first significant character of the line. */
  readonly leadChar: string;
}

interface ScanResult {
  /** Per-line metadata for every line that HAS significant content. */
  readonly lines: LineInfo[];
  /** Simple binding names declared at top level / per line, in source order. */
  readonly bindingsBefore: Map<number, string[]>;
  /** True if the lexer completed cleanly (balanced, no unterminated literal). */
  readonly clean: boolean;
}

/**
 * Lex the source once, recording for every line: its first significant char
 * offset, the previous significant token, whether a statement may begin there,
 * and the leading word/char. Also collects simple top-level binding names seen so
 * far (for the locals snapshot). Robust to strings, template literals with nested
 * `${}`, regex literals, and comments.
 */
function scanSource(source: string): ScanResult {
  const lines: LineInfo[] = [];
  const bindings: string[] = [];
  const bindingsBefore = new Map<number, string[]>();

  const state: LexState = { braceStack: [], parenDepth: 0, bracketDepth: 0 };
  // Template-literal interpolation stack: when we enter `${`, we push the current
  // brace-stack length so the matching `}` returns us to template mode.
  const templateBraceMarks: number[] = [];

  // --- Function-body + async-context tracking (for safe awaited-hook insertion).
  // We need to know, for each block `{`, whether code directly inside it runs in
  // an async context (await legal). A block is a FUNCTION BODY when it follows a
  // function header (`function (...)`, `name(...)` method, or `=>`); otherwise it
  // inherits the enclosing async-ness. We detect async positively (`async`
  // keyword) and otherwise assume non-async (the SAFE direction).
  //
  // `parenStack` mirrors paren nesting so we know when a `)` closes a FUNCTION
  // param list (its matching `(` was tagged "func-head"), making the next `{` a
  // function body. `funcHeadAsync` carries the `async`-ness of that header to the
  // body `{`. `arrowAsync` carries `async`-ness across an `=>` to its body `{`.
  const parenStack: Array<{
    kind: "func-head" | "other";
    async: boolean;
    /** The significant token this "(" followed — see the `(` case. */
    head: string;
  }> = [];
  let pendingFuncHeader = false; // saw `function` (maybe with name) awaiting `(`
  let pendingFuncAsync = false; // was that `function` prefixed by `async`?
  let sawAsyncKeyword = false; // the immediately-preceding token was `async`
  // One-token lookahead latch: set when a func-head `)` or an `=>` is seen; valid
  // ONLY for the IMMEDIATELY following significant token. If that token is `{`, the
  // block is a FUNCTION BODY with the latched async-ness; otherwise (e.g. `x => 1`
  // with an expression body, or any non-block) it lapses harmlessly.
  let pendingBody: { async: boolean } | null = null;
  // Async-ness of the most-recently-closed `(...)` of ANY kind — consumed ONLY by
  // a following `=>` (for `(...) =>` / `async (...) =>`). Unlike `pendingBody`,
  // setting this never mislabels an ordinary `if(...) {` block as a function body.
  let lastParenAsync = false;
  // True when the just-seen single bare identifier was prefixed by `async` (an
  // `async x => …` arrow head); consumed by the following `=>`.
  let arrowAsyncPending = false;

  let prevToken = ""; // previous significant token text ("" == start of input)
  let line = 1;
  let i = 0;
  const n = source.length;
  let clean = true;

  // Whether we are at "line start" for significant content on the current line —
  // i.e. we haven't yet seen a significant char on this line.
  let lineHasSignificant = false;

  function recordLineStart(offset: number): void {
    if (lineHasSignificant) return;
    lineHasSignificant = true;
    // Capture the leading word/char.
    let leadWord = "";
    const leadChar = codePointAt(source, offset);
    if (isIdentStartAt(source, offset)) {
      const j = readIdentifierEnd(source, offset);
      leadWord = source.slice(offset, j);
    }
    bindingsBefore.set(line, bindings.slice());
    lines.push({
      line,
      firstSignificantOffset: offset,
      prevToken,
      statementContext: inStatementContext(state),
      asyncContext: inAsyncContext(state),
      leadWord,
      leadChar,
    });
  }

  // Regex-vs-division disambiguation: a `/` starts a regex when the previous
  // significant token is NOT a value (i.e. not an identifier/number/`)`/`]`/`}`
  // /string/template/regex). We track this with `prevTokenIsValue`.
  let prevTokenIsValue = false;

  function setPrev(token: string, isValue: boolean): void {
    prevToken = token;
    prevTokenIsValue = isValue;
  }

  while (i < n) {
    const ch = source[i];

    // Newline handling.
    if (ch === "\n") {
      line++;
      lineHasSignificant = false;
      i++;
      continue;
    }
    // Whitespace.
    if (ch === " " || ch === "\t" || ch === "\r" || ch === "\f" || ch === "\v") {
      i++;
      continue;
    }

    // Line comment.
    if (ch === "/" && source[i + 1] === "/") {
      i += 2;
      while (i < n && source[i] !== "\n") i++;
      continue;
    }
    // Block comment.
    if (ch === "/" && source[i + 1] === "*") {
      i += 2;
      while (i < n && !(source[i] === "*" && source[i + 1] === "/")) {
        if (source[i] === "\n") {
          line++;
          lineHasSignificant = false;
        }
        i++;
      }
      if (i >= n) {
        clean = false;
        break;
      }
      i += 2;
      continue;
    }

    // From here on `ch` is significant content.
    recordLineStart(i);

    // Snapshot the one-token latches for THIS token, then clear the persistent
    // vars. Each latch is valid for exactly the token following the one that set
    // it; handlers below read the snapshot and re-set the persistent var (for the
    // NEXT token) as needed. This makes every latch strictly one-token-lived, so
    // a stale latch can never bleed across statements.
    const bodyMarker: { async: boolean } | null = pendingBody; // func-head `)` / `=>` → next `{` is a body
    const prevAsyncKeyword: boolean = sawAsyncKeyword; // `async` immediately before
    const prevArrowAsync: boolean = arrowAsyncPending; // `async x` before a `=>`
    const prevParenAsync: boolean = lastParenAsync; // `(...)` before a `=>`
    pendingBody = null;
    sawAsyncKeyword = false;
    arrowAsyncPending = false;
    lastParenAsync = false;

    // String literals.
    if (ch === '"' || ch === "'") {
      const quote = ch;
      i++;
      while (i < n && source[i] !== quote) {
        if (source[i] === "\\") {
          i += 2;
          continue;
        }
        if (source[i] === "\n") {
          // Unterminated string (no line continuation handling) — bail.
          clean = false;
          break;
        }
        i++;
      }
      if (i >= n || source[i] !== quote) {
        clean = false;
        break;
      }
      i++;
      setPrev("<string>", true);
      continue;
    }

    // Template literal.
    if (ch === "`") {
      i++;
      let terminated = false;
      while (i < n) {
        const c = source[i];
        if (c === "\\") {
          i += 2;
          continue;
        }
        if (c === "`") {
          i++;
          terminated = true;
          break;
        }
        if (c === "$" && source[i + 1] === "{") {
          // Enter interpolation: treat like a paren-balanced expression but with
          // its own brace marker so the matching `}` resumes the template.
          templateBraceMarks.push(state.braceStack.length);
          state.braceStack.push({
            kind: "object",
            asyncContext: inAsyncContext(state),
          }); // `${` opens an expression context
          i += 2;
          // Hand control back to the main loop to lex the interpolation; we mark
          // that the closing of THIS brace resumes template scanning via the
          // template marks. Simpler: recurse inline by continuing the outer loop.
          terminated = true; // we "exit" template scanning here; resumed on `}`
          break;
        }
        if (c === "\n") {
          line++;
          lineHasSignificant = false;
        }
        i++;
      }
      if (!terminated) {
        clean = false;
        break;
      }
      // If we exited because of `${`, we are now lexing the interpolation as
      // normal code; the matching `}` (handled below) will resume template mode.
      // If we exited because of a closing backtick, the template is a value.
      if (
        templateBraceMarks.length === 0 ||
        templateBraceMarks[templateBraceMarks.length - 1] !==
          state.braceStack.length - 1
      ) {
        setPrev("<template>", true);
      }
      continue;
    }

    // Regex literal (only when a regex is legal here, not division).
    if (ch === "/" && !prevTokenIsValue) {
      i++;
      let inClass = false;
      let terminated = false;
      while (i < n) {
        const c = source[i];
        if (c === "\\") {
          i += 2;
          continue;
        }
        if (c === "[") inClass = true;
        else if (c === "]") inClass = false;
        else if (c === "/" && !inClass) {
          i++;
          terminated = true;
          break;
        } else if (c === "\n") {
          break;
        }
        i++;
      }
      if (!terminated) {
        clean = false;
        break;
      }
      // Consume regex flags.
      while (i < n && isIdentPartAt(source, i)) {
        i += codePointAt(source, i).length;
      }
      setPrev("<regex>", true);
      continue;
    }

    // Identifier / keyword.
    if (isIdentStartAt(source, i)) {
      const j = readIdentifierEnd(source, i);
      const word = source.slice(i, j);
      i = j;
      // Collect simple binding names: `let x`, `const y`, `var z`, including
      // declarations in a `for (...)` header. Snapshot reads are individually
      // guarded, so a name that is no longer in scope becomes "<unavailable>"
      // instead of changing execution. We only record the FIRST identifier after
      // the keyword (single-binding common case); destructuring/multi are skipped.
      if (DECLARATION_KEYWORDS.has(word)) {
        // Skip whitespace to the binding name.
        let k = i;
        while (k < n && /\s/.test(source[k])) k++;
        if (
          bindings.length < MAX_LOCALS_PER_LINE &&
          k < n &&
          isIdentStartAt(source, k)
        ) {
          const m = readIdentifierEnd(source, k);
          const name = source.slice(k, m);
          if (!bindings.includes(name)) bindings.push(name);
        }
      }
      // --- Function-header / async tracking (reads this-token snapshots; sets
      // persistent latches for the NEXT token).
      if (word === "function") {
        pendingFuncHeader = true;
        pendingFuncAsync = prevAsyncKeyword; // `async function` → async body
      } else if (word === "async") {
        sawAsyncKeyword = true; // may prefix `function` or an arrow
      } else {
        // Any other word: if `async` preceded this bare identifier it is an arrow
        // head (`async x => …`); carry the async-ness to the following `=>`.
        arrowAsyncPending = prevAsyncKeyword;
      }
      // Keywords are NOT values for regex purposes, EXCEPT these which can be
      // followed by division (they yield values): this/super/true/false/null +
      // any identifier. We treat identifiers as values; keywords that introduce
      // expressions (return/typeof/…) are not values.
      const keywordNonValue = new Set([
        "return",
        "typeof",
        "instanceof",
        "in",
        "of",
        "new",
        "delete",
        "void",
        "throw",
        "do",
        "else",
        "yield",
        "await",
        "case",
        "async",
      ]);
      setPrev(word, !keywordNonValue.has(word));
      continue;
    }

    // Numbers (so `1 / 2` divides; a value precedes `/`).
    if (/[0-9]/.test(ch) || (ch === "." && /[0-9]/.test(source[i + 1] ?? ""))) {
      let j = i + 1;
      while (j < n && /[0-9a-fA-FxXoObBeE._]/.test(source[j])) j++;
      i = j;
      setPrev("<number>", true);
      continue;
    }

    // Punctuators / brackets.
    switch (ch) {
      case "{": {
        const kind = classifyBraceKind(prevToken);
        let asyncContext: boolean;
        if (kind === "object") {
          // Object literals don't host statements; async-ness is irrelevant
          // (never read because kind !== "block").
          asyncContext = inAsyncContext(state);
        } else if (bodyMarker) {
          // This block is a FUNCTION BODY (a func-head `)` or `=>` preceded it):
          // its async-ness is exactly what we proved.
          asyncContext = bodyMarker.async;
        } else {
          // Ordinary block (if/for/while/try/bare {}/class body): inherit.
          asyncContext = inAsyncContext(state);
        }
        state.braceStack.push({ kind, asyncContext });
        sawAsyncKeyword = false;
        setPrev("{", false);
        i++;
        break;
      }
      case "}": {
        // If this closes a `${` interpolation, resume template scanning.
        if (
          templateBraceMarks.length > 0 &&
          templateBraceMarks[templateBraceMarks.length - 1] ===
            state.braceStack.length - 1
        ) {
          templateBraceMarks.pop();
          if (state.braceStack.length) state.braceStack.pop();
          i++;
          // Resume the enclosing template literal until the next `${` or backtick.
          let terminated = false;
          while (i < n) {
            const c = source[i];
            if (c === "\\") {
              i += 2;
              continue;
            }
            if (c === "`") {
              i++;
              terminated = true;
              setPrev("<template>", true);
              break;
            }
            if (c === "$" && source[i + 1] === "{") {
              templateBraceMarks.push(state.braceStack.length);
              state.braceStack.push({
                kind: "object",
                asyncContext: inAsyncContext(state),
              });
              i += 2;
              terminated = true;
              break;
            }
            if (c === "\n") {
              line++;
              lineHasSignificant = false;
            }
            i++;
          }
          if (!terminated) {
            clean = false;
            i = n;
          }
          break;
        }
        if (state.braceStack.length) state.braceStack.pop();
        else clean = false;
        setPrev("}", true); // a closing brace can end a block/expression
        i++;
        break;
      }
      case "(": {
        state.parenDepth++;
        // A "(" is a FUNCTION param list when it directly follows `function`
        // (`pendingFuncHeader`, which spans an optional name). For arrow params
        // (`(x) => …`) we don't know yet; we record the `async`-ness (`async (...)
        // =>`) so a following `=>` can claim it. Any other "(" is "other".
        //
        // `head` is the token this "(" followed, which is what distinguishes a
        // METHOD head (`m(…) {`, `get x(…) {`, `constructor(…) {`) from a
        // CONTROL head (`if (…) {`, `for (…) {`). Without it a method body reads
        // as an ordinary block, inherits the top level's async-ness, and gets an
        // awaited hook inserted into a non-async function — which is a syntax
        // error in code the user never wrote.
        const isFuncHead = pendingFuncHeader;
        const async = pendingFuncHeader
          ? pendingFuncAsync
          : prevAsyncKeyword || prevArrowAsync;
        parenStack.push({
          kind: isFuncHead ? "func-head" : "other",
          async,
          head: prevToken,
        });
        pendingFuncHeader = false;
        pendingFuncAsync = false;
        setPrev("(", false);
        i++;
        break;
      }
      case ")": {
        if (state.parenDepth > 0) state.parenDepth--;
        else clean = false;
        const frame = parenStack.pop();
        lastParenAsync = frame ? frame.async : false;
        if (frame && frame.kind === "func-head") {
          // `function (...)` closed → the next token, if `{`, is a function body.
          pendingBody = { async: frame.async };
        } else if (frame && isMethodHeadToken(frame.head)) {
          // `m(...)` / `get x(...)` / `constructor(...)` closed → the next `{` is
          // a METHOD body. Only a positively-detected `async m(...)` is async;
          // anything else is assumed non-async, which is the safe direction (a
          // breakpoint that does not bind, never a hook that cannot compile).
          pendingBody = { async: frame.async };
        }
        // For any other paren we do NOT set pendingBody (so a plain `if(...) {`
        // block is correctly treated as an ordinary, inheriting block); only a
        // following `=>` (arrow) will promote it to a function body.
        setPrev(")", true);
        i++;
        break;
      }
      case "[":
        state.bracketDepth++;
        setPrev("[", false);
        i++;
        break;
      case "]":
        if (state.bracketDepth > 0) state.bracketDepth--;
        else clean = false;
        setPrev("]", true);
        i++;
        break;
      case "=":
        // Detect `=>` so brace classification can see it.
        if (source[i + 1] === ">") {
          // An arrow body `{` (if it follows) is a FUNCTION BODY. The arrow is
          // async iff `async (...) =>` (`prevParenAsync`, latched by the just-
          // closed paren) or `async x =>` (`prevArrowAsync`). (The token before
          // `=>` is a `)` of an "other" paren or a bare identifier, never a
          // func-head `)`, so `bodyMarker` is always null here.)
          const arrowAsync: boolean = prevArrowAsync || prevParenAsync;
          pendingBody = { async: arrowAsync };
          setPrev("=>", false);
          i += 2;
        } else {
          // A real assignment `=` ends any pending arrow/func-body expectation
          // (the snapshot already cleared the latches for this token).
          setPrev("=", false);
          i++;
        }
        break;
      case ";":
        setPrev(";", false);
        i++;
        break;
      default:
        // Any other punctuator: not a value (operators precede regex/expr).
        setPrev(ch, false);
        i++;
        break;
    }
  }

  if (state.parenDepth !== 0 || state.bracketDepth !== 0 || state.braceStack.length !== 0) {
    clean = false;
  }

  return { lines, bindingsBefore, clean };
}

/**
 * Build the locals-snapshot thunk source for a given line: an arrow that returns
 * an object literal reading each in-scope simple binding name, guarded so a
 * not-yet-initialised (TDZ) or otherwise-throwing read becomes the string
 * "<unavailable>" instead of crashing the user's program. Raw values stay in
 * the isolated execution worker so a genuinely paused frame can evaluate an
 * expression; only its bounded display copy crosses the worker boundary.
 *
 * Capped to keep the payload small and the inserted code short.
 */
function buildLocalsThunk(names: readonly string[]): string {
  const unique = Array.from(new Set(names)).slice(0, MAX_LOCALS_PER_LINE);
  if (unique.length === 0) return "()=>({})";
  // Each entry: `name:(()=>{try{return name}catch(_){return"<unavailable>"}})()`
  const entries = unique
    .map(
      (name) =>
        `${JSON.stringify(name)}:(()=>{try{return ${name}}catch(__e){return"<unavailable>"}})()`,
    )
    .join(",");
  return `()=>({${entries}})`;
}

/**
 * Whether `info` is a line we may safely instrument with a leading awaited hook.
 * Conservative — ALL must hold: statement context (not inside `()`/`[]`/object),
 * ASYNC context (await legal here — top level or a proven-async function body), a
 * statement-ending previous token, and a non-continuation, non-clause line lead.
 */
function safeToInstrument(info: LineInfo): boolean {
  if (!info.statementContext) return false;
  // The awaited hook is only legal where `await` is — top level (the runner wraps
  // the body in an async function) or inside a function we proved is async. Any
  // non-async function body is skipped (no breakpoint there) — never a miscompile.
  if (!info.asyncContext) return false;
  // Previous significant token must clearly END a statement or open a block.
  const okPrev = info.prevToken === "" ||
    info.prevToken === "{" ||
    info.prevToken === "}" ||
    info.prevToken === ";";
  if (!okPrev) return false;
  // The line must not START with a continuation char or an unsafe clause keyword.
  if (info.leadChar && CONTINUATION_FIRST_CHARS.has(info.leadChar)) return false;
  if (info.leadWord && UNSAFE_LINE_STARTS.includes(info.leadWord)) return false;
  // Don't instrument a line that is ONLY a closing brace (handled by prevToken
  // but leadChar `}` is also in CONTINUATION set, so already excluded).
  return true;
}

/**
 * Instrument `source` for debug execution. Inserts an `await __mrnoteDebugHook(...)`
 * at the start of every SAFE statement-boundary line (per the lexer + gates).
 * Returns the rewritten source plus the set of lines that received a hook.
 *
 * SAFETY: if the lexer did not complete cleanly (unterminated literal, unbalanced
 * brackets — i.e. the source likely has a syntax error or a construct we don't
 * model), this returns the ORIGINAL source unmodified with `instrumented:false`,
 * so the debug run executes exactly what the user wrote (a syntax error then
 * surfaces normally). We never emit half-instrumented code that could change
 * semantics.
 *
 * @param source The user's JS (already type-stripped if it was TS).
 * @param breakpointLines Optional: when provided, only these lines get a hook
 *   (smaller transform, fewer pause points). When omitted, ALL safe lines are
 *   hooked (so the worker can decide at runtime / support stepping anywhere).
 * @param conditions Optional per-line breakpoint conditions (VSCode 条件断点).
 *   A condition is baked into that line's hook call and evaluated IN THE USER
 *   SCOPE right before pausing; the hook skips the pause when it is falsy. An
 *   evaluation ERROR counts as a hit (VSCode semantics: never silently skip a
 *   breakpoint because its condition is broken). Conditions must pre-validate
 *   via {@link isValidBreakpointCondition} — a syntactically bad condition
 *   would otherwise break the whole instrumented program.
 */
export function instrumentForDebug(
  source: string,
  breakpointLines?: readonly number[],
  conditions?: Readonly<Record<number, string>>,
): InstrumentResult {
  const scan = scanSource(source);
  if (!scan.clean) {
    // Bail to verbatim source — never risk miscompiling.
    return { code: source, instrumented: false, hookedLines: [] };
  }

  const bpSet = breakpointLines ? new Set(breakpointLines) : null;
  const sourceLines = source.split("\n");
  const infoByLine = new Map<number, LineInfo>();
  for (const info of scan.lines) infoByLine.set(info.line, info);

  const hookedLines: number[] = [];
  const out: string[] = [];

  for (let lineNo = 1; lineNo <= sourceLines.length; lineNo++) {
    const text = sourceLines[lineNo - 1];
    const info = infoByLine.get(lineNo);
    const wants = !bpSet || bpSet.has(lineNo);
    if (info && wants && safeToInstrument(info)) {
      const names = scan.bindingsBefore.get(lineNo) ?? [];
      const thunk = buildLocalsThunk(names);
      // Preserve the line's leading indentation so column-sensitive nothing
      // breaks (whitespace is insignificant in JS but keeps output readable).
      const indent = text.slice(0, text.length - text.trimStart().length);
      const condition = conditions?.[lineNo];
      // The optional third hook argument is the condition VERDICT, computed in
      // the user scope at the pause site (it can read anything the line can —
      // strictly more capable than a debugger's frame-eval). try/catch makes a
      // throwing condition a hit, never a crash.
      const conditionArg =
        condition && isValidBreakpointCondition(condition)
          ? `, (()=>{try{return !!(${condition})}catch(__e){return true}})()`
          : "";
      const hookCall = `${indent}await ${DEBUG_HOOK_NAME}(${lineNo}, ${thunk}${conditionArg});`;
      out.push(hookCall);
      out.push(text);
      hookedLines.push(lineNo);
    } else {
      out.push(text);
    }
  }

  if (hookedLines.length === 0) {
    // Nothing safe to hook — return verbatim (still valid to "debug-run").
    return { code: source, instrumented: false, hookedLines: [] };
  }

  return { code: out.join("\n"), instrumented: true, hookedLines };
}
