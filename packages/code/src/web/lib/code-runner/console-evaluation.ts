/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Prompt semantics for the interactive console.
 *
 * Two things separate a prompt from running a file, and the console needs both.
 *
 * COMPLETION VALUE. The sandbox executes a submission as a function body and
 * reports that function's return value — right for a FILE, wrong for a prompt:
 * typing `1 + 2` would answer nothing at all. `eval` has exactly the semantics a
 * prompt wants (the value of the last expression, nothing for a declaration), so
 * handing the submission to `eval` is the whole trick.
 *
 * A SESSION. A prompt that forgets the previous line is a calculator, not a
 * console: `const rows = load()` followed by `rows.length` has to work. That
 * needs two things beyond keeping the worker alive (which the sandbox host
 * arranges):
 *
 *  - INDIRECT eval — `(0, eval)(…)` — evaluates in the worker's GLOBAL scope
 *    instead of the enclosing function body, so what it defines outlives the
 *    submission. A direct `eval` would define everything inside the per-run
 *    function and lose it the moment that function returned, no matter how long
 *    the worker lived.
 *  - Promoting top-level `let`/`const`/`class` to `var`, because lexical
 *    declarations in eval code go into a scope that is discarded when the eval
 *    returns — `var`, `function` and plain assignment are the only forms that
 *    survive. Without this the session would remember half of what you typed,
 *    which is worse than remembering none of it.
 *
 * The promotion costs the prompt its const-ness and its TDZ. That is the same
 * trade every browser devtools console makes, and it buys the thing people
 * actually want at a prompt: retyping a `const` line after a correction just
 * works instead of failing with "already declared".
 *
 * TOP-LEVEL AWAIT. `await fetchRows()` at a prompt is what a devtools console
 * or the Node REPL accept, and what eval refuses. A submission that awaits at
 * its top level is run as an async function instead, the way the Node REPL
 * rewrites it: its declarations become assignments to session globals
 * declared in front of it, its function declarations are published on the
 * global object, and its last expression is returned. The sandbox then waits
 * for that function and reports the value it settled to.
 */

import type { RunResult } from "@mrnote/core/web/lib/code-runner/js-runner.ts";

type ConsoleJavaScriptParser = (typeof import("@codemirror/lang-javascript"))["javascriptLanguage"]["parser"];

export type ConsoleDeclarationPromotionReceipt =
  /**
   * `awaits`: the source was rewritten to an async function to be awaited.
   * `firstLineEdits`: what the rewrite inserted or replaced on line 1, so an
   * error column there can be placed back on what was typed.
   */
  | {
      status: "ready";
      source: string;
      awaits?: boolean;
      firstLineEdits?: readonly ConsoleLineEdit[];
    }
  | { status: "unavailable"; source: string }
  | { status: "cancelled"; source: string };


export interface ConsoleDeclarationPromotionOptions {
  signal?: AbortSignal;
}

export type CodeConsoleEmptyMessageKey =
  | "code.console.placeholder"
  | "code.console.empty";

/**
 * Describe only continuity the active runtime really provides.
 *
 * JS/TS file runs deliberately use one-shot Workers while their Console owns
 * a separate long-lived realm, so replaying file variables here would be a
 * false promise. Page-scoped interpreter/database kernels genuinely share
 * their namespace between Run and Console and may retain the richer hint.
 */
export function codeConsoleEmptyMessageKey(
  language: string,
): CodeConsoleEmptyMessageKey {
  switch (language) {
    case "python":
    case "sql":
    case "lua":
    case "ruby":
      return "code.console.empty";
    default:
      return "code.console.placeholder";
  }
}

function consoleTableCell(value: unknown): string {
  if (value === null) return "null";
  if (value === undefined) return "";
  if (typeof value === "string") return value;
  if (
    typeof value === "number" ||
    typeof value === "boolean" ||
    typeof value === "bigint"
  ) {
    return String(value);
  }
  return "<complex output omitted>";
}

/**
 * Project structured run output into the Console's text-only transcript.
 * Rich data remains structured everywhere else; this session-only projection
 * avoids keeping legacy HTML/text duplicates merely to make it visible here.
 */
export function consoleResultForDisplay(result: RunResult): RunResult {
  const extraLogs: string[] = [];
  if (result.rich?.kind === "table") {
    extraLogs.push(result.rich.columns.join("\t"));
    for (const row of result.rich.rows) {
      extraLogs.push(row.map(consoleTableCell).join("\t"));
    }
  } else if (result.rich?.kind === "image") {
    extraLogs.push(`[Image output: ${result.rich.mime}]`);
  } else if (result.rich?.kind === "html") {
    extraLogs.push("[HTML output]");
  }
  if (result.html) {
    extraLogs.push("[HTML output]");
  }
  for (let index = 0; index < (result.images?.length ?? 0); index += 1) {
    extraLogs.push(
      `[Image output: ${result.imageMimes?.[index] ?? "image/png"}]`,
    );
  }
  if (extraLogs.length === 0) return result;

  return {
    logs: [...result.logs, ...extraLogs],
    logStreams: [
      ...result.logs.map((_, index) =>
        result.logStreams?.[index] === "stderr"
          ? ("stderr" as const)
          : ("stdout" as const),
      ),
      ...extraLogs.map(() => "stdout" as const),
    ],
    ...(result.notebookOutputs
      ? { notebookOutputs: result.notebookOutputs }
      : {}),
    ...(result.result === undefined ? {} : { result: result.result }),
    ...(result.error === undefined ? {} : { error: result.error }),
    ...(result.executionErrorCode === undefined
      ? {}
      : { executionErrorCode: result.executionErrorCode }),
    ...(result.timedOut ? { timedOut: true } : {}),
    ...(result.interrupted ? { interrupted: true } : {}),
  };
}

/** A rewrite on line 1: at authored offset `at`, the text grew by `delta`. */
export interface ConsoleLineEdit {
  readonly at: number;
  readonly delta: number;
}

/** The authored 1-based column of a 1-based column in rewritten line 1. */
function authoredColumn(column: number, edits: readonly ConsoleLineEdit[]): number {
  const rewritten = column - 1;
  let shift = 0;
  for (const edit of edits) {
    if (edit.at + shift + Math.max(0, edit.delta) > rewritten) break;
    shift += edit.delta;
  }
  return rewritten - shift + 1;
}

const CONSOLE_ERROR_FRAME =
  /^(\s*at (?:[^\s(]+ \()?)<console>:(\d+):(\d+)(\)?)\s*$/;

/**
 * Place the sandbox's `<console>:line:column` frames on what was typed: a
 * frame past the submission's last line is the evaluation wrapper's (the
 * `})()` that runs an awaiting submission) and is left out, and a column on
 * line 1 is mapped back through the rewrite's edits there and then loses
 * `firstLineOffset` characters of prelude (the TypeScript runtime's module
 * shim) that precede the typed text.
 */
export function placeConsoleErrorFrames(
  result: RunResult,
  submission: {
    lineCount: number;
    firstLineEdits?: readonly ConsoleLineEdit[];
    firstLineOffset?: number;
  },
): RunResult {
  if (typeof result.error !== "string" || !result.error.includes("<console>:")) {
    return result;
  }
  const edits = submission.firstLineEdits ?? [];
  const offset = submission.firstLineOffset ?? 0;
  const error = result.error
    .split("\n")
    .flatMap((text) => {
      const frame = CONSOLE_ERROR_FRAME.exec(text);
      if (!frame) return [text];
      const line = Number(frame[2]);
      if (line > submission.lineCount) return [];
      const column =
        line === 1
          ? authoredColumn(Number(frame[3]), edits) - offset
          : Number(frame[3]);
      return [`${frame[1]}<console>:${line}:${column >= 1 ? column : frame[3]}${frame[4]}`];
    })
    .join("\n");
  return { ...result, error };
}

/**
 * Wrap one console submission so the sandbox reports its completion value and
 * evaluates it in the session's global scope.
 *
 * The source is embedded as a JSON string literal, so it is data to the wrapper
 * and cannot terminate it early or inject syntax around it.
 */
export function consoleEvaluationSource(source: string): string {
  return `return (0, eval)(${JSON.stringify(source)});`;
}

/**
 * A Lua prompt as the standalone `lua` prompt reads it: first as an
 * expression list whose values are printed (`base + 2` shows 42, `f()` shows
 * whatever f returns), and only when that does not compile, as the statement
 * typed. The values go through a vararg call so `nil` results still print and
 * a call that returns nothing prints nothing; the closing parenthesis sits on
 * its own line so a trailing `--` comment cannot swallow it.
 */
export function luaConsoleExpressionSource(source: string): string {
  return `(function(...) if select("#", ...) > 0 then print(...) end end)(${source}\n)`;
}

/**
 * Whether a Lua run was refused at compile time, so nothing in it ran. The
 * prefix is written by the trusted runner from the worker's receipt status,
 * never by the program's own output.
 */
export function isLuaCompileFailure(result: RunResult): boolean {
  return (
    result.logs.length === 0 &&
    typeof result.error === "string" &&
    result.error.startsWith("Lua compile error:")
  );
}

/**
 * Evaluate one Ruby prompt in a stable top-level Binding and print its
 * completion value. The authored prompt is embedded as a string literal, so it
 * cannot terminate or rewrite the host wrapper. The binding lives only inside
 * the page's already-isolated Ruby Worker and disappears on Stop/Restart.
 */
export function rubyConsoleEvaluationSource(
  source: string,
  sessionToken: string,
): string {
  if (!/^[A-F0-9]{32}$/.test(sessionToken)) {
    throw new Error("Ruby console session authority is unavailable.");
  }
  const encoded = JSON.stringify(source);
  const host = `MRNoteConsoleHostV1_${sessionToken}`;
  return `
unless defined?(::${host})
  module ::${host}
    BINDING = TOPLEVEL_BINDING
    EVAL = Binding.instance_method(:eval)
    OUTPUT = Kernel.method(:puts)
    def self.evaluate(source)
      value = EVAL.bind_call(BINDING, source, "(console)", 1)
      OUTPUT.call("=> #{value.inspect}") unless value.nil?
    rescue Exception => error
      # Report the prompt's own frames, not this host's eval plumbing.
      begin
        frames = error.backtrace
        error.set_backtrace(frames.grep(/\\A\\(console\\):/)) if frames
      rescue Exception
      end
      raise error
    end
  end
  ::${host}.freeze
end
::${host}.evaluate(${encoded})
`.trim();
}

/** One text replacement, applied back-to-front so earlier offsets stay valid. */
interface Splice {
  from: number;
  to: number;
  text: string;
}

/** The edits a set of splices (plus a preamble) makes to line 1, in order. */
function firstLineEdits(
  source: string,
  splices: readonly Splice[],
  preamble = 0,
): ConsoleLineEdit[] {
  const firstBreak = source.indexOf("\n");
  const edits = splices
    .filter((splice) => firstBreak < 0 || splice.from <= firstBreak)
    .map((splice) => ({
      at: splice.from,
      delta: splice.text.length - (splice.to - splice.from),
    }))
    .filter((edit) => edit.delta !== 0)
    .sort((a, b) => a.at - b.at);
  return preamble > 0 ? [{ at: 0, delta: preamble }, ...edits] : edits;
}

type ConsoleSyntaxNode = ReturnType<
  ReturnType<ConsoleJavaScriptParser["parse"]>["topNode"]["childAfter"]
> & {};

/** Nodes whose body is its own function scope: an `await` in them is theirs. */
const FUNCTION_SCOPES = new Set([
  "FunctionDeclaration",
  "FunctionExpression",
  "ArrowFunction",
  "MethodDeclaration",
  "ClassBody",
]);

/** Whether `await` is used outside every function of the submission. */
function awaitsAtTopLevel(top: ConsoleSyntaxNode): boolean {
  let found = false;
  const visit = (node: ConsoleSyntaxNode) => {
    for (let child = node.firstChild; child && !found; child = child.nextSibling) {
      if (child.name === "await") {
        found = true;
        return;
      }
      if (!FUNCTION_SCOPES.has(child.name)) visit(child);
    }
  };
  visit(top);
  return found;
}

/**
 * The names a declaration binds: plain names, and the leaves of object and
 * array patterns (`{a, b: [c], ...rest}`), never names inside a default value
 * or an initializer.
 */
function boundNames(declaration: ConsoleSyntaxNode, source: string): string[] {
  const names: string[] = [];
  const target = (node: ConsoleSyntaxNode) => {
    if (node.name === "VariableDefinition") {
      names.push(source.slice(node.from, node.to));
    } else if (node.name === "ObjectPattern") {
      for (let property = node.firstChild; property; property = property.nextSibling) {
        if (property.name !== "PatternProperty") continue;
        // `{a: b}` binds `b`; `{a}`, `{a = 1}` bind `a`; `{...rest}` binds `rest`.
        const colon = property.getChild(":");
        const rest = property.getChild("VariableDefinition");
        const shorthand = property.getChild("PropertyName");
        if (colon?.nextSibling) target(colon.nextSibling);
        else if (rest) target(rest);
        else if (shorthand) names.push(source.slice(shorthand.from, shorthand.to));
      }
    } else if (node.name === "ArrayPattern") {
      targets(node);
    }
  };
  // Comma-separated targets, each optionally followed by `= value`: the
  // declarators of a declaration, or the elements of an array pattern.
  const targets = (list: ConsoleSyntaxNode) => {
    let inValue = false;
    for (let child = list.firstChild; child; child = child.nextSibling) {
      if (child.name === ",") inValue = false;
      else if (child.name === "Equals") inValue = true;
      else if (!inValue) target(child);
    }
  };
  targets(declaration);
  return [...new Set(names)];
}

/**
 * Rewrite a submission that awaits at its top level into an async function
 * the sandbox can wait for (see the module comment). Every authored line keeps
 * its line number, so an error frame still names the line that threw.
 */
function awaitingSubmission(
  source: string,
  top: ConsoleSyntaxNode,
): { source: string; firstLineEdits: ConsoleLineEdit[] } {
  const splices: Splice[] = [];
  const hoisted: string[] = [];
  const published: string[] = [];
  let last: ConsoleSyntaxNode | null = null;
  for (let node = top.firstChild; node; node = node.nextSibling) {
    if (node.name === "LineComment" || node.name === "BlockComment") continue;
    last = node;
    if (node.name === "VariableDeclaration") {
      const keyword = node.firstChild;
      if (!keyword) continue;
      hoisted.push(...boundNames(node, source));
      const semicolon = node.lastChild?.name === ";" ? node.lastChild : null;
      const end = semicolon ? semicolon.from : node.to;
      splices.push({ from: keyword.from, to: keyword.to, text: "void (" });
      splices.push({ from: end, to: semicolon ? semicolon.to : end, text: ");" });
    } else if (node.name === "ClassDeclaration") {
      const name = node.getChild("VariableDefinition");
      if (!name) continue;
      const identifier = source.slice(name.from, name.to);
      hoisted.push(identifier);
      splices.push({ from: node.from, to: node.from, text: `${identifier} = ` });
      splices.push({ from: node.to, to: node.to, text: ";" });
    } else if (node.name === "FunctionDeclaration") {
      const name = node.getChild("VariableDefinition");
      if (name) published.push(source.slice(name.from, name.to));
    }
  }
  if (last?.name === "ExpressionStatement") {
    const semicolon = last.lastChild?.name === ";" ? last.lastChild : null;
    splices.push({ from: last.from, to: last.from, text: "return (" });
    splices.push({ from: semicolon ? semicolon.from : last.to, to: semicolon ? semicolon.from : last.to, text: ")" });
  }
  let body = source;
  for (const splice of splices.sort((a, b) => b.from - a.from || b.to - a.to)) {
    body = body.slice(0, splice.from) + splice.text + body.slice(splice.to);
  }
  const names = [...new Set(hoisted)];
  // One line of preamble joined to the body's first line, so line N of the
  // submission is line N of what runs.
  const declare = names.length ? `var ${names.join(", ")}; ` : "";
  const publish = published.map((name) => `this.${name} = ${name}; `).join("");
  const preamble = `${declare}(async () => { ${publish}`;
  return {
    source: `${preamble}${body}\n})()`,
    firstLineEdits: firstLineEdits(source, splices, preamble.length),
  };
}

/** Rewrite declarations with one already-authorized parser instance. */
function promoteWithParser(
  source: string,
  parser: ConsoleJavaScriptParser,
): { source: string; awaits: boolean; firstLineEdits?: ConsoleLineEdit[] } {
  const tree = parser.parse(source);
  // Lezer recovers from broken input rather than refusing it, so a half-typed
  // `const` still comes back shaped like a declaration. Rewriting that would
  // edit source the user is in the middle of typing; leave it exactly as it is
  // and let the runtime report the syntax error it was always going to report.
  let broken = false;
  tree.iterate({
    enter: (node) => {
      if (node.type.isError) broken = true;
      return !broken;
    },
  });
  if (broken) return { source, awaits: false };
  if (/\bawait\b/.test(source) && awaitsAtTopLevel(tree.topNode)) {
    return { ...awaitingSubmission(source, tree.topNode), awaits: true };
  }

  const splices: Splice[] = [];
  const cursor = tree.cursor();
  if (!cursor.firstChild()) return { source, awaits: false };
  do {
    if (cursor.name === "VariableDeclaration") {
      const keyword = cursor.node.firstChild;
      if (keyword && (keyword.name === "let" || keyword.name === "const")) {
        splices.push({ from: keyword.from, to: keyword.to, text: "var" });
      }
      continue;
    }
    if (cursor.name === "ClassDeclaration") {
      // `class Foo {}` is lexical too. As an assignment it persists, and it
      // keeps its own name binding so recursion and `Foo.name` still work.
      const name = cursor.node.getChild("VariableDefinition");
      if (name) {
        const identifier = source.slice(name.from, name.to);
        splices.push({
          from: cursor.from,
          to: cursor.from,
          text: `var ${identifier} = `,
        });
      }
    }
  } while (cursor.nextSibling());

  if (splices.length === 0) return { source, awaits: false };
  let promoted = source;
  for (const splice of splices.sort((a, b) => b.from - a.from)) {
    promoted =
      promoted.slice(0, splice.from) + splice.text + promoted.slice(splice.to);
  }
  return {
    source: promoted,
    awaits: false,
    firstLineEdits: firstLineEdits(source, splices),
  };
}

const CANCELLED_PARSER_WAIT = Symbol("cancelled-console-parser-wait");

/**
 * Create the async declaration-promotion authority used by the interactive
 * console. Keeping the loader injectable makes the failure/retry boundary
 * deterministic without exposing mutable production globals.
 */
export function createConsoleDeclarationPromoter(
  load: () => Promise<ConsoleJavaScriptParser>,
): (
  source: string,
  options?: ConsoleDeclarationPromotionOptions,
) => Promise<ConsoleDeclarationPromotionReceipt> {
  let loaded: ConsoleJavaScriptParser | null = null;
  let inFlight: Promise<ConsoleJavaScriptParser | null> | null = null;

  const loadParser = (): Promise<ConsoleJavaScriptParser | null> => {
    if (loaded) return Promise.resolve(loaded);
    if (!inFlight) {
      inFlight = Promise.resolve()
        .then(load)
        .then((parser) => {
          if (!parser || typeof parser.parse !== "function") return null;
          loaded = parser;
          return parser;
        })
        .catch(() => null)
        .finally(() => {
          inFlight = null;
        });
    }
    return inFlight;
  };

  const waitForParser = (
    parserPromise: Promise<ConsoleJavaScriptParser | null>,
    signal?: AbortSignal,
  ): Promise<ConsoleJavaScriptParser | null | typeof CANCELLED_PARSER_WAIT> => {
    if (!signal) return parserPromise;
    if (signal.aborted) return Promise.resolve(CANCELLED_PARSER_WAIT);
    return new Promise((resolve) => {
      let settled = false;
      const finish = (
        value:
          | ConsoleJavaScriptParser
          | null
          | typeof CANCELLED_PARSER_WAIT,
      ) => {
        if (settled) return;
        settled = true;
        signal.removeEventListener("abort", onAbort);
        resolve(value);
      };
      const onAbort = () => finish(CANCELLED_PARSER_WAIT);
      signal.addEventListener("abort", onAbort, { once: true });
      if (signal.aborted) {
        onAbort();
        return;
      }
      void parserPromise.then(finish);
    });
  };

  return async (source, options = {}) => {
    if (options.signal?.aborted) return { status: "cancelled", source };
    if (!source.trim() || !/\b(?:let|const|class|await)\b/.test(source)) {
      return { status: "ready", source };
    }

    const parser = await waitForParser(loadParser(), options.signal);
    if (parser === CANCELLED_PARSER_WAIT || options.signal?.aborted) {
      return { status: "cancelled", source };
    }
    if (!parser) return { status: "unavailable", source };
    try {
      const rewritten = promoteWithParser(source, parser);
      const edits = rewritten.firstLineEdits?.length
        ? { firstLineEdits: rewritten.firstLineEdits }
        : {};
      return rewritten.awaits
        ? { status: "ready", source: rewritten.source, awaits: true, ...edits }
        : { status: "ready", source: rewritten.source, ...edits };
    } catch {
      return { status: "unavailable", source };
    }
  };
}

/**
 * Rewrite TOP-LEVEL lexical declarations so they survive the next console
 * submission. The JavaScript parser is loaded only for submissions that may
 * need promotion; expressions and ordinary statements stay on the zero-import
 * path. A failed parser load never executes a silently non-persistent fallback.
 */
export const promoteConsoleDeclarations = createConsoleDeclarationPromoter(
  async () => {
    const grammarModule = await import("@codemirror/lang-javascript");
    return grammarModule.javascriptLanguage.parser;
  },
);
