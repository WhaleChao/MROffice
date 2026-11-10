/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Bounded discovery, execution-harness and persistence helpers for the
 * notebook-scoped Code Test Explorer.
 *
 * This module deliberately does not invent a second runtime. It discovers a
 * small, honest convention (top-level `test*` functions) and
 * builds source for the existing selector-bounded execution providers. The
 * provider receipt remains the execution authority; the marker parsed here is
 * only the authored test result carried inside that already-bounded output.
 */

import type { RunResult } from "@mrnote/core/web/lib/code-runner/js-runner.ts";
import {
  EXECUTION_OUTPUT_TRUNCATION_MARKER,
} from "@mrnote/core/web/lib/code-runner/execution-output-budget.ts";

export const CODE_TEST_RUN_RECEIPT_KEY = "lastTestRun";
export const CODE_TEST_RUN_RECEIPT_VERSION = 1 as const;
export const CODE_TEST_RUN_TIMELINE_KEY = "testRunTimeline";
export const CODE_TEST_RUN_TIMELINE_VERSION = 1 as const;
export const CODE_TEST_WORKSPACE_FILE_LIMIT = 100;
export const CODE_TESTS_PER_FILE_LIMIT = 100;
export const CODE_TEST_TOTAL_LIMIT = 500;
export const CODE_TEST_SOURCE_CODE_UNIT_LIMIT = 1_000_000;
export const CODE_TEST_RUN_TIMELINE_ENTRY_LIMIT = 20;
export const CODE_TEST_RUN_TIMELINE_UTF8_BYTE_LIMIT = 384_000;
export const CODE_TEST_RUN_OUTPUT_UTF8_BYTE_LIMIT = 128_000;

const CODE_TEST_PATH_LIMIT = 1_024;
const CODE_TEST_NAME_LIMIT = 256;
const CODE_TEST_MESSAGE_LIMIT = 512;
const CODE_TEST_RUN_ID_LIMIT = 128;
const CODE_TEST_RESULT_LINE_LIMIT = 1_000_000;
const CODE_TEST_DURATION_LIMIT_MS = 24 * 60 * 60 * 1_000;
const CODE_TEST_RUN_TIMELINE_SCAN_LIMIT = 80;
const CODE_TEST_RUN_OUTPUT_LINE_LIMIT = 2_000;
const CODE_TEST_RUN_OUTPUT_LINE_CODE_UNIT_LIMIT = 20_000;
const CODE_TEST_MARKER_PREFIX = "__MRNOTE_CODE_TEST_RESULT__";
// Internal runner protocol. Only a harness produced here carries this prefix;
// js-runner uses its existing isolated async worker for that exact source.
// Authored code may contain the same comment without gaining any authority --
// it merely selects the same credential-free sandbox execution mode.
const CODE_TEST_ASYNC_HARNESS_PREFIX =
  "/*__MRNOTE_CODE_TEST_AWAIT_COMPLETION_V1__*/";
const SHA256_HEX_PATTERN = /^[0-9a-f]{64}$/;
const TEST_NAME_PATTERN =
  /^test[$\u200c\u200d\p{ID_Continue}]*$/u;
const PYTHON_TEST_NAME_PATTERN = /^test[\p{ID_Continue}]*$/u;
const LUA_TEST_NAME_PATTERN = /^test[A-Za-z0-9_]*$/;
const SHELL_TEST_NAME_PATTERN = /^test[A-Za-z0-9_]*$/;

export type CodeTestLanguage =
  | "javascript"
  | "typescript"
  | "python"
  | "lua"
  | "ruby"
  | "bash";
export type CodeTestRunScope = "selected" | "file" | "all";
export type CodeTestResultStatus = "passed" | "failed";

export interface CodeTestSource {
  readonly pageId: string;
  readonly path: string;
  readonly language: CodeTestLanguage;
  readonly contentVersion: number;
  readonly source: string;
}

export interface CodeTestSourceIdentity {
  readonly pageId: string;
  readonly path: string;
  readonly language: CodeTestLanguage;
  readonly sourceFingerprint: string;
}

export interface DiscoveredCodeTest {
  readonly id: string;
  readonly pageId: string;
  readonly path: string;
  readonly language: CodeTestLanguage;
  readonly name: string;
  readonly line: number;
  readonly execution: "sync" | "async";
}

export interface CodeTestFile {
  readonly source: CodeTestSource;
  readonly tests: readonly DiscoveredCodeTest[];
  /** True when this file's visible list is only a bounded prefix. */
  readonly truncated: boolean;
}

export interface CodeTestDiscovery {
  readonly files: readonly CodeTestFile[];
  readonly totalTests: number;
  readonly truncated: boolean;
}

export interface CodeTestResult {
  readonly id: string;
  readonly testId: string;
  readonly pageId: string;
  readonly path: string;
  readonly language: CodeTestLanguage;
  readonly name: string;
  readonly line: number;
  readonly status: CodeTestResultStatus;
  readonly message?: string;
  readonly durationMs: number;
}

export interface PersistedCodeTestRunReceipt {
  readonly version: typeof CODE_TEST_RUN_RECEIPT_VERSION;
  readonly runId: string;
  readonly scope: CodeTestRunScope;
  readonly completedAt: string;
  readonly sources: readonly CodeTestSourceIdentity[];
  readonly results: readonly CodeTestResult[];
}

export interface PersistedCodeTestSourceDigest {
  readonly pageId: string;
  readonly sourceSha256: string;
  readonly lineCount: number;
}

export interface CodeTestRunFileOutput {
  readonly pageId: string;
  readonly output: RunResult;
}

export interface PersistedCodeTestFileOutput {
  readonly pageId: string;
  readonly path: string;
  readonly language: CodeTestLanguage;
  readonly logs: readonly string[];
  readonly logStreams: readonly ("stdout" | "stderr")[];
}

/**
 * One durable Test Explorer result. The receipt keeps the bounded display
 * result while `sourceDigests` is the dispatch-time authority after reopen.
 * Authored source bytes are never copied into page metadata.
 */
export interface PersistedCodeTestRunTimelineEntry {
  readonly version: typeof CODE_TEST_RUN_TIMELINE_VERSION;
  readonly receipt: PersistedCodeTestRunReceipt;
  readonly sourceDigests: readonly PersistedCodeTestSourceDigest[];
  /** Marker-free, text-only output from the exact provider receipt per file. */
  readonly fileOutputs?: readonly PersistedCodeTestFileOutput[];
}

export interface PersistedCodeTestRunTimeline {
  readonly version: typeof CODE_TEST_RUN_TIMELINE_VERSION;
  readonly entries: readonly PersistedCodeTestRunTimelineEntry[];
}

export interface CodeTestHarness {
  readonly source: string;
  readonly marker: string;
  readonly tests: readonly DiscoveredCodeTest[];
  readonly execution: "sync" | "async";
}

export interface ParsedCodeTestHarnessResult {
  readonly results: readonly CodeTestResult[];
  readonly complete: boolean;
}

/**
 * One failed test, rebound to the exact authored source that produced its
 * persisted receipt. `executionSource` is transient debugger input: it keeps
 * every authored line at its original 1-based line and appends only the
 * selected invocation.
 */
export interface CodeTestDebugHarness {
  readonly source: CodeTestSource;
  readonly test: DiscoveredCodeTest;
  readonly result: CodeTestResult;
  readonly executionSource: string;
  readonly authoredLineCount: number;
}

export interface CodeTestDebugHandoff extends CodeTestDebugHarness {
  readonly requestId: string;
  readonly receipt: PersistedCodeTestRunReceipt;
  readonly targetBufferId?: string;
  readonly expiresAt: number;
}

interface HarnessPayload {
  readonly id: string;
  readonly status: CodeTestResultStatus;
  readonly line: number;
  readonly durationMs: number;
  readonly message?: string;
}

export function isCodeTestLanguage(value: string): value is CodeTestLanguage {
  return (
    value === "javascript" ||
    value === "typescript" ||
    value === "python" ||
    value === "lua" ||
    value === "ruby" ||
    value === "bash"
  );
}

function finiteInteger(
  value: unknown,
  minimum: number,
  maximum: number,
): number | null {
  return typeof value === "number" &&
    Number.isFinite(value) &&
    Number.isInteger(value) &&
    value >= minimum &&
    value <= maximum
    ? value
    : null;
}

function boundedString(value: unknown, maximum: number): string | null {
  return typeof value === "string" && value.length <= maximum ? value : null;
}

function validPageId(value: unknown): value is string {
  return typeof value === "string" && value.length > 0 && value.length <= 256;
}

function validPath(value: unknown): value is string {
  return typeof value === "string" && value.length > 0 && value.length <= CODE_TEST_PATH_LIMIT;
}

/**
 * Synchronous non-cryptographic fingerprint used only to mark a persisted UI
 * result fresh/stale after reopen. Dispatch-time execution still uses the
 * provider contract's SHA-256 and exact source equality.
 */
export function codeTestSourceFingerprint(source: string): string {
  let fnv = 0x811c9dc5;
  let djb = 0x1505;
  for (let index = 0; index < source.length; index += 1) {
    const unit = source.charCodeAt(index);
    fnv = Math.imul(fnv ^ unit, 0x01000193) >>> 0;
    djb = (Math.imul(djb, 33) ^ unit) >>> 0;
  }
  return `v1:${source.length}:${fnv.toString(36)}:${djb.toString(36)}`;
}

export function codeTestSourceIdentity(
  source: CodeTestSource,
): CodeTestSourceIdentity {
  return {
    pageId: source.pageId,
    path: source.path,
    language: source.language,
    sourceFingerprint: codeTestSourceFingerprint(source.source),
  };
}

function testId(source: CodeTestSource, name: string, line: number): string {
  return `${source.pageId}:${line}:${name}`;
}

function discoverPythonTests(source: CodeTestSource): DiscoveredCodeTest[] {
  const tests: DiscoveredCodeTest[] = [];
  const lines = source.source.split("\n");
  for (let index = 0; index < lines.length; index += 1) {
    // Leading whitespace means nested/class scope. Both declarations are
    // explicit: a synchronous test that happens to return an awaitable must
    // still fail closed instead of silently changing its execution contract.
    const match = /^(async\s+)?def\s+(test[\p{ID_Continue}]*)\s*\(/u.exec(
      lines[index],
    );
    const name = match?.[2];
    if (!name || !PYTHON_TEST_NAME_PATTERN.test(name)) continue;
    const line = index + 1;
    tests.push({
      id: testId(source, name, line),
      pageId: source.pageId,
      path: source.path,
      language: source.language,
      name,
      line,
      execution: match?.[1] ? "async" : "sync",
    });
    if (tests.length >= CODE_TESTS_PER_FILE_LIMIT) break;
  }
  return tests;
}

type JavaScriptLexicalState = "code" | "single" | "double" | "template" | "block";

function slashStartsJavaScriptRegularExpression(maskedPrefix: string): boolean {
  const prefix = maskedPrefix.trimEnd();
  if (prefix.length === 0) return true;
  if (prefix.endsWith("=>")) return true;
  if (/[({=,:;!?&|+\-*%^~<>\[]$/.test(prefix)) return true;
  return /(?:^|[^\w$])(?:return|throw|case|delete|void|typeof|instanceof|in|of|yield|await|new)\s*$/.test(
    prefix,
  );
}

function maskJavaScriptLine(
  line: string,
  initialState: JavaScriptLexicalState,
): { masked: string; state: JavaScriptLexicalState } {
  let state = initialState;
  let escaped = false;
  let masked = "";
  for (let index = 0; index < line.length; index += 1) {
    const char = line[index];
    const next = line[index + 1];
    if (state === "block") {
      if (char === "*" && next === "/") {
        masked += "  ";
        index += 1;
        state = "code";
      } else {
        masked += " ";
      }
      continue;
    }
    if (state !== "code") {
      masked += " ";
      if (escaped) {
        escaped = false;
        continue;
      }
      if (char === "\\") {
        escaped = true;
        continue;
      }
      if (
        (state === "single" && char === "'") ||
        (state === "double" && char === '"') ||
        (state === "template" && char === "`")
      ) {
        state = "code";
      }
      continue;
    }
    if (char === "/" && next === "/") {
      masked += " ".repeat(line.length - index);
      break;
    }
    if (char === "/" && next === "*") {
      masked += "  ";
      index += 1;
      state = "block";
      continue;
    }
    if (
      char === "/" &&
      next !== "/" &&
      next !== "*" &&
      slashStartsJavaScriptRegularExpression(masked)
    ) {
      // A regex body may contain braces, comment markers, and quoted text.
      // Mask it before the selected-test brace scan; otherwise `/}/` can make
      // the scanner stop early and skip a later ambiguous `return`.
      masked += " ";
      let inCharacterClass = false;
      let regexEscaped = false;
      let closed = false;
      for (index += 1; index < line.length; index += 1) {
        const regexChar = line[index];
        masked += " ";
        if (regexEscaped) {
          regexEscaped = false;
          continue;
        }
        if (regexChar === "\\") {
          regexEscaped = true;
          continue;
        }
        if (regexChar === "[") {
          inCharacterClass = true;
          continue;
        }
        if (regexChar === "]" && inCharacterClass) {
          inCharacterClass = false;
          continue;
        }
        if (regexChar === "/" && !inCharacterClass) {
          closed = true;
          break;
        }
      }
      if (closed) {
        while (index + 1 < line.length && /[A-Za-z]/.test(line[index + 1])) {
          index += 1;
          masked += " ";
        }
      }
      continue;
    }
    if (char === "'" || char === '"' || char === "`") {
      state = char === "'" ? "single" : char === '"' ? "double" : "template";
      masked += " ";
      escaped = false;
      continue;
    }
    masked += char;
  }
  // JavaScript string literals cannot continue past a newline without an
  // escape. Treat a malformed line as code on the next line; the provider will
  // remain the final syntax authority.
  if (state === "single" || state === "double") state = "code";
  return { masked, state };
}

function discoverJavaScriptTests(source: CodeTestSource): DiscoveredCodeTest[] {
  const tests: DiscoveredCodeTest[] = [];
  const lines = source.source.split("\n");
  let lexicalState: JavaScriptLexicalState = "code";
  let braceDepth = 0;
  for (let index = 0; index < lines.length; index += 1) {
    const startsInCode = lexicalState === "code";
    const maskedLine = maskJavaScriptLine(lines[index], lexicalState);
    lexicalState = maskedLine.state;
    const code = maskedLine.masked;
    if (braceDepth === 0 && startsInCode) {
      const functionMatch = /^\s*(?:export\s+)?(async\s+)?function(?!\s*\*)\s+(test[$\u200c\u200d\p{ID_Continue}]*)\s*\(/u.exec(code);
      const functionExpressionMatch = /^\s*(?:export\s+)?(?:const|let|var)\s+(test[$\u200c\u200d\p{ID_Continue}]*)\s*(?::[^=]+)?=\s*(async\s+)?function(?!\s*\*)\b/u.exec(code);
      const arrowMatch = /^\s*(?:export\s+)?(?:const|let|var)\s+(test[$\u200c\u200d\p{ID_Continue}]*)\s*(?::[^=]+)?=\s*(async\s+)?(?:\([^)]*\)|[$_\p{ID_Start}][$\u200c\u200d\p{ID_Continue}]*)\s*(?::[^=]+)?=>/u.exec(code);
      const name = functionMatch?.[2] ?? functionExpressionMatch?.[1] ?? arrowMatch?.[1];
      const execution =
        functionMatch?.[1] || functionExpressionMatch?.[2] || arrowMatch?.[2]
          ? "async"
          : "sync";
      if (name && TEST_NAME_PATTERN.test(name)) {
        const line = index + 1;
        tests.push({
          id: testId(source, name, line),
          pageId: source.pageId,
          path: source.path,
          language: source.language,
          name,
          line,
          execution,
        });
        if (tests.length >= CODE_TESTS_PER_FILE_LIMIT) break;
      }
    }
    for (const char of code) {
      if (char === "{") braceDepth += 1;
      else if (char === "}") braceDepth = Math.max(0, braceDepth - 1);
    }
  }
  return tests;
}

type LuaLexicalState =
  | { readonly kind: "code" }
  | { readonly kind: "single" | "double" }
  | { readonly kind: "long-string" | "long-comment"; readonly equals: number };

function luaLongBracketOpener(
  line: string,
  index: number,
): { readonly equals: number; readonly width: number } | null {
  if (line[index] !== "[") return null;
  let cursor = index + 1;
  while (line[cursor] === "=") cursor += 1;
  if (line[cursor] !== "[") return null;
  return { equals: cursor - index - 1, width: cursor - index + 1 };
}

function luaLongBracketCloser(
  line: string,
  index: number,
  equals: number,
): number {
  if (line[index] !== "]") return 0;
  let cursor = index + 1;
  for (let count = 0; count < equals; count += 1) {
    if (line[cursor] !== "=") return 0;
    cursor += 1;
  }
  return line[cursor] === "]" ? cursor - index + 1 : 0;
}

function maskLuaLine(
  line: string,
  initialState: LuaLexicalState,
): { readonly masked: string; readonly state: LuaLexicalState } {
  let state = initialState;
  let masked = "";
  let escaped = false;
  for (let index = 0; index < line.length; index += 1) {
    if (state.kind === "long-string" || state.kind === "long-comment") {
      const width = luaLongBracketCloser(line, index, state.equals);
      if (width > 0) {
        masked += " ".repeat(width);
        index += width - 1;
        state = { kind: "code" };
      } else {
        masked += " ";
      }
      continue;
    }
    const char = line[index];
    if (state.kind === "single" || state.kind === "double") {
      masked += " ";
      if (escaped) {
        escaped = false;
      } else if (char === "\\") {
        escaped = true;
      } else if (
        (state.kind === "single" && char === "'") ||
        (state.kind === "double" && char === '"')
      ) {
        state = { kind: "code" };
      }
      continue;
    }

    const next = line[index + 1];
    if (char === "-" && next === "-") {
      const opener = luaLongBracketOpener(line, index + 2);
      if (opener) {
        masked += "  ".concat(" ".repeat(opener.width));
        index += opener.width + 1;
        state = { kind: "long-comment", equals: opener.equals };
      } else {
        masked += " ".repeat(line.length - index);
        break;
      }
      continue;
    }
    const longString = luaLongBracketOpener(line, index);
    if (longString) {
      masked += " ".repeat(longString.width);
      index += longString.width - 1;
      state = { kind: "long-string", equals: longString.equals };
      continue;
    }
    if (char === "'" || char === '"') {
      masked += " ";
      state = { kind: char === "'" ? "single" : "double" };
      escaped = false;
      continue;
    }
    masked += char;
  }
  // Lua short strings cannot span a newline without an escape. A malformed
  // source is ultimately rejected by the provider; keep later discovery
  // conservative instead of letting a stale quote hide top-level structure.
  if (state.kind === "single" || state.kind === "double") {
    state = { kind: "code" };
  }
  return { masked, state };
}

function discoverLuaTests(source: CodeTestSource): DiscoveredCodeTest[] {
  const tests: DiscoveredCodeTest[] = [];
  const lines = source.source.split("\n");
  let lexicalState: LuaLexicalState = { kind: "code" };
  let blockDepth = 0;
  for (let index = 0; index < lines.length; index += 1) {
    const startsInCode = lexicalState.kind === "code";
    const maskedLine = maskLuaLine(lines[index], lexicalState);
    lexicalState = maskedLine.state;
    const code = maskedLine.masked;
    if (blockDepth === 0 && startsInCode) {
      const match = /^\s*(?:local\s+)?function\s+(test[A-Za-z0-9_]*)\s*\(/.exec(
        code,
      );
      const name = match?.[1];
      if (name && LUA_TEST_NAME_PATTERN.test(name)) {
        const line = index + 1;
        tests.push({
          id: testId(source, name, line),
          pageId: source.pageId,
          path: source.path,
          language: source.language,
          name,
          line,
          execution: "sync",
        });
        if (tests.length >= CODE_TESTS_PER_FILE_LIMIT) break;
      }
    }

    const tokens = code.match(/[A-Za-z_][A-Za-z0-9_]*/g) ?? [];
    let loopAwaitingDo = 0;
    for (const token of tokens) {
      if (token === "end" || token === "until") {
        blockDepth = Math.max(0, blockDepth - 1);
      } else if (token === "function" || token === "if" || token === "repeat") {
        blockDepth += 1;
      } else if (token === "for" || token === "while") {
        blockDepth += 1;
        loopAwaitingDo += 1;
      } else if (token === "do") {
        if (loopAwaitingDo > 0) loopAwaitingDo -= 1;
        else blockDepth += 1;
      }
    }
  }
  return tests;
}

type RubyLexicalState =
  | { readonly kind: "code" | "single" | "double" | "block-comment" }
  | { readonly kind: "heredoc"; readonly terminator: string };

function maskRubyLine(
  line: string,
  initialState: RubyLexicalState,
): { readonly masked: string; readonly state: RubyLexicalState } {
  if (initialState.kind === "heredoc") {
    return {
      masked: " ".repeat(line.length),
      state:
        line.trim() === initialState.terminator
          ? { kind: "code" }
          : initialState,
    };
  }
  if (initialState.kind === "block-comment") {
    return {
      masked: " ".repeat(line.length),
      state: /^=end(?:\s|$)/.test(line)
        ? { kind: "code" }
        : initialState,
    };
  }
  if (initialState.kind === "code" && /^=begin(?:\s|$)/.test(line)) {
    return { masked: " ".repeat(line.length), state: { kind: "block-comment" } };
  }

  let state: RubyLexicalState = initialState;
  let masked = "";
  let escaped = false;
  for (let index = 0; index < line.length; index += 1) {
    const character = line[index];
    if (state.kind === "single" || state.kind === "double") {
      masked += " ";
      if (escaped) {
        escaped = false;
      } else if (character === "\\") {
        escaped = true;
      } else if (
        (state.kind === "single" && character === "'") ||
        (state.kind === "double" && character === '"')
      ) {
        state = { kind: "code" };
      }
      continue;
    }
    if (character === "#") {
      masked += " ".repeat(line.length - index);
      break;
    }
    if (character === "'" || character === '"') {
      state = { kind: character === "'" ? "single" : "double" };
      masked += " ";
      escaped = false;
      continue;
    }
    masked += character;
  }
  if (state.kind === "code") {
    const heredoc = /<<[-~]?\s*(["']?)([A-Za-z_][A-Za-z0-9_]*)\1/.exec(
      masked,
    );
    if (heredoc?.[2]) state = { kind: "heredoc", terminator: heredoc[2] };
  }
  return { masked, state };
}

function discoverRubyTests(source: CodeTestSource): DiscoveredCodeTest[] {
  const tests: DiscoveredCodeTest[] = [];
  const lines = source.source.split("\n");
  let lexicalState: RubyLexicalState = { kind: "code" };
  let blockDepth = 0;
  for (let index = 0; index < lines.length; index += 1) {
    const startsInCode = lexicalState.kind === "code";
    const maskedLine = maskRubyLine(lines[index], lexicalState);
    lexicalState = maskedLine.state;
    const code = maskedLine.masked.trim();
    if (!startsInCode || code === "") continue;
    if (code === "__END__") break;
    if (/^end(?:\s|$)/.test(code)) {
      blockDepth = Math.max(0, blockDepth - 1);
      continue;
    }
    if (blockDepth === 0) {
      const match = /^def\s+(test[\p{ID_Continue}]*)\s*(?:\(\s*\))?\s*$/u.exec(
        code,
      );
      const name = match?.[1];
      if (name && PYTHON_TEST_NAME_PATTERN.test(name)) {
        const line = index + 1;
        tests.push({
          id: testId(source, name, line),
          pageId: source.pageId,
          path: source.path,
          language: source.language,
          name,
          line,
          execution: "sync",
        });
        if (tests.length >= CODE_TESTS_PER_FILE_LIMIT) break;
      }
    }
    const leadingBlock =
      /^(?:class|module|def|if|unless|case|begin|while|until|for)\b/.test(
        code,
      );
    if (leadingBlock) {
      blockDepth += 1;
    } else if (/\bdo(?:\s*\|[^|]*\|)?\s*$/.test(code)) {
      blockDepth += 1;
    }
  }
  return tests;
}

interface ShellLexicalState {
  readonly quote: "code" | "single" | "double";
  readonly parameterDepth: number;
  readonly heredoc: string | null;
}

function maskShellLine(
  line: string,
  initialState: ShellLexicalState,
): { readonly masked: string; readonly state: ShellLexicalState } {
  if (initialState.heredoc !== null) {
    return {
      masked: " ".repeat(line.length),
      state:
        line.trim() === initialState.heredoc
          ? { quote: "code", parameterDepth: 0, heredoc: null }
          : initialState,
    };
  }

  let quote = initialState.quote;
  let parameterDepth = initialState.parameterDepth;
  let masked = "";
  let escaped = false;
  for (let index = 0; index < line.length; index += 1) {
    const character = line[index];
    if (quote !== "code") {
      masked += " ";
      if (escaped && quote === "double") {
        escaped = false;
      } else if (character === "\\" && quote === "double") {
        escaped = true;
      } else if (
        (quote === "single" && character === "'") ||
        (quote === "double" && character === '"')
      ) {
        quote = "code";
      }
      continue;
    }
    if (parameterDepth > 0) {
      masked += " ";
      if (character === "$" && line[index + 1] === "{") {
        masked += " ";
        parameterDepth += 1;
        index += 1;
      } else if (character === "}") {
        parameterDepth -= 1;
      }
      continue;
    }
    if (escaped) {
      masked += " ";
      escaped = false;
      continue;
    }
    if (character === "\\") {
      masked += " ";
      escaped = true;
      continue;
    }
    if (character === "'" || character === '"') {
      quote = character === "'" ? "single" : "double";
      masked += " ";
      continue;
    }
    if (character === "$" && line[index + 1] === "{") {
      masked += "  ";
      parameterDepth = 1;
      index += 1;
      continue;
    }
    if (
      character === "#" &&
      (index === 0 || /[\s;|&()]/.test(line[index - 1] ?? ""))
    ) {
      masked += " ".repeat(line.length - index);
      break;
    }
    masked += character;
  }

  let heredoc: string | null = null;
  if (quote === "code" && parameterDepth === 0) {
    for (const match of line.matchAll(
      /<<-?\s*(["']?)([A-Za-z_][A-Za-z0-9_]*)\1/g,
    )) {
      if (
        match.index !== undefined &&
        masked.slice(match.index, match.index + 2) === "<<"
      ) {
        heredoc = match[2] ?? null;
        break;
      }
    }
  }
  return { masked, state: { quote, parameterDepth, heredoc } };
}

function discoverShellTests(source: CodeTestSource): DiscoveredCodeTest[] {
  const tests: DiscoveredCodeTest[] = [];
  const lines = source.source.split("\n");
  let lexicalState: ShellLexicalState = {
    quote: "code",
    parameterDepth: 0,
    heredoc: null,
  };
  let braceDepth = 0;
  for (let index = 0; index < lines.length; index += 1) {
    const startsInCode =
      lexicalState.quote === "code" &&
      lexicalState.parameterDepth === 0 &&
      lexicalState.heredoc === null;
    const maskedLine = maskShellLine(lines[index], lexicalState);
    lexicalState = maskedLine.state;
    const code = maskedLine.masked;
    if (braceDepth === 0 && startsInCode) {
      const declaration = /^\s*(?:function\s+(test[A-Za-z0-9_]*)(?:\s*\(\s*\))?|(test[A-Za-z0-9_]*)\s*\(\s*\))\s*\{/.exec(
        code,
      );
      const name = declaration?.[1] ?? declaration?.[2];
      if (name && SHELL_TEST_NAME_PATTERN.test(name)) {
        const line = index + 1;
        tests.push({
          id: testId(source, name, line),
          pageId: source.pageId,
          path: source.path,
          language: source.language,
          name,
          line,
          execution: "sync",
        });
        if (tests.length >= CODE_TESTS_PER_FILE_LIMIT) break;
      }
    }
    for (const character of code) {
      if (character === "{") braceDepth += 1;
      else if (character === "}") braceDepth = Math.max(0, braceDepth - 1);
    }
  }
  return tests;
}

export function discoverCodeTests(
  sources: readonly CodeTestSource[],
): CodeTestDiscovery {
  const files: CodeTestFile[] = [];
  let totalTests = 0;
  let truncated = sources.length > CODE_TEST_WORKSPACE_FILE_LIMIT;
  const seenPages = new Set<string>();
  for (const source of sources.slice(0, CODE_TEST_WORKSPACE_FILE_LIMIT)) {
    if (
      !validPageId(source.pageId) ||
      !validPath(source.path) ||
      !isCodeTestLanguage(source.language) ||
      !Number.isSafeInteger(source.contentVersion) ||
      source.contentVersion < 0 ||
      typeof source.source !== "string" ||
      source.source.length > CODE_TEST_SOURCE_CODE_UNIT_LIMIT ||
      seenPages.has(source.pageId)
    ) {
      truncated = true;
      continue;
    }
    seenPages.add(source.pageId);
    const discovered =
      source.language === "python"
        ? discoverPythonTests(source)
        : source.language === "lua"
          ? discoverLuaTests(source)
          : source.language === "ruby"
            ? discoverRubyTests(source)
            : source.language === "bash"
              ? discoverShellTests(source)
              : discoverJavaScriptTests(source);
    const fileTruncated = discovered.length === CODE_TESTS_PER_FILE_LIMIT;
    if (fileTruncated) truncated = true;
    const remaining = CODE_TEST_TOTAL_LIMIT - totalTests;
    const tests = discovered.slice(0, Math.max(0, remaining));
    const totalTruncated = tests.length < discovered.length;
    if (totalTruncated) truncated = true;
    if (tests.length > 0) {
      files.push({
        source,
        tests,
        truncated: fileTruncated || totalTruncated,
      });
      totalTests += tests.length;
    }
    if (totalTests >= CODE_TEST_TOTAL_LIMIT) {
      if (files.length < sources.length) truncated = true;
      break;
    }
  }
  return { files, totalTests, truncated };
}

function safeHarnessIdentifier(marker: string): string {
  const suffix = marker.replace(/[^A-Za-z0-9]/g, "").slice(-48);
  return `__mrnote_test_${suffix || "receipt"}`;
}

function createSelectedJavaScriptDebugSource(
  source: CodeTestSource,
  test: DiscoveredCodeTest,
  bounds: SelectedJavaScriptTestDebugBounds,
): string | null {
  const lines = source.source.split("\n");
  const lineIndex = test.line - 1;
  let line = lines[lineIndex];
  if (line === undefined || bounds.endLine < test.line) return null;
  if (/^\s*export\b/.test(line)) {
    // TypeScript's existing transpiler removes module syntax before the debug
    // worker runs. Plain JavaScript has no module worker contract, so keep it
    // fail-closed rather than executing a partly rewritten module.
    if (source.language !== "typescript") return null;
    line = line.replace(/^(\s*)export\s+/, "$1");
    lines[lineIndex] = line;
  }
  if (/=\s*function\s*\*/.test(line)) return null;

  const escapedName = test.name.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const asyncDeclaration = new RegExp(
    `^(\\s*)async\\s+function\\s+${escapedName}\\s*\\(`,
  );
  const asyncExpression = new RegExp(
    `^(\\s*(?:const|let|var)\\s+${escapedName}\\s*(?::[^=]+)?=\\s*)async\\s+function(?!\\s*\\*)\\b`,
  );
  const asyncArrow = new RegExp(
    `^(\\s*(?:const|let|var)\\s+${escapedName}\\s*(?::[^=]+)?=\\s*)async\\s+(?:(?:\\([^)]*\\)|[$_\\p{ID_Start}][$\\u200c\\u200d\\p{ID_Continue}]*)(?:\\s*:[^=]+)?\\s*=>)`,
    "u",
  );
  if (test.execution === "async") {
    if (
      !asyncDeclaration.test(line) &&
      !asyncExpression.test(line) &&
      !asyncArrow.test(line)
    ) {
      return null;
    }
    return `${lines.join("\n")}${source.source.endsWith("\n") ? "" : "\n"};await ${test.name}();`;
  }

  const declaration = new RegExp(
    `^(\\s*)function\\s+${escapedName}\\s*\\(`,
  );
  const expression = new RegExp(
    `^(\\s*(?:const|let|var)\\s+${escapedName}\\s*(?::[^=]+)?=\\s*)function\\b`,
  );
  const arrow = new RegExp(
    `^(\\s*(?:const|let|var)\\s+${escapedName}\\s*(?::[^=]+)?=\\s*)(?!async\\b)(?:(?:\\([^)]*\\)|[$_\\p{ID_Start}][$\\u200c\\u200d\\p{ID_Continue}]*)(?:\\s*:[^=]+)?\\s*=>)`,
    "u",
  );

  if (declaration.test(line)) {
    lines[lineIndex] = line.replace(declaration, `$1async function ${test.name}(`);
  } else if (expression.test(line)) {
    lines[lineIndex] = line.replace(expression, "$1async function");
  } else if (arrow.test(line)) {
    lines[lineIndex] = line.replace(arrow, (match, prefix: string) =>
      `${prefix}async ${match.slice(prefix.length)}`,
    );
  } else {
    return null;
  }

  // The debugger inserts awaited hooks only where `await` is legal. Promoting
  // a synchronous declaration on its existing line makes its body inspectable
  // while preserving every authored line number. Authored async declarations
  // already provide that boundary and reach the early branch above unchanged.
  return `${lines.join("\n")}${source.source.endsWith("\n") ? "" : "\n"};await ${test.name}();`;
}

function returnExpressionIsDefinitelyNonThenable(expression: string): boolean {
  const value = expression.trim().replace(/;\s*$/, "").trim();
  if (value.length === 0) return true;
  if (/=>/.test(value)) return false;
  if (
    /^(?:true|false|null|[-+]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][-+]?\d+)?n?)$/.test(
      value,
    )
  ) {
    return true;
  }
  if (/^(?:!+|typeof\s+|void\s+)/.test(value)) return true;
  return false;
}

/**
 * Debugging can execute an authored async JS/TS declaration directly. A sync
 * declaration is promoted in place so the existing awaited breakpoint hook can
 * enter it without shifting lines; because that promotion could hide a returned
 * thenable, sync tests accept only syntax that proves a primitive result.
 */
interface SelectedJavaScriptTestDebugBounds {
  readonly endLine: number;
}

function selectedJavaScriptTestDebugBounds(
  source: CodeTestSource,
  test: DiscoveredCodeTest,
  requirePrimitiveReturn: boolean,
): SelectedJavaScriptTestDebugBounds | null {
  const lines = source.source.split("\n");
  const start = test.line - 1;
  let lexicalState: JavaScriptLexicalState = "code";
  const maskedLines: string[] = [];
  let opened = false;
  let depth = 0;

  for (let index = 0; index < lines.length; index += 1) {
    const masked = maskJavaScriptLine(lines[index], lexicalState);
    lexicalState = masked.state;
    if (index < start) continue;

    const code = masked.masked;
    if (index === start) {
      const arrowIndex = code.indexOf("=>");
      if (arrowIndex >= 0) {
        const tail = code.slice(arrowIndex + 2).trim();
        if (!tail.startsWith("{")) {
          const returnIsSafe =
            !requirePrimitiveReturn ||
            returnExpressionIsDefinitelyNonThenable(tail);
          return returnIsSafe ? { endLine: test.line } : null;
        }
      }
    }
    maskedLines.push(code);
    for (const char of code) {
      if (char === "{") {
        opened = true;
        depth += 1;
      } else if (char === "}" && opened) {
        depth -= 1;
      }
    }
    if (opened && depth === 0) break;
  }
  if (!opened || depth !== 0) return null;

  if (requirePrimitiveReturn) {
    for (const code of maskedLines) {
      const pattern = /\breturn\b/g;
      let match: RegExpExecArray | null;
      while ((match = pattern.exec(code))) {
        const tail = code.slice(match.index + match[0].length);
        const terminator = tail.search(/[;}]/);
        if (terminator < 0 && tail.trim().length > 0) return null;
        const expression = terminator < 0 ? "" : tail.slice(0, terminator);
        if (!returnExpressionIsDefinitelyNonThenable(expression)) return null;
      }
    }
  }
  return { endLine: start + maskedLines.length };
}

function createSelectedPythonDebugSource(
  source: CodeTestSource,
  test: DiscoveredCodeTest,
): string | null {
  const identifier = safeHarnessIdentifier(
    `${source.pageId}-${test.line}-${test.name}`,
  );
  if (test.execution === "async") {
    const suffix = `
import inspect as ${identifier}_inspect
${identifier}_value = ${test.name}()
if not ${identifier}_inspect.isawaitable(${identifier}_value) or ${identifier}_inspect.isasyncgen(${identifier}_value):
    ${identifier}_close = getattr(${identifier}_value, "close", None)
    if callable(${identifier}_close):
        ${identifier}_close()
    raise TypeError("Async test did not return an awaitable")
__mrnote_debug_selected_test_awaitable__ = ${identifier}_value`;
    return `${source.source}${source.source.endsWith("\n") ? "" : "\n"}${suffix.trimStart()}`;
  }
  const suffix = `
import inspect as ${identifier}_inspect
${identifier}_value = ${test.name}()
if ${identifier}_inspect.isawaitable(${identifier}_value) or ${identifier}_inspect.isgenerator(${identifier}_value):
    ${identifier}_close = getattr(${identifier}_value, "close", None)
    if callable(${identifier}_close):
        ${identifier}_close()
    raise TypeError("Async tests are unavailable")`;
  return `${source.source}${source.source.endsWith("\n") ? "" : "\n"}${suffix.trimStart()}`;
}

/**
 * Re-authorize one failed persisted result against a fresh source snapshot and
 * build a selected-only debugger input without shifting authored lines.
 */
export function createCodeTestDebugHarness(
  source: CodeTestSource,
  receipt: PersistedCodeTestRunReceipt,
  testIdValue: string,
  exactReceiptSource: CodeTestSource,
): CodeTestDebugHarness | null {
  // Lua and Ruby have real bounded execution providers but no debugger
  // provider. Keep Debug Test fail-closed instead of feeding either source to
  // the JavaScript/Python debugger.
  if (
    source.language === "lua" ||
    source.language === "ruby" ||
    source.language === "bash"
  ) {
    return null;
  }
  const trustedReceipt = restoreCodeTestRunReceipt({
    [CODE_TEST_RUN_RECEIPT_KEY]: receipt,
  });
  if (!trustedReceipt) return null;

  const sourceIdentity = trustedReceipt.sources.find(
    (candidate) => candidate.pageId === source.pageId,
  );
  if (
    !sourceIdentity ||
    exactReceiptSource.pageId !== source.pageId ||
    exactReceiptSource.path !== source.path ||
    exactReceiptSource.language !== source.language ||
    exactReceiptSource.contentVersion !== source.contentVersion ||
    exactReceiptSource.source !== source.source ||
    sourceIdentity.path !== source.path ||
    sourceIdentity.language !== source.language ||
    sourceIdentity.sourceFingerprint !==
      codeTestSourceFingerprint(exactReceiptSource.source)
  ) {
    return null;
  }

  const matchingResults = trustedReceipt.results.filter(
    (candidate) => candidate.testId === testIdValue,
  );
  const result = matchingResults.length === 1 ? matchingResults[0] : null;
  if (
    !result ||
    result.status !== "failed" ||
    result.id !== `${testIdValue}:failed` ||
    result.pageId !== source.pageId ||
    result.path !== source.path ||
    result.language !== source.language ||
    result.line > source.source.split("\n").length
  ) {
    return null;
  }

  const discovery = discoverCodeTests([source]);
  const file = discovery.files[0];
  if (
    discovery.truncated ||
    discovery.files.length !== 1 ||
    !file ||
    file.truncated ||
    file.source.pageId !== source.pageId
  ) {
    return null;
  }
  const test = file.tests.find((candidate) => candidate.id === testIdValue);
  if (
    !test ||
    result.name !== test.name ||
    result.testId !== test.id ||
    result.pageId !== test.pageId ||
    result.path !== test.path ||
    result.language !== test.language
  ) {
    return null;
  }

  const javascriptBounds =
    source.language === "python"
      ? null
      : selectedJavaScriptTestDebugBounds(
          source,
          test,
          test.execution === "sync",
        );
  if (
    source.language !== "python" &&
    (!javascriptBounds ||
      result.line < test.line ||
      result.line > javascriptBounds.endLine)
  ) {
    return null;
  }
  const executionSource = source.language === "python"
    ? createSelectedPythonDebugSource(source, test)
    : createSelectedJavaScriptDebugSource(source, test, javascriptBounds!);
  if (
    executionSource === null ||
    executionSource.length > CODE_TEST_SOURCE_CODE_UNIT_LIMIT + 4_096
  ) {
    return null;
  }
  return {
    source: { ...source },
    test: { ...test },
    result: { ...result },
    executionSource,
    authoredLineCount: source.source.split("\n").length,
  };
}

const CODE_TEST_DEBUG_HANDOFF_LIMIT = 8;
const CODE_TEST_DEBUG_HANDOFF_TTL_MS = 30_000;
const pendingCodeTestDebugHandoffs = new Map<string, CodeTestDebugHandoff>();
const codeTestDebugHandoffListeners = new Map<string, Set<() => void>>();
let codeTestDebugHandoffSequence = 0;

function purgeExpiredCodeTestDebugHandoffs(now: number): void {
  for (const [pageId, handoff] of pendingCodeTestDebugHandoffs) {
    if (handoff.expiresAt <= now) pendingCodeTestDebugHandoffs.delete(pageId);
  }
  while (pendingCodeTestDebugHandoffs.size >= CODE_TEST_DEBUG_HANDOFF_LIMIT) {
    const oldestPageId = pendingCodeTestDebugHandoffs.keys().next().value;
    if (typeof oldestPageId !== "string") break;
    pendingCodeTestDebugHandoffs.delete(oldestPageId);
  }
}

/** Publish a short-lived, take-once target for an exact destination window. */
export function publishCodeTestDebugHandoff(input: {
  readonly harness: CodeTestDebugHarness;
  readonly receipt: PersistedCodeTestRunReceipt;
  readonly targetBufferId?: string;
}): CodeTestDebugHandoff {
  const now = Date.now();
  purgeExpiredCodeTestDebugHandoffs(now);
  codeTestDebugHandoffSequence += 1;
  const handoff: CodeTestDebugHandoff = {
    ...input.harness,
    requestId: `code-test-debug-${codeTestDebugHandoffSequence.toString(36)}`,
    receipt: input.receipt,
    ...(input.targetBufferId ? { targetBufferId: input.targetBufferId } : {}),
    expiresAt: now + CODE_TEST_DEBUG_HANDOFF_TTL_MS,
  };
  pendingCodeTestDebugHandoffs.delete(handoff.source.pageId);
  pendingCodeTestDebugHandoffs.set(handoff.source.pageId, handoff);
  for (const listener of [
    ...(codeTestDebugHandoffListeners.get(handoff.source.pageId) ?? []),
  ]) {
    listener();
  }
  return handoff;
}

/** Consume an exact target once its destination page has fully hydrated. */
export function takeCodeTestDebugHandoff(
  pageId: string,
  bufferId: string,
): CodeTestDebugHandoff | null {
  const now = Date.now();
  purgeExpiredCodeTestDebugHandoffs(now);
  const handoff = pendingCodeTestDebugHandoffs.get(pageId);
  if (
    !handoff ||
    (handoff.targetBufferId !== undefined &&
      handoff.targetBufferId !== bufferId)
  ) {
    return null;
  }
  pendingCodeTestDebugHandoffs.delete(pageId);
  return handoff;
}

export function subscribeCodeTestDebugHandoff(
  pageId: string,
  listener: () => void,
): () => void {
  const listeners = codeTestDebugHandoffListeners.get(pageId) ?? new Set();
  listeners.add(listener);
  codeTestDebugHandoffListeners.set(pageId, listeners);
  return () => {
    listeners.delete(listener);
    if (listeners.size === 0) codeTestDebugHandoffListeners.delete(pageId);
  };
}

export function discardCodeTestDebugHandoff(
  pageId: string,
  bufferId: string,
): void {
  const handoff = pendingCodeTestDebugHandoffs.get(pageId);
  if (
    !handoff ||
    (handoff.targetBufferId !== undefined &&
      handoff.targetBufferId !== bufferId)
  ) {
    return;
  }
  pendingCodeTestDebugHandoffs.delete(pageId);
}

/** Retire one exact publisher-owned request without affecting a replacement. */
export function discardCodeTestDebugHandoffRequest(requestId: string): void {
  for (const [pageId, handoff] of pendingCodeTestDebugHandoffs) {
    if (handoff.requestId !== requestId) continue;
    pendingCodeTestDebugHandoffs.delete(pageId);
    return;
  }
}

export function _clearCodeTestDebugHandoffsForTest(): void {
  pendingCodeTestDebugHandoffs.clear();
  codeTestDebugHandoffListeners.clear();
  codeTestDebugHandoffSequence = 0;
}

function createJavaScriptHarness(
  source: CodeTestSource,
  tests: readonly DiscoveredCodeTest[],
  marker: string,
): CodeTestHarness {
  const identifier = safeHarnessIdentifier(marker);
  const specs = tests.map((test) => ({
    id: test.id,
    name: test.name,
    line: test.line,
    async: test.execution === "async",
  }));
  const awaitsCompletion = specs.some((test) => test.async);
  const references = tests
    .map(
      (test) =>
        `${JSON.stringify(test.id)}:function(){return ${test.name};}`,
    )
    .join(",");
  // `_base` is the line this bootstrap runs on, as the runtime numbers it. The
  // authored source starts on the next line, so a failure frame at runtime
  // line N is authored line N - base, whatever wrapper the provider (worker,
  // sandbox iframe, rich-output shell, module bundle) put in front of it.
  const bootstrap = `const ${identifier}_emit=console.log.bind(console),${identifier}_json=JSON.stringify.bind(JSON),${identifier}_now=Date.now.bind(Date),${identifier}_string=String,${identifier}_base=(()=>{try{const m=/<anonymous>:(\\d+):\\d+/.exec(new Error().stack);return m?Number(m[1]):0;}catch{return 0;}})(),${identifier}_refs={${references}};`;
  const invokeTest = awaitsCompletion
    ? `if(spec.async){if(!returned||typeof returned.then!=="function")throw new TypeError("Async test is not awaitable");await returned;}else if(returned&&typeof returned.then==="function")throw new TypeError("Async tests require an async declaration");`
    : `if(returned&&typeof returned.then==="function")throw new TypeError("Async tests require an async declaration");`;
  const runSpecs = `const specs=${JSON.stringify(specs)};for(const spec of specs){const started=${identifier}_now();let status="passed",message,line=spec.line;try{const value=${identifier}_refs[spec.id]();if(typeof value!=="function")throw new TypeError("Test is not callable");const returned=value();${invokeTest}}catch(error){status="failed";try{message=error&&typeof error.message==="string"?(error.name?error.name+": ":"")+error.message:${identifier}_string(error);}catch{message="Test failed";}try{const stack=error&&typeof error.stack==="string"?error.stack:"";const match=/<anonymous>:(\\d+):\\d+/.exec(stack);const candidate=match&&${identifier}_base?Number(match[1])-${identifier}_base:0;if(candidate>=1&&candidate<=${source.source.split("\n").length})line=candidate;}catch{}}${identifier}_emit(${JSON.stringify(marker)}+${identifier}_json({id:spec.id,status,line,durationMs:Math.max(0,${identifier}_now()-started),...(message?{message}: {})}));}`;
  const harness = awaitsCompletion
    ? `\n;return (async()=>{${runSpecs}})();`
    : `\n;(()=>{${runSpecs}})();`;
  return {
    source: `${awaitsCompletion ? CODE_TEST_ASYNC_HARNESS_PREFIX : ""}${bootstrap}\n${source.source}${harness}`,
    marker,
    tests: [...tests],
    execution: awaitsCompletion ? "async" : "sync",
  };
}

function createPythonHarness(
  source: CodeTestSource,
  tests: readonly DiscoveredCodeTest[],
  marker: string,
): CodeTestHarness {
  const identifier = safeHarnessIdentifier(marker);
  const specs = tests.map((test) => ({
    id: test.id,
    name: test.name,
    line: test.line,
    // Keep the serialized object a valid Python literal: JSON's lowercase
    // booleans are not Python names, while 0/1 retain the same truth contract.
    async: test.execution === "async" ? 1 : 0,
  }));
  const awaitsCompletion = specs.some((test) => test.async);
  const harness = `
import builtins as ${identifier}_builtins, json as ${identifier}_json, time as ${identifier}_time, traceback as ${identifier}_traceback, inspect as ${identifier}_inspect
for ${identifier}_spec in ${JSON.stringify(specs)}:
    ${identifier}_started = ${identifier}_time.monotonic()
    ${identifier}_status = "passed"
    ${identifier}_message = None
    ${identifier}_line = ${identifier}_spec["line"]
    try:
        ${identifier}_fn = globals().get(${identifier}_spec["name"])
        if not callable(${identifier}_fn):
            raise TypeError("Test is not callable")
        ${identifier}_returned = ${identifier}_fn()
        if ${identifier}_spec["async"]:
            if not ${identifier}_inspect.isawaitable(${identifier}_returned):
                raise TypeError("Async test is not awaitable")
            await ${identifier}_returned
        elif ${identifier}_inspect.isawaitable(${identifier}_returned):
            ${identifier}_close = getattr(${identifier}_returned, "close", None)
            if callable(${identifier}_close):
                ${identifier}_close()
            raise TypeError("Async tests require an async declaration")
    except BaseException as ${identifier}_error:
        ${identifier}_status = "failed"
        ${identifier}_message = type(${identifier}_error).__name__ + ": " + str(${identifier}_error)
        ${identifier}_frames = ${identifier}_traceback.extract_tb(${identifier}_error.__traceback__)
        for ${identifier}_frame in reversed(${identifier}_frames):
            if 1 <= ${identifier}_frame.lineno <= ${source.source.split("\n").length}:
                ${identifier}_line = ${identifier}_frame.lineno
                break
    ${identifier}_payload = {"id": ${identifier}_spec["id"], "status": ${identifier}_status, "line": ${identifier}_line, "durationMs": max(0, int((${identifier}_time.monotonic() - ${identifier}_started) * 1000))}
    if ${identifier}_message:
        ${identifier}_payload["message"] = ${identifier}_message
    ${identifier}_builtins.print(${JSON.stringify(marker)} + ${identifier}_json.dumps(${identifier}_payload, separators=(",", ":")))
`;
  return {
    source: `${source.source}${source.source.endsWith("\n") ? "" : "\n"}${harness}`,
    marker,
    tests: [...tests],
    execution: awaitsCompletion ? "async" : "sync",
  };
}

function luaStringLiteral(value: string): string {
  return `"${value.replace(/[\u0000-\u001f\\"]/g, (character) => {
    if (character === "\\") return "\\\\";
    if (character === '"') return '\\"';
    if (character === "\n") return "\\n";
    if (character === "\r") return "\\r";
    if (character === "\t") return "\\t";
    return `\\x${character.charCodeAt(0).toString(16).padStart(2, "0")}`;
  })}"`;
}

function createLuaHarness(
  source: CodeTestSource,
  tests: readonly DiscoveredCodeTest[],
  marker: string,
): CodeTestHarness {
  const identifier = safeHarnessIdentifier(marker);
  const specs = tests
    .map(
      (test) =>
        `{id=${luaStringLiteral(test.id)},line=${test.line},resolve=function() return ${test.name} end}`,
    )
    .join(",");
  // Capture the safe standard-library surface before authored code can replace
  // globals. Keeping this bootstrap on one physical line makes the source-line
  // correction exact and deterministic.
  const bootstrap = `local ${identifier}_print,${identifier}_pcall,${identifier}_type,${identifier}_tostring,${identifier}_tonumber,${identifier}_error,${identifier}_ipairs,${identifier}_string,${identifier}_table=print,pcall,type,tostring,tonumber,error,ipairs,string,table`;
  const harness = `
local ${identifier}_specs={${specs}}
local function ${identifier}_quote(value)
  local text=${identifier}_tostring(value)
  local parts={'"'}
  for index=1,#text do
    local byte=${identifier}_string.byte(text,index)
    if byte==34 then ${identifier}_table.insert(parts,${identifier}_string.char(92,34))
    elseif byte==92 then ${identifier}_table.insert(parts,${identifier}_string.char(92,92))
    elseif byte==8 then ${identifier}_table.insert(parts,${identifier}_string.char(92)..'b')
    elseif byte==9 then ${identifier}_table.insert(parts,${identifier}_string.char(92)..'t')
    elseif byte==10 then ${identifier}_table.insert(parts,${identifier}_string.char(92)..'n')
    elseif byte==12 then ${identifier}_table.insert(parts,${identifier}_string.char(92)..'f')
    elseif byte==13 then ${identifier}_table.insert(parts,${identifier}_string.char(92)..'r')
    elseif byte<32 then ${identifier}_table.insert(parts,${identifier}_string.format(${identifier}_string.char(92)..'u%04x',byte))
    else ${identifier}_table.insert(parts,${identifier}_string.char(byte)) end
  end
  ${identifier}_table.insert(parts,'"')
  return ${identifier}_table.concat(parts)
end
for _,spec in ${identifier}_ipairs(${identifier}_specs) do
  local status='passed'
  local message=nil
  local line=spec.line
  local ok,returned=${identifier}_pcall(function()
    local fn=spec.resolve()
    if ${identifier}_type(fn)~='function' then ${identifier}_error('Test is not callable') end
    return fn()
  end)
  if not ok then
    status='failed'
    message=${identifier}_string.sub(${identifier}_tostring(returned),1,480)
    local candidate=${identifier}_string.match(message,'main%.lua:(%d+):')
    candidate=candidate and ${identifier}_tonumber(candidate) or nil
    if candidate and candidate>=2 and candidate<=${source.source.split("\n").length + 1} then line=candidate-1 end
  elseif ${identifier}_type(returned)=='thread' then
    status='failed'
    message='Async tests are unavailable'
  end
  local payload='{"id":'..${identifier}_quote(spec.id)..',"status":'..${identifier}_quote(status)..',"line":'..line..',"durationMs":0'
  if message then payload=payload..',"message":'..${identifier}_quote(message) end
  ${identifier}_print(${luaStringLiteral(marker)}..payload..'}')
end`;
  return {
    source: `${bootstrap}\n${source.source}${source.source.endsWith("\n") ? "" : "\n"}${harness.trimStart()}`,
    marker,
    tests: [...tests],
    execution: "sync",
  };
}

function rubyStringLiteral(value: string): string {
  const serialized = JSON.stringify(value);
  return typeof serialized === "string" ? serialized : '""';
}

function createRubyHarness(
  source: CodeTestSource,
  tests: readonly DiscoveredCodeTest[],
  marker: string,
): CodeTestHarness | null {
  if (source.source.split("\n").some((line) => line === "__END__")) {
    return null;
  }
  const identifier = safeHarnessIdentifier(marker);
  const specs = tests
    .map(
      (test) =>
        `{id:${rubyStringLiteral(test.id)},name:${rubyStringLiteral(test.name)},line:${test.line}}`,
    )
    .join(",");
  // Capture the small standard-library surface before authored Ruby can
  // redefine core methods. This bootstrap occupies exactly one physical line;
  // authored backtrace locations are therefore mapped back by subtracting one.
  const bootstrap = `require "json";${identifier}_puts=Kernel.method(:puts);${identifier}_json=JSON.method(:generate);${identifier}_resolve=Kernel.instance_method(:method);${identifier}_call=Method.instance_method(:call);${identifier}_hash_get=Hash.instance_method(:[]);${identifier}_hash_set=Hash.instance_method(:[]=);${identifier}_each=Array.instance_method(:each);${identifier}_to_sym=String.instance_method(:to_sym);${identifier}_backtrace=Exception.instance_method(:backtrace_locations);${identifier}_message=Exception.instance_method(:message);${identifier}_object_class=Object.instance_method(:class);${identifier}_module_name=Module.instance_method(:name);${identifier}_slice=String.instance_method(:slice);${identifier}_is_a=Object.instance_method(:is_a?);${identifier}_fiber=Fiber;${identifier}_thread=Thread;${identifier}_enumerator=Enumerator`;
  const harness = `
${identifier}_specs=[${specs}]
${identifier}_each.bind_call(${identifier}_specs) do |${identifier}_spec|
  ${identifier}_status="passed"
  ${identifier}_message_value=nil
  ${identifier}_line=${identifier}_hash_get.bind_call(${identifier}_spec,:line)
  begin
    ${identifier}_name=${identifier}_hash_get.bind_call(${identifier}_spec,:name)
    ${identifier}_symbol=${identifier}_to_sym.bind_call(${identifier}_name)
    ${identifier}_method=${identifier}_resolve.bind_call(self,${identifier}_symbol)
    ${identifier}_returned=${identifier}_call.bind_call(${identifier}_method)
    if ${identifier}_is_a.bind_call(${identifier}_returned,${identifier}_fiber) || ${identifier}_is_a.bind_call(${identifier}_returned,${identifier}_thread) || ${identifier}_is_a.bind_call(${identifier}_returned,${identifier}_enumerator)
      ${identifier}_status="failed"
      ${identifier}_message_value="Async tests are unavailable"
    end
  rescue Exception => ${identifier}_error
    ${identifier}_status="failed"
    ${identifier}_class=${identifier}_object_class.bind_call(${identifier}_error)
    ${identifier}_class_name=${identifier}_module_name.bind_call(${identifier}_class)
    ${identifier}_error_message=${identifier}_message.bind_call(${identifier}_error)
    ${identifier}_message_value=${identifier}_slice.bind_call("#{${identifier}_class_name}: #{${identifier}_error_message}",0,480)
    ${identifier}_locations=${identifier}_backtrace.bind_call(${identifier}_error)
    if ${identifier}_locations
      ${identifier}_each.bind_call(${identifier}_locations) do |${identifier}_location|
        ${identifier}_candidate=${identifier}_location.lineno
        if ${identifier}_candidate>=2 && ${identifier}_candidate<=${source.source.split("\n").length + 1}
          ${identifier}_line=${identifier}_candidate-1
          break
        end
      end
    end
  end
  ${identifier}_payload={"id"=>${identifier}_hash_get.bind_call(${identifier}_spec,:id),"status"=>${identifier}_status,"line"=>${identifier}_line,"durationMs"=>0}
  ${identifier}_hash_set.bind_call(${identifier}_payload,"message",${identifier}_message_value) if ${identifier}_message_value
  ${identifier}_encoded=${identifier}_call.bind_call(${identifier}_json,${identifier}_payload)
  ${identifier}_call.bind_call(${identifier}_puts,${rubyStringLiteral(marker)}+${identifier}_encoded)
end`;
  return {
    source: `${bootstrap}\n${source.source}${source.source.endsWith("\n") ? "" : "\n"}${harness.trimStart()}`,
    marker,
    tests: [...tests],
    execution: "sync",
  };
}

function shellSingleQuoted(value: string): string {
  return `'${value.replaceAll("'", `'"'"'`)}'`;
}

function createShellHarness(
  source: CodeTestSource,
  tests: readonly DiscoveredCodeTest[],
  marker: string,
): CodeTestHarness | null {
  if (tests.some((test) => !SHELL_TEST_NAME_PATTERN.test(test.name))) {
    return null;
  }
  const identifier = safeHarnessIdentifier(marker);
  const harness = tests
    .map((test, index) => {
      const statusVariable = `${identifier}_status_${index}`;
      const passed = `${marker}${JSON.stringify({
        id: test.id,
        status: "passed",
        line: test.line,
        durationMs: 0,
      })}`;
      const failedPrefix = `${marker}{"id":${JSON.stringify(test.id)},"status":"failed","line":${test.line},"durationMs":0,"message":"Shell test exited with status `;
      return `set +e
${test.name}
${statusVariable}=$?
if [ "$${statusVariable}" -eq 0 ]; then
  builtin printf '%s\\n' ${shellSingleQuoted(passed)}
else
  builtin printf '%s%s%s\\n' ${shellSingleQuoted(failedPrefix)} "$${statusVariable}" ${shellSingleQuoted('"}')}
fi`;
    })
    .join("\n");
  return {
    source: `${source.source}${source.source.endsWith("\n") ? "" : "\n"}${harness}\n`,
    marker,
    tests: [...tests],
    execution: "sync",
  };
}

export function createCodeTestHarness(
  source: CodeTestSource,
  tests: readonly DiscoveredCodeTest[],
  nonce: string,
): CodeTestHarness | null {
  if (
    !isCodeTestLanguage(source.language) ||
    source.source.length > CODE_TEST_SOURCE_CODE_UNIT_LIMIT ||
    tests.length === 0 ||
    tests.length > CODE_TESTS_PER_FILE_LIMIT ||
    !/^[A-Za-z0-9_-]{16,96}$/.test(nonce)
  ) {
    return null;
  }
  const discoveredById = new Map(
    discoverCodeTests([source]).files.flatMap((file) => file.tests).map((test) => [test.id, test] as const),
  );
  const selected: DiscoveredCodeTest[] = [];
  const seen = new Set<string>();
  for (const test of tests) {
    const discovered = discoveredById.get(test.id);
    if (
      !discovered ||
      seen.has(test.id) ||
      discovered.pageId !== source.pageId ||
      discovered.name !== test.name ||
      discovered.line !== test.line
    ) {
      return null;
    }
    seen.add(test.id);
    selected.push(discovered);
  }
  const marker = `${CODE_TEST_MARKER_PREFIX}:${nonce}:`;
  return source.language === "python"
    ? createPythonHarness(source, selected, marker)
    : source.language === "lua"
      ? createLuaHarness(source, selected, marker)
      : source.language === "ruby"
        ? createRubyHarness(source, selected, marker)
        : source.language === "bash"
          ? createShellHarness(source, selected, marker)
          : createJavaScriptHarness(source, selected, marker);
}

function parseHarnessPayload(value: unknown): HarnessPayload | null {
  if (!value || typeof value !== "object" || Array.isArray(value)) return null;
  const raw = value as Record<string, unknown>;
  const keys = Object.keys(raw);
  if (keys.some((key) => !["id", "status", "line", "durationMs", "message"].includes(key))) {
    return null;
  }
  const id = boundedString(raw.id, 512);
  const line = finiteInteger(raw.line, 1, CODE_TEST_RESULT_LINE_LIMIT);
  const durationMs = finiteInteger(raw.durationMs, 0, CODE_TEST_DURATION_LIMIT_MS);
  const message = raw.message === undefined ? undefined : boundedString(raw.message, CODE_TEST_MESSAGE_LIMIT);
  if (
    !id ||
    (raw.status !== "passed" && raw.status !== "failed") ||
    line === null ||
    durationMs === null ||
    (raw.message !== undefined && message === null) ||
    (raw.status === "passed" && message !== undefined)
  ) {
    return null;
  }
  return {
    id,
    status: raw.status,
    line,
    durationMs,
    ...(message ? { message } : {}),
  };
}

function fallbackResults(
  harness: CodeTestHarness,
  output: RunResult,
): CodeTestResult[] | null {
  const message =
    typeof output.error === "string" && output.error.length > 0
      ? output.error.slice(0, CODE_TEST_MESSAGE_LIMIT)
      : output.timedOut
        ? "Test run timed out"
        : null;
  if (!message || output.interrupted) return null;
  return harness.tests.map((test) => ({
    id: `${test.id}:failed`,
    testId: test.id,
    pageId: test.pageId,
    path: test.path,
    language: test.language,
    name: test.name,
    line: test.line,
    status: "failed",
    message,
    durationMs: 0,
  }));
}

export function parseCodeTestHarnessResult(
  harness: CodeTestHarness,
  output: RunResult,
): ParsedCodeTestHarnessResult {
  const byId = new Map(harness.tests.map((test) => [test.id, test] as const));
  const parsed = new Map<string, CodeTestResult>();
  for (const log of output.logs) {
    if (typeof log !== "string" || !log.startsWith(harness.marker)) continue;
    let payload: unknown;
    try {
      payload = JSON.parse(log.slice(harness.marker.length));
    } catch {
      return { results: [], complete: false };
    }
    const normalized = parseHarnessPayload(payload);
    const test = normalized ? byId.get(normalized.id) : null;
    if (!normalized || !test || parsed.has(test.id)) {
      return { results: [], complete: false };
    }
    parsed.set(test.id, {
      id: `${test.id}:${normalized.status}`,
      testId: test.id,
      pageId: test.pageId,
      path: test.path,
      language: test.language,
      name: test.name,
      line: normalized.line,
      status: normalized.status,
      ...(normalized.message ? { message: normalized.message } : {}),
      durationMs: normalized.durationMs,
    });
  }
  if (parsed.size !== harness.tests.length) {
    const fallback = fallbackResults(harness, output);
    return fallback
      ? { results: fallback, complete: true }
      : { results: [], complete: false };
  }
  return {
    results: harness.tests.map((test) => parsed.get(test.id)!),
    complete: true,
  };
}

function normalizeSourceIdentity(value: unknown): CodeTestSourceIdentity | null {
  if (!value || typeof value !== "object" || Array.isArray(value)) return null;
  const raw = value as Record<string, unknown>;
  if (Object.keys(raw).some((key) => !["pageId", "path", "language", "sourceFingerprint"].includes(key))) {
    return null;
  }
  const sourceFingerprint = boundedString(raw.sourceFingerprint, 64);
  if (
    !validPageId(raw.pageId) ||
    !validPath(raw.path) ||
    typeof raw.language !== "string" ||
    !isCodeTestLanguage(raw.language) ||
    !sourceFingerprint ||
    !/^v1:[0-9]+:[0-9a-z]+:[0-9a-z]+$/.test(sourceFingerprint)
  ) {
    return null;
  }
  return {
    pageId: raw.pageId,
    path: raw.path,
    language: raw.language,
    sourceFingerprint,
  };
}

function normalizeTestResult(value: unknown): CodeTestResult | null {
  if (!value || typeof value !== "object" || Array.isArray(value)) return null;
  const raw = value as Record<string, unknown>;
  if (Object.keys(raw).some((key) => !["id", "testId", "pageId", "path", "language", "name", "line", "status", "message", "durationMs"].includes(key))) {
    return null;
  }
  const id = boundedString(raw.id, 768);
  const testIdValue = boundedString(raw.testId, 512);
  const name = boundedString(raw.name, CODE_TEST_NAME_LIMIT);
  const line = finiteInteger(raw.line, 1, CODE_TEST_RESULT_LINE_LIMIT);
  const durationMs = finiteInteger(raw.durationMs, 0, CODE_TEST_DURATION_LIMIT_MS);
  const message = raw.message === undefined ? undefined : boundedString(raw.message, CODE_TEST_MESSAGE_LIMIT);
  if (
    !id ||
    !testIdValue ||
    !validPageId(raw.pageId) ||
    !validPath(raw.path) ||
    typeof raw.language !== "string" ||
    !isCodeTestLanguage(raw.language) ||
    !name ||
    !TEST_NAME_PATTERN.test(name) ||
    line === null ||
    durationMs === null ||
    (raw.status !== "passed" && raw.status !== "failed") ||
    (raw.message !== undefined && message === null) ||
    (raw.status === "passed" && message !== undefined)
  ) {
    return null;
  }
  return {
    id,
    testId: testIdValue,
    pageId: raw.pageId,
    path: raw.path,
    language: raw.language,
    name,
    line,
    status: raw.status,
    ...(message ? { message } : {}),
    durationMs,
  };
}

export function createPersistedCodeTestRunReceipt(input: {
  readonly runId: string;
  readonly scope: CodeTestRunScope;
  readonly completedAt: string;
  readonly sources: readonly CodeTestSource[];
  readonly results: readonly CodeTestResult[];
}): PersistedCodeTestRunReceipt | null {
  if (
    input.runId.length === 0 ||
    input.runId.length > CODE_TEST_RUN_ID_LIMIT ||
    !["selected", "file", "all"].includes(input.scope) ||
    !Number.isFinite(Date.parse(input.completedAt)) ||
    input.sources.length === 0 ||
    input.sources.length > CODE_TEST_WORKSPACE_FILE_LIMIT ||
    input.results.length === 0 ||
    input.results.length > CODE_TEST_TOTAL_LIMIT
  ) {
    return null;
  }
  const sources = input.sources.map(codeTestSourceIdentity);
  const sourcePages = new Set<string>();
  for (const source of sources) {
    if (sourcePages.has(source.pageId)) return null;
    sourcePages.add(source.pageId);
  }
  const results = input.results.map(normalizeTestResult);
  if (results.some((result) => result === null)) return null;
  const resultIds = new Set<string>();
  for (const result of results as CodeTestResult[]) {
    if (!sourcePages.has(result.pageId) || resultIds.has(result.testId)) return null;
    resultIds.add(result.testId);
  }
  return {
    version: CODE_TEST_RUN_RECEIPT_VERSION,
    runId: input.runId,
    scope: input.scope,
    completedAt: input.completedAt,
    sources,
    results: results as CodeTestResult[],
  };
}

function normalizePersistedCodeTestRunReceipt(
  value: unknown,
): PersistedCodeTestRunReceipt | null {
  if (!value || typeof value !== "object" || Array.isArray(value)) return null;
  const raw = value as Record<string, unknown>;
  if (Object.keys(raw).some((key) => !["version", "runId", "scope", "completedAt", "sources", "results"].includes(key))) {
    return null;
  }
  if (
    raw.version !== CODE_TEST_RUN_RECEIPT_VERSION ||
    typeof raw.runId !== "string" ||
    raw.runId.length === 0 ||
    raw.runId.length > CODE_TEST_RUN_ID_LIMIT ||
    (raw.scope !== "selected" && raw.scope !== "file" && raw.scope !== "all") ||
    typeof raw.completedAt !== "string" ||
    !Number.isFinite(Date.parse(raw.completedAt)) ||
    !Array.isArray(raw.sources) ||
    raw.sources.length === 0 ||
    raw.sources.length > CODE_TEST_WORKSPACE_FILE_LIMIT ||
    !Array.isArray(raw.results) ||
    raw.results.length === 0 ||
    raw.results.length > CODE_TEST_TOTAL_LIMIT
  ) {
    return null;
  }
  const sources = raw.sources.map(normalizeSourceIdentity);
  const results = raw.results.map(normalizeTestResult);
  if (sources.some((source) => source === null) || results.some((result) => result === null)) return null;
  const sourcePages = new Set<string>();
  for (const source of sources as CodeTestSourceIdentity[]) {
    if (sourcePages.has(source.pageId)) return null;
    sourcePages.add(source.pageId);
  }
  const resultIds = new Set<string>();
  for (const result of results as CodeTestResult[]) {
    if (!sourcePages.has(result.pageId) || resultIds.has(result.testId)) return null;
    resultIds.add(result.testId);
  }
  return {
    version: CODE_TEST_RUN_RECEIPT_VERSION,
    runId: raw.runId,
    scope: raw.scope,
    completedAt: raw.completedAt,
    sources: sources as CodeTestSourceIdentity[],
    results: results as CodeTestResult[],
  };
}

export function restoreCodeTestRunReceipt(
  contentJson: unknown,
): PersistedCodeTestRunReceipt | null {
  if (!contentJson || typeof contentJson !== "object" || Array.isArray(contentJson)) return null;
  return normalizePersistedCodeTestRunReceipt(
    (contentJson as Record<string, unknown>)[CODE_TEST_RUN_RECEIPT_KEY],
  );
}

export function codeTestRunReceiptIsCurrent(
  receipt: PersistedCodeTestRunReceipt,
  sources: readonly CodeTestSource[],
): boolean {
  if (receipt.sources.length !== sources.length) return false;
  const byPage = new Map(sources.map((source) => [source.pageId, source] as const));
  return receipt.sources.every((identity) => {
    const source = byPage.get(identity.pageId);
    return Boolean(
      source &&
        source.path === identity.path &&
        source.language === identity.language &&
        codeTestSourceFingerprint(source.source) === identity.sourceFingerprint,
    );
  });
}

async function sha256CodeTestSource(source: string): Promise<string | null> {
  try {
    const subtle = globalThis.crypto?.subtle;
    if (!subtle || typeof TextEncoder === "undefined") return null;
    const bytes = new TextEncoder().encode(source);
    const digest = await subtle.digest("SHA-256", bytes.slice().buffer);
    return Array.from(new Uint8Array(digest), (value) =>
      value.toString(16).padStart(2, "0"),
    ).join("");
  } catch {
    return null;
  }
}

function normalizeCodeTestSourceDigest(
  value: unknown,
): PersistedCodeTestSourceDigest | null {
  if (!value || typeof value !== "object" || Array.isArray(value)) return null;
  const raw = value as Record<string, unknown>;
  if (
    Object.keys(raw).some(
      (key) =>
        key !== "pageId" && key !== "sourceSha256" && key !== "lineCount",
    ) ||
    !validPageId(raw.pageId) ||
    typeof raw.sourceSha256 !== "string" ||
    !SHA256_HEX_PATTERN.test(raw.sourceSha256) ||
    finiteInteger(raw.lineCount, 1, CODE_TEST_RESULT_LINE_LIMIT) === null
  ) {
    return null;
  }
  return {
    pageId: raw.pageId,
    sourceSha256: raw.sourceSha256,
    lineCount: raw.lineCount as number,
  };
}

function normalizeCodeTestFileOutput(
  value: unknown,
  sourceByPage: ReadonlyMap<string, CodeTestSourceIdentity>,
): PersistedCodeTestFileOutput | null {
  if (!value || typeof value !== "object" || Array.isArray(value)) return null;
  const raw = value as Record<string, unknown>;
  if (
    Object.keys(raw).some(
      (key) =>
        key !== "pageId" &&
        key !== "path" &&
        key !== "language" &&
        key !== "logs" &&
        key !== "logStreams",
    ) ||
    !validPageId(raw.pageId) ||
    !validPath(raw.path) ||
    typeof raw.language !== "string" ||
    !isCodeTestLanguage(raw.language) ||
    !Array.isArray(raw.logs) ||
    raw.logs.length === 0 ||
    raw.logs.length > CODE_TEST_RUN_OUTPUT_LINE_LIMIT ||
    !Array.isArray(raw.logStreams) ||
    raw.logStreams.length !== raw.logs.length ||
    raw.logs.some(
      (line) =>
        typeof line !== "string" ||
        line.length > CODE_TEST_RUN_OUTPUT_LINE_CODE_UNIT_LIMIT,
    ) ||
    raw.logStreams.some(
      (stream) => stream !== "stdout" && stream !== "stderr",
    )
  ) {
    return null;
  }
  const source = sourceByPage.get(raw.pageId);
  if (
    !source ||
    source.path !== raw.path ||
    source.language !== raw.language
  ) {
    return null;
  }
  return {
    pageId: raw.pageId,
    path: raw.path,
    language: raw.language,
    logs: raw.logs as string[],
    logStreams: raw.logStreams as ("stdout" | "stderr")[],
  };
}

type MutableCodeTestFileOutput = {
  pageId: string;
  path: string;
  language: CodeTestLanguage;
  logs: string[];
  logStreams: ("stdout" | "stderr")[];
};

function boundCodeTestFileOutputs(
  receipt: PersistedCodeTestRunReceipt,
  values: readonly CodeTestRunFileOutput[],
): readonly PersistedCodeTestFileOutput[] | null {
  const sourceByPage = new Map(
    receipt.sources.map((source) => [source.pageId, source] as const),
  );
  const valueByPage = new Map<string, CodeTestRunFileOutput>();
  for (const value of values) {
    if (
      !validPageId(value.pageId) ||
      !sourceByPage.has(value.pageId) ||
      valueByPage.has(value.pageId) ||
      !value.output ||
      !Array.isArray(value.output.logs)
    ) {
      return null;
    }
    valueByPage.set(value.pageId, value);
  }

  const outputs: MutableCodeTestFileOutput[] = [];
  for (const source of receipt.sources) {
    const value = valueByPage.get(source.pageId);
    if (!value) continue;
    const logs: string[] = [];
    const logStreams: ("stdout" | "stderr")[] = [];
    for (
      let index = 0;
      index < value.output.logs.length &&
      logs.length < CODE_TEST_RUN_OUTPUT_LINE_LIMIT;
      index += 1
    ) {
      const line = value.output.logs[index];
      if (typeof line !== "string") return null;
      logs.push(line.slice(0, CODE_TEST_RUN_OUTPUT_LINE_CODE_UNIT_LIMIT));
      logStreams.push(
        value.output.logStreams?.[index] === "stderr" ? "stderr" : "stdout",
      );
    }
    if (logs.length === 0) continue;
    outputs.push({
      pageId: source.pageId,
      path: source.path,
      language: source.language,
      logs,
      logStreams,
    });
  }

  let truncated = false;
  while (
    outputs.length > 0 &&
    (codeTestRunTimelineValueByteLength(outputs) ?? Infinity) >
      CODE_TEST_RUN_OUTPUT_UTF8_BYTE_LIMIT
  ) {
    const last = outputs[outputs.length - 1]!;
    last.logs.pop();
    last.logStreams.pop();
    truncated = true;
    if (last.logs.length === 0) outputs.pop();
  }
  if (truncated) {
    const markerSource = outputs[outputs.length - 1] ?? receipt.sources[0];
    if (!markerSource) return null;
    const markerTarget =
      outputs[outputs.length - 1] ??
      ({
        pageId: markerSource.pageId,
        path: markerSource.path,
        language: markerSource.language,
        logs: [],
        logStreams: [],
      } satisfies MutableCodeTestFileOutput);
    if (outputs.length === 0) outputs.push(markerTarget);
    markerTarget.logs.push(EXECUTION_OUTPUT_TRUNCATION_MARKER);
    markerTarget.logStreams.push("stdout");
    while (
      markerTarget.logs.length > 1 &&
      (codeTestRunTimelineValueByteLength(outputs) ?? Infinity) >
        CODE_TEST_RUN_OUTPUT_UTF8_BYTE_LIMIT
    ) {
      markerTarget.logs.splice(markerTarget.logs.length - 2, 1);
      markerTarget.logStreams.splice(markerTarget.logStreams.length - 2, 1);
    }
  }
  return outputs;
}

function receiptResultRelationsAreValid(
  receipt: PersistedCodeTestRunReceipt,
  lineCountByPage: ReadonlyMap<string, number>,
): boolean {
  // A file run has one test target, but may consume several local modules.
  // Every consumed source stays fingerprint-bound in the same receipt.
  if (
    receipt.scope === "file" &&
    new Set(receipt.results.map((result) => result.pageId)).size !== 1
  )
    return false;
  const sourceByPage = new Map(
    receipt.sources.map((source) => [source.pageId, source] as const),
  );
  return receipt.results.every((result) => {
    const source = sourceByPage.get(result.pageId);
    const lineCount = lineCountByPage.get(result.pageId);
    if (!source || !lineCount) return false;
    const prefix = `${result.pageId}:`;
    if (!result.testId.startsWith(prefix)) return false;
    const target = /^(\d+):(test[$\u200c\u200d\p{ID_Continue}]*)$/u.exec(
      result.testId.slice(prefix.length),
    );
    const definitionLine = target
      ? finiteInteger(Number(target[1]), 1, lineCount)
      : null;
    return Boolean(
      target &&
        definitionLine !== null &&
        target[2] === result.name &&
        result.testId === `${result.pageId}:${definitionLine}:${result.name}` &&
        result.id === `${result.testId}:${result.status}` &&
        result.path === source.path &&
        result.language === source.language &&
        result.line <= lineCount,
    );
  });
}

function receiptMatchesCurrentDiscovery(
  receipt: PersistedCodeTestRunReceipt,
  sources: readonly CodeTestSource[],
): boolean {
  if (!codeTestRunReceiptIsCurrent(receipt, sources)) return false;
  const discovery = discoverCodeTests(
    receipt.scope === "file"
      ? sources.filter((source) => source.pageId === receipt.results[0]?.pageId)
      : sources,
  );
  if (
    receipt.scope !== "selected" &&
    (discovery.truncated || discovery.files.some((file) => file.truncated))
  ) {
    return false;
  }
  const discovered = new Map(
    discovery.files.flatMap((file) =>
      file.tests.map((test) => [test.id, test] as const),
    ),
  );
  const lineCountByPage = new Map(
    sources.map(
      (source) => [source.pageId, source.source.split("\n").length] as const,
    ),
  );
  if (!receiptResultRelationsAreValid(receipt, lineCountByPage)) return false;
  for (const result of receipt.results) {
    const test = discovered.get(result.testId);
    if (
      !test ||
      test.pageId !== result.pageId ||
      test.path !== result.path ||
      test.language !== result.language ||
      test.name !== result.name
    ) {
      return false;
    }
  }
  if (receipt.scope === "selected") return true;
  return (
    receipt.results.length === discovered.size &&
    receipt.results.every((result) => discovered.has(result.testId))
  );
}

function normalizeCodeTestRunTimelineEntry(
  value: unknown,
): PersistedCodeTestRunTimelineEntry | null {
  if (!value || typeof value !== "object" || Array.isArray(value)) return null;
  const raw = value as Record<string, unknown>;
  if (
    Object.keys(raw).some(
      (key) =>
        !["version", "receipt", "sourceDigests", "fileOutputs"].includes(
          key,
        ),
    ) ||
    raw.version !== CODE_TEST_RUN_TIMELINE_VERSION ||
    !Array.isArray(raw.sourceDigests) ||
    (raw.fileOutputs !== undefined && !Array.isArray(raw.fileOutputs))
  ) {
    return null;
  }
  const receipt = normalizePersistedCodeTestRunReceipt(raw.receipt);
  if (!receipt || raw.sourceDigests.length !== receipt.sources.length) {
    return null;
  }
  const sourceDigests = raw.sourceDigests.map(normalizeCodeTestSourceDigest);
  if (sourceDigests.some((digest) => digest === null)) return null;
  const digestByPage = new Map<string, PersistedCodeTestSourceDigest>();
  for (const digest of sourceDigests as PersistedCodeTestSourceDigest[]) {
    if (digestByPage.has(digest.pageId)) return null;
    digestByPage.set(digest.pageId, digest);
  }
  if (
    !receipt.sources.every((source) => digestByPage.has(source.pageId)) ||
    !receiptResultRelationsAreValid(
      receipt,
      new Map(
        [...digestByPage.values()].map(
          (digest) => [digest.pageId, digest.lineCount] as const,
        ),
      ),
    )
  ) {
    return null;
  }
  const sourceByPage = new Map(
    receipt.sources.map((source) => [source.pageId, source] as const),
  );
  const resultPageIds = new Set(
    receipt.results.map((result) => result.pageId),
  );
  const fileOutputs = (raw.fileOutputs ?? []).map((output) =>
    normalizeCodeTestFileOutput(output, sourceByPage),
  );
  if (
    fileOutputs.some((output) => output === null) ||
    new Set(
      (fileOutputs as PersistedCodeTestFileOutput[]).map(
        (output) => output.pageId,
      ),
    ).size !== fileOutputs.length ||
    (fileOutputs as PersistedCodeTestFileOutput[]).some(
      (output) => !resultPageIds.has(output.pageId),
    ) ||
    (codeTestRunTimelineValueByteLength(raw.fileOutputs ?? []) ?? Infinity) >
      CODE_TEST_RUN_OUTPUT_UTF8_BYTE_LIMIT
  ) {
    return null;
  }
  return {
    version: CODE_TEST_RUN_TIMELINE_VERSION,
    receipt,
    sourceDigests: receipt.sources.map(
      (source) => digestByPage.get(source.pageId)!,
    ),
    ...(fileOutputs.length > 0
      ? {
          fileOutputs: fileOutputs as PersistedCodeTestFileOutput[],
        }
      : {}),
  };
}

function codeTestRunTimelineByteLength(
  entries: readonly PersistedCodeTestRunTimelineEntry[],
): number | null {
  return codeTestRunTimelineValueByteLength({
    version: CODE_TEST_RUN_TIMELINE_VERSION,
    entries,
  } satisfies PersistedCodeTestRunTimeline);
}

function codeTestRunTimelineValueByteLength(value: unknown): number | null {
  try {
    if (typeof TextEncoder === "undefined") return null;
    return new TextEncoder().encode(JSON.stringify(value)).byteLength;
  } catch {
    return null;
  }
}

export async function createPersistedCodeTestRunTimelineEntry(
  receiptValue: PersistedCodeTestRunReceipt,
  sources: readonly CodeTestSource[],
  fileOutputValues: readonly CodeTestRunFileOutput[] = [],
): Promise<PersistedCodeTestRunTimelineEntry | null> {
  const receipt = normalizePersistedCodeTestRunReceipt(receiptValue);
  if (
    !receipt ||
    !receiptMatchesCurrentDiscovery(receipt, sources) ||
    sources.length !== receipt.sources.length
  ) {
    return null;
  }
  const sourceByPage = new Map(sources.map((source) => [source.pageId, source]));
  const sourceDigests: PersistedCodeTestSourceDigest[] = [];
  for (const identity of receipt.sources) {
    const source = sourceByPage.get(identity.pageId);
    if (!source) return null;
    const sourceSha256 = await sha256CodeTestSource(source.source);
    if (!sourceSha256) return null;
    sourceDigests.push({
      pageId: identity.pageId,
      sourceSha256,
      lineCount: source.source.split("\n").length,
    });
  }
  const fileOutputs = boundCodeTestFileOutputs(receipt, fileOutputValues);
  if (!fileOutputs) return null;
  const entry = normalizeCodeTestRunTimelineEntry({
    version: CODE_TEST_RUN_TIMELINE_VERSION,
    receipt,
    sourceDigests,
    ...(fileOutputs.length > 0 ? { fileOutputs } : {}),
  });
  if (
    !entry ||
    (codeTestRunTimelineByteLength([entry]) ?? Infinity) >
      CODE_TEST_RUN_TIMELINE_UTF8_BYTE_LIMIT
  ) {
    return null;
  }
  return entry;
}

export async function codeTestRunTimelineEntryIsCurrent(
  entryValue: PersistedCodeTestRunTimelineEntry,
  sources: readonly CodeTestSource[],
): Promise<boolean> {
  const entry = normalizeCodeTestRunTimelineEntry(entryValue);
  if (!entry || !receiptMatchesCurrentDiscovery(entry.receipt, sources)) {
    return false;
  }
  const sourceByPage = new Map(sources.map((source) => [source.pageId, source]));
  for (const digest of entry.sourceDigests) {
    const source = sourceByPage.get(digest.pageId);
    if (!source || digest.lineCount !== source.source.split("\n").length) {
      return false;
    }
    const currentDigest = await sha256CodeTestSource(source.source);
    if (!currentDigest || currentDigest !== digest.sourceSha256) return false;
  }
  return true;
}

export function appendCodeTestRunTimelineEntry(
  entries: readonly PersistedCodeTestRunTimelineEntry[],
  entryValue: PersistedCodeTestRunTimelineEntry,
): readonly PersistedCodeTestRunTimelineEntry[] {
  const entry = normalizeCodeTestRunTimelineEntry(entryValue);
  if (!entry) return entries;
  const candidates = [
    entry,
    ...entries
      .slice(0, CODE_TEST_RUN_TIMELINE_SCAN_LIMIT)
      .map(normalizeCodeTestRunTimelineEntry)
      .filter(
        (candidate): candidate is PersistedCodeTestRunTimelineEntry =>
          Boolean(candidate && candidate.receipt.runId !== entry.receipt.runId),
      ),
  ].slice(0, CODE_TEST_RUN_TIMELINE_ENTRY_LIMIT);
  while (candidates.length > 0) {
    const byteLength = codeTestRunTimelineByteLength(candidates);
    if (
      byteLength !== null &&
      byteLength <= CODE_TEST_RUN_TIMELINE_UTF8_BYTE_LIMIT
    ) {
      return candidates;
    }
    candidates.pop();
  }
  return [];
}

export function buildPersistedCodeTestRunTimeline(
  entries: readonly PersistedCodeTestRunTimelineEntry[],
): PersistedCodeTestRunTimeline | null {
  const normalized = entries
    .slice(0, CODE_TEST_RUN_TIMELINE_ENTRY_LIMIT)
    .map(normalizeCodeTestRunTimelineEntry)
    .filter(
      (entry): entry is PersistedCodeTestRunTimelineEntry => entry !== null,
    );
  if (
    normalized.length === 0 ||
    (codeTestRunTimelineByteLength(normalized) ?? Infinity) >
      CODE_TEST_RUN_TIMELINE_UTF8_BYTE_LIMIT
  ) {
    return null;
  }
  return { version: CODE_TEST_RUN_TIMELINE_VERSION, entries: normalized };
}

export function restoreCodeTestRunTimeline(
  contentJson: unknown,
): readonly PersistedCodeTestRunTimelineEntry[] {
  if (!contentJson || typeof contentJson !== "object" || Array.isArray(contentJson)) {
    return [];
  }
  const value = (contentJson as Record<string, unknown>)[CODE_TEST_RUN_TIMELINE_KEY];
  if (!value || typeof value !== "object" || Array.isArray(value)) return [];
  const raw = value as Record<string, unknown>;
  if (
    (codeTestRunTimelineValueByteLength(value) ?? Infinity) >
      CODE_TEST_RUN_TIMELINE_UTF8_BYTE_LIMIT ||
    Object.keys(raw).some((key) => key !== "version" && key !== "entries") ||
    raw.version !== CODE_TEST_RUN_TIMELINE_VERSION ||
    !Array.isArray(raw.entries) ||
    raw.entries.length > CODE_TEST_RUN_TIMELINE_ENTRY_LIMIT
  ) {
    return [];
  }
  const normalized = raw.entries.map(normalizeCodeTestRunTimelineEntry);
  if (normalized.some((entry) => entry === null)) return [];
  const entries = normalized as PersistedCodeTestRunTimelineEntry[];
  const runIds = new Set(entries.map((entry) => entry.receipt.runId));
  if (runIds.size !== entries.length) return [];
  return entries;
}

export function stripCodeTestHarnessLogs(
  harness: CodeTestHarness,
  output: RunResult,
): RunResult {
  const logs: string[] = [];
  const streams: ("stdout" | "stderr")[] = [];
  for (let index = 0; index < output.logs.length; index += 1) {
    const log = output.logs[index];
    if (typeof log === "string" && log.startsWith(harness.marker)) continue;
    logs.push(log);
    streams.push(output.logStreams?.[index] === "stderr" ? "stderr" : "stdout");
  }
  return {
    ...output,
    logs,
    ...(output.logStreams ? { logStreams: streams } : {}),
  };
}
