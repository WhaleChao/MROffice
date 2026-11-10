/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Pure editor key handling for the textarea-based code editor: auto-indent on
 * Enter, multi-line indent / dedent with Tab / Shift+Tab, and auto-closing
 * brackets & quotes (with selection-wrap and skip-over). Framework-free so the
 * behaviour is unit-tested without a DOM; CodeEditor applies the returned value
 * and selection.
 */

export interface EditorKeyInput {
  key: string;
  shiftKey: boolean;
  value: string;
  selectionStart: number;
  selectionEnd: number;
}

export interface EditorEdit {
  value: string;
  selectionStart: number;
  selectionEnd: number;
}

export interface ImeKeyInput {
  isComposing?: boolean;
  keyCode?: number;
}

/**
 * Browsers do not report IME candidate keystrokes consistently. Modern engines
 * expose `isComposing`; Safari and some Chromium input methods can instead send
 * the legacy keyCode 229 while composition is active.
 */
export function isImeCompositionKey(input: ImeKeyInput): boolean {
  return input.isComposing === true || input.keyCode === 229;
}

const INDENT = "  ";
const PAIRS: Record<string, string> = {
  "(": ")",
  "[": "]",
  "{": "}",
  '"': '"',
  "'": "'",
  "`": "`",
};
const CLOSERS = new Set(Object.values(PAIRS));

function lineStartIndex(value: string, index: number): number {
  const nl = value.lastIndexOf("\n", index - 1);
  return nl === -1 ? 0 : nl + 1;
}

function leadingWhitespace(line: string): string {
  const match = /^[ \t]*/.exec(line);
  return match ? match[0] : "";
}

/** Auto-indent: a new line keeps the current indent, and deepens it inside an
 * opener (`{`/`(`/`[`) or after a Python `:`. An empty bracket pair splits. */
function handleEnter(input: EditorKeyInput): EditorEdit {
  const { value, selectionStart, selectionEnd } = input;
  const ls = lineStartIndex(value, selectionStart);
  const currentLine = value.slice(ls, selectionStart);
  const indent = leadingWhitespace(currentLine);
  const before = value[selectionStart - 1];
  const after = value[selectionEnd];
  const opensBlock = before === "{" || before === "(" || before === "[" || before === ":";

  if (
    (before === "{" && after === "}") ||
    (before === "(" && after === ")") ||
    (before === "[" && after === "]")
  ) {
    // Split an empty pair onto three lines with the cursor on the middle one.
    const insertion = `\n${indent}${INDENT}\n${indent}`;
    const caret = selectionStart + 1 + indent.length + INDENT.length;
    return {
      value: value.slice(0, selectionStart) + insertion + value.slice(selectionEnd),
      selectionStart: caret,
      selectionEnd: caret,
    };
  }

  const nextIndent = opensBlock ? indent + INDENT : indent;
  const insertion = `\n${nextIndent}`;
  const caret = selectionStart + insertion.length;
  return {
    value: value.slice(0, selectionStart) + insertion + value.slice(selectionEnd),
    selectionStart: caret,
    selectionEnd: caret,
  };
}

function handleTab(input: EditorKeyInput): EditorEdit {
  const { shiftKey, value, selectionStart, selectionEnd } = input;
  const multiLine = value.slice(selectionStart, selectionEnd).includes("\n");

  if (!multiLine && !shiftKey) {
    // Insert a single indent at the caret.
    const caret = selectionStart + INDENT.length;
    return {
      value: value.slice(0, selectionStart) + INDENT + value.slice(selectionEnd),
      selectionStart: caret,
      selectionEnd: caret,
    };
  }

  // Operate on every line the selection touches.
  const blockStart = lineStartIndex(value, selectionStart);
  const blockEnd = selectionEnd;
  const head = value.slice(0, blockStart);
  const block = value.slice(blockStart, blockEnd);
  const tail = value.slice(blockEnd);
  const lines = block.split("\n");
  let startDelta = 0;
  let totalDelta = 0;
  const newLines = lines.map((line, i) => {
    if (shiftKey) {
      const removed = /^( {1,2}|\t)/.exec(line);
      const cut = removed ? removed[0].length : 0;
      if (i === 0) startDelta = -cut;
      totalDelta -= cut;
      return line.slice(cut);
    }
    if (i === 0) startDelta = INDENT.length;
    totalDelta += INDENT.length;
    return INDENT + line;
  });
  const newBlock = newLines.join("\n");
  return {
    value: head + newBlock + tail,
    selectionStart: Math.max(blockStart, selectionStart + startDelta),
    selectionEnd: Math.max(blockStart, selectionEnd + totalDelta),
  };
}

function shouldAutoClose(open: string, after: string | undefined): boolean {
  // Always pair brackets; only pair quotes when not butting against a word
  // char (so apostrophes inside identifiers/text are left alone).
  if (open === '"' || open === "'" || open === "`") {
    if (after && /[\w]/.test(after)) return false;
  }
  return after === undefined || after === "" || /[\s)\]}>,.;:]/.test(after);
}

function handlePair(input: EditorKeyInput): EditorEdit | null {
  const { key, value, selectionStart, selectionEnd } = input;
  const close = PAIRS[key];
  if (selectionStart !== selectionEnd) {
    // Wrap the selection.
    const selected = value.slice(selectionStart, selectionEnd);
    return {
      value:
        value.slice(0, selectionStart) +
        key +
        selected +
        close +
        value.slice(selectionEnd),
      selectionStart: selectionStart + 1,
      selectionEnd: selectionEnd + 1,
    };
  }
  if (!shouldAutoClose(key, value[selectionStart])) return null;
  const caret = selectionStart + 1;
  return {
    value: value.slice(0, selectionStart) + key + close + value.slice(selectionStart),
    selectionStart: caret,
    selectionEnd: caret,
  };
}

/**
 * Skip over an auto-inserted closer/quote when typing the same char right in
 * front of it (so `(|)` + `)` → `()|`, and `'x|'` + `'` → `'x'|`).
 */
function handleCloserSkip(input: EditorKeyInput): EditorEdit | null {
  const { key, value, selectionStart, selectionEnd } = input;
  if (selectionStart !== selectionEnd) return null;
  if (!CLOSERS.has(key)) return null;
  if (value[selectionStart] !== key) return null;
  const caret = selectionStart + 1;
  return { value, selectionStart: caret, selectionEnd: caret };
}

/** Backspace inside an empty pair removes both characters. */
function handleBackspace(input: EditorKeyInput): EditorEdit | null {
  const { value, selectionStart, selectionEnd } = input;
  if (selectionStart !== selectionEnd || selectionStart === 0) return null;
  const before = value[selectionStart - 1];
  const after = value[selectionStart];
  if (before in PAIRS && PAIRS[before] === after) {
    const caret = selectionStart - 1;
    return {
      value: value.slice(0, selectionStart - 1) + value.slice(selectionStart + 1),
      selectionStart: caret,
      selectionEnd: caret,
    };
  }
  return null;
}

/**
 * Compute the edit for a key, or null if the key should fall through to the
 * browser's default handling. `key` is the KeyboardEvent.key.
 */
export function handleEditorKey(input: EditorKeyInput): EditorEdit | null {
  if (input.key === "Enter") return handleEnter(input);
  if (input.key === "Tab") return handleTab(input);
  if (input.key === "Backspace") return handleBackspace(input);
  // Skip-over is checked before auto-pair because quotes are both opener and
  // closer: typing ' right before an existing ' must step over, not insert.
  if (input.key.length === 1 && CLOSERS.has(input.key)) {
    const skip = handleCloserSkip(input);
    if (skip) return skip;
  }
  if (input.key.length === 1 && input.key in PAIRS) return handlePair(input);
  return null;
}
