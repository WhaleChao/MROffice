/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Best-effort structural formatters for Go and Rust.
 *
 * These are PURE TEXT REINDENTATION passes — NOT the canonical gofmt / rustfmt
 * toolchains. They apply gofmt's most visible rules (tab indentation, brace on
 * same line, blank-line collapsing, trailing-whitespace trim) and the analogous
 * rustfmt rules (4-space indentation, same brace style) as a deterministic string
 * transform. They cannot parse Go/Rust syntax, so edge cases (multi-line strings,
 * complex macro invocations, raw string literals, etc.) may not be perfectly
 * idiomatic, but they will always:
 *   • Return a syntactically equivalent string for well-formatted input (idempotent)
 *   • Return the ORIGINAL text unchanged when an internal error occurs
 *   • Never throw
 *
 * FOLLOW-ON: Swapping in the real gofmt / rustfmt requires:
 *   1. A multi-MB WASM binary compiled from the Go / Rust toolchain
 *   2. A bundler plugin or CDN fetch to make it available in the browser
 *   3. A worker wrapper (matching the pattern in python-runner.ts) for the WASM call
 *   That is an ops/build decision outside the scope of this formatter module.
 */

// ---------------------------------------------------------------------------
// Shared helpers
// ---------------------------------------------------------------------------

/**
 * Which lines sit INSIDE a multi-line string literal.
 *
 * These formatters reindent text, and a Go raw string or a Rust raw/multi-line
 * string carries its own indentation as DATA: an embedded SQL query, a template,
 * a test fixture. Reindenting those lines changes what the program does, and
 * counting the braces inside them mis-nests everything after — a `{` in a SQL
 * string used to push the function's own closing brace a level deeper.
 *
 * So the lines a string owns are copied through untouched. Two masks, because
 * the opening line is not the same case as the interior:
 *   - `inside` — a literal was already open at the line's START, so the whole
 *     line is content: no reindent, no trim, no blank-line collapsing.
 *   - `keepTail` — a literal is open at the line's END. That line is still code
 *     up to the delimiter and gets reindented, but everything after the
 *     delimiter belongs to the string, so its trailing spaces must survive.
 */
interface StringMask {
  inside: boolean[];
  keepTail: boolean[];
}

function protectedLines(lines: string[], language: "go" | "rust"): StringMask {
  const inside: boolean[] = [];
  const keepTail: boolean[] = [];
  // The delimiter that would close the literal currently open, or null.
  let closing: string | null = null;

  for (const line of lines) {
    inside.push(closing !== null);
    let i = 0;
    while (i < line.length) {
      if (closing !== null) {
        const end = line.indexOf(closing, i);
        if (end < 0) break; // still open at end of line
        i = end + closing.length;
        closing = null;
        continue;
      }
      const ch = line[i];
      // A line comment ends the interesting part of the line.
      if (ch === "/" && line[i + 1] === "/") break;
      if (language === "go" && ch === "`") {
        closing = "`";
        i += 1;
        continue;
      }
      if (language === "rust" && ch === "r") {
        // r"…", r#"…"#, r##"…"## — the terminator carries the same hash count.
        let hashes = 0;
        while (line[i + 1 + hashes] === "#") hashes += 1;
        if (line[i + 1 + hashes] === '"') {
          closing = `"${"#".repeat(hashes)}`;
          i += 2 + hashes;
          continue;
        }
      }
      if (ch === '"') {
        // Rust allows a plain string to span lines; Go does not, so an unclosed
        // quote there is a syntax error rather than a continuation, and leaving
        // it unprotected keeps this pass from swallowing the rest of the file.
        let j = i + 1;
        let closed = false;
        while (j < line.length) {
          if (line[j] === "\\") {
            j += 2;
            continue;
          }
          if (line[j] === '"') {
            closed = true;
            break;
          }
          j += 1;
        }
        if (closed) {
          i = j + 1;
          continue;
        }
        if (language === "rust") closing = '"';
        break;
      }
      if (ch === "'") {
        // Go rune / Rust char or lifetime: never spans a line either way.
        break;
      }
      i += 1;
    }
    keepTail.push(closing !== null);
  }
  return { inside, keepTail };
}

/** Trim trailing whitespace from every line (both formatters require this). */
function trimTrailingWhitespace(
  lines: string[],
  mask: StringMask,
): string[] {
  // Trailing spaces inside a raw string are CONTENT — a fixture may depend on
  // them — so those lines are handed back exactly as they came. That includes
  // the line the literal OPENS on, whose tail is already inside the string.
  return lines.map((l, index) =>
    mask.inside[index] || mask.keepTail[index] ? l : l.trimEnd(),
  );
}

/**
 * Collapse runs of 3+ consecutive blank lines to exactly 2 blank lines (the
 * maximum gofmt / rustfmt permits at file level), and strip leading/trailing
 * blank lines from the whole file.
 */
function collapseBlankLines(
  lines: string[],
  maxConsecutive: number,
  mask: StringMask,
): string[] {
  const out: string[] = [];
  let blankRun = 0;
  for (const [index, line] of lines.entries()) {
    // A blank line inside a raw string is part of the string.
    if (mask.inside[index]) {
      blankRun = 0;
      out.push(line);
      continue;
    }
    if (line === "") {
      blankRun++;
      if (blankRun <= maxConsecutive) out.push("");
    } else {
      blankRun = 0;
      out.push(line);
    }
  }
  // Strip leading / trailing blank lines.
  while (out.length > 0 && out[0] === "") out.shift();
  while (out.length > 0 && out[out.length - 1] === "") out.pop();
  return out;
}

/**
 * Strip an arbitrary indent prefix from every line and re-indent using
 * `indentUnit`, maintaining relative nesting depth via brace / bracket
 * counting.
 *
 * Rules applied:
 *   - Opening braces/brackets at end of non-empty line → stay on same line
 *   - Lines that start with `}` or `]` → dedent before printing
 *   - Lines that end with `{` or `[` → increase indent for next line
 *
 * This is not a parser — it tracks depth purely by counting delimiters.
 */
function reindent(
  lines: string[],
  indentUnit: string,
  mask: StringMask,
): string[] {
  const out: string[] = [];
  let depth = 0;

  for (const [index, rawLine] of lines.entries()) {
    // Inside a multi-line string: emit the line exactly, and let none of its
    // characters change the nesting. A `{` in an embedded SQL query used to
    // push everything after it a level deeper, including the function's own
    // closing brace.
    if (mask.inside[index]) {
      out.push(rawLine);
      continue;
    }
    // On a line that OPENS a multi-line literal, everything after the delimiter
    // is already string content, so only the leading indentation may go.
    const line = mask.keepTail[index]
      ? rawLine.replace(/^\s+/, "")
      : rawLine.trim();

    // Count leading close-delimiters to dedent this line BEFORE printing.
    // We count the run of `}` / `)` / `]` at the START (possibly mixed with spaces).
    let closingLeadCount = 0;
    for (let i = 0; i < line.length; i++) {
      const ch = line[i];
      if (ch === "}" || ch === ")" || ch === "]") {
        closingLeadCount++;
      } else {
        break;
      }
    }
    // Dedent by the number of leading closers, floored at 0.
    depth = Math.max(0, depth - closingLeadCount);

    if (line === "") {
      out.push("");
    } else {
      out.push(indentUnit.repeat(depth) + line);
    }

    // Count net openers on this line to decide the indent of the NEXT line.
    // We skip characters inside string literals (simple heuristic: ignore
    // characters between matching quotes; handles the common case).
    let netOpen = 0;
    let inStr = false;
    let strChar = "";
    for (let i = 0; i < line.length; i++) {
      const ch = line[i];
      if (inStr) {
        if (ch === "\\" ) {
          i++; // skip escaped character
        } else if (ch === strChar) {
          inStr = false;
        }
        continue;
      }
      if (ch === '"' || ch === "'" || ch === "`") {
        inStr = true;
        strChar = ch;
        continue;
      }
      // Line comment: stop counting for the rest of this line.
      if (ch === "/" && line[i + 1] === "/") break;
      if (ch === "{" || ch === "(") netOpen++;
      else if (ch === "}" || ch === ")") netOpen--;
      // `[` and `]` — Go array/slice literals, Rust generics: count them too
      // but only when not already consumed as closingLeadCount (already done).
      else if (ch === "[") netOpen++;
      else if (ch === "]") netOpen--;
    }
    // Adjust for closingLeadCount already subtracted above (those `}` were
    // already counted in the loop above for netOpen — re-add them so they
    // don't double-subtract the next line's depth).
    netOpen += closingLeadCount;
    depth = Math.max(0, depth + netOpen);
  }

  return out;
}

// ---------------------------------------------------------------------------
// Go formatter (gofmt parity: tab indentation)
// ---------------------------------------------------------------------------

/**
 * Best-effort Go formatter.
 *
 * Applies gofmt's core visual rules:
 *   • Hard-tab indentation (one `\t` per nesting level)
 *   • Opening braces on the same line as the preceding statement (enforced by
 *     the reindent pass, which does not move braces between lines)
 *   • Trailing whitespace trimmed
 *   • Blank-line runs collapsed (≤ 2 consecutive blank lines)
 *
 * NOT applied (would require parsing): operator spacing, import grouping,
 * comment formatting, struct tag alignment. Those are handled by real gofmt.
 *
 * TOTAL: returns `original` unchanged on any internal error.
 */
export function formatGo(text: string): string {
  if (!text) return text;
  try {
    const lines = text.split("\n");
    // Computed once, from the ORIGINAL lines, and reused by all three passes:
    // every one of them would otherwise rewrite string content.
    const mask = protectedLines(lines, "go");
    const trimmed = trimTrailingWhitespace(lines, mask);
    const reindented = reindent(trimmed, "\t", mask);
    const collapsed = collapseBlankLines(reindented, 2, mask);
    const result = collapsed.join("\n") + "\n";
    return result;
  } catch {
    return text;
  }
}

// ---------------------------------------------------------------------------
// Rust formatter (rustfmt parity: 4-space indentation)
// ---------------------------------------------------------------------------

/**
 * Best-effort Rust formatter.
 *
 * Applies rustfmt's core visual rules:
 *   • 4-space indentation (no tabs)
 *   • Opening braces on the same line
 *   • Trailing whitespace trimmed
 *   • Blank-line runs collapsed (≤ 2 consecutive blank lines)
 *
 * NOT applied (would require parsing): line-length wrapping, import
 * (`use`) statement sorting, match-arm alignment, attribute formatting,
 * macro call formatting. Those are handled by real rustfmt.
 *
 * TOTAL: returns `original` unchanged on any internal error.
 */
export function formatRust(text: string): string {
  if (!text) return text;
  try {
    const lines = text.split("\n");
    // Computed once, from the ORIGINAL lines, and reused by all three passes:
    // every one of them would otherwise rewrite string content.
    const mask = protectedLines(lines, "rust");
    const trimmed = trimTrailingWhitespace(lines, mask);
    const reindented = reindent(trimmed, "    ", mask);
    const collapsed = collapseBlankLines(reindented, 2, mask);
    const result = collapsed.join("\n") + "\n";
    return result;
  } catch {
    return text;
  }
}
