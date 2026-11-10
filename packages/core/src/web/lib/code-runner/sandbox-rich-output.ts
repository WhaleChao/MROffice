/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { RunResult } from "./js-runner";

const SANDBOX_RICH_PREFIX = "__MRNOTE_SANDBOX_TABLE__";
const SANDBOX_RICH_PAYLOAD_CHARS = 45_000;
// Everything the wrapper puts before the person's first line.
const SANDBOX_RICH_HEAD = `
return (function () {
  "use strict";
  var __mrnoteValue = (function () {
`;

/**
 * Wrap authored JavaScript so the already-isolated sandbox can serialize a
 * recognised table through its existing string-only result channel.
 *
 * The wrapper deliberately stays synchronous: /code-sandbox.html currently
 * executes a synchronous Function and does not await returned promises. User
 * code still runs exclusively inside that sandbox; the host only parses JSON.
 */
export function wrapSandboxJavaScriptForRichOutput(
  code: string,
  requestToken: string,
): { code: string; marker: string; lines: SandboxSourceLines } {
  const marker = `${SANDBOX_RICH_PREFIX}:${requestToken}:`;
  return {
    marker,
    lines: {
      before: SANDBOX_RICH_HEAD.split("\n").length - 1,
      count: code.split("\n").length,
    },
    code: `${SANDBOX_RICH_HEAD}${code}
  })();
  if (!Array.isArray(__mrnoteValue) || __mrnoteValue.length === 0) {
    return __mrnoteValue;
  }

  try {
    var __mrnoteRowsSource = __mrnoteValue.slice(0, 1000);
    var __mrnoteColumns = [];
    var __mrnoteAccessKeys = [];
    var __mrnoteRows = [];
    var __mrnoteObjectRows = __mrnoteRowsSource.every(function (row) {
      if (row === null || typeof row !== "object" || Array.isArray(row)) return false;
      var proto = Object.getPrototypeOf(row);
      return proto === Object.prototype || proto === null;
    });
    var __mrnoteArrayRows = !__mrnoteObjectRows && __mrnoteRowsSource.every(Array.isArray);

    if (__mrnoteObjectRows) {
      var __mrnoteSeen = Object.create(null);
      for (var __mrnoteRowIndex = 0; __mrnoteRowIndex < __mrnoteRowsSource.length; __mrnoteRowIndex += 1) {
        var __mrnoteKeys = Object.keys(__mrnoteRowsSource[__mrnoteRowIndex]);
        for (var __mrnoteKeyIndex = 0; __mrnoteKeyIndex < __mrnoteKeys.length; __mrnoteKeyIndex += 1) {
          var __mrnoteKey = __mrnoteKeys[__mrnoteKeyIndex];
          if (__mrnoteSeen[__mrnoteKey]) continue;
          __mrnoteSeen[__mrnoteKey] = true;
          __mrnoteAccessKeys.push(__mrnoteKey);
          __mrnoteColumns.push(__mrnoteKey.slice(0, 1000));
          if (__mrnoteColumns.length >= 100) break;
        }
        if (__mrnoteColumns.length >= 100) break;
      }
    } else if (__mrnoteArrayRows) {
      var __mrnoteColumnCount = 0;
      for (var __mrnoteArrayIndex = 0; __mrnoteArrayIndex < __mrnoteRowsSource.length; __mrnoteArrayIndex += 1) {
        __mrnoteColumnCount = Math.max(__mrnoteColumnCount, Math.min(__mrnoteRowsSource[__mrnoteArrayIndex].length, 100));
      }
      for (var __mrnoteColumnIndex = 0; __mrnoteColumnIndex < __mrnoteColumnCount; __mrnoteColumnIndex += 1) {
        __mrnoteColumns.push(String(__mrnoteColumnIndex));
      }
    }

    if (__mrnoteColumns.length === 0) return __mrnoteValue;

    function __mrnoteCell(value) {
      if (value === null || typeof value === "boolean") return value;
      if (typeof value === "number") return Number.isFinite(value) ? value : String(value);
      if (value === undefined) return null;
      var rendered;
      if (typeof value === "string") rendered = value;
      else {
        try {
          rendered = String(value);
        } catch (_error) {
          rendered = "<unprintable>";
        }
      }
      return rendered.length <= 2000
        ? rendered
        : rendered.slice(0, 1982) + "[output truncated]";
    }

    var __mrnoteUsed = JSON.stringify({
      kind: "table",
      columns: __mrnoteColumns,
      rows: __mrnoteRows,
    }).length;
    for (var __mrnoteSourceIndex = 0; __mrnoteSourceIndex < __mrnoteRowsSource.length; __mrnoteSourceIndex += 1) {
      var __mrnoteSourceRow = __mrnoteRowsSource[__mrnoteSourceIndex];
      var __mrnoteRow = [];
      for (var __mrnoteCellIndex = 0; __mrnoteCellIndex < __mrnoteColumns.length; __mrnoteCellIndex += 1) {
        var __mrnoteRawCell = __mrnoteObjectRows
          ? (Object.prototype.hasOwnProperty.call(__mrnoteSourceRow, __mrnoteAccessKeys[__mrnoteCellIndex])
              ? __mrnoteSourceRow[__mrnoteAccessKeys[__mrnoteCellIndex]]
              : null)
          : __mrnoteSourceRow[__mrnoteCellIndex];
        __mrnoteRow.push(__mrnoteCell(__mrnoteRawCell));
      }
      var __mrnoteRowChars = JSON.stringify(__mrnoteRow).length + 1;
      if (__mrnoteUsed + __mrnoteRowChars > ${SANDBOX_RICH_PAYLOAD_CHARS}) break;
      __mrnoteRows.push(__mrnoteRow);
      __mrnoteUsed += __mrnoteRowChars;
    }

    var __mrnotePayload = JSON.stringify({
      kind: "table",
      columns: __mrnoteColumns,
      rows: __mrnoteRows,
    });
    if (__mrnotePayload.length > ${SANDBOX_RICH_PAYLOAD_CHARS}) return __mrnoteValue;
    return ${JSON.stringify(marker)} + __mrnotePayload;
  } catch (_error) {
    return __mrnoteValue;
  }
})();`,
  };
}

/** Where the person's code sits inside what the sandbox ran. */
export interface SandboxSourceLines {
  before: number;
  count: number;
}

const SANDBOX_ERROR_FRAME =
  /^(\s*at (?:[^\s(]+ \()?)<anonymous>:(\d+):(\d+)(\)?)\s*$/;

/**
 * The sandbox names an error's frames by the line of the code it ran, which
 * is the wrapper above plus the person's code. Put each frame back on the
 * person's line, and drop the wrapper's own frames: a traceback that reads
 * "line 7" for a throw on line 3, then "line 105", sends the caret to the
 * wrong place.
 */
export function placeSandboxErrorFrames(
  result: RunResult,
  lines: SandboxSourceLines | undefined,
): RunResult {
  if (!lines || typeof result.error !== "string") return result;
  let changed = false;
  const placed = result.error
    .split("\n")
    .flatMap((text) => {
      const frame = SANDBOX_ERROR_FRAME.exec(text);
      if (!frame) return [text];
      changed = true;
      const line = Number(frame[2]) - lines.before;
      if (line < 1 || line > lines.count) return [];
      return [`${frame[1]}<anonymous>:${line}:${frame[3]}${frame[4]}`];
    })
    .join("\n");
  return changed ? { ...result, error: placed } : result;
}

/**
 * Decode only the per-request marker produced by the sandbox wrapper. The
 * caller must still pass the result through boundExecutionResult, which applies
 * the canonical row/column/cell limits before UI exposure.
 */
export function decodeSandboxRichOutput(
  result: RunResult,
  marker: string | null,
): RunResult {
  if (
    marker === null ||
    typeof result.result !== "string" ||
    !result.result.startsWith(marker)
  ) {
    return result;
  }
  const encoded = result.result.slice(marker.length);
  if (encoded.length === 0 || encoded.length > SANDBOX_RICH_PAYLOAD_CHARS) {
    return result;
  }
  try {
    const parsed: unknown = JSON.parse(encoded);
    if (!parsed || typeof parsed !== "object") return result;
    const candidate = parsed as Record<string, unknown>;
    if (
      candidate.kind !== "table" ||
      !Array.isArray(candidate.columns) ||
      !Array.isArray(candidate.rows)
    ) {
      return result;
    }
    return {
      ...result,
      result: undefined,
      rich: {
        kind: "table",
        columns: candidate.columns as unknown as string[],
        rows: candidate.rows as unknown[][],
      },
    };
  } catch {
    return result;
  }
}
