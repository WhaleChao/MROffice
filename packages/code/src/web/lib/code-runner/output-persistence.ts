/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Pure helpers for PERSISTING a multi-buffer host cell's last run output (Jupyter
 * parity: outputs survive reload + .ipynb export).
 *
 * A run produces an ephemeral `RunResult` (logs/logStreams/result/error/rich/…).
 * To survive a reload it is distilled into a small, serializable
 * `StoredCellOutput` that:
 *   - is stored on the cell (`multi-buffer hostCellSchema.outputs`, OPTIONAL/additive);
 *   - round-trips to/from nbformat outputs (stream / execute_result /
 *     display_data / error) for .ipynb export and import.
 *
 * Everything here is synchronous and side-effect free so the snapshot<->nbformat
 * mapping can be unit-tested without a runtime or the React tree.
 *
 * Size discipline: a single giant log line or multi-megabyte figure must not
 * bloat the saved document. Text is truncated to explicit caps; an oversized
 * encoded image is dropped rather than sliced into corrupt base64. Re-running
 * the cell restores the full live output.
 */

import type { LogStream } from "@mrnote/core/web/lib/code-runner/cell-output-format.ts";
import type { RunResult, RichOutput } from "@mrnote/core/web/lib/code-runner/js-runner.ts";
import { boundExecutionResult } from "@mrnote/core/web/lib/code-runner/execution-output-budget.ts";
import {
  ALLOWED_PERSISTED_RASTER_IMAGE_MIMES,
  MAX_PERSISTED_CELL_OUTPUT_JSON_CHARS,
} from "./output-persistence-budgets";

export { MAX_PERSISTED_CELL_OUTPUT_JSON_CHARS } from "./output-persistence-budgets";

// Raster image MIME types a persisted/imported rich `image` may carry. Kept in
// lockstep with the schema whitelist (native-page-schemas.ts). SVG is excluded
// on purpose: it is markup (a script-injection vector), not opaque raster bytes.
const ALLOWED_RICH_IMAGE_MIME_SET = new Set<string>(
  ALLOWED_PERSISTED_RASTER_IMAGE_MIMES,
);

/** True when `mime` is one of the whitelisted raster image types. */
export function isAllowedRichImageMime(mime: unknown): mime is string {
  return typeof mime === "string" && ALLOWED_RICH_IMAGE_MIME_SET.has(mime);
}

// --- Size caps (anti-bloat). Generous enough for normal output; a runaway loop
// or a huge figure is bounded rather than persisted verbatim. ---
/** Max characters kept for a single log line. */
export const MAX_LOG_LINE_CHARS = 20_000;
/** Max number of log lines kept (older lines past this are dropped). */
export const MAX_LOG_LINES = 2_000;
/** Max characters kept for the text `result` / `error`. */
export const MAX_TEXT_CHARS = 50_000;
/** Max characters kept for a stored HTML payload. */
export const MAX_HTML_CHARS = 200_000;
/** Max base64 characters kept per image (≈ payload bytes * 4/3). */
export const MAX_IMAGE_CHARS = 1_000_000;
/** Max images kept on a single cell. */
export const MAX_IMAGES = 20;
/** Max rows / columns kept for a stored rich table. */
export const MAX_TABLE_ROWS = 1_000;
export const MAX_TABLE_COLS = 100;

/** Aggregate category budgets keep the complete snapshot below 2M. */
const MAX_PERSISTED_LOG_CHARS = 300_000;
const MAX_PERSISTED_IMAGE_CHARS = 1_000_000;
const MAX_PERSISTED_RICH_TABLE_JSON_CHARS = 500_000;
const MAX_PERSISTED_TABLE_CELL_CHARS = 2_000;
const MAX_PERSISTED_TABLE_COLUMN_CHARS = 1_000;

function truncate(text: string, max: number): string {
  return text.length > max ? text.slice(0, max) : text;
}

const INVALID_OUTPUT_FIELD = Symbol("invalid-output-field");

function ownOutputField(
  value: object,
  key: string,
): unknown | typeof INVALID_OUTPUT_FIELD {
  try {
    const descriptor = Object.getOwnPropertyDescriptor(value, key);
    if (!descriptor) return undefined;
    return "value" in descriptor ? descriptor.value : INVALID_OUTPUT_FIELD;
  } catch {
    return INVALID_OUTPUT_FIELD;
  }
}

/**
 * A serializable snapshot of a cell's last run, stored on the cell so the output
 * renders again on reload without re-executing. Mirrors the renderable subset of
 * `RunResult`; transient run-state flags (`timedOut`/`interrupted`) are NOT kept
 * — only the visible output (logs/result/error/rich) is persisted.
 */
export interface StoredCellOutput {
  /** Captured log lines (truncated to caps). */
  readonly logs: string[];
  /**
   * A strictly validated, ordered nbformat output array imported from Jupyter.
   * It is authoritative when present: the editor derives a safe render view at
   * runtime and re-emits this snapshot verbatim on export. Keeping the ordered
   * representation avoids collapsing interleaved streams, display bundles and
   * error fields into the native aggregate model and then pretending that the
   * original notebook can be reconstructed.
   */
  readonly notebookOutputs?: readonly NbformatOutput[];
  /** Per-line stream tags aligned to `logs`; absent ⇒ all stdout (legacy). */
  readonly logStreams?: LogStream[];
  /** The final expression's text value (Jupyter execute_result). */
  readonly result?: string;
  /** A raised error / traceback text. */
  readonly error?: string;
  /** Base64 raster figures (e.g. matplotlib), no data: prefix. */
  readonly images?: string[];
  /**
   * Per-image MIME types aligned to `images`; absent ⇒ all `image/png` (the
   * runtime + legacy snapshots). Lets a non-PNG `display_data` image imported
   * from a .ipynb (jpeg/gif/webp/bmp) keep its real type on render + export
   * instead of being mislabeled as PNG.
   */
  readonly imageMimes?: string[];
  /** Rich HTML repr (e.g. pandas DataFrame), rendered sandboxed. */
  readonly html?: string;
  /** A structured rich output (table / image / html) for the final value. */
  readonly rich?: RichOutput;
}

/** Clamp a rich output to the stored caps (table rows/cols, image/html size). */
function persistedTableCell(value: unknown): unknown {
  if (value === null || typeof value === "boolean") return value;
  if (typeof value === "number")
    return Number.isFinite(value) ? value : String(value);
  if (typeof value === "string") {
    return truncate(value, MAX_PERSISTED_TABLE_CELL_CHARS);
  }
  if (value === undefined) return "";
  if (
    typeof value === "bigint" ||
    typeof value === "symbol" ||
    typeof value === "function"
  ) {
    return truncate(String(value), MAX_PERSISTED_TABLE_CELL_CHARS);
  }
  return truncate(safeJson(value), MAX_PERSISTED_TABLE_CELL_CHARS);
}

function clampRich(rich: RichOutput): RichOutput | undefined {
  if (rich.kind === "table") {
    const columns = rich.columns
      .slice(0, MAX_TABLE_COLS)
      .map((column) =>
        truncate(String(column), MAX_PERSISTED_TABLE_COLUMN_CHARS),
      );
    const rows: unknown[][] = [];
    let used = JSON.stringify({ kind: "table", columns, rows: [] }).length;
    for (const rawRow of rich.rows.slice(0, MAX_TABLE_ROWS)) {
      const row = rawRow
        .slice(0, MAX_TABLE_COLS)
        .map((cell) => persistedTableCell(cell));
      const rowChars = JSON.stringify(row).length + 1;
      if (used + rowChars > MAX_PERSISTED_RICH_TABLE_JSON_CHARS) break;
      rows.push(row);
      used += rowChars;
    }
    return { kind: "table", columns, rows };
  }
  if (rich.kind === "image") {
    if (
      !isAllowedRichImageMime(rich.mime) ||
      typeof rich.data !== "string" ||
      rich.data.length > MAX_IMAGE_CHARS
    ) {
      return undefined;
    }
    return {
      kind: "image",
      mime: rich.mime,
      data: rich.data,
    };
  }
  return { kind: "html", html: truncate(rich.html, MAX_HTML_CHARS) };
}

type MutableStoredOutput = {
  logs: string[];
  notebookOutputs?: NbformatOutput[];
  logStreams?: LogStream[];
  result?: string;
  error?: string;
  images?: string[];
  imageMimes?: string[];
  html?: string;
  rich?: RichOutput;
};

function finalizeStoredOutput(
  raw: MutableStoredOutput,
): StoredCellOutput | null {
  const logs: string[] = [];
  const streams: LogStream[] = [];
  let logChars = 0;
  for (
    let index = 0;
    index < raw.logs.length && logs.length < MAX_LOG_LINES;
    index += 1
  ) {
    const line = truncate(String(raw.logs[index]), MAX_LOG_LINE_CHARS);
    const remaining = MAX_PERSISTED_LOG_CHARS - logChars;
    if (remaining <= 0) break;
    const kept = truncate(line, remaining);
    logs.push(kept);
    logChars += kept.length;
    if (raw.logStreams)
      streams.push(raw.logStreams[index] === "stderr" ? "stderr" : "stdout");
  }

  const next: MutableStoredOutput = { logs };
  if (raw.logStreams && streams.length > 0) next.logStreams = streams;
  if (raw.result !== undefined)
    next.result = truncate(raw.result, MAX_TEXT_CHARS);
  if (raw.error !== undefined) next.error = truncate(raw.error, MAX_TEXT_CHARS);
  if (raw.html !== undefined) next.html = truncate(raw.html, MAX_HTML_CHARS);

  if (raw.images && raw.images.length > 0) {
    const images: string[] = [];
    const mimes: string[] = [];
    let imageChars = 0;
    for (
      let index = 0;
      index < raw.images.length && images.length < MAX_IMAGES;
      index += 1
    ) {
      const image = raw.images[index];
      if (
        typeof image !== "string" ||
        image.length > MAX_IMAGE_CHARS ||
        imageChars + image.length > MAX_PERSISTED_IMAGE_CHARS
      ) {
        continue;
      }
      images.push(image);
      imageChars += image.length;
      const mime = raw.imageMimes?.[index];
      mimes.push(isAllowedRichImageMime(mime) ? mime : "image/png");
    }
    if (images.length > 0) {
      next.images = images;
      if (mimes.some((mime) => mime !== "image/png")) next.imageMimes = mimes;
    }
  }

  if (raw.rich) {
    const rich = clampRich(raw.rich);
    if (rich) next.rich = rich;
  }

  // Category budgets above are chosen to fit, but keep one final assertion at
  // the serialization boundary so future fields cannot silently invalidate the
  // transport contract. Optional representations are discarded in descending
  // recomputation cost; authored source is never involved here.
  if (JSON.stringify(next).length > MAX_PERSISTED_CELL_OUTPUT_JSON_CHARS) {
    delete next.images;
    delete next.imageMimes;
  }
  if (JSON.stringify(next).length > MAX_PERSISTED_CELL_OUTPUT_JSON_CHARS) {
    delete next.html;
  }
  if (JSON.stringify(next).length > MAX_PERSISTED_CELL_OUTPUT_JSON_CHARS) {
    delete next.rich;
  }

  const hasContent =
    next.logs.length > 0 ||
    next.result !== undefined ||
    next.error !== undefined ||
    (next.images?.length ?? 0) > 0 ||
    next.html !== undefined ||
    next.rich !== undefined;
  return hasContent ? next : null;
}

/**
 * Distill a live `RunResult` into a `StoredCellOutput`, or `null` when there is
 * nothing worth persisting (no logs/result/error/rich/images/html — e.g. a
 * cleared or still-loading placeholder). Pure; applies the size caps.
 */
export function runResultToStoredOutputs(
  result: RunResult | null | undefined,
): StoredCellOutput | null {
  if (!result) return null;
  const notebookOutputs = ownOutputField(result, "notebookOutputs");
  if (notebookOutputs === INVALID_OUTPUT_FIELD) return null;
  if (notebookOutputs !== undefined) {
    const snapshot = normalizeNbformatOutputsSnapshot(notebookOutputs);
    return snapshot && snapshot.length > 0
      ? { logs: [], notebookOutputs: snapshot }
      : null;
  }
  const sourceLogStreams = ownOutputField(result, "logStreams");
  if (sourceLogStreams === INVALID_OUTPUT_FIELD) return null;
  const persistLogStreams = (ownArrayLength(sourceLogStreams) ?? 0) > 0;
  const bounded = boundExecutionResult(result);
  const logs = bounded.logs
    .slice(0, MAX_LOG_LINES)
    .map((line) => truncate(String(line), MAX_LOG_LINE_CHARS));
  const stored: MutableStoredOutput = { logs };
  if (
    persistLogStreams &&
    bounded.logStreams &&
    bounded.logStreams.length > 0
  ) {
    stored.logStreams = bounded.logStreams.slice(
      0,
      MAX_LOG_LINES,
    ) as LogStream[];
  }
  if (typeof bounded.result === "string") {
    stored.result = truncate(bounded.result, MAX_TEXT_CHARS);
  }
  if (typeof bounded.error === "string") {
    stored.error = truncate(bounded.error, MAX_TEXT_CHARS);
  }
  if (bounded.images && bounded.images.length > 0) {
    stored.images = bounded.images.slice(0, MAX_IMAGES).map(String);
    if (
      bounded.imageMimes &&
      bounded.imageMimes.length === stored.images.length
    ) {
      stored.imageMimes = bounded.imageMimes.map((mime) =>
        isAllowedRichImageMime(mime) ? mime : "image/png",
      );
    }
  }
  if (typeof bounded.html === "string") {
    stored.html = truncate(bounded.html, MAX_HTML_CHARS);
  }
  if (bounded.rich) stored.rich = bounded.rich;
  return finalizeStoredOutput(stored);
}

/**
 * Coerce a persisted `outputs` value (from a saved doc / schema passthrough)
 * back into a clean `StoredCellOutput`, dropping malformed parts. Returns `null`
 * for anything unusable so a legacy cell with no outputs renders as before.
 * Pure + defensive (never throws).
 */
export function normalizeStoredOutputs(raw: unknown): StoredCellOutput | null {
  const detached = detachJsonTree(raw);
  if (
    detached === INVALID_JSON_TREE ||
    !detached ||
    typeof detached !== "object" ||
    Array.isArray(detached)
  ) {
    return null;
  }
  const record = detached as Record<string, unknown>;
  if (record.notebookOutputs !== undefined) {
    const snapshot = normalizeNbformatOutputsSnapshot(record.notebookOutputs);
    const hasNativeFields = Object.keys(record).some(
      (key) => key !== "logs" && key !== "notebookOutputs",
    );
    if (
      snapshot === null ||
      snapshot.length === 0 ||
      hasNativeFields ||
      !Array.isArray(record.logs) ||
      record.logs.length !== 0
    ) {
      return null;
    }
    return { logs: [], notebookOutputs: snapshot };
  }
  // Filter `logs` and `logStreams` TOGETHER, paired by index, so the surviving
  // stream tags stay aligned to the surviving log lines. Dropping a malformed
  // log entry without dropping its tag would shift every later tag onto the
  // wrong line (e.g. stderr coloring jumping to a stdout line). Tags are
  // optional/additive: a missing or shorter `logStreams` defaults to stdout.
  const rawLogs = Array.isArray(record.logs) ? record.logs : [];
  const rawStreams = Array.isArray(record.logStreams)
    ? record.logStreams
    : undefined;
  const logs: string[] = [];
  const logStreams: LogStream[] = [];
  for (let i = 0; i < rawLogs.length; i += 1) {
    if (logs.length >= MAX_LOG_LINES) break;
    const line = rawLogs[i];
    if (typeof line !== "string") continue;
    logs.push(truncate(line, MAX_LOG_LINE_CHARS));
    const tag = rawStreams?.[i];
    logStreams.push(tag === "stderr" ? "stderr" : "stdout");
  }
  const stored: MutableStoredOutput = { logs };
  // Preserve aligned stream tags when the source carried any (matches the
  // original "keep if present" semantic, now index-aligned to the kept logs).
  if (rawStreams !== undefined && logStreams.length > 0) {
    stored.logStreams = logStreams;
  }
  if (typeof record.result === "string") {
    stored.result = truncate(record.result, MAX_TEXT_CHARS);
  }
  if (typeof record.error === "string") {
    stored.error = truncate(record.error, MAX_TEXT_CHARS);
  }
  if (Array.isArray(record.images)) {
    const rawImages = record.images;
    const images = rawImages
      .filter((b64): b64 is string => typeof b64 === "string")
      .slice(0, MAX_IMAGES);
    if (images.length > 0) {
      stored.images = images;
      // Carry the aligned per-image MIMEs ONLY when nothing was filtered or
      // capped away (so index alignment is guaranteed); otherwise default to PNG
      // rather than risk mislabeling a different image.
      if (
        Array.isArray(record.imageMimes) &&
        rawImages.length === images.length &&
        record.imageMimes.length === images.length
      ) {
        const mimes = record.imageMimes.map((m) =>
          // Validate against the raster whitelist (a hand-crafted record could
          // smuggle image/svg+xml or a script-y type that bypasses the import's
          // own whitelist); fall back to PNG so render/export stay safe.
          typeof m === "string" && isAllowedRichImageMime(m) ? m : "image/png",
        );
        if (mimes.some((m) => m !== "image/png")) stored.imageMimes = mimes;
      }
    }
  }
  if (typeof record.html === "string") {
    stored.html = truncate(record.html, MAX_HTML_CHARS);
  }
  const rich = normalizeRich(record.rich);
  if (rich) stored.rich = rich;
  return finalizeStoredOutput(stored);
}

function normalizeRich(raw: unknown): RichOutput | undefined {
  if (!raw || typeof raw !== "object") return undefined;
  const record = raw as Record<string, unknown>;
  if (record.kind === "table") {
    const columns = Array.isArray(record.columns)
      ? record.columns.map((c) => String(c)).slice(0, MAX_TABLE_COLS)
      : [];
    const rows = Array.isArray(record.rows)
      ? record.rows
          .filter((row): row is unknown[] => Array.isArray(row))
          .slice(0, MAX_TABLE_ROWS)
          .map((row) => row.slice(0, MAX_TABLE_COLS))
      : [];
    return { kind: "table", columns, rows };
  }
  if (record.kind === "image") {
    // Reject a malformed/unsupported mime (e.g. svg) rather than persisting it.
    if (
      !isAllowedRichImageMime(record.mime) ||
      typeof record.data !== "string"
    ) {
      return undefined;
    }
    if (record.data.length > MAX_IMAGE_CHARS) return undefined;
    return { kind: "image", mime: record.mime, data: record.data };
  }
  if (record.kind === "html") {
    if (typeof record.html !== "string") return undefined;
    return { kind: "html", html: truncate(record.html, MAX_HTML_CHARS) };
  }
  return undefined;
}

/**
 * A stored output is also a renderable `RunResult` for `CellOutput`, since
 * `StoredCellOutput` is the renderable subset of `RunResult`. This adapter makes
 * that intent explicit at the call site (load → render without a run).
 */
export function storedOutputsToRunResult(
  stored: StoredCellOutput | null | undefined,
): RunResult | null {
  if (!stored) return null;
  if (stored.notebookOutputs) {
    return { logs: [], notebookOutputs: stored.notebookOutputs };
  }
  return {
    logs: stored.logs,
    logStreams: stored.logStreams,
    result: stored.result,
    error: stored.error,
    images: stored.images,
    imageMimes: stored.imageMimes,
    html: stored.html,
    rich: stored.rich,
  };
}

// --- nbformat (.ipynb) <-> StoredCellOutput ---

/**
 * Split a multi-line text payload into the canonical nbformat line list
 * (keep-ends: every line retains its trailing "\n"; a final line with no newline
 * keeps none). Unlike a plain split, a text ENDING in "\n" yields no spurious
 * trailing "" element — matching how real Jupyter writes stream/text blocks.
 */
function textToLines(text: string): string[] {
  if (text === "") return [];
  const out: string[] = [];
  let start = 0;
  for (let i = 0; i < text.length; i += 1) {
    if (text[i] === "\n") {
      out.push(text.slice(start, i + 1));
      start = i + 1;
    }
  }
  if (start < text.length) out.push(text.slice(start));
  return out;
}

/** Join an nbformat string|string[] payload back into one string. */
function joinText(value: unknown): string {
  if (typeof value === "string") return value;
  if (Array.isArray(value)) {
    return value.map((line) => (typeof line === "string" ? line : "")).join("");
  }
  return "";
}

/** One nbformat output object (loosely typed; we only emit the fields we use). */
export type NbformatOutput = Record<string, unknown>;

const MAX_NBFORMAT_OUTPUTS_PER_CELL = 2_000;

function isPlainJsonObject(value: unknown): value is Record<string, unknown> {
  if (!value || typeof value !== "object" || Array.isArray(value)) return false;
  const prototype = Object.getPrototypeOf(value);
  return prototype === Object.prototype || prototype === null;
}

const INVALID_JSON_TREE = Symbol("invalid-json-tree");

function ownArrayLength(value: unknown): number | null {
  try {
    if (!Array.isArray(value)) return null;
    const descriptor = Object.getOwnPropertyDescriptor(value, "length");
    return descriptor &&
      "value" in descriptor &&
      Number.isSafeInteger(descriptor.value) &&
      descriptor.value >= 0
      ? descriptor.value
      : null;
  } catch {
    return null;
  }
}

function detachJsonTree(value: unknown): unknown | typeof INVALID_JSON_TREE {
  const ancestors = new WeakSet<object>();
  let nodes = 0;

  const visit = (
    current: unknown,
    depth: number,
  ): unknown | typeof INVALID_JSON_TREE => {
    nodes += 1;
    if (nodes > 300_000 || depth > 256) return INVALID_JSON_TREE;
    if (
      current === null ||
      typeof current === "string" ||
      typeof current === "boolean"
    ) {
      return current;
    }
    if (typeof current === "number") {
      return Number.isFinite(current) ? current : INVALID_JSON_TREE;
    }
    if (!current || typeof current !== "object") return INVALID_JSON_TREE;
    if (ancestors.has(current)) return INVALID_JSON_TREE;
    ancestors.add(current);
    try {
      let isArray = false;
      let prototype: object | null;
      let descriptors: PropertyDescriptorMap;
      try {
        isArray = Array.isArray(current);
        prototype = Object.getPrototypeOf(current);
        descriptors = Object.getOwnPropertyDescriptors(current);
      } catch {
        return INVALID_JSON_TREE;
      }
      const keys = Reflect.ownKeys(descriptors);
      if (keys.some((key) => typeof key !== "string")) {
        return INVALID_JSON_TREE;
      }
      if (isArray) {
        const lengthDescriptor = descriptors.length;
        if (
          prototype !== Array.prototype ||
          !lengthDescriptor ||
          !("value" in lengthDescriptor) ||
          !Number.isSafeInteger(lengthDescriptor.value) ||
          lengthDescriptor.value < 0 ||
          lengthDescriptor.value > 100_000 ||
          keys.length !== lengthDescriptor.value + 1
        ) {
          return INVALID_JSON_TREE;
        }
        const detached: unknown[] = [];
        for (let index = 0; index < lengthDescriptor.value; index += 1) {
          const descriptor = descriptors[String(index)];
          if (
            !descriptor ||
            descriptor.enumerable !== true ||
            !("value" in descriptor)
          ) {
            return INVALID_JSON_TREE;
          }
          const child = visit(descriptor.value, depth + 1);
          if (child === INVALID_JSON_TREE) return INVALID_JSON_TREE;
          detached.push(child);
        }
        return detached;
      }
      if (
        (prototype !== Object.prototype && prototype !== null) ||
        keys.length > 10_000
      ) {
        return INVALID_JSON_TREE;
      }
      const detached = Object.create(null) as Record<string, unknown>;
      for (const key of keys as string[]) {
        const descriptor = descriptors[key];
        if (
          !descriptor ||
          descriptor.enumerable !== true ||
          !("value" in descriptor)
        ) {
          return INVALID_JSON_TREE;
        }
        const child = visit(descriptor.value, depth + 1);
        if (child === INVALID_JSON_TREE) return INVALID_JSON_TREE;
        detached[key] = child;
      }
      return detached;
    } finally {
      ancestors.delete(current);
    }
  };

  return visit(value, 0);
}

function isTextPayload(value: unknown): boolean {
  return (
    typeof value === "string" ||
    (Array.isArray(value) && value.every((part) => typeof part === "string"))
  );
}

const NBFORMAT_JSON_MIME = /^application\/(?:.*\+)?json$/;

function isValidMimeBundle(value: Record<string, unknown>): boolean {
  return Object.entries(value).every(([mime, payload]) =>
    NBFORMAT_JSON_MIME.test(mime) ? true : isTextPayload(payload),
  );
}

/**
 * Validate and detach an ordered nbformat-4 output array. `null` means invalid;
 * an empty array is valid and is represented by an absent cell snapshot.
 */
/**
 * One nbformat-4 output, normalized — or null when it cannot be understood.
 *
 * The single place the output rules live, so the strict snapshot validator and
 * the lenient import salvage can never disagree about what a valid output is.
 *
 * `repair` is what separates the two callers. Reading OUR OWN persisted
 * snapshot, a missing required field means the data is not what we wrote and
 * the snapshot is suspect, so nothing is filled in. Reading someone else's
 * .ipynb, the nbformat schema names exactly one legal value for each of those
 * omissions and Jupyter's reader fills the same one in — a stream with no
 * `name` is stdout, a display bundle with no `metadata` has none, an error with
 * no `traceback` has none — so repairing keeps data a real notebook wrote.
 *
 * Either way, only an ABSENT field is ever filled in. A field that is present
 * and the wrong type is garbage, not a default waiting to happen.
 */
function normalizeOneNbformatOutput(
  value: unknown,
  repair: boolean,
): NbformatOutput | null {
  if (!isPlainJsonObject(value)) return null;
  const type = value.output_type;
  if (type === "stream") {
    if (
      Object.keys(value).some(
        (key) => !["output_type", "name", "text"].includes(key),
      ) ||
      !isTextPayload(value.text) ||
      // Repair only an OMISSION. A name that is present and not a string is
      // garbage, not a default waiting to be filled in, and quietly turning it
      // into "stdout" would be the confidently-wrong answer this file avoids.
      (repair
        ? value.name !== undefined &&
          value.name !== null &&
          typeof value.name !== "string"
        : typeof value.name !== "string")
    ) {
      return null;
    }
    // Absent means stdout — the value the nbformat schema names and the one
    // Jupyter's own reader fills in. A name that IS there is kept verbatim.
    const name = typeof value.name === "string" ? value.name : "stdout";
    return { ...value, name } as unknown as NbformatOutput;
  }
  if (type === "error") {
    if (
      Object.keys(value).some(
        (key) =>
          !["output_type", "ename", "evalue", "traceback"].includes(key),
      ) ||
      typeof value.ename !== "string" ||
      typeof value.evalue !== "string" ||
      !(repair && value.traceback === undefined) &&
      (!Array.isArray(value.traceback) ||
        !value.traceback.every((line) => typeof line === "string"))
    ) {
      return null;
    }
    return {
      ...value,
      traceback: Array.isArray(value.traceback) ? value.traceback : [],
    } as unknown as NbformatOutput;
  }
  if (type === "display_data" || type === "execute_result") {
    const allowed =
      type === "display_data"
        ? ["output_type", "data", "metadata"]
        : ["output_type", "data", "metadata", "execution_count"];
    if (
      Object.keys(value).some((key) => !allowed.includes(key)) ||
      !isPlainJsonObject(value.data) ||
      !(repair && value.metadata === undefined
        ? true
        : isPlainJsonObject(value.metadata)) ||
      !isValidMimeBundle(value.data)
    ) {
      return null;
    }
    // `execution_count` is OPTIONAL on an execute_result — the v4 schema
    // requires only output_type, data and metadata. Absent is not the same as
    // invalid, and treating it as invalid rejected notebooks Jupyter opens.
    if (
      type === "execute_result" &&
      value.execution_count !== undefined &&
      value.execution_count !== null &&
      (typeof value.execution_count !== "number" ||
        !Number.isSafeInteger(value.execution_count) ||
        value.execution_count < 0)
    ) {
      return null;
    }
    return {
      ...value,
      metadata: isPlainJsonObject(value.metadata) ? value.metadata : {},
    } as unknown as NbformatOutput;
  }
  return null;
}

export function normalizeNbformatOutputsSnapshot(
  raw: unknown,
): NbformatOutput[] | null {
  const detachedRaw = detachJsonTree(raw);
  if (
    detachedRaw === INVALID_JSON_TREE ||
    !Array.isArray(detachedRaw) ||
    detachedRaw.length > MAX_NBFORMAT_OUTPUTS_PER_CELL
  ) {
    return null;
  }
  const normalized: NbformatOutput[] = [];
  for (const value of detachedRaw) {
    const one = normalizeOneNbformatOutput(value, false);
    if (one === null) return null;
    normalized.push(one);
  }
  try {
    const detached = detachJsonTree(normalized);
    if (detached === INVALID_JSON_TREE || !Array.isArray(detached)) return null;
    const serialized = JSON.stringify({ logs: [], notebookOutputs: detached });
    return serialized.length <= MAX_PERSISTED_CELL_OUTPUT_JSON_CHARS
      ? (detached as NbformatOutput[])
      : null;
  } catch {
    return null;
  }
}

/** What an import could keep, and how much it could not. */
export interface SalvagedNbformatOutputs {
  outputs: NbformatOutput[];
  /** Outputs that could not be understood and were left out. */
  dropped: number;
}

/**
 * Read an imported cell's outputs, keeping everything that can be understood.
 *
 * The difference from {@link normalizeNbformatOutputsSnapshot} is what happens
 * to the rest of the notebook when one output is strange. That function is the
 * gate for OUR OWN persisted data, where a single unreadable entry means the
 * snapshot is not what we wrote and all of it is suspect. This one reads a file
 * someone else produced, where the same strictness cost the user their code:
 * one output from a tool we do not know, and the whole notebook refused to
 * open.
 *
 * Outputs are the one part of a notebook that is REGENERABLE — running the cell
 * makes them again — and source is not. So an output nobody can parse is left
 * out and counted, and the notebook opens.
 */
export function salvageNbformatOutputs(raw: unknown): SalvagedNbformatOutputs {
  const detachedRaw = detachJsonTree(raw);
  if (detachedRaw === INVALID_JSON_TREE || !Array.isArray(detachedRaw)) {
    return { outputs: [], dropped: ownArrayLength(raw) ?? 0 };
  }
  const outputs: NbformatOutput[] = [];
  let dropped = 0;
  for (const value of detachedRaw) {
    if (outputs.length >= MAX_NBFORMAT_OUTPUTS_PER_CELL) {
      dropped += 1;
      continue;
    }
    const one = normalizeOneNbformatOutput(value, true);
    if (one === null) dropped += 1;
    else outputs.push(one);
  }
  // The persisted size cap applies to what survives. Trimming from the END
  // keeps the beginning of a run, which is where a traceback or the first
  // printed lines are — the part someone reads to find out what happened.
  while (
    outputs.length > 0 &&
    JSON.stringify({ logs: [], notebookOutputs: outputs }).length >
      MAX_PERSISTED_CELL_OUTPUT_JSON_CHARS
  ) {
    outputs.pop();
    dropped += 1;
  }
  return { outputs, dropped };
}

/**
 * Serialize a `StoredCellOutput` into nbformat-4 output objects:
 *   - logs → `stream` outputs, coalesced by consecutive stream (stdout/stderr);
 *   - error → an `error` output (ename/evalue/traceback);
 *   - rich table/html, legacy `html` → `display_data` (text/html);
 *   - images, rich image → `display_data` (image/png or the image mime);
 *   - text `result` → `execute_result` (text/plain).
 *
 * Order mirrors the renderer (error, logs, images, html, rich, result) so the
 * .ipynb reads top-to-bottom like the cell. Pure.
 */
export function storedOutputsToNbformat(
  stored: StoredCellOutput | null | undefined,
  executionCount?: number,
): NbformatOutput[] {
  if (!stored) return [];
  if (stored.notebookOutputs) {
    return normalizeNbformatOutputsSnapshot(stored.notebookOutputs) ?? [];
  }
  const outputs: NbformatOutput[] = [];

  // Coalesce consecutive log lines that share a stream into one stream output
  // (nbformat convention: one stream block per contiguous run).
  for (let i = 0; i < stored.logs.length; ) {
    const stream = stored.logStreams?.[i] === "stderr" ? "stderr" : "stdout";
    let j = i;
    const chunk: string[] = [];
    while (
      j < stored.logs.length &&
      (stored.logStreams?.[j] === "stderr" ? "stderr" : "stdout") === stream
    ) {
      // Each captured line is one printed line; restore the newline between them.
      chunk.push(stored.logs[j]);
      j += 1;
    }
    outputs.push({
      output_type: "stream",
      name: stream,
      // Emit a real block terminator (trailing \n, the nbformat norm) so the
      // import side's "drop one trailing empty line" is always absorbing the
      // terminator — never a genuinely-printed trailing blank line. Without it,
      // `join("\n")` produced no terminator and the import pop ate a real blank.
      text: textToLines(`${chunk.join("\n")}\n`),
    });
    i = j;
  }

  const storedImages = stored.images ?? [];
  storedImages.forEach((b64, index) => {
    // Emit each image under its real MIME (aligned imageMimes; default PNG for
    // runtime/legacy snapshots) so a non-PNG display_data image round-trips.
    const mime = stored.imageMimes?.[index] ?? "image/png";
    outputs.push({
      output_type: "display_data",
      data: { [mime]: textToLines(b64) },
      metadata: {},
    });
  });

  if (stored.html !== undefined) {
    outputs.push({
      output_type: "display_data",
      data: { "text/html": textToLines(stored.html) },
      metadata: {},
    });
  }

  const rich = stored.rich;
  if (rich) {
    if (rich.kind === "html") {
      outputs.push({
        output_type: "display_data",
        data: { "text/html": textToLines(rich.html) },
        metadata: {},
      });
    } else if (rich.kind === "image") {
      // Tag with our private metadata so a re-import restores it as `rich`
      // (a structured final-value image) rather than a standalone `images`
      // figure — both render identically, but this keeps the round-trip exact.
      outputs.push({
        output_type: "display_data",
        data: { [rich.mime]: textToLines(rich.data) },
        metadata: { "mrnote.rich": { kind: "image", mime: rich.mime } },
      });
    } else {
      // Render a table to text/plain so other notebook tools can read it.
      outputs.push({
        output_type: "display_data",
        data: { "text/plain": textToLines(richTableToText(rich)) },
        metadata: { "mrnote.rich": serializeRichTable(rich) },
      });
    }
  }

  if (stored.result !== undefined) {
    outputs.push({
      output_type: "execute_result",
      execution_count:
        typeof executionCount === "number" && executionCount > 0
          ? executionCount
          : null,
      data: { "text/plain": textToLines(stored.result) },
      metadata: {},
    });
  }

  // The error/traceback is emitted LAST: any stdout/stderr printed before the
  // exception — and any display_data shown during the run — precede it, matching
  // how Jupyter orders and renders a failed cell's outputs (the traceback comes
  // after the prints, not before them).
  if (stored.error !== undefined) {
    const lines = stored.error.replace(/\r\n/g, "\n").split("\n");
    // First line is the most useful summary; keep the whole text as traceback.
    const ename = lines[0]?.split(":")[0]?.trim() || "Error";
    outputs.push({
      output_type: "error",
      ename,
      evalue: stored.error,
      traceback: lines,
    });
  }

  return outputs;
}

/** Plain-text rendering of a rich table (TSV) for nbformat text/plain. */
function richTableToText(rich: Extract<RichOutput, { kind: "table" }>): string {
  const header = rich.columns.join("\t");
  const body = rich.rows
    .map((row) =>
      row
        .map((cell) =>
          cell === null || cell === undefined
            ? ""
            : typeof cell === "object"
              ? safeJson(cell)
              : String(cell),
        )
        .join("\t"),
    )
    .join("\n");
  return body ? `${header}\n${body}` : header;
}

function serializeRichTable(rich: Extract<RichOutput, { kind: "table" }>) {
  return { kind: "table", columns: rich.columns, rows: rich.rows };
}

function safeJson(value: unknown): string {
  try {
    const serialized = JSON.stringify(value);
    return typeof serialized === "string" ? serialized : String(value);
  } catch {
    return String(value);
  }
}

/**
 * Read nbformat outputs off a parsed `.ipynb` code cell back into a
 * `StoredCellOutput`, or `null` when there is nothing renderable. Inverse of
 * `storedOutputsToNbformat` (best-effort across other tools' notebooks):
 *   - `stream` → logs (+ stderr tags);
 *   - `error` → error text (joined traceback, falling back to ename/evalue);
 *   - `execute_result`/`display_data` text/plain → result (preferring a
 *     reconstructed rich table from our metadata);
 *   - text/html → html; image/* → images or a rich image.
 *
 * Pure + defensive; applies the same size caps as the live path.
 */
export function nbformatToStoredOutputs(
  cell: unknown,
): StoredCellOutput | null {
  if (!cell || typeof cell !== "object") return null;
  const raw = (cell as { outputs?: unknown }).outputs;
  if (!Array.isArray(raw)) return null;

  const logs: string[] = [];
  const logStreams: LogStream[] = [];
  let error: string | undefined;
  let result: string | undefined;
  let html: string | undefined;
  const images: string[] = [];
  const imageMimes: string[] = [];
  let rich: RichOutput | undefined;

  for (const out of raw) {
    if (!out || typeof out !== "object") continue;
    const o = out as Record<string, unknown>;
    const type = o.output_type;

    if (type === "stream") {
      const stream: LogStream = o.name === "stderr" ? "stderr" : "stdout";
      // nbformat stream text is the raw block; split into display lines,
      // dropping a single trailing empty line from the block's trailing \n.
      const text = joinText(o.text);
      const parts = text.split("\n");
      if (parts.length > 1 && parts[parts.length - 1] === "") parts.pop();
      for (const line of parts) {
        if (logs.length >= MAX_LOG_LINES) break;
        logs.push(truncate(line, MAX_LOG_LINE_CHARS));
        logStreams.push(stream);
      }
      continue;
    }

    if (type === "error") {
      const traceback = Array.isArray(o.traceback)
        ? o.traceback
            .map((line) => (typeof line === "string" ? line : ""))
            .join("\n")
        : "";
      const summary = [o.ename, o.evalue]
        .filter((part) => typeof part === "string" && part)
        .join(": ");
      error = truncate(traceback || summary || "Error", MAX_TEXT_CHARS);
      continue;
    }

    if (type === "execute_result" || type === "display_data") {
      const data = (o.data ?? {}) as Record<string, unknown>;
      const meta = (o.metadata ?? {}) as Record<string, unknown>;
      // Reconstruct a structured rich output from our own metadata when present.
      const metaRichRaw = meta["mrnote.rich"];
      if (metaRichRaw && typeof metaRichRaw === "object" && !rich) {
        const metaRecord = metaRichRaw as Record<string, unknown>;
        if (
          metaRecord.kind === "image" &&
          isAllowedRichImageMime(metaRecord.mime)
        ) {
          // The base64 lives in `o.data[mime]`, not the metadata.
          const b64 = joinText(data[metaRecord.mime]);
          if (b64 && b64.length <= MAX_IMAGE_CHARS) {
            rich = { kind: "image", mime: metaRecord.mime, data: b64 };
            continue;
          }
        } else {
          // Non-image rich (table/html), or an image with an unsupported mime:
          // normalizeRich gates it (a disallowed image mime → undefined → drop).
          const metaRich = normalizeRich(metaRichRaw);
          if (metaRich) {
            rich = metaRich;
            continue;
          }
        }
      }
      const htmlPayload = data["text/html"];
      if (typeof htmlPayload !== "undefined" && html === undefined && !rich) {
        html = truncate(joinText(htmlPayload), MAX_HTML_CHARS);
        continue;
      }
      // Only accept whitelisted raster image mimes; skip e.g. image/svg+xml
      // (markup, a script-injection vector) so it never imports as an image.
      const imageMime = Object.keys(data).find((key) =>
        isAllowedRichImageMime(key),
      );
      if (imageMime) {
        const b64 = joinText(data[imageMime]);
        if (b64 && b64.length <= MAX_IMAGE_CHARS) {
          if (type === "display_data") {
            if (images.length < MAX_IMAGES) {
              images.push(b64);
              imageMimes.push(imageMime);
            }
          } else if (!rich) {
            rich = { kind: "image", mime: imageMime, data: b64 };
          }
        }
        continue;
      }
      const plain = data["text/plain"];
      if (typeof plain !== "undefined") {
        const text = truncate(joinText(plain), MAX_TEXT_CHARS);
        // execute_result is the cell's value; display_data text/plain is a
        // secondary repr — only take it as `result` if we don't have one yet.
        if (type === "execute_result" || result === undefined) result = text;
      }
      continue;
    }
  }

  const stored: MutableStoredOutput = { logs };
  if (logStreams.some((s) => s === "stderr")) stored.logStreams = logStreams;
  if (result !== undefined) stored.result = result;
  if (error !== undefined) stored.error = error;
  if (images.length > 0) {
    stored.images = images;
    // Only persist per-image MIMEs when at least one isn't the default PNG, so
    // the common all-PNG case stays byte-identical to before.
    if (imageMimes.some((m) => m !== "image/png"))
      stored.imageMimes = imageMimes;
  }
  if (html !== undefined) stored.html = html;
  if (rich) stored.rich = rich;

  return finalizeStoredOutput(stored);
}
