/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Bounded last-run output persistence for standalone code pages.
 *
 * multi-buffer host cells already persist a normalized `StoredCellOutput`. Standalone
 * code pages use the same representation under an additive
 * `lastRunOutput` field so a completed local run survives save/reopen without
 * introducing another output format or widening the runner/sandbox boundary.
 */

import type { RunResult } from "@mrnote/core/web/lib/code-runner/js-runner.ts";
import type {
  ExecutionLineMapping,
  ExecutionSelection,
} from "./execution-selection";
import {
  normalizeCodeSourceEncoding,
  normalizeCodeSourceFilename,
  type CodeSourceEncoding,
} from "./code-file-import";
import {
  CODE_WORKSPACE_REFERENCE_KEY,
  persistedCodeWorkspaceReference,
  type PersistedCodeWorkspaceReference,
} from "./code-workspace";
import type { EolStyle } from "./editor-commands";
import {
  CODE_PAGE_DEBUG_CONFIGURATION_KEY,
  CODE_PAGE_DEBUG_TIMELINE_KEY,
  normalizeCodeDebugConfiguration,
  normalizeCodeDebugTimeline,
  type CodeDebugConfiguration,
  type PersistedCodeDebugTimelineEntry,
} from "./code-page-debug-persistence";
import {
  normalizeStoredOutputs,
  runResultToStoredOutputs,
  storedOutputsToRunResult,
  type StoredCellOutput,
} from "./output-persistence";

export const CODE_PAGE_LAST_RUN_OUTPUT_KEY = "lastRunOutput";
export const CODE_PAGE_LAST_RUN_NAVIGATION_KEY = "lastRunNavigation";
export const CODE_PAGE_RUN_TIMELINE_KEY = "runTimeline";
export const CODE_PAGE_RUN_OUTPUT_CLEARED_KEY = "lastRunOutputCleared";
export const CODE_PAGE_RUN_PROFILES_KEY = "runProfiles";
export const CODE_PAGE_SOURCE_FILENAME_KEY = "sourceFileName";
export const CODE_PAGE_LINE_ENDING_KEY = "lineEnding";
export const CODE_PAGE_SOURCE_ENCODING_KEY = "sourceEncoding";
export const CODE_PAGE_RUN_NAVIGATION_RANGE_LIMIT = 2_048;
export const CODE_PAGE_RUN_TIMELINE_ENTRY_LIMIT = 12;
export const CODE_PAGE_RUN_TIMELINE_UTF8_BYTE_LIMIT = 1_000_000;
export const CODE_PAGE_RUN_PROFILE_LIMIT = 8;
export const CODE_PAGE_RUN_PROFILES_UTF8_BYTE_LIMIT = 128_000;
export const CODE_PAGE_RUN_STDIN_UTF8_BYTE_LIMIT = 1_000_000;
const CODE_PAGE_RUN_NAVIGATION_SOURCE_LIMIT = 1_000_000;
const CODE_PAGE_RUN_NAVIGATION_VERSION = 1;
const CODE_PAGE_RUN_INPUT_IDENTITY_VERSION = 1;
const CODE_PAGE_RUN_TIMELINE_VERSION = 2;
const CODE_PAGE_RUN_PROFILE_VERSION = 2;
const CODE_PAGE_RUN_TIMELINE_PROVIDER_ID_LIMIT = 128;
const CODE_PAGE_RUN_TIMELINE_LANGUAGE_LIMIT = 32;
const MAX_JAVASCRIPT_TIME = 8_640_000_000_000_000;

interface PersistedExecutionSelection {
  readonly kind: ExecutionSelection["kind"];
  readonly from: number;
  readonly to: number;
  readonly lineFrom: number;
  readonly lineTo: number;
  readonly selectedChars: number;
  readonly rangeCount: number;
  readonly lineMappings: readonly ExecutionLineMapping[];
}

export interface PersistedCodeRunNavigation {
  readonly version: typeof CODE_PAGE_RUN_NAVIGATION_VERSION;
  readonly sourceFingerprint: string;
  /** `null` represents a full-document run. */
  readonly selection: PersistedExecutionSelection | null;
}

export interface RestoredCodeRunNavigation {
  readonly persisted: PersistedCodeRunNavigation;
  /** Source text is intentionally omitted; only line mapping is restored. */
  readonly selection: ExecutionSelection | null;
}

/**
 * Non-secret identity for the exact standard input used by one accepted run.
 * The input body is deliberately never persisted.
 */
export interface PersistedCodeRunInputIdentity {
  readonly version: typeof CODE_PAGE_RUN_INPUT_IDENTITY_VERSION;
  readonly stdinSha256: string;
  readonly stdinBytes: number;
}

export interface PersistedCodeRunTimelineEntry {
  readonly version: typeof CODE_PAGE_RUN_TIMELINE_VERSION;
  readonly completedAt: number;
  readonly language: string;
  readonly providerId: string;
  readonly sourceSha256: string;
  readonly outcome: "success" | "error";
  /** Exact process exit code from a validated provider receipt, when present. */
  readonly exitCode?: number | null;
  /** A successful run with no visible output is represented by `null`. */
  readonly output: StoredCellOutput | null;
  readonly input: PersistedCodeRunInputIdentity;
  readonly navigation: PersistedCodeRunNavigation;
}

/**
 * A user-saved rerun recipe. It stores no source, stdin, output, credentials,
 * or runtime state: dispatch always re-authorizes the current provider and the
 * SHA-256 binds the compact selection metadata to the exact source snapshot.
 */
export interface PersistedCodeRunProfile {
  readonly version: typeof CODE_PAGE_RUN_PROFILE_VERSION;
  readonly savedAt: number;
  readonly language: string;
  readonly providerId: string;
  readonly sourceSha256: string;
  readonly input: PersistedCodeRunInputIdentity;
  readonly navigation: PersistedCodeRunNavigation;
}

export interface PersistedCodePageContent extends Record<string, unknown> {
  readonly type: string;
  readonly language: string;
  readonly content: string;
  readonly sourceFileName?: string;
  readonly lineEnding?: EolStyle;
  readonly sourceEncoding?: CodeSourceEncoding;
  readonly lastRunOutput?: StoredCellOutput;
  readonly lastRunNavigation?: PersistedCodeRunNavigation;
  readonly runTimeline?: readonly PersistedCodeRunTimelineEntry[];
  readonly lastRunOutputCleared?: PersistedCodeRunOutputCleared;
  readonly runProfiles?: readonly PersistedCodeRunProfile[];
  readonly debugTimeline?: readonly PersistedCodeDebugTimelineEntry[];
  readonly debugConfiguration?: CodeDebugConfiguration;
  readonly codeWorkspace?: PersistedCodeWorkspaceReference;
}

function finiteInteger(
  value: unknown,
  minimum: number,
  maximum: number,
): number | null {
  if (
    typeof value !== "number" ||
    !Number.isFinite(value) ||
    !Number.isInteger(value)
  ) {
    return null;
  }
  return value >= minimum && value <= maximum ? value : null;
}

function exactObjectKeys(
  value: Record<string, unknown>,
  required: readonly string[],
): boolean {
  const allowed = new Set(required);
  return (
    required.every((key) => Object.prototype.hasOwnProperty.call(value, key)) &&
    Object.keys(value).every((key) => allowed.has(key))
  );
}

function codeSourceFingerprint(source: string): string {
  let fnv = 0x811c9dc5;
  let djb = 0x1505;
  for (let index = 0; index < source.length; index += 1) {
    const unit = source.charCodeAt(index);
    fnv = Math.imul(fnv ^ unit, 0x01000193) >>> 0;
    djb = (Math.imul(djb, 33) ^ unit) >>> 0;
  }
  return `v1:${source.length}:${fnv.toString(36)}:${djb.toString(36)}`;
}

/** One bounded full-source fingerprint shared by last-output and timeline authority. */
export function codePageRunSourceFingerprint(source: string): string | null {
  return source.length <= CODE_PAGE_RUN_NAVIGATION_SOURCE_LIMIT
    ? codeSourceFingerprint(source)
    : null;
}

function codeSourceLineCount(source: string): number {
  let count = 1;
  for (let index = 0; index < source.length; index += 1) {
    if (source.charCodeAt(index) === 10) count += 1;
  }
  return count;
}

function normalizePersistedSelection(
  value: unknown,
): PersistedExecutionSelection | null {
  if (!value || typeof value !== "object" || Array.isArray(value)) return null;
  const raw = value as Record<string, unknown>;
  const kind =
    raw.kind === "selection" || raw.kind === "line" ? raw.kind : null;
  const from = finiteInteger(
    raw.from,
    0,
    CODE_PAGE_RUN_NAVIGATION_SOURCE_LIMIT,
  );
  const to = finiteInteger(
    raw.to,
    0,
    CODE_PAGE_RUN_NAVIGATION_SOURCE_LIMIT,
  );
  const lineFrom = finiteInteger(
    raw.lineFrom,
    1,
    CODE_PAGE_RUN_NAVIGATION_SOURCE_LIMIT + 1,
  );
  const lineTo = finiteInteger(
    raw.lineTo,
    1,
    CODE_PAGE_RUN_NAVIGATION_SOURCE_LIMIT + 1,
  );
  const selectedChars = finiteInteger(
    raw.selectedChars,
    0,
    CODE_PAGE_RUN_NAVIGATION_SOURCE_LIMIT,
  );
  const rangeCount = finiteInteger(
    raw.rangeCount,
    1,
    CODE_PAGE_RUN_NAVIGATION_RANGE_LIMIT,
  );
  if (
    !kind ||
    from === null ||
    to === null ||
    from > to ||
    lineFrom === null ||
    lineTo === null ||
    lineFrom > lineTo ||
    selectedChars === null ||
    rangeCount === null ||
    !Array.isArray(raw.lineMappings) ||
    raw.lineMappings.length !== rangeCount
  ) {
    return null;
  }

  const lineMappings: ExecutionLineMapping[] = [];
  let previousExecutionLineTo = 0;
  let previousDocumentLineFrom = 0;
  for (const value of raw.lineMappings) {
    if (!value || typeof value !== "object" || Array.isArray(value)) {
      return null;
    }
    const mapping = value as Record<string, unknown>;
    const executionLineFrom = finiteInteger(
      mapping.executionLineFrom,
      1,
      CODE_PAGE_RUN_NAVIGATION_SOURCE_LIMIT +
        CODE_PAGE_RUN_NAVIGATION_RANGE_LIMIT,
    );
    const executionLineTo = finiteInteger(
      mapping.executionLineTo,
      1,
      CODE_PAGE_RUN_NAVIGATION_SOURCE_LIMIT +
        CODE_PAGE_RUN_NAVIGATION_RANGE_LIMIT,
    );
    const documentLineFrom = finiteInteger(
      mapping.documentLineFrom,
      1,
      CODE_PAGE_RUN_NAVIGATION_SOURCE_LIMIT + 1,
    );
    if (
      executionLineFrom === null ||
      executionLineTo === null ||
      executionLineFrom > executionLineTo ||
      executionLineFrom <= previousExecutionLineTo ||
      documentLineFrom === null ||
      documentLineFrom < previousDocumentLineFrom ||
      documentLineFrom < lineFrom ||
      documentLineFrom + executionLineTo - executionLineFrom > lineTo
    ) {
      return null;
    }
    lineMappings.push({
      executionLineFrom,
      executionLineTo,
      documentLineFrom,
    });
    previousExecutionLineTo = executionLineTo;
    previousDocumentLineFrom = documentLineFrom;
  }

  if (
    (kind === "line" &&
      (rangeCount !== 1 ||
        lineFrom !== lineTo ||
        lineMappings.length !== 1)) ||
    selectedChars > to - from
  ) {
    return null;
  }

  return {
    kind,
    from,
    to,
    lineFrom,
    lineTo,
    selectedChars,
    rangeCount,
    lineMappings,
  };
}

function normalizeCodeRunNavigation(
  value: unknown,
): PersistedCodeRunNavigation | null {
  if (!value || typeof value !== "object" || Array.isArray(value)) return null;
  const raw = value as Record<string, unknown>;
  if (
    raw.version !== CODE_PAGE_RUN_NAVIGATION_VERSION ||
    typeof raw.sourceFingerprint !== "string" ||
    raw.sourceFingerprint.length > 64 ||
    !/^v1:[0-9]+:[0-9a-z]+:[0-9a-z]+$/.test(raw.sourceFingerprint)
  ) {
    return null;
  }
  if (raw.selection === null) {
    return {
      version: CODE_PAGE_RUN_NAVIGATION_VERSION,
      sourceFingerprint: raw.sourceFingerprint,
      selection: null,
    };
  }
  const selection = normalizePersistedSelection(raw.selection);
  return selection
    ? {
        version: CODE_PAGE_RUN_NAVIGATION_VERSION,
        sourceFingerprint: raw.sourceFingerprint,
        selection,
      }
    : null;
}

/**
 * Build a bounded navigation snapshot for the source that actually ran.
 * Selection text is never duplicated into page JSON.
 */
export function createCodePageRunNavigation(
  source: string,
  selection: ExecutionSelection | null,
): PersistedCodeRunNavigation | null {
  if (source.length > CODE_PAGE_RUN_NAVIGATION_SOURCE_LIMIT) return null;
  const persistedSelection = selection
    ? normalizePersistedSelection({
        ...selection,
        lineMappings: selection.lineMappings,
      })
    : null;
  if (selection && !persistedSelection) return null;
  if (
    persistedSelection &&
    (persistedSelection.to > source.length ||
      persistedSelection.lineTo > codeSourceLineCount(source))
  ) {
    return null;
  }
  return {
    version: CODE_PAGE_RUN_NAVIGATION_VERSION,
    sourceFingerprint: codeSourceFingerprint(source),
    selection: persistedSelection,
  };
}

function normalizedRunTimelineIdentifier(
  value: unknown,
  maximum: number,
): string | null {
  return typeof value === "string" &&
    value.length > 0 &&
    value.length <= maximum &&
    /^[A-Za-z0-9._:+#-]+$/.test(value)
    ? value
    : null;
}

function normalizeCodeRunInputIdentity(
  value: unknown,
): PersistedCodeRunInputIdentity | null {
  if (!value || typeof value !== "object" || Array.isArray(value)) return null;
  const raw = value as Record<string, unknown>;
  const stdinBytes = finiteInteger(
    raw.stdinBytes,
    0,
    CODE_PAGE_RUN_STDIN_UTF8_BYTE_LIMIT,
  );
  if (
    !exactObjectKeys(raw, ["version", "stdinSha256", "stdinBytes"]) ||
    raw.version !== CODE_PAGE_RUN_INPUT_IDENTITY_VERSION ||
    typeof raw.stdinSha256 !== "string" ||
    !/^[0-9a-f]{64}$/.test(raw.stdinSha256) ||
    stdinBytes === null
  ) {
    return null;
  }
  return {
    version: CODE_PAGE_RUN_INPUT_IDENTITY_VERSION,
    stdinSha256: raw.stdinSha256,
    stdinBytes,
  };
}

export function codePageRunInputIdentityIsCurrent(
  input: PersistedCodeRunInputIdentity,
  stdinSha256: string | null,
  stdinBytes: number,
): boolean {
  return (
    stdinSha256 !== null &&
    input.stdinSha256 === stdinSha256 &&
    input.stdinBytes === stdinBytes
  );
}

function normalizeCodeRunTimelineEntry(
  value: unknown,
): PersistedCodeRunTimelineEntry | null {
  if (!value || typeof value !== "object" || Array.isArray(value)) return null;
  try {
    const raw = value as Record<string, unknown>;
    const completedAt = finiteInteger(raw.completedAt, 0, MAX_JAVASCRIPT_TIME);
    const language = normalizedRunTimelineIdentifier(
      raw.language,
      CODE_PAGE_RUN_TIMELINE_LANGUAGE_LIMIT,
    );
    const providerId = normalizedRunTimelineIdentifier(
      raw.providerId,
      CODE_PAGE_RUN_TIMELINE_PROVIDER_ID_LIMIT,
    );
    const outcome =
      raw.outcome === "success" || raw.outcome === "error"
        ? raw.outcome
        : null;
    const exitCode =
      raw.exitCode === undefined || raw.exitCode === null
        ? null
        : finiteInteger(raw.exitCode, 0, 255);
    const navigation = normalizeCodeRunNavigation(raw.navigation);
    const input = normalizeCodeRunInputIdentity(raw.input);
    const output =
      raw.output === null ? null : normalizeStoredOutputs(raw.output);
    const timelineKeys = [
      "version",
      "completedAt",
      "language",
      "providerId",
      "sourceSha256",
      "outcome",
      "output",
      "input",
      "navigation",
    ] as const;
    if (
      !(
        exactObjectKeys(raw, timelineKeys) ||
        exactObjectKeys(raw, [...timelineKeys, "exitCode"])
      ) ||
      raw.version !== CODE_PAGE_RUN_TIMELINE_VERSION ||
      completedAt === null ||
      !language ||
      !providerId ||
      typeof raw.sourceSha256 !== "string" ||
      !/^[0-9a-f]{64}$/.test(raw.sourceSha256) ||
      !outcome ||
      (raw.exitCode !== undefined &&
        raw.exitCode !== null &&
        exitCode === null) ||
      (outcome === "success" && exitCode !== null && exitCode !== 0) ||
      (outcome === "error" && exitCode === 0) ||
      !input ||
      !navigation ||
      (raw.output !== null && !output) ||
      (outcome === "error" && output?.error === undefined) ||
      (outcome === "success" && output?.error !== undefined)
    ) {
      return null;
    }
    return {
      version: CODE_PAGE_RUN_TIMELINE_VERSION,
      completedAt,
      language,
      providerId,
      sourceSha256: raw.sourceSha256,
      outcome,
      exitCode,
      output,
      input,
      navigation,
    };
  } catch {
    return null;
  }
}

function normalizeCodeRunProfile(
  value: unknown,
): PersistedCodeRunProfile | null {
  if (!value || typeof value !== "object" || Array.isArray(value)) return null;
  try {
    const raw = value as Record<string, unknown>;
    const savedAt = finiteInteger(raw.savedAt, 0, MAX_JAVASCRIPT_TIME);
    const language = normalizedRunTimelineIdentifier(
      raw.language,
      CODE_PAGE_RUN_TIMELINE_LANGUAGE_LIMIT,
    );
    const providerId = normalizedRunTimelineIdentifier(
      raw.providerId,
      CODE_PAGE_RUN_TIMELINE_PROVIDER_ID_LIMIT,
    );
    const navigation = normalizeCodeRunNavigation(raw.navigation);
    const input = normalizeCodeRunInputIdentity(raw.input);
    if (
      !exactObjectKeys(raw, [
        "version",
        "savedAt",
        "language",
        "providerId",
        "sourceSha256",
        "input",
        "navigation",
      ]) ||
      raw.version !== CODE_PAGE_RUN_PROFILE_VERSION ||
      savedAt === null ||
      !language ||
      !providerId ||
      typeof raw.sourceSha256 !== "string" ||
      !/^[0-9a-f]{64}$/.test(raw.sourceSha256) ||
      !input ||
      !navigation ||
      (navigation.selection !== null &&
        navigation.selection.rangeCount !== 1)
    ) {
      return null;
    }
    return {
      version: CODE_PAGE_RUN_PROFILE_VERSION,
      savedAt,
      language,
      providerId,
      sourceSha256: raw.sourceSha256,
      input,
      navigation,
    };
  } catch {
    return null;
  }
}

function runProfilesUtf8Bytes(value: unknown): number | null {
  try {
    return new TextEncoder().encode(JSON.stringify(value)).byteLength;
  } catch {
    return null;
  }
}

/** Strict all-or-nothing restore: malformed or oversized saved recipes vanish. */
export function normalizeCodePageRunProfiles(
  value: unknown,
): readonly PersistedCodeRunProfile[] {
  if (
    !Array.isArray(value) ||
    value.length > CODE_PAGE_RUN_PROFILE_LIMIT ||
    (runProfilesUtf8Bytes(value) ?? Infinity) >
      CODE_PAGE_RUN_PROFILES_UTF8_BYTE_LIMIT
  ) {
    return [];
  }
  const profiles: PersistedCodeRunProfile[] = [];
  const identities = new Set<string>();
  for (const candidate of value) {
    const profile = normalizeCodeRunProfile(candidate);
    if (!profile) return [];
    const identity = `${profile.language}:${profile.sourceSha256}:${profile.input.stdinSha256}:${profile.input.stdinBytes}:${JSON.stringify(profile.navigation.selection)}`;
    if (identities.has(identity)) return [];
    identities.add(identity);
    profiles.push(profile);
  }
  return profiles;
}

export function createCodePageRunProfile(input: {
  readonly savedAt: number;
  readonly entry: PersistedCodeRunTimelineEntry;
  readonly sourceSha256: string;
}): PersistedCodeRunProfile | null {
  return normalizeCodeRunProfile({
    version: CODE_PAGE_RUN_PROFILE_VERSION,
    savedAt: input.savedAt,
    language: input.entry.language,
    providerId: input.entry.providerId,
    sourceSha256: input.sourceSha256,
    input: input.entry.input,
    navigation: input.entry.navigation,
  });
}

type CodeRunProfileSlotIdentity = Pick<
  PersistedCodeRunProfile,
  "language" | "sourceSha256" | "input" | "navigation"
>;

/** One saved slot per exact source + stdin + execution scope. */
export function codePageRunProfilesShareSlot(
  left: CodeRunProfileSlotIdentity,
  right: CodeRunProfileSlotIdentity,
): boolean {
  return (
    left.language === right.language &&
    left.sourceSha256 === right.sourceSha256 &&
    left.input.stdinSha256 === right.input.stdinSha256 &&
    left.input.stdinBytes === right.input.stdinBytes &&
    JSON.stringify(left.navigation.selection) ===
      JSON.stringify(right.navigation.selection)
  );
}

/** Save newest-first and replace an existing recipe for the same exact scope. */
export function saveCodePageRunProfile(
  current: readonly PersistedCodeRunProfile[],
  value: PersistedCodeRunProfile,
): readonly PersistedCodeRunProfile[] {
  const profile = normalizeCodeRunProfile(value);
  if (!profile) return normalizeCodePageRunProfiles(current);
  const normalizedCurrent = normalizeCodePageRunProfiles(current);
  const replacesExisting = normalizedCurrent.some((candidate) =>
    codePageRunProfilesShareSlot(candidate, profile),
  );
  if (
    !replacesExisting &&
    normalizedCurrent.length >= CODE_PAGE_RUN_PROFILE_LIMIT
  ) {
    return normalizedCurrent;
  }
  const next = [
    profile,
    ...normalizedCurrent.filter(
      (candidate) => !codePageRunProfilesShareSlot(candidate, profile),
    ),
  ];
  return normalizeCodePageRunProfiles(next);
}

export function removeCodePageRunProfile(
  current: readonly PersistedCodeRunProfile[],
  index: number,
): readonly PersistedCodeRunProfile[] {
  const normalized = normalizeCodePageRunProfiles(current);
  return Number.isSafeInteger(index) && index >= 0 && index < normalized.length
    ? normalized.filter((_, candidateIndex) => candidateIndex !== index)
    : normalized;
}

export function restoreCodePageRunProfiles(
  contentJson: unknown,
): readonly PersistedCodeRunProfile[] {
  if (
    !contentJson ||
    typeof contentJson !== "object" ||
    Array.isArray(contentJson)
  ) {
    return [];
  }
  try {
    return normalizeCodePageRunProfiles(
      (contentJson as Record<string, unknown>)[CODE_PAGE_RUN_PROFILES_KEY],
    );
  } catch {
    return [];
  }
}

function runTimelineUtf8Bytes(
  entries: readonly PersistedCodeRunTimelineEntry[],
): number {
  return new TextEncoder().encode(JSON.stringify(entries)).byteLength;
}

/**
 * Normalize newest-first timeline entries under one aggregate UTF-8 budget.
 * A malformed or oversized restored envelope fails closed as a whole so a
 * corrupt newest item can never silently promote an older entry to "last".
 * When a new trusted receipt is prepended, only the oldest valid entries are
 * evicted to remain within the count and byte budgets.
 */
export function normalizeCodePageRunTimeline(
  value: unknown,
): readonly PersistedCodeRunTimelineEntry[] {
  if (
    !Array.isArray(value) ||
    value.length > CODE_PAGE_RUN_TIMELINE_ENTRY_LIMIT ||
    (runProfilesUtf8Bytes(value) ?? Infinity) >
      CODE_PAGE_RUN_TIMELINE_UTF8_BYTE_LIMIT
  ) {
    return [];
  }
  const normalized: PersistedCodeRunTimelineEntry[] = [];
  for (const valueEntry of value) {
    const entry = normalizeCodeRunTimelineEntry(valueEntry);
    if (!entry) return [];
    normalized.push(entry);
  }
  return normalized;
}

export function createCodePageRunTimelineEntry(input: {
  readonly completedAt: number;
  readonly language: string;
  readonly providerId: string;
  readonly sourceSha256: string;
  readonly source: string;
  readonly selection: ExecutionSelection | null;
  readonly stdinSha256: string;
  readonly stdinBytes: number;
  readonly output: RunResult;
  readonly exitCode?: number | null;
}): PersistedCodeRunTimelineEntry | null {
  const navigation = createCodePageRunNavigation(input.source, input.selection);
  if (!navigation) return null;
  return normalizeCodeRunTimelineEntry({
    version: CODE_PAGE_RUN_TIMELINE_VERSION,
    completedAt: input.completedAt,
    language: input.language,
    providerId: input.providerId,
    sourceSha256: input.sourceSha256,
    outcome: input.output.error === undefined ? "success" : "error",
    exitCode: input.exitCode ?? null,
    output: runResultToStoredOutputs(input.output),
    input: {
      version: CODE_PAGE_RUN_INPUT_IDENTITY_VERSION,
      stdinSha256: input.stdinSha256,
      stdinBytes: input.stdinBytes,
    },
    navigation,
  });
}

export function appendCodePageRunTimelineEntry(
  current: readonly PersistedCodeRunTimelineEntry[],
  entry: PersistedCodeRunTimelineEntry,
): readonly PersistedCodeRunTimelineEntry[] {
  const normalizedEntry = normalizeCodeRunTimelineEntry(entry);
  const normalizedCurrent = normalizeCodePageRunTimeline(current);
  if (!normalizedEntry) return normalizedCurrent;
  const candidate = [normalizedEntry, ...normalizedCurrent].slice(
    0,
    CODE_PAGE_RUN_TIMELINE_ENTRY_LIMIT,
  );
  while (
    candidate.length > 0 &&
    runTimelineUtf8Bytes(candidate) > CODE_PAGE_RUN_TIMELINE_UTF8_BYTE_LIMIT
  ) {
    candidate.pop();
  }
  return candidate[0] === normalizedEntry ? candidate : normalizedCurrent;
}

export function restoreCodePageRunTimeline(
  contentJson: unknown,
): readonly PersistedCodeRunTimelineEntry[] {
  if (
    !contentJson ||
    typeof contentJson !== "object" ||
    Array.isArray(contentJson)
  ) {
    return [];
  }
  try {
    return normalizeCodePageRunTimeline(
      (contentJson as Record<string, unknown>)[CODE_PAGE_RUN_TIMELINE_KEY],
    );
  } catch {
    return [];
  }
}

export function restoreCodePageRunTimelineNavigation(
  entry: PersistedCodeRunTimelineEntry,
  source: string,
): RestoredCodeRunNavigation | null {
  return restoreCodePageRunNavigation(
    { [CODE_PAGE_LAST_RUN_NAVIGATION_KEY]: entry.navigation },
    source,
  );
}

/**
 * The run Clear Output took off the page, named by its completion time and
 * source hash. A page with no saved output shows its newest timeline run
 * instead; without this the cleared run came straight back (the save's own
 * echo restored it within a second, and so did every reopen). The run stays in
 * the output history; a newer run no longer matches and shows as usual.
 */
export interface PersistedCodeRunOutputCleared {
  readonly completedAt: number;
  readonly sourceSha256: string;
}

export function codePageRunOutputClearedMarker(
  entry: PersistedCodeRunTimelineEntry | null | undefined,
): PersistedCodeRunOutputCleared | null {
  return entry
    ? { completedAt: entry.completedAt, sourceSha256: entry.sourceSha256 }
    : null;
}

export function codePageRunOutputWasCleared(
  marker: PersistedCodeRunOutputCleared | null | undefined,
  entry: PersistedCodeRunTimelineEntry | null | undefined,
): boolean {
  return Boolean(
    marker &&
      entry &&
      marker.completedAt === entry.completedAt &&
      marker.sourceSha256 === entry.sourceSha256,
  );
}

export function restoreCodePageRunOutputCleared(
  contentJson: unknown,
): PersistedCodeRunOutputCleared | null {
  if (
    !contentJson ||
    typeof contentJson !== "object" ||
    Array.isArray(contentJson)
  ) {
    return null;
  }
  const raw = (contentJson as Record<string, unknown>)[
    CODE_PAGE_RUN_OUTPUT_CLEARED_KEY
  ];
  if (!raw || typeof raw !== "object" || Array.isArray(raw)) return null;
  const { completedAt, sourceSha256 } = raw as Record<string, unknown>;
  return typeof completedAt === "number" &&
    Number.isFinite(completedAt) &&
    typeof sourceSha256 === "string" &&
    /^[0-9a-f]{64}$/.test(sourceSha256)
    ? { completedAt, sourceSha256 }
    : null;
}

export function codePageRunTimelineEntryToRunResult(
  entry: PersistedCodeRunTimelineEntry,
): RunResult {
  return storedOutputsToRunResult(entry.output) ?? { logs: [] };
}

/**
 * Build the complete standalone-code save payload. Passing `null` (or an empty
 * run result) deliberately omits `lastRunOutput`, which clears a previously
 * persisted result because page saves replace the full content object.
 */
export function buildPersistedCodePageContent(
  type: string,
  language: string,
  content: string,
  output: RunResult | null | undefined,
  debugConfiguration?: CodeDebugConfiguration | null,
  sourceFileName?: string | null,
  lineEnding?: EolStyle | null,
  sourceEncoding?: CodeSourceEncoding | null,
  runNavigation?: PersistedCodeRunNavigation | null,
  workspacePath?: string | null,
  runTimeline?: readonly PersistedCodeRunTimelineEntry[] | null,
  debugTimeline?: readonly PersistedCodeDebugTimelineEntry[] | null,
  runProfiles?: readonly PersistedCodeRunProfile[] | null,
  runOutputCleared?: PersistedCodeRunOutputCleared | null,
): PersistedCodePageContent {
  const lastRunOutput = runResultToStoredOutputs(output);
  const candidateRunNavigation = lastRunOutput
    ? normalizeCodeRunNavigation(runNavigation)
    : null;
  const normalizedRunNavigation =
    candidateRunNavigation &&
    content.length <= CODE_PAGE_RUN_NAVIGATION_SOURCE_LIMIT &&
    candidateRunNavigation.sourceFingerprint ===
      codeSourceFingerprint(content)
      ? candidateRunNavigation
      : null;
  const normalizedDebugConfiguration = normalizeCodeDebugConfiguration(
    debugConfiguration,
  );
  const normalizedSourceFileName = normalizeCodeSourceFilename(sourceFileName);
  const normalizedLineEnding = normalizeCodePageLineEnding(lineEnding);
  const normalizedSourceEncoding =
    normalizeCodeSourceEncoding(sourceEncoding);
  const workspaceReference = persistedCodeWorkspaceReference(workspacePath);
  const normalizedRunTimeline = normalizeCodePageRunTimeline(runTimeline);
  const normalizedDebugTimeline = normalizeCodeDebugTimeline(debugTimeline);
  const normalizedRunProfiles = normalizeCodePageRunProfiles(runProfiles);
  return {
    type,
    language,
    content,
    ...(normalizedSourceFileName
      ? { [CODE_PAGE_SOURCE_FILENAME_KEY]: normalizedSourceFileName }
      : {}),
    ...(normalizedLineEnding === "CRLF"
      ? { [CODE_PAGE_LINE_ENDING_KEY]: normalizedLineEnding }
      : {}),
    ...(normalizedSourceEncoding !== "utf8"
      ? { [CODE_PAGE_SOURCE_ENCODING_KEY]: normalizedSourceEncoding }
      : {}),
    ...(lastRunOutput ? { lastRunOutput } : {}),
    ...(normalizedRunNavigation
      ? { [CODE_PAGE_LAST_RUN_NAVIGATION_KEY]: normalizedRunNavigation }
      : {}),
    ...(normalizedRunTimeline.length > 0
      ? { [CODE_PAGE_RUN_TIMELINE_KEY]: normalizedRunTimeline }
      : {}),
    ...(!lastRunOutput &&
    codePageRunOutputWasCleared(runOutputCleared, normalizedRunTimeline[0])
      ? {
          [CODE_PAGE_RUN_OUTPUT_CLEARED_KEY]:
            codePageRunOutputClearedMarker(normalizedRunTimeline[0]) ?? undefined,
        }
      : {}),
    ...(normalizedRunProfiles.length > 0
      ? { [CODE_PAGE_RUN_PROFILES_KEY]: normalizedRunProfiles }
      : {}),
    ...(normalizedDebugTimeline.length > 0
      ? { [CODE_PAGE_DEBUG_TIMELINE_KEY]: normalizedDebugTimeline }
      : {}),
    ...(normalizedDebugConfiguration.breakpoints.length > 0 ||
    (normalizedDebugConfiguration.watches?.length ?? 0) > 0 ||
    normalizedDebugConfiguration.pauseOnUncaughtExceptions === true
      ? {
          [CODE_PAGE_DEBUG_CONFIGURATION_KEY]: normalizedDebugConfiguration,
        }
      : {}),
    ...(workspaceReference
      ? { [CODE_WORKSPACE_REFERENCE_KEY]: workspaceReference }
      : {}),
  };
}

/** Restore source encoding metadata; legacy and malformed data stays UTF-8. */
export function restoreCodePageSourceEncoding(
  contentJson: unknown,
): CodeSourceEncoding {
  if (
    !contentJson ||
    typeof contentJson !== "object" ||
    Array.isArray(contentJson)
  ) {
    return "utf8";
  }
  return normalizeCodeSourceEncoding(
    (contentJson as Record<string, unknown>)[CODE_PAGE_SOURCE_ENCODING_KEY],
  );
}

/** Keep the persisted EOL enum total; absent or malformed legacy data is LF. */
export function normalizeCodePageLineEnding(value: unknown): EolStyle {
  return value === "CRLF" ? "CRLF" : "LF";
}

/** Restore the source EOL independently from CodeMirror's LF text model. */
export function restoreCodePageLineEnding(contentJson: unknown): EolStyle {
  if (
    !contentJson ||
    typeof contentJson !== "object" ||
    Array.isArray(contentJson)
  ) {
    return "LF";
  }
  return normalizeCodePageLineEnding(
    (contentJson as Record<string, unknown>)[CODE_PAGE_LINE_ENDING_KEY],
  );
}

/** Restore one bounded, path-free source basename from a saved code page. */
export function restoreCodePageSourceFilename(
  contentJson: unknown,
): string | null {
  if (
    !contentJson ||
    typeof contentJson !== "object" ||
    Array.isArray(contentJson)
  ) {
    return null;
  }
  return normalizeCodeSourceFilename(
    (contentJson as Record<string, unknown>)[CODE_PAGE_SOURCE_FILENAME_KEY],
  );
}

/**
 * Restore a saved standalone-code result defensively. Malformed, oversized, or
 * legacy content without `lastRunOutput` degrades to no output.
 */
export function restoreCodePageOutput(
  contentJson: unknown,
): RunResult | null {
  if (
    !contentJson ||
    typeof contentJson !== "object" ||
    Array.isArray(contentJson)
  ) {
    return null;
  }
  const raw = (contentJson as Record<string, unknown>)[
    CODE_PAGE_LAST_RUN_OUTPUT_KEY
  ];
  return storedOutputsToRunResult(normalizeStoredOutputs(raw));
}

/**
 * Restore traceback navigation only when the saved source fingerprint matches
 * the complete live buffer. Legacy, malformed, stale, or oversized metadata
 * remains safely inert while the output itself can still render.
 */
export function restoreCodePageRunNavigation(
  contentJson: unknown,
  source: string,
): RestoredCodeRunNavigation | null {
  if (
    source.length > CODE_PAGE_RUN_NAVIGATION_SOURCE_LIMIT ||
    !contentJson ||
    typeof contentJson !== "object" ||
    Array.isArray(contentJson)
  ) {
    return null;
  }
  const persisted = normalizeCodeRunNavigation(
    (contentJson as Record<string, unknown>)[
      CODE_PAGE_LAST_RUN_NAVIGATION_KEY
    ],
  );
  if (
    !persisted ||
    persisted.sourceFingerprint !== codeSourceFingerprint(source)
  ) {
    return null;
  }
  if (
    persisted.selection &&
    (persisted.selection.to > source.length ||
      persisted.selection.lineTo > codeSourceLineCount(source))
  ) {
    return null;
  }
  return {
    persisted,
    selection: persisted.selection
      ? {
          ...persisted.selection,
          // Navigation needs only the compact line mapping; never reconstruct
          // or persist the selected source body.
          source: "",
        }
      : null,
  };
}
