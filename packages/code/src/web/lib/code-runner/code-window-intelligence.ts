/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Cross-window code intelligence for standalone CodeWindow buffers.
 *
 * The registry already gives every open code window a stable id, live source,
 * language, label and write-back function. This adapter projects a bounded
 * snapshot of those buffers into the existing TypeScript workspace service and
 * converts compiler paths back into window ids. It stays framework-free so
 * navigation and multi-file rename can be tested without mounting React.
 */
import {
  identifierAt,
  isLegalIdentifier,
  type SignatureHelpInfo,
} from "./code-intelligence";
import { codeExportFilename } from "./code-file-import";
import { normalizeCodeWorkspacePath } from "./code-workspace";
import { tsDiagnosticsToCm } from "./lsp-diagnostics";
import {
  createWorkspaceProject,
  isWorkspaceLanguage,
  materializeWorkspaceCodeFixAction,
  workspaceCodeFixes,
  workspaceDefinition,
  workspaceImplementation,
  workspaceQuickInfo,
  workspaceCallHierarchy,
  workspaceRefactorEdits,
  workspaceRefactors,
  workspaceReferences,
  workspaceRenameLocations,
  workspaceSignatureHelp,
  workspaceTypeDefinition,
  type WorkspaceFile,
  type WorkspaceLocation,
  type WorkspaceQuickInfo,
} from "./workspace-intelligence";
import type { Diagnostic } from "@codemirror/lint";

/** Maximum source size accepted for one compiler-backed buffer. */
export const CODE_WINDOW_INTELLIGENCE_FILE_CHAR_LIMIT = 2 * 1024 * 1024;
/** Maximum aggregate source kept in one short-lived workspace project. */
export const CODE_WINDOW_INTELLIGENCE_WORKSPACE_CHAR_LIMIT = 8 * 1024 * 1024;
/** Maximum open buffers kept in one short-lived workspace project. */
export const CODE_WINDOW_INTELLIGENCE_BUFFER_LIMIT = 64;
/** Maximum rows shown for one find-references request. */
export const CODE_WINDOW_REFERENCE_LIMIT = 2_000;
/** Maximum compiler locations accepted for one atomic multi-window rename. */
export const CODE_WINDOW_RENAME_LOCATION_LIMIT = 50_000;
/** Maximum project-aware Quick Fix choices surfaced for one request. */
export const CODE_WINDOW_CODE_FIX_ACTION_LIMIT = 100;
/** Maximum compiler text edits accepted for one atomic Quick Fix. */
export const CODE_WINDOW_CODE_FIX_EDIT_LIMIT = 5_000;
/** Per-keystroke sibling scan budget for completion identifiers. */
export const CODE_WINDOW_IDENTIFIER_SCAN_CHAR_LIMIT = 1_000_000;
/** Maximum distinct sibling identifiers returned to autocomplete. */
export const CODE_WINDOW_IDENTIFIER_LIMIT = 5_000;
/** Maximum project diagnostics returned to one CodeMirror buffer. */
export const CODE_WINDOW_DIAGNOSTIC_LIMIT = 200;

export interface CodeWindowBufferText {
  bufferId: string;
  label: string;
  language: string;
  source: string;
}

export interface CodeWindowIntelligenceLocation {
  bufferId: string;
  from: number;
  to: number;
  /** Exact source snapshot used to calculate this range. */
  expectedSource: string;
}

export interface CodeWindowReference extends CodeWindowIntelligenceLocation {
  label: string;
  line: number;
  column: number;
  preview: string;
}

export interface CodeWindowReferenceResult {
  refs: CodeWindowReference[];
  activeIndex: number;
  truncated: boolean;
}

export interface CodeWindowHoverResult extends WorkspaceQuickInfo {
  /** Active-buffer snapshot used by the asynchronous compiler query. */
  expectedSource: string;
  /** Every bounded source snapshot that influenced the hover result. */
  expectedSources: Array<{ bufferId: string; source: string }>;
}

export interface CodeWindowSignatureHelpResult extends SignatureHelpInfo {
  /** Every bounded source snapshot that influenced the parameter hint. */
  expectedSources: Array<{ bufferId: string; source: string }>;
}

export interface CodeWindowRenameEdit {
  bufferId: string;
  before: string;
  after: string;
}

export interface CodeWindowCodeFixAction {
  description: string;
  edits: CodeWindowRenameEdit[];
  /** Every bounded source snapshot that influenced the compiler action. */
  expectedSources: Array<{ bufferId: string; source: string }>;
}

export interface CodeWindowDiagnosticResult {
  diagnostics: Diagnostic[];
  /** Every source snapshot that influenced the compiler result. */
  expectedSources: Array<{ bufferId: string; source: string }>;
}

interface WorkspaceSnapshot {
  active: CodeWindowBufferText;
  activePath: string;
  buffersById: Map<string, CodeWindowBufferText>;
  idByPath: Map<string, string>;
  files: WorkspaceFile[];
}

function splitExtension(filename: string): { stem: string; extension: string } {
  const dot = filename.lastIndexOf(".");
  if (dot <= 0) return { stem: filename, extension: "" };
  return {
    stem: filename.slice(0, dot),
    extension: filename.slice(dot),
  };
}

function uniqueWorkspacePath(
  buffer: CodeWindowBufferText,
  usedPaths: Set<string>,
): string {
  const normalizedLabel = normalizeCodeWorkspacePath(buffer.label);
  const filename =
    normalizedLabel ??
    codeExportFilename(
      buffer.label,
      buffer.label || buffer.bufferId,
      buffer.language,
    );
  const basePath = `/${filename}`;
  if (!usedPaths.has(basePath)) {
    usedPaths.add(basePath);
    return basePath;
  }
  const slash = filename.lastIndexOf("/");
  const directory = slash < 0 ? "" : filename.slice(0, slash + 1);
  const basename = slash < 0 ? filename : filename.slice(slash + 1);
  const { stem, extension } = splitExtension(basename);
  let suffix = 2;
  let candidate = `/${directory}${stem}~${suffix}${extension}`;
  while (usedPaths.has(candidate)) {
    suffix += 1;
    candidate = `/${directory}${stem}~${suffix}${extension}`;
  }
  usedPaths.add(candidate);
  return candidate;
}

/**
 * Build a bounded, active-first snapshot. Oversized peers are skipped; an
 * oversized active file disables cross-window intelligence instead of feeding
 * an unbounded compiler project on the UI thread.
 */
function buildWorkspaceSnapshot(
  buffers: readonly CodeWindowBufferText[],
  activeBufferId: string,
): WorkspaceSnapshot | null {
  const active = buffers.find((buffer) => buffer.bufferId === activeBufferId);
  if (
    !active ||
    !isWorkspaceLanguage(active.language) ||
    active.source.length > CODE_WINDOW_INTELLIGENCE_FILE_CHAR_LIMIT
  ) {
    return null;
  }

  const ordered = [
    active,
    ...buffers.filter((buffer) => buffer.bufferId !== activeBufferId),
  ];
  const usedPaths = new Set<string>();
  const buffersById = new Map<string, CodeWindowBufferText>();
  const idByPath = new Map<string, string>();
  const files: WorkspaceFile[] = [];
  let totalChars = 0;
  let activePath = "";

  for (const buffer of ordered) {
    if (files.length >= CODE_WINDOW_INTELLIGENCE_BUFFER_LIMIT) break;
    if (!isWorkspaceLanguage(buffer.language)) continue;
    if (buffer.source.length > CODE_WINDOW_INTELLIGENCE_FILE_CHAR_LIMIT) {
      continue;
    }
    if (
      files.length > 0 &&
      totalChars + buffer.source.length >
        CODE_WINDOW_INTELLIGENCE_WORKSPACE_CHAR_LIMIT
    ) {
      continue;
    }
    const path = uniqueWorkspacePath(buffer, usedPaths);
    files.push({ path, content: buffer.source });
    buffersById.set(buffer.bufferId, buffer);
    idByPath.set(path, buffer.bufferId);
    totalChars += buffer.source.length;
    if (buffer.bufferId === activeBufferId) activePath = path;
  }

  if (!activePath) return null;
  return { active, activePath, buffersById, idByPath, files };
}

async function createSnapshotProject(
  buffers: readonly CodeWindowBufferText[],
  activeBufferId: string,
) {
  const snapshot = buildWorkspaceSnapshot(buffers, activeBufferId);
  if (!snapshot) return null;
  const project = await createWorkspaceProject(snapshot.files);
  return project ? { snapshot, project } : null;
}

function firstForeignLocation(
  snapshot: WorkspaceSnapshot,
  activeBufferId: string,
  locations: readonly WorkspaceLocation[],
): CodeWindowIntelligenceLocation | null {
  for (const location of locations) {
    const bufferId = snapshot.idByPath.get(location.path);
    if (!bufferId || bufferId === activeBufferId) continue;
    const buffer = snapshot.buffersById.get(bufferId);
    if (!buffer) continue;
    return {
      bufferId,
      from: location.from,
      to: location.to,
      expectedSource: buffer.source,
    };
  }
  return null;
}

function expectedSnapshotSources(
  snapshot: WorkspaceSnapshot,
): Array<{ bufferId: string; source: string }> {
  return [...snapshot.buffersById.values()].map((buffer) => ({
    bufferId: buffer.bufferId,
    source: buffer.source,
  }));
}

/** Resolve a definition only when it lives in another open code window. */
export async function resolveCodeWindowDefinition(
  buffers: readonly CodeWindowBufferText[],
  activeBufferId: string,
  offset: number,
): Promise<CodeWindowIntelligenceLocation | null> {
  const context = await createSnapshotProject(buffers, activeBufferId);
  if (!context) return null;
  const { snapshot, project } = context;
  const definitions = workspaceDefinition(project, snapshot.activePath, offset);
  return firstForeignLocation(snapshot, activeBufferId, definitions);
}

/** Resolve a type declaration only when it lives in another code window/file. */
export async function resolveCodeWindowTypeDefinition(
  buffers: readonly CodeWindowBufferText[],
  activeBufferId: string,
  offset: number,
): Promise<CodeWindowIntelligenceLocation | null> {
  const context = await createSnapshotProject(buffers, activeBufferId);
  if (!context) return null;
  const { snapshot, project } = context;
  const definitions = workspaceTypeDefinition(
    project,
    snapshot.activePath,
    offset,
  );
  return firstForeignLocation(snapshot, activeBufferId, definitions);
}

/** Resolve a concrete implementation only when it lives in another code file. */
export async function resolveCodeWindowImplementation(
  buffers: readonly CodeWindowBufferText[],
  activeBufferId: string,
  offset: number,
): Promise<CodeWindowIntelligenceLocation | null> {
  const context = await createSnapshotProject(buffers, activeBufferId);
  if (!context) return null;
  const { snapshot, project } = context;
  const implementations = workspaceImplementation(
    project,
    snapshot.activePath,
    offset,
  );
  return firstForeignLocation(
    snapshot,
    activeBufferId,
    implementations,
  );
}

function lineStarts(source: string): number[] {
  const starts = [0];
  for (let index = 0; index < source.length; index += 1) {
    if (source.charCodeAt(index) === 10) starts.push(index + 1);
  }
  return starts;
}

function lineIndexAt(starts: readonly number[], offset: number): number {
  let low = 0;
  let high = starts.length;
  while (low < high) {
    const middle = low + Math.floor((high - low) / 2);
    if (starts[middle] <= offset) low = middle + 1;
    else high = middle;
  }
  return Math.max(0, low - 1);
}

function locateReference(
  source: string,
  starts: readonly number[],
  offset: number,
): { line: number; column: number; preview: string } {
  const safeOffset = Math.max(0, Math.min(offset, source.length));
  const lineIndex = lineIndexAt(starts, safeOffset);
  const lineStart = starts[lineIndex] ?? 0;
  const rawLineEnd = source.indexOf("\n", safeOffset);
  const lineEnd = rawLineEnd < 0 ? source.length : rawLineEnd;
  const previewEnd = Math.min(lineEnd, lineStart + 300);
  const preview =
    source.slice(lineStart, previewEnd).trim() +
    (previewEnd < lineEnd ? "…" : "");
  return {
    line: lineIndex + 1,
    column: safeOffset - lineStart,
    preview,
  };
}

/**
 * Resolve a bounded call hierarchy across every included open code window.
 *
 * The rows use the SAME shape as find-all-references, so the editor renders and
 * navigates them through the path it already has: one direction of the
 * hierarchy is a list of places, and each row is "this call, in this file, on
 * this line". `direction` picks whether those places are the callers of the
 * symbol or the functions it calls.
 */
export async function resolveCodeWindowCallHierarchy(
  buffers: readonly CodeWindowBufferText[],
  activeBufferId: string,
  offset: number,
  direction: "incoming" | "outgoing",
): Promise<CodeWindowReferenceResult | null> {
  const context = await createSnapshotProject(buffers, activeBufferId);
  if (!context) return null;
  const { snapshot, project } = context;
  const hierarchy = workspaceCallHierarchy(project, snapshot.activePath, offset);
  if (!hierarchy) return null;
  const calls =
    direction === "incoming" ? hierarchy.incoming : hierarchy.outgoing;
  if (calls.length === 0) return null;

  const startsById = new Map<string, number[]>();
  const refs: CodeWindowReference[] = [];
  let seen = 0;
  for (const call of calls) {
    for (const site of call.sites) {
      seen += 1;
      if (refs.length >= CODE_WINDOW_REFERENCE_LIMIT) break;
      const bufferId = snapshot.idByPath.get(site.path);
      if (!bufferId) continue;
      const buffer = snapshot.buffersById.get(bufferId);
      if (!buffer) continue;
      let starts = startsById.get(bufferId);
      if (!starts) {
        starts = lineStarts(buffer.source);
        startsById.set(bufferId, starts);
      }
      const located = locateReference(buffer.source, starts, site.from);
      refs.push({
        bufferId,
        label: `${call.name} — ${buffer.label}`,
        from: site.from,
        to: site.to,
        expectedSource: buffer.source,
        ...located,
      });
    }
  }
  if (refs.length === 0) return null;
  return {
    refs,
    activeIndex: 0,
    truncated: hierarchy.truncated || seen > refs.length,
  };
}

/** Resolve a bounded reference list across every included open code window. */
export async function resolveCodeWindowReferences(
  buffers: readonly CodeWindowBufferText[],
  activeBufferId: string,
  offset: number,
): Promise<CodeWindowReferenceResult | null> {
  const context = await createSnapshotProject(buffers, activeBufferId);
  if (!context) return null;
  const { snapshot, project } = context;
  const raw = workspaceReferences(project, snapshot.activePath, offset);
  if (raw.length === 0) return null;

  const startsById = new Map<string, number[]>();
  const refs: CodeWindowReference[] = [];
  let activeIndex = 0;
  for (const location of raw) {
    if (refs.length >= CODE_WINDOW_REFERENCE_LIMIT) break;
    const bufferId = snapshot.idByPath.get(location.path);
    if (!bufferId) continue;
    const buffer = snapshot.buffersById.get(bufferId);
    if (!buffer) continue;
    let starts = startsById.get(bufferId);
    if (!starts) {
      starts = lineStarts(buffer.source);
      startsById.set(bufferId, starts);
    }
    const located = locateReference(buffer.source, starts, location.from);
    if (
      bufferId === activeBufferId &&
      offset >= location.from &&
      offset <= location.to
    ) {
      activeIndex = refs.length;
    }
    refs.push({
      bufferId,
      label: buffer.label,
      from: location.from,
      to: location.to,
      expectedSource: buffer.source,
      ...located,
    });
  }
  if (refs.length === 0) return null;
  return {
    refs,
    activeIndex,
    truncated: raw.length > refs.length,
  };
}

/** Resolve project-aware hover information for the active open code window. */
export async function resolveCodeWindowHover(
  buffers: readonly CodeWindowBufferText[],
  activeBufferId: string,
  offset: number,
): Promise<CodeWindowHoverResult | null> {
  const context = await createSnapshotProject(buffers, activeBufferId);
  if (!context) return null;
  const { snapshot, project } = context;
  const info = workspaceQuickInfo(project, snapshot.activePath, offset);
  return info
    ? {
        ...info,
        expectedSource: snapshot.active.source,
        expectedSources: expectedSnapshotSources(snapshot),
      }
    : null;
}

/** Resolve project-aware parameter hints for the active code window. */
export async function resolveCodeWindowSignatureHelp(
  buffers: readonly CodeWindowBufferText[],
  activeBufferId: string,
  offset: number,
): Promise<CodeWindowSignatureHelpResult | null> {
  const context = await createSnapshotProject(buffers, activeBufferId);
  if (!context) return null;
  const { snapshot, project } = context;
  const info = workspaceSignatureHelp(
    project,
    snapshot.activePath,
    offset,
  );
  return info.hasContent
    ? {
        ...info,
        expectedSources: expectedSnapshotSources(snapshot),
      }
    : null;
}

/**
 * Resolve syntactic + semantic diagnostics for the active file against the
 * whole open-window project. A single included file returns `null` so the
 * editor keeps its existing single-buffer linter and avoids duplicate work.
 */
export async function resolveCodeWindowDiagnostics(
  buffers: readonly CodeWindowBufferText[],
  activeBufferId: string,
): Promise<CodeWindowDiagnosticResult | null> {
  const context = await createSnapshotProject(buffers, activeBufferId);
  if (!context || context.snapshot.files.length < 2) return null;
  const { snapshot, project } = context;
  try {
    const raw = [
      ...project.service.getSyntacticDiagnostics(snapshot.activePath),
      ...project.service.getSemanticDiagnostics(snapshot.activePath),
    ];
    const flattened = raw.map((diagnostic) => ({
      start: diagnostic.start,
      length: diagnostic.length,
      category: diagnostic.category,
      code: diagnostic.code,
      messageText: project.ts.flattenDiagnosticMessageText(
        diagnostic.messageText,
        "\n",
      ),
    }));
    return {
      diagnostics: tsDiagnosticsToCm(flattened, snapshot.active.source).slice(
        0,
        CODE_WINDOW_DIAGNOSTIC_LIMIT,
      ),
      expectedSources: expectedSnapshotSources(snapshot),
    };
  } catch {
    return null;
  }
}

/** One refactoring offered for the current selection, plus its compiler ids. */
export interface CodeWindowRefactorAction {
  refactorName: string;
  actionName: string;
  description: string;
}

/**
 * List the refactorings applicable to `[from, to)` across the open windows.
 *
 * Listing is cheap and materializing is not, so this only names the offers;
 * {@link materializeCodeWindowRefactor} produces edits for the ONE the user
 * picks, and does so against a freshly built project so a buffer edited while
 * the menu was open cannot be rewritten with stale spans.
 */
export async function resolveCodeWindowRefactors(
  buffers: readonly CodeWindowBufferText[],
  activeBufferId: string,
  from: number,
  to: number,
): Promise<CodeWindowRefactorAction[] | null> {
  const context = await createSnapshotProject(buffers, activeBufferId);
  if (!context) return null;
  const actions = workspaceRefactors(
    context.project,
    context.snapshot.activePath,
    from,
    to,
  );
  return actions.length > 0 ? actions : null;
}

/**
 * Materialize one chosen refactoring into the same validated action shape a
 * project Quick Fix produces, so the host writes it through the identical
 * atomic snapshot-checked path.
 */
export async function materializeCodeWindowRefactor(
  buffers: readonly CodeWindowBufferText[],
  activeBufferId: string,
  from: number,
  to: number,
  refactorName: string,
  actionName: string,
): Promise<CodeWindowCodeFixAction | null> {
  const context = await createSnapshotProject(buffers, activeBufferId);
  if (!context) return null;
  const { snapshot, project } = context;
  const action = workspaceRefactorEdits(
    project,
    snapshot.activePath,
    from,
    to,
    refactorName,
    actionName,
  );
  if (!action) return null;
  const sourcesByPath = new Map(
    snapshot.files.map((file) => [file.path, file.content]),
  );
  const expectedSources = snapshot.files.flatMap((file) => {
    const bufferId = snapshot.idByPath.get(file.path);
    return bufferId ? [{ bufferId, source: file.content }] : [];
  });
  const materialized = materializeWorkspaceCodeFixAction(
    action,
    sourcesByPath,
    CODE_WINDOW_CODE_FIX_EDIT_LIMIT,
  );
  if (!materialized) return null;
  const edits = materialized.flatMap((edit) => {
    const bufferId = snapshot.idByPath.get(edit.path);
    return bufferId
      ? [{ bufferId, before: edit.before, after: edit.after }]
      : [];
  });
  if (edits.length === 0) return null;
  return { description: action.description, edits, expectedSources };
}

/**
 * Calculate project-aware Quick Fixes that touch at least one OTHER open code
 * window. Every action carries complete before/after snapshots so the host can
 * preflight all targets and apply it all-or-nothing.
 */
export async function resolveCodeWindowCodeFixes(
  buffers: readonly CodeWindowBufferText[],
  activeBufferId: string,
  from: number,
  to: number,
): Promise<CodeWindowCodeFixAction[] | null> {
  // The normal single-window path already owns same-file fixes. Avoid building
  // a second language-service project unless at least one bounded peer could
  // actually receive a project edit.
  const hasEligiblePeer = buffers.some(
    (buffer) =>
      buffer.bufferId !== activeBufferId &&
      isWorkspaceLanguage(buffer.language) &&
      buffer.source.length <= CODE_WINDOW_INTELLIGENCE_FILE_CHAR_LIMIT,
  );
  if (!hasEligiblePeer) return null;
  const context = await createSnapshotProject(buffers, activeBufferId);
  if (!context || context.snapshot.files.length < 2) return null;
  const { snapshot, project } = context;
  const sourcesByPath = new Map(
    snapshot.files.map((file) => [file.path, file.content]),
  );
  const expectedSources = snapshot.files.flatMap((file) => {
    const bufferId = snapshot.idByPath.get(file.path);
    return bufferId ? [{ bufferId, source: file.content }] : [];
  });
  const projectActions = workspaceCodeFixes(
    project,
    snapshot.activePath,
    from,
    to,
  );
  const actions: CodeWindowCodeFixAction[] = [];
  for (const action of projectActions.slice(
    0,
    CODE_WINDOW_CODE_FIX_ACTION_LIMIT,
  )) {
    const materialized = materializeWorkspaceCodeFixAction(
      action,
      sourcesByPath,
      CODE_WINDOW_CODE_FIX_EDIT_LIMIT,
    );
    if (!materialized) continue;
    const edits: CodeWindowRenameEdit[] = [];
    let hasForeignEdit = false;
    for (const edit of materialized) {
      const bufferId = snapshot.idByPath.get(edit.path);
      if (!bufferId) continue;
      edits.push({
        bufferId,
        before: edit.before,
        after: edit.after,
      });
      if (bufferId !== activeBufferId) hasForeignEdit = true;
    }
    if (!hasForeignEdit || edits.length !== materialized.length) continue;
    actions.push({
      description: action.description,
      edits,
      expectedSources,
    });
  }
  return actions.length > 0 ? actions : null;
}

/**
 * Calculate an all-or-nothing multi-window rename. The caller must verify every
 * `before` snapshot still matches before applying any `after` value.
 */
export async function resolveCodeWindowRename(
  buffers: readonly CodeWindowBufferText[],
  activeBufferId: string,
  offset: number,
  newName: string,
): Promise<CodeWindowRenameEdit[] | null> {
  if (!isLegalIdentifier(newName)) return null;
  const context = await createSnapshotProject(buffers, activeBufferId);
  if (!context) return null;
  const { snapshot, project } = context;
  if (identifierAt(snapshot.active.source, offset) === newName) return null;

  const locations = workspaceRenameLocations(
    project,
    snapshot.activePath,
    offset,
  );
  if (
    !locations ||
    locations.length === 0 ||
    locations.length > CODE_WINDOW_RENAME_LOCATION_LIMIT
  ) {
    return null;
  }

  const spansById = new Map<
    string,
    Array<{
      from: number;
      to: number;
      prefix: string;
      suffix: string;
    }>
  >();
  const seen = new Set<string>();
  for (const location of locations) {
    const bufferId = snapshot.idByPath.get(location.path);
    const buffer = bufferId ? snapshot.buffersById.get(bufferId) : undefined;
    if (
      !bufferId ||
      !buffer ||
      location.from < 0 ||
      location.to < location.from ||
      location.to > buffer.source.length
    ) {
      continue;
    }
    const key = `${bufferId}:${location.from}:${location.to}`;
    if (seen.has(key)) continue;
    seen.add(key);
    const spans = spansById.get(bufferId) ?? [];
    spans.push({
      from: location.from,
      to: location.to,
      prefix: location.prefixText ?? "",
      suffix: location.suffixText ?? "",
    });
    spansById.set(bufferId, spans);
  }

  const edits: CodeWindowRenameEdit[] = [];
  for (const [bufferId, spans] of spansById) {
    const buffer = snapshot.buffersById.get(bufferId);
    if (!buffer) continue;
    spans.sort((left, right) => right.from - left.from);
    let after = buffer.source;
    for (const span of spans) {
      after =
        after.slice(0, span.from) +
        span.prefix +
        newName +
        span.suffix +
        after.slice(span.to);
    }
    if (after !== buffer.source) {
      edits.push({ bufferId, before: buffer.source, after });
    }
  }
  if (edits.length === 0) return null;
  if (edits.every((edit) => edit.bufferId === activeBufferId)) return null;
  return edits;
}

const IDENTIFIER_RE = /[\p{ID_Start}_$][\p{ID_Continue}$]*/gu;

/**
 * Harvest bounded sibling identifiers for autocomplete without constructing a
 * compiler project per keystroke.
 */
export function harvestOpenCodeIdentifiers(
  buffers: readonly CodeWindowBufferText[],
  activeBufferId: string,
): string[] {
  const identifiers = new Set<string>();
  let remainingChars = CODE_WINDOW_IDENTIFIER_SCAN_CHAR_LIMIT;
  for (const buffer of buffers) {
    if (
      buffer.bufferId === activeBufferId ||
      !isWorkspaceLanguage(buffer.language) ||
      remainingChars <= 0
    ) {
      continue;
    }
    const source = buffer.source.slice(0, remainingChars);
    remainingChars -= source.length;
    IDENTIFIER_RE.lastIndex = 0;
    let match: RegExpExecArray | null;
    while ((match = IDENTIFIER_RE.exec(source))) {
      identifiers.add(match[0]);
      if (identifiers.size >= CODE_WINDOW_IDENTIFIER_LIMIT) {
        return [...identifiers];
      }
    }
  }
  return [...identifiers];
}
