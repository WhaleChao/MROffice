/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * CROSS-FILE (project-wide) code intelligence for the MRNote code editor — the
 * foundation of LSP-grade go-to-definition / find-all-references that spans
 * MULTIPLE files, not just the active buffer.
 *
 * `code-intelligence.ts` drives the real TypeScript `LanguageService` over a
 * SINGLE virtual file (the live buffer) and explicitly notes "cross-file
 * definitions are N/A here". This module lifts that limitation: it builds the
 * same real `LanguageService` over a WHOLE workspace of files, so a reference in
 * `main.ts` to `import { foo } from "./util"` resolves to `foo`'s declaration in
 * `util.ts`, and "find all references" returns every use across every file.
 *
 * It is the genuine cross-file engine (the actual tsserver/VSCode engine, in
 * memory, no language-server sidecar, no new dependency — `typescript` is already
 * installed and pulled in lazily). Everything here is PURE / framework-agnostic
 * and unit-tested against the REAL compiler; the editor wiring (open the target
 * file's window and jump there) consumes these functions.
 *
 * Scope: the JS/TS family (js/jsx/ts/tsx). Other languages have no TS service and
 * are simply absent from a workspace project.
 */
import type * as TsModule from "typescript";
import { localizeTypeScript } from "./typescript-locale";
import { TYPESCRIPT_AMBIENT_LIB } from "./typescript-ambient-lib";

import {
  completionsAt,
  inlayHints,
  signatureHelpAt,
  type BufferEdit,
  type CodeLens,
  type InlayHintItem,
  type SemanticCompletionEntry,
  type SignatureHelpInfo,
} from "./code-intelligence";
import { normalizeLanguage } from "./language-metadata";

/** A workspace file fed to the cross-file project. */
export interface WorkspaceFile {
  /** A stable path used for imports + result attribution, e.g. "/main.ts".
   *  Must start with "/" and carry a real .js/.jsx/.ts/.tsx extension. */
  path: string;
  /** The file's full source text. */
  content: string;
}

/** A cross-file location result: which file, and the char span within it. */
export interface WorkspaceLocation {
  /** The `path` of the file the location lives in. */
  path: string;
  /** 0-based char offset of the span start (inclusive). */
  from: number;
  /** 0-based char offset just past the span end (exclusive). */
  to: number;
}

/** One file's edits within a project-wide TypeScript code action. */
export interface WorkspaceCodeFixChange {
  path: string;
  edits: BufferEdit[];
}

/** One project-aware Quick Fix, potentially touching multiple workspace files. */
export interface WorkspaceCodeFixAction {
  description: string;
  changes: WorkspaceCodeFixChange[];
}

/** One fully materialized file rewrite for a project-aware code action. */
export interface WorkspaceCodeFixFileEdit {
  path: string;
  before: string;
  after: string;
}

/** One compiler-backed completion item returned from an authorized project. */
export type WorkspaceCompletionEntry = SemanticCompletionEntry;

/** Project completion range and bounded semantic entries. */
export interface WorkspaceCompletionResult {
  /** Start of the identifier/member fragment to replace. */
  from: number;
  /** End of the replacement range (normally the request caret). */
  to: number;
  /** True when TypeScript resolved a member-access completion. */
  isMember: boolean;
  /** De-duplicated semantic entries, capped by the caller's budget. */
  entries: WorkspaceCompletionEntry[];
}

/** Compiler symbols shown as Code Lens blocks for one workspace file. */
export const WORKSPACE_CODE_LENS_LIMIT = 200;
/** Visible-range hint ceiling before data crosses into CodeMirror. */
export const WORKSPACE_INLAY_HINT_LIMIT = 500;

/**
 * Validate, merge and materialize one compiler code action against exact source
 * snapshots. Shared by standalone files and multi-buffer host cells so both surfaces use
 * the same overlap, bounds and edit-count safety rules.
 */
export function materializeWorkspaceCodeFixAction(
  action: WorkspaceCodeFixAction,
  sources: ReadonlyMap<string, string>,
  maxEdits: number,
): WorkspaceCodeFixFileEdit[] | null {
  if (!Number.isSafeInteger(maxEdits) || maxEdits < 1) return null;
  const editsByPath = new Map<string, BufferEdit[]>();
  let editCount = 0;
  for (const change of action.changes) {
    const source = sources.get(change.path);
    if (source === undefined) return null;
    const edits = editsByPath.get(change.path) ?? [];
    for (const edit of change.edits) {
      editCount += 1;
      if (
        editCount > maxEdits ||
        edit.from < 0 ||
        edit.to < edit.from ||
        edit.to > source.length
      ) {
        return null;
      }
      edits.push(edit);
    }
    editsByPath.set(change.path, edits);
  }

  const materialized: WorkspaceCodeFixFileEdit[] = [];
  for (const [path, edits] of editsByPath) {
    const before = sources.get(path);
    if (before === undefined) return null;
    edits.sort(
      (left, right) =>
        left.from - right.from || left.to - right.to,
    );
    let previousTo = -1;
    for (const edit of edits) {
      if (edit.from < previousTo) return null;
      previousTo = edit.to;
    }
    let after = before;
    for (let index = edits.length - 1; index >= 0; index -= 1) {
      const edit = edits[index];
      after =
        after.slice(0, edit.from) +
        edit.insert +
        after.slice(edit.to);
    }
    if (after !== before) materialized.push({ path, before, after });
  }
  return materialized.length > 0 ? materialized : null;
}

/**
 * A rename location: a {@link WorkspaceLocation} plus the TS-provided
 * `prefixText`/`suffixText` that MUST bracket the new name to rename safely —
 * e.g. an object-shorthand `{ port }` renamed to `PORT` needs prefix `"port: "`
 * so it becomes `{ port: PORT }` instead of silently rewriting the property key.
 */
export interface WorkspaceRenameLocation extends WorkspaceLocation {
  prefixText?: string;
  suffixText?: string;
}

const LIB_FILE_NAME = "/__mrnote_ws_lib__.d.ts";

/** The ambient built-ins the single-buffer services use too, so Problems agree. */
const WS_LIB = TYPESCRIPT_AMBIENT_LIB;

function scriptKindFor(ts: typeof TsModule, path: string): TsModule.ScriptKind {
  if (path.endsWith(".tsx")) return ts.ScriptKind.TSX;
  if (path.endsWith(".ts")) return ts.ScriptKind.TS;
  if (path.endsWith(".jsx")) return ts.ScriptKind.JSX;
  return ts.ScriptKind.JS;
}

/** Whether a language id belongs in a cross-file TS project. */
export function isWorkspaceLanguage(language: string): boolean {
  const lang = normalizeLanguage(language);
  return (
    lang === "javascript" || lang === "jsx" || lang === "typescript" || lang === "tsx"
  );
}

/**
 * A live cross-file project: a real TS `LanguageService` over the given files,
 * plus the query helpers. Files are normalized so a path missing a leading "/"
 * still works. `update(path, content)` re-versions a single file in place (for
 * the live editor buffer) without rebuilding the whole project.
 */
export interface WorkspaceProject {
  service: TsModule.LanguageService;
  ts: typeof TsModule;
  /** The normalized paths in the project (excludes the synthetic lib). */
  paths: string[];
  update(path: string, content: string): void;
}

function normalizePath(path: string): string {
  return path.startsWith("/") ? path : `/${path}`;
}

/**
 * Build a cross-file project from a set of files. Returns null if `typescript`
 * cannot be loaded (so callers cleanly skip cross-file features). PURE aside
 * from the lazy `import("typescript")`.
 */
export async function createWorkspaceProject(
  files: WorkspaceFile[],
): Promise<WorkspaceProject | null> {
  let ts: typeof TsModule;
  try {
    ts = (await import("typescript")).default ?? (await import("typescript"));
    await localizeTypeScript(ts);
  } catch {
    return null;
  }
  const contents = new Map<string, string>();
  const versions = new Map<string, number>();
  contents.set(LIB_FILE_NAME, WS_LIB);
  versions.set(LIB_FILE_NAME, 0);
  for (const f of files) {
    const p = normalizePath(f.path);
    contents.set(p, f.content);
    versions.set(p, 0);
  }
  const projectPaths = (): string[] =>
    [...contents.keys()].filter((p) => p !== LIB_FILE_NAME);

  const host: TsModule.LanguageServiceHost = {
    getScriptFileNames: () => [...contents.keys()],
    getScriptVersion: (fileName) => String(versions.get(fileName) ?? 0),
    getScriptSnapshot: (fileName) => {
      const text = contents.get(fileName);
      return text === undefined ? undefined : ts.ScriptSnapshot.fromString(text);
    },
    getCurrentDirectory: () => "/",
    getCompilationSettings: () => ({
      noLib: true,
      allowJs: true,
      checkJs: false,
      target: ts.ScriptTarget.ESNext,
      module: ts.ModuleKind.ESNext,
      // Every code page runs on its own: its top-level names are its own, and
      // another page's are reachable only through an import. Without this a
      // page with no import/export is a global script, so two pages that each
      // declare `main` were told "Duplicate function implementation", and a
      // call to a function only another page declares passed the checker and
      // then threw at run time.
      moduleDetection: ts.ModuleDetectionKind.Force,
      moduleResolution: ts.ModuleResolutionKind.Bundler,
      jsx: ts.JsxEmit.Preserve,
      allowNonTsExtensions: true,
      esModuleInterop: true,
      isolatedModules: false,
    }),
    getDefaultLibFileName: () => LIB_FILE_NAME,
    fileExists: (fileName) => contents.has(fileName),
    readFile: (fileName) => contents.get(fileName),
    directoryExists: () => true,
    getDirectories: () => [],
    getScriptKind: (fileName) => scriptKindFor(ts, fileName),
    getNewLine: () => "\n",
    useCaseSensitiveFileNames: () => true,
  };
  const service = ts.createLanguageService(host, ts.createDocumentRegistry());
  return {
    service,
    ts,
    paths: projectPaths(),
    update(path: string, content: string) {
      const p = normalizePath(path);
      contents.set(p, content);
      versions.set(p, (versions.get(p) ?? 0) + 1);
    },
  };
}

/**
 * Cross-file go-to-definition: every declaration of the symbol at `offset` in
 * `path`, ACROSS files. A definition that resolves into the synthetic lib (a
 * built-in) is dropped — only real workspace declarations are returned. Returns
 * `[]` when nothing resolves (Excel of an unknown name).
 */
export function workspaceDefinition(
  project: WorkspaceProject,
  path: string,
  offset: number,
): WorkspaceLocation[] {
  const p = normalizePath(path);
  const defs = project.service.getDefinitionAtPosition(p, offset);
  if (!defs) return [];
  const out: WorkspaceLocation[] = [];
  for (const d of defs) {
    if (d.fileName === LIB_FILE_NAME) continue;
    out.push({
      path: d.fileName,
      from: d.textSpan.start,
      to: d.textSpan.start + d.textSpan.length,
    });
  }
  return out;
}

/**
 * Cross-file go-to-type-definition: declarations of the TYPE of the symbol at
 * `offset`, including imported interfaces/classes/types in another authorized
 * workspace file. Synthetic-lib locations are omitted like normal definitions.
 */
export function workspaceTypeDefinition(
  project: WorkspaceProject,
  path: string,
  offset: number,
): WorkspaceLocation[] {
  const p = normalizePath(path);
  const definitions = project.service.getTypeDefinitionAtPosition(p, offset);
  if (!definitions) return [];
  const out: WorkspaceLocation[] = [];
  for (const definition of definitions) {
    if (definition.fileName === LIB_FILE_NAME) continue;
    out.push({
      path: definition.fileName,
      from: definition.textSpan.start,
      to: definition.textSpan.start + definition.textSpan.length,
    });
  }
  return out;
}

/**
 * Cross-file go-to-implementation: concrete implementations of the symbol at
 * `offset`, including classes/members declared in another authorized workspace
 * file. Synthetic-lib locations are omitted like other navigation results.
 */
export function workspaceImplementation(
  project: WorkspaceProject,
  path: string,
  offset: number,
): WorkspaceLocation[] {
  const p = normalizePath(path);
  const implementations =
    project.service.getImplementationAtPosition(p, offset);
  if (!implementations) return [];
  const out: WorkspaceLocation[] = [];
  for (const implementation of implementations) {
    if (implementation.fileName === LIB_FILE_NAME) continue;
    out.push({
      path: implementation.fileName,
      from: implementation.textSpan.start,
      to:
        implementation.textSpan.start +
        implementation.textSpan.length,
    });
  }
  return out;
}

/**
 * Turn compiler file changes into validated, bounded edits, or reject the whole
 * action. Every target must be a file this project owns, every span must land
 * inside that file, and edits within a file must not overlap — a partially
 * valid rewrite is refused rather than applied.
 */
function materializeCompilerFileChanges(
  program: TsModule.Program | undefined,
  allowedPaths: ReadonlySet<string>,
  fileChanges: readonly TsModule.FileTextChanges[],
): WorkspaceCodeFixChange[] | null {
  const changes: WorkspaceCodeFixChange[] = [];
  for (const change of fileChanges) {
    if (!allowedPaths.has(change.fileName)) return null;
    const targetSource = program?.getSourceFile(change.fileName)?.text;
    if (targetSource === undefined) return null;
    const edits: BufferEdit[] = [];
    for (const textChange of change.textChanges) {
      const editFrom = textChange.span.start;
      const editTo = editFrom + textChange.span.length;
      if (editFrom < 0 || editTo < editFrom || editTo > targetSource.length) {
        return null;
      }
      edits.push({ from: editFrom, to: editTo, insert: textChange.newText });
    }
    edits.sort((left, right) => left.from - right.from || left.to - right.to);
    let previousTo = -1;
    for (const edit of edits) {
      if (edit.from < previousTo) return null;
      previousTo = edit.to;
    }
    if (edits.length > 0) changes.push({ path: change.fileName, edits });
  }
  changes.sort((left, right) => left.path.localeCompare(right.path));
  return changes;
}

/**
 * Project-aware Quick Fixes for diagnostics overlapping `[from, to]`.
 *
 * Unlike the single-buffer adapter, this retains changes across authorized
 * project files. An action is rejected whole when any target is outside the
 * project, stale/out of bounds, overlapping, or command-only.
 */
export function workspaceCodeFixes(
  project: WorkspaceProject,
  path: string,
  from: number,
  to: number,
): WorkspaceCodeFixAction[] {
  try {
    const p = normalizePath(path);
    const program = project.service.getProgram();
    const activeSource = program?.getSourceFile(p)?.text;
    if (activeSource === undefined) return [];
    const start = Math.min(
      activeSource.length,
      Math.max(0, Math.min(from, to)),
    );
    const end = Math.min(
      activeSource.length,
      Math.max(0, Math.max(from, to)),
    );
    const diagnostics = [
      ...(project.service.getSemanticDiagnostics(p) ?? []),
      ...(project.service.getSyntacticDiagnostics(p) ?? []),
    ];
    const codes = new Set<number>();
    for (const diagnostic of diagnostics) {
      if (diagnostic.start == null) continue;
      const diagnosticEnd =
        diagnostic.start + (diagnostic.length ?? 0);
      if (diagnostic.start <= end && diagnosticEnd >= start) {
        codes.add(diagnostic.code);
      }
    }
    if (codes.size === 0) return [];

    const rawFixes = project.service.getCodeFixesAtPosition(
      p,
      start,
      end,
      [...codes],
      project.ts.getDefaultFormatCodeSettings("\n"),
      {},
    );
    const allowedPaths = new Set(project.paths);
    const actions: WorkspaceCodeFixAction[] = [];
    const seen = new Set<string>();
    for (const fix of rawFixes ?? []) {
      let rejected = false;
      const changes: WorkspaceCodeFixChange[] = [];
      for (const change of fix.changes ?? []) {
        if (!allowedPaths.has(change.fileName)) {
          rejected = true;
          break;
        }
        const targetSource = program?.getSourceFile(
          change.fileName,
        )?.text;
        if (targetSource === undefined) {
          rejected = true;
          break;
        }
        const edits: BufferEdit[] = [];
        for (const textChange of change.textChanges) {
          const editFrom = textChange.span.start;
          const editTo = editFrom + textChange.span.length;
          if (
            editFrom < 0 ||
            editTo < editFrom ||
            editTo > targetSource.length
          ) {
            rejected = true;
            break;
          }
          edits.push({
            from: editFrom,
            to: editTo,
            insert: textChange.newText,
          });
        }
        if (rejected) break;
        edits.sort(
          (left, right) =>
            left.from - right.from || left.to - right.to,
        );
        let previousTo = -1;
        for (const edit of edits) {
          if (edit.from < previousTo) {
            rejected = true;
            break;
          }
          previousTo = edit.to;
        }
        if (rejected) break;
        if (edits.length > 0) {
          changes.push({ path: change.fileName, edits });
        }
      }
      if (rejected || changes.length === 0) continue;
      changes.sort((left, right) =>
        left.path.localeCompare(right.path),
      );
      const key = `${fix.description}\0${changes
        .map(
          (change) =>
            `${change.path}:${change.edits
              .map(
                (edit) =>
                  `${edit.from}:${edit.to}:${edit.insert}`,
              )
              .join("|")}`,
        )
        .join("\0")}`;
      if (seen.has(key)) continue;
      seen.add(key);
      actions.push({ description: fix.description, changes });
    }
    return actions;
  } catch {
    return [];
  }
}

/** Refactorings offered for one selection; the menu is a menu, not a dump. */
export const WORKSPACE_REFACTOR_LIMIT = 50;

/**
 * One refactoring the compiler is willing to perform here. The two names are
 * the compiler's own identifiers for it and are what {@link workspaceRefactorEdits}
 * needs to produce the actual edits — the description is what a human reads.
 */
export interface WorkspaceRefactorAction {
  refactorName: string;
  actionName: string;
  description: string;
}

/**
 * Refactorings applicable to `[from, to)` (VSCode's Refactor… menu): extract a
 * function or constant, convert a function form, move a symbol to another file,
 * and the rest of what the TypeScript service offers for this selection.
 *
 * Quick Fix answers "this is broken, repair it". A refactoring answers "this is
 * fine, restructure it" — different question, different menu, same compiler.
 * Nothing is applied here; this only lists what could be.
 */
export function workspaceRefactors(
  project: WorkspaceProject,
  path: string,
  from: number,
  to: number,
): WorkspaceRefactorAction[] {
  try {
    const p = normalizePath(path);
    const program = project.service.getProgram();
    const activeSource = program?.getSourceFile(p)?.text;
    if (activeSource === undefined) return [];
    const start = Math.min(activeSource.length, Math.max(0, Math.min(from, to)));
    const end = Math.min(activeSource.length, Math.max(0, Math.max(from, to)));
    const raw = project.service.getApplicableRefactors(
      p,
      start === end ? start : { pos: start, end },
      {},
    );
    const actions: WorkspaceRefactorAction[] = [];
    const seen = new Set<string>();
    for (const refactor of raw ?? []) {
      for (const action of refactor.actions ?? []) {
        if (actions.length >= WORKSPACE_REFACTOR_LIMIT) return actions;
        // A refactor the compiler already knows it cannot run here (for
        // example "Move to a new file" with nothing selected) is not an offer.
        if (action.notApplicableReason) continue;
        const key = `${refactor.name}\0${action.name}`;
        if (seen.has(key)) continue;
        seen.add(key);
        actions.push({
          refactorName: refactor.name,
          actionName: action.name,
          description: action.description,
        });
      }
    }
    return actions;
  } catch {
    return [];
  }
}

/**
 * Materialize ONE refactoring into the same validated, bounded file changes a
 * Quick Fix produces, so both apply through the identical atomic host path.
 * Returns null when the compiler declines, or when any edit falls outside the
 * authorized project or overlaps another.
 */
export function workspaceRefactorEdits(
  project: WorkspaceProject,
  path: string,
  from: number,
  to: number,
  refactorName: string,
  actionName: string,
): WorkspaceCodeFixAction | null {
  try {
    const p = normalizePath(path);
    const program = project.service.getProgram();
    const activeSource = program?.getSourceFile(p)?.text;
    if (activeSource === undefined) return null;
    const start = Math.min(activeSource.length, Math.max(0, Math.min(from, to)));
    const end = Math.min(activeSource.length, Math.max(0, Math.max(from, to)));
    const edits = project.service.getEditsForRefactor(
      p,
      project.ts.getDefaultFormatCodeSettings("\n"),
      start === end ? start : { pos: start, end },
      refactorName,
      actionName,
      {},
    );
    if (!edits) return null;
    // A refactor that wants to CREATE a file (Move to a new file) is declined:
    // this workspace has no path the compiler may invent a file at.
    if (edits.edits.some((change) => change.isNewFile)) return null;
    const changes = materializeCompilerFileChanges(
      program,
      new Set(project.paths),
      edits.edits,
    );
    if (!changes || changes.length === 0) return null;
    const description =
      workspaceRefactors(project, path, from, to).find(
        (action) =>
          action.refactorName === refactorName &&
          action.actionName === actionName,
      )?.description ?? actionName;
    return { description, changes };
  } catch {
    return null;
  }
}

/**
 * Cross-file rename: every occurrence of the symbol at `offset` in `path` that a
 * rename must rewrite, ACROSS files. Synthetic-lib hits are dropped. Returns
 * `null` when the position is not renameable (e.g. a keyword, a literal, or a
 * symbol from the ambient lib) so callers can leave the buffer untouched.
 */
export function workspaceRenameLocations(
  project: WorkspaceProject,
  path: string,
  offset: number,
): WorkspaceRenameLocation[] | null {
  const p = normalizePath(path);
  // providePrefixAndSuffixTextForRename: true — REQUIRED so TS returns the
  // prefix/suffix that keeps object-shorthand / aliased renames correct (the
  // deprecated 5th-boolean overload `false` drops it and corrupts shorthand).
  const locs = project.service.findRenameLocations(p, offset, false, false, {
    providePrefixAndSuffixTextForRename: true,
  });
  if (!locs) return null;
  const out: WorkspaceRenameLocation[] = [];
  for (const l of locs) {
    if (l.fileName === LIB_FILE_NAME) continue;
    out.push({
      path: l.fileName,
      from: l.textSpan.start,
      to: l.textSpan.start + l.textSpan.length,
      prefixText: l.prefixText,
      suffixText: l.suffixText,
    });
  }
  return out.length > 0 ? out : null;
}

/** Quick-info (hover) content: the symbol signature and its doc, as plain text. */
export interface WorkspaceQuickInfo {
  /** The symbol's signature line (e.g. "function add(a: number, b: number): number"). */
  signature: string;
  /** The symbol's documentation, or "" when it has none. */
  documentation: string;
}

/**
 * Cross-file hover quick-info: the signature + docs of the symbol at `offset` in
 * `path`, resolved through the WHOLE project — so a symbol declared in another
 * file/cell shows its real signature (a single buffer can only say "any").
 * Returns plain text (caller renders via textContent — no markup), or `null`
 * when there's nothing to show.
 */
export function workspaceQuickInfo(
  project: WorkspaceProject,
  path: string,
  offset: number,
): WorkspaceQuickInfo | null {
  const p = normalizePath(path);
  const info = project.service.getQuickInfoAtPosition(p, offset);
  if (!info) return null;
  const signature = project.ts.displayPartsToString(info.displayParts ?? []);
  const documentation = project.ts.displayPartsToString(info.documentation ?? []);
  if (!signature && !documentation) return null;
  return { signature, documentation };
}

/**
 * Compiler-backed completion over the whole authorized workspace project.
 *
 * The result is fail-closed for unknown files, invalid offsets, invalid limits,
 * or compiler failures. `maxEntries` is applied before results cross into the UI
 * so a large ambient/type surface cannot create an unbounded completion popup.
 */
export function workspaceCompletions(
  project: WorkspaceProject,
  path: string,
  offset: number,
  maxEntries: number,
): WorkspaceCompletionResult | null {
  if (!Number.isSafeInteger(maxEntries) || maxEntries < 1) return null;
  const normalizedPath = normalizePath(path);
  const source =
    project.service.getProgram()?.getSourceFile(normalizedPath)?.text;
  if (
    source === undefined ||
    !Number.isSafeInteger(offset) ||
    offset < 0 ||
    offset > source.length
  ) {
    return null;
  }
  const result = completionsAt(
    {
      service: project.service,
      fileName: normalizedPath,
      markDirty: () => {},
      ts: project.ts,
    },
    offset,
  );
  if (!result || result.entries.length === 0) return null;
  const entries = result.entries.slice(0, maxEntries);
  if (entries.length === 0) return null;
  return {
    from: result.from,
    to: result.to,
    isMember: result.isMember,
    entries,
  };
}

/**
 * Project-aware inlay hints for one authorized file.
 *
 * The single-buffer mapper is reused unchanged, but its LanguageService now
 * sees every workspace file. That lets a call in one multi-buffer host cell display the
 * parameter name declared in a sibling cell. The caller controls a hard result
 * ceiling in addition to the visible-range span enforced by `inlayHints`.
 */
export function workspaceInlayHints(
  project: WorkspaceProject,
  path: string,
  fromOffset: number,
  toOffset: number,
  maxHints = WORKSPACE_INLAY_HINT_LIMIT,
): InlayHintItem[] {
  if (!Number.isSafeInteger(maxHints) || maxHints < 1) return [];
  const normalizedPath = normalizePath(path);
  if (!project.paths.includes(normalizedPath)) return [];
  const hints = inlayHints(
    {
      service: project.service,
      fileName: normalizedPath,
      markDirty: () => {},
      ts: project.ts,
    },
    fromOffset,
    toOffset,
  );
  return hints.slice(0, maxHints);
}

/**
 * Project-aware Code Lens reference counts for one authorized file.
 *
 * The navigation tree is traversed once to collect at most `maxLenses`
 * definition sites. Every reference query then reuses this SAME project and
 * counts uses across all files/cells, excluding the symbol's own declaration.
 * No project is created per symbol.
 */
export function workspaceCodeLenses(
  project: WorkspaceProject,
  path: string,
  maxLenses = WORKSPACE_CODE_LENS_LIMIT,
): CodeLens[] {
  if (!Number.isSafeInteger(maxLenses) || maxLenses < 1) return [];
  const normalizedPath = normalizePath(path);
  try {
    const sourceFile = project.service
      .getProgram()
      ?.getSourceFile(normalizedPath);
    const tree =
      project.service.getNavigationTree(normalizedPath);
    if (!sourceFile || !tree) return [];
    const source = sourceFile.text;
    const lineStarts = [0];
    for (let index = 0; index < source.length; index += 1) {
      const code = source.charCodeAt(index);
      if (code === 10) {
        lineStarts.push(index + 1);
      } else if (code === 13) {
        if (source.charCodeAt(index + 1) === 10) index += 1;
        lineStarts.push(index + 1);
      }
    }
    const lineAt = (offset: number): number => {
      let low = 0;
      let high = lineStarts.length - 1;
      while (low < high) {
        const middle = (low + high + 1) >> 1;
        if (lineStarts[middle] <= offset) low = middle;
        else high = middle - 1;
      }
      return low + 1;
    };
    const definitions: Array<{
      name: string;
      line: number;
      nameOffset: number;
    }> = [];
    const walk = (node: TsModule.NavigationTree): void => {
      if (node.nameSpan) {
        definitions.push({
          name: node.text,
          line: lineAt(node.nameSpan.start),
          nameOffset: node.nameSpan.start,
        });
      }
      for (const child of node.childItems ?? []) walk(child);
    };
    for (const child of tree.childItems ?? []) walk(child);
    definitions.sort(
      (left, right) => left.nameOffset - right.nameOffset,
    );

    const lenses: CodeLens[] = [];
    const seenLines = new Set<number>();
    for (const definition of definitions) {
      if (seenLines.has(definition.line)) continue;
      seenLines.add(definition.line);
      const refs = workspaceReferences(
        project,
        normalizedPath,
        definition.nameOffset,
      );
      const uniqueUses = new Set<string>();
      for (const ref of refs) {
        if (
          ref.path === normalizedPath &&
          ref.from === definition.nameOffset
        ) {
          continue;
        }
        uniqueUses.add(`${ref.path}:${ref.from}:${ref.to}`);
      }
      lenses.push({
        ...definition,
        refCount: uniqueUses.size,
      });
      if (lenses.length >= maxLenses) break;
    }
    return lenses;
  } catch {
    return [];
  }
}

/**
 * Project-aware parameter hints for a call in `path`. Reuses the single-buffer
 * signature conversion/sanitisation contract over the workspace LanguageService
 * so imported functions get their real overload, parameter types and active
 * argument without maintaining a second formatter.
 */
export function workspaceSignatureHelp(
  project: WorkspaceProject,
  path: string,
  offset: number,
): SignatureHelpInfo {
  return signatureHelpAt(
    {
      service: project.service,
      fileName: normalizePath(path),
      markDirty: () => {},
      ts: project.ts,
    },
    offset,
  );
}

/** Calls listed per direction; a hot symbol must not flood the panel. */
export const WORKSPACE_CALL_HIERARCHY_LIMIT = 200;

export interface WorkspaceCall {
  /** The function on the other end of the call. */
  readonly name: string;
  /** Where that function is declared. */
  readonly path: string;
  readonly from: number;
  readonly to: number;
  /** Call sites tying the two together, in the caller's file. */
  readonly sites: readonly WorkspaceLocation[];
}

export interface WorkspaceCallHierarchy {
  /** The symbol the hierarchy was prepared for. */
  readonly name: string;
  readonly incoming: readonly WorkspaceCall[];
  readonly outgoing: readonly WorkspaceCall[];
  /** Either direction hit the limit. */
  readonly truncated: boolean;
}

function callHierarchySpan(
  item: TsModule.CallHierarchyItem,
): { path: string; from: number; to: number } {
  const span = item.selectionSpan ?? item.span;
  return {
    path: item.file,
    from: span.start,
    to: span.start + span.length,
  };
}

/**
 * Cross-file call hierarchy for the symbol at `offset` (VSCode's "Show Call
 * Hierarchy"): who calls this function, and what this function calls.
 *
 * References answer "where is this name used"; a call hierarchy answers the
 * question people actually chase — how does control reach here, and where does
 * it go next. Both directions come from the real TypeScript service, so they
 * follow the resolver rather than matching text.
 *
 * Returns null when the position is not on something callable.
 */
export function workspaceCallHierarchy(
  project: WorkspaceProject,
  path: string,
  offset: number,
): WorkspaceCallHierarchy | null {
  const p = normalizePath(path);
  const service = project.service;
  if (
    typeof service.prepareCallHierarchy !== "function" ||
    typeof service.provideCallHierarchyIncomingCalls !== "function" ||
    typeof service.provideCallHierarchyOutgoingCalls !== "function"
  ) {
    return null;
  }
  let prepared: TsModule.CallHierarchyItem | undefined;
  try {
    const raw = service.prepareCallHierarchy(p, offset);
    prepared = Array.isArray(raw) ? raw[0] : (raw ?? undefined);
  } catch {
    return null;
  }
  if (!prepared) return null;

  let truncated = false;
  const incoming: WorkspaceCall[] = [];
  const outgoing: WorkspaceCall[] = [];
  try {
    const calls = service.provideCallHierarchyIncomingCalls(
      prepared.file,
      prepared.selectionSpan.start,
    );
    for (const call of calls) {
      if (call.from.file === LIB_FILE_NAME) continue;
      if (incoming.length >= WORKSPACE_CALL_HIERARCHY_LIMIT) {
        truncated = true;
        break;
      }
      const located = callHierarchySpan(call.from);
      incoming.push({
        name: call.from.name,
        ...located,
        sites: call.fromSpans.map((span) => ({
          path: call.from.file,
          from: span.start,
          to: span.start + span.length,
        })),
      });
    }
  } catch {
    /* a service that cannot answer one direction still answers the other */
  }
  try {
    const calls = service.provideCallHierarchyOutgoingCalls(
      prepared.file,
      prepared.selectionSpan.start,
    );
    for (const call of calls) {
      if (call.to.file === LIB_FILE_NAME) continue;
      if (outgoing.length >= WORKSPACE_CALL_HIERARCHY_LIMIT) {
        truncated = true;
        break;
      }
      const located = callHierarchySpan(call.to);
      outgoing.push({
        name: call.to.name,
        ...located,
        // Outgoing call sites live in the PREPARED file, not the callee's.
        sites: call.fromSpans.map((span) => ({
          path: prepared.file,
          from: span.start,
          to: span.start + span.length,
        })),
      });
    }
  } catch {
    /* see above */
  }
  if (incoming.length === 0 && outgoing.length === 0) return null;
  return { name: prepared.name, incoming, outgoing, truncated };
}

/**
 * Cross-file find-all-references: every use of the symbol at `offset` in `path`,
 * ACROSS files (declaration included, matching VSCode). Synthetic-lib hits are
 * dropped. Returns `[]` when the position is not on a referenceable symbol.
 */
export function workspaceReferences(
  project: WorkspaceProject,
  path: string,
  offset: number,
): WorkspaceLocation[] {
  const p = normalizePath(path);
  const refs = project.service.getReferencesAtPosition(p, offset);
  if (!refs) return [];
  const out: WorkspaceLocation[] = [];
  for (const r of refs) {
    if (r.fileName === LIB_FILE_NAME) continue;
    out.push({
      path: r.fileName,
      from: r.textSpan.start,
      to: r.textSpan.start + r.textSpan.length,
    });
  }
  return out;
}
