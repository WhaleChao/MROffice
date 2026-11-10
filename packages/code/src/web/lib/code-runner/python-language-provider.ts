/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import {
  codePointToUtf16Col,
  identifierAt,
  identifierSpanAt,
  isRenamablePythonIdentifier,
  utf16ColToCodePointCol,
} from "./python-source";
import type {
  PythonCompletionItem,
  PythonDefinitionInfo,
  PythonHoverInfo,
  PythonReference,
} from "@mrnote/core/web/lib/code-runner/python-runner.ts";
import type {
  LanguageServiceProvider,
  LanguageServiceRequest,
  ProviderCapabilities,
  ProviderCompletion,
  ProviderHover,
  ProviderLocation,
  ProviderRenameEdit,
} from "./language-service-provider";

const LANGUAGES = ["python", "py"] as const;

/** jedi query position: 1-based line, 0-based code-point column. */
interface JediPosition {
  line: number;
  column: number;
}

interface RuntimeOptions {
  namespace?: string;
}

/**
 * The slice of `python-runner` this provider uses.
 *
 * Named as a type so the module can stay a lazy `import()` — a page that never
 * asks Python anything should not pull the runner into its chunk.
 */
export interface PythonRuntimeOps {
  hasPythonKernel: (namespace?: string) => boolean;
  completePython: (
    code: string,
    line: number,
    column: number,
    options?: RuntimeOptions,
  ) => Promise<PythonCompletionItem[]>;
  hoverPython: (
    code: string,
    line: number,
    column: number,
    options?: RuntimeOptions,
  ) => Promise<PythonHoverInfo | null>;
  definitionPython: (
    code: string,
    line: number,
    column: number,
    options?: RuntimeOptions,
  ) => Promise<PythonDefinitionInfo | null>;
  referencesPython: (
    code: string,
    line: number,
    column: number,
    options?: RuntimeOptions,
  ) => Promise<PythonReference[]>;
}

export interface PythonLanguageProviderOptions {
  /**
   * Which kernel to ask. Read per call, not captured once: the host switches
   * pages, and asking a stale namespace would answer about another notebook's
   * variables. `null` means "no kernel to ask" and every op degrades to empty.
   */
  namespace: () => string | null;
  /** Seam for tests; defaults to the real lazy `python-runner` import. */
  loadRuntime?: () => Promise<PythonRuntimeOps>;
}

export interface PythonLanguageServiceProvider extends LanguageServiceProvider {
  /**
   * Load the runtime module so `capabilities()` can start answering.
   *
   * `capabilities()` is synchronous and the runtime import is not, so until the
   * module resolves the provider reports nothing — "I cannot answer yet" rather
   * than a guess that might be wrong. The host MUST call this (from an effect,
   * not from render) or Python's capabilities stay empty forever; the ops
   * themselves prime on demand, so a call still works, but the affordances
   * gated on capabilities would never appear.
   */
  prime(): Promise<void>;
}

/** One line of a buffer: its start offset and its text (no newline). */
interface IndexedLine {
  start: number;
  text: string;
}

/** Split `text` into lines with their start offsets. */
function indexLines(text: string): IndexedLine[] {
  const lines: IndexedLine[] = [];
  let start = 0;
  for (;;) {
    const nl = text.indexOf("\n", start);
    if (nl === -1) {
      lines.push({ start, text: text.slice(start) });
      return lines;
    }
    lines.push({ start, text: text.slice(start, nl) });
    start = nl + 1;
  }
}

/** A UTF-16 offset as the (1-based line, 0-based code-point column) jedi wants. */
function toJediPosition(lines: IndexedLine[], offset: number): JediPosition {
  let lo = 0;
  let hi = lines.length - 1;
  while (lo < hi) {
    const mid = (lo + hi + 1) >> 1;
    if (lines[mid].start <= offset) lo = mid;
    else hi = mid - 1;
  }
  const line = lines[lo];
  const within = Math.max(0, offset - line.start);
  return { line: lo + 1, column: utf16ColToCodePointCol(line.text, within) };
}

/**
 * jedi's (1-based line, 0-based code-point column) as a UTF-16 buffer offset.
 * Clamps: a position from a round trip the user has since edited past must not
 * throw or point outside the buffer.
 */
function toOffset(
  lines: IndexedLine[],
  line: number,
  codePointColumn: number,
): number {
  const index = Math.min(Math.max(Math.trunc(line), 1), lines.length) - 1;
  const target = lines[index];
  return (
    target.start + codePointToUtf16Col(target.text, Math.max(codePointColumn, 0))
  );
}

const NO_CAPABILITIES: ProviderCapabilities = {};

/**
 * What this provider can do once a kernel exists. jedi answers all five from the
 * same live namespace, so they arrive together or not at all.
 */
const KERNEL_CAPABILITIES: ProviderCapabilities = {
  completions: true,
  hover: true,
  definition: true,
  references: true,
  rename: true,
};

async function loadPythonRuntime(): Promise<PythonRuntimeOps> {
  const mod = await import("@mrnote/core/web/lib/code-runner/python-runner.ts");
  return {
    hasPythonKernel: mod.hasPythonKernel,
    completePython: mod.completePython,
    hoverPython: mod.hoverPython,
    definitionPython: mod.definitionPython,
    referencesPython: mod.referencesPython,
  };
}

export function createPythonLanguageProvider(
  options: PythonLanguageProviderOptions,
): PythonLanguageServiceProvider {
  const loadRuntime = options.loadRuntime ?? loadPythonRuntime;
  let runtime: PythonRuntimeOps | null = null;
  let priming: Promise<void> | null = null;

  const prime = (): Promise<void> => {
    if (runtime) return Promise.resolve();
    if (!priming) {
      priming = loadRuntime()
        .then(
          (ops) => {
            runtime = ops;
          },
          () => {
            // A runtime that failed to load leaves capabilities empty, which is
            // the truth. Clearing `priming` below lets a later ask retry.
          },
        )
        .finally(() => {
          priming = null;
        });
    }
    return priming;
  };

  /**
   * The runtime + namespace an op needs, or null when the answer is "nothing".
   *
   * The kernel check is the load-bearing one: without it, a keystroke could boot
   * Pyodide. Every op funnels through here so that rule cannot be forgotten in
   * one of them.
   */
  const ready = async (): Promise<{
    ops: PythonRuntimeOps;
    namespace: string;
  } | null> => {
    await prime();
    const ops = runtime;
    if (!ops) return null;
    const namespace = options.namespace();
    if (!namespace) return null;
    if (!ops.hasPythonKernel(namespace)) return null;
    return { ops, namespace };
  };

  /**
   * jedi runs in a worker that cannot be cancelled mid-request, so an abort
   * cannot stop the round trip — but it must still stop the RESULT from being
   * used, or a completion popup renders answers for a caret the user has left.
   */
  const aborted = (signal?: AbortSignal): boolean => signal?.aborted === true;

  const referenceLocations = async (
    request: LanguageServiceRequest,
  ): Promise<ProviderLocation[]> => {
    const context = await ready();
    if (!context) return [];
    const lines = indexLines(request.text);
    const position = toJediPosition(lines, request.offset);
    const refs = await context.ops.referencesPython(
      request.text,
      position.line,
      position.column,
      { namespace: context.namespace },
    );
    if (aborted(request.signal)) return [];
    return refs
      .map((ref) => {
        const index =
          Math.min(Math.max(Math.trunc(ref.line), 1), lines.length) - 1;
        const line = lines[index];
        // Both ends go through the code-point mapping. Adding `length` in UTF-16
        // space would be wrong whenever an astral char sits before — or inside —
        // the identifier.
        const startColumn = Math.max(ref.column, 0);
        const from =
          line.start + codePointToUtf16Col(line.text, startColumn);
        const to =
          line.start +
          codePointToUtf16Col(line.text, startColumn + Math.max(ref.length, 0));
        return { bufferId: null, from, to };
      })
      .filter((location) => location.to > location.from);
  };

  return {
    id: "python-jedi",
    languages: LANGUAGES,

    prime,

    capabilities(language: string): ProviderCapabilities {
      if (!(LANGUAGES as readonly string[]).includes(language)) {
        return NO_CAPABILITIES;
      }
      const ops = runtime;
      // Not loaded yet: say nothing. This must NOT kick the load itself —
      // `capabilities()` is read during render, and a query that starts an
      // import is a side effect in render. The host primes explicitly instead.
      if (!ops) return NO_CAPABILITIES;
      const namespace = options.namespace();
      if (!namespace || !ops.hasPythonKernel(namespace)) return NO_CAPABILITIES;
      return KERNEL_CAPABILITIES;
    },

    async completions(
      request: LanguageServiceRequest & { language: string },
    ): Promise<ProviderCompletion[]> {
      const context = await ready();
      if (!context) return [];
      const position = toJediPosition(indexLines(request.text), request.offset);
      const items = await context.ops.completePython(
        request.text,
        position.line,
        position.column,
        { namespace: context.namespace },
      );
      if (aborted(request.signal)) return [];
      // jedi's `kind` is already CodeMirror's vocabulary ("function", "module",
      // …); `detail` would need a second round trip, so it stays absent rather
      // than being filled with something that is not the detail.
      return items.map((item) => ({
        label: item.name,
        ...(item.kind ? { kind: item.kind } : {}),
      }));
    },

    async hover(
      request: LanguageServiceRequest & { language: string },
    ): Promise<ProviderHover | null> {
      const context = await ready();
      if (!context) return null;
      const position = toJediPosition(indexLines(request.text), request.offset);
      const info = await context.ops.hoverPython(
        request.text,
        position.line,
        position.column,
        { namespace: context.namespace },
      );
      if (aborted(request.signal)) return null;
      if (!info || (!info.signature && !info.doc)) return null;
      const contents =
        info.signature && info.doc
          ? `${info.signature}\n\n${info.doc}`
          : (info.signature ?? info.doc ?? "");
      if (!contents) return null;
      // jedi reports no span, only a name. Rather than invent one, the caret's
      // identifier is offered ONLY when it reads back as exactly that name — the
      // same read-back guard the cross-cell rename uses.
      const span = identifierSpanAt(request.text, request.offset);
      const verified =
        span && request.text.slice(span.from, span.to) === info.name
          ? span
          : undefined;
      return verified ? { contents, span: verified } : { contents };
    },

    async definition(
      request: LanguageServiceRequest & { language: string },
    ): Promise<ProviderLocation | null> {
      const context = await ready();
      if (!context) return null;
      const lines = indexLines(request.text);
      const position = toJediPosition(lines, request.offset);
      const info = await context.ops.definitionPython(
        request.text,
        position.line,
        position.column,
        { namespace: context.namespace },
      );
      if (aborted(request.signal)) return null;
      // An out-of-buffer definition (stdlib, an installed package) is real but
      // has no location this interface can name — `bufferId` addresses buffers
      // the host owns, and the host does not own CPython's source. Reporting
      // null keeps the caller's "not found" notice, which is what the editor
      // already shows for this case; inventing an in-buffer offset would send
      // the user to an unrelated line.
      if (!info || !info.inBuffer) return null;
      const from = toOffset(lines, info.line, info.column);
      // jedi gives the name's start, not its extent. Extend to the end of the
      // identifier only when one actually starts there; otherwise a caret.
      const span = identifierSpanAt(request.text, from);
      const to = span && span.from === from ? span.to : from;
      return { bufferId: null, from, to };
    },

    references(
      request: LanguageServiceRequest & { language: string },
    ): Promise<ProviderLocation[]> {
      return referenceLocations(request);
    },

    /**
     * Rename via jedi's references, SAFE BY CONSTRUCTION: the new name must be a
     * legal Python identifier, and every span must read back as exactly the old
     * name. Any mismatch returns no edits at all — a rename that is half right
     * is worse than one that did not happen, and a stale span from a buffer the
     * user kept typing into is precisely how that happens.
     */
    async rename(
      request: LanguageServiceRequest & { language: string; newName: string },
    ): Promise<ProviderRenameEdit[]> {
      if (!isRenamablePythonIdentifier(request.newName)) return [];
      const oldName = identifierAt(request.text, request.offset);
      if (!oldName || oldName === request.newName) return [];
      const locations = await referenceLocations(request);
      if (locations.length === 0) return [];
      // The worker can report the same occurrence twice; two edits over one span
      // would splice into the freshly written name.
      const seen = new Set<string>();
      const edits: ProviderRenameEdit[] = [];
      for (const location of locations) {
        const key = `${location.from}:${location.to}`;
        if (seen.has(key)) continue;
        seen.add(key);
        if (request.text.slice(location.from, location.to) !== oldName) {
          return [];
        }
        edits.push({
          bufferId: null,
          span: { from: location.from, to: location.to },
          newText: request.newName,
        });
      }
      return edits;
    },
  };
}
