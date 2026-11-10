/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { isDoNotTrackEnabled, sendErrorBeacon } from "./errors/report-beacon";
import { currentWebBuildId } from "./web-build";

export type PageSaveDiagnosticAttempt = "explicit" | "automatic";

/** Everything a caller may report. Unknown keys and non-matching values are dropped. */
export interface PageSaveDiagnosticInput {
  pageId: string;
  /** Where the save stopped, e.g. `autosave.causal_model` or `editor.canonical_prepare`. */
  stage: string;
  /** Stable failure code, e.g. `page_causal_model_failed`. */
  code: string;
  /** A finer code for the same stage, e.g. `unbound_draft_at_preparation`. */
  reason?: string | null;
  attempt?: PageSaveDiagnosticAttempt | null;
  /** Whether an HTTP request for this save had been sent. */
  requestSent?: boolean | null;
  windowId?: string | null;
   
  saveGeneration?: number | null;
  /** The editor's local edit generation and the one last queued for save. */
  localGeneration?: number | null;
  queuedGeneration?: number | null;
  queuedGenerationMatches?: boolean | null;
  localDirty?: boolean | null;
  canonicalReady?: boolean | null;
  imeComposing?: boolean | null;
  /** Which part of the editor asked for this save (export, Copilot, ...). */
  caller?: string | null;
  causalStatus?: string | null;
  causalFailure?: string | null;
  /** True once the original model was ever bound to a durable identity. */
  causalBound?: boolean | null;
  causalUnboundDraft?: boolean | null;
  /** Where a body the room refused first differs from it: block indices and
   * schema attribute names only (lib/collab/document-page-causal-model
   * lastDocumentMirrorMiss), e.g. `content[29].attrs.source.*`. */
  causalMirrorMiss?: string | null;
  /** Where the capture's room-body fallback stopped (`no_baseline`,
   * `view_unproven:<why>`, `room_body_not_mirrored`, ...). */
  causalCaptureFallback?: string | null;
  durableContentVersion?: number | null;
  durableCausalRevision?: number | null;
  baseContentVersion?: number | null;
  receiptContentVersion?: number | null;
  quarantined?: boolean | null;
}

export interface PageSaveDiagnostic extends Readonly<Partial<Omit<PageSaveDiagnosticInput, "pageId" | "stage" | "code">>> {
  readonly at: number;
  readonly webBuild: string;
  readonly pageId: string;
  readonly stage: string;
  readonly code: string;
}

const MAX_ENTRIES = 100;
/** Codes and stage names: short identifiers, never free text. */
const CODE = /^[A-Za-z0-9_.:-]{1,96}$/u;
/** Page and window ids are opaque identifiers (UUIDs in practice). */
const IDENTIFIER = /^[A-Za-z0-9_-]{1,80}$/u;
/** A structural path: identifiers, indices and the separators the miss path uses. */
const STRUCTURE_PATH = /^[A-Za-z0-9_.:*|+/[\] -]{1,200}$/u;
const NUMBERS = ["saveGeneration", "localGeneration", "queuedGeneration", "durableContentVersion",
  "durableCausalRevision", "baseContentVersion", "receiptContentVersion"] as const;
const FLAGS = ["requestSent", "queuedGenerationMatches", "localDirty", "canonicalReady", "imeComposing",
  "causalBound", "causalUnboundDraft", "quarantined"] as const;
const CODES = ["reason", "caller", "causalStatus", "causalFailure", "causalCaptureFallback"] as const;

const entries: PageSaveDiagnostic[] = [];
let lastPrinted: string | null = null;

function installReader(): void {
  if (typeof window === "undefined") return;
  if (Object.hasOwn(window, "__mrnoteSaveDiagnostics")) return;
  try {
    Object.defineProperty(window, "__mrnoteSaveDiagnostics", {
      configurable: true,
      value: Object.freeze({
        dump: (): PageSaveDiagnostic[] => entries.map((entry) => ({ ...entry })),
        clear: (): void => { entries.length = 0; lastPrinted = null; },
      }),
    });
  } catch {
    // A locked-down global must not break saving.
  }
}

/** Build the allow-listed, content-free entry. Exported for tests. */
export function buildPageSaveDiagnostic(input: PageSaveDiagnosticInput, now = Date.now()): PageSaveDiagnostic | null {
  if (!input || typeof input !== "object") return null;
  if (typeof input.pageId !== "string" || !IDENTIFIER.test(input.pageId)) return null;
  if (typeof input.stage !== "string" || !CODE.test(input.stage)) return null;
  if (typeof input.code !== "string" || !CODE.test(input.code)) return null;
  const entry: Record<string, unknown> = {
    at: now, webBuild: currentWebBuildId() || "development",
    pageId: input.pageId, stage: input.stage, code: input.code,
  };
  for (const key of CODES) {
    const value = input[key];
    if (typeof value === "string" && CODE.test(value)) entry[key] = value;
  }
  if (typeof input.causalMirrorMiss === "string" && STRUCTURE_PATH.test(input.causalMirrorMiss)) {
    entry.causalMirrorMiss = input.causalMirrorMiss;
  }
  if (input.attempt === "explicit" || input.attempt === "automatic") entry.attempt = input.attempt;
  if (typeof input.windowId === "string" && IDENTIFIER.test(input.windowId)) entry.windowId = input.windowId;
  for (const key of NUMBERS) {
    const value = input[key];
    if (typeof value === "number" && Number.isSafeInteger(value)) entry[key] = value;
  }
  for (const key of FLAGS) {
    const value = input[key];
    if (typeof value === "boolean") entry[key] = value;
  }
  return Object.freeze(entry) as unknown as PageSaveDiagnostic;
}

/** Record one refusal. Never throws. Returns the stored entry, or null. */
export function recordPageSaveDiagnostic(input: PageSaveDiagnosticInput): PageSaveDiagnostic | null {
  try {
    const entry = buildPageSaveDiagnostic(input);
    if (!entry) return null;
    entries.push(entry);
    if (entries.length > MAX_ENTRIES) entries.splice(0, entries.length - MAX_ENTRIES);
    installReader();
    // One console line per distinct verdict: an explicit Save pressed ten
    // times for the same reason must not flood the console.
    const key = `${entry.pageId}|${entry.stage}|${entry.code}|${entry.reason ?? ""}|${entry.causalStatus ?? ""}`;
    if (key !== lastPrinted) {
      lastPrinted = key;
      console.warn("[page-save]", entry.requestSent ? "receipt refused" : "not sent", entry);
      reportPageSaveRefusal(entry);
    }
    return entry;
  } catch {
    return null;
  }
}







function reportPageSaveRefusal(entry: PageSaveDiagnostic): void {
  try {
    if (typeof window === "undefined" || isDoNotTrackEnabled()) return;
    sendErrorBeacon({
      name: "PageSaveRefusal",
      message: JSON.stringify(entry).slice(0, 900),
      path: window.location?.pathname || "/",
      source: "window",
    });
  } catch {
    // A report must never get in the way of saving.
  }
}

export function readPageSaveDiagnostics(): PageSaveDiagnostic[] {
  return entries.map((entry) => ({ ...entry }));
}

/** Test seam. */
export function resetPageSaveDiagnosticsForTests(): void {
  entries.length = 0;
  lastPrinted = null;
}
