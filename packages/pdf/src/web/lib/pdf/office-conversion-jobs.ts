/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * PDF → Word / Excel / PPT conversion jobs, followed outside any component.
 *
 * A conversion runs on the server for seconds to minutes. It keeps being
 * followed here when its dialog closes ("continue in the background") or the
 * PDF window that started it unmounts, so its result still reaches the
 * person. One long-poll loop per job, ending when the job ends; the job's
 * result is handed to exactly one registered handler, exactly once.
 */
import { isApiRequestError } from "@mrnote/core/web/lib/api.ts";
import {
  cancelPdfConversionJob as requestPdfConversionCancel,
  getPdfConversionJob as requestPdfConversionJob,
  listPdfConversionJobs,
  type PdfConversionJob,
  type PdfConversionJobError,
} from "@mrnote/core/web/lib/resources-sdk.ts";

/** How long the server may hold one poll before answering unchanged. */
export const PDF_CONVERSION_POLL_WAIT_MS = 4000;
/** Pause after a poll that failed in a way worth retrying. */
const PDF_CONVERSION_RETRY_DELAY_MS = 1000;
/** Failed polls in a row for this long end the job here with a failure. */
const PDF_CONVERSION_GIVE_UP_MS = 60_000;
/** A poll that came back unchanged sooner than this waits before the next. */
const PDF_CONVERSION_MIN_POLL_INTERVAL_MS = 500;
/** A poll the server has not answered within this is dropped and retried. */
const PDF_CONVERSION_POLL_TIMEOUT_MS = PDF_CONVERSION_POLL_WAIT_MS + 15_000;
/** Ended jobs kept for a dialog that still shows one. */
const PDF_CONVERSION_ENDED_JOBS_KEPT = 16;

/** Code of the failure this client reports when it lost the job. */
export const PDF_CONVERSION_UNREACHABLE_CODE = "pdf_conversion_unreachable";

/**
 * Who takes a job's result. Several windows may follow one job; the result
 * goes to the one whose dialog shows it (foreground), else to one still
 * mounted, else to the last one registered: one result, one notice.
 */
export interface PdfConversionJobSettleHandler {
  /** A mounted dialog shows this job now. */
  readonly isForeground: () => boolean;
  /** The window that registered this handler is still mounted. */
  readonly isMounted: () => boolean;
  readonly settle: (job: PdfConversionJob, foreground: boolean) => void;
}

type PdfConversionJobListener = (job: PdfConversionJob) => void;

interface TrackedPdfConversionJob {
  job: PdfConversionJob;
  readonly listeners: Set<PdfConversionJobListener>;
  readonly handlers: PdfConversionJobSettleHandler[];
  settled: boolean;
  polling: boolean;
}

const trackedJobs = new Map<string, TrackedPdfConversionJob>();
const changeListeners = new Set<() => void>();
const reattachFlights = new Map<string, Promise<PdfConversionJob[]>>();
let trackerGeneration = 0;

export function isPdfConversionJobEnded(job: PdfConversionJob): boolean {
  return (
    job.status === "succeeded" ||
    job.status === "failed" ||
    job.status === "cancelled"
  );
}

function notifyChange(): void {
  for (const listener of [...changeListeners]) {
    try {
      listener();
    } catch {
      // A listener's failure must not stop the job being followed.
    }
  }
}

function notify(entry: TrackedPdfConversionJob): void {
  for (const listener of [...entry.listeners]) {
    try {
      listener(entry.job);
    } catch {
      // As above.
    }
  }
  notifyChange();
}

/** Take `next` when it is newer news about the same job; an ended job is final. */
function accept(entry: TrackedPdfConversionJob, next: PdfConversionJob): boolean {
  const current = entry.job;
  if (
    next.job_id !== current.job_id ||
    next.resource_id !== current.resource_id ||
    isPdfConversionJobEnded(current) ||
    next.revision < current.revision ||
    (next.revision === current.revision && !isPdfConversionJobEnded(next))
  ) {
    return false;
  }
  entry.job = next;
  notify(entry);
  return true;
}

function pruneEndedJobs(): void {
  let ended = 0;
  for (const entry of trackedJobs.values()) {
    if (entry.settled) ended += 1;
  }
  for (const [jobId, entry] of trackedJobs) {
    if (ended <= PDF_CONVERSION_ENDED_JOBS_KEPT) break;
    if (!entry.settled) continue;
    trackedJobs.delete(jobId);
    ended -= 1;
  }
}

function settle(entry: TrackedPdfConversionJob): void {
  if (entry.settled || !isPdfConversionJobEnded(entry.job)) return;
  entry.settled = true;
  const job = entry.job;
  const generation = trackerGeneration;
  // After the caller's own step: a job that is over the moment it is
  // tracked still reaches its handler after the caller recorded it.
  queueMicrotask(() => {
    if (generation !== trackerGeneration) return;
    const handlers = entry.handlers;
    const foreground = handlers.find((handler) => handler.isForeground());
    const chosen =
      foreground ??
      [...handlers].reverse().find((handler) => handler.isMounted()) ??
      handlers[handlers.length - 1];
    try {
      chosen?.settle(job, foreground !== undefined);
    } catch {
      // The handler's own failure is its own.
    }
  });
  pruneEndedJobs();
}

function failLocally(
  entry: TrackedPdfConversionJob,
  error: PdfConversionJobError,
): void {
  if (isPdfConversionJobEnded(entry.job)) return;
  entry.job = { ...entry.job, status: "failed", error };
  notify(entry);
  settle(entry);
}

function delay(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

/** A failure the server meant (the job is not there for us), not a hiccup. */
function isFinalPollRefusal(error: unknown): boolean {
  return (
    isApiRequestError(error) &&
    error.status >= 400 &&
    error.status < 500 &&
    error.status !== 408 &&
    error.status !== 429
  );
}

async function pollUntilEnded(entry: TrackedPdfConversionJob): Promise<void> {
  if (entry.polling) return;
  entry.polling = true;
  const generation = trackerGeneration;
  let failingSince: number | null = null;
  try {
    while (generation === trackerGeneration && !isPdfConversionJobEnded(entry.job)) {
      const askedAt = Date.now();
      const controller = new AbortController();
      const timeout = setTimeout(
        () => controller.abort(),
        PDF_CONVERSION_POLL_TIMEOUT_MS,
      );
      try {
        const next = await requestPdfConversionJob(
          entry.job.resource_id,
          entry.job.job_id,
          {
            waitMs: PDF_CONVERSION_POLL_WAIT_MS,
            since: entry.job.revision,
            signal: controller.signal,
          },
        );
        if (generation !== trackerGeneration) return;
        failingSince = null;
        if (accept(entry, next)) {
          settle(entry);
        } else if (Date.now() - askedAt < PDF_CONVERSION_MIN_POLL_INTERVAL_MS) {
          // A server that answers at once without news is not long-polling;
          // do not spin on it.
          await delay(PDF_CONVERSION_RETRY_DELAY_MS);
        }
      } catch (error) {
        if (generation !== trackerGeneration) return;
        if (isPdfConversionJobEnded(entry.job)) break;
        if (isFinalPollRefusal(error)) {
          failLocally(entry, {
            code: PDF_CONVERSION_UNREACHABLE_CODE,
            message: "",
            status: isApiRequestError(error) ? error.status : 0,
          });
          break;
        }
        const now = Date.now();
        failingSince ??= now;
        if (now - failingSince >= PDF_CONVERSION_GIVE_UP_MS) {
          failLocally(entry, {
            code: PDF_CONVERSION_UNREACHABLE_CODE,
            message: "",
            status: 0,
          });
          break;
        }
        await delay(PDF_CONVERSION_RETRY_DELAY_MS);
      } finally {
        clearTimeout(timeout);
      }
    }
  } finally {
    entry.polling = false;
  }
}

/**
 * Follow `job` until it ends (a job already followed takes the newer news),
 * and register `handler` for its result. Returns the job as followed now.
 */
export function trackPdfConversionJob(
  job: PdfConversionJob,
  handler?: PdfConversionJobSettleHandler,
): PdfConversionJob {
  let entry = trackedJobs.get(job.job_id);
  if (!entry) {
    entry = {
      job,
      listeners: new Set(),
      handlers: [],
      settled: false,
      polling: false,
    };
    trackedJobs.set(job.job_id, entry);
    notifyChange();
  } else {
    accept(entry, job);
  }
  if (handler && !entry.handlers.includes(handler)) entry.handlers.push(handler);
  if (isPdfConversionJobEnded(entry.job)) settle(entry);
  else void pollUntilEnded(entry);
  return entry.job;
}

/** Hear every change of one followed job; returns the unsubscribe. */
export function subscribePdfConversionJob(
  jobId: string,
  listener: PdfConversionJobListener,
): () => void {
  const entry = trackedJobs.get(jobId);
  if (!entry) return () => undefined;
  entry.listeners.add(listener);
  return () => {
    entry.listeners.delete(listener);
  };
}

/** Hear that some followed job started, changed or ended. */
export function subscribePdfConversionJobs(listener: () => void): () => void {
  changeListeners.add(listener);
  return () => {
    changeListeners.delete(listener);
  };
}

/** The followed job as known now (the same object until it changes). */
export function getTrackedPdfConversionJob(
  jobId: string,
): PdfConversionJob | null {
  return trackedJobs.get(jobId)?.job ?? null;
}

/** The newest job of this resource that has not ended, if one is followed. */
export function getActivePdfConversionJobForResource(
  resourceId: string,
): PdfConversionJob | null {
  let newest: PdfConversionJob | null = null;
  for (const { job } of trackedJobs.values()) {
    if (job.resource_id !== resourceId || isPdfConversionJobEnded(job)) continue;
    if (!newest || job.created_at >= newest.created_at) newest = job;
  }
  return newest;
}

/** Ask the server to stop a followed job; it is followed until it stops. */
export async function cancelTrackedPdfConversionJob(
  jobId: string,
): Promise<PdfConversionJob | null> {
  const entry = trackedJobs.get(jobId);
  if (!entry) return null;
  if (isPdfConversionJobEnded(entry.job)) return entry.job;
  const next = await requestPdfConversionCancel(entry.job.resource_id, jobId);
  if (accept(entry, next)) settle(entry);
  return entry.job;
}

/**
 * Follow again this person's running conversions of a resource (after its
 * window reopened, or the page reloaded): the server lists them, and each
 * gets the handler `handlerFor` gives it.
 */
export async function reattachPdfConversionJobs(
  resourceId: string,
  handlerFor?: (job: PdfConversionJob) => PdfConversionJobSettleHandler,
): Promise<PdfConversionJob[]> {
  let flight = reattachFlights.get(resourceId);
  if (!flight) {
    flight = listPdfConversionJobs(resourceId, { activeOnly: true }).finally(
      () => {
        reattachFlights.delete(resourceId);
      },
    );
    reattachFlights.set(resourceId, flight);
  }
  const generation = trackerGeneration;
  const jobs = await flight;
  if (generation !== trackerGeneration) return [];
  return jobs
    .filter((job) => !isPdfConversionJobEnded(job))
    .map((job) => trackPdfConversionJob(job, handlerFor?.(job)));
}

/** Forget every followed job and stop their loops (tests only). */
export function resetPdfConversionJobsForTests(): void {
  trackerGeneration += 1;
  trackedJobs.clear();
  changeListeners.clear();
  reattachFlights.clear();
}
