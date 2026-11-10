"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The browser's one way to an engine that runs on the server.
 *
 * The editors' algorithmic engines (SmartArt layout first) run in the API's
 * Node workers (ENGINE-HOST-CONTRACT.md); the browser sends a method's input
 * to `POST /api/v1/engine/{engine}/{method}` and gets its JSON result. Every
 * engine caller goes through {@link callEngine} (or the hook in
 * `./use-engine-call`), so all of them share:
 *
 * - the session + CSRF, with a guest retry when there is no session (the
 *   marketing guest editor; `lib/document-export/docx-server-export.ts`'s
 *   pattern);
 * - one closed error vocabulary ({@link EngineUnavailableError} with a
 *   reason, {@link EngineRefusedError} with the engine's own reason), and the
 *   offline case answered at once instead of after a timeout;
 * - an input-hash cache for pure methods and one shared request for identical
 *   calls in flight;
 * - abort through the caller's signal.
 */

import { apiFetchRaw, isApiRequestError } from "../api";
import { EngineBlobAnswerError, readEngineBlobAnswer } from "./blob-answer";
import type {
  EngineInput,
  EngineMethodName,
  EngineName,
  EngineOutput,
  EngineSessionInput,
  EngineSessionMethodName,
  EngineSessionName,
  EngineSessionOpenInput,
  EngineSessionOpenOutput,
  EngineSessionOutput,
} from "./methods";

export type {
  EngineInput,
  EngineMethodName,
  EngineName,
  EngineOutput,
  EngineSessionInput,
  EngineSessionMethodName,
  EngineSessionName,
  EngineSessionOpenInput,
  EngineSessionOpenOutput,
  EngineSessionOutput,
} from "./methods";

export const ENGINE_HOST_PATH = "/api/v1/engine";

export type EngineUnavailableReason =
  | "offline"
  | "busy"
  | "timeout"
  | "rate_limited"
  | "signed_out"
  | "failed";

/** The engine could not answer (no connection, busy, failed); nothing changed. */
export class EngineUnavailableError extends Error {
  readonly reason: EngineUnavailableReason;
  readonly retryAfterMs?: number;
  /** The HTTP status and error code behind a "failed" answer, when there
   * was one: a reload of the API answers a call 404/500/503 for a moment,
   * and a caller may ask again (pdf-native-remote). */
  status?: number;
  code?: string;

  constructor(reason: EngineUnavailableReason, detail?: string, retryAfterMs?: number) {
    super(detail ? `engine ${reason}: ${detail}` : `engine ${reason}`);
    this.name = "EngineUnavailableError";
    this.reason = reason;
    this.retryAfterMs = retryAfterMs;
  }
}

/** The engine refused this input; `reason` is the engine's closed code. */
export class EngineRefusedError extends Error {
  readonly reason: string;
  readonly details: Record<string, unknown>;

  constructor(reason: string, details: Record<string, unknown> = {}) {
    super(`engine refused: ${reason}`);
    this.name = "EngineRefusedError";
    this.reason = reason;
    this.details = details;
  }
}

/**
 * The session a call named is gone on the server (idle expiry, a recycled
 * worker, another API process, a quota): re-open with the snapshot and retry
 * once (ENGINE-HOST-CONTRACT 12.2; `./session.ts` does this).
 */
export class EngineSessionUnknownError extends Error {
  constructor() {
    super("engine session unknown");
    this.name = "EngineSessionUnknownError";
  }
}

/** The session's call was cancelled (an explicit cancel, or a newer call
 * that superseded it); the session itself is intact. */
export class EngineCallCancelledError extends Error {
  constructor() {
    super("engine call cancelled");
    this.name = "EngineCallCancelledError";
  }
}

export function isEngineAbort(error: unknown): boolean {
  return error instanceof DOMException && error.name === "AbortError";
}

/** The message key for an engine failure (`console-notebooks.json`). */
export function engineErrorMessageKey(
  error: unknown,
): "engineHost.needsConnection" | "engineHost.busy" | "engineHost.failed" | "engineHost.refused" {
  if (error instanceof EngineRefusedError) return "engineHost.refused";
  if (error instanceof EngineUnavailableError) {
    if (error.reason === "offline") return "engineHost.needsConnection";
    if (error.reason === "busy" || error.reason === "rate_limited") return "engineHost.busy";
  }
  return "engineHost.failed";
}

function unavailable(error: unknown): Error {
  if (isEngineAbort(error)) return error as Error;
  if (isApiRequestError(error)) {
    if (error.status === 0 || error.code === "network_unreachable") return new EngineUnavailableError("offline");
    if (error.code === "engine_invalid_input") {
      const details = { ...(error.details ?? {}) };
      const reason = typeof details.reason === "string" ? details.reason : "invalid_input";
      delete details.reason;
      return new EngineRefusedError(reason, details);
    }
    if (error.code === "payload_too_large") return new EngineRefusedError("too_large");
    // The server never saw the bytes a call named (12.1): the caller may ask
    // once more with them (`upload`).
    if (error.code === "engine_blob_unknown") return new EngineRefusedError("blob_unknown", { ...(error.details ?? {}) });
    if (error.code === "engine_session_unknown") return new EngineSessionUnknownError();
    if (error.code === "engine_cancelled") return new EngineCallCancelledError();
    if (error.status === 401 || error.status === 403) return new EngineUnavailableError("signed_out");
    if (error.status === 429) return new EngineUnavailableError("rate_limited", undefined, error.retryAfterMs);
    if (error.code === "engine_busy") return new EngineUnavailableError("busy", undefined, error.retryAfterMs ?? 1000);
    if (error.code === "engine_timeout") return new EngineUnavailableError("timeout");
    const failed = new EngineUnavailableError("failed", `${error.status} ${error.code ?? ""}`.trim());
    failed.status = error.status;
    if (error.code) failed.code = error.code;
    return failed;
  }
  if (error instanceof TypeError) return new EngineUnavailableError("offline", error.message);
  return new EngineUnavailableError("failed", error instanceof Error ? error.message : undefined);
}

/** A caller with no session (the CSRF bootstrap refused, or the session did). */
function noSession(error: unknown): boolean {
  if (isEngineAbort(error)) return false;
  if (isApiRequestError(error)) return error.status === 401;
  return error instanceof Error && !(error instanceof TypeError);
}

/** The API names a session that holds no CSRF token yet: ask signed in. */
function needsSession(error: unknown): boolean {
  return isApiRequestError(error) && error.status === 403 && /csrf/iu.test(error.code ?? "");
}

// Learned once per page: a guest's calls skip the CSRF bootstrap that cannot
// succeed for them, and a guest who signs in is caught by the API's CSRF
// refusal and asked again with the session.
let guestMode = false;

/** A request body: JSON, or JSON plus one uploaded blob (12.1). */
interface EngineRequestBody {
  readonly body: BodyInit;
  readonly contentType: string;
  readonly headers?: Record<string, string>;
  /** Session control: `DELETE` closes a session (12.2). */
  readonly method?: "POST" | "DELETE";
  /** Outlives the page (closing a session from `pagehide`). */
  readonly keepalive?: boolean;
}

/** The body that carries one blob: `[u32 BE JSON length][JSON][bytes]`. */
export const ENGINE_BLOB_BODY_TYPE = "application/x-mrai-engine-call";

/** Several blobs in one body (12.1): the JSON names them in order
 * (`uploads: [{sha256, size, role}]`), their bytes follow. `role: "source"`
 * is a working head the server never saw; `"input"` the method's other
 * inputs (a picture to embed), which live only for this call. */
function blobsBody(input: unknown, uploads: readonly EngineBlobUploadPart[]): EngineRequestBody {
  if (!input || typeof input !== "object" || Array.isArray(input)) {
    throw new EngineUnavailableError("failed", "uploads need an object input");
  }
  const text = new TextEncoder().encode(JSON.stringify({
    ...(input as Record<string, unknown>),
    uploads: uploads.map((upload) => ({ sha256: upload.sha256, size: upload.bytes.byteLength, role: upload.role })),
  }));
  const total = uploads.reduce((sum, upload) => sum + upload.bytes.byteLength, 0);
  const framed = new Uint8Array(4 + text.byteLength + total);
  new DataView(framed.buffer).setUint32(0, text.byteLength, false);
  framed.set(text, 4);
  let offset = 4 + text.byteLength;
  for (const upload of uploads) {
    framed.set(upload.bytes, offset);
    offset += upload.bytes.byteLength;
  }
  return { body: framed, contentType: ENGINE_BLOB_BODY_TYPE };
}

function blobBody(json: string, upload: EngineBlobUpload): EngineRequestBody {
  const text = new TextEncoder().encode(json);
  const framed = new Uint8Array(4 + text.byteLength + upload.bytes.byteLength);
  new DataView(framed.buffer).setUint32(0, text.byteLength, false);
  framed.set(text, 4);
  framed.set(upload.bytes, 4 + text.byteLength);
  return { body: framed, contentType: ENGINE_BLOB_BODY_TYPE, headers: {
    "X-Mrai-Blob-Sha256": upload.sha256,
    ...(upload.base ? { "X-Mrai-Blob-Base": `${upload.base.sha256}:${upload.base.size}` } : {}),
  } };
}









const SOFT_BUSY_QUERY = "busy=soft";

function softBusyPath(path: string): string {
  return `${path}${path.includes("?") ? "&" : "?"}${SOFT_BUSY_QUERY}`;
}

function retryAfterHeaderMs(response: Response): number | undefined {
  const seconds = Number(response.headers.get("retry-after"));
  return Number.isFinite(seconds) && seconds >= 0 ? seconds * 1000 : undefined;
}

/** The busy refusal a 202 carries, or null for any other 2xx. */
async function softBusyRefusal(response: Response): Promise<EngineUnavailableError | null> {
  if (response.status !== 202) return null;
  const data = (await response.clone().json().catch(() => null)) as { error?: { code?: unknown } } | null;
  if (data?.error?.code !== "engine_busy") return null;
  return new EngineUnavailableError("busy", undefined, retryAfterHeaderMs(response) ?? 1000);
}

function post(path: string, body: EngineRequestBody, signal: AbortSignal | undefined, signedIn: boolean): Promise<Response> {
  return apiFetchRaw(
    (body.method ?? "POST") === "POST" ? softBusyPath(path) : path,
    {
      method: body.method ?? "POST", body: body.body, signal,
      ...(body.keepalive ? { keepalive: true } : {}),
      ...(body.headers ? { headers: body.headers } : {}),
    },
    {
      contentType: body.contentType,
      // An editor command never bounces the page to the login screen.
      suppressUnauthorizedHandling: true,
      ...(signedIn ? {} : { requireCsrf: false }),
    },
  );
}

/** Engine requests on the wire right now. The engine host serves one at a
 * time per caller (a second is refused `engine_busy`): the homepage demo
 * waits for this to reach zero before it asks for its own. */
let requestsOnTheWire = 0;
export function engineRequestsInFlight(): number {
  return requestsOnTheWire;
}

async function request(path: string, json: string | EngineRequestBody, signal: AbortSignal | undefined): Promise<unknown> {
  requestsOnTheWire += 1;
  try {
    return await requestOnce(path, json, signal);
  } finally {
    requestsOnTheWire -= 1;
  }
}

async function requestOnce(path: string, json: string | EngineRequestBody, signal: AbortSignal | undefined): Promise<unknown> {
  const body: EngineRequestBody = typeof json === "string" ? { body: json, contentType: "application/json" } : json;
  let raw: Response;
  try {
    if (guestMode) {
      try {
        raw = await post(path, body, signal, false);
      } catch (error) {
        if (!needsSession(error)) throw error;
        guestMode = false;
        raw = await post(path, body, signal, true);
      }
    } else {
      try {
        raw = await post(path, body, signal, true);
      } catch (error) {
        if (!noSession(error)) throw error;
        guestMode = true;
        raw = await post(path, body, signal, false);
      }
    }
    const busy = await softBusyRefusal(raw);
    if (busy) throw busy;
    // A method that answers whole files (binary_answer): the framed body.
    if ((raw.headers.get("content-type") ?? "").split(";", 1)[0].trim().toLowerCase() === ENGINE_BLOB_BODY_TYPE) {
      return await readEngineBlobAnswer(await raw.arrayBuffer());
    }
    const answer = (await raw.json()) as { result?: unknown };
    if (!answer || typeof answer !== "object" || !("result" in answer)) {
      throw new EngineUnavailableError("failed", "malformed response");
    }
    return answer.result;
  } catch (error) {
    if (
      error instanceof EngineUnavailableError || error instanceof EngineRefusedError ||
      error instanceof EngineSessionUnknownError || error instanceof EngineCallCancelledError
    ) throw error;
    if (error instanceof EngineBlobAnswerError) throw new EngineUnavailableError("failed", error.message);
    throw unavailable(error);
  }
}

/** JSON with every object's keys sorted: one spelling per value, for the hash. */
export function canonicalJson(value: unknown): string {
  if (value === null || typeof value !== "object") return JSON.stringify(value) ?? "null";
  if (Array.isArray(value)) return `[${value.map((item) => (item === undefined ? "null" : canonicalJson(item))).join(",")}]`;
  const record = value as Record<string, unknown>;
  const keys = Object.keys(record).filter((key) => record[key] !== undefined).sort();
  return `{${keys.map((key) => `${JSON.stringify(key)}:${canonicalJson(record[key])}`).join(",")}}`;
}

async function sha256Hex(text: string): Promise<string | null> {
  const subtle = typeof crypto !== "undefined" ? crypto.subtle : undefined;
  if (!subtle) return null;
  const digest = await subtle.digest("SHA-256", new TextEncoder().encode(text));
  return Array.from(new Uint8Array(digest), (byte) => byte.toString(16).padStart(2, "0")).join("");
}

const CACHE_ENTRIES = 256;
const cache = new Map<string, unknown>();
const inFlight = new Map<string, Promise<unknown>>();

function remember(key: string, value: unknown): void {
  cache.delete(key);
  cache.set(key, value);
  while (cache.size > CACHE_ENTRIES) cache.delete(cache.keys().next().value as string);
}

/** Forget every cached engine answer (tests; a sign-out). */
export function clearEngineCache(): void {
  cache.clear();
  inFlight.clear();
}

/** Bytes a call sends along because the server answered `blob_unknown`
 * for them (a working head it never saw). */
export interface EngineBlobUpload {
  readonly sha256: string;
  readonly bytes: Uint8Array;
  /** `bytes` is only what was appended to this base (a PDF working head is
   * its predecessor plus an incremental update): the server puts a base it
   * holds in front and checks `sha256` on the whole, or answers
   * `blob_unknown` (then the whole head goes). */
  readonly base?: Readonly<{ sha256: string; size: number }>;
}

/** One blob of a several-blob body (`CallEngineOptions.uploads`). */
export interface EngineBlobUploadPart extends EngineBlobUpload {
  readonly role: "source" | "input";
}

export interface CallEngineOptions {
  signal?: AbortSignal;
  /** Several blobs with the request (methods whose policy declares
   * `uploads`): the working head (`role: "source"`, at most one) and inputs. */
  uploads?: readonly EngineBlobUploadPart[];
  /** Reuse an earlier answer to the same input. Only for pure methods. */
  cache?: boolean;
  /** One blob sent with the request (methods whose policy allows it). */
  upload?: EngineBlobUpload;
}

function abortable<T>(promise: Promise<T>, signal: AbortSignal | undefined): Promise<T> {
  if (!signal) return promise;
  if (signal.aborted) return Promise.reject(new DOMException("The operation was aborted", "AbortError"));
  return new Promise<T>((resolve, reject) => {
    const onAbort = () => reject(new DOMException("The operation was aborted", "AbortError"));
    signal.addEventListener("abort", onAbort, { once: true });
    promise.then(
      (value) => {
        signal.removeEventListener("abort", onAbort);
        resolve(value);
      },
      (error) => {
        signal.removeEventListener("abort", onAbort);
        reject(error);
      },
    );
  });
}

/**
 * Run one engine method on the server.
 *
 * Rejects with {@link EngineUnavailableError} (offline, busy, timeout, rate
 * limited, signed out, failed), {@link EngineRefusedError} (the engine's own
 * closed reason) or the caller's `AbortError`. With `cache`, an identical
 * input answers from memory and identical calls in flight share one request
 * (an abort then only detaches that caller).
 */
export async function callEngine<E extends EngineName, M extends EngineMethodName<E>>(
  engine: E,
  method: M,
  input: EngineInput<E, M>,
  options: CallEngineOptions = {},
): Promise<EngineOutput<E, M>> {
  if (typeof navigator !== "undefined" && navigator.onLine === false) {
    throw new EngineUnavailableError("offline");
  }
  const path = `${ENGINE_HOST_PATH}/${encodeURIComponent(engine)}/${encodeURIComponent(method)}`;
  if (options.uploads?.length) {
    return (await request(path, blobsBody(input, options.uploads), options.signal)) as EngineOutput<E, M>;
  }
  if (options.upload) {
    return (await request(path, blobBody(JSON.stringify(input), options.upload), options.signal)) as EngineOutput<E, M>;
  }
  const body = options.cache ? canonicalJson(input) : JSON.stringify(input);
  if (!options.cache) return (await request(path, body, options.signal)) as EngineOutput<E, M>;
  const hash = await sha256Hex(`${engine}\u0000${method}\u0000${body}`);
  if (hash === null) return (await request(path, body, options.signal)) as EngineOutput<E, M>;
  if (cache.has(hash)) {
    const value = cache.get(hash);
    remember(hash, value);
    return value as EngineOutput<E, M>;
  }
  let shared = inFlight.get(hash);
  if (!shared) {
    // The shared request is not tied to any one caller's signal: one caller
    // aborting must not fail the others waiting on the same answer.
    shared = request(path, body, undefined).then(
      (value) => {
        inFlight.delete(hash);
        remember(hash, value);
        return value;
      },
      (error: unknown) => {
        inFlight.delete(hash);
        throw error;
      },
    );
    inFlight.set(hash, shared);
  }
  return (await abortable(shared, options.signal)) as EngineOutput<E, M>;
}

// -- soft sessions (ENGINE-HOST-CONTRACT 12.2) ---------------------------------
//
// The transport only; `./session.ts` keeps a session open and re-opens it
// when the server lost it. Never cached: a session call reads server state.

function offlineCheck(): void {
  if (typeof navigator !== "undefined" && navigator.onLine === false) throw new EngineUnavailableError("offline");
}

const sessionPath = (engine: string, sessionId?: string, tail?: string) =>
  `${ENGINE_HOST_PATH}/${encodeURIComponent(engine)}/sessions` +
  (sessionId ? `/${encodeURIComponent(sessionId)}` : "") + (tail ? `/${encodeURIComponent(tail)}` : "");

export interface OpenEngineSessionOptions {
  signal?: AbortSignal;
  /** What the caller's per-scope quota counts (the page id). */
  scope?: string;
}

/** Opens a session from `snapshot`; answers its id and the engine's `opened`. */
export async function openEngineSession<E extends EngineSessionName>(
  engine: E,
  snapshot: EngineSessionOpenInput<E>,
  options: OpenEngineSessionOptions = {},
): Promise<{ sessionId: string; opened: EngineSessionOpenOutput<E> }> {
  offlineCheck();
  const query = options.scope ? `?scope=${encodeURIComponent(options.scope)}` : "";
  const answer = (await request(`${sessionPath(engine)}${query}`, JSON.stringify(snapshot), options.signal)) as {
    session?: unknown;
    opened?: unknown;
  } | null;
  if (!answer || typeof answer.session !== "string") throw new EngineUnavailableError("failed", "malformed session");
  return { sessionId: answer.session, opened: (answer.opened ?? null) as EngineSessionOpenOutput<E> };
}

/** One method against the open session (rejects with
 * {@link EngineSessionUnknownError} when the server lost it). */
export async function callEngineSession<E extends EngineSessionName, M extends EngineSessionMethodName<E>>(
  engine: E,
  sessionId: string,
  method: M,
  input: EngineSessionInput<E, M>,
  options: { signal?: AbortSignal } = {},
): Promise<EngineSessionOutput<E, M>> {
  offlineCheck();
  return (await request(sessionPath(engine, sessionId, method), JSON.stringify(input), options.signal)) as EngineSessionOutput<E, M>;
}

/** Stops the session's running call at its next yield. Never throws. */
export async function cancelEngineSession(engine: EngineSessionName, sessionId: string): Promise<boolean> {
  try {
    offlineCheck();
    const answer = (await request(sessionPath(engine, sessionId, "cancel"), "{}", undefined)) as { cancelled?: unknown } | null;
    return answer?.cancelled === true;
  } catch {
    return false;
  }
}

/** Frees the session on the server (`keepalive` from `pagehide`). Never throws. */
export async function closeEngineSession(
  engine: EngineSessionName,
  sessionId: string,
  options: { keepalive?: boolean } = {},
): Promise<boolean> {
  try {
    offlineCheck();
    const answer = (await request(
      sessionPath(engine, sessionId),
      { body: "", contentType: "application/json", method: "DELETE", keepalive: options.keepalive },
      undefined,
    )) as { closed?: unknown } | null;
    return answer?.closed === true;
  } catch {
    return false;
  }
}
