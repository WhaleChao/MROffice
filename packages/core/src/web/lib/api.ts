/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
// MR Note's service API, as the open-source editors see it: there is no
// server. Every request is answered in this browser by a local answerer (the
// editors' own browser-only hosts register them, as does the host adapter's
// engine host), or refused with a 503 "host_unavailable" that the editors
// already handle as "this needs the service".

export * from "@mrnote/core/web/__stubs__/lib/api.ts";

import { noteDesktopFeature } from "@mrnote/core/host/desktop-app.ts";

// A request refused here is a feature of MR Note's service: it is provided
// by the MR Office desktop app's installer (free: the full service),
// https://github.com/XinaTian/MROffice/releases. The
// person is told so (with the download link) instead of the button doing
// nothing.
const noteNeedsService = () => noteDesktopFeature();

export class ApiRequestError extends Error {
  readonly status: number;
  readonly code?: string;
  readonly details?: Record<string, unknown>;
  readonly retryAfterMs?: number;

  constructor(
    message: string,
    options: { status: number; code?: string; details?: Record<string, unknown>; retryAfterMs?: number },
  ) {
    super(message);
    this.name = "ApiRequestError";
    this.status = options.status;
    this.code = options.code;
    this.details = options.details;
    this.retryAfterMs = options.retryAfterMs;
  }
}

export function isApiRequestError(error: unknown): error is ApiRequestError {
  return error instanceof ApiRequestError;
}

export type LocalApiAnswerer = (request: {
  readonly method: string;
  readonly path: string;
  readonly body: unknown;
}) => Promise<unknown> | undefined;

const answerers = new Set<LocalApiAnswerer>();
const firstAnswerers = new Set<LocalApiAnswerer>();

/** Claim requests for this browser to answer; returns the release. */
export function registerLocalApiAnswerer(answerer: LocalApiAnswerer, options: { first?: boolean } = {}): () => void {
  (options.first ? firstAnswerers : answerers).add(answerer);
  return () => {
    firstAnswerers.delete(answerer);
    answerers.delete(answerer);
  };
}

function parseBody(body: unknown): unknown {
  if (typeof body !== "string") return body;
  try {
    return JSON.parse(body);
  } catch {
    return body;
  }
}

function answer(path: string, method: string, body: unknown): Promise<unknown> | undefined {
  const request = { method: method.toUpperCase(), path, body: parseBody(body) };
  for (const answerer of [...firstAnswerers, ...answerers]) {
    const answered = answerer(request);
    if (answered) return answered;
  }
  return undefined;
}

export function answerApiRequestLocally(path: string, method: string, body?: unknown): Promise<unknown> | undefined {
  return answer(path, method, body);
}

/** A request the editors make on their own (a gallery's previews fetched in
 * the background, a soft availability probe), not for something the person
 * asked for: refused without telling them. */
const BACKGROUND_REQUEST = /[?&]busy=soft\b|\/previews(\?|$)|\/capabilit(y|ies)(\?|\/|$)|-availability(\?|$)|\/office-data\//;

/** Whether the person has just acted (a click or a key the browser counts as
 * activation). A request an editor makes on its own as it opens -- reading its
 * font registry, say -- has none: telling the person then put "This is in the
 * MR Office desktop app" over every new document before they had done
 * anything. A browser that cannot tell keeps the old behaviour. */
function personJustActed(): boolean {
  const activation = typeof navigator !== "undefined" ? (navigator as Navigator & { userActivation?: { isActive: boolean } }).userActivation : undefined;
  return activation ? activation.isActive : true;
}

function unavailable(path: string): ApiRequestError {
  if (!BACKGROUND_REQUEST.test(path) && personJustActed()) noteNeedsService();
  if (typeof localStorage !== "undefined") {
    try {
      if (localStorage.getItem("mrnote-editors:trace-inert") === "1") console.debug("[api] not available:", path);
    } catch {
      // storage unavailable
    }
  }
  return new ApiRequestError("This needs the MR Note service, which the open-source editors do not include.", {
    status: 503,
    code: "host_unavailable",
  });
}

/** Not the service's own path: a file's URL (blob:, data:, a static asset). */
function isPlainUrl(path: string): boolean {
  return !path.startsWith("/api/") && !/^https?:\/\/[^/]+\/api\//.test(path);
}

async function request<T>(path: string, init: RequestInit = {}): Promise<T> {
  const local = answer(path, init.method || "GET", init.body);
  if (!local && isPlainUrl(path)) {
    const response = await fetch(path, { ...init, credentials: "omit" });
    if (!response.ok) throw new ApiRequestError(`HTTP ${response.status}`, { status: response.status });
    return ((response.headers.get("content-type") ?? "").includes("json") ? await response.json() : await response.arrayBuffer()) as T;
  }
  if (!local) throw unavailable(path);
  const value = await local;
  if (value instanceof Response) {
    if (!value.ok) throw new ApiRequestError(`HTTP ${value.status}`, { status: value.status });
    const type = value.headers.get("content-type") ?? "";
    return (type.includes("json") ? await value.json() : await value.arrayBuffer()) as T;
  }
  return value as T;
}

export function apiGet<T>(path: string, init?: RequestInit): Promise<T> {
  return request<T>(path, { ...init, method: "GET" });
}

export function apiPost<T>(path: string, body?: unknown, init?: RequestInit): Promise<T> {
  return request<T>(path, { ...init, method: "POST", body: body as BodyInit });
}

export function apiPut<T>(path: string, body?: unknown, init?: RequestInit): Promise<T> {
  return request<T>(path, { ...init, method: "PUT", body: body as BodyInit });
}

export function apiPatch<T>(path: string, body?: unknown, init?: RequestInit): Promise<T> {
  return request<T>(path, { ...init, method: "PATCH", body: body as BodyInit });
}

export function apiDelete<T>(path: string, init?: RequestInit): Promise<T> {
  return request<T>(path, { ...init, method: "DELETE" });
}

export function apiPostFormData<T>(path: string, body: FormData, init?: RequestInit): Promise<T> {
  return request<T>(path, { ...init, method: "POST", body });
}

export async function apiFetchRaw(path: string, init: RequestInit = {}): Promise<Response> {
  const local = answer(path, init.method || "GET", init.body);
  if (!local && isPlainUrl(path)) return fetch(path, { ...init, credentials: "omit" });
  if (!local) throw unavailable(path);
  const value = await local;
  return value instanceof Response
    ? value
    : new Response(JSON.stringify(value ?? null), { headers: { "Content-Type": "application/json" } });
}

export function localizedApiErrorCodeMessage(_code: unknown): string | null {
  return null;
}

export function personalWorkspaceId(): string | null {
  return null;
}

export function activeWorkspaceHeaders(): Record<string, string> {
  return {};
}

export function personalWorkspaceHeaders(): Record<string, string> {
  return {};
}

export function readAvailableApiCsrfToken(): string | null {
  return null;
}

export async function fetchWorkspaceScoped(path: string, init?: RequestInit): Promise<Response> {
  return apiFetchRaw(path, init);
}

export async function uploadToPresignedUrl(): Promise<never> {
  throw unavailable("upload");
}
