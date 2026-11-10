/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { getFileTransferOrigin } from "./private-file-transfer";
import { noteAccessGateRequired } from "./access-gate";

/** `ApiRequestError.code` of a file transfer the file host refused for its
 * own gate session and that could not be repaired. */
export const FILE_ACCESS_REQUIRED_CODE = "file_access_required";
export const FILE_ACCESS_REQUIRED_EVENT = "mrnote:file-access-required";
export const FILE_ACCESS_RESTORED_EVENT = "mrnote:file-access-restored";

/** The gate's plain-text refusals on the file host (server.mjs `/verify`). */
export const FILE_GATE_ACCESS_REQUIRED_BODY = "Access required";
export const FILE_GATE_SIGN_IN_REQUIRED_BODY = "Sign-in required";

/** How long after an automatic reconnect another one is not attempted. */
export const RECONNECT_GUARD_MS = 2 * 60_000;
const RECONNECT_GUARD_KEY = "mrnote:file-access-reconnected-at";

/**
 * The probe: an allowlisted file route (private-file-routes.json) that the
 * file host's `/verify` authenticates like any transfer, for an attachment
 * that cannot exist. The gate answers its refusal before the product is
 * reached; past the gate the API answers 404 and nothing is written.
 */
export const FILE_ACCESS_PROBE_PATH =
  "/api/v1/attachments/00000000-0000-0000-0000-000000000000/download";
const BRIDGE_PATH = "/__beta/connect";
const HANDOFF_PATH = "/__beta/file-connect";
/** The static page that only says the original window works again. */
const RECONNECTED_LANDING_PATH = "/access-restored.html";
const PROBE_TIMEOUT_MS = 10_000;
const HANDOFF_TIMEOUT_MS = 8_000;
 
const HANDOFF_PATTERN = /^[A-Za-z0-9_-]{1,3955}\.[A-Za-z0-9_-]{43}$/u;

/** What the file host said about this browser's file session. */
export type FileAccessState =
  | "ok"
  | "access-missing"
  | "sign-in-missing"
  | "unreachable"
  | "not-applicable";

export type FileAccessRepairOutcome =
  | "repaired"
  /** An automatic reconnect already ran in this tab a moment ago. */
  | "guarded"
  /** The site's own gate pass is gone too (lib/access-gate.ts takes over). */
  | "gate-locked"
   
  | "handoff-unavailable"
   
  | "refused"
  | "unreachable"
  | "not-applicable";

export type FileAccessCheck = FileAccessRepairOutcome | Exclude<FileAccessState, "access-missing">;

/** Which of the gate's refusals a file-host response is, if any. */
export type FileGateRefusal = "access" | "sign-in" | null;

let required = false;
let lastRequiredNotice = 0;
let repairFlight: Promise<FileAccessRepairOutcome> | null = null;
let pageLoadCheck: Promise<FileAccessCheck> | null = null;

/**
 * Whether this page may repair the file session at all: a browser top-level
 * window of a deployment that has a private file host. A frame (an embed,
 * a render or print frame) never does; the window that holds it does.
 */
export function fileAccessRepairApplies(): boolean {
  if (typeof window === "undefined" || typeof fetch !== "function") return false;
  try {
    if (window.top !== window.self) return false;
  } catch {
    return false;
  }
  return configuredFileOrigin() !== null;
}

function configuredFileOrigin(): string | null {
  try {
    return getFileTransferOrigin();
  } catch {
    
    return null;
  }
}

/** Read which gate refusal a response from the file host is, without
 * consuming the caller's body. */
export async function readFileGateRefusal(response: Response): Promise<FileGateRefusal> {
  if (response.status !== 401) return null;
  const length = Number(response.headers?.get("content-length") ?? "");
  if (Number.isFinite(length) && length > 64) return null;
  let text: string;
  try {
    const copy = typeof response.clone === "function" ? response.clone() : response;
    text = await copy.text();
  } catch {
    return null;
  }
  const body = text.trim();
  if (body === FILE_GATE_ACCESS_REQUIRED_BODY) return "access";
  if (body === FILE_GATE_SIGN_IN_REQUIRED_BODY) return "sign-in";
  return null;
}

function timeoutSignal(ms: number): { signal: AbortSignal; done: () => void } {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), ms);
  return { signal: controller.signal, done: () => clearTimeout(timer) };
}

/** Ask the file host whether this browser's file session passes its gate. */
export async function probeFileAccess(): Promise<FileAccessState> {
  if (!fileAccessRepairApplies()) return "not-applicable";
  const origin = configuredFileOrigin();
  if (!origin) return "not-applicable";
  const deadline = timeoutSignal(PROBE_TIMEOUT_MS);
  try {
    // No custom header: a simple CORS request, no preflight.
    const response = await fetch(`${origin}${FILE_ACCESS_PROBE_PATH}`, {
      method: "GET",
      credentials: "include",
      cache: "no-store",
      redirect: "error",
      signal: deadline.signal,
    });
    const refusal = await readFileGateRefusal(response);
    void response.body?.cancel().catch(() => undefined);
    if (refusal === "access") return "access-missing";
    if (refusal === "sign-in") return "sign-in-missing";
    // Anything past the gate (the API's 404) means the file session works.
    return "ok";
  } catch {
    return "unreachable";
  } finally {
    deadline.done();
  }
}






export function readBridgeHandoff(html: string, filesOrigin: string): string | null {
  let raw: string | null = null;
  if (typeof DOMParser === "function") {
    try {
      const page = new DOMParser().parseFromString(html, "text/html");
      raw = page.getElementById("bridge-data")?.textContent ?? null;
    } catch {
      raw = null;
    }
  }
  if (raw === null) {
    raw = /<script id="bridge-data" type="application\/json">([\s\S]*?)<\/script>/u.exec(html)?.[1] ?? null;
  }
  if (!raw) return null;
  let data: unknown;
  try {
    data = JSON.parse(raw);
  } catch {
    return null;
  }
  if (!data || typeof data !== "object" || Array.isArray(data)) return null;
  const { filesOrigin: origin, handoff, action } = data as Record<string, unknown>;
  if (origin !== filesOrigin || action !== "unlock") return null;
  return typeof handoff === "string" && HANDOFF_PATTERN.test(handoff) ? handoff : null;
}

function readReconnectedAt(): number | null {
  try {
    const value = Number(window.sessionStorage.getItem(RECONNECT_GUARD_KEY));
    return Number.isFinite(value) && value > 0 ? value : null;
  } catch {
    return null;
  }
}

function recordReconnect(now: number): void {
  try {
    window.sessionStorage.setItem(RECONNECT_GUARD_KEY, String(now));
  } catch {
    // Without storage the in-memory copy still guards this page.
  }
  memoryReconnectedAt = now;
}

let memoryReconnectedAt: number | null = null;

/** Was there an automatic reconnect in this tab within the guard window? */
export function reconnectedRecently(now = Date.now()): boolean {
  const at = Math.max(readReconnectedAt() ?? 0, memoryReconnectedAt ?? 0);
  return at > 0 && now - at >= 0 && now - at < RECONNECT_GUARD_MS;
}

async function runRepair(): Promise<FileAccessRepairOutcome> {
  if (!fileAccessRepairApplies()) return "not-applicable";
  const filesOrigin = configuredFileOrigin();
  if (!filesOrigin) return "not-applicable";
  if (reconnectedRecently()) return "guarded";

  let handoff: string | null;
  try {
    const response = await fetch(`${BRIDGE_PATH}?next=${encodeURIComponent("/")}`, {
      method: "GET",
      credentials: "same-origin",
      cache: "no-store",
      redirect: "error",
      headers: { Accept: "text/html" },
    });
    if (response.status === 401) {
      void response.body?.cancel().catch(() => undefined);
      // The site's own pass ran out as well; its notice offers the unlock,
      // and unlocking runs the file bridge too.
      noteAccessGateRequired();
      return "gate-locked";
    }
    if (!response.ok) {
      void response.body?.cancel().catch(() => undefined);
      return "handoff-unavailable";
    }
    handoff = readBridgeHandoff(await response.text(), filesOrigin);
  } catch {
    return "unreachable";
  }
  if (!handoff) return "handoff-unavailable";

  const deadline = timeoutSignal(HANDOFF_TIMEOUT_MS);
  try {
    const response = await fetch(`${filesOrigin}${HANDOFF_PATH}`, {
      method: "POST",
      credentials: "include",
      cache: "no-store",
      redirect: "error",
      signal: deadline.signal,
      body: new URLSearchParams({ handoff, response: "json" }),
    });
    
    // counts as this tab's reconnect, whatever the answer.
    recordReconnect(Date.now());
    if (!response.ok || !response.headers.get("content-type")?.startsWith("application/json")) {
      void response.body?.cancel().catch(() => undefined);
      return "refused";
    }
    const result: unknown = await response.json().catch(() => null);
    if (!result || typeof result !== "object" ||
        (result as { ok?: unknown }).ok !== true ||
        (result as { action?: unknown }).action !== "unlock") {
      return "refused";
    }
  } catch {
    return "unreachable";
  } finally {
    deadline.done();
  }
  const after = await probeFileAccess();
  if (after === "ok") return "repaired";
  return after === "unreachable" ? "unreachable" : "refused";
}

/** Run the file bridge from this page; concurrent callers share one run. */
export function repairFileAccess(): Promise<FileAccessRepairOutcome> {
  if (!repairFlight) {
    repairFlight = runRepair()
      .catch((): FileAccessRepairOutcome => "unreachable")
      .then((outcome) => {
        if (outcome === "repaired") noteFileAccessRestored();
        return outcome;
      })
      .finally(() => {
        repairFlight = null;
      });
  }
  return repairFlight;
}

/** Probe the file session and repair it when the file host refuses it. */
export async function checkFileAccess(): Promise<FileAccessCheck> {
  const state = await probeFileAccess();
  if (state === "ok") {
    noteFileAccessRestored();
    return "ok";
  }
  if (state !== "access-missing") return state;
  const outcome = await repairFileAccess();
  if (outcome !== "repaired" && outcome !== "gate-locked" && outcome !== "not-applicable") {
    noteFileAccessRequired();
  }
  return outcome;
}

/** The check the signed-in shell makes once per page load. */
export function checkFileAccessOnPageLoad(): Promise<FileAccessCheck> {
  if (!pageLoadCheck) pageLoadCheck = checkFileAccess();
  return pageLoadCheck;
}

/**
 * For a response the file host gave a request: when it is the gate's
 * "Access required", repair the file session. "retry" means the refused
 * request (which never reached the product) may be sent once more; any
 * other answer says why it may not, and the caller reports it.
 */
export async function recoverFromFileGateRefusal(
  response: Response,
): Promise<"not-refused" | "retry" | "gate-locked" | "unrepaired"> {
  if ((await readFileGateRefusal(response)) !== "access") return "not-refused";
  const outcome = await repairFileAccess();
  if (outcome === "repaired") return "retry";
  if (outcome === "gate-locked") return "gate-locked";
  noteFileAccessRequired();
  return "unrepaired";
}

/** Can this request body be sent a second time? */
export function requestBodyIsReplayable(init: RequestInit): boolean {
  if (init.signal?.aborted) return false;
  return !(typeof ReadableStream !== "undefined" && init.body instanceof ReadableStream);
}

export function isFileAccessRequired(): boolean {
  return required;
}

/** Say (again) that file transfers need the session reconnected. Repeats
 * within a few seconds are folded into one announcement. */
export function noteFileAccessRequired(): void {
  required = true;
  if (typeof window === "undefined") return;
  const now = Date.now();
  if (now - lastRequiredNotice < 5_000) return;
  lastRequiredNotice = now;
  window.dispatchEvent(new CustomEvent(FILE_ACCESS_REQUIRED_EVENT));
}

function noteFileAccessRestored(): void {
  if (!required) return;
  required = false;
  lastRequiredNotice = 0;
  if (typeof window !== "undefined") window.dispatchEvent(new CustomEvent(FILE_ACCESS_RESTORED_EVENT));
}

/** The gate's bridge in a tab of its own, landing on a page that says the
 * original window works again (the person's way when automatic repair
 * could not). */
export function fileAccessReconnectUrl(): string {
  return `${BRIDGE_PATH}?next=${encodeURIComponent(RECONNECTED_LANDING_PATH)}`;
}

/** The message of an error that is an unrepaired file-session refusal. */
export function fileAccessErrorMessage(error: unknown): string | null {
  if (!error || typeof error !== "object") return null;
  const { code, message } = error as { code?: unknown; message?: unknown };
  return code === FILE_ACCESS_REQUIRED_CODE && typeof message === "string" && message ? message : null;
}

/** Test seam. */
export function resetFileAccessRepairForTests(): void {
  required = false;
  lastRequiredNotice = 0;
  repairFlight = null;
  pageLoadCheck = null;
  memoryReconnectedAt = null;
  try {
    window.sessionStorage.removeItem(RECONNECT_GUARD_KEY);
  } catch {
    // nothing stored
  }
}
