/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
export const ACCESS_GATE_ERROR_CODE = "beta_access_required";
export const ACCESS_GATE_REQUIRED_EVENT = "mrnote:access-gate-required";
export const ACCESS_GATE_RESTORED_EVENT = "mrnote:access-gate-restored";
/** Served to anyone who is through the gate; holds nothing private. */
const PROBE_ENDPOINT = "/web-build.json";
const UNLOCK_PATH = "/__beta";
/** A static page that only says the original window can be used again. */
const UNLOCKED_LANDING_PATH = "/access-restored.html";

let required = false;
let probing: Promise<boolean> | null = null;

/** Is this response the gate's own refusal rather than the product's? */
export function isAccessGateRefusal(status: number, body: unknown): boolean {
  return status === 401 && !!body && typeof body === "object" &&
    (body as { error?: unknown }).error === ACCESS_GATE_ERROR_CODE;
}

export function isAccessGateRequired(): boolean {
  return required;
}

/** Record a gate refusal; the first one of an episode is announced. */
export function noteAccessGateRequired(): void {
  if (required) return;
  required = true;
  if (typeof window !== "undefined") window.dispatchEvent(new CustomEvent(ACCESS_GATE_REQUIRED_EVENT));
}

function noteAccessGateRestored(): void {
  if (!required) return;
  required = false;
  if (typeof window !== "undefined") window.dispatchEvent(new CustomEvent(ACCESS_GATE_RESTORED_EVENT));
}

/** Where to pass the gate again without leaving the current window. */
export function accessGateUnlockUrl(): string {
  return `${UNLOCK_PATH}?next=${encodeURIComponent(UNLOCKED_LANDING_PATH)}`;
}

/**
 * Ask whether requests pass the gate again. Resolves to true when they do.
 * A network failure or any other answer leaves the verdict as it was.
 */
export function probeAccessGate(): Promise<boolean> {
  if (!required || typeof fetch !== "function") return Promise.resolve(!required);
  if (probing) return probing;
  probing = fetch(PROBE_ENDPOINT, { cache: "no-store", credentials: "same-origin", headers: { accept: "application/json" } })
    .then((response) => {
      if (response.ok) noteAccessGateRestored();
      return !required;
    })
    .catch(() => !required)
    .finally(() => { probing = null; });
  return probing;
}

/** Test seam. */
export function resetAccessGateForTests(): void {
  required = false;
  probing = null;
}
