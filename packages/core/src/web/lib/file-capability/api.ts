/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Q7-PR0 UI Affordance Gate — capability resolver API client.
 *
 * Single-purpose wrapper around the canonical
 * {@code GET /api/v1/resources/&#123;resource_id&#125;/capability} endpoint
 * (spec §2.1) that maps HTTP responses to the
 * {@link FileCapabilityState} discriminated union (spec §2.3 + §3.1).
 *
 * Q7-PR0 is a consumer-side PR (vocab §17 red line #5); this client
 * never writes to backend state, never invents capabilities on
 * degraded transport, and never optimistic-renders edit chrome on the
 * caller's behalf.
 */
import { apiGet, ApiRequestError, isApiRequestError } from "../api";
import type {
  CapabilityResolutionUiProjection,
  FileCapabilityResponse,
  FileCapabilityState,
} from "@/lib/file-capability/types";

/**
 * Build the resolver path. Encoded so callers can pass arbitrary opaque
 * resource IDs without crafting their own URL.
 */
function buildCapabilityPath(resourceId: string): string {
  return `/api/v1/resources/${encodeURIComponent(resourceId)}/capability`;
}

/**
 * Type-guard for the consumer-side projection. Validates the closed
 * top-level keys we read (spec §2.2 invariants 1, 7, 8). Per Q7-PR0
 * acceptance #11, deeper shape validation is performed by the
 * backend's Pydantic contract test
 * (`apps/api/tests/api/test_file_capability_endpoint.py`); the frontend
 * only confirms the envelope is plausible enough to switch on safely.
 */
function isCapabilityResolutionShape(
  value: unknown,
): value is CapabilityResolutionUiProjection {
  if (typeof value !== "object" || value === null) return false;
  const v = value as Record<string, unknown>;
  return (
    typeof v.resolver_schema_version === "string" &&
    typeof v.claim_lock_revision === "string" &&
    typeof v.family_version === "number" &&
    typeof v.family === "string" &&
    typeof v.extension_display === "string" &&
    typeof v.editing_path === "string" &&
    typeof v.ui === "object" &&
    v.ui !== null &&
    typeof v.claim_trace === "object" &&
    v.claim_trace !== null &&
    typeof v.eval_evidence_summary === "object" &&
    v.eval_evidence_summary !== null
  );
}

function isFileCapabilityResponse(
  value: unknown,
): value is FileCapabilityResponse {
  if (typeof value !== "object" || value === null) return false;
  const v = value as Record<string, unknown>;
  return (
    typeof v.resource_id === "string" &&
    isCapabilityResolutionShape(v.resolution)
  );
}

/**
 * Normalize an unknown thrown value into an {@code Error} instance so the
 * {@code resolver_unavailable} state always carries a real Error (the
 * discriminated union requires it; W-F lint guards against `any`).
 */
function toError(value: unknown, fallbackMessage: string): Error {
  if (value instanceof Error) return value;
  if (typeof value === "string") return new Error(value);
  return new Error(fallbackMessage);
}

/**
 * Fetch a file capability resolution and project the HTTP outcome into
 * the {@link FileCapabilityState} discriminated union.
 *
 * Transport → state mapping (spec §2.3):
 * - {@code 200} + envelope-plausible → {@code &#123; kind: "ready", resolution &#125;}
 * - {@code 401 / 403} → {@code &#123; kind: "auth_required" &#125;}
 * - {@code 404} → {@code &#123; kind: "not_found" &#125;}
 * - {@code 5xx} / network failure / malformed envelope →
 *   {@code &#123; kind: "resolver_unavailable", error &#125;}
 *
 * The 401 path passes {@code suppressUnauthorizedHandling: true} so the
 * shared {@code apiGet} does NOT auto-redirect to the login page —
 * Q7-PR0 surfaces auth-required as a UI state, not a navigation event.
 * The caller (hook + UI shell) decides whether to render the auth
 * fallback or to defer to a sibling auth boundary.
 */
export async function fetchFileCapability(
  resourceId: string,
): Promise<FileCapabilityState> {
  const path = buildCapabilityPath(resourceId);
  try {
    const raw = await apiGet<unknown>(
      path,
      undefined,
      { suppressUnauthorizedHandling: true },
    );
    if (!isFileCapabilityResponse(raw)) {
      return {
        kind: "resolver_unavailable",
        error: new Error(
          "File capability response shape mismatch (Q7-PR0 §2.2 contract violation).",
        ),
      };
    }
    return { kind: "ready", resolution: raw.resolution };
  } catch (err: unknown) {
    if (isApiRequestError(err)) {
      return mapApiRequestError(err);
    }
    return {
      kind: "resolver_unavailable",
      error: toError(err, "File capability fetch failed."),
    };
  }
}

/**
 * Map an {@code ApiRequestError} to a {@link FileCapabilityState}. The
 * 402 plan-required status is treated as {@code auth_required} for UI
 * gating purposes (the plan-required modal is a sibling concern wired
 * through {@code mrnote:plan-required}; Q7-PR0 must not silently render
 * edit chrome while a plan upgrade is pending).
 *
 * Network-unreachable failures throw an {@code ApiRequestError} with
 * {@code status: 0} and {@code code: "network_unreachable"} (see
 * {@code apps/web/lib/api.ts} {@code toApiRequestError}); those map to
 * {@code resolver_unavailable}.
 */
function mapApiRequestError(err: ApiRequestError): FileCapabilityState {
  const status = err.status;
  if (status === 401 || status === 403 || status === 402) {
    return { kind: "auth_required" };
  }
  if (status === 404) {
    return { kind: "not_found" };
  }
  // status 0 (network), 5xx, and any other non-2xx outcome fall through
  // to the conservative "resolver_unavailable" branch per spec §2.3.
  return { kind: "resolver_unavailable", error: err };
}
