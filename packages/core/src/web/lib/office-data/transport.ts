/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * How the browser asks the API for an Office reference table
 * (`GET /api/v1/office-data/{name}`, app/routers/office_data.py).
 *
 * The same client every API read uses (`lib/api.ts`: session cookie, the
 * notebook mount, the error vocabulary), with the browser's HTTP cache left
 * on: the API answers with a strong ETag and `Cache-Control: private`, so a
 * reload revalidates or reuses what this browser already holds. A 401 never
 * bounces the page to the sign-in screen (the editors draw their fallback),
 * and a guest asks for nothing but the tables the API serves a guest (the
 * font registry's mapping, `officeDataServedToGuests`).
 */
import { apiGet, isApiRequestError } from "../api";
import { isLoggedIn } from "../auth-state";
import {
  OfficeDataUnavailableError,
  officeDataServedToGuests,
  type OfficeDataPartitionedName,
  type OfficeDataTransport,
  type OfficeDataWholeName,
} from "./store";

export const OFFICE_DATA_PATH = "/api/v1/office-data";

function unavailable(error: unknown): OfficeDataUnavailableError {
  if (error instanceof OfficeDataUnavailableError) return error;
  if (isApiRequestError(error)) {
    if (error.status === 0 || error.code === "network_unreachable") return new OfficeDataUnavailableError("offline");
    if (error.status === 401 || error.status === 403) return new OfficeDataUnavailableError("signed_out");
    if (error.status === 429) return new OfficeDataUnavailableError("rate_limited");
    return new OfficeDataUnavailableError("failed", `${error.status} ${error.code ?? ""}`.trim());
  }
  if (error instanceof TypeError) return new OfficeDataUnavailableError("offline");
  return new OfficeDataUnavailableError("failed");
}

async function read(path: string, servedToGuests = false): Promise<Record<string, unknown>> {
  if (!servedToGuests && !isLoggedIn()) throw new OfficeDataUnavailableError("signed_out");
  if (typeof navigator !== "undefined" && navigator.onLine === false) throw new OfficeDataUnavailableError("offline");
  try {
    const answer = await apiGet<unknown>(path, { cache: "default" }, { suppressUnauthorizedHandling: true });
    if (!answer || typeof answer !== "object" || Array.isArray(answer)) {
      throw new OfficeDataUnavailableError("failed", "malformed response");
    }
    return answer as Record<string, unknown>;
  } catch (error) {
    throw unavailable(error);
  }
}

export const officeDataApiTransport: OfficeDataTransport = {
  signedIn: () => isLoggedIn(),
  async whole(name: OfficeDataWholeName) {
    const answer = await read(`${OFFICE_DATA_PATH}/${encodeURIComponent(name)}`);
    if (!("data" in answer)) throw new OfficeDataUnavailableError("failed", "malformed response");
    return answer.data;
  },
  async parts(name: OfficeDataPartitionedName, keys: readonly string[]) {
    const query = keys.map((key) => `part=${encodeURIComponent(key)}`).join("&");
    const answer = await read(`${OFFICE_DATA_PATH}/${encodeURIComponent(name)}?${query}`, officeDataServedToGuests(name));
    const data = answer.data;
    if (!data || typeof data !== "object" || Array.isArray(data)) {
      throw new OfficeDataUnavailableError("failed", "malformed response");
    }
    return data as Record<string, unknown>;
  },
};
