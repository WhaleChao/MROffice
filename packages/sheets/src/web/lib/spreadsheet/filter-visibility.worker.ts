/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { filterVisibilitySteps, exhaustFilterVisibility, type PreparedFilterVisibility } from "./filter-visibility";
const scope = globalThis as unknown as { onmessage: ((event: MessageEvent<PreparedFilterVisibility>) => void) | null; postMessage(value: unknown): void };
scope.onmessage = event => { scope.postMessage({ rows: exhaustFilterVisibility(filterVisibilitySteps(event.data)) }); };
