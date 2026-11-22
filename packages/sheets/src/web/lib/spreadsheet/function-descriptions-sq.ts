/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
// Excel function help in sq: the description a function shows in AutoComplete, Insert Function
// and Function Arguments, and the line for each argument box (same shape as function-descriptions-de.ts;
// loaded lazily by function-docs-locale.ts).

export type LocalizedFunctionDoc = readonly [summary: string, args?: Readonly<Record<string, string>>];

export const COMMON_ARGUMENTS: Readonly<Record<string, string>> = /* own data: COMMON_ARGUMENTS */ {};

export const FUNCTION_DOCS: Readonly<Record<string, LocalizedFunctionDoc>> = /* own data: FUNCTION_DOCS */ {};
