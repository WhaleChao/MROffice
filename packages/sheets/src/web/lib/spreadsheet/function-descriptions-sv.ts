/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
// ---------------------------------------------------------------------------
// Swedish function help (XLS-A094 / XLS-A067): the description a
// function shows in AutoComplete, Insert Function and Function Arguments, and
// the line for each argument box, following Excel's own wording in this
// language. Same shape as function-descriptions-zh.ts; PURE data, loaded
// lazily by function-docs-locale.ts so only a session in this language
// downloads it.
// ---------------------------------------------------------------------------

export type LocalizedFunctionDoc = readonly [summary: string, args?: Readonly<Record<string, string>>];

/** Arguments with one meaning across every function that takes them. */
export const COMMON_ARGUMENTS: Readonly<Record<string, string>> = /* own data: COMMON_ARGUMENTS */ {};

export const FUNCTION_DOCS: Readonly<Record<string, LocalizedFunctionDoc>> = /* own data: FUNCTION_DOCS */ {};
