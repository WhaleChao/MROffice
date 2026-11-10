/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/** Closed machine-readable failures emitted by the attested executor. */
export const LANGUAGE_EXECUTION_ERROR_CODES = [
  "compile-failed",
  "runtime-failed",
  "time-limit",
  "cpu-limit",
  "memory-limit",
  "output-limit",
  "interrupted",
  "sandbox-policy",
  "toolchain-failure",
] as const;

export type LanguageExecutionErrorCode =
  (typeof LANGUAGE_EXECUTION_ERROR_CODES)[number];

const ERROR_CODES = new Set<string>(LANGUAGE_EXECUTION_ERROR_CODES);

export function isLanguageExecutionErrorCode(
  value: unknown,
): value is LanguageExecutionErrorCode {
  return typeof value === "string" && ERROR_CODES.has(value);
}

const MESSAGE_KEYS = {
  "compile-failed": "code.executionError.compileFailed",
  "runtime-failed": "code.executionError.runtimeFailed",
  "time-limit": "code.executionError.timeLimit",
  "cpu-limit": "code.executionError.cpuLimit",
  "memory-limit": "code.executionError.memoryLimit",
  "output-limit": "code.executionError.outputLimit",
  interrupted: "codeOutput.interrupted",
  "sandbox-policy": "code.executionError.sandboxPolicy",
  "toolchain-failure": "code.executionError.toolchainFailure",
} as const satisfies Record<LanguageExecutionErrorCode, string>;

export type LanguageExecutionErrorMessageKey =
  (typeof MESSAGE_KEYS)[LanguageExecutionErrorCode];

export function languageExecutionErrorMessageKey(
  code: LanguageExecutionErrorCode,
): LanguageExecutionErrorMessageKey {
  return MESSAGE_KEYS[code];
}
