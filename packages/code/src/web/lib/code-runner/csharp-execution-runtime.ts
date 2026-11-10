/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
export const CSHARP_EXECUTION_PROTOCOL = "mrnote.csharp-execution.v1";
export const CSHARP_EXECUTION_RUNTIME = "dotnet-10.0.12-roslyn-5.9.0";
export const CSHARP_RUNTIME_PACK_SHA256 =
  "ce88d675f9de16ef50f037e51abc66ee50e1eef89f6c67dab2cfaab00680f036";
export const CSHARP_EXECUTION_POLICY_SHA256 =
  "dd8032324605681492e246f1ad65333a8e95dbd82cc69df53a09e4225b455383";
export const CSHARP_TOOLCHAIN_PROFILE = "dotnet-10.0.12-roslyn-5.9.0-aot-v1";
export const CSHARP_SANDBOX_PROFILE =
  "dedicated-worker-no-network-fresh-namespace-v1";
export const CSHARP_WORKER_URL = "/api/code-runtimes/csharp/runtime-worker.js";

export const CSHARP_MAX_SOURCE_BYTES = 2_000_000;
export const CSHARP_MAX_STDIN_BYTES = 100_000;
export const CSHARP_MAX_OUTPUT_BYTES = 300_000;
export const CSHARP_MAX_MEMORY_BYTES = 268_435_456;
export const CSHARP_DEFAULT_TIMEOUT_MS = 10_000;

export const CSHARP_EXECUTION_POLICY_MANIFEST = Object.freeze({
  allowedReferences: Object.freeze([
    "System.Collections",
    "System.Console",
    "System.Linq",
    "System.Runtime",
    "netstandard",
  ]),
  freshWorkerPerRun: true,
  maxMemoryBytes: CSHARP_MAX_MEMORY_BYTES,
  maxOutputBytes: CSHARP_MAX_OUTPUT_BYTES,
  maxSourceBytes: CSHARP_MAX_SOURCE_BYTES,
  maxStdinBytes: CSHARP_MAX_STDIN_BYTES,
  network: "denied-after-same-origin-bootstrap",
  packSha256: CSHARP_RUNTIME_PACK_SHA256,
  runtime: CSHARP_EXECUTION_RUNTIME,
  schemaVersion: 1,
  stdinAdapter: "Console.ReadLine-Read-exact-line-rewrite",
  workerRoute: CSHARP_WORKER_URL,
});
