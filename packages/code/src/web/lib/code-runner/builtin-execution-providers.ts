/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/** Browser-local adapters for the execution runtimes that already ship. */

import { runJavaScript } from "@mrnote/core/web/lib/code-runner/js-runner.ts";
import {
  getPythonExecutionAvailability,
  runPython,
} from "@mrnote/core/web/lib/code-runner/python-runner.ts";
import {
  runPythonInCloudGui,
  runsInCloudGui,
  type CloudGuiMessage,
} from "./cloud-gui";
import { runSql } from "./sql-runner";
import { runTypeScript } from "./ts-runner";
import { runLua } from "./lua-runner";
import {
  getCSharpRuntimeAvailability,
  runCSharp,
  subscribeCSharpRuntimeAvailability,
} from "./csharp-runner";
import {
  CSHARP_EXECUTION_POLICY_SHA256,
  CSHARP_MAX_MEMORY_BYTES,
  CSHARP_RUNTIME_PACK_SHA256,
  CSHARP_SANDBOX_PROFILE,
  CSHARP_TOOLCHAIN_PROFILE,
} from "./csharp-execution-runtime";
import {
  getRubyRuntimeAvailability,
  runRuby,
  subscribeRubyRuntimeAvailability,
} from "./ruby-runner";
import {
  getShellRuntimeAvailability,
  runShell,
  subscribeShellRuntimeAvailability,
} from "./shell-runner";
import {
  SHELL_EXECUTION_POLICY_PROFILE,
  SHELL_EXECUTION_POLICY_SHA256,
  SHELL_HARDENED_WASM_SHA256,
  SHELL_TOOLCHAIN_PROFILE,
} from "./shell-execution-runtime";
import {
  createRemoteExecutionProvider,
  type RemoteExecutionTarget,
} from "./remote-execution-provider";
import type { RunResult } from "@mrnote/core/web/lib/code-runner/js-runner.ts";
import {
  LANGUAGE_EXECUTION_SCHEMA_VERSION,
  LanguageExecutionProviderRegistry,
  outcomeForRunResult,
  phaseReceipt,
  type LanguageExecutionAvailability,
  type LanguageExecutionProvider,
  type LanguageExecutionReceipt,
  type LanguageExecutionRequest,
  type LanguageExecutionContext,
} from "./language-execution-provider";

const TOOLCHAIN_DIGESTS = {
  // Digests identify the immutable runtime/toolchain inputs that are present
  // in this build. Remote providers must supply OCI/toolchain digests from an
  // authenticated capability receipt instead of inventing a local label.
  javascript: "a648fee4e1f8e2e65dd5f143df28e5a5792ed12b9143b3492d300880e8917726",
  typescript: "ed1401d81a1a4161f62db494b6e820238e3a3dfe84967254ef13db0d7a9af9c7",
  python: "7a78603ef97cfa7b8cce4f0a404e82bec490e90a7d9a93026886724c8917b1b5",
  sql: "8c0d4061652824f2a0688ecb7c62fc1389e2aa0f2665749fbdc2c47ce76516f8",
  lua: "0e7a82c07d90b9ae967a52b1f458400f37c58a9c98a6f7790a4edb0e3b88f8db",
  csharp: CSHARP_RUNTIME_PACK_SHA256,
  ruby: "9fe3c749730a16da9b0d1a973bb77bec164f8d958efcb7f5b81e4275ea8c5439",
  bash: SHELL_HARDENED_WASM_SHA256,
} as const;

const TOOLCHAIN_PROFILES = {
  javascript: "js-code-sandbox-v1",
  typescript: "ts-sucrase-3.35.1-code-sandbox-v1",
  python: "python-pyodide-0.26.2-worker-v1",
  sql: "sql-sqljs-1.14.1-wasm-v1",
  lua: "lua-5.5.0-wasm-emscripten-3.1.6-v1",
  csharp: CSHARP_TOOLCHAIN_PROFILE,
  ruby: "ruby-4.0-wasm-wasi-2.10.1-v1",
  bash: SHELL_TOOLCHAIN_PROFILE,
} as const;

const POLICY_DIGESTS = {
  javascript: "b6023ff931df450a466627e1a259e2c1eafacd87ba65cce8756c72b8d0b0ac30",
  typescript: "b6023ff931df450a466627e1a259e2c1eafacd87ba65cce8756c72b8d0b0ac30",
  python: "336659236c056c5adf6934dbc8aeb49d386dceb5ec264a64b145be6d63c0bc6a",
  sql: "7a77355a5de40167ef3a9c8ee3f110a382b144f2cd35e714a4a64670521e205c",
  lua: "812b70db68606a3e4e77d7d848640821e2b1bfa54639e427e428ab6027977d08",
  csharp: CSHARP_EXECUTION_POLICY_SHA256,
  ruby: "19a27d35ceea90e02fd406ef688f185db6a4f69140fe363ed142707105358f46",
  bash: SHELL_EXECUTION_POLICY_SHA256,
} as const;

const POLICY_PROFILES = {
  javascript: "cross-origin-worker-no-app-credentials-v1",
  typescript: "cross-origin-worker-no-app-credentials-v1",
  python: "cross-origin-pyodide-worker-no-app-credentials-v1",
  sql: "origin-sqljs-worker-no-app-credentials-v1",
  lua: "origin-lua-wasm-worker-safe-libs-v1",
  csharp: CSHARP_SANDBOX_PROFILE,
  ruby: "origin-ruby-wasi-worker-no-host-interop-v1",
  bash: SHELL_EXECUTION_POLICY_PROFILE,
} as const;

const INPUT_BYTES = 2_000_000;
const OUTPUT_BYTES = 300_000;

function limitsFor(
  language: keyof typeof TOOLCHAIN_DIGESTS,
): import("./language-execution-provider").LanguageExecutionLimits {
  return {
    sourceBytes: INPUT_BYTES,
    stdinBytes:
      language === "python" ||
      language === "lua" ||
      language === "csharp" ||
      language === "ruby" ||
      language === "bash"
        ? 100_000
        : 0,
    wallTimeMs:
      language === "javascript" || language === "typescript"
        ? 5_000
        : language === "python"
          ? 20_000
          : language === "sql"
            ? 5_000
            : language === "lua" ||
                language === "csharp" ||
                language === "ruby" ||
                language === "bash"
              ? language === "csharp"
                ? 10_000
                : 5_000
              : null,
    cpuTimeMs: null,
    memoryBytes:
      language === "lua"
        ? 67_108_864
        : language === "csharp"
          ? CSHARP_MAX_MEMORY_BYTES
          : language === "ruby"
            ? 268_435_456
            : language === "bash"
              ? 67_108_864
              : null,
    stdoutBytes: OUTPUT_BYTES,
    stderrBytes: OUTPUT_BYTES,
    outputBytes: OUTPUT_BYTES,
  };
}

const PROVIDER_ID = "mrai.browser.v1";

export interface BuiltinExecutionProviderContext {
  readonly javaScriptSandboxOwner: string;
  readonly pageNamespace: string;
  /**
   * A signed-out window. The browser providers still run — a guest may run all
   * eight of them — but the remote provider has no authority to ask for, so it
   * answers `unavailable` without polling an endpoint that must refuse it.
   */
  readonly guest?: boolean;
  /** Remote runs go here instead of the code page's route (a code draft). */
  readonly remoteTarget?: RemoteExecutionTarget | null;
  /**
   * A saved code page's Python that needs a desktop window, a terminal or
   * processes runs in a cloud window when the deployment offers them
   * (cloud-gui.ts): the page it runs as, and the page's words for its end.
   * Absent for a draft, a guest and a test run.
   */
  readonly cloudGui?: { readonly pageId: string; readonly message: CloudGuiMessage } | null;
}

function capabilityFor(
  language: keyof typeof TOOLCHAIN_DIGESTS,
): Extract<LanguageExecutionAvailability, { status: "available" }> {
  return {
    status: "available",
    providerId: PROVIDER_ID,
    mode: "browser",
    toolchainProfile: TOOLCHAIN_PROFILES[language],
    toolchainDigest: TOOLCHAIN_DIGESTS[language],
    sandboxProfile: POLICY_PROFILES[language],
    sandboxPolicyDigest: POLICY_DIGESTS[language],
    supports: {
      stdin:
        language === "python" ||
        language === "lua" ||
        language === "csharp" ||
        language === "ruby" ||
        language === "bash",
      partialExecution: language !== "csharp",
      streaming:
        language !== "sql" &&
        language !== "csharp" &&
        language !== "ruby" &&
        language !== "bash",
      cancel: true,
      restart: true,
      interactiveConsole: language !== "bash" && language !== "csharp",
      debug:
        language === "javascript" ||
        language === "typescript" ||
        language === "python",
    },
    limits: limitsFor(language),
  };
}

function receiptFor(
  request: LanguageExecutionRequest,
  output: RunResult,
  startedAt: string,
): LanguageExecutionReceipt {
  const outcome = outcomeForRunResult(output);
  const compileFailed =
    (request.language === "typescript" ||
      request.language === "lua" ||
      request.language === "csharp" ||
      request.language === "ruby" ||
      request.language === "bash") &&
    typeof output.error === "string" &&
    output.error.startsWith(
      request.language === "lua"
        ? "Lua compile error:"
        : request.language === "csharp"
          ? "C# compile error:"
          : request.language === "ruby"
            ? "Ruby compile error:"
            : request.language === "bash"
              ? "Shell compile error:"
              : "TypeScript compile error",
    );
  const quotaFailed =
    (request.language === "lua" ||
      request.language === "csharp" ||
      request.language === "ruby" ||
      request.language === "bash") &&
    typeof output.error === "string" &&
    output.error.startsWith(
      request.language === "ruby"
        ? "Ruby execution quota exceeded:"
        : request.language === "csharp"
          ? "C# execution quota exceeded:"
          : request.language === "bash"
            ? "Shell execution quota exceeded:"
            : "Lua execution quota exceeded:",
    );
  const runFailed = !compileFailed && outcome !== "succeeded";
  const rawRuntimeExitCode =
    request.language === "bash" || request.language === "csharp"
      ? (output as RunResult & { readonly exitCode?: unknown }).exitCode
      : undefined;
  const runtimeExitCode =
    Number.isSafeInteger(rawRuntimeExitCode) &&
    Number(rawRuntimeExitCode) >= 0 &&
    Number(rawRuntimeExitCode) <= 255
      ? Number(rawRuntimeExitCode)
      : null;
  const runPhase = phaseReceipt(
    compileFailed ? "skipped" : runFailed ? "failed" : "succeeded",
  );
  return {
    schemaVersion: LANGUAGE_EXECUTION_SCHEMA_VERSION,
    requestId: request.requestId,
    pageId: request.pageId,
    contentVersion: request.contentVersion,
    language: request.language,
    sourceSha256: request.sourceSha256,
    providerId: request.providerId,
    toolchainProfile: request.toolchainProfile,
    toolchainDigest: request.toolchainDigest,
    sandboxProfile: request.sandboxProfile,
    sandboxPolicyDigest: request.sandboxPolicyDigest,
    outcome: compileFailed ? "compile-error" : quotaFailed ? "quota-exceeded" : outcome,
    compile: phaseReceipt(compileFailed ? "failed" : "succeeded"),
    run:
      (request.language === "bash" || request.language === "csharp") &&
      !compileFailed &&
      runtimeExitCode !== null
        ? { ...runPhase, exitCode: runtimeExitCode }
        : runPhase,
    output,
    startedAt,
    completedAt: new Date().toISOString(),
  };
}

export function createBuiltinExecutionProvider(
  owner: BuiltinExecutionProviderContext,
): LanguageExecutionProvider {
  const languages = [
    "javascript",
    "typescript",
    "python",
    "sql",
    "lua",
    "csharp",
    "ruby",
    "bash",
  ] as const;
  return {
    id: PROVIDER_ID,
    mode: "browser",
    languages,
    subscribe(listener) {
      const unsubscribeRuby = subscribeRubyRuntimeAvailability(
        owner.pageNamespace,
        listener,
      );
      const unsubscribeShell = subscribeShellRuntimeAvailability(
        owner.pageNamespace,
        listener,
      );
      const unsubscribeCSharp = subscribeCSharpRuntimeAvailability(
        owner.pageNamespace,
        listener,
      );
      return () => {
        unsubscribeRuby();
        unsubscribeShell();
        unsubscribeCSharp();
      };
    },
    capability(language) {
      if (!languages.includes(language as (typeof languages)[number])) {
        return { status: "unavailable", reason: "provider-unavailable" };
      }
      if (
        language === "python" &&
        getPythonExecutionAvailability() === "unavailable"
      ) {
        return { status: "unavailable", reason: "provider-unavailable" };
      }
      if (language === "ruby") {
        const runtimeStatus = getRubyRuntimeAvailability(owner.pageNamespace);
        if (runtimeStatus === "checking") {
          return {
            status: "checking",
            providerId: PROVIDER_ID,
            mode: "browser",
            supports: {
              stdin: false,
              partialExecution: false,
              streaming: false,
              cancel: false,
              restart: false,
              interactiveConsole: false,
              debug: false,
            },
            reason: "capability-checking",
          };
        }
        if (runtimeStatus === "unavailable") {
          return { status: "unavailable", reason: "provider-unavailable" };
        }
      }
      if (language === "bash") {
        const runtimeStatus = getShellRuntimeAvailability(owner.pageNamespace);
        if (runtimeStatus === "checking") {
          return {
            status: "checking",
            providerId: PROVIDER_ID,
            mode: "browser",
            supports: {
              stdin: false,
              partialExecution: false,
              streaming: false,
              cancel: false,
              restart: false,
              interactiveConsole: false,
              debug: false,
            },
            reason: "capability-checking",
          };
        }
        if (runtimeStatus === "unavailable") {
          return { status: "unavailable", reason: "provider-unavailable" };
        }
      }
      if (language === "csharp") {
        const runtimeStatus = getCSharpRuntimeAvailability(owner.pageNamespace);
        if (runtimeStatus === "checking") {
          return {
            status: "checking",
            providerId: PROVIDER_ID,
            mode: "browser",
            supports: {
              stdin: false,
              partialExecution: false,
              streaming: false,
              cancel: false,
              restart: false,
              interactiveConsole: false,
              debug: false,
            },
            reason: "capability-checking",
          };
        }
        if (runtimeStatus === "unavailable") {
          return { status: "unavailable", reason: "provider-unavailable" };
        }
      }
      return capabilityFor(language as keyof typeof TOOLCHAIN_DIGESTS);
    },
    async execute(
      request: LanguageExecutionRequest,
      context: LanguageExecutionContext,
    ): Promise<LanguageExecutionReceipt> {
      if (!languages.includes(request.language as (typeof languages)[number])) {
        throw new Error("Execution language is not owned by this provider.");
      }
      const capability = capabilityFor(
        request.language as keyof typeof TOOLCHAIN_DIGESTS,
      );
      if (
        request.providerId !== PROVIDER_ID ||
        request.toolchainProfile !== capability.toolchainProfile ||
        request.toolchainDigest !== capability.toolchainDigest ||
        request.sandboxProfile !== capability.sandboxProfile ||
        request.sandboxPolicyDigest !== capability.sandboxPolicyDigest
      ) {
        throw new Error("Execution request authority does not match provider.");
      }
      const startedAt = new Date().toISOString();
      let output: RunResult;
      if (request.language === "javascript") {
        output = await runJavaScript(request.source, {
          signal: context.signal,
          onOutput: context.onOutput,
          sandboxOwner: owner.javaScriptSandboxOwner,
          timeoutMs: request.limits.wallTimeMs ?? undefined,
        });
      } else if (request.language === "typescript") {
        output = await runTypeScript(request.source, {
          signal: context.signal,
          onOutput: context.onOutput,
          sandboxOwner: owner.javaScriptSandboxOwner,
          timeoutMs: request.limits.wallTimeMs ?? undefined,
        });
      } else if (request.language === "python" && owner.cloudGui && runsInCloudGui(request.source)) {
        // Its window plays in the run output, as a pygame game's does; it runs
        // until it ends or the reader stops it, not to the 20 s stop.
        output = await runPythonInCloudGui(request.source, {
          pageId: owner.cloudGui.pageId,
          namespace: owner.pageNamespace,
          signal: context.signal,
          onOutput: context.onOutput,
          message: owner.cloudGui.message,
          ...(request.stdin.trim() !== "" ? { stdin: request.stdin } : {}),
        });
      } else if (request.language === "python") {
        output = await runPython(request.source, {
          namespace: owner.pageNamespace,
          signal: context.signal,
          onOutput: context.onOutput,
          timeoutMs: request.limits.wallTimeMs ?? undefined,
          ...(request.stdin.trim() !== "" ? { stdin: request.stdin } : {}),
        });
      } else if (request.language === "sql") {
        output = await runSql(request.source, {
          namespace: owner.pageNamespace,
          signal: context.signal,
          timeoutMs: request.limits.wallTimeMs ?? undefined,
          outputBytes: request.limits.outputBytes,
        });
      } else if (request.language === "lua") {
        output = await runLua(request.source, {
          namespace: owner.pageNamespace,
          signal: context.signal,
          onOutput: context.onOutput,
          timeoutMs: request.limits.wallTimeMs ?? undefined,
          ...(request.stdin !== "" ? { stdin: request.stdin } : {}),
        });
      } else if (request.language === "csharp") {
        output = await runCSharp(request.source, {
          namespace: owner.pageNamespace,
          signal: context.signal,
          timeoutMs: request.limits.wallTimeMs ?? undefined,
          outputBytes: request.limits.outputBytes,
          ...(request.stdin !== "" ? { stdin: request.stdin } : {}),
        });
      } else if (request.language === "ruby") {
        output = await runRuby(request.source, {
          namespace: owner.pageNamespace,
          signal: context.signal,
          timeoutMs: request.limits.wallTimeMs ?? undefined,
          outputBytes: request.limits.outputBytes,
          ...(request.stdin !== "" ? { stdin: request.stdin } : {}),
        });
      } else {
        output = await runShell(request.source, {
          namespace: owner.pageNamespace,
          signal: context.signal,
          onOutput: context.onOutput,
          timeoutMs: request.limits.wallTimeMs ?? undefined,
          outputBytes: request.limits.outputBytes,
          ...(request.stdin !== "" ? { stdin: request.stdin } : {}),
        });
      }
      return receiptFor(request, output, startedAt);
    },
  };
}

export function createBuiltinExecutionProviderRegistry(
  owner: BuiltinExecutionProviderContext,
): LanguageExecutionProviderRegistry & { dispose(): void } {
  const registry = new LanguageExecutionProviderRegistry();
  const remoteProvider = createRemoteExecutionProvider({
    guest: owner.guest,
    target: owner.remoteTarget ?? null,
  });
  let unregisterBuiltin: (() => void) | null = registry.register(
    createBuiltinExecutionProvider(owner),
  );
  let unregisterRemote: (() => void) | null = registry.register(remoteProvider);
  return Object.assign(registry, {
    dispose() {
      unregisterRemote?.();
      unregisterRemote = null;
      remoteProvider.dispose();
      unregisterBuiltin?.();
      unregisterBuiltin = null;
    },
  });
}
