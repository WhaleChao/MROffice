/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/** Compile-only Roslyn diagnostics backed by the pinned .NET 10 Worker. */

import type {
  LanguageServiceProvider,
  ProviderCapabilities,
  ProviderDiagnostic,
} from "./language-service-provider";
import {
  compileCSharp,
  type CSharpDiagnostic,
} from "./csharp-runner";

export const CSHARP_FILE_LEVEL_DIAGNOSTIC_CODE = "csharp.fileLevel";
export const CSHARP_FILE_LEVEL_MESSAGE_PREFIX =
  "File-level C# compile error (no source location available)";

const CAPABILITIES: ProviderCapabilities = { diagnostics: true };
let providerSequence = 0;

// Every C# window on the page shares one analysis namespace and takes turns.
// Each compile still runs in its own single-use Worker generation; what is
// shared is the queue, so a notebook with many C# windows boots one runtime at
// a time instead of one per window at once.
const ANALYSIS_NAMESPACE = "csharp-analysis";
let analysisTail: Promise<unknown> = Promise.resolve();
let analysisInFlight: { owner: number; controller: AbortController } | null =
  null;

function withAnalysisTurn<T>(task: () => Promise<T>): Promise<T> {
  const turn = analysisTail.then(task, task);
  analysisTail = turn.catch(() => undefined);
  return turn;
}

function sourceSpanAt(
  source: string,
  requestedLine: number,
  requestedColumn: number,
): { from: number; to: number } | null {
  if (
    !Number.isSafeInteger(requestedLine) ||
    !Number.isSafeInteger(requestedColumn) ||
    requestedLine < 1 ||
    requestedColumn < 1
  ) {
    return null;
  }
  let line = 1;
  let start = 0;
  while (line < requestedLine) {
    const newline = source.indexOf("\n", start);
    if (newline === -1) return null;
    start = newline + 1;
    line += 1;
  }
  const newline = source.indexOf("\n", start);
  const rawEnd = newline === -1 ? source.length : newline;
  const end =
    rawEnd > start && source.charCodeAt(rawEnd - 1) === 13
      ? rawEnd - 1
      : rawEnd;
  const from = start + requestedColumn - 1;
  if (from < start || from > end) return null;
  return { from, to: from < end ? from + 1 : from };
}

function fileLevelDiagnostic(
  detail: string,
  severity: "error" | "warning" = "error",
): ProviderDiagnostic {
  return {
    span: { from: 0, to: 0 },
    severity,
    message: detail
      ? `${CSHARP_FILE_LEVEL_MESSAGE_PREFIX}: ${detail}`
      : CSHARP_FILE_LEVEL_MESSAGE_PREFIX,
    code: CSHARP_FILE_LEVEL_DIAGNOSTIC_CODE,
  };
}

function mapDiagnostic(
  source: string,
  diagnostic: CSharpDiagnostic,
): ProviderDiagnostic {
  if (diagnostic.line !== null && diagnostic.column !== null) {
    const span = sourceSpanAt(source, diagnostic.line, diagnostic.column);
    if (span) {
      return {
        span,
        severity: "error",
        message: diagnostic.message,
        code: diagnostic.code,
      };
    }
  }
  return fileLevelDiagnostic(`${diagnostic.code}: ${diagnostic.message}`);
}

export function createCSharpLanguageProvider(): LanguageServiceProvider {
  const owner = ++providerSequence;
  let diagnosticsEpoch = 0;

  return {
    id: "csharp-roslyn-compile",
    languages: ["csharp"],
    capabilities(language) {
      return language === "csharp" ? CAPABILITIES : {};
    },
    async diagnostics(request) {
      if (request.language !== "csharp") return [];
      const epoch = ++diagnosticsEpoch;
      // A newer edit in this window physically retires its older compile
      // (aborting the request retires that Worker generation); another
      // window's compile is left to finish. Normal Run remains untouched.
      if (analysisInFlight?.owner === owner) analysisInFlight.controller.abort();
      const result = await withAnalysisTurn(async () => {
        // Superseded while it waited: nothing to compile.
        if (epoch !== diagnosticsEpoch || request.signal?.aborted) {
          return { diagnostics: [], interrupted: true };
        }
        const controller = new AbortController();
        const forward = () => controller.abort();
        request.signal?.addEventListener("abort", forward, { once: true });
        const flight = { owner, controller };
        analysisInFlight = flight;
        try {
          return await compileCSharp(request.text, {
            namespace: ANALYSIS_NAMESPACE,
            signal: controller.signal,
          });
        } finally {
          request.signal?.removeEventListener("abort", forward);
          if (analysisInFlight === flight) analysisInFlight = null;
        }
      });
      if (
        epoch !== diagnosticsEpoch ||
        request.signal?.aborted ||
        result.interrupted
      ) {
        return [];
      }
      if (result.diagnostics.length > 0) {
        return result.diagnostics.map((diagnostic) =>
          mapDiagnostic(request.text, diagnostic),
        );
      }
      if (result.error) {
        const severity = /unavailable|timed out|transport/i.test(result.error)
          ? "warning"
          : "error";
        return [fileLevelDiagnostic(result.error, severity)];
      }
      return [];
    },
  };
}
