/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The TypeScript/JavaScript intelligence, behind the provider interface.
 *
 * Nothing here is new capability: `code-intelligence.ts` already drives the real
 * `LanguageService` for js/jsx/ts/tsx. What this adds is that the editor can now
 * ask for it the same way it will ask a Go or Rust server later — and the proof
 * that the interface can carry a real engine, not just a fallback.
 *
 * The compiler is pulled in lazily, as it already was: a page that never asks
 * for intelligence never pays for it.
 */

import { localizeTypeScript } from "./typescript-locale";
import {
  intelligenceLanguage,
  createTsService,
  quickInfoAt,
  completionsAt,
  definitionSpan,
  referenceSpans,
  type TsIntelligenceService,
} from "./code-intelligence";
import type {
  LanguageServiceProvider,
  LanguageServiceRequest,
  ProviderCapabilities,
  ProviderCompletion,
  ProviderHover,
  ProviderLocation,
} from "./language-service-provider";

const LANGUAGES = ["javascript", "jsx", "typescript", "tsx"] as const;

/**
 * Quick info arrives as the compiler's display parts. The interface promises
 * PLAIN TEXT, so the parts are joined rather than rendered — a caller that wants
 * the editor's HTML rendering still has `quickInfoToHtml` for it.
 */
function quickInfoText(info: {
  displayParts?: readonly { text: string }[];
  documentation?: readonly { text: string }[];
}): string {
  const signature = (info.displayParts ?? []).map((part) => part.text).join("");
  const docs = (info.documentation ?? []).map((part) => part.text).join("");
  return docs ? `${signature}\n\n${docs}` : signature;
}

/**
 * Build a provider over one buffer.
 *
 * The service is created per request from the CURRENT text: the caller owns the
 * buffer, and a provider that cached a service would answer about a document
 * the user has already changed — the same staleness that makes a fold cache or
 * a compare baseline lie.
 */
export function createTypeScriptLanguageProvider(): LanguageServiceProvider {
  const serviceFor = async (
    text: string,
    language: string,
  ): Promise<TsIntelligenceService | null> => {
    const resolved = intelligenceLanguage(language);
    if (!resolved) return null;
    const ts = await import("typescript");
    await localizeTypeScript(ts);
    return createTsService(ts, () => text, resolved);
  };

  return {
    id: "typescript",
    languages: LANGUAGES,

    capabilities(language: string): ProviderCapabilities {
      if (!intelligenceLanguage(language)) return {};
      // Rename and diagnostics exist in this codebase too, but they are wired
      // through their own surfaces (the rename affordance, the lint extension).
      // Advertising them here before they are routed through the registry would
      // be exactly the lie this interface exists to prevent.
      return {
        completions: true,
        hover: true,
        definition: true,
        references: true,
      };
    },

    async completions(
      request: LanguageServiceRequest & { language: string },
    ): Promise<ProviderCompletion[]> {
      const service = await serviceFor(request.text, request.language);
      if (!service) return [];
      const result = completionsAt(service, request.offset);
      if (!result) return [];
      // The engine's entry carries `type` (CodeMirror's icon name) and a sort
      // weight; there is no per-entry detail without a second round trip to
      // `completionEntryDetail`, so the interface's optional `detail` is left
      // absent rather than filled with something that is not the detail.
      return result.entries.map((entry) => ({
        label: entry.name,
        kind: entry.type,
      }));
    },

    async hover(
      request: LanguageServiceRequest & { language: string },
    ): Promise<ProviderHover | null> {
      const service = await serviceFor(request.text, request.language);
      if (!service) return null;
      const info = quickInfoAt(service, request.offset);
      if (!info) return null;
      const contents = quickInfoText(info);
      if (!contents) return null;
      return {
        contents,
        span: info.textSpan
          ? {
              from: info.textSpan.start,
              to: info.textSpan.start + info.textSpan.length,
            }
          : undefined,
      };
    },

    async definition(
      request: LanguageServiceRequest & { language: string },
    ): Promise<ProviderLocation | null> {
      const service = await serviceFor(request.text, request.language);
      if (!service) return null;
      const span = definitionSpan(service, request.offset);
      if (!span) return null;
      return { bufferId: null, from: span.from, to: span.to };
    },

    async references(
      request: LanguageServiceRequest & { language: string },
    ): Promise<ProviderLocation[]> {
      const service = await serviceFor(request.text, request.language);
      if (!service) return [];
      return referenceSpans(service, request.offset).map((span) => ({
        bufferId: null,
        from: span.from,
        to: span.to,
      }));
    },
  };
}
