/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The editor's single answer to "what does this language know about itself",
 * and the one place a new language gets added.
 *
 * Until now the editor asked that question by name. `intelligenceLanguage(...)`
 * meant TypeScript, `isPythonEditor` meant jedi, and every affordance that had
 * to know — the rename prompt, Cmd+click, the signature-help trigger, the
 * context menu — repeated the pair. A third language meant finding all of them
 * again, which is why "more languages" kept being a rewrite.
 *
 * Here the question is asked once, of a registry:
 *
 *   - `serves(language)` — is anyone offering to answer for this language at
 *     all? This is what the affordances gate on. It is deliberately the STATIC
 *     question: whether a particular answer is available right now (Python's
 *     jedi needs a live kernel) is settled when the call is actually made, so a
 *     menu item does not blink in and out while a kernel boots.
 *   - `capabilitiesFor(language)` / `providerFor(...)` — the live question, for
 *     the call itself.
 *
 * `registerLanguageServiceProvider` is the whole of "add a language": register
 * a provider, and every surface routed through here starts serving it —
 * completions, hover, definition, type definition, implementation, references,
 * diagnostics, rename, formatting and signature help, which is every capability
 * the interface declares.
 *
 * Providers are bounded on the way in (`boundProviderResults`), including the
 * built-ins. What comes back from one is the newest edge where an unsized
 * answer could wedge a popup, a panel or the document itself, and a bridge to a
 * remote language server is exactly the shape that edge will take next.
 *
 * Some paths stay bespoke, and that is a consequence of the interface rather
 * than an exception to it: TypeScript quick info is structured display parts
 * rendered as sanitized HTML, and Python completions resolve against the whole
 * notebook's concatenated cells, while the interface promises PLAIN TEXT and a
 * single buffer on purpose. Where the interface loses nothing, the built-ins go
 * through it — Python's go-to-definition and rename do, which is how the rename
 * read-back guard reached them.
 *
 * Note the two built-ins deliberately UNDER-declare: TypeScript's rename runs
 * on the editor's own path, so its provider offers none. An affordance gate
 * therefore asks "bespoke engine OR registry", never the registry alone.
 */

import { normalizeLanguage } from "./language-metadata";
import {
  LanguageServiceRegistry,
  boundProviderResults,
  type LanguageServiceProvider,
  type ProviderCapabilities,
} from "./language-service-provider";
import { createCSharpLanguageProvider } from "./csharp-language-provider";
import { createGoWasmLanguageProvider } from "./go-wasm-toolchain";
import { createLuaLanguageProvider } from "./lua-language-provider";
import { createPythonLanguageProvider } from "./python-language-provider";
import { createShellLanguageProvider } from "./shell-language-provider";
import { createStructuralLanguageProvider } from "./structural-language-provider";
import { createTypeScriptLanguageProvider } from "./typescript-language-provider";

/**
 * Providers contributed from outside this module — a language server bridge, a
 * future extension host. Held at module scope because registering a language is
 * an application-wide fact, not a per-editor one.
 */
const contributed: LanguageServiceProvider[] = [];

/**
 * Add a language. Returns the unregister function.
 *
 * Contributed providers are consulted BEFORE the built-ins, because
 * registration order is priority and a provider added later is the one that
 * means to take over — that is how a real language server replaces a fallback
 * without either of them knowing about the other.
 *
 * Only editors created after this call see the new provider; a provider meant
 * for the whole app should register at module load.
 */
export function registerLanguageServiceProvider(
  provider: LanguageServiceProvider,
): () => void {
  contributed.push(provider);
  return () => {
    const at = contributed.indexOf(provider);
    if (at !== -1) contributed.splice(at, 1);
  };
}

/** Drop every contributed provider. Tests only. */
export function __resetContributedLanguageProvidersForTests(): void {
  contributed.length = 0;
}

export interface EditorLanguageServices {
  /** The composed registry, for a caller that needs the provider itself. */
  readonly registry: LanguageServiceRegistry;
  /**
   * Whether ANY provider claims this language.
   *
   * The gate for affordances: a language nobody serves must not get a rename
   * prompt or a Cmd+click that goes nowhere. Static by design — see the module
   * comment.
   */
  serves(language: string): boolean;
  /**
   * Whether anyone COULD serve `capability` for this language.
   *
   * Static, like `serves`, and for the same reason: affordances must not blink
   * while a kernel boots. But per capability, because providers are no longer
   * all-or-nothing — the outline fallback can navigate a Go file and cannot
   * rename anything in it, and gating F2 on "somebody serves Go" would open a
   * prompt nothing can fulfil. Method presence is the question; whether it can
   * answer THIS instant is settled at the call.
   */
  servesCapability(
    language: string,
    capability: keyof ProviderCapabilities,
  ): boolean;
  /** What can be answered for this language RIGHT NOW. */
  capabilitiesFor(language: string): ProviderCapabilities;
  /** The provider that should answer `capability`, or null. */
  providerFor(
    language: string,
    capability: keyof ProviderCapabilities,
  ): LanguageServiceProvider | null;
  /**
   * Let every provider that loads something load it.
   *
   * Call from an EFFECT, never from render. `capabilitiesFor` is read while
   * rendering, so it has to be a pure query; the loading it used to trigger
   * lives here instead. Resolves once every provider has settled, so a host can
   * re-read capabilities that were "not yet known" a moment ago.
   */
  prime(): Promise<void>;
}

export interface EditorLanguageServicesOptions {
  /**
   * The Python kernel namespace this editor is wired to, if any. Absent means
   * this editor has no Python kernel to ask, so no Python provider is composed
   * in and `serves("python")` is false — matching what the editor could
   * actually do before this registry existed.
   */
  pythonNamespace?: string | null;
}

type PrimeOperation = () => Promise<void>;

/** Capture one provider startup method without invoking an accessor twice. */
function providerPrimeOperation(value: unknown): PrimeOperation | null {
  if (!value || (typeof value !== "object" && typeof value !== "function")) {
    return null;
  }
  let owner: object | null = value as object;
  for (let depth = 0; owner && depth < 8; depth += 1) {
    let descriptor: PropertyDescriptor | undefined;
    try {
      descriptor = Object.getOwnPropertyDescriptor(owner, "prime");
    } catch {
      return null;
    }
    if (descriptor) {
      if (!("value" in descriptor) || typeof descriptor.value !== "function") {
        return null;
      }
      const prime = descriptor.value as () => unknown;
      return () =>
        Promise.resolve()
          .then(() => Reflect.apply(prime, value, []))
          .then(() => undefined);
    }
    try {
      owner = Object.getPrototypeOf(owner) as object | null;
    } catch {
      return null;
    }
  }
  return null;
}

/**
 * Compose the providers one editor should consult.
 *
 * Per editor rather than global because Python's provider is bound to a kernel
 * namespace, and asking another notebook's kernel would answer about another
 * notebook's variables.
 */
export function createEditorLanguageServices(
  options: EditorLanguageServicesOptions = {},
): EditorLanguageServices {
  const registry = new LanguageServiceRegistry();
  const primeOperations: PrimeOperation[] = [];
  // EVERY provider is bounded on the way in, including the two built-ins. They
  // are already bounded by their own engines, so this is a no-op for them — and
  // that is the point: there is one path, and nobody has to remember which
  // providers were the trustworthy ones.
  for (const provider of contributed) {
    const prime = providerPrimeOperation(provider);
    if (prime) primeOperations.push(prime);
    registry.register(boundProviderResults(provider));
  }
  const pythonNamespace = options.pythonNamespace ?? null;
  if (pythonNamespace) {
    const python = createPythonLanguageProvider({
      namespace: () => pythonNamespace,
    });
    const prime = providerPrimeOperation(python);
    if (prime) primeOperations.push(prime);
    registry.register(boundProviderResults(python));
  }
  // Real Go formatting and compiler diagnostics. The provider itself is tiny;
  // its version-pinned WASM worker is created only when a Go buffer asks a
  // question, so opening JavaScript/Python does not download the Go toolchain.
  const go = boundProviderResults(createGoWasmLanguageProvider());
  registry.register(go);
  registry.register(boundProviderResults(createCSharpLanguageProvider()));
  registry.register(boundProviderResults(createLuaLanguageProvider()));
  // The same pinned Bashkit asset used by Run also exposes an analyze-only
  // parser. It lives in a separate Worker and never executes authored source.
  registry.register(boundProviderResults(createShellLanguageProvider()));
  registry.register(boundProviderResults(createTypeScriptLanguageProvider()));
  // LAST: the outline-based fallback for languages with no engine at all. Any
  // provider that actually understands the language is registered ahead of it
  // and wins, so this only ever fills a hole.
  registry.register(boundProviderResults(createStructuralLanguageProvider()));

  return {
    registry,
    serves(language: string): boolean {
      return registry.providersFor(normalizeLanguage(language)).length > 0;
    },
    servesCapability(language, capability) {
      return registry
        .providersFor(normalizeLanguage(language))
        .some((provider) => typeof provider[capability] === "function");
    },
    capabilitiesFor(language: string): ProviderCapabilities {
      return registry.capabilitiesFor(normalizeLanguage(language));
    },
    providerFor(language, capability) {
      const normalized = normalizeLanguage(language);
      // Go formatting is a fixed toolchain operation, not an extensibility
      // hook. A contributed language server may add richer navigation or
      // hover, but it must never displace the pinned gofmt-compatible worker
      // and gain whole-document rewrite authority.
      if (
        normalized === "go" &&
        capability === "formatting" &&
        typeof go.formatting === "function"
      ) {
        return go;
      }
      return registry.providerFor(normalized, capability);
    },
    async prime() {
      // One provider failing to load must not stop the others from being ready.
      await Promise.allSettled(primeOperations.map((prime) => prime()));
    },
  };
}
