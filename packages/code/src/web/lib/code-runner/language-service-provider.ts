/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/** A half-open range of the buffer, in UTF-16 code units (CodeMirror's units). */
export interface ProviderSpan {
  from: number;
  to: number;
}

export interface ProviderCompletion {
  label: string;
  /** CodeMirror's completion type ("function", "variable", …); free-form. */
  kind?: string;
  detail?: string;
  /** Documentation as PLAIN TEXT. Rendering is the caller's business. */
  documentation?: string;
}

export interface ProviderHover {
  /** Plain text; a provider that only has HTML should strip it, not embed it. */
  contents: string;
  span?: ProviderSpan;
}

export interface ProviderDiagnostic {
  span: ProviderSpan;
  message: string;
  severity: "error" | "warning" | "info";
  /** Stable identifier for the rule, when the language has one. */
  code?: string;
}

export interface ProviderLocation extends ProviderSpan {
  /**
   * Which buffer this location is in. `null` means "this one" — the common
   * case. A cross-file provider returns the host's own id for the other buffer.
   */
  bufferId: string | null;
}

/**
 * Parameter hints for the call the caret sits in. PLAIN TEXT, like everything
 * else here: the editor escapes it and decides how to mark the active one.
 */
export interface ProviderSignature {
  /** The whole signature, e.g. `add(a, b, *, c=1)`. */
  label: string;
  /** Per-parameter text, in order. */
  parameters: string[];
  /** 0-based index of the argument slot the caret is in. */
  activeParameter: number;
}

/** One text edit: replace `span` in `bufferId` with `newText`. */
export interface ProviderTextEdit {
  bufferId: string | null;
  span: ProviderSpan;
  newText: string;
}

/** A rename is a list of text edits; the name is kept for what calls it. */
export type ProviderRenameEdit = ProviderTextEdit;

/**
 * One offered fix or refactor — a title the user picks and the edits it makes.
 *
 * Applied all-or-nothing, like every other document change here: an action the
 * caller could only partly apply is one it should not apply at all.
 */
export interface ProviderCodeAction {
  /** Shown in the Quick Fix menu. PLAIN TEXT. */
  title: string;
  edits: ProviderTextEdit[];
}

/**
 * What a provider can do. Absent or false means the editor must not offer it.
 *
 * Kept as data rather than "does the method exist", so a provider can advertise
 * a capability it only supports for some buffers (a server that is still
 * starting, a language whose parser failed) by flipping the flag at runtime.
 */
export interface ProviderCapabilities {
  completions?: boolean;
  hover?: boolean;
  definition?: boolean;
  references?: boolean;
  diagnostics?: boolean;
  rename?: boolean;
  formatting?: boolean;
  signatureHelp?: boolean;
  typeDefinition?: boolean;
  implementation?: boolean;
  codeActions?: boolean;
}

const PROVIDER_CAPABILITY_KEYS = [
  "completions",
  "hover",
  "definition",
  "references",
  "diagnostics",
  "rename",
  "formatting",
  "signatureHelp",
  "typeDefinition",
  "implementation",
  "codeActions",
] as const satisfies readonly (keyof ProviderCapabilities)[];

export interface LanguageServiceRequest {
  /** The live buffer text. Providers must not assume it is on disk anywhere. */
  text: string;
  /** Caret offset in UTF-16 code units. */
  offset: number;
  /** Opaque id of the buffer, for providers that see more than one. */
  bufferId?: string;
  /** Aborts a slow provider (a language server, a WASM cold start). */
  signal?: AbortSignal;
}

export interface LanguageServiceProvider {
  /** Stable id, e.g. "typescript" or "gopls". Used in logs and in tests. */
  readonly id: string;
  /** Editor language ids this provider serves, e.g. ["javascript", "jsx"]. */
  readonly languages: readonly string[];
  capabilities(language: string): ProviderCapabilities;

  completions?(
    request: LanguageServiceRequest & { language: string },
  ): Promise<ProviderCompletion[]>;
  hover?(
    request: LanguageServiceRequest & { language: string },
  ): Promise<ProviderHover | null>;
  definition?(
    request: LanguageServiceRequest & { language: string },
  ): Promise<ProviderLocation | null>;
  references?(
    request: LanguageServiceRequest & { language: string },
  ): Promise<ProviderLocation[]>;
  diagnostics?(
    request: Omit<LanguageServiceRequest, "offset"> & { language: string },
  ): Promise<ProviderDiagnostic[]>;
  rename?(
    request: LanguageServiceRequest & { language: string; newName: string },
  ): Promise<ProviderRenameEdit[]>;
  /**
   * Pretty-print the whole buffer, or null when it will not.
   *
   * Whole text rather than a list of edits, because that is what this editor's
   * formatters already produce and what its one apply path already takes. A
   * provider that computes ranged edits composes them itself; a provider that
   * cannot format THIS buffer (a parse error, a formatter that is not installed)
   * returns null so the caller leaves the document exactly as it was — the same
   * never-corrupt contract every other formatter here holds to.
   */
  formatting?(
    request: Omit<LanguageServiceRequest, "offset"> & { language: string },
  ): Promise<string | null>;
  signatureHelp?(
    request: LanguageServiceRequest & { language: string },
  ): Promise<ProviderSignature | null>;
  /**
   * Where the TYPE of the symbol under the caret is declared — distinct from
   * `definition`, which goes to the symbol itself.
   */
  typeDefinition?(
    request: LanguageServiceRequest & { language: string },
  ): Promise<ProviderLocation | null>;
  /** Where an abstract symbol is actually implemented. */
  implementation?(
    request: LanguageServiceRequest & { language: string },
  ): Promise<ProviderLocation | null>;
  /**
   * Fixes and refactors offered for the selection. `offset` is its start and
   * `to` its end; they are equal for a bare caret.
   */
  codeActions?(
    request: LanguageServiceRequest & { language: string; to: number },
  ): Promise<ProviderCodeAction[]>;
}

/**
 * The language → provider map.
 *
 * Registration order is priority: the first provider that claims a language AND
 * advertises the capability wins. That lets a real language server take over a
 * language a fallback already serves, without either of them knowing about the
 * other.
 */
export class LanguageServiceRegistry {
  private readonly providers: LanguageServiceProvider[] = [];

  register(provider: LanguageServiceProvider): () => void {
    this.providers.push(provider);
    return () => {
      const at = this.providers.indexOf(provider);
      if (at !== -1) this.providers.splice(at, 1);
    };
  }

  /** Every provider serving `language`, in registration order. */
  providersFor(language: string): LanguageServiceProvider[] {
    return this.providers.filter((provider) =>
      provider.languages.includes(language),
    );
  }

  /**
   * The provider that should answer `capability` for `language`, or null.
   *
   * Both halves matter: a provider must serve the language AND advertise the
   * capability AND actually implement the method. A provider that claims a
   * capability it did not implement is a bug in that provider, and returning
   * null here turns it into a missing menu item rather than a crash.
   */
  providerFor(
    language: string,
    capability: keyof ProviderCapabilities,
  ): LanguageServiceProvider | null {
    for (const provider of this.providersFor(language)) {
      if (!provider.capabilities(language)[capability]) continue;
      if (typeof provider[capability] !== "function") continue;
      return provider;
    }
    return null;
  }

  /** What the editor may offer for `language` right now. */
  capabilitiesFor(language: string): ProviderCapabilities {
    const merged: ProviderCapabilities = {};
    for (const capability of [
      "completions",
      "hover",
      "definition",
      "references",
      "diagnostics",
      "rename",
      "formatting",
      "signatureHelp",
      "typeDefinition",
      "implementation",
      "codeActions",
    ] as const) {
      if (this.providerFor(language, capability)) merged[capability] = true;
    }
    return merged;
  }
}

/**
 * What a provider is allowed to hand back.
 *
 * Every other engine on this surface is bounded where it crosses into the
 * editor — jedi at the worker's message boundary, the merge view at its own
 * limits table — because an answer nobody sized can wedge a popup, a panel or
 * the whole document. A provider is the newest such edge, and the one most
 * likely to be a bridge to something remote, so it gets the same treatment.
 */
export const PROVIDER_LIMITS = {
  /** Suggestions kept. The engines here return tens; this is headroom, not a target. */
  completions: 200,
  completionLabelChars: 256,
  completionKindChars: 64,
  completionDetailChars: 4_096,
  /** Hover text. jedi's own signature + docstring caps sum to under this. */
  hoverChars: 8_000,
  /** Locations kept, matching the Python protocol's own per-message item cap. */
  locations: 500,
  locationBufferIdChars: 1_024,
  diagnostics: 500,
  diagnosticMessageChars: 1_000,
  diagnosticCodeChars: 256,
  /** Rename edits. Beyond this a single-buffer rename is not what is happening. */
  renameEdits: 2_000,
  editTextChars: 4 * 1024 * 1024,
  editAggregateChars: 4 * 1024 * 1024,
  /** Formatted document size, matching the merge view's output ceiling. */
  formattedChars: 4 * 1024 * 1024,
  /** Signature text, and how many parameters one hint may list. */
  signatureChars: 4_096,
  signatureParameters: 128,
  /** Offered actions, and the edits any ONE of them may make. */
  codeActions: 50,
  codeActionTitleChars: 200,
  codeActionEdits: 1_000,
} as const;

/** Clamp a string, or null when it is not one. */
function boundedText(value: unknown, maxChars: number): string | null {
  if (typeof value !== "string") return null;
  return value.length > maxChars ? value.slice(0, maxChars) : value;
}

function ownDataFields(
  value: unknown,
  keys: readonly string[],
): Record<string, unknown> | null {
  if (!value || typeof value !== "object") return null;
  try {
    if (Array.isArray(value)) return null;
  } catch {
    return null;
  }
  const fields: Record<string, unknown> = Object.create(null);
  for (const key of keys) {
    let descriptor: PropertyDescriptor | undefined;
    try {
      descriptor = Object.getOwnPropertyDescriptor(value, key);
    } catch {
      return null;
    }
    if (!descriptor) continue;
    if (!("value" in descriptor)) return null;
    fields[key] = descriptor.value;
  }
  return fields;
}

function ownArrayWindow(
  value: unknown,
  maxItems: number,
): { items: unknown[]; length: number } | null {
  let isArray = false;
  try {
    isArray = Array.isArray(value);
  } catch {
    return null;
  }
  if (!isArray) return null;
  let lengthDescriptor: PropertyDescriptor | undefined;
  try {
    lengthDescriptor = Object.getOwnPropertyDescriptor(value, "length");
  } catch {
    return null;
  }
  if (
    !lengthDescriptor ||
    !("value" in lengthDescriptor) ||
    !Number.isSafeInteger(lengthDescriptor.value) ||
    lengthDescriptor.value < 0
  ) {
    return null;
  }
  const length = lengthDescriptor.value;
  const items: unknown[] = [];
  for (let index = 0; index < Math.min(length, maxItems); index += 1) {
    let descriptor: PropertyDescriptor | undefined;
    try {
      descriptor = Object.getOwnPropertyDescriptor(value, String(index));
    } catch {
      return null;
    }
    if (!descriptor) {
      items.push(undefined);
      continue;
    }
    if (!("value" in descriptor)) return null;
    items.push(descriptor.value);
  }
  return { items, length };
}

function boundedSpan(value: unknown): ProviderSpan | null {
  const raw = ownDataFields(value, ["from", "to"]);
  if (
    !raw ||
    !Number.isSafeInteger(raw.from) ||
    !Number.isSafeInteger(raw.to) ||
    Number(raw.from) < 0 ||
    Number(raw.to) < Number(raw.from)
  ) {
    return null;
  }
  return { from: Number(raw.from), to: Number(raw.to) };
}

function boundedLocation(value: unknown): ProviderLocation | null {
  const raw = ownDataFields(value, ["bufferId", "from", "to"]);
  if (!raw) return null;
  const span = boundedSpan(raw);
  if (!span) return null;
  if (
    raw.bufferId !== null &&
    (typeof raw.bufferId !== "string" ||
      raw.bufferId.length > PROVIDER_LIMITS.locationBufferIdChars)
  ) {
    return null;
  }
  return { ...span, bufferId: raw.bufferId as string | null };
}

function boundedTextEdit(value: unknown): ProviderTextEdit | null {
  const raw = ownDataFields(value, ["bufferId", "span", "newText"]);
  if (!raw) return null;
  const span = boundedSpan(raw.span);
  if (
    !span ||
    (raw.bufferId !== null &&
      (typeof raw.bufferId !== "string" ||
        raw.bufferId.length > PROVIDER_LIMITS.locationBufferIdChars)) ||
    typeof raw.newText !== "string" ||
    raw.newText.length > PROVIDER_LIMITS.editTextChars
  ) {
    return null;
  }
  return {
    bufferId: raw.bufferId as string | null,
    span,
    newText: raw.newText,
  };
}

function boundedTextEdits(
  value: unknown,
  limit: number,
): ProviderTextEdit[] | null {
  const raw = ownArrayWindow(value, limit);
  if (!raw || raw.length > limit) return null;
  const edits: ProviderTextEdit[] = [];
  let aggregateChars = 0;
  for (const candidate of raw.items) {
    const edit = boundedTextEdit(candidate);
    if (!edit) return null;
    aggregateChars += edit.newText.length;
    if (aggregateChars > PROVIDER_LIMITS.editAggregateChars) return null;
    edits.push(edit);
  }
  return edits;
}

function boundedCapabilities(value: unknown): ProviderCapabilities {
  const raw = ownDataFields(value, PROVIDER_CAPABILITY_KEYS);
  if (!raw) return {};
  const capabilities: ProviderCapabilities = {};
  for (const key of PROVIDER_CAPABILITY_KEYS) {
    if (raw[key] === true) capabilities[key] = true;
  }
  return capabilities;
}

















export function boundProviderResults(
  provider: LanguageServiceProvider,
): LanguageServiceProvider {
  const bounded: LanguageServiceProvider = {
    id: provider.id,
    languages: provider.languages,
    capabilities(language: string) {
      try {
        return boundedCapabilities(provider.capabilities(language));
      } catch {
        return {};
      }
    },
  };

  if (provider.completions) {
    bounded.completions = async (request) => {
      const reported = ownArrayWindow(
        await provider.completions!(request),
        PROVIDER_LIMITS.completions,
      );
      if (!reported) return [];
      return reported.items
        .map((candidate) => {
          const item = ownDataFields(candidate, [
            "label",
            "kind",
            "detail",
            "documentation",
          ]);
          if (!item) return null;
          const label = boundedText(
            item.label,
            PROVIDER_LIMITS.completionLabelChars,
          );
          if (!label) return null;
          const kind = boundedText(
            item.kind,
            PROVIDER_LIMITS.completionKindChars,
          );
          const detail = boundedText(
            item.detail,
            PROVIDER_LIMITS.completionDetailChars,
          );
          const documentation = boundedText(
            item.documentation,
            PROVIDER_LIMITS.completionDetailChars,
          );
          return {
            label,
            ...(kind ? { kind } : {}),
            ...(detail ? { detail } : {}),
            ...(documentation ? { documentation } : {}),
          };
        })
        .filter((item): item is ProviderCompletion => item !== null);
    };
  }

  if (provider.hover) {
    bounded.hover = async (request) => {
      const raw = ownDataFields(await provider.hover!(request), [
        "contents",
        "span",
      ]);
      if (!raw) return null;
      const contents = boundedText(raw.contents, PROVIDER_LIMITS.hoverChars);
      if (!contents) return null;
      if (raw.span === undefined) return { contents };
      const span = boundedSpan(raw.span);
      return span ? { contents, span } : null;
    };
  }

  // Normalize the location's structure here. Every caller still resolves its
  // offsets against its own live document, which is the only thing that knows
  // how long the buffer is by the time the answer lands.
  if (provider.definition) {
    bounded.definition = async (request) =>
      boundedLocation(await provider.definition!(request));
  }
  if (provider.typeDefinition) {
    bounded.typeDefinition = async (request) =>
      boundedLocation(await provider.typeDefinition!(request));
  }
  if (provider.implementation) {
    bounded.implementation = async (request) =>
      boundedLocation(await provider.implementation!(request));
  }

  if (provider.references) {
    bounded.references = async (request) => {
      const reported = ownArrayWindow(
        await provider.references!(request),
        PROVIDER_LIMITS.locations,
      );
      if (!reported) return [];
      return reported.items
        .map(boundedLocation)
        .filter((location): location is ProviderLocation => location !== null);
    };
  }

  if (provider.diagnostics) {
    bounded.diagnostics = async (request) => {
      const reported = ownArrayWindow(
        await provider.diagnostics!(request),
        PROVIDER_LIMITS.diagnostics,
      );
      if (!reported) return [];
      return reported.items
        .map((candidate) => {
          const diagnostic = ownDataFields(candidate, [
            "span",
            "message",
            "severity",
            "code",
          ]);
          if (!diagnostic) return null;
          const span = boundedSpan(diagnostic.span);
          const message = boundedText(
            diagnostic.message,
            PROVIDER_LIMITS.diagnosticMessageChars,
          );
          const severity = diagnostic.severity;
          if (
            !span ||
            !message ||
            (severity !== "error" &&
              severity !== "warning" &&
              severity !== "info")
          ) {
            return null;
          }
          const code = boundedText(
            diagnostic.code,
            PROVIDER_LIMITS.diagnosticCodeChars,
          );
          if (diagnostic.code !== undefined && code === null) return null;
          return {
            span,
            message,
            severity,
            ...(code ? { code } : {}),
          };
        })
        .filter(
          (diagnostic): diagnostic is ProviderDiagnostic => diagnostic !== null,
        );
    };
  }

  if (provider.rename) {
    bounded.rename = async (request) => {
      // Refused whole: a truncated rename rewrites some occurrences of a symbol
      // and leaves the rest, which reads as working and is not.
      return (
        boundedTextEdits(
          await provider.rename!(request),
          PROVIDER_LIMITS.renameEdits,
        ) ?? []
      );
    };
  }

  if (provider.signatureHelp) {
    bounded.signatureHelp = async (request) => {
      const signature = ownDataFields(await provider.signatureHelp!(request), [
        "label",
        "parameters",
        "activeParameter",
      ]);
      if (!signature) return null;
      const label = boundedText(
        signature.label,
        PROVIDER_LIMITS.signatureChars,
      );
      if (!label) return null;
      const rawParameters = ownArrayWindow(
        signature.parameters,
        PROVIDER_LIMITS.signatureParameters,
      );
      if (!rawParameters) return null;
      const parameters = rawParameters.items
        .map((parameter) => boundedText(parameter, PROVIDER_LIMITS.signatureChars))
        .filter((parameter): parameter is string => parameter !== null);
      // A hint that points past its own parameter list would bold nothing, or
      // the wrong thing; clamp it into the list it came with.
      const activeParameter =
        parameters.length === 0
          ? 0
          : Math.max(
              0,
              Math.min(
                typeof signature.activeParameter === "number" &&
                  Number.isFinite(signature.activeParameter)
                  ? Math.trunc(signature.activeParameter)
                  : 0,
                parameters.length - 1,
              ),
            );
      return { label, parameters, activeParameter };
    };
  }

  if (provider.codeActions) {
    bounded.codeActions = async (request) => {
      const reported = ownArrayWindow(
        await provider.codeActions!(request),
        PROVIDER_LIMITS.codeActions,
      );
      if (!reported) return [];
      return reported.items
        .map((candidate) => {
          const action = ownDataFields(candidate, ["title", "edits"]);
          if (!action) return null;
          const title = boundedText(
            action.title,
            PROVIDER_LIMITS.codeActionTitleChars,
          );
          if (!title) return null;
          // Per action, all-or-nothing: an action too big to apply safely is
          // dropped whole, and the others are still offered. Truncating ITS
          // edits would leave a half-applied fix behind the user's back.
          const edits = boundedTextEdits(
            action.edits,
            PROVIDER_LIMITS.codeActionEdits,
          );
          return !edits || edits.length === 0 ? null : { title, edits };
        })
        .filter((action): action is ProviderCodeAction => action !== null);
    };
  }

  if (provider.formatting) {
    bounded.formatting = async (request) => {
      const formatted = await provider.formatting!(request);
      if (typeof formatted !== "string") return null;
      // Declined rather than clipped: half a file is not a formatted file.
      if (formatted.length > PROVIDER_LIMITS.formattedChars) return null;
      return formatted;
    };
  }

  return bounded;
}
