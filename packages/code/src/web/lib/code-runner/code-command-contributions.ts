/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The code editor's first extension contribution point: a command in the
 * palette.
 *
 * The workspace already hosts extensions — install lifecycle, version
 * compatibility, a sandboxed widget host, admin approval, supply-chain checks.
 * What it has never had is a way for one to reach the CODE editor. This is that
 * opening, and it is deliberately the smallest one that is still useful: an
 * extension names a command, the palette lists it, the user runs it.
 *
 * Two rules shape everything here.
 *
 * **A command never receives the editor.** Handing over `EditorView` would be
 * handing over the document, the undo history, every extension in the state,
 * and the DOM — root access dressed up as a callback. A command gets
 * {@link CodeCommandHost}: read the selection, read the buffer, replace the
 * selection, ask for a panel. Nothing else exists to be called.
 *
 * **A command gets only what it declared.** The host is built per invocation
 * from the command's own permission list, and a method it did not ask for
 * throws instead of working. So a command that declares only `read.selection`
 * cannot quietly start writing later, and what it may do is auditable from the
 * manifest without reading its code.
 *
 * Everything in this module is pure data and validation, so the rules can be
 * tested without an editor. The editor's side is the palette merge, which
 * labels every contributed command with the extension it came from — a command
 * that looked built-in would be the first thing worth abusing.
 */

/**
 * What a command may do. CLOSED, like the manifest's capability enum: a new
 * permission is a deliberate amendment, not something a contribution can invent
 * by naming one.
 */
export const CODE_COMMAND_PERMISSIONS = [
  /** Read the selected text (empty string when nothing is selected). */
  "read.selection",
  /** Read the whole buffer. Strictly more than `read.selection`. */
  "read.document",
  /** Replace the selection, or insert at the caret when it is empty. */
  "edit.insertText",
  /** Ask the editor to open one of ITS OWN panels. Not arbitrary UI. */
  "ui.openPanel",
] as const;

export type CodeCommandPermission = (typeof CODE_COMMAND_PERMISSIONS)[number];

/**
 * The panels a command may ask for — the editor's own, by name, never arbitrary
 * UI. Only panels that open without arguments are listed: References needs a
 * position, and a command that could ask for it would be asking the editor to
 * guess what it meant.
 */
export const CODE_COMMAND_PANELS = ["problems", "outline", "find"] as const;

export type CodeCommandPanel = (typeof CODE_COMMAND_PANELS)[number];

/**
 * Everything a contributed command can reach. There is no escape hatch to the
 * editor by design: what is not a method here cannot be done.
 *
 * Every method throws {@link CodeCommandPermissionError} when the command did
 * not declare the matching permission — loudly, because a command asking for
 * something it never declared is a bug in the extension, and failing silently
 * would make it look like the editor's.
 */
export interface CodeCommandHost {
  /** The buffer's language id, so one command can serve several languages. */
  readonly language: string;
  /** Requires `read.selection`. */
  getSelection(): string;
  /** Requires `read.document`. */
  getText(): string;
  /** Requires `edit.insertText`. Replaces the selection, or inserts at the caret. */
  replaceSelection(text: string): void;
  /** Requires `ui.openPanel`. */
  openPanel(panel: CodeCommandPanel): void;
}

export class CodeCommandPermissionError extends Error {
  constructor(readonly permission: CodeCommandPermission) {
    super(`code command is missing the "${permission}" permission`);
    this.name = "CodeCommandPermissionError";
  }
}

export interface CodeCommandContribution {
  /**
   * Stable, extension-namespaced id: `<extensionId>.<command>`. Namespaced so
   * two extensions cannot collide, and so a command's origin is legible in a
   * log without a lookup.
   */
  readonly id: string;
  /** The extension that contributed it. */
  readonly extensionId: string;
  /** The extension's display name. Shown next to the command, always. */
  readonly extensionTitle: string;
  /** The command's own title. Supplied by the extension, in its own words. */
  readonly title: string;
  readonly permissions: readonly CodeCommandPermission[];
  /**
   * Language ids this command applies to. Empty or absent means every language
   * — a formatter for one language should say so rather than appear everywhere.
   */
  readonly languages?: readonly string[];
  run(host: CodeCommandHost): void | Promise<void>;
}

/** Bounds. A contribution is untrusted input even after an admin approved it. */
export const MAX_CODE_COMMAND_ID_CHARS = 128;
export const MAX_CODE_COMMAND_TITLE_CHARS = 80;
export const MAX_CODE_COMMANDS_PER_EXTENSION = 24;
/** What one `replaceSelection` may write. A command is not an import path. */
export const MAX_CODE_COMMAND_INSERT_CHARS = 100_000;

/** `acme.wrap-in-try` — a namespace, a dot, a name. */
const COMMAND_ID = /^[a-z0-9][a-z0-9-]*(\.[a-z0-9][a-z0-9-]*)+$/;

/**
 * Text that is safe to render as a label: no control characters, which would
 * let a title carry line breaks or bidi overrides and draw itself somewhere it
 * does not belong in the palette list.
 */
function displayable(value: unknown, maxChars: number): string | null {
  if (typeof value !== "string") return null;
  const trimmed = value.trim();
  if (!trimmed || trimmed.length > maxChars) return null;
  // C0/C1 controls, plus the bidi overrides that would let a label draw
  // itself somewhere it does not belong in the palette list.
  const UNSAFE =
    /[\u0000-\u001F\u007F-\u009F\u200E\u200F\u202A-\u202E\u2066-\u2069]/;
  if (UNSAFE.test(trimmed)) return null;
  return trimmed;
}

export type CodeCommandValidation =
  | { ok: true; command: CodeCommandContribution }
  | { ok: false; reason: string };

/**
 * Validate one contribution.
 *
 * Rejects rather than repairs. A contribution the editor had to fix up is one
 * whose author does not agree with the editor about what it does, and quietly
 * running a corrected version of someone's command is worse than not running it.
 */
export function validateCodeCommandContribution(
  candidate: unknown,
  extensionId: string,
): CodeCommandValidation {
  if (!candidate || typeof candidate !== "object" || Array.isArray(candidate)) {
    return { ok: false, reason: "not an object" };
  }
  const record = candidate as Record<string, unknown>;

  const id = displayable(record.id, MAX_CODE_COMMAND_ID_CHARS);
  if (!id || !COMMAND_ID.test(id)) {
    return { ok: false, reason: "id must be namespaced, e.g. acme.do-thing" };
  }
  // The namespace must be the extension's own: otherwise one extension could
  // publish `git.commit` and be indistinguishable from the extension that owns
  // that namespace in every log and every id-keyed lookup.
  if (!id.startsWith(`${extensionId}.`)) {
    return { ok: false, reason: `id must start with "${extensionId}."` };
  }

  const title = displayable(record.title, MAX_CODE_COMMAND_TITLE_CHARS);
  if (!title) return { ok: false, reason: "title is missing or unrenderable" };

  const extensionTitle = displayable(
    record.extensionTitle,
    MAX_CODE_COMMAND_TITLE_CHARS,
  );
  if (!extensionTitle) {
    return { ok: false, reason: "extensionTitle is missing or unrenderable" };
  }

  if (!Array.isArray(record.permissions)) {
    return { ok: false, reason: "permissions must be an array" };
  }
  const permissions: CodeCommandPermission[] = [];
  for (const permission of record.permissions) {
    if (
      typeof permission !== "string" ||
      !(CODE_COMMAND_PERMISSIONS as readonly string[]).includes(permission)
    ) {
      return { ok: false, reason: `unknown permission ${String(permission)}` };
    }
    if (!permissions.includes(permission as CodeCommandPermission)) {
      permissions.push(permission as CodeCommandPermission);
    }
  }

  let languages: string[] | undefined;
  if (record.languages !== undefined) {
    if (!Array.isArray(record.languages)) {
      return { ok: false, reason: "languages must be an array" };
    }
    languages = [];
    for (const language of record.languages) {
      if (typeof language !== "string" || !language.trim()) {
        return { ok: false, reason: "languages must be non-empty strings" };
      }
      languages.push(language.trim().toLowerCase());
    }
  }

  if (typeof record.run !== "function") {
    return { ok: false, reason: "run must be a function" };
  }

  const frozenPermissions = Object.freeze([...permissions]);
  const frozenLanguages = languages
    ? Object.freeze([...languages])
    : undefined;

  return {
    ok: true,
    command: Object.freeze({
      id,
      extensionId,
      extensionTitle,
      title,
      permissions: frozenPermissions,
      ...(frozenLanguages ? { languages: frozenLanguages } : {}),
      run: record.run as CodeCommandContribution["run"],
    }),
  };
}

/**
 * What the host needs in order to serve a command, expressed WITHOUT the editor.
 *
 * The editor supplies these four closures; the command never learns what is
 * behind them. That is the whole reason this indirection exists rather than
 * passing `EditorView` and trusting everyone.
 */
export interface CodeCommandHostBinding {
  language: string;
  readSelection: () => string;
  readText: () => string;
  replaceSelection: (text: string) => void;
  openPanel: (panel: CodeCommandPanel) => void;
}

/**
 * Build the host one command sees, gated on the permissions it declared.
 *
 * The gate is per invocation and reads the command's own list, so revoking a
 * permission takes effect on the next run with nothing to invalidate.
 */
export function createCodeCommandHost(
  command: Pick<CodeCommandContribution, "permissions">,
  binding: CodeCommandHostBinding,
): CodeCommandHost {
  const require = (permission: CodeCommandPermission): void => {
    if (!command.permissions.includes(permission)) {
      throw new CodeCommandPermissionError(permission);
    }
  };
  return {
    language: binding.language,
    getSelection() {
      require("read.selection");
      return binding.readSelection();
    },
    getText() {
      require("read.document");
      return binding.readText();
    },
    replaceSelection(text: string) {
      require("edit.insertText");
      if (typeof text !== "string") return;
      // A command is not an import path. Truncating would corrupt the buffer
      // with a half-written edit, so an oversized write is refused whole.
      if (text.length > MAX_CODE_COMMAND_INSERT_CHARS) {
        throw new RangeError(
          `code command insert exceeds ${MAX_CODE_COMMAND_INSERT_CHARS} characters`,
        );
      }
      binding.replaceSelection(text);
    },
    openPanel(panel: CodeCommandPanel) {
      require("ui.openPanel");
      if (!(CODE_COMMAND_PANELS as readonly string[]).includes(panel)) return;
      binding.openPanel(panel);
    },
  };
}

/**
 * The commands contributed by installed extensions.
 *
 * Keyed by extension so uninstalling takes its commands with it — a palette
 * entry that outlives its extension is a button that throws.
 */
export class CodeCommandRegistry {
  private readonly byExtension = new Map<
    string,
    readonly CodeCommandContribution[]
  >();
  private readonly listeners = new Set<() => void>();
  private version = 0;

  /**
   * Notify when the set of commands changes.
   *
   * Extensions are installed and uninstalled while editors are open, so a
   * palette that only read this registry at mount would keep offering a
   * command whose extension is gone — the button that throws, again.
   */
  subscribe(listener: () => void): () => void {
    this.listeners.add(listener);
    return () => {
      this.listeners.delete(listener);
    };
  }

  /** A value that changes whenever the commands do. For `useSyncExternalStore`. */
  getVersion(): number {
    return this.version;
  }

  /**
   * Every contributed command, as ONE array whose identity survives until the
   * commands actually change.
   *
   * `useSyncExternalStore` compares snapshots by identity, so a getter that
   * built a fresh array each call would report a change on every render and
   * spin. The cache is invalidated by {@link changed}, which is the only thing
   * that can make this list wrong.
   */
  all(): readonly CodeCommandContribution[] {
    if (!this.snapshot) {
      this.snapshot = Object.freeze(
        Array.from(this.byExtension.values()).flat(),
      );
    }
    return this.snapshot;
  }

  /**
   * Whether this exact validated command still belongs to its extension.
   *
   * Identity is intentional: reinstalling an extension may reuse an id while
   * changing its permissions or implementation. A host retained by the old
   * command must not regain authority merely because the new command has the
   * same label.
   */
  isRegistered(command: CodeCommandContribution): boolean {
    return (
      this.byExtension.get(command.extensionId)?.includes(command) ?? false
    );
  }

  private snapshot: readonly CodeCommandContribution[] | null = null;

  private changed(): void {
    this.version += 1;
    this.snapshot = null;
    for (const listener of this.listeners) listener();
  }

  /**
   * Register an extension's commands, replacing anything it registered before.
   *
   * Invalid contributions are DROPPED and reported, not thrown: one malformed
   * command in a manifest should cost that command, not the extension's other
   * ones, and certainly not the editor.
   */
  register(
    extensionId: string,
    candidates: readonly unknown[],
  ): { accepted: CodeCommandContribution[]; rejected: string[] } {
    const accepted: CodeCommandContribution[] = [];
    const rejected: string[] = [];
    const seen = new Set<string>();
    for (const candidate of candidates.slice(
      0,
      MAX_CODE_COMMANDS_PER_EXTENSION,
    )) {
      const result = validateCodeCommandContribution(candidate, extensionId);
      if (!result.ok) {
        rejected.push(result.reason);
        continue;
      }
      if (seen.has(result.command.id)) {
        rejected.push(`duplicate id ${result.command.id}`);
        continue;
      }
      seen.add(result.command.id);
      accepted.push(result.command);
    }
    if (candidates.length > MAX_CODE_COMMANDS_PER_EXTENSION) {
      rejected.push(
        `only the first ${MAX_CODE_COMMANDS_PER_EXTENSION} commands were read`,
      );
    }
    // Do not store the mutable array returned to the installer. Without this
    // copy, a caller could append an unvalidated command after registration and
    // bypass both the per-extension bound and the registry change signal.
    this.byExtension.set(extensionId, Object.freeze([...accepted]));
    this.changed();
    return { accepted, rejected };
  }

  /** Remove an extension's commands. Uninstall, or a kill broadcast. */
  unregister(extensionId: string): void {
    if (this.byExtension.delete(extensionId)) this.changed();
  }

  /** Every contributed command applicable to `language`, in extension order. */
  commandsFor(language: string): CodeCommandContribution[] {
    return this.all().filter((command) =>
      codeCommandAppliesTo(command, language),
    );
  }

  clear(): void {
    if (this.byExtension.size === 0) return;
    this.byExtension.clear();
    this.changed();
  }
}

/**
 * Whether a command applies to `language`. Absent or empty `languages` means
 * every language — a formatter for one language should say which.
 */
export function codeCommandAppliesTo(
  command: Pick<CodeCommandContribution, "languages">,
  language: string,
): boolean {
  if (!command.languages || command.languages.length === 0) return true;
  return command.languages.includes((language || "").trim().toLowerCase());
}

/**
 * The registry the editor reads. One per application: which extensions are
 * installed is not a per-editor fact.
 */
export const codeCommandRegistry = new CodeCommandRegistry();

/**
 * How a contributed command is labelled in the palette.
 *
 * VS Code's convention, and the reason for it: a command's origin belongs in
 * the label itself, not in a chip a busy user reads past. An extension that
 * names its command "Format Document" still reads as its own.
 */
export function codeCommandPaletteLabel(
  command: Pick<CodeCommandContribution, "extensionTitle" | "title">,
): string {
  return `${command.extensionTitle}: ${command.title}`;
}
