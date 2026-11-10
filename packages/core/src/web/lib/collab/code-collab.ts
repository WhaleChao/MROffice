/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * CodeMirror 6 binding for the real-time collaborative editing foundation.
 *
 * This is the code-editor analogue of collab-extensions.ts (the TipTap document
 * editor's collab wiring). It is the ONLY place that imports the
 * `y-codemirror.next` package, and it is consumed by {@link CodeMirrorEditor}
 * ONLY when the `USE_COLLAB_EDITING` flag is ON — so the OFF path never loads
 * the yCollab extension and the editor's existing `value`/`onChange` path is
 * byte-for-byte the current single-user behaviour with ZERO regression.
 *
 * Split, deliberately, into:
 *  - PURE helpers (no React / no DOM): {@link getCodeYText},
 *    {@link buildCodeCollabExtensions}, {@link codeCollabAwarenessUser}. These
 *    are exhaustively unit-testable headlessly (Y.Doc convergence, [] when no
 *    session, stable Y.Text identity).
 *  - The side-effectful Y.Doc/provider lifecycle stays in collab-provider.ts;
 *    the editor wiring stays in the component.
 *
 * Security: the awareness `user` payload broadcast to (and rendered by) every
 * peer is routed through the SAME strict sanitizers as the document editor
 * (collab-state.ts) — the display name is control-stripped + length-capped and
 * the cursor color is forced through the strict `#rrggbb` hex gate. y-codemirror
 * derives a translucent selection color as `color + "33"`; appending to a
 * canonical 6-digit hex keeps that derivation a valid 8-digit hex, so an
 * injection payload can never reach the remote DOM as CSS.
 */

import type { Awareness } from "y-protocols/awareness";
import type { Extension } from "@codemirror/state";
import type * as Y from "yjs";
import { yCollab } from "y-codemirror.next";

import { isLocalUsersOtherConnection, sanitizeCollabUser, type CollabUser } from "./collab-state";
import type { CollabSession } from "./collab-provider";
import { CODE_COLLAB_TEXT_KEY } from "./code-collab-model";

/**
 * Default Y.Text field name a code document binds to. A single, stable key per
 * doc means every peer editing the same code page shares one CRDT text type.
 * Mirrors the y-codemirror.next examples (`ydoc.getText("codemirror")`).
 */
export { CODE_COLLAB_TEXT_KEY } from "./code-collab-model";

/**
 * Resolve the shared {@link Y.Text} a code editor binds to. `Y.Doc.getText` is
 * idempotent: calling it again with the same key returns the SAME Y.Text
 * instance, so this is safe to call on every render/effect and always yields a
 * stable shared type. An empty/blank key falls back to {@link CODE_COLLAB_TEXT_KEY}
 * so a misconfigured caller still binds to a single well-known field rather than
 * the unnamed root.
 */
export function getCodeYText(ydoc: Y.Doc, key?: string | null): Y.Text {
  const field =
    typeof key === "string" && key.trim().length > 0 ? key : CODE_COLLAB_TEXT_KEY;
  return ydoc.getText(field);
}

/**
 * Map arbitrary identity input to the sanitized awareness `user` payload that
 * y-codemirror.next reads when rendering remote carets/selections
 * (`state.user.{name,color}`; `colorLight` is derived as `color + "33"`). Reuses
 * the document editor's sanitizers so the name is safe + capped and the color is
 * a canonical `#rrggbb` value.
 */
export function codeCollabAwarenessUser(input: {
  name?: unknown;
  color?: unknown;
}): CollabUser {
  return sanitizeCollabUser(input);
}

export interface CodeCollabExtensionOptions {
  /** The shared CRDT text the editor binds to (from {@link getCodeYText}). */
  ytext: Y.Text;
  /** Awareness instance carrying local + remote presence (cursors/names/colors). */
  awareness: Awareness;
  /**
   * Optional Yjs undo manager. When supplied, yCollab installs its undo/redo
   * keymap scoped to local changes; when omitted (or `false`) CM6's own history
   * keeps handling undo/redo. Passing `false` is the explicit "no Yjs undo".
   */
  undoManager?: Y.UndoManager | false;
}

/**
 * The awareness yCollab draws remote selections from, minus this user's own
 * other connections (another tab, window or device): y-codemirror draws every
 * state but its own clientID, so those read as a stray caret carrying the
 * user's own name. Only `getStates` is narrowed; every other member, including
 * the local state the binding publishes and the change events it listens to,
 * is the session's awareness itself.
 */
function withoutOwnOtherConnections(awareness: Awareness): Awareness {
  return new Proxy(awareness, {
    get(target, key) {
      if (key === "getStates") {
        return () => {
          const states = target.getStates() as Map<number, Record<string, unknown>>;
          const own = [...states].filter(([clientId, state]) =>
            clientId !== target.clientID &&
            isLocalUsersOtherConnection(states, target.clientID, state));
          if (own.length === 0) return states;
          const narrowed = new Map(states);
          for (const [clientId] of own) narrowed.delete(clientId);
          return narrowed;
        };
      }
      const value = Reflect.get(target, key, target);
      return typeof value === "function" ? value.bind(target) : value;
    },
  });
}

/**
 * Build the yCollab extension(s) for a live code-collab binding. Returns the
 * `yCollab(ytext, awareness, { undoManager })` extension wrapped in an array so
 * the caller can spread it into CM6's extension list exactly like the editor's
 * other extension groups.
 *
 * Returns an EMPTY array when (defensively) no `ytext` or no `awareness` is
 * available, so a caller that flips the flag on before a session exists simply
 * gets the plain (non-collab) editor — never a thrown error.
 */
export function buildCodeCollabExtensions(
  options: CodeCollabExtensionOptions | null | undefined,
): Extension[] {
  if (!options?.ytext || !options?.awareness) return [];
  const { ytext, awareness, undoManager } = options;
  return [
    yCollab(ytext, withoutOwnOtherConnections(awareness), {
      undoManager: undoManager ?? false,
    }),
  ];
}

/**
 * Convenience: derive the full code-collab extension list from a
 * {@link CollabSession}. Resolves the shared Y.Text by `key`, then delegates to
 * {@link buildCodeCollabExtensions}. Returns `[]` for a missing/torn-down
 * session so the editor degrades to single-user with no extra branching at the
 * call site.
 */
export function codeCollabExtensionsForSession(
  session: CollabSession | null | undefined,
  options?: { key?: string | null; undoManager?: Y.UndoManager | false },
): Extension[] {
  if (!session?.ydoc || !session.awareness) return [];
  const ytext = getCodeYText(session.ydoc, options?.key);
  return buildCodeCollabExtensions({
    ytext,
    awareness: session.awareness,
    undoManager: options?.undoManager,
  });
}
