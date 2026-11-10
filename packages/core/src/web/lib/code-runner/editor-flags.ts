/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Feature flag for the textarea -> CodeMirror 6 editor migration.
 *
 * Stage 1 shipped the new {@link CodeMirrorEditor} alongside the existing
 * {@link CodeEditor} (textarea overlay). Stage 2 brought CodeMirrorEditor to
 * parity (find/replace, go-to-line, autocomplete, minimap, JSON+JS linters).
 *
 * Stage 3 ACTIVATED CodeMirror 6: this flag now defaults to `true`, so the
 * Code surfaces render the CM6 editor. The textarea {@link CodeEditor}
 * remains in the tree behind the `false` branch as the instant-rollback
 * fallback — flip this back to `false` to restore the textarea engine with zero
 * other changes.
 *
 * BOTH engines stay under regression test. The window-level component tests
 * (`tests/unit/code-window.test.tsx`) are engine-parametrized and mock THIS
 * module to force the flag value per describe block. The original textarea-contract assertions
 * (`(getByTestId("code-editor") as HTMLTextAreaElement).value`) run with the
 * flag forced `false` so the fallback keeps its window-level coverage, while
 * parallel CodeMirror describe blocks force the flag `true` and assert the same
 * behaviours through the CM6 surface (document text via `.cm-content` /
 * `EditorView` and run callbacks).
 */
export const USE_CODEMIRROR_EDITOR = true;

/**
 * Feature flag for the real-time COLLABORATIVE editing foundation in the MRNote
 * TipTap document editor (Stage 1, Yjs CRDT).
 *
 * DEFAULT OFF. The flag is intentionally additive: while `false`, NoteEditor
 * builds its extension list exactly as before (StarterKit history intact, the
 * existing autosave content/load path unchanged) and the Yjs + TipTap
 * collaboration extensions are NEVER loaded — so single-user editing is
 * byte-for-byte the current behaviour with ZERO regression.
 *
 * When flipped to `true`, NoteEditor:
 *   1. creates a {@link createCollabSession} (local-only unless a room+URL is
 *      configured — see lib/collab/collab-provider.ts),
 *   2. disables StarterKit's built-in `history` (TipTap's `Collaboration`
 *      extension replaces undo/redo with the Yjs undo manager — they cannot
 *      coexist), and
 *   3. appends the `Collaboration` + `CollaborationCaret` extensions
 *      (lib/collab/collab-extensions.ts) bound to the session's Y.Doc +
 *      awareness, rendering remote cursors + presence (name + color).
 *
 * The history swap is the ONLY behavioural change, and only on the ON path.
 *
 * Stage 2 (LIVE): createCollabSession now auto-wires the app's ticketed
 * websocket transport (lib/collab/ws-provider.ts + the page-yjs room relay)
 * whenever a socket origin is resolvable, so sessions with peers converge in
 * real time. Environments without a socket origin (headless tests, guest
 * demos, misconfigured deploys) still degrade to the local-only session.
 * The headless Yjs merge (two Y.Docs converging via applyUpdate) is proven
 * by lib/collab/__tests__/collab-provider.test.ts.
 *
 * Default ON. `NEXT_PUBLIC_COLLAB_EDITING=0` is the build-time kill switch
 * back to the single-user editing path.
 */
export const USE_COLLAB_EDITING =
  process.env.NEXT_PUBLIC_COLLAB_EDITING !== "0";










export const USE_LEGACY_ROOM_LOGICAL_BINDING =
  process.env.NEXT_PUBLIC_LEGACY_ROOM_LOGICAL_BINDING !== "0";

/**
 * Share one local text model, undo history and autosave writer between
 * duplicate standalone CodeWindow views of the same page.
 *
 * This is deliberately separate from {@link USE_COLLAB_EDITING}: it never
 * creates a room or network provider and does not affect Docs. The
 * local page session exists because duplicate Code windows are already
 * available through the note titlebar; without one, two full-body autosave
 * drafts can race inside the debounce window.
 *
 * Rollback keeps the existing fail-closed CAS/conflict behavior: set this false
 * to restore independent per-window buffers without changing either editor
 * engine.
 */
export const ENABLE_CODE_PAGE_SHARED_BUFFER: boolean = true;




























































export const ENABLE_INLINE_PROPOSAL_DIFF: boolean = true;







































// Retain the implementation and saved configuration; normal Run stays enabled.
export const ENABLE_CODE_DEBUG_V1: boolean = false;

/**
 * Feature flag for loading a code page's GRAMMAR on demand rather than shipping
 * every grammar to every code page.
 *
 * DEFAULT ON after the isolated bundle and behavior gates. The OFF path still
 * dynamically loads the former all-language implementation after mounting the
 * same usable gap fallback, so rollback does not pin every parser table into
 * CodeWindow's initial chunk.
 *
 * Measured on the installed packages: the twelve `@codemirror/lang-*` grammars
 * with their `@lezer` parser tables are ~1.77 MB of unminified JavaScript and
 * the eight `legacy-modes` stream parsers add ~99 KB. `CodeWindow` is already
 * behind `next/dynamic`, so this is not in the app's first load — it is in the
 * chunk fetched when a user opens their FIRST code page of a session, and every
 * page after that reuses it. So the cost is once per session, and the saving is
 * on that first open: a Python user pays for Python instead of also paying for
 * PHP, Rust, Java and C++.
 *
 * In both modes the language sits in a CM6 `Compartment`, seeded with a small
 * comment/indent/fold fallback until the grammar resolves. ON fetches one
 * language; OFF fetches the bundled fallback. Within a session only the first
 * page of a language has a gap because a cache hit is synchronous. A failed
 * chunk leaves an explicitly degraded but editable page and remains retryable.
 *
 * Typed `boolean` (not the literal `false`) so flipping it is a one-token change
 * and tests can force either path.
 */
export const ENABLE_LAZY_LANGUAGE_GRAMMARS: boolean = true;
