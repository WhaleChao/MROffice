/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { EditorView } from "@codemirror/view";

















export type CodeSoftKey =
  | {
      kind: "key";
      key: "Escape" | "Tab" | "ArrowLeft" | "ArrowUp" | "ArrowDown" | "ArrowRight";
      shift?: boolean;
      face: string;
      /** The name a screen reader says (the face is a glyph). */
      name: string;
    }
  | { kind: "text"; text: string };

export const CODE_SOFT_KEYS: readonly CodeSoftKey[] = [
  { kind: "key", key: "Escape", face: "esc", name: "Escape" },
  { kind: "key", key: "Tab", face: "⇥", name: "Tab" },
  { kind: "key", key: "Tab", shift: true, face: "⇤", name: "Shift Tab" },
  { kind: "text", text: "(" },
  { kind: "text", text: ")" },
  { kind: "text", text: "{" },
  { kind: "text", text: "}" },
  { kind: "text", text: "[" },
  { kind: "text", text: "]" },
  { kind: "text", text: ";" },
  { kind: "key", key: "ArrowLeft", face: "←", name: "Left" },
  { kind: "key", key: "ArrowUp", face: "↑", name: "Up" },
  { kind: "key", key: "ArrowDown", face: "↓", name: "Down" },
  { kind: "key", key: "ArrowRight", face: "→", name: "Right" },
];

const KEY_CODES = {
  Escape: 27,
  Tab: 9,
  ArrowLeft: 37,
  ArrowUp: 38,
  ArrowRight: 39,
  ArrowDown: 40,
} as const;

/** Types `text` as the keyboard would at the editor's selection. */
function typeText(view: EditorView, text: string): void {
  const { from, to } = view.state.selection.main;
  const insert = () =>
    view.state.update(view.state.replaceSelection(text), {
      scrollIntoView: true,
      userEvent: "input.type",
    });
  const handlers = view.state.facet(EditorView.inputHandler);
  if (handlers.some((handler) => handler(view, from, to, text, insert))) return;
  view.dispatch(insert());
}

export function pressCodeSoftKey(view: EditorView, key: CodeSoftKey): void {
  if (key.kind === "text") {
    typeText(view, key.text);
    return;
  }
  view.contentDOM.dispatchEvent(
    new KeyboardEvent("keydown", {
      key: key.key,
      code: key.key,
      keyCode: KEY_CODES[key.key],
      shiftKey: Boolean(key.shift),
      bubbles: true,
      cancelable: true,
    }),
  );
}
