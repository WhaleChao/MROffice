/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Before every commit, React DOM saves the selection of the focused element
 * when that element "has selection capabilities" (an input, a textarea, or an
 * element whose `contentEditable` reads "true"), to restore it should the
 * commit move the focus. For a contenteditable it counts the characters of
 * EVERY text node of the element, from its first to its last, whatever the
 * caret's place (react-dom getSelectionInformation / getModernOffsetsFromPoints,
 * which only stops at the element's end). A Docs editor is one contenteditable
 * holding the whole document: in the 296-page book (~100k elements) each
 * commit of any React root on the page walked all of it, and 30 keys typed
 * mid-book made 54 commits (the toolbar, the status bar, the editor shell):
 * 1.1 s of the 30 keys in a dev build, ~20 ms a commit.
 *
 * ProseMirror owns that element's selection and restores it itself on focus;
 * React never renders inside it. Its own `contentEditable` therefore reads
 * "inherit" to JavaScript (the attribute, which is what the browser edits by,
 * is untouched, as is `isContentEditable`), and React leaves it out. Writes
 * still reach the element. Nothing in this app or in ProseMirror/Tiptap reads
 * that property of the editor root (they read the attribute, `isContentEditable`,
 * or `contentEditable == "false"` of nodes inside it).
 */
import { Extension } from "@tiptap/core";
import { Plugin, PluginKey } from "@tiptap/pm/state";

const nativeContentEditable =
  typeof HTMLElement === "undefined" ? undefined : Object.getOwnPropertyDescriptor(HTMLElement.prototype, "contentEditable");

/** Make React's pre-commit selection save pass over `element`. Returns the undo. */
export function hideEditableRootFromReactSelectionSave(element: HTMLElement): () => void {
  const native = nativeContentEditable;
  if (!native?.get || !native.set || Object.getOwnPropertyDescriptor(element, "contentEditable")) return () => {};
  Object.defineProperty(element, "contentEditable", {
    configurable: true,
    enumerable: native.enumerable,
    get(this: HTMLElement) {
      const value = native.get!.call(this) as string;
      return value === "true" ? "inherit" : value;
    },
    set(this: HTMLElement, value: string) {
      native.set!.call(this, value);
    },
  });
  return () => {
    delete (element as { contentEditable?: string }).contentEditable;
  };
}

export const ReactCommitSelectionGuard = Extension.create({
  name: "reactCommitSelectionGuard",
  addProseMirrorPlugins() {
    return [
      new Plugin({
        key: new PluginKey("reactCommitSelectionGuard"),
        view: (view) => {
          const release = hideEditableRootFromReactSelectionSave(view.dom as HTMLElement);
          return { destroy: release };
        },
      }),
    ];
  },
});
