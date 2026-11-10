"use client";























/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { useLayoutEffect, useRef, useState } from "react";

import {
  presentationPreviewPlainText,
  presentationPreviewSelection,
  setPresentationPreviewSelection,
} from "@mrnote/core/web/lib/presentation/text-selection.ts";

/** The letters typed in place since the host's words were last committed:
 * those words, what the host holds now, the caret after the last letter. */
interface TypedInPlace {
  readonly before: string;
  readonly after: string;
  readonly caret: number;
}

/** A plain letter by what the input says (the window holds its render for these). */
export function inputIsLetterInPlace(event: Event): boolean {
  const input = event as InputEvent;
  return (
    input.inputType === "insertText" &&
    !input.isComposing &&
    Boolean(input.data) &&
    !/[\r\n\t\v\f\u2028\u2029]/.test(input.data ?? "")
  );
}

/** Whether the browser's input is a letter typed in place in `host`. */
function typedInPlace(event: InputEvent, host: HTMLElement, before: string, words: string): TypedInPlace | null {
  if (!inputIsLetterInPlace(event) || event.defaultPrevented) return null;
  const data = event.data ?? "";
  const selection = host.ownerDocument.defaultView?.getSelection();
  if (!selection || selection.rangeCount !== 1 || !selection.isCollapsed) return null;
  const anchor = selection.anchorNode;
  if (!anchor || anchor.nodeType !== Node.TEXT_NODE || !host.contains(anchor)) return null;
  if (anchor.parentElement?.closest('[contenteditable="false"]')) return null;
  if (presentationPreviewPlainText(host) !== words) return null;
  const range = presentationPreviewSelection(host, words.length);
  if (!range || range.start !== range.end) return null;
  return {
    before,
    after: words.slice(0, range.start) + data + words.slice(range.start),
    caret: range.start + data.length,
  };
}

export function usePresentationRichTypingKey({
  blockId,
  text,
  enabled,
  host,
  pendingCaret,
}: {
  blockId: string;
  /** The words the host draws (the model's). */
  text: string;
  /** A host is drawn for typing (the box is selected and its words drawn). */
  enabled: boolean;
  host: () => HTMLElement | null;
  /** The caret this person's commit puts back, if it is putting one back;
   * else the caret after the letter typed. */
  pendingCaret?: (blockId: string) => { start: number; end: number } | null;
}): string | undefined {
  // A host whose words came apart from the model's is replaced: a new key.
  const [repairs, setRepairs] = useState(0);
  /** What the host last committed: its words, and the words its key is of. */
  const committedRef = useRef<{ text: string; keyText: string; repairs: number } | null>(null);
  const typedRef = useRef<TypedInPlace | null>(null);
  const committed = enabled ? committedRef.current : null;
  const typed = typedRef.current;
  const keptInPlace =
    committed !== null &&
    committed.repairs === repairs &&
    typed !== null &&
    typed.before === committed.text &&
    text === typed.after;
  // Words unchanged (a render for anything else) keep the key too.
  const keyText =
    committed !== null && committed.repairs === repairs && (keptInPlace || committed.text === text)
      ? committed.keyText
      : text;
  // An input method composing in the host keeps it: a collaborator's words
  // arriving meanwhile changed the key, the host was mounted anew under the
  // composition and the keyboard left it -- the member's next words went
  
  // of 8 words). The words are drawn into the same host; once the
  // composition ends the key follows the words again.
  const composingKeyRef = useRef<string | null>(null);
  const lastKeyRef = useRef<string | undefined>(undefined);
  const wordsKey = enabled ? `imported-rich:${blockId}:${repairs}:${keyText}` : undefined;
  const key = enabled && composingKeyRef.current !== null ? composingKeyRef.current : wordsKey;
  lastKeyRef.current = key;
  useLayoutEffect(() => {
    const inHost = (event: Event) => {
      const node = host();
      return node !== null && event.target instanceof Node && node.contains(event.target);
    };
    const onStart = (event: Event) => {
      if (inHost(event) && lastKeyRef.current !== undefined) composingKeyRef.current = lastKeyRef.current;
    };
    const onEnd = () => {
      composingKeyRef.current = null;
    };
    document.addEventListener("compositionstart", onStart, true);
    document.addEventListener("compositionend", onEnd, true);
    return () => {
      document.removeEventListener("compositionstart", onStart, true);
      document.removeEventListener("compositionend", onEnd, true);
    };
  }, [host]);

  useLayoutEffect(() => {
    if (!enabled) {
      committedRef.current = null;
      typedRef.current = null;
      return;
    }
    const previous = committedRef.current;
    committedRef.current = { text, keyText, repairs };
    if (previous?.text !== text) typedRef.current = null;
    if (!keptInPlace) return;
    const node = host();
    if (!node) return;
    // React wrote the model's words into the nodes the browser typed in;
    // they must read back as the model's, or the host is rebuilt now.
    if (presentationPreviewPlainText(node) !== text) {
      setRepairs((count) => count + 1);
      return;
    }
    // Writing a text node's words moves a caret inside it to its start: the
    // caret the commit leaves goes back before the frame is painted.
    if (!typed) return;
    const caret = pendingCaret?.(blockId) ?? { start: typed.caret, end: typed.caret };
    if (node.contains(node.ownerDocument.activeElement)) {
      setPresentationPreviewSelection(node, caret.start, caret.end);
    }
  });

  useLayoutEffect(() => {
    const node = enabled ? host() : null;
    if (!node) return;
    const onBeforeInput = (event: Event) => {
      const committedWords = committedRef.current?.text;
      // Letters the window has not drawn yet are in the host: the chain goes on
      // from them, and one commit takes them all.
      const held = typedRef.current?.before === committedWords ? typedRef.current : null;
      typedRef.current =
        committedWords === undefined
          ? null
          : typedInPlace(event as InputEvent, node, committedWords, held?.after ?? committedWords);
    };
    node.addEventListener("beforeinput", onBeforeInput);
    return () => node.removeEventListener("beforeinput", onBeforeInput);
  }, [enabled, host, key]);

  return key;
}
