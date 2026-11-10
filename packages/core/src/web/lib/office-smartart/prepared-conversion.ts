/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Opening a file writes nothing: a preserved Office SmartArt is converted with
 * its first real edit, not as the file opens.
 *
 * The file's own diagram stays the file's until the author changes it. When one
 * is selected its conversion is asked of the engine and kept beside it, written
 * nowhere; the first command, gesture or typed word that really changes it lands
 * the conversion together with the change, as one Undo step. This is what the
 * three hosts share: the state of that request, what the author sees and may
 * press while it is on its way, and the answer when it can never be given.
 *
 * How the conversion is held until then is each host's own: the deck and the
 * worksheet keep it beside their model (`presentation-smartart-prepared`,
 * `spreadsheetSmartArtPrepared`), and Word beside the document's node
 * (`documentSmartArtPrepared`), whose document is shared the moment it changes.
 */

import { useCallback, useEffect, useRef, useSyncExternalStore } from "react";

import type { SmartArtRefusal } from "./convert-refusal-message";

/** A diagram the engine could not answer for is asked again after this long. */
export const SMARTART_PREPARE_RETRY_MS = 3000;

export type SmartArtPreparation<Value> =
  | { readonly state: "preparing" }
  | { readonly state: "ready"; readonly value: Value }
  /** It can never convert as it is: it stays as Office drew it, and says why. */
  | { readonly state: "refused"; readonly refusal: SmartArtRefusal }
  /** No answer (offline, the engine busy): not the diagram's fault, ask again. */
  | { readonly state: "failed"; readonly error: unknown; readonly at: number };

/** What asking for one diagram's conversion comes back with. */
export type SmartArtPreparationAnswer<Value> =
  | { readonly value: Value }
  | { readonly refusal: SmartArtRefusal };

const listeners = new Set<() => void>();
let version = 0;

function announce(): void {
  version += 1;
  for (const listener of listeners) listener();
}

export function subscribeSmartArtPreparation(listener: () => void): () => void {
  listeners.add(listener);
  return () => {
    listeners.delete(listener);
  };
}

/** Changes whenever a diagram's preparation changes, so a host reads it again. */
export function useSmartArtPreparationVersion(): number {
  return useSyncExternalStore(subscribeSmartArtPreparation, () => version, () => 0);
}

/** The preparations of one host's diagrams, by the object that holds each one. */
export interface SmartArtPreparations<Key extends object, Value> {
  read(key: Key): SmartArtPreparation<Value> | undefined;
  /**
   * Asks for `key`'s conversion unless it has an answer already: an answer is
   * kept for as long as the object lives, and an edit makes a new object. A
   * failure is asked again only after {@link SMARTART_PREPARE_RETRY_MS}.
   */
  prepare(key: Key, ask: () => Promise<SmartArtPreparationAnswer<Value>>): SmartArtPreparation<Value>;
  /** Asks again now, for a press that must have its answer. */
  retry(key: Key, ask: () => Promise<SmartArtPreparationAnswer<Value>>): SmartArtPreparation<Value>;
  /** The answer for `key` is known without asking. */
  seed(key: Key, value: Value): void;
}

export function createSmartArtPreparations<Key extends object, Value>(): SmartArtPreparations<Key, Value> {
  const entries = new WeakMap<Key, SmartArtPreparation<Value>>();
  const start = (key: Key, ask: () => Promise<SmartArtPreparationAnswer<Value>>) => {
    const preparing: SmartArtPreparation<Value> = { state: "preparing" };
    entries.set(key, preparing);
    void Promise.resolve()
      .then(ask)
      .then(
        (answer): SmartArtPreparation<Value> =>
          "refusal" in answer
            ? { state: "refused", refusal: answer.refusal }
            : { state: "ready", value: answer.value },
        (error: unknown): SmartArtPreparation<Value> => ({ state: "failed", error, at: Date.now() }),
      )
      .then((settled) => {
        entries.set(key, settled);
        announce();
      });
    // Read by whoever asked; a press that asks again is seen to wait.
    queueMicrotask(announce);
    return preparing;
  };
  return {
    read: (key) => entries.get(key),
    prepare: (key, ask) => {
      const entry = entries.get(key);
      if (entry && (entry.state !== "failed" || Date.now() - entry.at < SMARTART_PREPARE_RETRY_MS)) return entry;
      return start(key, ask);
    },
    retry: (key, ask) => {
      const entry = entries.get(key);
      if (entry && entry.state !== "failed") return entry;
      return start(key, ask);
    },
    seed: (key, value) => {
      entries.set(key, { state: "ready", value });
      queueMicrotask(announce);
    },
  };
}

/** The commands of the SmartArt tab's diagram that can wait for the conversion. */
export interface SmartArtDiagramCommands {
  onToggleTextPane: () => void;
  onPromote: () => void;
  onDemote: () => void;
  onRightToLeft: () => void;
  onChangeLayout: (...args: never[]) => void;
  onChangeColors: (...args: never[]) => void;
  onChangeStyle: (...args: never[]) => void;
  onReset: () => void;
  onAltText: (trigger: HTMLElement | null) => void;
}

/**
 * The tab's whole diagram while its conversion is on its way, its state read
 * from the preserved view (`state`): every press waits for the answer and then
 * runs on the real diagram. Nothing that needs an item (Add Shape, Promote) is
 * offered yet: which item it is, is the conversion's to say.
 */
export function smartArtDiagramWhilePreparing<Diagram extends SmartArtDiagramCommands, State extends object>(
  state: State,
  whenReady: (press: (diagram: Diagram) => void) => void,
): State & Pick<Diagram, keyof SmartArtDiagramCommands> {
  const later =
    <Args extends unknown[]>(run: (diagram: Diagram, ...args: Args) => void) =>
    (...args: Args) => {
      whenReady((diagram) => run(diagram, ...args));
    };
  return {
    ...state,
    onToggleTextPane: later((diagram) => diagram.onToggleTextPane()),
    onPromote: later((diagram) => diagram.onPromote()),
    onDemote: later((diagram) => diagram.onDemote()),
    onRightToLeft: later((diagram) => diagram.onRightToLeft()),
    onChangeLayout: later((diagram, ...args: unknown[]) =>
      (diagram.onChangeLayout as (...values: unknown[]) => void)(...args),
    ),
    onChangeColors: later((diagram, ...args: unknown[]) =>
      (diagram.onChangeColors as (...values: unknown[]) => void)(...args),
    ),
    onChangeStyle: later((diagram, ...args: unknown[]) =>
      (diagram.onChangeStyle as (...values: unknown[]) => void)(...args),
    ),
    onReset: later((diagram) => diagram.onReset()),
    onAltText: later((diagram, trigger: HTMLElement | null) => diagram.onAltText(trigger)),
  } as unknown as State & Pick<Diagram, keyof SmartArtDiagramCommands>;
}

/**
 * A press made while the conversion was on its way runs when the diagram is
 * ready, one at a time, each on the diagram the one before it left; if the answer
 * is no, the author is told, and the press is not lost in silence.
 * `scope` is the diagram in hand: another one drops what was waiting.
 */
export function useSmartArtPressesWhilePreparing<Diagram>(
  input: () => {
    readonly scope: unknown;
    /** The diagram's real tab state, once it has one. */
    readonly diagram: Diagram | null;
    readonly preparation: SmartArtPreparation<unknown> | null | undefined;
    readonly report: (preparation: Extract<SmartArtPreparation<unknown>, { state: "refused" | "failed" }>) => void;
  },
  /** The author pressed while it was on its way: say it is being prepared. */
  waiting?: () => void,
): (press: (diagram: Diagram) => void) => void {
  const queue = useRef<Array<(diagram: Diagram) => void>>([]);
  const inputRef = useRef(input);
  inputRef.current = input;
  const waitingRef = useRef(waiting);
  waitingRef.current = waiting;
  const scopeRef = useRef<unknown>(null);
  const reportedRef = useRef<unknown>(null);
  // After every render, reading what that render made: the host computes its tab
  // late in the render, past where a hook may be called.
  useEffect(() => {
    const { scope, diagram, preparation, report } = inputRef.current();
    if (scope !== scopeRef.current) {
      scopeRef.current = scope;
      queue.current = [];
      return;
    }
    if (!queue.current.length) return;
    if (diagram) {
      queue.current.shift()?.(diagram);
      return;
    }
    if (
      preparation &&
      (preparation.state === "refused" || preparation.state === "failed") &&
      reportedRef.current !== preparation
    ) {
      reportedRef.current = preparation;
      queue.current = [];
      report(preparation);
    }
  });
  return useCallback((press: (diagram: Diagram) => void) => {
    if (queue.current.length === 0) waitingRef.current?.();
    queue.current.push(press);
  }, []);
}

/**
 * A press on a prepared diagram's conversion (`selector`, in a window that lays it
 * unseen over the file's drawing): it is being worked on, so it is drawn as what it
 * will be. Nothing is saved until it changes, and leaving it puts the drawing back.
 */
export function useSmartArtRevealOnPress(active: boolean, selector: string, reveal: () => void): void {
  const revealRef = useRef(reveal);
  revealRef.current = reveal;
  useEffect(() => {
    if (!active) return;
    const onPress = (event: Event) => {
      if (event.target instanceof Element && event.target.closest(selector)) revealRef.current();
    };
    document.addEventListener("pointerdown", onPress, true);
    return () => document.removeEventListener("pointerdown", onPress, true);
  }, [active, selector]);
}
