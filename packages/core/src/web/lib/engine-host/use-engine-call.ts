"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Latest-wins engine calls for a control, with a pending state.
 *
 * An editor command that asks the server (a SmartArt relayout) must not apply
 * an answer to a document that moved on: a newer command supersedes an older
 * one, and a closed window aborts what it asked. {@link EngineCallSequence}
 * is that rule outside React (an editor's command callbacks hold one in a
 * ref); {@link useEngineCall} is the same with a `pending` flag a control
 * renders (shown after a short delay, so a fast answer never flickers).
 * Callers still check their own authority (same selection, same content)
 * before applying a defined result.
 */

import { useCallback, useEffect, useRef, useState, useSyncExternalStore } from "react";
import {
  callEngine,
  isEngineAbort,
  type CallEngineOptions,
  type EngineInput,
  type EngineMethodName,
  type EngineName,
  type EngineOutput,
} from "./client";

/** A pending state shows only after this long, so a fast answer never flickers. */
export const ENGINE_PENDING_DELAY_MS = 150;

export class EngineCallSequence {
  private latest = 0;
  private controller: AbortController | null = null;
  private pendingTimer: ReturnType<typeof setTimeout> | null = null;
  private pendingValue = false;
  private readonly listeners = new Set<() => void>();

  /** True while the latest call has been running longer than the delay. */
  get pending(): boolean {
    return this.pendingValue;
  }

  subscribe = (listener: () => void): (() => void) => {
    this.listeners.add(listener);
    return () => this.listeners.delete(listener);
  };

  private setPending(value: boolean): void {
    if (this.pendingValue === value) return;
    this.pendingValue = value;
    for (const listener of this.listeners) listener();
  }

  /**
   * Run `call` as the latest: resolves its answer, or `undefined` when a
   * newer run started or {@link abort} was called before it answered.
   * Rejects with the engine's error otherwise.
   */
  async run<T>(call: (signal: AbortSignal) => Promise<T>): Promise<T | undefined> {
    this.controller?.abort();
    const controller = new AbortController();
    const token = ++this.latest;
    this.controller = controller;
    if (this.pendingTimer) clearTimeout(this.pendingTimer);
    this.pendingTimer = setTimeout(() => {
      if (this.latest === token) this.setPending(true);
    }, ENGINE_PENDING_DELAY_MS);
    // A superseded or aborted call answers `undefined` at once, whether or not
    // the work it started honours the signal.
    const dropped = new Promise<undefined>((resolve) =>
      controller.signal.addEventListener("abort", () => resolve(undefined), { once: true }),
    );
    try {
      const value = await Promise.race([call(controller.signal), dropped]);
      return this.latest === token && !controller.signal.aborted ? value : undefined;
    } catch (error) {
      if (isEngineAbort(error) || this.latest !== token || controller.signal.aborted) return undefined;
      throw error;
    } finally {
      if (this.latest === token) {
        if (this.pendingTimer) clearTimeout(this.pendingTimer);
        this.pendingTimer = null;
        this.controller = null;
        this.setPending(false);
      }
    }
  }

  /** Drop whatever is in flight; its answer will never be applied. */
  abort(): void {
    this.latest += 1;
    this.controller?.abort();
    this.controller = null;
    if (this.pendingTimer) clearTimeout(this.pendingTimer);
    this.pendingTimer = null;
    this.setPending(false);
  }
}














export class EngineCallQueue {
  private tail: Promise<unknown> = Promise.resolve();
  private controller = new AbortController();
  private depth = 0;
  private readonly listeners = new Set<() => void>();

  /** True while a step is running or waiting. */
  get busy(): boolean {
    return this.depth > 0;
  }

  subscribe = (listener: () => void): (() => void) => {
    this.listeners.add(listener);
    return () => this.listeners.delete(listener);
  };

  private notify(): void {
    for (const listener of this.listeners) listener();
  }

  /**
   * Queue `step`: it runs after every step queued before it has settled.
   * Resolves its result, or `undefined` when the queue was aborted first;
   * rejects with the step's own error (an engine abort resolves `undefined`).
   */
  run<T>(step: (signal: AbortSignal) => Promise<T>): Promise<T | undefined> {
    const signal = this.controller.signal;
    this.depth += 1;
    if (this.depth === 1) this.notify();
    const result = this.tail.then(() => {
      if (signal.aborted) return undefined;
      // An aborted step answers `undefined` at once, whether or not the work
      // it started honours the signal, so the steps after it are not held.
      const dropped = new Promise<undefined>((resolve) =>
        signal.addEventListener("abort", () => resolve(undefined), { once: true }),
      );
      return Promise.race([step(signal), dropped]);
    });
    this.tail = result
      .catch(() => undefined)
      .finally(() => {
        this.depth -= 1;
        if (this.depth === 0) this.notify();
      });
    return result.catch((error: unknown) => {
      if (isEngineAbort(error) || signal.aborted) return undefined;
      throw error;
    });
  }

  /** Abort what is running and drop what waits. */
  abort(): void {
    this.controller.abort();
    this.controller = new AbortController();
  }
}

/** One {@link EngineCallQueue} for a component, aborted on unmount. */
export function useEngineCallQueue(): EngineCallQueue {
  const ref = useRef<EngineCallQueue | null>(null);
  if (!ref.current) ref.current = new EngineCallQueue();
  const queue = ref.current;
  useEffect(() => () => queue.abort(), [queue]);
  return queue;
}

/** Whether `queue` has a step running or waiting. */
export function useEngineCallQueueBusy(queue: EngineCallQueue): boolean {
  return useSyncExternalStore(queue.subscribe, () => queue.busy, () => false);
}

/** One {@link EngineCallSequence} for a component, aborted on unmount. */
export function useEngineCallSequence(): EngineCallSequence {
  const ref = useRef<EngineCallSequence | null>(null);
  if (!ref.current) ref.current = new EngineCallSequence();
  const sequence = ref.current;
  useEffect(() => () => sequence.abort(), [sequence]);
  return sequence;
}

export function useEngineCallPending(sequence: EngineCallSequence): boolean {
  return useSyncExternalStore(sequence.subscribe, () => sequence.pending, () => false);
}

export interface EngineCallHandle<E extends EngineName, M extends EngineMethodName<E>> {
  run(input: EngineInput<E, M>, options?: Omit<CallEngineOptions, "signal">): Promise<EngineOutput<E, M> | undefined>;
  pending: boolean;
  error: unknown;
  abort(): void;
}

export function useEngineCall<E extends EngineName, M extends EngineMethodName<E>>(
  engine: E,
  method: M,
): EngineCallHandle<E, M> {
  const sequence = useEngineCallSequence();
  const pending = useEngineCallPending(sequence);
  const [error, setError] = useState<unknown>(null);
  const run = useCallback(
    async (input: EngineInput<E, M>, options?: Omit<CallEngineOptions, "signal">) => {
      setError(null);
      try {
        return await sequence.run((signal) => callEngine(engine, method, input, { ...options, signal }));
      } catch (caught) {
        setError(caught);
        throw caught;
      }
    },
    [engine, method, sequence],
  );
  const abort = useCallback(() => sequence.abort(), [sequence]);
  return { run, pending, error, abort };
}
