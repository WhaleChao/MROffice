"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * One soft session on the engine host (ENGINE-HOST-CONTRACT 12.2), kept open
 * for the browser and re-opened when the server lost it.
 *
 * A session is server state built from a snapshot the browser owns (a
 * workbook's calculation input). It is SOFT: idle expiry, the server's
 * memory caps and quotas, a recycled worker or a request that lands on
 * another API process all answer `session_unknown`, and the client re-opens
 * with its CURRENT snapshot and retries once. Correctness never depends on a
 * session surviving; only speed does.
 *
 * The transport lives in `./client` (mocked by the unit tests' in-process
 * host); this class only sequences it.
 */

import {
  EngineSessionUnknownError,
  callEngineSession,
  cancelEngineSession,
  closeEngineSession,
  openEngineSession,
  type EngineSessionInput,
  type EngineSessionMethodName,
  type EngineSessionName,
  type EngineSessionOpenInput,
  type EngineSessionOpenOutput,
  type EngineSessionOutput,
} from "./client";

export interface EngineSessionOptions<E extends EngineSessionName> {
  /** The state to open (or re-open) the session from, read at that moment. */
  snapshot(): EngineSessionOpenInput<E>;
  /** What the server's per-scope quota counts (the page id). */
  scope?: string;
  /** The scope read off the state an open sends, when `scope` is not given. */
  scopeOf?(input: EngineSessionOpenInput<E>): string | undefined;
  /** Every successful open (the first and each re-open). */
  onOpened?(opened: EngineSessionOpenOutput<E>, reopened: boolean): void;
}

export interface EngineSessionCallOptions<I> {
  signal?: AbortSignal;
  /**
   * The input to retry with after a re-open. The snapshot a re-open sends is
   * current, so a call that carried a change the snapshot already holds
   * retries without it. Default: the same input.
   */
  retryInput?(): I;
}








const LAST_SESSION_KEY = "mrai:engine-session:";

function rememberedSession(engine: string, scope: string): string | null {
  try {
    return typeof sessionStorage === "undefined" ? null : sessionStorage.getItem(`${LAST_SESSION_KEY}${engine}:${scope}`);
  } catch {
    return null;
  }
}

function rememberSession(engine: string, scope: string, id: string | null): void {
  try {
    if (typeof sessionStorage === "undefined") return;
    const key = `${LAST_SESSION_KEY}${engine}:${scope}`;
    if (id) sessionStorage.setItem(key, id);
    else sessionStorage.removeItem(key);
  } catch {
    /* a tab without storage keeps only the idle expiry */
  }
}

function abortError(): DOMException {
  return new DOMException("The operation was aborted", "AbortError");
}

export class EngineSession<E extends EngineSessionName> {
  readonly engine: E;
  readonly #options: EngineSessionOptions<E>;
  #id: string | null = null;
  #opening: Promise<string> | null = null;
  #generation = 0;
  #closed = false;
  /** The scope the open session counts toward (and this tab remembers it by). */
  #scope: string | undefined;
  /** Opens after the first, because the server had lost the session. */
  reopens = 0;

  constructor(engine: E, options: EngineSessionOptions<E>) {
    this.engine = engine;
    this.#options = options;
  }

  /** The open session's id (null before the first open or after a loss). */
  get id(): string | null {
    return this.#id;
  }

  get closed(): boolean {
    return this.#closed;
  }

  /** Forget the server's state; the next call opens from a fresh snapshot
   * (`keepalive` from `pagehide`: a page restored from the back/forward
   * cache simply opens again). */
  reset(options: { keepalive?: boolean } = {}): void {
    const id = this.#id;
    this.#id = null;
    this.#opening = null;
    this.#generation += 1;
    if (id) {
      this.#forgetRemembered(id);
      void closeEngineSession(this.engine, id, options);
    }
  }

  #forgetRemembered(id: string): void {
    if (this.#scope && rememberedSession(this.engine, this.#scope) === id) rememberSession(this.engine, this.#scope, null);
  }

  /** Opens the session now unless it is open (the open's snapshot is read
   * now); resolves its id. */
  open(signal?: AbortSignal): Promise<string> {
    return this.#ensure(signal, false);
  }

  /** Close for good (`keepalive` from `pagehide`). */
  close(options: { keepalive?: boolean } = {}): void {
    const id = this.#id;
    this.#closed = true;
    this.#id = null;
    this.#opening = null;
    this.#generation += 1;
    if (id) {
      this.#forgetRemembered(id);
      void closeEngineSession(this.engine, id, options);
    }
  }

  /** Stop the session's running call at its next yield; the state stays. */
  cancel(): void {
    if (this.#id) void cancelEngineSession(this.engine, this.#id);
  }

  async #ensure(signal: AbortSignal | undefined, reopened: boolean): Promise<string> {
    for (;;) {
      if (this.#closed) throw abortError();
      if (signal?.aborted) throw abortError();
      if (this.#id) return this.#id;
      const generation = this.#generation;
      if (!this.#opening) {
        // Shared by every caller waiting for this generation's open; one
        // caller's abort does not abort the open the others wait for.
        const input = this.#options.snapshot();
        const scope = this.#options.scope ?? this.#options.scopeOf?.(input);
        // The one this tab opened before for the same scope goes first.
        const stale = scope ? rememberedSession(this.engine, scope) : null;
        if (scope && stale) {
          rememberSession(this.engine, scope, null);
          void closeEngineSession(this.engine, stale);
        }
        const opening = openEngineSession(this.engine, input, { scope }).then(
          ({ sessionId, opened }) => {
            if (generation !== this.#generation) {
              // Reset while opening: that state is not the current snapshot.
              void closeEngineSession(this.engine, sessionId);
              return "";
            }
            this.#id = sessionId;
            this.#scope = scope;
            if (scope) rememberSession(this.engine, scope, sessionId);
            this.#opening = null;
            this.#options.onOpened?.(opened, reopened || this.reopens > 0);
            return sessionId;
          },
          (error: unknown) => {
            if (this.#opening === opening) this.#opening = null;
            throw error;
          },
        );
        this.#opening = opening;
      }
      const id = await this.#abortable(this.#opening, signal);
      if (id) return id;
    }
  }

  #abortable<T>(promise: Promise<T>, signal: AbortSignal | undefined): Promise<T> {
    if (!signal) return promise;
    return new Promise<T>((resolve, reject) => {
      const onAbort = () => reject(abortError());
      if (signal.aborted) return onAbort();
      signal.addEventListener("abort", onAbort, { once: true });
      promise.then(
        (value) => { signal.removeEventListener("abort", onAbort); resolve(value); },
        (error) => { signal.removeEventListener("abort", onAbort); reject(error); },
      );
    });
  }

  /**
   * Runs `method` against the session, opening it first when needed. A
   * `session_unknown` answer re-opens with the current snapshot and retries
   * once; a second loss rejects with {@link EngineSessionUnknownError}.
   */
  async call<M extends EngineSessionMethodName<E>>(
    method: M,
    input: EngineSessionInput<E, M>,
    options: EngineSessionCallOptions<EngineSessionInput<E, M>> = {},
  ): Promise<EngineSessionOutput<E, M>> {
    const id = await this.#ensure(options.signal, false);
    try {
      return await callEngineSession(this.engine, id, method, input, { signal: options.signal });
    } catch (error) {
      if (!(error instanceof EngineSessionUnknownError)) throw error;
      if (this.#id === id) {
        this.#id = null;
        this.#generation += 1;
        this.#opening = null;
      }
      this.reopens += 1;
      const again = await this.#ensure(options.signal, true);
      const retry = options.retryInput ? options.retryInput() : input;
      return await callEngineSession(this.engine, again, method, retry, { signal: options.signal });
    }
  }
}
