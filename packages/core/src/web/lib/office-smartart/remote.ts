"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * SmartArt layout, asked of the server.
 *
 * The layout engine and every relayout it drives (insert, change layout,
 * colours, quick style, Reset, the outline commands, the conversion of an
 * Office diagram) run in the engine host (`engine-host/engines/smartart.ts`);
 * the browser keeps the members a page stores and draws them. This is the one
 * door every editor uses, on top of the shared engine client, adding the two
 * things a SmartArt answer needs:
 *
 * - new ids: the engine answers placeholders and each distinct one becomes a
 *   fresh id here, so a cached answer still gives every insert its own ids;
 * - picture slots: a member's picture is an inline image the engine only
 *   copies, so it travels as a short stand-in that still reads as an image
 *   and is put back in the answer, keeping requests small.
 */

import { callEngine, type EngineInput, type EngineMethodName, type EngineOutput } from "../engine-host/client";
import { EngineRefusedError } from "../engine-host/client";
import { SMARTART_ENGINE_ID_PREFIX } from "../engine-host/smartart-types";
import { smartArtDefinitionReference, smartArtLayoutDefinition } from "./layout-definition-client";

export {
  EngineRefusedError,
  EngineUnavailableError,
  engineErrorMessageKey,
  isEngineAbort,
} from "../engine-host/client";

type SmartArtMethod = EngineMethodName<"smartart">;

/** Pictures shorter than this ride as they are. */
const STAND_IN_MIN_LENGTH = 4096;
const PICTURE = /^data:image\/(?:png|jpeg);base64,/;

function nonce(): string {
  const bytes = new Uint8Array(9);
  crypto.getRandomValues(bytes);
  return Array.from(bytes, (byte) => "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789"[byte % 62]).join("");
}

function mapStrings(value: unknown, map: (text: string) => string): unknown {
  if (typeof value === "string") return map(value);
  if (Array.isArray(value)) {
    let changed = false;
    const next = value.map((item) => {
      const mapped = mapStrings(item, map);
      if (mapped !== item) changed = true;
      return mapped;
    });
    return changed ? next : value;
  }
  if (value && typeof value === "object") {
    let changed = false;
    const next: Record<string, unknown> = {};
    for (const [key, item] of Object.entries(value as Record<string, unknown>)) {
      const mapped = mapStrings(item, map);
      if (mapped !== item) changed = true;
      next[key] = mapped;
    }
    return changed ? next : value;
  }
  return value;
}

/** The request with every large picture replaced by a stand-in, and the way back. */
export function withPictureStandIns<T>(input: T): { input: T; restore: (text: string) => string } {
  const originals = new Map<string, string>();
  const byOriginal = new Map<string, string>();
  const tag = nonce();
  const replaced = mapStrings(input, (text) => {
    if (text.length < STAND_IN_MIN_LENGTH) return text;
    const header = PICTURE.exec(text)?.[0];
    if (!header) return text;
    let token = byOriginal.get(text);
    if (!token) {
      token = `${header}TVJBSXNtYXJ0YXJ0${tag}${originals.size}`;
      originals.set(token, text);
      byOriginal.set(text, token);
    }
    return token;
  }) as T;
  return { input: replaced, restore: (text) => originals.get(text) ?? text };
}

function freshId(): string {
  return typeof crypto !== "undefined" && "randomUUID" in crypto
    ? crypto.randomUUID()
    : `smartart-${Date.now().toString(36)}-${nonce()}`;
}

/** The answer with stand-ins put back and every placeholder id made real. */
export function withRealIds<T>(answer: T, restore: (text: string) => string = (text) => text): T {
  const ids = new Map<string, string>();
  return mapStrings(answer, (text) => {
    if (text.startsWith(SMARTART_ENGINE_ID_PREFIX)) {
      let id = ids.get(text);
      if (!id) {
        id = freshId();
        ids.set(text, id);
      }
      return id;
    }
    return restore(text);
  }) as T;
}

/**
 * Run one SmartArt engine method. Rejects with the engine client's errors
 * (`EngineUnavailableError` -- no connection, busy, failed; or
 * `EngineRefusedError`) or the caller's `AbortError`; the caller leaves the
 * diagram as it is and says why.
 */
export async function smartArtEngine<M extends SmartArtMethod>(
  method: M,
  input: EngineInput<"smartart", M>,
  options: { signal?: AbortSignal } = {},
): Promise<EngineOutput<"smartart", M>> {
  const { input: sent, restore } = withPictureStandIns(await withLentDefinition(input));
  const answer = await callEngine("smartart", method, sent, { signal: options.signal, cache: true });
  return withStatedDirection(withRealIds(answer, restore));
}

/**
 * A request about a diagram laid out from its own file's definition carries
 * that part (`layoutDefinition`, SMARTART-GAP.md §8.6), from the open page's
 * package; one nobody can supply is refused here, before the engine would
 * lay the diagram out as nothing.
 */
async function withLentDefinition<T>(input: T): Promise<T> {
  const reference = smartArtDefinitionReference(input);
  if (!reference || (input as { layoutDefinition?: unknown }).layoutDefinition) return input;
  const definition = await smartArtLayoutDefinition(reference.sha256);
  if (!definition || definition.urn !== reference.urn) {
    throw new EngineRefusedError("smartart_layout_definition_unavailable", { urn: reference.urn });
  }
  const { layoutDefinitionRef: _ref, ...rest } = input as Record<string, unknown>;
  return { ...rest, layoutDefinition: { urn: definition.urn, sha256: definition.sha256, xml: definition.xml } } as T;
}

/**
 * A Docs patch always STATES the diagram's direction, because the host merges
 * a patch onto the node's attrs: left to right is `smartArtDirection:
 * undefined`, which JSON cannot carry. Put it back, or a diagram turned back
 * to left to right would keep the `rtl` it no longer draws.
 */
function withStatedDirection<T>(answer: T): T {
  const patch = (answer as { patch?: unknown } | null)?.patch;
  if (!patch || typeof patch !== "object") return answer;
  // …and the Office alias it is drawn as (`smartArtLayoutUrn`) and its colour
  // variation (`smartArtColors`), likewise: a layout switch must not keep the
  // alias it left behind, nor a Reset Graphic the variation it cleared.
  const stated = ["smartArtDirection", "smartArtLayoutUrn", "smartArtColors"].filter((key) => !(key in patch));
  if (stated.length === 0) return answer;
  return {
    ...answer,
    patch: { ...patch, ...Object.fromEntries(stated.map((key) => [key, undefined])) },
  } as T;
}
