/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * `JSON.stringify` of a large plain-JSON value in slices of main-thread time.
 *
 * Keeping a copy of a long document (lib/saved-page-hot-cache) stringified
 * its whole page in one task: 8.5 MB of body and checkpoint, hundreds of
 * milliseconds during which nothing else on the site could run. The text
 * here is exactly `JSON.stringify(value)`; large arrays and objects are
 * written member by member, yielding to the main thread whenever a slice has
 * run for `sliceMs`.
 */

import { createBackgroundLane, type BackgroundLane } from "./background-work";

/** Arrays with fewer members than this are stringified whole. */
const SLICE_MIN_MEMBERS = 32;
/** Containers this deep or deeper are stringified whole. */
const SLICE_MAX_DEPTH = 4;

export function yieldToMainThread(): Promise<void> {
  const scheduler = (globalThis as { scheduler?: { yield?: () => Promise<void> } }).scheduler;
  if (typeof scheduler?.yield === "function") return scheduler.yield();
  return new Promise(resolve => setTimeout(resolve, 0));
}

export async function stringifyJsonInSlices(
  value: unknown,
  options: { sliceMs?: number; yieldTask?: () => Promise<void> } = {},
): Promise<string> {
  return inSlices(value, options, async (parts) => parts.join(""));
}

/** Text parts encoded together before the next slice check. */
const ENCODE_BATCH_CHARS = 64 * 1024;

/**
 * Exactly `new TextEncoder().encode(JSON.stringify(value))`, in slices: the
 * encoding of a long document's 10 MB text was a task of its own of 150 ms.
 * Every part is whole JSON text (never half of a surrogate pair), so the
 * parts' bytes, end to end, are the bytes of the whole.
 */
export async function encodeJsonInSlices(
  value: unknown,
  options: { sliceMs?: number; yieldTask?: () => Promise<void> } = {},
): Promise<Uint8Array<ArrayBuffer>> {
  return inSlices(value, options, async (parts, maybeYield) => {
    const encoder = new TextEncoder();
    const chunks: Array<Uint8Array<ArrayBuffer>> = [];
    let total = 0;
    let batch: string[] = [];
    let batchChars = 0;
    const flush = () => {
      if (batch.length === 0) return;
      const bytes = encoder.encode(batch.join(""));
      chunks.push(bytes);
      total += bytes.byteLength;
      batch = [];
      batchChars = 0;
    };
    for (const part of parts) {
      batch.push(part);
      batchChars += part.length;
      if (batchChars < ENCODE_BATCH_CHARS) continue;
      flush();
      await maybeYield();
    }
    flush();
    if (chunks.length === 1) return chunks[0];
    const out = new Uint8Array(total);
    let offset = 0;
    for (const chunk of chunks) {
      out.set(chunk, offset);
      offset += chunk.byteLength;
    }
    return out;
  });
}

async function inSlices<T>(
  value: unknown,
  options: { sliceMs?: number; yieldTask?: () => Promise<void> },
  finish: (parts: string[], maybeYield: () => Promise<void>) => Promise<T>,
): Promise<T> {
  const sliceMs = options.sliceMs ?? 8;
  // Without a caller's own yield, the slices are background work
  // (lib/background-work): they give way to the person's input.
  const lane: BackgroundLane | null = options.yieldTask ? null : createBackgroundLane({ label: "json-slices" });
  const yieldTask = options.yieldTask ?? (() => lane!.turn());
  let sliceStart = performance.now();
  const parts: string[] = [];
  const maybeYield = async () => {
    if (lane ? !lane.shouldYield() : performance.now() - sliceStart < sliceMs) return;
    await yieldTask();
    sliceStart = performance.now();
  };
  // The page, its body, the body's blocks: below that a member is written
  // whole (one block of a document is small).
  const write = async (item: unknown, depth: number): Promise<void> => {
    if (depth < SLICE_MAX_DEPTH && item !== null && typeof item === "object" &&
        typeof (item as { toJSON?: unknown }).toJSON !== "function") {
      if (Array.isArray(item)) {
        if (item.length >= SLICE_MIN_MEMBERS) {
          parts.push("[");
          for (let index = 0; index < item.length; index += 1) {
            if (index > 0) parts.push(",");
            const member = item[index];
            // JSON.stringify writes null for an array's undefined/function.
            if (member === undefined || typeof member === "function" || typeof member === "symbol") parts.push("null");
            else await write(member, depth + 1);
            await maybeYield();
          }
          parts.push("]");
          return;
        }
      } else {
        parts.push("{");
        let first = true;
        for (const key of Object.keys(item)) {
          const member = (item as Record<string, unknown>)[key];
          if (member === undefined || typeof member === "function" || typeof member === "symbol") continue;
          if (!first) parts.push(",");
          first = false;
          parts.push(JSON.stringify(key), ":");
          await write(member, depth + 1);
          await maybeYield();
        }
        parts.push("}");
        return;
      }
    }
    parts.push(JSON.stringify(item) ?? "null");
  };
  try {
    if (lane) await lane.turn();
    await write(value, 0);
    return await finish(parts, maybeYield);
  } finally {
    lane?.dispose();
  }
}
