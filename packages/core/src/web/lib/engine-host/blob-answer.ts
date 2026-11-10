/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Binary answers (ENGINE-HOST-CONTRACT 12.1, answer direction): a method
 * whose policy opts in (`binary_answer`) answers whole files after its JSON,
 * in the same framing an upload uses -- `[u32 BE JSON length][JSON][bytes]`
 * -- where the JSON is `{ result, ms, blobs: [{ sha256, size }] }` and the
 * bytes follow in that order. In `result` each file is `{"$blob": sha256}`;
 * the caller gets a `Uint8Array` in its place, checked against its digest.
 */

export const ENGINE_BLOB_ANSWER_TAG = "$blob";

export class EngineBlobAnswerError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "EngineBlobAnswerError";
  }
}

async function sha256Hex(bytes: Uint8Array): Promise<string> {
  const digest = await crypto.subtle.digest("SHA-256", bytes as Uint8Array<ArrayBuffer>);
  return Array.from(new Uint8Array(digest), (byte) => byte.toString(16).padStart(2, "0")).join("");
}

/** `{"$blob": sha}` objects in `value` replaced by their bytes from `blobs`. */
export function resolveEngineAnswerBlobs(
  value: unknown,
  blobs: Pick<ReadonlyMap<string, Uint8Array>, "get">,
): unknown {
  if (Array.isArray(value)) return value.map((item) => resolveEngineAnswerBlobs(item, blobs));
  if (!value || typeof value !== "object") return value;
  const record = value as Record<string, unknown>;
  const keys = Object.keys(record);
  if (keys.length === 1 && keys[0] === ENGINE_BLOB_ANSWER_TAG && typeof record[ENGINE_BLOB_ANSWER_TAG] === "string") {
    const bytes = blobs.get(record[ENGINE_BLOB_ANSWER_TAG] as string);
    if (!bytes) throw new EngineBlobAnswerError("The answer names a file it did not carry.");
    return bytes;
  }
  return Object.fromEntries(keys.map((key) => [key, resolveEngineAnswerBlobs(record[key], blobs)]));
}

/** The result of a framed binary answer, its files verified and in place. */
export async function readEngineBlobAnswer(body: ArrayBuffer): Promise<unknown> {
  const bytes = new Uint8Array(body);
  if (bytes.byteLength < 4) throw new EngineBlobAnswerError("Malformed engine answer.");
  const length = new DataView(body).getUint32(0, false);
  if (4 + length > bytes.byteLength) throw new EngineBlobAnswerError("Malformed engine answer.");
  const head = JSON.parse(new TextDecoder().decode(bytes.subarray(4, 4 + length))) as {
    result?: unknown; blobs?: { sha256?: unknown; size?: unknown }[];
  };
  if (!head || typeof head !== "object" || !("result" in head) || !Array.isArray(head.blobs)) {
    throw new EngineBlobAnswerError("Malformed engine answer.");
  }
  const files = new Map<string, Uint8Array>();
  let offset = 4 + length;
  for (const blob of head.blobs) {
    const size = blob?.size;
    const sha256 = blob?.sha256;
    if (typeof sha256 !== "string" || !/^[a-f0-9]{64}$/.test(sha256) ||
      typeof size !== "number" || !Number.isSafeInteger(size) || size < 0 || offset + size > bytes.byteLength) {
      throw new EngineBlobAnswerError("Malformed engine answer.");
    }
    // A copy: the answer's buffer is one allocation for all of its files.
    const file = bytes.slice(offset, offset + size);
    offset += size;
    if ((await sha256Hex(file)) !== sha256) throw new EngineBlobAnswerError("An engine answer's file failed its digest.");
    files.set(sha256, file);
  }
  if (offset !== bytes.byteLength) throw new EngineBlobAnswerError("Malformed engine answer.");
  return resolveEngineAnswerBlobs(head.result, files);
}
