/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The SHA-256 of a file of any size, without holding it whole.
 *
 * WebCrypto digests one buffer: a file up to {@link WHOLE_DIGEST_MAX_BYTES}
 * is read once and handed to it; a larger one is read in slices and fed to
 * the incremental digest below (FIPS 180-4), so memory stays one slice.
 */

const WHOLE_DIGEST_MAX_BYTES = 64 * 1024 * 1024;
const SLICE_BYTES = 8 * 1024 * 1024;

const K = new Uint32Array([
  0x428a2f98, 0x71374491, 0xb5c0fbcf, 0xe9b5dba5, 0x3956c25b, 0x59f111f1, 0x923f82a4, 0xab1c5ed5,
  0xd807aa98, 0x12835b01, 0x243185be, 0x550c7dc3, 0x72be5d74, 0x80deb1fe, 0x9bdc06a7, 0xc19bf174,
  0xe49b69c1, 0xefbe4786, 0x0fc19dc6, 0x240ca1cc, 0x2de92c6f, 0x4a7484aa, 0x5cb0a9dc, 0x76f988da,
  0x983e5152, 0xa831c66d, 0xb00327c8, 0xbf597fc7, 0xc6e00bf3, 0xd5a79147, 0x06ca6351, 0x14292967,
  0x27b70a85, 0x2e1b2138, 0x4d2c6dfc, 0x53380d13, 0x650a7354, 0x766a0abb, 0x81c2c92e, 0x92722c85,
  0xa2bfe8a1, 0xa81a664b, 0xc24b8b70, 0xc76c51a3, 0xd192e819, 0xd6990624, 0xf40e3585, 0x106aa070,
  0x19a4c116, 0x1e376c08, 0x2748774c, 0x34b0bcb5, 0x391c0cb3, 0x4ed8aa4a, 0x5b9cca4f, 0x682e6ff3,
  0x748f82ee, 0x78a5636f, 0x84c87814, 0x8cc70208, 0x90befffa, 0xa4506ceb, 0xbef9a3f7, 0xc67178f2,
]);

/** An incremental SHA-256: update() with the bytes in order, then hex(). */
export class IncrementalSha256 {
  private readonly state = new Uint32Array([
    0x6a09e667, 0xbb67ae85, 0x3c6ef372, 0xa54ff53a, 0x510e527f, 0x9b05688c, 0x1f83d9ab, 0x5be0cd19,
  ]);
  private readonly block = new Uint8Array(64);
  private readonly words = new Uint32Array(64);
  private blockLength = 0;
  private totalBytes = 0;

  update(bytes: Uint8Array): void {
    let offset = 0;
    this.totalBytes += bytes.length;
    if (this.blockLength > 0) {
      const take = Math.min(64 - this.blockLength, bytes.length);
      this.block.set(bytes.subarray(0, take), this.blockLength);
      this.blockLength += take;
      offset = take;
      if (this.blockLength < 64) return;
      this.compress(this.block, 0);
      this.blockLength = 0;
    }
    for (; offset + 64 <= bytes.length; offset += 64) this.compress(bytes, offset);
    if (offset < bytes.length) {
      this.block.set(bytes.subarray(offset), 0);
      this.blockLength = bytes.length - offset;
    }
  }

  hex(): string {
    const bitsHigh = Math.floor((this.totalBytes * 8) / 0x1_0000_0000);
    const bitsLow = (this.totalBytes * 8) >>> 0;
    const padding = new Uint8Array(((this.blockLength < 56 ? 56 : 120) - this.blockLength) + 8);
    padding[0] = 0x80;
    const view = new DataView(padding.buffer);
    view.setUint32(padding.length - 8, bitsHigh, false);
    view.setUint32(padding.length - 4, bitsLow, false);
    const total = this.totalBytes;
    this.update(padding);
    this.totalBytes = total;
    return Array.from(this.state, (word) => word.toString(16).padStart(8, "0")).join("");
  }

  private compress(bytes: Uint8Array, offset: number): void {
    const w = this.words;
    for (let index = 0; index < 16; index += 1) {
      const at = offset + index * 4;
      w[index] = (bytes[at] << 24) | (bytes[at + 1] << 16) | (bytes[at + 2] << 8) | bytes[at + 3];
    }
    for (let index = 16; index < 64; index += 1) {
      const a = w[index - 15];
      const b = w[index - 2];
      const s0 = ((a >>> 7) | (a << 25)) ^ ((a >>> 18) | (a << 14)) ^ (a >>> 3);
      const s1 = ((b >>> 17) | (b << 15)) ^ ((b >>> 19) | (b << 13)) ^ (b >>> 10);
      w[index] = (w[index - 16] + s0 + w[index - 7] + s1) | 0;
    }
    let [a, b, c, d, e, f, g, h] = this.state;
    for (let index = 0; index < 64; index += 1) {
      const S1 = ((e >>> 6) | (e << 26)) ^ ((e >>> 11) | (e << 21)) ^ ((e >>> 25) | (e << 7));
      const ch = (e & f) ^ (~e & g);
      const t1 = (h + S1 + ch + K[index] + w[index]) | 0;
      const S0 = ((a >>> 2) | (a << 30)) ^ ((a >>> 13) | (a << 19)) ^ ((a >>> 22) | (a << 10));
      const maj = (a & b) ^ (a & c) ^ (b & c);
      const t2 = (S0 + maj) | 0;
      h = g;
      g = f;
      f = e;
      e = (d + t1) | 0;
      d = c;
      c = b;
      b = a;
      a = (t1 + t2) | 0;
    }
    const state = this.state;
    state[0] += a;
    state[1] += b;
    state[2] += c;
    state[3] += d;
    state[4] += e;
    state[5] += f;
    state[6] += g;
    state[7] += h;
  }
}

async function wholeDigestHex(blob: Blob): Promise<string> {
  const digest = await crypto.subtle.digest("SHA-256", await blob.arrayBuffer());
  return Array.from(new Uint8Array(digest), (byte) => byte.toString(16).padStart(2, "0")).join("");
}

/** Lowercase hex SHA-256 of `blob`, read in slices past 64 MiB. */
export async function blobSha256Hex(
  blob: Blob,
  options: { readonly sliceBytes?: number; readonly forceIncremental?: boolean } = {},
): Promise<string> {
  if (!options.forceIncremental && blob.size <= WHOLE_DIGEST_MAX_BYTES && globalThis.crypto?.subtle) {
    return wholeDigestHex(blob);
  }
  const hash = new IncrementalSha256();
  const slice = Math.max(64, options.sliceBytes ?? SLICE_BYTES);
  for (let start = 0; start < blob.size; start += slice) {
    hash.update(new Uint8Array(await blob.slice(start, Math.min(blob.size, start + slice)).arrayBuffer()));
  }
  return hash.hex();
}
