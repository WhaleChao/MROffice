/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Excel's protection password verifiers (sheetProtection / workbookProtection).
 *
 * New passwords get the hash current Excel writes, the ISO write-protection
 * method of [MS-OFFCRYPTO] §2.4.2.4 that ECMA-376 §18.3.1.85 points to:
 *
 *   H0 = H(salt + password)          password = UTF-16LE code units
 *   Hn = H(Hn-1 + iterator)          iterator = uint32 little-endian, 0…spin-1
 *   hashValue = base64(H_final)
 *
 * with SHA-512, a random 16-byte salt and spinCount 100000. Imported files may
 * also carry the 16-bit legacy verifier (`password` / `workbookPassword`),
 * checked by the ECMA-376 Part 4 §15.3.3 routine in `protection.ts`.
 *
 * Not encryption: like Excel's, the verifier only gates the Unprotect command.
 * The plaintext is hashed and dropped; only the four hash fields are stored.
 *
 * WebCrypto's `subtle.digest` resolves once per call, and 100,000 dependent
 * digests took 40 s in Node on the shared machine, so SHA-512 runs here
 * synchronously over pre-split 32-bit words, yielding to the event loop
 * between chunks. WebCrypto supplies the salt, and checks the rarer SHA-1/
 * SHA-256/SHA-384 hashes an imported file may carry.
 */

/** MS-OFFCRYPTO: 1 to 255 Unicode characters. */
export const PROTECTION_PASSWORD_MAX_LENGTH = 255;
/** What Excel writes for a new password. */
export const PROTECTION_PASSWORD_ALGORITHM = "SHA-512";
export const PROTECTION_PASSWORD_SPIN_COUNT = 100_000;
export const PROTECTION_PASSWORD_SALT_BYTES = 16;
/** MS-OFFCRYPTO: SpinCount MUST NOT be greater than 10,000,000. */
export const MAX_PROTECTION_PASSWORD_SPIN_COUNT = 10_000_000;

/** The four ISO hash attributes, exactly as SpreadsheetML stores them. */
export interface ProtectionPasswordHash {
  algorithmName: string;
  /** base64 */
  hashValue: string;
  /** base64 */
  saltValue: string;
  spinCount: number;
}

export type ProtectionPasswordHashCheck =
  | "verified"
  | "invalid-password"
  | "unsupported-algorithm";

const BASE64 = /^(?:[A-Za-z0-9+/]{4})*(?:[A-Za-z0-9+/]{2}==|[A-Za-z0-9+/]{3}=)?$/;
const MAX_BASE64_LENGTH = 1_024;

/** Strict normalizer for a persisted/imported ISO hash; never guesses. */
export function normalizeProtectionPasswordHash(
  value: unknown,
): ProtectionPasswordHash | undefined {
  if (!value || typeof value !== "object" || Array.isArray(value)) return undefined;
  const candidate = value as Record<string, unknown>;
  const { algorithmName, hashValue, saltValue, spinCount } = candidate;
  if (
    typeof algorithmName !== "string" ||
    !/^[A-Za-z0-9-]{1,32}$/.test(algorithmName) ||
    typeof hashValue !== "string" ||
    hashValue.length === 0 ||
    hashValue.length > MAX_BASE64_LENGTH ||
    !BASE64.test(hashValue) ||
    typeof saltValue !== "string" ||
    saltValue.length > MAX_BASE64_LENGTH ||
    !BASE64.test(saltValue) ||
    typeof spinCount !== "number" ||
    !Number.isInteger(spinCount) ||
    spinCount < 0 ||
    spinCount > MAX_PROTECTION_PASSWORD_SPIN_COUNT
  ) {
    return undefined;
  }
  return { algorithmName, hashValue, saltValue, spinCount };
}

export function base64FromBytes(bytes: Uint8Array): string {
  let binary = "";
  for (let index = 0; index < bytes.length; index += 1) {
    binary += String.fromCharCode(bytes[index]);
  }
  return btoa(binary);
}

export function bytesFromBase64(value: string): Uint8Array {
  const binary = atob(value);
  const bytes = new Uint8Array(binary.length);
  for (let index = 0; index < binary.length; index += 1) {
    bytes[index] = binary.charCodeAt(index);
  }
  return bytes;
}

/** The password as SpreadsheetML hashes it: UTF-16LE code units. */
export function utf16leBytes(text: string): Uint8Array {
  const bytes = new Uint8Array(text.length * 2);
  for (let index = 0; index < text.length; index += 1) {
    const unit = text.charCodeAt(index);
    bytes[index * 2] = unit & 0xff;
    bytes[index * 2 + 1] = unit >>> 8;
  }
  return bytes;
}

// ---------------------------------------------------------------------------
// SHA-512 (FIPS 180-4) over 32-bit halves. The constants are derived, not
// transcribed: K = first 64 fractional bits of the cube roots of the first 80
// primes, IV = the same of the square roots of the first 8.

let K: Int32Array | null = null;
let IV: Int32Array | null = null;

function integerRoot(value: bigint, degree: 2n | 3n): bigint {
  // Newton's method from an overestimate; exact floor root for positive n.
  let x = 1n << (BigInt(value.toString(2).length) / degree + 1n);
  while (true) {
    const y =
      ((degree - 1n) * x + value / x ** (degree - 1n)) / degree;
    if (y >= x) return x;
    x = y;
  }
}

function sha512Constants(): { k: Int32Array; iv: Int32Array } {
  if (K && IV) return { k: K, iv: IV };
  const primes: number[] = [];
  for (let candidate = 2; primes.length < 80; candidate += 1) {
    if (primes.every((prime) => candidate % prime !== 0)) primes.push(candidate);
  }
  const mask = (1n << 64n) - 1n;
  const split = (value: bigint, out: Int32Array, index: number) => {
    out[index * 2] = Number((value >> 32n) & 0xffffffffn) | 0;
    out[index * 2 + 1] = Number(value & 0xffffffffn) | 0;
  };
  const k = new Int32Array(160);
  primes.forEach((prime, index) =>
    split(integerRoot(BigInt(prime) << 192n, 3n) & mask, k, index),
  );
  const iv = new Int32Array(16);
  primes
    .slice(0, 8)
    .forEach((prime, index) =>
      split(integerRoot(BigInt(prime) << 128n, 2n) & mask, iv, index),
    );
  K = k;
  IV = iv;
  return { k, iv };
}

const TWO_32 = 0x1_0000_0000;

/** One SHA-512 compression of `block` (32 int32 words, big-endian halves)
 * into `state` (16 int32 words). `w` is a 160-word scratch schedule. */
function compress(state: Int32Array, block: Int32Array, w: Int32Array): void {
  const { k } = sha512Constants();
  for (let index = 0; index < 32; index += 1) w[index] = block[index];
  for (let t = 16; t < 80; t += 1) {
    // σ0(w[t-15]) = ROTR1 ^ ROTR8 ^ SHR7
    let xh = w[(t - 15) * 2];
    let xl = w[(t - 15) * 2 + 1];
    const s0h =
      ((xh >>> 1) | (xl << 31)) ^ ((xh >>> 8) | (xl << 24)) ^ (xh >>> 7);
    const s0l =
      ((xl >>> 1) | (xh << 31)) ^
      ((xl >>> 8) | (xh << 24)) ^
      ((xl >>> 7) | (xh << 25));
    // σ1(w[t-2]) = ROTR19 ^ ROTR61 ^ SHR6
    xh = w[(t - 2) * 2];
    xl = w[(t - 2) * 2 + 1];
    const s1h =
      ((xh >>> 19) | (xl << 13)) ^ ((xl >>> 29) | (xh << 3)) ^ (xh >>> 6);
    const s1l =
      ((xl >>> 19) | (xh << 13)) ^
      ((xh >>> 29) | (xl << 3)) ^
      ((xl >>> 6) | (xh << 26));
    const lo =
      (s0l >>> 0) +
      (s1l >>> 0) +
      (w[(t - 7) * 2 + 1] >>> 0) +
      (w[(t - 16) * 2 + 1] >>> 0);
    w[t * 2] =
      (s0h + s1h + w[(t - 7) * 2] + w[(t - 16) * 2] + Math.floor(lo / TWO_32)) |
      0;
    w[t * 2 + 1] = lo | 0;
  }
  let ah = state[0], al = state[1], bh = state[2], bl = state[3];
  let ch = state[4], cl = state[5], dh = state[6], dl = state[7];
  let eh = state[8], el = state[9], fh = state[10], fl = state[11];
  let gh = state[12], gl = state[13], hh = state[14], hl = state[15];
  for (let t = 0; t < 80; t += 1) {
    // Σ1(e) = ROTR14 ^ ROTR18 ^ ROTR41
    const S1h =
      ((eh >>> 14) | (el << 18)) ^
      ((eh >>> 18) | (el << 14)) ^
      ((el >>> 9) | (eh << 23));
    const S1l =
      ((el >>> 14) | (eh << 18)) ^
      ((el >>> 18) | (eh << 14)) ^
      ((eh >>> 9) | (el << 23));
    const chh = (eh & fh) ^ (~eh & gh);
    const chl = (el & fl) ^ (~el & gl);
    const t1l =
      (hl >>> 0) + (S1l >>> 0) + (chl >>> 0) + (k[t * 2 + 1] >>> 0) +
      (w[t * 2 + 1] >>> 0);
    const t1h =
      (hh + S1h + chh + k[t * 2] + w[t * 2] + Math.floor(t1l / TWO_32)) | 0;
    const t1lo = t1l >>> 0;
    // Σ0(a) = ROTR28 ^ ROTR34 ^ ROTR39
    const S0h =
      ((ah >>> 28) | (al << 4)) ^
      ((al >>> 2) | (ah << 30)) ^
      ((al >>> 7) | (ah << 25));
    const S0l =
      ((al >>> 28) | (ah << 4)) ^
      ((ah >>> 2) | (al << 30)) ^
      ((ah >>> 7) | (al << 25));
    const majh = (ah & bh) ^ (ah & ch) ^ (bh & ch);
    const majl = (al & bl) ^ (al & cl) ^ (bl & cl);
    const t2l = (S0l >>> 0) + (majl >>> 0);
    const t2h = (S0h + majh + Math.floor(t2l / TWO_32)) | 0;
    const t2lo = t2l >>> 0;
    hh = gh; hl = gl;
    gh = fh; gl = fl;
    fh = eh; fl = el;
    const el2 = (dl >>> 0) + t1lo;
    eh = (dh + t1h + Math.floor(el2 / TWO_32)) | 0;
    el = el2 | 0;
    dh = ch; dl = cl;
    ch = bh; cl = bl;
    bh = ah; bl = al;
    const al2 = t1lo + t2lo;
    ah = (t1h + t2h + Math.floor(al2 / TWO_32)) | 0;
    al = al2 | 0;
  }
  const add = (index: number, high: number, low: number) => {
    const sum = (state[index + 1] >>> 0) + (low >>> 0);
    state[index] = (state[index] + high + Math.floor(sum / TWO_32)) | 0;
    state[index + 1] = sum | 0;
  };
  add(0, ah, al); add(2, bh, bl); add(4, ch, cl); add(6, dh, dl);
  add(8, eh, el); add(10, fh, fl); add(12, gh, gl); add(14, hh, hl);
}

function digestState(message: Uint8Array): Int32Array {
  const { iv } = sha512Constants();
  const state = new Int32Array(iv);
  const blockCount = Math.ceil((message.length + 17) / 128);
  const padded = new Uint8Array(blockCount * 128);
  padded.set(message);
  padded[message.length] = 0x80;
  const view = new DataView(padded.buffer);
  view.setUint32(padded.length - 8, Math.floor(message.length / 0x2000_0000));
  view.setUint32(padded.length - 4, (message.length * 8) >>> 0);
  const block = new Int32Array(32);
  const schedule = new Int32Array(160);
  for (let offset = 0; offset < padded.length; offset += 128) {
    for (let index = 0; index < 32; index += 1) {
      block[index] = view.getInt32(offset + index * 4);
    }
    compress(state, block, schedule);
  }
  return state;
}

function stateBytes(state: Int32Array): Uint8Array {
  const out = new Uint8Array(64);
  const view = new DataView(out.buffer);
  for (let index = 0; index < 16; index += 1) view.setInt32(index * 4, state[index]);
  return out;
}

/** SHA-512 of `message`. */
export function sha512(message: Uint8Array): Uint8Array {
  return stateBytes(digestState(message));
}

const yieldToEventLoop = () =>
  new Promise<void>((resolve) => setTimeout(resolve, 0));

/** ISO spin over SHA-512: every iteration hashes exactly 68 bytes (H + i),
 * one padded block whose words are written in place. */
async function spinSha512(
  initial: Int32Array,
  spinCount: number,
  chunk: number,
): Promise<Uint8Array> {
  const { iv } = sha512Constants();
  const block = new Int32Array(32);
  const schedule = new Int32Array(160);
  const state = new Int32Array(16);
  let hash = initial;
  // Byte 68 is the 0x80 pad; the message is 68 * 8 = 544 bits long.
  block[17] = 0x8000_0000 | 0;
  block[31] = 544;
  for (let iterator = 0; iterator < spinCount; iterator += 1) {
    for (let index = 0; index < 16; index += 1) block[index] = hash[index];
    // The iterator's little-endian bytes, read big-endian into word 16.
    block[16] =
      ((iterator & 0xff) << 24) |
      (((iterator >>> 8) & 0xff) << 16) |
      (((iterator >>> 16) & 0xff) << 8) |
      ((iterator >>> 24) & 0xff);
    state.set(iv);
    compress(state, block, schedule);
    hash = state.slice();
    if (chunk > 0 && (iterator + 1) % chunk === 0 && iterator + 1 < spinCount) {
      await yieldToEventLoop();
    }
  }
  return stateBytes(hash);
}

// ---------------------------------------------------------------------------
// MD5 (RFC 1321), one of the ISO write-protection algorithms ECMA-376 lists
// (MS-OFFCRYPTO 2.4.2.4) that WebCrypto does not offer. Excel itself writes
// SHA-512; MD5 hashes come from other writers. The per-round constants are
// derived (floor(abs(sin(i + 1)) * 2^32)), as the RFC defines them.

let MD5_T: Uint32Array | null = null;
const MD5_SHIFTS = [
  7, 12, 17, 22, 5, 9, 14, 20, 4, 11, 16, 23, 6, 10, 15, 21,
];

export function md5(bytes: Uint8Array): Uint8Array {
  if (!MD5_T) {
    MD5_T = new Uint32Array(64);
    for (let index = 0; index < 64; index += 1) {
      MD5_T[index] = Math.floor(Math.abs(Math.sin(index + 1)) * 2 ** 32) >>> 0;
    }
  }
  const t = MD5_T;
  const length = bytes.length;
  const padded = new Uint8Array((((length + 8) >>> 6) + 1) << 6);
  padded.set(bytes);
  padded[length] = 0x80;
  const view = new DataView(padded.buffer);
  view.setUint32(padded.length - 8, (length * 8) >>> 0, true);
  view.setUint32(padded.length - 4, Math.floor(length / 0x20000000), true);
  let a0 = 0x67452301;
  let b0 = 0xefcdab89;
  let c0 = 0x98badcfe;
  let d0 = 0x10325476;
  const m = new Uint32Array(16);
  for (let offset = 0; offset < padded.length; offset += 64) {
    for (let word = 0; word < 16; word += 1) {
      m[word] = view.getUint32(offset + word * 4, true);
    }
    let a = a0;
    let b = b0;
    let c = c0;
    let d = d0;
    for (let index = 0; index < 64; index += 1) {
      const round = index >>> 4;
      let f: number;
      let g: number;
      if (round === 0) {
        f = (b & c) | (~b & d);
        g = index;
      } else if (round === 1) {
        f = (d & b) | (~d & c);
        g = (5 * index + 1) & 15;
      } else if (round === 2) {
        f = b ^ c ^ d;
        g = (3 * index + 5) & 15;
      } else {
        f = c ^ (b | ~d);
        g = (7 * index) & 15;
      }
      const shift = MD5_SHIFTS[round * 4 + (index & 3)]!;
      const sum = (a + f + t[index]! + m[g]!) >>> 0;
      a = d;
      d = c;
      c = b;
      b = (b + ((sum << shift) | (sum >>> (32 - shift)))) >>> 0;
    }
    a0 = (a0 + a) >>> 0;
    b0 = (b0 + b) >>> 0;
    c0 = (c0 + c) >>> 0;
    d0 = (d0 + d) >>> 0;
  }
  const out = new Uint8Array(16);
  const outView = new DataView(out.buffer);
  [a0, b0, c0, d0].forEach((value, index) => outView.setUint32(index * 4, value, true));
  return out;
}

async function spinMd5(
  initial: Uint8Array,
  spinCount: number,
  chunk: number,
): Promise<Uint8Array> {
  let hash = md5(initial);
  const buffer = new Uint8Array(20);
  const view = new DataView(buffer.buffer);
  for (let iterator = 0; iterator < spinCount; iterator += 1) {
    buffer.set(hash, 0);
    view.setUint32(16, iterator, true);
    hash = md5(buffer);
    if (chunk > 0 && (iterator + 1) % chunk === 0 && iterator + 1 < spinCount) {
      await yieldToEventLoop();
    }
  }
  return hash;
}

const WEBCRYPTO_ALGORITHMS: Record<string, string> = {
  "SHA-1": "SHA-1",
  "SHA-256": "SHA-256",
  "SHA-384": "SHA-384",
};

async function spinWebCrypto(
  algorithm: string,
  initial: Uint8Array,
  spinCount: number,
): Promise<Uint8Array> {
  const subtle = globalThis.crypto?.subtle;
  if (!subtle) throw new Error("WebCrypto unavailable");
  const digest = async (bytes: Uint8Array) =>
    new Uint8Array(await subtle.digest(algorithm, bytes as BufferSource));
  let hash = await digest(initial);
  const buffer = new Uint8Array(hash.length + 4);
  const view = new DataView(buffer.buffer);
  for (let iterator = 0; iterator < spinCount; iterator += 1) {
    buffer.set(hash, 0);
    view.setUint32(hash.length, iterator, true);
    hash = await digest(buffer);
  }
  return hash;
}

/** True when this browser can check a hash made with `algorithmName`. */
export function protectionPasswordAlgorithmSupported(algorithmName: string): boolean {
  const name = algorithmName.toUpperCase();
  return name === "SHA-512" || name === "MD5" || name in WEBCRYPTO_ALGORITHMS;
}

/**
 * The base64 hashValue of `password` under `algorithmName`/salt/spin, or
 * undefined when the algorithm cannot be computed here. `chunk` is how many
 * SHA-512 iterations run between yields (0 never yields).
 */
export async function computeProtectionPasswordHash(
  password: string,
  params: Pick<ProtectionPasswordHash, "algorithmName" | "saltValue" | "spinCount">,
  options: { chunk?: number } = {},
): Promise<string | undefined> {
  const name = params.algorithmName.toUpperCase();
  const salt = bytesFromBase64(params.saltValue);
  const passwordBytes = utf16leBytes(password);
  const initial = new Uint8Array(salt.length + passwordBytes.length);
  initial.set(salt, 0);
  initial.set(passwordBytes, salt.length);
  if (name === "SHA-512") {
    const hash = await spinSha512(
      digestState(initial),
      params.spinCount,
      options.chunk ?? 10_000,
    );
    return base64FromBytes(hash);
  }
  if (name === "MD5") {
    return base64FromBytes(
      await spinMd5(initial, params.spinCount, options.chunk ?? 10_000),
    );
  }
  const webCrypto = WEBCRYPTO_ALGORITHMS[name];
  if (!webCrypto) return undefined;
  return base64FromBytes(await spinWebCrypto(webCrypto, initial, params.spinCount));
}

/** Excel's verifier for a NEW password: SHA-512, 16 random bytes, 100000. */
export async function createProtectionPasswordHash(
  password: string,
  options: { salt?: Uint8Array; chunk?: number } = {},
): Promise<ProtectionPasswordHash> {
  let salt = options.salt;
  if (!salt) {
    salt = new Uint8Array(PROTECTION_PASSWORD_SALT_BYTES);
    globalThis.crypto.getRandomValues(salt);
  }
  const saltValue = base64FromBytes(salt);
  const hashValue = await computeProtectionPasswordHash(
    password,
    {
      algorithmName: PROTECTION_PASSWORD_ALGORITHM,
      saltValue,
      spinCount: PROTECTION_PASSWORD_SPIN_COUNT,
    },
    { chunk: options.chunk },
  );
  return {
    algorithmName: PROTECTION_PASSWORD_ALGORITHM,
    // SHA-512 is always computable here.
    hashValue: hashValue as string,
    saltValue,
    spinCount: PROTECTION_PASSWORD_SPIN_COUNT,
  };
}

/** Check `password` against a stored ISO hash. */
export async function checkProtectionPasswordHash(
  password: string,
  stored: ProtectionPasswordHash,
  options: { chunk?: number } = {},
): Promise<ProtectionPasswordHashCheck> {
  if (!protectionPasswordAlgorithmSupported(stored.algorithmName)) {
    return "unsupported-algorithm";
  }
  const computed = await computeProtectionPasswordHash(password, stored, options);
  if (computed === undefined) return "unsupported-algorithm";
  return computed === stored.hashValue ? "verified" : "invalid-password";
}
