/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
type CryptoLike = {
  randomUUID?: () => string;
  getRandomValues?: <T extends ArrayBufferView | null>(array: T) => T;
};

function runtimeCrypto(): CryptoLike | undefined {
  return (globalThis as { crypto?: CryptoLike }).crypto;
}

const HEX: readonly string[] = Array.from({ length: 256 }, (_, byte) => byte.toString(16).padStart(2, "0"));

/** A v4 UUID from 16 random bytes (RFC 9562 §5.4: version and variant bits set). */
function uuidFromBytes(bytes: Uint8Array): string {
  bytes[6] = (bytes[6] & 0x0f) | 0x40;
  bytes[8] = (bytes[8] & 0x3f) | 0x80;
  const h = (index: number) => HEX[bytes[index]];
  return (
    h(0) + h(1) + h(2) + h(3) + "-" +
    h(4) + h(5) + "-" +
    h(6) + h(7) + "-" +
    h(8) + h(9) + "-" +
    h(10) + h(11) + h(12) + h(13) + h(14) + h(15)
  );
}

/** A v4 UUID that does not use `crypto.randomUUID` (the fallback itself). */
export function fallbackRandomUUID(): string {
  const bytes = new Uint8Array(16);
  const cryptoApi = runtimeCrypto();
  if (typeof cryptoApi?.getRandomValues === "function") {
    cryptoApi.getRandomValues(bytes);
  } else {
    for (let index = 0; index < bytes.length; index += 1) bytes[index] = Math.floor(Math.random() * 256);
  }
  return uuidFromBytes(bytes);
}

/** `crypto.randomUUID()` where it exists, else the same shape of id from `getRandomValues`. */
export function randomUUID(): string {
  const cryptoApi = runtimeCrypto();
  if (typeof cryptoApi?.randomUUID === "function") return cryptoApi.randomUUID();
  return fallbackRandomUUID();
}

/**
 * Gives `crypto` a `randomUUID` where the context withholds it, so the many
 * call sites (and libraries) that call `crypto.randomUUID()` directly work
 * there too. Installed once, before the app's code runs
 * (`instrumentation-client.ts`). A no-op where the native one exists.
 */
export function installRandomUUIDFallback(): void {
  const cryptoApi = runtimeCrypto();
  if (!cryptoApi || typeof cryptoApi.randomUUID === "function") return;
  try {
    Object.defineProperty(cryptoApi, "randomUUID", {
      value: fallbackRandomUUID,
      configurable: true,
      writable: true,
    });
  } catch {
    // A frozen crypto object: callers that use `randomUUID()` above still work.
  }
}
