/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Feature detection and sampling for the browser's native EyeDropper API
 * (Chromium only; absent in jsdom and Firefox).
 *
 * Moved verbatim out of `components/notebook/contents/PresentationWindow.tsx`
 * (structural split, no behaviour change).
 */



/**
 * Minimal structural type for the browser's native EyeDropper API (Chromium /
 * not in the DOM lib types, absent in jsdom / Firefox). We never assume it
 * exists — `eyeDropperSupported` feature-detects before any use.
 */
interface EyeDropperResult {
  sRGBHex: string;
}

interface EyeDropperInstance {
  open: (options?: { signal?: AbortSignal }) => Promise<EyeDropperResult>;
}

type EyeDropperCtor = new () => EyeDropperInstance;

/**
 * Whether the native EyeDropper color picker is available in this environment.
 * False in jsdom / Firefox / older browsers — callers HIDE the eyedropper
 * affordance when this is false so it never throws.
 */
export function eyeDropperSupported(): boolean {
  return (
    typeof window !== "undefined" &&
    typeof (window as { EyeDropper?: EyeDropperCtor }).EyeDropper === "function"
  );
}

/**
 * Open the native eyedropper and resolve to the sampled six-digit hex colour
 * (lower-cased), or `null` when unsupported, cancelled (the user pressed Esc),
 * or the sampled value is not a clean six-digit hex. NEVER throws — a rejected
 * `open()` (cancellation) resolves to `null`. Pure feature-detect at the top so
 * a missing API is a graceful no-op.
 */
export async function sampleEyeDropperColor(): Promise<string | null> {
  if (!eyeDropperSupported()) return null;
  try {
    const Ctor = (window as unknown as { EyeDropper: EyeDropperCtor })
      .EyeDropper;
    const result = await new Ctor().open();
    const hex =
      typeof result?.sRGBHex === "string" ? result.sRGBHex.trim() : "";
    return /^#[0-9a-fA-F]{6}$/.test(hex) ? hex.toLowerCase() : null;
  } catch {
    // Cancellation (Esc) rejects the promise — treat as a no-op, never throw.
    return null;
  }
}
