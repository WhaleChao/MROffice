/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Sounds an imported deck keeps in its own package (readings rows 559,
 * 560): a slide's transition sound (`sourceKind` "package") and its
 * effects' sounds (`nativeTimingTree.sounds`) are named by their digest and
 * played from the page's package. A slide taken into another deck (Reuse
 * Slides, the AI's reuse proposal) leaves that package behind, so its
 * sounds are read from the source page first -- through the same route the
 * show plays them from, checked by digest -- and travel inline, as they did
 * before they were references: the target deck plays and saves them.
 */

import { getApiHttpBaseUrl } from "@mrnote/core/web/lib/env.ts";
import {
  MAX_PRESENTATION_TRANSITION_SOUND_BYTES,
  normalizePresentationTransitionSound,
} from "@mrnote/core/web/lib/presentation-transitions.ts";

/** An effect's sound past this is not carried (PowerPoint's are tens of kilobytes). */
const MAX_EFFECT_SOUND_BYTES = 20 * 1024 * 1024;
const WAV_DATA_URL = /^data:audio\/(?:wav|x-wav|wave);base64,[A-Za-z0-9+/]+={0,2}$/;

type Fetcher = (url: string, init: RequestInit) => Promise<Response>;

function base64(bytes: Uint8Array): string {
  let binary = "";
  for (let at = 0; at < bytes.length; at += 0x8000) binary += String.fromCharCode(...bytes.subarray(at, at + 0x8000));
  return btoa(binary);
}

async function sha256Hex(bytes: Uint8Array): Promise<string> {
  const digest = await crypto.subtle.digest("SHA-256", bytes as BufferSource);
  return Array.from(new Uint8Array(digest), (byte) => byte.toString(16).padStart(2, "0")).join("");
}

/** An effect sound's inline copy, as the tree player takes it (`sounds[digest].dataUrl`). */
export function inlineEffectSoundUrl(tree: unknown, digest: string): string | null {
  const sounds = tree && typeof tree === "object" ? (tree as { sounds?: unknown }).sounds : null;
  const entry = sounds && typeof sounds === "object" ? (sounds as Record<string, unknown>)[digest] : null;
  const url = entry && typeof entry === "object" ? (entry as { dataUrl?: unknown }).dataUrl : null;
  return typeof url === "string" && WAV_DATA_URL.test(url) ? url : null;
}

/**
 * ``content``'s slides with every package sound read from ``pageId``'s
 * package and inlined (a copy; ``content`` is not changed). A sound that
 * cannot be read stays a reference: silent in another deck, as before.
 */
export async function inlinePackageSounds<T extends { slides?: unknown }>(
  content: T,
  pageId: string,
  fetcher: Fetcher = (url, init) => fetch(url, init),
): Promise<T> {
  const slides = Array.isArray(content?.slides) ? (content.slides as Array<Record<string, unknown>>) : [];
  const wanted = new Map<string, number>();
  for (const slide of slides) {
    const sound = normalizePresentationTransitionSound(slide?.transitionSound);
    if (sound && "fingerprint" in sound) wanted.set(sound.fingerprint, MAX_PRESENTATION_TRANSITION_SOUND_BYTES);
    const sounds = (slide?.nativeTimingTree as { sounds?: Record<string, unknown> } | undefined)?.sounds;
    for (const digest of sounds && typeof sounds === "object" ? Object.keys(sounds) : []) {
      if (/^[0-9a-f]{64}$/.test(digest) && !wanted.has(digest)) wanted.set(digest, MAX_EFFECT_SOUND_BYTES);
    }
  }
  if (wanted.size === 0) return content;
  const read = new Map<string, string>();
  const base = getApiHttpBaseUrl();
  await Promise.all(
    [...wanted].map(async ([digest, limit]) => {
      try {
        const response = await fetcher(
          `${base}/api/v1/pages/${encodeURIComponent(pageId)}/presentation-native-media/snd-${digest}`,
          { credentials: "include" },
        );
        if (!response.ok) return;
        const bytes = new Uint8Array(await response.arrayBuffer());
        if (
          bytes.length < 12 ||
          bytes.length > limit ||
          String.fromCharCode(...bytes.subarray(0, 4)) !== "RIFF" ||
          String.fromCharCode(...bytes.subarray(8, 12)) !== "WAVE" ||
          (await sha256Hex(bytes)) !== digest
        ) {
          return;
        }
        read.set(digest, `data:audio/wav;base64,${base64(bytes)}`);
      } catch {
        // Not readable: it stays a reference.
      }
    }),
  );
  if (read.size === 0) return content;
  const next = slides.map((slide) => {
    let out = slide;
    const sound = normalizePresentationTransitionSound(slide?.transitionSound);
    if (sound && "fingerprint" in sound && read.has(sound.fingerprint)) {
      out = {
        ...out,
        transitionSound: { dataUrl: read.get(sound.fingerprint)!, mimeType: "audio/wav", name: sound.name, sizeBytes: sound.sizeBytes },
      };
    }
    const tree = slide?.nativeTimingTree as { sounds?: Record<string, Record<string, unknown>> } | undefined;
    if (tree?.sounds && Object.keys(tree.sounds).some((digest) => read.has(digest))) {
      const sounds = Object.fromEntries(
        Object.entries(tree.sounds).map(([digest, entry]) => [digest, read.has(digest) ? { ...entry, dataUrl: read.get(digest)! } : entry]),
      );
      out = { ...out, nativeTimingTree: { ...tree, sounds } };
    }
    return out;
  });
  return { ...content, slides: next };
}
