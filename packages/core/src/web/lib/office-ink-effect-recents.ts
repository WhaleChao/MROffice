/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The custom pen effects the user built lately, newest first, kept in this
 * browser as the pens themselves are and shared by every editor's pen menu
 * (Docs, the deck, the worksheet): a custom effect made on a slide is there
 * to pick in a document too. Another tab's change follows (`storage`).
 */

import { useSyncExternalStore } from "react";

import { isCustomInkEffect, parseCustomInkEffect, type CustomInkEffect } from "./office-ink-effects";

export const INK_EFFECT_RECENTS_STORAGE_KEY = "mrai:ink:custom-effects:v1";
export const INK_EFFECT_RECENTS_MAX = 7;

const EMPTY: readonly CustomInkEffect[] = [];
let memory: readonly CustomInkEffect[] | null = null;
const listeners = new Set<() => void>();

/** A stored list: custom effects only, each once, newest first, at most the max. */
export function normalizeInkEffectRecents(value: unknown): CustomInkEffect[] {
  if (!Array.isArray(value)) return [];
  const out: CustomInkEffect[] = [];
  for (const entry of value) {
    if (!isCustomInkEffect(entry) || out.includes(entry)) continue;
    out.push(entry);
    if (out.length >= INK_EFFECT_RECENTS_MAX) break;
  }
  return out;
}

/** `effect` moved to the head of `recents`. */
export function pushInkEffectRecent(
  recents: readonly CustomInkEffect[],
  effect: unknown,
): CustomInkEffect[] {
  if (!isCustomInkEffect(effect)) return [...recents];
  return normalizeInkEffectRecents([effect, ...recents.filter((entry) => entry !== effect)]);
}

function read(): readonly CustomInkEffect[] {
  if (memory) return memory;
  try {
    const stored = window.localStorage.getItem(INK_EFFECT_RECENTS_STORAGE_KEY);
    memory = stored ? normalizeInkEffectRecents(JSON.parse(stored)) : EMPTY;
  } catch {
    memory = EMPTY;
  }
  return memory;
}

export function inkEffectRecents(): readonly CustomInkEffect[] {
  if (typeof window === "undefined") return EMPTY;
  return read();
}

/** Remember a custom effect the user made (anything else is ignored). */
export function rememberInkEffect(effect: unknown): void {
  if (typeof window === "undefined" || !isCustomInkEffect(effect)) return;
  const next = pushInkEffectRecent(read(), effect);
  memory = next;
  try {
    window.localStorage.setItem(INK_EFFECT_RECENTS_STORAGE_KEY, JSON.stringify(next));
  } catch {
    /* storage unavailable: this page's memory */
  }
  for (const listener of listeners) listener();
}

/** The colours of the recent custom effects, newest first, each once (a palette's recents). */
export function inkEffectRecentColors(recents: readonly CustomInkEffect[]): string[] {
  const colors: string[] = [];
  for (const effect of recents) {
    for (const color of parseCustomInkEffect(effect)?.colors ?? []) {
      if (!colors.includes(color)) colors.push(color);
    }
  }
  return colors.slice(0, 10);
}

function subscribe(listener: () => void): () => void {
  listeners.add(listener);
  const onStorage = (event: StorageEvent) => {
    if (event.key !== INK_EFFECT_RECENTS_STORAGE_KEY) return;
    memory = null;
    listener();
  };
  if (typeof window !== "undefined") window.addEventListener("storage", onStorage);
  return () => {
    listeners.delete(listener);
    if (typeof window !== "undefined") window.removeEventListener("storage", onStorage);
  };
}

export function useInkEffectRecents(): readonly CustomInkEffect[] {
  return useSyncExternalStore(subscribe, inkEffectRecents, () => EMPTY);
}
