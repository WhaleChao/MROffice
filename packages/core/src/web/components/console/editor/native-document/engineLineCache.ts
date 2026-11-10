/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The engine's line corrections of a block, kept by its node (text engine on,
 * documentEngineLineBreaks).
 *
 * The page ledger moves blocks from sheet to sheet many times while a long
 * document opens; each move replaces the block, drops its corrections and
 * makes the line pass read it again: R286-D01 opened with the engine on spent
 * 26.5 s in that pass (24 s measuring blocks), most of it on blocks it had
 * already measured, unchanged, at the same width. A block whose drawn lines
 * agreed with the engine's is the same block when it lands: the same node
 * (content, marks and attributes), drawn as wide, in the same faces, under the
 * same document attributes and the same kind of parent. Its corrections are
 * put back from here, without reading the page.
 */
import type { Node as ProseMirrorNode } from "@tiptap/pm/model";
import { newFontFacesOf } from "../../../../lib/font-face-news";

/** What the lines of a block depend on besides the node itself. */
export interface EngineLineKey {
  /** The block's drawn width, px. */
  readonly width: number;
  /** The engine's faces (textEngineEpoch). */
  readonly epoch: string;
  /** The document's attributes (compatibility mode, styles), by identity. */
  readonly docAttrs: unknown;
  /** The parent the block inherits from: a sheet (any), or a cell or list
   * item by its attributes' identity. */
  readonly parent: unknown;
}

/** A correction span, its positions relative to the block's start. */
export interface CachedSpan {
  readonly from: number;
  readonly to: number;
  readonly [field: string]: unknown;
}

interface Entry {
  readonly key: EngineLineKey;
  readonly spans: readonly CachedSpan[];
  /** The families the block's runs name, with the faces of each the browser
   * had settled when it was measured (FamilyNews). */
  readonly faces: ReadonlyArray<readonly [string, number]>;
}








export interface FamilyNews {
  count(family: string): number;
}

const familyNewsByDocument = new WeakMap<Document, FamilyNews>();

export function normalizeFamily(family: string): string {
  return family.replace(/["']/g, "").trim().toLowerCase();
}

export function familyNews(document: Document): FamilyNews {
  const known = familyNewsByDocument.get(document);
  if (known) return known;
  const counts = new Map<string, number>();
  const news: FamilyNews = { count: (family) => counts.get(normalizeFamily(family)) ?? 0 };
  const onSettled = (event: Event) => {
    const faces = newFontFacesOf(event) ?? (event as Event & { fontfaces?: ReadonlyArray<{ family?: string }> }).fontfaces ?? [];
    for (const face of faces) {
      if (!face.family) continue;
      const family = normalizeFamily(face.family);
      counts.set(family, (counts.get(family) ?? 0) + 1);
    }
  };
  document.fonts?.addEventListener?.("loadingdone", onSettled);
  document.fonts?.addEventListener?.("loadingerror", onSettled);
  familyNewsByDocument.set(document, news);
  return news;
}

const WIDTH_EPSILON_PX = 0.01;

function sameFaces(entry: Entry, news: FamilyNews): boolean {
  return entry.faces.every(([family, count]) => news.count(family) === count);
}

function sameKey(a: EngineLineKey, b: EngineLineKey): boolean {
  return a.epoch === b.epoch && a.docAttrs === b.docAttrs &&
    a.parent === b.parent && Math.abs(a.width - b.width) <= WIDTH_EPSILON_PX;
}

export interface EngineLineCache {
  /** The block's spans at `pos`, when it was measured as it is now. */
  lookup<T extends { from: number; to: number }>(node: ProseMirrorNode, key: EngineLineKey, pos: number, news: FamilyNews): T[] | null;
  /** Keep the spans of a block at `pos` whose drawn lines agreed, with the
   * families its runs name. */
  store<T extends { from: number; to: number }>(node: ProseMirrorNode, key: EngineLineKey, pos: number, spans: readonly T[], families: readonly string[], news: FamilyNews): void;
  forget(node: ProseMirrorNode): void;
  /** Why `lookup` found nothing for a block (probes): "node" (never kept),
   * or the first part of the key that differed. */
  missReason(node: ProseMirrorNode, key: EngineLineKey, news: FamilyNews): string;
}

export function createEngineLineCache(): EngineLineCache {
  const entries = new WeakMap<ProseMirrorNode, Entry>();
  return {
    lookup<T extends { from: number; to: number }>(node: ProseMirrorNode, key: EngineLineKey, pos: number, news: FamilyNews): T[] | null {
      const entry = entries.get(node);
      if (!entry || !sameKey(entry.key, key) || !sameFaces(entry, news)) return null;
      return entry.spans.map((span) => ({ ...span, from: span.from + pos, to: span.to + pos }) as unknown as T);
    },
    store(node, key, pos, spans, families, news) {
      const faces = [...new Set(families.map(normalizeFamily))].map((family) => [family, news.count(family)] as const);
      entries.set(node, { key, faces, spans: spans.map((span) => ({ ...span, from: span.from - pos, to: span.to - pos })) });
    },
    forget(node) {
      entries.delete(node);
    },
    missReason(node, key, news) {
      const entry = entries.get(node);
      if (!entry) return "node";
      const kept = entry.key;
      if (!sameFaces(entry, news)) return "faces";
      if (kept.epoch !== key.epoch) return "epoch";
      if (kept.docAttrs !== key.docAttrs) return "docAttrs";
      if (kept.parent !== key.parent) return "parent";
      if (Math.abs(kept.width - key.width) > WIDTH_EPSILON_PX) return "width";
      return "same";
    },
  };
}
