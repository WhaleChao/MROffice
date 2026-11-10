/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The drawing a LEGACY Docs `smartArt` block keeps with it.
 *
 * A legacy block stores only its layout id and item list; its picture is the
 * layout engine's, and the engine is the server's. So the first time an
 * editable copy of the document is open online, the editor asks the server
 * once for the scene and stores it on the block (`attrs.scene`), and every
 * later opening draws that stored scene with no request -- read-only, as a
 * guest, offline. The scene is keyed by what it was drawn FROM (layout, item
 * texts, outline levels): an item edit makes it stale, and the next online,
 * editable moment draws and stores it again.
 *
 * Pure: types, the key, the request a block's scene is asked with, and the
 * reader that accepts a stored scene only when it is well formed and current.
 */

import type { OfficeSmartArtNode, SmartArtKind } from "./model";
import { DOCUMENT_SMARTART_FRAME_EMU, DOCUMENT_SMARTART_STYLE } from "./document-shapes-model";
import type { SmartArtPreviewRequest } from "../engine-host/smartart-types";

export const LEGACY_SMARTART_SCENE_VERSION = 1;
/** Mirrors the API's bound on a stored scene (`_validate_smart_art_node_attrs`). */
export const MAX_LEGACY_SMARTART_SCENE_NODES = 1_000;

export interface LegacySmartArtScene {
  v: typeof LEGACY_SMARTART_SCENE_VERSION;
  /** {@link legacySmartArtSceneKey} of what the scene was drawn from. */
  key: string;
  nodes: OfficeSmartArtNode[];
}

 
function fingerprint(text: string): string {
  let a = 0x811c9dc5;
  let b = 0x01000193 ^ 0x5bd1e995;
  for (let index = 0; index < text.length; index += 1) {
    const code = text.charCodeAt(index);
    a = Math.imul(a ^ code, 0x01000193) >>> 0;
    b = Math.imul(b ^ code, 0x5bd1e995) >>> 0;
  }
  return a.toString(16).padStart(8, "0") + b.toString(16).padStart(8, "0");
}

/** The key of the scene drawn from this layout, these item texts and levels. */
export function legacySmartArtSceneKey(
  kind: SmartArtKind,
  items: readonly string[],
  levels: readonly number[],
): string {
  return fingerprint(JSON.stringify([kind, items, levels]));
}

/** The one request a legacy block's scene is asked with (the `previews` method). */
export function legacySmartArtSceneRequest(
  kind: SmartArtKind,
  items: readonly string[],
  levels: readonly number[],
): SmartArtPreviewRequest["previews"][number] {
  return {
    kind,
    items: [...items],
    levels: [...levels],
    bounds: { width: DOCUMENT_SMARTART_FRAME_EMU.width, height: DOCUMENT_SMARTART_FRAME_EMU.height },
    style: DOCUMENT_SMARTART_STYLE,
  };
}

function record(value: unknown): Record<string, unknown> | null {
  return value && typeof value === "object" && !Array.isArray(value)
    ? (value as Record<string, unknown>)
    : null;
}

const finite = (value: unknown) => typeof value === "number" && Number.isFinite(value);

/**
 * The stored scene's nodes when `value` is a well-formed scene drawn from
 * `key`; otherwise null (absent, stale, or not a scene). A stored picture
 * member keeps only an inline image: the document never makes a reader fetch.
 */
export function storedLegacySmartArtScene(value: unknown, key: string): OfficeSmartArtNode[] | null {
  const scene = record(value);
  if (!scene || scene.v !== LEGACY_SMARTART_SCENE_VERSION || scene.key !== key) return null;
  const nodes = scene.nodes;
  if (!Array.isArray(nodes) || nodes.length === 0 || nodes.length > MAX_LEGACY_SMARTART_SCENE_NODES) {
    return null;
  }
  const read: OfficeSmartArtNode[] = [];
  for (const entry of nodes) {
    const node = record(entry);
    if (
      !node ||
      (node.type !== "shape" && node.type !== "text" && node.type !== "image") ||
      !finite(node.left) ||
      !finite(node.top) ||
      !finite(node.width) ||
      !finite(node.height)
    ) {
      return null;
    }
    if (node.type === "image") {
      const src = typeof node.src === "string" && /^data:image\//i.test(node.src) ? node.src : "";
      read.push({ ...(node as unknown as OfficeSmartArtNode), src } as OfficeSmartArtNode);
    } else {
      read.push(node as unknown as OfficeSmartArtNode);
    }
  }
  return read;
}
