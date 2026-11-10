/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Designer's graphics ideas, built (SPEC designer-ms-parity §1.3 items 3–5):
 * a slide's list becomes a SmartArt diagram through the same engine call as
 * Home > Convert to SmartArt, and an icon layout's picture slots are filled
 * with icons from the Insert > Icons library, as PowerPoint fills them from
 * the Microsoft 365 icons.
 */

import type { PresentationBlock } from "@mrnote/core/web/components/console/editor/native-document/documentCommands.ts";
import { officeIconCatalog } from "@mrnote/core/web/lib/office-icons/catalog.ts";
import { officeIconSvg } from "@mrnote/core/web/lib/office-icons/icon-svg.ts";

const ICON_PIXELS = 192;

/** An icon of the library as a PNG data URL (null when it cannot be drawn here). */
export async function designerIconPng(iconId: string, color: string): Promise<string | null> {
  const entry = officeIconCatalog().find((candidate) => candidate.id === iconId);
  const svg = entry ? officeIconSvg(entry, color) : null;
  if (!svg || typeof document === "undefined" || typeof Image === "undefined") return null;
  const image = new Image();
  const loaded = await new Promise<boolean>((resolve) => {
    const timer = setTimeout(() => resolve(false), 2_000);
    image.onload = () => {
      clearTimeout(timer);
      resolve(true);
    };
    image.onerror = () => {
      clearTimeout(timer);
      resolve(false);
    };
    image.src = `data:image/svg+xml;charset=utf-8,${encodeURIComponent(svg)}`;
  });
  if (!loaded) return null;
  const canvas = document.createElement("canvas");
  canvas.width = ICON_PIXELS;
  canvas.height = ICON_PIXELS;
  const context = canvas.getContext("2d");
  if (!context) return null;
  context.drawImage(image, 0, 0, ICON_PIXELS, ICON_PIXELS);
  try {
    return canvas.toDataURL("image/png");
  } catch {
    return null;
  }
}

/**
 * The diagram with each Level 1 item's picture slot holding its icon. `icons`
 * are in Level 1 order; `levels` is the outline the diagram was built from.
 * A slot whose icon cannot be drawn stays empty (the diagram is still valid).
 */
export async function designerDiagramWithIcons(
  nodes: readonly PresentationBlock[],
  icons: readonly string[],
  levels: readonly number[],
  color: string,
): Promise<PresentationBlock[]> {
  // Outline row index -> Level 1 ordinal.
  const top = Math.min(...levels);
  const ordinalOfRow = new Map<number, number>();
  let ordinal = -1;
  levels.forEach((level, row) => {
    if (level === top) ordinal += 1;
    ordinalOfRow.set(row, ordinal);
  });
  const pictures = new Map<string, string | null>();
  await Promise.all(
    [...new Set(icons)].map(async (id) => pictures.set(id, await designerIconPng(id, color))),
  );
  return nodes.map((node) => {
    const record = node as Record<string, unknown>;
    if (node.type !== "image" || record.smartArtRole !== "node") return node;
    const row = typeof record.smartArtItemIndex === "number" ? record.smartArtItemIndex : null;
    const at = row === null ? undefined : ordinalOfRow.get(row);
    const icon = at === undefined ? undefined : icons[at];
    const src = icon ? pictures.get(icon) : null;
    return src ? ({ ...node, src } as PresentationBlock) : node;
  });
}
