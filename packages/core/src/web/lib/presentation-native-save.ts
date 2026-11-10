/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { pageDocumentJsonEquals } from "./page-save-acknowledgement";

const record = (value: unknown): Record<string, unknown> | null =>
  value !== null && typeof value === "object" && !Array.isArray(value)
    ? value as Record<string, unknown> : null;

/** Match only the native writer's established text-representation equivalences.
 * This compares documents; it never certifies native materialization or clears
 * a pending draft. Unknown fields and all package receipts remain significant. */
export function presentationNativeSavedContentMatches(left: unknown, right: unknown): boolean {
  const normalize = (value: unknown) => {
    const source = record(value);
    if (source?.type !== "mrnote.presentation.v1" || !Array.isArray(source.slides)) return null;
    const content = structuredClone(source);
    for (const rawSlide of content.slides as unknown[]) {
      const slide = record(rawSlide);
      if (!slide || !Array.isArray(slide.blocks)) return null;
      for (const rawBlock of slide.blocks) {
        const block = record(rawBlock);
        if (!block || !["text", "shape"].includes(String(block.type))) continue;
        if (block.style != null && !record(block.style)) continue;
        const style = { ...(record(block.style) ?? {}) };
        const allowed = {
          align: ["left", "center", "right", "justify"],
          verticalAlign: ["top", "middle", "bottom"],
          writingMode: ["horizontal-tb", "vertical-rl", "rot90", "rot270", "stacked"],
        };
        for (const [key, values] of Object.entries(allowed)) {
          const raw = block[key];
          if (typeof raw === "string" && values.includes(raw) &&
              (!Object.hasOwn(style, key) || style[key] === raw)) {
            style[key] = raw;
            delete block[key];
          }
        }
        const noAutofit = (value: unknown) => value == null || value === "none";
        if (noAutofit(block.autofit) && noAutofit(style.autofit)) {
          delete block.autofit;
          delete style.autofit;
        }
        if (Object.keys(style).length) block.style = style;
        else delete block.style;
      }
    }
    return content;
  };
  const a = normalize(left), b = normalize(right);
  return a !== null && b !== null && pageDocumentJsonEquals(a, b);
}
