/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type {
  PdfNativeDeviceColor,
  PdfNativeTextFontOption,
  PdfNativeTextPropertyDraft,
} from "./engine/native/list-text-fidelity";
import type { PdfTextStyleRun } from "@mrnote/core/web/lib/pdf/text-style-runs.ts";

function hexOf(color: PdfNativeDeviceColor | undefined): string | undefined {
  if (!color) return undefined;
  const byte = (value: number) => Math.round(Math.max(0, Math.min(1, value)) * 255).toString(16).padStart(2, "0");
  if (color.kind === "gray") return `#${byte(color.components[0]).repeat(3)}`;
  if (color.kind === "rgb") return `#${color.components.map(byte).join("")}`;
  const [c, m, y, k] = color.components;
  return `#${[c, m, y].map((value) => byte((1 - value) * (1 - k))).join("")}`;
}

export function pdfNativeDraftTextRuns(input: {
  readonly text: string;
  readonly draft: PdfNativeTextPropertyDraft;
  readonly fonts: readonly PdfNativeTextFontOption[];
  /** The box's own size, a fraction of the page's height. */
  readonly boxFontSize: number;
}): PdfTextStyleRun[] | null {
  const segments = (input.draft.segments ?? []).filter((segment) =>
    segment.end > segment.start && segment.start >= 0 && segment.end <= input.text.length);
  if (segments.length === 0) return null;
  const face = (name: string | undefined) => input.fonts.find((font) => font.resourceName === name);
  const lineFace = face(input.draft.fontResourceName);
  const runs: PdfTextStyleRun[] = [];
  const plain = (text: string) => ({
    text,
    ...(lineFace?.bold || input.draft.syntheticBold ? { bold: true } : {}),
    ...(lineFace?.italic ? { italic: true } : {}),
  });
  let at = 0;
  for (const segment of [...segments].sort((left, right) => left.start - right.start)) {
    if (segment.start < at) return null;
    if (segment.start > at) runs.push(plain(input.text.slice(at, segment.start)));
    const own = face(segment.fontResourceName) ?? lineFace;
    const bold = Boolean(own?.bold || (segment.syntheticBold ?? input.draft.syntheticBold));
    const italic = Boolean(own?.italic || segment.slanted);
    const color = hexOf(segment.color);
    const size = segment.fontSize && input.draft.fontSize > 0
      ? input.boxFontSize * segment.fontSize / input.draft.fontSize : undefined;
    runs.push({
      text: input.text.slice(segment.start, segment.end),
      ...(bold ? { bold: true } : {}),
      ...(italic ? { italic: true } : {}),
      ...(color ? { color } : {}),
      ...(size && Math.abs(size - input.boxFontSize) > 1e-9 ? { fontSize: size } : {}),
      ...(segment.textRise ? { script: segment.textRise > 0 ? "superscript" as const : "subscript" as const } : {}),
    });
    at = segment.end;
  }
  if (at < input.text.length) runs.push(plain(input.text.slice(at)));
  return runs;
}
