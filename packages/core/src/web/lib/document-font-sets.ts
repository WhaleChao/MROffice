/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
export const DOCUMENT_FONT_SETS = Object.freeze([
  {
    id: "office",
    family:
      'Aptos, Calibri, DengXian, "等线", "PingFang SC", "Microsoft YaHei", sans-serif',
  },
  {
    id: "classic",
    family: '"Times New Roman", Times, SimSun, "宋体", "Songti SC", serif',
  },
  {
    id: "modern",
    family: 'Arial, Helvetica, "PingFang SC", "Microsoft YaHei", sans-serif',
  },
  {
    id: "cjkSerif",
    family: 'SimSun, "宋体", "Songti SC", "Noto Serif SC", serif',
  },
  {
    id: "cjkSans",
    family: '"Microsoft YaHei", "微软雅黑", "PingFang SC", sans-serif',
  },
  {
    id: "mono",
    family: '"Courier New", Courier, monospace',
  },
] as const);

export type DocumentFontSetId = (typeof DOCUMENT_FONT_SETS)[number]["id"];

/**
 * Word's heading font ("+Headings": the Office theme's major font — Aptos
 * Display / Calibri Light in a Latin Office, 等线 Light in a Chinese one), on
 * the faces the font layer ships. The gallery's Heading 1–3, Title and
 * Subtitle are set in it, and the DOCX writer states the same faces.
 */
export const DOCUMENT_WORD_HEADING_FONT_FAMILY =
  '"Aptos Display", "Calibri Light", "DengXian Light", "等线 Light", Aptos, Calibri, DengXian, "等线", "PingFang SC", "Microsoft YaHei", sans-serif';

export function documentFontSet(
  id: unknown,
): (typeof DOCUMENT_FONT_SETS)[number] | null {
  return DOCUMENT_FONT_SETS.find((preset) => preset.id === id) ?? null;
}

function fontTokens(value: unknown): string[] {
  if (typeof value !== "string") return [];
  return value
    .split(",")
    .map((token) =>
      token
        .trim()
        .replace(/^["']|["']$/g, "")
        .toLowerCase(),
    )
    .filter(
      (token) =>
        token &&
        ![
          "serif",
          "sans-serif",
          "monospace",
          "system-ui",
          "ui-serif",
          "ui-sans-serif",
          "ui-monospace",
        ].includes(token),
    );
}

/** Match both the full authored stack and its effective OOXML Latin/CJK faces. */
export function documentFontSetFor(value: unknown): DocumentFontSetId | null {
  const active = fontTokens(value);
  if (active.length === 0) return null;
  return (
    DOCUMENT_FONT_SETS.find((preset) => {
      const expected = fontTokens(preset.family);
      return (
        active[0] === expected[0] &&
        active.every((token) => expected.includes(token))
      );
    })?.id ?? null
  );
}
