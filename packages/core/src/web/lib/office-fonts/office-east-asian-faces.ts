/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import table from "./office-east-asian-faces.generated.json";

const FACES = new Map<string, number | null>(
  Object.entries((table as { faces: Record<string, number | null> }).faces),
);
const LINES = new Map<string, readonly [number, number]>(
  Object.entries((table as unknown as { lines?: Record<string, readonly [number, number]> }).lines ?? {}),
);

/**
 * An East Asian Office face's GDI ascent and descent (its file's usWinAscent
 * and usWinDescent), em; undefined for a face the table does not hold. Windows
 * Excel lays a worksheet shape's line of such a face out on them (the API's
 * office_east_asian_line; xlsx backlog bl8/lnex8-9): a single line 1.3 x their
 * sum, its baseline three quarters down less (descent - 0.175 x the sum).
 */
export function officeEastAsianWinLine(fontFamily: unknown): { ascent: number; descent: number } | undefined {
  const line = LINES.get(primaryFace(fontFamily));
  return line ? { ascent: line[0], descent: line[1] } : undefined;
}

function primaryFace(fontFamily: unknown): string {
  return typeof fontFamily === "string"
    ? fontFamily.split(",")[0]!.trim().replace(/^["']|["']$/g, "").trim().toLowerCase()
    : "";
}

/** Whether the face (a CSS stack's first family) is an Office East Asian face. */
export function officeFaceIsEastAsian(fontFamily: unknown): boolean {
  const face = primaryFace(fontFamily);
  return face !== "" && FACES.has(face);
}
