/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import table from "./office-east-asian-faces.generated.json";

const FACES = new Map<string, number | null>(
  Object.entries((table as { faces: Record<string, number | null> }).faces),
);

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
