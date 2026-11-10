/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
// The family a face's full or PostScript name belongs to.
//
// A workbook (or a document) can name one face of a family rather than the
// family: "Lucida Sans Regular", "Arial Bold", "ArialMT". Excel for Mac opens
// such a name as that face (P28_DateFormatTests draws "Lucida Sans Regular"

// name would fall through to the fallback face. face-names.generated.json
// (scripts/sync-office-fonts.mjs, from local-names.json) lists every Office
// for Mac face name that is not itself a family. Kept apart from
// substitutes.ts so pages that never meet such a name do not carry the table.
import table from "./face-names.generated.json";

export interface OfficeFaceName {
  readonly family: string;
  readonly weight: number;
  readonly italic: boolean;
}

// The JSON module's rows type as (string | number)[], not as the tuple each
// row is; the table is the generator's, so the tuple type is asserted.
const faces = table as unknown as Record<string, readonly [string, number, 0 | 1]>;

/** The family, weight and style a face name stands for, or undefined when
 * the name is a family (or unknown). */
export function officeFaceName(name: string | null | undefined): OfficeFaceName | undefined {
  const bare = name ? name.trim().replace(/^["']|["']$/g, "").trim() : "";
  const entry = bare ? faces[bare.toLowerCase()] : undefined;
  if (entry) return { family: entry[0], weight: entry[1], italic: entry[2] === 1 };
  // Windows names a family's character set after it ("Arial CE", "Times
  // New Roman CYR", "Arial (Hebrew)"; the FontSubstitutes registry maps them
  // to the family), and Excel for Mac draws them as the family: "Arial CE"
  
  // PH072_issue.2301, X-258). The traits stay the cell's.
  const charset = WINDOWS_CHARSET_SUFFIX.exec(bare);
  return charset ? { family: charset[1], weight: 400, italic: false } : undefined;
}

const WINDOWS_CHARSET_SUFFIX =
  /^(.+?)\s+(?:CE|CYR|Cyr|Baltic|Greek|TUR|Tur|\((?:Hebrew|Arabic|Baltic|Central Europe|Cyrillic|Greek|Turkish|Vietnamese)\))$/;
