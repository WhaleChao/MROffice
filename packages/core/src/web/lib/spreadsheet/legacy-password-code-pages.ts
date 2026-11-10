/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
interface CodePage {
  /** Windows code page number (what Excel's system uses). */
  readonly codePage: number;
  /** The WHATWG Encoding Standard label that decodes it. */
  readonly label: string;
  /** Double-byte lead / trail ranges; absent for a single-byte code page. */
  readonly lead?: readonly (readonly [number, number])[];
  readonly trail?: readonly (readonly [number, number])[];
}

const CODE_PAGES: readonly CodePage[] = [
  { codePage: 1252, label: "windows-1252" },
  { codePage: 1250, label: "windows-1250" },
  { codePage: 1251, label: "windows-1251" },
  { codePage: 1253, label: "windows-1253" },
  { codePage: 1254, label: "windows-1254" },
  { codePage: 1255, label: "windows-1255" },
  { codePage: 1256, label: "windows-1256" },
  { codePage: 1257, label: "windows-1257" },
  { codePage: 1258, label: "windows-1258" },
  { codePage: 874, label: "windows-874" },
  {
    codePage: 936,
    label: "gbk",
    lead: [[0x81, 0xfe]],
    trail: [[0x40, 0x7e], [0x80, 0xfe]],
  },
  {
    codePage: 932,
    label: "shift_jis",
    lead: [[0x81, 0x9f], [0xe0, 0xfc]],
    trail: [[0x40, 0x7e], [0x80, 0xfc]],
  },
  {
    codePage: 949,
    label: "euc-kr",
    lead: [[0x81, 0xfe]],
    trail: [[0x41, 0x5a], [0x61, 0x7a], [0x81, 0xfe]],
  },
  {
    // Big5's own rows first, so a character HKSCS repeats lower down keeps
    // the code page 950 spelling.
    codePage: 950,
    label: "big5",
    lead: [[0xa1, 0xfe], [0x81, 0xa0]],
    trail: [[0x40, 0x7e], [0xa1, 0xfe]],
  },
];

type EncodeTable = ReadonlyMap<string, Uint8Array>;
const tables = new Map<number, EncodeTable | null>();

function decoderFor(label: string): TextDecoder | null {
  try {
    return new TextDecoder(label, { fatal: true });
  } catch {
    return null;
  }
}

function encodeTable(page: CodePage): EncodeTable | null {
  if (tables.has(page.codePage)) return tables.get(page.codePage) ?? null;
  const decoder = decoderFor(page.label);
  if (!decoder) {
    tables.set(page.codePage, null);
    return null;
  }
  const table = new Map<string, Uint8Array>();
  const add = (bytes: Uint8Array) => {
    let text: string;
    try {
      text = decoder.decode(bytes);
    } catch {
      return;
    }
    // One character back, and never one an earlier spelling already owns.
    if ([...text].length !== 1 || text === "�" || table.has(text)) return;
    if (text.charCodeAt(0) < 0x80) return;
    table.set(text, bytes.slice());
  };
  for (let byte = 0x80; byte <= 0xff; byte += 1) {
    add(Uint8Array.of(byte));
  }
  if (page.lead && page.trail) {
    for (const [leadFrom, leadTo] of page.lead) {
      for (let lead = leadFrom; lead <= leadTo; lead += 1) {
        for (const [trailFrom, trailTo] of page.trail) {
          for (let trail = trailFrom; trail <= trailTo; trail += 1) {
            add(Uint8Array.of(lead, trail));
          }
        }
      }
    }
  }
  tables.set(page.codePage, table);
  return table;
}

function encode(password: string, table: EncodeTable): Uint8Array | null {
  const bytes: number[] = [];
  for (const character of password) {
    const code = character.codePointAt(0) ?? 0;
    if (code > 0 && code < 0x80) {
      bytes.push(code);
      continue;
    }
    const spelled = table.get(character);
    if (!spelled) return null;
    bytes.push(...spelled);
  }
  return Uint8Array.from(bytes);
}

export interface LegacyPasswordBytes {
  readonly codePage: number | "ascii";
  readonly bytes: Uint8Array;
}

/**
 * Every byte spelling of `password` a Windows system could have hashed: one
 * for an ASCII password, one per ANSI code page that has all its characters
 * otherwise (identical spellings collapse). Empty when no code page can
 * spell it. NUL is never a password byte (the native input is C-style).
 */
export function legacyPasswordByteCandidates(
  password: string,
): LegacyPasswordBytes[] {
  if (password.includes("\0")) return [];
  if (/^[\x01-\x7F]*$/u.test(password)) {
    return [
      {
        codePage: "ascii",
        bytes: Uint8Array.from(password, (character) => character.charCodeAt(0)),
      },
    ];
  }
  const candidates: LegacyPasswordBytes[] = [];
  const seen = new Set<string>();
  for (const page of CODE_PAGES) {
    const table = encodeTable(page);
    if (!table) continue;
    const bytes = encode(password, table);
    if (!bytes) continue;
    const key = Array.from(bytes).join(",");
    if (seen.has(key)) continue;
    seen.add(key);
    candidates.push({ codePage: page.codePage, bytes });
  }
  return candidates;
}
