/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
export const PDF_SOURCE_NAME = String.raw`(?:[^\x00\t\n\f\r /\[\]<>(){}%#]|#[0-9A-Fa-f]{2}){1,127}`;

/** One whole source name with its slash, as the content lexer cuts it. */
export const PDF_SOURCE_NAME_TOKEN = new RegExp(String.raw`^/${PDF_SOURCE_NAME}$`, "u");

/** A source name's value: its `#xx` escapes read as the bytes they spell
 * (Latin-1, one character per byte, as every content walk reads a stream).
 * A leading slash is dropped. A `#` not followed by two hexadecimal digits
 * is kept as it stands; the patterns above never match one. Mirrors
 * replay's `_decode_source_name`. */
export function decodePdfSourceName(name: string): string {
  const value = name.startsWith("/") ? name.slice(1) : name;
  return value.includes("#")
    ? value.replace(/#([0-9A-Fa-f]{2})/gu, (_escape, hex: string) => String.fromCharCode(Number.parseInt(hex, 16)))
    : value;
}

/** A marked-content property list (or any operand text) with each name read
 * as its value, for the tests that look a key up in it -- `/ActualText`,
 * `/MCID n`: `/Actual#54ext` is the same key. A name whose value holds white
 * space or a delimiter keeps its spelling, so the view cuts into the same
 * tokens as the source. Only ever read, never written. Mirrors replay's
 * `_source_names_read`. */
export function pdfSourceNamesRead(text: string): string {
  return text.includes("#")
    ? text.replace(/\/[^\x00\t\n\f\r /[\]<>(){}%]+/gu, (name) => {
        const value = decodePdfSourceName(name);
        return /[\x00\t\n\f\r /[\]<>(){}%]/u.test(value) ? name : `/${value}`;
      })
    : text;
}
