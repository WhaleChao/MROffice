/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The outlines of a CFF table (Adobe TN 5176, Type 2 charstrings TN 5177):
 * each glyph's path in font units, moves, lines and cubics. Name-keyed and
 * CID-keyed fonts (an FDArray of private dictionaries chosen per glyph by
 * FDSelect), local and global subroutines, hint operators skipped. What a
 * font from the reader's computer is drawn with when it is CFF (PingFang),
 * read here because fontkit cannot read its font dictionaries.
 */

import { sfntFail } from "./sfnt-reader";

export type CffCommand =
  | Readonly<{ kind: "M"; x: number; y: number }>
  | Readonly<{ kind: "L"; x: number; y: number }>
  | Readonly<{ kind: "C"; x1: number; y1: number; x2: number; y2: number; x: number; y: number }>
  | Readonly<{ kind: "Z" }>;

export interface CffOutlines {
  readonly glyphCount: number;
  path(glyphId: number): readonly CffCommand[];
}

interface CffIndex {
  readonly count: number;
  item(index: number): Uint8Array;
  /** Where the bytes after the index start. */
  readonly end: number;
}

function readIndex(data: Uint8Array, at: number): CffIndex {
  if (at + 2 > data.length) sfntFail("invalid-font", "A CFF index is outside the table.");
  const count = (data[at] << 8) | data[at + 1];
  if (count === 0) return { count: 0, item: () => sfntFail("invalid-font"), end: at + 2 };
  const offSize = data[at + 2];
  if (offSize < 1 || offSize > 4) sfntFail("invalid-font", "A CFF index has an invalid offset size.");
  const offsetAt = (index: number) => {
    let value = 0;
    const start = at + 3 + index * offSize;
    for (let byte = 0; byte < offSize; byte += 1) value = value * 256 + data[start + byte];
    return value;
  };
  const base = at + 3 + (count + 1) * offSize - 1;
  const end = base + offsetAt(count);
  if (end > data.length) sfntFail("invalid-font", "A CFF index is outside the table.");
  return {
    count,
    end,
    item(index: number) {
      if (index < 0 || index >= count) sfntFail("invalid-font", "A CFF index item is missing.");
      return data.subarray(base + offsetAt(index), base + offsetAt(index + 1));
    },
  };
}

/** A DICT's operators and their operands (escaped operators as 1200 + b). */
function readDict(data: Uint8Array): Map<number, number[]> {
  const dict = new Map<number, number[]>();
  let operands: number[] = [];
  let at = 0;
  while (at < data.length) {
    const b0 = data[at];
    if (b0 <= 21) {
      let operator = b0;
      at += 1;
      if (b0 === 12) { operator = 1200 + data[at]; at += 1; }
      dict.set(operator, operands);
      operands = [];
    } else if (b0 === 28) {
      operands.push(((data[at + 1] << 24) | (data[at + 2] << 16)) >> 16);
      at += 3;
    } else if (b0 === 29) {
      operands.push((data[at + 1] << 24) | (data[at + 2] << 16) | (data[at + 3] << 8) | data[at + 4]);
      at += 5;
    } else if (b0 === 30) {
      let text = "";
      at += 1;
      for (let done = false; !done && at < data.length; at += 1) {
        for (const nibble of [data[at] >> 4, data[at] & 15]) {
          if (nibble === 15) { done = true; break; }
          text += nibble <= 9 ? String(nibble) : ["", "", "", "", "", "", "", "", "", "", ".", "E", "E-", "", "-"][nibble];
        }
      }
      operands.push(Number.parseFloat(text) || 0);
    } else if (b0 >= 32 && b0 <= 246) {
      operands.push(b0 - 139);
      at += 1;
    } else if (b0 >= 247 && b0 <= 250) {
      operands.push((b0 - 247) * 256 + data[at + 1] + 108);
      at += 2;
    } else if (b0 >= 251 && b0 <= 254) {
      operands.push(-(b0 - 251) * 256 - data[at + 1] - 108);
      at += 2;
    } else {
      sfntFail("invalid-font", "A CFF dictionary holds an invalid operand.");
    }
  }
  return dict;
}

function subrBias(count: number): number {
  return count < 1240 ? 107 : count < 33900 ? 1131 : 32768;
}

interface PrivateData {
  readonly subrs: CffIndex | null;
}

function readPrivate(table: Uint8Array, entry: number[] | undefined): PrivateData {
  if (!entry || entry.length < 2) return { subrs: null };
  const [size, offset] = entry;
  if (offset < 0 || offset + size > table.length) sfntFail("invalid-font", "A CFF private dictionary is outside the table.");
  const dict = readDict(table.subarray(offset, offset + size));
  const subrs = dict.get(19);
  return { subrs: subrs && subrs.length ? readIndex(table, offset + subrs[0]) : null };
}

export function readCffOutlines(table: Uint8Array): CffOutlines {
  if (table.length < 4 || table[0] !== 1) sfntFail("unsupported-outlines", "Only CFF version 1 outlines are read.");
  const names = readIndex(table, table[2]);
  const topDicts = readIndex(table, names.end);
  const strings = readIndex(table, topDicts.end);
  const globalSubrs = readIndex(table, strings.end);
  if (topDicts.count < 1) sfntFail("invalid-font", "The CFF table has no font.");
  const top = readDict(topDicts.item(0));
  if ((top.get(1206)?.[0] ?? 2) !== 2) sfntFail("unsupported-outlines", "Only Type 2 charstrings are read.");
  const charStringsAt = top.get(17)?.[0];
  if (charStringsAt === undefined) sfntFail("invalid-font", "The CFF font has no charstrings.");
  const charStrings = readIndex(table, charStringsAt);
  // A CID-keyed font chooses each glyph's private dictionary (FDSelect).
  const fdArrayAt = top.get(1236)?.[0];
  const fdSelectAt = top.get(1237)?.[0];
  let privates: PrivateData[];
  let fdOf: (glyphId: number) => number = () => 0;
  if (fdArrayAt !== undefined && fdSelectAt !== undefined) {
    const fdArray = readIndex(table, fdArrayAt);
    privates = Array.from({ length: fdArray.count }, (_, index) => readPrivate(table, readDict(fdArray.item(index)).get(18)));
    const format = table[fdSelectAt];
    if (format === 0) {
      fdOf = (glyphId) => table[fdSelectAt + 1 + glyphId];
    } else if (format === 3) {
      const ranges = (table[fdSelectAt + 1] << 8) | table[fdSelectAt + 2];
      const firsts: number[] = [], fds: number[] = [];
      for (let index = 0; index < ranges; index += 1) {
        const at = fdSelectAt + 3 + index * 3;
        firsts.push((table[at] << 8) | table[at + 1]);
        fds.push(table[at + 2]);
      }
      fdOf = (glyphId) => {
        let low = 0, high = firsts.length - 1;
        while (low < high) {
          const middle = (low + high + 1) >> 1;
          if (firsts[middle] <= glyphId) low = middle; else high = middle - 1;
        }
        return fds[low];
      };
    } else {
      sfntFail("unsupported-outlines", "This FDSelect format is not read.");
    }
  } else {
    privates = [readPrivate(table, top.get(18))];
  }
  return {
    glyphCount: charStrings.count,
    path(glyphId: number) {
      if (glyphId < 0 || glyphId >= charStrings.count) sfntFail("missing-glyph");
      const local = privates[fdOf(glyphId)]?.subrs ?? null;
      return runCharString(charStrings.item(glyphId), local, globalSubrs);
    },
  };
}

/** A Type 2 charstring run: the path it draws (hints and widths read past). */
function runCharString(code: Uint8Array, localSubrs: CffIndex | null, globalSubrs: CffIndex): CffCommand[] {
  const commands: CffCommand[] = [];
  const stack: number[] = [];
  let x = 0, y = 0, stems = 0, depth = 0, open = false, sawWidth = false, ended = false;
  const localBias = localSubrs ? subrBias(localSubrs.count) : 0;
  const globalBias = subrBias(globalSubrs.count);
  const close = () => { if (open) commands.push({ kind: "Z" }); open = false; };
  const moveTo = (dx: number, dy: number) => {
    close();
    x += dx; y += dy;
    commands.push({ kind: "M", x, y });
    open = true;
  };
  const lineTo = (dx: number, dy: number) => { x += dx; y += dy; commands.push({ kind: "L", x, y }); };
  const curveTo = (dx1: number, dy1: number, dx2: number, dy2: number, dx3: number, dy3: number) => {
    const x1 = x + dx1, y1 = y + dy1, x2 = x1 + dx2, y2 = y1 + dy2;
    x = x2 + dx3; y = y2 + dy3;
    commands.push({ kind: "C", x1, y1, x2, y2, x, y });
  };
  // The first stack-clearing operator may carry the glyph's width first.
  const width = (even: boolean) => {
    if (!sawWidth && (stack.length % 2 === 1) === even) stack.shift();
    sawWidth = true;
  };
  const run = (bytes: Uint8Array) => {
    if (++depth > 10) sfntFail("invalid-font", "CFF subroutines nest too deep.");
    let at = 0;
    while (at < bytes.length && !ended) {
      const b0 = bytes[at];
      if (b0 >= 32 || b0 === 28) {
        if (b0 === 28) { stack.push(((bytes[at + 1] << 24) | (bytes[at + 2] << 16)) >> 16); at += 3; }
        else if (b0 <= 246) { stack.push(b0 - 139); at += 1; }
        else if (b0 <= 250) { stack.push((b0 - 247) * 256 + bytes[at + 1] + 108); at += 2; }
        else if (b0 <= 254) { stack.push(-(b0 - 251) * 256 - bytes[at + 1] - 108); at += 2; }
        else { stack.push(((bytes[at + 1] << 24) | (bytes[at + 2] << 16) | (bytes[at + 3] << 8) | bytes[at + 4]) / 65536); at += 5; }
        continue;
      }
      at += 1;
      switch (b0) {
        case 1: case 3: case 18: case 23: // hstem vstem hstemhm vstemhm
          width(true);
          stems += stack.length >> 1;
          stack.length = 0;
          break;
        case 19: case 20: // hintmask cntrmask
          width(true);
          stems += stack.length >> 1;
          stack.length = 0;
          at += (stems + 7) >> 3;
          break;
        case 21: // rmoveto
          width(true);
          moveTo(stack[0] ?? 0, stack[1] ?? 0);
          stack.length = 0;
          break;
        case 22: // hmoveto
          width(false);
          moveTo(stack[0] ?? 0, 0);
          stack.length = 0;
          break;
        case 4: // vmoveto
          width(false);
          moveTo(0, stack[0] ?? 0);
          stack.length = 0;
          break;
        case 5: // rlineto
          for (let index = 0; index + 1 < stack.length; index += 2) lineTo(stack[index], stack[index + 1]);
          stack.length = 0;
          break;
        case 6: case 7: { // hlineto vlineto
          let horizontal = b0 === 6;
          for (const value of stack) { if (horizontal) lineTo(value, 0); else lineTo(0, value); horizontal = !horizontal; }
          stack.length = 0;
          break;
        }
        case 8: // rrcurveto
          for (let index = 0; index + 5 < stack.length; index += 6) {
            curveTo(stack[index], stack[index + 1], stack[index + 2], stack[index + 3], stack[index + 4], stack[index + 5]);
          }
          stack.length = 0;
          break;
        case 24: { // rcurveline
          let index = 0;
          for (; index + 7 < stack.length; index += 6) {
            curveTo(stack[index], stack[index + 1], stack[index + 2], stack[index + 3], stack[index + 4], stack[index + 5]);
          }
          lineTo(stack[index], stack[index + 1]);
          stack.length = 0;
          break;
        }
        case 25: { // rlinecurve
          let index = 0;
          for (; index + 7 < stack.length; index += 2) lineTo(stack[index], stack[index + 1]);
          curveTo(stack[index], stack[index + 1], stack[index + 2], stack[index + 3], stack[index + 4], stack[index + 5]);
          stack.length = 0;
          break;
        }
        case 26: { // vvcurveto
          let index = 0, dx1 = 0;
          if (stack.length % 2 === 1) { dx1 = stack[0]; index = 1; }
          for (; index + 3 < stack.length; index += 4) {
            curveTo(dx1, stack[index], stack[index + 1], stack[index + 2], 0, stack[index + 3]);
            dx1 = 0;
          }
          stack.length = 0;
          break;
        }
        case 27: { // hhcurveto
          let index = 0, dy1 = 0;
          if (stack.length % 2 === 1) { dy1 = stack[0]; index = 1; }
          for (; index + 3 < stack.length; index += 4) {
            curveTo(stack[index], dy1, stack[index + 1], stack[index + 2], stack[index + 3], 0);
            dy1 = 0;
          }
          stack.length = 0;
          break;
        }
        case 30: case 31: { // vhcurveto hvcurveto
          let horizontal = b0 === 31;
          for (let index = 0; index + 3 < stack.length; index += 4) {
            const last = index + 4 >= stack.length - 1 && stack.length - index === 5 ? stack[index + 4] : 0;
            if (horizontal) curveTo(stack[index], 0, stack[index + 1], stack[index + 2], last, stack[index + 3]);
            else curveTo(0, stack[index], stack[index + 1], stack[index + 2], stack[index + 3], last);
            horizontal = !horizontal;
          }
          stack.length = 0;
          break;
        }
        case 10: { // callsubr
          if (!localSubrs) sfntFail("invalid-font", "A CFF charstring calls a missing subroutine.");
          const index = stack.pop()! + localBias;
          run(localSubrs.item(index));
          break;
        }
        case 29: { // callgsubr
          const index = stack.pop()! + globalBias;
          run(globalSubrs.item(index));
          break;
        }
        case 11: // return
          depth -= 1;
          return;
        case 14: // endchar
          width(true);
          close();
          ended = true;
          break;
        case 12: {
          const b1 = bytes[at];
          at += 1;
          if (b1 === 35) { // flex
            curveTo(stack[0], stack[1], stack[2], stack[3], stack[4], stack[5]);
            curveTo(stack[6], stack[7], stack[8], stack[9], stack[10], stack[11]);
          } else if (b1 === 34) { // hflex
            const startY = y;
            curveTo(stack[0], 0, stack[1], stack[2], stack[3], 0);
            curveTo(stack[4], 0, stack[5], startY - y, stack[6], 0);
          } else if (b1 === 36) { // hflex1
            const startY = y;
            curveTo(stack[0], stack[1], stack[2], stack[3], stack[4], 0);
            curveTo(stack[5], 0, stack[6], stack[7], stack[8], startY - (y + stack[7]));
          } else if (b1 === 37) { // flex1
            const startX = x, startY = y;
            let dx = 0, dy = 0;
            for (let index = 0; index < 10; index += 2) { dx += stack[index]; dy += stack[index + 1]; }
            curveTo(stack[0], stack[1], stack[2], stack[3], stack[4], stack[5]);
            const last = stack[10];
            if (Math.abs(dx) > Math.abs(dy)) curveTo(stack[6], stack[7], stack[8], stack[9], last, startY - (y + stack[7] + stack[9]));
            else curveTo(stack[6], stack[7], stack[8], stack[9], startX - (x + stack[6] + stack[8]), last);
          } else {
            sfntFail("unsupported-outlines", "This CFF charstring operator is not read.");
          }
          stack.length = 0;
          break;
        }
        default:
          sfntFail("unsupported-outlines", "This CFF charstring operator is not read.");
      }
    }
    depth -= 1;
  };
  run(code);
  close();
  return commands;
}
