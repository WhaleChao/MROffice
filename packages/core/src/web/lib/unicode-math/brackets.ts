/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Bracket pairing for brackets written as ordinary run text. Word builds a
 * matched pair typed in linear text into a delimiter object, so the parser,
 * the serializer and the equivalence check all pair them the same way.
 */

import type { MathList } from "./ast";
import { CLOSERS, OPENERS, SELF_DELIMITERS } from "./characters";

export interface BracketPosition {
  readonly node: number;
  /** Code-point index within the run text. */
  readonly index: number;
}

export interface BracketPair {
  readonly open: BracketPosition;
  readonly close: BracketPosition;
}

/** Pairs of matched bracket characters in a container's plain runs. */
export function bracketPairs(list: MathList): BracketPair[] {
  const pairs: BracketPair[] = [];
  const stack: { character: string; at: BracketPosition; self: boolean; content: boolean }[] = [];
  const touch = () => {
    const top = stack.at(-1);
    if (top) top.content = true;
  };
  list.forEach((node, nodeIndex) => {
    if (node.kind !== "text" || node.style.nor || node.style.lit) {
      touch();
      return;
    }
    Array.from(node.text).forEach((character, index) => {
      const at = { node: nodeIndex, index };
      if (OPENERS.has(character)) {
        touch();
        stack.push({ character, at, self: false, content: false });
        return;
      }
      if (CLOSERS.has(character)) {
        while (stack.length && stack[stack.length - 1].self) stack.pop();
        const open = stack.pop();
        if (open) {
          pairs.push({ open: open.at, close: at });
          touch();
        }
        return;
      }
      if (SELF_DELIMITERS.has(character)) {
        const top = stack.at(-1);
        if (top && top.self && top.character === character && top.content) {
          stack.pop();
          pairs.push({ open: top.at, close: at });
          touch();
          return;
        }
        touch();
        stack.push({ character, at, self: true, content: false });
        return;
      }
      touch();
    });
  });
  return pairs;
}

export function pairedPositions(list: MathList): Set<string> {
  const positions = new Set<string>();
  for (const pair of bracketPairs(list)) {
    positions.add(`${pair.open.node}:${pair.open.index}`);
    positions.add(`${pair.close.node}:${pair.close.index}`);
  }
  return positions;
}
