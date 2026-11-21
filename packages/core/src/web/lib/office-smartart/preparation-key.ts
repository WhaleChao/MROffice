/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { canonicalJson } from "../engine-host/client";

/**
 * The key a diagram's prepared conversion is kept under (`createSmartArtPreparations`):
 * its members as they are -- their words, their runs' formatting, every field --
 * in one stable spelling (`canonicalJson`, whatever the keys' order), not its
 * first member alone (TRIAGE S11l: the first member is a shape another
 * window's words do not touch, so the conversion asked for before their words
 * arrived was taken again after, drawn without their words and ready to be
 * written over them). The same member objects answer at once; members equal
 * in every field share one key.
 */
export function createSmartArtPreparationKeys<Block extends object>(): (members: readonly Block[]) => object {
  const held = new WeakMap<Block, { members: readonly Block[]; print: string; key: object }>();
  return (members) => {
    const first = members[0];
    const last = held.get(first);
    if (last && last.members.length === members.length && last.members.every((member, index) => member === members[index])) return last.key;
    const print = canonicalJson(members);
    const key = last && last.print === print ? last.key : {};
    held.set(first, { members: [...members], print, key });
    return key;
  };
}
