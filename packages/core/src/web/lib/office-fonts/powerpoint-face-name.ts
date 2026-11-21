/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { windowsOfficeHas } from "./windows-office-faces";

























export function powerPointFaceName(name: string): string {
  const trimmed = name.trim();
  const comma = trimmed.indexOf(",");
  const paren = trimmed.indexOf("(");
  const cut = comma < 0 ? paren : paren < 0 ? comma : Math.min(comma, paren);
  if (cut < 0) return trimmed;
  const head = trimmed.slice(0, cut).trim();
  if (!head) return trimmed.replace(/[(),]/g, " ").split(/\s+/).filter(Boolean).join(" ");
  return regularOfFullFace(head);
}

/** A full face name ("Georgia Bold", "Arial Bold Italic") of a family Windows has: that family. */
const FULL_FACE_STYLE = /\s+(?:bold\s+italic|bold|italic)$/i;
function regularOfFullFace(head: string): string {
  const match = FULL_FACE_STYLE.exec(head);
  if (!match) return head;
  const family = head.slice(0, match.index).trim();
  return family && windowsOfficeHas(family) ? family : head;
}
