/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * What a hint says under a finger.
 *
 * "Draw around the content · Esc to cancel": an iPad without a keyboard has no
 * Esc, and each of these surfaces has a control of its own that leaves it. A
 * hint's clauses are divided by " · " in every language, and every language
 * names the key, so the clause goes without a second copy of each hint.
 */
const CLAUSE_SEPARATOR = " · ";
const NAMES_ESCAPE_KEY = /(^|[^\p{L}])(Esc|Échap)(?![a-zA-Z])/u;

export function withoutEscapeClause(hint: string): string {
  const clauses = hint.split(CLAUSE_SEPARATOR);
  const kept = clauses.filter((clause) => !NAMES_ESCAPE_KEY.test(clause));
  return kept.length > 0 ? kept.join(CLAUSE_SEPARATOR) : hint;
}
