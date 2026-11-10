/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The characters an equation draws, asked of the server before the editor
 * puts the equation on a slide (`POST /api/v1/presentation/equation-text`).
 *
 * The page's equation run must carry exactly the characters the server's
 * Office Math writer draws, or the next save finds the shape's text stale; only
 * the server knows what its writer draws, so the editor never computes them
 * itself. A refusal or a network failure inserts nothing.
 */

import { apiPost, isApiRequestError } from "@mrnote/core/web/lib/api.ts";
import { presentationEquationTextIsUsable, type PresentationEquation } from "../presentation-equation-box";

export type PresentationEquationTextAnswer =
  | {
      readonly ok: true;
      readonly equation: PresentationEquation;
      /** False: the file keeps it in Office's linear format (still an equation). */
      readonly native: boolean;
    }
  | { readonly ok: false; readonly reason: "unreadable" | "unavailable" };

export const PRESENTATION_EQUATION_TEXT_PATH = "/api/v1/presentation/equation-text";

export async function fetchPresentationEquationText(
  latex: string,
  display: boolean,
  post: typeof apiPost = apiPost,
): Promise<PresentationEquationTextAnswer> {
  const trimmed = latex.trim();
  if (!trimmed) return { ok: false, reason: "unreadable" };
  try {
    const answer = await post<{ text?: unknown; latex?: unknown; native?: unknown }>(
      PRESENTATION_EQUATION_TEXT_PATH,
      { latex: trimmed, display },
    );
    const text = typeof answer?.text === "string" ? answer.text : "";
    const answeredLatex = typeof answer?.latex === "string" && answer.latex.trim() ? answer.latex : trimmed;
    if (!presentationEquationTextIsUsable(text)) return { ok: false, reason: "unreadable" };
    return {
      ok: true,
      equation: { text, latex: answeredLatex, display },
      native: answer?.native !== false,
    };
  } catch (error) {
    if (isApiRequestError(error) && error.status === 422) return { ok: false, reason: "unreadable" };
    return { ok: false, reason: "unavailable" };
  }
}
