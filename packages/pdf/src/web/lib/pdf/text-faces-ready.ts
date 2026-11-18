/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/** How long the faces still loading after a refused load are waited for. */
export const PDF_TEXT_FACES_SETTLE_MS = 5_000;

export type PdfTextFaceSet = Pick<FontFaceSet, "load" | "ready">;

/**
 * Ask for the faces of each `[font, text]` (a CSS font shorthand and the
 * characters drawn in it). Never rejects. Answers how many of the loads were
 * refused: 0 when every face asked for loaded.
 */
export async function pdfTextFacesReady(
  fonts: PdfTextFaceSet | null | undefined,
  samples: readonly (readonly [font: string, text: string])[],
  settleMs: number = PDF_TEXT_FACES_SETTLE_MS,
): Promise<number> {
  if (!fonts || typeof fonts.load !== "function" || samples.length === 0) return 0;
  const answers = await Promise.all(samples.map(async ([font, text]) => {
    try {
      await fonts.load(font, text);
      return true;
    } catch {
      return false;
    }
  }));
  const refused = answers.filter((loaded) => !loaded).length;
  if (refused === 0) return 0;
  let timer: ReturnType<typeof setTimeout> | undefined;
  try {
    await Promise.race([
      Promise.resolve(fonts.ready).catch(() => undefined),
      new Promise<void>((resolve) => { timer = setTimeout(resolve, Math.max(0, settleMs)); }),
    ]);
  } catch {
    // A set without `ready`: the loads above are all the wait there is.
  } finally {
    if (timer !== undefined) clearTimeout(timer);
  }
  return refused;
}
