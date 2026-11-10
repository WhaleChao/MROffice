/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { withScriptFallbacks } from "./office-fonts/script-fallbacks";
import { inlinePresentationDeckReceipts } from "./presentation-deck-receipts";
import { apiPost } from "./api";
import { pageDocumentJsonEquals } from "./page-save-acknowledgement";
import { presentationNativeSavedContentMatches } from "./presentation-native-save";
import { presentationPageContentSchema } from "./native-page-schemas";
import {
  normalizeOfficeFontFace,
  type OfficeFontFace,
  type OfficeFontRole,
} from "../components/office-ribbon/officeFontFace";
import { officeFontSubstitution, withOfficeFontSubstitutes } from "./office-fonts/substitutes";
export * from "./presentation-font-face-reads";
import { digest, normalizePresentationFontFaces, record } from "./presentation-font-face-reads";
import type { PresentationFontAliases } from "./presentation-font-face-reads";

export function presentationFontFacesForPersist(
  value: readonly OfficeFontFace[] | undefined,
): OfficeFontFace[] | undefined {
  if (value === undefined) return undefined;
  const faces = normalizePresentationFontFaces(value);
  if (!faces) throw new Error("presentation_font_binding_invalid");
  return faces.length ? faces : undefined;
}

export function presentationFontWitnessFamily(
  family: string | undefined,
  aliases: PresentationFontAliases,
): string | undefined {
  const primary = family
    ?.split(",")[0]
    .trim()
    .replace(/^["']|["']$/g, "");
  return primary
    ? Object.keys(aliases).find((name) => aliases[name] === primary)
    : undefined;
}

export function mergePresentationFontFace(
  current: readonly OfficeFontFace[] | undefined,
  face: OfficeFontFace,
): OfficeFontFace[] {
  const result = normalizePresentationFontFaces([
    ...(current ?? []).filter(
      (old) => old.family.toLowerCase() !== face.family.toLowerCase(),
    ),
    face,
  ]);
  if (!result) throw new Error("presentation_font_binding_invalid");
  return result;
}

export type PresentationFontResolutionRequest = {
  pageId: string;
  contentVersion: number;
  sourceContent: Record<string, unknown>;
  fontFaces: OfficeFontFace[];
};
export type PresentationFontResolutionPlan = {
  content: Record<string, unknown>;
  receipt: Readonly<Record<string, unknown>>;
};
const receiptKeys = new Set([
  "schemaVersion",
  "planId",
  "pageId",
  "sourceContentVersion",
  "sourceContentSha256",
  "sourcePackageFingerprint",
  "sourceResourceId",
  "sourceResourceVersionId",
  "sourceResourceVersionKind",
  "modelSlideId",
  "sourceSlideId",
  "fontManifestSha256",
  "outputContentSha256",
]);
function topology(content: Record<string, unknown>): unknown {
  if (!Array.isArray(content.slides)) return null;
  return content.slides.map((raw) => {
    const slide = record(raw);
    return slide && Array.isArray(slide.blocks)
      ? [
          slide.id,
          slide.blocks.map((block) => [record(block)?.id, record(block)?.type]),
        ]
      : null;
  });
}
export function readPresentationFontResolutionPlan(
  value: unknown,
  request: PresentationFontResolutionRequest,
): PresentationFontResolutionPlan | null {
  const response = record(value),
    content = record(inlinePresentationDeckReceipts(response?.content)),
    receipt = record(response?.receipt);
  if (
    !response ||
    !content ||
    !receipt ||
    Object.keys(receipt).some((key) => !receiptKeys.has(key)) ||
    receipt.schemaVersion !== "pptx-native-font-resolution-v1" ||
    typeof receipt.planId !== "string" ||
    !receipt.planId ||
    receipt.planId !== response.plan_id ||
    receipt.pageId !== request.pageId ||
    receipt.sourceContentVersion !== request.contentVersion ||
    response.expected_content_version !== request.contentVersion ||
    ![
      "sourceContentSha256",
      "sourcePackageFingerprint",
      "fontManifestSha256",
      "outputContentSha256",
    ].every((key) => digest(receipt[key])) ||
    typeof receipt.sourceResourceId !== "string" ||
    !receipt.sourceResourceId ||
    !(
      receipt.sourceResourceVersionId === null ||
      (typeof receipt.sourceResourceVersionId === "string" &&
        receipt.sourceResourceVersionId.length > 0)
    ) ||
    typeof receipt.sourceResourceVersionKind !== "string" ||
    !["original", "export"].includes(receipt.sourceResourceVersionKind) ||
    (receipt.sourceResourceVersionKind === "export" &&
      !receipt.sourceResourceVersionId) ||
    !presentationPageContentSchema.safeParse(content).success ||
    !pageDocumentJsonEquals(
      topology(content),
      topology(request.sourceContent),
    ) ||
    !pageDocumentJsonEquals(
      normalizePresentationFontFaces(content.fontFaces),
      normalizePresentationFontFaces(request.fontFaces),
    )
  )
    return null;
  const slides = content.slides as Record<string, unknown>[];
  if (
    slides[0]?.id !== receipt.modelSlideId ||
    typeof receipt.sourceSlideId !== "string" ||
    !/^sld[1-9][0-9]*$/.test(receipt.sourceSlideId) ||
    record(slides[0]?.source)?.slideId !== receipt.sourceSlideId
  )
    return null;
  const pending = slides.flatMap((slide) =>
    Object.hasOwn(record(slide.source) ?? {}, "nativeFontResolutions")
      ? [{ id: slide.id, values: record(slide.source)?.nativeFontResolutions }]
      : [],
  );
  if (
    pending.length !== 1 ||
    pending[0].id !== receipt.modelSlideId ||
    !Array.isArray(pending[0].values) ||
    pending[0].values.length !== 1 ||
    !pageDocumentJsonEquals(pending[0].values[0], receipt)
  )
    return null;
  return { content, receipt };
}

/** A successful exact-draft transport ACK may update native-derived slide data.
 * Keep the font manifest, deck fields and authored object topology unchanged. */
export function readPresentationFontCanonicalSave(
  draft: Record<string, unknown>, value: unknown,
): Record<string, unknown> | null {
  const content = record(value), faces = normalizePresentationFontFaces(draft.fontFaces);
  if (!content || !faces?.length || !presentationPageContentSchema.safeParse(content).success ||
    !pageDocumentJsonEquals(normalizePresentationFontFaces(content.fontFaces), faces) ||
    !pageDocumentJsonEquals(topology(content), topology(draft))) return null;
  const withoutSlides = (source: Record<string, unknown>) =>
    Object.fromEntries(Object.entries(source).filter(([key]) => key !== "slides"));
  return pageDocumentJsonEquals(withoutSlides(content), withoutSlides(draft)) ? content : null;
}

export async function requestPresentationFontResolutionPlan(
  request: PresentationFontResolutionRequest,
  signal: AbortSignal,
): Promise<PresentationFontResolutionPlan> {
  const response = await apiPost<unknown>(
    `/api/v1/pages/${encodeURIComponent(request.pageId)}/native-pptx-font-resolution-plan`,
    {
      expected_content_version: request.contentVersion,
      font_faces: request.fontFaces,
    },
    { signal },
  );
  const plan = readPresentationFontResolutionPlan(response, request);
  if (!plan) throw new Error("presentation_font_plan_invalid");
  return plan;
}

/** Only consume the exact server-created pending receipt, never all source metadata. */
export function presentationFontResolutionSavedMatches(
  plan: PresentationFontResolutionPlan,
  saved: unknown,
): boolean {
  const value = record(saved);
  if (
    !value ||
    !Array.isArray(value.slides) ||
    value.slides.some((slide) =>
      Object.hasOwn(
        record(record(slide)?.source) ?? {},
        "nativeFontResolutions",
      ),
    )
  )
    return false;
  const expected = structuredClone(plan.content);
  const carrier = (expected.slides as Record<string, unknown>[]).find(
    (slide) => slide.id === plan.receipt.modelSlideId,
  );
  const source = record(carrier?.source),
    receipts = source?.nativeFontResolutions;
  if (
    !source ||
    !Array.isArray(receipts) ||
    receipts.length !== 1 ||
    !pageDocumentJsonEquals(receipts[0], plan.receipt)
  )
    return false;
  delete source.nativeFontResolutions;
  return presentationNativeSavedContentMatches(expected, value);
}


/** Same exact transport-snapshot refresh used by native-copy preflight. */
export function presentationFontRefreshedSnapshot(refresh: {
  generation: number; before: Record<string, unknown>; after: Record<string, unknown>; saved: unknown;
} | null, generation: number, snapshot: Record<string, unknown>, current: Record<string, unknown>, saved: unknown): Record<string, unknown> | undefined {
  return refresh?.generation === generation && pageDocumentJsonEquals(refresh.before, snapshot) &&
    pageDocumentJsonEquals(refresh.after, current) && pageDocumentJsonEquals(refresh.saved, saved)
    ? structuredClone(refresh.after) : undefined;
}
