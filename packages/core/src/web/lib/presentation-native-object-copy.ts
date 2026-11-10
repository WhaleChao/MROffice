/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { inlinePresentationDeckReceipts } from "./presentation-deck-receipts";
import { apiPost } from "./api";
import { presentationPageContentSchema } from "./native-page-schemas";
import { pageDocumentJsonEquals } from "./page-save-acknowledgement";
import { presentationNativeSavedContentMatches } from "./presentation-native-save";

type RecordValue = Record<string, unknown>;
const record = (value: unknown): RecordValue | null =>
  value !== null && typeof value === "object" && !Array.isArray(value)
    ? value as RecordValue
    : null;

export function hasNativePresentationObject(block: unknown): boolean {
  const source = record(record(block)?.source);
  const view = record(source?.nativeSmartArtView);
  return Boolean(source && (
    typeof source.shapeId === "string" && /^shp[1-9]\d*$/.test(source.shapeId)
    || view?.schemaVersion === "pptx-smartart-drawing-v1"
  ));
}

export type PresentationNativeCopyRequest = {
  pageId: string;
  contentVersion: number;
  slideId: string;
  blockIds: string[];
  offsetEmu: number;
  /** A different vertical offset (Duplicate's smart offset); `offsetEmu`
   *  when omitted. */
  offsetYEmu?: number;
  /** Paste onto another slide of the page: the copies land there. Omitted
   *  (or `slideId`), they land on the slide they come from. */
  targetSlideId?: string;
};

export type PresentationNativeCopyPlan = {
  content: RecordValue;
  insertedBlockIds: string[];
  blockIdMap: Readonly<Record<string, string>>;
  slideId: string;
};

export function readPresentationNativeCopyPlan(
  value: unknown,
  request: PresentationNativeCopyRequest,
): PresentationNativeCopyPlan | null {
  const response = record(value);
  const receipt = record(response?.receipt);
  const content = record(inlinePresentationDeckReceipts(response?.content));
  const ids = response?.inserted_block_ids;
  const selected = receipt?.sourceBlockIds;
  const mapping = record(receipt?.modelBlockIdMap);
  if (!response || !receipt || !content || !mapping
    || response.expected_content_version !== request.contentVersion
    || response.plan_id !== receipt.planId
    || receipt.schemaVersion !== "pptx-native-object-copy-v1"
    || receipt.pageId !== request.pageId
    || receipt.sourceContentVersion !== request.contentVersion
    || receipt.modelSlideId !== request.slideId
    || (crossSlide(request)
      ? receipt.targetModelSlideId !== request.targetSlideId || response.slide_id !== request.targetSlideId
      : receipt.targetModelSlideId !== undefined && receipt.targetModelSlideId !== request.slideId)
    || !Array.isArray(ids) || ids.length === 0
    || ids.some((id) => typeof id !== "string" || !id)
    || new Set(ids).size !== ids.length
    || !Array.isArray(selected) || selected.length !== request.blockIds.length
    || new Set(selected).size !== selected.length
    || selected.some((id) => !request.blockIds.includes(id))
    || Object.values(mapping).length !== ids.length
    || new Set(Object.values(mapping)).size !== ids.length
    || Object.keys(mapping).some((id) => !id || ids.includes(id))
    || selected.some((id) => typeof mapping[id] !== "string")
    || Object.values(mapping).some((id) => !ids.includes(id))) return null;
  const validated = presentationPageContentSchema.safeParse(content);
  if (!validated.success) return null;
  const landedOn = crossSlide(request) ? request.targetSlideId! : request.slideId;
  const slide = validated.data.slides.find((item) => item.id === landedOn);
  if (!slide || ids.some((id) => !slide.blocks.some((block) => block.id === id))) return null;
  const embedded = record(slide.source)?.nativeObjectCopies;
  if (!Array.isArray(embedded) || embedded.length !== 1
    || JSON.stringify(embedded[0]) !== JSON.stringify(receipt)) return null;
  // Validate without rewriting the server snapshot. Normalized renderer values
  // must not replace the version-bound native provenance in the pending plan.
  return { content, insertedBlockIds: ids as string[],
    blockIdMap: { ...mapping } as Record<string, string>, slideId: landedOn };
}

function crossSlide(request: PresentationNativeCopyRequest): boolean {
  return typeof request.targetSlideId === "string" && request.targetSlideId !== request.slideId;
}

/** The server cannot copy these objects to another slide (a member of a
 *  turned group, a server that does not know the target): Paste makes its
 *  redrawn copy instead. */
export function isPresentationNativeCopyUnsupported(error: unknown): boolean {
  const record = error !== null && typeof error === "object" ? error as { status?: unknown; code?: unknown } : null;
  return record?.status === 422 && record.code === "native_copy_unsupported";
}

export type PresentationNativeCopyPreviewResult =
  | { ok: true; plan: PresentationNativeCopyPlan }
  | { ok: false; reason: "missing-member" | "changed-native-structure" | "invalid-geometry" | "unknown-reference" | "invalid-content" };

/** Project private preview edits onto the server's independent native copies.
 * This does not save the page or establish native-writer support. The caller
 * must publish through normal Save and report success only after acknowledgement.
 */
export function projectPresentationNativeCopyPreview(
  plan: PresentationNativeCopyPlan,
  sourceContent: RecordValue,
  previewContent: RecordValue,
  targetSlideId?: string,
): PresentationNativeCopyPreviewResult {
  const slideOf = (content: RecordValue, id = plan.slideId) => Array.isArray(content.slides)
    ? content.slides.map(record).find((slide) => slide?.id === id)
    : null;
  const rowsOf = (slide: RecordValue | null | undefined): Map<string, RecordValue> | null => {
    if (!slide || !Array.isArray(slide.blocks)) return null;
    const result = new Map<string, RecordValue>();
    for (const value of slide.blocks) {
      const block = record(value);
      if (!block || typeof block.id !== "string" || result.has(block.id)) return null;
      result.set(block.id, block);
    }
    return result;
  };
  const baseline = rowsOf(slideOf(sourceContent));
  const preview = rowsOf(slideOf(previewContent));
  if (!baseline || !preview) return { ok: false, reason: "missing-member" };
  if (targetSlideId !== undefined) {
    // Only a verified whole-slide plan may target another slide. An arbitrary
    // existing destination must never turn preview projection into a writer.
    const pending = record(slideOf(plan.content)?.source)?.nativeSlideCopies;
    const receipt = Array.isArray(pending) && pending.length === 1 ? record(pending[0]) : null;
    const target = slideOf(plan.content, targetSlideId);
    if (preview.size !== baseline.size || [...baseline.keys()].some(id => !preview.has(id)) ||
      !receipt || receipt.schemaVersion !== "pptx-native-slide-copy-v1" ||
      receipt.modelSlideId !== plan.slideId || receipt.targetModelSlideId !== targetSlideId ||
      targetSlideId === plan.slideId || slideOf(sourceContent, targetSlideId) ||
      !target || record(target.source)?.slideId !== receipt.targetSlideId ||
      record(slideOf(sourceContent)?.source)?.slideId !== receipt.sourceSlideId ||
      receipt.targetSlideId === receipt.sourceSlideId ||
      !pageDocumentJsonEquals(receipt.modelBlockIdMap, plan.blockIdMap) ||
      Object.keys(plan.blockIdMap).length !== baseline.size ||
      [...baseline.keys()].some(id => !Object.hasOwn(plan.blockIdMap, id))) {
      return { ok: false, reason: "unknown-reference" };
    }
  }
  const content = JSON.parse(JSON.stringify(plan.content)) as RecordValue;
  const copies = rowsOf(slideOf(content, targetSlideId));
  if (!copies) return { ok: false, reason: "missing-member" };
  if (targetSlideId !== undefined && copies.size !== baseline.size)
    return { ok: false, reason: "changed-native-structure" };
  const targets = Object.values(plan.blockIdMap);
  if (targets.length === 0 || new Set(targets).size !== targets.length
    || targets.length !== plan.insertedBlockIds.length
    || targets.some((id) => baseline.has(id) || !plan.insertedBlockIds.includes(id))) {
    return { ok: false, reason: "changed-native-structure" };
  }
  const protectedKeys = new Set(["id", "type", "source", "sourceOrder", "zIndex", "groupId"]);
  for (const [sourceId, copyId] of Object.entries(plan.blockIdMap)) {
    const original = baseline.get(sourceId);
    const draft = preview.get(sourceId);
    const copied = copies.get(copyId);
    if (!original || !draft || !copied) return { ok: false, reason: "missing-member" };
    for (const key of protectedKeys) {
      if (!pageDocumentJsonEquals(original[key], draft[key])) {
        return { ok: false, reason: "changed-native-structure" };
      }
    }
    // Keep the materialized shape/frame/group identities, point ids and part
    // custody. Only actual preview differences replace authored properties.
    for (const key of new Set([...Object.keys(original), ...Object.keys(draft)])) {
      if (protectedKeys.has(key) || pageDocumentJsonEquals(original[key], draft[key])) continue;
      const value = draft[key];
      if (key === "left" || key === "top") {
        const before = original[key], offset = copied[key];
        if (typeof value !== "number" || typeof before !== "number" || typeof offset !== "number"
          || ![value, before, offset].every(Number.isFinite)) return { ok: false, reason: "invalid-geometry" };
        copied[key] = offset + value - before;
      } else if (key === "startShapeId" || key === "endShapeId") {
        if (value === undefined) delete copied[key];
        else if (typeof value !== "string" || !baseline.has(value)) return { ok: false, reason: "unknown-reference" };
        else copied[key] = plan.blockIdMap[value] ?? value;
      } else if (value === undefined) {
        delete copied[key];
      } else {
        copied[key] = JSON.parse(JSON.stringify(value)) as unknown;
      }
    }
  }
  if (targetSlideId !== undefined) {
    const originalSlide = slideOf(sourceContent), copiedSlide = slideOf(content, targetSlideId);
    const placeholders = record(record(originalSlide?.source)?.layoutCascade)?.placeholders;
    // Native reimport derives the slide label from its title placeholder, or
    // the first text-bearing object when no title placeholder exists. Keep a
    // provably derived label in sync with the copied text before Save, rather
    // than weakening the full authored-content acknowledgement comparison.
    // Ambiguous/empty/non-text titles remain subject to the exact comparison.
    if (originalSlide && copiedSlide && Array.isArray(placeholders) && placeholders.every(record) &&
      originalSlide.title === copiedSlide.title) {
      const titlePlaceholders = placeholders.map(record).filter(item =>
        item?.placeholderType === "title" || item?.placeholderType === "ctrTitle");
      const titleBlock = titlePlaceholders.length === 1
        ? [...baseline.values()].find(block => record(block.source)?.shapeId === titlePlaceholders[0]?.shapeId)
        : titlePlaceholders.length === 0 ? baseline.values().next().value : undefined;
      const copiedBlock = typeof titleBlock?.id === "string" ? copies.get(plan.blockIdMap[titleBlock.id]) : undefined;
      if (titleBlock?.type === "text" && copiedBlock?.type === "text" &&
        typeof titleBlock.text === "string" && titleBlock.text.trim() === originalSlide.title &&
        typeof copiedBlock.text === "string" && copiedBlock.text.trim()) {
        copiedSlide.title = copiedBlock.text.trim();
      }
    }
  }
  if (!presentationPageContentSchema.safeParse(content).success) return { ok: false, reason: "invalid-content" };
  return { ok: true, plan: { ...plan, content } };
}

export async function requestPresentationNativeCopy(
  request: PresentationNativeCopyRequest,
): Promise<PresentationNativeCopyPlan> {
  const response = await apiPost<unknown>(
    `/api/v1/pages/${encodeURIComponent(request.pageId)}/native-pptx-copy-plan`,
    {
      expected_content_version: request.contentVersion,
      model_slide_id: request.slideId,
      block_ids: request.blockIds,
      offset_x_emu: request.offsetEmu,
      offset_y_emu: request.offsetYEmu ?? request.offsetEmu,
      ...(crossSlide(request) ? { target_model_slide_id: request.targetSlideId } : {}),
    },
  );
  const plan = readPresentationNativeCopyPlan(response, request);
  if (!plan) throw new Error("native_copy_response_mismatch");
  return plan;
}

export function canRequestNativeSlideCopy(slide: unknown): boolean {
  const source = record(record(slide)?.source), cascade = record(source?.layoutCascade);
  return Boolean(source && /^sld[1-9]\d*$/.test(String(source.slideId)) && cascade &&
    /^[a-f0-9]{64}$/.test(String(cascade.packageFingerprint)) &&
    !("nativeSlideInsertion" in source));
}

/** Only use after the exact native copy export receipt has been verified.
 * Rebinding source custody after Save is not a second authored document edit. */
export function presentationCopyAuthoredContentMatches(left: RecordValue, right: RecordValue): boolean {
  const withoutSources = (input: RecordValue): RecordValue => {
    const copy = structuredClone(input);
    if (Array.isArray(copy.slides)) for (const raw of copy.slides) {
      const slide = record(raw);
      if (!slide) continue;
      delete slide.source;
      if (Array.isArray(slide.blocks)) for (const raw of slide.blocks) {
        const block = record(raw);
        if (block) delete block.source;
      }
    }
    return copy;
  };
  return presentationNativeSavedContentMatches(withoutSources(left), withoutSources(right));
}

export type PresentationNativeSlideCopyRequest = Pick<PresentationNativeCopyRequest,
  "pageId" | "contentVersion" | "slideId"> & { sourceContent: RecordValue };

/** A copy the server made from the page, its package not lining up with the page
 * (`rebuilt_from_page`): a New Slide after the original, every other slide as it was. */
function readPresentationSlideCopyRebuiltFromPage(response: RecordValue, request: PresentationNativeSlideCopyRequest):
  { content: RecordValue; insertedSlideId: string; rebuiltFromPage: true } | null {
  const content = record(inlinePresentationDeckReceipts(response.content));
  const sourceSlides = request.sourceContent.slides;
  const insertedId = response.inserted_slide_id;
  if (!content || !Array.isArray(sourceSlides) || typeof insertedId !== "string" ||
    response.expected_content_version !== request.contentVersion) return null;
  const parsed = presentationPageContentSchema.safeParse(content);
  if (!parsed.success || parsed.data.slides.length !== sourceSlides.length + 1) return null;
  const originalIndex = sourceSlides.findIndex(slide => record(slide)?.id === request.slideId);
  const copy = parsed.data.slides[originalIndex + 1];
  if (originalIndex < 0 || copy?.id !== insertedId || !record(record(copy.source)?.nativeSlideInsertion)) return null;
  const others = parsed.data.slides.filter(slide => slide.id !== insertedId);
  if (others.some((slide, index) => slide.id !== record(sourceSlides[index])?.id)) return null;
  return { content, insertedSlideId: insertedId, rebuiltFromPage: true };
}

export function readPresentationNativeSlideCopyPlan(value: unknown, request: PresentationNativeSlideCopyRequest):
  { content: RecordValue; insertedSlideId: string; rebuiltFromPage?: true } | null {
  const rebuilt = record(value);
  if (rebuilt?.rebuilt_from_page === true) return readPresentationSlideCopyRebuiltFromPage(rebuilt, request);
  const response = record(value), receipt = record(response?.receipt), content = record(inlinePresentationDeckReceipts(response?.content));
  const sourceSlides = request.sourceContent.slides;
  const mapping = record(receipt?.modelBlockIdMap);
  if (!response || !receipt || !content || !mapping || !Array.isArray(sourceSlides) ||
    receipt.schemaVersion !== "pptx-native-slide-copy-v1" || receipt.planId !== response.plan_id ||
    receipt.pageId !== request.pageId || receipt.sourceContentVersion !== request.contentVersion ||
    response.expected_content_version !== request.contentVersion || receipt.modelSlideId !== request.slideId ||
    typeof response.inserted_slide_id !== "string" || receipt.targetModelSlideId !== response.inserted_slide_id ||
    receipt.targetModelSlideId === request.slideId || receipt.targetSlideId === receipt.sourceSlideId) return null;
  const parsed = presentationPageContentSchema.safeParse(content);
  if (!parsed.success || parsed.data.slides.length !== sourceSlides.length + 1) return null;
  const originalIndex = sourceSlides.findIndex(slide => record(slide)?.id === request.slideId);
  const original = record(sourceSlides[originalIndex]), originalSource = record(original?.source);
  const clone = parsed.data.slides[originalIndex + 1];
  if (!original || originalIndex < 0 || !Array.isArray(original.blocks) || !clone ||
    clone.id !== response.inserted_slide_id || record(clone.source)?.slideId !== receipt.targetSlideId ||
    originalSource?.slideId !== receipt.sourceSlideId ||
    parsed.data.slides.filter(slide => slide.id !== clone.id).some((slide, i) => slide.id !== record(sourceSlides[i])?.id)) return null;
  const originals = original.blocks.map(block => record(block)?.id);
  const ids = Object.values(mapping), live = new Set(sourceSlides.flatMap(slide => {
    const row = record(slide); return [row?.id, ...(Array.isArray(row?.blocks) ? row.blocks.map(block => record(block)?.id) : [])];
  }));
  if (Object.keys(mapping).length !== originals.length || originals.some(id => typeof id !== "string" || !(id in mapping)) ||
    new Set(ids).size !== ids.length || ids.length !== clone.blocks.length || live.has(clone.id) ||
    ids.some(id => typeof id !== "string" || !id || live.has(id) || !clone.blocks.some(block => block.id === id))) return null;
  const pending = record(parsed.data.slides[originalIndex].source)?.nativeSlideCopies;
  if (!Array.isArray(pending) || pending.length !== 1 || !pageDocumentJsonEquals(pending[0], receipt)) return null;
  return { content, insertedSlideId: clone.id };
}

export async function requestPresentationNativeSlideCopy(request: PresentationNativeSlideCopyRequest) {
  const response = await apiPost<unknown>(`/api/v1/pages/${encodeURIComponent(request.pageId)}/native-pptx-slide-copy-plan`, {
    expected_content_version: request.contentVersion, model_slide_id: request.slideId,
  });
  const plan = readPresentationNativeSlideCopyPlan(response, request);
  if (!plan) throw new Error("native_slide_copy_response_mismatch");
  return plan;
}
