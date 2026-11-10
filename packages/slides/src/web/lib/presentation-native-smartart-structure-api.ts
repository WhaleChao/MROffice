/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { inlinePresentationDeckReceipts } from "@mrnote/core/web/lib/presentation-deck-receipts.ts";
import { apiPost } from "@mrnote/core/web/lib/api.ts";
import type { PresentationBlock } from "@mrnote/core/web/components/console/editor/native-document/documentCommands.ts";
import { presentationBlockSchema, presentationPageContentSchema } from "@mrnote/core/web/lib/native-page-schemas.ts";
import { pageDocumentJsonEquals } from "@mrnote/core/web/lib/page-save-acknowledgement.ts";
import { presentationNativeSavedContentMatches } from "@mrnote/core/web/lib/presentation-native-save.ts";
import { readNativeSmartArtGroup, type NativeSmartArtGroup } from "@mrnote/core/web/lib/presentation-native-smartart.ts";
import { readNativeSmartArtStructure, type NativeSmartArtStructureReceipt, type NativeSmartArtStructureRequest } from "./presentation-native-smartart-structure";

type RecordValue = Record<string, unknown>;
const record = (v: unknown): RecordValue | null => v && typeof v === "object" && !Array.isArray(v) ? v as RecordValue : null;
const digest = (v: unknown) => typeof v === "string" && /^[a-f0-9]{64}$/.test(v);
export type NativeSmartArtProfileInput = { pageId: string; contentVersion: number; slideId: string; frameShapeId: string };
export type NativeSmartArtProfile = { supported: false } | { supported: true; profile: NativeSmartArtStructureReceipt; group: NativeSmartArtGroup };
export type NativeSmartArtPlanInput = Omit<NativeSmartArtProfileInput, "frameShapeId"> & { request: NativeSmartArtStructureRequest };
export type NativeSmartArtPlan = { content: RecordValue; receipt: RecordValue; selectedBlockId: string; focusPointId: string };

export function readNativeSmartArtProfileResponse(value: unknown, input: NativeSmartArtProfileInput): NativeSmartArtProfile | null {
  const response = record(value);
  if (!response || response.expected_content_version !== input.contentVersion || response.model_slide_id !== input.slideId || response.frame_shape_id !== input.frameShapeId) return null;
  if (response.supported === false) return { supported: false };
  if (response.supported !== true || !Array.isArray(response.blocks) || response.blocks.length < 1 || response.blocks.length > 200 ||
    response.blocks.some(b => !presentationBlockSchema.safeParse(b).success)) return null;
  const blocks = response.blocks as PresentationBlock[], group = readNativeSmartArtGroup(blocks, blocks[0]);
  const profile = readNativeSmartArtStructure(group, response.profile);
  if (!group || !profile || group.members.length !== blocks.length || profile.frameShapeId !== input.frameShapeId) return null;
  return { supported: true, profile, group };
}

export function readNativeSmartArtPlanResponse(value: unknown, input: NativeSmartArtPlanInput): NativeSmartArtPlan | null {
  const response = record(value), receipt = record(response?.receipt), content = record(inlinePresentationDeckReceipts(response?.content));
  const schema = input.request.schemaVersion, operation = input.request.operation;
  if (!(schema === "pptx-smartart-structure-v1" && (operation === "insert-after" || operation === "delete") ||
    schema === "pptx-smartart-hierarchy-v1" && (operation === "promote" || operation === "demote"))) return null;
  if (!response || !receipt || !content || response.expected_content_version !== input.contentVersion ||
    receipt.schemaVersion !== "pptx-native-smartart-structure-plan-v1" || typeof receipt.planId !== "string" || !receipt.planId || receipt.planId !== response.plan_id ||
    receipt.pageId !== input.pageId || receipt.sourceContentVersion !== input.contentVersion || receipt.modelSlideId !== input.slideId ||
    !pageDocumentJsonEquals(receipt.request, input.request) ||
    !["sourceContentSha256", "sourcePackageFingerprint", "fontManifestSha256", "outputPackageFingerprint", "outputContentSha256"].every(k => digest(receipt[k])) ||
    typeof receipt.sourceResourceId !== "string" || !receipt.sourceResourceId ||
    !["original", "export"].includes(String(receipt.sourceResourceVersionKind)) ||
    !(receipt.sourceResourceVersionId === undefined || receipt.sourceResourceVersionId === null || typeof receipt.sourceResourceVersionId === "string" && receipt.sourceResourceVersionId.length > 0) ||
    receipt.sourceResourceVersionKind === "export" && !receipt.sourceResourceVersionId ||
    typeof response.selected_block_id !== "string" || typeof response.focus_point_id !== "string" ||
    !presentationPageContentSchema.safeParse(content).success) return null;
  const slides = content.slides as RecordValue[], slide = slides.find(s => s.id === input.slideId), source = record(slide?.source);
  const embedded = source?.nativeSmartArtStructures;
  if (!slide || source?.slideId !== receipt.sourceSlideId || !Array.isArray(embedded) || embedded.length !== 1 ||
    !pageDocumentJsonEquals(embedded[0], receipt) || !Array.isArray(slide.blocks)) return null;
  const selected = (slide.blocks as PresentationBlock[]).find(b => b.id === response.selected_block_id);
  const group = readNativeSmartArtGroup(slide.blocks as PresentationBlock[], selected);
  if (!group || group.receipt.frameShapeId !== input.request.frameShapeId || !group.outline.some(p => p.id === response.focus_point_id)) return null;
  if (input.request.operation === "insert-after" && response.focus_point_id !== input.request.newPointId ||
    input.request.operation === "delete" && group.outline.some(p => p.id === input.request.pointId) ||
    (input.request.operation === "promote" || input.request.operation === "demote") && response.focus_point_id !== input.request.pointId) return null;
  // Keep the exact server snapshot; renderer normalization must not replace its
  // version-bound provenance or fabricate pending plan metadata.
  return { content, receipt, selectedBlockId: response.selected_block_id, focusPointId: response.focus_point_id };
}

export function hasPendingNativeSmartArtStructure(slide: unknown): boolean {
  const source = record(record(slide)?.source), pending = source?.nativeSmartArtStructures;
  return pending !== undefined && (!Array.isArray(pending) || pending.length > 0);
}

/** Consume a confirmed native Save, never a preview or an authored edit. */
export function readNativeSmartArtSavedContent(plan: Pick<NativeSmartArtPlan, "content" | "receipt">, value: unknown): RecordValue | null {
  const page = record(value), native = record(page?.editable_page_patchback), content = record(page?.content_json), receipt = plan.receipt;
  if (!page || !native || !content || page.id !== receipt.pageId || page.page_type !== "presentation" ||
    !Number.isSafeInteger(page.content_version) || typeof page.content_version !== "number" ||
    typeof receipt.sourceContentVersion !== "number" || page.content_version <= receipt.sourceContentVersion ||
    native.status !== "succeeded" || native.adapter !== "pptx" || native.loss_entry_count !== 0 ||
    native.resource_id !== receipt.sourceResourceId || typeof native.resource_version_id !== "string" || !native.resource_version_id ||
    native.output_sha256 !== receipt.outputPackageFingerprint || !digest(native.output_sha256) ||
    !presentationPageContentSchema.safeParse(content).success || !Array.isArray(content.slides) ||
    content.slides.some(hasPendingNativeSmartArtStructure)) return null;
  const expected = structuredClone(plan.content), slides = expected.slides;
  if (!Array.isArray(slides)) return null;
  const carrier = slides.map(record).find(s => s?.id === receipt.modelSlideId), source = record(carrier?.source);
  const pending = source?.nativeSmartArtStructures;
  if (!source || !Array.isArray(pending) || pending.length !== 1 || !pageDocumentJsonEquals(pending[0], receipt)) return null;
  delete source.nativeSmartArtStructures;
  return presentationNativeSavedContentMatches(expected, content) ? content : null;
}

export async function requestNativeSmartArtProfile(input: NativeSmartArtProfileInput, signal?: AbortSignal): Promise<NativeSmartArtProfile> {
  const response = await apiPost<unknown>(`/api/v1/pages/${encodeURIComponent(input.pageId)}/native-pptx-smartart-structure-profile`,
    { expected_content_version: input.contentVersion, model_slide_id: input.slideId, frame_shape_id: input.frameShapeId }, { signal });
  const profile = readNativeSmartArtProfileResponse(response, input);
  if (!profile) throw new Error("native_smartart_profile_mismatch");
  return profile;
}
export async function requestNativeSmartArtPlan(input: NativeSmartArtPlanInput, signal?: AbortSignal): Promise<NativeSmartArtPlan> {
  const response = await apiPost<unknown>(`/api/v1/pages/${encodeURIComponent(input.pageId)}/native-pptx-smartart-structure-plan`,
    { expected_content_version: input.contentVersion, model_slide_id: input.slideId, request: input.request }, { signal });
  const plan = readNativeSmartArtPlanResponse(response, input);
  if (!plan) throw new Error("native_smartart_plan_mismatch");
  return plan;
}
