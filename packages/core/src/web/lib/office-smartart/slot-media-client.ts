/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The browser side of `GET /api/v1/pages/{pageId}/smartart-slot-media`.
 *
 * 转换为可编辑 on a PICTURED preserved diagram needs the diagram's slot
 * pictures, and the page does not store them: four slots inline is half a Docs
 * page's whole string budget and 80% of a deck, so
 * `docs/office-graphics/SMARTART-IMPORT.md` §5.24 ruled inline storage out at
 * any cap and decided the carrier is a REQUEST, made at the moment the author
 * presses the command. This is that request.
 *
 * It is deliberately NOT the `native-projection` route: that one is a pure
 * transformation and carries the five parts in its own body, while this one
 * needs the page's stored source package, which is workspace-bound. So it is a
 * page route, authorised by the same helpers its neighbours use.
 *
 * Kept in its OWN module, like `native-projection-client`, so each host imports
 * one narrow function (and a test can replace it) instead of the API client.
 */

import { apiGet } from "../api";

import { officeSmartArtLayout } from "./catalog";
import type { NativeSmartArtDocPictures } from "./doc-pictures";
import { SMARTART_ITEM_PICTURE_FILLED_LAYOUTS, smartArtPictureLayoutConvertible } from "./native-convert-model";

/** `r:embed` → inline `data:` URL, exactly what `convertNativeSmartArt` takes. */
export type SmartArtSlotMedia = Readonly<Record<string, string>>;

/**
 * Why 转换为可编辑 brought no pictures, in the author's terms.
 *
 * Three outcomes, because the three read differently to the person who pressed
 * the command and Office never leaves a command doing nothing without a word:
 *
 * * `no-page` — there is no page to ask at all. The guest gallery renders a
 *   diagram under the placeholder page id `qa-gallery-note`, which no route
 *   can resolve, so the pictures are simply not reachable from here.
 * * `unavailable` — there IS a page, and its answer did not arrive: the
 *   request failed, the page refused it, the response was about another
 *   diagram, or the stored node cannot name its own data part.
 * * `incomplete` — the answer arrived and does not fill every slot this
 *   layout draws, so converting would empty the rest (§5.28).
 */
export type SmartArtConversionRefusal = "no-page" | "unavailable" | "incomplete" | "picture-unplaced";

/**
 * The media decision, whole: either convert (with the slot pictures, when the
 * diagram has any) or do not, and say why.
 *
 * `ok: false` is still FAIL-CLOSED — the caller converts nothing and the
 * diagram stays preserved, Office's own drawing with its pictures intact. The
 * only thing `reason` adds is that the host can now SAY so.
 */
export type SmartArtConversionMedia =
  | { readonly ok: true; readonly images?: SmartArtSlotMedia }
  | { readonly ok: false; readonly reason: SmartArtConversionRefusal };

/**
 * Each refusal is spoken through the key `toolbar.smartartContext.convertRefused.<reason>`
 * — the shape the SmartArt picture command's own refusals already take
 * (`…pictureRefused.<reason>`), so the three hosts say one sentence each and
 * every host spells it the same way.
 */

export interface SmartArtSlotMediaRequest {
  /** The page the preserved diagram lives on. */
  readonly pageId: string;
  /** The diagram's data part, as the preserved node's `parts.data` names it. */
  readonly dataPart: string;
  /** That part's fingerprint, as the node's `fingerprints.data` records it. */
  readonly dataFingerprint: string;
  /**
   * DISPLAY only (TRIAGE S14c): the longest side, in pixels, the diagram's
   * pictures are drawn at. A picture past the inline cap then comes scaled
   * down to it instead of refusing the diagram. The open-time layout of a
   * cache-less diagram sends it; 转换为可编辑 never does, so nothing scaled
   * down is ever converted or saved.
   */
  readonly displaySide?: number;
}

/**
 * Where one preserved diagram's media lives, read off the node's own view.
 *
 * ONE reader for three hosts: the Docs group, the worksheet object and the
 * slide block all carry the same `nativeSmartArtView.parts` /
 * `.fingerprints` the server projection wrote (`NativeSmartArtProjection`), so
 * the address is the same question in every host and is asked once.
 *
 * `null` when the view cannot name its own data part — an older stored page,
 * or a projection that could not record one — and the caller then keeps the
 * diagram preserved rather than converting it without its pictures.
 */
export function smartArtSlotMediaAddress(
  view: unknown,
): { readonly dataPart: string; readonly dataFingerprint: string } | null {
  const record =
    view && typeof view === "object" && !Array.isArray(view)
      ? (view as Record<string, unknown>)
      : null;
  const parts = record?.parts;
  const fingerprints = record?.fingerprints;
  if (
    !parts ||
    typeof parts !== "object" ||
    Array.isArray(parts) ||
    !fingerprints ||
    typeof fingerprints !== "object" ||
    Array.isArray(fingerprints)
  ) {
    return null;
  }
  const dataPart = (parts as Record<string, unknown>).data;
  const dataFingerprint = (fingerprints as Record<string, unknown>).data;
  return typeof dataPart === "string" &&
    dataPart &&
    typeof dataFingerprint === "string" &&
    /^[0-9a-f]{64}$/.test(dataFingerprint)
    ? { dataPart, dataFingerprint }
    : null;
}

export function smartArtSlotMediaPath(request: SmartArtSlotMediaRequest): string {
  const query = new URLSearchParams({
    part: request.dataPart,
    fingerprint: request.dataFingerprint,
  });
  if (request.displaySide !== undefined) query.set("display", String(request.displaySide));
  return `/api/v1/pages/${encodeURIComponent(request.pageId)}/smartart-slot-media?${query}`;
}

/**
 * One diagram's slot pictures, or `null` when the server refuses or is
 * unreachable.
 *
 * **Whole-diagram, never per slot.** `null` means "do not convert this
 * diagram": it stays preserved — Office's own drawing, pictures and all — with
 * its 转换为可编辑 still offered. A partial answer would fill some slots and
 * silently empty the rest, which is exactly the loss §5.28 found and fixed, so
 * a response this reader cannot prove whole is treated as no answer at all.
 *
 * An EMPTY map is a different, real answer: this diagram carries no picture, so
 * there is nothing to bring and the conversion proceeds.
 */
export async function fetchSmartArtSlotMedia(
  request: SmartArtSlotMediaRequest,
): Promise<SmartArtSlotMedia | null> {
  if (
    !request.pageId ||
    !request.dataPart ||
    !/^[0-9a-f]{64}$/.test(request.dataFingerprint)
  ) {
    return null;
  }
  try {
    const answer = await apiGet<{
      readonly dataPart?: unknown;
      readonly dataFingerprint?: unknown;
      readonly images?: unknown;
    }>(smartArtSlotMediaPath(request));
    // The answer has to be about the diagram that was asked for. A response
    // addressed to another part is not a partial answer to this one.
    if (
      !answer ||
      typeof answer !== "object" ||
      answer.dataPart !== request.dataPart ||
      answer.dataFingerprint !== request.dataFingerprint
    ) {
      return null;
    }
    const images = answer.images;
    if (!images || typeof images !== "object" || Array.isArray(images)) {
      return null;
    }
    const media: Record<string, string> = {};
    for (const [relId, url] of Object.entries(images as Record<string, unknown>)) {
      if (typeof relId !== "string" || !relId || typeof url !== "string" || !url) {
        return null;
      }
      media[relId] = url;
    }
    return media;
  } catch {
    return null;
  }
}

/**
 * The media 转换为可编辑 must bring before converting THIS diagram, fetched.
 *
 * ONE rule for the three hosts, so none of them has to remember it and a
 * fourth is right by construction — the same reason §5.28 moved the gate
 * itself into `convertNativeSmartArt`:
 *
 * * `{}` — nothing to fetch. The catalog layout the author picked draws no
 *   pictures, or no point of this diagram carries one, so the conversion
 *   proceeds with no request at all. This is the overwhelmingly common case:
 *   the Office corpus holds 27 picture layouts and not one of its diagrams has
 *   a picture in it.
 * * `{ images }` — the diagram's slots, resolved out of the page's own source
 *   package.
 * * `{ ok: false, reason }` — this diagram needs pictures and they could not be
 *   brought. The caller converts NOTHING and the diagram stays preserved,
 *   which is Office's own drawing with its pictures intact and its command
 *   still offered. A half answer would fill some slots and silently empty the
 *   rest. `reason` exists so the host can SAY that: the rule refuses, and the
 *   author is told which of the three refusals it was
 *   ({@link SmartArtConversionRefusal}).
 */
export async function smartArtConversionMedia(input: {
  readonly pageId: unknown;
  /** The preserved node's stored `nativeSmartArtView`. */
  readonly view: unknown;
  /** The catalog layout the author picked, or the one the urn resolved to. */
  readonly layoutId: string;
  readonly outline: readonly { readonly imageRelId?: string; readonly imageOnTransition?: true }[];
  /**
   * The diagram's own pictures (`docPictures`, readings row 1005: Radial
   * List's centre): every one must come too, or nothing converts.
   */
  readonly docPictures?: NativeSmartArtDocPictures | null;
  /** Display only (`SmartArtSlotMediaRequest.displaySide`); a conversion never passes it. */
  readonly displaySide?: number;
  /** Injectable for tests; the app never passes it. */
  readonly fetchMedia?: (
    request: SmartArtSlotMediaRequest,
  ) => Promise<SmartArtSlotMedia | null>;
}): Promise<SmartArtConversionMedia> {
  // A page read under row 1005 draws a Radial List item's own picture in its
  // circle (`SMARTART_ITEM_PICTURE_FILLED_LAYOUTS`): wanted like the diagram's.
  const filled = input.docPictures !== undefined && input.docPictures !== null && SMARTART_ITEM_PICTURE_FILLED_LAYOUTS.has(input.layoutId);
  const items: SmartArtConversionMedia = filled ? { ok: true } : await smartArtItemMedia(input);
  const wanted = [
    ...Object.values(input.docPictures ?? {}).map((picture) => picture.imageRelId),
    ...(filled ? input.outline.flatMap((item) => (item.imageRelId ? [item.imageRelId] : [])) : []),
  ];
  if (!items.ok || !wanted.length) return items;
  // The diagram's own pictures resolve against the same data part, and the
  // route brings every picture it holds: asked once more only when the items
  // needed none.
  let images = items.images;
  if (!images || wanted.some((relId) => !images![relId])) {
    if (typeof input.pageId !== "string" || !input.pageId) return { ok: false, reason: "no-page" };
    const address = smartArtSlotMediaAddress(input.view);
    if (!address) return { ok: false, reason: "unavailable" };
    const fetched = await (input.fetchMedia ?? fetchSmartArtSlotMedia)({
      pageId: input.pageId,
      ...address,
      ...(input.displaySide !== undefined ? { displaySide: input.displaySide } : {}),
    });
    if (!fetched) return { ok: false, reason: "unavailable" };
    images = { ...fetched, ...(images ?? {}) };
  }
  return wanted.every((relId) => images![relId]) ? { ok: true, images } : { ok: false, reason: "incomplete" };
}

/** {@link smartArtConversionMedia} for the items' slots alone. */
async function smartArtItemMedia(input: Parameters<typeof smartArtConversionMedia>[0]): Promise<SmartArtConversionMedia> {
  // Asked WITHOUT media first: a diagram whose layout draws no pictures, or
  // whose points carry none, is already convertible and needs no request.
  if (smartArtPictureLayoutConvertible(input.layoutId, input.outline)) {
    return { ok: true };
  }
  // A layout with no picture slot keeps none of the points' pictures, and no
  // media can change that: nothing is asked, and the diagram stays preserved.
  // (Unless every picture is only on a transition: whether the layout draws its
  // transitions as lines, so that none of them is ever seen, is the engine's
  
  // and it refuses the rest.)
  if (officeSmartArtLayout(input.layoutId)?.picture !== true) {
    if (input.outline.every((item) => !item.imageRelId || item.imageOnTransition === true)) return { ok: true };
    return { ok: false, reason: "picture-unplaced" };
  }
  // No page to ask comes FIRST, because it is the refusal the author can do
  // something about: this diagram is being read somewhere its saved file is
  // not reachable from (the guest gallery), not somewhere its pictures are
  // broken.
  if (typeof input.pageId !== "string" || !input.pageId) {
    return { ok: false, reason: "no-page" };
  }
  const address = smartArtSlotMediaAddress(input.view);
  // A node that cannot name its own data part is a page whose pictures cannot
  // be fetched at all — the same thing, to the author, as a request that came
  // back empty.
  if (!address) return { ok: false, reason: "unavailable" };
  const images = await (input.fetchMedia ?? fetchSmartArtSlotMedia)({
    pageId: input.pageId,
    ...address,
    ...(input.displaySide !== undefined ? { displaySide: input.displaySide } : {}),
  });
  if (!images) return { ok: false, reason: "unavailable" };
  // Fail closed for the WHOLE diagram: media that does not in fact fill every
  // slot is media that would empty the rest, so the gate is asked again with
  // what came back rather than trusted because a response arrived.
  return smartArtPictureLayoutConvertible(input.layoutId, input.outline, images)
    ? { ok: true, images }
    : { ok: false, reason: "incomplete" };
}

/**
 * The refusal that is already certain BEFORE the command is pressed, or `null`
 * when pressing it could still work.
 *
 * Only one of the three can be known that early, and it is the one that never
 * resolves by trying again: a diagram carrying pictures, read somewhere with no
 * page to ask. Office greys a command out rather than offering one that must
 * fail, so the hosts hand this to the SmartArt tab and the picture layouts in
 * the 转换为可编辑 gallery are shown disabled, carrying the reason — the same
 * shape the tab already uses for `textBlockedReason` and for a Picture layout a
 * pictured diagram may not change into.
 *
 * It is deliberately NOT per layout: the outline either carries a blip or it
 * does not, and every PICTURE layout in the catalog would need it. A non-picture
 * layout converts with no request at all and stays offered.
 */
export function smartArtPictureConversionUnreachable(input: {
  readonly pageId: unknown;
  /** The preserved node's stored `nativeSmartArtView`. */
  readonly view: unknown;
  readonly outline: readonly { readonly imageRelId?: string }[];
}): SmartArtConversionRefusal | null {
  if (!input.outline.some((item) => item.imageRelId)) return null;
  if (typeof input.pageId !== "string" || !input.pageId) return "no-page";
  return smartArtSlotMediaAddress(input.view) === null ? "unavailable" : null;
}











export async function smartArtSceneMedia(input: {
  readonly pageId: unknown;
  readonly view: unknown;
  readonly layoutId: string;
  readonly outline: readonly { readonly imageRelId?: string; readonly imageOnTransition?: true }[];
  /**
   * The diagram's frame (`frameEmu`): its pictures are DRAWN, so one past the
   * inline cap may come scaled down to the size the frame is drawn at
   * (`smartArtDisplaySide`, TRIAGE S14c) rather than leave the frame empty.
   */
  readonly frameEmu?: { readonly width?: number; readonly height?: number } | null;
  /** The diagram's own pictures (`docPictures`, readings row 1005). */
  readonly docPictures?: NativeSmartArtDocPictures | null;
  /** Injectable for tests; the app never passes it. */
  readonly fetchMedia?: (request: SmartArtSlotMediaRequest) => Promise<SmartArtSlotMedia | null>;
}): Promise<{ readonly images?: SmartArtSlotMedia } | null> {
  if (!input.outline.some((point) => point.imageRelId) && !Object.keys(input.docPictures ?? {}).length) return {};
  const displaySide = smartArtDisplaySide(input.frameEmu);
  const media = await smartArtConversionMedia({ ...input, ...(displaySide !== undefined ? { displaySide } : {}) });
  return media.ok ? (media.images ? { images: media.images } : {}) : null;
}

/** Device pixels per CSS pixel a displayed slot picture is drawn with (the 859 rule: supersampled). */
const DISPLAY_SUPERSAMPLE = 2;
const EMU_PER_PX = 9525;

/**
 * The longest side, in pixels, a diagram's slot picture is drawn at, at most:
 * its frame at 96 dpi, supersampled, within the route's 64..4096 -- the size a
 * picture past the inline cap is scaled down to for display (TRIAGE S14c).
 * The print asks with the same size (`smartart_print_pictures.smartart_display_side`).
 */
export function smartArtDisplaySide(
  frame: { readonly width?: number; readonly height?: number } | null | undefined,
): number | undefined {
  const width = frame?.width;
  const height = frame?.height;
  if (typeof width !== "number" || typeof height !== "number" || !(width > 0) || !(height > 0)) return undefined;
  return Math.max(64, Math.min(4096, Math.round((Math.max(width, height) / EMU_PER_PX) * DISPLAY_SUPERSAMPLE)));
}
