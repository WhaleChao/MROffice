/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Putting a picture INTO a SmartArt picture slot: the one primitive behind the
 * command, shared by every host that has one.
 *
 * §5.18a gave the Docs group member and the worksheet shape object a `src` and
 * taught both canvases to draw it, but no command ever put one there — which is
 * why the two insert galleries still refused the four picture layouts (§5.18,
 * "what this leaves open"). The command is the same in both hosts and in the
 * presentation, so it is written once here:
 *
 *   file pick  →  {@link readSmartArtPictureFile}  →  the member's `src`
 *
 * The file leg is NOT a new reader. `preparePresentationImageFile` is the
 * picture-file gate the presentation host already puts every local picture
 * through — magic bytes, PNG chunk walk, JPEG segment walk, EXIF orientation,
 * a real browser decode, then a `data:` URL — and its two accepted codecs
 * (`image/png`, `image/jpeg`) and its 2,000,000-character cap are exactly what
 * {@link smartArtPictureSrc} accepts and what
 * `office_smartart_diagram_write._inline_image` decodes into the point's
 * `a:blipFill`. Reusing it means a slot refuses in the browser precisely what
 * the page-content validator would have refused on save, instead of letting an
 * author fill a slot that fails to store.
 *
 * What is deliberately NOT here: the file INPUT, the transaction and the patch.
 * The shared Design tab holds no state and owns no transaction
 * (`OfficeSmartArtDesignTab`), so each host opens its own picker and commits
 * its own undo step; this module only decides what a slot is and what may go
 * into one.
 */

import {
  PresentationImageFileError,
  preparePresentationImageFile,
} from "../presentation-image-file";

import { SMARTART_PICTURE_KINDS, type SmartArtKind } from "./model";
import { smartArtPictureSrc } from "./picture";
import PICTURE_BINDINGS from "./picture-bindings.json";

/**
 * Why a chosen file cannot become a slot's picture. Each one is a distinct
 * sentence to the author, so the host's i18n can say which it was rather than
 * "that did not work".
 */
export type SmartArtPictureFileRefusal =
  /** Not a PNG or a JPEG — the only two codecs a diagram blip may carry. */
  | "unsupported-type"
  /** Over the inline cap the page-content validator enforces. */
  | "too-large"
  /** Empty, truncated, mis-declared, or the browser would not decode it. */
  | "unreadable";

export type SmartArtPictureFileResult =
  | { readonly ok: true; readonly src: string }
  | { readonly ok: false; readonly reason: SmartArtPictureFileRefusal };

/**
 * The presentation gate's failure vocabulary, narrowed to the three answers a
 * slot can give. `bad-magic`, `bad-dimensions`, `decode-failed`, `cache-failed`
 * and `empty` are all one thing to the author — the file is not a picture this
 * slot can hold — and telling them apart would be describing MRAI's reader
 * rather than their file.
 */
const REFUSAL_BY_CODE: Record<string, SmartArtPictureFileRefusal> = {
  "unsupported-type": "unsupported-type",
  "too-large": "too-large",
};

/**
 * Read one local file into the `data:` URL a picture slot stores, or say why
 * it cannot be one. Never throws: a refusal is a value, because every caller
 * is a UI command that has to put a sentence on the screen.
 */
export async function readSmartArtPictureFile(
  file: File,
): Promise<SmartArtPictureFileResult> {
  let src: string;
  try {
    src = (await preparePresentationImageFile(file)).src;
  } catch (error) {
    const code =
      error instanceof PresentationImageFileError ? error.code : "decode-failed";
    return { ok: false, reason: REFUSAL_BY_CODE[code] ?? "unreadable" };
  }
  // The gate above and the slot's own model agree by construction, so this is
  // a belt-and-braces re-check rather than a second opinion: it is what stops a
  // later change to either one from putting a payload on a page that the
  // server's validator, or the DiagramML writer, would then refuse.
  const accepted = smartArtPictureSrc(src);
  return accepted
    ? { ok: true, src: accepted }
    : { ok: false, reason: "unsupported-type" };
}

/**
 * Where Office's own `layoutDef` binds each picture slot's blip, by catalog
 * layout and by OUTLINE ITEM index — the whole table, read from the file that
 * is its only copy (`./picture-bindings.json`).
 *
 * The same file is the source of the server writer's `_PICTURE_SLOT_BINDINGS`
 * (`office_smartart_diagram_write.py`), which is what makes this ONE fact
 * rather than two: the browser half decides which slot the menu may offer and
 * the Python half decides which point the blip is written onto, and until now
 * nothing held the two to each other. §5.25 named that as the trap worth
 * naming, harmless only while one side was empty.
 *
 * Its one entry today is `circular-picture-callout`'s hub: Office binds the big
 * central picture (`picture_1`) with
 * `<dgm:forEach axis="ch" ptType="sibTrans" cnt="1">` whose `pictureRepeatNode`
 * carries `<dgm:presOf axis="self"/>`, so the hub's picture lives on the first
 * SIBLING TRANSITION point and not on a node — and a blip MRAI wrote onto that
 * node was discarded by Word and by Excel on their first re-save (§5.19a).
 */
export const SMARTART_PICTURE_SLOT_BINDINGS: Readonly<
  Record<string, Readonly<Record<string, string>>>
> = PICTURE_BINDINGS.layouts;

/**
 * Slots whose picture OFFICE WOULD THROW AWAY, by catalog layout and by the
 * slot's item index: the `"none"` rows of the table above.
 *
 * **Empty today, and derived rather than declared.** The callout's hub used to
 * be here; the writer now puts item 0's picture on its transition point and
 * both outline readers read a transition blip back (§5.21), so the hub keeps
 * its picture through Office and through reopen and nothing is refused.
 *
 * The rule stays because the NEXT layout Office draws no picture in has to
 * degrade visibly — offered and disabled with a reason — rather than silently
 * accepting a picture that will not survive. Adding that layout is now one row
 * in one file, and the day it is added the Python half cannot fail to see it.
 */
const SLOTS_OFFICE_DOES_NOT_DRAW: ReadonlyMap<string, ReadonlySet<number>> =
  new Map(
    Object.entries(SMARTART_PICTURE_SLOT_BINDINGS)
      .map(
        ([kind, slots]) =>
          [
            kind,
            new Set(
              Object.entries(slots)
                .filter(([, where]) => where === "none")
                .map(([itemIndex]) => Number(itemIndex)),
            ),
          ] as const,
      )
      .filter(([, indices]) => indices.size > 0),
  );

/** One picture slot of a diagram, as the command addresses it. */
export interface SmartArtPictureSlot {
  /**
   * Where the slot sits in the diagram's OWN member list. That list is the
   * address both hosts already use — a Docs group's `groupMembers` index and
   * the worksheet's objects sharing one `groupId`, in the same order — so the
   * command needs no second identity.
   */
  readonly memberIndex: number;
  /** 1-based position among the SLOTS, which is what the author is shown. */
  readonly ordinal: number;
  /** Whether a picture is already in it (the menu says so; Office does too). */
  readonly filled: boolean;
  /**
   * False when Office's own layout binds no picture to this slot's point, so a
   * picture put here would be drawn by MRAI and discarded by Office on its
   * first re-save ({@link SLOTS_OFFICE_DOES_NOT_DRAW}). The command offers such
   * a slot disabled rather than accepting a picture it knows will be lost.
   */
  readonly exportable: boolean;
}

/**
 * A slot as the caller reads one off its own member. Both hosts spell the
 * marker identically (`smartArtPicturePlaceholder`), but they keep the picture
 * in different places — a Docs member's own `src`, a worksheet object's shape
 * STYLE `src` (§5.18a) — so the caller supplies the payload and this decides
 * nothing about where it lives.
 */
export interface SmartArtPictureSlotInput {
  readonly smartArtPicturePlaceholder?: unknown;
  readonly src?: unknown;
  /** Which outline item the slot belongs to, for the exportability rule. */
  readonly smartArtItemIndex?: unknown;
}

/**
 * The diagram's picture slots, in member order.
 *
 * A slot is a member the projection marked `smartArtPicturePlaceholder` — the
 * marker every host writes for an engine `image` node and the one the identity
 * extension carries as `picture="1"`. A diagram with none is not a picture
 * layout, and the command is not offered at all.
 */
export function smartArtPictureSlots(
  members: readonly SmartArtPictureSlotInput[],
  kind?: string,
): readonly SmartArtPictureSlot[] {
  const unexportable = kind ? SLOTS_OFFICE_DOES_NOT_DRAW.get(kind) : undefined;
  const slots: SmartArtPictureSlot[] = [];
  members.forEach((member, memberIndex) => {
    if (member.smartArtPicturePlaceholder !== true) return;
    const itemIndex = member.smartArtItemIndex;
    slots.push({
      memberIndex,
      ordinal: slots.length + 1,
      filled: smartArtPictureSrc(member.src) !== null,
      // No `kind` means the caller did not say which layout this is, and an
      // unknown layout is not a licence to offer a slot Office would empty —
      // but it is also not a reason to disable every slot, so the rule only
      // fires on a slot this map actually names.
      exportable: !(
        unexportable &&
        typeof itemIndex === "number" &&
        unexportable.has(itemIndex)
      ),
    });
  });
  return slots;
}

/**
 * Whether changing this diagram to `targetKind` would THROW A PICTURE AWAY.
 *
 * The engine's `smartArtPictureLayoutChangeBlocked` asks the same question of
 * the presentation's blocks, where a picture member is an `image` block with
 * its own custody fields. A Docs group member and a worksheet member are not
 * that — they are shapes carrying a slot marker and an inline `src` (§5.18a) —
 * so the two models answer in their own terms and the shared Design tab takes
 * either. Consolidating them further would mean one page model, which is the
 * very thing the three hosts do not share.
 *
 * An EMPTY slot blocks nothing: there is no picture to lose, and Office lets a
 * picture layout with nothing in it become anything.
 */
export function smartArtPictureSlotChangeBlocked(
  slots: readonly SmartArtPictureSlot[],
  targetKind: SmartArtKind,
): boolean {
  return (
    !SMARTART_PICTURE_KINDS.some((kind) => kind === targetKind) &&
    slots.some((slot) => slot.filled)
  );
}

/**
 * The pictures a diagram is holding, keyed by the DURABLE item id of the slot
 * holding each one.
 *
 * Every Design-tab command but this one rebuilds the diagram's members from the
 * engine, and an engine-built slot is empty — so without carrying them across,
 * a Change Colors or a Promote on a filled picture layout would silently empty
 * its slots. The item id is the right key for the same reason
 * `preserveSmartArtPictureCustody` uses it on the presentation's blocks: it is
 * what survives a relayout, where a member's position does not.
 */
export function smartArtSlotPicturesByItemId(
  members: readonly (SmartArtPictureSlotInput & {
    readonly smartArtItemId?: unknown;
  })[],
): ReadonlyMap<string, string> {
  const pictures = new Map<string, string>();
  for (const member of members) {
    if (member.smartArtPicturePlaceholder !== true) continue;
    const src = smartArtPictureSrc(member.src);
    const itemId = member.smartArtItemId;
    if (src && typeof itemId === "string" && itemId) pictures.set(itemId, src);
  }
  return pictures;
}
