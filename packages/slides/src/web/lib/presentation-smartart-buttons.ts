/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The deck editor's SmartArt Design commands on a slide's blocks: the outline
 * commands (Add Shape, Delete, Promote, Demote, Move Up / Down, a node's
 * text), Change Layout, Right to Left, Change Colors, SmartArt Styles and
 * Reset Graphic. The window's Design tab and an editor command run the SAME
 * command, in three steps:
 *
 *   1. a PLAN, here: which diagram the member belongs to, its frame and
 *      outline, and the one rebuild (or restyle) it needs -- or a refusal;
 *   2. the LAYOUT, which is the engine's: the browser asks the engine host
 *      (`smartArtEngine("presentation.rebuild" | "presentation.restyle")`),
 *      the headless command host runs the engine in process
 *      (`./presentation-smartart-buttons-engine`, server only);
 *   3. the APPLY, here: the rebuilt members keep the graphic's Alt Text and
 *      are spliced in where the diagram's first member was, in one change.
 *
 * Nothing in this module lays a diagram out, so the browser can import it
 * (ENGINE-HOST-CONTRACT.md section 11).
 */

import type { PresentationBlock } from "@mrnote/core/web/components/console/editor/native-document/documentCommands.ts";
import {
  readSmartArtGroupDirection,
  readSmartArtGroupFrame,
  readSmartArtGroupMembers,
  smartArtToggledDirection,
  type OfficeSmartArtFrame,
} from "@mrnote/core/web/lib/office-smartart/design-model.ts";
import { DEFAULT_SMARTART_STYLE, readSmartArtBlockMarker, type SmartArtKind } from "@mrnote/core/web/lib/office-smartart/model.ts";
import { readSmartArtOutline, type SmartArtOutlineItem } from "@mrnote/core/web/lib/office-smartart/outline.ts";
import { DEFAULT_SMARTART_STYLE_ID } from "@mrnote/core/web/lib/office-smartart/styles-model.ts";
import type {
  SmartArtPresentationRebuildRequest,
  SmartArtRestyleRequest,
} from "@mrnote/core/web/lib/engine-host/smartart-types.ts";
import { isVerifiedAuthoredSmartArtGroup } from "@mrnote/core/web/lib/presentation-authored-smartart.ts";
import { blockClaimsNativeGroup } from "@mrnote/core/web/lib/presentation-object-buttons.ts";

/**
 * A diagram the deck keeps as the file drew it (a foreign deck's `p:grpSp`
 * claiming a native group) and does not rebuild: only a verified authored
 * group may take the Design tab's structure commands.
 */
export function smartArtGroupIsPreserveOnly(members: readonly PresentationBlock[]): boolean {
  return !isVerifiedAuthoredSmartArtGroup(members) && members.some(blockClaimsNativeGroup);
}

/**
 * A rebuilt diagram keeps the graphic's Alt Text. The deck states it on the
 * graphic's first member (`groupAlt` / `groupDecorative`, as the PPTX states
 * it on the frame), and a rebuild makes new members, so Add Shape emptied the
 * description the author had written (PPT-B165).
 */
export function withPresentationSmartArtAlt(
  members: readonly PresentationBlock[],
  generated: PresentationBlock[] | null,
): PresentationBlock[] | null {
  const carrier = members.find((member) => {
    const record = member as unknown as Record<string, unknown>;
    return "groupAlt" in record || "groupDecorative" in record;
  }) as unknown as Record<string, unknown> | undefined;
  if (!generated?.length || !carrier) return generated;
  const alt = {
    ...(typeof carrier.groupAlt === "string" ? { groupAlt: carrier.groupAlt } : {}),
    ...(carrier.groupDecorative === true ? { groupDecorative: true } : {}),
  };
  return generated.map((block, index) => {
    const record = { ...(block as unknown as Record<string, unknown>) };
    delete record.groupAlt;
    delete record.groupDecorative;
    return (index === 0 ? { ...record, ...alt } : record) as unknown as PresentationBlock;
  });
}

/** The slide's blocks with a diagram's members replaced by its rebuilt ones, at the first member's place. */
export function spliceSmartArtGroup(
  blocks: readonly PresentationBlock[],
  members: readonly PresentationBlock[],
  generated: readonly PresentationBlock[],
): PresentationBlock[] {
  const firstIndex = Math.min(...members.map((member) => blocks.findIndex((block) => block.id === member.id)));
  const memberIds = new Set(members.map((block) => block.id));
  return blocks.flatMap((block, index) =>
    index === firstIndex ? [...generated] : memberIds.has(block.id) ? [] : [block],
  );
}

/** The diagram `primary` belongs to, its frame and outline, when every member may change. */
function writableDiagram(
  blocks: readonly PresentationBlock[],
  primary: PresentationBlock | null | undefined,
  canManipulate: (block: PresentationBlock) => boolean,
): { members: PresentationBlock[]; groupId: string; kind: SmartArtKind; frame: OfficeSmartArtFrame; outline: SmartArtOutlineItem[] } | null {
  const group = readSmartArtGroupMembers(blocks, primary);
  const frame = group ? readSmartArtGroupFrame(group.members) : null;
  const outline = group ? readSmartArtOutline(blocks, group.marker.groupId) : null;
  if (!group || !frame || !outline?.length || group.members.some((block) => !canManipulate(block))) return null;
  return { members: group.members, groupId: group.marker.groupId, kind: group.marker.kind, frame, outline };
}

/**
 * One rebuild of one diagram: the members it replaces and the engine request
 * (`presentation.rebuild`, ENGINE-HOST-CONTRACT.md section 11) that lays the
 * replacement out.
 */
export interface SmartArtRebuildPlan {
  members: PresentationBlock[];
  request: SmartArtPresentationRebuildRequest;
}

function rebuildPlan(
  diagram: NonNullable<ReturnType<typeof writableDiagram>>,
  request: Omit<SmartArtPresentationRebuildRequest, "members" | "groupId">,
): SmartArtRebuildPlan {
  return {
    members: diagram.members,
    request: {
      ...request,
      members: diagram.members as unknown as SmartArtPresentationRebuildRequest["members"],
      groupId: diagram.groupId,
    },
  };
}

/**
 * Step 3 of every rebuild: the slide's blocks with the rebuilt members (Alt
 * Text kept) in the diagram's place, and those members; null when the engine
 * built nothing.
 */
export function smartArtBlocksWithRebuilt(
  blocks: readonly PresentationBlock[],
  plan: SmartArtRebuildPlan,
  rebuilt: readonly PresentationBlock[] | null,
): { blocks: PresentationBlock[]; generated: PresentationBlock[] } | null {
  const generated = withPresentationSmartArtAlt(plan.members, rebuilt ? [...rebuilt] : null);
  if (!generated?.length) return null;
  return { blocks: spliceSmartArtGroup(blocks, plan.members, generated), generated };
}

export type SmartArtOutlineMutation = (
  items: readonly SmartArtOutlineItem[],
  selectedItemId: string,
  kind: SmartArtKind,
) => { items: SmartArtOutlineItem[]; selectedItemId: string } | null;

/**
 * One outline command on the diagram `primary` belongs to, acting on the
 * first of `preferredItemIds` the outline holds (the Text Pane's row, the
 * selected node) or else its first item: the rebuild it needs and the item
 * selected afterwards; null when refused.
 */
export function smartArtOutlinePlan(
  blocks: readonly PresentationBlock[],
  primary: PresentationBlock | null | undefined,
  preferredItemIds: ReadonlyArray<string | null | undefined>,
  mutate: SmartArtOutlineMutation,
  canManipulate: (block: PresentationBlock) => boolean,
): { plan: SmartArtRebuildPlan; selectedItemId: string } | null {
  const diagram = writableDiagram(blocks, primary, canManipulate);
  if (!diagram) return null;
  const selectedItemId = [...preferredItemIds, diagram.outline[0].id].find(
    (candidate): candidate is string =>
      typeof candidate === "string" && diagram.outline.some((item) => item.id === candidate),
  );
  if (!selectedItemId) return null;
  const result = mutate(diagram.outline, selectedItemId, diagram.kind);
  if (!result) return null;
  return {
    plan: rebuildPlan(diagram, {
      kind: diagram.kind,
      frame: diagram.frame,
      outline: result.items,
      allowItemChanges: true,
    }),
    selectedItemId: result.selectedItemId,
  };
}

/**
 * Change Layout (or, with `toggleDirection`, Right to Left) on the diagram
 * `primary` belongs to: rebuilt as `kind` in its frame with its outline.
 * A layout switch keeps the diagram's direction; Right to Left mirrors the
 * layout and leaves the text pane alone (Office's `dgm:dir="rev"`). A layout
 * that draws fewer items than the diagram has keeps the rest as hidden
 * members (`smartArtHidden`), as Office keeps unused text for a later switch.
 */
export function smartArtLayoutPlan(
  blocks: readonly PresentationBlock[],
  primary: PresentationBlock | null | undefined,
  kind: SmartArtKind,
  toggleDirection: boolean,
  canManipulate: (block: PresentationBlock) => boolean,
  /** The pick's Office alias (its own gallery entry), `null` a catalog layout. */
  layoutUrn?: string | null,
): SmartArtRebuildPlan | null {
  const diagram = writableDiagram(blocks, primary, canManipulate);
  if (!diagram) return null;
  const { outline } = diagram;
  return rebuildPlan(diagram, {
    kind,
    frame: diagram.frame,
    outline,
    ...(toggleDirection ? { direction: smartArtToggledDirection(readSmartArtGroupDirection(diagram.members)) } : {}),
    ...(!toggleDirection && layoutUrn !== undefined ? { layoutUrn } : {}),
  });
}

/**
 * Reset Graphic: the diagram rebuilt with the catalog's colours and quick
 * style, discarding the shapes' own fills and lines and every move and
 * resize, then handed back to the catalog palette (`resetPaint`). A Right to
 * Left applied to the whole graphic and its background colour are NOT
 * removed (Microsoft Support, "Restore the default layout and color of your
 * SmartArt graphic", Customizations not removed; live Word audit WUA 260):
 * the rebuild keeps the direction it has, and the host keeps the graphic's
 * fill while taking off its own line and effect.
 */
export function smartArtResetPlan(
  blocks: readonly PresentationBlock[],
  primary: PresentationBlock | null | undefined,
  canManipulate: (block: PresentationBlock) => boolean,
): SmartArtRebuildPlan | null {
  const diagram = writableDiagram(blocks, primary, canManipulate);
  if (!diagram) return null;
  return rebuildPlan(diagram, {
    kind: diagram.kind,
    frame: diagram.frame,
    outline: diagram.outline,
    colors: DEFAULT_SMARTART_STYLE,
    styleId: DEFAULT_SMARTART_STYLE_ID,
    keepAuthorPaint: false,
    resetCustomGeometry: true,
    resetPaint: true,
  });
}

/** The restyle change Change Colors / SmartArt Styles asks for. */
export type SmartArtRestyleChange = Omit<SmartArtRestyleRequest, "members" | "groupId">;

/**
 * Change Colors and the SmartArt Styles gallery: the diagram's members are
 * restyled in place (geometry is not rebuilt). The request for
 * `presentation.restyle`, or null when a member may not change.
 */
export function smartArtRestylePlan(
  blocks: readonly PresentationBlock[],
  primary: PresentationBlock | null | undefined,
  change: SmartArtRestyleChange,
  canManipulate: (block: PresentationBlock) => boolean,
): { members: PresentationBlock[]; request: SmartArtRestyleRequest } | null {
  const group = readSmartArtGroupMembers(blocks, primary);
  if (!group || group.members.some((block) => !canManipulate(block))) return null;
  return {
    members: group.members,
    request: {
      ...change,
      members: group.members as unknown as SmartArtRestyleRequest["members"],
      groupId: group.marker.groupId,
    },
  };
}

/**
 * The slide's blocks with the restyled members in place of the old ones (same
 * ids, same order); every other block untouched.
 */
export function smartArtBlocksWithRestyled(
  blocks: readonly PresentationBlock[],
  restyled: readonly PresentationBlock[],
): PresentationBlock[] {
  const byId = new Map(restyled.map((block) => [block.id, block]));
  return blocks.map((block) => byId.get(block.id) ?? block);
}

/** The outline item a SmartArt member draws (a node or its label), or null. */
export function smartArtMemberItemId(block: PresentationBlock | null | undefined): string | null {
  return readSmartArtBlockMarker(block)?.itemId ?? null;
}
