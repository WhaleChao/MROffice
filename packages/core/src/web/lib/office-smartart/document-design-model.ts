/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The Docs SmartArt MODEL: a `documentShape` group read back as the Design
 * tab's diagram, and the edits that rewrite members the group already has --
 * a picture into a slot, Larger / Smaller, moving or resizing one shape,
 * typing into one shape, the preserved diagram's shape walk -- each verified
 * against the shape model before it is committed. None of them lays the
 * diagram out again.
 *
 * A NEW layout (change layout, colours, quick style, Reset, an outline
 * change) is the server's (`./document-design-ops`, run by the engine host;
 * the browser asks through `./remote`), which re-exports everything here.
 */

import {
  DOCUMENT_SHAPE_MIN_EDGE_EMU,
  documentShapeIsPreservedNativeSmartArt,
  documentShapeSmartArtIdentity,
  normalizeDocumentShapeAttrs,
  type DocumentShapeAttrs,
  type DocumentShapeAttrsInput,
  type DocumentShapeGroupMember,
} from "../document-shapes";
import { DOCUMENT_SMARTART_STYLE } from "./document-shapes-model";
import {
  isSmartArtHierBranch,
  type SmartArtDirection,
  type SmartArtHierBranch,
  type SmartArtKind,
  type SmartArtStyle,
} from "./model";
import {
  MAX_SMARTART_OUTLINE_ITEMS,
  MAX_SMARTART_OUTLINE_TEXT_LENGTH,
  type SmartArtOutlineItem,
} from "./outline";
import {
  resizeSmartArtItem,
  smartArtOutlineAfterTextEdit,
  smartArtConnectorColor,
  smartArtPaletteNode,
  smartArtVisibleFill,
} from "./design-model";
import {
  applyMultiMove,
  applyMultiResize,
  type TransformMember,
} from "../presentation-multi-transform";
import type { ResizeHandle } from "../presentation-drag";
import type { TextRun } from "../presentation-runs";
import { smartArtPictureSrc } from "./picture";
import { smartArtConnectorsFollowNodes } from "./connector-follow";
import { smartArtMembersAfterBoxEdit } from "./custom-geometry";
import { smartArtPictureSlots } from "./picture-command";
import { type SmartArtStyleId } from "./styles-model";

/** One Docs SmartArt, as the Design tab reads it. */
export interface DocumentSmartArtDiagram {
  /** The catalog layout the group IS, checked in the layouts gallery. */
  readonly kind: SmartArtKind;
  /** The one id every member of this diagram shares. */
  readonly groupId: string;
  /** The group's normalized attrs — what a patch is computed against. */
  readonly attrs: DocumentShapeAttrs;
  /** Its members, in the model's own order. */
  readonly members: readonly DocumentShapeGroupMember[];
  /** Its outline, in item order. */
  readonly outline: readonly SmartArtOutlineItem[];
  /** The colours the members actually wear, never inferred from a style. */
  readonly colors: SmartArtStyle;
  /** The quick style the members agree on; `null` when they carry none. */
  readonly styleId: SmartArtStyleId | null;
  /** Which way it runs: the group's own `smartArtDirection`, else left to right. */
  readonly direction: SmartArtDirection;
  /** Its Change Colors variation (Office's id); `null` is Office's default. */
  readonly colorsId: string | null;
  /** The Office alias it is laid out from (`smartArtLayoutUrn`), or null. */
  readonly layoutUrn: string | null;
}

/** The fields one tab command writes back onto the `documentShape` node. */
export type DocumentSmartArtPatch = Pick<
  DocumentShapeAttrs,
  | "groupMembers"
  | "groupCoordinateWidthEmu"
  | "groupCoordinateHeightEmu"
  | "smartArtKind"
  | "smartArtGroupId"
  | "smartArtStyle"
  | "smartArtDirection"
  | "smartArtColors"
  | "smartArtLayoutUrn"
> &
  // The graphic's own fill, line and effect (`dgm:bg`, `dgm:whole`), stated
  // only by Reset Graphic, which clears them.
  Partial<Pick<DocumentShapeAttrs, "fill" | "outline" | "outlineWidthPt" | "effect">>;

/** `Record`-shaped read access to a member, which is an interface. */
function record(member: DocumentShapeGroupMember): Record<string, unknown> {
  return member as unknown as Record<string, unknown>;
}

function colorOrNull(value: unknown): string | null {
  return typeof value === "string" && value !== "none" && value.length > 0
    ? value
    : null;
}

/**
 * The outline a Docs group carries.
 *
 * A Docs member merges an item's label into the shape that hosts it, so the
 * item's text normally lives on its `node` member; a picture point has no text
 * host, and its label stays its own `label` member, which then owns the text.
 * Refused (fail closed) unless the members describe ONE complete outline: item
 * indices 0…n-1 exactly once each, a stable id and level per item, and the
 * nesting the generator itself enforces.
 */
export function readDocumentSmartArtOutline(
  members: readonly DocumentShapeGroupMember[],
): SmartArtOutlineItem[] | null {
  const rows = new Map<
    number,
    {
      id: string;
      level: number;
      nodeText: string | null;
      labelText: string | null;
      assistant: boolean;
      hierBranch?: SmartArtHierBranch;
    }
  >();
  for (const member of members) {
    const role = member.smartArtRole;
    if (role !== "node" && role !== "label") continue;
    const { smartArtItemId: id, smartArtItemIndex: index, smartArtLevel: level } =
      member;
    if (
      typeof id !== "string" ||
      typeof index !== "number" ||
      typeof level !== "number" ||
      typeof member.text !== "string" ||
      member.text.length > MAX_SMARTART_OUTLINE_TEXT_LENGTH
    ) {
      return null;
    }
    const current = rows.get(index);
    if (!current) {
      rows.set(index, {
        id,
        level,
        nodeText: role === "node" ? member.text : null,
        labelText: role === "label" ? member.text : null,
        assistant: member.smartArtAssistant === true,
        ...(role === "node" && isSmartArtHierBranch(member.smartArtHierBranch)
          ? { hierBranch: member.smartArtHierBranch }
          : {}),
      });
      continue;
    }
    if (member.smartArtAssistant === true) current.assistant = true;
    if (role === "node" && isSmartArtHierBranch(member.smartArtHierBranch)) {
      current.hierBranch = member.smartArtHierBranch;
    }
    // Two members of one item must agree on who the item is.
    if (current.id !== id || current.level !== level) return null;
    if (role === "node" && current.nodeText === null) current.nodeText = member.text;
    else if (role === "label" && current.labelText === null) {
      current.labelText = member.text;
    } else return null;
  }
  if (rows.size === 0 || rows.size > MAX_SMARTART_OUTLINE_ITEMS) return null;
  // Titles (Name and Title): the words of the members that name an item.
  const titles = new Map<number, string>();
  for (const member of members) {
    const at = member.smartArtTitleIndex;
    if (
      member.smartArtRole === "connector" &&
      typeof at === "number" &&
      typeof member.text === "string" &&
      member.text.length > 0 &&
      member.text.length <= MAX_SMARTART_OUTLINE_TEXT_LENGTH
    ) {
      titles.set(at, member.text);
    }
  }
  const items: SmartArtOutlineItem[] = [];
  const ids = new Set<string>();
  for (let index = 0; index < rows.size; index += 1) {
    const row = rows.get(index);
    if (
      !row ||
      ids.has(row.id) ||
      !Number.isInteger(row.level) ||
      row.level < 0 ||
      (index === 0 && row.level !== 0) ||
      (index > 0 && row.level > items[index - 1].level + 1)
    ) {
      return null;
    }
    ids.add(row.id);
    const title = titles.get(index);
    items.push({
      id: row.id,
      text: row.labelText ?? row.nodeText ?? "",
      level: row.level,
      ...(row.assistant && row.level > 0 ? { assistant: true as const } : {}),
      ...(title !== undefined ? { title } : {}),
      ...(row.hierBranch ? { hierBranch: row.hierBranch } : {}),
    });
  }
  return items;
}

/**
 * The colour set the diagram actually wears, read off its own members — so a
 * quick-style change keeps the author's colours, exactly as PowerPoint has it.
 * Each role falls back to the Docs insert palette when no member supplies it.
 */
export function readDocumentSmartArtColors(
  members: readonly DocumentShapeGroupMember[],
): SmartArtStyle {
  const node = smartArtPaletteNode(
    members.filter(
      (member) =>
        member.smartArtRole === "node" && member.smartArtPicturePlaceholder !== true,
    ),
    (member) => smartArtVisibleFill(member.fill, member.opacity),
    members.filter(
      (member) =>
        member.smartArtRole !== "label" && member.smartArtPicturePlaceholder !== true,
    ),
  );
  const connector = members.find((member) => member.smartArtRole === "connector");
  return {
    nodeFill:
      colorOrNull(smartArtVisibleFill(node?.fill, node?.opacity)) ??
      DOCUMENT_SMARTART_STYLE.nodeFill,
    nodeStroke: colorOrNull(node?.outline) ?? DOCUMENT_SMARTART_STYLE.nodeStroke,
    connector: smartArtConnectorColor(
      connector?.fill,
      connector?.outline,
      DOCUMENT_SMARTART_STYLE.connector,
    ),
    textColor:
      colorOrNull(node?.textFill) ?? DOCUMENT_SMARTART_STYLE.textColor,
    fontSize: DOCUMENT_SMARTART_STYLE.fontSize,
  };
}

/**
 * The selected group read back as one diagram, or `null` when it is not a
 * catalog SmartArt the tab owns (an ordinary shape group, a preserved Office
 * diagram, or a group whose members no longer describe one outline).
 */
export function readDocumentSmartArtDiagram(
  value: DocumentShapeAttrsInput,
): DocumentSmartArtDiagram | null {
  const identity = documentShapeSmartArtIdentity(value);
  if (!identity) return null;
  const attrs = normalizeDocumentShapeAttrs(value);
  const members = attrs.groupMembers;
  if (!members?.length) return null;
  const outline = readDocumentSmartArtOutline(members);
  if (!outline?.length) return null;
  return {
    kind: identity.kind as SmartArtKind,
    groupId: identity.groupId,
    attrs,
    members,
    outline,
    colors: readDocumentSmartArtColors(members),
    // The group SAYS which of the fourteen it wears; nothing here infers a
    // style from how the members happen to be painted.
    styleId: attrs.smartArtStyle ?? null,
    // …and which way it runs, likewise.
    direction: attrs.smartArtDirection === "rtl" ? "rtl" : "ltr",
    // …and which colour variation, likewise.
    colorsId: attrs.smartArtColors ?? null,
    // …and which Office alias it is drawn as, likewise.
    layoutUrn: attrs.smartArtLayoutUrn ?? null,
  };
}


/**
 * Put a picture into ONE of the diagram's picture slots.
 *
 * Nothing is rebuilt: a slot's picture changes no geometry, no outline and no
 * paint, so the patch is the members the group already has with `src` set on
 * the one member the slot names. That also means it is the only Design-tab
 * command that cannot fail on a skeleton mismatch — there is no second scene to
 * agree with.
 *
 * Fails closed, and each refusal is a real one rather than a guard for its own
 * sake: a member index outside the group, a member that is not a picture SLOT
 * (the validator refuses a `src` anywhere else, so writing one would cost the
 * whole group its identity on save), and a payload the shared spelling does not
 * accept ({@link smartArtPictureSrc} — the same answer the page-content
 * validator and the DiagramML writer give).
 */
export function documentSmartArtPicturePatch(
  diagram: DocumentSmartArtDiagram,
  memberIndex: number,
  src: string,
): DocumentSmartArtPatch | null {
  const target = diagram.members[memberIndex];
  if (!target || target.smartArtPicturePlaceholder !== true) return null;
  // A slot Office's own layout binds no picture to is refused HERE and not
  // only in the menu that disables it: a picture MRAI stores and Office empties
  // on its first re-save is exactly the silent loss this campaign exists to
  // stop, so the model is where it fails closed.
  const slot = smartArtPictureSlots(diagram.members, diagram.kind).find(
    (candidate) => candidate.memberIndex === memberIndex,
  );
  if (!slot?.exportable) return null;
  const accepted = smartArtPictureSrc(src);
  if (!accepted) return null;
  const members = diagram.members.map((member, index) =>
    index === memberIndex ? { ...member, src: accepted } : member,
  );
  const patch = documentSmartArtVerifiedPatch(
    diagram,
    diagram.kind,
    diagram.styleId ?? undefined,
    {
      width: diagram.attrs.groupCoordinateWidthEmu ?? 0,
      height: diagram.attrs.groupCoordinateHeightEmu ?? 0,
    },
    members,
  );
  // `verifiedPatch` compares the fields a REBUILD can move; the picture is not
  // one of them, so the one thing this command actually writes is checked here
  // or it is checked nowhere. A model that dropped the `src` would otherwise
  // hand back a patch that looks like a success and changes nothing.
  return patch?.groupMembers?.[memberIndex]?.src === accepted ? patch : null;
}


/**
 * Fail closed: hand back the members the SHAPE MODEL accepts, and only when it
 * accepted every one of them with its identity intact. A diagram that lost its
 * identity on the way into the model would land as an anonymous shape group —
 * no SmartArt tab, no layout gallery — and one whose members were clamped
 * would be a different picture; both are refusals, not results.
 */
/**
 * SmartArt Format > Shapes > Larger / Smaller on one item, through the shared
 * {@link resizeSmartArtItem}. A member lives in the group's own child space,
 * which is the graphic's frame: the grown shape stays inside it.
 */
export function documentSmartArtResizePatch(
  diagram: DocumentSmartArtDiagram,
  itemId: string,
  factor: number,
): DocumentSmartArtPatch | null {
  const coordinate = {
    width: diagram.attrs.groupCoordinateWidthEmu ?? 0,
    height: diagram.attrs.groupCoordinateHeightEmu ?? 0,
  };
  if (coordinate.width < 1 || coordinate.height < 1) return null;
  const resized = resizeSmartArtItem(
    diagram.members,
    itemId,
    factor,
    (member) => ({
      left: member.xEmu,
      top: member.yEmu,
      width: member.widthEmu,
      height: member.heightEmu,
    }),
    (member, box) => {
      const width = Math.min(coordinate.width, Math.round(box.width));
      const height = Math.min(coordinate.height, Math.round(box.height));
      return {
        ...member,
        xEmu: Math.min(Math.max(0, Math.round(box.left)), coordinate.width - width),
        yEmu: Math.min(Math.max(0, Math.round(box.top)), coordinate.height - height),
        widthEmu: width,
        heightEmu: height,
      };
    },
    12_700,
    { left: 0, top: 0, width: coordinate.width, height: coordinate.height },
    diagram.kind,
  );
  if (!resized) return null;
  return documentSmartArtVerifiedPatch(
    diagram,
    diagram.kind,
    diagram.styleId ?? undefined,
    coordinate,
    resized,
  );
}

/**
 * The members that ARE one SmartArt shape on the canvas: every node / label
 * member carrying the same `smartArtItemId` as `memberId` (a picture point's
 * frame and its caption travel together). Empty for a member that is not an
 * item's (a connector, a background), which the canvas does not select.
 *
 * Office's model, as its keyboard shortcuts and Format tab describe it: a
 * SmartArt graphic is a frame holding shapes; clicking a shape selects THAT
 * shape inside the frame, and it can then be moved, resized (Shift+arrows)
 * and typed into (Enter / F2) on its own, while the frame stays the graphic's
 * (Microsoft Support, "Keyboard shortcuts for SmartArt graphics"; "Change the
 * size or format of a SmartArt graphic").
 */
export function documentSmartArtItemMemberIds(
  members: readonly DocumentShapeGroupMember[],
  memberId: string,
): string[] {
  const owner = members.find((member) => member.memberId === memberId);
  const itemId = owner?.smartArtItemId;
  if (
    !owner ||
    typeof itemId !== "string" ||
    (owner.smartArtRole !== "node" && owner.smartArtRole !== "label")
  ) {
    return [];
  }
  return members
    .filter(
      (member) =>
        member.smartArtItemId === itemId &&
        (member.smartArtRole === "node" || member.smartArtRole === "label"),
    )
    .map((member) => member.memberId);
}

/**
 * The members a press on `memberId` selects on the canvas: an item's shape
 * (`documentSmartArtItemMemberIds`) for a node or a label, and the member
 * alone for a shape of the layout's own drawing — the diamond behind Basic
 * Matrix, a cycle's arrow, a hierarchy's connecting line.
 *
 * Office lets every shape of a SmartArt graphic be clicked and formatted;
 * connecting and background shapes cannot be deleted, and their fill and
 * outline are how an author hides them (Microsoft Support, "Delete a SmartArt
 * graphic or a shape in a SmartArt graphic"). Such a shape holds no words
 * (Office's text pane has no line for it), so it is selected, painted,
 * moved, resized and rotated here, never typed into or removed. Its moved
 * box rides on the member itself; a relayout (Add Shape, a layout switch)
 * draws the layout's own arrows again.
 */
export function documentSmartArtCanvasShapeMemberIds(
  members: readonly DocumentShapeGroupMember[],
  memberId: string,
): string[] {
  const item = documentSmartArtItemMemberIds(members, memberId);
  if (item.length) return item;
  return members.some(
    (member) =>
      member.memberId === memberId && member.smartArtRole === "connector",
  )
    ? [memberId]
    : [];
}

/** The item a canvas member belongs to (its outline row), or `null`. */
export function documentSmartArtMemberItemId(
  members: readonly DocumentShapeGroupMember[],
  memberId: string,
): string | null {
  return documentSmartArtItemMemberIds(members, memberId).length
    ? (members.find((member) => member.memberId === memberId)?.smartArtItemId ??
        null)
    : null;
}

/**
 * The item Delete on a SmartArt member takes out: its own item, or -- for a
 * shape drawn for an item without its words (`smartArtOwnerIndex`, Vertical
 * Circle List's circles) -- that item, found as its node among the members.
 * Null for a connecting arrow, which Office does not delete.
 */
export function documentSmartArtMemberDeletedItemId(
  members: readonly DocumentShapeGroupMember[],
  memberId: string,
): string | null {
  const member = members.find((candidate) => candidate.memberId === memberId);
  if (member?.smartArtRole !== "connector") return documentSmartArtMemberItemId(members, memberId);
  const owner = member.smartArtOwnerIndex;
  if (typeof owner !== "number") return null;
  return (
    members.find(
      (candidate) =>
        candidate.smartArtRole === "node" &&
        candidate.smartArtItemIndex === owner &&
        typeof candidate.smartArtItemId === "string",
    )?.smartArtItemId ?? null
  );
}

/** One pointer (or arrow-key) transform of the selected shape, in child EMU. */
export type DocumentSmartArtMemberTransform =
  | { readonly kind: "move"; readonly dx: number; readonly dy: number }
  | {
      readonly kind: "resize";
      readonly handle: ResizeHandle;
      readonly dx: number;
      readonly dy: number;
      readonly maintainAspect?: boolean;
    }
  | {
      /**
       * Turn ONE shape by `degrees`, around the centre of the member the
       * handle belongs to (Office: "each individual shape is rotated around
       * its own center"); its item's other members (its label, its picture)
       * ride around that centre and turn with it.
       */
      readonly kind: "rotate";
      readonly pivotMemberId: string;
      readonly degrees: number;
    };

/**
 * Move or resize ONE SmartArt shape (all of its item's members, as one set)
 * inside the graphic's frame, through the shared multi-transform the slide
 * canvas uses (`applyMultiMove` / `applyMultiResize`) in the group's own
 * child space, which IS the frame: nothing leaves it. Every other member, the
 * outline, the colours and the quick style are untouched, and the result is
 * the ordinary Docs group patch the writer already takes.
 */
export function documentSmartArtMemberTransformPatch(
  diagram: DocumentSmartArtDiagram,
  memberId: string,
  transform: DocumentSmartArtMemberTransform,
): DocumentSmartArtPatch | null {
  const coordinate = {
    width: diagram.attrs.groupCoordinateWidthEmu ?? 0,
    height: diagram.attrs.groupCoordinateHeightEmu ?? 0,
  };
  if (coordinate.width < 1 || coordinate.height < 1) return null;
  // An item's shape moves with its own members; a connecting arrow (or a
  // background shape) moves, resizes and rotates on its own, as every
  // SmartArt shape does in Office (Microsoft Support, "Keyboard shortcuts
  // for SmartArt graphics": nudge, resize and rotate the selected shape).
  const ids = new Set(documentSmartArtCanvasShapeMemberIds(diagram.members, memberId));
  if (!ids.size) return null;
  const members = transformedGroupMembers(diagram.members, ids, coordinate, transform);
  if (!members) return null;
  return documentSmartArtVerifiedPatch(
    diagram,
    diagram.kind,
    diagram.styleId ?? undefined,
    coordinate,
    documentSmartArtCustomized(diagram.kind, coordinate, diagram.members, members),
  );
}

/**
 * The members after one shape (the members named by `ids`) was moved or
 * resized inside the group's child space, through the slide canvas's own
 * multi-transform; `null` when nothing changed.
 */
function transformedGroupMembers(
  groupMembers: readonly DocumentShapeGroupMember[],
  ids: ReadonlySet<string>,
  coordinate: { width: number; height: number },
  transform: DocumentSmartArtMemberTransform,
): DocumentShapeGroupMember[] | null {
  const moving: TransformMember[] = groupMembers
    .filter((member) => ids.has(member.memberId))
    .map((member) => ({
      blockId: member.memberId,
      startGeom: {
        left: member.xEmu,
        top: member.yEmu,
        width: member.widthEmu,
        height: member.heightEmu,
      },
    }));
  if (transform.kind === "rotate") {
    return rotatedGroupMembers(groupMembers, ids, coordinate, transform);
  }
  const boxes =
    transform.kind === "move"
      ? applyMultiMove(moving, Math.round(transform.dx), Math.round(transform.dy), coordinate)
      : applyMultiResize(
          moving,
          transform.handle,
          transform.dx,
          transform.dy,
          coordinate,
          { maintainAspect: transform.maintainAspect, minSize: 12_700 },
        );
  let changed = false;
  const members = groupMembers.map((member) => {
    const box = boxes.get(member.memberId);
    if (!box) return member;
    const width = Math.min(coordinate.width, Math.max(1, Math.round(box.width)));
    const height = Math.min(coordinate.height, Math.max(1, Math.round(box.height)));
    const next = {
      ...member,
      xEmu: Math.min(Math.max(0, Math.round(box.left)), coordinate.width - width),
      yEmu: Math.min(Math.max(0, Math.round(box.top)), coordinate.height - height),
      widthEmu: width,
      heightEmu: height,
    };
    if (
      next.xEmu !== member.xEmu ||
      next.yEmu !== member.yEmu ||
      next.widthEmu !== member.widthEmu ||
      next.heightEmu !== member.heightEmu
    ) {
      changed = true;
    }
    return next;
  });
  return changed ? members : null;
}

/**
 * The members after one shape was turned: the pivot member keeps its box and
 * adds the angle; each other member of the item (a label laid over the
 * shape, a caption beside it) has its centre carried around the pivot's and
 * adds the same angle, so the item turns as one shape would. Boxes stay in
 * the child space; `null` when nothing changed.
 */
function rotatedGroupMembers(
  groupMembers: readonly DocumentShapeGroupMember[],
  ids: ReadonlySet<string>,
  coordinate: { width: number; height: number },
  transform: Extract<DocumentSmartArtMemberTransform, { kind: "rotate" }>,
): DocumentShapeGroupMember[] | null {
  const pivot = groupMembers.find((member) => member.memberId === transform.pivotMemberId);
  if (!pivot || !Number.isFinite(transform.degrees)) return null;
  const degrees = Math.round(transform.degrees * 10) / 10;
  if (degrees === 0) return null;
  const centerX = pivot.xEmu + pivot.widthEmu / 2;
  const centerY = pivot.yEmu + pivot.heightEmu / 2;
  const radians = (degrees * Math.PI) / 180;
  const cosine = Math.cos(radians);
  const sine = Math.sin(radians);
  let changed = false;
  const members = groupMembers.map((member) => {
    if (!ids.has(member.memberId)) return member;
    const memberCenterX = member.xEmu + member.widthEmu / 2;
    const memberCenterY = member.yEmu + member.heightEmu / 2;
    const dx = memberCenterX - centerX;
    const dy = memberCenterY - centerY;
    const rotatedX = centerX + dx * cosine - dy * sine;
    const rotatedY = centerY + dx * sine + dy * cosine;
    const next = {
      ...member,
      xEmu: Math.min(
        Math.max(0, Math.round(rotatedX - member.widthEmu / 2)),
        Math.max(0, coordinate.width - member.widthEmu),
      ),
      yEmu: Math.min(
        Math.max(0, Math.round(rotatedY - member.heightEmu / 2)),
        Math.max(0, coordinate.height - member.heightEmu),
      ),
      rotation: Math.round(((((member.rotation + degrees) % 360) + 360) % 360) * 10) / 10,
    };
    if (
      next.xEmu !== member.xEmu ||
      next.yEmu !== member.yEmu ||
      next.rotation !== member.rotation
    ) {
      changed = true;
    }
    return next;
  });
  return changed ? members : null;
}

/**
 * A PRESERVED Office diagram's shapes, moved and resized one at a time.
 *
 * Office lets every shape of any SmartArt graphic be moved and resized on its
 * own (Microsoft Support, "Learn more about SmartArt graphics"; "Keyboard
 * shortcuts for SmartArt graphics": arrows nudge, Shift+arrows resize), and
 * keeps the change as `dgm:prSet` factors on the shape's presentation point.
 * A preserved diagram's members are its cached drawing, one `dsp:sp` split
 * into a shape member and a text member that share its `modelId`; that pair
 * is ONE shape on the canvas and moves as one. The save writes the rectangle
 * back into the drawing and the factors onto that point
 * (`office_smartart_preserved_geometry`); nothing else about the diagram may
 * change, so these return only `groupMembers`.
 */
function preservedNativeModelIds(
  attrs: DocumentShapeAttrs,
): Map<string, string> | null {
  if (!documentShapeIsPreservedNativeSmartArt(attrs)) return null;
  const view = attrs.source?.nativeSmartArtView as Record<string, unknown> | undefined;
  const members = Array.isArray(view?.members) ? view.members : [];
  const byMember = new Map<string, string>();
  for (const entry of members) {
    if (!entry || typeof entry !== "object") continue;
    const record = entry as Record<string, unknown>;
    if (typeof record.memberId === "string") {
      byMember.set(
        record.memberId,
        typeof record.modelId === "string" ? record.modelId : "",
      );
    }
  }
  return byMember;
}

/** The members that are ONE shape of a preserved diagram with `memberId`. */
export function documentPreservedSmartArtShapeMemberIds(
  value: DocumentShapeAttrsInput,
  memberId: string,
): string[] {
  const attrs = normalizeDocumentShapeAttrs(value);
  const models = preservedNativeModelIds(attrs);
  const members = attrs.groupMembers ?? [];
  if (!models || !members.some((member) => member.memberId === memberId)) return [];
  const model = models.get(memberId) ?? "";
  if (!model) return [memberId];
  return members
    .filter((member) => models.get(member.memberId) === model)
    .map((member) => member.memberId);
}

/** One member per shape, in drawing order: what Tab / Shift+Tab walk. */
export function documentPreservedSmartArtShapeOrder(
  value: DocumentShapeAttrsInput,
): string[] {
  const attrs = normalizeDocumentShapeAttrs(value);
  const models = preservedNativeModelIds(attrs);
  if (!models) return [];
  const seen = new Set<string>();
  const order: string[] = [];
  for (const member of attrs.groupMembers ?? []) {
    const model = models.get(member.memberId) || `member:${member.memberId}`;
    if (seen.has(model)) continue;
    seen.add(model);
    order.push(member.memberId);
  }
  return order;
}

/** Move or resize one shape of a preserved diagram, in the group's child EMU. */
export function documentPreservedSmartArtMemberTransformPatch(
  value: DocumentShapeAttrsInput,
  memberId: string,
  transform: DocumentSmartArtMemberTransform,
): Pick<DocumentShapeAttrs, "groupMembers"> | null {
  const attrs = normalizeDocumentShapeAttrs(value);
  const ids = new Set(documentPreservedSmartArtShapeMemberIds(attrs, memberId));
  const coordinate = {
    width: attrs.groupCoordinateWidthEmu ?? 0,
    height: attrs.groupCoordinateHeightEmu ?? 0,
  };
  if (!ids.size || coordinate.width < 1 || coordinate.height < 1) return null;
  const members = transformedGroupMembers(
    attrs.groupMembers ?? [],
    ids,
    coordinate,
    transform,
  );
  return members ? { groupMembers: members } : null;
}

/**
 * The shapes a member edit moved or resized, carrying Office's customization
 * factors (`./custom-geometry`), with the arrows and lines that join them
 * re-laid onto them (`./connector-follow`). A Docs member lives in the
 * group's child space, which is the graphic's frame, so every change here is
 * a shape's own, and a re-laid connector stays inside that space.
 */
function documentSmartArtCustomized(
  kind: string | null | undefined,
  coordinate: { width: number; height: number },
  before: readonly DocumentShapeGroupMember[],
  after: readonly DocumentShapeGroupMember[],
): DocumentShapeGroupMember[] {
  const box = (member: DocumentShapeGroupMember) => ({
    left: member.xEmu,
    top: member.yEmu,
    width: member.widthEmu,
    height: member.heightEmu,
  });
  const customized = smartArtMembersAfterBoxEdit(before, after, {
    id: (member) => member.memberId,
    box,
    absolute: false,
  });
  return documentSmartArtConnectorsFollow(kind, coordinate, before, customized);
}

/**
 * `after` with the arrows and lines that join a shape the edit (or a
 * rebuild's re-applied factors) moved or resized re-laid onto it, as Office's
 * connector algorithm re-routes them (`./connector-follow`), in the group's
 * child space.
 */
export function documentSmartArtConnectorsFollow(
  kind: string | null | undefined,
  coordinate: { width: number; height: number },
  before: readonly DocumentShapeGroupMember[],
  after: readonly DocumentShapeGroupMember[],
): DocumentShapeGroupMember[] {
  return smartArtConnectorsFollowNodes(kind, before, after, {
    id: (member) => member.memberId,
    box: (member) => ({
      left: member.xEmu,
      top: member.yEmu,
      width: member.widthEmu,
      height: member.heightEmu,
    }),
    rotation: (member) => member.rotation,
    shape: (member) => member.preset,
    write: (member, next, rotation) =>
      documentMemberWithBox(member, next, rotation, coordinate),
  });
}

/**
 * A Docs member laid at `box` and turned to `rotation`, kept inside the
 * group's child space (the model refuses a member outside it) and no thinner
 * than the model's minimum edge, grown about its own centre.
 */
function documentMemberWithBox(
  member: DocumentShapeGroupMember,
  box: { left: number; top: number; width: number; height: number },
  rotation: number,
  coordinate: { width: number; height: number },
): DocumentShapeGroupMember {
  const widthEmu = Math.min(coordinate.width, Math.max(DOCUMENT_SHAPE_MIN_EDGE_EMU, Math.round(box.width)));
  const heightEmu = Math.min(coordinate.height, Math.max(DOCUMENT_SHAPE_MIN_EDGE_EMU, Math.round(box.height)));
  const left = box.left + box.width / 2 - widthEmu / 2;
  const top = box.top + box.height / 2 - heightEmu / 2;
  return {
    ...member,
    xEmu: Math.min(Math.max(0, Math.round(left)), coordinate.width - widthEmu),
    yEmu: Math.min(Math.max(0, Math.round(top)), coordinate.height - heightEmu),
    widthEmu,
    heightEmu,
    rotation: Math.round(((((rotation % 360) + 360) % 360)) * 10) / 10,
  };
}

/**
 * Type into the selected SmartArt shape. The words go through the shared
 * outline verb the Text Pane uses (`smartArtOutlineAfterTextEdit`: its cap,
 * its no-op refusal), and land on the member that HOSTS the item's text (the
 * caption of a picture point, the node otherwise), so the outline read back
 * off the members is exactly the one the verb produced. Unlike a Text Pane
 * rebuild, nothing is re-laid out: a shape the author moved or resized stays
 * where it is, as it does in Office.
 */
export function documentSmartArtMemberTextPatch(
  diagram: DocumentSmartArtDiagram,
  memberId: string,
  text: string,
  runs?: readonly TextRun[] | null,
): DocumentSmartArtPatch | null {
  const itemId = documentSmartArtMemberItemId(diagram.members, memberId);
  if (!itemId) return null;
  const host =
    diagram.members.find(
      (member) =>
        member.smartArtItemId === itemId && member.smartArtRole === "label",
    ) ??
    diagram.members.find(
      (member) =>
        member.smartArtItemId === itemId && member.smartArtRole === "node",
    );
  if (!host) return null;
  if (text === host.text) {
    // The words are as they were and only their formatting changed (a word
    // made bold or larger): the outline verb would refuse the no-op, so the
    // runs land on the host directly, outline and geometry untouched.
    const runsNow = JSON.stringify(runs?.length ? runs : null);
    if (runsNow === JSON.stringify(host.textRuns ?? null)) return null;
    const members = diagram.members.map((member) => {
      if (member.memberId !== host.memberId) return member;
      const { textRuns: _formatted, ...rest } = member;
      return runs?.length ? { ...rest, textRuns: runs } : rest;
    });
    return documentSmartArtVerifiedPatch(
      diagram,
      diagram.kind,
      diagram.styleId ?? undefined,
      {
        width: diagram.attrs.groupCoordinateWidthEmu ?? 0,
        height: diagram.attrs.groupCoordinateHeightEmu ?? 0,
      },
      members,
    );
  }
  const change = smartArtOutlineAfterTextEdit(diagram.outline, itemId, text);
  const nextText = change?.items.find((item) => item.id === itemId)?.text;
  if (!change || nextText === undefined) return null;
  const members = diagram.members.map((member) => {
    if (member.memberId !== host.memberId) return member;
    // The imported paragraph reading described the words it was read WITH;
    // the runs the field typed go with the words (dropped when they no
    // longer spell them: the verb may have clipped the text).
    const { textBody: _stale, textRuns: _formatted, ...rest } = member;
    return runs?.length && nextText === text
      ? { ...rest, text: nextText, textRuns: runs }
      : { ...rest, text: nextText };
  });
  const patch = documentSmartArtVerifiedPatch(
    diagram,
    diagram.kind,
    diagram.styleId ?? undefined,
    {
      width: diagram.attrs.groupCoordinateWidthEmu ?? 0,
      height: diagram.attrs.groupCoordinateHeightEmu ?? 0,
    },
    members,
  );
  const outline = patch?.groupMembers
    ? readDocumentSmartArtOutline(patch.groupMembers)
    : null;
  return outline &&
    outline.length === change.items.length &&
    outline.every(
      (item, index) =>
        item.id === change.items[index].id &&
        item.text === change.items[index].text &&
        item.level === change.items[index].level,
    )
    ? patch
    : null;
}



export function documentSmartArtVerifiedPatch(
  diagram: DocumentSmartArtDiagram,
  kind: SmartArtKind,
  styleId: SmartArtStyleId | undefined,
  coordinate: { width: number; height: number },
  members: readonly DocumentShapeGroupMember[],
  // A command that does not rebuild keeps the direction the group has.
  direction: SmartArtDirection = diagram.direction,
  // …and its colour variation.
  colorsId: string | null = diagram.colorsId,
  // …and the Office alias it is drawn as.
  layoutUrn: string | null = diagram.layoutUrn,
): DocumentSmartArtPatch | null {
  const smartArtDirection = direction === "rtl" ? ("rtl" as const) : undefined;
  const smartArtColors = colorsId ?? undefined;
  const smartArtLayoutUrn = layoutUrn ?? undefined;
  const verified = normalizeDocumentShapeAttrs({
    ...diagram.attrs,
    groupCoordinateWidthEmu: coordinate.width,
    groupCoordinateHeightEmu: coordinate.height,
    groupMembers: members,
    smartArtKind: kind,
    smartArtGroupId: diagram.groupId,
    smartArtStyle: styleId,
    smartArtDirection,
    smartArtColors,
    smartArtLayoutUrn,
  });
  const accepted = verified.groupMembers;
  if (
    !accepted ||
    accepted.length !== members.length ||
    verified.smartArtKind !== kind ||
    verified.smartArtGroupId !== diagram.groupId ||
    verified.smartArtStyle !== styleId ||
    verified.smartArtDirection !== smartArtDirection ||
    verified.smartArtColors !== smartArtColors ||
    verified.smartArtLayoutUrn !== smartArtLayoutUrn ||
    verified.groupCoordinateWidthEmu !== coordinate.width ||
    verified.groupCoordinateHeightEmu !== coordinate.height ||
    accepted.some(
      (member, index) =>
        member.preset !== members[index].preset ||
        member.xEmu !== members[index].xEmu ||
        member.yEmu !== members[index].yEmu ||
        member.widthEmu !== members[index].widthEmu ||
        member.heightEmu !== members[index].heightEmu ||
        member.text !== members[index].text ||
        member.fill !== members[index].fill ||
        member.outline !== members[index].outline ||
        member.smartArtItemId !== members[index].smartArtItemId,
    )
  ) {
    return null;
  }
  return {
    groupMembers: accepted,
    groupCoordinateWidthEmu: verified.groupCoordinateWidthEmu,
    groupCoordinateHeightEmu: verified.groupCoordinateHeightEmu,
    smartArtKind: kind,
    smartArtGroupId: diagram.groupId,
    ...(styleId ? { smartArtStyle: styleId } : {}),
    // Always stated, because the host MERGES a patch onto the node's attrs:
    // left out, a diagram turned back to left to right would keep the old
    // `rtl` it no longer draws.
    smartArtDirection,
    // …and so is the colour variation (a Reset puts it back to the default).
    smartArtColors,
    // Likewise: a layout switch must not leave the alias it left behind.
    smartArtLayoutUrn,
  };
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

/**
 * Whether the LIVE node still is the diagram a queued command was computed
 * for. A command must never land on a different diagram after a concurrent
 * edit, so the host re-reads the selection and asks this before dispatching.
 */
export function sameDocumentSmartArtDiagram(
  current: unknown,
  expected: DocumentSmartArtDiagram,
): boolean {
  if (!isRecord(current)) return false;
  const identity = documentShapeSmartArtIdentity(
    current as DocumentShapeAttrsInput,
  );
  return (
    identity?.groupId === expected.groupId && identity.kind === expected.kind
  );
}
