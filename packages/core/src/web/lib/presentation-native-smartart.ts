/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { nativeLayoutToCatalogId } from "./office-smartart/native-layouts";
import { nativeSmartArtHostTheme } from "./office-smartart/document-native-model";
import type { PresentationBlock } from "../components/console/editor/native-document/documentCommands";
import { smartArtColorVariantIdFromUrn, type SmartArtSchemeColor } from "./office-smartart/colors";
import { readSmartArtCustom } from "./office-smartart/custom-geometry";
import { isShapeKind } from "./presentation-shapes";
import { normalizeOfficeShapeEffectPreset } from "./office-shape-effects";
import { normalizeOfficePictureEffects } from "./office-picture-effects";
import {
  normalizeOfficeShapeGradient,
  normalizeOfficeShapeShadow,
} from "./office-shape-paint";
import { normalizeOfficeTextPaint, type OfficeTextPaint } from "./office-text-paint";
import { normalizeOfficeShapeTextFont } from "./office-shape-text-font";
import { normalizeRuns } from "./presentation-runs";
import type { SmartArtAuthorFormat } from "@/lib/office-smartart/native-outline";
import { MAX_SMARTART_OUTLINE_LEVEL } from "./office-smartart/model";
import {
  smartArtBlockDrawingBoxes,
  smartArtDrawingLikeness,
  smartArtConversionLooksLikeDrawing,
} from "./office-smartart/drawing-likeness";
import {
  nativeSmartArtFrameInBounds,
  type NativeSmartArtFrame,
} from "./office-smartart/native-convert-model";
import type { NativeSmartArtOutlineItem } from "@/lib/office-smartart/native-outline";
import {
  readSmartArtIdentityRecord,
  type SmartArtIdentity,
} from "./office-smartart/smartart-identity";
import {
  MAX_SMARTART_OUTLINE_ITEMS,
  MAX_SMARTART_OUTLINE_TEXT_LENGTH,
  type SmartArtOutlineItem,
} from "./presentation-smartart";

type RecordValue = Record<string, unknown>;
const record = (v: unknown): RecordValue | null =>
  v && typeof v === "object" && !Array.isArray(v) ? (v as RecordValue) : null;
const pointId = (v: unknown): v is string =>
  typeof v === "string" && /^[A-Za-z0-9{}_-]{1,128}$/.test(v);
export interface NativeSmartArtGroup {
  groupId: string;
  members: readonly PresentationBlock[];
  receipt: RecordValue;
  outline: readonly SmartArtOutlineItem[];
  selectedPointId: string | null;
}

/**
 * ONE preserved native diagram, as the conversion reads it.
 *
 * This is deliberately weaker than {@link NativeSmartArtGroup}: it needs only
 * the drawing view every projected member carries, not the text-edit receipt,
 * because a diagram whose points are empty placeholders projects no text member
 * at all and still has to be convertible.
 */
export interface NativeSmartArtView {
  groupId: string;
  members: readonly PresentationBlock[];
  /** The graphic frame this diagram is; the replacement declares it on save. */
  frameShapeId: string;
  /**
   * The DiagramML layout part's `uniqueId` (`office-smartart/native-layouts`),
   * or the stocked layout it is a second spelling of (`layoutEquivalent`).
   */
  layoutIdentity: string;
  /**
   * The layout part this diagram's own file carries, by address and digest,
   * when MRAI stocks no definition of its layout (`layoutDefinition`, reading
   * row 370, SMARTART-GAP.md §8.6): the conversion lays it out from that part.
   */
  layoutDefinition?: { readonly part: string; readonly sha256: string } | null;
  frameEmu: NativeSmartArtFrame;
  outline: readonly NativeSmartArtOutlineItem[];
  /** The diagram's quick style, as the projection resolved it. */
  smartArtStyle: string | null;
  /**
   * MRAI's OWN identity record, when this diagram is one MRAI exported as a
   * real SmartArt and the projection could prove it whole. The two document
   * hosts read the same record out of the data part they hold; a slide's view
   * does not carry the part, so the projection ships the record itself
   * (`smart_art_identity_record`). Absent for every Office-made diagram.
   */
  identity: SmartArtIdentity | null;
  /**
   * The five parts this diagram is, and their fingerprints, exactly as the
   * projection recorded them (`NativeSmartArtProjection.parts` /
   * `.fingerprints`). Carried so 转换为可编辑 can ASK the page's own source
   * package for a pictured diagram's slots — the media never rides on the
   * stored page (SMARTART-IMPORT.md §5.24: four slots inline is 80% of a
   * deck), so the address has to travel instead. The two document hosts hold
   * the same two records on their own preserved nodes; this is the slide
   * host's copy of the same fact, passed to the SHARED reader rather than
   * interpreted here.
   */
  slotMediaView: Readonly<Record<string, unknown>> | null;
  /**
   * Office's `dgm:dir` for this diagram, `rev` when it runs Right to Left, as
   * the projection resolved it off the data part (`layoutDirection`). The
   * projection states it beside MRAI's own identity record (a diagram MRAI
   * exported mirrored), and the conversion on open builds it mirrored again;
   * absent is Office's default.
   */
  layoutDirection: "rev" | null;
  /**
   * The diagram's colour variation (Office's id, `accent1_3`, `colorful2`,
   * …) as the projection read it off the data part's `csTypeId`
   * (`colorsIdentity`), so the conversion on open paints it the way
   * PowerPoint does and keeps it. Null when the projection named none.
   */
  colorsId: string | null;
  /**
   * The theme that colour variation resolves against when the diagram's
   * slide master has a theme of its own (not the first slide's, the one the
   * page holds): dk1, lt1, dk2, lt2 and accent1..6 as the projection read
   * them (`themeColors`). Null on a deck of one theme.
   */
  themeColors: Readonly<Record<string, string>> | null;
  /**
   * The diagram's master theme's fill and effect style lists, as the
   * projection read them (`hostTheme`, reading `smartArtViewHostTheme`, on
   * the diagram's first member): what its quick style's `a:fillRef` /
   * `a:effectRef` draw while the deck still wears that theme. Null when the
   * projection named none.
   */
  hostTheme?: { readonly effects?: string; readonly xml: string } | null;
  





  unkeptAuthorFormat: readonly string[];
  /**
   * The author's hand changes on the diagram's shapes (prSet cust*, custT
   * sizes) as the projection read them (`customizations`), handed to the
   * conversion unread: the engine checks them (`nativeSmartArtCustomizations`).
   */
  customizations: unknown;
  /**
   * The graphic's OWN fill, line and effect (`dgm:bg`, `dgm:whole`) as the
   * projection read them (`graphicFormat`, reading `smartArtGraphicFormat`):
   * the converted diagram's `smartArtBackground`. Null when it states none.
   */
  graphicFormat?: unknown;
  /**
   * The graphic's OWN 3-D Rotation (its document point's `a:scene3d`, reading
   * `smartArtViewGraphicCamera`), which the conversion draws as one scene
   * over its SmartArt Style's. Null when it states none.
   */
  graphicScene3d?: unknown;
}

const THEME_COLOR_SLOTS = ["dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6"] as const;

function readThemeColors(value: unknown): Readonly<Record<string, string>> | null {
  const colors = record(value);
  if (!colors) return null;
  const out: Record<string, string> = {};
  for (const slot of THEME_COLOR_SLOTS) {
    const color = colors[slot];
    if (typeof color !== "string" || !/^#[0-9a-f]{6}$/i.test(color)) return null;
    out[slot] = color.toLowerCase();
  }
  return out;
}

// The frame's offsets are signed (`ST_Coordinate`): a diagram may sit above

// a frame this reader accepts is one 转换为可编辑 accepts too.
function frameRect(value: unknown): NativeSmartArtFrame | null {
  const rect = record(value);
  if (!rect) return null;
  const frame = {
    left: rect.left,
    top: rect.top,
    width: rect.width,
    height: rect.height,
  };
  return nativeSmartArtFrameInBounds(frame) ? frame : null;
}

function viewOutline(value: unknown): NativeSmartArtOutlineItem[] | null {
  if (
    !Array.isArray(value) ||
    value.length === 0 ||
    value.length > MAX_SMARTART_OUTLINE_ITEMS
  ) {
    return null;
  }
  const items: NativeSmartArtOutlineItem[] = [];
  const ids = new Set<string>();
  for (const entry of value) {
    const item = record(entry);
    const id = item?.id;
    const text = item?.text;
    const level = item?.level;
    if (
      !item ||
      !pointId(id) ||
      ids.has(id) ||
      typeof text !== "string" ||
      text.length > MAX_SMARTART_OUTLINE_TEXT_LENGTH ||
      typeof level !== "number" ||
      !Number.isInteger(level) ||
      level < 0 ||
      level > MAX_SMARTART_OUTLINE_LEVEL ||
      (items.length === 0
        ? level !== 0
        : level > items[items.length - 1].level + 1)
    ) {
      return null;
    }
    ids.add(id);
    // A slot's picture relationship travels WITH the item. Dropping it here
    // was how a pictured Office diagram converted on this host into empty
    // slots: the outline reached `convertNativeSmartArt` claiming no pictures,
    // so nothing could fill a slot and nothing could refuse to
    // (`smartArtPictureLayoutConvertible` keys off exactly this field). The
    // two document hosts carry it already; the slide host now does too
    // (SMARTART-IMPORT.md §5.28).
    const imageRelId = item.imageRelId;
    // Words on the item's transitions (Name and Title Organization Chart's
    // titles): carried, so the conversion knows it cannot keep them.
    const transitionWords = (value: unknown) =>
      typeof value === "string" && value.trim() && value.length <= MAX_SMARTART_OUTLINE_TEXT_LENGTH ? value : undefined;
    items.push({
      ...(transitionWords(item.transitionText) ? { transitionText: transitionWords(item.transitionText) } : {}),
      ...(transitionWords(item.parentTransitionText)
        ? { parentTransitionText: transitionWords(item.parentTransitionText) }
        : {}),
      id,
      text,
      level,
      ...(item.assistant === true && level > 0 ? { assistant: true as const } : {}),
      ...(item.textLanguage === "other" ? { textLanguage: "other" as const } : {}),
      ...(readSmartArtCustom(item.custom) ? { custom: readSmartArtCustom(item.custom)! } : {}),
      ...(typeof imageRelId === "string" &&
      imageRelId.length > 0 &&
      imageRelId.length <= 128
        ? { imageRelId }
        : {}),
      // ...and whether that picture is only the item's transition's (row 426).
      ...(typeof imageRelId === "string" && imageRelId && item.imageOnTransition === true ? { imageOnTransition: true as const } : {}),
      // The point's OWN paint and shape (SmartArt Format on one shape, in
      // Office or in MRAI's export): an answer, not a default, so the
      // conversion keeps them on the item's node (PPT-A010). Dropping them
      // here is how a recoloured or reshaped slide shape came back in the
      // layout's own on reopen.
      ...(paintValue(item.fill) ? { fill: paintValue(item.fill)! } : {}),
      // Paint in a THEME colour (row 212): resolved against the deck's theme
      // by the conversion, as the document hosts' reader of the same part
      
      ...(schemeValue(item.fillScheme) ? { fillScheme: schemeValue(item.fillScheme)! } : {}),
      ...(schemeValue(item.outlineScheme) ? { outlineScheme: schemeValue(item.outlineScheme)! } : {}),
      ...(schemeValue(item.textColorScheme) ? { textColorScheme: schemeValue(item.textColorScheme)! } : {}),
      ...(paintValue(item.outline) ? { outline: paintValue(item.outline)! } : {}),
      ...(hexColour(item.outline) &&
      typeof item.outlineWidthPt === "number" &&
      Number.isFinite(item.outlineWidthPt) &&
      item.outlineWidthPt >= 0 &&
      item.outlineWidthPt <= 1584
        ? { outlineWidthPt: item.outlineWidthPt }
        : {}),
      ...(isShapeKind(item.shape) ? { shape: item.shape } : {}),
      // The point's own effects and WordArt (PPT-A010), bounded by the
      // shared readers; a value they refuse is simply not carried.
      ...authorFormat(item.format),
      ...(paintValue(item.fill) &&
      typeof item.fillTransparency === "number" &&
      Number.isInteger(item.fillTransparency) &&
      item.fillTransparency >= 1 &&
      item.fillTransparency <= 100
        ? { fillTransparency: item.fillTransparency }
        : {}),
      ...textPaintOf(item.textPaint),
      // The point's words as Office formats them (its runs' face, weight and
      // slant; the size where Office marks it the author's, `custT`; the
      // colour): the conversion puts them on the label, the export writes
      
      // and saved plain.
      ...textFormatOf(item, text),
    });
  }
  return items;
}

function textFormatOf(
  item: Record<string, unknown>,
  text: string,
): Pick<NativeSmartArtOutlineItem, "textFont" | "textSizePt" | "textColor" | "textRuns" | "textAlign" | "textLineSpacing" | "textCapitals"> {
  const font = normalizeOfficeShapeTextFont(item.textFont);
  const size = item.textSizePt;
  const color = hexColour(item.textColor);
  const runs = normalizeRuns(item.textRuns);
  return {
    ...(Object.keys(font).length ? { textFont: font } : {}),
    ...(typeof size === "number" && Number.isFinite(size) && size >= 1 && size <= 4000 ? { textSizePt: size } : {}),
    ...(color ? { textColor: color } : {}),
    ...(runs && runs.map((run) => run.text).join("") === text ? { textRuns: runs } : {}),
    ...(item.textAlign === "l" || item.textAlign === "ctr" || item.textAlign === "r" || item.textAlign === "just"
      ? { textAlign: item.textAlign }
      : {}),
    // The author's line spacing (row 380), a multiple of a single line.
    ...(typeof item.textLineSpacing === "number" && Number.isFinite(item.textLineSpacing) &&
    item.textLineSpacing >= 0.05 && item.textLineSpacing <= 9.6
      ? { textLineSpacing: item.textLineSpacing }
      : {}),
    // The author's capitals (`a:defRPr@cap`).
    ...(item.textCapitals === true ? { textCapitals: true as const } : {}),
  };
}

function authorFormat(value: unknown): { format?: SmartArtAuthorFormat } {
  if (!value || typeof value !== "object" || Array.isArray(value)) return {};
  const record = value as Record<string, unknown>;
  const format: SmartArtAuthorFormat = {};
  const effects = normalizeOfficePictureEffects(record.shapeEffects);
  if (effects) format.shapeEffects = effects;
  const effect = normalizeOfficeShapeEffectPreset(record.effect);
  if (effect) format.effect = effect;
  const shadow = normalizeOfficeShapeShadow(record.shapeShadow);
  if (shadow) format.shapeShadow = shadow;
  const gradient = normalizeOfficeShapeGradient(record.fillGradient);
  if (gradient) format.fillGradient = gradient;
  return { format };
}

function textPaintOf(value: unknown): { textPaint?: OfficeTextPaint } {
  try {
    const paint = normalizeOfficeTextPaint(value);
    return paint ? { textPaint: paint } : {};
  } catch {
    return {};
  }
}

/** A point's own colour, or its No Fill / No Outline (PPT-A010). */
const paintValue = (value: unknown): string | null =>
  value === "none" ? "none" : hexColour(value);

const hexColour = (value: unknown): string | null =>
  typeof value === "string" && /^#[0-9a-f]{6}$/i.test(value) ? value.toLowerCase() : null;

/**
 * Read the diagram the given block belongs to, or `null` when the block is not
 * part of one complete, self-consistent native SmartArt projection.
 */
/**
 * The ids a diagram's first-edit conversion gives its blocks: the same in
 * every window (`idPrefix` + the layout's own count), from its slide and its
 * frame, so two people converting it at once make one set of members. A
 * name the id grammar cannot carry, or a long one, is hashed.
 */
export function nativeSmartArtConversionIdPrefix(slideId: string, frameShapeId: unknown): string {
  const raw = `${slideId}-${String(frameShapeId ?? "")}`;
  if (/^[A-Za-z0-9_-]{1,96}$/.test(raw)) return `smartart-${raw}-`;
  let hash = 0x811c9dc5;
  for (let i = 0; i < raw.length; i += 1) {
    hash ^= raw.charCodeAt(i);
    hash = Math.imul(hash, 0x01000193) >>> 0;
  }
  return `smartart-${hash.toString(16).padStart(8, "0")}-${raw.length}-`;
}

/** A theme colour an outline item states (`{ scheme, mods }`), or `undefined`. */
function schemeValue(value: unknown): SmartArtSchemeColor | undefined {
  const colour = record(value);
  const scheme = colour?.scheme;
  const mods = colour?.mods;
  if (
    typeof scheme !== "string" ||
    !/^(dk1|lt1|dk2|lt2|tx1|bg1|tx2|bg2|accent[1-6])$/.test(scheme) ||
    !Array.isArray(mods) ||
    mods.length > 12 ||
    !mods.every(
      (mod) =>
        Array.isArray(mod) &&
        mod.length === 2 &&
        typeof mod[0] === "string" &&
        /^[a-zA-Z]{1,24}$/.test(mod[0]) &&
        Number.isInteger(mod[1]) &&
        Math.abs(mod[1]) <= 999_999_999,
    )
  ) {
    return undefined;
  }
  return { scheme, mods: mods.map((mod) => [mod[0] as string, mod[1] as number] as const) };
}

export function readNativeSmartArtView(
  blocks: readonly PresentationBlock[],
  primary: PresentationBlock | null | undefined,
): NativeSmartArtView | null {
  const groupId = primary?.groupId;
  if (typeof groupId !== "string" || !groupId) return null;
  const members = blocks.filter((block) => block.groupId === groupId);
  if (
    !members.length ||
    members.length > 200 ||
    new Set(members.map((block) => block.id)).size !== members.length
  ) {
    return null;
  }
  const first = record(record(members[0].source)?.nativeSmartArtView);
  const frameShapeId = first?.frameShapeId;
  const fileLayoutIdentity = first?.layoutIdentity;
  // A second spelling of a stocked layout (PowerPoint's numbered
  // `…/vList3#1`, Plus and Minus without its trailing space) whose part the
  // import proved IS that layout's definition: converted and edited as that
  // layout (`equivalent_layout_identity`, reading row 362). Only a layout the
  // catalog knows is taken; anything else is the file's own urn.
  const equivalent = first?.layoutEquivalent;
  const layoutIdentity =
    typeof equivalent === "string" && nativeLayoutToCatalogId(equivalent) !== null
      ? equivalent
      : fileLayoutIdentity;
  const frameEmu = frameRect(first?.frameEmu);
  const outline = viewOutline(first?.outline);
  if (
    !first ||
    first.schemaVersion !== "pptx-smartart-drawing-v1" ||
    typeof frameShapeId !== "string" ||
    !/^shp[1-9][0-9]*$/.test(frameShapeId) ||
    typeof layoutIdentity !== "string" ||
    !layoutIdentity ||
    !frameEmu ||
    !outline
  ) {
    return null;
  }
  // Every member of one diagram repeats the diagram's own identity; a group
  // whose members disagree is not one this conversion may consume.
  for (const block of members) {
    const view = record(record(block.source)?.nativeSmartArtView);
    if (
      !view ||
      view.frameShapeId !== frameShapeId ||
      view.layoutIdentity !== fileLayoutIdentity ||
      view.layoutEquivalent !== first.layoutEquivalent ||
      JSON.stringify(view.layoutDefinition ?? null) !== JSON.stringify(first.layoutDefinition ?? null) ||
      JSON.stringify(view.frameEmu) !== JSON.stringify(first.frameEmu) ||
      JSON.stringify(view.outline) !== JSON.stringify(first.outline)
    ) {
      return null;
    }
  }
  const styles = new Set(
    members.map((block) => (block as unknown as RecordValue).smartArtStyle),
  );
  const style = (members[0] as unknown as RecordValue).smartArtStyle;
  return {
    groupId,
    members,
    frameShapeId,
    layoutIdentity,
    layoutDefinition: ownLayoutDefinition(first.layoutDefinition),
    frameEmu,
    outline,
    smartArtStyle: styles.size === 1 && typeof style === "string" ? style : null,
    // Every member repeats the diagram's record, exactly as it repeats the
    // outline; a group whose members disagree about what the diagram IS is not
    // one this conversion may consume, so the record is read off the first and
    // proven identical on the rest.
    identity: members.every(
      (block) =>
        JSON.stringify(
          record(record(block.source)?.nativeSmartArtView)?.smartArtIdentity ??
            null,
        ) === JSON.stringify(first.smartArtIdentity ?? null),
    )
      ? readSmartArtIdentityRecord(first.smartArtIdentity)
      : null,
    // Handed over whole, never interpreted here: `smartArtSlotMediaAddress`
    // is the ONE reader of this pair and all three hosts call it.
    slotMediaView:
      record(first.parts) && record(first.fingerprints)
        ? { parts: first.parts, fingerprints: first.fingerprints }
        : null,
    layoutDirection: first.layoutDirection === "rev" ? "rev" : null,
    colorsId: smartArtColorVariantIdFromUrn(first.colorsIdentity),
    themeColors: readThemeColors(first.themeColors),
    hostTheme:
      nativeSmartArtHostTheme(
        members.map((block) => record(record(block.source)?.nativeSmartArtView)?.hostTheme).find((value) => value !== undefined),
      ) ?? null,
    unkeptAuthorFormat: Array.isArray(first.unkeptAuthorFormat)
      ? first.unkeptAuthorFormat.filter((kind): kind is string => typeof kind === "string").slice(0, 8)
      : [],
    customizations: first.customizations ?? null,
    graphicFormat: first.graphicFormat ?? null,
    graphicScene3d: first.graphicScene3d ?? null,
  };
}

/** A view's `layoutDefinition`, when it is a well-formed address and digest. */
function ownLayoutDefinition(value: unknown): { part: string; sha256: string } | null {
  const definition = record(value);
  return definition &&
    typeof definition.part === "string" &&
    definition.part.length > 0 &&
    definition.part.length <= 512 &&
    typeof definition.sha256 === "string" &&
    /^[0-9a-f]{64}$/.test(definition.sha256)
    ? { part: definition.part, sha256: definition.sha256 }
    : null;
}

/** Whether a block is a member of a file's SmartArt drawing (its native view). */
export function isNativeSmartArtMember(block: unknown): boolean {
  return record(record(record(block)?.source)?.nativeSmartArtView)?.schemaVersion === "pptx-smartart-drawing-v1";
}

/**
 * Whether a block is the one frame member the importer keeps for a diagram
 * whose cached drawing is absent or empty (`drawingCache: "absent"`): the
 * diagram is drawn as the engine lays it out, and this frame is the object.
 */
export function nativeSmartArtDrawingAbsent(block: unknown): boolean {
  return record(record(record(block)?.source)?.nativeSmartArtView)?.drawingCache === "absent";
}

/**
 * The shape of a diagram's conversion whose words a double click on the file's
 * own drawing opens: the item of the topmost drawn member under the point
 * (slide units), as its data points name it, and that item's one converted
 * node. PowerPoint opens the words of the shape the author double-clicked;
 * hit again on the conversion, which is laid out anew and drawn over the same
 * place, the press landed on whatever the conversion has there -- a column
 * behind the words (Horizontal Bullet List), the node's own outline (a label
 * as large as its node) -- and opened nothing. Null where the drawing has no
 * single item there or the conversion no single node for it.
 */
export function nativeSmartArtConvertedShapeAt(
  view: Pick<NativeSmartArtView, "members" | "outline">,
  converted: readonly PresentationBlock[],
  point: { x: number; y: number },
): string | null {
  for (let at = view.members.length - 1; at >= 0; at -= 1) {
    const member = view.members[at] as unknown as Record<string, unknown>;
    const { left, top, width, height } = member;
    if (
      typeof left !== "number" || typeof top !== "number" ||
      typeof width !== "number" || typeof height !== "number" ||
      point.x < left || point.x > left + width || point.y < top || point.y > top + height
    ) {
      continue;
    }
    const node = convertedNodeOfMember(view, converted, member);
    if (node === undefined) continue;
    return node;
  }
  return null;
}

/**
 * The conversion's shape standing for the drawing's member `memberId` (its
 * item's node), whose words open where that member's were open: null when
 * the member stands for no single item.
 */
export function nativeSmartArtConvertedShapeOf(
  view: Pick<NativeSmartArtView, "members" | "outline">,
  converted: readonly PresentationBlock[],
  memberId: string,
): string | null {
  const member = view.members.find((block) => block.id === memberId);
  return member ? convertedNodeOfMember(view, converted, member as unknown as Record<string, unknown>) ?? null : null;
}

/** The node of `member`'s item in the conversion; undefined when it names no
 * item of the outline, null when the conversion has no single node for it. */
function convertedNodeOfMember(
  view: Pick<NativeSmartArtView, "outline">,
  converted: readonly PresentationBlock[],
  member: Record<string, unknown>,
): string | null | undefined {
  const ids = record(record(member.source)?.nativeSmartArtView)?.dataPointIds;
  const index = Array.isArray(ids) && typeof ids[0] === "string"
    ? view.outline.findIndex((item) => item.id === ids[0])
    : -1;
  if (index < 0) return undefined;
  const nodes = converted.filter((block) => {
    const own = block as { smartArtItemIndex?: unknown; smartArtRole?: unknown };
    return block.type === "shape" && own.smartArtItemIndex === index && own.smartArtRole === "node";
  });
  return nodes.length === 1 ? nodes[0].id : null;
}

/**
 * Whether a diagram is laid out from its own file's definition when it
 * converts: a layout MRAI does not own (no catalog id, alias or second
 * spelling) whose view names that part, drawn in the file (there is a
 * PowerPoint drawing to hold the result to).
 */
export function nativeSmartArtConvertsFromOwnDefinition(
  view: Pick<NativeSmartArtView, "layoutIdentity" | "layoutDefinition" | "members">,
): boolean {
  return (
    !!view.layoutDefinition &&
    nativeLayoutToCatalogId(view.layoutIdentity) === null &&
    !view.members.some(nativeSmartArtDrawingAbsent)
  );
}















export function nativeSmartArtConversionAtDrawnBoxes(
  view: Pick<NativeSmartArtView, "members" | "outline" | "customizations">,
  converted: readonly PresentationBlock[],
): PresentationBlock[] {
  if (!view.customizations) return [...converted];
  const indexById = new Map(view.outline.map((item, index) => [item.id, index] as const));
  type Box = { left: number; top: number; width: number; height: number };
  const drawn = new Map<string, Box | null>();
  for (const member of view.members) {
    const native = record(record(member.source)?.nativeSmartArtView);
    const ids = native?.dataPointIds;
    const role = native?.role === "shape" ? "node" : native?.role === "text" ? "label" : null;
    if (!role || !Array.isArray(ids) || ids.length !== 1 || typeof ids[0] !== "string") continue;
    const index = indexById.get(ids[0]);
    const { left, top, width, height } = member as unknown as Record<string, unknown>;
    if (
      index === undefined ||
      typeof left !== "number" || typeof top !== "number" ||
      typeof width !== "number" || typeof height !== "number" || width <= 0 || height <= 0
    ) {
      continue;
    }
    const key = `${role}:${index}`;
    drawn.set(key, drawn.has(key) ? null : { left, top, width, height });
  }
  const keyOf = (block: PresentationBlock): string | null => {
    const index = (block as { smartArtItemIndex?: unknown }).smartArtItemIndex;
    const role = (block as { smartArtRole?: unknown }).smartArtRole;
    return typeof index === "number" && (role === "node" || role === "label") ? `${role}:${index}` : null;
  };
  const converts = new Map<string, number>();
  for (const block of converted) {
    const key = keyOf(block);
    if (key) converts.set(key, (converts.get(key) ?? 0) + 1);
  }
  return converted.map((block) => {
    const key = keyOf(block);
    const box = key && converts.get(key) === 1 ? drawn.get(key) : null;
    return box ? ({ ...block, ...box } as PresentationBlock) : block;
  });
}

/**
 * Whether a diagram converted as the deck OPENS still looks like the drawing
 * PowerPoint saved for it. An author's hand changes the engine does not
 * reproduce (a dragged or resized shape) would turn it into another picture,
 * so the open keeps such a diagram as PowerPoint drew it; 转换为可编辑 still
 * converts it when the author asks.
 */
export function nativeSmartArtConversionLooksLikeDrawing(
  view: Pick<NativeSmartArtView, "members" | "frameEmu">,
  converted: readonly unknown[],
): boolean {
  return smartArtConversionLooksLikeDrawing(
    smartArtBlockDrawingBoxes(view.members),
    smartArtBlockDrawingBoxes(converted),
    view.frameEmu.width,
  );
}

/** How far a converted diagram stands from PowerPoint's drawing of it (mean / worst box error of the frame width), or `null`. */
export function nativeSmartArtConversionLikeness(
  view: Pick<NativeSmartArtView, "members" | "frameEmu">,
  converted: readonly unknown[],
): { mean: number; worst: number } | null {
  const likeness = smartArtDrawingLikeness(
    smartArtBlockDrawingBoxes(view.members),
    smartArtBlockDrawingBoxes(converted),
    view.frameEmu.width,
  );
  return likeness ? { mean: likeness.mean, worst: likeness.worst } : null;
}

/** Every complete native diagram on one slide, in first-member order. */
export function listNativeSmartArtViews(
  blocks: readonly PresentationBlock[],
): NativeSmartArtView[] {
  const seen = new Set<string>();
  const views: NativeSmartArtView[] = [];
  for (const block of blocks) {
    const groupId = block.groupId;
    if (typeof groupId !== "string" || !groupId || seen.has(groupId)) continue;
    seen.add(groupId);
    const view = readNativeSmartArtView(blocks, block);
    if (view) views.push(view);
  }
  return views;
}

function textSlots(block: PresentationBlock): string[] | null {
  if (block.type !== "text") return null;
  if (Array.isArray(block.paragraphs) && block.paragraphs.length) {
    const values = block.paragraphs.map((p) => p.text);
    return values.every((v) => typeof v === "string") &&
      values.join("\n") === block.text
      ? values
      : null;
  }
  return typeof block.text === "string" ? [block.text] : null;
}

export function readNativeSmartArtGroup(
  blocks: readonly PresentationBlock[],
  primary: PresentationBlock | null | undefined,
): NativeSmartArtGroup | null {
  if (!primary || typeof primary.groupId !== "string") return null;
  const source = record(primary.source),
    receipt = record(source?.nativeSmartArtEdit);
  if (
    !receipt ||
    receipt.schemaVersion !== "pptx-smartart-edit-v2" ||
    !Array.isArray(receipt.points) ||
    receipt.points.length < 1 ||
    receipt.points.length > 20 ||
    !/^shp[1-9][0-9]*$/.test(String(receipt.frameShapeId)) ||
    !/^ppt\/slides\/slide[1-9][0-9]*\.xml$/.test(String(receipt.slidePart))
  )
    return null;
  const parts = record(receipt.parts),
    hashes = record(receipt.fingerprints);
  const keys = ["data", "layout", "style", "colors", "drawing"];
  if (
    !parts ||
    !hashes ||
    keys.some(
      (key) =>
        !/^ppt\/diagrams\/[A-Za-z0-9_-]+\.xml$/.test(String(parts[key])) ||
        !/^[a-f0-9]{64}$/.test(String(hashes[key])),
    )
  )
    return null;
  const members = blocks.filter((b) => b.groupId === primary.groupId);
  if (
    !members.length ||
    members.length > 200 ||
    new Set(members.map((b) => b.id)).size !== members.length
  )
    return null;
  const original = new Map<string, { text: string; level: number }>();
  for (const value of receipt.points) {
    const p = record(value);
    if (
      !p ||
      !pointId(p.modelId) ||
      original.has(p.modelId) ||
      typeof p.text !== "string" ||
      p.text.length > 2000 ||
      typeof p.level !== "number" ||
      !Number.isInteger(p.level) ||
      p.level < 0 ||
      p.level > 20
    )
      return null;
    original.set(p.modelId, { text: p.text, level: p.level });
  }
  const current = new Map<string, string>();
  for (const block of members) {
    const s = record(block.source),
      view = record(s?.nativeSmartArtView);
    if (
      !view ||
      view.frameShapeId !== receipt.frameShapeId ||
      JSON.stringify(s?.nativeSmartArtEdit) !== JSON.stringify(receipt) ||
      JSON.stringify(view.parts) !== JSON.stringify(parts) ||
      JSON.stringify(view.fingerprints) !== JSON.stringify(hashes)
    )
      return null;
    if (view.role !== "text") continue;
    let values = textSlots(block);
    const ids = view.dataPointIds;
    // One point of several paragraphs drawn in one member: its text is the
    // member's paragraphs, line by line (the save's text channel maps it row
    
    if (
      values &&
      Array.isArray(ids) &&
      ids.length === 1 &&
      values.length > 1 &&
      typeof ids[0] === "string" &&
      original.get(ids[0])?.text.includes("\n")
    )
      values = [values.join("\n")];
    if (!values || !Array.isArray(ids) || ids.length !== values.length)
      return null;
    for (let i = 0; i < ids.length; i++) {
      if (
        !pointId(ids[i]) ||
        !original.has(ids[i]) ||
        values[i].length > 2000 ||
        (current.has(ids[i]) && current.get(ids[i]) !== values[i])
      )
        return null;
      current.set(ids[i], values[i]);
    }
  }
  if (current.size !== original.size) return null;
  const selectedIds = record(source?.nativeSmartArtView)?.dataPointIds;
  return {
    groupId: primary.groupId,
    members,
    receipt,
    outline: [...original].map(([id, p]) => ({
      id,
      text: current.get(id)!,
      level: p.level,
    })),
    selectedPointId:
      Array.isArray(selectedIds) &&
      pointId(selectedIds[0]) &&
      original.has(selectedIds[0])
        ? selectedIds[0]
        : null,
  };
}

/**
 * The run texts that make ``runs`` (joined: ``expected``) read ``text``, each
 * run keeping its own formatting: the words before and after the change stay
 * in their runs, the run the change starts in takes the new words, runs
 * wholly inside it are emptied -- what typing over a selection does in
 * PowerPoint, and what the save writes into the package (`_span_edit`).
 */
export function nativeSmartArtRunSpanEdit(
  runs: readonly string[],
  expected: string,
  text: string,
): string[] | null {
  if (runs.join("") !== expected) return null;
  let prefix = 0;
  const limit = Math.min(expected.length, text.length);
  while (prefix < limit && expected[prefix] === text[prefix]) prefix++;
  let suffix = 0;
  while (
    suffix < limit - prefix &&
    expected[expected.length - 1 - suffix] === text[text.length - 1 - suffix]
  )
    suffix++;
  const start = prefix,
    end = expected.length - suffix,
    middle = text.slice(prefix, text.length - suffix);
  const result: string[] = [];
  let offset = 0,
    placed = false;
  runs.forEach((value, index) => {
    const lo = offset,
      hi = offset + value.length;
    offset = hi;
    if (hi < start || lo > end || (hi === start && start !== end && index + 1 < runs.length)) {
      result.push(value);
      return;
    }
    const head = lo <= start ? value.slice(0, Math.max(0, start - lo)) : "";
    const tail = hi >= end ? value.slice(Math.max(0, end - lo)) : "";
    if (!placed && lo <= start && start <= hi) {
      result.push(head + middle + tail);
      placed = true;
    } else result.push(head + tail);
  });
  return placed && result.join("") === text ? result : null;
}

type NativeTextRun = { text?: string } & Record<string, unknown>;

function withRunText<T extends { text?: unknown; runs?: NativeTextRun[] }>(row: T, text: string): T {
  const runs = row.runs;
  if (!runs) return { ...row, text };
  const spans = nativeSmartArtRunSpanEdit(
    runs.map((run) => String(run.text ?? "")),
    String(row.text ?? ""),
    text,
  );
  return {
    ...row,
    text,
    runs: runs.map((run, j) => ({ ...run, text: spans ? spans[j] : j === 0 ? text : "" })),
  };
}

export function updateNativeSmartArtText(
  blocks: readonly PresentationBlock[],
  group: NativeSmartArtGroup,
  id: string,
  text: string,
): PresentationBlock[] | null {
  const current = group.outline.find((item) => item.id === id);
  if (!current || current.text === text || text.length > 2000) return null;
  let changed = false;
  const updated = blocks.map((block) => {
    if (block.groupId !== group.groupId || block.type !== "text") return block;
    const view = record(record(block.source)?.nativeSmartArtView),
      ids = view?.dataPointIds;
    if (!Array.isArray(ids) || !ids.includes(id)) return block;
    const values = textSlots(block);
    if (!values) return block;
    const next = { ...block };
    if (block.paragraphs?.length && ids.length === 1 && values.length > 1) {
      // One point of several paragraphs: line by line, its count kept (a new
      // or removed paragraph is a structural edit of the diagram).
      const parts = text.split("\n");
      if (parts.length !== block.paragraphs.length) return block;
      next.paragraphs = block.paragraphs.map((paragraph, i) =>
        withRunText(paragraph as typeof paragraph & { runs?: NativeTextRun[] }, parts[i]),
      );
      next.text = text;
    } else if (ids.length !== values.length) {
      return block;
    } else if (block.paragraphs?.length) {
      next.paragraphs = block.paragraphs.map((paragraph, i) =>
        ids[i] === id ? withRunText(paragraph as typeof paragraph & { runs?: NativeTextRun[] }, text) : paragraph,
      );
      next.text = next.paragraphs.map((p) => p.text).join("\n");
    } else {
      Object.assign(next, withRunText(block as typeof block & { runs?: NativeTextRun[] }, text));
    }
    changed = true;
    return next;
  });
  return changed ? updated : null;
}

/** A text block reading `text`, its paragraphs and runs keeping their formatting (`withRunText`); null if the words do not fit its paragraphs. */
function textBlockWithWords(block: PresentationBlock, text: string): PresentationBlock | null {
  type Row = { text?: unknown; runs?: NativeTextRun[] };
  const paragraphs = (block as { paragraphs?: unknown }).paragraphs;
  if (Array.isArray(paragraphs) && paragraphs.length) {
    const parts = text.split("\n");
    if (parts.length !== paragraphs.length) return null;
    return {
      ...block,
      text,
      paragraphs: (paragraphs as Row[]).map((paragraph, i) => withRunText(paragraph, parts[i])),
    } as PresentationBlock;
  }
  return withRunText(block as PresentationBlock & Row, text) as PresentationBlock;
}

/**
 * The words another window wrote on a file's diagram -- its preserved members
 * as they were (`before`) and as they arrived (`after`) -- carried onto this
 * window's conversion of it: each point whose words changed gives them to the
 * one converted block of its place in the diagram's `outline` (the place the
 * conversion's `smartArtItemIndex` names, as `nativeSmartArtWithWords` reads
 * it the other way). Null when they cannot be carried: a point with no single
 * block of words, or words whose lines its paragraphs cannot hold.
 */
export function nativeSmartArtWordsOntoConversion(
  converted: readonly PresentationBlock[],
  before: readonly PresentationBlock[],
  after: readonly PresentationBlock[],
  outline: readonly { id: string }[],
): PresentationBlock[] | null {
  const was = readNativeSmartArtGroup(before, before.find((block) => block.type === "text"));
  const now = readNativeSmartArtGroup(after, after.find((block) => block.type === "text"));
  if (!was || !now) return null;
  const previous = new Map(was.outline.map((item) => [item.id, item.text]));
  const out = [...converted];
  for (const item of now.outline) {
    if (!previous.has(item.id)) return null;
    if (previous.get(item.id) === item.text) continue;
    const index = outline.findIndex((entry) => entry.id === item.id);
    const holders = out.flatMap((block, at) =>
      (block as { smartArtItemIndex?: unknown }).smartArtItemIndex === index &&
      typeof (block as { text?: unknown }).text === "string" ? [at] : []);
    if (index < 0 || holders.length !== 1) return null;
    const written = textBlockWithWords(out[holders[0]], item.text);
    if (!written) return null;
    out[holders[0]] = written;
  }
  return out;
}

/**
 * A served answer's blocks (a SmartArt rebuild or restyle) with words another
 * window wrote while it was on its way: each outline place's words on the one
 * block of words its items name (`smartArtItemIndex`), runs keeping their
 * formatting; null when a place has no single such block or the words do not
 * fit its paragraphs.
 */
export function smartArtAnswerWithWords(
  blocks: readonly PresentationBlock[],
  words: ReadonlyMap<number, string>,
): PresentationBlock[] | null {
  const out = [...blocks];
  for (const [index, text] of words) {
    const holders = out.flatMap((block, at) =>
      (block as { smartArtItemIndex?: unknown }).smartArtItemIndex === index &&
      typeof (block as { text?: unknown }).text === "string" ? [at] : []);
    if (holders.length !== 1) return null;
    const written = textBlockWithWords(out[holders[0]], text);
    if (!written) return null;
    out[holders[0]] = written;
  }
  return out;
}
