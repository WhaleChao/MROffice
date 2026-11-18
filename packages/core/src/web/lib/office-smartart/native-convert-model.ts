/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * What a host asks BEFORE converting a native Office SmartArt: the frame it
 * accepts, whether a picture layout can convert with the pictures it has, the
 * quick style a urn names, and the sentence a refusal is told in. The
 * conversion itself (`./native-convert`) lays the diagram out and runs on the
 * server; the browser asks for it through `./remote`. `./native-convert`
 * re-exports everything here.
 */

import type { SmartArtColorTheme } from "./colors";
import { officeSmartArtLayout } from "./catalog";
import { MAX_SMARTART_OUTLINE_LEVEL } from "./model";
import { smartArtPictureSrc } from "./picture";
import type { NativeSmartArtOutlineItem } from "./native-outline";
import {
  MAX_SMARTART_OUTLINE_ITEMS,
  MAX_SMARTART_OUTLINE_TEXT_LENGTH,
} from "./outline";
import {
  DEFAULT_SMARTART_STYLE_ID,
  isSmartArtStyleId,
  SMARTART_STYLES,
  type SmartArtStyleId,
} from "./styles-model";

/** The diagram's rectangle in the host's own EMU coordinates. */
export interface NativeSmartArtFrame {
  left: number;
  top: number;
  width: number;
  height: number;
}

export interface NativeSmartArtCustomizations {
  shapes?: Readonly<Record<string, Readonly<Record<string, number>>>>;
  fixedFontPt?: Readonly<Record<string, number>>;
}

/** The projection's customizations, checked; null when it states none it can use. */
export function nativeSmartArtCustomizations(value: unknown): NativeSmartArtCustomizations | null {
  if (!value || typeof value !== "object" || Array.isArray(value)) return null;
  const record = value as Record<string, unknown>;
  const shapes: Record<string, Record<string, number>> = {};
  const fixedFontPt: Record<string, number> = {};
  const rawShapes = record.shapes;
  if (rawShapes && typeof rawShapes === "object" && !Array.isArray(rawShapes)) {
    for (const [key, values] of Object.entries(rawShapes as Record<string, unknown>).slice(0, 512)) {
      if (typeof key !== "string" || key.length > 300 || !values || typeof values !== "object") continue;
      const entry: Record<string, number> = {};
      for (const [name, raw] of Object.entries(values as Record<string, unknown>)) {
        if (name.startsWith("cust") && typeof raw === "number" && Number.isFinite(raw) && Math.abs(raw) <= 1e8) entry[name] = raw;
      }
      if (Object.keys(entry).length) shapes[key] = entry;
    }
  }
  const rawFixed = record.fixedFontPt;
  if (rawFixed && typeof rawFixed === "object" && !Array.isArray(rawFixed)) {
    for (const [id, raw] of Object.entries(rawFixed as Record<string, unknown>).slice(0, 512)) {
      if (typeof raw === "number" && raw >= 1 && raw <= 4000) fixedFontPt[id] = raw;
    }
  }
  if (!Object.keys(shapes).length && !Object.keys(fixedFontPt).length) return null;
  return { ...(Object.keys(shapes).length ? { shapes } : {}), ...(Object.keys(fixedFontPt).length ? { fixedFontPt } : {}) };
}

export interface ConvertNativeSmartArtInput {
  /** The layout part's `uniqueId` urn (`./native-layouts` owns the table). */
  layoutIdentity: string;
  /**
   * The host draws layouts Office's own definition covers from it (slides):
   * the diagram is laid out in its own frame as Office lays it out, so the
   * converted drawing is PowerPoint's and an edit does not move it.
   */
  officeLayout?: boolean;
  /** The data model's outline (`./native-outline`, or the PPT view's copy). */
  outline: readonly NativeSmartArtOutlineItem[];
  /** Where the diagram sits in the host: slide EMU, inline extent, anchor. */
  frameEmu: NativeSmartArtFrame;
  /** The quick-style urn or its short id; anything else falls back to the default. */
  smartArtStyle?: unknown;
  /**
   * Which way the diagram runs: Office's `dgm:dir` layout variable (`rev` is
   * Right to Left, ISO/IEC 29500-1 §21.4.7.26), as the data part states it
   * (`readNativeSmartArtDataDirection`) or a projection carried it. The
   * converted diagram is built mirrored, exactly as Right to Left would have
   * drawn it; anything else is Office's default.
   */
  direction?: unknown;
  /** Fresh-id factory: `crypto.randomUUID` in the app, a counter in tests. */
  nextId: () => string;
  /**
   * The host's own insert placeholders, used wherever the Office point's text
   * is the empty `[文本]` placeholder. Same providers the host already passes
   * to `officeSmartArtSeed`, so a converted diagram reads exactly like an
   * inserted one instead of showing blank label boxes.
   */
  itemLabel?: (index: number) => string;
  rootLabel?: string;
  /**
   * Keep an empty Office point EMPTY instead of seeding it. An empty point is
   * Office's `[Text]` prompt (`dgm:prSet phldr="1"`), not content: PowerPoint
   * draws the prompt while editing and prints nothing. A host that converts on
   * open and then saves must not turn that prompt into words the author never
   * typed, because the next export writes them into the file as real text
   * (PPT-G013). The host draws its own non-printing prompt instead. When set,
   * `itemLabel` / `rootLabel` are not consulted.
   */
  keepEmptyText?: boolean;
  /**
   * The colour variation the file names for this diagram (the data part's
   * `csTypeId`, as Office's id) and the document's theme colours. Either one
   * given, the converted diagram is painted the way PowerPoint paints that
   * variation on that theme -- `accent1_2` when the file names none -- and a
   * variation other than the default is kept on its members
   * (`smartArtColors`). Both omitted, the catalog palette, as before.
   */
  colorsId?: string | null;
  colorTheme?: Partial<SmartArtColorTheme> | null;
  /**
   * The diagram's OWN pictures, by the layout node keyed to its document
   * point (`docPictures`, readings row 1005: Radial List's centre `visible`),
   * each resolved against `images` like an item's. Given, every one must land
   * on its slot with its picture, or the diagram is not converted.
   */
  docPictures?: import("./doc-pictures").NativeSmartArtDocPictures | null;
  /**
   * The author's hand changes on the diagram's shapes, as the projection read
   * them off the data part (`nativeSmartArtView.customizations`): `shapes`
   * keyed `presName|<Office point id>` (`doc`, `<id>:sib`, `<id>:par`),
   * `fixedFontPt` by Office point id. Only a layout drawn from Office's own
   * definition takes them.
   */
  customizations?: NativeSmartArtCustomizations | null;
  /**
   * The graphic's OWN 3-D Rotation (its document point's `a:scene3d`, as the
   * projection read it: `nativeSmartArtView.graphicScene3d`), drawn as one
   * scene over its SmartArt Style's (`./member-camera`).
   */
  graphicScene3d?: unknown;
  /** The shape kinds the host draws as presets (a document's few); omitted, all. */
  drawableShape?: ((kind: string) => boolean) | null;
  /** The workbook's / document's body face (`+mn-lt`), which sets how tall the labels' lines stand. */
  bodyFace?: string | null;
  /**
   * Optional `r:embed` → `data:` payloads for the outline's picture points. A
   * member whose relationship is missing keeps the layout's own placeholder,
   * exactly like an inserted picture layout before the author picks a file.
   */
  images?: Readonly<Record<string, string>>;
  /**
   * What MRAI's OWN identity record says this diagram is, when the diagram is
   * one MRAI exported and is reopening (`{@link nativeSmartArtIdentityOverride}`
   * proved it against this outline). The layout urn alone cannot say it: two
   * catalog layouts share Office's `…/layout/default`, a quick style is not
   * derivable from a data model, and item ids are MRAI's own naming.
   *
   * Absent — an Office-made diagram, or a record this conversion could not
   * prove — and everything is read from the urn exactly as before.
   */
  identity?: NativeSmartArtIdentityOverride | null;
  








  stretchToFrame?: boolean;
  /**
   * The catalog layout the CALLER has already decided — 转换为可编辑, where the
   * author picked the layout out of the gallery.
   *
   * It exists because the urn is a lossy spelling of a catalog id and every
   * host was round-tripping the author's pick through it: `card-grid` became
   * Office's `…/layout/default`, which `nativeLayoutToCatalogId` reads back as
   * `basic-list` (the two are ONE Office layout, SMARTART-IMPORT.md §1), so the
   * author asked for a card grid and got a basic list. A layout id the caller
   * names is an answer already given; nothing may launder it.
   *
   * The record still wins where there is one, because an automatic reopen has
   * no author pick to honour and the two are never both set.
   */
  catalogLayoutId?: string | null;
  /**
   * Is this group id already taken on the page the diagram is landing on?
   *
   * Only consulted for the group id an identity record names. Two Office
   * diagrams can carry the SAME record — copying a diagram inside Word or Excel
   * copies its data part verbatim — and two page groups sharing a group id are
   * one group to every host's shape model, so the second one keeps the group id
   * the conversion generated instead. The layout, the style and the item ids
   * are already in it; only the byte-stability of a second export is lost.
   */
  groupIdInUse?: (groupId: string) => boolean;
}

/**
 * The part of MRAI's identity record that a REOPEN can act on, once the record
 * has been proven to describe the outline being converted.
 */
export interface NativeSmartArtIdentityOverride {
  /** The catalog layout id the record names, resolved through the catalog. */
  readonly catalogId: string;
  /**
   * The group id the record names — the id the diagram's members shared when
   * MRAI wrote the file.
   *
   * Without it a reopen mints a fresh one (`buildSmartArt` calls `nextId` for
   * it), and the point ids the writer derives from it move with it, so
   * "export" and "reopen then export" of one unchanged page differ byte for
   * byte. On a SLIDE it costs more than bytes: the save that follows an
   * insert finds the diagram the file already holds by exactly this id
   * (SMARTART-IMPORT.md §5.10), so a reopened deck whose group id changed
   * falls back to writing shapes.
   */
  readonly groupId: string;
  /** The quick style the record names; absent when the record omitted it. */
  readonly smartArtStyle?: SmartArtStyleId;
  /**
   * One entry per outline item, in outline order — and `null` where the record
   * names no member for that item, which is an item OFFICE ADDED since MRAI
   * wrote the file. Such an item takes a fresh id and the level the data model
   * gives it; everything the record DOES name is still the diagram MRAI wrote.
   */
  readonly items: readonly ({
    readonly itemId: string;
    readonly level: number;
  } | null)[];
}










export const NATIVE_SMARTART_FRAME_BOUND_EMU = 100_000_000;
const EMU_LIMIT = NATIVE_SMARTART_FRAME_BOUND_EMU;








export const SMARTART_ITEM_PICTURE_FILLED_LAYOUTS: ReadonlySet<string> = new Set(["radial-list"]);

/**
 * May this outline become a CATALOG diagram of this layout, picture-wise?
 *
 * The one loss converting a picture layout can cause: a picture point whose
 * picture the host's member model cannot be given comes back as an empty slot,
 * and the diagram has quietly lost it. So a point that HAS a picture is
 * convertible only when the caller brought that picture — `images` resolves its
 * relationship to a payload the member models accept — and otherwise the
 * diagram stays PRESERVED, which is Office's own drawing, pictures and all.
 *
 * Host-neutral on purpose: the Docs group member and the worksheet shape object
 * hold a picture the same way and lose it the same way, and a rule written
 * twice is a rule that drifts. A caller that reads no media passes none and
 * gets the strict old answer — which is what keeps a host not yet taught to
 * read them from converting pictures away.
 */
export function smartArtPictureLayoutConvertible(
  layoutId: string,
  outline: readonly {
    readonly imageRelId?: string;
    readonly imageOnTransition?: true;
    readonly level?: number;
  }[],
  images?: Readonly<Record<string, string>>,
  transitionsDrawnAsLines = false,
): boolean {
  const entry = officeSmartArtLayout(layoutId);
  // A layout with no picture slot draws a point's picture nowhere MRAI can
  
  // p6) converted into the empty box, and the save rewrote the diagram
  // without the photo. Such a diagram stays PRESERVED too.
  // (Not for a picture only on a transition the layout draws as a line: Office's own
  
  // save keeps that transition point as the file has it, `smartart_data_merge`.)
  if (entry?.picture !== true) {
    return outline.every((item) => !item.imageRelId || (transitionsDrawnAsLines && item.imageOnTransition === true));
  }
  // A layout draws a picture at ONE outline level -- the depth its own
  // `<dgm:forEach axis="ch">` frames put the `blipPhldr` shape at -- so MRAI
  // draws a slot at that level and nowhere else (SMARTART-IMPORT.md §5.39,
  // §5.46). A file whose OTHER point carries a blip anyway -- Office keeps a
  // picture through a layout change, "unused text does not appear" and nor
  // does an unused picture -- is therefore a diagram MRAI would convert with
  // that picture quietly gone, so it stays PRESERVED instead: Office's own
  // drawing, every picture in it.
  const pictureLevel = entry.officePictureLevel;
  // ...and a layout may draw a picture for only the FIRST so many items at
  // that level -- Radial List's centre shape is the whole of its picture
  // (SMARTART-IMPORT.md §5.51). Counted over the items AT the picture level,
  // because that is what the layout's own `presOf ch/node cnt="1"` selects.
  const pictureItems = entry.officePictureItems;
  let atLevel = 0;
  return outline.every((item) => {
    const onLevel =
      pictureLevel === undefined ||
      typeof item.level !== "number" ||
      item.level === pictureLevel;
    const slot = onLevel && (pictureItems === undefined || atLevel < pictureItems);
    if (onLevel) atLevel += 1;
    return (
      !item.imageRelId ||
      (smartArtPictureSrc(images?.[item.imageRelId]) !== null && slot)
    );
  });
}

const STYLE_BY_URN: ReadonlyMap<string, SmartArtStyleId> = new Map(
  SMARTART_STYLES.map((entry) => [entry.urn, entry.id] as const),
);

/** The catalog style id for a quick-style urn or short id; unknown → `simple1`. */
export function nativeSmartArtStyleId(value: unknown): SmartArtStyleId {
  if (isSmartArtStyleId(value)) return value;
  return typeof value === "string"
    ? (STYLE_BY_URN.get(value) ?? DEFAULT_SMARTART_STYLE_ID)
    : DEFAULT_SMARTART_STYLE_ID;
}

/**
 * Is this a frame the converter (and every host that reads one) accepts: four
 * finite numbers, a positive extent, and both edges inside the signed bound?
 */
export function nativeSmartArtFrameInBounds(
  frame: unknown,
): frame is NativeSmartArtFrame {
  const rect = frame as NativeSmartArtFrame | null;
  return (
    !!rect &&
    typeof rect === "object" &&
    [rect.left, rect.top, rect.width, rect.height].every(
      (value) => typeof value === "number" && Number.isFinite(value),
    ) &&
    // Signed offsets: a frame above/left of the slide converts too (the
    // bound `presentation-native-smartart.ts` and the save path state).
    rect.left >= -EMU_LIMIT &&
    rect.top >= -EMU_LIMIT &&
    rect.width >= 1 &&
    rect.height >= 1 &&
    rect.left + rect.width <= EMU_LIMIT &&
    rect.top + rect.height <= EMU_LIMIT
  );
}

export function validNativeSmartArtOutline(
  outline: readonly NativeSmartArtOutlineItem[],
): outline is readonly NativeSmartArtOutlineItem[] {
  return (
    Array.isArray(outline) &&
    outline.length > 0 &&
    outline.length <= MAX_SMARTART_OUTLINE_ITEMS &&
    outline.every(
      (item, index) =>
        !!item &&
        typeof item.text === "string" &&
        item.text.length <= MAX_SMARTART_OUTLINE_TEXT_LENGTH &&
        Number.isInteger(item.level) &&
        item.level >= 0 &&
        item.level <= MAX_SMARTART_OUTLINE_LEVEL &&
        (index === 0 ? item.level === 0 : item.level <= outline[index - 1].level + 1),
    )
  );
}


/**
 * Why converting a diagram into a layout brings nothing back.
 *
 * `convertNativeSmartArt` answers `null` for several different things and the
 * three hosts all just returned, so 转换为可编辑 did nothing and said nothing —
 * the same defect §5.57 fixed one layer up, in the layer below it
 * (`SMARTART-IMPORT.md` §5.60). What reaches an author who pressed the
 * command is `no-scene`: the diagram cannot be redrawn in that layout at all
 * (no seed labels, a scene that builds empty, or one that will not fit its
 * frame).
 *
 * A layout's documented capacity is NOT a refusal: the items a layout does not
 * draw stay in the outline as hidden members, as Office keeps them ("unused
 * text does not appear, but remains available if you switch layouts").
 *
 * A layout MRAI does not own is not in this vocabulary: that diagram simply
 * stays preserved and the command was never offered for it.
 */
export type SmartArtConvertRefusalReason =
  | "no-scene"
  // A conversion kept as Office drew it, by why (SMARTART-GAP.md §8.6): its
  // drawing would differ from PowerPoint's (more than 0.06 of the frame), its
  // own file's layout definition could not be read or is not the one the page
  // recorded, the engine refused the definition, or the author set formatting
  // the conversion cannot keep.
  | "drawing-mismatch"
  | "definition-unavailable"
  | "definition-changed"
  | "engine-refused"
  | "unkept-author-format";

export interface SmartArtConvertRefusal {
  readonly reason: SmartArtConvertRefusalReason;
  /** `drawing-mismatch`: the mean and worst box error, as fractions of the frame width. */
  readonly likeness?: { readonly mean: number; readonly worst: number };
  /** `unkept-author-format`: what the author set (`presShapeFormat`, `runCapitals`, `paragraphAlign`, `paragraphSpacing`). */
  readonly kinds?: readonly string[];
}

/**
 * Why a conversion that came back empty came back empty: by default a scene
 * that could not be built; a host that knows better names the reason.
 */
export function smartArtConvertRefusal(
  reason: SmartArtConvertRefusalReason = "no-scene",
  detail: Omit<SmartArtConvertRefusal, "reason"> = {},
): SmartArtConvertRefusal {
  return { reason, ...detail };
}

const SMART_ART_AS_READ_FIELDS = [
  "shape", "fill", "fillGradient", "fillTransparency", "stroke", "strokeWidth", "strokeStyle", "effect",
  "shadow", "shapeShadow", "shapeEffects", "opacity", "shape3d", "scene3d", "style", "smartArtCustom",
  "smartArtAuthorPaint", "smartArtAuthorShape", "smartArtAuthorText", "smartArtAuthorTextSize",
  "smartArtAuthorTextFont", "smartArtAuthorTextAlign", "smartArtAuthorCapitals",
] as const;

/**
 * What a slide diagram's conversion read, kept on its first member. Its shapes state
 * the file's paint and format as values, so nothing on the page alone tells them from
 * what the person changes later; against this the save does, and writes only that
 * into the file's own data part (`smartart_data_merge.py`). Compared, never drawn.
 */
export function smartArtAsReadOf(members: readonly Record<string, unknown>[]): Record<string, unknown> | null {
  const first = members[0];
  if (!first) return null;
  const items: Record<string, Record<string, unknown>> = {};
  for (const member of members) {
    const item = member.smartArtItemId;
    const role = member.smartArtRole;
    if (typeof item !== "string" || (role !== "node" && role !== "label")) continue;
    const stated: Record<string, unknown> = {};
    for (const name of SMART_ART_AS_READ_FIELDS) if (member[name] !== undefined) stated[name] = member[name];
    if (Array.isArray(member.runs)) {
      stated.runStyles = member.runs.map((run: { style?: unknown }) => run?.style ?? null);
    }
    (items[item] ??= {})[role] = stated;
  }
  return {
    kind: first.smartArtKind ?? null,
    style: first.smartArtStyle ?? null,
    colors: first.smartArtColors ?? null,
    direction: first.smartArtDirection ?? null,
    background: first.smartArtBackground ?? null,
    items,
  };
}
