/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A SmartArt shape's TEXT BOX, apart from the shape's own box.
 *
 * Office lays a SmartArt shape's words out in a rectangle of the diagram's own
 * (the cached drawing's `dsp:txXfrm`), not in the shape's box: a Segmented
 * Cycle wedge's words sit inside the wedge, a Cycle Matrix quadrant's toward
 * its outer corner, a Grouped List item's in its group. A host that draws a
 * label as the text OF its shape (a Docs group member, a worksheet shape)
 * keeps that rectangle here; without it the words were centred in the shape's
 * box -- the whole disc for every wedge -- and five labels drew on top of each
 * other in the middle of the diagram.
 *
 * The rectangle is kept relative to the shape's box, in Office's percentage
 * unit (`100000` = 100 %, the unit `smartArtCustom` uses), so a shape the
 * author moves or resizes takes its words with it. Absent: the words fill the
 * shape's box, as before. Server mirror: `office_smartart_text_box.py`.
 */

import type { OfficeSmartArtNode } from "./model";

/** Office's percentage unit: `100000` = 100 % of the shape's box. */
export const SMARTART_TEXT_BOX_UNIT = 100_000;
/** How far outside its shape a text box may start (the engine's 2 % slack, and room). */
const SMARTART_TEXT_BOX_LOW = -20_000;
/** …and end. */
const SMARTART_TEXT_BOX_HIGH = 120_000;
/** The smallest edge a text box may have (0.5 % of the shape). */
const SMARTART_TEXT_BOX_MIN_EDGE = 500;
/**
 * A label whose centre is within this much of its shape's centre (on both
 * axes) is the shape's own centred text: the shape's box stays its text box.
 * The layouts that place words elsewhere (Segmented Cycle, Cycle Matrix,
 * Grouped List, Segmented Pyramid, Increasing Arrows) put them 11 % to 35 %
 * off; a centred label smaller than its shape (a list block's band) is not
 * recorded, so its words keep the whole shape to wrap in.
 */
const SMARTART_TEXT_BOX_CENTRED = 5_000;

export interface SmartArtTextBox {
  /** Left edge, from the shape's left, in {@link SMARTART_TEXT_BOX_UNIT}. */
  readonly left: number;
  /** Top edge, from the shape's top. */
  readonly top: number;
  readonly width: number;
  readonly height: number;
  /**
   * Drawn only, never stored: the words are set at a size the author fixed
   * (Office's `custT`), so Office breaks them by its own measures exactly,
   * with no room past the box (`SMARTART_LABEL_WRAP_SLACK` is for words the
   * engine fitted).
   */
  readonly exact?: true;
}

const FIELDS = ["left", "top", "width", "height"] as const;

/**
 * The persisted text box, or `null` for none / a record that is not one.
 * Same rule as the server's `text_box_error`.
 */
export function readSmartArtTextBox(value: unknown): SmartArtTextBox | null {
  if (!value || typeof value !== "object" || Array.isArray(value)) return null;
  const record = value as Record<string, unknown>;
  const keys = Object.keys(record);
  if (keys.length !== FIELDS.length || keys.some((key) => !(FIELDS as readonly string[]).includes(key))) {
    return null;
  }
  for (const field of FIELDS) {
    const number = record[field];
    if (typeof number !== "number" || !Number.isInteger(number)) return null;
  }
  const { left, top, width, height } = record as unknown as SmartArtTextBox;
  if (width < SMARTART_TEXT_BOX_MIN_EDGE || height < SMARTART_TEXT_BOX_MIN_EDGE) return null;
  if (left < SMARTART_TEXT_BOX_LOW || top < SMARTART_TEXT_BOX_LOW) return null;
  if (left + width > SMARTART_TEXT_BOX_HIGH || top + height > SMARTART_TEXT_BOX_HIGH) return null;
  return { left, top, width, height };
}

/**
 * The text box a label gives the shape it is merged into, or `null` when the
 * label is centred on the shape (within 5 % on both axes): the shape's own box
 * is then its text box and nothing is recorded.
 */
/** Shapes whose ends are no place for words (a chevron's notch and point). */
const POINTED_LABEL_HOSTS: ReadonlySet<string> = new Set(["chevron", "home-plate", "homePlate"]);

export function smartArtLabelTextBox(
  label: Pick<OfficeSmartArtNode, "left" | "top" | "width" | "height">,
  host: Pick<OfficeSmartArtNode, "left" | "top" | "width" | "height"> & { readonly shape?: string },
  /** The words are not centred (a bullet row flush left after its dot): a
   * box only nearly the host's is still theirs, or they start under the dot. */
  options: {
    readonly aligned?: boolean;
    /** The label is the shape's text area with its insets already off (an
     * Office scene's `dsp:txXfrm` less `a:bodyPr`'s insets): it is kept even
     * centred, or the host drew the words over its whole box inside insets of
     * its own a second time. */
    readonly exact?: boolean;
  } = {},
): SmartArtTextBox | null {
  const values = [label.left, label.top, label.width, label.height, host.left, host.top, host.width, host.height];
  if (values.some((value) => !Number.isFinite(value))) return null;
  if (host.width <= 0 || host.height <= 0) return null;
  const unit = (part: number, whole: number) => Math.round((part / whole) * SMARTART_TEXT_BOX_UNIT);
  const box = {
    left: unit(label.left - host.left, host.width),
    top: unit(label.top - host.top, host.height),
    width: unit(label.width, host.width),
    height: unit(label.height, host.height),
  };
  if (options.exact) return readSmartArtTextBox(box);
  const half = SMARTART_TEXT_BOX_UNIT / 2;
  const centred =
    Math.abs(box.left + box.width / 2 - half) <= SMARTART_TEXT_BOX_CENTRED &&
    Math.abs(box.top + box.height / 2 - half) <= SMARTART_TEXT_BOX_CENTRED;
  // A pointed shape's centred label keeps its width: a chevron's words sit
  // on its body, between the notch and the point. As the shape's own box
  // they ran into the notch (live audit S-170, "结果789" cut by a Right to
  // Left V-shaped process). The full height stays the shape's, so the words
  // still wrap and centre on it.
  if (centred && host.shape !== undefined && POINTED_LABEL_HOSTS.has(host.shape) &&
      box.width < SMARTART_TEXT_BOX_UNIT - 2 * SMARTART_TEXT_BOX_CENTRED) {
    return readSmartArtTextBox({ ...box, top: 0, height: SMARTART_TEXT_BOX_UNIT });
  }
  if (centred && options.aligned) {
    const whole = box.left === 0 && box.top === 0 && box.width === SMARTART_TEXT_BOX_UNIT && box.height === SMARTART_TEXT_BOX_UNIT;
    return whole ? null : readSmartArtTextBox(box);
  }
  return centred ? null : readSmartArtTextBox(box);
}

/**
 * The text box as CSS insets of the shape's box, in percent (`inset: t r b l`),
 * for a host that draws the words in an element laid over the shape.
 */
export function smartArtTextBoxInsetPercents(box: SmartArtTextBox): {
  top: number;
  right: number;
  bottom: number;
  left: number;
} {
  const percent = (value: number) => (value / SMARTART_TEXT_BOX_UNIT) * 100;
  return {
    top: percent(box.top),
    right: percent(SMARTART_TEXT_BOX_UNIT - box.left - box.width),
    bottom: percent(SMARTART_TEXT_BOX_UNIT - box.top - box.height),
    left: percent(box.left),
  };
}

/** The text box in the shape's own space: `{x, y, width, height}` of a shape at `shape`. */
export function smartArtTextBoxRect(
  box: SmartArtTextBox,
  shape: { x: number; y: number; width: number; height: number },
): { x: number; y: number; width: number; height: number } {
  const along = (value: number, whole: number) => (value / SMARTART_TEXT_BOX_UNIT) * whole;
  return {
    x: shape.x + along(box.left, shape.width),
    y: shape.y + along(box.top, shape.height),
    width: along(box.width, shape.width),
    height: along(box.height, shape.height),
  };
}
