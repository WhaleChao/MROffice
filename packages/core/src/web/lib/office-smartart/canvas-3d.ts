/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import cameraTable from "../office-camera-rotations.shared.json";
import { OFFICE_CAMERA_DISTANCE_EMU_AT_45 } from "../office-camera-distance";
import {
  type SmartArtLightRigType,
  type SmartArtMaterial,
  type SmartArtScene3d,
  type SmartArtShape3d,
} from "./styles-model";

const EMU_PER_PX = 9525;
/** DrawingML angles are 60000ths of a degree. */
const ANGLE_UNITS_PER_DEGREE = 60000;

const emuToPx = (emu: number): number => emu / EMU_PER_PX;

/** `#rrggbb` with each DRAWN (encoded) channel scaled by a measured ratio. */
function scaleDrawnColor(hex: string, ratio: number): string {
  const digits = /^#?([0-9a-fA-F]{6})$/.exec(hex.trim())?.[1] ?? "000000";
  return (
    "#" +
    [0, 2, 4]
      .map((at) =>
        Math.max(
          0,
          Math.min(255, Math.round(parseInt(digits.slice(at, at + 2), 16) * ratio)),
        )
          .toString(16)
          .padStart(2, "0"),
      )
      .join("")
  );
}

/** Map a DrawingML angle onto (-180, 180], so 355° reads as the -5° it looks like. */
export function smartArtAngleDegrees(units: number | undefined): number {
  if (typeof units !== "number" || !Number.isFinite(units)) return 0;
  const degrees = ((units / ANGLE_UNITS_PER_DEGREE) % 360 + 360) % 360;
  return degrees > 180 ? degrees - 360 : degrees;
}


































export interface SmartArtLightRigPaint {
  /** Per-channel multiplier on the node's own fill. */
  readonly face: readonly [number, number, number];
  /** Per-channel offset added after it, 0-255 (`chilly`'s light grey). */
  readonly offset?: readonly [number, number, number];
  /** Per-channel multiplier on the extruded faces, relative to `face`. */
  readonly body?: readonly [number, number, number];
}

const SMARTART_LIGHT_RIG_PAINT: Readonly<
  Partial<Record<SmartArtLightRigType, SmartArtLightRigPaint>>
> = {
  // `flat` fits at 0.998 / 0.996 / 0.998 — the identity, to within what the
  // PDF rasteriser itself moves a colour — so it is recorded as the identity.
  flat: { face: [1, 1, 1], body: [1, 1, 1] },
  morning: { face: [0.881, 0.848, 0.765] },
  soft: { face: [0.966, 0.969, 0.97], body: [0.434, 0.434, 0.434] },
  
  
  // #F2F2F2, #1F1F1F) as `out = a·in + b`, max error 1.2/255. Not a plain
  // multiplier: a cool ambient lifts every face (Powder's pale blue).
  chilly: { face: [0.691, 0.691, 0.731], offset: [78.4, 78.4, 78.2] },
  balanced: { face: [0.953, 0.953, 0.953], offset: [-1.5, -1.5, -1.5] },
};









const CHILLY_WITHOUT_POWDER: SmartArtLightRigPaint = { face: [1, 1, 1.046] };







const CHILLY_UNDER_POWDER: SmartArtLightRigPaint = { face: [1, 1, 1.058] };








const METAL_THREE_POINT_FACE = { face: [0.98, 0.975, 0.976] as const, top: 108, bottom: 42 };










function metalFaceLight(scene3d: SmartArtScene3d | undefined, shape3d?: SmartArtShape3d): boolean {
  if (shape3d?.material !== "metal" || lightRigOf(scene3d, shape3d) !== "threePt") return false;
  const camera = scene3d?.camera;
  return (
    !!camera &&
    (camera.preset !== "orthographicFront" || !!camera.latitude || !!camera.longitude)
  );
}

 
function lightRigOf(scene3d: SmartArtScene3d | undefined, shape3d?: SmartArtShape3d): SmartArtLightRigType | undefined {
  return scene3d?.lightRig.rig ?? shape3d?.lightRig?.rig;
}

 
export function smartArtLightRigPaint(
  scene3d: SmartArtScene3d | undefined,
  shape3d?: SmartArtShape3d,
): SmartArtLightRigPaint | null {
  const rig = lightRigOf(scene3d, shape3d);
  if (rig === "chilly" && shape3d?.material !== "translucentPowder") return CHILLY_WITHOUT_POWDER;
  // Powder's "lift" was the ground showing through: the face keeps its
  
  // ({@link smartArtMaterialOpacityPrimitives}).
  if (rig === "chilly" && smartArtMaterialOpacity(shape3d) !== null) return CHILLY_UNDER_POWDER;
  return rig ? SMARTART_LIGHT_RIG_PAINT[rig] ?? null : null;
}












export function smartArtMaterialOpacity(shape3d: SmartArtShape3d | undefined): number | null {
  return shape3d?.material === "translucentPowder" ? 0.7 : null;
}

/**
 * The closing steps that draw `input` (the member's lit face) at its
 * material's opacity ({@link smartArtMaterialOpacity}): its colour where the
 * member is, and an alpha of the material's opacity over the member's
 * coverage -- the fill's own transparency dropped. None for another material.
 */
/**
 * Whether a member's material opacity is drawn AFTER its bevel (the closing
 * {@link smartArtMaterialOpacityPrimitives}) rather than under it: a bevel read
 * side by side is laid on the composited face ({@link smartArtBevelLightPrimitives}'
 * `faceOpacity`), any other on the opaque face it lights.
 */
export function smartArtMaterialOpacityAfterBevel(light: SmartArtBevelLight | null | undefined): boolean {
  return !light?.edges;
}

export function smartArtMaterialOpacityPrimitives(
  shape3d: SmartArtShape3d | undefined,
  input: string,
): SmartArtFilterPrimitive[] {
  const opacity = smartArtMaterialOpacity(shape3d);
  if (opacity === null) return [];
  return [
    {
      tag: "feComponentTransfer",
      attrs: { in: "SourceAlpha", result: "materialCoverage" },
      children: [{ tag: "feFuncA", attrs: { type: "linear", slope: 50, intercept: 0 } }],
    },
    {
      tag: "feColorMatrix",
      attrs: { in: input, type: "matrix", values: "1 0 0 0 0  0 1 0 0 0  0 0 1 0 0  0 0 0 0 1", result: "materialFace" },
    },
    { tag: "feComposite", attrs: { in: "materialFace", in2: "materialCoverage", operator: "in", result: "materialCovered" } },
    {
      tag: "feComponentTransfer",
      attrs: { in: "materialCovered", result: "materialResult" },
      children: [{ tag: "feFuncA", attrs: { type: "linear", slope: opacity, intercept: 0 } }],
    },
  ];
}










export function smartArtLightRigMatrix(
  scene3d: SmartArtScene3d | undefined,
  shape3d?: SmartArtShape3d,
): string | null {
  const metal = metalFaceLight(scene3d, shape3d);
  const paint = metal ? { face: METAL_THREE_POINT_FACE.face } : smartArtLightRigPaint(scene3d, shape3d);
  if (!paint) return null;
  const [r, g, b] = paint.face;
  const [or, og, ob] = (paint as SmartArtLightRigPaint).offset ?? [0, 0, 0];
  if (r === 1 && g === 1 && b === 1 && !or && !og && !ob) return null;
  const off = (value: number) => fixed(value / 255);
  return `${r} 0 0 0 ${off(or)}  0 ${g} 0 0 ${off(og)}  0 0 ${b} 0 ${off(ob)}  0 0 0 1 0`;
}























function metalLightAt(u: number, v: number): number {
  const down = 1.2 - Math.max(0, Math.min(1, v));
  const across = 0.2 + 1.08 / (1 + Math.exp((Math.max(0, Math.min(1, u)) - 0.63) / 0.12));
  return (4 + 63 * down * down) * across;
}


















function balancedLightAt(u: number, v: number): number {
  return (5.4654 + 14.7079 * Math.max(0, Math.min(1, v))) * (Math.max(0, Math.min(1, u)) + 0.3099) ** 4.1151;
}

/** Where a member stands in its diagram, as fractions of the diagram's frame. */
export type SmartArtFaceSpan = {
  readonly top: number;
  readonly bottom: number;
  readonly left?: number;
  readonly right?: number;
  /**
   * The member's own turn (degrees, clockwise) and mirror, which the face
   * filter is drawn under: the light is the diagram's, so it is read back
   * through them. Picture Strips' strips are `round2SameRect` turned 180°:
   * read as if upright, each strip was lit from its left where Windows lights
   * it from its right (truth office-truth/probes/sa36-3dcam pList2-3d6).
   */
  readonly rotation?: number;
  readonly flipH?: boolean;
  readonly flipV?: boolean;
  /** The diagram frame's width over its height, for a turn that is not a half turn. */
  readonly aspect?: number;
};








function lightImage(span: SmartArtFaceSpan, lightAt: (u: number, v: number) => number): string | null {
  const { left, right, top, bottom } = span;
  if (left === undefined || right === undefined || !(right > left) || !(bottom > top)) return null;
  const width = right - left;
  const height = bottom - top;
  // A point of the member's own (upright, unmirrored) box, x and y from 0 to
  // 1, where it stands in the diagram once the member is mirrored and turned
  // about its centre -- the place the diagram's light is read at.
  const turn = (((span.rotation ?? 0) % 360) + 360) % 360;
  const aspect = span.aspect && span.aspect > 0 ? span.aspect : 1;
  const radians = (turn * Math.PI) / 180;
  const cos = Math.cos(radians);
  const sin = Math.sin(radians);
  const centreU = (left + right) / 2;
  const centreV = (top + bottom) / 2;
  const at = (x: number, y: number): [number, number] => {
    const dx = (x - 0.5) * (span.flipH ? -1 : 1);
    const dy = (y - 0.5) * (span.flipV ? -1 : 1);
    if (!turn) return [centreU + dx * width, centreV + dy * height];
    return [
      centreU + dx * width * cos - (dy * height * sin) / aspect,
      centreV + dx * width * sin * aspect + dy * height * cos,
    ];
  };
  const strips = 32;
  let lit = false;
  const rects: string[] = [];
  for (let index = 0; index < strips; index += 1) {
    const x = -0.1 + (1.2 * (index + 0.5)) / strips;
    const a = Math.round(Math.max(0, Math.min(255, lightAt(...at(x, -0.1)))));
    const b = Math.round(Math.max(0, Math.min(255, lightAt(...at(x, 1.1)))));
    if (a || b) lit = true;
    rects.push(
      `<linearGradient id="g${index}" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="rgb(${a},${a},${a})"/><stop offset="1" stop-color="rgb(${b},${b},${b})"/></linearGradient>` +
        `<rect x="${fixed(index / strips)}" y="0" width="${fixed(1 / strips + 0.001)}" height="1" fill="url(#g${index})"/>`,
    );
  }
  if (!lit) return null;
  return (
    "data:image/svg+xml," +
    encodeURIComponent(
      `<svg xmlns="http://www.w3.org/2000/svg" width="256" height="256" viewBox="0 0 1 1" preserveAspectRatio="none">${rects.join("")}</svg>`,
    )
  );
}

export function smartArtFaceTintPrimitives(
  scene3d: SmartArtScene3d | undefined,
  shape3d?: SmartArtShape3d,
  /** Where the member stands in the diagram's frame (fractions; left / right for 3d6's light). */
  span?: SmartArtFaceSpan | null,
): SmartArtFilterPrimitive[] {
  const matrix = smartArtLightRigMatrix(scene3d, shape3d);
  if (!matrix) return [];
  const metal = metalFaceLight(scene3d, shape3d);
  const balanced = !metal && span ? lightRigOf(scene3d, shape3d) === "balanced" && lightImage(span, balancedLightAt) : null;
  // Metal's light, across the member as well as down it, where the host says
  // where the member stands in its diagram.
  const metalLit = metal && span ? lightImage(span, metalLightAt) : null;
  // The light is added to the face's own colour, and the member's alpha put
  // back once: the face is taken OPAQUE for the sum (its alpha row x255), or a
  
  // closing `in SourceAlpha` with its colour multiplied by its alpha twice,
  // dark where Windows draws it light. An opaque face is as it was.
  const opaque = matrix.replace(/0 0 0 1 0$/, "0 0 0 255 0");
  if (balanced || metalLit) {
    return [
      { tag: "feColorMatrix", attrs: { in: "SourceGraphic", type: "matrix", values: opaque, result: "faceColour" } },
      { tag: "feImage", attrs: { href: (balanced || metalLit) as string, preserveAspectRatio: "none", result: "faceLight" } },
      { tag: "feComposite", attrs: { in: "faceColour", in2: "faceLight", operator: "arithmetic", k1: 0, k2: 1, k3: 1, k4: 0, result: "faceLit" } },
      { tag: "feComposite", attrs: { in: "faceLit", in2: "SourceAlpha", operator: "in", result: "tinted" } },
    ];
  }
  if (!metal) {
    return [{ tag: "feColorMatrix", attrs: { in: "SourceGraphic", type: "matrix", values: matrix, result: "tinted" } }];
  }
  // The light falls across the filter region, which reaches 10% past the
  // shape on each side: the ramp's ends are extended to match.
  // Without the diagram's frame (a gallery thumbnail, a host that does not
  // say), the ramp measured down one tall shape in the upper half.
  const extent = span && span.bottom > span.top ? span.bottom - span.top : null;
  const at = (t: number) =>
    Math.round(
      Math.max(
        0,
        Math.min(
          255,
          extent !== null && span
            ? metalLightAt(0.5, span.top + extent * t)
            : METAL_THREE_POINT_FACE.top + (METAL_THREE_POINT_FACE.bottom - METAL_THREE_POINT_FACE.top) * t,
        ),
      ),
    );
  const grey = (v: number) => `rgb(${v},${v},${v})`;
  const ramp =
    "data:image/svg+xml," +
    encodeURIComponent(
      `<svg xmlns="http://www.w3.org/2000/svg" width="64" height="256" viewBox="0 0 1 1" preserveAspectRatio="none">` +
        `<linearGradient id="r" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="${grey(at(-0.1))}"/><stop offset="1" stop-color="${grey(at(1.1))}"/></linearGradient>` +
        `<rect width="1" height="1" fill="url(#r)"/></svg>`,
    );
  return [
    { tag: "feColorMatrix", attrs: { in: "SourceGraphic", type: "matrix", values: opaque, result: "faceColour" } },
    { tag: "feImage", attrs: { href: ramp, preserveAspectRatio: "none", result: "faceLight" } },
    { tag: "feComposite", attrs: { in: "faceColour", in2: "faceLight", operator: "arithmetic", k1: 0, k2: 1, k3: 1, k4: 0, result: "faceLit" } },
    { tag: "feComposite", attrs: { in: "faceLit", in2: "SourceAlpha", operator: "in", result: "tinted" } },
  ];
}

/**
 * How hard a material reads. PowerPoint varies specular response; we only vary
 * the contrast of the bevel rim, which is the part a flat renderer can show.
 */
const MATERIAL_CONTRAST: Readonly<Record<SmartArtMaterial, number>> = {
  matte: 0.3,
  translucentPowder: 0.26,
  plastic: 0.48,
  dkEdge: 0.62,
  metal: 0.72,
};
















































interface CameraPreset {
  /** Degrees about X (tilt, positive leans the top away). */
  latitude: number;
  /** Degrees about Y (yaw, positive turns the left edge away). */
  longitude: number;
  /** Degrees about Z (roll). */
  revolution: number;
  perspective: boolean;
  /**
   * CSS `perspective()` distance in EMU — an absolute length, measured; see the
   * note above. Absent falls back to {@link DEFAULT_PERSPECTIVE_DISTANCE_EMU}.
   */
  perspectiveDistanceEmu?: number;
}

/** What a perspective preset with no fitted distance of its own gets. */
const DEFAULT_PERSPECTIVE_DISTANCE_EMU = 13_840_000;

export const SMARTART_CAMERA_PRESETS: Readonly<Record<string, CameraPreset>> = {
  orthographicFront: { latitude: 0, longitude: 0, revolution: 0, perspective: false },
  // Both angles NEGATIVE: the pair's overall sign leaves an orthographic
  // projection untouched (`cos` is even and `sin lat · sin lon` keeps its
  // sign), so only the extruded body can read it, and the body says this one.
  isometricOffAxis2Left: {
    latitude: -17.92, longitude: -26.12, revolution: 0, perspective: false,
  },
  // 3d9 is the only style on this preset and it overrides the angles; the
  // distance is measured and is what 3d9 actually uses.
  perspectiveRelaxed: {
    latitude: 12, longitude: 0, revolution: 0, perspective: true,
    perspectiveDistanceEmu: 13_861_743,
  },
  perspectiveRelaxedModerately: {
    latitude: 34.64, longitude: 0, revolution: 0, perspective: true,
    perspectiveDistanceEmu: 13_785_491,
  },
  // The yaw is NEGATIVE, which only the perspective fit could tell: with a zero
  // latitude an affine sees `cos(lon)` alone and cannot read the sign at all.
  perspectiveLeft: {
    latitude: 0, longitude: -19.7, revolution: 0, perspective: true,
    perspectiveDistanceEmu: 13_878_855,
  },
  // Estimate: 3d8 is the only style on this preset and it overrides the angles.
  // Its distance is measured, and is half of what the other three cameras use.
  perspectiveHeroicExtremeRightFacing: {
    latitude: 6, longitude: 24, revolution: 0, perspective: true,
    perspectiveDistanceEmu: 6_935_639,
  },
};



















function sceneCameraPreset(name: string): CameraPreset {
  const measured = SMARTART_CAMERA_PRESETS[name];
  if (measured) return measured;
  const row = (cameraTable.cameras as unknown as Record<string, readonly number[] | undefined>)[name];
  if (!row) return SMARTART_CAMERA_PRESETS.orthographicFront!;
  const [x = 0, y = 0, z = 0, fov = 45] = row;
  const perspective = name.startsWith("perspective");
  const half = (Math.min(179, Math.max(1, fov)) * Math.PI) / 360;
  return {
    latitude: y ? -y : 0,
    longitude: x,
    revolution: z,
    perspective,
    ...(perspective
      ? { perspectiveDistanceEmu: Math.round((OFFICE_CAMERA_DISTANCE_EMU_AT_45 * Math.tan(Math.PI / 8)) / Math.tan(half)) }
      : {}),
  };
}

export interface SmartArtCameraAngles {
  latitude: number;
  longitude: number;
  revolution: number;
  /** 1 = unzoomed. */
  scale: number;
  perspective: boolean;
}

/**
 * Resolve a diagram camera to concrete degrees: the preset's orientation, then
 * the style's explicit `a:rot` override, then `zoom`.
 *
 * The three angles come back NEGATED from the file's `a:rot/@lat/@lon/@rev`.
 * DrawingML says only that they are "a rotation in 3-D space ... a latitude
 * coordinate, a longitude coordinate, and a revolution about the axis as the
 * latitude and longitude coordinates" (ECMA-376 §20.1.5, and [MS-DOCX]
 * CT_SphereCoords repeats it) — neither the axis order nor the sense is
 * written down anywhere, so both were measured.
 *
 * PowerPoint's own render of a group carrying 3d9's `a:scene3d` verbatim fits
 *
 *     Rz(-rev) · Rx(-lat) · Ry(-lon),  CSS perspective 13.86M EMU
 *
 * to within 1.0px on a 3840px page, and the rotation it recovers blind — with
 * no angle held fixed — is 0.057° from that product of the file's own numbers.
 * Nothing else is close: the same angles taken at face value (`Rz(+rev)·
 * Rx(+lat)·Ry(+lon)`, what this module used to do) are 817px out, and every
 * other axis order is 170px or worse. 3d8, the only other style carrying an
 * `a:rot`, agrees at 0.69°.
 *
 * That sign is the whole of the 3d9 defect: its roll is ~26°, and rolling the
 * diagram the wrong way about its centre threw the members furthest from the
 * centre out of the frame.
 */
export function smartArtCameraAngles(scene3d: SmartArtScene3d): SmartArtCameraAngles {
  const preset = sceneCameraPreset(scene3d.camera.preset);
  const hasExplicitRotation =
    scene3d.camera.latitude !== undefined ||
    scene3d.camera.longitude !== undefined ||
    scene3d.camera.revolution !== undefined;
  const zoom = scene3d.camera.zoom;
  const fromFile = (units: number | undefined) => -smartArtAngleDegrees(units);
  return {
    latitude: hasExplicitRotation
      ? fromFile(scene3d.camera.latitude)
      : preset.latitude,
    longitude: hasExplicitRotation
      ? fromFile(scene3d.camera.longitude)
      : preset.longitude,
    revolution: hasExplicitRotation
      ? fromFile(scene3d.camera.revolution)
      : preset.revolution,
    scale: typeof zoom === "number" && zoom > 0 ? zoom / 100000 : 1,
    perspective: preset.perspective,
  };
}

export interface SmartArtBox {
  left: number;
  top: number;
  width: number;
  height: number;
}

/**
 * The CSS a member needs so the whole diagram tilts as ONE object: the same
 * rotation on every member, with each member's transform origin moved onto the
 * diagram's shared centre.
 *
 * Both boxes must be in the SAME units (the block model's EMU); the origin is
 * returned as a PERCENTAGE of the member's own box, which makes it independent
 * of whatever units the host positions the member in and of the canvas zoom.
 * Only the perspective distance needs a real length, so it is converted from
 * EMU to px — i.e. it is calibrated for 100% zoom and the foreshortening is
 * slightly off at other zooms. That is inside the stated fidelity of §5.
 */
export function smartArtCameraCss(
  scene3d: SmartArtScene3d,
  group: SmartArtBox,
  member: SmartArtBox,
): { transform: string; transformOrigin: string; transformStyle: "preserve-3d" } | null {
  const angles = smartArtCameraAngles(scene3d);
  if (
    angles.latitude === 0 &&
    angles.longitude === 0 &&
    angles.revolution === 0 &&
    angles.scale === 1
  ) {
    return null;
  }
  if (!(member.width > 0) || !(member.height > 0)) return null;
  const originX = ((group.left + group.width / 2 - member.left) / member.width) * 100;
  const originY = ((group.top + group.height / 2 - member.top) / member.height) * 100;
  const preset = sceneCameraPreset(scene3d.camera.preset);
  // A field of view the camera states itself (Format Shape > 3-D Rotation >
  // Perspective, `a:camera@fov`) stands the eye at Office's absolute distance
  // for that view instead of the preset's.
  const ownFov = typeof scene3d.camera.fov === "number" ? scene3d.camera.fov / 60_000 : undefined;
  const ownDistanceEmu =
    preset.perspective && ownFov !== undefined
      ? Math.round(
          (OFFICE_CAMERA_DISTANCE_EMU_AT_45 * Math.tan(Math.PI / 8)) /
            Math.tan((Math.min(179, Math.max(1, ownFov)) * Math.PI) / 360),
        )
      : undefined;
  // An ABSOLUTE distance, not a multiple of the diagram: measured across a 4x
  // range of diagram sizes and two slide shapes (see SMARTART_CAMERA_PRESETS).
  // The diagram's own size still decides how much of that distance it spans,
  // which is exactly how PowerPoint gives a bigger diagram more perspective.
  const distance = emuToPx(
    ownDistanceEmu ?? preset.perspectiveDistanceEmu ?? DEFAULT_PERSPECTIVE_DISTANCE_EMU,
  );
  // ORDER MATTERS, and it is Office's: `rotateX(lat) rotateY(lon)` composes as
  // Rx·Ry, which shears the diagram VERTICALLY (a tilted diagram's right-hand
  // side sinks) — measured off PowerPoint's own render, and the same matrix
  // {@link smartArtCameraSvgTransform} builds for thumbnails. Written the other
  // way round (`rotateY rotateX`, which is Ry·Rx) the same angles shear
  // HORIZONTALLY instead, so the canvas leaned one way and every thumbnail of
  // it leaned the other.
  //
  // `scale()` is OUTSIDE the `perspective()`, because `zoom` magnifies the
  // PROJECTED image rather than the object: inside, CSS would divide by the
  // zoomed depth and the same distance would foreshorten by 1/zoom too much
  // (8-18% for the three zoomed styles).
  const parts = [
    ...(angles.scale === 1 ? [] : [`scale(${angles.scale.toFixed(4)})`]),
    ...(angles.perspective ? [`perspective(${distance.toFixed(2)}px)`] : []),
    ...(angles.revolution ? [`rotateZ(${angles.revolution.toFixed(2)}deg)`] : []),
    ...(angles.latitude ? [`rotateX(${angles.latitude.toFixed(2)}deg)`] : []),
    ...(angles.longitude ? [`rotateY(${angles.longitude.toFixed(2)}deg)`] : []),
  ];
  return {
    transform: parts.join(" "),
    transformOrigin: `${originX.toFixed(3)}% ${originY.toFixed(3)}%`,
    transformStyle: "preserve-3d",
  };
}

/**
 * A diagram member's transform under the diagram camera: the camera about
 * the diagram's centre (the `transformOrigin` {@link smartArtCameraCss}
 * gives), then the member's own rotate / flip about the member's OWN centre.
 * CSS has one origin per element, so the own part is moved to the member's
 * centre and back: written bare after the camera it pivoted about the
 * diagram's centre, and Basic Cycle's turned arrows left the ring under every
 * camera style (3d5-3d9), while PowerPoint keeps them between the circles.
 */
export function smartArtCameraWithOwnTransform(
  camera: { transform: string; transformOrigin: string },
  own: string | null | undefined,
): string {
  const self = typeof own === "string" && own !== "none" ? own.trim() : "";
  if (!self) return camera.transform;
  const origin = /^(-?[\d.]+)% (-?[\d.]+)%$/.exec(camera.transformOrigin.trim());
  if (!origin) return `${camera.transform} ${self}`.trim();
  const dx = 50 - Number(origin[1]);
  const dy = 50 - Number(origin[2]);
  return `${camera.transform} translate(${dx.toFixed(3)}%, ${dy.toFixed(3)}%) ${self} translate(${(-dx).toFixed(3)}%, ${(-dy).toFixed(3)}%)`.trim();
}

/**
 * The static-SVG equivalent: a true axonometric projection of the same three
 * rotations, as an affine matrix about the diagram's centre.
 *
 * Rotating by `lon` about Y then `lat` about X and dropping Z gives
 *   x' = cos(lon)·x
 *   y' = sin(lon)·sin(lat)·x + cos(lat)·y
 * which is exactly `matrix(cos lon, sin lon · sin lat, 0, cos lat, 0, 0)`; the
 * roll is the 2-D `rotate()` in front of it, because `Rz(rev)` multiplies the
 * whole projected plane. SVG applies the list left to right, so `rotate` before
 * `matrix` is `Rz · Rx · Ry` — the order the canvas builds and the order
 * PowerPoint's render decomposes into.
 *
 * The ROLL is what a tile of 3d9 gets wrong when the sign is wrong: `lat` and
 * `lon` reach this matrix only through `cos` and `sin·sin`, both even in the
 * pair's sign, so a wrong sign on them is invisible here — but `rev` is a plain
 * rotation and a 26° roll the wrong way swings the members furthest from the
 * centre out of the tile.
 *
 * Perspective is not expressible in an affine transform and is dropped — the
 * reason thumbnails read as flatter than the canvas.
 */
export function smartArtCameraSvgTransform(
  scene3d: SmartArtScene3d,
  centerX: number,
  centerY: number,
): string | null {
  const angles = smartArtCameraAngles(scene3d);
  if (
    angles.latitude === 0 &&
    angles.longitude === 0 &&
    angles.revolution === 0 &&
    angles.scale === 1
  ) {
    return null;
  }
  const rad = (degrees: number) => (degrees * Math.PI) / 180;
  const a = Math.cos(rad(angles.longitude)) * angles.scale;
  const b = Math.sin(rad(angles.longitude)) * Math.sin(rad(angles.latitude)) * angles.scale;
  const d = Math.cos(rad(angles.latitude)) * angles.scale;
  return [
    `translate(${centerX.toFixed(3)} ${centerY.toFixed(3)})`,
    ...(angles.revolution ? [`rotate(${angles.revolution.toFixed(3)})`] : []),
    `matrix(${a.toFixed(5)} ${b.toFixed(5)} 0 ${d.toFixed(5)} 0 0)`,
    `translate(${(-centerX).toFixed(3)} ${(-centerY).toFixed(3)})`,
  ].join(" ");
}

/**
 * Where a point lands under {@link smartArtCameraSvgTransform} — the same
 * affine, as numbers, so a static surface can measure what the camera does to
 * its own drawing (a style tile has to keep a tilted diagram inside its frame).
 * Identity when the scene has no camera.
 */
export function smartArtCameraSvgPoint(
  scene3d: SmartArtScene3d | undefined,
  centerX: number,
  centerY: number,
): (x: number, y: number) => { x: number; y: number } {
  if (!scene3d) return (x, y) => ({ x, y });
  const angles = smartArtCameraAngles(scene3d);
  const rad = (degrees: number) => (degrees * Math.PI) / 180;
  const a = Math.cos(rad(angles.longitude)) * angles.scale;
  const b = Math.sin(rad(angles.longitude)) * Math.sin(rad(angles.latitude)) * angles.scale;
  const d = Math.cos(rad(angles.latitude)) * angles.scale;
  const cosR = Math.cos(rad(angles.revolution));
  const sinR = Math.sin(rad(angles.revolution));
  return (x, y) => {
    // translate(-c) → matrix(a b 0 d) → rotate(rev) → translate(c), read right
    // to left exactly as the SVG transform list applies them.
    const px = x - centerX;
    const py = y - centerY;
    const mx = a * px;
    const my = b * px + d * py;
    return {
      x: cosR * mx - sinR * my + centerX,
      y: sinR * mx + cosR * my + centerY,
    };
  };
}

/**
 * Where a member's extruded body sits, as an offset in the member's OWN
 * (pre-camera) pixels — `null` when the scene shows no side of it.
 *
 * `a:sp3d/@extrusionH` extrudes AWAY from the viewer, along −Z: the front face
 * stays where the flat shape is and the solid runs back from it. The camera
 * sends −Z to the screen vector `−(sin lon, −sin lat · cos lon)`. Our filter
 * runs INSIDE the element the camera transform is on, so the offset it needs is
 * the one that lands there after the camera, i.e. `M⁻¹` of that:
 *
 *     depth · ( −tan lon , tan lat / cos lon ).
 *
 * The camera's roll drops out of this on its own — it multiplies the screen
 * vector and `M` by the same 2-D rotation, which `M⁻¹` then cancels — so the
 * body is correct under 3d8's and 3d9's rolls without carrying `rev` here.
 *
 * MEASURED on four styles, each against a slide carrying the SAME camera with
 * NO `a:sp3d`, so the body vector is the difference of two support functions
 * rather than a guess at where a face ends (px on a 3840px page, unit vectors):
 *
 *   3d5  body (+51.1, −32.2) | −Z predicts (+0.845, −0.535), measured (+0.846, −0.533)
 *   3d7  body (+13.5,  −0.2) | −Z predicts (+1.000,  0.000), measured (+1.000, −0.016)
 *   3d8  body (−20.3,  +0.3) | −Z predicts (−0.960, +0.279), measured (−1.000, +0.013)
 *   3d9  body ( −9.2, +51.1) | −Z predicts (−0.158, +0.987), measured (−0.178, +0.984)
 *
 * and the length comes back as `extrusionH` PLUS the top bevel's height —
 * 3d9's 152250 + 76200 EMU projects to 52.3px against 51.9px measured — since
 * the bevel raises the plateau above the base and the silhouette sweeps both.
 * That extra is not added here: it is a percent-level effect on the body's
 * length, and `a:bevelT/@h` is not something this renderer reads today.
 *
 * `null` when the diagram has no camera at all: PowerPoint draws an
 * `orthographicFront` extrusion as nothing — you are looking straight at the
 * front face — and a deck exported with `a:sp3d` but no `a:scene3d` comes back
 * from PowerPoint perfectly flat.
 */
export function smartArtExtrusionOffset(
  scene3d: SmartArtScene3d | undefined,
  depthPx: number,
): { dx: number; dy: number } | null {
  if (!scene3d || !(depthPx > 0)) return null;
  const angles = smartArtCameraAngles(scene3d);
  const rad = (degrees: number) => (degrees * Math.PI) / 180;
  const lat = rad(angles.latitude);
  const lon = rad(angles.longitude);
  const cosLat = Math.cos(lat);
  const cosLon = Math.cos(lon);
  // A camera looking along the plane makes `M` singular and the inverse
  // meaningless; no preset comes near it, so refuse rather than guess.
  if (Math.abs(cosLat) < 0.2 || Math.abs(cosLon) < 0.2) return null;
  const dx = -depthPx * Math.tan(lon);
  const dy = (depthPx * Math.tan(lat)) / cosLon;
  return dx === 0 && dy === 0 ? null : { dx, dy };
}

/**
 * The CSS `filter` that fakes a member's bevel, extrusion and nothing else
 * (contour is a stroke, see {@link smartArtContourStroke}). Returns `null` when
 * the style asks for neither.
 *
 * `fill` is the member's own colour, so the extruded body is a copy of what the
 * face actually is rather than a generic grey.
 */
/**
 * How far a member's solid runs back from its front face, in EMU.
 *
 * Measured on PowerPoint's own render of style 3d9: the body is `extrusionH`
 * PLUS the top bevel's height (152250 + 76200 EMU predicts 52.3px, the render
 * gives 51.9px), because the front face is the top of the bevel, which sits on
 * the extrusion. Reading `extrusionH` alone left every such body about 1.5%
 * short, and more for a taller bevel. A member with no extrusion keeps no body:
 * that is the case the measurement did not cover.
 */
export function smartArtBodyDepthEmu(shape3d: SmartArtShape3d): number {
  const extrusion = shape3d.extrusionHeightEmu ?? 0;
  if (!(extrusion > 0)) return 0;
  return extrusion + Math.max(0, shape3d.bevelTop?.heightEmu ?? 0);
}

export function smartArtShape3dFilter(
  shape3d: SmartArtShape3d,
  fill: string | undefined,
  scene3d?: SmartArtScene3d,
): string | null {
  const layers: string[] = [];
  const offset = smartArtExtrusionOffset(scene3d, emuToPx(smartArtBodyDepthEmu(shape3d)));
  if (offset) {
    // A SWEPT body, not one offset copy: an extruded solid's silhouette is the
    // shape smeared along the whole extrusion vector, and one copy at the far
    // end leaves the middle empty — which is why a 3-D node read as a flat card
    // with a shadow behind it. Chaining `drop-shadow`s doubles the set of
    // offsets each time (a step of v, then 2v, then 4v… reaches every multiple
    // 0…2ⁿ-1 of v), so n layers fill the sweep at a step of ~1.5px or finer.
    const length = Math.hypot(offset.dx, offset.dy);
    const layerCount = Math.max(1, Math.min(6, Math.ceil(Math.log2(length / 1.5 + 1))));
    const steps = 2 ** layerCount - 1;
    
    // `flat` paints every face like the front (3d5's solid is a single colour,
    // and reads as a box only because of the white contour between the faces),
    
    // shade this renderer has always used — see SMARTART_LIGHT_RIG_PAINT.
    
    // flooded with their own colour after it, so the body carries BOTH factors.
    const paint = smartArtLightRigPaint(scene3d);
    const lit = paint?.body
      ? paint.body.map((channel, index) => channel * paint.face[index])
      : null;
    // A measured ratio of DRAWN pixels (side face over front face), so it
    // scales the encoded channels. It is not DrawingML's `a:shade`, which
    // mixes linear light and would paint the body a third lighter.
    const shade = lit ? (lit[0] + lit[1] + lit[2]) / 3 : 0.62;
    if (fill === "none" || fill === "transparent") {
      // A solid with No Fill has see-through walls: PowerPoint draws only its
      
      // three unfilled 30 pt deep boxes under their group's Isometric: Off
      // Axis 2 Left: white wireframes). Swept, the outline became a slab, and
      // a drop shadow of an unfilled shape took black for its colour.
      const edge = shape3d.contourColor;
      return edge && (shape3d.contourWidthEmu ?? 0) > 0
        ? `drop-shadow(${offset.dx.toFixed(2)}px ${offset.dy.toFixed(2)}px 0 ${scaleDrawnColor(edge, shade)})`
        : null;
    }
    const face = scaleDrawnColor(fill ?? "#000000", shade);
    for (let layer = 0; layer < layerCount; layer += 1) {
      const reach = 2 ** layer / steps;
      layers.push(
        `drop-shadow(${(offset.dx * reach).toFixed(2)}px ${(offset.dy * reach).toFixed(2)}px 0 ${face})`,
      );
    }
  }
  return layers.length ? layers.join(" ") : null;
}
























export function smartArtBevelRim(
  shape3d: SmartArtShape3d,
): { widthPx: number; light: string; dark: string } | null {
  const bevel = shape3d.bevelTop ?? shape3d.bevelBottom;
  if (!bevel || !(bevel.widthEmu > 0)) return null;
  const contrast = MATERIAL_CONTRAST[shape3d.material ?? "matte"];
  return {
    widthPx: emuToPx(bevel.widthEmu),
    light: `rgba(255,255,255,${contrast.toFixed(2)})`,
    dark: `rgba(0,0,0,${(contrast * 0.7).toFixed(2)})`,
  };
}

/**
 * The SVG `<filter>` primitives that paint {@link smartArtBevelRim} — one band
 * hugging the top-left inner edge and one the bottom-right, each clipped to the
 * shape's real alpha.
 *
 * `SourceAlpha` minus a copy of itself shifted by `d` is exactly the part of the
 * shape the shifted copy does not cover: an L of width `d` along the two edges
 * the shift came from. That is an inner rim on any silhouette — which is why
 * this cannot be a CSS `drop-shadow`, whose every layer lands OUTSIDE the shape.
 * The outer half and the inner half of each side come from two shifts, `w` and
 * `w/2`, so the band ramps instead of reading as a printed border.
 *
 * Returned as data rather than markup so the three surfaces that draw SmartArt
 * can each render it into their own `<defs>` with their own id.
 */
export interface SmartArtRimPrimitive {
  /** Shift that carves this band, in the member's own px. */
  dx: number;
  dy: number;
  color: string;
}

export function smartArtBevelRimPrimitives(
  rim: { widthPx: number; light: string; dark: string },
): SmartArtRimPrimitive[] {
  const outer = rim.widthPx;
  const inner = rim.widthPx / 2;
  // Each side is painted twice: the whole band, then its outer half over the
  // top. A rounded inset turns away from the light halfway across, so the two
  // halves take opposite sides of the face and the band averages back to it.
  return [
    // Bottom-right: shifted up-left, so the uncovered L is the bottom and
    // right sides. Outer half away from the light, inner half towards it.
    { dx: -outer, dy: -outer, color: rim.light },
    { dx: -inner, dy: -inner, color: rim.dark },
    // Top-left, drawn second so it owns the shared corners.
    { dx: outer, dy: outer, color: rim.dark },
    { dx: inner, dy: inner, color: rim.light },
  ];
}
















































export interface SmartArtBevelLight {
  /** The silhouette's blur that becomes the bevel's height field, in px. */
  readonly sigmaPx: number;
  /** `surfaceScale`: how steep that height field is. */
  readonly surfaceScale: number;
  /** The key light, SVG convention: degrees clockwise from +x, y down. */
  readonly azimuth: number;
  readonly elevation: number;
  /** The share of a flat face's light that does not depend on the slope. */
  readonly ambient: number;
  readonly specular: number;
  readonly specularExponent: number;
  /** The share of the light coming from the key light's mirror about the vertical. */
  readonly mirror: number;
  /**
   * The `dkEdge` material's dark rim, where it was measured: the face darkened
   * by `depth` times `slope` x (1 - the silhouette blurred by `sigmaPx`), so
   * the outermost pixel goes darkest and the face a pixel in keeps its colour.
   */
  readonly darkEdge?: { readonly sigmaPx: number; readonly slope: number; readonly depth: number };
  /**
   * The bevel drawn side by side, where Windows' own render of the style was
   * read side by side ({@link SmartArtBevelEdges}): in place of the light, on
   * a member whose outline is a box (see {@link smartArtBevelLight}).
   */
  readonly edges?: SmartArtBevelEdges;
  /**
   * How far inside the painted silhouette the member's own edge stands, in
   * px, where it wears a contour: the contour is stroked on the outline, half
   * of it outside, and the sides' depths are read from the member's edge
   * (Windows draws 3d3's contour wholly outside the bevel, probes/sa36-s2b).
   */
  readonly edgeInsetPx?: number;
}

/**
 * How far each side of a bevelled box goes toward white (positive) or black
 * (negative) by its depth in from that side's edge, every quarter px from 0
 * (px of the member's own space), as Windows PowerPoint draws it; past the
 * last entry the face keeps its colour.
 */
export interface SmartArtBevelEdges {
  readonly top: readonly number[];
  readonly right: readonly number[];
  readonly bottom: readonly number[];
  readonly left: readonly number[];
}

/**
 * 3d1 Polished's bevel (plastic, bevelT 120900 x 88900 EMU) side by side, read
 * off Windows PowerPoint's own pictures at 2.67 px/pt (truth office-truth/
 * probes/sa36-s2: Basic Block List, Basic Process and Vertical Bullet List in
 * two fills, every side the same to 0.03 across all six): a soft highlight
 * 3.5 px in along the top, the right edge's outer 3 px nearly white, a faint
 * one along the left, nothing but a dark hairline along the bottom. The light
 * it replaces on a box drew only the top (bevel bands 11.2 -> 8.3/255; held
 * out on 3dcam's Basic Block List, 13.1 -> 9.5).
 */
const POLISHED_EDGES: SmartArtBevelEdges = {
  top: [0.036, 0.038, 0.042, 0.065, 0.102, 0.164, 0.241, 0.322, 0.407, 0.479, 0.538, 0.592, 0.644, 0.686, 0.719, 0.714, 0.685, 0.653, 0.619, 0.58, 0.535, 0.489, 0.441, 0.394, 0.346, 0.297, 0.248, 0.2, 0.154, 0.11, 0.069, 0.045, 0.036, 0.031, 0.03, 0.027, 0.023, 0.02, 0.018, 0.015, 0.005],
  right: [0.583, 0.615, 0.823, 0.946, 0.97, 0.96, 0.909, 0.835, 0.739, 0.636, 0.525, 0.409, 0.29, 0.181, 0.088, 0.035, 0.023, 0.017, 0.014, 0.012, 0.012, 0.012, 0.011, 0.009, 0.007, 0.006, 0.005, 0.004, 0.001],
  bottom: [-0.062, -0.062, -0.058, -0.049, -0.044, -0.04, -0.032, -0.018],
  left: [0, 0, 0, 0.004, 0.012, 0.024, 0.039, 0.052, 0.066, 0.078, 0.088, 0.097, 0.106, 0.116, 0.123, 0.12, 0.113, 0.106, 0.099, 0.092, 0.085, 0.077, 0.067, 0.059, 0.051, 0.044, 0.037, 0.028, 0.018, 0.009, 0.001],
};

/**
 * 3d2 Inset's bevel (plastic, relaxedInset bevelT 127000 x 25400 EMU) side by
 * side, read the same way off probes/sa36-s2b (Basic Block List, Basic Process
 * and Vertical Bullet List in two fills, every side within 0.05 across the
 * six): a dark lip along the top and left, the bottom dark 11 px in, the
 * right edge's outer 6 px lit toward white, and along the left the groove's
 * highlight 6.5-11 px in -- the inset groove the light could not draw (bevel
 * bands 16.7 -> 10.2/255; Basic Cycle's circles 13.9 -> 9.4).
 */
const INSET_EDGES: SmartArtBevelEdges = {
  top: [
    -0.26, -0.243, -0.213, -0.171, -0.126, -0.091, -0.061, -0.045, -0.039, -0.033, -0.027, -0.019, -0.011, -0.006,
    -0.002, 0, 0.001, 0.002, 0.004, 0.005, 0.006, 0.007, 0.008, 0.009, 0.011, 0.012, 0.011, 0.01, 0.009, 0.009,
    0.01, 0.01, 0.009, 0.01, 0.01, 0.011, 0.012, 0.014, 0.016, 0.018, 0.02, 0.021, 0.021, 0.022, 0.023, 0.024,
    0.024, 0.023, 0.011
  ],
  right: [
    0.177, 0.123, 0.337, 0.528, 0.676, 0.77, 0.8, 0.802, 0.773, 0.745, 0.716, 0.68, 0.638, 0.602, 0.572, 0.542,
    0.511, 0.477, 0.441, 0.408, 0.376, 0.338, 0.298, 0.261, 0.221, 0.163, 0.087, 0.015, -0.038, -0.063, -0.065,
    -0.066, -0.068, -0.068, -0.069, -0.07, -0.072, -0.075, -0.079, -0.082, -0.084, -0.083, -0.079, -0.074, -0.069,
    -0.065, -0.059, -0.05, -0.022
  ],
  bottom: [
    -0.431, -0.357, -0.283, -0.24, -0.206, -0.18, -0.158, -0.141, -0.126, -0.115, -0.102, -0.087, -0.073, -0.062,
    -0.052, -0.049, -0.046, -0.039, -0.032, -0.025, -0.02, -0.019, -0.018, -0.018, -0.021, -0.034, -0.054, -0.084,
    -0.108, -0.111, -0.109, -0.107, -0.105, -0.102, -0.1, -0.099, -0.096, -0.092, -0.088, -0.086, -0.082, -0.073,
    -0.065, -0.058, -0.052, -0.049, -0.045, -0.034, -0.013
  ],
  left: [
    -0.406, -0.341, -0.233, -0.133, -0.039, 0.014, 0.034, 0.034, 0.022, 0.009, -0.004, -0.018, -0.034, -0.052,
    -0.071, -0.077, -0.076, -0.077, -0.078, -0.078, -0.078, -0.077, -0.077, -0.077, -0.069, 0.019, 0.172, 0.365,
    0.545, 0.609, 0.611, 0.605, 0.6, 0.592, 0.58, 0.551, 0.512, 0.471, 0.428, 0.383, 0.337, 0.288, 0.238, 0.194,
    0.154, 0.118, 0.086, 0.068, 0.029
  ],
};

/**
 * 3d3 Cartoon's bevel (metal, bevelT 88900 x 203200 EMU, a 1.5 pt contour)
 * side by side, off probes/sa36-s2b: every side lit toward white from its
 * edge for 3 px, and along the left and right a second, brighter band 3-7 px
 * in (bevel bands 26.5 -> 19.0/255; Basic Cycle's circles 19.3 -> 13.5).
 */
const CARTOON_EDGES: SmartArtBevelEdges = {
  top: [
    0.499, 0.488, 0.468, 0.42, 0.356, 0.304, 0.262, 0.226, 0.194, 0.165, 0.14, 0.119, 0.103, 0.088, 0.072, 0.056,
    0.042, 0.034, 0.03, 0.025, 0.02, 0.014, 0.008, 0.005, 0.005, 0.004, 0.003, 0.002, 0.001, 0.001
  ],
  right: [
    0.264, 0.498, 0.461, 0.419, 0.37, 0.319, 0.272, 0.236, 0.211, 0.188, 0.17, 0.179, 0.216, 0.254, 0.293, 0.335,
    0.384, 0.46, 0.561, 0.648, 0.718, 0.775, 0.807, 0.778, 0.689, 0.587, 0.474, 0.343, 0.206, 0.108, 0.053, 0.026,
    0.021, 0.017, 0.012, 0.008, 0.005, 0.002, 0.001
  ],
  bottom: [
    0.494, 0.477, 0.447, 0.388, 0.328, 0.28, 0.239, 0.207, 0.178, 0.152, 0.127, 0.106, 0.087, 0.071, 0.057, 0.046,
    0.035, 0.028, 0.023, 0.018, 0.014, 0.01, 0.006, 0.005, 0.004, 0.003, 0.003, 0.002, 0.001
  ],
  left: [
    0.499, 0.481, 0.454, 0.412, 0.359, 0.31, 0.264, 0.231, 0.205, 0.183, 0.165, 0.182, 0.224, 0.262, 0.296, 0.341,
    0.396, 0.48, 0.581, 0.66, 0.724, 0.775, 0.812, 0.773, 0.686, 0.574, 0.448, 0.309, 0.169, 0.087, 0.043, 0.026,
    0.021, 0.016, 0.012, 0.008, 0.005, 0.002, 0.001
  ],
};

/**
 * 3d4 Powder's bevel (translucentPowder, softRound bevelT 127000 x 25400 EMU)
 * side by side, off probes/sa36-s2b: a bright lip at every edge, along the
 * top and right lit toward white 8 px in, along the bottom and left only 2 px
 * and then a little darker (bevel bands 6.3 -> 5.6/255; Basic Cycle's circles
 * 5.5 -> 4.5).
 */
const POWDER_EDGES: SmartArtBevelEdges = {
  top: [
    0.678, 0.565, 0.385, 0.242, 0.125, 0.059, 0.024, 0.01, 0.011, 0.049, 0.109, 0.156, 0.192, 0.209, 0.212, 0.21,
    0.204, 0.201, 0.2, 0.195, 0.188, 0.182, 0.174, 0.163, 0.15, 0.14, 0.131, 0.122, 0.113, 0.107, 0.103, 0.096,
    0.087, 0.079, 0.074, 0.069, 0.066, 0.058, 0.047, 0.039, 0.032, 0.027, 0.024, 0.022, 0.02, 0.018, 0.015, 0.013,
    0.006
  ],
  right: [
    0.753, 0.604, 0.419, 0.271, 0.167, 0.089, 0.043, 0.019, 0.026, 0.061, 0.122, 0.184, 0.241, 0.272, 0.273, 0.267,
    0.259, 0.252, 0.247, 0.24, 0.232, 0.219, 0.204, 0.19, 0.179, 0.169, 0.16, 0.151, 0.142, 0.134, 0.126, 0.118,
    0.108, 0.099, 0.092, 0.088, 0.083, 0.073, 0.061, 0.051, 0.044, 0.039, 0.035, 0.031, 0.028, 0.024, 0.021, 0.018,
    0.008
  ],
  bottom: [
    0.761, 0.668, 0.57, 0.491, 0.405, 0.295, 0.181, 0.077, -0.002, -0.023, -0.033, -0.043, -0.052, -0.056, -0.059,
    -0.061, -0.062, -0.061, -0.061, -0.064, -0.065, -0.062, -0.059, -0.057, -0.055, -0.054, -0.053, -0.052, -0.05,
    -0.046, -0.043, -0.044, -0.043, -0.04, -0.036, -0.032, -0.03, -0.031, -0.03, -0.028, -0.025, -0.023, -0.021,
    -0.019, -0.017, -0.015, -0.014, -0.011, -0.004
  ],
  left: [
    0.825, 0.767, 0.675, 0.598, 0.527, 0.42, 0.288, 0.162, 0.046, -0.008, -0.023, -0.035, -0.046, -0.052, -0.055,
    -0.057, -0.059, -0.059, -0.057, -0.057, -0.058, -0.061, -0.064, -0.065, -0.065, -0.064, -0.064, -0.064, -0.063,
    -0.06, -0.056, -0.054, -0.052, -0.05, -0.047, -0.045, -0.042, -0.042, -0.043, -0.042, -0.039, -0.036, -0.034,
    -0.03, -0.027, -0.024, -0.021, -0.018, -0.007
  ],
};

/**
 * The outline a bevel read side by side may use: the member's preset, when it
 * stands unturned and unmirrored (Windows lights the slide's sides, so a turned
 * box would carry its highlight round with it); else null.
 */
export function smartArtBevelOutline(
  preset: string | null | undefined,
  turn: { readonly rotation?: unknown; readonly flipH?: unknown; readonly flipV?: unknown },
): string | null {
  const rotation = typeof turn.rotation === "number" ? ((turn.rotation % 360) + 360) % 360 : 0;
  if (!preset || rotation > 0.01 || turn.flipH === true || turn.flipV === true) return null;
  return preset;
}

/** The bevels read side by side, keyed like {@link SMARTART_BEVEL_LIGHTS}. */
const SMARTART_BEVEL_EDGES: Readonly<Record<string, SmartArtBevelEdges>> = {
  "plastic:circle:120900:88900": POLISHED_EDGES,
  "plastic:relaxedInset:127000:25400": INSET_EDGES,
  "metal:circle:88900:203200": CARTOON_EDGES,
  "translucentPowder:softRound:127000:25400": POWDER_EDGES,
};

/**
 * The bevels whose side-by-side reading also fits an ellipse, read at its four
 * sides' middles (Basic Cycle's circles measured nearer than the light: 3d2,
 * 3d3 and 3d4; 3d1's did not).
 */
const ROUND_BEVEL_EDGES: ReadonlySet<string> = new Set([
  "plastic:relaxedInset:127000:25400",
  "metal:circle:88900:203200",
  "translucentPowder:softRound:127000:25400",
]);

/**
 * Outlines that are a box: four straight sides along the member's own axes
 * (DrawingML preset names and the editors' own kind names), the ones a bevel
 * read side by side fits. A circle's rim is lit round its curve, which this
 * does not draw (3d1 on Basic Cycle's circles: the light 10.5, side by side
 * 10.6), so every other outline keeps the light -- except an ellipse under a
 * bevel of {@link ROUND_BEVEL_EDGES}.
 */
const BOX_OUTLINES: ReadonlySet<string> = new Set([
  "rect", "rectangle", "roundRect", "rounded-rectangle",
  "round1Rect", "single-round-corner", "round2SameRect", "same-side-corners-rounded",
  "round2DiagRect", "diagonal-corners-rounded", "snip1Rect", "single-corner-snipped",
  "snip2SameRect", "same-side-corners-snipped", "snip2DiagRect", "diagonal-corners-snipped",
  "snipRoundRect", "single-corner-rounded-snipped",
  "flowChartProcess", "flowchart-process", "flowChartAlternateProcess", "flowchart-alternate-process",
]);

const DEFAULT_BEVEL_HEIGHT_EMU = 76_200;

function bevelKey(shape3d: SmartArtShape3d): string | null {
  const bevel = shape3d.bevelTop ?? shape3d.bevelBottom;
  if (!bevel || !(bevel.widthEmu > 0)) return null;
  return [
    shape3d.material ?? "",
    bevel.preset ?? "circle",
    bevel.widthEmu,
    bevel.heightEmu ?? DEFAULT_BEVEL_HEIGHT_EMU,
  ].join(":");
}

const light = (
  sigmaPx: number,
  surfaceScale: number,
  azimuth: number,
  elevation: number,
  ambient: number,
  specular: number,
  specularExponent: number,
  mirror: number,
): SmartArtBevelLight => ({
  sigmaPx, surfaceScale, azimuth, elevation, ambient, specular, specularExponent, mirror,
});

/** Fitted per style (see {@link SmartArtBevelLight}), keyed by the style's own bevel. */
const SMARTART_BEVEL_LIGHTS: Readonly<Record<string, SmartArtBevelLight>> = {
  // simple3 (Subtle Effect): dkEdge, bevelT w 8200 h 38100 -- under a pixel
  // wide, and no light at all: Windows PowerPoint draws the face its own
  // colour to the edge and a dark rim round it, the outermost pixel at about
  // a third of the face (truth office-truth/probes/sa36-s2: Basic Block List,
  // Basic Process and Vertical Bullet List in two fills at 2.67 px/pt, their
  // bevel bands within 36 px of each edge 9.8 -> 5.4/255 against the lit
  // slope this used to be, which brightened the top 24 px of every member by
  // 13; held out on Basic Cycle's circles, 4.4 -> 2.4).
  "dkEdge:circle:8200:38100": {
    ...light(0, 0, 315, 31, 0, 0, 30, 0),
    darkEdge: { sigmaPx: 0.511, slope: 1.73, depth: 0.997 },
  },
  // 3d1 Polished.
  "plastic:circle:120900:88900": light(4.55, 3.79, 255.07, 42.82, 0.956, 0.675, 60, 0.215),
  // 3d2 Inset.
  "plastic:relaxedInset:127000:25400": light(4.15, 4.12, 315, 21.59, 0.911, 2.5, 31.87, 0),
  // 3d3 Cartoon: lit from both sides, so the mirror light carries almost half.
  "metal:circle:88900:203200": light(2.78, 2.88, 225, 15, 1, 2.5, 26.31, 0.448),
  // 3d4 Powder.
  "translucentPowder:softRound:127000:25400": light(4.43, 13.48, 315, 15, 0.924, 0, 58.48, 0),
  // 3d6 Bird's Eye Scene: PowerPoint draws its face almost flat.
  "plastic:circle:50800:50800": light(1.17, 0.32, 315, 15.03, 0.533, 0, 59.53, 0),
  // 3d7 Sunset Scene.
  "metal:relaxedInset:101600:80600": light(2.4, 30, 315, 47.32, 0.506, 0, 41.78, 0),
  // 3d8 Flat Scene.
  "matte:relaxedInset:120650:38100": light(4.84, 8.65, 310.06, 40.89, 1, 1.738, 60, 0),
};








export function smartArtBevelLight(
  shape3d: SmartArtShape3d,
  /**
   * The member's outline (its preset), when it stands unturned and unmirrored:
   * a box takes the bevel read side by side where there is one. Without it,
   * or for any other outline, the light.
   */
  outline?: string | null,
): SmartArtBevelLight | null {
  const key = bevelKey(shape3d);
  if (!key) return null;
  const measured = SMARTART_BEVEL_LIGHTS[key];
  const edges = outline && (BOX_OUTLINES.has(outline) || (outline === "ellipse" && ROUND_BEVEL_EDGES.has(key)))
    ? SMARTART_BEVEL_EDGES[key]
    : undefined;
  if (measured && edges) {
    const contour = shape3d.contourWidthEmu && shape3d.contourWidthEmu > 0 ? emuToPx(shape3d.contourWidthEmu) / 2 : 0;
    return contour > 0 ? { ...measured, edges, edgeInsetPx: contour } : { ...measured, edges };
  }
  if (measured) return measured;
  const bevel = (shape3d.bevelTop ?? shape3d.bevelBottom)!;
  const widthPx = emuToPx(bevel.widthEmu);
  const heightPx = emuToPx(bevel.heightEmu ?? DEFAULT_BEVEL_HEIGHT_EMU);
  return light(
    Math.max(0.75, Math.min(6, widthPx / 3)),
    Math.max(1, Math.min(12, (heightPx / Math.max(1, widthPx)) * 6)),
    315,
    35,
    0.85,
    0.4 * (MATERIAL_CONTRAST[shape3d.material ?? "matte"] / MATERIAL_CONTRAST.plastic),
    40,
    0,
  );
}

/**
 * One SVG filter primitive as data — tag, attributes (their SVG names) and
 * children — so the React canvases and the Docs canvas (plain DOM) build the
 * same filter from one description.
 */
export interface SmartArtFilterPrimitive {
  readonly tag: string;
  readonly attrs: Readonly<Record<string, string | number>>;
  readonly children?: readonly SmartArtFilterPrimitive[];
}

const fixed = (value: number, digits = 4) => Number(value.toFixed(digits));






export function smartArtBevelLightPrimitives(
  light: SmartArtBevelLight,
  input = "SourceGraphic",
  /**
   * What the rim is raised from: the painted face's own alpha, unless the
   * host names the GEOMETRY's (a picture fill, see `OfficeSmartArtBevelFilter`).
   */
  height = "SourceAlpha",
  /**
   * The material's own face opacity (`smartArtMaterialOpacity`), laid under a
   * bevel read side by side; the caller then adds no closing opacity step
   * (`smartArtMaterialOpacityAfterBevel`). The light ignores it.
   */
  faceOpacity: number | null = null,
): SmartArtFilterPrimitive[] {
  if (light.edges) return bevelEdgePrimitives(light.edges, input, height, light.edgeInsetPx ?? 0, faceOpacity);
  if (light.surfaceScale <= 0 && light.specular <= 0.001) {
    // No slope to light (simple3): the face as it is, with its dark rim.
    return [
      {
        tag: "feColorMatrix",
        attrs: { in: input, type: "matrix", values: "1 0 0 0 0  0 1 0 0 0  0 0 1 0 0  0 0 0 0 1", result: "bevelFace" },
      },
      ...darkEdgePrimitives(light, "bevelFace", height),
      { tag: "feComposite", attrs: { in: light.darkEdge ? "edgeLit" : "bevelFace", in2: "SourceAlpha", operator: "in", result: "bevelResult" } },
    ];
  }
  const elevation = Math.max(5, Math.min(90, light.elevation));
  const sin = Math.sin((elevation * Math.PI) / 180);
  const slope = (1 - light.ambient) / sin;
  const lights = [
    { azimuth: light.azimuth, weight: 1 - light.mirror },
    { azimuth: (540 - light.azimuth) % 360, weight: light.mirror },
  ].filter((entry) => entry.weight > 0.001);
  const distant = (azimuth: number): SmartArtFilterPrimitive => ({
    tag: "feDistantLight",
    attrs: { azimuth: fixed(azimuth, 2), elevation: fixed(elevation, 2) },
  });
  const out: SmartArtFilterPrimitive[] = [
    {
      tag: "feGaussianBlur",
      attrs: { in: height, stdDeviation: fixed(light.sigmaPx, 3), result: "bevelHeight" },
    },
  ];
  // The shade is built at HALF scale and doubled when it meets the face, so a
  // slope facing the light can brighten the face (a lighting result is capped
  // at 1). Each map is opaque, so these sums need no premultiplied care.
  const maps: string[] = [];
  lights.forEach((entry, index) => {
    const gain = (slope * entry.weight) / 2;
    const offset = index === 0 ? light.ambient / 2 : 0;
    out.push(
      {
        tag: "feDiffuseLighting",
        attrs: {
          in: "bevelHeight",
          surfaceScale: fixed(light.surfaceScale, 3),
          diffuseConstant: 1,
          "lighting-color": "#ffffff",
          result: `bevelDiffuse${index}`,
        },
        children: [distant(entry.azimuth)],
      },
      {
        tag: "feColorMatrix",
        attrs: {
          in: `bevelDiffuse${index}`,
          type: "matrix",
          values: `${fixed(gain)} 0 0 0 ${fixed(offset)}  0 ${fixed(gain)} 0 0 ${fixed(offset)}  0 0 ${fixed(gain)} 0 ${fixed(offset)}  0 0 0 0 1`,
          result: `bevelShade${index}`,
        },
      },
    );
    maps.push(`bevelShade${index}`);
  });
  let shade = maps[0];
  if (maps.length > 1) {
    out.push({
      tag: "feComposite",
      attrs: { in: maps[0], in2: maps[1], operator: "arithmetic", k1: 0, k2: 1, k3: 1, k4: 0, result: "bevelShade" },
    });
    shade = "bevelShade";
  }
  // The face made opaque first, so its anti-aliased edge is shaded in its own
  // colour and only then given its coverage back by the final clip.
  out.push(
    {
      tag: "feColorMatrix",
      attrs: {
        in: input,
        type: "matrix",
        values: "1 0 0 0 0  0 1 0 0 0  0 0 1 0 0  0 0 0 0 1",
        result: "bevelFace",
      },
    },
    {
      tag: "feComposite",
      attrs: { in: "bevelFace", in2: shade, operator: "arithmetic", k1: 2, k2: 0, k3: 0, k4: 0, result: "bevelLit" },
    },
  );
  let lit = "bevelLit";
  if (light.specular > 0.001) {
    lights.forEach((entry, index) => {
      out.push(
        {
          tag: "feSpecularLighting",
          attrs: {
            in: "bevelHeight",
            surfaceScale: fixed(light.surfaceScale, 3),
            specularConstant: fixed(light.specular * entry.weight),
            specularExponent: fixed(Math.max(1, Math.min(128, light.specularExponent)), 2),
            "lighting-color": "#ffffff",
            result: `bevelGloss${index}`,
          },
          children: [distant(entry.azimuth)],
        },
        {
          tag: "feComposite",
          attrs: { in: lit, in2: `bevelGloss${index}`, operator: "arithmetic", k1: 0, k2: 1, k3: 1, k4: 0, result: `bevelShine${index}` },
        },
      );
      lit = `bevelShine${index}`;
    });
  }
  if (light.darkEdge) {
    out.push(...darkEdgePrimitives(light, lit, height));
    lit = "edgeLit";
  }
  out.push({
    tag: "feComposite",
    attrs: { in: lit, in2: "SourceAlpha", operator: "in", result: "bevelResult" },
  });
  return out;
}






function coveragePrimitives(height: string): SmartArtFilterPrimitive[] {
  return [
    {
      tag: "feComponentTransfer",
      attrs: { in: height, result: "edgeCoverage" },
      children: [{ tag: "feFuncA", attrs: { type: "linear", slope: 50, intercept: 0 } }],
    },
  ];
}

/** Depth fields of {@link bevelEdgePrimitives}: the alpha blurred along a side's normal by this much... */
const EDGE_FIELD_SIGMA = 3;
/** ...and moved this far in, so the field runs from 0 just outside the edge to 1 well inside. */
const EDGE_FIELD_INSET = 4.5;
const EDGE_TABLE_SIZE = 64;

/** The standard normal distribution's quantile (Acklam's rational approximation, |error| < 1.2e-9). */
function normalQuantile(p: number): number {
  if (p <= 0) return -Infinity;
  if (p >= 1) return Infinity;
  const a = [-39.6968302866538, 220.946098424521, -275.928510446969, 138.357751867269, -30.6647980661472, 2.50662827745924];
  const b = [-54.4760987982241, 161.585836858041, -155.698979859887, 66.8013118877197, -13.2806815528857];
  const c = [-0.00778489400243029, -0.322396458041136, -2.40075827716184, -2.54973253934373, 4.37466414146497, 2.93816398269878];
  const d = [0.00778469570904146, 0.32246712907004, 2.445134137143, 3.75440866190742];
  const tail = (q: number) =>
    (((((c[0]! * q + c[1]!) * q + c[2]!) * q + c[3]!) * q + c[4]!) * q + c[5]!) / ((((d[0]! * q + d[1]!) * q + d[2]!) * q + d[3]!) * q + 1);
  if (p < 0.02425) return tail(Math.sqrt(-2 * Math.log(p)));
  if (p > 1 - 0.02425) return -tail(Math.sqrt(-2 * Math.log(1 - p)));
  const q = p - 0.5;
  const r = q * q;
  return ((((((a[0]! * r + a[1]!) * r + a[2]!) * r + a[3]!) * r + a[4]!) * r + a[5]!) * q) /
    (((((b[0]! * r + b[1]!) * r + b[2]!) * r + b[3]!) * r + b[4]!) * r + 1);
}

/**
 * A side's profile as the `feFuncA` table over its depth field: entry i is
 * the field's value i/(n-1), the depth it stands at INSET + SIGMA * quantile,
 * and the profile there (`sign` keeps one direction: +1 toward white, -1
 * toward black).
 */
function edgeTable(profile: readonly number[], sign: 1 | -1): string | null {
  const values: number[] = [];
  for (let index = 0; index < EDGE_TABLE_SIZE; index += 1) {
    const depth = EDGE_FIELD_INSET + EDGE_FIELD_SIGMA * normalQuantile(index / (EDGE_TABLE_SIZE - 1));
    const at = depth / 0.25;
    const low = Math.floor(at);
    const value = depth < 0 || low >= profile.length
      ? 0
      : (profile[low] ?? 0) + ((profile[low + 1] ?? 0) - (profile[low] ?? 0)) * (at - low);
    values.push(Math.max(0, Math.min(1, sign * value)));
  }
  return values.some((value) => value > 0.002) ? values.map((value) => fixed(value, 4)).join(" ") : null;
}

/**
 * {@link SmartArtBevelEdges} on `input`: per side, the alpha blurred along that
 * side's normal only and moved inward is a depth field that reads 0 outside
 * the edge and rises to 1 inside, unaffected by the opposite side; the side's
 * profile, as a table over that field, is how much white (or black) is laid
 * over the face there. Clipped to the shape's alpha at the end.
 */
function bevelEdgePrimitives(
  edges: SmartArtBevelEdges,
  input: string,
  height: string,
  edgeInsetPx = 0,
  /**
   * The face's own opacity when its material sets one (`smartArtMaterialOpacity`):
   * the sides are laid on the face as Windows composites it -- read off its
   * pictures toward white or black FROM that composited face -- so the face is
   * made translucent first and the member's coverage, not the fill's alpha,
   * closes it.
   */
  faceOpacity: number | null = null,
): SmartArtFilterPrimitive[] {
  const inset = fixed(EDGE_FIELD_INSET + Math.max(0, edgeInsetPx), 3);
  const out: SmartArtFilterPrimitive[] = [
    {
      tag: "feColorMatrix",
      attrs: { in: input, type: "matrix", values: "1 0 0 0 0  0 1 0 0 0  0 0 1 0 0  0 0 0 0 1", result: "edgeFace0" },
    },
  ];
  // Depth from where the member IS, whatever its fill's alpha (`coveragePrimitives`).
  out.push(...coveragePrimitives(height));
  let face = "edgeFace0";
  if (faceOpacity !== null) {
    out.push(
      {
        tag: "feColorMatrix",
        attrs: { in: "edgeFace0", type: "matrix", values: "1 0 0 0 0  0 1 0 0 0  0 0 1 0 0  0 0 0 0 1", result: "edgeFaceOpaque" },
      },
      { tag: "feComposite", attrs: { in: "edgeFaceOpaque", in2: "edgeCoverage", operator: "in", result: "edgeFaceCovered" } },
      {
        tag: "feComponentTransfer",
        attrs: { in: "edgeFaceCovered", result: "edgeFaceMaterial" },
        children: [{ tag: "feFuncA", attrs: { type: "linear", slope: faceOpacity, intercept: 0 } }],
      },
    );
    face = "edgeFaceMaterial";
  }
  let step = 0;
  const sides = [
    ["top", 0, inset, `0 ${EDGE_FIELD_SIGMA}`],
    ["bottom", 0, -inset, `0 ${EDGE_FIELD_SIGMA}`],
    ["left", inset, 0, `${EDGE_FIELD_SIGMA} 0`],
    ["right", -inset, 0, `${EDGE_FIELD_SIGMA} 0`],
  ] as const;
  for (const [side, dx, dy, deviation] of sides) {
    const tables = ([1, -1] as const)
      .map((sign) => [sign, edgeTable(edges[side], sign)] as const)
      .filter((entry): entry is readonly [1 | -1, string] => entry[1] !== null);
    if (!tables.length) continue;
    out.push(
      { tag: "feGaussianBlur", attrs: { in: "edgeCoverage", stdDeviation: deviation, result: `${side}Blur` } },
      { tag: "feOffset", attrs: { in: `${side}Blur`, dx, dy, result: `${side}Depth` } },
    );
    for (const [sign, table] of tables) {
      const name = `${side}${sign > 0 ? "Light" : "Dark"}`;
      step += 1;
      out.push(
        {
          tag: "feComponentTransfer",
          attrs: { in: `${side}Depth`, result: `${name}Amount` },
          children: [{ tag: "feFuncA", attrs: { type: "table", tableValues: table } }],
        },
        { tag: "feFlood", attrs: { "flood-color": sign > 0 ? "#ffffff" : "#000000", result: `${name}Colour` } },
        { tag: "feComposite", attrs: { in: `${name}Colour`, in2: `${name}Amount`, operator: "in", result: `${name}Layer` } },
        { tag: "feComposite", attrs: { in: `${name}Layer`, in2: face, operator: "over", result: `edgeFace${step}` } },
      );
      face = `edgeFace${step}`;
    }
  }
  out.push({
    tag: "feComposite",
    attrs: { in: face, in2: faceOpacity !== null ? "edgeCoverage" : "SourceAlpha", operator: "in", result: "bevelResult" },
  });
  return out;
}

/** {@link SmartArtBevelLight.darkEdge} over `input`, as `edgeLit` (nothing when the light has none). */
function darkEdgePrimitives(light: SmartArtBevelLight, input: string, height: string): SmartArtFilterPrimitive[] {
  const edge = light.darkEdge;
  if (!edge) return [];
  const depth = fixed(-edge.depth);
  return [
    ...coveragePrimitives(height),
    { tag: "feGaussianBlur", attrs: { in: "edgeCoverage", stdDeviation: fixed(edge.sigmaPx, 3), result: "edgeBlur" } },
    {
      tag: "feComponentTransfer",
      attrs: { in: "edgeBlur", result: "edgeNear" },
      children: [{ tag: "feFuncA", attrs: { type: "linear", slope: fixed(-edge.slope), intercept: fixed(edge.slope) } }],
    },
    {
      tag: "feColorMatrix",
      attrs: {
        in: "edgeNear",
        type: "matrix",
        values: `0 0 0 ${depth} 1  0 0 0 ${depth} 1  0 0 0 ${depth} 1  0 0 0 0 1`,
        result: "edgeShade",
      },
    },
    { tag: "feComposite", attrs: { in: input, in2: "edgeShade", operator: "arithmetic", k1: 1, k2: 0, k3: 0, k4: 0, result: "edgeLit" } },
  ];
}

/**
 * The outer stroke that stands in for a 3-D contour, in points. Only used when
 * the style itself draws no line, so a real line is never overwritten.
 */
export function smartArtContourStroke(
  shape3d: SmartArtShape3d,
): { stroke: string; strokeWidth: number } | null {
  if (!shape3d.contourWidthEmu || shape3d.contourWidthEmu <= 0) return null;
  return {
    // No `a:contourClr` is black: PowerPoint draws 3d3's contour as a dark
    // line round the shape (#1d1d1d at its outer pixel), and the white this
    // used to default to drew a pale ghost outline round every Cartoon node.
    stroke: shape3d.contourColor ?? "#000000",
    // EMU → points (12700 EMU per point), floored at a hairline so it shows.
    strokeWidth: Math.max(0.5, shape3d.contourWidthEmu / 12700),
  };
}
