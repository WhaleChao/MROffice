/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
export const OFFICE_PICTURE_STYLE_CATALOG = [
  {
    id: "simple-frame-white",
    officeId: "SimpleFrameWhite",
    zh: "简单框架，白色",
    en: "Simple Frame, White",
  },
  {
    id: "beveled-matte-white",
    officeId: "BeveledMatteWhite",
    zh: "棱台亚光，白色",
    en: "Beveled Matte, White",
  },
  {
    id: "metal-frame",
    officeId: "MetalFrame",
    zh: "金属框架",
    en: "Metal Frame",
  },
  {
    id: "drop-shadow-rectangle",
    officeId: "DropShadowRectangle",
    zh: "矩形投影",
    en: "Drop Shadow Rectangle",
  },
  {
    id: "reflected-rounded-rectangle",
    officeId: "ReflectedRoundedRectangle",
    zh: "映像圆角矩形",
    en: "Reflected Rounded Rectangle",
  },
  {
    id: "soft-edge-rectangle",
    officeId: "SoftEdgeRectangle",
    zh: "柔化边缘矩形",
    en: "Soft Edge Rectangle",
  },
  {
    id: "double-frame-black",
    officeId: "DoubleFrameBlack",
    zh: "双框架，黑色",
    en: "Double Frame, Black",
  },
  {
    id: "thick-matte-black",
    officeId: "ThickMatteBlack",
    zh: "厚重亚光，黑色",
    en: "Thick Matte, Black",
  },
  {
    id: "simple-frame-black",
    officeId: "SimpleFrameBlack",
    zh: "简单框架，黑色",
    en: "Simple Frame, Black",
  },
  {
    id: "beveled-oval-black",
    officeId: "BeveledOvalBlack",
    zh: "棱台形椭圆，黑色",
    en: "Beveled Oval, Black",
  },
  {
    id: "compound-frame-black",
    officeId: "CompoundFrameBlack",
    zh: "复杂框架，黑色",
    en: "Compound Frame, Black",
  },
  {
    id: "moderate-frame-black",
    officeId: "ModerateFrameBlack",
    zh: "中等复杂框架，黑色",
    en: "Moderate Frame, Black",
  },
  {
    id: "center-shadow-rectangle",
    officeId: "CenterShadowRectangle",
    zh: "居中矩形阴影",
    en: "Center Shadow Rectangle",
  },
  {
    id: "rounded-diagonal-corner-white",
    officeId: "RoundedDiagonalCornerWhite",
    zh: "圆形对角，白色",
    en: "Rounded Diagonal Corner, White",
  },
  {
    id: "snip-diagonal-corner-white",
    officeId: "SnipDiagonalCornerWhite",
    zh: "剪裁对角线，白色",
    en: "Snip Diagonal Corner, White",
  },
  {
    id: "moderate-frame-white",
    officeId: "ModerateFrameWhite",
    zh: "中等复杂框架，白色",
    en: "Moderate Frame, White",
  },
  {
    id: "rotated-white",
    officeId: "RotatedWhite",
    zh: "旋转，白色",
    en: "Rotated, White",
  },
  {
    id: "perspective-shadow-white",
    officeId: "PerspectiveShadowWhite",
    zh: "透视阴影，白色",
    en: "Perspective Shadow, White",
  },
  {
    id: "relaxed-perspective-white",
    officeId: "RelaxedPerspectiveWhite",
    zh: "松散透视，白色",
    en: "Relaxed Perspective, White",
  },
  {
    id: "soft-edge-oval",
    officeId: "SoftEdgeOval",
    zh: "柔化边缘椭圆",
    en: "Soft Edge Oval",
  },
  {
    id: "bevel-edge-rectangle",
    officeId: "BevelEdgeRectangle",
    zh: "棱台矩形",
    en: "Bevel Rectangle",
  },
  {
    id: "bevel-perspective",
    officeId: "BevelPerspective",
    zh: "棱台透视",
    en: "Bevel Perspective",
  },
  {
    id: "reflected-perspective-right",
    officeId: "ReflectedPerspectiveRight",
    zh: "映像右透视",
    en: "Reflected Perspective Right",
  },
  {
    id: "bevel-perspective-left-white",
    officeId: "BevelPerspectiveLeftWhite",
    zh: "棱台左透视，白色",
    en: "Bevel Perspective Left, White",
  },
  {
    id: "reflected-bevel-black",
    officeId: "ReflectedBevelBlack",
    zh: "映像棱台，黑色",
    en: "Reflected Bevel, Black",
  },
  {
    id: "reflected-bevel-white",
    officeId: "ReflectedBevelWhite",
    zh: "映像棱台，白色",
    en: "Reflected Bevel, White",
  },
  {
    id: "metal-rounded-rectangle",
    officeId: "MetalRoundedRectangle",
    zh: "金属圆角矩形",
    en: "Metal Rounded Rectangle",
  },
  { id: "metal-oval", officeId: "MetalOval", zh: "金属椭圆", en: "Metal Oval" },
] as const;

export type OfficePictureStyleId =
  (typeof OFFICE_PICTURE_STYLE_CATALOG)[number]["id"];
export type OfficePictureStyleEntry =
  (typeof OFFICE_PICTURE_STYLE_CATALOG)[number];

const BY_ID = new Map(
  OFFICE_PICTURE_STYLE_CATALOG.map((entry) => [entry.id, entry]),
);
export function officePictureStyle(id: unknown): OfficePictureStyleEntry | null {
  return typeof id === "string"
    ? (BY_ID.get(id as OfficePictureStyleId) ?? null)
    : null;
}
export function isOfficePictureStyleId(id: unknown): id is OfficePictureStyleId {
  return officePictureStyle(id) !== null;
}
/** The style's name as a `console-notebooks` catalogue key (Office's own names). */
export function officePictureStyleLabelKey(id: OfficePictureStyleId): string {
  return `office.pictureStyle.${id}`;
}
