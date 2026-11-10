"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { useTranslations } from "@mrnote/core/host/next-intl.tsx";
import { useState, type ComponentProps } from "react";

import {
  OfficeHomeCommand,
  OfficeHomeCommandRow,
  OfficeHomeCommandStack,
  OfficeHomeMenuTrigger,
  OfficeHomeRibbonGroup,
} from "@mrnote/core/web/components/office-ribbon/OfficeHomeRibbon.tsx";
import {
  OfficeShapeFormatTab,
  type OfficeShapeEffect,
  type OfficeShapeFormatAdapter,
  type OfficeShapeFormatLabels,
  OfficeSmartArtEditIn2dCommand,
  type OfficeSmartArtEditIn2dProps,
} from "@mrnote/core/web/components/office-ribbon/OfficeShapeFormatTab.tsx";
import type { OfficeWordArtRecipe } from "@mrnote/core/web/lib/office-wordart-style-catalog.ts";
import type { OfficeColorPalettePanel } from "@mrnote/core/web/components/office-ribbon/OfficeColorPalettePanel.tsx";
import { OfficeSemanticIcon } from "@mrnote/core/web/components/office-ribbon/icons/OfficeSemanticIcon.tsx";
import { Maximize2, Minimize2 } from "@mrnote/core/web/components/office-ribbon/icons/houseIcons.tsx";
import {
  DropdownMenu,
  DropdownMenuContent,
  DropdownMenuItem,
  DropdownMenuLabel,
  DropdownMenuTrigger,
} from "@mrnote/core/web/components/ui/dropdown-menu.tsx";
import type { PresentationImageBlock } from "@mrnote/core/web/components/console/editor/native-document/documentCommands.ts";
import { canApplyNativePictureStyle } from "@mrnote/core/web/lib/presentation-picture-style-native.ts";
import { IconPictureCrop } from "@mrnote/core/web/components/notebook/contents/PresentationRibbonIcons.tsx";
import { OfficePictureMenuIcon } from "@mrnote/core/web/components/office-ribbon/icons/OfficePictureMenuIcon.tsx";
import { PresentationPictureStyleGallery } from "@mrnote/core/web/components/notebook/contents/ribbon/PresentationPictureStyleGallery.tsx";
import { officeTextPaintEffect } from "@mrnote/core/web/lib/document-wordart-style-adapter.ts";
import { OFFICE_SHAPE_EFFECT_PRESETS } from "@mrnote/core/web/lib/office-shape-effects.ts";
import type { OfficeShapeStyleRecipe } from "@mrnote/core/web/lib/office-shape-style-catalog.ts";
import { patchOfficeTextPaint, type OfficeTextPaint } from "@mrnote/core/web/lib/office-text-paint.ts";
import type { ShapeKind } from "@mrnote/core/web/lib/presentation-shapes.ts";

import { PresentationHomeShapeGallery } from "@mrnote/core/web/components/notebook/contents/ribbon/PresentationHomeShapeGallery.tsx";

/** What the targeted shapes and words of the diagram currently wear. */
export interface PresentationSmartArtFormatState {
  readonly shape: ShapeKind | null;
  readonly fill: string | null;
  readonly outline: string | null;
  readonly outlineWeightPt: number | null;
  readonly effect: OfficeShapeEffect | null;
  readonly styleId: string | null;
  readonly textColor: string | null;
  readonly textPaint: OfficeTextPaint | null;
  readonly width: string | null;
  readonly height: string | null;
}

/** The commands the host runs, each one undo step on the live selection. */
export interface PresentationSmartArtFormatCommands {
  changeShape: (shape: ShapeKind) => boolean | void;
  resize: (direction: "larger" | "smaller") => void;
  fill: (color: string | null) => void;
  outline: (color: string | null) => void;
  outlineWeight: (pt: number) => void;
  effect: (effect: OfficeShapeEffect) => void;
  style: (recipe: OfficeShapeStyleRecipe) => void;
  textPaint: (paint: OfficeTextPaint | null) => void;
  textFill: (color: string) => void;
  size: (axis: "width" | "height", value: string, lockAspect: boolean) => void;
  altText: () => void;
  selectionPane: () => void;
  /** Bring Forward / Send Backward on the WHOLE graphic, in the slide's stack. */
  zOrder?: (action: "bringToFront" | "sendToBack" | "bringForward" | "sendBackward") => void;
}

/**
 * SmartArt ▸ Format (PPT-A010): PowerPoint's second SmartArt tab, which
 * formats the SHAPES selected inside the graphic -- Shapes (Change Shape,
 * Larger, Smaller), Shape Styles (the deck's style gallery, Shape Fill,
 * Shape Outline, Shape Effects), WordArt Styles (the deck's gallery, Text
 * Fill, Text Outline, Text Effects), Accessibility, Arrange and Size.
 *
 * It is the SHARED Shape Format tab (`OfficeShapeFormatTab`) with the
 * SmartArt adapter: the same galleries, menus and fields an ordinary shape's
 * tab draws, writing through the host's SmartArt commands
 * (`lib/presentation-smartart-format`). What PowerPoint greys for a shape
 * inside a graphic (moving it in the z-order, aligning, grouping) is greyed
 * here with the reason; Rotate has no field on one diagram shape yet and says so.
 */
export function PresentationSmartArtFormatTab({
  labels,
  writable,
  hasTarget,
  disabledReason,
  state,
  styleRecipes,
  themeColumns,
  recents,
  canResize,
  resizeDisabledReason,
  editIn2d,
  sizeStep,
  sizeMax,
  wordArt,
  commands,
  graphicSelected = false,
  graphicOutline = null,
  picture = null,
}: {
  labels: OfficeShapeFormatLabels;
  /** The page and every targeted member may change. */
  writable: boolean;
  /** At least one shape of the diagram is targeted. */
  hasTarget: boolean;
  /** Why the commands are greyed, when they are. */
  disabledReason?: string;
  state: PresentationSmartArtFormatState;
  styleRecipes: readonly OfficeShapeStyleRecipe[];
  themeColumns?: ComponentProps<typeof OfficeColorPalettePanel>["themeColumns"];
  recents: readonly string[];
  canResize: boolean;
  /** Why Larger / Smaller and the size boxes are greyed with a shape targeted. */
  resizeDisabledReason?: string;
  /** Shapes ▸ Edit in 2-D, the shared toggle (`OfficeShapeFormatTab`). */
  editIn2d?: OfficeSmartArtEditIn2dProps;
  sizeStep: number;
  sizeMax: number;
  /** The deck's WordArt gallery (PowerPoint's twenty under its theme, E026). */
  wordArt: {
    readonly recipes: readonly OfficeWordArtRecipe[];
    readonly labelFor: (recipe: OfficeWordArtRecipe) => string;
    readonly value: string | null;
    readonly onCommit: (recipe: OfficeWordArtRecipe) => void;
  };
  commands: PresentationSmartArtFormatCommands;
  /**
   * The WHOLE graphic is selected: Fill, Outline and Effects format the
   * graphic itself, whose background holds the effect-list effects only
   * (`dgm:bg`: no Bevel or 3-D).
   */
  graphicSelected?: boolean;
  /** The whole graphic's own line colour (`dgm:whole`), when it has one. */
  graphicOutline?: string | null;
  /**
   * Picture Format on the picture in the targeted slot (PowerPoint's
   * Picture Format tab for a picture placeholder holding a picture): Crop >
   * Fill / Fit, the Picture Styles gallery and Reset Picture, through the
   * deck's own picture machinery. Absent when no filled slot is targeted.
   */
  picture?: {
    readonly image: PresentationImageBlock;
    readonly accentColors: readonly string[];
    readonly onCrop: (mode: "fill" | "fit") => void;
    readonly onStyle: (styleId: string) => void;
    readonly onReset: () => void;
  } | null;
}) {
  const t = useTranslations("console-notebooks");
  const [aspectLocked, setAspectLocked] = useState(false);
  const enabled = writable && hasTarget;
  const off = enabled ? { disabled: false } : { disabled: true, disabledReason };
  const arranged = {
    disabled: true,
    disabledReason: t("presentation.smartartContext.formatArrangedByLayout"),
  } as const;
  const unsupported = {
    disabled: true,
    disabledReason: t("presentation.smartartContext.formatUnavailable"),
  } as const;
  const paintBase: OfficeTextPaint = state.textPaint ?? {
    fill: state.textColor ?? "#ffffff",
    outline: "none",
    outlineWidthPt: 0,
  };
  const adapter: OfficeShapeFormatAdapter = {
    application: "presentation",
    labels,
    // The SmartArt tab draws its own Shapes group; these two stay unused.
    insert: { onInsert: () => false, recentStorageKey: "mrai.smartart.format.insert", ...arranged },
    textBox: { onInsert: () => undefined, ...arranged },
    changeShape: {
      onChange: () => false,
      recentStorageKey: "mrai.smartart.format.change",
      ...arranged,
    },
    styles: {
      quickCount: 7,
      quickRow: "light-gradient",
      recipes: styleRecipes,
      value: state.styleId,
      onCommit: commands.style,
      ...off,
    },
    fill: {
      value: state.fill ?? undefined,
      themeColumns,
      recents,
      onPick: commands.fill,
      ...off,
    },
    outline: {
      // The whole graphic: its own line, and Office's default black on the
      // button while it has none.
      value: graphicSelected ? (graphicOutline ?? "none") : (state.outline ?? undefined),
      ...(graphicSelected ? { lastColor: graphicOutline ?? "#000000" } : {}),
      themeColumns,
      recents,
      onPick: commands.outline,
      weightPt: state.outlineWeightPt,
      onWeight: commands.outlineWeight,
      lineStyle: null,
      onLineStyle: () => undefined,
      // A dash or arrow on one diagram shape has no field yet.
      lineStyleCapability: unsupported,
      lineEnds: { headEnd: "none", tailEnd: "none" },
      onLineEnds: () => undefined,
      lineEndsCapability: unsupported,
      ...off,
    },
    effects: {
      value: state.effect ?? "none",
      available: graphicSelected
        ? ["none", "shadow", "reflection", "glow", "softEdge"]
        : ["none", "shadow", ...OFFICE_SHAPE_EFFECT_PRESETS],
      onPick: commands.effect,
      cascade: {
        presets: t("office.shapeEffects.presets"),
        category: (effect) =>
          effect === "threeD"
            ? t("office.shapeEffects.threeDRotation")
            : t(`toolbar.shapeFormatOption.effect.${effect}` as "toolbar.shapeFormatOption.effect.none"),
        none: (effect) => t(`office.shapeEffects.none.${effect}` as "office.shapeEffects.none.shadow"),
      },
      ...off,
    },
    // WordArt Styles: the deck's gallery on the targeted shapes' words, as
    // PowerPoint's SmartArt Format tab offers it (E026).
    wordArt: {
      recipes: wordArt.recipes,
      labelFor: wordArt.labelFor,
      sample: "Aa",
      value: wordArt.value,
      onCommit: wordArt.onCommit,
      ...off,
    },
    textFill: {
      value: state.textPaint?.fill ?? state.textColor ?? undefined,
      themeColumns,
      recents,
      onPick: (color) => {
        if (!color) return;
        if (state.textPaint) commands.textPaint(patchOfficeTextPaint(paintBase, { fill: color }));
        else commands.textFill(color);
      },
      ...off,
    },
    textOutline: {
      value: state.textPaint && state.textPaint.outline !== "none" ? state.textPaint.outline : undefined,
      themeColumns,
      recents,
      onPick: (color) =>
        commands.textPaint(
          patchOfficeTextPaint(paintBase, { outline: color ?? "none", outlineWidthPt: color ? 1 : 0 }),
        ),
      ...off,
    },
    textEffects: {
      value: officeTextPaintEffect(state.textPaint),
      available: ["none", "shadow", "reflection", "glow", "bevel"],
      unavailableReason: t("presentation.ui.shapeFormat.effectNoWriter"),
      onPick: (effect) => commands.textPaint(patchOfficeTextPaint(paintBase, { effect })),
      ...off,
    },
    altText: { active: false, onOpen: commands.altText, disabled: !writable, disabledReason },
    // The whole graphic moves in the slide's stack like any object; one shape
    // of it does not, since its place is the layout's and the file has no
    // field for another (`dgm:prSet` states no z-order; `zOrderOff` is the
    // layout's own).
    zOrder:
      graphicSelected && commands.zOrder
        ? { onAction: commands.zOrder, ...(writable ? { disabled: false } : { disabled: true, disabledReason }) }
        : { onAction: () => undefined, ...arranged },
    selectionPane: { onSelect: commands.selectionPane },
    alignObjects: { onAction: () => undefined, ...arranged },
    group: { grouped: false, onGroup: () => undefined, onUngroup: () => undefined, ...arranged },
    rotate: { onAction: () => undefined, ...unsupported },
    size: {
      width: state.width ?? "",
      height: state.height ?? "",
      mixed: { width: state.width === null, height: state.height === null },
      step: sizeStep,
      max: sizeMax,
      heightFirst: true,
      onCommit: (axis, value) => commands.size(axis, value, aspectLocked),
      lockAspect: {
        value: aspectLocked,
        onToggle: () => setAspectLocked((locked) => !locked),
        ...(enabled && canResize ? {} : off),
      },
      ...(enabled && canResize
        ? { disabled: false }
        : { disabled: true, disabledReason: enabled ? resizeDisabledReason ?? disabledReason : disabledReason }),
    },
  };
  const resizeReason = enabled && !canResize ? resizeDisabledReason ?? off.disabledReason : off.disabledReason;
  const shapesGroup = (
    <OfficeHomeRibbonGroup
      label={t("presentation.smartartContext.shapesGroup")}
      testId="presentation-ribbon-smartart-format-shapes-group"
    >
      {/* Change Shape beside its stack (Edit in 2-D, Larger, Smaller), as
          Office lays the group out; the group itself stacks its children. */}
      <OfficeHomeCommandRow>
      <DropdownMenu modal={false}>
        <DropdownMenuTrigger asChild>
          <OfficeHomeMenuTrigger
            variant="large"
            hasMenuIndicator
            icon={<OfficeSemanticIcon semanticId="shape.change" size={24} />}
            label={t("presentation.ui.shapeFormat.changeShape")}
            disabled={!enabled}
            disabledReason={off.disabledReason}
            testId="presentation-ribbon-smartart-format-change-shape"
          />
        </DropdownMenuTrigger>
        <DropdownMenuContent align="start" className="docs-cupertino-menu">
          <DropdownMenuLabel>{t("presentation.ui.shapeFormat.changeShape")}</DropdownMenuLabel>
          <PresentationHomeShapeGallery
            mode="change"
            disabled={!enabled}
            selectedKind={state.shape}
            labelFor={(selection) => String(selection)}
            onSelect={(selection) => commands.changeShape(selection as ShapeKind)}
          />
        </DropdownMenuContent>
      </DropdownMenu>
      <OfficeHomeCommandStack>
        {editIn2d ? (
          <OfficeSmartArtEditIn2dCommand
            editIn2d={editIn2d}
            testId="presentation-ribbon-smartart-format-edit-in-2d"
          />
        ) : null}
        <OfficeHomeCommand
          variant="compact"
          icon={<Maximize2 size={16} />}
          label={t("presentation.smartartContext.larger")}
          disabled={!enabled || !canResize}
          disabledReason={resizeReason}
          onClick={() => commands.resize("larger")}
          testId="presentation-ribbon-smartart-format-larger"
        />
        <OfficeHomeCommand
          variant="compact"
          icon={<Minimize2 size={16} />}
          label={t("presentation.smartartContext.smaller")}
          disabled={!enabled || !canResize}
          disabledReason={resizeReason}
          onClick={() => commands.resize("smaller")}
          testId="presentation-ribbon-smartart-format-smaller"
        />
      </OfficeHomeCommandStack>
      </OfficeHomeCommandRow>
    </OfficeHomeRibbonGroup>
  );
  const pictureGroup = picture ? (
    <OfficeHomeRibbonGroup
      label={t("presentation.home.ribbonPictureFormat")}
      size="gallery"
      testId="presentation-ribbon-smartart-format-picture-group"
    >
      <OfficeHomeCommandRow>
        <DropdownMenu modal={false}>
          <DropdownMenuTrigger asChild>
            <OfficeHomeMenuTrigger
              variant="large"
              hasMenuIndicator
              icon={<IconPictureCrop />}
              label={t("presentation.crop.toggle")}
              disabled={!writable}
              disabledReason={disabledReason}
              testId="presentation-ribbon-smartart-format-picture-crop"
            />
          </DropdownMenuTrigger>
          <DropdownMenuContent align="start" className="docs-cupertino-menu">
            <DropdownMenuItem
              onSelect={() => picture.onCrop("fill")}
              data-testid="presentation-ribbon-smartart-format-picture-crop-fill"
            >
              {t("presentation.crop.fill")}
            </DropdownMenuItem>
            <DropdownMenuItem
              onSelect={() => picture.onCrop("fit")}
              data-testid="presentation-ribbon-smartart-format-picture-crop-fit"
            >
              {t("presentation.crop.fit")}
            </DropdownMenuItem>
          </DropdownMenuContent>
        </DropdownMenu>
        <PresentationPictureStyleGallery
          accentColors={picture.accentColors}
          disabled={!writable || !canApplyNativePictureStyle(picture.image)}
          disabledReason={disabledReason}
          image={picture.image}
          labels={{
            ariaLabel: t("presentation.image.stylesGroup"),
            more: t("toolbar.quickOverflow"),
          }}
          onCommit={(style) => picture.onStyle(style.id)}
          onPreviewChange={() => undefined}
        />
        <OfficeHomeCommandStack>
          <OfficeHomeCommand
            variant="compact"
            icon={<OfficePictureMenuIcon icon={{ kind: "picture-reset" }} />}
            label={t("presentation.pictureRibbon.resetPicture")}
            disabled={!writable || (!picture.image.cropRect && !picture.image.pictureStylePreset)}
            disabledReason={disabledReason}
            onClick={picture.onReset}
            testId="presentation-ribbon-smartart-format-picture-reset"
          />
        </OfficeHomeCommandStack>
      </OfficeHomeCommandRow>
    </OfficeHomeRibbonGroup>
  ) : null;
  return (
    <OfficeShapeFormatTab
      adapter={adapter}
      shapesGroup={pictureGroup ? <>{shapesGroup}{pictureGroup}</> : shapesGroup}
      testIdPrefix="presentation-ribbon-smartart-format"
      menuClassName="docs-cupertino-menu"
      ids={{
        fillPrimary: "presentation-ribbon-smartart-format-fill",
        fillTrigger: "presentation-ribbon-smartart-format-fill-menu",
        fill: "presentation-ribbon-smartart-format-fill-palette",
        outlinePrimary: "presentation-ribbon-smartart-format-outline",
        outlineTrigger: "presentation-ribbon-smartart-format-outline-menu",
        outline: "presentation-ribbon-smartart-format-outline-palette",
        effects: "presentation-ribbon-smartart-format-effects",
        effect: "presentation-ribbon-smartart-format-effect",
        textFill: "presentation-ribbon-smartart-format-text-fill",
        textOutline: "presentation-ribbon-smartart-format-text-outline",
        textEffects: "presentation-ribbon-smartart-format-text-effects",
        altText: "presentation-ribbon-smartart-format-alt-text",
      }}
    />
  );
}
