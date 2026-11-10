"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { ReactNode } from "react";

import {
  OfficeGalleryMenuContent,
  OfficeGalleryMenuGrid,
  OfficeGalleryMenuSection,
  OfficeGalleryMenuTileBody,
  officeGalleryMenuStyles,
} from "../../../office-ribbon/OfficeGalleryMenu";
import { OfficeHomeCommand } from "../../../office-ribbon/OfficeHomeRibbon";
import { OfficeHomeMenuTrigger } from "../../../office-ribbon/OfficeHomeRibbon";
import { OfficeRibbonGallery } from "../../../office-ribbon/OfficeRibbonGallery";
import { OfficeSemanticIcon } from "../../../office-ribbon/icons/OfficeSemanticIcon";
import {
  DropdownMenu,
  DropdownMenuItem,
  DropdownMenuTrigger,
} from "../../../ui/dropdown-menu";
import {
  PRESENTATION_ANIMATION_EFFECT_GROUPS,
  presentationAnimationEffect,
  presentationAnimationEffectsByCategory,
  type PresentationAnimationEffectId,
  type TimelineAnimation,
} from "../../../../lib/presentation-anim-timeline";

import { PresentationAnimationIcon } from "./PresentationAnimationGallery";

/**
 * One tile of PowerPoint's own animation galleries. `label` names a gallery
 * title PowerPoint gives an effect that differs from the effect's catalogue
 * name: since PowerPoint 2010 the gallery calls Circle "Shape" (its Effect
 * Options pick the shape) and the Ascend exit "Float Out".
 */
export interface PresentationAnimationRibbonTile {
  readonly id: PresentationAnimationEffectId;
  readonly label?: "shape" | "floatOut";
}

/**
 * The in-ribbon Entrance Effects gallery of PowerPoint for Mac 16.86, in its
 * order (accessibility tree of the Animations tab, a shape selected): Basic,
 * then Expand and Fade, then the Exciting row.
 */
export const PRESENTATION_ANIMATION_RIBBON_ENTRANCE = [
  { id: "appear" },
  { id: "blinds-in" },
  { id: "checkerboard-in" },
  { id: "dissolve-in" },
  { id: "fly-in" },
  { id: "peek-in" },
  { id: "random-bars-in" },
  { id: "circle-in", label: "shape" },
  { id: "split-in" },
  { id: "strips-in" },
  { id: "wedge-in" },
  { id: "wheel-in" },
  { id: "wipe-in" },
  { id: "expand" },
  { id: "fade" },
  { id: "bounce-in" },
  { id: "credits-in" },
  { id: "curve-up-in" },
  { id: "float-in" },
  { id: "pinwheel-in" },
  { id: "spiral-in" },
  { id: "basic-swivel-in" },
] as const satisfies readonly PresentationAnimationRibbonTile[];






export const PRESENTATION_ANIMATION_RIBBON_EMPHASIS = [
  { id: "fill-color" },
  { id: "grow-shrink" },
  { id: "line-color" },
  { id: "spin" },
  { id: "transparency" },
  { id: "complementary-color" },
  { id: "complementary-color-2" },
  { id: "contrasting-color" },
  { id: "darken" },
  { id: "desaturate" },
  { id: "lighten" },
  { id: "object-color" },
  { id: "flash-bulb" },
  { id: "flicker" },
  { id: "teeter" },
] as const satisfies readonly PresentationAnimationRibbonTile[];

/** PowerPoint for Mac 16.86's Exit Effects menu: four titled sections. */
export const PRESENTATION_ANIMATION_RIBBON_EXIT = [
  {
    group: "basic",
    items: [
      { id: "blinds-out" },
      { id: "checkerboard-out" },
      { id: "disappear" },
      { id: "dissolve-out" },
      { id: "fly-out" },
      { id: "peek-out" },
      { id: "random-bars-out" },
      { id: "circle-out", label: "shape" },
      { id: "split-out" },
      { id: "strips-out" },
      { id: "wedge-out" },
      { id: "wheel-out" },
      { id: "wipe-out" },
    ],
  },
  {
    group: "subtle",
    items: [
      { id: "contract-out" },
      { id: "fade-out" },
      { id: "swivel-out" },
      { id: "zoom-out" },
    ],
  },
  {
    group: "moderate",
    items: [
      { id: "center-revolve-out" },
      { id: "stretch-out" },
      { id: "float-down-out", label: "floatOut" },
      { id: "shrink-and-turn" },
      { id: "sink-down" },
      { id: "spinner-out" },
      { id: "basic-zoom-out" },
      { id: "stretchy-out" },
    ],
  },
  {
    group: "exciting",
    items: [
      { id: "boomerang-out" },
      { id: "bounce-out" },
      { id: "credits-out" },
      { id: "curve-down-out" },
      { id: "float-out" },
      { id: "pinwheel-out" },
      { id: "spiral-out" },
      { id: "basic-swivel-out" },
    ],
  },
] as const satisfies readonly {
  readonly group: "basic" | "subtle" | "moderate" | "exciting";
  readonly items: readonly PresentationAnimationRibbonTile[];
}[];

/** The motion-path presets by PowerPoint's own path groups. */
export function presentationAnimationRibbonMotionPaths(): readonly {
  readonly group: "basic" | "lines" | "special";
  readonly items: readonly PresentationAnimationRibbonTile[];
}[] {
  return (["basic", "lines", "special"] as const).map((group) => ({
    group,
    items: presentationAnimationEffectsByCategory("emphasis")
      .filter((entry) => entry.motionPath && entry.group === group)
      .map((entry) => ({ id: entry.id })),
  }));
}

/**
 * Every emphasis effect the editor offers, in PowerPoint's More Emphasis
 * Effects groups (Basic, Subtle, Moderate, Exciting): what the emphasis
 * strip's chevron opens, beyond the fifteen tiles the ribbon has room for.
 * Motion paths are emphasis-timed but have their own menu (Path Animation).
 */
export function presentationAnimationRibbonEmphasisSections(): readonly {
  readonly group: (typeof PRESENTATION_ANIMATION_EFFECT_GROUPS)[number];
  readonly items: readonly PresentationAnimationRibbonTile[];
}[] {
  const emphasis = presentationAnimationEffectsByCategory("emphasis").filter(
    (entry) => !entry.motionPath,
  );
  return PRESENTATION_ANIMATION_EFFECT_GROUPS.map((group) => ({
    group,
    items: emphasis.filter((entry) => entry.group === group).map((entry) => ({ id: entry.id })),
  }));
}

type Translate = (key: string) => string;

function tileLabel(tile: PresentationAnimationRibbonTile, t: Translate): string {
  return tile.label
    ? t(`presentation.animations.galleryName.${tile.label}`)
    : t(`presentation.animation.${tile.id}`);
}

/**
 * An in-ribbon effect gallery (Entrance or Emphasis): large tiles that apply
 * the effect to the selected object, and a chevron that opens `more`.
 */
export function PresentationAnimationRibbonStrip({
  ariaLabel,
  className,
  testId,
  leading,
  tiles,
  canSelect,
  onSelect,
  selected,
  more,
  t,
}: {
  ariaLabel: string;
  className?: string;
  testId: string;
  leading?: ReactNode;
  tiles: readonly PresentationAnimationRibbonTile[];
  canSelect: (effect: TimelineAnimation) => boolean;
  onSelect: (effect: TimelineAnimation) => void;
  selected?: string;
  more?: ReactNode;
  t: Translate;
}) {
  return (
    <OfficeRibbonGallery
      ariaLabel={ariaLabel}
      className={className}
      testId={testId}
      viewport={
        <div>
          {leading}
          {tiles.map((tile) => {
            const entry = presentationAnimationEffect(tile.id);
            if (!entry) return null;
            return (
              <OfficeHomeCommand
                key={tile.id}
                variant="large"
                active={selected === tile.id}
                controlId={entry.controlId}
                icon={<PresentationAnimationIcon effect={tile.id} size={24} />}
                label={tileLabel(tile, t)}
                onClick={() => onSelect(tile.id)}
                disabled={!canSelect(tile.id)}
                testId={`presentation-animation-quick-${tile.id}`}
              />
            );
          })}
        </div>
      }
      more={more}
    />
  );
}

/**
 * A large menu button (Exit Effects, Path Animation) whose menu lists the
 * effects in PowerPoint's titled sections.
 */
export function PresentationAnimationRibbonMenu({
  label,
  icon,
  testId,
  controlId,
  sections,
  sectionLabel,
  canSelect,
  onSelect,
  selected,
  t,
}: {
  label: string;
  icon: ReactNode;
  testId: string;
  controlId: string;
  sections: readonly {
    readonly group: string;
    readonly items: readonly PresentationAnimationRibbonTile[];
  }[];
  sectionLabel: (group: string) => string;
  canSelect: (effect: TimelineAnimation) => boolean;
  onSelect: (effect: TimelineAnimation) => void;
  selected?: string;
  t: Translate;
}) {
  const anyEnabled = sections.some((section) =>
    section.items.some((tile) => canSelect(tile.id)),
  );
  return (
    <DropdownMenu>
      <DropdownMenuTrigger asChild>
        <OfficeHomeMenuTrigger
          data-control-id={controlId}
          disabled={!anyEnabled}
          displayLabel={label}
          hasMenuIndicator
          icon={icon}
          label={label}
          testId={testId}
          variant="large"
        />
      </DropdownMenuTrigger>
      <PresentationAnimationSectionsMenuContent
        label={label}
        testId={testId}
        sections={sections}
        sectionLabel={sectionLabel}
        canSelect={canSelect}
        onSelect={onSelect}
        selected={selected}
        t={t}
      />
    </DropdownMenu>
  );
}

/**
 * The titled sections of effects a ribbon menu opens (Exit Effects, Path
 * Animation, the emphasis strip's chevron): one card per effect, the
 * selected object's own effect marked.
 */
export function PresentationAnimationSectionsMenuContent({
  label,
  testId,
  sections,
  sectionLabel,
  canSelect,
  onSelect,
  selected,
  t,
}: {
  label: string;
  testId: string;
  sections: readonly {
    readonly group: string;
    readonly items: readonly PresentationAnimationRibbonTile[];
  }[];
  sectionLabel: (group: string) => string;
  canSelect: (effect: TimelineAnimation) => boolean;
  onSelect: (effect: TimelineAnimation) => void;
  selected?: string;
  t: Translate;
}) {
  return (
    <OfficeGalleryMenuContent aria-label={label} data-testid={`${testId}-menu`}>
      {sections.map((section) =>
        section.items.length === 0 ? null : (
          <div
            key={section.group}
            role="group"
            aria-label={sectionLabel(section.group)}
            data-animation-group={section.group}
          >
            <OfficeGalleryMenuSection>{sectionLabel(section.group)}</OfficeGalleryMenuSection>
            {/* PowerPoint's effect sheets: seven icon tiles a row. */}
            <OfficeGalleryMenuGrid columns={7} tileWidth={72}>
              {section.items.map((tile) => {
                const enabled = canSelect(tile.id);
                const name = tileLabel(tile, t);
                return (
                  <DropdownMenuItem
                    asChild
                    key={tile.id}
                    disabled={!enabled}
                    onSelect={() => {
                      if (canSelect(tile.id)) onSelect(tile.id);
                    }}
                  >
                    <button
                      type="button"
                      disabled={!enabled}
                      className={officeGalleryMenuStyles.tile}
                      data-selected={selected === tile.id ? "true" : "false"}
                      aria-pressed={selected === tile.id}
                      data-testid={`${testId}-${tile.id}`}
                    >
                      <OfficeGalleryMenuTileBody
                        thumbnail={<PresentationAnimationIcon effect={tile.id} />}
                        name={name}
                      />
                    </button>
                  </DropdownMenuItem>
                );
              })}
            </OfficeGalleryMenuGrid>
          </div>
        ),
      )}
    </OfficeGalleryMenuContent>
  );
}

export function PresentationAnimationMoreChevron({
  label,
  disabled,
  testId,
}: {
  label: string;
  disabled: boolean;
  testId: string;
}) {
  return (
    <DropdownMenuTrigger asChild>
      <OfficeHomeMenuTrigger
        geometry="gallery"
        variant="icon"
        icon={<OfficeSemanticIcon semanticId="menu.chevron.down" size={14} />}
        label={label}
        disabled={disabled}
        testId={testId}
      />
    </DropdownMenuTrigger>
  );
}
