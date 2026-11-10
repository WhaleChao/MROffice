"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import {
  useMemo,
  useRef,
  useState,
  type KeyboardEvent,
} from "react";

import { PresentationTransitionThumbnail } from "./PresentationTransitionThumbnails";
import {
  OfficeGalleryMenuContent,
  OfficeGalleryMenuGrid,
  OfficeGalleryMenuSection,
  OfficeGalleryMenuTileBody,
  officeGalleryMenuStyles,
} from "../../../office-ribbon/OfficeGalleryMenu";
import { OfficeHomeMenuTrigger } from "../../../office-ribbon/OfficeHomeRibbon";
import homeStyles from "../../../office-ribbon/OfficeHomeRibbon.module.css";
import { OfficeRibbonButton } from "../../../office-ribbon/OfficeRibbonChassis";
import { OfficeRibbonGallery } from "../../../office-ribbon/OfficeRibbonGallery";
import { OfficeSemanticIcon } from "../../../office-ribbon/icons/OfficeSemanticIcon";
import {
  DropdownMenu,
  DropdownMenuItem,
  DropdownMenuTrigger,
} from "../../../ui/dropdown-menu";
import {
  normalizeSlideTransition,
  type AuthorableSlideTransition,
  type SlideTransition,
} from "../../../../lib/presentation-transitions";

import styles from "./PresentationTransitionGallery.module.css";

export type PresentationTransitionGalleryLabels = Readonly<
  Record<AuthorableSlideTransition, string>
>;

interface PresentationTransitionGalleryCatalogEntry {
  readonly id: string;
  readonly defaultLabel: string;
  readonly kind: AuthorableSlideTransition | null;
  readonly currentValues: readonly SlideTransition[];
}

/** Every transition the gallery offers, with its stable control identity. */
export const PRESENTATION_TRANSITION_GALLERY_CATALOG = [
  {
    id: "NONE",
    defaultLabel: "None",
    kind: "none",
    currentValues: ["none"],
  },
  {
    id: "MORPH",
    defaultLabel: "Morph",
    kind: "morph",
    currentValues: ["morph"],
  },
  {
    id: "FADE",
    defaultLabel: "Fade",
    kind: "fade",
    currentValues: ["fade"],
  },
  {
    id: "PUSH",
    defaultLabel: "Push",
    kind: "push",
    currentValues: [
      "push",
      "push-left",
      "push-right",
      "push-up",
      "push-down",
      "slide-left",
      "slide-right",
      "slide-up",
      "slide-down",
    ],
  },
  {
    id: "WIPE",
    defaultLabel: "Wipe",
    kind: "wipe",
    currentValues: [
      "wipe",
      "wipe-left",
      "wipe-right",
      "wipe-up",
      "wipe-down",
    ],
  },
  {
    id: "SPLIT",
    defaultLabel: "Split",
    kind: "split",
    currentValues: ["split"],
  },
  {
    id: "REVEAL",
    defaultLabel: "Reveal",
    kind: "reveal",
    currentValues: ["reveal"],
  },
  {
    id: "CUT",
    defaultLabel: "Cut",
    kind: "cut",
    currentValues: ["cut"],
  },
  {
    id: "RANDOM_BARS",
    defaultLabel: "Random Bars",
    kind: "random-bars",
    currentValues: ["random-bars"],
  },
  {
    id: "SHAPE",
    defaultLabel: "Shape",
    kind: "shape",
    currentValues: ["shape"],
  },
  {
    id: "UNCOVER",
    defaultLabel: "Uncover",
    kind: "uncover",
    currentValues: ["uncover"],
  },
  {
    id: "COVER",
    defaultLabel: "Cover",
    kind: "cover",
    currentValues: ["cover"],
  },
  {
    id: "FLASH",
    defaultLabel: "Flash",
    kind: "flash",
    currentValues: ["flash"],
  },
  {
    id: "FALL_OVER",
    defaultLabel: "Fall Over",
    kind: "fall-over",
    currentValues: ["fall-over"],
  },
  {
    id: "DRAPE",
    defaultLabel: "Drape",
    kind: "drape",
    currentValues: ["drape"],
  },
  {
    id: "CURTAINS",
    defaultLabel: "Curtains",
    kind: "curtains",
    currentValues: ["curtains"],
  },
  {
    id: "WIND",
    defaultLabel: "Wind",
    kind: "wind",
    currentValues: ["wind"],
  },
  {
    id: "PRESTIGE",
    defaultLabel: "Prestige",
    kind: "prestige",
    currentValues: ["prestige"],
  },
  {
    id: "FRACTURE",
    defaultLabel: "Fracture",
    kind: "fracture",
    currentValues: ["fracture"],
  },
  {
    id: "CRUSH",
    defaultLabel: "Crush",
    kind: "crush",
    currentValues: ["crush"],
  },
  {
    id: "PEEL_OFF",
    defaultLabel: "Peel Off",
    kind: "peel-off",
    currentValues: ["peel-off"],
  },
  {
    id: "PAGE_CURL",
    defaultLabel: "Page Curl",
    kind: "page-curl",
    currentValues: ["page-curl"],
  },
  {
    id: "AIRPLANE",
    defaultLabel: "Airplane",
    kind: "airplane",
    currentValues: ["airplane"],
  },
  {
    id: "ORIGAMI",
    defaultLabel: "Origami",
    kind: "origami",
    currentValues: ["origami"],
  },
  {
    id: "DISSOLVE",
    defaultLabel: "Dissolve",
    kind: "dissolve",
    currentValues: ["dissolve"],
  },
  {
    id: "CHECKERBOARD",
    defaultLabel: "Checkerboard",
    kind: "checkerboard",
    currentValues: ["checkerboard"],
  },
  {
    id: "BLINDS",
    defaultLabel: "Blinds",
    kind: "blinds",
    currentValues: ["blinds"],
  },
  {
    id: "CLOCK",
    defaultLabel: "Clock",
    kind: "clock",
    currentValues: ["clock"],
  },
  {
    id: "RIPPLE",
    defaultLabel: "Ripple",
    kind: "ripple",
    currentValues: ["ripple"],
  },
  {
    id: "HONEYCOMB",
    defaultLabel: "Honeycomb",
    kind: "honeycomb",
    currentValues: ["honeycomb"],
  },
  {
    id: "GLITTER",
    defaultLabel: "Glitter",
    kind: "glitter",
    currentValues: ["glitter"],
  },
  {
    id: "VORTEX",
    defaultLabel: "Vortex",
    kind: "vortex",
    currentValues: ["vortex"],
  },
  {
    id: "SHRED",
    defaultLabel: "Shred",
    kind: "shred",
    currentValues: ["shred"],
  },
  {
    id: "SWITCH",
    defaultLabel: "Switch",
    kind: "switch",
    currentValues: ["switch"],
  },
  {
    id: "FLIP",
    defaultLabel: "Flip",
    kind: "flip",
    currentValues: ["flip"],
  },
  {
    id: "GALLERY",
    defaultLabel: "Gallery",
    kind: "gallery",
    currentValues: ["gallery"],
  },
  { id: "CUBE", defaultLabel: "Cube", kind: "cube", currentValues: ["cube"] },
  {
    id: "DOORS",
    defaultLabel: "Doors",
    kind: "doors",
    currentValues: ["doors"],
  },
  { id: "BOX", defaultLabel: "Box", kind: "box", currentValues: ["box"] },
  {
    id: "COMB",
    defaultLabel: "Comb",
    kind: "comb",
    currentValues: ["comb"],
  },
  {
    id: "ZOOM",
    defaultLabel: "Zoom",
    kind: "zoom",
    currentValues: ["zoom"],
  },
  {
    id: "RANDOM",
    defaultLabel: "Random",
    kind: "random",
    currentValues: ["random"],
  },
  {
    id: "PAN",
    defaultLabel: "Pan",
    kind: "pan",
    currentValues: ["pan"],
  },
  {
    id: "FERRIS_WHEEL",
    defaultLabel: "Ferris Wheel",
    kind: "ferris-wheel",
    currentValues: ["ferris-wheel"],
  },
  {
    id: "CONVEYOR",
    defaultLabel: "Conveyor",
    kind: "conveyor",
    currentValues: ["conveyor"],
  },
  {
    id: "ROTATE",
    defaultLabel: "Rotate",
    kind: "rotate",
    currentValues: ["rotate"],
  },
  {
    id: "WINDOW",
    defaultLabel: "Window",
    kind: "window",
    currentValues: ["window"],
  },
  {
    id: "ORBIT",
    defaultLabel: "Orbit",
    kind: "orbit",
    currentValues: ["orbit"],
  },
  {
    id: "FLY_THROUGH",
    defaultLabel: "Fly Through",
    kind: "fly-through",
    currentValues: ["fly-through"],
  },
] as const satisfies readonly PresentationTransitionGalleryCatalogEntry[];

export type PresentationTransitionGalleryCatalogId =
  (typeof PRESENTATION_TRANSITION_GALLERY_CATALOG)[number]["id"];

/**
 * PowerPoint desktop's Transition to This Slide gallery: Subtle, Exciting and
 * Dynamic Content, in the order PowerPoint 16.86 lists them (accessibility
 * tree of the Transitions tab, `GalTransitionScheme_Control_<group>_<n>`).
 * All 49 effects; the in-ribbon row and the expanded catalogue share it.
 */
export const PRESENTATION_TRANSITION_GROUPS = [
  {
    id: "subtle",
    items: [
      "NONE",
      "MORPH",
      "FADE",
      "PUSH",
      "WIPE",
      "SPLIT",
      "REVEAL",
      "CUT",
      "RANDOM_BARS",
      "SHAPE",
      "UNCOVER",
      "COVER",
      "FLASH",
    ],
  },
  {
    id: "exciting",
    items: [
      "FALL_OVER",
      "DRAPE",
      "CURTAINS",
      "WIND",
      "PRESTIGE",
      "FRACTURE",
      "CRUSH",
      "PEEL_OFF",
      "PAGE_CURL",
      "AIRPLANE",
      "ORIGAMI",
      "DISSOLVE",
      "CHECKERBOARD",
      "BLINDS",
      "CLOCK",
      "RIPPLE",
      "HONEYCOMB",
      "GLITTER",
      "VORTEX",
      "SHRED",
      "SWITCH",
      "FLIP",
      "GALLERY",
      "CUBE",
      "DOORS",
      "BOX",
      "COMB",
      "ZOOM",
      "RANDOM",
    ],
  },
  {
    id: "dynamic",
    items: [
      "PAN",
      "FERRIS_WHEEL",
      "CONVEYOR",
      "ROTATE",
      "WINDOW",
      "ORBIT",
      "FLY_THROUGH",
    ],
  },
] as const satisfies readonly {
  readonly id: "subtle" | "exciting" | "dynamic";
  readonly items: readonly PresentationTransitionGalleryCatalogId[];
}[];

// The in-ribbon row scrolls through the same 49 effects in the same order.
const INLINE_TRANSITIONS: readonly AuthorableSlideTransition[] =
  PRESENTATION_TRANSITION_GROUPS.flatMap((group) =>
    group.items.flatMap((id) => {
      const kind = PRESENTATION_TRANSITION_GALLERY_CATALOG.find(
        (item) => item.id === id,
      )?.kind;
      return kind ? [kind] : [];
    }),
  );

export interface PresentationTransitionGalleryProps {
  readonly value?: unknown;
  readonly labels: PresentationTransitionGalleryLabels;
  readonly catalogLabels?: Readonly<
    Partial<Record<PresentationTransitionGalleryCatalogId, string>>
  >;
  readonly unavailableLabel: string;
  readonly disabled?: boolean;
  readonly isItemDisabled?: (kind: AuthorableSlideTransition) => boolean;
  readonly onSelect: (kind: AuthorableSlideTransition) => void;
  readonly ariaLabel?: string;
  readonly groupLabels?: Readonly<
    Record<"subtle" | "exciting" | "dynamic", string>
  >;
}

const CONTROL_SUFFIX = Object.freeze({
  cube: "CUBE",
  box: "BOX",
  rotate: "ROTATE",
  orbit: "ORBIT",
  none: "NONE",
  morph: "MORPH",
  fade: "FADE",
  push: "PUSH",
  wipe: "WIPE",
  dissolve: "DISSOLVE",
  cut: "CUT",
  shape: "SHAPE",
  split: "SPLIT",
  zoom: "ZOOM",
  "random-bars": "RANDOM_BARS",
  checkerboard: "CHECKERBOARD",
  blinds: "BLINDS",
  clock: "CLOCK",
  comb: "COMB",
  "fall-over": "FALL_OVER",
  drape: "DRAPE",
  curtains: "CURTAINS",
  wind: "WIND",
  prestige: "PRESTIGE",
  fracture: "FRACTURE",
  crush: "CRUSH",
  "peel-off": "PEEL_OFF",
  "page-curl": "PAGE_CURL",
  airplane: "AIRPLANE",
  origami: "ORIGAMI",
  switch: "SWITCH",
  flip: "FLIP",
  gallery: "GALLERY",
  "ferris-wheel": "FERRIS_WHEEL",
  conveyor: "CONVEYOR",
  uncover: "UNCOVER",
  cover: "COVER",
  doors: "DOORS",
  window: "WINDOW",
  vortex: "VORTEX",
  pan: "PAN",
  flash: "FLASH",
  honeycomb: "HONEYCOMB",
  reveal: "REVEAL",
  ripple: "RIPPLE",
  glitter: "GLITTER",
  shred: "SHRED",
  "fly-through": "FLY_THROUGH",
  random: "RANDOM",
} satisfies Record<AuthorableSlideTransition, string>);

/**
 * Presentation-owned first-row transition gallery.
 *
 * The component owns only visible gallery interaction. The host supplies the
 * live command gate and typed dispatch so a disabled tile can never manufacture
 * authoring custody. Native buttons preserve Enter/Space; the roving Arrow,
 * Home and End keys match Office gallery navigation without trapping Tab.
 */
export function PresentationTransitionGallery({
  value,
  labels,
  catalogLabels,
  unavailableLabel,
  disabled = false,
  isItemDisabled,
  onSelect,
  ariaLabel = "Transition to This Slide",
  groupLabels = {
    subtle: "Subtle",
    exciting: "Exciting",
    dynamic: "Dynamic Content",
  },
}: PresentationTransitionGalleryProps) {
  const normalized = normalizeSlideTransition(value);
  const selectedCatalogItem = PRESENTATION_TRANSITION_GALLERY_CATALOG.find(
    (item) =>
      (item.currentValues as readonly SlideTransition[]).includes(normalized),
  );
  const selectedCatalogId = selectedCatalogItem?.id;
  const selected = selectedCatalogItem?.kind ?? null;
  const enabledQuickKinds = useMemo<readonly AuthorableSlideTransition[]>(
    () =>
      INLINE_TRANSITIONS.filter(
        (kind) => !disabled && !isItemDisabled?.(kind),
      ),
    [disabled, isItemDisabled],
  );
  const [rovingKind, setRovingKind] = useState<AuthorableSlideTransition>(
    selected && enabledQuickKinds.includes(selected)
      ? selected
      : (enabledQuickKinds[0] ?? "none"),
  );
  const buttonRefs = useRef<
    Partial<Record<AuthorableSlideTransition, HTMLButtonElement | null>>
  >({});
  const effectiveRovingKind = enabledQuickKinds.includes(rovingKind)
    ? rovingKind
    : selected && enabledQuickKinds.includes(selected)
      ? selected
      : enabledQuickKinds[0];

  const moveFocus = (
    current: AuthorableSlideTransition,
    event: KeyboardEvent<HTMLButtonElement>,
  ) => {
    if (enabledQuickKinds.length === 0) return;
    let nextIndex: number | null = null;
    const currentIndex = Math.max(0, enabledQuickKinds.indexOf(current));
    if (event.key === "ArrowRight" || event.key === "ArrowDown") {
      nextIndex = (currentIndex + 1) % enabledQuickKinds.length;
    } else if (event.key === "ArrowLeft" || event.key === "ArrowUp") {
      nextIndex =
        (currentIndex - 1 + enabledQuickKinds.length) %
        enabledQuickKinds.length;
    } else if (event.key === "Home") {
      nextIndex = 0;
    } else if (event.key === "End") {
      nextIndex = enabledQuickKinds.length - 1;
    }
    if (nextIndex === null) return;
    event.preventDefault();
    const next = enabledQuickKinds[nextIndex];
    setRovingKind(next);
    buttonRefs.current[next]?.focus();
  };

  return (
    <OfficeRibbonGallery
      className={styles.gallery}
      ariaLabel={ariaLabel}
      role="toolbar"
      aria-orientation="horizontal"
      testId="presentation-transition-gallery"
      viewport={<div>
      {INLINE_TRANSITIONS.map((kind) => {
        const itemDisabled = disabled || Boolean(isItemDisabled?.(kind));
        const active = selectedCatalogId === CONTROL_SUFFIX[kind];
        const label = labels[kind];
        // The same large command face the Animations strip uses (icon over
        // label, no tile or preview frame), so the two galleries read as one
        // control; the roving Arrow/Home/End keys ride on the button props.
        return (
          <OfficeRibbonButton
            key={kind}
            big
            active={active}
            buttonRef={(node) => {
              buttonRefs.current[kind] = node;
            }}
            className={`${homeStyles.command} ${homeStyles.commandLarge}`}
            disabled={itemDisabled}
            label={label}
            testId={`presentation-transition-gallery-${kind}`}
            tabIndex={!itemDisabled && effectiveRovingKind === kind ? 0 : -1}
            data-kind={kind}
            data-selected={active ? "true" : "false"}
            data-control-id={`PPTCTL.RIBBON.TRANSITIONS.GALLERY.${CONTROL_SUFFIX[kind]}`}
            onFocus={() => setRovingKind(kind)}
            onKeyDown={(event) => moveFocus(kind, event)}
            onClick={() => {
              if (!itemDisabled) {
                setRovingKind(kind);
                onSelect(kind);
              }
            }}
          >
            {/* The effect's picture, as PowerPoint's gallery shows it: a
                small slide doing the transition, filling the tile. */}
            <span className={`${homeStyles.commandGlyph} ${styles.thumb}`} aria-hidden>
              <PresentationTransitionThumbnail kind={kind} width={48} />
            </span>
            <span className={`${homeStyles.commandLabel} ${styles.tileLabel}`}>{label}</span>
          </OfficeRibbonButton>
        );
      })}
      </div>}
      more={<DropdownMenu>
        <DropdownMenuTrigger asChild>
          <OfficeHomeMenuTrigger
            geometry="gallery"
            icon={
              <OfficeSemanticIcon semanticId="menu.chevron.down" size={14} />
            }
            label={ariaLabel}
            testId="presentation-ribbon-transition-gallery-open"
            variant="icon"
          />
        </DropdownMenuTrigger>
        <OfficeGalleryMenuContent data-testid="presentation-transition-gallery-menu">
          <div data-testid="presentation-transition-gallery-catalog">
            {PRESENTATION_TRANSITION_GROUPS.map((group) => (
              <section
                key={group.id}
                data-testid={`presentation-transition-category-${group.id}`}
              >
                <OfficeGalleryMenuSection>{groupLabels[group.id]}</OfficeGalleryMenuSection>
                {/* PowerPoint's Transitions sheet: eight pictures a row, each
                    as large as the ribbon strip draws it. */}
                <OfficeGalleryMenuGrid columns={8} tileWidth={76}>
                  {group.items
                    .map((id) =>
                      PRESENTATION_TRANSITION_GALLERY_CATALOG.find(
                        (item) => item.id === id,
                      )!,
                    )
                    .map((item) => {
                      const kind = item.kind;
                      const writerEnabled =
                        kind !== null &&
                        !disabled &&
                        !Boolean(isItemDisabled?.(kind));
                      const active = selectedCatalogId === item.id;
                      const label = catalogLabels?.[item.id] ?? labels[kind];
                      const accessibleLabel = writerEnabled
                        ? label
                        : `${label}: ${unavailableLabel}`;
                      return (
                        <DropdownMenuItem
                          asChild
                          disabled={!writerEnabled}
                          key={item.id}
                          onSelect={
                            kind === null
                              ? undefined
                              : () => {
                                  setRovingKind(kind);
                                  onSelect(kind);
                                }
                          }
                        >
                          <button
                            type="button"
                            className={officeGalleryMenuStyles.tile}
                            disabled={!writerEnabled}
                            aria-label={accessibleLabel}
                            aria-current={active ? "true" : undefined}
                            data-control-id={`PPTCTL.RIBBON.TRANSITIONS.GALLERY.${item.id}`}
                            data-disposition={
                              writerEnabled
                                ? "writer"
                                : "native-writer-unavailable"
                            }
                            data-selected={active ? "true" : "false"}
                            data-testid={`presentation-transition-gallery-catalog-${item.id.toLowerCase().replaceAll("_", "-")}`}
                            title={accessibleLabel}
                          >
                            <OfficeGalleryMenuTileBody
                              thumbnail={<PresentationTransitionThumbnail kind={kind ?? "unavailable"} width={48} />}
                              name={label}
                            />
                          </button>
                        </DropdownMenuItem>
                      );
                    })}
                </OfficeGalleryMenuGrid>
              </section>
            ))}
          </div>
        </OfficeGalleryMenuContent>
      </DropdownMenu>}
    />
  );
}
