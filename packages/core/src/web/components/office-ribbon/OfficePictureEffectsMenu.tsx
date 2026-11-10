"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Picture Format > Picture Effects: PowerPoint's cascade of Preset, Shadow,
 * Reflection, Glow, Soft Edges, Bevel and 3-D Rotation, each a submenu with
 * "No <effect>" over its gallery of variations (PPT-A127); Shadow's are the
 * shared Outer / Inner / Perspective galleries (`OfficeShadowGallery`). The variations and
 * their parameters are `lib/office-picture-effects`; a host gives the current
 * effects and gets back one pick at a time. Editor-neutral: the deck wires it
 * on its picture tab, and Docs or the worksheet can wire the same menu.
 */

import type { CSSProperties, ReactNode } from "react";

import {
  DropdownMenu,
  DropdownMenuContent,
  DropdownMenuItem,
  DropdownMenuLabel,
  DropdownMenuSeparator,
  DropdownMenuSub,
  DropdownMenuSubContent,
  DropdownMenuSubTrigger,
  DropdownMenuTrigger,
} from "../ui/dropdown-menu";
import { OFFICE_PICTURE_EFFECT_PRESETS } from "../../lib/office-picture-effect-presets";
import {
  OFFICE_BEVEL_PRESETS,
  OFFICE_GLOW_ACCENTS,
  OFFICE_GLOW_SIZES_PT,
  OFFICE_REFLECTION_VARIATIONS,
  OFFICE_ROTATION_GALLERY,
  OFFICE_SOFT_EDGE_SIZES_PT,
  normalizeOfficePictureEffects,
  officeBevelVariation,
  officeGlowVariation,
  officePictureEffectsCss,
  officeSoftEdgeVariation,
  sameOfficePictureEffect,
  type OfficeBevelPreset,
  type OfficePictureEffectKind,
  type OfficePictureEffectValues,
} from "../../lib/office-picture-effects";

import { OfficeHomeMenuTrigger } from "./OfficeHomeRibbon";
import styles from "./OfficePictureEffectsMenu.module.css";
import { OfficeShadowGallery, type OfficeShadowGalleryLabels } from "./OfficeShadowGallery";
import { OfficeSemanticIcon } from "./icons/OfficeSemanticIcon";
import type { OfficeSemanticIconId } from "./icons/officeIconRegistry";

type Accent = (typeof OFFICE_GLOW_ACCENTS)[number];

export interface OfficePictureEffectsMenuLabels {
  readonly button: string;
  readonly presets: string;
  /** "Preset N" (PowerPoint's names). */
  readonly preset: (preset: number) => string;
  readonly noPresets: string;
  readonly threeDOptions: string;
  /** Why 3-D Options... is not offered (no 3-D Format pane). */
  readonly threeDOptionsUnavailable: string;
  /** 3-D Rotation > 3-D Rotation Options...; absent, the 3-D Rotation
   * submenu lists no options row. */
  readonly threeDRotationOptions?: string;
  readonly category: (kind: OfficePictureEffectKind) => string;
  readonly none: (kind: OfficePictureEffectKind) => string;
  readonly variations: (kind: "reflection" | "glow" | "softEdge") => string;
  readonly shadowGroup: OfficeShadowGalleryLabels["group"];
  readonly shadowVariation: OfficeShadowGalleryLabels["variation"];
  readonly reflection: (id: string) => string;
  readonly glow: (sizePt: number, accent: number) => string;
  readonly softEdge: (sizePt: number) => string;
  readonly bevel: (preset: OfficeBevelPreset) => string;
  readonly rotationSection: (section: "parallel" | "perspective" | "oblique") => string;
  readonly camera: (camera: string) => string;
}

export interface OfficePictureEffectsMenuProps {
  readonly effects: unknown;
  /** The picture's older shadow switch (a shadow none of the galleries names). */
  readonly shadow: boolean;
  /** The deck theme's Accent 1-6, for the Glow gallery. */
  readonly accentColors: Readonly<Record<Accent, string>>;
  /** The effects the picture's Picture Style draws (its "No <effect>" states "none"). */
  readonly styleEffects?: ReadonlySet<OfficePictureEffectKind>;
  readonly labels: OfficePictureEffectsMenuLabels;
  readonly disabled?: boolean;
  readonly disabledReason?: string;
  readonly controlId?: string;
  readonly testId: string;
  /** The trigger's visible word when it is shorter than its name (the
   * Shape Format tab's "Effects"); absent, the name. */
  readonly displayLabel?: string;
  /** The categories listed, in the menu's order; absent, all seven (Text
   * Effects lists Shadow, Reflection, Glow, Bevel and 3-D Rotation). */
  readonly kinds?: readonly OfficePictureEffectKind[];
  /** Host rows after the categories (Text Effects > Transform). */
  readonly extra?: ReactNode;
  /** The trigger's glyph; absent, Shape Effects'. */
  readonly iconSemanticId?: OfficeSemanticIconId;
  readonly onPick: <K extends OfficePictureEffectKind>(kind: K, value: OfficePictureEffectValues[K] | null) => void;
  /** 3-D Options... (Preset, Bevel) and 3-D Rotation Options...: the host's
   * Format pane at its Effects page, as Office opens it ("format" for 3-D
   * Format, "rotation" for 3-D Rotation). Absent, the rows are disabled
   * with `labels.threeDOptionsUnavailable`. */
  readonly onThreeDOptions?: (section: "format" | "rotation") => void;
}

const PREVIEW = { width: 34, height: 26 };
/** Gallery thumbnails draw each effect at this fraction of its size. */
const PREVIEW_SCALE = 0.4;

function scaled<K extends OfficePictureEffectKind>(kind: K, value: OfficePictureEffectValues[K]) {
  const scale = (n: number) => Math.round(n * PREVIEW_SCALE);
  if (kind === "glow" || kind === "softEdge") {
    const round = value as OfficePictureEffectValues["glow"] | OfficePictureEffectValues["softEdge"];
    return { ...round, radius: Math.max(1, scale(round.radius)) };
  }
  if (kind === "reflection") {
    const reflection = value as OfficePictureEffectValues["reflection"];
    return { ...reflection, dist: scale(reflection.dist), blurRad: scale(reflection.blurRad) };
  }
  if (kind === "bevel") {
    const bevel = value as OfficePictureEffectValues["bevel"];
    return { ...bevel, width: scale(bevel.width), height: scale(bevel.height) };
  }
  return value;
}

function Preview<K extends OfficePictureEffectKind>({ kind, value }: { kind: K; value: OfficePictureEffectValues[K] }) {
  const css =
    kind === "preset"
      ? officePictureEffectsCss({ preset: value }, PREVIEW, { scale: PREVIEW_SCALE })
      : officePictureEffectsCss({ [kind]: scaled(kind, value) }, PREVIEW);
  return (
    <span className={styles.previewCell} aria-hidden="true">
      <span className={styles.preview} style={css?.wrapper as CSSProperties | undefined}>
        {css?.bevel ? <span style={css.bevel} /> : null}
      </span>
    </span>
  );
}

export function OfficePictureEffectsMenu({
  effects,
  shadow,
  accentColors,
  styleEffects,
  labels,
  disabled = false,
  disabledReason,
  controlId,
  testId,
  displayLabel,
  kinds,
  extra,
  iconSemanticId,
  onPick,
  onThreeDOptions,
}: OfficePictureEffectsMenuProps) {
  const shows = (kind: OfficePictureEffectKind) => !kinds || kinds.includes(kind);
  const current = normalizeOfficePictureEffects(effects) ?? {};
  const active = shadow || Object.values(current).some((value) => value !== "none");
  const has = (kind: OfficePictureEffectKind) => {
    const value = current[kind];
    return value !== undefined ? value !== "none" : styleEffects?.has(kind) === true;
  };

  // Office ends Preset and Bevel with 3-D Options... and 3-D Rotation with
  // 3-D Rotation Options...: the Format pane at its Effects page.
  function threeDOptionsRow(section: "format" | "rotation", label: string, rowTestId: string) {
    return (
      <>
        <DropdownMenuSeparator />
        <DropdownMenuItem
          disabled={!onThreeDOptions || disabled}
          data-testid={rowTestId}
          data-disabled-reason={onThreeDOptions ? undefined : labels.threeDOptionsUnavailable}
          title={onThreeDOptions ? undefined : labels.threeDOptionsUnavailable}
          onSelect={() => onThreeDOptions?.(section)}
        >
          {label}
        </DropdownMenuItem>
      </>
    );
  }

  function category(
    kind: OfficePictureEffectKind,
    gallery: ReactNode,
    variations?: "reflection" | "glow" | "softEdge",
    after?: ReactNode,
  ) {
    if (!shows(kind)) return null;
    return (
      <DropdownMenuSub key={kind}>
        <DropdownMenuSubTrigger data-testid={`${testId}-${kind}-menu`}>
          {labels.category(kind)}
        </DropdownMenuSubTrigger>
        <DropdownMenuSubContent className={styles.menu}>
          <DropdownMenuItem
            data-testid={`${testId}-${kind}-none`}
            data-state={has(kind) ? "unchecked" : "checked"}
            onSelect={() => {
              if (has(kind)) onPick(kind, null);
            }}
          >
            {labels.none(kind)}
          </DropdownMenuItem>
          <DropdownMenuSeparator />
          {variations ? <DropdownMenuLabel>{labels.variations(variations)}</DropdownMenuLabel> : null}
          {gallery}
          {after}
        </DropdownMenuSubContent>
      </DropdownMenuSub>
    );
  }

  function tile<K extends OfficePictureEffectKind>(kind: K, id: string, value: OfficePictureEffectValues[K], label: string) {
    const checked = sameOfficePictureEffect(current[kind], value);
    return (
      <DropdownMenuItem
        key={id}
        className={styles.tile}
        data-testid={`${testId}-${kind}-${id}`}
        data-checked={checked ? "true" : undefined}
        aria-label={label}
        title={label}
        onSelect={() => onPick(kind, value)}
      >
        <Preview kind={kind} value={value} />
      </DropdownMenuItem>
    );
  }

  return (
    <DropdownMenu modal={false}>
      <DropdownMenuTrigger asChild>
        <OfficeHomeMenuTrigger
          active={active}
          controlId={controlId}
          disabled={disabled}
          disabledReason={disabledReason}
          hasMenuIndicator
          icon={<OfficeSemanticIcon semanticId={iconSemanticId ?? "shape.effects"} size={18} />}
          label={labels.button}
          displayLabel={displayLabel ?? labels.button}
          testId={testId}
          variant="compact"
        />
      </DropdownMenuTrigger>
      <DropdownMenuContent align="start" className={styles.menu}>
        {shows("preset") ? (
        <DropdownMenuSub>
          <DropdownMenuSubTrigger data-testid={`${testId}-presets`}>{labels.presets}</DropdownMenuSubTrigger>
          <DropdownMenuSubContent className={styles.menu}>
            <DropdownMenuItem
              data-testid={`${testId}-preset-none`}
              data-state={current.preset === undefined ? "checked" : "unchecked"}
              onSelect={() => {
                if (current.preset !== undefined) onPick("preset", null);
              }}
            >
              {labels.noPresets}
            </DropdownMenuItem>
            <DropdownMenuSeparator />
            <div className={styles.grid} style={{ "--columns": 4 } as CSSProperties}>
              {OFFICE_PICTURE_EFFECT_PRESETS.map((preset) => tile("preset", String(preset), preset, labels.preset(preset)))}
            </div>
            {threeDOptionsRow("format", labels.threeDOptions, `${testId}-three-d-options`)}
          </DropdownMenuSubContent>
        </DropdownMenuSub>
        ) : null}
        {shows("shadow") ? (
        <DropdownMenuSub>
          <DropdownMenuSubTrigger data-testid={`${testId}-shadow-menu`}>
            {labels.category("shadow")}
          </DropdownMenuSubTrigger>
          <DropdownMenuSubContent className={styles.menu}>
            <OfficeShadowGallery
              value={current.shadow}
              hasShadow={shadow || has("shadow")}
              labels={{
                none: labels.none("shadow"),
                group: labels.shadowGroup,
                variation: labels.shadowVariation,
              }}
              testId={testId}
              onPick={(value) => onPick("shadow", value)}
            />
          </DropdownMenuSubContent>
        </DropdownMenuSub>
        ) : null}
        {category(
          "reflection",
          <div className={styles.grid} style={{ "--columns": 3 } as CSSProperties}>
            {OFFICE_REFLECTION_VARIATIONS.map((variation) =>
              tile("reflection", variation.id, variation.value, labels.reflection(variation.id)),
            )}
          </div>,
          "reflection",
        )}
        {category(
          "glow",
          <div className={styles.grid} style={{ "--columns": 6 } as CSSProperties}>
            {OFFICE_GLOW_SIZES_PT.flatMap((sizePt) =>
              OFFICE_GLOW_ACCENTS.map((accent, index) =>
                tile(
                  "glow",
                  `${sizePt}pt-${accent}`,
                  officeGlowVariation(sizePt, accent, accentColors[accent]),
                  labels.glow(sizePt, index + 1),
                ),
              ),
            )}
          </div>,
          "glow",
        )}
        {category(
          "softEdge",
          <div className={styles.list}>
            {OFFICE_SOFT_EDGE_SIZES_PT.map((sizePt) => {
              const value = officeSoftEdgeVariation(sizePt);
              return (
                <DropdownMenuItem
                  key={sizePt}
                  data-testid={`${testId}-softEdge-${sizePt}`}
                  data-checked={sameOfficePictureEffect(current.softEdge, value) ? "true" : undefined}
                  onSelect={() => onPick("softEdge", value)}
                >
                  <Preview kind="softEdge" value={value} />
                  {labels.softEdge(sizePt)}
                </DropdownMenuItem>
              );
            })}
          </div>,
          "softEdge",
        )}
        {category(
          "bevel",
          <div className={styles.grid} style={{ "--columns": 4 } as CSSProperties}>
            {OFFICE_BEVEL_PRESETS.map((preset) =>
              tile("bevel", preset, officeBevelVariation(preset), labels.bevel(preset)),
            )}
          </div>,
          undefined,
          threeDOptionsRow("format", labels.threeDOptions, `${testId}-bevel-three-d-options`),
        )}
        {category(
          "rotation3d",
          <>
            {OFFICE_ROTATION_GALLERY.map((section) => (
              <div key={section.section}>
                <DropdownMenuLabel>{labels.rotationSection(section.section)}</DropdownMenuLabel>
                <div className={styles.grid} style={{ "--columns": 4 } as CSSProperties}>
                  {section.cameras.map((camera) =>
                    tile("rotation3d", camera, { camera }, labels.camera(camera)),
                  )}
                </div>
              </div>
            ))}
          </>,
          undefined,
          labels.threeDRotationOptions
            ? threeDOptionsRow("rotation", labels.threeDRotationOptions, `${testId}-rotation3d-options`)
            : null,
        )}
        {extra ? (
          <>
            <DropdownMenuSeparator />
            {extra}
          </>
        ) : null}
      </DropdownMenuContent>
    </DropdownMenu>
  );
}
