"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The body of an Office outline menu: Shape Outline, Picture Border and
 * Graphics Outline are one menu in PowerPoint -- the colour palette (theme
 * colours, standard colours, recent, No Outline, More Outline Colors...),
 * Eyedropper, then Weight, Sketched, Dashes and, for lines, Arrows, each a
 * submenu (Microsoft Support, "Add or remove a border on a text box, shape,
 * or SmartArt graphic"; "Add a border to a picture"). Shared by the Shape
 * Format tab's outline split and the deck's Picture Border so both draw the
 * same structure; a host leaves out what its object has not (a picture has
 * no Arrows).
 */

import type { ReactNode } from "react";

import {
  DropdownMenuItem,
  DropdownMenuRadioGroup,
  DropdownMenuRadioItem,
  DropdownMenuSeparator,
  DropdownMenuSub,
  DropdownMenuSubContent,
  DropdownMenuSubTrigger,
} from "../ui/dropdown-menu";
import {
  OFFICE_LINE_WEIGHT_PT,
  formatOfficeLineWeightPt,
} from "../../lib/office-line-weight";

import {
  OfficeColorPalettePanel,
  type OfficeColorPalettePanelProps,
} from "./OfficeColorPalettePanel";

export type OfficeOutlineDash = "solid" | "dashed" | "dotted";

/** Shape Outline > Sketched, in Office's order. */
export const OFFICE_OUTLINE_SKETCHES = ["none", "curved", "freehand", "scribble"] as const;
export type OfficeOutlineSketch = (typeof OFFICE_OUTLINE_SKETCHES)[number];

/** A short line drawn the way each Sketched choice draws one. */
function SketchIcon({ sketch }: { sketch: OfficeOutlineSketch }) {
  const d =
    sketch === "none"
      ? "M2 8 L30 8"
      : sketch === "curved"
        ? "M2 9 C9 6 15 10 22 7 S28 8 30 8"
        : sketch === "freehand"
          ? "M2 8 C5 5 7 11 10 8 S14 5 17 9 S22 6 25 9 S29 7 30 8"
          : "M2 8 L5 6 L7 10 L10 5 L12 11 L15 6 L17 10 L20 5 L22 11 L25 6 L27 10 L30 7";
  return (
    <svg width="32" height="16" viewBox="0 0 32 16" aria-hidden="true" style={{ flex: "none" }}>
      <path d={d} fill="none" stroke="currentColor" strokeWidth="1.4" strokeLinecap="round" strokeLinejoin="round" />
    </svg>
  );
}

interface Capability {
  readonly disabled?: boolean;
  readonly disabledReason?: string;
  readonly controlId?: string;
}

export interface OfficeOutlineMenuContentProps<D extends string = OfficeOutlineDash> {
  readonly palette: OfficeColorPalettePanelProps;
  /** Eyedropper, after the palette; absent where the browser has none. */
  readonly eyedropper?: { readonly label: string; readonly testId: string; readonly onPick: () => void };
  readonly weight: Capability & {
    readonly label: string;
    readonly testId: string;
    readonly valuePt: number | null;
    readonly onPick: (pt: number) => void;
    readonly field?: ReactNode;
  };
  /**
   * Shape Outline > Sketched: None, Curved, Freehand, Scribble (Office's
   * order). Without `onPick` it is listed and refused with its reason.
   */
  readonly sketched?: Capability & {
    readonly label: string;
    readonly testId: string;
    readonly value?: OfficeOutlineSketch;
    readonly optionLabel?: (sketch: OfficeOutlineSketch) => string;
    readonly onPick?: (sketch: OfficeOutlineSketch) => void;
  };
  readonly dashes: Capability & {
    readonly label: string;
    readonly testId: string;
    readonly value: D | null;
    /**
     * The dashes listed, in order; absent, the three a picture border and a
     * shape outline model (solid, dashed, dotted). Graphics Outline lists
     * PowerPoint's eight by their DrawingML names (PPT-A161).
     */
    readonly options?: readonly D[];
    readonly optionLabel: (style: D) => string;
    /** A picture of each dash beside its name, where the host draws one. */
    readonly optionIcon?: (style: D) => ReactNode;
    readonly onPick: (style: D) => void;
  };
  readonly arrows?: Capability & {
    readonly label: string;
    readonly testId: string;
    readonly content: ReactNode;
  };
  readonly menuClassName?: string;
}

function trigger(capability: Capability, label: string, testId: string) {
  return (
    <DropdownMenuSubTrigger
      data-testid={testId}
      data-control-id={capability.controlId}
      data-disabled-reason={capability.disabled ? capability.disabledReason : undefined}
      title={capability.disabled ? capability.disabledReason : undefined}
      disabled={capability.disabled}
    >
      {label}
    </DropdownMenuSubTrigger>
  );
}

const DEFAULT_DASHES = ["solid", "dashed", "dotted"] as const;

export function OfficeOutlineMenuContent<D extends string = OfficeOutlineDash>({
  palette,
  eyedropper,
  weight,
  sketched,
  dashes,
  arrows,
  menuClassName,
}: OfficeOutlineMenuContentProps<D>) {
  const dashOptions = dashes.options ?? (DEFAULT_DASHES as unknown as readonly D[]);
  return (
    <>
      <OfficeColorPalettePanel {...palette} />
      {eyedropper ? (
        <DropdownMenuItem data-testid={eyedropper.testId} onSelect={eyedropper.onPick}>
          {eyedropper.label}
        </DropdownMenuItem>
      ) : null}
      <DropdownMenuSeparator />
      <DropdownMenuSub>
        {trigger(weight, weight.label, weight.testId)}
        <DropdownMenuSubContent className={menuClassName}>
          <DropdownMenuRadioGroup value={weight.valuePt === null ? "" : String(weight.valuePt)}>
            {OFFICE_LINE_WEIGHT_PT.map((pt) => (
              <DropdownMenuRadioItem
                key={pt}
                data-testid={`${weight.testId}-${pt}`}
                disabled={weight.disabled}
                onSelect={() => weight.onPick(pt)}
                value={String(pt)}
              >
                {formatOfficeLineWeightPt(pt)} pt
              </DropdownMenuRadioItem>
            ))}
          </DropdownMenuRadioGroup>
          {weight.field ? (
            <>
              <DropdownMenuSeparator />
              {weight.field}
            </>
          ) : null}
        </DropdownMenuSubContent>
      </DropdownMenuSub>
      {sketched ? (
        <DropdownMenuSub>
          {sketched.onPick
            ? trigger(sketched, sketched.label, sketched.testId)
            : trigger({ disabled: true, disabledReason: sketched.disabledReason }, sketched.label, sketched.testId)}
          {sketched.onPick ? (
            <DropdownMenuSubContent className={menuClassName}>
              <DropdownMenuRadioGroup value={sketched.value ?? "none"}>
                {OFFICE_OUTLINE_SKETCHES.map((sketch) => (
                  <DropdownMenuRadioItem
                    key={sketch}
                    data-testid={`${sketched.testId}-${sketch}`}
                    disabled={sketched.disabled}
                    onSelect={() => sketched.onPick?.(sketch)}
                    value={sketch}
                  >
                    <SketchIcon sketch={sketch} />
                    {sketched.optionLabel?.(sketch) ?? sketch}
                  </DropdownMenuRadioItem>
                ))}
              </DropdownMenuRadioGroup>
            </DropdownMenuSubContent>
          ) : null}
        </DropdownMenuSub>
      ) : null}
      <DropdownMenuSub>
        {trigger(dashes, dashes.label, dashes.testId)}
        <DropdownMenuSubContent className={menuClassName}>
          <DropdownMenuRadioGroup value={dashes.value ?? ""}>
            {dashOptions.map((style) => (
              <DropdownMenuRadioItem
                key={style}
                data-testid={`${dashes.testId}-${style}`}
                disabled={dashes.disabled}
                onSelect={() => dashes.onPick(style)}
                value={style}
              >
                {dashes.optionIcon?.(style)}
                {dashes.optionLabel(style)}
              </DropdownMenuRadioItem>
            ))}
          </DropdownMenuRadioGroup>
        </DropdownMenuSubContent>
      </DropdownMenuSub>
      {arrows ? (
        <DropdownMenuSub>
          {trigger(arrows, arrows.label, arrows.testId)}
          <DropdownMenuSubContent className={menuClassName}>{arrows.content}</DropdownMenuSubContent>
        </DropdownMenuSub>
      ) : null}
    </>
  );
}
