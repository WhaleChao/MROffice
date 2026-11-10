"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Text Effects > Transform, as Word, PowerPoint and Excel list it: No
 * Transform, Follow Path (4) and Warp (36), each a picture of its paths. A
 * pick is a preset text warp (`a:prstTxWarp`, lib/office-text-warp), which
 * all three editors draw and write.
 */

import type { CSSProperties } from "react";

import {
  DropdownMenuItem,
  DropdownMenuLabel,
  DropdownMenuSeparator,
  DropdownMenuSub,
  DropdownMenuSubContent,
  DropdownMenuSubTrigger,
} from "../ui/dropdown-menu";
import { textWarpPaths } from "../../lib/office-text-warp";

/** Follow Path, in Office's order. */
export const OFFICE_TEXT_TRANSFORM_FOLLOW = ["textArchUp", "textArchDown", "textCircle", "textButton"] as const;

/** Warp, in Office's order (PowerPoint for Mac 16.86, six to a row). */
export const OFFICE_TEXT_TRANSFORM_WARP = [
  "textPlain",
  "textStop",
  "textTriangle",
  "textTriangleInverted",
  "textChevron",
  "textChevronInverted",
  "textRingInside",
  "textRingOutside",
  "textArchUpPour",
  "textArchDownPour",
  "textCirclePour",
  "textButtonPour",
  "textCurveUp",
  "textCurveDown",
  "textCanUp",
  "textCanDown",
  "textWave1",
  "textWave2",
  "textDoubleWave1",
  "textWave4",
  "textInflate",
  "textDeflate",
  "textInflateBottom",
  "textDeflateBottom",
  "textInflateTop",
  "textDeflateTop",
  "textDeflateInflate",
  "textDeflateInflateDeflate",
  "textFadeRight",
  "textFadeLeft",
  "textFadeUp",
  "textFadeDown",
  "textSlantUp",
  "textSlantDown",
  "textCascadeUp",
  "textCascadeDown",
] as const;

const W = 44;
const H = 26;

function TransformPreview({ token }: { token: string }) {
  const paths = textWarpPaths(token, W - 6, H - 6, null, 12) ?? [];
  return (
    <svg width={W} height={H} viewBox={`-3 -3 ${W} ${H}`} aria-hidden="true">
      {paths.map((path, index) => (
        <polyline
          key={index}
          points={path.map(([x, y]) => `${x.toFixed(1)},${y.toFixed(1)}`).join(" ")}
          fill="none"
          stroke="currentColor"
          strokeWidth="1.3"
          strokeLinecap="round"
          strokeLinejoin="round"
        />
      ))}
    </svg>
  );
}

export interface OfficeTextTransformLabels {
  readonly transform: string;
  readonly noTransform: string;
  readonly followPath: string;
  readonly warp: string;
  readonly name: (token: string) => string;
}

export function OfficeTextTransformSubmenu({
  value,
  onPick,
  labels,
  testId,
  disabled = false,
  disabledReason,
  menuClassName,
}: {
  /** The warp the text wears (a preset token), or null for none. */
  readonly value: string | null;
  readonly onPick: (token: string | null) => void;
  readonly labels: OfficeTextTransformLabels;
  readonly testId: string;
  readonly disabled?: boolean;
  readonly disabledReason?: string;
  readonly menuClassName?: string;
}) {
  const grid = (columns: number): CSSProperties => ({
    display: "grid",
    gridTemplateColumns: `repeat(${columns}, auto)`,
    gap: 2,
    padding: "2px 6px",
  });
  const tile = (token: string) => (
    <DropdownMenuItem
      key={token}
      data-testid={`${testId}-${token}`}
      data-checked={value === token ? "true" : undefined}
      aria-label={labels.name(token)}
      title={labels.name(token)}
      onSelect={() => onPick(token)}
      style={{
        padding: 2,
        justifyContent: "center",
        outline: value === token ? "1.5px solid var(--office-accent, #c43e1c)" : undefined,
      }}
    >
      <TransformPreview token={token} />
    </DropdownMenuItem>
  );
  return (
    <DropdownMenuSub>
      <DropdownMenuSubTrigger
        data-testid={`${testId}-menu`}
        disabled={disabled}
        data-disabled-reason={disabled ? disabledReason : undefined}
        title={disabled ? disabledReason : undefined}
      >
        {labels.transform}
      </DropdownMenuSubTrigger>
      <DropdownMenuSubContent className={menuClassName}>
        <DropdownMenuLabel>{labels.noTransform}</DropdownMenuLabel>
        <DropdownMenuItem
          data-testid={`${testId}-none`}
          data-state={value ? "unchecked" : "checked"}
          onSelect={() => {
            if (value) onPick(null);
          }}
        >
          {labels.noTransform}
        </DropdownMenuItem>
        <DropdownMenuSeparator />
        <DropdownMenuLabel>{labels.followPath}</DropdownMenuLabel>
        <div style={grid(4)}>{OFFICE_TEXT_TRANSFORM_FOLLOW.map(tile)}</div>
        <DropdownMenuSeparator />
        <DropdownMenuLabel>{labels.warp}</DropdownMenuLabel>
        <div style={grid(6)}>{OFFICE_TEXT_TRANSFORM_WARP.map(tile)}</div>
      </DropdownMenuSubContent>
    </DropdownMenuSub>
  );
}
