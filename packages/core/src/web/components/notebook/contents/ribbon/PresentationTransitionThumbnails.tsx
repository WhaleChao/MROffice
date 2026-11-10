/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { useId, type ReactNode } from "react";

import type { AuthorableSlideTransition } from "../../../../lib/presentation-transitions";

/**
 * The Transitions gallery's effect pictures: a small slide that shows what
 * the effect does, the slide being left in pale blue and the slide arriving
 * in deep blue, with the direction it moves. PowerPoint's gallery shows the
 * effect the same way (a preview, not a glyph), so a tile reads at a glance.
 * Drawn on a 48 x 32 slide (3:2), sized by the caller.
 */
export type TransitionThumbnailKind = AuthorableSlideTransition | "unavailable";

const W = 48;
const H = 32;
const OLD = "#e4eef9";
const OLD_EDGE = "#b7cde6";
const FRAME = "#9db8d6";
const ARROW_ON_OLD = "#3f7cc4";

interface Paint {
  newFill: string;
  newDeep: string;
  fadeH: string;
  softLeft: string;
  iris: string;
  flash: string;
  shadow: string;
}

/** A right-pointing arrow centred on the origin, turned by `angle`. */
function Arrow({ x, y, angle = 0, scale = 1, color = "white" }: {
  x: number; y: number; angle?: number; scale?: number; color?: string;
}) {
  return (
    <path
      d="M-4.2-1.3H0V-3.6L4.4 0 0 3.6V1.3H-4.2Z"
      fill={color}
      transform={`translate(${x} ${y}) rotate(${angle}) scale(${scale})`}
    />
  );
}

function Slide({ x, y, w, h, fill, stroke, rx = 1.6, opacity }: {
  x: number; y: number; w: number; h: number; fill: string; stroke?: string; rx?: number; opacity?: number;
}) {
  return <rect x={x} y={y} width={w} height={h} rx={rx} fill={fill} stroke={stroke} strokeWidth={stroke ? 0.8 : undefined} opacity={opacity} />;
}

function artwork(kind: TransitionThumbnailKind, p: Paint): ReactNode {
  const full = (fill: string) => <rect x="0" y="0" width={W} height={H} fill={fill} />;
  switch (kind) {
    case "none":
      return full(OLD);
    case "unavailable":
      return (
        <>
          {full(OLD)}
          <rect x="3" y="3" width={W - 6} height={H - 6} rx="2" fill="none" stroke={OLD_EDGE} strokeDasharray="2 2" />
        </>
      );
    case "morph":
      return (
        <>
          {full(OLD)}
          <rect x="6" y="6" width="13" height="11" rx="2" fill="#b9d1ec" />
          <path d="M14 21c5 5 12 5 17 0" fill="none" stroke={ARROW_ON_OLD} strokeWidth="1.3" strokeDasharray="1.6 1.6" />
          <rect x="26" y="9" width="15" height="15" rx="3" fill={p.newFill} transform="rotate(18 33.5 16.5)" />
        </>
      );
    case "fade":
      return full(p.fadeH);
    case "push":
      return (
        <>
          {full(OLD)}
          <rect x="0" y="13" width={W} height={H - 13} fill={p.newFill} />
          <Arrow x={24} y={22.5} angle={-90} />
        </>
      );
    case "wipe":
      return (
        <>
          {full(OLD)}
          <rect x="12" y="0" width={W - 12} height={H} fill={p.softLeft} />
          <Arrow x={34} y={16} angle={180} />
        </>
      );
    case "split":
      return (
        <>
          {full(OLD)}
          <rect x="17" y="0" width="14" height={H} fill={p.newFill} />
          <Arrow x={8.5} y={16} angle={180} color={ARROW_ON_OLD} />
          <Arrow x={39.5} y={16} color={ARROW_ON_OLD} />
        </>
      );
    case "reveal":
      return (
        <>
          {full(OLD)}
          <rect x="10" y="0" width="10" height={H} fill={p.shadow} />
          <Slide x={16} y={0} w={32} h={H} fill={p.newFill} rx={0} />
        </>
      );
    case "cut":
      return (
        <>
          {full(OLD)}
          <Slide x={4} y={4} w={26} h={17} fill="#c9dcf0" stroke={OLD_EDGE} />
          <Slide x={17} y={11} w={27} h={18} fill={p.newFill} />
        </>
      );
    case "random-bars": {
      const widths = [2, 1.5, 3, 1, 2.5, 1.5, 2, 3.5, 1, 2, 1.5, 3, 2, 1, 2.5, 1.5, 2, 3, 1.5, 2.5];
      let x = 0;
      return (
        <>
          {full(OLD)}
          {widths.map((width, index) => {
            const at = x;
            x += width + (index % 3 === 0 ? 0.8 : 0.5);
            return index % 2 === 0 && at < W ? <rect key={index} x={at} y="0" width={width} height={H} fill={p.newFill} /> : null;
          })}
        </>
      );
    }
    case "shape":
      return (
        <>
          {full(OLD)}
          <circle cx="24" cy="16" r="11.5" fill={p.iris} />
        </>
      );
    case "uncover":
      return (
        <>
          {full(p.newFill)}
          <rect x="0" y="0" width="22" height={H} fill={OLD} />
          <rect x="22" y="0" width="5" height={H} fill={p.shadow} />
          <Arrow x={11} y={16} angle={180} color={ARROW_ON_OLD} />
        </>
      );
    case "cover":
      return (
        <>
          {full(OLD)}
          <rect x="15" y="0" width="5" height={H} fill={p.shadow} />
          <rect x="19" y="0" width={W - 19} height={H} fill={p.newFill} />
          <Arrow x={33.5} y={16} angle={180} />
        </>
      );
    case "flash":
      return (
        <>
          {full(p.newFill)}
          <rect x="0" y="0" width={W} height={H} fill={p.flash} />
          <path d="M24 7l1.6 6.4L32 15l-6.4 1.6L24 23l-1.6-6.4L16 15l6.4-1.6Z" fill="white" />
        </>
      );
    case "dissolve": {
      const cells: ReactNode[] = [];
      for (let row = 0; row < 8; row += 1) {
        for (let column = 0; column < 12; column += 1) {
          if ((row * 7 + column * 5 + row * column) % 3 !== 0) continue;
          cells.push(<rect key={`${row}-${column}`} x={column * 4} y={row * 4} width="4" height="4" fill={p.newFill} />);
        }
      }
      return (
        <>
          {full(OLD)}
          {cells}
        </>
      );
    }
    case "zoom":
      return (
        <>
          {full(OLD)}
          <Slide x={15} y={10} w={18} h={12} fill={p.newFill} />
          <Arrow x={9} y={6} angle={-146} scale={0.8} color={ARROW_ON_OLD} />
          <Arrow x={39} y={6} angle={-34} scale={0.8} color={ARROW_ON_OLD} />
          <Arrow x={9} y={26} angle={146} scale={0.8} color={ARROW_ON_OLD} />
          <Arrow x={39} y={26} angle={34} scale={0.8} color={ARROW_ON_OLD} />
        </>
      );
    case "checkerboard": {
      const cells: ReactNode[] = [];
      for (let row = 0; row < 4; row += 1) {
        for (let column = 0; column < 6; column += 1) {
          if ((row + column) % 2 === 0) {
            cells.push(<rect key={`${row}-${column}`} x={column * 8} y={row * 8} width={column < 3 ? 8 : 4} height="8" fill={p.newFill} />);
          }
        }
      }
      return (
        <>
          {full(OLD)}
          {cells}
        </>
      );
    }
    case "blinds":
      return (
        <>
          {full(OLD)}
          {[0, 1, 2, 3, 4].map((band) => (
            <rect key={band} x="0" y={band * 6.4} width={W} height="3.8" fill={p.newFill} />
          ))}
        </>
      );
    case "clock":
      return (
        <>
          {full(OLD)}
          <path d="M24 16 24-24A40 40 0 0 1 44 50.6Z" fill={p.newFill} />
          <path d="M24 16 24 4" stroke="white" strokeWidth="1.4" strokeLinecap="round" />
          <circle cx="24" cy="16" r="1.4" fill="white" />
        </>
      );
    case "comb":
      return (
        <>
          {full(OLD)}
          {[0, 1, 2, 3, 4, 5, 6, 7].map((row) => (
            <rect key={row} x={row % 2 === 0 ? 0 : 16} y={row * 4} width="32" height="4" fill={p.newFill} />
          ))}
        </>
      );
    case "fall-over":
      return (
        <>
          {full(OLD)}
          <path d="M8 28h32l-6-17H14Z" fill={p.newFill} />
          <path d="M14 7c3-3 8-4 12-3" fill="none" stroke={ARROW_ON_OLD} strokeWidth="1.3" />
          <Arrow x={27} y={4.2} angle={8} scale={0.7} color={ARROW_ON_OLD} />
        </>
      );
    case "drape":
      return (
        <>
          {full(p.newFill)}
          <path d="M0 0h48v8c-4 6-6 14-6 24H30c0-9-2-16-6-20-4 4-6 11-6 20H6C6 18 3 10 0 6Z" fill={OLD} />
        </>
      );
    case "curtains":
      return (
        <>
          {full(p.newFill)}
          <path d="M0 0h17c-2 8-3 20 1 32H0Z" fill={OLD} />
          <path d="M48 0H31c2 8 3 20-1 32h18Z" fill={OLD} />
          <path d="M5 0v32M10 0c-1 10-1 22 1 32M38 0v32M43 0c1 10 1 22-1 32" stroke={OLD_EDGE} strokeWidth="0.8" fill="none" />
        </>
      );
    case "wind":
      return (
        <>
          {full(OLD)}
          <path d="M14 4h32v24H8Z" fill={p.newFill} />
          <path d="M2 9h9M4 15h7M1 21h8" stroke={ARROW_ON_OLD} strokeWidth="1.3" strokeLinecap="round" />
        </>
      );
    case "prestige":
      return (
        <>
          {full(p.newFill)}
          <path d="M10 27h28l-4-21H14Z" fill={OLD} />
          <path d="M10 27h28" stroke={OLD_EDGE} strokeWidth="0.8" />
        </>
      );
    case "fracture":
      return (
        <>
          {full(OLD)}
          <path d="M6 5l10 2-4 9-8-2Z" fill={p.newFill} />
          <path d="M20 4l12 1-3 10-11-1Z" fill={p.newFill} />
          <path d="M35 6l9 3-2 9-9-3Z" fill={p.newFill} />
          <path d="M10 19l11 2-2 9-11-2Z" fill={p.newFill} />
          <path d="M25 18l12 1-1 10-12-1Z" fill={p.newFill} />
        </>
      );
    case "crush":
      return (
        <>
          {full(p.newFill)}
          <path d="M17 9l7-3 8 3 3 7-2 8-8 3-7-2-4-7Z" fill={OLD} />
          <path d="M24 6l-1 8 9 2M17 9l6 5-2 10M35 16l-12-2" stroke={OLD_EDGE} strokeWidth="0.8" fill="none" />
        </>
      );
    case "peel-off":
      return (
        <>
          {full(p.newFill)}
          <path d="M0 0h48v12L26 32H0Z" fill={OLD} />
          <path d="M48 12 26 32c6-1 12-6 16-11Z" fill="#c4d8ee" stroke={OLD_EDGE} strokeWidth="0.6" />
        </>
      );
    case "page-curl":
      return (
        <>
          {full(p.newFill)}
          <path d="M0 0h34c-4 8-4 24 0 32H0Z" fill={OLD} />
          <path d="M34 0c-4 8-4 24 0 32 5-6 6-26 0-32Z" fill="#c4d8ee" stroke={OLD_EDGE} strokeWidth="0.6" />
        </>
      );
    case "airplane":
      return (
        <>
          {full(p.newFill)}
          <path d="M10 22 38 8 22 26l-3-5Z" fill={OLD} />
          <path d="M19 21 38 8" stroke={OLD_EDGE} strokeWidth="0.8" />
          <path d="M4 28c3-1 5-2 7-4" stroke="white" strokeWidth="1.1" strokeDasharray="1.5 1.5" fill="none" />
        </>
      );
    case "origami":
      return (
        <>
          {full(p.newFill)}
          <path d="M8 24 20 8l6 10 14-8-6 16Z" fill={OLD} />
          <path d="M20 8l6 10-3 8M26 18l8 8" stroke={OLD_EDGE} strokeWidth="0.8" fill="none" />
        </>
      );
    case "switch":
      return (
        <>
          {full(OLD)}
          <Slide x={4} y={7} w={22} h={16} fill="#c9dcf0" stroke={OLD_EDGE} />
          <rect x="20" y="8" width="23" height="17" rx="1.6" fill={p.newFill} transform="rotate(8 31.5 16.5)" />
        </>
      );
    case "flip":
      return (
        <>
          {full(OLD)}
          <path d="M12 7l22 3v12L12 25Z" fill={p.newFill} />
          <path d="M11 4c8-3 18-3 26 1" fill="none" stroke={ARROW_ON_OLD} strokeWidth="1.3" />
          <Arrow x={37} y={5.3} angle={20} scale={0.7} color={ARROW_ON_OLD} />
        </>
      );
    case "gallery":
      return (
        <>
          {full(OLD)}
          <path d="M3 6l18 2v16L3 26Z" fill={p.newFill} />
          <path d="M24 9l11 1.3v11.4L24 23Z" fill="#c9dcf0" />
          <path d="M37 11l8 1v8l-8 1Z" fill="#d6e4f3" />
        </>
      );
    case "ferris-wheel":
      return (
        <>
          {full(OLD)}
          <circle cx="24" cy="16" r="11" fill="none" stroke={OLD_EDGE} strokeWidth="1" />
          <Slide x={19} y={2} w={10} h={7} fill={p.newFill} rx={1} />
          <Slide x={31} y={13} w={9} h={6} fill="#c9dcf0" rx={1} />
          <Slide x={19.5} y={23.5} w={9} h={6} fill="#c9dcf0" rx={1} />
          <Slide x={8} y={13} w={9} h={6} fill="#c9dcf0" rx={1} />
        </>
      );
    case "conveyor":
      return (
        <>
          {full(OLD)}
          <path d="M2 26h44" stroke={OLD_EDGE} strokeWidth="1.4" />
          <path d="M6 9l15-1v16L6 24Z" fill="#c9dcf0" />
          <path d="M25 6l19 2v16l-19 2Z" fill={p.newFill} />
          <Arrow x={23} y={29} angle={180} scale={0.7} color={ARROW_ON_OLD} />
        </>
      );
    case "doors":
      return (
        <>
          {full(p.newFill)}
          <path d="M0 0l12 5v22L0 32Z" fill={OLD} />
          <path d="M48 0l-12 5v22l12 5Z" fill={OLD} />
        </>
      );
    case "window":
      return (
        <>
          {full(p.newFill)}
          <path d="M0 0l12 5v22L0 32Z" fill={OLD} />
          <path d="M48 0l-12 5v22l12 5Z" fill={OLD} />
          <path d="M6 2.5v27M0 16h12M42 2.5v27M36 16h12" stroke={OLD_EDGE} strokeWidth="0.8" />
        </>
      );
    case "vortex":
      return (
        <>
          {full(OLD)}
          {[
            [24, 16, 9, 0], [33, 10, 6, 30], [36, 20, 5, 60], [27, 26, 4, 90], [16, 24, 4, 120], [12, 13, 3, 150], [19, 7, 3, 180],
          ].map(([x, y, s, a]) => (
            <rect key={`${x}-${y}`} x={x - s / 2} y={y - s / 2} width={s} height={s} rx="0.6" fill={p.newFill} transform={`rotate(${a} ${x} ${y})`} />
          ))}
        </>
      );
    case "pan":
      return (
        <>
          {full(OLD)}
          <Slide x={14} y={2} w={32} h={22} fill={p.newFill} />
          <Arrow x={9} y={26} angle={-35} color={ARROW_ON_OLD} />
        </>
      );
    case "honeycomb": {
      const hex = (cx: number, cy: number, fill: string) => (
        <path key={`${cx}-${cy}`} d={`M${cx - 4} ${cy}l2-3.5h4l2 3.5-2 3.5h-4Z`} fill={fill} stroke="white" strokeWidth="0.5" />
      );
      const cells: ReactNode[] = [];
      for (let column = 0; column < 9; column += 1) {
        for (let row = 0; row < 5; row += 1) {
          const cx = 2 + column * 6;
          const cy = (column % 2 === 0 ? 3.5 : 7) + row * 7;
          cells.push(hex(cx, cy, (column * 3 + row * 5) % 4 === 0 ? OLD : p.newFill));
        }
      }
      return (
        <>
          {full(OLD)}
          {cells}
        </>
      );
    }
    case "ripple":
      return (
        <>
          {full(p.newFill)}
          {[4, 8, 12, 16, 20].map((r) => (
            <circle key={r} cx="24" cy="16" r={r} fill="none" stroke="white" strokeOpacity={1 - r / 26} strokeWidth="1.2" />
          ))}
        </>
      );
    case "glitter": {
      const cells: ReactNode[] = [];
      for (let column = 0; column < 8; column += 1) {
        for (let row = 0; row < 5; row += 1) {
          const cx = 3 + column * 6 + (row % 2) * 3;
          const cy = 3 + row * 6.5;
          const fill = (column + row * 3) % 3 === 0 ? OLD : p.newFill;
          cells.push(<path key={`${column}-${row}`} d={`M${cx} ${cy - 3.2}l3 3.2-3 3.2-3-3.2Z`} fill={fill} />);
        }
      }
      return (
        <>
          {full(OLD)}
          {cells}
          <path d="M37 5l.9 2.3 2.3.9-2.3.9-.9 2.3-.9-2.3-2.3-.9 2.3-.9Z" fill="white" />
        </>
      );
    }
    case "shred":
      return (
        <>
          {full(p.newFill)}
          {[0, 1, 2, 3, 4, 5, 6, 7].map((strip) => (
            <rect key={strip} x={strip * 6 + 0.6} y={strip % 2 === 0 ? -2 : 3} width="4.6" height={H} fill={OLD} />
          ))}
        </>
      );
    case "fly-through":
      return (
        <>
          {full(OLD)}
          <Slide x={6} y={3} w={36} h={26} fill="none" stroke={OLD_EDGE} />
          <Slide x={11} y={7} w={26} h={18} fill="none" stroke="#9fbfe2" />
          <Slide x={16} y={10.5} w={16} h={11} fill={p.newFill} />
        </>
      );
    case "random":
      return (
        <>
          {full(p.newFill)}
          <path d="M8 21h7c6 0 9-10 15-10h6" fill="none" stroke="white" strokeWidth="1.6" />
          <path d="M8 11h7c6 0 9 10 15 10h6" fill="none" stroke="white" strokeWidth="1.6" />
          <Arrow x={38} y={11} scale={0.7} />
          <Arrow x={38} y={21} scale={0.7} />
        </>
      );
    case "cube":
      return (
        <>
          {full(OLD)}
          <path d="M8 9l14 3v17L8 26Z" fill="#c9dcf0" />
          <path d="M22 12l18-4v18l-18 3Z" fill={p.newFill} />
          <path d="M8 9l15-5 17 4-18 4Z" fill={p.newDeep} opacity="0.55" />
        </>
      );
    case "box":
      return (
        <>
          {full(OLD)}
          <path d="M6 7l16 4v18L6 25Z" fill="#c9dcf0" />
          <path d="M22 11l20-4v18l-20 4Z" fill={p.newFill} />
        </>
      );
    case "rotate":
      return (
        <>
          {full(OLD)}
          <path d="M16 6l16 2v16l-16 2Z" fill={p.newFill} />
          <path d="M7 16c0-6 7-10 17-10" fill="none" stroke={ARROW_ON_OLD} strokeWidth="1.3" />
          <path d="M41 16c0 6-7 10-17 10" fill="none" stroke={ARROW_ON_OLD} strokeWidth="1.3" />
          <Arrow x={24.5} y={6} scale={0.7} color={ARROW_ON_OLD} />
          <Arrow x={23.5} y={26} angle={180} scale={0.7} color={ARROW_ON_OLD} />
        </>
      );
    case "orbit":
      return (
        <>
          {full(OLD)}
          <ellipse cx="24" cy="17" rx="19" ry="8" fill="none" stroke={OLD_EDGE} strokeWidth="1.1" />
          <Slide x={15} y={6} w={18} h={13} fill={p.newFill} />
          <Arrow x={42} y={18} angle={95} scale={0.7} color={ARROW_ON_OLD} />
        </>
      );
    default:
      return full(OLD);
  }
}

export function PresentationTransitionThumbnail({
  kind,
  width = 50,
}: {
  kind: TransitionThumbnailKind;
  width?: number;
}) {
  const id = useId().replaceAll(":", "");
  const paint: Paint = {
    newFill: `url(#${id}n)`,
    newDeep: "#2d65a8",
    fadeH: `url(#${id}f)`,
    softLeft: `url(#${id}w)`,
    iris: `url(#${id}i)`,
    flash: `url(#${id}l)`,
    shadow: `url(#${id}s)`,
  };
  return (
    <svg
      width={width}
      height={(width * H) / W}
      // A menu row forces every icon inside it to 16 px (`[&_svg]:size-4`);
      // the picture keeps the size it was asked for.
      style={{ width, height: (width * H) / W }}
      viewBox={`0 0 ${W} ${H}`}
      aria-hidden="true"
      data-presentation-transition-thumbnail={kind}
    >
      <defs>
        <linearGradient id={`${id}n`} x1="0" y1="0" x2="0" y2="1">
          <stop offset="0" stopColor="#7fb2e5" />
          <stop offset="1" stopColor="#3570b6" />
        </linearGradient>
        <linearGradient id={`${id}f`} x1="0" y1="0" x2="1" y2="0">
          <stop offset="0" stopColor={OLD} />
          <stop offset="0.35" stopColor="#c6dbf1" />
          <stop offset="1" stopColor="#3f7bc0" />
        </linearGradient>
        <linearGradient id={`${id}w`} x1="0" y1="0" x2="1" y2="0">
          <stop offset="0" stopColor="#5d95d3" stopOpacity="0" />
          <stop offset="0.3" stopColor="#5d95d3" stopOpacity="0.85" />
          <stop offset="1" stopColor="#3570b6" />
        </linearGradient>
        <radialGradient id={`${id}i`} cx="0.5" cy="0.5" r="0.5">
          <stop offset="0" stopColor="#3570b6" />
          <stop offset="0.75" stopColor="#6aa3dd" />
          <stop offset="1" stopColor="#9cc2ea" stopOpacity="0.4" />
        </radialGradient>
        <radialGradient id={`${id}l`} cx="0.5" cy="0.47" r="0.62">
          <stop offset="0" stopColor="white" />
          <stop offset="0.45" stopColor="white" stopOpacity="0.75" />
          <stop offset="1" stopColor="white" stopOpacity="0" />
        </radialGradient>
        <linearGradient id={`${id}s`} x1="0" y1="0" x2="1" y2="0">
          <stop offset="0" stopColor="#1f4f8a" stopOpacity="0" />
          <stop offset="1" stopColor="#1f4f8a" stopOpacity="0.35" />
        </linearGradient>
        <clipPath id={`${id}c`}>
          <rect x="0" y="0" width={W} height={H} rx="2.4" />
        </clipPath>
      </defs>
      <g clipPath={`url(#${id}c)`}>{artwork(kind, paint)}</g>
      <rect x="0.4" y="0.4" width={W - 0.8} height={H - 0.8} rx="2.2" fill="none" stroke={FRAME} strokeWidth="0.8" />
    </svg>
  );
}
