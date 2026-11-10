/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { SVGProps } from "react";

export type DocumentHeaderFooterRibbonIconKind =
  | "page-number"
  | "go-header"
  | "go-footer"
  | "previous"
  | "next"
  | "delete-header"
  | "delete-footer"
  | "close-preview"
  | "date-time"
  | "document-info"
  | "picture";

const INK = "#475569";
const PAPER = "#ffffff";
const PAPER_SOFT = "#f8fafc";
const BLUE = "#0b8ed8";
const ORANGE = "#f97316";
const RED = "#ef4444";

function PageFrame({ band }: { band?: "header" | "footer" }) {
  return (
    <>
      <path
        d="M7.5 3.5h11l6 6v19h-17z"
        fill={PAPER}
        stroke={INK}
        strokeWidth="1.5"
        strokeLinejoin="round"
      />
      <path
        d="M18.5 3.5v6h6"
        fill={PAPER_SOFT}
        stroke={INK}
        strokeWidth="1.5"
        strokeLinejoin="round"
      />
      {band ? (
        <path
          d={band === "header" ? "M11 10.5h9" : "M11 24.5h9"}
          stroke={ORANGE}
          strokeWidth="2.4"
          strokeLinecap="round"
        />
      ) : null}
    </>
  );
}

function DirectionBadge({
  direction,
  compact = false,
}: {
  direction: "up" | "down" | "left" | "right";
  compact?: boolean;
}) {
  const centerX = compact ? 24 : 8;
  const centerY = compact ? 22 : direction === "up" ? 22 : 9;
  const arrow =
    direction === "up"
      ? "M8 24v-5m-2 2 2-2 2 2"
      : direction === "down"
        ? "M8 7v5m-2-2 2 2 2-2"
        : direction === "left"
          ? "M26 22h-5m2-2-2 2 2 2"
          : "M22 22h5m-2-2 2 2-2 2";
  return (
    <>
      <circle cx={centerX} cy={centerY} r="5" fill={BLUE} />
      <path
        d={arrow}
        fill="none"
        stroke="#fff"
        strokeWidth="1.5"
        strokeLinecap="round"
        strokeLinejoin="round"
      />
    </>
  );
}

function DeleteBadge() {
  return (
    <>
      <circle cx="24" cy="10" r="5.2" fill={RED} />
      <path
        d="m21.8 7.8 4.4 4.4m0-4.4-4.4 4.4"
        stroke="#fff"
        strokeWidth="1.5"
        strokeLinecap="round"
      />
    </>
  );
}

export function DocumentHeaderFooterRibbonIcon({
  kind,
  size = 28,
  ...props
}: {
  kind: DocumentHeaderFooterRibbonIconKind;
  size?: number;
} & Omit<SVGProps<SVGSVGElement>, "height" | "width">) {
  if (kind === "close-preview") {
    return (
      <svg
        {...props}
        aria-hidden="true"
        data-docs-header-footer-icon={kind}
        height={size}
        viewBox="0 0 32 32"
        width={size}
      >
        <rect x="4" y="4" width="24" height="24" rx="4" fill={RED} />
        <path
          d="m10 10 12 12m0-12L10 22"
          stroke="#fff"
          strokeWidth="1.8"
          strokeLinecap="round"
        />
      </svg>
    );
  }

  return (
    <svg
      {...props}
      aria-hidden="true"
      data-docs-header-footer-icon={kind}
      height={size}
      viewBox="0 0 32 32"
      width={size}
    >
      <PageFrame
        band={
          kind === "go-header" || kind === "delete-header"
            ? "header"
            : kind === "go-footer" || kind === "delete-footer"
              ? "footer"
              : undefined
        }
      />
      {kind === "page-number" ? (
        <path
          d="M12 15h8m-7 5h8m-6.2-7-1.6 10m5-10-1.6 10"
          fill="none"
          stroke={BLUE}
          strokeWidth="1.6"
          strokeLinecap="round"
        />
      ) : null}
      {kind === "go-header" ? <DirectionBadge direction="up" /> : null}
      {kind === "go-footer" ? <DirectionBadge direction="down" /> : null}
      {kind === "previous" ? <DirectionBadge compact direction="left" /> : null}
      {kind === "next" ? <DirectionBadge compact direction="right" /> : null}
      {kind === "delete-header" || kind === "delete-footer" ? (
        <DeleteBadge />
      ) : null}
      {kind === "date-time" ? (
        // A small calendar leaf with a clock hand: Word's Date & Time.
        <>
          <rect x="11" y="13" width="11" height="10" rx="1.5" fill={PAPER} stroke={BLUE} strokeWidth="1.4" />
          <path d="M11 16.5h11M14 12v2.5M19 12v2.5" stroke={BLUE} strokeWidth="1.4" strokeLinecap="round" />
          <circle cx="22.5" cy="23.5" r="4" fill={PAPER} stroke={ORANGE} strokeWidth="1.4" />
          <path d="M22.5 21.5v2.2l1.5 1" stroke={ORANGE} strokeWidth="1.3" strokeLinecap="round" fill="none" />
        </>
      ) : null}
      {kind === "document-info" ? (
        // Text lines with an information badge: Word's Document Info.
        <>
          <path d="M11 14h9M11 18h9M11 22h6" stroke={BLUE} strokeWidth="1.5" strokeLinecap="round" />
          <circle cx="22.5" cy="23.5" r="4" fill={BLUE} />
          <path d="M22.5 22.2v3.2" stroke="#fff" strokeWidth="1.4" strokeLinecap="round" />
          <circle cx="22.5" cy="20.7" r="0.8" fill="#fff" />
        </>
      ) : null}
      {kind === "picture" ? (
        // A framed landscape: Word's Pictures.
        <>
          <rect x="10.5" y="13.5" width="12" height="10" rx="1" fill={PAPER} stroke={BLUE} strokeWidth="1.4" />
          <path d="m11.5 22 3.2-3.6 2.4 2.4 2-2.2 2.9 3.4z" fill={BLUE} />
          <circle cx="19.2" cy="16.5" r="1.3" fill={ORANGE} />
        </>
      ) : null}
    </svg>
  );
}
