/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The Design tab's thumbnail of an Office theme or variant: a small 16:9
 * slide drawn here from the theme's own colours and fonts -- the master's
 * ground, a title in the heading face, body text in the body face, one
 * the theme's decoration and the accent 1..6 strip. Everything is MRAI's own
 * drawing; no Microsoft thumbnail or artwork is used. The decoration is the
 * one the slide canvas and an exported package show (the title layout's,
 * lib/presentation-office-theme-artwork); a theme MRAI has not drawn yet
 * gets a simple geometric motif so neighbouring tiles read apart.
 */

import { presentationOfficeThemeArtwork } from "../../../../lib/presentation-office-theme-artwork";
import { themeFontSchemeFamily, type DeckTheme } from "@mrnote/core/web/lib/presentation-theme.ts";

import { PresentationOfficeThemeArtwork } from "./PresentationOfficeThemeArtwork";

type Motif = "none" | "band" | "corner" | "rule" | "frame" | "side" | "dots";

/** One small decoration per theme, drawn by us so the tiles differ. */
const THEME_MOTIF: Readonly<Record<string, Motif>> = {
  office: "none",
  "office-2013": "none",
  atlas: "rule",
  badge: "frame",
  berlin: "side",
  celestial: "dots",
  circuit: "band",
  crop: "corner",
  dividend: "side",
  droplet: "dots",
  facet: "corner",
  frame: "frame",
  gallery: "rule",
  integral: "side",
  ion: "band",
  "ion-boardroom": "band",
  madison: "dots",
  "main-event": "band",
  mesh: "corner",
  organic: "rule",
  parcel: "frame",
  quotable: "rule",
  savon: "frame",
  "vapor-trail": "band",
  wisp: "corner",
  "wood-type": "side",
};

function isDark(color: string): boolean {
  const value = /^#[0-9a-f]{6}$/i.test(color) ? color : "#ffffff";
  const red = Number.parseInt(value.slice(1, 3), 16);
  const green = Number.parseInt(value.slice(3, 5), 16);
  const blue = Number.parseInt(value.slice(5, 7), 16);
  return red * 0.2126 + green * 0.7152 + blue * 0.0722 < 140;
}

function MotifShape({ motif, color }: { motif: Motif; color: string }) {
  switch (motif) {
    case "band":
      return <polygon points="112,0 136,0 88,90 64,90" fill={color} opacity={0.85} />;
    case "corner":
      return <polygon points="160,52 160,90 118,90" fill={color} />;
    case "rule":
      return <rect x={14} y={47} width={58} height={2} fill={color} />;
    case "frame":
      return (
        <rect
          x={5}
          y={5}
          width={150}
          height={70}
          fill="none"
          stroke={color}
          strokeWidth={2}
        />
      );
    case "side":
      return <rect x={0} y={0} width={8} height={80} fill={color} />;
    case "dots":
      return (
        <g fill={color}>
          <circle cx={138} cy={18} r={6} />
          <circle cx={124} cy={30} r={3.5} opacity={0.7} />
          <circle cx={146} cy={36} r={2.5} opacity={0.5} />
        </g>
      );
    default:
      return null;
  }
}

export function PresentationOfficeThemeThumbnail({
  theme,
  themeId,
  variantId: variantIdProp,
  className,
}: {
  /** Deck theme fields of the theme / variant (lib/presentation-theme). */
  theme: DeckTheme;
  /** Office theme id: picks the decoration. */
  themeId: string;
  /** Office variant (`v1`..): picks the variant's own decoration. */
  variantId?: string | null;
  className?: string;
}) {
  const slots = theme.colorSlots;
  const background = theme.background ?? slots?.lt1 ?? "#ffffff";
  const dark = isDark(background);
  const titleColor = dark ? (slots?.lt1 ?? "#ffffff") : (theme.titleColor ?? slots?.dk1 ?? "#000000");
  const bodyColor = dark ? (slots?.lt2 ?? "#e8e8e8") : (theme.bodyTextColor ?? slots?.dk2 ?? "#404040");
  const accents = (theme.accentColors ?? []).slice(0, 6);
  const accent1 = accents[0] ?? titleColor;
  const variantId =
    variantIdProp ??
    (theme.identity?.selectionKind === "office" ? theme.identity.variantId : "v1");
  const hasArtwork =
    presentationOfficeThemeArtwork(themeId, variantId, "title").length > 0;
  const motif = hasArtwork ? "none" : (THEME_MOTIF[themeId] ?? "none");
  return (
    <svg
      aria-hidden
      className={className}
      data-office-theme-thumbnail={themeId}
      viewBox="0 0 160 90"
      preserveAspectRatio="xMidYMid meet"
      style={{ display: "block", height: "100%", width: "100%" }}
    >
      <rect width={160} height={90} fill={background} />
      {hasArtwork ? (
        <PresentationOfficeThemeArtwork
          theme={theme}
          themeId={themeId}
          variantId={variantId}
          layoutId="title"
          nested={{ width: 160, height: 90 }}
        />
      ) : null}
      <MotifShape motif={motif} color={accent1} />
      <text
        data-theme-preview-aa
        x={14}
        y={40}
        fill={titleColor}
        style={{
          fontFamily: themeFontSchemeFamily(theme, true),
          fontSize: 24,
        }}
      >
        Aa
      </text>
      <text
        x={14}
        y={60}
        fill={bodyColor}
        style={{ fontFamily: themeFontSchemeFamily(theme, false), fontSize: 9 }}
      >
        Aa Bb Cc
      </text>
      <rect x={14} y={66} width={46} height={2.5} rx={1} fill={bodyColor} opacity={0.55} />
      <g data-theme-preview-accents>
        {accents.map((accent, index) => (
          <rect
            key={`${index}-${accent}`}
            x={(index * 160) / Math.max(accents.length, 1)}
            y={80}
            width={160 / Math.max(accents.length, 1)}
            height={10}
            fill={accent}
          />
        ))}
      </g>
    </svg>
  );
}
