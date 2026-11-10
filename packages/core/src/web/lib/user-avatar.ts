/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Shared vocabulary for a person's avatar: what the profile API sends
 * (`avatar_kind` / `avatar_doodle` / `avatar_url` / `avatar_url_small`) and
 * what every surface that draws a person (sidebar, member lists, presence,
 * cursors, profile card) hands to `UserAvatar`.
 *
 * Kept free of React so SDK mappers (collaboration-sdk, collab-state) can
 * import it without pulling a component into their module graph.
 */

/**
 * Fired on `window` after the signed-in user changes their own profile
 * (detail: the new `/auth/me` record). Surfaces that show "me" — the console
 * sidebar, the marketing header — redraw from it instead of refetching.
 */
export const PROFILE_UPDATED_EVENT = "mrnote:profile-updated";

/**
 * The doodles a person may wear instead of a photo, in the picker's groups.
 * The workspace group is the product's own sidebar glyphs, tinted with the
 * person's identity colour; the other three are drawings on their own
 * colours. The API keeps the same list (`AvatarDoodleValue` in
 * apps/api/app/schemas/auth.py): a name must be added there first.
 */
export const AVATAR_DOODLE_GROUPS = {
  workspace: [
    "pages",
    "assistant",
    "search",
    "notebooks",
    "messages",
    "memory",
    "memory-graph",
    "calendar",
  ],
  animals: [
    "cat",
    "dog",
    "rabbit",
    "panda",
    "bird",
    "fish",
    "turtle",
    "squirrel",
    "snail",
    "shell",
    "bug",
    "paw-print",
  ],
  nature: [
    "sun",
    "moon",
    "cloud",
    "rainbow",
    "snowflake",
    "flower",
    "leaf",
    "clover",
    "tree-pine",
    "mountain",
    "waves",
    "flame",
  ],
  hobbies: [
    "coffee",
    "rocket",
    "palette",
    "music",
    "guitar",
    "camera",
    "headphones",
    "gamepad",
    "telescope",
    "sailboat",
    "tent",
    "ice-cream",
  ],
} as const;

export type AvatarDoodleGroup = keyof typeof AVATAR_DOODLE_GROUPS;

export const AVATAR_DOODLES = [
  ...AVATAR_DOODLE_GROUPS.workspace,
  ...AVATAR_DOODLE_GROUPS.animals,
  ...AVATAR_DOODLE_GROUPS.nature,
  ...AVATAR_DOODLE_GROUPS.hobbies,
] as const;

export type AvatarDoodle = (typeof AVATAR_DOODLES)[number];
export type ProductAvatarDoodle =
  (typeof AVATAR_DOODLE_GROUPS.workspace)[number];

export function isProductAvatarDoodle(
  value: string,
): value is ProductAvatarDoodle {
  return (AVATAR_DOODLE_GROUPS.workspace as readonly string[]).includes(value);
}

export type UserAvatarKind = "photo" | "doodle" | null;

export interface UserAvatarSource {
  kind: UserAvatarKind;
  /** Photo URL at the size the surface needs (64px or 256px rendition). */
  url?: string | null;
  doodle?: string | null;
}

/** The avatar fields as the API sends them; every field is optional because
 * older records (and older API builds) do not carry them yet. */
export interface UserAvatarWire {
  avatar_kind?: string | null;
  avatar_doodle?: string | null;
  avatar_url?: string | null;
  avatar_url_small?: string | null;
}

export function isAvatarDoodle(value: unknown): value is AvatarDoodle {
  return (
    typeof value === "string" &&
    (AVATAR_DOODLES as readonly string[]).includes(value)
  );
}

/**
 * Only same-origin paths and http(s) URLs may become an `<img src>`: the URL
 * comes from the server, but a record written by an older client or a proxy
 * must never smuggle a `javascript:` or `data:` URL into the page.
 */
export function safeAvatarUrl(value: unknown): string | null {
  if (typeof value !== "string") return null;
  const trimmed = value.trim();
  if (!trimmed) return null;
  if (trimmed.startsWith("/") && !trimmed.startsWith("//")) return trimmed;
  return /^https?:\/\//i.test(trimmed) ? trimmed : null;
}

/**
 * Map the wire fields to an avatar source. `size` picks the rendition: the
 * 64px `avatar_url_small` for everything up to 40px, the 256px `avatar_url`
 * above that; each falls back to the other when only one is present.
 */
export function avatarFromWire(
  wire: UserAvatarWire | null | undefined,
  size: "small" | "large" = "small",
): UserAvatarSource | null {
  if (!wire) return null;
  const kind =
    wire.avatar_kind === "photo" || wire.avatar_kind === "doodle"
      ? wire.avatar_kind
      : null;
  const small = safeAvatarUrl(wire.avatar_url_small);
  const large = safeAvatarUrl(wire.avatar_url);
  const url = size === "small" ? (small ?? large) : (large ?? small);
  const doodle = isAvatarDoodle(wire.avatar_doodle) ? wire.avatar_doodle : null;
  if (kind === null && !url && !doodle) return null;
  return { kind, url, doodle };
}

/** Strict `#rgb` / `#rrggbb`; anything else is not a colour we paint with. */
export function normalizeAccentHex(value: unknown): string | null {
  if (typeof value !== "string") return null;
  const normalized = value.trim().toLowerCase();
  if (!/^#([0-9a-f]{3}|[0-9a-f]{6})$/.test(normalized)) return null;
  if (normalized.length === 4) {
    const [, r, g, b] = normalized;
    return `#${r}${r}${g}${g}${b}${b}`;
  }
  return normalized;
}

/**
 * The letters an initials avatar shows. Up to two initials (first letters of
 * the first two words) when `max` is 2 and the name has two words; otherwise
 * the first character. Falls back to the email's local part, then "?".
 */
export function avatarInitials(
  name: string | null | undefined,
  email?: string | null,
  max: 1 | 2 = 1,
): string {
  const source =
    (name ?? "").trim() || (email ?? "").trim().split("@", 1)[0] || "";
  if (!source) return "?";
  const words = source.split(/\s+/).filter(Boolean);
  if (max === 2 && words.length >= 2) {
    return words
      .slice(0, 2)
      .map((word) => Array.from(word)[0] ?? "")
      .join("")
      .toUpperCase();
  }
  return (Array.from(source)[0] ?? "?").toUpperCase();
}
