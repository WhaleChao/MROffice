/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Pure logic for the real-time collaborative editing foundation.
 *
 * This module is deliberately free of any Yjs / TipTap / React / network
 * imports so the security-sensitive normalization (user-name + color
 * sanitization), the provider-kind decision, and the awareness-state mapping
 * can be unit-tested headlessly and reused by both the
 * {@link createCollabSession} provider factory (collab-provider.ts) and the
 * editor presence UI.
 *
 * Security model mirrors the document editor's other untrusted-string paths
 * (table cell shading, inline image alt, text color): the user color is forced
 * through a STRICT hex gate that canonicalizes to `#rrggbb` and rejects anything
 * that is not a single pure hex color, and the display name is length-capped and
 * stripped of control characters. A collaborator's awareness payload is shared
 * peer-to-peer with every other client, so it is untrusted input that must never
 * be allowed to smuggle CSS, markup, or unbounded strings into a remote DOM.
 */

import type { UserAvatarSource } from "../user-avatar";

/**
 * A participant's avatar as the presence UI draws it. It comes from the
 * server's notebook member record, never from a peer's awareness payload:
 * awareness stays `{ id, name, color }`, so no peer can hand another client
 * an image URL to load.
 */
export type CollabParticipantAvatar = UserAvatarSource;

/** Max length for a sanitized collaborator display name. */
export const COLLAB_USER_NAME_MAX_LENGTH = 40;

/** Fallback name used when the supplied name is empty/invalid after sanitizing. */
export const COLLAB_DEFAULT_USER_NAME = "Anonymous";

/**
 * Curated, accessible cursor colors. A deterministic pick from this palette is
 * used when no valid color is supplied, so two clients with the same seed get a
 * stable color and the palette stays within a readable contrast range.
 */
export const COLLAB_CURSOR_PALETTE = [
  "#2563eb", // blue
  "#dc2626", // red
  "#16a34a", // green
  "#9333ea", // purple
  "#ea580c", // orange
  "#0891b2", // cyan
  "#db2777", // pink
  "#ca8a04", // amber
] as const;

/** Strict 3- or 6-digit hex (with leading #). Mirrors the editor color gate. */
const STRICT_HEX_RE = /^#([0-9a-f]{3}|[0-9a-f]{6})$/;

/** Control characters (incl. DEL) that must be stripped from a display name. */
const CONTROL_CHARS_RE = /[\x00-\x1f\x7f]/g;

/** Collapse any run of whitespace (incl. the stripped controls' gaps) to one space. */
const WHITESPACE_RUN_RE = /\s+/g;

/** Max length for a sanitized collaborator user id. */
export const COLLAB_USER_ID_MAX_LENGTH = 64;

/** A fully-validated collaborator identity safe to broadcast via awareness. */
export interface CollabUser {
  








  id: string | null;
  name: string;
  color: string;
}

/**
 * Expand a 3-digit hex (`#abc`) to its canonical 6-digit form (`#aabbcc`) and
 * lowercase a 6-digit hex. Input MUST already match {@link STRICT_HEX_RE}.
 */
function canonicalHex(value: string): string {
  if (value.length === 4) {
    const [, r, g, b] = value;
    return `#${r}${r}${g}${g}${b}${b}`;
  }
  return value;
}

/**
 * Coerce an arbitrary value into a canonical `#rrggbb` cursor color, or `null`
 * when it is not a single pure hex color. Routes through the SAME strict gate as
 * the editor's text/cell color path, so an injection payload like
 * `"#fff;background:url(x)"` or `"red"` is rejected outright.
 */
export function normalizeCollabColor(value: unknown): string | null {
  if (typeof value !== "string") return null;
  const normalized = value.trim().toLowerCase();
  if (!STRICT_HEX_RE.test(normalized)) return null;
  return canonicalHex(normalized);
}

/**
 * Deterministically pick a palette color from a seed string (the user name or
 * id). Same seed → same color, so a collaborator keeps a stable cursor color
 * across reconnects even when no explicit color was supplied. An empty seed
 * yields the first palette entry.
 */
export function pickCollabColor(seed: string): string {
  let hash = 0;
  for (let i = 0; i < seed.length; i += 1) {
    // Simple, stable, non-cryptographic string hash (djb2-ish).
    hash = (hash * 31 + seed.charCodeAt(i)) | 0;
  }
  const index = Math.abs(hash) % COLLAB_CURSOR_PALETTE.length;
  return COLLAB_CURSOR_PALETTE[index];
}

/**
 * Sanitize a possibly-untrusted display name: drop control characters, collapse
 * whitespace, trim, and cap the length. Returns {@link COLLAB_DEFAULT_USER_NAME}
 * when nothing usable remains.
 */
export function sanitizeCollabName(value: unknown): string {
  if (typeof value !== "string") return COLLAB_DEFAULT_USER_NAME;
  const cleaned = value
    .replace(CONTROL_CHARS_RE, " ")
    .replace(WHITESPACE_RUN_RE, " ")
    .trim()
    .slice(0, COLLAB_USER_NAME_MAX_LENGTH)
    // A trailing partial space can survive the slice; trim once more.
    .trim();
  return cleaned || COLLAB_DEFAULT_USER_NAME;
}

/**
 * Sanitize a possibly-untrusted user id. It arrives over the wire from a peer
 * and is only ever compared for equality and used as a map key, so the gate is
 * about bounding it rather than about shape: control characters out, trimmed,
 * length capped. Anything unusable becomes null, which means "cannot identify
 * this person" — never a shared bucket that would merge two of them.
 */
export function sanitizeCollabUserId(value: unknown): string | null {
  if (typeof value !== "string") return null;
  const cleaned = value
    .replace(CONTROL_CHARS_RE, "")
    .trim()
    .slice(0, COLLAB_USER_ID_MAX_LENGTH);
  return cleaned || null;
}

/**
 * Produce a fully-validated {@link CollabUser} from arbitrary input. The name is
 * sanitized + capped; the color is forced through the strict hex gate and, when
 * invalid/absent, a deterministic palette color is picked from the (sanitized)
 * name so the identity is stable and always renderable.
 */
export function sanitizeCollabUser(input: {
  id?: unknown;
  name?: unknown;
  color?: unknown;
}): CollabUser {
  const name = sanitizeCollabName(input?.name);
  const color = normalizeCollabColor(input?.color) ?? pickCollabColor(name);
  return { id: sanitizeCollabUserId(input?.id), name, color };
}

/** Configuration the provider factory uses to decide what kind of session to build. */
export interface CollabProviderConfig {
  /** Stable room/document name shared by all peers editing the same doc. */
  roomName?: string | null;
  /**
   * Optional explicit provider URL. When present (a `ws://`/`wss://` URL) the
   * session attaches a websocket-style provider; otherwise a configured room
   * falls back to serverless WebRTC P2P.
   */
  providerUrl?: string | null;
}

/**
 * The three session shapes the factory can return:
 * - `local`  : no room configured → single-user Y.Doc, no network provider.
 * - `webrtc` : a room (and no usable websocket URL) → serverless P2P.
 * - `websocket` : an explicit ws:// / wss:// URL → server-relayed provider.
 */
export type CollabProviderKind = "local" | "webrtc" | "websocket";

/** Only ws/wss URLs are accepted as a websocket provider endpoint. */
function isWebsocketUrl(value: unknown): value is string {
  if (typeof value !== "string") return false;
  const trimmed = value.trim();
  return /^wss?:\/\/[^\s]+$/i.test(trimmed);
}

/** A room name is usable only when it is a non-empty, trimmable string. */
function hasUsableRoom(value: unknown): value is string {
  return typeof value === "string" && value.trim().length > 0;
}

/**
 * Decide which provider kind a config maps to. Pure + total:
 * - no room → always `local` (single-user; the editor still works).
 * - room + ws/wss URL → `websocket`.
 * - room + no URL (or a non-ws URL) → `webrtc` (serverless P2P).
 */
export function providerKindFor(config: CollabProviderConfig | null | undefined): CollabProviderKind {
  if (!config || !hasUsableRoom(config.roomName)) {
    return "local";
  }
  if (isWebsocketUrl(config.providerUrl)) {
    return "websocket";
  }
  return "webrtc";
}

/** A remote collaborator's presence as surfaced to the editor UI. */
export interface CollabPresenceEntry {
  clientId: number;
  user: CollabUser;
}

/** Raw awareness state map shape (as returned by y-protocols `getStates()`). */
export type AwarenessStates = Map<number, Record<string, unknown>>;

/**
 * Whether a peer's awareness state is the local user's own other connection
 * (another tab, window or device of the same account). Such a caret is still
 * you, and drawing it beside the one you type with read as a stray second
 * cursor carrying your own name. A peer without an id is never taken for you.
 */
export function isLocalUsersOtherConnection(
  states: AwarenessStates | null | undefined,
  localClientId: number,
  state: unknown,
): boolean {
  if (!states) return false;
  const localUserId = sanitizeCollabUserId(
    (states.get(localClientId)?.user as { id?: unknown } | undefined)?.id,
  );
  if (localUserId === null) return false;
  const peer = (state as { user?: { id?: unknown } } | null | undefined)?.user;
  return sanitizeCollabUserId(peer?.id) === localUserId;
}























export function mapAwarenessStates(
  states: AwarenessStates | null | undefined,
  localClientId: number,
): CollabPresenceEntry[] {
  if (!states) return [];
  // Read the local identity from awareness itself — the same value the
  // provider published, so there is nothing to keep in sync.
  const localUserId = sanitizeCollabUserId(
    (states.get(localClientId)?.user as { id?: unknown } | undefined)?.id,
  );
  const candidates: CollabPresenceEntry[] = [];
  states.forEach((state, clientId) => {
    if (clientId === localClientId) return;
    const rawUser = state?.user;
    if (!rawUser || typeof rawUser !== "object") return;
    const user = sanitizeCollabUser(
      rawUser as { id?: unknown; name?: unknown; color?: unknown },
    );
    if (user.id !== null && user.id === localUserId) return;
    candidates.push({ clientId, user });
  });
  // Sort BEFORE collapsing, so the connection that represents a person is
  // their lowest clientId rather than whichever one the awareness map
  // happened to be holding first.
  candidates.sort((a, b) => a.clientId - b.clientId);
  const seen = new Set<string>();
  return candidates.filter((entry) => {
    if (entry.user.id === null) return true;
    if (seen.has(entry.user.id)) return false;
    seen.add(entry.user.id);
    return true;
  });
}

/**
 * Build the local awareness `user` field payload from a sanitized user. Kept as a
 * tiny pure mapper so the provider factory and tests agree on the exact shape
 * written to awareness (`{ user: { name, color } }`).
 */
export function localAwarenessUserState(user: CollabUser): { user: CollabUser } {
  // The id travels with the name and colour: it is what lets every peer tell
  // one person's several connections apart from several people.
  return { user: { id: user.id, name: user.name, color: user.color } };
}
