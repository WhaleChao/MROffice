"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import {
  ApiRequestError,
  apiDelete,
  apiGet,
  apiPatch,
  apiPost,
  personalWorkspaceId,
} from "./api";
import { registerNotebookMountGate } from "./notebook-mount-gate";
import {
  collaborationRequestAbortError,
  runCollaborationRequestWithTimeout,
} from "./collaboration-windows/request-timeout";
import {
  mountWorkspaceFor,
  reconcileSharedNotebookMounts,
} from "./workspace-mounts";
import { subscribeAuthState } from "./auth-state";
import type { ActivitySegment } from "@/lib/workspace-activity";
import { avatarFromWire, type UserAvatarSource } from "./user-avatar";

/**
 * Client for the notebook collaboration surface: membership management and
 * the uploaded activity-segment feed. Server roles are viewer/editor/manager;
 * the members dialog speaks read/edit/manage — mapped here so neither side
 * bends.
 */

export type NotebookServerRole = "viewer" | "editor" | "manager";
export type NotebookPermission = "read" | "edit" | "manage";

const ROLE_TO_PERMISSION: Record<NotebookServerRole, NotebookPermission> = {
  viewer: "read",
  editor: "edit",
  manager: "manage",
};

const PERMISSION_TO_ROLE: Record<NotebookPermission, NotebookServerRole> = {
  read: "viewer",
  edit: "editor",
  manage: "manager",
};

export function roleToPermission(role: string): NotebookPermission {
  return ROLE_TO_PERMISSION[role as NotebookServerRole] ?? "read";
}

export function permissionToRole(
  permission: NotebookPermission,
): NotebookServerRole {
  return PERMISSION_TO_ROLE[permission];
}

export interface NotebookMemberRecord {
  id: string;
  userId: string;
  email: string;
  displayName: string | null;
  role: NotebookServerRole;
  isOwner: boolean;
  isSelf: boolean;
  /** The member's chosen identity colour, or null if they never picked one. */
  accentColor: string | null;
  /** Photo (64px rendition) or doodle; null when the member has neither. */
  avatar?: UserAvatarSource | null;
}

interface MemberWire {
  id: string;
  user_id: string;
  email: string;
  display_name: string | null;
  role: string;
  is_owner: boolean;
  is_self: boolean;
  accent_color?: string | null;
  avatar_kind?: string | null;
  avatar_doodle?: string | null;
  avatar_url_small?: string | null;
}

function memberFromWire(wire: MemberWire): NotebookMemberRecord {
  // Present only when the API sent avatar fields, so records from older
  // builds keep their exact shape.
  const avatar = avatarFromWire(wire, "small");
  return {
    id: wire.id,
    userId: wire.user_id,
    email: wire.email,
    displayName: wire.display_name,
    role: (wire.role as NotebookServerRole) ?? "viewer",
    isOwner: wire.is_owner,
    isSelf: wire.is_self,
    accentColor: wire.accent_color ?? null,
    ...(avatar ? { avatar } : {}),
  };
}

function normalizedEmail(value: string): string {
  return value.trim().toLowerCase();
}

/**
 * Cross-workspace targeting for member management: a MOUNTED share (ARCH8)
 * lives in the owner's workspace, so its member calls must carry that
 * workspace explicitly instead of the browser's personal-workspace default.
 * The server treats this only as a locator and rechecks the exact
 * NotebookMember grant; it never becomes workspace membership.
 */
export interface MemberApiOptions {
  workspaceId?: string;
  signal?: AbortSignal;
}

function memberInit(
  options?: MemberApiOptions,
  signal?: AbortSignal,
): RequestInit {
  return {
    ...(options?.workspaceId
      ? { headers: { "X-Workspace-ID": options.workspaceId } }
      : {}),
    ...(signal ? { signal } : {}),
  };
}

async function runMemberRequest<T>(
  options: MemberApiOptions | undefined,
  request: (signal: AbortSignal) => Promise<T>,
  timeoutMessage: string,
): Promise<T> {
  const controller = new AbortController();
  const cancelFromCaller = () => {
    controller.abort(
      options?.signal?.reason ??
        collaborationRequestAbortError("Notebook member request cancelled"),
    );
  };
  if (options?.signal?.aborted) cancelFromCaller();
  else
    options?.signal?.addEventListener("abort", cancelFromCaller, {
      once: true,
    });
  try {
    return await runCollaborationRequestWithTimeout(
      controller,
      request,
      timeoutMessage,
    );
  } finally {
    options?.signal?.removeEventListener("abort", cancelFromCaller);
  }
}

function ambiguousMemberMutation(
  error: unknown,
  options?: MemberApiOptions,
): boolean {
  if (options?.signal?.aborted) return false;
  if (error instanceof DOMException && error.name === "AbortError") return true;
  if (!(error instanceof ApiRequestError)) return true;
  return (
    error.status === 0 ||
    error.status === 408 ||
    error.status === 425 ||
    error.status >= 500
  );
}

function reconciliationOptions(
  options?: MemberApiOptions,
): MemberApiOptions | undefined {
  if (!options) return undefined;
  return {
    ...(options.workspaceId ? { workspaceId: options.workspaceId } : {}),
    ...(options.signal ? { signal: options.signal } : {}),
  };
}

export async function listNotebookMembers(
  notebookId: string,
  options?: MemberApiOptions,
): Promise<NotebookMemberRecord[]> {
  const data = await runMemberRequest(
    options,
    (signal) =>
      apiGet<{ items: MemberWire[] }>(
        `/api/v1/notebooks/${encodeURIComponent(notebookId)}/members`,
        memberInit(options, signal),
      ),
    "Notebook member list timed out",
  );
  return data.items.map(memberFromWire);
}

export async function inviteNotebookMember(
  notebookId: string,
  email: string,
  role: NotebookServerRole,
  options?: MemberApiOptions,
): Promise<NotebookMemberRecord> {
  try {
    const data = await runMemberRequest(
      options,
      (signal) =>
        apiPost<MemberWire>(
          `/api/v1/notebooks/${encodeURIComponent(notebookId)}/members`,
          { email, role },
          memberInit(options, signal),
        ),
      "Notebook member invite timed out",
    );
    const member = memberFromWire(data);
    if (
      normalizedEmail(member.email) !== normalizedEmail(email) ||
      member.role !== role
    ) {
      throw new Error(
        "Notebook member invite response changed identity or role",
      );
    }
    return member;
  } catch (error) {
    if (!ambiguousMemberMutation(error, options)) throw error;
    try {
      const expectedEmail = normalizedEmail(email);
      const members = await listNotebookMembers(
        notebookId,
        reconciliationOptions(options),
      );
      const durable = members.find(
        (member) =>
          normalizedEmail(member.email) === expectedEmail &&
          member.role === role,
      );
      if (durable) return durable;
    } catch {
      // Only a fresh exact member projection may turn ambiguity into success.
    }
    throw error;
  }
}

export async function updateNotebookMemberRole(
  notebookId: string,
  memberId: string,
  role: NotebookServerRole,
  options?: MemberApiOptions,
): Promise<NotebookMemberRecord> {
  try {
    const data = await runMemberRequest(
      options,
      (signal) =>
        apiPatch<MemberWire>(
          `/api/v1/notebooks/${encodeURIComponent(notebookId)}/members/${encodeURIComponent(memberId)}`,
          { role },
          memberInit(options, signal),
        ),
      "Notebook member role update timed out",
    );
    const member = memberFromWire(data);
    if (member.id !== memberId || member.role !== role) {
      throw new Error("Notebook member role response changed identity or role");
    }
    return member;
  } catch (error) {
    if (!ambiguousMemberMutation(error, options)) throw error;
    try {
      const members = await listNotebookMembers(
        notebookId,
        reconciliationOptions(options),
      );
      const durable = members.find(
        (member) => member.id === memberId && member.role === role,
      );
      if (durable) return durable;
    } catch {
      // Preserve the mutation failure without a fresh exact role projection.
    }
    throw error;
  }
}

export async function removeNotebookMember(
  notebookId: string,
  memberId: string,
  options?: MemberApiOptions,
): Promise<void> {
  const path = `/api/v1/notebooks/${encodeURIComponent(notebookId)}/members/${encodeURIComponent(memberId)}`;
  const remove = () =>
    runMemberRequest(
      options,
      (signal) => apiDelete(path, undefined, memberInit(options, signal)),
      "Notebook member removal timed out",
    );
  try {
    await remove();
    return;
  } catch (error) {
    if (!ambiguousMemberMutation(error, options)) throw error;
    try {
      await remove();
      return;
    } catch (replayError) {
      if (replayError instanceof ApiRequestError && replayError.status === 404)
        return;
    }
    try {
      const members = await listNotebookMembers(
        notebookId,
        reconciliationOptions(options),
      );
      if (!members.some((member) => member.id === memberId)) return;
    } catch {
      // A self-removal may make the list unreadable; that is not proof of the
      // exact target outcome, so retain the original ambiguous failure.
    }
    throw error;
  }
}

// ---------------------------------------------------------------------------
// Invite links and pending (e-mail bound) invites
// ---------------------------------------------------------------------------





export type NotebookInviteRole = "viewer" | "editor" | "manager";

/** Any other role on the wire is a malformed invite, refused as such. */
function isNotebookInviteRole(value: unknown): value is NotebookInviteRole {
  return value === "viewer" || value === "editor" || value === "manager";
}
export const NOTEBOOK_INVITE_EXPIRY_DAYS = [1, 7, 30] as const;
export type NotebookInviteExpiryDays =
  (typeof NOTEBOOK_INVITE_EXPIRY_DAYS)[number];

export interface NotebookInviteRecord {
  id: string;
  kind: "link" | "email";
  role: NotebookInviteRole;
  /** Only for pending invites: the address the invite is bound to. */
  email: string | null;
  expiresAt: string;
  maxUses: number | null;
  useCount: number;
  createdAt: string;
  createdByName: string | null;
}

/**
 * What became of a new pending invite's e-mail: the SMTP server took it,
 * it could not be sent, or collaboration mail is switched off.
 */
export type NotebookInviteEmailDelivery = "sent" | "failed" | "disabled";

/** The token is returned once, when an invite is created or reissued. */
export interface IssuedNotebookInvite extends NotebookInviteRecord {
  token: string;
  /** Locale-free app path of the accept page, e.g. `/invite/<token>`. */
  path: string;
  /** Only when a pending invite was just created. */
  emailDelivery?: NotebookInviteEmailDelivery;
}

interface InviteWire {
  id?: unknown;
  kind?: unknown;
  role?: unknown;
  email?: unknown;
  expires_at?: unknown;
  max_uses?: unknown;
  use_count?: unknown;
  created_at?: unknown;
  created_by?: { display_name?: unknown } | null;
  token?: unknown;
  path?: unknown;
  email_delivery?: unknown;
}

function inviteFromWire(wire: InviteWire): NotebookInviteRecord {
  if (
    typeof wire.id !== "string" ||
    (wire.kind !== "link" && wire.kind !== "email") ||
    !isNotebookInviteRole(wire.role) ||
    typeof wire.expires_at !== "string"
  ) {
    throw new Error("Malformed notebook invite");
  }
  return {
    id: wire.id,
    kind: wire.kind,
    role: wire.role,
    email: typeof wire.email === "string" ? wire.email : null,
    expiresAt: wire.expires_at,
    maxUses: typeof wire.max_uses === "number" ? wire.max_uses : null,
    useCount: typeof wire.use_count === "number" ? wire.use_count : 0,
    createdAt: typeof wire.created_at === "string" ? wire.created_at : "",
    createdByName:
      typeof wire.created_by?.display_name === "string"
        ? wire.created_by.display_name
        : null,
  };
}

function issuedInviteFromWire(wire: InviteWire): IssuedNotebookInvite {
  const invite = inviteFromWire(wire);
  if (
    typeof wire.token !== "string" ||
    typeof wire.path !== "string" ||
    !wire.path.startsWith("/invite/")
  ) {
    throw new Error("Malformed notebook invite token");
  }
  const delivery = wire.email_delivery;
  return {
    ...invite,
    token: wire.token,
    path: wire.path,
    ...(delivery === "sent" || delivery === "failed" || delivery === "disabled"
      ? { emailDelivery: delivery }
      : {}),
  };
}

function invitesPath(notebookId: string, inviteId?: string): string {
  const base = `/api/v1/notebooks/${encodeURIComponent(notebookId)}/invites`;
  return inviteId ? `${base}/${encodeURIComponent(inviteId)}` : base;
}

export async function listNotebookInvites(
  notebookId: string,
  options?: MemberApiOptions,
): Promise<NotebookInviteRecord[]> {
  const data = await runMemberRequest(
    options,
    (signal) =>
      apiGet<{ items: InviteWire[] }>(
        invitesPath(notebookId),
        memberInit(options, signal),
      ),
    "Notebook invite list timed out",
  );
  return (Array.isArray(data.items) ? data.items : []).map(inviteFromWire);
}

export async function createNotebookInviteLink(
  notebookId: string,
  settings: {
    role: NotebookInviteRole;
    expiresInDays: NotebookInviteExpiryDays;
    maxUses: number | null;
  },
  options?: MemberApiOptions,
): Promise<IssuedNotebookInvite> {
  const data = await runMemberRequest(
    options,
    (signal) =>
      apiPost<InviteWire>(
        invitesPath(notebookId),
        {
          kind: "link",
          role: settings.role,
          expires_in_days: settings.expiresInDays,
          max_uses: settings.maxUses,
        },
        memberInit(options, signal),
      ),
    "Notebook invite link timed out",
  );
  return issuedInviteFromWire(data);
}

/**
 * A pending invite for an address that has no account yet. `locale` is the
 * inviter's UI language, which the invitation e-mail is written in.
 */
export async function createNotebookEmailInvite(
  notebookId: string,
  email: string,
  role: NotebookInviteRole,
  options?: MemberApiOptions,
  locale?: string,
): Promise<IssuedNotebookInvite> {
  const data = await runMemberRequest(
    options,
    (signal) =>
      apiPost<InviteWire>(
        invitesPath(notebookId),
        { kind: "email", email, role, ...(locale ? { locale } : {}) },
        memberInit(options, signal),
      ),
    "Notebook invite timed out",
  );
  return issuedInviteFromWire(data);
}

/** A fresh link for an active invite; the previous link stops working. */
export async function reissueNotebookInvite(
  notebookId: string,
  inviteId: string,
  options?: MemberApiOptions,
): Promise<IssuedNotebookInvite> {
  const data = await runMemberRequest(
    options,
    (signal) =>
      apiPost<InviteWire>(
        `${invitesPath(notebookId, inviteId)}/reissue`,
        undefined,
        memberInit(options, signal),
      ),
    "Notebook invite reissue timed out",
  );
  return issuedInviteFromWire(data);
}

export async function revokeNotebookInvite(
  notebookId: string,
  inviteId: string,
  options?: MemberApiOptions,
): Promise<void> {
  const revoke = () =>
    runMemberRequest(
      options,
      (signal) =>
        apiDelete(
          invitesPath(notebookId, inviteId),
          undefined,
          memberInit(options, signal),
        ),
      "Notebook invite removal timed out",
    );
  try {
    await revoke();
  } catch (error) {
    // Removal is idempotent server-side, so an ambiguous outcome is replayed.
    if (!ambiguousMemberMutation(error, options)) throw error;
    await revoke();
  }
}

export interface NotebookInvitePreview {
  /** `member`: the person already has at least this access. */
  status: "active" | "member";
  kind: "link" | "email";
  role: NotebookInviteRole;
  notebookTitle: string;
  inviterName: string | null;
  /** The inviter's own photo or doodle; null draws their initials. */
  inviterAvatar?: UserAvatarSource | null;
  inviterAccent?: string | null;
  expiresAt: string;
}

export interface NotebookInviteAcceptance {
  notebookId: string;
  workspaceId: string;
  role: NotebookServerRole;
  joined: boolean;
}

export async function previewNotebookInvite(
  token: string,
): Promise<NotebookInvitePreview> {
  const data = await apiPost<Record<string, unknown>>(
    "/api/v1/invites/preview",
    { token },
  );
  if (
    (data.status !== "active" && data.status !== "member") ||
    (data.kind !== "link" && data.kind !== "email") ||
    !isNotebookInviteRole(data.role)
  ) {
    throw new Error("Malformed invite preview");
  }
  return {
    status: data.status,
    kind: data.kind,
    role: data.role,
    notebookTitle:
      typeof data.notebook_title === "string" ? data.notebook_title : "",
    inviterName:
      typeof data.inviter_name === "string" ? data.inviter_name : null,
    inviterAvatar: avatarFromWire({
      avatar_kind: optionalString(data.inviter_avatar_kind),
      avatar_doodle: optionalString(data.inviter_avatar_doodle),
      avatar_url_small: optionalString(data.inviter_avatar_url_small),
    }),
    inviterAccent: optionalString(data.inviter_accent_color),
    expiresAt: typeof data.expires_at === "string" ? data.expires_at : "",
  };
}

export async function acceptNotebookInvite(
  token: string,
): Promise<NotebookInviteAcceptance> {
  const data = await apiPost<Record<string, unknown>>(
    "/api/v1/invites/accept",
    { token },
  );
  const role = data.role;
  if (
    typeof data.notebook_id !== "string" ||
    typeof data.workspace_id !== "string" ||
    (role !== "viewer" && role !== "editor" && role !== "manager")
  ) {
    throw new Error("Malformed invite acceptance");
  }
  return {
    notebookId: data.notebook_id,
    workspaceId: data.workspace_id,
    role,
    joined: data.joined === true,
  };
}

/** The shareable URL of an invite path on this site. */
export function notebookInviteUrl(path: string): string {
  if (typeof window === "undefined") return path;
  return `${window.location.origin}${path}`;
}

// ---------------------------------------------------------------------------
// Activity segment upload + shared feed
// ---------------------------------------------------------------------------

export interface RemoteActivitySegment extends ActivitySegment {
  actor: {
    userId: string;
    displayName: string;
    isSelf: boolean;
    accentColor: string | null;
  };
}

interface SegmentWire {
  date: string;
  start_second: number;
  end_second: number;
  lane_key: string;
  context: string;
  notebook_id: string | null;
  window_type: string | null;
  window_title: string | null;
  hint: string | null;
  clicks: number;
  scrolls: number;
  keys: number;
  typing: number;
  actor: {
    user_id: string;
    display_name: string;
    is_self: boolean;
    accent_color?: string | null;
  };
}

const TEXT_LIMIT = 160;




export function laneKeyOf(segment: ActivitySegment): string {
  const win = segment.windowId
    ? `${segment.windowType ?? "window"}#${segment.windowId}`
    : "shell";
  return `${segment.context}|${segment.notebookId ?? ""}|${win}`.slice(
    0,
    TEXT_LIMIT,
  );
}

function segmentToWire(segment: ActivitySegment) {
  return {
    start_second: segment.startSecond,
    end_second: segment.endSecond,
    lane_key: laneKeyOf(segment),
    context: segment.context,
    notebook_id: segment.notebookId,
    window_type: segment.windowType,
    window_title: segment.windowTitle?.slice(0, TEXT_LIMIT) ?? null,
    hint: segment.hint?.slice(0, TEXT_LIMIT) ?? null,
    clicks: segment.clicks,
    scrolls: segment.scrolls,
    keys: segment.keys,
    typing: segment.typing,
  };
}

/**
 * Uploads one finished day, grouped by the tenant each segment was CAPTURED
 * in — activity done inside a shared workspace lands in that workspace's
 * feed even when the uploader has since switched back to their own.
 * Groups whose workspace now rejects the caller (403 forbidden — access
 * revoked) are dropped for good; any other failure rethrows so the caller
 * keeps the day unmarked and retries later.
 */
export async function uploadActivityDay(
  date: string,
  segments: readonly ActivitySegment[],
): Promise<void> {
  const groups = new Map<string, ActivitySegment[]>();
  for (const segment of segments.slice(0, 500)) {
    // "" = legacy/unknown capture tenant → the active workspace, as before.
    const key = segment.workspaceId ?? "";
    const group = groups.get(key);
    if (group) group.push(segment);
    else groups.set(key, [segment]);
  }
  let retriableFailure: unknown = null;
  for (const [workspaceId, group] of groups) {
    try {
      const options = workspaceId ? { workspaceId } : undefined;
      await runMemberRequest(
        options,
        (signal) =>
          apiPost(
            "/api/v1/activity/segments",
            { date, segments: group.map(segmentToWire) },
            memberInit(options, signal),
          ),
        "Activity segment upload timed out",
      );
    } catch (error) {
      const status = (error as { status?: number }).status;
      const code = (error as { code?: string }).code;
      if (status === 403 && code === "forbidden") {
        // Membership in that workspace is gone — the data has no home.
        continue;
      }
      retriableFailure = error;
    }
  }
  if (retriableFailure) {
    throw retriableFailure;
  }
}

export async function fetchNotebookActivityFeed(
  notebookId: string,
  dateFrom: string,
  dateTo: string,
): Promise<RemoteActivitySegment[]> {
  const params = new URLSearchParams({
    date_from: dateFrom,
    date_to: dateTo,
    notebook_id: notebookId,
  });
  const data = await runMemberRequest(
    undefined,
    (signal) =>
      apiGet<{ items: SegmentWire[] }>(
        `/api/v1/activity/segments?${params.toString()}`,
        { signal },
      ),
    "Notebook activity feed timed out",
  );
  return data.items.map((wire) => ({
    date: wire.date,
    startSecond: wire.start_second,
    endSecond: wire.end_second,
    context: (wire.context as ActivitySegment["context"]) ?? "workspace",
    notebookId: wire.notebook_id,
    hint: wire.hint,
    windowType: wire.window_type,
    windowId: null,
    windowTitle: wire.window_title,
    clicks: wire.clicks,
    scrolls: wire.scrolls,
    keys: wire.keys,
    typing: wire.typing,
    live: false,
    actor: {
      userId: wire.actor.user_id,
      displayName: wire.actor.display_name,
      isSelf: wire.actor.is_self,
      accentColor: wire.actor.accent_color ?? null,
    },
  }));
}

// ---------------------------------------------------------------------------
// Mounted shares (ARCH8) + workspace lifecycle (FEAT6)
// ---------------------------------------------------------------------------

export interface SharedWithMeNotebook {
  id: string;
  membershipId: string;
  title: string;
  workspaceId: string;
  myRole: NotebookServerRole;
  ownerName: string;
  ownerAccent: string | null;
  /** The owner's own avatar (photo, doodle or initials), for the cover. */
  ownerAvatar: UserAvatarSource | null;
   
  ownerId: string | null;
  

  memberCount: number | null;
  pageCount: number | null;
  /** The notebook's OWN cover — a shared notebook keeps its color for
   *  everyone; the owner's accent is attribution chrome, not the book. */
  coverPreference: unknown | null;
  notebookType: string;
  sharedAt: string;
}

export interface SharedNotebooksView {
  viewerId: string;
  sharedWithMe: SharedWithMeNotebook[];
  /** My notebooks (current workspace) that have members, with counts. */
  sharedByMe: { id: string; memberCount: number }[];
}

const sharedNotebooksInFlight = new Map<
  string,
  { promise: Promise<SharedNotebooksView>; controller: AbortController }
>();

type SharedNotebookWire = {
  id: string;
  membership_id: string;
  title: string;
  workspace_id: string;
  my_role: NotebookServerRole;
  owner_name: string;
  owner_accent: string | null;
  cover_preference: unknown | null;
  notebook_type: string | null;
  shared_at: string;
  owner_id?: string | null;
  owner_avatar_kind?: string | null;
  owner_avatar_doodle?: string | null;
  owner_avatar_url_small?: string | null;
  member_count?: number | null;
  page_count?: number | null;
};

/** Keys a newer server adds to a shared notebook; an older one omits them. */
const SHARED_NOTEBOOK_OPTIONAL_KEYS = [
  "owner_id",
  "owner_avatar_kind",
  "owner_avatar_doodle",
  "owner_avatar_url_small",
  "member_count",
  "page_count",
] as const;

type SharedByMeWire = { id: string; member_count: number };

function isPlainRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function hasExactKeys(
  value: Record<string, unknown>,
  expected: readonly string[],
): boolean {
  const actual = Object.keys(value).sort();
  const required = [...expected].sort();
  return (
    actual.length === required.length &&
    actual.every((key, index) => key === required[index])
  );
}

function isNonEmptyString(value: unknown): value is string {
  return typeof value === "string" && value.length > 0;
}

function optionalString(value: unknown): string | null {
  return typeof value === "string" && value ? value : null;
}

function optionalCount(value: unknown): number | null {
  return typeof value === "number" && Number.isFinite(value) && value >= 0
    ? Math.floor(value)
    : null;
}

function parseSharedNotebookWire(value: unknown): SharedNotebookWire {
  if (
    !isPlainRecord(value) ||
    !hasExactKeys(
      Object.fromEntries(
        Object.entries(value).filter(
          ([key]) =>
            !(SHARED_NOTEBOOK_OPTIONAL_KEYS as readonly string[]).includes(key),
        ),
      ),
      [
        "cover_preference",
        "id",
        "membership_id",
        "my_role",
        "notebook_type",
        "owner_accent",
        "owner_name",
        "shared_at",
        "title",
        "workspace_id",
      ],
    ) ||
    !isNonEmptyString(value.id) ||
    !isNonEmptyString(value.membership_id) ||
    typeof value.title !== "string" ||
    !isNonEmptyString(value.workspace_id) ||
    (value.my_role !== "viewer" &&
      value.my_role !== "editor" &&
      value.my_role !== "manager") ||
    typeof value.owner_name !== "string" ||
    (value.owner_accent !== null && typeof value.owner_accent !== "string") ||
    (value.notebook_type !== null && typeof value.notebook_type !== "string") ||
    !isNonEmptyString(value.shared_at)
  ) {
    throw new TypeError("shared notebook response item is malformed");
  }
  return {
    id: value.id,
    membership_id: value.membership_id,
    title: value.title,
    workspace_id: value.workspace_id,
    my_role: value.my_role,
    owner_name: value.owner_name,
    owner_accent: value.owner_accent,
    cover_preference: value.cover_preference,
    notebook_type: value.notebook_type,
    shared_at: value.shared_at,
    owner_id: optionalString(value.owner_id),
    owner_avatar_kind: optionalString(value.owner_avatar_kind),
    owner_avatar_doodle: optionalString(value.owner_avatar_doodle),
    owner_avatar_url_small: optionalString(value.owner_avatar_url_small),
    member_count: optionalCount(value.member_count),
    page_count: optionalCount(value.page_count),
  };
}

function parseSharedByMeWire(value: unknown): SharedByMeWire {
  if (
    !isPlainRecord(value) ||
    !hasExactKeys(value, ["id", "member_count"]) ||
    !isNonEmptyString(value.id) ||
    typeof value.member_count !== "number" ||
    !Number.isSafeInteger(value.member_count) ||
    value.member_count < 0
  ) {
    throw new TypeError("shared notebook owner response item is malformed");
  }
  return { id: value.id, member_count: value.member_count };
}

function parseSharedNotebooksWire(value: unknown): {
  viewerId: string;
  sharedWithMe: SharedNotebookWire[];
  sharedByMe: SharedByMeWire[];
} {
  if (
    !isPlainRecord(value) ||
    !hasExactKeys(value, ["viewer_id", "shared_with_me", "shared_by_me"]) ||
    !isNonEmptyString(value.viewer_id) ||
    !Array.isArray(value.shared_with_me) ||
    !Array.isArray(value.shared_by_me)
  ) {
    throw new TypeError("shared notebooks response is malformed");
  }
  return {
    viewerId: value.viewer_id,
    sharedWithMe: value.shared_with_me.map(parseSharedNotebookWire),
    sharedByMe: value.shared_by_me.map(parseSharedByMeWire),
  };
}

async function fetchSharedNotebooksFresh(
  signal: AbortSignal,
): Promise<SharedNotebooksView> {
  const raw = await runMemberRequest(
    { signal },
    (signal) => apiGet<unknown>("/api/v1/shared-notebooks", { signal }),
    "Shared notebook list timed out",
  );
  if (signal.aborted) {
    throw (
      signal.reason ??
      collaborationRequestAbortError("Shared notebook account changed")
    );
  }
  const data = parseSharedNotebooksWire(raw);
  const view: SharedNotebooksView = {
    viewerId: data.viewerId,
    sharedWithMe: data.sharedWithMe.map((wire) => ({
      id: wire.id,
      membershipId: wire.membership_id,
      title: wire.title,
      workspaceId: wire.workspace_id,
      myRole: wire.my_role,
      ownerName: wire.owner_name,
      ownerAccent: wire.owner_accent,
      ownerAvatar: avatarFromWire({
        avatar_kind: wire.owner_avatar_kind ?? null,
        avatar_doodle: wire.owner_avatar_doodle ?? null,
        avatar_url_small: wire.owner_avatar_url_small ?? null,
      }),
      ownerId: wire.owner_id ?? null,
      memberCount: wire.member_count ?? null,
      pageCount: wire.page_count ?? null,
      coverPreference: wire.cover_preference ?? null,
      notebookType: wire.notebook_type || "personal",
      sharedAt: wire.shared_at,
    })),
    sharedByMe: data.sharedByMe.map((wire) => ({
      id: wire.id,
      memberCount: wire.member_count,
    })),
  };
  // `shared_with_me` is a complete cross-workspace snapshot for this user.
  reconcileSharedNotebookMounts(
    view.sharedWithMe.map((item) => ({
      notebookId: item.id,
      workspaceId: item.workspaceId,
      role: item.myRole,
    })),
  );
  return view;
}

/** Workspaces whose complete shared-notebook list this page has received. */
const sharedNotebookListsLoaded = new Set<string>();

/** Whether this page has loaded the complete shared-notebook list for the
 * current workspace, i.e. whether the mount map reflects every share. */
export function sharedNotebookListLoaded(): boolean {
  return sharedNotebookListsLoaded.has(personalWorkspaceId() ?? "<server-default>");
}

export function fetchSharedNotebooks(): Promise<SharedNotebooksView> {
  // Coalesce concurrent readers so a slower stale response cannot resurrect a
  // removed grant after a newer role snapshot has already reached this tab.
  const workspaceKey = personalWorkspaceId() ?? "<server-default>";
  const pending = sharedNotebooksInFlight.get(workspaceKey);
  if (pending) return pending.promise;
  const controller = new AbortController();
  const retireFlight = (): void => {
    if (sharedNotebooksInFlight.get(workspaceKey)?.controller === controller) {
      sharedNotebooksInFlight.delete(workspaceKey);
    }
  };
  const unsubscribe = subscribeAuthState(() => {
    // A workspace locator is not an account identity. Retire the request
    // synchronously so a new account never reuses its promise or mounts.
    retireFlight();
    controller.abort(
      collaborationRequestAbortError("Shared notebook account changed"),
    );
  });
  const request = fetchSharedNotebooksFresh(controller.signal);
  if (!controller.signal.aborted) {
    sharedNotebooksInFlight.set(workspaceKey, { promise: request, controller });
  }
  const settle = (): void => {
    unsubscribe();
    retireFlight();
  };
  void request.then(() => {
    if (!controller.signal.aborted) sharedNotebookListsLoaded.add(workspaceKey);
    settle();
  }, settle);
  return request;
}

/**
 * Wait until a notebook-scoped request for `notebookId` is addressed to the
 * right workspace.
 *
 * A notebook shared with the viewer lives in its owner's workspace, and the
 * request's workspace header comes from the mount map. Until this page has
 * the complete shared list, such a notebook is unknown and would be read in
 * the viewer's own workspace (404). This resolves at once when the list is
 * loaded or the notebook is already mounted; otherwise it joins the coalesced
 * list read (normally already in flight from the console's mount sync), so an
 * own notebook adds no serial round trip of its own. A failed list read does
 * not block the caller. Resolves `true` when it had to wait for the list.
 */
export function awaitNotebookMountReady(notebookId: string): Promise<boolean> {
  const id = notebookId.trim();
  if (
    !id ||
    id === "guest" ||
    sharedNotebookListLoaded() ||
    mountWorkspaceFor(id)
  ) {
    return Promise.resolve(false);
  }
  return fetchSharedNotebooks().then(
    () => true,
    () => true,
  );
}

// Requests addressed to a notebook wait for this before choosing a workspace.
registerNotebookMountGate(awaitNotebookMountReady);

export async function deleteWorkspace(
  workspaceId: string,
  password: string,
): Promise<void> {
  await runMemberRequest(
    undefined,
    (signal) =>
      apiPost(
        `/api/v1/workspaces/${encodeURIComponent(workspaceId)}/delete`,
        { password },
        { signal },
      ),
    "Workspace deletion timed out",
  );
}
