"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { apiGet, apiPost, apiPut, isApiRequestError } from "./api";
import { avatarFromWire, type UserAvatarSource } from "./user-avatar";
import {
  parseCollaborationSourcePagePolicy,
  type CollaborationSourcePageGrant,
  type CollaborationSourcePagePolicy,
} from "./collaboration-windows/sharing-policy";

/** Page-level collaboration client (FEAT5): visibility, exclusive
 *  checkout, transfer. Mirrors routers/page_collab.py. */

export type PageVisibility = "private" | "shared_readonly" | "shared_editable";

export interface PageEditorMembership {
  userId: string;
  membershipIncarnation: string | null;
}

export type PageAccessPolicy = CollaborationSourcePagePolicy;
export type PageAccessGrant = CollaborationSourcePageGrant;

export interface PageAccessPolicyBinding {
  readonly notebookId: string;
  readonly ownerUserId: string;
}

export interface PageAccessPolicyReplacement {
  readonly visibility: PageVisibility;
  /** Required complete recipient set, not a patch or inherited audience. */
  readonly grants: readonly PageAccessGrant[];
  readonly confirmLegacyConversion: boolean;
}

export interface PageCollabState {
  pageId: string;
  visibility: PageVisibility;
  owner: {
    userId: string;
    name: string;
    accentColor: string | null;
    /** Photo (64px) or doodle from the owner's profile; null without one. */
    avatar?: UserAvatarSource | null;
  };
  isOwner: boolean;
  canEdit: boolean;
  /**
   * Who the OWNER granted write access to, for a `shared_editable` page.
   * Empty is a real answer — published, owner-only edits — not "everyone".
   */
  editorUserIds: string[];
  checkout: {
    userId: string;
    name: string;
    accentColor: string | null;
    isSelf: boolean;
  } | null;
}

const COLLAB_KEYS = new Set([
  "can_edit",
  "checkout",
  "editor_user_ids",
  "is_owner",
  "owner",
  "page_id",
  "visibility",
]);
const OWNER_KEYS = new Set(["accent_color", "name", "user_id"]);

// object. Both spellings are the exact wire shape of one server build; any
// other key set is still a changed binding.
const OWNER_AVATAR_KEYS = new Set([
  ...OWNER_KEYS,
  "avatar_doodle",
  "avatar_kind",
  "avatar_url_small",
]);
const CHECKOUT_KEYS = new Set(["accent_color", "is_self", "name", "user_id"]);
// The same profile API build adds the holder's avatar summary to the
// checkout object (a page someone is editing); either exact shape is one
// server build, any other key set a changed binding.
const CHECKOUT_AVATAR_KEYS = new Set([
  ...CHECKOUT_KEYS,
  "avatar_doodle",
  "avatar_kind",
  "avatar_url_small",
]);

export class InvalidPageCollabResponseError extends Error {}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function exactKeys(
  value: Record<string, unknown>,
  expected: ReadonlySet<string>,
): boolean {
  const keys = Object.keys(value);
  return (
    keys.length === expected.size && keys.every((key) => expected.has(key))
  );
}

function identity(value: unknown, field: string): string {
  if (
    typeof value !== "string" ||
    value.length === 0 ||
    value.length > 256 ||
    value.normalize("NFC") !== value ||
    /\s|\p{C}/u.test(value)
  ) {
    throw new InvalidPageCollabResponseError(`${field} is invalid`);
  }
  return value;
}

function displayText(value: unknown, field: string): string {
  if (typeof value !== "string" || value.length > 512) {
    throw new InvalidPageCollabResponseError(`${field} is invalid`);
  }
  return value;
}

function booleanField(value: unknown, field: string): boolean {
  if (typeof value !== "boolean") {
    throw new InvalidPageCollabResponseError(`${field} is invalid`);
  }
  return value;
}

function accentColor(value: unknown): string | null {
  if (
    value !== null &&
    (typeof value !== "string" ||
      value.length > 128 ||
      value.normalize("NFC") !== value ||
      /\p{C}/u.test(value))
  ) {
    throw new InvalidPageCollabResponseError("accent color is invalid");
  }
  return value as string | null;
}

function optionalText(value: unknown, field: string): string | null {
  if (value === undefined || value === null) return null;
  if (typeof value !== "string" || value.length > 2048 || /\p{C}/u.test(value)) {
    throw new InvalidPageCollabResponseError(`${field} is invalid`);
  }
  return value;
}

/** The owner's avatar summary as the profile API spells it; absent = none. */
function ownerAvatarWire(owner: Record<string, unknown>) {
  return {
    avatar_kind: optionalText(owner.avatar_kind, "owner avatar kind"),
    avatar_doodle: optionalText(owner.avatar_doodle, "owner avatar doodle"),
    avatar_url_small: optionalText(owner.avatar_url_small, "owner avatar url"),
  };
}

function fromWire(wireValue: unknown, expectedPageId: string): PageCollabState {
  if (!isRecord(wireValue) || !exactKeys(wireValue, COLLAB_KEYS)) {
    throw new InvalidPageCollabResponseError(
      "page collaboration response is malformed",
    );
  }
  const pageId = identity(wireValue.page_id, "page id");
  const visibility = wireValue.visibility;
  const ownerValue = wireValue.owner;
  const checkoutValue = wireValue.checkout;
  if (
    pageId !== expectedPageId ||
    (visibility !== "private" &&
      visibility !== "shared_readonly" &&
      visibility !== "shared_editable") ||
    typeof wireValue.is_owner !== "boolean" ||
    typeof wireValue.can_edit !== "boolean" ||
    !isRecord(ownerValue) ||
    !(exactKeys(ownerValue, OWNER_KEYS) || exactKeys(ownerValue, OWNER_AVATAR_KEYS)) ||
    !Array.isArray(wireValue.editor_user_ids) ||
    (checkoutValue !== null &&
      (!isRecord(checkoutValue) ||
        !(exactKeys(checkoutValue, CHECKOUT_KEYS) ||
          exactKeys(checkoutValue, CHECKOUT_AVATAR_KEYS))))
  ) {
    throw new InvalidPageCollabResponseError(
      "page collaboration response changed its page binding",
    );
  }
  const ownerId = identity(ownerValue.user_id, "owner id");
  const editorUserIds = wireValue.editor_user_ids.map((value) =>
    identity(value, "editor id"),
  );
  if (
    new Set(editorUserIds).size !== editorUserIds.length ||
    editorUserIds.includes(ownerId)
  ) {
    throw new InvalidPageCollabResponseError(
      "page collaboration editor list is ambiguous",
    );
  }
  const checkout =
    checkoutValue === null
      ? null
      : {
          userId: identity(checkoutValue.user_id, "checkout user id"),
          name: displayText(checkoutValue.name, "checkout name"),
          accentColor: accentColor(checkoutValue.accent_color),
          isSelf: booleanField(checkoutValue.is_self, "checkout self flag"),
        };
  return {
    pageId,
    visibility,
    owner: {
      userId: ownerId,
      name: displayText(ownerValue.name, "owner name"),
      accentColor: accentColor(ownerValue.accent_color),
      avatar: avatarFromWire(ownerAvatarWire(ownerValue), "small"),
    },
    isOwner: booleanField(wireValue.is_owner, "owner flag"),
    canEdit: booleanField(wireValue.can_edit, "edit flag"),
    editorUserIds,
    checkout,
  };
}

const base = (pageId: string) =>
  `/api/v1/pages/${encodeURIComponent(identity(pageId, "page id"))}`;

function withWorkspaceHeader(
  init: RequestInit | undefined,
  workspaceId: string | null | undefined,
): RequestInit | undefined {
  if (!workspaceId) return init;
  const headers = new Headers(init?.headers);
  headers.set("X-Workspace-ID", identity(workspaceId, "workspace id"));
  return { ...init, headers };
}







export async function getPageCollab(
  pageId: string,
  init?: RequestInit,
  workspaceId?: string | null,
  options?: { admission?: boolean },
): Promise<PageCollabState> {
  const expectedPageId = identity(pageId, "page id");
  return fromWire(
    await apiGet<unknown>(
      `${base(pageId)}/collab${options?.admission ? "?admission=1" : ""}`,
      withWorkspaceHeader(init, workspaceId),
    ),
    expectedPageId,
  );
}

function boundPageAccessPolicy(
  wire: unknown,
  pageId: string,
  binding: PageAccessPolicyBinding,
): PageAccessPolicy {
  const parsed = parseCollaborationSourcePagePolicy(wire);
  if (
    parsed.pageId !== identity(pageId, "page id") ||
    parsed.notebookId !== identity(binding.notebookId, "notebook id") ||
    parsed.ownerUserId !== identity(binding.ownerUserId, "owner id") ||
    parsed.grants.some((grant) => grant.userId === parsed.ownerUserId)
  ) {
    throw new InvalidPageCollabResponseError(
      "file access policy changed its source binding",
    );
  }
  return parsed;
}

/** Owner-only policy read; opening this endpoint never converts a legacy audience. */
export async function getPageAccessPolicy(
  pageId: string,
  binding: PageAccessPolicyBinding,
  init: RequestInit | undefined,
  workspaceId: string,
): Promise<PageAccessPolicy> {
  identity(workspaceId, "workspace id");
  const expectedBinding = Object.freeze({
    notebookId: identity(binding.notebookId, "notebook id"),
    ownerUserId: identity(binding.ownerUserId, "owner id"),
  });
  return boundPageAccessPolicy(
    await apiGet<unknown>(
      `${base(pageId)}/access-policy`,
      withWorkspaceHeader(init, workspaceId),
    ),
    pageId,
    expectedBinding,
  );
}

/** Revision-bound full replacement. This receipt is not a writable/cache verdict. */
export async function replacePageAccessPolicy(
  current: PageAccessPolicy,
  replacement: PageAccessPolicyReplacement,
  init: RequestInit | undefined,
  workspaceId: string,
): Promise<PageAccessPolicy> {
  identity(workspaceId, "workspace id");
  if (
    !isRecord(replacement) ||
    !exactKeys(
      replacement,
      new Set(["visibility", "grants", "confirmLegacyConversion"]),
    ) ||
    typeof replacement.confirmLegacyConversion !== "boolean" ||
    !Array.isArray(replacement.grants) ||
    replacement.grants.some(
      (grant) =>
        !isRecord(grant) ||
        !exactKeys(
          grant,
          new Set(["userId", "grantRole", "membershipIncarnation"]),
        ),
    ) ||
    (current.viewScope !== "notebook_members_legacy" &&
      current.viewScope !== "explicit_grants") ||
    (current.viewScope === "notebook_members_legacy" &&
      !replacement.confirmLegacyConversion)
  ) {
    throw new TypeError("a complete confirmed file access policy is required");
  }
  // Reuse the source parser for role/incarnation/revision/visibility validation.
  // Freeze intent before yielding to CSRF/network work; caller arrays can change.
  const requested = boundPageAccessPolicy(
    {
      page_id: current.pageId,
      notebook_id: current.notebookId,
      owner_user_id: current.ownerUserId,
      visibility: replacement.visibility,
      view_scope: "explicit_grants",
      acl_revision: current.aclRevision,
      grants: replacement.grants.map((grant) => ({
        user_id: grant.userId,
        grant_role: grant.grantRole,
        membership_incarnation: grant.membershipIncarnation,
      })),
      changed: false,
    },
    current.pageId,
    current,
  );
  const response = boundPageAccessPolicy(
    await apiPost<unknown>(
      `${base(requested.pageId)}/access-policy`,
      {
        expected_acl_revision: requested.aclRevision,
        visibility: requested.visibility,
        grants: requested.grants.map((grant) => ({
          user_id: grant.userId,
          grant_role: grant.grantRole,
          membership_incarnation: grant.membershipIncarnation,
        })),
        confirm_legacy_conversion: replacement.confirmLegacyConversion,
      },
      withWorkspaceHeader(init, workspaceId),
    ),
    requested.pageId,
    requested,
  );
  if (
    response.viewScope !== "explicit_grants" ||
    response.aclRevision !== requested.aclRevision + Number(response.changed) ||
    response.visibility !== requested.visibility ||
    response.grants.length !== requested.grants.length ||
    response.grants.some(
      (grant) =>
        !requested.grants.some(
          (candidate) =>
            candidate.userId === grant.userId &&
            candidate.grantRole === grant.grantRole &&
            candidate.membershipIncarnation === grant.membershipIncarnation,
        ),
    )
  ) {
    throw new InvalidPageCollabResponseError(
      "file access policy did not confirm this replacement",
    );
  }
  return response;
}

export async function setPageVisibility(
  pageId: string,
  visibility: PageVisibility,
  /**
   * Who may write it. Omit to leave the current list alone; `[]` publishes
   * the page with the owner as its only editor.
   */
  editorUserIds?: readonly string[],
  editorMemberships?: readonly PageEditorMembership[],
  init?: RequestInit,
): Promise<PageCollabState> {
  const expectedPageId = identity(pageId, "page id");
  return fromWire(
    await apiPost<unknown>(
      `${base(pageId)}/visibility`,
      {
        visibility,
        ...(editorUserIds ? { editor_user_ids: [...editorUserIds] } : {}),
        ...(editorUserIds
          ? {
              editor_memberships: (editorMemberships ?? []).map(
                (membership) => ({
                  user_id: membership.userId,
                  membership_incarnation: membership.membershipIncarnation,
                }),
              ),
            }
          : {}),
      },
      init,
    ),
    expectedPageId,
  );
}

export interface HandOverResult {
  operationId: string;
  /** ONE recipient moves ownership; SEVERAL each receive their own copy. */
  mode: "transferred" | "copied";
  state: PageCollabState;
  copies: { pageId: string; userId: string }[];
}

const CANONICAL_UUID_PATTERN =
  /^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/u;

export function newPageHandOverOperationId(): string {
  const operationId = globalThis.crypto?.randomUUID?.();
  if (!operationId || !CANONICAL_UUID_PATTERN.test(operationId)) {
    throw new Error("secure page hand-over operation identity is unavailable");
  }
  return operationId;
}

/**
 * Give a page away.
 *
 * Ownership is singular, so the two cases are different actions and the API
 * says which one happened: one recipient MOVES the page, several hand each
 * of them their own private copy while the original stays yours.
 */
export async function handOverPage(
  pageId: string,
  toUserIds: readonly string[],
  /**
   * `move` gives the page away (ownership moves, one recipient); `copy`
   * sends each recipient their own and keeps the original. Always send it:
   * the server can infer from the head count for older clients, but making
   * an irreversible choice depend on how many people were ticked is not
   * something a member should have to know.
   */
  mode: "move" | "copy",
  operationId: string,
  init?: RequestInit,
): Promise<HandOverResult> {
  if (!CANONICAL_UUID_PATTERN.test(operationId)) {
    throw new TypeError("page hand-over operation identity is invalid");
  }
  const data = await apiPost<{
    operation_id: string;
    mode: "transferred" | "copied";
    page: unknown;
    copies: { page_id: string; user_id: string }[];
  }>(
    `${base(pageId)}/hand-over`,
    {
      to_user_ids: [...toUserIds],
      mode,
      operation_id: operationId,
    },
    init,
  );
  if (data.operation_id !== operationId) {
    throw new Error("page hand-over response changed operation identity");
  }
  return {
    operationId,
    mode: data.mode,
    state: fromWire(data.page, identity(pageId, "page id")),
    copies: (data.copies ?? []).map((row) => ({
      pageId: row.page_id,
      userId: row.user_id,
    })),
  };
}

export interface CopiedPage {
  id: string;
  notebook_id: string;
  title: string;
  page_type: string;
  visibility: PageVisibility;
}

/** Fork a viewable page into an independent PRIVATE copy owned by the caller. */
export async function copyPage(pageId: string): Promise<CopiedPage> {
  return apiPost<CopiedPage>(`${base(pageId)}/copy`, {});
}

export async function acquirePageCheckout(
  pageId: string,
): Promise<PageCollabState> {
  const expectedPageId = identity(pageId, "page id");
  return fromWire(
    await apiPost<unknown>(`${base(pageId)}/checkout`, {}),
    expectedPageId,
  );
}

export async function heartbeatPageCheckout(
  pageId: string,
  init?: RequestInit,
  workspaceId?: string | null,
): Promise<PageCollabState> {
  const expectedPageId = identity(pageId, "page id");
  return fromWire(
    await apiPost<unknown>(
      `${base(pageId)}/checkout/heartbeat`,
      {},
      withWorkspaceHeader(init, workspaceId),
    ),
    expectedPageId,
  );
}

export async function releasePageCheckout(
  pageId: string,
  init?: RequestInit,
  workspaceId?: string | null,
): Promise<PageCollabState> {
  const expectedPageId = identity(pageId, "page id");
  return fromWire(
    await apiPost<unknown>(
      `${base(pageId)}/checkout/release`,
      {},
      withWorkspaceHeader(init, workspaceId),
    ),
    expectedPageId,
  );
}

export async function transferPage(
  pageId: string,
  toUserId: string,
): Promise<PageCollabState> {
  const expectedPageId = identity(pageId, "page id");
  return fromWire(
    await apiPost<unknown>(`${base(pageId)}/transfer`, {
      to_user_id: toUserId,
    }),
    expectedPageId,
  );
}

export interface WorkbookCollabV2Shard {
  key: string;
  kind: "manifest" | "sheet" | "tile" | "objects";
  projectionRevision: number;
  headSequence: number;
  snapshotSequence: number;
  snapshotPresent: boolean;
}

export interface WorkbookCollabV2State {
  pageId: string;
  eligible: boolean;
  recoverable: boolean;
  rolloutState: "legacy" | "shadow" | "active" | "blocked";
  protocolVersion: 2;
  generation: number | null;
  workbookRevision: number | null;
  materializedContentVersion: number | null;
  shadowValidationCount: number;
  activationReady: boolean;
  actorMembershipIncarnation: string;
  canActivate: boolean;
  shards: WorkbookCollabV2Shard[];
  /** The server is making the rooms ready after an import: wait for it. */
  roomsBeingPrepared?: boolean;
  /** The API reads a mutation's `baseCells` (an API from before it would
   * refuse the unknown field, so a window sends them only when told). */
  cellBaseGuard?: boolean;
}

export interface WorkbookNativeChartCacheAck {
  schemaVersion: "workbook-native-chart-cache-ack-v1";
  sourceOperationId: string | null;
  derivedOperationId: string;
  entries: { sheetId: string; objectId: string; part: string }[];
}

export interface WorkbookShardProjectionReceipt {
  pageId: string;
  shardKey: string;
  projectionRevision: number;
  workbookRevision: number;
  canonicalWorkbookRevision?: number;
  nativeChartCacheAck?: WorkbookNativeChartCacheAck;
  contentVersion: number;
  sharedContentSha256: string;
  contentJson: Record<string, unknown>;
  actorMembershipIncarnation: string;
}

export interface WorkbookProjectionOperationReceipt {
  pageId: string;
  operationId: string;
  duplicate: boolean;
  workbookRevision: number;
  canonicalWorkbookRevision?: number;
  nativeChartCacheAck?: WorkbookNativeChartCacheAck;
  projectionRevisions: Record<string, number>;
  contentVersion: number;
  sharedContentSha256: string;
  contentJson: Record<string, unknown>;
  actorMembershipIncarnation: string;
}

const WORKBOOK_V2_STATE_KEYS = new Set([
  "activation_ready",
  "actor_membership_incarnation",
  "can_activate",
  "eligible",
  "generation",
  "materialized_content_version",
  "page_id",
  "protocol_version",
  "recoverable",
  "rollout_state",
  "shadow_validation_count",
  "shards",
  "workbook_revision",
]);
const WORKBOOK_V2_SHARD_KEYS = new Set([
  "head_sequence",
  "key",
  "kind",
  "projection_revision",
  "snapshot_present",
  "snapshot_sequence",
]);
const WORKBOOK_V2_PROJECTION_KEYS = new Set([
  "native_chart_cache_ack",
  "canonical_workbook_revision",
  "actor_membership_incarnation",
  "content_json",
  "content_version",
  "page_id",
  "projection_revision",
  "shard_key",
  "shared_content_sha256",
  "workbook_revision",
]);
const WORKBOOK_V2_OPERATION_KEYS = new Set([
  "native_chart_cache_ack",
  "canonical_workbook_revision",
  "actor_membership_incarnation",
  "content_json",
  "content_version",
  "duplicate",
  "operation_id",
  "page_id",
  "projection_revisions",
  "shared_content_sha256",
  "workbook_revision",
]);

function workbookReceiptKeys(
  wire: Record<string, unknown>,
  expected: ReadonlySet<string>,
): boolean {
  const keys = new Set(expected);
  for (const optional of [
    "canonical_workbook_revision",
    "native_chart_cache_ack",
  ]) {
    if (!Object.hasOwn(wire, optional)) keys.delete(optional);
  }
  return exactKeys(wire, keys);
}

function nativeChartCacheReceiptFields(
  wire: Record<string, unknown>,
  operation: boolean,
): { nativeChartCacheAck?: WorkbookNativeChartCacheAck } {
  const value = wire.native_chart_cache_ack;
  if (value === undefined) return {};
  const reject = (): never => {
    throw new InvalidPageCollabResponseError(
      "native chart cache acknowledgement is malformed",
    );
  };
  if (
    !isRecord(value) ||
    !exactKeys(
      value,
      new Set([
        "schema_version",
        "source_operation_id",
        "derived_operation_id",
        "entries",
      ]),
    ) ||
    value.schema_version !== "workbook-native-chart-cache-ack-v1" ||
    value.source_operation_id !== (operation ? wire.operation_id : null) ||
    typeof value.derived_operation_id !== "string" ||
    !value.derived_operation_id.trim() ||
    value.derived_operation_id === value.source_operation_id ||
    !Array.isArray(value.entries) ||
    value.entries.length === 0 ||
    value.entries.length > 1024 ||
    !isRecord(wire.content_json) ||
    !Array.isArray(wire.content_json.sheets) ||
    nonnegativeInteger(
      wire.canonical_workbook_revision,
      "canonical workbook revision",
    ) <= nonnegativeInteger(wire.workbook_revision, "workbook revision")
  )
    return reject();
  const entries: WorkbookNativeChartCacheAck["entries"] = [];
  const seen = new Set<string>();
  for (const raw of value.entries) {
    if (
      !isRecord(raw) ||
      !exactKeys(raw, new Set(["sheet_id", "object_id", "part"]))
    )
      return reject();
    const sheetId = identity(raw.sheet_id, "cache sheet id"),
      objectId = identity(raw.object_id, "cache object id"),
      part = identity(raw.part, "cache part");
    const key = JSON.stringify([sheetId, objectId]);
    if (seen.has(key) || !/^xl\/charts\/[^/\\]+\.xml$/.test(part))
      return reject();
    seen.add(key);
    const sheets = wire.content_json.sheets.filter(
      (sheet) => isRecord(sheet) && sheet.id === sheetId,
    );
    if (
      sheets.length !== 1 ||
      !isRecord(sheets[0]) ||
      !Array.isArray(sheets[0].objects)
    )
      return reject();
    const objects = sheets[0].objects.filter(
      (obj) => isRecord(obj) && obj.id === objectId,
    );
    const obj = objects[0];
    if (
      objects.length !== 1 ||
      !isRecord(obj) ||
      obj.type !== "chart" ||
      !isRecord(obj.source) ||
      obj.source.part !== part ||
      !isRecord(obj.data)
    )
      return reject();
    entries.push({ sheetId, objectId, part });
  }
  return {
    nativeChartCacheAck: {
      schemaVersion: "workbook-native-chart-cache-ack-v1",
      sourceOperationId: operation ? String(wire.operation_id) : null,
      derivedOperationId: identity(
        value.derived_operation_id,
        "native derived operation id",
      ),
      entries,
    },
  };
}

function canonicalWorkbookReceiptFields(wire: Record<string, unknown>): {
  canonicalWorkbookRevision?: number;
} {
  if (wire.canonical_workbook_revision === undefined) return {};
  const revision = nonnegativeInteger(
    wire.canonical_workbook_revision,
    "canonical workbook revision",
  );
  if (
    revision < nonnegativeInteger(wire.workbook_revision, "workbook revision")
  ) {
    throw new InvalidPageCollabResponseError(
      "canonical workbook revision precedes the accepted operation",
    );
  }
  return { canonicalWorkbookRevision: revision };
}

function nonnegativeInteger(value: unknown, field: string): number {
  if (typeof value !== "number" || !Number.isSafeInteger(value) || value < 0) {
    throw new InvalidPageCollabResponseError(field + " is invalid");
  }
  return value;
}

function positiveInteger(value: unknown, field: string): number {
  const normalized = nonnegativeInteger(value, field);
  if (normalized === 0) {
    throw new InvalidPageCollabResponseError(field + " is invalid");
  }
  return normalized;
}

function workbookV2StateFromWire(
  wireValue: unknown,
  expectedPageId: string,
): WorkbookCollabV2State {
  if (
    !isRecord(wireValue) ||
    // Every key it needs; one it does not know is ignored. Releases roll the
    // API first: a field added there refused by a window of the last release
    
    ![...WORKBOOK_V2_STATE_KEYS].every((key) => key in wireValue) ||
    wireValue.page_id !== expectedPageId ||
    typeof wireValue.eligible !== "boolean" ||
    typeof wireValue.recoverable !== "boolean" ||
    typeof wireValue.activation_ready !== "boolean" ||
    typeof wireValue.can_activate !== "boolean" ||
    wireValue.protocol_version !== 2 ||
    !Array.isArray(wireValue.shards)
  ) {
    throw new InvalidPageCollabResponseError(
      "workbook collaboration response is malformed",
    );
  }
  const rolloutState = wireValue.rollout_state;
  if (
    rolloutState !== "legacy" &&
    rolloutState !== "shadow" &&
    rolloutState !== "active" &&
    rolloutState !== "blocked"
  ) {
    throw new InvalidPageCollabResponseError(
      "workbook collaboration rollout state is invalid",
    );
  }
  const nullableInteger = (
    value: unknown,
    field: string,
    positive: boolean,
  ): number | null => {
    if (value === null) return null;
    return positive
      ? positiveInteger(value, field)
      : nonnegativeInteger(value, field);
  };
  const seen = new Set<string>();
  const shards = wireValue.shards.map((raw): WorkbookCollabV2Shard => {
    if (
      !isRecord(raw) ||
      ![...WORKBOOK_V2_SHARD_KEYS].every((key) => key in raw) ||
      (raw.kind !== "manifest" &&
        raw.kind !== "sheet" &&
        raw.kind !== "tile" &&
        raw.kind !== "objects") ||
      typeof raw.snapshot_present !== "boolean"
    ) {
      throw new InvalidPageCollabResponseError(
        "workbook shard descriptor is malformed",
      );
    }
    const key = identity(raw.key, "workbook shard key");
    if (seen.has(key)) {
      throw new InvalidPageCollabResponseError(
        "workbook shard descriptor is duplicated",
      );
    }
    seen.add(key);
    return {
      key,
      kind: raw.kind,
      projectionRevision: nonnegativeInteger(
        raw.projection_revision,
        "projection revision",
      ),
      headSequence: nonnegativeInteger(raw.head_sequence, "head sequence"),
      snapshotSequence: nonnegativeInteger(
        raw.snapshot_sequence,
        "snapshot sequence",
      ),
      snapshotPresent: raw.snapshot_present,
    };
  });
  return {
    pageId: expectedPageId,
    eligible: wireValue.eligible,
    recoverable: wireValue.recoverable,
    roomsBeingPrepared: wireValue.rooms_being_prepared === true,
    rolloutState,
    protocolVersion: 2,
    generation: nullableInteger(wireValue.generation, "generation", true),
    workbookRevision: nullableInteger(
      wireValue.workbook_revision,
      "workbook revision",
      false,
    ),
    materializedContentVersion: nullableInteger(
      wireValue.materialized_content_version,
      "materialized content version",
      true,
    ),
    shadowValidationCount: nonnegativeInteger(
      wireValue.shadow_validation_count,
      "shadow validation count",
    ),
    activationReady: wireValue.activation_ready,
    ...(wireValue.cell_base_guard === true ? { cellBaseGuard: true } : {}),
    actorMembershipIncarnation: identity(
      wireValue.actor_membership_incarnation,
      "actor membership incarnation",
    ),
    canActivate: wireValue.can_activate,
    shards,
  };
}

const workbookV2Base = (pageId: string) => base(pageId) + "/workbook-collab-v2";

/** How long ago the rooms last changed, when that was after the last save. */
export async function getWorkbookRoomsUnsavedSeconds(pageId: string): Promise<number | null> {
  const wire = await apiGet<unknown>(`${workbookV2Base(pageId)}/rooms-unsaved`);
  const seconds = isRecord(wire) ? wire.seconds : null;
  return typeof seconds === "number" && Number.isFinite(seconds) && seconds >= 0 ? seconds : null;
}

/** A closing window's edits its rooms never got (workbook_closing_edits): sent
 * as the window goes, so nothing waits for the answer but a window still open. */
export async function sendWorkbookClosingEdits(pageId: string, edits: unknown): Promise<number | null> {
  const wire = await apiPost<unknown>(`${workbookV2Base(pageId)}/closing-edits`, edits, { keepalive: true });
  const version = isRecord(wire) ? wire.content_version : null;
  return typeof version === "number" ? version : null;
}

/** What a page held as one shared document had beyond its checkpoint as its
 * window closed (services/page_closing_edits): sent as the window goes. */
export interface PageClosingEditsRequest {
  closing_id: string;
  base_content_version: number;
  /** A shared document: its original operations. */
  epoch?: number;
  model?: string;
  update_base64?: string;
  /** A document saved as JSON (held as parts): its changed blocks. */
  /** `inflight_base`: the version (a part's revision) a save still on its
   * way was made from; landing after these edits, it is refused. */
  body?: { changes: unknown[]; attrs?: Record<string, unknown> | null; inflight_base?: number };
  parts?: Array<{ index: number; changes: unknown[]; inflight_base?: number }>;
}

export async function sendPageClosingEdits(pageId: string, body: PageClosingEditsRequest): Promise<number | null> {
  const wire = await apiPost<unknown>(`${base(pageId)}/closing-edits`, body, { keepalive: true });
  const version = isRecord(wire) ? wire.content_version : null;
  return typeof version === "number" ? version : null;
}

/** The server took a newer request of this window instead of this one
 * (`custody_gate`): nothing of it was kept on its own. */
export const CUSTODY_SUPERSEDED = "superseded" as const;

/** A deck window's edits its room has not confirmed, handed to the room's
 * durable custody while the window is open (services/page_room_durable
 * `take_edits_into_custody`): the closing edits' causal form. Resolves with
 * the custody's sequence, or CUSTODY_SUPERSEDED (not kept: a newer request of
 * this window was taken instead); rejects with the request's error. `keepalive` for a
 * body small enough to outlive the page. */
// A server from before custody_gate refuses the window's client id (its
// payload forbids unknown fields): from then on the requests go without it.
let custodyClientIdRefused = false;

export async function sendPageCustodyEdits(
  pageId: string,
  body: { epoch: number; model: string; update_base64: string; client_id?: number },
  signal?: AbortSignal,
): Promise<number | null | typeof CUSTODY_SUPERSEDED> {
  const keepalive = body.update_base64.length < 60 * 1024;
  // High priority: waiting for a connection, it goes before the window's
  // other requests (HTTP/1.1 dispatches the queue by priority).
  const post = (payload: typeof body) => apiPost<unknown>(`${base(pageId)}/custody-edits`, payload,
    { ...(keepalive ? { keepalive: true } : {}), ...(signal ? { signal } : {}), priority: "high" });
  const { client_id: clientId, ...anonymous } = body;
  const withId = clientId !== undefined && !custodyClientIdRefused;
  let wire: unknown;
  try {
    wire = await post(withId ? body : anonymous);
  } catch (error) {
    if (!withId || !isApiRequestError(error) || error.status !== 422 || error.code !== "validation_error"
      || !JSON.stringify(error.details ?? {}).includes("client_id")) throw error;
    custodyClientIdRefused = true;
    wire = await post(anonymous);
  }
  if (isRecord(wire) && wire.outcome === CUSTODY_SUPERSEDED) return CUSTODY_SUPERSEDED;
  const sequence = isRecord(wire) ? wire.sequence : null;
  return typeof sequence === "number" ? sequence : null;
}

/** What became of closing edits this account sent: unknown (not heard of),
 * received, applied, unchanged or refused. Null from a server without the read. */
export async function getPageClosingEditsState(
  pageId: string,
  closingId: string,
): Promise<{ state: string; conflicts: Array<{ kept: string; copy: string }> } | null> {
  try {
    const wire = await apiGet<unknown>(`${base(pageId)}/closing-edits/${encodeURIComponent(closingId)}`, { cache: "no-store" });
    const state = isRecord(wire) ? wire.state : null;
    if (typeof state !== "string") return null;
    const listed = isRecord(wire) && Array.isArray(wire.conflicts) ? wire.conflicts : [];
    const conflicts = listed.filter((pair): pair is { kept: string; copy: string } =>
      isRecord(pair) && typeof pair.kept === "string" && typeof pair.copy === "string");
    return { state, conflicts };
  } catch (error) {
    if (isApiRequestError(error) && (error.status === 404 || error.status === 405)) return null;
    throw error;
  }
}

/** Edits made on a page that has moved on since, laid on it cell by cell. */
export async function mergeWorkbookEdits(pageId: string, edits: Record<string, unknown>): Promise<number | null> {
  const wire = await apiPost<unknown>(`${workbookV2Base(pageId)}/closing-edits`, { ...edits, merge: true });
  const version = isRecord(wire) ? wire.content_version : null;
  return typeof version === "number" ? version : null;
}

/** The page's content version, without its body, and whether its rooms are still
 * being started from it. Null from a server without this read. */
export async function getWorkbookPageVersion(
  pageId: string,
): Promise<{ version: number; roomsBehind: boolean } | null> {
  try {
    const wire = await apiGet<unknown>(`${workbookV2Base(pageId)}/page-version`);
    if (!isRecord(wire) || typeof wire.content_version !== "number") return null;
    return { version: wire.content_version, roomsBehind: wire.rooms_behind === true };
  } catch (error) {
    if (isApiRequestError(error) && (error.status === 404 || error.status === 405)) return null;
    throw error;
  }
}

export async function getWorkbookCollabV2(
  pageId: string,
  init?: RequestInit,
): Promise<WorkbookCollabV2State> {
  const expectedPageId = identity(pageId, "page id");
  return workbookV2StateFromWire(
    await apiGet<unknown>(workbookV2Base(pageId), init),
    expectedPageId,
  );
}

export async function bootstrapWorkbookCollabV2(
  pageId: string,
): Promise<WorkbookCollabV2State> {
  const expectedPageId = identity(pageId, "page id");
  return workbookV2StateFromWire(
    await apiPost<unknown>(workbookV2Base(pageId) + "/bootstrap", {}),
    expectedPageId,
  );
}

export async function activateWorkbookCollabV2(
  pageId: string,
): Promise<WorkbookCollabV2State> {
  const expectedPageId = identity(pageId, "page id");
  return workbookV2StateFromWire(
    await apiPost<unknown>(workbookV2Base(pageId) + "/activate", {}),
    expectedPageId,
  );
}

export async function deactivateWorkbookCollabV2(
  pageId: string,
): Promise<WorkbookCollabV2State> {
  const expectedPageId = identity(pageId, "page id");
  return workbookV2StateFromWire(
    await apiPost<unknown>(workbookV2Base(pageId) + "/deactivate", {}),
    expectedPageId,
  );
}

export async function updateWorkbookShardProjection(
  pageId: string,
  shardKey: string,
  input: {
    baseProjectionRevision: number;
    basePageContentVersion: number;
    projection: Record<string, unknown>;
  },
): Promise<WorkbookShardProjectionReceipt> {
  const expectedPageId = identity(pageId, "page id");
  const expectedShardKey = identity(shardKey, "workbook shard key");
  const wire = await apiPut<unknown>(
    workbookV2Base(pageId) +
      "/shards/" +
      encodeURIComponent(expectedShardKey) +
      "/projection?canonical_receipt=1&native_cache_ack=1",
    {
      base_projection_revision: nonnegativeInteger(
        input.baseProjectionRevision,
        "projection revision",
      ),
      base_page_content_version: positiveInteger(
        input.basePageContentVersion,
        "page content version",
      ),
      projection: input.projection,
    },
  );
  if (
    !isRecord(wire) ||
    !workbookReceiptKeys(wire, WORKBOOK_V2_PROJECTION_KEYS) ||
    wire.page_id !== expectedPageId ||
    wire.shard_key !== expectedShardKey ||
    !isRecord(wire.content_json) ||
    typeof wire.shared_content_sha256 !== "string" ||
    !/^[0-9a-f]{64}$/u.test(wire.shared_content_sha256)
  ) {
    throw new InvalidPageCollabResponseError(
      "workbook projection response is malformed",
    );
  }
  return {
    pageId: expectedPageId,
    shardKey: expectedShardKey,
    ...canonicalWorkbookReceiptFields(wire),
    ...nativeChartCacheReceiptFields(wire, false),
    projectionRevision: nonnegativeInteger(
      wire.projection_revision,
      "projection revision",
    ),
    workbookRevision: nonnegativeInteger(
      wire.workbook_revision,
      "workbook revision",
    ),
    contentVersion: positiveInteger(
      wire.content_version,
      "page content version",
    ),
    sharedContentSha256: wire.shared_content_sha256,
    contentJson: wire.content_json,
    actorMembershipIncarnation: identity(
      wire.actor_membership_incarnation,
      "actor membership incarnation",
    ),
  };
}

/**
 * The operation receipt with ``content_json`` rebuilt from a
 * ``content_delta`` answer and the base body it was computed from (see the
 * API's ``workbook_content_delta``): the base with, per changed sheet, the
 * members set or removed and the cells set or removed. Any other answer
 * passes through unchanged.
 */
export function withWorkbookContentDelta(
  wire: unknown,
  base: { contentVersion: number; contentJson: Record<string, unknown> } | undefined,
): unknown {
  if (!isRecord(wire) || !Object.hasOwn(wire, "content_delta")) return wire;
  const malformed = () =>
    new InvalidPageCollabResponseError("workbook operation delta is malformed");
  const delta = wire.content_delta;
  const baseSheets = base?.contentJson.sheets;
  if (
    !base ||
    Object.hasOwn(wire, "content_json") ||
    !isRecord(delta) ||
    !exactKeys(delta, new Set(["base_content_version", "sheets"])) ||
    delta.base_content_version !== base.contentVersion ||
    !isRecord(delta.sheets) ||
    !Array.isArray(baseSheets)
  ) {
    throw malformed();
  }
  const entries = delta.sheets;
  for (const index of Object.keys(entries)) {
    if (!/^(0|[1-9][0-9]*)$/u.test(index) || Number(index) >= baseSheets.length) throw malformed();
  }
  const strings = (value: unknown): string[] => {
    if (!Array.isArray(value) || value.some((item) => typeof item !== "string")) throw malformed();
    return value as string[];
  };
  const sheets = baseSheets.map((sheet, index) => {
    const entry = entries[String(index)];
    if (entry === undefined) return sheet;
    if (
      !isRecord(entry) ||
      !exactKeys(entry, new Set(["fields", "removed_fields", "cells", "removed"])) ||
      !isRecord(entry.fields) ||
      !isRecord(entry.cells) ||
      !isRecord(sheet) ||
      !isRecord(sheet.cells) ||
      Object.hasOwn(entry.fields, "cells")
    ) {
      throw malformed();
    }
    const next: Record<string, unknown> = { ...sheet, ...entry.fields };
    for (const field of strings(entry.removed_fields)) {
      if (field === "cells") throw malformed();
      delete next[field];
    }
    const cells: Record<string, unknown> = { ...sheet.cells };
    for (const key of strings(entry.removed)) delete cells[key];
    for (const [key, value] of Object.entries(entry.cells)) cells[key] = value;
    next.cells = cells;
    return next;
  });
  const rest: Record<string, unknown> = { ...wire };
  delete rest.content_delta;
  return { ...rest, content_json: { ...base.contentJson, sheets } };
}

export async function applyWorkbookProjectionOperation(
  pageId: string,
  input: {
    baseWorkbookRevision: number;
    basePageContentVersion: number;
    idempotencyKey: string;
    operationType: string;
    mutations: Array<{
      shardKey: string;
      baseProjectionRevision: number | null;
      projection?: Record<string, unknown>;
      cellPatch?: Record<string, Record<string, unknown> | null>;
      /** What this window held in each cell the mutation changes (value and
       * formula; null: no cell): the server refuses the save when the shard
       * holds something else there. */
      baseCells?: Record<string, Record<string, unknown> | null>;
    }>;
    removeShardKeys: string[];
    /** Shards whose rooms this window could not pass the change through (it
     * is not connected to them): the server writes it into those rooms. */
    unrelayedShardKeys?: readonly string[];
    /** The page body of ``basePageContentVersion`` exactly as an earlier
     * receipt gave it: the server may then answer with what changed from it
     * (``content_delta``) instead of the whole body, and the receipt's
     * ``contentJson`` is rebuilt from it here. */
    durableBase?: {
      contentVersion: number;
      contentJson: Record<string, unknown>;
      /** The page read's ``content_sha256`` when the body came from a read of
       * the page, not from a receipt. */
      contentSha256?: string;
    };
  },
): Promise<WorkbookProjectionOperationReceipt> {
  const expectedPageId = identity(pageId, "page id");
  if (!/^[0-9a-f]{32,64}$/u.test(input.idempotencyKey)) {
    throw new TypeError("workbook operation idempotency key is invalid");
  }
  const durableBase =
    input.durableBase &&
    input.durableBase.contentVersion === input.basePageContentVersion &&
    Array.isArray(input.durableBase.contentJson.sheets)
      ? input.durableBase
      : undefined;
  const received = await apiPost<unknown>(
    workbookV2Base(pageId) +
      "/operations?canonical_receipt=1&native_cache_ack=1" +
      (durableBase ? "&content_delta=1" : "") +
      (durableBase?.contentSha256 ? `&base_content_sha256=${durableBase.contentSha256}` : "") +
      (input.unrelayedShardKeys ?? []).map((key) => `&unrelayed=${encodeURIComponent(key)}`).join(""),
    {
      base_workbook_revision: nonnegativeInteger(
        input.baseWorkbookRevision,
        "workbook revision",
      ),
      base_page_content_version: positiveInteger(
        input.basePageContentVersion,
        "page content version",
      ),
      idempotency_key: input.idempotencyKey,
      operation_type: input.operationType,
      mutations: input.mutations.map((mutation) => {
        const hasProjection = isRecord(mutation.projection);
        const hasCellPatch = isRecord(mutation.cellPatch);
        if (hasProjection === hasCellPatch) {
          throw new TypeError(
            "workbook mutation must carry one projection or cell patch",
          );
        }
        return {
          shard_key: identity(mutation.shardKey, "workbook shard key"),
          base_projection_revision:
            mutation.baseProjectionRevision === null
              ? null
              : nonnegativeInteger(
                  mutation.baseProjectionRevision,
                  "projection revision",
                ),
          ...(hasProjection
            ? { projection: mutation.projection }
            : { cell_patch: mutation.cellPatch }),
          ...(isRecord(mutation.baseCells) && Object.keys(mutation.baseCells).length > 0
            ? { base_cells: mutation.baseCells }
            : {}),
        };
      }),
      remove_shard_keys: input.removeShardKeys.map((key) =>
        identity(key, "workbook shard key"),
      ),
    },
  );
  const wire = withWorkbookContentDelta(received, durableBase);
  if (
    !isRecord(wire) ||
    !workbookReceiptKeys(wire, WORKBOOK_V2_OPERATION_KEYS) ||
    wire.page_id !== expectedPageId ||
    typeof wire.operation_id !== "string" ||
    typeof wire.duplicate !== "boolean" ||
    !isRecord(wire.projection_revisions) ||
    !isRecord(wire.content_json) ||
    typeof wire.shared_content_sha256 !== "string" ||
    !/^[0-9a-f]{64}$/u.test(wire.shared_content_sha256)
  ) {
    throw new InvalidPageCollabResponseError(
      "workbook operation response is malformed",
    );
  }
  const projectionRevisions: Record<string, number> = {};
  for (const [key, value] of Object.entries(wire.projection_revisions)) {
    projectionRevisions[identity(key, "workbook shard key")] =
      nonnegativeInteger(value, "projection revision");
  }
  return {
    pageId: expectedPageId,
    operationId: identity(wire.operation_id, "workbook operation id"),
    ...canonicalWorkbookReceiptFields(wire),
    ...nativeChartCacheReceiptFields(wire, true),
    duplicate: wire.duplicate,
    workbookRevision: nonnegativeInteger(
      wire.workbook_revision,
      "workbook revision",
    ),
    projectionRevisions,
    contentVersion: positiveInteger(
      wire.content_version,
      "page content version",
    ),
    sharedContentSha256: wire.shared_content_sha256,
    contentJson: wire.content_json,
    actorMembershipIncarnation: identity(
      wire.actor_membership_incarnation,
      "actor membership incarnation",
    ),
  };
}
