/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Mounted shared notebooks (ARCH8): a notebook shared WITH me lives in its
 * owner's workspace, but I open it from MY OWN library without ever
 * "switching" — every API call made while the app is inside that notebook
 * simply carries the notebook's HOST workspace id.
 *
 * The mapping (notebookId → host workspace id) is written whenever the
 * shared-notebooks endpoint is read (library, background sync) and consulted:
 *  - by lib/api's notebook-scoped header injection, keyed off the current URL
 *    AND the requested API path;
 *  - by the activity tracker, so work done inside a mounted notebook is
 *    captured (and uploaded) against the host workspace's feed.
 *
 * The durable cache is paired with a small external-store contract. Writes in
 * this tab and `storage` events from another tab therefore update mounted
 * surfaces immediately instead of waiting for a remount. A mount is resource
 * authority, not a global workspace switch: calendar, memory, settings,
 * messaging and other personal services must keep using the caller's private
 * workspace while this page is open.
 */

const STORAGE_KEY = "mrn:notebook-workspace-mounts";
const ROLE_STORAGE_KEY = "mrn:notebook-workspace-mount-roles";
const AUTHORITY_STORAGE_KEY = "mrn:notebook-workspace-mount-authority-v2";

export const NOTEBOOK_MOUNTS_CHANGED_EVENT =
  "mrnote:notebook-workspace-mounts-changed";

export type MountedNotebookRole = "viewer" | "editor" | "manager";

interface MountedNotebookAuthority {
  workspaceId: string;
  role: MountedNotebookRole | null;
}

export interface NotebookMountSnapshot {
  /** Corrupt or unavailable storage must never be interpreted as "my notebook". */
  integrity: "ready" | "corrupt";
  mounts: Readonly<Record<string, MountedNotebookAuthority>>;
}

interface PersistedMountAuthority {
  version: 2;
  mounts: Record<
    string,
    { workspaceId: string; role: MountedNotebookRole | null }
  >;
}

type MountMap = Record<string, string>;

const SERVER_SNAPSHOT: NotebookMountSnapshot = Object.freeze({
  integrity: "ready",
  mounts: Object.freeze({}),
});

let cachedStorageFingerprint: string | null = null;
let cachedSnapshot: NotebookMountSnapshot = SERVER_SNAPSHOT;

function validRole(value: unknown): value is MountedNotebookRole {
  return value === "viewer" || value === "editor" || value === "manager";
}

function plainRecord(value: unknown): value is Record<string, unknown> {
  return Boolean(value) && typeof value === "object" && !Array.isArray(value);
}

function corruptSnapshot(): NotebookMountSnapshot {
  return Object.freeze({
    integrity: "corrupt" as const,
    mounts: Object.freeze({}),
  });
}

function readySnapshot(
  mounts: Record<string, MountedNotebookAuthority>,
): NotebookMountSnapshot {
  return Object.freeze({
    integrity: "ready" as const,
    mounts: Object.freeze(mounts),
  });
}

function parseAuthorityStorage(raw: string): NotebookMountSnapshot {
  try {
    const parsed = JSON.parse(raw) as unknown;
    if (!plainRecord(parsed) || parsed.version !== 2) {
      return corruptSnapshot();
    }
    if (!plainRecord(parsed.mounts)) return corruptSnapshot();
    const mounts: Record<string, MountedNotebookAuthority> = {};
    for (const [notebookId, candidate] of Object.entries(parsed.mounts)) {
      if (!notebookId.trim() || !plainRecord(candidate)) {
        return corruptSnapshot();
      }
      const workspaceId = candidate.workspaceId;
      const role = candidate.role;
      if (
        typeof workspaceId !== "string" ||
        !workspaceId.trim() ||
        (role !== null && !validRole(role))
      ) {
        return corruptSnapshot();
      }
      mounts[notebookId] = { workspaceId: workspaceId.trim(), role };
    }
    return readySnapshot(mounts);
  } catch {
    return corruptSnapshot();
  }
}

function parseLegacyStorage(
  mountsRaw: string | null,
  rolesRaw: string | null,
): NotebookMountSnapshot {
  try {
    const parsedMounts = mountsRaw ? (JSON.parse(mountsRaw) as unknown) : {};
    const parsedRoles = rolesRaw ? (JSON.parse(rolesRaw) as unknown) : {};
    if (!plainRecord(parsedMounts) || !plainRecord(parsedRoles)) {
      return corruptSnapshot();
    }
    const roles: Record<string, MountedNotebookRole> = {};
    for (const [notebookId, role] of Object.entries(parsedRoles)) {
      if (!notebookId.trim() || !validRole(role)) return corruptSnapshot();
      roles[notebookId] = role;
    }
    const mounts: Record<string, MountedNotebookAuthority> = {};
    for (const [notebookId, workspaceId] of Object.entries(parsedMounts)) {
      if (
        !notebookId.trim() ||
        typeof workspaceId !== "string" ||
        !workspaceId.trim()
      ) {
        return corruptSnapshot();
      }
      mounts[notebookId] = {
        workspaceId: workspaceId.trim(),
        role: roles[notebookId] ?? null,
      };
    }
    return readySnapshot(mounts);
  } catch {
    return corruptSnapshot();
  }
}

/** Stable snapshot getter for React's `useSyncExternalStore`. */
export function notebookMountSnapshot(): NotebookMountSnapshot {
  if (typeof window === "undefined") return SERVER_SNAPSHOT;
  let authorityRaw: string | null;
  let mountsRaw: string | null;
  let rolesRaw: string | null;
  try {
    authorityRaw = window.localStorage.getItem(AUTHORITY_STORAGE_KEY);
    mountsRaw = authorityRaw ? null : window.localStorage.getItem(STORAGE_KEY);
    rolesRaw = authorityRaw
      ? null
      : window.localStorage.getItem(ROLE_STORAGE_KEY);
  } catch {
    authorityRaw = "<storage-unavailable>";
    mountsRaw = null;
    rolesRaw = null;
  }
  const fingerprint = authorityRaw
    ? `authority:${authorityRaw}`
    : `legacy:${mountsRaw ?? ""}\n${rolesRaw ?? ""}`;
  if (fingerprint === cachedStorageFingerprint) return cachedSnapshot;
  cachedStorageFingerprint = fingerprint;
  cachedSnapshot = authorityRaw
    ? authorityRaw === "<storage-unavailable>"
      ? corruptSnapshot()
      : parseAuthorityStorage(authorityRaw)
    : parseLegacyStorage(mountsRaw, rolesRaw);
  return cachedSnapshot;
}

export function notebookMountServerSnapshot(): NotebookMountSnapshot {
  return SERVER_SNAPSHOT;
}

function invalidateSnapshot(): void {
  cachedStorageFingerprint = null;
}

function emitMountChange(): void {
  invalidateSnapshot();
  if (typeof window !== "undefined") {
    window.dispatchEvent(new Event(NOTEBOOK_MOUNTS_CHANGED_EVENT));
  }
}

/** Subscribe to same-tab writes and cross-tab durable-cache changes. */
export function subscribeNotebookMounts(onStoreChange: () => void): () => void {
  if (typeof window === "undefined") return () => undefined;
  const onStorage = (event: StorageEvent) => {
    if (
      event.key !== AUTHORITY_STORAGE_KEY &&
      event.key !== STORAGE_KEY &&
      event.key !== ROLE_STORAGE_KEY &&
      event.key !== null
    ) {
      return;
    }
    invalidateSnapshot();
    onStoreChange();
  };
  window.addEventListener(NOTEBOOK_MOUNTS_CHANGED_EVENT, onStoreChange);
  window.addEventListener("storage", onStorage);
  return () => {
    window.removeEventListener(NOTEBOOK_MOUNTS_CHANGED_EVENT, onStoreChange);
    window.removeEventListener("storage", onStorage);
  };
}

function persistMounts(
  mounts: Readonly<Record<string, MountedNotebookAuthority>>,
): boolean {
  if (typeof window === "undefined") return false;
  const durable: PersistedMountAuthority = {
    version: 2,
    mounts: Object.fromEntries(
      Object.entries(mounts).map(([notebookId, mount]) => [
        notebookId,
        {
          workspaceId: mount.workspaceId,
          role: mount.role,
        },
      ]),
    ),
  };
  try {
    // One canonical write keeps workspace + role atomic for other tabs.
    window.localStorage.setItem(AUTHORITY_STORAGE_KEY, JSON.stringify(durable));
  } catch {
    return false;
  }

  // Keep legacy readers alive during a rolling frontend update. Failure here
  // cannot invalidate the already-committed canonical authority record.
  try {
    const workspaceMap: MountMap = {};
    const roleMap: Record<string, MountedNotebookRole> = {};
    for (const [notebookId, mount] of Object.entries(mounts)) {
      workspaceMap[notebookId] = mount.workspaceId;
      if (mount.role) roleMap[notebookId] = mount.role;
    }
    window.localStorage.setItem(STORAGE_KEY, JSON.stringify(workspaceMap));
    window.localStorage.setItem(ROLE_STORAGE_KEY, JSON.stringify(roleMap));
  } catch {
    // Canonical v2 storage above remains authoritative.
  }
  emitMountChange();
  return true;
}

export function rememberNotebookMounts(
  entries: readonly {
    notebookId: string;
    workspaceId: string;
    role?: "viewer" | "editor" | "manager";
  }[],
): void {
  if (typeof window === "undefined" || entries.length === 0) return;
  const current = notebookMountSnapshot();
  // A partial hint cannot safely reconstruct unseen grants after corruption.
  // Only the complete shared-notebooks snapshot may repair that state.
  if (current.integrity !== "ready") return;
  const mounts: Record<string, MountedNotebookAuthority> = {
    ...current.mounts,
  };
  let changed = false;
  for (const entry of entries) {
    const notebookId = entry.notebookId.trim();
    const workspaceId = entry.workspaceId.trim();
    if (!notebookId || !workspaceId) continue;
    const previous = mounts[notebookId];
    const role =
      entry.role ??
      (previous?.workspaceId === workspaceId ? previous.role : null);
    if (previous?.workspaceId === workspaceId && previous.role === role) {
      continue;
    }
    mounts[notebookId] = { workspaceId, role };
    changed = true;
  }
  if (changed) persistMounts(mounts);
}

export function forgetNotebookMount(notebookId: string): void {
  if (typeof window === "undefined") return;
  const current = notebookMountSnapshot();
  if (current.integrity !== "ready" || !(notebookId in current.mounts)) return;
  const mounts: Record<string, MountedNotebookAuthority> = {
    ...current.mounts,
  };
  delete mounts[notebookId];
  persistMounts(mounts);
}

/**
 * Replace all shared-with-me mounts from one complete server response.
 * This is the only mutation allowed to recover corrupt storage, and it makes
 * demotions/removals authoritative while retaining every grant still present
 * in the same complete result.
 */
export function reconcileSharedNotebookMounts(
  entries: readonly {
    notebookId: string;
    workspaceId: string;
    role: MountedNotebookRole;
  }[],
): void {
  if (typeof window === "undefined") return;
  const mounts: Record<string, MountedNotebookAuthority> = {};
  for (const entry of entries) {
    const notebookId = entry.notebookId.trim();
    const workspaceId = entry.workspaceId.trim();
    if (!notebookId || !workspaceId || !validRole(entry.role)) continue;
    mounts[notebookId] = { workspaceId, role: entry.role };
  }
  persistMounts(mounts);
}

/**
 * Reconcile the workspace returned by an authoritative notebook read with the
 * caller's selected workspace.
 *
 * A notebook in the caller's own workspace is not a mount, even if an older
 * client accidentally cached it as one. A different workspace is a genuine
 * foreign mount; when its member role has not arrived yet, consumers must
 * keep failing closed as viewer.
 */
export function reconcileNotebookWorkspaceMount(input: {
  notebookId: string;
  notebookWorkspaceId: string;
  personalWorkspaceId: string | null;
}): void {
  const notebookId = input.notebookId.trim();
  const notebookWorkspaceId = input.notebookWorkspaceId.trim();
  const personalWorkspaceId = input.personalWorkspaceId?.trim() ?? "";
  if (!notebookId || !notebookWorkspaceId || !personalWorkspaceId) return;

  if (notebookWorkspaceId === personalWorkspaceId) {
    if (notebookMountSnapshot().integrity === "corrupt") {
      // This exact notebook GET proves personal ownership. Discard an
      // unreadable cache rather than leaving every notebook surface blocked;
      // the next complete shared-list refresh repopulates valid foreign mounts.
      persistMounts({});
      return;
    }
    forgetNotebookMount(notebookId);
    return;
  }

  rememberNotebookMounts([{ notebookId, workspaceId: notebookWorkspaceId }]);
}

const NOTEBOOK_PATH = /\/(?:app|workspace)\/notebooks\/([a-zA-Z0-9-]+)/;

const NOTEBOOK_ENTITY_PATHS = [
  /^\/api\/v1\/notebooks\/(?!home(?:\/|$))[^/]+(?:\/|$)/,
  /^\/api\/v1\/pages\/(?!search(?:\/|$))[^/]+(?:\/|$)/,
  /^\/api\/v1\/resources\/[^/]+(?:\/|$)/,
  /^\/api\/v1\/attachments\/[^/]+(?:\/|$)/,
  /^\/api\/v1\/data-items\/[^/]+(?:\/|$)/,
  /^\/api\/v1\/annotations\/[^/]+(?:\/|$)/,
  /^\/api\/v1\/datasets(?:\/|$)/,
  /^\/api\/v1\/universal-documents(?:\/|$)/,
  /^\/api\/v1\/study-assets\/[^/]+(?:\/|$)/,
  /^\/api\/v1\/decks\/[^/]+(?:\/|$)/,
  /^\/api\/v1\/cards\/[^/]+(?:\/|$)/,
] as const;

const NOTEBOOK_SERVICE_PATHS = [
  /^\/api\/v1\/ai\/notebook(?:\/|$)/,
  // An editor Copilot ask run durably is the same page-scoped selection
  // action (start, a page's pending result, settle): it is scoped as that
  
  // notebook in another workspace was "not found").
  /^\/api\/v1\/ai\/turn-runs\/copilot(?:\/|$)/,
  /^\/api\/v1\/ai\/study(?:\/|$)/,
  /^\/api\/v1\/book-learning(?:\/|$)/,
  /^\/api\/v1\/notebook-ai(?:\/|$)/,
  /^\/api\/v1\/uploads\/(?:presign|complete)(?:\/|$)/,
] as const;

const NOTEBOOK_FILTERED_COLLECTION_PATHS = new Set([
  "/api/v1/resources",
  "/api/v1/activity/page-events",
  "/api/v1/activity/segments",
]);

function parseApiUrl(path: string): URL | null {
  try {
    return new URL(path, "https://mrnote.invalid");
  } catch {
    return null;
  }
}

/**
 * Whether a JSON API request is unambiguously owned by the notebook currently
 * mounted in the page.
 *
 * This is deliberately an allow-list. A new global service therefore defaults
 * to the caller's private workspace instead of silently inheriting a foreign
 * notebook host. Collection endpoints are mounted only when their URL carries
 * an explicit notebook filter; entity endpoints carry their own notebook-owned
 * identifier. Binary notebook resources use `activeWorkspaceHeaders` and do
 * not depend on this classifier.
 */
export function apiPathUsesMountedNotebookWorkspace(path: string): boolean {
  const url = parseApiUrl(path);
  if (!url) return false;
  const { pathname, searchParams } = url;

  if (
    NOTEBOOK_ENTITY_PATHS.some((pattern) => pattern.test(pathname)) ||
    NOTEBOOK_SERVICE_PATHS.some((pattern) => pattern.test(pathname))
  ) {
    return true;
  }

  return (
    NOTEBOOK_FILTERED_COLLECTION_PATHS.has(pathname) &&
    Boolean(searchParams.get("notebook_id")?.trim())
  );
}

/**
 * Host workspace override for the CURRENT page, or null when the page is
 * not inside a mounted foreign notebook. Reads window.location so it is
 * always consistent with what the user is looking at.
 */
export function activeMountWorkspaceId(): string | null {
  if (typeof window === "undefined") return null;
  const match = NOTEBOOK_PATH.exec(window.location.pathname);
  if (!match) return null;
  return mountWorkspaceFor(match[1]);
}

/** The notebook the current page is inside, by its address; null elsewhere. */
export function activeRouteNotebookId(): string | null {
  if (typeof window === "undefined") return null;
  return NOTEBOOK_PATH.exec(window.location.pathname)?.[1] ?? null;
}

/** Mapping lookup by notebook id (for links). */
export function mountWorkspaceFor(notebookId: string): string | null {
  const snapshot = notebookMountSnapshot();
  return snapshot.integrity === "ready"
    ? (snapshot.mounts[notebookId]?.workspaceId ?? null)
    : null;
}

/** The exact NotebookMember role cached with a foreign notebook mount. */
export function mountRoleFor(notebookId: string): MountedNotebookRole | null {
  const snapshot = notebookMountSnapshot();
  return snapshot.integrity === "ready"
    ? (snapshot.mounts[notebookId]?.role ?? null)
    : null;
}

/**
 * Resolve the cached foreign-mount availability hint for one notebook.
 * Same-workspace cache entries are ignored defensively so they can never turn
 * an owner's private notebook into a mounted viewer while cleanup completes.
 */
export function mountedNotebookAccessFor(
  notebookId: string,
  personalWorkspaceId: string | null | undefined,
): { workspaceId: string; role: MountedNotebookRole } | null {
  return mountedNotebookAccessFromSnapshot(
    notebookMountSnapshot(),
    notebookId,
    personalWorkspaceId,
  );
}

/** Pure resolver for consumers already subscribed to the external store. */
export function mountedNotebookAccessFromSnapshot(
  snapshot: NotebookMountSnapshot,
  notebookId: string,
  personalWorkspaceId: string | null | undefined,
): { workspaceId: string; role: MountedNotebookRole } | null {
  if (snapshot.integrity !== "ready") return null;
  const mount = snapshot.mounts[notebookId];
  if (!mount) return null;
  const personal = personalWorkspaceId?.trim() ?? "";
  if (personal && mount.workspaceId === personal) return null;
  return {
    workspaceId: mount.workspaceId,
    role: mount.role ?? "viewer",
  };
}
