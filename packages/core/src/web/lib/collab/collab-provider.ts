/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Provider factory for the real-time collaborative editing foundation.
 *
 * {@link createCollabSession} returns a Yjs document plus an awareness instance
 * and (optionally) a network provider, behind a pluggable abstraction:
 *
 * - No room configured  → a LOCAL-ONLY session: a bare `Y.Doc` + a detached
 *   `Awareness`, with `provider = null`. This is the single-user case — the
 *   editor binds to the Y.Doc and works exactly as before, with NO network.
 * - A room + a `ws://`/`wss://` URL → a websocket-style provider (constructed
 *   lazily; the concrete client is injected to keep this module
 *   dependency-light and serverless by default).
 * - A room + no usable URL → serverless WebRTC P2P via `y-webrtc`.
 *
 * The provider construction is lazy and the network packages are imported only
 * on the branch that needs them, so importing this module (e.g. for the
 * local-only path or for tests) never eagerly pulls a signaling/websocket
 * client.
 *
 * The pure decision/normalization logic lives in collab-state.ts; this file is
 * the thin, side-effectful (Y.Doc + provider lifecycle) layer.
 */

import * as Y from "yjs";
import { Awareness } from "y-protocols/awareness";
import {
  providerKindFor,
  sanitizeCollabUser,
  type CollabUser,
  localAwarenessUserState,
  type CollabProviderKind,
} from "./collab-state";
import { fetchSelfCollabIdentity } from "./self-identity";
import { autoRecoverTabId } from "../autorecover-store";
import {
  collabSocketOrigin,
  ticketedWebsocketProviderFactory,
  type CollabNetworkState,
} from "./ws-provider";
import { installYJsonWriteGuard } from "./y-json-values";

// Every session's document takes JSON only: an `undefined` the API would read
// as null never reaches the room (y-json-values.ts).
installYJsonWriteGuard();

/** Minimal shape a network provider must satisfy for the editor + presence UI. */
export interface CollabNetworkProvider {
  awareness: Awareness;
  destroy: () => void;
  /** A deck's causal room whose server keeps each update in the database
   * and confirms it (`update_durable`): see PageYjsProvider.durableUpdates. */
  durableUpdates?: () => { active: boolean; settled: boolean; refusal: string | null; sequence?: number | null };
  subscribeDurableUpdates?: (listener: () => void) => () => void;
  /** Seed only an elected empty page room from a verified durable checkpoint. */
  seedDurablePageState?: () => boolean;
  /** The durable checkpoint this window loaded with the page: its digest, and
   * its bytes once verified (`null` when they are not that digest's). A room
   * whose snapshot is those same bytes (the server seeds an empty causal
   * room from its stored checkpoint) is not downloaded again. */
  offerDurablePageSnapshot?: (offer: {
    readonly sha256: string;
    readonly update: Uint8Array | PromiseLike<Uint8Array | null>;
  }) => void;
  /** Preserve a history origin while excluding verified server bytes from the
   * local transport queue. The marker is scoped to this synchronous apply. */
  withDurablePageUpdate?: <T>(origin: object, apply: () => T) => T;
  /** Re-mint the ticket/socket while preserving the Y.Doc and awareness.
   * `options.becameWritable`: the caller learned the page turned writable. */
  refreshAuthorization?: (options?: { becameWritable?: boolean }) => Promise<boolean>;
  /**
   * Synchronously stop accepting LOCAL Y.Doc updates for network custody.
   * Remote replay and awareness remain active; only the same caller's explicit
   * resume may release this hold after authorization is current.
   */
  suspendLocalDocumentUpdates?: () => void;
  /** Release the caller-owned hold after its authority snapshot is current. */
  resumeLocalDocumentUpdates?: () => void;
  /** Subscribe to the relay's live writable verdict for this connection. */
  subscribeWritable?: (listener: (writable: boolean) => void) => () => void;
  /** Sync/seeding state for providers that support the readiness contract
   * (the Stage-2 ticketed websocket provider). Absent on e.g. WebRTC. */
  network?: CollabNetworkState;
}

/**
 * Factory for a websocket-style provider. Injected (rather than imported) so the
 * default serverless build never bundles a websocket client and tests can supply
 * a fake. Mirrors the `y-websocket` `WebsocketProvider(url, room, doc, opts)`
 * signature.
 */
export type WebsocketProviderFactory = (args: {
  url: string;
  roomName: string;
  doc: Y.Doc;
  awareness: Awareness;
  pageJsonCausalModel?: boolean;
  pageJsonCausalEpoch?: number;
  pageJsonTextModel?: boolean;
}) => CollabNetworkProvider;

export interface CreateCollabSessionOptions {
  /** Stable document identifier (used as the default room name). */
  docId: string;
  /** Supports the page's declared durable original checkpoint model (JSON
   * or native Code). The wire capability keeps its original JSON-era name. */
  pageJsonCausalModel?: boolean;
  /** Migrated original history requires a fresh document bound to this epoch. */
  pageJsonCausalEpoch?: number;
  /** The room stores collaborative text as Y.Text (json-crdt.ts): join only
   * when the server acknowledges keeping older clients out of it. */
  pageJsonTextModel?: boolean;
  /**
   * Explicit room name shared by all peers editing the same doc. When omitted
   * the session is LOCAL-ONLY even if a providerUrl is supplied, because a room
   * is what binds peers together.
   */
  roomName?: string | null;
  /** Local collaborator display name (sanitized + capped before broadcast). */
  userName?: string | null;
  /** Local collaborator cursor color (forced through the strict hex gate). */
  userColor?: string | null;
  /** Optional `ws://`/`wss://` URL selecting a websocket provider over WebRTC. */
  providerUrl?: string | null;
  /**
   * Optional comma-separated signaling server list for the WebRTC provider. When
   * omitted, `y-webrtc`'s built-in public signaling defaults are used.
   */
  webrtcSignaling?: string[] | null;
  /**
   * Injected websocket provider factory. Required only when a `ws://`/`wss://`
   * `providerUrl` is supplied; absent, a websocket config gracefully degrades to
   * a local-only session (the editor still works, just unsynced).
   */
  websocketProviderFactory?: WebsocketProviderFactory;
}

/** The handle returned by {@link createCollabSession}. */
export interface CollabSession {
  /** The shared CRDT document the editor binds to. */
  ydoc: Y.Doc;
  /** Awareness instance carrying local + remote presence (cursors/names/colors). */
  awareness: Awareness;
  /** The network provider, or `null` for a local-only (single-user) session. */
  provider: CollabNetworkProvider | null;
  /** Which kind of session was built (`local` | `webrtc` | `websocket`). */
  kind: CollabProviderKind;
  /** Readiness/seeding contract of the network transport, or null when the
   * session is local (or the provider kind does not support it). Editors
   * gate their initial-content seeding on this — see ws-provider.ts. */
  network: CollabNetworkState | null;
  /** Tear down the provider (if any), the awareness, and the Y.Doc. Idempotent. */
  destroy: () => void;
}

/**
 * Lazily construct a WebRTC provider. Imported only on this branch so the
 * serverless/local path and tests never bundle the signaling client.
 */
function createWebrtcProvider(args: {
  roomName: string;
  doc: Y.Doc;
  awareness: Awareness;
  signaling?: string[] | null;
}): CollabNetworkProvider {
  // Synchronous require keeps createCollabSession synchronous while still only
  // loading y-webrtc when a WebRTC session is actually requested.
  // eslint-disable-next-line @typescript-eslint/no-require-imports
  throw new Error("Live collaboration is not part of the open-source editors.");
  const { WebrtcProvider } = {} as typeof import("y-webrtc");
  const provider = new WebrtcProvider(args.roomName, args.doc, {
    awareness: args.awareness,
    ...(args.signaling && args.signaling.length > 0
      ? { signaling: args.signaling }
      : {}),
  });
  return provider as unknown as CollabNetworkProvider;
}

/**
 * Create a collaborative editing session.
 *
 * Always returns a usable {@link CollabSession}; when no room is configured (or a
 * websocket URL is given without a factory) the session is local-only — a plain
 * Y.Doc with a detached awareness and `provider = null` — so the editor keeps
 * working as a single-user surface with zero network.
 */
/** Awareness -> identity guard listener, so destroy() can detach it. */
const identityGuards = new WeakMap<Awareness, (payload: { updated: number[] }) => void>();

export function createCollabSession(
  options: CreateCollabSessionOptions,
): CollabSession {
  const ydoc = new Y.Doc();
  // A standalone Awareness is valid even with no provider: the local user state
  // can still be set, and a provider (if attached) shares this same instance.
  const awareness = new Awareness(ydoc);

  // Seed the local presence so cursors/names/colors are populated from the start.
  const user = sanitizeCollabUser({
    name: options.userName,
    color: options.userColor,
  });
  awareness.setLocalState(localAwarenessUserState(user));

  // Stage-2 auto-wiring: when the caller passes NO explicit network config
  // (the editors' plain `{docId}` form), resolve the app's own ticketed
  // websocket transport — the docId doubles as the server room (a page id).
  // Headless/test environments have no socket origin and stay local-only;
  // explicit options always win untouched.
  const autoConfigure =
    options.roomName === undefined &&
    options.providerUrl === undefined &&
    options.websocketProviderFactory === undefined;
  const autoUrl = autoConfigure ? collabSocketOrigin() : null;
  const effectiveRoomName = options.roomName ?? (autoUrl ? options.docId : null);
  const effectiveProviderUrl = options.providerUrl ?? autoUrl;
  const effectiveFactory =
    options.websocketProviderFactory ??
    (autoUrl ? ticketedWebsocketProviderFactory : undefined);

  const roomName = effectiveRoomName ?? options.docId;
  const kind = providerKindFor({
    roomName: effectiveRoomName,
    providerUrl: effectiveProviderUrl,
  });

  let provider: CollabNetworkProvider | null = null;

  if (kind === "websocket" && effectiveFactory && effectiveProviderUrl) {
    provider = effectiveFactory({
      url: effectiveProviderUrl,
      roomName,
      doc: ydoc,
      awareness,
      ...(options.pageJsonCausalModel ? { pageJsonCausalModel: true } : {}),
      ...(options.pageJsonCausalEpoch ? { pageJsonCausalEpoch: options.pageJsonCausalEpoch } : {}),
      ...(options.pageJsonTextModel ? { pageJsonTextModel: true } : {}),
    });
  } else if (kind === "webrtc") {
    provider = createWebrtcProvider({
      roomName,
      doc: ydoc,
      awareness,
      signaling: options.webrtcSignaling,
    });
  }
  // `kind === "local"` (or websocket without a factory) leaves provider null:
  // local-only single-user session, editor still fully functional.

  let destroyed = false;
  const destroy = () => {
    if (destroyed) return;
    destroyed = true;
    try {
      const guard = identityGuards.get(awareness);
      if (guard) {
        awareness.off("update", guard);
        identityGuards.delete(awareness);
      }
      provider?.destroy();
    } finally {
      awareness.destroy();
      ydoc.destroy();
    }
  };

  // Networked sessions publish the user's REAL nickname + accent; peers
  // otherwise render the caller's localized "you" placeholder.
  //
  // This is ENFORCED, not applied once. TipTap's CollaborationCaret re-stamps
  // the value it was CONFIGURED with every time its ProseMirror plugin view
  // initializes — which happens on every editor rebuild, including the
  // keyed remount that fires when a PEER saves. A one-shot upgrade therefore
  // got overwritten with "you" at exactly the moment collaboration was
  // happening. The guard below re-asserts the resolved identity whenever the
  // local awareness user drifts away from it (two writes, then it matches
  // and stops — no loop).
  if (provider && typeof window !== "undefined") {
    let enforced: CollabUser | null = null;
    const reassert = () => {
      if (destroyed || !enforced) return;
      const current = awareness.getLocalState()?.user as
        | { id?: unknown; name?: unknown; color?: unknown }
        | undefined;
      if (
        current?.id === enforced.id &&
        current?.name === enforced.name &&
        current?.color === enforced.color
      ) {
        return;
      }
      setLocalCollabUser(awareness, enforced);
    };
    const onAwareness = ({ updated }: { updated: number[] }) => {
      if (updated.includes(awareness.clientID)) reassert();
    };
    void fetchSelfCollabIdentity().then((self) => {
      if (destroyed || (!self.name && !self.color)) return;
      enforced = sanitizeCollabUser({
        // Publishing WHO, not just what to draw: without this every one of
        // this person's connections looks like a different collaborator.
        id: self.id,
        name: self.name ?? options.userName,
        color: self.color ?? options.userColor,
      });
      reassert();
      awareness.on("update", onAwareness);
      identityGuards.set(awareness, onAwareness);
    });
  }

  return {
    ydoc,
    awareness,
    provider,
    kind: provider ? kind : "local",
    network: provider?.network ?? null,
    destroy,
  };
}

/**
 * Update the local collaborator's presence (id + name + color) on a live
 * session. The values are re-sanitized so a caller cannot inject an unsafe
 * color/name, or an unbounded id.
 */
export function setLocalCollabUser(
  awareness: Awareness,
  input: { id?: unknown; name?: unknown; color?: unknown },
): void {
  awareness.setLocalStateField("user", sanitizeCollabUser(input));
  // Which tab this connection is: a tab that closed says so on this device,
  // and a view opening after it does not wait out its lingering presence
  // before setting the room to the stored page (autorecover-store).
  awareness.setLocalStateField("tab", autoRecoverTabId());
}
