"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { apiGet, apiPost } from "./api";
import { subscribeAuthState } from "./auth-state";
import {
  avatarFromWire,
  normalizeAccentHex,
  type UserAvatarSource,
  type UserAvatarWire,
} from "./user-avatar";














/** The server's cap on mentioned people per comment. */
export const MAX_MENTIONS_PER_COMMENT = 20;
/** The inbox excerpt the server keeps (it trims again). */
export const MENTION_EXCERPT_MAX = 280;
/** Longest text typed after "@" that still counts as a search. */
export const MENTION_QUERY_MAX = 40;
/** The roster request size; fewer rows back means the roster is complete. */
export const MENTION_ROSTER_LIMIT = 50;
const CACHE_TTL_MS = 60_000;

export interface MentionPerson {
  userId: string;
  displayName: string;
}

export interface MentionCandidate extends MentionPerson {
  /** False: on the notebook but cannot open this page (no notification). */
  canView: boolean;
  email: string | null;
  accentColor: string | null;
  avatar: UserAvatarSource | null;
}

export interface MentionSpan {
  /** UTF-16 offset of the "@". */
  start: number;
  /** UTF-16 offset just after the name. */
  end: number;
  displayName: string;
  /** Null when two people on the roster share this exact name. */
  userId: string | null;
}

export interface MentionTrigger {
  /** Offset of the "@" that opened the search. */
  start: number;
  /** What was typed after it, up to the caret. */
  query: string;
}

// --- text rules ----------------------------------------------------------------

const CJK = /[\p{Script=Han}\p{Script=Hiragana}\p{Script=Katakana}\p{Script=Hangul}]/u;
const BREAK = /[\s\p{P}\p{S}]/u;

/** May an "@" right after `previous` start a mention? (Not inside a@b.) */
function opensMention(previous: string | undefined): boolean {
  return previous === undefined || BREAK.test(previous) || CJK.test(previous);
}

/** Does a mention end cleanly before `next`? */
function closesMention(next: string | undefined): boolean {
  return next === undefined || BREAK.test(next);
}

function codePointBefore(text: string, index: number): string | undefined {
  if (index <= 0) return undefined;
  const low = text.charCodeAt(index - 1);
  if (low >= 0xdc00 && low <= 0xdfff && index >= 2) {
    return text.slice(index - 2, index);
  }
  return text[index - 1];
}

function codePointAt(text: string, index: number): string | undefined {
  if (index >= text.length) return undefined;
  const code = text.codePointAt(index);
  return code === undefined ? undefined : String.fromCodePoint(code);
}

function cleanName(value: string): string {
  return value.normalize("NFC").trim();
}

/**
 * Every mention in `text`: "@" at a word start followed by the exact display
 * name of someone on `roster` (the page's people plus the viewer), ending at
 * the end of the text, a space, punctuation or a symbol. Longest name wins.
 * `picks` (people chosen in the composer) override the roster for their name.
 */
export function resolveCommentMentions(
  text: string,
  roster: readonly MentionPerson[],
  picks: readonly MentionPerson[] = [],
): MentionSpan[] {
  if (!text || !text.includes("@")) return [];
  const byName = new Map<string, Set<string>>();
  for (const person of roster) {
    const name = cleanName(person.displayName);
    if (!name || !person.userId) continue;
    const ids = byName.get(name) ?? new Set<string>();
    ids.add(person.userId);
    byName.set(name, ids);
  }
  for (const person of picks) {
    const name = cleanName(person.displayName);
    if (name && person.userId) byName.set(name, new Set([person.userId]));
  }
  if (byName.size === 0) return [];
  const names = [...byName.keys()].sort((left, right) => right.length - left.length);
  const source = text;
  const spans: MentionSpan[] = [];
  let index = source.indexOf("@");
  while (index !== -1) {
    let matched: MentionSpan | null = null;
    if (opensMention(codePointBefore(source, index))) {
      for (const name of names) {
        if (!source.startsWith(name, index + 1)) continue;
        const end = index + 1 + name.length;
        if (!closesMention(codePointAt(source, end))) continue;
        const ids = byName.get(name)!;
        matched = {
          start: index,
          end,
          displayName: name,
          userId: ids.size === 1 ? [...ids][0] : null,
        };
        break;
      }
    }
    if (matched) {
      spans.push(matched);
      index = source.indexOf("@", matched.end);
    } else {
      index = source.indexOf("@", index + 1);
    }
  }
  return spans;
}

/** The distinct people mentioned, in order, at most the server's cap. */
export function mentionedUserIds(spans: readonly MentionSpan[]): string[] {
  const ids: string[] = [];
  for (const span of spans) {
    if (span.userId && !ids.includes(span.userId)) ids.push(span.userId);
  }
  return ids.slice(0, MAX_MENTIONS_PER_COMMENT);
}

/**
 * People to notify for a posted text: everyone it mentions, minus everyone
 * the previous version already mentioned (an edit only notifies additions).
 * A name that was ambiguous before counts as mentioned by everyone with it.
 */
export function newlyMentionedUserIds(
  text: string,
  previousText: string | null | undefined,
  roster: readonly MentionPerson[],
  picks: readonly MentionPerson[] = [],
): string[] {
  const now = mentionedUserIds(resolveCommentMentions(text, roster, picks));
  if (!previousText) return now;
  const before = new Set<string>();
  const byName = new Map<string, string[]>();
  for (const person of [...roster, ...picks]) {
    const name = cleanName(person.displayName);
    byName.set(name, [...(byName.get(name) ?? []), person.userId]);
  }
  for (const span of resolveCommentMentions(previousText, roster)) {
    for (const id of span.userId ? [span.userId] : (byName.get(span.displayName) ?? [])) {
      before.add(id);
    }
  }
  return now.filter((id) => !before.has(id));
}

/**
 * The "@query" being typed at the caret, or null. The query may hold a name's
 * spaces (at most two), never a line break, and is at most 40 characters.
 * An "@" that already begins a finished mention (`spans`) whose name the caret
 * has moved past does not reopen the search.
 */
export function findMentionTrigger(
  text: string,
  caret: number,
  spans: readonly MentionSpan[] = [],
): MentionTrigger | null {
  if (caret < 1 || caret > text.length) return null;
  const floor = Math.max(0, caret - MENTION_QUERY_MAX - 1);
  for (let index = caret - 1; index >= floor; index -= 1) {
    const char = text[index];
    if (char === "\n" || char === "\r") return null;
    if (char !== "@") continue;
    if (!opensMention(codePointBefore(text, index))) return null;
    const finished = spans.find((span) => span.start === index);
    if (finished && caret > finished.end) return null;
    const query = text.slice(index + 1, caret);
    if (/^\s/.test(query) || /\s{2}/.test(query)) return null;
    if ((query.match(/\s/g) ?? []).length > 2) return null;
    return { start: index, query };
  }
  return null;
}

/** Replace the "@query" at the caret with "@Name " and say where the caret goes. */
export function applyMentionPick(
  text: string,
  trigger: MentionTrigger,
  caret: number,
  person: MentionPerson,
): { text: string; caret: number } {
  const before = text.slice(0, trigger.start);
  const after = text.slice(caret);
  const label = `@${cleanName(person.displayName)}`;
  if (/^\s/.test(after)) {
    return { text: `${before}${label}${after}`, caret: before.length + label.length + 1 };
  }
  return {
    text: `${before}${label} ${after}`,
    caret: before.length + label.length + 1,
  };
}

/** Backspace right after a mention removes the whole mention. */
export function removeMentionEndingAt(
  text: string,
  caret: number,
  spans: readonly MentionSpan[],
): { text: string; caret: number } | null {
  const span = spans.find((candidate) => candidate.end === caret);
  if (!span) return null;
  return {
    text: text.slice(0, span.start) + text.slice(span.end),
    caret: span.start,
  };
}

/** Does a candidate match what was typed after "@"? */
export function candidateMatches(
  candidate: MentionCandidate,
  query: string,
): boolean {
  const needle = query.trim().toLocaleLowerCase();
  if (!needle) return true;
  return (
    candidate.displayName.toLocaleLowerCase().includes(needle) ||
    (candidate.email?.toLocaleLowerCase().includes(needle) ?? false)
  );
}

/** One plain line for the inbox (the server trims to 280 as well). */
export function mentionExcerpt(text: string): string {
  const collapsed = text.replace(/\s+/g, " ").trim();
  if (collapsed.length <= MENTION_EXCERPT_MAX) return collapsed;
  return `${collapsed.slice(0, MENTION_EXCERPT_MAX - 1).trimEnd()}…`;
}

function randomRef(): string {
  try {
    return crypto.randomUUID().replace(/-/g, "").slice(0, 12);
  } catch {
    return Math.random().toString(36).slice(2, 14);
  }
}

/**
 * The `comment_id` the producer API dedupes on: `<thread>:<message>`. The
 * thread part locates the comment on the page; the message part keeps each
 * posted message (or the host's stable message id) separate, so a later reply
 * that mentions the same person notifies again while a retry does not.
 */
export function commentMentionId(
  threadRef: string,
  messageRef?: string | null,
): string {
  const clean = (value: string) => value.replace(/[^A-Za-z0-9_.-]/g, "_");
  const thread = clean(threadRef).slice(0, 80) || "comment";
  const message = clean(messageRef || randomRef()).slice(0, 44) || randomRef();
  return `${thread}:${message}`;
}

// --- server ------------------------------------------------------------------

type Wire = Record<string, unknown>;

function record(value: unknown): Wire | null {
  return value && typeof value === "object" && !Array.isArray(value)
    ? (value as Wire)
    : null;
}

function text(value: unknown): string | null {
  return typeof value === "string" ? value : null;
}

function parseCandidate(value: unknown): MentionCandidate | null {
  const wire = record(value);
  const userId = text(wire?.user_id)?.trim();
  const displayName = cleanName(text(wire?.display_name) ?? "");
  if (!wire || !userId || !displayName) return null;
  return {
    userId,
    displayName,
    canView: wire.can_view !== false,
    email: text(wire.email),
    accentColor: normalizeAccentHex(wire.accent_color),
    avatar: avatarFromWire(wire as UserAvatarWire),
  };
}

function scopedInit(
  workspaceId: string | null | undefined,
  init?: RequestInit,
): RequestInit | undefined {
  if (!workspaceId) return init;
  const headers = new Headers(init?.headers);
  headers.set("X-Workspace-ID", workspaceId);
  return { ...init, headers };
}

export interface MentionCandidatePage {
  items: MentionCandidate[];
  /** True when the server had no more people than it returned. */
  complete: boolean;
}

/**
 * Where an embedded editor finds the people a comment may mention, and tells
 * of a comment that names them (components/embed/EmbedFrame): the partner's
 * own directory, through its page (embed.js `onRequestUsers`), instead of the
 * page's MR Office roster. Unset outside an embed.
 */
export interface MentionDirectory {
  search(query: string, limit: number): Promise<MentionCandidate[]>;
  notify(pageId: string, body: { commentId: string; userIds: readonly string[]; excerpt: string }): Promise<CommentMentionReceipt>;
}

let directory: MentionDirectory | null = null;

export function setMentionDirectory(next: MentionDirectory | null): void {
  directory = next;
  candidateCache.clear();
}

export async function fetchMentionCandidates(
  pageId: string,
  query: string,
  options: { workspaceId?: string | null; limit?: number; signal?: AbortSignal } = {},
): Promise<MentionCandidatePage> {
  const limit = options.limit ?? MENTION_ROSTER_LIMIT;
  if (directory) {
    const items = (await directory.search(query.trim().slice(0, 80), limit)).slice(0, limit);
    return { items, complete: items.length < limit };
  }
  const params = new URLSearchParams({ limit: String(limit) });
  const q = query.trim().slice(0, 80);
  if (q) params.set("q", q);
  const data = await apiGet<{ items?: unknown }>(
    `/api/v1/pages/${encodeURIComponent(pageId)}/mention-candidates?${params}`,
    scopedInit(options.workspaceId, { signal: options.signal }),
    // A lookup behind a comment box never signs anyone out by itself.
    { suppressUnauthorizedHandling: true },
  );
  const items = Array.isArray(data?.items)
    ? data.items
        .map(parseCandidate)
        .filter((item): item is MentionCandidate => item !== null)
    : [];
  return { items, complete: items.length < limit };
}

export interface CommentMentionReceipt {
  notified: string[];
  needsAccess: MentionCandidate[];
  notebookId: string | null;
  /** The author owns the page and may share it with `needsAccess`. */
  canShare: boolean;
}

export async function postCommentMentions(
  pageId: string,
  body: { commentId: string; userIds: readonly string[]; excerpt: string },
  options: { workspaceId?: string | null } = {},
): Promise<CommentMentionReceipt> {
  if (directory) return directory.notify(pageId, body);
  const data = await apiPost<Wire>(
    `/api/v1/pages/${encodeURIComponent(pageId)}/comment-mentions`,
    {
      comment_id: body.commentId,
      mentioned_user_ids: [...body.userIds].slice(0, MAX_MENTIONS_PER_COMMENT),
      excerpt: mentionExcerpt(body.excerpt),
    },
    scopedInit(options.workspaceId),
  );
  const notified = Array.isArray(data?.notified)
    ? data.notified.filter((id): id is string => typeof id === "string")
    : [];
  const needsAccess = Array.isArray(data?.needs_access)
    ? data.needs_access
        .map((item) => {
          const person = parseCandidate(item);
          return person ? { ...person, canView: false } : null;
        })
        .filter((item): item is MentionCandidate => item !== null)
    : [];
  return {
    notified,
    needsAccess,
    notebookId: text(data?.notebook_id),
    canShare: data?.can_share === true,
  };
}

// --- per-page cache ------------------------------------------------------------

interface CacheEntry {
  at: number;
  promise: Promise<MentionCandidatePage>;
}

const candidateCache = new Map<string, CacheEntry>();
let authSubscribed = false;

function cacheKey(workspaceId: string | null | undefined, pageId: string, query: string) {
  return `${workspaceId ?? ""}\u0000${pageId}\u0000${query.trim().toLocaleLowerCase()}`;
}

/** Drop every cached candidate list (sign-out, account switch, tests). */
export function resetMentionCandidateCache(): void {
  candidateCache.clear();
}

/**
 * Candidates for `query` on one page, cached for a minute per workspace +
 * page + query and shared by every composer and thread of that page. A failed
 * request is not cached.
 */
export function cachedMentionCandidates(
  pageId: string,
  query: string,
  workspaceId?: string | null,
): Promise<MentionCandidatePage> {
  if (!authSubscribed && typeof window !== "undefined") {
    authSubscribed = true;
    subscribeAuthState(resetMentionCandidateCache);
  }
  const key = cacheKey(workspaceId, pageId, query);
  const now = Date.now();
  const hit = candidateCache.get(key);
  if (hit && now - hit.at < CACHE_TTL_MS) return hit.promise;
  const promise = fetchMentionCandidates(pageId, query, { workspaceId });
  candidateCache.set(key, { at: now, promise });
  promise.catch(() => {
    if (candidateCache.get(key)?.promise === promise) candidateCache.delete(key);
  });
  return promise;
}
