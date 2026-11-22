/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { onlyofficeAllowsComment, onlyofficeAllowsReview, onlyofficeRestrictsComments } from "./onlyoffice-frame";

export type EmbedSessionMode = "edit" | "view" | "comment" | "review" | "fillForms";

export interface EmbedPermissions {
  readonly download: boolean;
  readonly print: boolean;
  readonly copy: boolean;
  /** The formats File > Export may write (lower case, no dot); null: every one. */
  readonly export: readonly string[] | null;
  readonly acceptRejectChanges: boolean;
  readonly comment: boolean;
  readonly manageComments: boolean;
  /** The Code editor may run its file (the token's features.run on an app
   * that may run code; app/services/embed/access.code_run_granted). */
  readonly run: boolean;
}

export interface EmbedWatermark {
  readonly text: string;
  readonly opacity: number;
  readonly angle: number;
  readonly fontSize: number;
  readonly color: string;
  readonly inExports: boolean;
}

export interface EmbedAccess {
  /** The token's mode. */
  readonly mode: EmbedSessionMode;
  /** The mode the file opens in (a workbook, deck or PDF opens review as
   * comment, and fillForms on anything but a PDF read only). */
  readonly effectiveMode: EmbedSessionMode;
  readonly permissions: EmbedPermissions;
  readonly watermark: EmbedWatermark | null;
  /** The session's account only reads the page; what it may change goes
   * through the session's own routes (a PDF's form: lib/embed/form-session;
   * apps/api app/services/embed/forms.py). */
  readonly mediated: boolean;
  /** Who the person is on a form whose fields are assigned to roles (the
   * first of the token's user.roles), or null. */
  readonly role: string | null;
}

const MODES: readonly EmbedSessionMode[] = ["edit", "view", "comment", "review", "fillForms"];

function mode(value: unknown, fallback: EmbedSessionMode): EmbedSessionMode {
  return typeof value === "string" && (MODES as readonly string[]).includes(value) ? (value as EmbedSessionMode) : fallback;
}

function flag(value: unknown, fallback: boolean): boolean {
  return typeof value === "boolean" ? value : fallback;
}

function number(value: unknown, fallback: number, low: number, high: number): number {
  return typeof value === "number" && Number.isFinite(value) ? Math.min(high, Math.max(low, value)) : fallback;
}

/** The access a session view describes (an older API, which says only
 * `mode: edit | view`, reads as that mode with every permission). */
export function readEmbedAccess(view: unknown): EmbedAccess {
  const wire = view && typeof view === "object" ? (view as Record<string, unknown>) : {};
  const requested = mode(wire.mode, "view");
  const effective = mode(wire.effective_mode, requested);
  const given = wire.permissions && typeof wire.permissions === "object" ? (wire.permissions as Record<string, unknown>) : {};
  const exportList = Array.isArray(given.export)
    ? given.export.filter((item): item is string => typeof item === "string").map((item) => item.toLowerCase())
    : null;
  const readOnly = effective === "view";
  const permissions: EmbedPermissions = {
    download: flag(given.download, true),
    print: flag(given.print, true),
    copy: flag(given.copy, true),
    export: exportList,
    acceptRejectChanges: !readOnly && flag(given.acceptRejectChanges, effective === "edit"),
    comment: !readOnly && effective !== "fillForms" && flag(given.comment, true),
    manageComments: flag(given.manageComments, true),
    // Off unless the session says so: running costs a sandbox.
    run: given.run === true,
  };
  const mark = wire.watermark && typeof wire.watermark === "object" ? (wire.watermark as Record<string, unknown>) : null;
  const text = typeof mark?.text === "string" ? mark.text.trim().slice(0, 240) : "";
  const watermark: EmbedWatermark | null = mark && text
    ? {
        text,
        opacity: number(mark.opacity, 0.15, 0.02, 1),
        angle: number(mark.angle, -30, -90, 90),
        fontSize: number(mark.fontSize, 22, 8, 120),
        color: typeof mark.color === "string" && /^#[0-9a-fA-F]{6}$/.test(mark.color) ? mark.color : "#808080",
        inExports: mark.inExports === true,
      }
    : null;
  const user = wire.user && typeof wire.user === "object" ? (wire.user as Record<string, unknown>) : {};
  const firstRole = Array.isArray(user.roles) && typeof user.roles[0] === "string" ? user.roles[0].trim() : "";
  return {
    mode: requested,
    effectiveMode: effective,
    permissions,
    watermark,
    mediated: wire.mediated === true,
    role: firstRole || null,
  };
}

let current: EmbedAccess | null = null;
const listeners = new Set<() => void>();

export function setEmbedAccess(access: EmbedAccess | null): void {
  current = access;
  for (const listener of [...listeners]) listener();
}

export function getEmbedAccess(): EmbedAccess | null {
  return current;
}

export function subscribeEmbedAccess(listener: () => void): () => void {
  listeners.add(listener);
  return () => listeners.delete(listener);
}

/** The Word restriction a session's mode lays over the document (never
 * written into it): null for `edit` and `view` (view is the page's own
 * read-only verdict), and outside an embed. */
export function embedImposedDocumentProtection(): { enabled: true; mode: "comments" | "trackedChanges" | "forms" } | null {
  switch (current?.effectiveMode) {
    case "comment":
      return { enabled: true, mode: "comments" };
    case "review":
      return { enabled: true, mode: "trackedChanges" };
    case "fillForms":
      return { enabled: true, mode: "forms" };
    default:
      return null;
  }
}

/** Whether tracked changes may be accepted or rejected here (true outside an embed). */
/** Whether this frame fills in a PDF's form and changes nothing else of
 * the file: the PDF window then sends each fill and signature through the
 * session's form routes instead of saving the page. */
export function embedFillsPdfForm(): boolean {
  return current !== null && current.mediated && current.effectiveMode === "fillForms";
}

/** Whether this frame comments on a workbook, a deck or a PDF and changes
 * nothing else of the file (the token's comment, or review there): the
 * session's account only reads the page, so the editor sends each comment
 * gesture through the session's comment route (lib/embed/comment-session)
 * instead of saving the page. */
export function embedCommentsMediated(): boolean {
  return current !== null && current.mediated && current.effectiveMode === "comment";
}

/** The role this person fills a form's fields as, when the session names one. */
export function embedFormRole(): string | null {
  return current?.role ?? null;
}

export function embedAllowsRevisionResolution(): boolean {
  return current === null || current.permissions.acceptRejectChanges;
}

/** Whether comments may be written here (true outside an embed). */
export function embedAllowsComments(): boolean {
  return current === null || current.permissions.comment;
}

/** Whether a comment by `author` may be resolved or deleted by this person:
 * always, unless the session's `manageComments` is off and it is someone
 * else's (ONLYOFFICE's editCommentAuthorOnly / deleteCommentAuthorOnly). */
export function embedAllowsManagingComment(
  author: string | null | undefined,
  self: string | null | undefined,
  action: "edit" | "remove" = "remove",
): boolean {
  // An ONLYOFFICE config's commentGroups (lib/embed/onlyoffice-frame): the
  // groups whose comments this person may resolve (edit) or delete.
  // (Every comment at once, author null: only where no group is excluded.)
  if (author === null ? onlyofficeRestrictsComments(action) : !onlyofficeAllowsComment(action, author, self)) return false;
  if (current === null || current.permissions.manageComments) return true;
  return Boolean(author && self && author.trim() === self.trim());
}

/** Whether a comment by `author` is shown to this person (commentGroups.view). */
export function embedShowsComment(author: string | null | undefined, self: string | null | undefined): boolean {
  return onlyofficeAllowsComment("view", author, self);
}

/** Whether a tracked change by `author` may be accepted or rejected by this
 * person (the session's acceptRejectChanges, and an ONLYOFFICE config's
 * reviewGroups). */
export function embedAllowsResolvingChangeBy(author: string | null | undefined, self: string | null | undefined): boolean {
  return embedAllowsRevisionResolution() && onlyofficeAllowsReview(author, self);
}

/** Whether a file of `format` may leave the frame as an export or download. */
export function embedAllowsExport(format: string): boolean {
  if (current === null) return true;
  if (!current.permissions.download) return false;
  const clean = format.toLowerCase().replace(/^\./, "");
  if (current.watermark?.inExports && clean !== "pdf") return false;
  return current.permissions.export === null || current.permissions.export.includes(clean);
}

/** Whether the Code editor may run code here (its Run, Debug, tests and
 * console): always outside an embed; in one, only when the session grants
 * it (the server holds its runtime to the same grant: services/embed/scope). */
export function embedAllowsCodeRun(): boolean {
  return current === null || current.permissions.run;
}

export function embedAllowsPrint(): boolean {
  return current === null || current.permissions.print;
}

export function embedAllowsCopy(): boolean {
  return current === null || current.permissions.copy;
}

/** The watermark an export or print must carry, or null. */
export function embedExportWatermark(): EmbedWatermark | null {
  return current?.watermark?.inExports ? current.watermark : null;
}
