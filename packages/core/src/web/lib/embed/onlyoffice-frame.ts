/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * An embed session the server made from an ONLYOFFICE Docs API config
 * (apps/api app/services/embed/onlyoffice; the session view's `onlyoffice`):
 * what MR Office's editor shows and does for the integration's page.
 *
 * One store for the frame (components/embed/EmbedFrame sets it), read by the
 * frame's header (components/embed/OnlyofficeBar) and by the editors' shared
 * code where a setting lands:
 *
 *   - the ribbon opens folded (`customization.compactToolbar`; the shared
 *     chassis, components/office-ribbon/OfficeRibbonChassis);
 *   - the editors open at `customization.zoom`;
 *   - comments by people of other groups are read, edited and removed as
 *     `permissions.commentGroups` says (the Word editor's comments pane);
 *   - tracked changes by other groups are accepted or rejected as
 *     `permissions.reviewGroups` says;
 *   - `permissions.protect` false keeps Restrict Editing / Protect Sheet out
 *     of reach, `modifyFilter` false the shared filter, `modifyContentControl`
 *     false the content controls.
 *
 * Outside such a session nothing is set and every answer is the product's own.
 */

export type OnlyofficeEvent =
  | "onRequestHistory"
  | "onRequestHistoryData"
  | "onRequestHistoryClose"
  | "onRequestRestore"
  | "onRequestRename"
  | "onRequestSaveAs"
  | "onRequestInsertImage"
  | "onRequestSelectSpreadsheet"
  | "onRequestSelectDocument"
  | "onRequestEditRights"
  | "onRequestClose"
  | "onRequestStartFilling"
  | "onMetaChange"
  | "onRequestRefreshFile"
  | "onOutdatedVersion"
  | "onRequestReferenceData";

/** A group permission: null for every group, else the groups named ("" is
 * the people in no group; [] nobody's). */
export type GroupList = readonly string[] | null;

export interface OnlyofficeView {
  readonly events: readonly OnlyofficeEvent[];
  readonly key: string;
  readonly user: string;
  readonly goback: { readonly url: string | null; readonly text: string | null; readonly blank: boolean; readonly requestClose: boolean } | null;
  readonly close: { readonly text: string | null } | null;
  readonly help: boolean;
  readonly feedback: { readonly url: string | null } | null;
  readonly logo: { readonly image: string | null; readonly imageDark: string | null; readonly url: string | null } | null;
  readonly compactToolbar: boolean;
  readonly hideFileName: boolean;
  /** Percent, or -1 (fit the page) / -2 (fit the width). */
  readonly zoom: number | null;
  readonly review: { readonly trackChanges: boolean | null; readonly display: "markup" | "simple" | "final" | "original" | null; readonly hideDisplay: boolean; readonly showChanges: boolean };
  readonly autosave: boolean;
  /** customization.submitForm (null: not said, the Submit button shows). */
  readonly submitForm: boolean | null;
  readonly groups: {
    readonly comment?: { readonly view: GroupList; readonly edit: GroupList; readonly remove: GroupList };
    readonly review?: GroupList;
    readonly userInfo?: GroupList;
  };
  /** The person's own groups (user.group). */
  readonly group: readonly string[];
  readonly modifyFilter: boolean;
  readonly modifyContentControl: boolean;
  readonly protect: boolean;
  /** A view whose page handles onRequestEditRights. */
  readonly editRights: boolean;
  /** document.info.favorite (null: no star). */
  readonly favorite: boolean | null;
  readonly theme: "light" | "dark" | "high-contrast" | "system" | null;
  /** The key was saved and taken: opened to read (onOutdatedVersion). */
  readonly outdated: boolean;
  readonly metaVersion: number;
  /** config.type "mobile" (customization.mobile, Docs 8.2), else null:
   * `editable` -- opened to read first (forceView), an Edit button away from
   * the mode the token vouches for. */
  readonly mobile: OnlyofficeMobile | null;
  /** editorConfig.region (a BCP 47 tag): the spreadsheet editor's number
   * separators; null, the interface language's. */
  readonly region: string | null;
}

export interface OnlyofficeMobile {
  readonly forceView: boolean;
  readonly standardView: boolean;
  readonly disableForceDesktop: boolean;
  readonly info: boolean;
  readonly editable: boolean;
}

function strings(value: unknown): string[] {
  return Array.isArray(value) ? value.filter((item): item is string => typeof item === "string").slice(0, 100) : [];
}

function groupList(value: unknown): GroupList {
  return Array.isArray(value) ? strings(value) : null;
}

function text(value: unknown, max = 200): string | null {
  return typeof value === "string" && value.trim() ? value.trim().slice(0, max) : null;
}

function url(value: unknown): string | null {
  if (typeof value !== "string" || value.length > 2048) return null;
  return /^https:\/\//.test(value) || /^http:\/\/(localhost|127\.0\.0\.1)(:\d+)?\//.test(value) ? value : null;
}

/** The session view's `onlyoffice`, or null (another kind of session). */
export function readOnlyofficeView(view: unknown): OnlyofficeView | null {
  const wire = view && typeof view === "object" ? ((view as Record<string, unknown>).onlyoffice as Record<string, unknown> | undefined) : undefined;
  if (!wire || typeof wire !== "object") return null;
  const goback = wire.goback && typeof wire.goback === "object" ? (wire.goback as Record<string, unknown>) : null;
  const close = wire.close && typeof wire.close === "object" ? (wire.close as Record<string, unknown>) : null;
  const feedback = wire.feedback && typeof wire.feedback === "object" ? (wire.feedback as Record<string, unknown>) : null;
  const logo = wire.logo && typeof wire.logo === "object" ? (wire.logo as Record<string, unknown>) : null;
  const review = wire.review && typeof wire.review === "object" ? (wire.review as Record<string, unknown>) : {};
  const groups = wire.groups && typeof wire.groups === "object" ? (wire.groups as Record<string, unknown>) : {};
  const comment = groups.comment && typeof groups.comment === "object" ? (groups.comment as Record<string, unknown>) : null;
  const display = review.display;
  const theme = wire.theme;
  return {
    events: strings(wire.events) as OnlyofficeEvent[],
    key: typeof wire.key === "string" ? wire.key : "",
    user: typeof wire.user === "string" ? wire.user : "",
    goback: goback ? { url: url(goback.url), text: text(goback.text, 80), blank: goback.blank !== false, requestClose: goback.requestClose === true } : null,
    close: close ? { text: text(close.text, 80) } : null,
    help: wire.help !== false,
    feedback: feedback ? { url: url(feedback.url) } : null,
    logo: logo ? { image: url(logo.image), imageDark: url(logo.imageDark), url: url(logo.url) } : null,
    compactToolbar: wire.compactToolbar === true,
    hideFileName: wire.hideFileName === true,
    region: typeof wire.region === "string" && /^[a-z]{2,3}(?:-[A-Za-z]{4})?(?:-(?:[A-Z]{2}|\d{3}))?$/.test(wire.region) ? wire.region : null,
    zoom: typeof wire.zoom === "number" && Number.isFinite(wire.zoom) ? wire.zoom : null,
    review: {
      trackChanges: typeof review.trackChanges === "boolean" ? review.trackChanges : null,
      display: display === "markup" || display === "simple" || display === "final" || display === "original" ? display : null,
      hideDisplay: review.hideDisplay === true,
      showChanges: review.showChanges === true,
    },
    autosave: wire.autosave !== false,
    submitForm: typeof wire.submitForm === "boolean" ? wire.submitForm : null,
    groups: {
      ...(comment ? { comment: { view: groupList(comment.view), edit: groupList(comment.edit), remove: groupList(comment.remove) } } : {}),
      ...("review" in groups ? { review: groupList(groups.review) } : {}),
      ...("userInfo" in groups ? { userInfo: groupList(groups.userInfo) } : {}),
    },
    group: strings(wire.group),
    modifyFilter: wire.modifyFilter !== false,
    modifyContentControl: wire.modifyContentControl !== false,
    protect: wire.protect !== false,
    editRights: wire.editRights === true,
    favorite: typeof wire.favorite === "boolean" ? wire.favorite : null,
    theme: theme === "light" || theme === "dark" || theme === "high-contrast" || theme === "system" ? theme : null,
    outdated: wire.outdated === true,
    metaVersion: typeof wire.meta_version === "number" ? wire.meta_version : 0,
    mobile: readMobile(wire.mobile),
  };
}

function readMobile(value: unknown): OnlyofficeMobile | null {
  if (!value || typeof value !== "object") return null;
  const mobile = value as Record<string, unknown>;
  return {
    forceView: mobile.forceView !== false,
    standardView: mobile.standardView === true,
    disableForceDesktop: mobile.disableForceDesktop === true,
    info: mobile.info === true,
    editable: mobile.editable === true,
  };
}

let current: OnlyofficeView | null = null;
/** The groups of the people in the file, by the name the editors show
 * them under (a comment's or a change's author), for the group permissions. */
let groupsByName = new Map<string, readonly string[]>();
const listeners = new Set<() => void>();

let selfName: string | null = null;

/** The person's own name (as their comments and changes carry it). */
export function setEmbedSelfName(name: string | null): void {
  selfName = name;
}

export function embedSelfName(): string | null {
  return selfName;
}

export function setOnlyofficeView(view: OnlyofficeView | null): void {
  current = view;
  for (const listener of [...listeners]) listener();
}

export function getOnlyofficeView(): OnlyofficeView | null {
  return current;
}

export function subscribeOnlyofficeView(listener: () => void): () => void {
  listeners.add(listener);
  return () => listeners.delete(listener);
}

/** The people's groups, as the frame last heard them (EmbedFrame's presence). */
export function setOnlyofficePeopleGroups(people: readonly { readonly name?: string | null; readonly groups?: readonly string[] }[]): void {
  const next = new Map<string, readonly string[]>();
  for (const person of people) {
    if (person.name) next.set(person.name.trim(), person.groups ?? []);
  }
  groupsByName = next;
}

export function handlesEvent(name: OnlyofficeEvent): boolean {
  return current !== null && current.events.includes(name);
}

/** Whether `allowed` admits a person of `groups` (ONLYOFFICE's group lists:
 * null every group, "" in the list the people in no group). */
export function groupAllows(allowed: GroupList | undefined, groups: readonly string[]): boolean {
  if (allowed === null || allowed === undefined) return true;
  if (!groups.length) return allowed.includes("");
  return groups.some((group) => allowed.includes(group));
}

function authorGroups(author: string | null | undefined): readonly string[] | null {
  if (!author) return null;
  return groupsByName.get(author.trim()) ?? null;
}

/** commentGroups: whether this person may `action` a comment by `author`
 * (their own always; an author whose groups are unknown counts as in none). */
export function onlyofficeAllowsComment(action: "view" | "edit" | "remove", author: string | null | undefined, self: string | null | undefined): boolean {
  const rule = current?.groups.comment;
  if (!rule) return true;
  if (author && self && author.trim() === self.trim()) return true;
  return groupAllows(rule[action], authorGroups(author) ?? []);
}

/** Whether commentGroups keeps some people's comments from `action`. */
export function onlyofficeRestrictsComments(action: "view" | "edit" | "remove"): boolean {
  const rule = current?.groups.comment;
  return Boolean(rule && rule[action] !== null);
}

/** reviewGroups: whether this person may accept or reject a change by `author`. */
export function onlyofficeAllowsReview(author: string | null | undefined, self: string | null | undefined): boolean {
  if (!current || !("review" in current.groups)) return true;
  if (author && self && author.trim() === self.trim()) return true;
  return groupAllows(current.groups.review, authorGroups(author) ?? []);
}

/** customization.compactToolbar: the ribbons open folded to their tab row. */
export function onlyofficeRibbonFolded(): boolean {
  return current?.compactToolbar === true;
}

/** type "mobile" without customization.mobile.standardView: a document
 * opens reflowed to the screen (ONLYOFFICE's Mobile view; Web Layout, whose
 * phone face flows the page's placed objects into the text). */
export function onlyofficeOpensReflowed(): boolean {
  return Boolean(current?.mobile && !current.mobile.standardView);
}

/** customization.zoom as the editors' scale (1 = 100 %), or null. */
export function onlyofficeInitialZoom(): number | null {
  const zoom = current?.zoom;
  return typeof zoom === "number" && zoom >= 10 ? Math.min(5, Math.max(0.1, zoom / 100)) : null;
}

export function onlyofficeAllowsProtect(): boolean {
  return current === null || current.protect;
}

export function onlyofficeAllowsFilterChange(): boolean {
  return current === null || current.modifyFilter;
}

export function onlyofficeAllowsContentControlChange(): boolean {
  return current === null || current.modifyContentControl;
}

/** customization.submitForm false: the form's Submit is not offered. */
export function onlyofficeOffersSubmit(): boolean {
  return current === null || current.submitForm !== false;
}

// ---------------------------------------------------------------------------
// The channel to the page (integrationCall / integrationEvent)
// ---------------------------------------------------------------------------

export interface IntegrationCall {
  readonly name: string;
  readonly data: unknown;
}

type CallListener = (call: IntegrationCall) => void;
const callListeners = new Set<CallListener>();

/** The page's answer to what the editor asked (refreshHistory, insertImage…). */
export function deliverIntegrationCall(call: IntegrationCall): void {
  for (const listener of [...callListeners]) listener(call);
}

export function subscribeIntegrationCalls(listener: CallListener): () => void {
  callListeners.add(listener);
  return () => callListeners.delete(listener);
}

let emitter: ((name: string, data?: unknown) => void) | null = null;

/** EmbedFrame's post of `integrationEvent` to the page. */
export function setIntegrationEmitter(next: ((name: string, data?: unknown) => void) | null): void {
  emitter = next;
}

/** Tell the page (its ONLYOFFICE event handler) what the person asked for. */
export function emitIntegrationEvent(name: string, data?: unknown): void {
  emitter?.(name, data);
}
