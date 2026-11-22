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
  | "onRequestReferenceData"
  | "onRequestCreateNew"
  | "onRequestSharingSettings"
  | "onMakeActionLink"
  | "onRequestReferenceSource"
  | "onRequestOpen"
  | "onRequestUsers"
  | "onRequestSendNotify";

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
  /** customization.about (+ customer, white label). */
  readonly about: OnlyofficeAbout;
  /** customization.hideRulers / hideNotes / hideRightMenu (null: not said). */
  readonly hideRulers: boolean | null;
  readonly hideNotes: boolean | null;
  readonly hideRightMenu: boolean | null;
  /** customization.unit (location / region before it), null: the editors' own. */
  readonly unit: "cm" | "pt" | "inch" | null;
  /** customization.font's size (px), white label. */
  readonly fontSize: number | null;
  readonly anonymous: { readonly request: boolean; readonly label: string | null };
  readonly features: OnlyofficeFeatures;
  /** customization.layout's switches that are off (dotted keys, e.g. "statusBar", "toolbar.draw"). */
  readonly hidden: readonly string[];
  /** createUrl / onRequestCreateNew, templates, recent. */
  readonly create: { readonly url: string | null; readonly event: boolean };
  readonly templates: readonly { readonly title: string | null; readonly url: string; readonly image: string | null }[];
  readonly recent: readonly { readonly title: string | null; readonly url: string; readonly folder: string | null }[];
  readonly coEditing: { readonly mode: "fast" | "strict"; readonly change: boolean };
  /** type "embedded": the embedded viewer's buttons and where its bar sits. */
  readonly embedded: OnlyofficeEmbedded | null;
  /** editorConfig.actionLink: where the file opens (a bookmark, a comment). */
  readonly actionLink: { readonly type: "bookmark" | "comment"; readonly data: string } | null;
   
  readonly info: OnlyofficeInfo;
}

export interface OnlyofficeAbout {
  readonly visible: boolean;
  readonly customer: {
    readonly name: string | null;
    readonly address: string | null;
    readonly mail: string | null;
    readonly phone: string | null;
    readonly www: string | null;
    readonly info: string | null;
    readonly logo: string | null;
    readonly logoDark: string | null;
  } | null;
}

/** A setting the integration fixes at open (`mode`) and may let the person change (`change`). */
export interface OnlyofficeSetting<T> {
  readonly mode: T | null;
  readonly change: boolean;
}

export interface OnlyofficeFeatures {
  readonly spellcheck: OnlyofficeSetting<boolean> | null;
  readonly roles: boolean;
  readonly tabStyle: OnlyofficeSetting<"fill" | "line"> | null;
  readonly tabBackground: OnlyofficeSetting<"header" | "toolbar"> | null;
}

export interface OnlyofficeEmbedded {
  readonly embedUrl: string | null;
  readonly fullscreenUrl: string | null;
  readonly saveUrl: string | null;
  readonly shareUrl: string | null;
  readonly toolbarDocked: "top" | "bottom";
}

export interface OnlyofficeSharing {
  readonly user: string;
  readonly permissions: string | null;
  readonly isLink: boolean;
}

export interface OnlyofficeInfo {
  readonly owner: string | null;
  readonly uploaded: string | null;
  readonly folder: string | null;
  /** null: the config names none (the Info pane lists none). */
  readonly sharingSettings: readonly OnlyofficeSharing[] | null;
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

/** A picture: an address (shown through the session's picture route) or a data: image. */
function picture(value: unknown): string | null {
  if (typeof value === "string" && /^data:image\/[a-z0-9.+-]+;base64,/i.test(value) && value.length <= 400_000) return value;
  return url(value);
}

function setting<T extends string | boolean>(value: unknown, allowed: readonly T[]): OnlyofficeSetting<T> | null {
  if (!value || typeof value !== "object") return null;
  const wire = value as Record<string, unknown>;
  return { mode: (allowed as readonly unknown[]).includes(wire.mode) ? (wire.mode as T) : null, change: wire.change !== false };
}

function record(value: unknown): Record<string, unknown> {
  return value && typeof value === "object" && !Array.isArray(value) ? (value as Record<string, unknown>) : {};
}

function readAbout(value: unknown): OnlyofficeAbout {
  const wire = record(value);
  const customer = wire.customer && typeof wire.customer === "object" ? record(wire.customer) : null;
  return {
    visible: wire.visible !== false,
    customer: customer
      ? {
          name: text(customer.name, 120),
          address: text(customer.address, 240),
          mail: text(customer.mail, 160),
          phone: text(customer.phone, 60),
          www: text(customer.www, 240),
          info: text(customer.info, 500),
          logo: picture(customer.logo),
          logoDark: picture(customer.logoDark),
        }
      : null,
  };
}

function readFeatures(value: unknown): OnlyofficeFeatures {
  const wire = record(value);
  return {
    spellcheck: setting(wire.spellcheck, [true, false] as const),
    roles: wire.roles !== false,
    tabStyle: setting(wire.tabStyle, ["fill", "line"] as const),
    tabBackground: setting(wire.tabBackground, ["header", "toolbar"] as const),
  };
}

function readEmbedded(value: unknown): OnlyofficeEmbedded | null {
  if (!value || typeof value !== "object") return null;
  const wire = record(value);
  return {
    embedUrl: url(wire.embedUrl),
    fullscreenUrl: url(wire.fullscreenUrl),
    saveUrl: url(wire.saveUrl),
    shareUrl: url(wire.shareUrl),
    toolbarDocked: wire.toolbarDocked === "bottom" ? "bottom" : "top",
  };
}

/** setSharingSettings' list (or document.info.sharingSettings). */
export function readSharing(value: unknown): OnlyofficeSharing[] | null {
  if (!Array.isArray(value)) return null;
  const out: OnlyofficeSharing[] = [];
  for (const item of value.slice(0, 50)) {
    const wire = record(item);
    const user = text(wire.user, 160);
    if (user) out.push({ user, permissions: text(wire.permissions, 80), isLink: wire.isLink === true });
  }
  return out;
}

function readInfo(value: unknown): OnlyofficeInfo {
  const wire = record(value);
  return {
    owner: text(wire.owner, 160),
    uploaded: text(wire.uploaded, 80),
    folder: typeof wire.folder === "string" ? wire.folder.slice(0, 240) : null,
    sharingSettings: readSharing(wire.sharingSettings),
  };
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
    about: readAbout(wire.about),
    hideRulers: typeof wire.hideRulers === "boolean" ? wire.hideRulers : null,
    hideNotes: typeof wire.hideNotes === "boolean" ? wire.hideNotes : null,
    hideRightMenu: typeof wire.hideRightMenu === "boolean" ? wire.hideRightMenu : null,
    unit: wire.unit === "cm" || wire.unit === "pt" || wire.unit === "inch" ? wire.unit : null,
    fontSize: typeof record(wire.font).size === "number" ? (record(wire.font).size as number) : null,
    anonymous: { request: record(wire.anonymous).request !== false, label: text(record(wire.anonymous).label, 40) },
    features: readFeatures(wire.features),
    hidden: strings(wire.hidden),
    create: { url: url(record(wire.create).url), event: record(wire.create).event === true },
    templates: (Array.isArray(wire.templates) ? wire.templates : []).slice(0, 30).flatMap((item) => {
      const entry = record(item);
      const target = url(entry.url);
      return target ? [{ title: text(entry.title, 240), url: target, image: picture(entry.image) }] : [];
    }),
    recent: (Array.isArray(wire.recent) ? wire.recent : []).slice(0, 30).flatMap((item) => {
      const entry = record(item);
      const target = url(entry.url);
      return target ? [{ title: text(entry.title, 240), url: target, folder: text(entry.folder, 240) }] : [];
    }),
    coEditing: { mode: record(wire.coEditing).mode === "strict" ? "strict" : "fast", change: record(wire.coEditing).change !== false },
    embedded: readEmbedded(wire.embedded),
    actionLink:
      (record(wire.actionLink).type === "bookmark" || record(wire.actionLink).type === "comment") && text(record(wire.actionLink).data, 200)
        ? { type: record(wire.actionLink).type as "bookmark" | "comment", data: text(record(wire.actionLink).data, 200)! }
        : null,
    info: readInfo(wire.info),
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

/** customization.layout: whether the switch `key` (e.g. "statusBar",
 * "toolbar.file.info") is on. Outside a session, everything is. */
export function onlyofficeShows(key: string): boolean {
  return current === null || !current.hidden.includes(key);
}

/** customization.hideRulers as the ruler's state at open: Docs' document
 * editor shows it unless hidden, its presentation editor hides it unless
 * hideRulers is false. null outside a session (the editor's own default). */
export function onlyofficeRulerAtOpen(editor: "document" | "presentation"): boolean | null {
  if (!current) return null;
  if (current.hideRulers !== null) return !current.hideRulers;
  return editor === "document";
}

/** customization.hideNotes: the presentation editor's notes pane at open (null: its own). */
export function onlyofficeNotesAtOpen(): boolean | null {
  return current && current.hideNotes !== null ? !current.hideNotes : null;
}

/** features.roles false: a PDF form's role manager and role picker are not offered. */
export function onlyofficeAllowsRoles(): boolean {
  return current === null || current.features.roles;
}

// ---------------------------------------------------------------------------
// Action links (onMakeActionLink → setActionLink)
// ---------------------------------------------------------------------------

let actionLinkWaiter: ((link: string | null) => void) | null = null;

/** Whether the page makes links to a bookmark (onMakeActionLink): the
 * Bookmark dialog's Get Link is offered. */
export function onlyofficeMakesActionLinks(): boolean {
  return handlesEvent("onMakeActionLink");
}

/** Ask the page for a link that opens the file at `action` (its
 * setActionLink answers; null after 15 s without an answer). */
export function requestOnlyofficeActionLink(action: { readonly type: "bookmark" | "comment"; readonly data: string }): Promise<string | null> {
  actionLinkWaiter?.(null);
  return new Promise((resolve) => {
    const timer = window.setTimeout(() => {
      actionLinkWaiter = null;
      resolve(null);
    }, 15_000);
    actionLinkWaiter = (link) => {
      window.clearTimeout(timer);
      actionLinkWaiter = null;
      resolve(link);
    };
    emitIntegrationEvent("onMakeActionLink", { action });
  });
}

/** setActionLink's link (the integration channel hands it here). */
export function deliverOnlyofficeActionLink(link: string): void {
  actionLinkWaiter?.(typeof link === "string" && link ? link.slice(0, 4096) : null);
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
