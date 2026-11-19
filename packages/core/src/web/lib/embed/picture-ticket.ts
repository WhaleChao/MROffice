/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The embedded editor's picture ticket (apps/api app/services/embed/
 * sessions.py): a page's pictures are drawn from plain URLs (`<img>`, a
 * slide's background, a narration's `<audio>`) that carry no header and,
 * where the browser blocks third-party cookies, no cookie either. In an
 * embedded editor each such URL carries the session's ticket instead
 * (lib/page-picture-objects.pagePictureUrl): it reads the session's own
 * page's pictures, while the session lives, and nothing else.
 *
 * Held through globalThis, so every copy of this module (one per bundle
 * that imports it) reads the same ticket. Outside an embedded editor it is
 * never set and no URL changes.
 */

export const EMBED_PICTURE_TICKET_PARAM = "embed_ticket";

const KEY = "__mroEmbedPictureTicket";

type Holder = { [KEY]?: string | null };

export function setEmbedPictureTicket(ticket: string | null): void {
  (globalThis as Holder)[KEY] = ticket && /^[0-9A-Za-z-]{8,64}\.[A-Za-z0-9_-]{43}$/.test(ticket) ? ticket : null;
}

export function embedPictureTicket(): string | null {
  return (globalThis as Holder)[KEY] ?? null;
}

/** A file the page holds, read by its id (lib/page-attachment-upload): the
 * URL an inserted picture is drawn from until the page saves it as one of
 * its own pictures. */
const ATTACHMENT_PATH = /^\/api\/v1\/attachments\/[0-9A-Za-z-]{8,64}\/download$/;

/**
 * ``url`` with the session's picture ticket, when it is one of the page's
 * files on one of ``origins`` (the frame's and the API's) and carries none
 * yet; otherwise ``url`` as it is.
 */
export function withEmbedPictureTicket(url: string, origins: ReadonlySet<string>, base: string): string {
  const ticket = embedPictureTicket();
  if (!ticket || !url || url.startsWith("data:") || url.startsWith("blob:")) return url;
  let parsed: URL;
  try {
    parsed = new URL(url, base);
  } catch {
    return url;
  }
  if (!origins.has(parsed.origin) || !ATTACHMENT_PATH.test(parsed.pathname)) return url;
  if (parsed.searchParams.has(EMBED_PICTURE_TICKET_PARAM)) return url;
  parsed.searchParams.set(EMBED_PICTURE_TICKET_PARAM, ticket);
  // Keep a relative URL relative (the editors compare what they wrote).
  return /^[a-z][a-z0-9+.-]*:/i.test(url) ? parsed.href : `${parsed.pathname}${parsed.search}${parsed.hash}`;
}

const MEDIA_ATTRIBUTES = new Set(["src", "poster"]);

/**
 * Every picture, video or audio element of the frame gets the ticket on a
 * page file's URL as it is given one (an inserted picture is drawn from its
 * attachment before the page holds it as a picture): a browser that blocks
 * third-party cookies (Safari) sends the frame's partitioned session cookie
 * nowhere, and drew the picture broken. Returns the release.
 */
export function installEmbedMediaTickets(origins: ReadonlySet<string>): () => void {
  if (typeof window === "undefined" || typeof Element === "undefined") return () => undefined;
  const base = () => window.location.href;
  const media = (element: Element) =>
    element instanceof HTMLImageElement || element instanceof HTMLMediaElement || element instanceof HTMLSourceElement;
  const elementProto = Element.prototype;
  const setAttribute = elementProto.setAttribute;
  const heldSetAttribute = function (this: Element, name: string, value: string) {
    if (media(this) && MEDIA_ATTRIBUTES.has(String(name).toLowerCase())) {
      value = withEmbedPictureTicket(String(value), origins, base());
    }
    return setAttribute.call(this, name, value);
  };
  elementProto.setAttribute = heldSetAttribute;
  const restores: (() => void)[] = [];
  for (const proto of [HTMLImageElement.prototype, HTMLMediaElement.prototype, HTMLSourceElement.prototype]) {
    const descriptor = Object.getOwnPropertyDescriptor(proto, "src");
    if (!descriptor?.set || !descriptor.get) continue;
    const { get, set } = descriptor;
    Object.defineProperty(proto, "src", {
      ...descriptor,
      get,
      set(this: Element, value: string) {
        set.call(this, withEmbedPictureTicket(String(value), origins, base()));
      },
    });
    restores.push(() => Object.defineProperty(proto, "src", descriptor));
  }
  return () => {
    if (elementProto.setAttribute === heldSetAttribute) elementProto.setAttribute = setAttribute;
    for (const restore of restores) restore();
  };
}
