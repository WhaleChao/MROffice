/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A deck's own embedded fonts (PowerPoint's "Embed fonts in the file",
 * `p:embeddedFontLst`), drawn by the canvas the way PowerPoint draws them.
 *
 * The server decodes each `ppt/fonts/*.fntdata` part (EOT, usually MicroType
 * Express) of the page's original package to TrueType
 * (`GET /pages/{id}/presentation-embedded-fonts`). Each face is registered on
 * the document under a family name of its own deck (`MRAI-Deck-…`), limited to
 * the characters it carries (`unicodeRange`: an embedded face is usually a
 * subset), and the deck's typeface names that family first, its own name
 * second -- so a character the subset lacks still draws in the machine's face
 * of that name. Read-only: nothing here touches the page.
 */
import { apiGet } from "@mrnote/core/web/lib/api.ts";

export type PresentationEmbeddedFontRole = "regular" | "bold" | "italic" | "boldItalic";

type EmbeddedFontRow = {
  family: string;
  role: PresentationEmbeddedFontRole;
  sha256: string;
  unicodeRange: string;
  /** The face, base64; or, on the server's print page, `url`. */
  data?: string;
  /** Where the server's deck printer hands the face to its print page
   * (`/__mrai-print-font/<sha256>`, answered by the printer's own browser). */
  url?: string;
  lineRatio?: number;
};

/** The one address a face may come from instead of its data: the deck
 * printer's own (`page_preview_render.ROUTED_PRINT_FONT_PATH`). */
const PRINT_FONT_PATH = "/__mrai-print-font/";

/** Typeface (lower case) -> the CSS family list that draws it. */
export type PresentationEmbeddedFontAliases = Readonly<Record<string, string>>;

const ROLES = new Set(["regular", "bold", "italic", "boldItalic"]);
const SAFE_FAMILY = /^[^"\\<>;{}\u0000-\u001f]{1,200}$/u;
const UNICODE_RANGE = /^U\+[0-9A-F]{1,6}(-[0-9A-F]{1,6})?(,U\+[0-9A-F]{1,6}(-[0-9A-F]{1,6})?)*$/;

const loadedFaces = new Map<string, Promise<boolean>>();
const pageAliases = new Map<string, Promise<PresentationEmbeddedFontAliases>>();

function readRows(value: unknown): EmbeddedFontRow[] {
  const fonts = (value as { fonts?: unknown } | null)?.fonts;
  if (!Array.isArray(fonts) || fonts.length > 128) return [];
  return fonts.filter((row): row is EmbeddedFontRow => {
    const record = row as Record<string, unknown> | null;
    return (
      !!record &&
      typeof record.family === "string" &&
      SAFE_FAMILY.test(record.family) &&
      typeof record.role === "string" &&
      ROLES.has(record.role) &&
      typeof record.sha256 === "string" &&
      /^[a-f0-9]{64}$/.test(record.sha256) &&
      typeof record.unicodeRange === "string" &&
      UNICODE_RANGE.test(record.unicodeRange) &&
      (typeof record.data === "string" || record.url === `${PRINT_FONT_PATH}${record.sha256}`)
    );
  });
}

function decodeBase64(data: string): ArrayBuffer {
  const binary = atob(data);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i += 1) bytes[i] = binary.charCodeAt(i);
  return bytes.buffer;
}

/** A face's bytes: its data, or what the deck printer answers at its address. */
async function faceBytes(row: EmbeddedFontRow): Promise<ArrayBuffer> {
  if (typeof row.data === "string") return decodeBase64(row.data);
  const answer = await fetch(row.url as string);
  if (!answer.ok) throw new Error(`embedded face ${row.sha256} not handed: ${answer.status}`);
  return answer.arrayBuffer();
}

async function sha256Hex(buffer: ArrayBuffer): Promise<string> {
  const digest = await crypto.subtle.digest("SHA-256", buffer);
  return [...new Uint8Array(digest)].map((byte) => byte.toString(16).padStart(2, "0")).join("");
}

/** The family a deck's embedded typeface is registered under. */
export function presentationEmbeddedFamily(typeface: string, digests: readonly string[]): string {
  let hash = 0x811c9dc5;
  for (const char of `${typeface.toLowerCase()}|${[...digests].sort().join(",")}`) {
    hash = Math.imul(hash ^ char.charCodeAt(0), 0x01000193) >>> 0;
  }
  return `MRAI-Deck-${hash.toString(16).padStart(8, "0")}`;
}

/** The CSS family list a typeface draws with: the deck's face, then its own name. */
export function presentationEmbeddedFamilyCss(alias: string, typeface: string): string {
  return `"${alias}", "${typeface}"`;
}

/** A face handed by address that could not be drawn (not answered, or not the
 * bytes its digest names): counted for the server's deck printer, which then
 * prints again with every face inside the deck instead of in a substitute. */
function noteFaceByAddressFailed(row: EmbeddedFontRow): void {
  if (row.url === undefined || typeof window === "undefined") return;
  const held = window as { __MRAI_DECK_PRINT_FACES_FAILED?: number };
  held.__MRAI_DECK_PRINT_FACES_FAILED = (held.__MRAI_DECK_PRINT_FACES_FAILED ?? 0) + 1;
}

function registerFace(alias: string, row: EmbeddedFontRow): Promise<boolean> {
  const key = `${alias}|${row.role}|${row.sha256}`;
  let loaded = loadedFaces.get(key);
  if (!loaded) {
    loaded = (async () => {
      if (typeof FontFace === "undefined" || typeof document === "undefined") return false;
      const buffer = await faceBytes(row).catch((error: unknown) => {
        noteFaceByAddressFailed(row);
        throw error;
      });
      if ((await sha256Hex(buffer)) !== row.sha256) {
        noteFaceByAddressFailed(row);
        return false;
      }
      const face = new FontFace(alias, buffer, {
        weight: row.role === "bold" || row.role === "boldItalic" ? "700" : "400",
        style: row.role === "italic" || row.role === "boldItalic" ? "italic" : "normal",
        unicodeRange: row.unicodeRange,
      });
      await face.load();
      document.fonts.add(face);
      return true;
    })().catch(() => false);
    loadedFaces.set(key, loaded);
  }
  return loaded;
}












export function presentationEmbeddedRegularSentinels(
  roles: readonly PresentationEmbeddedFontRole[],
): readonly ("normal" | "italic")[] {
  if (!roles.length || roles.some((role) => role === "regular" || role === "italic")) return [];
  return ["normal", "italic"];
}

function registerSentinel(alias: string, style: "normal" | "italic", row: EmbeddedFontRow): Promise<boolean> {
  const key = `${alias}|sentinel-${style}|${row.sha256}`;
  let loaded = loadedFaces.get(key);
  if (!loaded) {
    loaded = (async () => {
      if (typeof FontFace === "undefined" || typeof document === "undefined") return false;
      const face = new FontFace(alias, await faceBytes(row), { weight: "400", style, unicodeRange: "U+FFFF" });
      document.fonts.add(face);
      return true;
    })().catch(() => false);
    loadedFaces.set(key, loaded);
  }
  return loaded;
}

/**
 * Register the embedded faces an answer of the page's
 * `presentation-embedded-fonts` route lists (its `{ fonts: [...] }` body);
 * the aliases of the typefaces that are ready. The server's print page is
 * handed that answer whole, as it may reach the Web app alone.
 */
export async function registerPresentationEmbeddedFonts(body: unknown): Promise<PresentationEmbeddedFontAliases> {
  const rows = readRows(body);
  const byTypeface = new Map<string, EmbeddedFontRow[]>();
  for (const row of rows) {
    const key = row.family.toLowerCase();
    byTypeface.set(key, [...(byTypeface.get(key) ?? []), row]);
  }
  const aliases: Record<string, string> = Object.create(null);
  await Promise.all(
    [...byTypeface].map(async ([key, faces]) => {
      const alias = presentationEmbeddedFamily(faces[0].family, faces.map((face) => face.sha256));
      const ready = await Promise.all(faces.map((face) => registerFace(alias, face)));
      if (!ready.some(Boolean)) return;
      await Promise.all(
        presentationEmbeddedRegularSentinels(faces.map((face) => face.role)).map((style) =>
          registerSentinel(alias, style, faces[0]),
        ),
      );
      aliases[key] = presentationEmbeddedFamilyCss(alias, faces[0].family);
    }),
  );
  return Object.freeze(aliases);
}

/** How long a page asks again for faces the server is still unpacking
 * (`{ pending: true }`: the import's own unpacking is under way), and how often. */
const PENDING_RETRY_MS = [2_000, 3_000, 5_000, 8_000];
const PENDING_GIVE_UP_MS = 10 * 60_000;

/**
 * The page's `presentation-embedded-fonts` answer, asked for again while the
 * server answers that another process is still unpacking the faces: they are
 * unpacked once, and the page draws stand-in faces meanwhile. `maxWaitMs`: the
 * longest it asks again (the answer is then `{ fonts: [] }`).
 */
export async function readPresentationEmbeddedFontsAnswer(
  pageId: string,
  { maxWaitMs = PENDING_GIVE_UP_MS }: { maxWaitMs?: number } = {},
): Promise<unknown> {
  const started = Date.now();
  for (let attempt = 0; ; attempt += 1) {
    const body = await apiGet<unknown>(
      `/api/v1/pages/${encodeURIComponent(pageId)}/presentation-embedded-fonts`,
    );
    if ((body as { pending?: unknown } | null)?.pending !== true) return body;
    const delay = PENDING_RETRY_MS[Math.min(attempt, PENDING_RETRY_MS.length - 1)];
    if (Date.now() - started + delay > maxWaitMs) return { fonts: [] };
    await new Promise((resolve) => setTimeout(resolve, delay));
  }
}

/** Load (once per page) and register a page's embedded faces; the aliases of
 * the typefaces that are ready. */
export function loadPresentationEmbeddedFonts(pageId: string): Promise<PresentationEmbeddedFontAliases> {
  let pending = pageAliases.get(pageId);
  if (!pending) {
    pending = (async () => {
      const body = await readPresentationEmbeddedFontsAnswer(pageId);
      return registerPresentationEmbeddedFonts(body);
    })().catch((error) => {
      pageAliases.delete(pageId);
      throw error;
    });
    pageAliases.set(pageId, pending);
  }
  return pending;
}
