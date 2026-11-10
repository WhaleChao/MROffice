/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The worksheet's object clipboard on the system clipboard (Excel's Cut /
 * Copy / Paste of a picture, shape or other drawing object).
 *
 * Excel puts a copied object on the clipboard twice: as its own object, which
 * an Excel paste restores exactly (size, crop, effects, Alt Text), and as a
 * picture, which any other application pastes. The page keeps the exact
 * object in memory; the system clipboard receives the picture (`image/png`,
 * and an `<img>` for applications that paste HTML) whose element names the
 * copy by a token, so a paste back into the sheet can tell its own copy from
 * a picture copied anywhere else since: the token matches only while the
 * clipboard still holds this copy.
 */

/** The attribute the clipboard HTML names a copy by. */
export const SHEET_OBJECT_CLIPBOARD_ATTRIBUTE = "data-mrai-sheet-objects";

/** A fresh token for one copy. */
export function createSheetObjectClipboardToken(): string {
  const random =
    typeof crypto !== "undefined" && typeof crypto.randomUUID === "function"
      ? crypto.randomUUID()
      : `${Date.now().toString(36)}-${Math.random().toString(36).slice(2)}`;
  return `sheet-objects-${random}`;
}

function escapeAttribute(value: string): string {
  return value
    .replace(/&/g, "&amp;")
    .replace(/"/g, "&quot;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

/**
 * The HTML a copy writes: the picture as an `<img>` when there is one (so a
 * rich-text paste elsewhere receives the picture), else an element holding
 * the Alt Text. Either way the element carries the token.
 */
export function sheetObjectClipboardHtml(
  token: string,
  altText = "",
  pictureDataUrl?: string | null,
): string {
  const alt = escapeAttribute(altText.slice(0, 2_000));
  const marker = `${SHEET_OBJECT_CLIPBOARD_ATTRIBUTE}="${escapeAttribute(token)}"`;
  return pictureDataUrl && /^data:image\/png;base64,/i.test(pictureDataUrl)
    ? `<img ${marker} alt="${alt}" src="${pictureDataUrl}">`
    : `<span ${marker}>${alt}</span>`;
}

/** The copy token an HTML clipboard payload names, or null. */
export function readSheetObjectClipboardToken(html: string): string | null {
  if (!html || !html.includes(SHEET_OBJECT_CLIPBOARD_ATTRIBUTE)) return null;
  const match = new RegExp(
    `${SHEET_OBJECT_CLIPBOARD_ATTRIBUTE}\\s*=\\s*"([^"]{1,200})"`,
  ).exec(html);
  return match ? match[1] : null;
}

/**
 * Whether an HTML clipboard payload holds pictures and nothing to read: what
 * a browser or an image viewer writes beside the picture's bytes when an
 * image is copied (`<img src=…>`). Such a paste is a picture, not text.
 */
export function clipboardHtmlHoldsOnlyPictures(html: string): boolean {
  if (!html || !/<img\b/i.test(html)) return false;
  const text = html
    .replace(/<!--[\s\S]*?-->/g, "")
    .replace(/<(style|script|head|title)\b[\s\S]*?<\/\1>/gi, "")
    .replace(/<[^>]*>/g, "")
    .replace(/&nbsp;|&#160;/gi, " ")
    .trim();
  return text === "";
}

/** The first picture file a paste carries (its files, then its items). */
export function clipboardPictureFile(
  data: Pick<DataTransfer, "files" | "items"> | null | undefined,
): File | null {
  if (!data) return null;
  try {
    const file = Array.from(data.files ?? []).find((candidate) =>
      candidate.type.toLowerCase().startsWith("image/"),
    );
    if (file) return file;
    for (const item of Array.from(data.items ?? [])) {
      if (item.kind !== "file" || !item.type.toLowerCase().startsWith("image/")) {
        continue;
      }
      const found = item.getAsFile();
      if (found) return found;
    }
  } catch {
    return null;
  }
  return null;
}

/** What a paste reads off the clipboard for the object layer. */
export interface ObjectClipboardPayload {
  html: string;
  text: string;
  picture: File | null;
}

/**
 * The system clipboard as a context menu's Paste reads it (B348): the HTML
 * that names the sheet's own copy, the plain text, and a picture. Null when
 * the browser will not let the page read it, so the caller says so instead
 * of pasting a copy the clipboard no longer holds.
 */
export async function readSystemObjectClipboard(): Promise<ObjectClipboardPayload | null> {
  if (typeof navigator === "undefined" || typeof navigator.clipboard?.read !== "function") {
    return null;
  }
  try {
    const items = await navigator.clipboard.read();
    const payload: ObjectClipboardPayload = { html: "", text: "", picture: null };
    for (const item of items) {
      if (!payload.html && item.types.includes("text/html")) {
        payload.html = await (await item.getType("text/html")).text();
      }
      if (!payload.text && item.types.includes("text/plain")) {
        payload.text = await (await item.getType("text/plain")).text();
      }
      const imageType = item.types.find((type) => type.toLowerCase().startsWith("image/"));
      if (!payload.picture && imageType && typeof File !== "undefined") {
        const blob = await item.getType(imageType);
        payload.picture = new File([blob], `clipboard.${imageType.slice(6) || "png"}`, { type: imageType });
      }
    }
    return payload;
  } catch {
    return null;
  }
}

/**
 * The PNG bytes of a `data:image/png;base64,` URL (a chart's picture as the
 * chart engine renders it), or null for anything else.
 */
export function pngDataUrlBlob(dataUrl: string | null | undefined): Blob | null {
  const prefix = "data:image/png;base64,";
  if (typeof dataUrl !== "string" || !dataUrl.startsWith(prefix) || typeof Blob === "undefined") {
    return null;
  }
  try {
    const binary = atob(dataUrl.slice(prefix.length));
    const bytes = new Uint8Array(binary.length);
    for (let index = 0; index < binary.length; index += 1) {
      bytes[index] = binary.charCodeAt(index);
    }
    return bytes.length ? new Blob([bytes], { type: "image/png" }) : null;
  } catch {
    return null;
  }
}

function blobDataUrl(blob: Blob): Promise<string | null> {
  return new Promise((resolve) => {
    if (typeof FileReader === "undefined") {
      resolve(null);
      return;
    }
    const reader = new FileReader();
    reader.onerror = () => resolve(null);
    reader.onload = () =>
      resolve(typeof reader.result === "string" ? reader.result : null);
    reader.readAsDataURL(blob);
  });
}

/**
 * Put one copy on the system clipboard: the token's HTML always, the picture
 * as `image/png` (and inside the HTML) when there is one. The item is built
 * at once -- inside the keystroke or menu click that asked, as Safari
 * requires -- with the picture still being drawn; a clipboard that refuses
 * the picture is given the HTML alone, so a paste back into the sheet still
 * finds its copy. Never throws; resolves whether anything was written.
 */
export async function writeSheetObjectClipboard(
  clipboard: Pick<Clipboard, "write"> | undefined,
  token: string,
  altText: string,
  picture: Promise<Blob | null> | null,
): Promise<boolean> {
  if (!clipboard || typeof clipboard.write !== "function") return false;
  const Item = (globalThis as { ClipboardItem?: typeof ClipboardItem })
    .ClipboardItem;
  if (typeof Item !== "function" || typeof Blob === "undefined") return false;
  const htmlBlob = (dataUrl?: string | null) =>
    new Blob([sheetObjectClipboardHtml(token, altText, dataUrl)], {
      type: "text/html",
    });
  const write = async (withPicture: boolean): Promise<boolean> => {
    try {
      const png = withPicture && picture
        ? picture.then((blob) => {
            if (!blob) throw new Error("sheet_object_clipboard_picture_unavailable");
            return blob;
          })
        : null;
      const html = png
        ? png.then(async (blob) => htmlBlob(await blobDataUrl(blob)))
        : null;
      // A clipboard that gives up early must not leave a rejection unheard.
      png?.catch(() => undefined);
      html?.catch(() => undefined);
      const data: Record<string, Blob | Promise<Blob>> =
        png && html
          ? { "image/png": png, "text/html": html }
          : { "text/html": htmlBlob() };
      await clipboard.write([new Item(data)]);
      return true;
    } catch {
      return false;
    }
  };
  if (picture && (await write(true))) return true;
  return write(false);
}
