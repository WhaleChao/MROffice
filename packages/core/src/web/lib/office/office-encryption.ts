/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Password-encrypted Office files: open one Word encrypted, save one Word opens.
 *
 * Word's File > Info > Protect Document > Encrypt with Password writes the
 * package into an OLE compound file ([MS-OFFCRYPTO] Agile Encryption). The
 * browser recognises such a file by its OLE signature under an Office name,
 * asks for the password in its own dialog (the workspace's
 * `OfficePasswordHost`), and sends the file with the password in the request
 * BODY to `POST /api/v1/office/encryption/decrypt`
 * (`app.routers.office_encryption`); the server answers with the package,
 * which the import takes on as if the person had dropped it unencrypted.
 * An export of a document marked "Encrypt with Password" goes the other way
 * through `.../encrypt` before it is saved.
 *
 * The password is never put in a URL, a header, a log, local or session
 * storage, or the page: it lives in this module's memory for the open
 * session only -- as Word keeps an open document's password -- so the
 * document's next export is encrypted with it again, and is forgotten with
 * the tab, at sign-out, or at once when the person removes it.
 *
 * Every file of a page opened with its password is written encrypted
 * (`resolvePageExportPassword`, which every export goes through), or refused:
 * the server marks the page and refuses a plain Word file to any caller that
 * has not said it encrypts it (`office_encryption_required`).
 */

import { apiFetchRaw, apiGet, apiPost, isApiRequestError } from "../api";
import { heldPagePassword, holdPagePassword } from "./office-password-store";

export const OFFICE_DECRYPT_PATH = "/api/v1/office/encryption/decrypt";
export const OFFICE_ENCRYPT_PATH = "/api/v1/office/encryption/encrypt";
/** The longest open password accepted (the server's own bound). */
export const MAX_OFFICE_PASSWORD_CHARS = 255;

const OLE2_SIGNATURE = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1] as const;
const ENCRYPTABLE_OFFICE_NAME = /\.(docx|docm|dotx|dotm|xlsx|xlsm|pptx|pptm)$/i;
const FRAME_MAGIC = [0x4d, 0x52, 0x50, 0x57] as const; // "MRPW"

/** The reasons the server names for an encryption it does not open. */
export type OfficeEncryptionUnsupportedReason =
  | "extensible_encryption"
  | "irm_rights_management"
  | "certificate_encryption"
  | "unsupported_cipher"
  | "unknown_encryption_version"
  | "doc_rc4"
  | "doc_rc4_cryptoapi"
  | "doc_xor_obfuscation"
  | "legacy_binary_format"
  | "odf_password";

export type OfficeEncryptionErrorCode =
  | "office_password_required"
  | "office_password_incorrect"
  | "office_password_invalid"
  | "office_integrity_failed"
  | "office_encryption_unsupported"
  | "office_not_encrypted"
  | "office_already_encrypted"
  | "office_file_unrecognized"
  | "office_file_damaged"
  | "office_encryption_failed"
  | "office_encryption_cancelled"
  | "office_encryption_required"
  | "office_encryption_unavailable";

export class OfficeEncryptionError extends Error {
  readonly code: OfficeEncryptionErrorCode | string;
  /** The HTTP status the refusal came with (422 for one made here), so the
   *  upload-failure lines (lib/errors/upload-failure) read it like any other. */
  readonly status: number;
  readonly reason?: string;
  readonly method?: string;
  readonly details: Record<string, unknown>;
  constructor(code: string, message: string, options: { reason?: string; method?: string; status?: number } = {}) {
    super(message);
    this.name = "OfficeEncryptionError";
    this.code = code;
    this.status = options.status ?? 422;
    this.details = {
      ...(options.reason ? { reason: options.reason } : {}),
      ...(options.method ? { method: options.method } : {}),
    };
    this.reason = options.reason;
    this.method = options.method;
  }
}

export function isOfficeEncryptionError(error: unknown): error is OfficeEncryptionError {
  return error instanceof OfficeEncryptionError;
}

/**
 * Whether `file` is an Office file Word encrypted: an Office name over an
 * OLE compound file (an unencrypted .docx/.xlsx/.pptx is a ZIP package).
 */
export async function isEncryptedOfficeFile(file: Blob & { readonly name?: string }): Promise<boolean> {
  if (!ENCRYPTABLE_OFFICE_NAME.test(file.name ?? "") || file.size < OLE2_SIGNATURE.length) return false;
  const head = new Uint8Array(await file.slice(0, OLE2_SIGNATURE.length).arrayBuffer());
  return OLE2_SIGNATURE.every((byte, index) => head[index] === byte);
}

/** The request body: the magic, the password's UTF-8 length, the password, the file. */
function frame(password: string, file: Blob): Blob {
  const encoded = new TextEncoder().encode(password);
  if (encoded.length > 0xffff) throw new OfficeEncryptionError("office_password_invalid", "password too long");
  const header = new Uint8Array(FRAME_MAGIC.length + 2 + encoded.length);
  header.set(FRAME_MAGIC, 0);
  header[4] = encoded.length >> 8;
  header[5] = encoded.length & 0xff;
  header.set(encoded, 6);
  return new Blob([header, file], { type: "application/octet-stream" });
}

function refusal(error: unknown): unknown {
  if (!isApiRequestError(error) || typeof error.code !== "string" || !error.code.startsWith("office_")) return error;
  const details = error.details ?? {};
  return new OfficeEncryptionError(error.code, error.message, {
    reason: typeof details.reason === "string" ? details.reason : undefined,
    method: typeof details.method === "string" ? details.method : undefined,
    status: error.status,
  });
}

async function send(path: string, password: string, file: Blob, signal?: AbortSignal): Promise<Response> {
  try {
    return await apiFetchRaw(
      path,
      { method: "POST", body: frame(password, file), ...(signal ? { signal } : {}) },
      { contentType: "application/octet-stream", suppressUnauthorizedHandling: true },
    );
  } catch (error) {
    throw refusal(error);
  }
}

/**
 * How `file` is encrypted, asked without a password: resolves with the
 * method ("agile" / "standard") when a password opens it, rejects with
 * `office_encryption_unsupported` (and its reason) when none can.
 */
export async function probeEncryptedOfficeFile(file: Blob, signal?: AbortSignal): Promise<{ method: string }> {
  try {
    await send(OFFICE_DECRYPT_PATH, "", file, signal);
  } catch (error) {
    if (isOfficeEncryptionError(error) && error.code === "office_password_required") {
      return { method: error.method ?? "agile" };
    }
    throw error;
  }
  // The server opened it without a password: not encrypted after all.
  throw new OfficeEncryptionError("office_not_encrypted", "The file is not encrypted");
}

/** The package inside an encrypted Office file, as a File of the same name. */
export async function decryptOfficeFile(file: File, password: string, signal?: AbortSignal): Promise<File> {
  const response = await send(OFFICE_DECRYPT_PATH, password, file, signal);
  const blob = await response.blob();
  const type = response.headers.get("content-type")?.split(";", 1)[0]?.trim() || file.type;
  return new File([blob], file.name, { type, lastModified: file.lastModified });
}

/** A Word package encrypted with `password` the way Word 2010+ encrypts. */
export async function encryptWordPackage(blob: Blob, password: string, signal?: AbortSignal): Promise<Blob> {
  const response = await send(OFFICE_ENCRYPT_PATH, password, blob, signal);
  return response.blob();
}

// ---------------------------------------------------------------------------
// The password dialog (the workspace's OfficePasswordHost answers it).

export type OfficePasswordRequest =
  /** Open an encrypted file: one field; `incorrect` after a wrong attempt. */
  | { readonly mode: "open"; readonly fileName: string; readonly incorrect?: boolean }
  /** Encrypt Document: the password and the same again (Word's two fields). */
  | { readonly mode: "set"; readonly fileName?: string; readonly change?: boolean };

/** The typed password, or null when the person cancelled. */
export type OfficePasswordAsker = (request: OfficePasswordRequest) => Promise<string | null>;

let asker: OfficePasswordAsker | null = null;

export function registerOfficePasswordAsker(next: OfficePasswordAsker): () => void {
  asker = next;
  return () => {
    if (asker === next) asker = null;
  };
}

export function askOfficePassword(request: OfficePasswordRequest): Promise<string | null> {
  if (!asker) {
    return Promise.reject(
      new OfficeEncryptionError("office_password_required", "No password dialog is available here"),
    );
  }
  return asker(request);
}

// ---------------------------------------------------------------------------
// The open documents' passwords: memory of this tab only.

const importedPasswords = new WeakMap<Blob, string>();

/** The password `pageId`'s exports are encrypted with this session, if any. */
export function documentOpenPassword(pageId: string): string | null {
  return heldPagePassword(pageId);
}

/** Set (or, with null, forget) the password of `pageId` for this session. */
export function setDocumentOpenPassword(pageId: string, password: string | null): void {
  holdPagePassword(pageId, password);
}

/** Whether a page was opened from an encrypted file this session (its
 *  password is held), so its editor marks it Encrypt with Password. */
export function pageOpenedFromEncryptedFile(pageId: string): boolean {
  return heldPagePassword(pageId) !== null;
}

/**
 * Each encrypted Office file of `files` opened with a password the person
 * types (asked again after a wrong one), the others as they are. Rejects
 * with `office_encryption_cancelled` when the person cancels, or with the
 * server's refusal for an encryption no password opens.
 */
export async function unlockEncryptedOfficeFiles(
  files: readonly File[],
  signal?: AbortSignal,
): Promise<File[]> {
  const out: File[] = [];
  for (const file of files) {
    if (!(await isEncryptedOfficeFile(file))) {
      out.push(file);
      continue;
    }
    try {
      await probeEncryptedOfficeFile(file, signal);
    } catch (error) {
      // An Office name over a compound file that is no encrypted package
      // (an unencrypted Word 97-2003 file renamed): the import reads it.
      if (isOfficeEncryptionError(error) && error.code === "office_not_encrypted") {
        out.push(file);
        continue;
      }
      throw error;
    }
    let incorrect = false;
    for (;;) {
      const password = await askOfficePassword({ mode: "open", fileName: file.name, incorrect });
      if (password === null) {
        throw new OfficeEncryptionError("office_encryption_cancelled", "Opening the encrypted file was cancelled");
      }
      try {
        const opened = await decryptOfficeFile(file, password, signal);
        importedPasswords.set(opened, password);
        out.push(opened);
        break;
      } catch (error) {
        if (isOfficeEncryptionError(error) && error.code === "office_password_incorrect") {
          incorrect = true;
          continue;
        }
        throw error;
      }
    }
  }
  return out;
}

/** The Word-family formats written encrypted with a page's password. */
const ENCRYPTED_EXPORT_FORMATS: ReadonlySet<string> = new Set(["docx", "xlsx", "pptx"]);
/** The files a page marked Encrypt with Password asks a password for. */
const MARKED_EXPORT_FORMATS: ReadonlySet<string> = new Set(["docx", "xlsx"]);

function pageEncryptionPath(pageId: string): string {
  return `/api/v1/pages/${encodeURIComponent(pageId)}/office-encryption`;
}

/** Whether the server has `pageId` marked Encrypt with Password. A page it
 *  will not say (a guest's, a passing failure) is unmarked here: the server's
 *  own refusal still stands in front of its file. */
export async function pageRequiresEncryption(pageId: string): Promise<boolean> {
  try {
    const answer = await apiGet<{ required?: boolean }>(pageEncryptionPath(pageId));
    return answer.required === true;
  } catch {
    return false;
  }
}

/**
 * The password `pageId`'s `format` file is encrypted with: the one this
 * session holds or, for a file of a page marked Encrypt with Password,
 * one asked for now. `null` when the file is not encrypted. Rejects, and no
 * file is written, when the person cancels or no dialog can ask.
 */
export async function resolvePageExportPassword(pageId: string, format: string): Promise<string | null> {
  if (!ENCRYPTED_EXPORT_FORMATS.has(format)) return null;
  const held = heldPagePassword(pageId);
  if (held) return held;
  if (!MARKED_EXPORT_FORMATS.has(format) || !(await pageRequiresEncryption(pageId))) return null;
  return askPageExportPassword(pageId);
}

/** Whether `pageId`'s files may only leave encrypted (its stored upload,
 *  which is plain, may not leave at all). */
export async function pageIsProtected(pageId: string): Promise<boolean> {
  return heldPagePassword(pageId) !== null || (await pageRequiresEncryption(pageId));
}

/** Ask for the password a marked page's file is encrypted with. */
export async function askPageExportPassword(pageId: string): Promise<string> {
  const held = heldPagePassword(pageId);
  if (held) return held;
  const asked = await askOfficePassword({ mode: "set" });
  if (!asked) throw new OfficeEncryptionError("office_encryption_cancelled", "The export was cancelled");
  holdPagePassword(pageId, asked);
  return asked;
}

/** After `file` (one {@link unlockEncryptedOfficeFiles} opened) became
 *  `pageId`: that page's exports are encrypted with its password again, in
 *  this session by the password held and in any other by the page's mark,
 *  written now: a page imported and closed unedited holds no editor to mark
 *  it. The mark failing leaves the password held, and the editor marks the
 *  page when it opens. */
export async function adoptImportedPassword(file: Blob, pageId: string): Promise<void> {
  const password = importedPasswords.get(file);
  if (!password) return;
  importedPasswords.delete(file);
  setDocumentOpenPassword(pageId, password);
  try {
    await apiPost(pageEncryptionPath(pageId), {});
  } catch (error) {
    console.warn("The encrypted file's page could not be marked yet.", error);
  }
}
