/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A form's filling roles (ONLYOFFICE's Manage Roles / View form as a role):
 * who fills each field. A field names its role itself, in the PDF (the
 * editor's /MRNoteRole and /MRNoteRoleColor on the field dictionary, which
 * the native field writer sets and clears); one naming none is the default
 * role, anyone. The form's roles are the ones its fields name.
 *
 * Filling as a role, only that role's fields (and anyone's) take input; the
 * others stay shown, marked in their role's colour, and locked.
 */
import { PDFDict, PDFDocument, PDFHexString, PDFName, PDFRef, PDFString } from "pdf-lib";

export interface PdfFormRole {
  readonly name: string;
  /** `#RRGGBB`, upper case: the colour the editor marks the role's fields in. */
  readonly color: string;
}

/** A role's name as the field writer takes it: trimmed, 1-64 characters, no
 * control characters. */
export const PDF_FORM_ROLE_MAX_NAME = 64;

/** The colours offered to a new role, in order (ONLYOFFICE's and Office's
 * accent order: blue, orange, green, red, purple, brown, pink, teal). */
export const PDF_FORM_ROLE_PALETTE: readonly string[] = Object.freeze([
  "#4472C4", "#ED7D31", "#70AD47", "#C00000", "#7030A0", "#8B5A2B", "#E15A97", "#00A3A3",
]);

export function isPdfFormRoleName(value: unknown): value is string {
  return typeof value === "string" && value.length > 0 && value === value.trim() &&
    value.length <= PDF_FORM_ROLE_MAX_NAME && ![...value].some((character) => {
      const code = character.codePointAt(0) ?? 0;
      return code < 0x20 || code === 0x7f;
    });
}

export function readPdfFormRole(value: unknown): PdfFormRole | null {
  if (!value || typeof value !== "object" || Array.isArray(value)) return null;
  const candidate = value as Record<string, unknown>;
  if (Object.keys(candidate).sort().join(",") !== "color,name") return null;
  if (!isPdfFormRoleName(candidate.name) || typeof candidate.color !== "string" ||
    !/^#[0-9A-Fa-f]{6}$/u.test(candidate.color)) return null;
  return { name: candidate.name, color: candidate.color.toUpperCase() };
}

/** The colour a new role takes: the first of the palette no role has. */
export function nextPdfFormRoleColor(roles: readonly PdfFormRole[]): string {
  const taken = new Set(roles.map((role) => role.color.toUpperCase()));
  return PDF_FORM_ROLE_PALETTE.find((color) => !taken.has(color)) ??
    PDF_FORM_ROLE_PALETTE[roles.length % PDF_FORM_ROLE_PALETTE.length]!;
}

/** The form's roles, in the order its fields first name them (one name, one
 * role: a later colour for the same name is the first's). */
export function pdfFormRolesOf(fieldRoles: Iterable<PdfFormRole | null | undefined>): PdfFormRole[] {
  const roles = new Map<string, PdfFormRole>();
  for (const role of fieldRoles) {
    if (role && !roles.has(role.name)) roles.set(role.name, role);
  }
  return [...roles.values()];
}

/** Whether a field takes input when the form is filled as `fillingAs` (null:
 * as anyone -- every field, as without roles). */
export function pdfFormFieldFillableAs(fieldRole: PdfFormRole | null | undefined,
  fillingAs: string | null): boolean {
  return fillingAs === null || !fieldRole || fieldRole.name === fillingAs;
}

const OBJECT_REF_RE = /^([1-9][0-9]{0,9})R([0-9]{1,5})?$/u;
const documents = new WeakMap<object, Promise<PDFDocument | null>>();

function textOf(value: unknown): string | null {
  return value instanceof PDFString || value instanceof PDFHexString ? value.decodeText() : null;
}

/**
 * The role each widget's field names, read from the bytes (pdf.js reports no
 * key of the editor's): keyed by the widget's object reference (pdf.js's
 * annotation id, `12R`), inherited from a parent field as a field's other
 * entries are. A widget whose field names none is left out.
 */
export async function readPdfFormFieldRoles(
  owner: object,
  readBytes: () => Promise<Uint8Array>,
  widgetRefs: readonly string[],
): Promise<Map<string, PdfFormRole>> {
  const roles = new Map<string, PdfFormRole>();
  if (widgetRefs.length === 0) return roles;
  let pending = documents.get(owner);
  if (!pending) {
    pending = readBytes()
      .then((bytes) => PDFDocument.load(bytes, { ignoreEncryption: true, updateMetadata: false, throwOnInvalidObject: false }))
      .catch(() => null);
    documents.set(owner, pending);
  }
  const document = await pending;
  if (!document) return roles;
  for (const ref of widgetRefs) {
    const match = OBJECT_REF_RE.exec(ref);
    if (!match) continue;
    try {
      let node: unknown = document.context.lookup(PDFRef.of(Number(match[1]), Number(match[2] ?? 0)));
      for (let depth = 0; node instanceof PDFDict && depth < 32; depth += 1) {
        const name = textOf(node.lookup(PDFName.of("MRNoteRole")));
        if (name !== null) {
          const role = readPdfFormRole({ name, color: textOf(node.lookup(PDFName.of("MRNoteRoleColor"))) ?? "" });
          if (role) roles.set(ref, role);
          break;
        }
        node = node.lookup(PDFName.of("Parent"));
      }
    } catch {
      // An unreadable widget names no role: anyone fills it.
    }
  }
  return roles;
}
