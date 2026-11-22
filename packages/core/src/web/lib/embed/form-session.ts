/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A `fillForms` embed session's writes to a PDF form (apps/api
 * app/services/embed/forms.py): the session's account only reads the page,
 * so the PDF window sends each fill and each signature of a signature field
 * through these routes, which check the value against the file's field and
 * the person's role and write nothing else. A refusal carries the server's
 * own words (a read-only field, another role's field, a value the field does
 * not take).
 */

import { ApiRequestError, apiGet, apiPost } from "../api";

export type EmbedFormValue = string | boolean | string[];

export interface EmbedFormField {
  readonly name: string;
  readonly kind: string;
  readonly page?: number;
  /** The field's place on its page as shown: x, y, width, height in
   * fractions of the page from its top left. */
  readonly rect?: readonly number[];
  readonly required: boolean;
  readonly read_only: boolean;
  readonly role?: string;
  readonly fillable: boolean;
  readonly value?: EmbedFormValue;
  readonly signed?: boolean;
}

export interface EmbedFormView {
  readonly content_version: number;
  readonly role: string | null;
  readonly roles: readonly string[];
  readonly fields: readonly EmbedFormField[];
}

export type EmbedFormWrite =
  | { readonly ok: true; readonly contentVersion: number }
  | { readonly ok: false; readonly code: string; readonly message: string; readonly fields?: readonly string[] };

/** The fills and signatures on their way: Submit waits for them, so the
 * check and the file it sends hold the last field the person left. */
const inFlight = new Set<Promise<unknown>>();

function tracked<T>(write: Promise<T>): Promise<T> {
  inFlight.add(write);
  void write.finally(() => inFlight.delete(write)).catch(() => undefined);
  return write;
}

/** Resolves once no fill or signature is being written. */
export async function embedFormWritesSettled(): Promise<void> {
  while (inFlight.size) await Promise.allSettled([...inFlight]);
}

function refusal(error: unknown): EmbedFormWrite {
  if (error instanceof ApiRequestError) {
    const fields = error.details?.fields;
    return {
      ok: false,
      code: error.code ?? "embed_form_write_failed",
      message: error.message,
      ...(Array.isArray(fields) ? { fields: fields.filter((item): item is string => typeof item === "string") } : {}),
    };
  }
  return { ok: false, code: "network", message: error instanceof Error ? error.message : "" };
}

export async function readEmbedForm(): Promise<EmbedFormView> {
  return apiGet<EmbedFormView>("/api/v1/embed/session/form");
}

/** Fill in fields; a field the form computes or keeps read only is left
 * to the file (`skip_read_only`), as the editor's own save leaves it. */
export async function fillEmbedFormFields(
  entries: readonly { readonly name: string; readonly value: EmbedFormValue }[],
): Promise<EmbedFormWrite> {
  try {
    const result = await tracked(apiPost<{ content_version: number }>("/api/v1/embed/session/form/fields", {
      fields: entries,
      skip_read_only: true,
    }));
    return { ok: true, contentVersion: result.content_version };
  } catch (error) {
    return refusal(error);
  }
}

/** Place the person's signature picture on a signature field (null takes it off). */
export async function signEmbedFormField(
  field: string,
  image: string | null,
  kind: "drawn" | "typed" | "image",
): Promise<EmbedFormWrite> {
  try {
    const result = await tracked(
      apiPost<{ content_version: number }>("/api/v1/embed/session/form/signature", { field, image, kind }),
    );
    return { ok: true, contentVersion: result.content_version };
  } catch (error) {
    return refusal(error);
  }
}

/** Complete & Submit: refused with the missing required fields' names. */
export async function submitEmbedForm(): Promise<EmbedFormWrite> {
  await embedFormWritesSettled();
  try {
    const result = await apiPost<{ content_version: number }>("/api/v1/embed/session/form/submit", {});
    return { ok: true, contentVersion: result.content_version };
  } catch (error) {
    return refusal(error);
  }
}
