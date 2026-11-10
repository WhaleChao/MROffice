/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { alertText } from "../../components/ui/modal-dialog";
import { isApiRequestError } from "../api";

/**
 * The server's refusals of a name that is already taken: a page's in its
 * folder (a code file's in its folder and directory), a notebook's among its
 * owner's, an AI window's in its notebook. A rename onto a taken name meets
 * them, so does filing a page into a folder that has its name, and so does a
 * new notebook whose name the person typed in the create dialog; anything
 * else made new is numbered " (2)" by the server instead.
 */
const NAME_CONFLICT_CODES = new Set([
  "page_title_conflict",
  "notebook_title_conflict",
  "window_title_conflict",
]);

export function isNameConflictError(error: unknown): boolean {
  return (
    isApiRequestError(error) &&
    error.status === 409 &&
    NAME_CONFLICT_CODES.has(error.code ?? "")
  );
}

/**
 * Says in a dialog that names cannot repeat, for a taken-name refusal, and
 * returns true; returns false for any other error so the caller reports it
 * its own way. The message is the localized `code.<code>` line the API
 * client already put on the error.
 */
export function showNameConflictDialog(error: unknown): boolean {
  if (!isNameConflictError(error)) return false;
  void alertText((error as Error).message);
  return true;
}
