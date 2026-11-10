/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * PDF encryption policy (WPS PDF ▸ 文档加密).
 *
 * WPS asks for two different passwords and a set of permissions: an OPEN
 * password (nobody reads the file without it) and a PERMISSIONS password (the
 * file opens freely, but printing/copying/editing are restricted until it is
 * supplied). This module is the policy half — which combinations mean
 * something, which are a trap, and what the resulting permission bits are — so
 * the dialog and the writer agree and neither invents a guarantee.
 *
 * What encryption actually buys, stated plainly because users assume more:
 * an OPEN password genuinely keeps the content unreadable without it; a
 * PERMISSIONS password does not — every reader can decrypt the file, so the
 * restrictions are honoured by well-behaved viewers only. This module refuses
 * to pretend otherwise.
 */

export type PdfPermission =
  | "print"
  | "print-high-resolution"
  | "copy"
  | "modify"
  | "annotate"
  | "fill-forms"
  | "extract-for-accessibility"
  | "assemble";

export interface PdfEncryptionPolicy {
  /** Required to OPEN the document. Empty means the file opens freely. */
  openPassword: string;
  /** Required to LIFT the restrictions. Empty means they cannot be lifted. */
  permissionsPassword: string;
  /** What a reader may do without the permissions password. */
  allow: readonly PdfPermission[];
}

export type EncryptionPolicyProblem =
  | "no-password"
  | "same-password"
  | "permissions-without-password"
  | "password-too-long";

export interface EncryptionPolicyVerdict {
  ok: boolean;
  problems: EncryptionPolicyProblem[];
  /** True when the content itself is protected, not merely restricted. */
  contentProtected: boolean;
  /** Permissions a well-behaved viewer will enforce. */
  effectiveAllow: PdfPermission[];
  /**
   * An open password short enough to guess. The assurance below claims the
   * content cannot be read without it, and that claim is only as good as the
   * password — so the caller is told when it is not good.
   */
  weakOpenPassword: boolean;
}

/** PDF's own limit for a standard-security password. */
export const MAX_PDF_PASSWORD_LENGTH = 127;
/** Below this the assurance stops promising the content cannot be read. */
export const MIN_STRONG_PASSWORD_LENGTH = 8;

export const ALL_PDF_PERMISSIONS: readonly PdfPermission[] = [
  "print",
  "print-high-resolution",
  "copy",
  "modify",
  "annotate",
  "fill-forms",
  "extract-for-accessibility",
  "assemble",
];

/**
 * Judge a policy before anything is written. Never throws; a policy with
 * problems is reported rather than silently repaired, because "silently
 * repaired" is how a document ends up less protected than its author believes.
 */
export function reviewEncryptionPolicy(
  policy: PdfEncryptionPolicy,
): EncryptionPolicyVerdict {
  const problems: EncryptionPolicyProblem[] = [];
  const open = policy.openPassword.trim();
  const owner = policy.permissionsPassword.trim();

  if (!open && !owner) problems.push("no-password");
  if (open && owner && open === owner) problems.push("same-password");
  if (
    open.length > MAX_PDF_PASSWORD_LENGTH ||
    owner.length > MAX_PDF_PASSWORD_LENGTH
  ) {
    problems.push("password-too-long");
  }
  // Restricting permissions with no permissions password means nobody — not
  // even the author — can lift them later.
  const restricts = ALL_PDF_PERMISSIONS.some(
    (permission) => !policy.allow.includes(permission),
  );
  if (restricts && !owner) problems.push("permissions-without-password");

  const effectiveAllow = ALL_PDF_PERMISSIONS.filter((permission) =>
    policy.allow.includes(permission),
  );
  // High-resolution printing cannot outrank printing itself.
  const normalized = effectiveAllow.includes("print")
    ? effectiveAllow
    : effectiveAllow.filter(
        (permission) => permission !== "print-high-resolution",
      );

  return {
    ok: problems.length === 0,
    problems,
    contentProtected: open.length > 0,
    effectiveAllow: normalized,
    // Not a refusal: WPS accepts a short password too, and blocking here would
    // be a parity regression. It changes what we are entitled to CLAIM.
    weakOpenPassword: open.length > 0 && open.length < MIN_STRONG_PASSWORD_LENGTH,
  };
}

/** What the dialog should tell the user this policy actually guarantees. */
export function encryptionAssurance(
  verdict: EncryptionPolicyVerdict,
  t: (key: string) => string,
): string {
  // A policy with problems is not going to be written, so it guarantees
  // nothing. It used to answer for the password it CARRIES rather than for the
  // policy that will be applied: a password one character past the standard's
  // limit was refused by the review and still told the user the document could
  // not be read without it.
  if (!verdict.ok) {
    return t("pdf.encrypt.assuranceNotApplied");
  }
  if (verdict.contentProtected) {
    if (verdict.weakOpenPassword) {
      return t("pdf.encrypt.assuranceWeakPassword");
    }
    return t("pdf.encrypt.assuranceProtected");
  }
  return t("pdf.encrypt.assurancePermissionsOnly");
}
