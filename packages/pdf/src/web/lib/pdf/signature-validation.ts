/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Certificate signature status (WPS PDF ▸ 数字签名).
 *
 * A signed PDF is judged on two independent questions, and conflating them is
 * how a viewer misleads people: is the SIGNED BYTES' digest still intact, and
 * is the SIGNER someone you have reason to trust? A document can be
 * cryptographically intact and signed by a stranger, or signed by a trusted
 * certificate and modified afterwards. Acrobat and WPS report those as
 * different states, and so does this.
 *
 * This module is the judgement, not the cryptography: the host supplies what a
 * verifier found, and this turns it into the status a user is shown. It never
 * upgrades a status on a missing input — an unknown is reported as unknown,
 * because "probably fine" is the one answer a signature panel must not give.
 */

export type SignatureStatus =
  | "valid"
  | "valid-with-later-changes"
  | "digest-mismatch"
  | "untrusted-signer"
  | "expired-certificate"
  | "revoked-certificate"
  | "unknown";

export interface SignatureEvidence {
  /** Whether the digest over the signed byte range still matches. */
  digestIntact?: boolean;
  /** Whether the whole file is covered, or only a revision of it. */
  coversWholeDocument?: boolean;
  /** Whether the signer's chain reached a trust anchor. */
  chainTrusted?: boolean;
  /** Whether the certificate was valid AT SIGNING TIME. */
  certificateValidAtSigning?: boolean;
  /**
   * Whether revocation was checked and the certificate had NOT been revoked.
   * Revocation is how a compromised key is disowned, so a chain that still
   * reaches a trust anchor says nothing on its own — and an unchecked
   * revocation status is an unknown, not a pass.
   */
  notRevoked?: boolean;
  /** A trusted timestamp, when one was embedded. */
  timestamped?: boolean;
}

export interface SignatureVerdict {
  status: SignatureStatus;
  /** True only for a signature that is intact AND trusted AND unmodified. */
  trustworthy: boolean;
  /** Why, in the reader's language — always populated. */
  reason: string;
}

const REASONS: Record<SignatureStatus, [zh: string, en: string]> = {
  valid: [
    "签名有效,文档自签名后未被修改。",
    "The signature is valid and the document has not changed since signing.",
  ],
  "valid-with-later-changes": [
    "签名本身有效,但文档在签名后被修改过。",
    "The signature is valid, but the document was changed after it was signed.",
  ],
  "digest-mismatch": [
    "文档内容与签名不符,签名无效。",
    "The document does not match its signature.",
  ],
  "untrusted-signer": [
    "签名完整,但签名者不在受信任的证书链中。",
    "The signature is intact, but the signer is not trusted.",
  ],
  "expired-certificate": [
    "签名时证书已失效。",
    "The certificate was not valid when the document was signed.",
  ],
  "revoked-certificate": [
    "签署证书已被吊销,签名不可信。",
    "The signing certificate was revoked, so the signature cannot be trusted.",
  ],
  unknown: ["无法判定签名状态。", "This signature could not be judged."],
};

/**
 * Judge one signature. Order matters: a broken digest outranks everything (the
 * bytes changed, so nothing else is worth saying), and trust is only meaningful
 * once integrity holds.
 */
export function judgeSignature(
  evidence: SignatureEvidence,
  locale: "en" | "zh" = "zh",
): SignatureVerdict {
  const say = (status: SignatureStatus): SignatureVerdict => ({
    status,
    trustworthy: status === "valid",
    reason: REASONS[status][locale === "zh" ? 0 : 1],
  });

  if (evidence.digestIntact === false) return say("digest-mismatch");
  // A revoked certificate is disowned by the authority that issued it, which
  // outranks every question about the chain: the chain may still reach an
  // anchor, and that is exactly the case revocation exists to answer. Said
  // before the digest is known because the wording claims nothing about it.
  if (evidence.notRevoked === false) return say("revoked-certificate");
  // Every status below this line asserts the signed bytes are intact — the
  // untrusted-signer wording says so in as many words — so none of them can be
  // reached without knowing that.
  if (evidence.digestIntact === undefined) return say("unknown");
  // A definite negative is reported from the evidence that establishes it. What
  // must never happen is an upgrade on a missing input, not a refusal to state
  // a failure the verifier did find.
  if (evidence.certificateValidAtSigning === false) {
    return say("expired-certificate");
  }
  if (evidence.chainTrusted === false) return say("untrusted-signer");
  // Everything past here is a PASS, so EVERY input it rests on has to be
  // present: a missing verifier result is not a good result.
  if (
    evidence.chainTrusted === undefined ||
    evidence.certificateValidAtSigning === undefined ||
    evidence.notRevoked === undefined
  ) {
    return say("unknown");
  }
  if (evidence.coversWholeDocument === false) {
    return say("valid-with-later-changes");
  }
  if (evidence.coversWholeDocument === undefined) return say("unknown");
  return say("valid");
}

/**
 * What a VISUAL signature is, stated so a panel cannot imply more. The visual
 * signature feature draws an image; it makes no cryptographic claim at all.
 */
export function visualSignatureDisclaimer(locale: "en" | "zh"): string {
  return locale === "zh"
    ? "这是手写外观的图形签名,不是数字证书签名,无法验证签署人或文档是否被改动。"
    : "This is a drawn signature image, not a certificate signature: it proves neither who signed nor whether the document changed.";
}

/** The worst status across a document's signatures — what the badge shows. */
export function documentSignatureStatus(
  verdicts: readonly SignatureVerdict[],
): SignatureStatus {
  if (verdicts.length === 0) return "unknown";
  const order: SignatureStatus[] = [
    "digest-mismatch",
    "revoked-certificate",
    "expired-certificate",
    "untrusted-signer",
    "unknown",
    "valid-with-later-changes",
    "valid",
  ];
  for (const status of order) {
    if (verdicts.some((verdict) => verdict.status === status)) return status;
  }
  return "unknown";
}
