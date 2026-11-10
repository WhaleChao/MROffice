/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * XFA form handling (WPS PDF ▸ XFA 表单).
 *
 * XFA is Adobe's XML form format, and the dangerous part is that an XFA
 * document usually ALSO carries an AcroForm. For a STATIC XFA the two agree,
 * so filling the AcroForm is safe. For a DYNAMIC one the AcroForm is a stale
 * shell — often empty, often missing pages the form generates at runtime — and
 * writing values into it produces a file that looks filled in every viewer
 * except the one the form was built for.
 *
 * So the decision this module makes is not "can we render it" but "what may we
 * let the user do without quietly ruining their form": fill, read only, or
 * nothing at all. It refuses rather than guesses, because a corrupted form is
 * discovered by whoever receives it.
 */

export type XfaKind = "none" | "static" | "dynamic" | "unknown";

export interface XfaDescriptor {
  /** Whether the document carries an /XFA entry at all. */
  present: boolean;
  /**
   * The XFA packet's declared kind. Static forms have a fixed layout;
   * dynamic ones generate layout at runtime.
   */
  kind?: XfaKind;
  /** Whether a usable AcroForm field set exists alongside the XFA packet. */
  acroFormFieldCount?: number;
  /** Whether the document was flagged as having been filled in a viewer. */
  hasFieldValues?: boolean;
}

export type XfaCapability =
  /** No XFA: the normal AcroForm path applies. */
  | "acroform"
  /** Static XFA with a matching AcroForm: filling is safe. */
  | "fill-via-acroform"
  /** The form can be shown but must not be edited here. */
  | "read-only"
  /** Nothing can be shown faithfully; say so instead of rendering a shell. */
  | "unsupported";

export interface XfaVerdict {
  capability: XfaCapability;
  /** True when saving would risk a form that only some viewers read correctly. */
  saveWouldRisk: boolean;
  /** Catalogue key (console-notebooks) of the sentence that explains it. */
  reasonKey: string;
}

const REASON_KEYS: Record<XfaCapability, string> = {
  acroform: "pdf.xfa.reasonAcroform",
  "fill-via-acroform": "pdf.xfa.reasonFillViaAcroform",
  "read-only": "pdf.xfa.reasonReadOnly",
  unsupported: "pdf.xfa.reasonUnsupported",
};

/** Decide what may be done with a document's forms. */
export function judgeXfa(descriptor: XfaDescriptor): XfaVerdict {
  const say = (
    capability: XfaCapability,
    saveWouldRisk: boolean,
  ): XfaVerdict => ({
    capability,
    saveWouldRisk,
    reasonKey: REASON_KEYS[capability],
  });

  if (!descriptor.present) return say("acroform", false);

  const fields = descriptor.acroFormFieldCount ?? 0;

  // Dynamic XFA generates its layout at runtime; its AcroForm is a shell.
  // Once that form has been FILLED, the shell's values are not authoritative
  // either, so showing it would display numbers that may not be the form's —
  // worse than showing nothing, and the reason this module refuses rather than
  // guesses.
  if (descriptor.kind === "dynamic") {
    if (descriptor.hasFieldValues) return say("unsupported", true);
    return say(fields > 0 ? "read-only" : "unsupported", true);
  }

  if (descriptor.kind === "static") {
    // A static form with no AcroForm fields has nothing to fill through.
    return fields > 0
      ? say("fill-via-acroform", false)
      : say("read-only", true);
  }

  // An XFA packet we could not classify is treated as the worse case: showing
  // it is fine, writing to it is not.
  return say(fields > 0 ? "read-only" : "unsupported", true);
}

/**
 * The verdict for a file whose source profile names its XFA kind (``xfaDynamic``:
 * "none" static, "present" dynamic, anything else not known).
 *
 * The widgets are counted from the pages on screen, and a save's new revision
 * clears them until its pages paint. A static form's kind is its file's, so a
 * count of none there does not make it read-only: that took the Fill tool away
 * after every save of the IRS's Form 8822. With truly no fields there is nothing
 * to offer anyway. A dynamic or unclassified form still counts what it shows.
 */
export function judgeXfaOfProfile(
  xfaDynamic: "none" | "present" | "unknown" | undefined,
  acroFormFieldCount: number,
): XfaVerdict {
  if (xfaDynamic === "none") {
    return judgeXfa({
      present: true,
      kind: "static",
      acroFormFieldCount: Math.max(acroFormFieldCount, 1),
    });
  }
  return judgeXfa({
    present: true,
    kind: xfaDynamic === "present" ? "dynamic" : "unknown",
    acroFormFieldCount,
  });
}

/** Whether the editor should offer form-filling controls at all. */
export function allowsFormFilling(verdict: XfaVerdict): boolean {
  return (
    verdict.capability === "acroform" ||
    verdict.capability === "fill-via-acroform"
  );
}

/** Whether the document can be displayed, even if it cannot be edited. */
export function allowsDisplay(verdict: XfaVerdict): boolean {
  return verdict.capability !== "unsupported";
}
