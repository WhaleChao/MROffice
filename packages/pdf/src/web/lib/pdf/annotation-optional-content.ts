/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import {
  PDFArray,
  PDFDict,
  PDFDocument,
  PDFName,
  PDFRef,
} from "pdf-lib";
import { pdfLibParseSpeed } from "@mrnote/core/web/lib/pdf/pdf-lib-pacing.ts";

export type OcVisibilityPolicy = "AnyOn" | "AllOn" | "AnyOff" | "AllOff";

export interface AnnotationOcMembership {
  /** pdf.js group ids ("12R") the annotation's /OC names. */
  readonly groups: readonly string[];
  readonly policy: OcVisibilityPolicy;
}

const MAX_ANNOTATIONS = 20_000;

function refId(ref: PDFRef): string {
  return ref.generationNumber === 0
    ? `${ref.objectNumber}R`
    : `${ref.objectNumber}R${ref.generationNumber}`;
}

export async function readAnnotationOptionalContent(
  bytes: Uint8Array,
): Promise<Map<string, AnnotationOcMembership>> {
  const memberships = new Map<string, AnnotationOcMembership>();
  try {
    const doc = await PDFDocument.load(bytes, {
      parseSpeed: pdfLibParseSpeed(),
      ignoreEncryption: true,
      updateMetadata: false,
      throwOnInvalidObject: false,
    });
    let seen = 0;
    for (const page of doc.getPages()) {
      const annots = page.node.lookupMaybe(PDFName.of("Annots"), PDFArray);
      if (!annots) continue;
      for (let index = 0; index < annots.size(); index += 1) {
        if (++seen > MAX_ANNOTATIONS) return memberships;
        const ref = annots.get(index);
        if (!(ref instanceof PDFRef)) continue;
        const annotation = doc.context.lookup(ref);
        if (!(annotation instanceof PDFDict)) continue;
        const oc = annotation.get(PDFName.of("OC"));
        if (!oc) continue;
        const target = oc instanceof PDFRef ? doc.context.lookup(oc) : oc;
        if (!(target instanceof PDFDict)) continue;
        const type = target.get(PDFName.of("Type"));
        if (type === PDFName.of("OCG") && oc instanceof PDFRef) {
          memberships.set(refId(ref), { groups: [refId(oc)], policy: "AnyOn" });
          continue;
        }
        if (type !== PDFName.of("OCMD")) continue;
        // A visibility expression (/VE) is not evaluated here; the group
        // list and policy cover the common membership dictionaries.
        if (target.get(PDFName.of("VE"))) continue;
        const raw = target.get(PDFName.of("OCGs"));
        const list =
          raw instanceof PDFArray
            ? raw.asArray()
            : raw instanceof PDFRef
              ? [raw]
              : [];
        const groups = list.filter((item): item is PDFRef => item instanceof PDFRef).map(refId);
        if (groups.length === 0) continue;
        const policyName = target.get(PDFName.of("P"));
        const policy =
          policyName instanceof PDFName &&
          ["AnyOn", "AllOn", "AnyOff", "AllOff"].includes(policyName.decodeText())
            ? (policyName.decodeText() as OcVisibilityPolicy)
            : "AnyOn";
        memberships.set(refId(ref), { groups, policy });
      }
    }
  } catch {
    return memberships;
  }
  return memberships;
}

/** Whether a membership is visible under the given group states; a group
 * the reader does not know counts as on. */
export function annotationOcVisible(
  membership: AnnotationOcMembership,
  groupVisible: (id: string) => boolean | undefined,
): boolean {
  const states = membership.groups.map((id) => groupVisible(id) ?? true);
  switch (membership.policy) {
    case "AllOn":
      return states.every(Boolean);
    case "AnyOff":
      return states.some((state) => !state);
    case "AllOff":
      return states.every((state) => !state);
    default:
      return states.some(Boolean);
  }
}
