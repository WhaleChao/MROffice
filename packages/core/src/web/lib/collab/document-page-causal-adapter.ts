/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { omitUndefinedCanonicalXmlAttributes } from "./canonical-xml-attributes";
import { omitUndefinedMapValues } from "./y-json-values";
import { COLLABORATIVE_DOCUMENT_ATTRS_MAP } from "../../components/console/editor/native-document/collaborationDocumentAttrs";
import {
  documentCausalContentMirrored,
  documentCausalContentMirrors,
  documentContentSaysProjection,
  documentSaveBodyFor,
  documentSaveCopy,
  lastDocumentMirrorMiss,
  readDocumentCausalProjection,
  readSharedDocumentCausalProjection,
} from "./document-page-causal-model";
import { createNativePageCausalAdapter, type NativePageCausalHost } from "./native-page-causal-adapter";

/** The paged bridge owns authoring XML and document attrs synchronously.
 * Storage JSON is a verification input, never a second author model. */
export function createDocumentPageCausalAdapter(host: NativePageCausalHost) {
  return createNativePageCausalAdapter(host, {
    model: "document-canonical-v1",
    projection: readDocumentCausalProjection,
    // The mirror's own baseline: read once per document revision and kept
    // inside the adapter (consumers still receive a fresh projection).
    baselineProjection: readSharedDocumentCausalProjection,
    // A non-null contentForSave is a copy of the requested body proven equal
    // to the baseline, so a capture made at the same revision needs no
    // second projection check (see capture in native-page-causal-adapter).
    saveProvesProjection: true,
    canMirror: (content, _baseline, doc) => documentCausalContentMirrors(content, doc),
    // Neither canMirror nor apply reads the baseline: the body is proven
    // against the live room itself.
    mirrorsAgainstRoom: true,
    mirrorMiss: lastDocumentMirrorMiss,
    roomBody: doc => structuredClone(readDocumentCausalProjection(doc)),
    // mirror() has already checked exact body equivalence and the live writer.
    // Repair legacy pending XML at that existing non-history save boundary;
    // never rebuild the body from JSON or alter already-sent update identities.
    apply: doc => {
      const xml = omitUndefinedCanonicalXmlAttributes(doc.getXmlFragment("default"));
      // The document's own attributes are JSON values in a map: the same rule.
      return omitUndefinedMapValues(doc.getMap(COLLABORATIVE_DOCUMENT_ATTRS_MAP)) || xml;
    },
    contentForSave(projection, requested) {
      try {
        // The mirror check has just compared exactly this body with exactly
        // this projection object; comparing them again gives the same answer.
        // The body sent carries what the room passes through for a newer
        // client (documentSaveBodyFor): the save is then the room's own.
        const body = documentSaveBodyFor(requested, projection);
        if (documentCausalContentMirrored(requested, projection)) return documentSaveCopy(body, projection);
        return documentContentSaysProjection(body, projection) ? documentSaveCopy(body, projection) : null;
      } catch { return null; }
    },
    shouldDeliver: (session, transaction) => !transaction.local ||
      [...transaction.changed.keys()].some(type => Object.is(type, session.ydoc.getMap(COLLABORATIVE_DOCUMENT_ATTRS_MAP))),
    // No seedLegacy: a paginated old room is not a flat canonical original.
    // The shared loader must obtain the server's fenced inactive epoch first.
  });
}
