/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Detecting certificate signatures before a mutation destroys them.
 *
 * Every export this editor produces is a full rewrite: pdf-lib re-serialises
 * the document, so the byte ranges every /Sig digest was computed over no
 * longer exist, and every certificate signature in the file arrives at the
 * recipient INVALID. That is unavoidable given the pipeline — what is not
 * acceptable is doing it silently. Acrobat and WPS both name the signatures
 * and ask; so must we, and asking starts with detecting.
 *
 * Visual signatures (our own stamp-like ones) are not certificate signatures
 * and are untouched by this: they carry no digest to break.
 */

import {
  PDFArray,
  PDFDict,
  PDFHexString,
  PDFInvalidObject,
  PDFName,
  PDFNull,
  PDFNumber,
  PDFObjectParser,
  PDFRawStream,
  PDFRef,
  PDFStream,
  PDFString,
  type PDFDocument,
} from "pdf-lib";

const MAX_FIELD_NODES = 10_000;
const MAX_FIELD_DEPTH = 64;
const MAX_INDIRECT_OBJECTS = 100_000;
const MAX_DIRECT_CONTAINERS = 200_000;
const MAX_SIGNATURES = 1_000;
const MAX_SIGNATURE_REFERENCES = 64;
const MAX_SIGNATURE_CONTENT_BYTES = 16 * 1024 * 1024;
const MAX_TEXT_BYTES = 4 * 1024;

const NAME_ACRO_FORM = PDFName.of("AcroForm");
const NAME_BYTE_RANGE = PDFName.of("ByteRange");
const NAME_CONTENTS = PDFName.of("Contents");
const NAME_DOC_MDP = PDFName.of("DocMDP");
const NAME_ENCRYPT = PDFName.of("Encrypt");
const NAME_FIELDS = PDFName.of("Fields");
const NAME_FILTER = PDFName.of("Filter");
const NAME_FIRST = PDFName.of("First");
const NAME_FT = PDFName.of("FT");
const NAME_KIDS = PDFName.of("Kids");
const NAME_M = PDFName.of("M");
const NAME_N = PDFName.of("N");
const NAME_NAME = PDFName.of("Name");
const NAME_PARENT = PDFName.of("Parent");
const NAME_PERMS = PDFName.of("Perms");
const NAME_REFERENCE = PDFName.of("Reference");
const NAME_SUB_FILTER = PDFName.of("SubFilter");
const NAME_T = PDFName.of("T");
const NAME_TRANSFORM_METHOD = PDFName.of("TransformMethod");
const NAME_TRANSFORM_PARAMS = PDFName.of("TransformParams");
const NAME_TYPE = PDFName.of("Type");
const NAME_V = PDFName.of("V");
const NAME_XREF = PDFName.of("XRef");

export interface DetectedCertificateSignature {
  /** /T of the signature field, or a stable fallback. */
  fieldName: string;
  /** /Name in the signature dictionary — the signer as recorded. */
  signerName: string;
  /** /M — the recorded signing time, as the PDF date string. */
  signedAt: string;
  /** /SubFilter, e.g. adbe.pkcs7.detached or ETSI.CAdES.detached. */
  subFilter: string;
  /**
   * True when a /DocMDP reference marks this as a CERTIFICATION signature:
   * the author's declaration of what may change. Breaking one of these is
   * more than invalidating an approval — it voids the document's own rules.
   */
  certification: boolean;
}

interface ParsedSignature {
  signerName: string;
  signedAt: string;
  subFilter: string;
  certification: boolean;
}

interface SignatureRecord extends ParsedSignature {
  ref: PDFRef;
  fieldName: string | null;
}

function malformed(detail: string): never {
  throw new Error(`Ambiguous PDF certificate signature structure: ${detail}`);
}

function name(value: unknown, label: string): string {
  if (!(value instanceof PDFName)) malformed(`${label} must be a name`);
  const decoded = value.decodeText();
  if (!decoded || decoded.length > 256) malformed(`${label} is invalid`);
  return decoded;
}

function optionalText(value: unknown, label: string): string {
  if (value === undefined) return "";
  if (!(value instanceof PDFHexString) && !(value instanceof PDFString)) {
    malformed(`${label} must be a text string`);
  }
  if (value.asBytes().byteLength > MAX_TEXT_BYTES) {
    malformed(`${label} is too large`);
  }
  return value.decodeText();
}

function resolveDict(doc: PDFDocument, value: unknown, label: string): PDFDict {
  if (value instanceof PDFDict) return value;
  if (!(value instanceof PDFRef)) malformed(`${label} must be a dictionary`);
  const resolved = doc.context.lookupMaybe(value, PDFDict);
  if (!resolved) malformed(`${label} reference is missing`);
  return resolved;
}

function resolveArray(
  doc: PDFDocument,
  value: unknown,
  label: string,
): PDFArray {
  if (value instanceof PDFArray) return value;
  if (!(value instanceof PDFRef)) malformed(`${label} must be an array`);
  const resolved = doc.context.lookupMaybe(value, PDFArray);
  if (!resolved) malformed(`${label} reference is missing`);
  return resolved;
}

function parseByteRange(value: unknown): void {
  if (!(value instanceof PDFArray) || value.size() !== 4) {
    malformed("/ByteRange must be an exact four-number array");
  }
  const parts: number[] = [];
  for (let index = 0; index < value.size(); index += 1) {
    const part = value.get(index);
    if (!(part instanceof PDFNumber)) {
      malformed("/ByteRange entries must be direct numbers");
    }
    const numeric = part.asNumber();
    if (!Number.isSafeInteger(numeric) || numeric < 0) {
      malformed("/ByteRange entries must be non-negative safe integers");
    }
    parts.push(numeric);
  }
  const [firstOffset, firstLength, secondOffset, secondLength] = parts;
  if (
    firstOffset !== 0 ||
    firstLength < 1 ||
    secondOffset <= firstLength ||
    !Number.isSafeInteger(secondOffset + secondLength)
  ) {
    malformed("/ByteRange is not ordered around a signature gap");
  }
}

function parseSignatureReferences(doc: PDFDocument, value: unknown): boolean {
  if (value === undefined) return false;
  if (!(value instanceof PDFArray)) {
    malformed("/Reference must be a direct array");
  }
  if (value.size() > MAX_SIGNATURE_REFERENCES) {
    malformed("too many signature transform references");
  }
  let certification = false;
  const seenRefs = new Set<string>();
  const seenDirect = new Set<PDFDict>();
  const seenMethods = new Set<string>();
  for (let index = 0; index < value.size(); index += 1) {
    const rawReference = value.get(index);
    if (rawReference instanceof PDFRef) {
      if (seenRefs.has(rawReference.tag)) {
        malformed("duplicate signature transform reference");
      }
      seenRefs.add(rawReference.tag);
    }
    const reference = resolveDict(
      doc,
      rawReference,
      "signature transform reference",
    );
    if (!(rawReference instanceof PDFRef)) {
      if (seenDirect.has(reference)) {
        malformed("duplicate signature transform dictionary");
      }
      seenDirect.add(reference);
    }
    const referenceType = reference.get(NAME_TYPE);
    if (
      referenceType !== undefined &&
      name(referenceType, "/Reference /Type") !== "SigRef"
    ) {
      malformed("unsupported signature transform reference type");
    }
    const method = name(
      reference.get(NAME_TRANSFORM_METHOD),
      "/TransformMethod",
    );
    if (method !== "DocMDP" && method !== "FieldMDP") {
      malformed("unsupported signature transform method");
    }
    if (seenMethods.has(method)) {
      malformed("duplicate signature transform method");
    }
    seenMethods.add(method);
    certification ||= method === "DocMDP";

    const params = reference.get(NAME_TRANSFORM_PARAMS);
    if (params !== undefined) {
      const paramsDict = resolveDict(doc, params, "/TransformParams");
      const paramsType = paramsDict.get(NAME_TYPE);
      if (
        paramsType !== undefined &&
        name(paramsType, "/TransformParams /Type") !== "TransformParams"
      ) {
        malformed("unsupported transform-parameter type");
      }
      const permission = paramsDict.get(PDFName.of("P"));
      if (permission !== undefined) {
        if (!(permission instanceof PDFNumber)) {
          malformed("transform permission must be a direct number");
        }
        const permissionValue = permission.asNumber();
        if (
          !Number.isSafeInteger(permissionValue) ||
          permissionValue < 1 ||
          permissionValue > 3
        ) {
          malformed("transform permission is outside the supported domain");
        }
      }
    }
  }
  return certification;
}

function parseSignatureDictionary(
  doc: PDFDocument,
  value: PDFDict,
): ParsedSignature {
  const signatureType = name(value.get(NAME_TYPE), "signature /Type");
  if (signatureType !== "Sig" && signatureType !== "DocTimeStamp") {
    malformed("unsupported signature dictionary type");
  }
  name(value.get(NAME_FILTER), "signature /Filter");
  const subFilter = name(value.get(NAME_SUB_FILTER), "signature /SubFilter");
  parseByteRange(value.get(NAME_BYTE_RANGE));

  const contents = value.get(NAME_CONTENTS);
  if (!(contents instanceof PDFHexString) && !(contents instanceof PDFString)) {
    malformed("signature /Contents must be a direct string");
  }
  const contentsLength = contents.asBytes().byteLength;
  if (contentsLength < 1 || contentsLength > MAX_SIGNATURE_CONTENT_BYTES) {
    malformed("signature /Contents size is unsupported");
  }

  return {
    signerName: optionalText(value.get(NAME_NAME), "signature /Name"),
    signedAt: optionalText(value.get(NAME_M), "signature /M"),
    subFilter,
    certification: parseSignatureReferences(doc, value.get(NAME_REFERENCE)),
  };
}

function isSignatureLike(value: PDFDict): boolean {
  const type = value.get(NAME_TYPE);
  if (type instanceof PDFName) {
    const decoded = type.decodeText();
    if (decoded === "Sig" || decoded === "DocTimeStamp") return true;
  }
  return (
    value.get(NAME_BYTE_RANGE) !== undefined ||
    value.get(NAME_SUB_FILTER) !== undefined
  );
}

/**
 * Walk every supported signature anchor and authenticate the exact structure.
 *
 * Presence is intentionally narrower than validity: this proves only that a
 * byte-range certificate-signature container exists. Trust, CMS digest,
 * revocation, timestamps, and DocMDP permissions need a separate validator.
 * Any structure that could conceal a signature but cannot be authenticated is
 * an exception so source-risk inspection records UNKNOWN instead of "none".
 */
/**
 * A usage-rights grant no /Perms names any more: a form filled here gives up the
 * grant its values invalidate (Form 8822's /UR3), and the earlier revision keeps
 * the dictionary. It signs nothing a person signed, and read as an unknown
 * signature it made every filled form read-only with "signature unverified".
 */
function grantsUsageRights(doc: PDFDocument, signature: PDFDict): boolean {
  const references = doc.context.lookup(signature.get(NAME_REFERENCE));
  if (!(references instanceof PDFArray) || references.size() === 0) return false;
  for (let index = 0; index < references.size(); index += 1) {
    const reference = doc.context.lookup(references.get(index));
    if (!(reference instanceof PDFDict)) return false;
    const method = doc.context.lookup(reference.get(NAME_TRANSFORM_METHOD));
    if (!(method instanceof PDFName)) return false;
    if (method.decodeText() !== "UR3" && method.decodeText() !== "UR") return false;
  }
  return true;
}

/**
 * What an object pdf-lib kept opaque provably is: nothing (null), or the
 * dictionary of a cross-reference stream. Anything else refuses.
 *
 * Opacity is not proof that an object is harmless or unsigned, and its bytes
 * are never searched for one. Two shapes are read for what they are:
 *
 * - No bytes between `obj` and `endobj` (ImageMagick writes `13 0 obj endobj`
 *   for a slot it leaves unused). Nothing is there to be a dictionary, and a
 *   signature field or /Perms entry that names it still refuses below, since
 *   it does not resolve to one.
 * - A cross-reference stream whose table pdf-lib could not decode (the one a
 *   decrypting tool leaves behind, its rows garbage, under the new one).
 *   pdf-lib's own object parser reads the same bytes as a stream: what failed
 *   was reading its payload as a table. The dictionary is then judged as any
 *   stream's is, and the payload is cross-reference rows, never objects.
 *
 * Such a stream is also a trailer pdf-lib did not read. If the /Encrypt it
 * names exists, the file may be ciphertext that `isEncrypted` missed and no
 * walk of it proves anything; only a reference to an object the file does
 * not hold is known to encrypt nothing. An object stream that did not decode
 * hides members never seen, whatever /Type it states, so /N or /First refuse.
 *
 * The attachment inspection reads an opaque object the same way
 * (save-policy's hasEmbeddedFileRisk): one rule for what such an object is.
 */
export function opaqueObjectDictionary(
  doc: PDFDocument,
  object: PDFInvalidObject,
): PDFDict | null {
  const size = object.sizeInBytes();
  if (size === 0) return null;
  const bytes = new Uint8Array(size);
  object.copyBytesInto(bytes, 0);
  let parsed: unknown;
  try {
    parsed = PDFObjectParser.forBytes(bytes, doc.context).parseObject();
  } catch {
    parsed = undefined;
  }
  if (
    !(parsed instanceof PDFRawStream) ||
    parsed.dict.get(NAME_TYPE) !== NAME_XREF ||
    parsed.dict.get(NAME_N) !== undefined ||
    parsed.dict.get(NAME_FIRST) !== undefined
  ) {
    malformed("opaque indirect object prevents signature presence proof");
  }
  const encrypt = parsed.dict.get(NAME_ENCRYPT);
  if (
    encrypt !== undefined &&
    !(encrypt instanceof PDFRef && doc.context.lookup(encrypt) === undefined)
  ) {
    malformed("unread cross-reference stream declares encryption");
  }
  return parsed.dict;
}

export function detectCertificateSignatures(
  doc: PDFDocument,
): DetectedCertificateSignature[] {
  const records = new Map<string, SignatureRecord>();
  const signedFieldNames = new Map<string, string>();

  const recordSignature = (
    ref: unknown,
    fieldName: string | null,
    certification: boolean,
  ) => {
    if (!(ref instanceof PDFRef)) {
      malformed("signature dictionary must be an indirect reference");
    }
    const signature = doc.context.lookupMaybe(ref, PDFDict);
    if (!signature) malformed("signature dictionary reference is missing");
    const parsed = parseSignatureDictionary(doc, signature);
    const current = records.get(ref.tag);
    if (current) {
      if (fieldName && current.fieldName && current.fieldName !== fieldName) {
        malformed("one signature is attached to multiple field names");
      }
      current.fieldName ||= fieldName;
      current.certification ||= certification || parsed.certification;
      return;
    }
    if (records.size >= MAX_SIGNATURES) malformed("too many signatures");
    if (fieldName) {
      const existingRef = signedFieldNames.get(fieldName);
      if (existingRef && existingRef !== ref.tag) {
        malformed("duplicate signed field name");
      }
      signedFieldNames.set(fieldName, ref.tag);
    }
    records.set(ref.tag, {
      ref,
      fieldName,
      ...parsed,
      certification: certification || parsed.certification,
    });
  };

  const acroFormObject = doc.catalog.get(NAME_ACRO_FORM);
  if (acroFormObject !== undefined) {
    const acroForm = resolveDict(doc, acroFormObject, "catalog /AcroForm");
    const fieldsObject = acroForm.get(NAME_FIELDS);
    if (fieldsObject !== undefined) {
      const fields = resolveArray(doc, fieldsObject, "/AcroForm /Fields");
      const seenFields = new Set<string>();
      let fieldCount = 0;
      const visit = (
        fieldRef: unknown,
        inheritedName: string,
        inheritedType: string | null,
        parentRef: PDFRef | null,
        depth: number,
      ) => {
        if (depth > MAX_FIELD_DEPTH) malformed("field tree is too deep");
        if (!(fieldRef instanceof PDFRef)) {
          malformed("field-tree entries must be indirect references");
        }
        if (seenFields.has(fieldRef.tag)) {
          malformed("duplicate or cyclic field-tree reference");
        }
        seenFields.add(fieldRef.tag);
        fieldCount += 1;
        if (fieldCount > MAX_FIELD_NODES) malformed("too many form fields");
        const node = doc.context.lookupMaybe(fieldRef, PDFDict);
        if (!node) malformed("field reference is missing");

        const declaredParent = node.get(NAME_PARENT);
        if (declaredParent !== undefined) {
          if (!(declaredParent instanceof PDFRef)) {
            malformed("field /Parent must be an indirect reference");
          }
          if (!parentRef || declaredParent.tag !== parentRef.tag) {
            malformed("field /Parent does not match its field-tree owner");
          }
        }

        const ownNameObject = node.get(NAME_T);
        const ownName =
          ownNameObject === undefined
            ? ""
            : optionalText(ownNameObject, "field /T");
        const fullName = inheritedName
          ? ownName
            ? `${inheritedName}.${ownName}`
            : inheritedName
          : ownName;
        const ownType = node.get(NAME_FT);
        const effectiveType =
          ownType === undefined ? inheritedType : name(ownType, "field /FT");

        const value = node.get(NAME_V);
        if (effectiveType === "Sig") {
          // A signature field a form maker left with an empty string (or
          // any string) for /V holds no signature: only a signature
          // dictionary signs. Read as malformed, the whole form was shown
          // as 「包含证书签名」 and its fields refused (user audit PDF 0137).
          if (
            value !== undefined &&
            value !== PDFNull &&
            !(value instanceof PDFString) &&
            !(value instanceof PDFHexString)
          ) {
            recordSignature(value, fullName || null, false);
          }
        } else if (value instanceof PDFRef) {
          const resolvedValue = doc.context.lookupMaybe(value, PDFDict);
          if (resolvedValue && isSignatureLike(resolvedValue)) {
            malformed("signature dictionary belongs to a non-signature field");
          }
        }

        const kidsObject = node.get(NAME_KIDS);
        if (kidsObject !== undefined) {
          const kids = resolveArray(doc, kidsObject, "field /Kids");
          for (let index = 0; index < kids.size(); index += 1) {
            visit(
              kids.get(index),
              fullName,
              effectiveType,
              fieldRef,
              depth + 1,
            );
          }
        }
      };

      for (let index = 0; index < fields.size(); index += 1) {
        visit(fields.get(index), "", null, null, 0);
      }
    }
  }

  // Usage rights (/UR, /UR3: Adobe Reader extensions) sit under /Perms
  // too. They are a grant a reader honours, not a signature over the
  // document's content: an edit only ends the grant, as WPS edits such a
  
  // its text stayed read-only). Only /DocMDP certifies.
  // Held as the reference or, as Acrobat writes it, the direct dictionary.
  const usageRights = new Set<unknown>();
  const permissionsObject = doc.catalog.get(NAME_PERMS);
  if (permissionsObject !== undefined) {
    const permissions = resolveDict(doc, permissionsObject, "catalog /Perms");
    for (const [permissionName, value] of permissions.entries()) {
      const name = permissionName.decodeText();
      if (name === "UR" || name === "UR3") {
        usageRights.add(value);
        continue;
      }
      if (name !== "DocMDP") {
        malformed("unsupported catalog permission signature");
      }
    }
    const docMdp = permissions.get(NAME_DOC_MDP);
    if (docMdp !== undefined) recordSignature(docMdp, null, true);
    else if (usageRights.size === 0 && permissions.keys().length === 0) {
      malformed("empty catalog /Perms dictionary");
    }
  }

  const indirectObjects = doc.context.enumerateIndirectObjects();
  if (indirectObjects.length > MAX_INDIRECT_OBJECTS) {
    malformed("too many indirect objects to authenticate signatures");
  }
  const seenDirectContainers = new Set<PDFArray | PDFDict | PDFStream>();
  const activeDirectContainers = new Set<PDFArray | PDFDict | PDFStream>();
  let directContainerCount = 0;
  const inspectDirectValue = (value: unknown, depth: number): void => {
    if (value instanceof PDFRef || value === undefined || usageRights.has(value)) return;
    if (
      !(value instanceof PDFArray) &&
      !(value instanceof PDFDict) &&
      !(value instanceof PDFStream)
    ) {
      return;
    }
    if (depth > MAX_FIELD_DEPTH) {
      malformed("direct signature-like structure is too deep");
    }
    if (activeDirectContainers.has(value)) {
      malformed("cyclic direct signature-like structure");
    }
    if (seenDirectContainers.has(value)) {
      malformed("duplicate direct signature-like structure");
    }
    seenDirectContainers.add(value);
    activeDirectContainers.add(value);
    directContainerCount += 1;
    if (directContainerCount > MAX_DIRECT_CONTAINERS) {
      malformed("too many direct objects to authenticate signatures");
    }
    if (value instanceof PDFArray) {
      for (let index = 0; index < value.size(); index += 1) {
        inspectDirectValue(value.get(index), depth + 1);
      }
    } else {
      const dictionary = value instanceof PDFStream ? value.dict : value;
      if (isSignatureLike(dictionary)) {
        malformed("signature-like dictionaries must be indirect non-streams");
      }
      for (const [, child] of dictionary.entries()) {
        inspectDirectValue(child, depth + 1);
      }
    }
    activeDirectContainers.delete(value);
  };
  for (const [ref, object] of indirectObjects) {
    if (object instanceof PDFInvalidObject) {
      // pdf-lib deliberately preserves an indirect object it cannot parse as
      // an opaque PDFInvalidObject. Only an empty one and an undecoded
      // cross-reference stream are read (opaqueObjectDictionary); any other
      // refuses there.
      const unread = opaqueObjectDictionary(doc, object);
      if (unread) {
        if (isSignatureLike(unread)) {
          malformed("signature-like streams are unsupported");
        }
        for (const [, child] of unread.entries()) {
          inspectDirectValue(child, 1);
        }
      }
      continue;
    }
    if (object instanceof PDFArray) {
      for (let index = 0; index < object.size(); index += 1) {
        inspectDirectValue(object.get(index), 1);
      }
      continue;
    }
    if (
      object instanceof PDFDict &&
      isSignatureLike(object) &&
      !usageRights.has(ref) &&
      !grantsUsageRights(doc, object)
    ) {
      recordSignature(ref, null, false);
    }
    const dictionary =
      object instanceof PDFDict
        ? object
        : object instanceof PDFStream
          ? object.dict
          : null;
    if (object instanceof PDFStream && isSignatureLike(object.dict)) {
      malformed("signature-like streams are unsupported");
    }
    if (dictionary) {
      for (const [, child] of dictionary.entries()) {
        inspectDirectValue(child, 1);
      }
    }
  }

  return Array.from(records.values(), (record, index) => ({
    fieldName: record.fieldName || `signature-${index + 1}`,
    signerName: record.signerName,
    signedAt: record.signedAt,
    subFilter: record.subFilter,
    certification: record.certification,
  }));
}

export interface SignatureMutationImpact {
  signatures: DetectedCertificateSignature[];
  /** Any signature at all — the export will invalidate every one of them. */
  willInvalidate: boolean;
  /** A certification signature is among them — the stronger warning. */
  breaksCertification: boolean;
}

/**
 * What a full-rewrite export does to this document's signatures. There is no
 * partial answer: the pipeline re-serialises everything, so ANY signature
 * present means every signature invalid on arrival.
 */
export function signatureMutationImpact(
  doc: PDFDocument,
): SignatureMutationImpact {
  const signatures = detectCertificateSignatures(doc);
  return {
    signatures,
    willInvalidate: signatures.length > 0,
    breaksCertification: signatures.some(
      (signature) => signature.certification,
    ),
  };
}
