/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import {
  PDFArray,
  PDFDict,
  PDFDocument,
  PDFHexString,
  PDFName,
  PDFNumber,
  PDFRef,
  PDFStream,
  PDFString,
} from "pdf-lib";

import type { NativePdfFreeTextSourceFacts } from "./native-free-text-annotations";

const OBJECT_REF_RE = /^([1-9][0-9]{0,9})R([0-9]{1,5})?$/u;
const MAX_RD = 14_400;

const documents = new WeakMap<object, Promise<PDFDocument | null>>();

/** One pdf-lib document per shown pdf.js document, loaded on first need. */
function documentFor(
  owner: object,
  readBytes: () => Promise<Uint8Array>,
): Promise<PDFDocument | null> {
  let pending = documents.get(owner);
  if (!pending) {
    pending = readBytes()
      .then((bytes) =>
        PDFDocument.load(bytes, {
          ignoreEncryption: true,
          updateMetadata: false,
          throwOnInvalidObject: false,
        }),
      )
      .catch(() => null);
    documents.set(owner, pending);
  }
  return pending;
}

function textOf(value: unknown): string | null {
  return value instanceof PDFString || value instanceof PDFHexString
    ? value.decodeText()
    : null;
}

function fontResourceName(defaultAppearance: string | null): string | null {
  if (!defaultAppearance) return null;
  const match = /\/([^\s/[\]()<>{}%]+)\s+[-+]?[\d.]+\s+Tf/u.exec(defaultAppearance);
  return match ? match[1] : null;
}

function baseFontIn(resources: unknown, name: string): string | null {
  if (!(resources instanceof PDFDict)) return null;
  const fonts = resources.lookupMaybe(PDFName.of("Font"), PDFDict);
  const font = fonts?.lookup(PDFName.of(name));
  if (!(font instanceof PDFDict)) return null;
  const base = font.lookup(PDFName.of("BaseFont"));
  return base instanceof PDFName ? base.decodeText() : null;
}

function alphaOf(value: unknown): number | undefined {
  if (!(value instanceof PDFNumber)) return undefined;
  const number = value.asNumber();
  return Number.isFinite(number) && number >= 0 && number <= 1 ? number : undefined;
}







function sourceOpacity(dict: PDFDict): number | undefined {
  const appearance = dict.lookupMaybe(PDFName.of("AP"), PDFDict)?.lookup(PDFName.of("N"));
  if (!(appearance instanceof PDFStream)) {
    return appearance === undefined ? alphaOf(dict.lookup(PDFName.of("CA"))) ?? 1 : undefined;
  }
  const resources = appearance.dict.lookup(PDFName.of("Resources"));
  const states =
    resources instanceof PDFDict
      ? resources.lookupMaybe(PDFName.of("ExtGState"), PDFDict)
      : undefined;
  if (!states) return 1;
  const alphas = new Set<number>();
  for (const [, value] of states.entries()) {
    const state = dict.context.lookup(value);
    if (!(state instanceof PDFDict)) continue;
    for (const key of ["ca", "CA"]) {
      const alpha = alphaOf(state.lookup(PDFName.of(key)));
      if (alpha !== undefined) alphas.add(Math.round(alpha * 10_000) / 10_000);
    }
  }
  if (alphas.size === 0) return 1;
  return alphas.size === 1 ? [...alphas][0] : undefined;
}

export async function readNativeFreeTextSourceFacts(
  owner: object,
  readBytes: () => Promise<Uint8Array>,
  objectRefs: readonly string[],
): Promise<Map<string, NativePdfFreeTextSourceFacts>> {
  const facts = new Map<string, NativePdfFreeTextSourceFacts>();
  if (objectRefs.length === 0) return facts;
  const document = await documentFor(owner, readBytes);
  if (!document) return facts;
  const formResources = document.catalog
    .lookupMaybe(PDFName.of("AcroForm"), PDFDict)
    ?.lookup(PDFName.of("DR"));
  for (const objectRef of objectRefs) {
    const match = OBJECT_REF_RE.exec(objectRef);
    if (!match) continue;
    try {
      const dict = document.context.lookup(
        PDFRef.of(Number(match[1]), Number(match[2] ?? 0)),
      );
      if (!(dict instanceof PDFDict)) continue;
      const entry: {
        rectangleDifferences?: readonly [number, number, number, number];
        baseFont?: string;
        appearanceOpacity?: number;
        sourceTextOwner?: boolean;
        managedObject?: boolean;
        managedTextBox?: boolean;
      } = {};
      if (dict.has(PDFName.of("MRNoteSourceTextIdentity"))) entry.sourceTextOwner = true;
      // A table, chart or SmartArt made in MRAI (managed-table.ts,
      
      if (dict.has(PDFName.of("MRNoteTable")) || dict.has(PDFName.of("MRNoteChart")) ||
        dict.has(PDFName.of("MRNoteSmartArt"))) entry.managedObject = true;
      
      else if (dict.has(PDFName.of("MRNoteTextBox"))) entry.managedTextBox = true;
      const opacity = sourceOpacity(dict);
      if (opacity !== undefined && opacity < 1) entry.appearanceOpacity = opacity;
      const rd = dict.lookup(PDFName.of("RD"));
      if (rd instanceof PDFArray && rd.size() === 4) {
        const values = rd
          .asArray()
          .map((value) => (value instanceof PDFNumber ? value.asNumber() : Number.NaN));
        if (values.every((value) => Number.isFinite(value) && value >= 0 && value <= MAX_RD)) {
          entry.rectangleDifferences = values as unknown as readonly [number, number, number, number];
        }
      }
      const name = fontResourceName(textOf(dict.lookup(PDFName.of("DA"))));
      if (name) {
        const appearance = dict
          .lookupMaybe(PDFName.of("AP"), PDFDict)
          ?.lookup(PDFName.of("N"));
        const appearanceResources =
          appearance instanceof PDFStream
            ? appearance.dict.lookup(PDFName.of("Resources"))
            : undefined;
        const baseFont =
          baseFontIn(appearanceResources, name) ?? baseFontIn(formResources, name);
        if (baseFont) entry.baseFont = baseFont.slice(0, 200);
      }
      // Every readable object has facts, an empty entry included: it says
      
      facts.set(objectRef, entry);
    } catch {
      /* one unreadable object gives no facts */
    }
  }
  return facts;
}
