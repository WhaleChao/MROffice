/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Where a face added through the FontFace API came from.
 *
 * A face declared in a stylesheet keeps its `src` in the rule; a face built
 * with `new FontFace(family, source)` does not expose its source again. A
 * picture of a page (an exported slide, a presenter preview, a standalone
 * HTML file) has to carry the faces its text is set in, so whoever builds a
 * face from a URL notes the URL here, keyed by the face itself.
 */

// One registry per page, shared by every bundle that reads it: a probe
// injected into the page (the text engine's audits) sees the faces the app
// noted, not an empty copy of its own.
interface FontFaceRegistry {
  readonly sources: WeakMap<FontFace, string>;
  /** The descriptors a face was built with (size-adjust, ascent-override...):
   * WebKit does not expose every one on the FontFace object. */
  readonly descriptors: Map<string, Readonly<Record<string, string>>>;
  /** The same, by the face's descriptors: WebKit hands out a new wrapper for
   * a face when document.fonts is iterated, so the object itself is not a
   * key there. */
  readonly sourcesByDescriptors: Map<string, string>;
  /** The noted faces of each document, in the order noted, with what their
   * builder declared; faces noted for no document in particular, and how
   * many faces were noted in all (absent on a registry made by an older
   * bundle). A document's list goes with the document. */
  notedByDoc?: WeakMap<Document, NotedFontFace[]>;
  notedAnywhere?: NotedFontFace[];
  notedCount?: number;
  /** Noted faces whose descriptor keys are not written yet. */
  unkeyed?: NotedFontFace[];
}

/** What a face's builder declared: the descriptors as given to the FontFace
 * constructor (CSS syntax, not the engine's serialization of them). */
export interface DeclaredFontFace {
  readonly family: string;
  readonly weight?: string;
  readonly style?: string;
  readonly unicodeRange?: string;
}

export interface NotedFontFace {
  readonly face: FontFace;
  readonly src: string;
  readonly built?: Readonly<Record<string, string>>;
  /** The document the face is added to, when the builder said. */
  readonly doc?: Document;
  readonly declared?: DeclaredFontFace;
}

const registry: FontFaceRegistry = (() => {
  const holder = globalThis as typeof globalThis & { __mraiFontFaceRegistry?: FontFaceRegistry };
  return (holder.__mraiFontFaceRegistry ??= { sources: new WeakMap(), descriptors: new Map(), sourcesByDescriptors: new Map() });
})();
const { sources, descriptors, sourcesByDescriptors } = registry;
const notedByDoc = (registry.notedByDoc ??= new WeakMap());
const notedAnywhere = (registry.notedAnywhere ??= []);
const unkeyed = (registry.unkeyed ??= []);
const builtByFace = new WeakMap<FontFace, Readonly<Record<string, string>>>();
const descriptorKey = (face: FontFace) =>
  [face.family, face.weight, face.style, face.stretch, face.unicodeRange].join("\u0000");











function keyNotedFaces(): void {
  if (unkeyed.length === 0) return;
  for (const entry of unkeyed.splice(0)) {
    try {
      const key = descriptorKey(entry.face);
      sourcesByDescriptors.set(key, entry.src);
      if (entry.built) descriptors.set(key, entry.built);
    } catch { /* a face without descriptors */ }
  }
}

/** Note the CSS `src` list (`url(...)`, `local(...)`) a face was built from;
 * `doc` and `declared` (the constructor's descriptors) let a reader of the
 * noted faces skip the engine's FontFace getters ({@link notedFontFaces}). */
export function noteFontFaceSource(
  face: FontFace,
  src: string,
  built?: Readonly<Record<string, string>>,
  where?: { readonly doc?: Document; readonly declared?: DeclaredFontFace },
): void {
  sources.set(face, src);
  if (built) builtByFace.set(face, built);
  const entry: NotedFontFace = { face, src, ...(built ? { built } : {}), ...(where?.doc ? { doc: where.doc } : {}), ...(where?.declared ? { declared: where.declared } : {}) };
  if (where?.doc) {
    const list = notedByDoc.get(where.doc);
    if (list) list.push(entry);
    else notedByDoc.set(where.doc, [entry]);
  } else {
    notedAnywhere.push(entry);
  }
  registry.notedCount = (registry.notedCount ?? 0) + 1;
  unkeyed.push(entry);
  queueFacesNotice();
}














const facesListeners = new Set<() => void>();
let facesNoticeQueued = false;

function queueFacesNotice(): void {
  if (facesNoticeQueued || facesListeners.size === 0) return;
  facesNoticeQueued = true;
  setTimeout(() => {
    facesNoticeQueued = false;
    for (const listener of [...facesListeners]) {
      try {
        listener();
      } catch {
        /* one listener's failure leaves the others told */
      }
    }
  }, 0);
}

/** Hear when faces are declared or a family's registration settles. */
export function onFontFacesNoted(listener: () => void): () => void {
  facesListeners.add(listener);
  return () => {
    facesListeners.delete(listener);
  };
}

const familyKey = (family: string): string => family.trim().replace(/^["']|["']$/g, "").trim().toLowerCase();
/** Families (normalized) whose faces are being registered, and how many registrations each. */
const familiesRegistering = new Map<string, number>();
let familyRegistrationsSettled = 0;

/**
 * A family's faces are being registered (the font registry asked, its faces
 * not all declared yet): until the returned call, a reader that finds no face
 * declared for it waits rather than passing on to the stack's next family.
 * The call is idempotent.
 */
export function noteFamilyRegistration(family: string): () => void {
  const key = familyKey(family);
  if (!key) return () => {};
  familiesRegistering.set(key, (familiesRegistering.get(key) ?? 0) + 1);
  let settled = false;
  return () => {
    if (settled) return;
    settled = true;
    const left = (familiesRegistering.get(key) ?? 1) - 1;
    if (left > 0) familiesRegistering.set(key, left);
    else familiesRegistering.delete(key);
    familyRegistrationsSettled += 1;
    queueFacesNotice();
  };
}

/** Whether `family`'s faces are still being registered (`noteFamilyRegistration`). */
export function familyRegistrationPending(family: string): boolean {
  return familiesRegistering.size > 0 && familiesRegistering.has(familyKey(family));
}

/** Moves each time a family's registration settles: a cache of answers that
 * waited on one keys on it. */
export function familyRegistrationGeneration(): number {
  return familyRegistrationsSettled;
}

/** The faces noted as added to `doc`, in the order noted (each builder
 * notes a face right after adding it, so this is the order they were
 * added in); null when some face was noted without its document, as only
 * the document's own set can then say which faces are its. */
export function notedFontFaces(doc: Document): readonly NotedFontFace[] | null {
  if (notedAnywhere.length > 0) return null;
  return notedByDoc.get(doc) ?? [];
}

/** How many faces have been noted (any document): changes exactly when
 * {@link notedFontFaces} can. */
export function notedFontFaceCount(): number {
  return registry.notedCount ?? 0;
}

/** The descriptors noted with a face's source, when its builder gave them. */
export function fontFaceDescriptors(face: FontFace): Readonly<Record<string, string>> | undefined {
  const own = builtByFace.get(face);
  if (own !== undefined) return own;
  keyNotedFaces();
  try {
    return descriptors.get(descriptorKey(face));
  } catch {
    return undefined;
  }
}

/** The CSS `src` list a face was built from, when its builder noted one. */
export function fontFaceSource(face: FontFace): string | undefined {
  const own = sources.get(face);
  if (own !== undefined) return own;
  keyNotedFaces();
  try {
    return sourcesByDescriptors.get(descriptorKey(face));
  } catch {
    return undefined;
  }
}

const BUILT_DESCRIPTORS: ReadonlyArray<readonly [string, string]> = [
  ["sizeAdjust", "size-adjust"],
  ["ascentOverride", "ascent-override"],
  ["descentOverride", "descent-override"],
  ["lineGapOverride", "line-gap-override"],
  ["variationSettings", "font-variation-settings"],
];

/**
 * The faces noted as added to `doc`, as stylesheet rules: what a window of its
 * own (a print copy) needs for its text to be set in them. Copying stylesheets
 * brings no face built with `new FontFace`: a family the page drew from one fell
 * to whatever the machine had (an Amharic report in macOS's Kefa where the page
 * measured it in a metric-matched face: lines longer than the page's, sheets cut
 * at the paper's edge). Faces whose builder did not declare them are left out.
 */
export function notedFontFaceRules(doc: Document): string[] {
  const rules = new Set<string>();
  for (const { src, built, declared } of notedFontFaces(doc) ?? []) {
    if (!declared || !src) continue;
    const family = declared.family.replace(/^["']|["']$/g, "").replace(/["\\]/g, "\\$&");
    const parts = [`font-family:"${family}"`, `src:${src}`, "font-display:swap"];
    if (declared.weight) parts.push(`font-weight:${declared.weight}`);
    if (declared.style) parts.push(`font-style:${declared.style}`);
    if (declared.unicodeRange) parts.push(`unicode-range:${declared.unicodeRange}`);
    for (const [key, property] of BUILT_DESCRIPTORS) {
      const value = built?.[key];
      if (value) parts.push(`${property}:${value}`);
    }
    rules.add(`@font-face{${parts.join(";")}}`);
  }
  return [...rules];
}
