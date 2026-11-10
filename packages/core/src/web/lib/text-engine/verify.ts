/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Whether a face declared with `local()` sources draws in the served file.
 *
 * `src: local("Calibri"), url(carlito.woff2)` draws in the visitor's own
 * Calibri when one is installed. The served file's widths are only the drawn
 * widths when that font is absent, or is metric-compatible with the file
 * (Calibri and Carlito are). Checked once per face: a handful of the face's
 * own characters measured by the browser at 1000px against the file. A face
 * that differs is reported as such and the engine leaves its text to the
 * browser's own measurement, so an installed font is never measured with
 * another font's widths.
 */
import { faceCovers, selectComposite, declaredFaces, normalizeFamily, parseWeight, type FaceDeclaration } from "./face-index";
import { faceMetrics, type FaceMetrics } from "./face-metrics";
import { apiFaceOrder } from "./face-order";
import { glyphAdvanceUnits } from "./variation";

export type FaceVerification = "same" | "differs" | "pending";

const results = new Map<string, FaceVerification>();
/** The last check of each face: sample, browser and file advances (em). */
const details = new Map<string, unknown>();
export function faceVerificationDetail(face: FaceDeclaration): unknown {
  return details.get(face.url + "\u0000" + face.family);
}
let generation = 0;
const listeners = new Set<() => void>();
/** Answers that land together are told once: each told listener (a
 * spreadsheet window) draws its whole grid again, and a desk of sheets
 * opening asked dozens of faces whose answers came one by one. */
const ANNOUNCE_AFTER_MS = 150;
let announcing: ReturnType<typeof setTimeout> | undefined;
function announce(): void {
  if (announcing !== undefined) return;
  announcing = setTimeout(() => {
    announcing = undefined;
    for (const listener of [...listeners]) {
      try { listener(); } catch { /* its own */ }
    }
  }, ANNOUNCE_AFTER_MS);
}
let checker: ((face: FaceDeclaration, metrics: FaceMetrics, samples: readonly number[]) => Promise<boolean>) | null = null;

export function faceVerificationGeneration(): number {
  return generation;
}
export function onFaceVerified(listener: () => void): () => void {
  listeners.add(listener);
  return () => listeners.delete(listener);
}

/** Tests: decide verification without a browser. */
export function setFaceChecker(next: typeof checker): void {
  checker = next;
  results.clear();
  localOnly.clear();
  generation += 1;
}

const PREFERRED_SAMPLES = [..."HAMOWamnoeil0158"].map((char) => char.codePointAt(0)!);

function samplesOf(face: FaceDeclaration, metrics: FaceMetrics): number[] {
  const order = apiFaceOrder();
  const composite = selectComposite(declaredFaces(face.family), face.weightMin, face.style !== "normal", order === "pending" ? "last-wins" : order);
  const winner = (codePoint: number) => {
    for (let index = composite.length - 1; index >= 0; index -= 1) {
      if (faceCovers(composite[index], codePoint)) return composite[index];
    }
    return undefined;
  };
  const out: number[] = [];
  const consider = (codePoint: number) => {
    if (out.length >= 8 || out.includes(codePoint)) return;
    if (!metrics.cmap.has(codePoint) || !faceCovers(face, codePoint) || winner(codePoint) !== face) return;
    out.push(codePoint);
  };
  for (const codePoint of PREFERRED_SAMPLES) consider(codePoint);
  if (out.length < 4) {
    for (const codePoint of metrics.cmap.keys()) {
      if (codePoint > 0x20 && codePoint !== 0xa0 && codePoint !== 0xad) consider(codePoint);
      if (out.length >= 8) break;
    }
  }
  return out;
}

async function browserCheck(face: FaceDeclaration, metrics: FaceMetrics, samples: readonly number[]): Promise<boolean> {
  if (typeof document === "undefined" || !document.fonts) return false;
  const style = face.style === "normal" ? "normal" : "italic";
  const family = `"${face.family.replace(/["\\]/g, "\\$&")}"`;
  const font = `${style} ${face.weightMin} 1000px ${family}`;
  const text = String.fromCodePoint(...samples);
  await document.fonts.load(font, text);
  const context = document.createElement("canvas").getContext("2d");
  if (!context) return false;
  context.font = font;
  (context as CanvasRenderingContext2D & { fontKerning?: string }).fontKerning = "none";
  const rows: [string, number, number][] = [];
  let same = true;
  for (const codePoint of samples) {
    const browser = context.measureText(String.fromCodePoint(codePoint)).width / 1000;
    const served = glyphAdvanceUnits(face, metrics, metrics.cmap.get(codePoint)!, face.weightMin) / metrics.unitsPerEm * face.sizeAdjust;
    rows.push([String.fromCodePoint(codePoint), browser, served]);
    if (!Number.isFinite(browser) || Math.abs(browser - served) > 0.0008) same = false;
  }
  details.set(face.url + "\u0000" + face.family, { font, rows });
  return same;
}

/**
 * What decides whether the browser draws a face in an installed font: its
 * family, style and local() names. One check answers for every slice that
 * shares them (a CJK stand-in is a hundred slices with the same local
 * names), so a page asks the browser a few times, not once per file.
 */
function shadowKey(face: FaceDeclaration): string {
  return [face.family, face.weightMin, face.weightMax, face.style, face.localNames.join("|")].join("\u0000");
}

/** A check that has not answered in this long is asked again later. */
const CHECK_TIMEOUT_MS = 15_000;

/** "same" for a face with no local() source, or one checked equal. */
export function faceVerification(face: FaceDeclaration): FaceVerification {
  if (!face.localNames.length || !face.url) return "same";
  if (apiFaceOrder() === "pending") return "pending";
  const key = shadowKey(face);
  const known = results.get(key);
  if (known) return known;
  const metrics = faceMetrics(face.url);
  if (metrics === "pending") return "pending";
  if (metrics === "failed") return "differs";
  const samples = samplesOf(face, metrics);
  // A slice that wins none of its characters is never drawn: nothing to ask.
  if (!samples.length) return "same";
  return startCheck(key, face, metrics, samples);
}

/**
 * Whether an installed font a local()-only face draws (`localOnlyFaceState`
 * "installed") is the served file's metric twin: the family's characters the
 * file has, measured by the browser at 1000px (the installed font draws them)
 * against the file. "same": the file's widths are the drawn widths, so the
 * engine lays the text out with them (an installed Calibri or Arial, whose
 * advances Carlito and Arimo reproduce). The served face itself is one the
 * installed font shadows, so `faceVerification` has no samples of it.
 */
export function installedTwinVerification(served: FaceDeclaration, metrics: FaceMetrics): FaceVerification {
  if (!served.url) return "differs";
  const key = "twin\u0000" + served.url + "\u0000" + shadowKey(served);
  const known = results.get(key);
  if (known) return known;
  const samples: number[] = [];
  const consider = (codePoint: number) => {
    if (samples.length < 8 && !samples.includes(codePoint) && metrics.cmap.has(codePoint) && faceCovers(served, codePoint)) samples.push(codePoint);
  };
  for (const codePoint of PREFERRED_SAMPLES) consider(codePoint);
  if (samples.length < 4) {
    for (const codePoint of metrics.cmap.keys()) {
      if (codePoint > 0x20 && codePoint !== 0xa0 && codePoint !== 0xad) consider(codePoint);
      if (samples.length >= 8) break;
    }
  }
  if (!samples.length) return "differs";
  return startCheck(key, served, metrics, samples);
}

function startCheck(key: string, face: FaceDeclaration, metrics: FaceMetrics, samples: readonly number[]): FaceVerification {
  results.set(key, "pending");
  let timer: ReturnType<typeof setTimeout> | undefined;
  const timeout = new Promise<"timeout">((resolve) => { timer = setTimeout(() => resolve("timeout"), CHECK_TIMEOUT_MS); });
  void Promise.race([(checker ?? browserCheck)(face, metrics, samples), timeout])
    .then(
      (answer) => {
        if (answer === "timeout") results.delete(key);
        else results.set(key, answer ? "same" : "differs");
      },
      () => results.set(key, "differs"),
    )
    .finally(() => {
      if (timer) clearTimeout(timer);
      generation += 1;
      announce();
    });
  return "pending";
}

const localOnly = new Map<string, "missing" | "installed" | "pending">();
const FULL_RANGE = /^\s*u\+0+-10ffff\s*$/i;

/**
 * A face with only local() sources (`src: local("Arial")`, declared after the
 * served slices of the same family so an installed Office font wins): the
 * browser loads it when a character needs it, and when the font is not
 * installed the load fails and the character falls through to the family's
 * earlier faces -- the served slices. The CSS-connected FontFace says which:
 * "missing" (failed: fall through), "installed" (the engine cannot measure
 * it), "pending" (asked; the answer bumps the verification generation).
 *
 * Asked with a FontFace of its own local() names (never added to the page):
 * it loads exactly when the browser has such a font. The page's own faces
 * cannot say: a served face declared for the whole range with the same
 * descriptors (`src: local("Segoe UI"), url(selawik.woff2)` before
 * `src: local("Segoe UI")`) is indistinguishable through the FontFace API,
 * and its file loading read as an installed Segoe UI on every machine.
 */
export function localOnlyFaceState(face: FaceDeclaration): "missing" | "installed" | "pending" {
  if (typeof document === "undefined" || !document.fonts) return "missing";
  const key = shadowKey(face);
  const known = localOnly.get(key);
  if (known) return known;
  const settle = (state: "missing" | "installed") => {
    localOnly.set(key, state);
    generation += 1;
    announce();
  };
  const probe = localProbe(face);
  if (probe) {
    localOnly.set(key, "pending");
    probe.load().then(() => settle("installed"), () => settle("missing"));
    return "pending";
  }
  let match: FontFace | undefined;
  document.fonts.forEach((candidate) => {
    if (match || normalizeFamily(candidate.family) !== face.family) return;
    const style = candidate.style.startsWith("italic") ? "italic" : candidate.style.startsWith("oblique") ? "oblique" : "normal";
    if (style !== face.style) return;
    const [min, max] = parseWeight(candidate.weight);
    if (min !== face.weightMin || max !== face.weightMax) return;
    if (face.ranges === null ? !FULL_RANGE.test(candidate.unicodeRange || "U+0-10FFFF") : FULL_RANGE.test(candidate.unicodeRange)) return;
    match = candidate;
  });
  // No CSS-connected face to ask: the engine cannot tell, and keeps the
  // browser (as before local-only faces were followed).
  if (!match) {
    localOnly.set(key, "installed");
    return "installed";
  }
  if (match.status === "loaded") {
    localOnly.set(key, "installed");
    return "installed";
  }
  if (match.status === "error") {
    localOnly.set(key, "missing");
    return "missing";
  }
  localOnly.set(key, "pending");
  match.load().then(() => settle("installed"), () => settle("missing"));
  return "pending";
}

/** A detached FontFace of a local()-only face's names, or null where the browser has none to make. */
function localProbe(face: FaceDeclaration): FontFace | null {
  if (typeof FontFace !== "function" || !face.localNames.length) return null;
  const src = face.localNames.map((name) => `local("${name.replace(/["\\]/g, "\\$&")}")`).join(", ");
  try {
    return new FontFace("mrai local probe", src, { style: face.style === "normal" ? "normal" : "italic", weight: String(face.weightMin) });
  } catch {
    return null;
  }
}

const systemFamilies = new Map<string, "absent" | "present">();
const PROBE_TEXT = "mmmmmmmmmmlli1WQ中文あ한";

/**
 * Whether a family the page declares no face for is a font the browser has
 * (a system or user-installed font): a probe text drawn in "<family>,
 * monospace" and in "<family>, serif" has the fallback's own widths when the
 * name draws nothing. Asked once per family.
 */
export function systemFamilyState(family: string): "absent" | "present" {
  const known = systemFamilies.get(family);
  if (known) return known;
  let answer: "absent" | "present" = "present";
  try {
    const context = typeof document !== "undefined" ? document.createElement("canvas").getContext("2d") : null;
    if (context) {
      const quoted = `"${family.replace(/["\\]/g, "\\$&")}"`;
      const width = (stack: string) => { context.font = `100px ${stack}`; return context.measureText(PROBE_TEXT).width; };
      const mono = width("monospace"), serif = width("serif");
      if (Math.abs(width(`${quoted}, monospace`) - mono) < 0.01 && Math.abs(width(`${quoted}, serif`) - serif) < 0.01) answer = "absent";
    }
  } catch { /* keep "present": the engine stays out */ }
  systemFamilies.set(family, answer);
  return answer;
}
