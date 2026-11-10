/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { noteFontFaceSource } from "@mrnote/core/web/lib/font-face-declarations.ts";
import type { PresentationTextBlock } from "@mrnote/core/web/components/console/editor/native-document/documentCommands.ts";
import type { PresentationFontAliases } from "@mrnote/core/web/lib/presentation-font-faces.ts";

export const SMARTART_CJK_POLICY_ID = "pptx-short-cjk-label-v1";
export const SMARTART_CJK_CSS_FAMILY = "MRNoteSmartArtCJK-v1";
export const SMARTART_CJK_FONT_SHA256 = "4b2be8791bb2f58f7d9c1a2977ed82b5bda0d475dc480f1a6ad7782c07858280";
export const SMARTART_CJK_FONT_BYTES = 10_595_916;
type FontStatus = "idle" | "loading" | "available" | "unavailable";
type FontState = { status: FontStatus; promise?: Promise<boolean>; listeners: Set<() => void> };
const documents = new WeakMap<Document, FontState>();
const stateFor = (doc: Document): FontState => {
  let state = documents.get(doc);
  if (!state) { state = { status: "idle", listeners: new Set() }; documents.set(doc, state); }
  return state;
};
export function packagedSmartArtFontStatus(doc: Document): FontStatus {
  return stateFor(doc).status;
}
export function subscribePackagedSmartArtFont(doc: Document, listener: () => void): () => void {
  const state = stateFor(doc);
  state.listeners.add(listener);
  return () => { state.listeners.delete(listener); };
}

/** Public product asset only, never a URL or font name supplied by page JSON.
 * The original pinned bytes match the backend's measurement face exactly.
 * Retained per document because this licensed public font has no page ACL.
 */
export function ensurePackagedSmartArtFont(doc: Document): Promise<boolean> {
  const state = stateFor(doc);
  if (state.promise) return state.promise;
  state.status = "loading";
  state.promise = (async () => {
    try {
      const response = await fetch("/fonts/NotoSansSC-Regular.ttf", { credentials: "same-origin", cache: "force-cache" });
      if (!response.ok) throw new Error("smartart_cjk_font_unavailable");
      const bytes = await response.arrayBuffer();
      if (bytes.byteLength !== SMARTART_CJK_FONT_BYTES) throw new Error("smartart_cjk_font_size");
      const digest = await crypto.subtle.digest("SHA-256", bytes);
      const hash = Array.from(new Uint8Array(digest), b => b.toString(16).padStart(2, "0")).join("");
      if (hash !== SMARTART_CJK_FONT_SHA256) throw new Error("smartart_cjk_font_hash");
      const face = new FontFace(SMARTART_CJK_CSS_FAMILY, bytes, { style: "normal", weight: "400", display: "block" });
      await face.load();
      doc.fonts.add(face);
      noteFontFaceSource(face, `url("/fonts/NotoSansSC-Regular.ttf") format("truetype")`, undefined, {
        doc,
        declared: { family: SMARTART_CJK_CSS_FAMILY, weight: "400", style: "normal" },
      });
      state.status = "available";
      return true;
    } catch {
      state.status = "unavailable";
      return false;
    } finally {
      for (const listener of state.listeners) listener();
    }
  })();
  return state.promise;
}

type FontPolicy = {
  version: 1;
  policyId: typeof SMARTART_CJK_POLICY_ID;
  kind: "packaged-cjk" | "bound-eastAsia";
  family: string;
  cssFamily: string;
  fontSha256: string;
  role: "regular" | "bold" | "italic" | "boldItalic";
  layout: "single-line";
};
const record = (value: unknown): Record<string, unknown> | undefined =>
  value !== null && typeof value === "object" && !Array.isArray(value) ? value as Record<string, unknown> : undefined;
function readPolicy(value: unknown): FontPolicy | undefined {
  const p = record(value);
  if (!p || Object.keys(p).length !== 8 || p.version !== 1 || p.policyId !== SMARTART_CJK_POLICY_ID || p.layout !== "single-line" ||
    typeof p.role !== "string" || !["regular", "bold", "italic", "boldItalic"].includes(p.role) ||
    typeof p.fontSha256 !== "string" || !/^[0-9a-f]{64}$/.test(p.fontSha256) ||
    ![p.family, p.cssFamily].every(v => typeof v === "string" && v.length > 0 && v.length <= 200 && !/[\u0000-\u001f\u007f"'\\;{}<>]/u.test(v))) return undefined;
  if (p.kind === "packaged-cjk") {
    if (p.family !== "Noto Sans SC" || p.cssFamily !== SMARTART_CJK_CSS_FAMILY || p.role !== "regular" || p.fontSha256 !== SMARTART_CJK_FONT_SHA256) return undefined;
  } else if (p.kind !== "bound-eastAsia" || p.cssFamily !== p.family) return undefined;
  return p as FontPolicy;
}

/** The one font family every run of ``block`` names, when they agree. */
function namedEastAsiaFamily(block: PresentationTextBlock): string | undefined {
  const names = new Set<string>();
  const note = (style: unknown) => {
    const value = record(style)?.fontFamily;
    if (typeof value === "string" && value.trim()) names.add(value.trim());
  };
  for (const run of block.runs ?? []) note(run.style);
  for (const paragraph of block.paragraphs ?? []) for (const run of paragraph.runs ?? []) note(run.style);
  if (!names.size) note(block.style);
  return names.size === 1 ? [...names][0] : undefined;
}

/** Use ONLY for presentationTextVisualLayout's detached input. Never persist
 * this clone or treat the display policy as source/font attachment authority.
 * Hosts must hide text while ready=false instead of painting a system fallback.
 * An old snapshot with no policy remains untouched for legacy admission.
 */
export function nativeSmartArtFontProjection(
  block: PresentationTextBlock,
  aliases: PresentationFontAliases,
  doc?: Document,
): { block: PresentationTextBlock; ready: boolean; packaged: boolean } {
  const view = record(record(block.source)?.nativeSmartArtView);
  if (!view || !Object.hasOwn(view, "fontPolicy")) return { block, ready: true, packaged: false };
  const policy = readPolicy(view.fontPolicy);
  if (!policy) return { block, ready: false, packaged: false };
  const packaged = policy.kind === "packaged-cjk";
  const family = packaged ? policy.cssFamily : aliases[policy.family.toLowerCase()] ??
    (Object.values(aliases).includes(policy.family) ? policy.family : undefined);
  const ready = packaged ? Boolean(doc && packagedSmartArtFontStatus(doc) === "available") : Boolean(family);
  // The packaged face measured the label; the words are drawn in the face
  // the file names for them when it is there, as PowerPoint draws them (a
  // deck's 等线 labels came out in the heavier packaged Noto Sans SC beside
  // PowerPoint's DengXian). Its registered stand-in is that same Noto Sans
  // SC, and a CJK character is one em in both, so the measured line holds.
  const named = packaged ? namedEastAsiaFamily(block) : undefined;
  const namedCss = named ? aliases[named.toLowerCase()] ?? named : undefined;
  const fontFamily = namedCss && /^[^"'\\;{}<>\u0000-\u001f\u007f]{1,200}$/u.test(namedCss)
    ? `"${namedCss}", ${policy.cssFamily}`
    : family ?? policy.cssFamily;
  // The established render edge already accepts fontFamily on imported block
  // styles. Keep this extension confined to a detached style value.
  const style = { ...block.style, fontFamily };
  const runs = (values: NonNullable<PresentationTextBlock["runs"]>) =>
    values.map(run => ({ ...run, style: { ...run.style, fontFamily } }));
  return {
    ready, packaged,
    block: { ...block, style,
      ...(block.runs ? { runs: runs(block.runs) } : {}),
      ...(block.paragraphs ? { paragraphs: block.paragraphs.map(p => ({ ...p,
        ...(p.runs ? { runs: runs(p.runs) } : {}) })) } : {}),
    },
  };
}
