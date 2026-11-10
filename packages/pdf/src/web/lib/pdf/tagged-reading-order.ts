/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { mcidFromPdfJsContentId } from "./actual-text";

const MAX_TAGGED_TEXT_ITEMS = 50_000;
const MAX_STRUCT_NODES = 10_000;
const MAX_STRUCT_DEPTH = 64;
const MAX_READING_TEXT_CHARS = 2_000_000;
const MIN_TAGGED_TEXT_COVERAGE = 0.7;
const CONTROL_CHARACTERS = /[\u0000-\u0008\u000b\u000c\u000e-\u001f\u007f]/gu;
/** Structure roles whose /Alt replaces the whole element (ISO 32000-1 §14.9.3). */
const ALT_SUBSTITUTED_ROLES = new Set(["figure", "formula", "form"]);






export const INLINE_ROLES: ReadonlySet<string> = new Set([
  "span",
  "quote",
  "note",
  "reference",
  "bibentry",
  "code",
  "link",
  "annot",
  "ruby",
  "rb",
  "rt",
  "rp",
  "warichu",
  "wt",
  "wp",
  "lbl",
  "lbody",
  "em",
  "strong",
  "sub",
]);
/**
 * Grouping elements (ISO 32000-1 §14.8.4.2) hold blocks, not running text:
 * content one of them references directly is a block of its own, as when
 * each referenced sequence was its own line.
 */
export const GROUPING_ROLES: ReadonlySet<string> = new Set([
  "",
  "document",
  "documentfragment",
  "part",
  "art",
  "sect",
  "div",
  "blockquote",
  "toc",
  "toci",
  "index",
  "nonstruct",
  "private",
  "l",
  "table",
  "tr",
  "thead",
  "tbody",
  "tfoot",
]);



const MAX_SUBSTITUTE_TEXT = 1_000_000;

type TaggedTextContentItem = Readonly<{
  type?: unknown;
  id?: unknown;
  tag?: unknown;
  str?: unknown;
  hasEOL?: unknown;
}>;

type TaggedStructNode = Readonly<{
  type?: unknown;
  id?: unknown;
  role?: unknown;
  alt?: unknown;
  children?: unknown;
}>;

export interface PdfTaggedReadingOrderResult {
  readonly text: string;
  /** Visible text items a structure element referenced (or substituted). */
  readonly taggedTextItemCount: number;
  /** Visible, non-artifact text items on the page. */
  readonly totalTextItemCount: number;
  



  readonly untaggedTextItemCount: number;
  /** Text items inside /Artifact marked content (pagination, decoration). */
  readonly artifactTextItemCount: number;
  /** True when every visible text item was placed by the structure tree. */
  readonly complete: boolean;
  readonly usedAltTextCount: number;
   
  readonly usedActualTextCount: number;
  readonly semanticObjectCount: number;
  readonly missingAltTextCount: number;
  





  readonly ambiguousContentIdCount: number;
  

  readonly reversedTextItemCount: number;
}

function safeText(value: unknown, maxLength: number): string {
  return typeof value === "string"
    ? value.replace(CONTROL_CHARACTERS, "").slice(0, maxLength)
    : "";
}

interface TextRun {
  readonly ownerId: string;
  readonly text: string;
}

/**
 * The logical text of an item a /ReversedChars sequence holds: each show
 * string is stored in reverse (ISO 32000-1 §14.8.2.3.3). pdf.js joins
 * neighbouring show strings into one item with a space between them, so each
 * space-separated run is turned around in place and the runs keep their
 * order: "OLLEH DLROW" reads "HELLO WORLD". Characters are reversed whole
 * (surrogate pairs and combining marks stay with their base).
 */
function unreverseShowStrings(text: string): string {
  return text
    .split(/(\s+)/u)
    .map((part) => {
      if (/^\s*$/u.test(part)) return part;
      const clusters = part.match(/\P{M}\p{M}*|\p{M}+/gu) ?? [part];
      return clusters.reverse().join("");
    })
    .join("");
}
















export function pdfTaggedReadingOrderText(input: {
  readonly structTree: unknown;
  readonly markedContentItems: readonly unknown[];
  /**
   * `/ActualText` by MCID for this page (see `actual-text.ts`): a referenced
   * sequence carrying one is read as that text instead of its glyphs.
   */
  readonly actualTextByMcid?: ReadonlyMap<number, string> | null;
}): PdfTaggedReadingOrderResult | null {
  if (
    !input.structTree ||
    typeof input.structTree !== "object" ||
    !Array.isArray(input.markedContentItems) ||
    input.markedContentItems.length > MAX_TAGGED_TEXT_ITEMS * 3
  ) {
    return null;
  }

  // Every visible, non-artifact text item in paint order with the marked
  // content id that owns it ("" when none does).
  const runs: TextRun[] = [];
  const fragmentsById = new Map<string, string[]>();
  const frames: Array<{ id: string; artifact: boolean; reversed: boolean }> =
    [];
  // How many separate sequences opened each id: pdf.js prefixes a form
  // XObject's local MCID with the page id, so two streams' "MCID 0" collide.
  const openedById = new Map<string, number>();
  let totalTextItemCount = 0;
  let artifactTextItemCount = 0;
  let reversedTextItemCount = 0;
  for (const raw of input.markedContentItems) {
    if (!raw || typeof raw !== "object") continue;
    const item = raw as TaggedTextContentItem;
    if (
      item.type === "beginMarkedContentProps" ||
      item.type === "beginMarkedContent"
    ) {
      const parentArtifact = frames.some((frame) => frame.artifact);
      const tag = safeText(item.tag, 64);
      const id =
        item.type === "beginMarkedContentProps" ? safeText(item.id, 256) : "";
      if (id) openedById.set(id, (openedById.get(id) ?? 0) + 1);
      frames.push({
        id,
        artifact: parentArtifact || tag === "Artifact",
        reversed: tag === "ReversedChars",
      });
      continue;
    }
    if (item.type === "endMarkedContent") {
      if (frames.length === 0) return null;
      frames.pop();
      continue;
    }
    let text = safeText(item.str, MAX_READING_TEXT_CHARS);
    if (!text) continue;
    if (frames.some((frame) => frame.artifact)) {
      artifactTextItemCount += 1;
      continue;
    }
    totalTextItemCount += 1;
    if (totalTextItemCount > MAX_TAGGED_TEXT_ITEMS) return null;
    // The innermost ReversedChars decides; nested ones would flip back, which
    // the standard does not define, so any enclosing one counts once.
    if (frames.some((frame) => frame.reversed)) {
      text = unreverseShowStrings(text);
      reversedTextItemCount += 1;
    }
    const ownerId = [...frames].reverse().find((frame) => frame.id)?.id ?? "";
    const run = `${text}${item.hasEOL === true ? "\n" : ""}`;
    runs.push({ ownerId, text: run });
    if (ownerId) {
      const fragments = fragmentsById.get(ownerId) ?? [];
      fragments.push(run);
      fragmentsById.set(ownerId, fragments);
    }
  }
  if (frames.length !== 0) return null;

  const ambiguousIds = new Set(
    [...openedById].filter(([, count]) => count > 1).map(([id]) => id),
  );
  const chunks: string[] = [];
  const consumedIds = new Set<string>();
  let taggedTextItemCount = 0;
  let usedAltTextCount = 0;
  let usedActualTextCount = 0;
  let semanticObjectCount = 0;
  let missingAltTextCount = 0;
  let nodeCount = 0;
  let outputCharacters = 0;
  // Text of the block being read: inline content continues it, a block
  
  let block = "";
  const flush = () => {
    const text = block.trim();
    block = "";
    if (!text) return;
    chunks.push(text);
  };
  const append = (value: string) => {
    if (!value) return;
    outputCharacters += value.length;
    if (outputCharacters > MAX_READING_TEXT_CHARS) {
      throw new Error("tagged reading text exceeds limit");
    }
    block += value;
  };
  /** A substitute (/ActualText, /Alt) read whole or not at all. */
  const substitute = (value: unknown): string | null => {
    if (typeof value !== "string") return null;
    if (value.length > MAX_SUBSTITUTE_TEXT) {
      throw new Error("tagged substitute text exceeds limit");
    }
    return value.replace(CONTROL_CHARACTERS, "");
  };
  const readContent = (node: TaggedStructNode, substituted: boolean) => {
    const id = safeText(node.id, 256);
    if (!id || consumedIds.has(id)) {
      throw new Error("invalid tagged content identity");
    }
    consumedIds.add(id);
    const fragments = fragmentsById.get(id);
    taggedTextItemCount += fragments?.length ?? 0;
    // A page-wide MCID map cannot say which of several same-id sequences
    
    const mcid =
      input.actualTextByMcid && !ambiguousIds.has(id)
        ? mcidFromPdfJsContentId(id)
        : null;
    const actualText =
      mcid === null ? undefined : input.actualTextByMcid?.get(mcid);
    if (actualText !== undefined) {
      
      if (!substituted) {
        append(substitute(actualText) ?? "");
        usedActualTextCount += 1;
      }
      return;
    }
    if (!fragments) return;
    if (!substituted) append(fragments.join(""));
  };
  // `substituted`: an ancestor's /Alt already stands for this content, so its
  // glyphs are consumed (they are placed) but not read a second time.
  const walk = (
    raw: unknown,
    depth: number,
    substituted: boolean,
    parentRole: string,
  ): void => {
    if (!raw || typeof raw !== "object" || depth > MAX_STRUCT_DEPTH) {
      throw new Error("invalid tagged structure");
    }
    nodeCount += 1;
    if (nodeCount > MAX_STRUCT_NODES) {
      throw new Error("tagged structure exceeds limit");
    }
    const node = raw as TaggedStructNode;
    if (node.type === "content") {
      // Inside a paragraph or span the sequence continues the running text;
      // straight under a grouping element it is a block of its own.
      const own = GROUPING_ROLES.has(parentRole);
      if (own) flush();
      readContent(node, substituted);
      if (own) flush();
      return;
    }
    const role = safeText(node.role, 128).toLowerCase();
    const inline = INLINE_ROLES.has(role);
    if (!inline) flush();
    const alt = substitute(node.alt) ?? "";
    let childrenSubstituted = substituted;
    if (ALT_SUBSTITUTED_ROLES.has(role)) {
      semanticObjectCount += 1;
      if (alt) {
        if (!substituted) {
          append(alt);
          usedAltTextCount += 1;
        }
        childrenSubstituted = true;
      } else {
        missingAltTextCount += 1;
      }
    }
    if (node.children !== undefined) {
      if (!Array.isArray(node.children)) {
        throw new Error("invalid tagged structure children");
      }
      for (const child of node.children) {
        walk(child, depth + 1, childrenSubstituted, inline ? parentRole : role);
      }
    }
    if (!inline) flush();
  };
  try {
    walk(input.structTree, 0, false, "");
    flush();
  } catch {
    return null;
  }
  if (
    totalTextItemCount > 0 &&
    taggedTextItemCount / totalTextItemCount < MIN_TAGGED_TEXT_COVERAGE
  ) {
    return null;
  }
  // Visible text the tree never placed (no marked content, or a marked
  // content id no structure element references) follows in paint order.
  const leftover = runs.filter(
    (run) => !run.ownerId || !consumedIds.has(run.ownerId),
  );
  const untaggedTextItemCount = leftover.length;
  if (untaggedTextItemCount > 0) {
    try {
      append(leftover.map((run) => run.text).join(""));
      flush();
    } catch {
      return null;
    }
  }
  // A page whose only text is /Artifact decoration, or only a Figure, reads
  // as verified empty text with its diagnostics, not as "could not read the
  
  return Object.freeze({
    text: chunks.join("\n"),
    taggedTextItemCount,
    totalTextItemCount,
    untaggedTextItemCount,
    artifactTextItemCount,
    complete: untaggedTextItemCount === 0,
    usedAltTextCount,
    usedActualTextCount,
    semanticObjectCount,
    missingAltTextCount,
    ambiguousContentIdCount: ambiguousIds.size,
    reversedTextItemCount,
  });
}
