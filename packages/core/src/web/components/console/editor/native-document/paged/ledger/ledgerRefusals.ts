/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { Node as PmNode } from "@tiptap/pm/model";
import { collectPages } from "../pageRebalanceExtension";
import type { LedgerBuild } from "./ledgerAtoms";

/** Enough for every refusal a document can hold at once; the oldest go first. */
const MAX_REFUSALS = 64;
/** Capacity reads of one unchanged page agree to well under a pixel. */
const CAPACITY_TOLERANCE_PX = 0.5;

/** Enough of a block's start to tell it from its neighbours. */
const OPENER_TEXT_CHARS = 48;

/**
 * The block opening a page, as the plan will keep seeing it. Not the node
 * itself: the head fragment the verifier moved is merged with its tail by the
 * very next reconcile (a new node), and that must not end the refusal.
 */
interface OpenerSignature {
  type: string;
  id: string | null;
  text: string;
}

function openerSignature(node: PmNode): OpenerSignature {
  return {
    type: node.type.name,
    id: typeof node.attrs?.block_id === "string" ? node.attrs.block_id : null,
    text: node.textContent.slice(0, OPENER_TEXT_CHARS),
  };
}

/** The same block, whole or as its head fragment (a head can be shorter). */
function sameOpener(a: OpenerSignature, b: OpenerSignature): boolean {
  return a.type === b.type && a.id === b.id && (a.text.startsWith(b.text) || b.text.startsWith(a.text));
}

interface Refusal {
  /** The page the physical verifier pushed content off, as it stood after. */
  pageNode: PmNode;
  /** The block that now opens the next page. */
  opener: OpenerSignature;
  /** The refused page's content capacity when the verifier measured it. */
  capacityPx: number;
}

/**
 * Page boundaries the physical verifier refused, which the plan keeps.
 *
 * The plan fills a page from measured atoms; the verifier then checks the
 * painted sheet and pushes a suffix that still overflows onto the next page.
 * When the plan's own measure of that suffix is wrong (a text cut inside a
 * right-to-left paragraph whose head paints two lines taller than planned,
 * D010 pages 106 and 125), the next plan put the same head back, the
 * verifier pushed it off again, and the two layouts alternated every pass
 * until the 24-pass limit -- 4-6 s each on a 158-page document, restarted by
 * every keystroke.
 *
 * Word settles a page once and keeps it. So does the ledger now: after a
 * repair, the block that opens the next page opens it in every later plan
 * (a page break before it), for as long as that page is the same node, the
 * same block follows it, and the page keeps its capacity. An edit of the page
 * (ProseMirror replaces an edited node, untouched ones are shared), of the
 * block's start, or of the page's geometry drops the refusal and the plan is
 * free again.
 */
export class LedgerPhysicalRefusals {
  private entries: Refusal[] = [];

  get size(): number {
    return this.entries.length;
  }

  clear(): void {
    this.entries = [];
  }

  /** After a repair dispatched: the pages it pushed content off, with the
   * content capacity it measured them at. */
  record(doc: PmNode, repaired: ReadonlyArray<{ sheet: number; capacityPx: number }>): void {
    const pages = collectPages(doc);
    if (!pages) return;
    for (const { sheet, capacityPx } of repaired) {
      const pageNode = pages[sheet]?.node;
      const blockNode = pages[sheet + 1]?.node.firstChild;
      if (!pageNode || !blockNode) continue;
      this.entries = this.entries.filter((entry) => entry.pageNode !== pageNode);
      this.entries.push({ pageNode, opener: openerSignature(blockNode), capacityPx });
    }
    if (this.entries.length > MAX_REFUSALS) this.entries.splice(0, this.entries.length - MAX_REFUSALS);
  }

  /**
   * The atoms (positions in `build.atoms`) each refused boundary keeps at the
   * top of a page. A refusal whose page or block changed is dropped.
   */
  breakBefore(doc: PmNode, build: LedgerBuild): ReadonlySet<number> {
    const result = new Set<number>();
    if (this.entries.length === 0) return result;
    const pages = collectPages(doc);
    if (!pages) {
      this.clear();
      return result;
    }
    const pageIndex = new Map<PmNode, number>(pages.map((page, index) => [page.node, index]));
    const atomPosition = new Map<number, number>(build.atoms.map((atom, position) => [atom.index, position]));
    // The logical block that opens each sheet, by the sheet it opens.
    const openers = new Map<number, { atom: number; node: PmNode }>();
    for (const block of build.logical) {
      const first = block.parts[0];
      const atom = atomPosition.get(block.index);
      if (first?.childIndex === 0 && atom !== undefined) openers.set(first.sheet, { atom, node: first.node });
    }
    this.entries = this.entries.filter((entry) => {
      const page = pageIndex.get(entry.pageNode);
      const opener = page === undefined ? undefined : openers.get(page + 1);
      if (page === undefined || !opener || !sameOpener(openerSignature(opener.node), entry.opener)) return false;
      if (!(Math.abs(entry.capacityPx - build.capacities[page]) <= CAPACITY_TOLERANCE_PX)) return false;
      result.add(opener.atom);
      return true;
    });
    return result;
  }
}
