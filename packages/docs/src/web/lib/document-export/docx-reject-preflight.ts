/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import {
  DOCX_EXPORT_ATOM_POLICIES,
  SERVER_DOCX_ATOM_POLICIES,
  type DocumentAtomName,
  type DocxExportAtomPolicyMap,
} from "@mrnote/core/web/lib/document-feature-capabilities.ts";
import {
  projectTrackedChanges,
  type RevisionExportPolicy,
} from "@mrnote/core/web/lib/document-export/revision-policy.ts";

export interface DocumentExportTreeNode {
  readonly type?: string;
  readonly attrs?: Readonly<Record<string, unknown>> | null;
  readonly content?: DocumentExportTreeNode[] | null;
}

export type DocxRejectSurface = "offline" | "server";

/**
 * Ephemeral attribute used only on a cloned editor snapshot so a projected
 * reject can be mapped back to the exact live ProseMirror node. It is never
 * dispatched to the editor or passed to an exporter.
 */
export const DOCX_REJECT_PREFLIGHT_LOCATOR_ATTR =
  "__mrnoteDocxRejectPreflightLocator";

export interface RejectedDocxAtom {
  readonly atomName: DocumentAtomName;
  readonly path: string;
  readonly mediaKind: "iframe" | "video" | "unknown" | null;
  readonly sourceLocator: string | null;
}

export interface RejectedDocxAtomSummary {
  readonly atomName: DocumentAtomName;
  readonly count: number;
  readonly firstPath: string;
}

// The editor nodes a document page may hold (the API's `MAX_DOCUMENT_NODES`,
// half its 600,000-node JSON envelope): a page that saves is a page whose
// export this walk reads whole. It was 100,000, which refused the Word export
// of a long document the page itself held.
const MAX_DOCX_PREFLIGHT_NODES = 300_000;
const MAX_DOCX_PREFLIGHT_DEPTH = 64;

export class DocxRejectPreflightError extends Error {
  readonly code = "docx_reject_preflight_limit";

  constructor(message: string) {
    super(message);
    this.name = "DocxRejectPreflightError";
  }
}

function policyForSurface(
  surface: DocxRejectSurface,
): DocxExportAtomPolicyMap {
  return surface === "server"
    ? SERVER_DOCX_ATOM_POLICIES
    : DOCX_EXPORT_ATOM_POLICIES;
}

/**
 * Enumerate every atom that the selected DOCX renderer must reject.
 *
 * The walk is iterative, source ordered, and bounded. Callers receive every
 * exact occurrence (rather than the old first-error-only answer), while an
 * unexpectedly deep or large tree fails closed instead of returning a partial
 * report. The server retains its independent schema/export gate.
 */
export function collectRejectedDocxAtoms(
  root: DocumentExportTreeNode,
  surface: DocxRejectSurface,
): readonly RejectedDocxAtom[] {
  const policy = policyForSurface(surface);
  const rejected: RejectedDocxAtom[] = [];
  const stack: Array<{
    readonly node: DocumentExportTreeNode;
    readonly path: string;
    readonly depth: number;
  }> = [{ node: root, path: "", depth: 0 }];
  let visited = 0;

  while (stack.length > 0) {
    const current = stack.pop()!;
    visited += 1;
    if (visited > MAX_DOCX_PREFLIGHT_NODES) {
      throw new DocxRejectPreflightError(
        "DOCX export preflight exceeded the bounded node limit.",
      );
    }
    if (current.depth > MAX_DOCX_PREFLIGHT_DEPTH) {
      throw new DocxRejectPreflightError(
        "DOCX export preflight exceeded the bounded nesting limit.",
      );
    }

    const atomName =
      typeof current.node.type === "string" &&
      current.node.type in policy
        ? (current.node.type as DocumentAtomName)
        : null;
    if (atomName && policy[atomName] === "reject") {
      const kind = current.node.attrs?.kind;
      const rawSourceLocator =
        current.node.attrs?.[DOCX_REJECT_PREFLIGHT_LOCATOR_ATTR];
      rejected.push({
        atomName,
        path: current.path || "/",
        mediaKind:
          atomName === "documentMedia"
            ? kind === "iframe" || kind === "video"
              ? kind
              : "unknown"
            : null,
        sourceLocator:
          typeof rawSourceLocator === "string" ? rawSourceLocator : null,
      });
    }

    const children = Array.isArray(current.node.content)
      ? current.node.content
      : [];
    for (let index = children.length - 1; index >= 0; index -= 1) {
      const child = children[index];
      if (!child || typeof child !== "object") continue;
      stack.push({
        node: child,
        path: `${current.path}/content/${index}`,
        depth: current.depth + 1,
      });
    }
  }

  return rejected;
}

export function collectRejectedDocxAtomsForExport(
  root: DocumentExportTreeNode,
  options: {
    readonly surface: DocxRejectSurface;
    readonly revisionPolicy: RevisionExportPolicy;
  },
): readonly RejectedDocxAtom[] {
  const projected = projectTrackedChanges(root, options.revisionPolicy);
  return collectRejectedDocxAtoms(projected, options.surface);
}

export function summarizeRejectedDocxAtoms(
  rejected: readonly RejectedDocxAtom[],
): readonly RejectedDocxAtomSummary[] {
  const summaries = new Map<DocumentAtomName, RejectedDocxAtomSummary>();
  for (const item of rejected) {
    const current = summaries.get(item.atomName);
    summaries.set(
      item.atomName,
      current
        ? { ...current, count: current.count + 1 }
        : { atomName: item.atomName, count: 1, firstPath: item.path },
    );
  }
  return Array.from(summaries.values());
}
