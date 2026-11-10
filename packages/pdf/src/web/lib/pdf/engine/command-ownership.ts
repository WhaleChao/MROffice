/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import {
  findPdfFeatureCapability,
  type PdfCanonicalId,
} from "../feature-capabilities";
import type { PdfEngineAuthority, PdfEngineOutput } from "./contracts";

export type PdfEngineMutationKind =
  "none" | "shared-state" | "native-bytes" | "external-effect";

export interface PdfCommandOwnershipEntry {
  readonly featureId: PdfCanonicalId;
  readonly action: string;
  readonly authority: PdfEngineAuthority;
  readonly mutation: PdfEngineMutationKind;
  readonly output: PdfEngineOutput;
}

export interface PdfCommandOwnershipManifest {
  readonly entries: readonly PdfCommandOwnershipEntry[];
  find(
    featureId: PdfCanonicalId,
    action: string,
  ): PdfCommandOwnershipEntry | null;
}

const AUTHORITIES = new Set<PdfEngineAuthority>([
  "mrai-native",
  "apryse-client",
  "apryse-worker",
  "mrai-service",
  "browser-platform",
  "legacy-unmigrated",
  "blocked",
]);

const MUTATIONS = new Set<PdfEngineMutationKind>([
  "none",
  "shared-state",
  "native-bytes",
  "external-effect",
]);

const OUTPUTS = new Set<PdfEngineOutput>([
  "view",
  "current-pdf",
  "new-artifact",
  "external-effect",
]);

function ownershipKey(featureId: PdfCanonicalId, action: string): string {
  return `${featureId}\u0000${action}`;
}

function readAction(value: unknown): string | null {
  if (typeof value !== "string") return null;
  const action = value.trim();
  return action.length > 0 && action.length <= 160 ? action : null;
}

export function readPdfCommandOwnershipManifest(
  entries: readonly PdfCommandOwnershipEntry[],
): PdfCommandOwnershipManifest {
  const owners = new Map<string, PdfCommandOwnershipEntry>();
  const normalized: PdfCommandOwnershipEntry[] = [];

  for (const entry of entries) {
    const feature = findPdfFeatureCapability(entry.featureId);
    const action = readAction(entry.action);
    if (
      !feature ||
      !action ||
      !AUTHORITIES.has(entry.authority) ||
      !MUTATIONS.has(entry.mutation) ||
      !OUTPUTS.has(entry.output)
    ) {
      throw new Error("Invalid PDF command ownership entry.");
    }

    const key = ownershipKey(feature.id, action);
    if (owners.has(key)) {
      throw new Error(
        `Duplicate PDF command writer for ${feature.id}:${action}.`,
      );
    }

    const owner = Object.freeze({
      featureId: feature.id,
      action,
      authority: entry.authority,
      mutation: entry.mutation,
      output: entry.output,
    });
    owners.set(key, owner);
    normalized.push(owner);
  }

  const frozenEntries = Object.freeze(normalized);
  return Object.freeze({
    entries: frozenEntries,
    find(featureId: PdfCanonicalId, action: string) {
      const normalizedAction = readAction(action);
      if (!normalizedAction) return null;
      return owners.get(ownershipKey(featureId, normalizedAction)) ?? null;
    },
  });
}

export function pdfCommandRequiresWritable(
  owner: PdfCommandOwnershipEntry,
): boolean {
  return owner.mutation === "shared-state" || owner.mutation === "native-bytes";
}
