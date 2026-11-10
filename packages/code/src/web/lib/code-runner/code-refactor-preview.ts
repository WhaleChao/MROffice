/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
const MAX_REFACTOR_DESCRIPTION_CHARS = 512;
const MAX_REFACTOR_TARGETS = 32;
const MAX_REFACTOR_SOURCE_BYTES = 256 * 1024;
const MAX_REFACTOR_PLAN_BYTES = 768 * 1024;

export interface CodeRefactorSourceGuard {
  readonly bufferId: string;
  readonly source: string;
  readonly authorityGeneration: number;
  readonly documentGeneration: string;
}

export interface CodeRefactorPreviewChange {
  readonly bufferId: string;
  readonly label: string;
  readonly language: string;
  readonly before: string;
  readonly after: string;
}

export interface CodeRefactorPreviewPlan {
  readonly id: string;
  readonly description: string;
  readonly guards: readonly CodeRefactorSourceGuard[];
  readonly changes: readonly CodeRefactorPreviewChange[];
  readonly totalBytes: number;
}

export interface CodeRefactorTargetSnapshot {
  readonly bufferId: string;
  readonly label: string;
  readonly language: string;
  readonly source: string;
  readonly allowed: boolean;
  readonly authorityGeneration: number;
  readonly documentGeneration: string;
  readonly bindingGeneration?: string;
}

export interface CodeRefactorPreviewInput {
  readonly id: string;
  readonly description: string;
  readonly expectedSources: ReadonlyArray<{
    readonly bufferId: string;
    readonly source: string;
  }>;
  readonly edits: ReadonlyArray<{
    readonly bufferId: string;
    readonly before: string;
    readonly after: string;
  }>;
  readonly targets: readonly CodeRefactorTargetSnapshot[];
}

export interface CodeRefactorUndoReceipt {
  readonly id: string;
  readonly description: string;
  readonly guards: readonly CodeRefactorSourceGuard[];
  readonly changes: readonly CodeRefactorPreviewChange[];
}

const encoder = new TextEncoder();

function utf8Bytes(value: string): number {
  return encoder.encode(value).byteLength;
}

function boundedIdentity(value: string, max: number): boolean {
  return value.length > 0 && value.length <= max && !value.includes("\u0000");
}

function frozenGuard(
  target: CodeRefactorTargetSnapshot,
  source: string,
): CodeRefactorSourceGuard {
  return Object.freeze({
    bufferId: target.bufferId,
    source,
    authorityGeneration: target.authorityGeneration,
    documentGeneration: target.documentGeneration,
  });
}

/**
 * Turn one compiler-materialized refactor into a bounded, immutable preview.
 *
 * The compiler's complete project snapshot remains attached as `guards`, not
 * just the files it edits. A peer that influenced the refactor can therefore
 * invalidate Apply even when the compiler did not need to rewrite that peer.
 */
export function buildCodeRefactorPreviewPlan(
  input: CodeRefactorPreviewInput,
): CodeRefactorPreviewPlan | null {
  const description = input.description.trim();
  if (
    !boundedIdentity(input.id, 160) ||
    !boundedIdentity(description, MAX_REFACTOR_DESCRIPTION_CHARS) ||
    input.expectedSources.length === 0 ||
    input.expectedSources.length > MAX_REFACTOR_TARGETS ||
    input.edits.length === 0 ||
    input.edits.length > MAX_REFACTOR_TARGETS ||
    input.targets.length > MAX_REFACTOR_TARGETS
  ) {
    return null;
  }

  const targets = new Map<string, CodeRefactorTargetSnapshot>();
  for (const target of input.targets) {
    if (
      !boundedIdentity(target.bufferId, 512) ||
      !boundedIdentity(target.label, 1_024) ||
      !boundedIdentity(target.language, 64) ||
      !target.allowed ||
      !Number.isSafeInteger(target.authorityGeneration) ||
      target.authorityGeneration < 0 ||
      !boundedIdentity(target.documentGeneration, 256) ||
      targets.has(target.bufferId)
    ) {
      return null;
    }
    targets.set(target.bufferId, target);
  }

  let totalBytes = utf8Bytes(input.id) + utf8Bytes(description);
  const guards: CodeRefactorSourceGuard[] = [];
  const guardsById = new Map<string, CodeRefactorSourceGuard>();
  for (const expected of input.expectedSources) {
    const target = targets.get(expected.bufferId);
    const sourceBytes = utf8Bytes(expected.source);
    if (
      !target ||
      guardsById.has(expected.bufferId) ||
      target.source !== expected.source ||
      sourceBytes > MAX_REFACTOR_SOURCE_BYTES
    ) {
      return null;
    }
    totalBytes += sourceBytes;
    if (totalBytes > MAX_REFACTOR_PLAN_BYTES) return null;
    const guard = frozenGuard(target, expected.source);
    guards.push(guard);
    guardsById.set(expected.bufferId, guard);
  }

  const changes: CodeRefactorPreviewChange[] = [];
  const changedIds = new Set<string>();
  for (const edit of input.edits) {
    const target = targets.get(edit.bufferId);
    const guard = guardsById.get(edit.bufferId);
    const afterBytes = utf8Bytes(edit.after);
    if (
      !target ||
      !guard ||
      changedIds.has(edit.bufferId) ||
      edit.before !== guard.source ||
      edit.before === edit.after ||
      afterBytes > MAX_REFACTOR_SOURCE_BYTES
    ) {
      return null;
    }
    totalBytes += afterBytes + utf8Bytes(target.label) + utf8Bytes(target.language);
    if (totalBytes > MAX_REFACTOR_PLAN_BYTES) return null;
    changedIds.add(edit.bufferId);
    changes.push(
      Object.freeze({
        bufferId: edit.bufferId,
        label: target.label,
        language: target.language,
        before: edit.before,
        after: edit.after,
      }),
    );
  }

  return Object.freeze({
    id: input.id,
    description,
    guards: Object.freeze(guards),
    changes: Object.freeze(changes),
    totalBytes,
  });
}

export function codeRefactorPreviewPlanIsCurrent(
  plan: CodeRefactorPreviewPlan,
  targets: readonly CodeRefactorTargetSnapshot[],
): boolean {
  const current = new Map(targets.map((target) => [target.bufferId, target]));
  return plan.guards.every((guard) => {
    const target = current.get(guard.bufferId);
    return Boolean(
      target?.allowed &&
        target.source === guard.source &&
        target.authorityGeneration === guard.authorityGeneration &&
        target.documentGeneration === guard.documentGeneration,
    );
  });
}

/** Build the one session receipt that reverses every changed file together. */
export function createCodeRefactorUndoReceipt(
  plan: CodeRefactorPreviewPlan,
): CodeRefactorUndoReceipt {
  const changed = new Set(plan.changes.map((change) => change.bufferId));
  return Object.freeze({
    id: plan.id,
    description: plan.description,
    guards: Object.freeze(
      plan.guards
        .filter((guard) => changed.has(guard.bufferId))
        .map((guard) => Object.freeze({ ...guard })),
    ),
    changes: Object.freeze(
      plan.changes.map((change) => Object.freeze({ ...change })),
    ),
  });
}

export function codeRefactorUndoReceiptIsCurrent(
  receipt: CodeRefactorUndoReceipt,
  targets: readonly CodeRefactorTargetSnapshot[],
): boolean {
  const current = new Map(targets.map((target) => [target.bufferId, target]));
  const changes = new Map(
    receipt.changes.map((change) => [change.bufferId, change]),
  );
  return receipt.guards.every((guard) => {
    const target = current.get(guard.bufferId);
    const change = changes.get(guard.bufferId);
    return Boolean(
      target?.allowed &&
        change &&
        target.source === change.after &&
        target.authorityGeneration === guard.authorityGeneration &&
        target.documentGeneration === guard.documentGeneration,
    );
  });
}
