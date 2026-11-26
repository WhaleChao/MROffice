/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
export interface NativeSaveReceiptSnapshot {
  readonly status: unknown;
  readonly adapter: unknown;
  readonly resourceId: unknown;
  readonly resourceVersionId: unknown;
  readonly outputSha256: unknown;
  readonly patchedAnchorCount: unknown;
  readonly lossEntryCount: unknown;
  readonly reason: unknown;
  /** The server proved the file already is the page and wrote it unchanged. */
  readonly verifiedUnchanged?: unknown;
}

export interface LosslessDurableNativeSaveReceiptOptions {
  readonly expectedAdapter?: string;
  readonly expectedResourceId?: string;
  readonly expectedOutputSha256?: string;
  readonly previousResourceVersionId?: string;
  /** Defaults to 1. Explicit 0 requires all three expected native identities. */
  readonly minimumPatchedAnchorCount?: number;
  /**
   * With `minimumPatchedAnchorCount: 0`, stands in for `expectedOutputSha256`:
   * the receipt itself must say the server proved the file unchanged
   * (`verified_unchanged: true`) and name a well-formed output digest. A
   * zero-anchor receipt without that proof still fails (TRIAGE 159).
   */
  readonly acceptVerifiedUnchanged?: boolean;
}

type OwnDataProperty =
  | { readonly kind: "absent" }
  | { readonly kind: "invalid" }
  | { readonly kind: "value"; readonly value: unknown };

const ABSENT_PROPERTY: OwnDataProperty = Object.freeze({ kind: "absent" });
const INVALID_PROPERTY: OwnDataProperty = Object.freeze({ kind: "invalid" });
const EMPTY_OPTIONS: LosslessDurableNativeSaveReceiptOptions = Object.freeze(
  {},
);

function isRecord(value: unknown): value is object {
  if (typeof value !== "object" || value === null) return false;
  try {
    return !Array.isArray(value);
  } catch {
    return false;
  }
}

function readOwnEnumerableDataProperty(
  source: object,
  key: string,
): OwnDataProperty {
  const descriptor = Object.getOwnPropertyDescriptor(source, key);
  if (!descriptor) return ABSENT_PROPERTY;
  if (!descriptor.enumerable || !("value" in descriptor)) {
    return INVALID_PROPERTY;
  }
  return { kind: "value", value: descriptor.value };
}

function readSnapshotValue(source: object, key: string): unknown {
  const property = readOwnEnumerableDataProperty(source, key);
  if (property.kind !== "value") return undefined;
  const value = property.value;
  if (
    value === null ||
    value === undefined ||
    typeof value === "string" ||
    typeof value === "number" ||
    typeof value === "boolean"
  ) {
    return value;
  }
  return undefined;
}

/**
 * Capture the server's native-package publication receipt without evaluating
 * caller-controlled accessors or trusting inherited values.
 */
export function readNativeSaveReceiptSnapshot(
  page: unknown,
): NativeSaveReceiptSnapshot | null {
  if (!isRecord(page)) return null;
  try {
    const patchbackProperty = readOwnEnumerableDataProperty(
      page,
      "editable_page_patchback",
    );
    if (
      patchbackProperty.kind !== "value" ||
      !isRecord(patchbackProperty.value)
    ) {
      return null;
    }
    const receipt = patchbackProperty.value;
    return Object.freeze({
      status: readSnapshotValue(receipt, "status"),
      adapter: readSnapshotValue(receipt, "adapter"),
      resourceId: readSnapshotValue(receipt, "resource_id"),
      resourceVersionId: readSnapshotValue(receipt, "resource_version_id"),
      outputSha256: readSnapshotValue(receipt, "output_sha256"),
      patchedAnchorCount: readSnapshotValue(receipt, "patched_anchor_count"),
      lossEntryCount: readSnapshotValue(receipt, "loss_entry_count"),
      reason: readSnapshotValue(receipt, "reason"),
      verifiedUnchanged: readSnapshotValue(receipt, "verified_unchanged"),
    });
  } catch {
    return null;
  }
}

function isCanonicalAdapter(value: unknown): value is string {
  return typeof value === "string" && /^[a-z][a-z0-9_-]{0,63}$/.test(value);
}

function isNonblankIdentifier(value: unknown): value is string {
  return (
    typeof value === "string" &&
    value.length > 0 &&
    value === value.trim() &&
    !value.includes("\0")
  );
}

function isSha256(value: unknown): value is string {
  return typeof value === "string" && /^[a-f\d]{64}$/i.test(value);
}

function readOption(
  options: LosslessDurableNativeSaveReceiptOptions,
  key: keyof LosslessDurableNativeSaveReceiptOptions,
): OwnDataProperty {
  return readOwnEnumerableDataProperty(options, key);
}

/**
 * Prove only the cross-editor durable receipt envelope. Page acknowledgement,
 * editor generation, writer loss accounting and native target identity remain
 * the responsibility of each editor adapter.
 */
export function isLosslessDurableNativeSaveReceipt(
  snapshot: NativeSaveReceiptSnapshot | null,
  options: LosslessDurableNativeSaveReceiptOptions = EMPTY_OPTIONS,
): boolean {
  if (!isRecord(snapshot) || !isRecord(options)) return false;
  try {
    const expectedAdapterProperty = readOption(options, "expectedAdapter");
    const expectedResourceProperty = readOption(options, "expectedResourceId");
    const expectedOutputSha256Property = readOption(
      options,
      "expectedOutputSha256",
    );
    const previousResourceVersionProperty = readOption(
      options,
      "previousResourceVersionId",
    );
    const minimumAnchorProperty = readOption(
      options,
      "minimumPatchedAnchorCount",
    );
    const acceptVerifiedUnchangedProperty = readOption(
      options,
      "acceptVerifiedUnchanged",
    );
    if (
      expectedAdapterProperty.kind === "invalid" ||
      expectedResourceProperty.kind === "invalid" ||
      expectedOutputSha256Property.kind === "invalid" ||
      previousResourceVersionProperty.kind === "invalid" ||
      minimumAnchorProperty.kind === "invalid" ||
      acceptVerifiedUnchangedProperty.kind === "invalid"
    ) {
      return false;
    }
    const acceptVerifiedUnchanged =
      acceptVerifiedUnchangedProperty.kind === "value" &&
      acceptVerifiedUnchangedProperty.value === true;

    const expectedAdapter =
      expectedAdapterProperty.kind === "value"
        ? expectedAdapterProperty.value
        : undefined;
    const expectedResourceId =
      expectedResourceProperty.kind === "value"
        ? expectedResourceProperty.value
        : undefined;
    const minimumPatchedAnchorCount =
      minimumAnchorProperty.kind === "value" &&
      minimumAnchorProperty.value !== undefined
        ? minimumAnchorProperty.value
        : 1;
    const expectedOutputSha256 =
      expectedOutputSha256Property.kind === "value"
        ? expectedOutputSha256Property.value
        : undefined;
    const previousResourceVersionId =
      previousResourceVersionProperty.kind === "value"
        ? previousResourceVersionProperty.value
        : undefined;
    if (
      (expectedAdapter !== undefined && !isCanonicalAdapter(expectedAdapter)) ||
      (expectedResourceId !== undefined &&
        !isNonblankIdentifier(expectedResourceId)) ||
      (expectedOutputSha256 !== undefined &&
        !isSha256(expectedOutputSha256)) ||
      (previousResourceVersionId !== undefined &&
        !isNonblankIdentifier(previousResourceVersionId)) ||
      !Number.isSafeInteger(minimumPatchedAnchorCount) ||
      (minimumPatchedAnchorCount as number) < 0 ||
      // A zero-change export can only prove a known original byte sequence,
      // or carry the server's own proof that the file is unchanged.
      // Keep ordinary mutation checks at the default positive anchor count.
      (minimumPatchedAnchorCount === 0 &&
        (expectedAdapter === undefined || expectedResourceId === undefined ||
          (expectedOutputSha256 === undefined && !acceptVerifiedUnchanged)))
    ) {
      return false;
    }

    const status = readOwnEnumerableDataProperty(snapshot, "status");
    const adapter = readOwnEnumerableDataProperty(snapshot, "adapter");
    const resourceId = readOwnEnumerableDataProperty(snapshot, "resourceId");
    const resourceVersionId = readOwnEnumerableDataProperty(
      snapshot,
      "resourceVersionId",
    );
    const patchedAnchorCount = readOwnEnumerableDataProperty(
      snapshot,
      "patchedAnchorCount",
    );
    const outputSha256 = readOwnEnumerableDataProperty(
      snapshot,
      "outputSha256",
    );
    const lossEntryCount = readOwnEnumerableDataProperty(
      snapshot,
      "lossEntryCount",
    );
    const verifiedUnchanged = readOwnEnumerableDataProperty(
      snapshot,
      "verifiedUnchanged",
    );
    // Zero anchors with no expected digest: only the server's proof that the
    // file already is the page, on a receipt naming a well-formed digest.
    if (
      minimumPatchedAnchorCount === 0 &&
      expectedOutputSha256 === undefined &&
      !(
        verifiedUnchanged.kind === "value" &&
        verifiedUnchanged.value === true &&
        outputSha256.kind === "value" &&
        isSha256(outputSha256.value)
      )
    ) {
      return false;
    }
    if (
      status.kind !== "value" ||
      adapter.kind !== "value" ||
      resourceId.kind !== "value" ||
      resourceVersionId.kind !== "value" ||
      patchedAnchorCount.kind !== "value" ||
      lossEntryCount.kind !== "value"
    ) {
      return false;
    }

    return (
      status.value === "succeeded" &&
      isCanonicalAdapter(adapter.value) &&
      (expectedAdapter === undefined || adapter.value === expectedAdapter) &&
      isNonblankIdentifier(resourceId.value) &&
      (expectedResourceId === undefined ||
        resourceId.value === expectedResourceId) &&
      isNonblankIdentifier(resourceVersionId.value) &&
      (previousResourceVersionId === undefined ||
        resourceVersionId.value !== previousResourceVersionId) &&
      (expectedOutputSha256 === undefined ||
        (outputSha256.kind === "value" &&
          isSha256(outputSha256.value) &&
          outputSha256.value.toLowerCase() ===
            expectedOutputSha256.toLowerCase())) &&
      lossEntryCount.value === 0 &&
      Number.isSafeInteger(patchedAnchorCount.value) &&
      (patchedAnchorCount.value as number) >=
        (minimumPatchedAnchorCount as number)
    );
  } catch {
    return false;
  }
}
