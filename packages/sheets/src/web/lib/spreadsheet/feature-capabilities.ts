/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import {
  readFeatureCommandBindingManifest,
  type FeatureCommandBindingContract,
  type FeatureCommandBindingManifest,
  type FeatureCommandBindingRegistry,
  type FeatureCommandRegistryLineage,
} from "@mrnote/core/web/components/office-ribbon/featureCommandRegistry.ts";

/**
 * Versioned, fail-closed capability contract for Spreadsheet controls.
 *
 * This module deliberately contains no product catalog yet. A feature package
 * must provide one complete row per exact Atlas Control and prove its live
 * command bindings before the row can be executable. Missing rows and missing
 * dimensions resolve to `unknown`; no helper default upgrades them to native.
 *
 * Keep UI placement, command/domain ownership, import, both XLSX save routes,
 * and Excel/WPS outcomes separate. A working editor handler is not evidence
 * that either serializer can preserve the same feature.
 */

export const SPREADSHEET_CAPABILITY_CATALOG_SCHEMA_VERSION =
  "spreadsheet-capability-catalog-v1" as const;

/** Higher than the current 1,043-Control Atlas, but still hostile-input safe. */
export const MAX_SPREADSHEET_CAPABILITY_ROWS = 2_048;
export const MAX_SPREADSHEET_HANDLER_BINDINGS = MAX_SPREADSHEET_CAPABILITY_ROWS;
const MAX_LIST_ITEMS = 256;
const MAX_TEXT_CHARS = 4_096;

export const SPREADSHEET_WEB_ATLAS_COVERAGE_SCHEMA_VERSION =
  "spreadsheet-web-atlas-coverage-v1" as const;







export const SPREADSHEET_WEB_SPARKLINE_AUTHORING_ENABLED = true;
export const SPREADSHEET_WEB_SPARKLINE_EVIDENCE = Object.freeze({
  supportLevel: "unverified" as const,
  reference:
    "WPS_PDF与Office_Web顶栏Tab审查_2026-08-25.docx#P0101",
});

/**
 * Camera has no frozen authoring command in the approved Excel-for-the-web
 * base-tab fixture. Its model/import/round-trip code remains for
 * compatibility, but the Web Ribbon and compact All Tools surface must not
 * expose desktop-only writers or disabled museum entries.
 *
 * SmartArt is no longer one of those: the shared Office SmartArt layer
 * (docs/office-graphics/UNIFICATION-PLAN.md) writes every catalog layout as
 * ordinary worksheet drawing shapes, which export as native `xdr:sp` and
 * reimport as editable objects, so Insert ▸ Illustrations ▸ SmartArt is a real
 * authoring command here — the same one Word and PowerPoint offer — rather
 * than a disabled museum entry.
 */
export const SPREADSHEET_WEB_SMARTART_AUTHORING_ENABLED = true;
export const SPREADSHEET_WEB_CAMERA_AUTHORING_ENABLED = false;
export const SPREADSHEET_WEB_EXCLUDED_AUTHORING_EVIDENCE = Object.freeze({
  smartArt: "mr-spreadsheet-web-guide-2026-08-25:insert-no-smartart",
  camera: "mr-spreadsheet-web-guide-2026-08-25:P0196-camera",
});

export const SPREADSHEET_WEB_ATLAS_TABS = Object.freeze([
  "Home",
  "Insert",
  "Share",
  "Page Layout",
  "Formulas",
  "Data",
  "Review",
  "View",
  "Draw",
] as const);

export type SpreadsheetWebAtlasTab =
  (typeof SPREADSHEET_WEB_ATLAS_TABS)[number];

export type SpreadsheetCoverageEvidenceProvenance =
  | "mrai-renderer"
  | "native-mac-observed"
  | "official-spec"
  | "official-web";

export interface SpreadsheetCoverageEvidenceRef {
  readonly provenance: SpreadsheetCoverageEvidenceProvenance;
  readonly ref: string;
}

export interface SpreadsheetControlIdGroup {
  readonly controlIds: readonly string[];
}

export type SpreadsheetCanonicalParentEvidenceDisposition =
  | "verified/tree-def"
  | "unverified/no-canonical-evidence";

export interface SpreadsheetCanonicalParentControlGroup
  extends SpreadsheetControlIdGroup {
  readonly canonicalParentId: string;
  /** Independent tree proof, or an explicit statement that none exists. */
  readonly evidenceDisposition: SpreadsheetCanonicalParentEvidenceDisposition;
}

export type SpreadsheetOfficialWebEvidenceScope =
  | "exact-control"
  | "feature-level"
  | "generic-m365";

export type SpreadsheetOfficialWebEvidenceSupportLevel =
  | "verified"
  | "unverified"
  | "official-absent"
  /**
   * The web edition has the command, observed in the product itself, because
   * Microsoft's published pages do not say so yet — they can lag a shipped
   * ribbon by months, and one of them still calls a feature that is already on
   * the Data tab "coming soon". Its artifact is a dated capture of the tab, not
   * a page excerpt, so it carries a product URL instead of a support one. It
   * counts as evidence that the web edition offers the command, exactly like
   * `verified`; it is a separate level so a reader can see which claims rest on
   * a document and which on a screenshot.
   */
  | "product-observed";

export type SpreadsheetOfficialWebEvidenceArtifactKind =
  | "page-scope"
  | "redirect-observation"
  /** A dated capture of the live ribbon, named in `sourceScope`. */
  | "product-capture";

export interface SpreadsheetOfficialWebEvidenceGroup
  extends SpreadsheetControlIdGroup {
  readonly ref: string;
  readonly requestedUrl: string;
  readonly reviewedUrl: string;
  readonly pageTitle: string;
  readonly product: string;
  readonly platform: string;
  readonly artifactKind: SpreadsheetOfficialWebEvidenceArtifactKind;
  /** Bounded wording or section scope actually reviewed at `reviewedUrl`. */
  readonly sourceScope: string;
  /** Short artifact excerpt or redirect observation covered by the hash. */
  readonly sourceArtifact: string;
  readonly evidenceScope: SpreadsheetOfficialWebEvidenceScope;
  readonly supportLevel: SpreadsheetOfficialWebEvidenceSupportLevel;
  readonly scopeSha256: string;
}

export type SpreadsheetRenderedControlOccurrenceRole =
  | "semantic"
  | "auxiliary-session"
  | "unverified-source-only";

/**
 * A physical control is identified by its surface-local occurrence key, not
 * by semantic Control ID alone. The same command may intentionally appear on
 * the primary Home ribbon and in the compact All Tools surface.
 */
export interface SpreadsheetRenderedControlOccurrence {
  readonly tab: SpreadsheetWebAtlasTab;
  readonly surface: string;
  readonly occurrenceId: string;
  readonly controlId?: string;
  readonly role: SpreadsheetRenderedControlOccurrenceRole;
}

const SPREADSHEET_OCCURRENCE_TOKEN_PATTERN = /^[a-z0-9][a-z0-9._-]*$/;

export function defineSpreadsheetRenderedControlOccurrences(
  occurrences: readonly SpreadsheetRenderedControlOccurrence[],
): readonly SpreadsheetRenderedControlOccurrence[] {
  const seenSurfaceOccurrences = new Set<string>();
  for (const occurrence of occurrences) {
    const identity = `${occurrence.surface}\u0000${occurrence.occurrenceId}`;
    if (
      !SPREADSHEET_WEB_ATLAS_TABS.includes(occurrence.tab) ||
      !SPREADSHEET_OCCURRENCE_TOKEN_PATTERN.test(occurrence.surface) ||
      !SPREADSHEET_OCCURRENCE_TOKEN_PATTERN.test(occurrence.occurrenceId) ||
      seenSurfaceOccurrences.has(identity) ||
      (occurrence.controlId !== undefined &&
        !SPREADSHEET_CONTROL_ID_PATTERN.test(occurrence.controlId)) ||
      (occurrence.role === "semantic" && !occurrence.controlId)
    ) {
      throw new Error("Invalid spreadsheet rendered control occurrence");
    }
    seenSurfaceOccurrences.add(identity);
  }
  return Object.freeze(
    occurrences.map((occurrence) => Object.freeze({ ...occurrence })),
  );
}

/**
 * Keep source catalogs, verified Web evidence, the current MRAI renderer and
 * executable command authority separate. Native Excel-for-Mac observations
 * are useful catalog evidence, but are never promoted to an Excel-for-the-Web
 * physical denominator without an explicit `official-web` evidence ref.
 */
export interface SpreadsheetWebAtlasTabCoverage {
  readonly canonicalParentControlGroups: readonly SpreadsheetCanonicalParentControlGroup[];
  readonly nativeMacEvidenceRef: string;
  readonly nativeMacObservedControlIds: readonly string[];
  readonly officialWebCoverageCompleteness: "complete" | "partial";
  readonly officialWebEvidenceGroups: readonly SpreadsheetOfficialWebEvidenceGroup[];
  readonly mraiRenderedSemanticControlIds: readonly string[];
  readonly mraiRendererRef: string;
  readonly dynamicAggregateRootIds: readonly string[];
  readonly catalogDefinitionControlIds: readonly string[];
  readonly renderAuthorityControlIds: readonly string[];
  readonly handlerAuthorityControlIds: readonly string[];
  readonly mraiRenderedWithoutAuthorityControlIds: readonly string[];
  readonly mraiRenderedBoundaryControlIds: readonly string[];
  readonly openSupportableGapControlIds: readonly string[];
  readonly unverifiedSourceOnlyGapControlIds: readonly string[];
  readonly unverifiedSourceOnlyBoundaryControlIds: readonly string[];
  readonly externalProviderBoundaryControlIds: readonly string[];
  readonly officialWebAbsentBoundaryControlIds: readonly string[];
  readonly sourceCatalogOnlyControlIds: readonly string[];
  /**
   * Catalogued controls this TAB does not render. Tab-scoped, not workbook-wide:
   * a control listed here may still be reachable from a shared surface — the
   * Formulas name and auditing entries all render inside Home > All Tools — and
   * that is not a contradiction. Reading it as "renders nowhere" has already
   * produced one false audit finding, so verify a change against the live
   * ribbon, tab by tab (open the tab, read every [data-control-id], then open
   * each menu on it), rather than against a source-wide search for the id.
   */
  readonly catalogOnlyHiddenControlIds: readonly string[];
}

export interface SpreadsheetWebAtlasCoverageOverlay {
  readonly schemaVersion: typeof SPREADSHEET_WEB_ATLAS_COVERAGE_SCHEMA_VERSION;
  readonly tabs: Readonly<
    Record<SpreadsheetWebAtlasTab, SpreadsheetWebAtlasTabCoverage>
  >;
}

const SPREADSHEET_ATLAS_CANONICAL_ID_PATTERN = /^XLS(?:-[A-Z0-9]+)+-\d{2}$/;
const SPREADSHEET_CANONICAL_PARENT_EVIDENCE_DISPOSITIONS = Object.freeze([
  "verified/tree-def",
  "unverified/no-canonical-evidence",
] as const satisfies readonly SpreadsheetCanonicalParentEvidenceDisposition[]);
const SPREADSHEET_OFFICIAL_WEB_EVIDENCE_SCOPES = Object.freeze([
  "exact-control",
  "feature-level",
  "generic-m365",
] as const satisfies readonly SpreadsheetOfficialWebEvidenceScope[]);
const SPREADSHEET_OFFICIAL_WEB_EVIDENCE_SUPPORT_LEVELS = Object.freeze([
  "verified",
  "unverified",
  "official-absent",
  "product-observed",
] as const satisfies readonly SpreadsheetOfficialWebEvidenceSupportLevel[]);
const SPREADSHEET_OFFICIAL_WEB_EVIDENCE_ARTIFACT_KINDS = Object.freeze([
  "page-scope",
  "redirect-observation",
  "product-capture",
] as const satisfies readonly SpreadsheetOfficialWebEvidenceArtifactKind[]);
const SPREADSHEET_OFFICIAL_SUPPORT_URL_PATTERN =
  /^https:\/\/support\.microsoft\.com\//;
/** Where the live web edition is reached, for a product-observed capture. */
const SPREADSHEET_OFFICIAL_PRODUCT_URL_PATTERN =
  /^https:\/\/(excel\.cloud\.microsoft|www\.microsoft365\.com|www\.office\.com)\//;
const SHA256_HEX_PATTERN = /^[a-f0-9]{64}$/;

export function defineSpreadsheetWebAtlasCoverageOverlay(
  overlay: SpreadsheetWebAtlasCoverageOverlay,
): SpreadsheetWebAtlasCoverageOverlay {
  if (
    overlay.schemaVersion !== SPREADSHEET_WEB_ATLAS_COVERAGE_SCHEMA_VERSION
  ) {
    throw new Error("Unsupported spreadsheet Web Atlas coverage version");
  }
  const seenRenderedControlIds = new Set<string>();
  const tabs = Object.create(null) as Record<
    SpreadsheetWebAtlasTab,
    SpreadsheetWebAtlasTabCoverage
  >;
  for (const tab of SPREADSHEET_WEB_ATLAS_TABS) {
    const source = overlay.tabs[tab];
    if (!source) throw new Error(`Missing spreadsheet Web Atlas tab: ${tab}`);
    const rendered = [...source.mraiRenderedSemanticControlIds];
    const renderedPartitions = [
      ...source.renderAuthorityControlIds,
      ...source.mraiRenderedWithoutAuthorityControlIds,
      ...source.mraiRenderedBoundaryControlIds,
    ];
    if (
      rendered.length === 0 ||
      new Set(rendered).size !== rendered.length ||
      new Set(renderedPartitions).size !== renderedPartitions.length ||
      rendered.length !== renderedPartitions.length ||
      rendered.some((controlId) => !renderedPartitions.includes(controlId))
    ) {
      throw new Error(`Incomplete MRAI rendered-control coverage: ${tab}`);
    }
    const controlSets = [
      source.nativeMacObservedControlIds,
      ...source.officialWebEvidenceGroups.map((group) => group.controlIds),
      source.mraiRenderedSemanticControlIds,
      source.dynamicAggregateRootIds,
      source.catalogDefinitionControlIds,
      source.renderAuthorityControlIds,
      source.handlerAuthorityControlIds,
      source.mraiRenderedWithoutAuthorityControlIds,
      source.mraiRenderedBoundaryControlIds,
      source.openSupportableGapControlIds,
      source.unverifiedSourceOnlyGapControlIds,
      source.unverifiedSourceOnlyBoundaryControlIds,
      source.externalProviderBoundaryControlIds,
      source.officialWebAbsentBoundaryControlIds,
      source.sourceCatalogOnlyControlIds,
      source.catalogOnlyHiddenControlIds,
    ];
    for (const controlIds of controlSets) {
      if (
        new Set(controlIds).size !== controlIds.length ||
        controlIds.some(
          (controlId) => !SPREADSHEET_CONTROL_ID_PATTERN.test(controlId),
        )
      ) {
        throw new Error(`Invalid spreadsheet Web Atlas Control ID set: ${tab}`);
      }
    }
    for (const controlId of rendered) {
      if (seenRenderedControlIds.has(controlId)) {
        throw new Error(`Duplicate spreadsheet Web Atlas Control ID: ${controlId}`);
      }
      seenRenderedControlIds.add(controlId);
    }
    const renderedSet = new Set(rendered);
    const catalogSet = new Set(source.catalogDefinitionControlIds);
    // Both levels answer the same question — does the web edition have this —
    // so a control observed in the product carries the same weight here as one
    // read off a documented page.
    const officialWebIds = source.officialWebEvidenceGroups
      .filter(
        (group) =>
          group.supportLevel === "verified" ||
          group.supportLevel === "product-observed",
      )
      .flatMap((group) => group.controlIds);
    const officialWebAbsentEvidenceIds =
      source.officialWebEvidenceGroups
        .filter((group) => group.supportLevel === "official-absent")
        .flatMap((group) => group.controlIds);
    const officialWebPartitions = [
      ...source.renderAuthorityControlIds.filter((id) =>
        officialWebIds.includes(id),
      ),
      ...source.openSupportableGapControlIds.filter((id) =>
        officialWebIds.includes(id),
      ),
      ...source.externalProviderBoundaryControlIds.filter((id) =>
        officialWebIds.includes(id),
      ),
    ];
    const seenOfficialWebEvidenceRefs = new Set<string>();
    const invalidOfficialWebEvidence = source.officialWebEvidenceGroups.some(
      (group) => {
        // A capture of the live ribbon is read the same way a documented page
        // is: it says the web edition has the command, so it must name which
        // controls it saw — and it is reached at the product, not on the
        // support site, so its URL answers to a different pattern.
        const productObserved = group.supportLevel === "product-observed";
        const expectedEvidenceUrlPattern = productObserved
          ? SPREADSHEET_OFFICIAL_PRODUCT_URL_PATTERN
          : SPREADSHEET_OFFICIAL_SUPPORT_URL_PATTERN;
        const positiveExactControlEvidence =
          group.evidenceScope === "exact-control" &&
          (group.supportLevel === "verified" || productObserved);
        const exactOfficialAbsence =
          group.evidenceScope === "exact-control" &&
          group.supportLevel === "official-absent";
        const duplicateRef = seenOfficialWebEvidenceRefs.has(group.ref);
        seenOfficialWebEvidenceRefs.add(group.ref);
        return (
          !group.ref ||
          duplicateRef ||
          !expectedEvidenceUrlPattern.test(group.requestedUrl) ||
          !expectedEvidenceUrlPattern.test(group.reviewedUrl) ||
          !group.pageTitle?.trim() ||
          !group.product?.trim() ||
          !group.platform?.trim() ||
          !SPREADSHEET_OFFICIAL_WEB_EVIDENCE_ARTIFACT_KINDS.includes(
            group.artifactKind,
          ) ||
          !group.sourceScope?.trim() ||
          !group.sourceArtifact?.trim() ||
          !SPREADSHEET_OFFICIAL_WEB_EVIDENCE_SCOPES.includes(
            group.evidenceScope,
          ) ||
          !SPREADSHEET_OFFICIAL_WEB_EVIDENCE_SUPPORT_LEVELS.includes(
            group.supportLevel,
          ) ||
          !SHA256_HEX_PATTERN.test(group.scopeSha256) ||
          (positiveExactControlEvidence || exactOfficialAbsence
            ? group.controlIds.length === 0
            : group.controlIds.length !== 0) ||
          (group.artifactKind === "redirect-observation" &&
            (group.supportLevel !== "unverified" ||
              group.evidenceScope !== "generic-m365"))
        );
      },
    );
    if (
      new Set(officialWebIds).size !== officialWebIds.length ||
      new Set(officialWebPartitions).size !== officialWebPartitions.length ||
      officialWebIds.length !== officialWebPartitions.length ||
      officialWebIds.some((id) => !officialWebPartitions.includes(id)) ||
      source.officialWebEvidenceGroups.length === 0 ||
      invalidOfficialWebEvidence
    ) {
      throw new Error(`Incomplete official Web coverage: ${tab}`);
    }
    const nativeDispositionIds = [
      ...officialWebIds.filter((id) =>
        source.nativeMacObservedControlIds.includes(id) &&
        !source.externalProviderBoundaryControlIds.includes(id),
      ),
      ...source.externalProviderBoundaryControlIds.filter((id) =>
        source.nativeMacObservedControlIds.includes(id),
      ),
      ...source.officialWebAbsentBoundaryControlIds.filter((id) =>
        source.nativeMacObservedControlIds.includes(id),
      ),
      ...source.renderAuthorityControlIds.filter(
        (id) =>
          source.nativeMacObservedControlIds.includes(id) &&
          !officialWebIds.includes(id) &&
          !source.externalProviderBoundaryControlIds.includes(id) &&
          !source.officialWebAbsentBoundaryControlIds.includes(id),
      ),
      ...source.unverifiedSourceOnlyGapControlIds.filter((id) =>
        source.nativeMacObservedControlIds.includes(id),
      ),
      ...source.unverifiedSourceOnlyBoundaryControlIds.filter((id) =>
        source.nativeMacObservedControlIds.includes(id),
      ),
      ...source.sourceCatalogOnlyControlIds,
    ];
    if (
      new Set(nativeDispositionIds).size !== nativeDispositionIds.length ||
      nativeDispositionIds.length !== source.nativeMacObservedControlIds.length ||
      source.nativeMacObservedControlIds.some(
        (id) => !nativeDispositionIds.includes(id),
      )
    ) {
      throw new Error(`Incomplete native Mac source disposition: ${tab}`);
    }
    if (
      source.dynamicAggregateRootIds.some((id) => !renderedSet.has(id)) ||
      source.renderAuthorityControlIds.some(
        (id) => !renderedSet.has(id) || !catalogSet.has(id),
      ) ||
      source.handlerAuthorityControlIds.some(
        (id) => !source.renderAuthorityControlIds.includes(id),
      ) ||
      source.mraiRenderedWithoutAuthorityControlIds.some(
        (id) =>
          !source.openSupportableGapControlIds.includes(id) &&
          !source.unverifiedSourceOnlyGapControlIds.includes(id),
      ) ||
      source.openSupportableGapControlIds.some(
        (id) => !officialWebIds.includes(id),
      ) ||
      source.unverifiedSourceOnlyGapControlIds.some(
        (id) => officialWebIds.includes(id) || !renderedSet.has(id),
      ) ||
      source.catalogOnlyHiddenControlIds.some(
        (id) => !catalogSet.has(id) || renderedSet.has(id),
      ) ||
      source.mraiRenderedBoundaryControlIds.some(
        (id) =>
          !source.externalProviderBoundaryControlIds.includes(id) &&
          !source.officialWebAbsentBoundaryControlIds.includes(id) &&
          !source.unverifiedSourceOnlyBoundaryControlIds.includes(id),
      ) ||
      source.unverifiedSourceOnlyBoundaryControlIds.some(
        (id) =>
          officialWebIds.includes(id) ||
          source.officialWebAbsentBoundaryControlIds.includes(id) ||
          source.externalProviderBoundaryControlIds.includes(id),
      ) ||
      source.officialWebAbsentBoundaryControlIds.some((id) =>
        officialWebIds.includes(id),
      ) ||
      new Set(officialWebAbsentEvidenceIds).size !==
        officialWebAbsentEvidenceIds.length ||
      officialWebAbsentEvidenceIds.some(
        (id) => !source.officialWebAbsentBoundaryControlIds.includes(id),
      ) ||
      officialWebAbsentEvidenceIds.length !==
        source.officialWebAbsentBoundaryControlIds.length ||
      source.officialWebAbsentBoundaryControlIds.some(
        (id) => !officialWebAbsentEvidenceIds.includes(id),
      )
    ) {
      throw new Error(`Invalid spreadsheet Web Atlas authority join: ${tab}`);
    }
    const parentControlIds: string[] = [];
    const canonicalParentByControlId = new Map<
      string,
      { readonly canonicalParentId: string; readonly evidenceDisposition: SpreadsheetCanonicalParentEvidenceDisposition }
    >();
    for (const group of source.canonicalParentControlGroups) {
      if (
        group.controlIds.length === 0 ||
        !SPREADSHEET_ATLAS_CANONICAL_ID_PATTERN.test(group.canonicalParentId) ||
        !SPREADSHEET_CANONICAL_PARENT_EVIDENCE_DISPOSITIONS.includes(
          group.evidenceDisposition,
        )
      ) {
        throw new Error(`Invalid spreadsheet Web Atlas parent group: ${tab}`);
      }
      for (const controlId of group.controlIds) {
        const existing = canonicalParentByControlId.get(controlId);
        if (existing) {
          const kind =
            existing.canonicalParentId === group.canonicalParentId &&
            existing.evidenceDisposition === group.evidenceDisposition
              ? "duplicate"
              : "conflicting";
          throw new Error(
            `${kind} spreadsheet Web Atlas parent mapping: ${controlId}`,
          );
        }
        canonicalParentByControlId.set(controlId, {
          canonicalParentId: group.canonicalParentId,
          evidenceDisposition: group.evidenceDisposition,
        });
        parentControlIds.push(controlId);
      }
    }
    const parentCoverageIds = new Set([
      ...source.nativeMacObservedControlIds,
      ...officialWebIds,
      ...source.mraiRenderedSemanticControlIds,
      ...source.catalogDefinitionControlIds,
    ]);
    if (
      source.canonicalParentControlGroups.length === 0 ||
      parentControlIds.length !== parentCoverageIds.size ||
      [...parentCoverageIds].some(
        (id) => !canonicalParentByControlId.has(id),
      )
    ) {
      throw new Error(`Invalid spreadsheet Web Atlas parent set: ${tab}`);
    }
    if (
      !source.nativeMacEvidenceRef ||
      !source.mraiRendererRef
    ) {
      throw new Error(`Missing spreadsheet Web Atlas evidence: ${tab}`);
    }
    tabs[tab] = Object.freeze({
      canonicalParentControlGroups: Object.freeze(
        source.canonicalParentControlGroups.map((group) =>
          Object.freeze({
            canonicalParentId: group.canonicalParentId,
            evidenceDisposition: group.evidenceDisposition,
            controlIds: Object.freeze([...group.controlIds]),
          }),
        ),
      ),
      nativeMacEvidenceRef: source.nativeMacEvidenceRef,
      nativeMacObservedControlIds: Object.freeze([
        ...source.nativeMacObservedControlIds,
      ]),
      officialWebCoverageCompleteness:
        source.officialWebCoverageCompleteness,
      officialWebEvidenceGroups: Object.freeze(
        source.officialWebEvidenceGroups.map((group) =>
          Object.freeze({
            ref: group.ref,
            requestedUrl: group.requestedUrl,
            reviewedUrl: group.reviewedUrl,
            pageTitle: group.pageTitle,
            product: group.product,
            platform: group.platform,
            artifactKind: group.artifactKind,
            sourceScope: group.sourceScope,
            sourceArtifact: group.sourceArtifact,
            evidenceScope: group.evidenceScope,
            supportLevel: group.supportLevel,
            scopeSha256: group.scopeSha256,
            controlIds: Object.freeze([...group.controlIds]),
          }),
        ),
      ),
      mraiRenderedSemanticControlIds: Object.freeze(rendered),
      mraiRendererRef: source.mraiRendererRef,
      dynamicAggregateRootIds: Object.freeze([
        ...source.dynamicAggregateRootIds,
      ]),
      catalogDefinitionControlIds: Object.freeze([
        ...source.catalogDefinitionControlIds,
      ]),
      renderAuthorityControlIds: Object.freeze([
        ...source.renderAuthorityControlIds,
      ]),
      handlerAuthorityControlIds: Object.freeze([
        ...source.handlerAuthorityControlIds,
      ]),
      mraiRenderedWithoutAuthorityControlIds: Object.freeze([
        ...source.mraiRenderedWithoutAuthorityControlIds,
      ]),
      mraiRenderedBoundaryControlIds: Object.freeze([
        ...source.mraiRenderedBoundaryControlIds,
      ]),
      openSupportableGapControlIds: Object.freeze([
        ...source.openSupportableGapControlIds,
      ]),
      unverifiedSourceOnlyGapControlIds: Object.freeze([
        ...source.unverifiedSourceOnlyGapControlIds,
      ]),
      unverifiedSourceOnlyBoundaryControlIds: Object.freeze([
        ...source.unverifiedSourceOnlyBoundaryControlIds,
      ]),
      externalProviderBoundaryControlIds: Object.freeze([
        ...source.externalProviderBoundaryControlIds,
      ]),
      officialWebAbsentBoundaryControlIds: Object.freeze([
        ...source.officialWebAbsentBoundaryControlIds,
      ]),
      sourceCatalogOnlyControlIds: Object.freeze([
        ...source.sourceCatalogOnlyControlIds,
      ]),
      catalogOnlyHiddenControlIds: Object.freeze([
        ...source.catalogOnlyHiddenControlIds,
      ]),
    });
  }
  if (Object.keys(overlay.tabs).length !== SPREADSHEET_WEB_ATLAS_TABS.length) {
    throw new Error("Spreadsheet Web Atlas overlay declares an unknown tab");
  }
  return Object.freeze({
    schemaVersion: SPREADSHEET_WEB_ATLAS_COVERAGE_SCHEMA_VERSION,
    tabs: Object.freeze(tabs),
  });
}

export type SpreadsheetCapabilitySourceVersion =
  typeof SPREADSHEET_CAPABILITY_CATALOG_SCHEMA_VERSION | "unregistered";

export type SpreadsheetCapabilityControlKind =
  | "button"
  | "toggle"
  | "menu"
  | "gallery"
  | "combobox"
  | "field"
  | "dialog-launcher"
  | "internal"
  | "unknown";

export type SpreadsheetCapabilityCommandDisposition =
  "bound" | "parameterized" | "external" | "unsupported" | "unknown";

export type SpreadsheetCapabilityUndoPolicy =
  "single-transaction" | "content-none" | "external" | "unknown";

export type SpreadsheetCapabilityAuthoringDisposition =
  "native" | "projection" | "external" | "unsupported" | "unknown";

export type SpreadsheetCapabilityDomainDisposition =
  "typed" | "view-only" | "external" | "none" | "unknown";

export type SpreadsheetCapabilityPersistence =
  | "workbook"
  | "worksheet"
  | "object"
  | "session"
  | "external"
  | "none"
  | "unknown";

export type SpreadsheetCapabilityImportDisposition =
  | "native"
  | "preserved"
  | "collapsed"
  | "unsupported"
  | "not-applicable"
  | "unknown";

export type SpreadsheetCapabilityExportDisposition =
  "native" | "declared-loss" | "reject" | "not-applicable" | "unknown";

export type SpreadsheetCapabilityOfficeOutcome =
  "editable" | "visible" | "unsupported" | "unknown";

export type SpreadsheetCapabilityScalar = string | number | boolean | null;

export interface SpreadsheetCapabilityOfficeBaseline {
  readonly platform: string | null;
  readonly version: string | null;
  /** Microsoft/Open XML, Atlas evidence, or exact debt/fixture references. */
  readonly evidenceRefs: readonly string[];
}

export interface SpreadsheetCapabilitySurfaceContract {
  readonly disposition: "cataloged" | "unknown";
  readonly physicalPath: readonly string[];
  readonly group: string | null;
  readonly order: number | null;
  readonly controlKind: SpreadsheetCapabilityControlKind;
  /** Stable predicate key. `unknown` is inert, never an optimistic default. */
  readonly availabilityPredicate: string | "unknown";
}

export interface SpreadsheetCapabilityCommandContract {
  readonly disposition: SpreadsheetCapabilityCommandDisposition;
  readonly key: string | null;
  readonly handlerRef: string | null;
  readonly undoPolicy: SpreadsheetCapabilityUndoPolicy;
  readonly availabilityPredicate: string | "unknown";
}

export interface SpreadsheetCapabilitySemanticsContract {
  readonly defaults:
    Readonly<Record<string, SpreadsheetCapabilityScalar>> | "unknown";
  readonly limits:
    Readonly<Record<string, SpreadsheetCapabilityScalar>> | "unknown";
  readonly interactions: readonly string[];
  readonly shortcutIds: readonly string[];
}

export interface SpreadsheetCapabilityDomainContract {
  readonly disposition: SpreadsheetCapabilityDomainDisposition;
  readonly schemaPaths: readonly string[];
  readonly version: string | null;
  readonly persistence: SpreadsheetCapabilityPersistence;
}

export interface SpreadsheetCapabilityImportContract {
  readonly disposition: SpreadsheetCapabilityImportDisposition;
  readonly projectionRefs: readonly string[];
  readonly preservedUnknownParts: boolean | "unknown";
  readonly lossCodes: readonly string[];
  readonly flags: readonly string[];
}

export interface SpreadsheetCapabilityExportContract {
  readonly disposition: SpreadsheetCapabilityExportDisposition;
  readonly serializerRefs: readonly string[];
  readonly lossCodes: readonly string[];
  readonly flags: readonly string[];
}

export interface SpreadsheetCapabilityPatchbackContract extends SpreadsheetCapabilityExportContract {
  /** Exact owned-subtree/preserve rule; mandatory for a writable patchback. */
  readonly preservePolicy: string | null;
}

export interface SpreadsheetFeatureCapabilityContract {
  /** Shared command semantic. Physical aliases may intentionally reuse it. */
  readonly featureId: string;
  readonly canonicalId: string;
  /** Exact Control ID from the Spreadsheet Atlas. */
  readonly controlId: string;
  readonly sourceVersion: SpreadsheetCapabilitySourceVersion;
  readonly officeBaseline: SpreadsheetCapabilityOfficeBaseline;
  readonly surface: SpreadsheetCapabilitySurfaceContract;
  readonly command: SpreadsheetCapabilityCommandContract;
  readonly authoring: {
    readonly disposition: SpreadsheetCapabilityAuthoringDisposition;
  };
  readonly semantics: SpreadsheetCapabilitySemanticsContract;
  readonly domain: SpreadsheetCapabilityDomainContract;
  readonly import: SpreadsheetCapabilityImportContract;
  readonly fromScratchExport: SpreadsheetCapabilityExportContract;
  readonly patchbackExport: SpreadsheetCapabilityPatchbackContract;
  readonly officeOutcome: {
    readonly excel: SpreadsheetCapabilityOfficeOutcome;
    readonly wps: SpreadsheetCapabilityOfficeOutcome;
  };
  readonly acceptanceContractIds: readonly string[];
  readonly owner: string | null;
}

export interface SpreadsheetFeatureCapabilityCatalog {
  readonly schemaVersion: typeof SPREADSHEET_CAPABILITY_CATALOG_SCHEMA_VERSION;
  readonly sourceVersion: typeof SPREADSHEET_CAPABILITY_CATALOG_SCHEMA_VERSION;
  readonly features: readonly SpreadsheetFeatureCapabilityContract[];
  readonly controlIds: readonly string[];
  readonly featureIds: readonly string[];
  /** True only after every bound semantic has an exact live handler. */
  readonly handlerBindingsVerified: boolean;
}

export interface ReadSpreadsheetFeatureCapabilityCatalogOptions {
  /** Reserved for the later production command-registry adapter package. */
  readonly handlerRegistry?: never;
  readonly handlerBindingManifest?: never;
  readonly handlerFeatureIds?: never;
}

export interface VerifySpreadsheetFeatureCapabilityCatalogOptions {
  readonly expectedLineage: FeatureCommandRegistryLineage;
  readonly expectedRegistry: FeatureCommandBindingRegistry;
  readonly expectedBindingContract: FeatureCommandBindingContract;
}

const verifiedSpreadsheetCapabilityCatalogs =
  new WeakSet<SpreadsheetFeatureCapabilityCatalog>();

const SPREADSHEET_FEATURE_ID_PATTERN =
  /^spreadsheet(?:\.[A-Za-z][A-Za-z0-9_]*)+$/;
const SPREADSHEET_CANONICAL_ID_PATTERN = /^XLS(?:-[A-Z0-9]+)+-\d{2}$/;
const SPREADSHEET_CONTROL_ID_PATTERN = /^XLSCTL(?:\.[A-Z0-9_]+)+$/;
const RESERVED_CONTROL_SEGMENTS = new Set([
  "DEBT",
  "PROBE",
  "COV",
  "CLOSE",
  "CAP",
  "REG",
  "ALIAS",
]);

const CONTROL_KINDS = [
  "button",
  "toggle",
  "menu",
  "gallery",
  "combobox",
  "field",
  "dialog-launcher",
  "internal",
  "unknown",
] as const satisfies readonly SpreadsheetCapabilityControlKind[];
const COMMAND_DISPOSITIONS = [
  "bound",
  "parameterized",
  "external",
  "unsupported",
  "unknown",
] as const satisfies readonly SpreadsheetCapabilityCommandDisposition[];
const UNDO_POLICIES = [
  "single-transaction",
  "content-none",
  "external",
  "unknown",
] as const satisfies readonly SpreadsheetCapabilityUndoPolicy[];
const AUTHORING_DISPOSITIONS = [
  "native",
  "projection",
  "external",
  "unsupported",
  "unknown",
] as const satisfies readonly SpreadsheetCapabilityAuthoringDisposition[];
const DOMAIN_DISPOSITIONS = [
  "typed",
  "view-only",
  "external",
  "none",
  "unknown",
] as const satisfies readonly SpreadsheetCapabilityDomainDisposition[];
const PERSISTENCE_DISPOSITIONS = [
  "workbook",
  "worksheet",
  "object",
  "session",
  "external",
  "none",
  "unknown",
] as const satisfies readonly SpreadsheetCapabilityPersistence[];
const IMPORT_DISPOSITIONS = [
  "native",
  "preserved",
  "collapsed",
  "unsupported",
  "not-applicable",
  "unknown",
] as const satisfies readonly SpreadsheetCapabilityImportDisposition[];
const EXPORT_DISPOSITIONS = [
  "native",
  "declared-loss",
  "reject",
  "not-applicable",
  "unknown",
] as const satisfies readonly SpreadsheetCapabilityExportDisposition[];
const OFFICE_OUTCOMES = [
  "editable",
  "visible",
  "unsupported",
  "unknown",
] as const satisfies readonly SpreadsheetCapabilityOfficeOutcome[];

function isRecord(value: unknown): value is Record<string, unknown> {
  return value !== null && typeof value === "object" && !Array.isArray(value);
}

function requireRecord(value: unknown, label: string): Record<string, unknown> {
  if (!isRecord(value)) {
    throw new Error(`Missing spreadsheet capability ${label}`);
  }
  return value;
}

function requireString(value: unknown, label: string): string {
  if (
    typeof value !== "string" ||
    value.trim() === "" ||
    value.length > MAX_TEXT_CHARS
  ) {
    throw new Error(`Invalid spreadsheet capability ${label}`);
  }
  return value;
}

function requireSpreadsheetFeatureId(value: unknown, label: string): string {
  const id = requireString(value, label);
  if (id !== id.trim() || !SPREADSHEET_FEATURE_ID_PATTERN.test(id)) {
    throw new Error(`Invalid spreadsheet capability ${label}`);
  }
  return id;
}

function requireSpreadsheetCanonicalId(value: unknown, label: string): string {
  const id = requireString(value, label);
  if (id !== id.trim() || !SPREADSHEET_CANONICAL_ID_PATTERN.test(id)) {
    throw new Error(`Invalid spreadsheet capability ${label}`);
  }
  return id;
}

function requireSpreadsheetControlId(value: unknown, label: string): string {
  const id = requireString(value, label);
  const segments = id.split(".").slice(1);
  if (
    id !== id.trim() ||
    !SPREADSHEET_CONTROL_ID_PATTERN.test(id) ||
    segments.some((segment) => RESERVED_CONTROL_SEGMENTS.has(segment))
  ) {
    throw new Error(`Invalid spreadsheet capability ${label}`);
  }
  return id;
}

function requireNullableString(value: unknown, label: string): string | null {
  if (value === null) return null;
  return requireString(value, label);
}

function requireNullableOrder(value: unknown, label: string): number | null {
  if (value === null) return null;
  if (!Number.isSafeInteger(value) || (value as number) < 0) {
    throw new Error(`Invalid spreadsheet capability ${label}`);
  }
  return value as number;
}

function snapshotDenseArray(
  value: unknown,
  label: string,
  maxItems: number,
  oversizedMessage?: string,
): readonly unknown[] {
  if (!Array.isArray(value)) {
    throw new Error(`Invalid spreadsheet capability ${label}`);
  }
  const lengthDescriptor = Object.getOwnPropertyDescriptor(value, "length");
  const length = lengthDescriptor?.value;
  if (!Number.isSafeInteger(length) || length < 0) {
    throw new Error(`Invalid spreadsheet capability ${label}`);
  }
  if (length > maxItems) {
    throw new Error(
      oversizedMessage ?? `Invalid spreadsheet capability ${label}`,
    );
  }
  const snapshot: unknown[] = [];
  for (let index = 0; index < length; index += 1) {
    const descriptor = Object.getOwnPropertyDescriptor(value, String(index));
    if (!descriptor || !("value" in descriptor) || !descriptor.enumerable) {
      throw new Error(`Sparse spreadsheet capability ${label}[${index}]`);
    }
    snapshot.push(descriptor.value);
  }
  return Object.freeze(snapshot);
}

function requireStringList(value: unknown, label: string): readonly string[] {
  const values = snapshotDenseArray(value, label, MAX_LIST_ITEMS);
  const result: string[] = [];
  for (let index = 0; index < values.length; index += 1) {
    result.push(requireString(values[index], `${label}[${index}]`));
  }
  if (new Set(result).size !== result.length) {
    throw new Error(`Duplicate spreadsheet capability ${label}`);
  }
  return Object.freeze(result);
}

function requireDisposition<T extends readonly string[]>(
  value: unknown,
  allowed: T,
  label: string,
): T[number] {
  if (value === undefined) {
    throw new Error(`Missing spreadsheet capability ${label} disposition`);
  }
  if (typeof value !== "string" || !allowed.includes(value)) {
    throw new Error(
      `Undeclared spreadsheet capability ${label} disposition: ${String(value)}`,
    );
  }
  return value as T[number];
}

function requirePredicate(value: unknown, label: string): string | "unknown" {
  return value === "unknown" ? value : requireString(value, label);
}

function requireScalarMap(
  value: unknown,
  label: string,
): Readonly<Record<string, SpreadsheetCapabilityScalar>> | "unknown" {
  if (value === "unknown") return value;
  if (!isRecord(value) || Object.keys(value).length > MAX_LIST_ITEMS) {
    throw new Error(`Invalid spreadsheet capability ${label}`);
  }
  // A null-prototype record keeps hostile keys such as `__proto__` as data
  // instead of letting a catalog row mutate the result object's prototype.
  const result = Object.create(null) as Record<
    string,
    SpreadsheetCapabilityScalar
  >;
  const entries = Object.entries(value).sort(([left], [right]) =>
    left < right ? -1 : left > right ? 1 : 0,
  );
  for (const [key, candidate] of entries) {
    const normalizedKey = requireString(key, `${label}.key`);
    if (
      candidate !== null &&
      typeof candidate !== "string" &&
      typeof candidate !== "number" &&
      typeof candidate !== "boolean"
    ) {
      throw new Error(`Invalid spreadsheet capability ${label}.${key}`);
    }
    if (
      (typeof candidate === "number" && !Number.isFinite(candidate)) ||
      (typeof candidate === "string" && candidate.length > MAX_TEXT_CHARS)
    ) {
      throw new Error(`Invalid spreadsheet capability ${label}.${key}`);
    }
    result[normalizedKey] = candidate as SpreadsheetCapabilityScalar;
  }
  return Object.freeze(result);
}

function normalizeExportContract(
  value: unknown,
  label: string,
): SpreadsheetCapabilityExportContract {
  const source = requireRecord(value, label);
  const disposition = requireDisposition(
    source.disposition,
    EXPORT_DISPOSITIONS,
    label,
  );
  const serializerRefs = requireStringList(
    source.serializerRefs,
    `${label}.serializerRefs`,
  );
  const lossCodes = requireStringList(source.lossCodes, `${label}.lossCodes`);
  const flags = requireStringList(source.flags, `${label}.flags`);

  if (
    (disposition === "native" || disposition === "declared-loss") &&
    serializerRefs.length === 0
  ) {
    throw new Error(
      `Writable spreadsheet capability ${label} lacks serializer`,
    );
  }
  if (disposition === "native" && lossCodes.length > 0) {
    throw new Error(`Native spreadsheet capability ${label} declares loss`);
  }
  if (disposition === "declared-loss" && lossCodes.length === 0) {
    throw new Error(`Lossy spreadsheet capability ${label} lacks loss code`);
  }
  if (
    (disposition === "reject" ||
      disposition === "not-applicable" ||
      disposition === "unknown") &&
    serializerRefs.length > 0
  ) {
    throw new Error(
      `Non-writable spreadsheet capability ${label} has serializer`,
    );
  }

  return Object.freeze({ disposition, serializerRefs, lossCodes, flags });
}

function normalizeFeature(
  value: unknown,
  index: number,
): SpreadsheetFeatureCapabilityContract {
  const label = `features[${index}]`;
  const source = requireRecord(value, label);
  const officeBaseline = requireRecord(
    source.officeBaseline,
    `${label}.officeBaseline`,
  );
  const surface = requireRecord(source.surface, `${label}.surface`);
  const command = requireRecord(source.command, `${label}.command`);
  const authoring = requireRecord(source.authoring, `${label}.authoring`);
  const semantics = requireRecord(source.semantics, `${label}.semantics`);
  const domain = requireRecord(source.domain, `${label}.domain`);
  const importContract = requireRecord(source.import, `${label}.import`);
  const officeOutcome = requireRecord(
    source.officeOutcome,
    `${label}.officeOutcome`,
  );

  if (source.sourceVersion !== SPREADSHEET_CAPABILITY_CATALOG_SCHEMA_VERSION) {
    throw new Error(`Invalid spreadsheet capability ${label}.sourceVersion`);
  }

  const normalizedSurface: SpreadsheetCapabilitySurfaceContract = Object.freeze(
    {
      disposition: requireDisposition(
        surface.disposition,
        ["cataloged", "unknown"] as const,
        `${label}.surface`,
      ),
      physicalPath: requireStringList(
        surface.physicalPath,
        `${label}.surface.physicalPath`,
      ),
      group: requireNullableString(surface.group, `${label}.surface.group`),
      order: requireNullableOrder(surface.order, `${label}.surface.order`),
      controlKind: requireDisposition(
        surface.controlKind,
        CONTROL_KINDS,
        `${label}.surface.controlKind`,
      ),
      availabilityPredicate: requirePredicate(
        surface.availabilityPredicate,
        `${label}.surface.availabilityPredicate`,
      ),
    },
  );
  if (
    normalizedSurface.disposition === "cataloged" &&
    (normalizedSurface.physicalPath.length === 0 ||
      normalizedSurface.group === null ||
      normalizedSurface.order === null ||
      normalizedSurface.controlKind === "unknown" ||
      normalizedSurface.availabilityPredicate === "unknown")
  ) {
    throw new Error(`Incomplete spreadsheet capability ${label}.surface`);
  }

  const normalizedCommand: SpreadsheetCapabilityCommandContract = Object.freeze(
    {
      disposition: requireDisposition(
        command.disposition,
        COMMAND_DISPOSITIONS,
        `${label}.command`,
      ),
      key: requireNullableString(command.key, `${label}.command.key`),
      handlerRef: requireNullableString(
        command.handlerRef,
        `${label}.command.handlerRef`,
      ),
      undoPolicy: requireDisposition(
        command.undoPolicy,
        UNDO_POLICIES,
        `${label}.command.undoPolicy`,
      ),
      availabilityPredicate: requirePredicate(
        command.availabilityPredicate,
        `${label}.command.availabilityPredicate`,
      ),
    },
  );
  const commandIsBound =
    normalizedCommand.disposition === "bound" ||
    normalizedCommand.disposition === "parameterized";
  if (
    commandIsBound &&
    (normalizedCommand.key === null ||
      normalizedCommand.handlerRef === null ||
      normalizedCommand.undoPolicy === "unknown" ||
      normalizedCommand.availabilityPredicate === "unknown")
  ) {
    throw new Error(`Incomplete spreadsheet capability ${label}.command`);
  }
  if (
    !commandIsBound &&
    normalizedCommand.disposition !== "external" &&
    normalizedCommand.handlerRef !== null
  ) {
    throw new Error(
      `Inert spreadsheet capability ${label}.command has handler`,
    );
  }
  if (
    commandIsBound &&
    normalizedSurface.disposition === "cataloged" &&
    normalizedSurface.availabilityPredicate !==
      normalizedCommand.availabilityPredicate
  ) {
    throw new Error(
      `Spreadsheet capability ${label} availability predicate drift`,
    );
  }

  const normalizedAuthoring = Object.freeze({
    disposition: requireDisposition(
      authoring.disposition,
      AUTHORING_DISPOSITIONS,
      `${label}.authoring`,
    ),
  });
  if (
    commandIsBound &&
    normalizedAuthoring.disposition !== "native" &&
    normalizedAuthoring.disposition !== "projection"
  ) {
    throw new Error(
      `Non-authorable spreadsheet capability ${label} has bound command`,
    );
  }

  const normalizedDomain: SpreadsheetCapabilityDomainContract = Object.freeze({
    disposition: requireDisposition(
      domain.disposition,
      DOMAIN_DISPOSITIONS,
      `${label}.domain`,
    ),
    schemaPaths: requireStringList(
      domain.schemaPaths,
      `${label}.domain.schemaPaths`,
    ),
    version: requireNullableString(domain.version, `${label}.domain.version`),
    persistence: requireDisposition(
      domain.persistence,
      PERSISTENCE_DISPOSITIONS,
      `${label}.domain.persistence`,
    ),
  });
  if (
    normalizedDomain.disposition === "typed" &&
    (normalizedDomain.schemaPaths.length === 0 ||
      normalizedDomain.version === null ||
      normalizedDomain.persistence === "unknown")
  ) {
    throw new Error(`Incomplete spreadsheet capability ${label}.domain`);
  }
  if (
    normalizedDomain.disposition === "none" &&
    (normalizedDomain.schemaPaths.length > 0 ||
      normalizedDomain.persistence !== "none")
  ) {
    throw new Error(
      `Invalid spreadsheet capability ${label}.domain none contract`,
    );
  }

  const importDisposition = requireDisposition(
    importContract.disposition,
    IMPORT_DISPOSITIONS,
    `${label}.import`,
  );
  const normalizedImport: SpreadsheetCapabilityImportContract = Object.freeze({
    disposition: importDisposition,
    projectionRefs: requireStringList(
      importContract.projectionRefs,
      `${label}.import.projectionRefs`,
    ),
    preservedUnknownParts:
      importContract.preservedUnknownParts === "unknown"
        ? "unknown"
        : Boolean(importContract.preservedUnknownParts),
    lossCodes: requireStringList(
      importContract.lossCodes,
      `${label}.import.lossCodes`,
    ),
    flags: requireStringList(importContract.flags, `${label}.import.flags`),
  });
  if (
    importContract.preservedUnknownParts !== true &&
    importContract.preservedUnknownParts !== false &&
    importContract.preservedUnknownParts !== "unknown"
  ) {
    throw new Error(
      `Invalid spreadsheet capability ${label}.import.preservedUnknownParts`,
    );
  }
  if (
    normalizedImport.disposition === "native" &&
    normalizedImport.projectionRefs.length === 0
  ) {
    throw new Error(
      `Native spreadsheet capability ${label}.import lacks projection`,
    );
  }
  if (
    normalizedImport.preservedUnknownParts === false &&
    normalizedImport.lossCodes.length === 0
  ) {
    throw new Error(
      `Lossy spreadsheet capability ${label}.import lacks loss code`,
    );
  }

  const normalizedFromScratch = normalizeExportContract(
    source.fromScratchExport,
    `${label}.fromScratchExport`,
  );
  const patchbackSource = requireRecord(
    source.patchbackExport,
    `${label}.patchbackExport`,
  );
  const normalizedPatchbackBase = normalizeExportContract(
    patchbackSource,
    `${label}.patchbackExport`,
  );
  const normalizedPatchback: SpreadsheetCapabilityPatchbackContract =
    Object.freeze({
      ...normalizedPatchbackBase,
      preservePolicy: requireNullableString(
        patchbackSource.preservePolicy,
        `${label}.patchbackExport.preservePolicy`,
      ),
    });
  const patchbackWritable =
    normalizedPatchback.disposition === "native" ||
    normalizedPatchback.disposition === "declared-loss";
  if (patchbackWritable && normalizedPatchback.preservePolicy === null) {
    throw new Error(
      `Writable spreadsheet capability ${label}.patchbackExport lacks preserve policy`,
    );
  }
  if (!patchbackWritable && normalizedPatchback.preservePolicy !== null) {
    throw new Error(
      `Non-writable spreadsheet capability ${label}.patchbackExport has preserve policy`,
    );
  }

  const normalized = Object.freeze({
    featureId: requireSpreadsheetFeatureId(
      source.featureId,
      `${label}.featureId`,
    ),
    canonicalId: requireSpreadsheetCanonicalId(
      source.canonicalId,
      `${label}.canonicalId`,
    ),
    controlId: requireSpreadsheetControlId(
      source.controlId,
      `${label}.controlId`,
    ),
    sourceVersion: SPREADSHEET_CAPABILITY_CATALOG_SCHEMA_VERSION,
    officeBaseline: Object.freeze({
      platform: requireNullableString(
        officeBaseline.platform,
        `${label}.officeBaseline.platform`,
      ),
      version: requireNullableString(
        officeBaseline.version,
        `${label}.officeBaseline.version`,
      ),
      evidenceRefs: requireStringList(
        officeBaseline.evidenceRefs,
        `${label}.officeBaseline.evidenceRefs`,
      ),
    }),
    surface: normalizedSurface,
    command: normalizedCommand,
    authoring: normalizedAuthoring,
    semantics: Object.freeze({
      defaults: requireScalarMap(
        semantics.defaults,
        `${label}.semantics.defaults`,
      ),
      limits: requireScalarMap(semantics.limits, `${label}.semantics.limits`),
      interactions: requireStringList(
        semantics.interactions,
        `${label}.semantics.interactions`,
      ),
      shortcutIds: requireStringList(
        semantics.shortcutIds,
        `${label}.semantics.shortcutIds`,
      ),
    }),
    domain: normalizedDomain,
    import: normalizedImport,
    fromScratchExport: normalizedFromScratch,
    patchbackExport: normalizedPatchback,
    officeOutcome: Object.freeze({
      excel: requireDisposition(
        officeOutcome.excel,
        OFFICE_OUTCOMES,
        `${label}.officeOutcome.excel`,
      ),
      wps: requireDisposition(
        officeOutcome.wps,
        OFFICE_OUTCOMES,
        `${label}.officeOutcome.wps`,
      ),
    }),
    acceptanceContractIds: requireStringList(
      source.acceptanceContractIds,
      `${label}.acceptanceContractIds`,
    ),
    owner: requireNullableString(source.owner, `${label}.owner`),
  }) satisfies SpreadsheetFeatureCapabilityContract;

  if (
    normalized.officeBaseline.evidenceRefs.length === 0 ||
    normalized.acceptanceContractIds.length === 0 ||
    normalized.owner === null
  ) {
    throw new Error(`Unowned spreadsheet capability ${label}`);
  }
  return normalized;
}

export function spreadsheetFeatureCommandSemanticContract(
  capability: SpreadsheetFeatureCapabilityContract,
): string {
  return JSON.stringify({
    canonicalId: capability.canonicalId,
    officeBaseline: capability.officeBaseline,
    surface: {
      disposition: capability.surface.disposition,
      availabilityPredicate: capability.surface.availabilityPredicate,
    },
    command: capability.command,
    authoring: capability.authoring,
    semantics: capability.semantics,
    domain: capability.domain,
    import: capability.import,
    fromScratchExport: capability.fromScratchExport,
    patchbackExport: capability.patchbackExport,
    officeOutcome: capability.officeOutcome,
    acceptanceContractIds: capability.acceptanceContractIds,
    owner: capability.owner,
  });
}

/**
 * Parse a capability catalog into immutable, validated contracts.
 *
 * No production Spreadsheet command-registry adapter is wired in this
 * architecture package. Detached handler IDs, manifests, and shared test
 * registries are rejected; every command and writer-readiness helper stays
 * fail-closed until a separately leased production adapter supplies proof.
 */
export function readSpreadsheetFeatureCapabilityCatalog(
  input: unknown,
  options: ReadSpreadsheetFeatureCapabilityCatalogOptions = {},
): SpreadsheetFeatureCapabilityCatalog {
  const source = requireRecord(input, "catalog");
  const untrustedOptions = options as {
    readonly handlerRegistry?: unknown;
    readonly handlerBindingManifest?: unknown;
    readonly handlerFeatureIds?: unknown;
  };
  const handlerRegistry = untrustedOptions.handlerRegistry;
  const handlerBindingManifest = untrustedOptions.handlerBindingManifest;
  const handlerFeatureIds = untrustedOptions.handlerFeatureIds;
  if (
    handlerRegistry !== undefined ||
    handlerBindingManifest !== undefined ||
    handlerFeatureIds !== undefined
  ) {
    throw new Error(
      `Spreadsheet production command-registry adapter is not installed`,
    );
  }
  if (source.version !== SPREADSHEET_CAPABILITY_CATALOG_SCHEMA_VERSION) {
    throw new Error(`Unsupported spreadsheet capability catalog version`);
  }
  if (!Array.isArray(source.features)) {
    throw new Error(`Missing spreadsheet capability catalog features`);
  }
  const sourceFeatures = snapshotDenseArray(
    source.features,
    "features",
    MAX_SPREADSHEET_CAPABILITY_ROWS,
    `Spreadsheet capability catalog exceeds ${MAX_SPREADSHEET_CAPABILITY_ROWS} rows`,
  );

  const features: SpreadsheetFeatureCapabilityContract[] = [];
  for (let index = 0; index < sourceFeatures.length; index += 1) {
    features.push(normalizeFeature(sourceFeatures[index], index));
  }
  const byControl = new Map<string, SpreadsheetFeatureCapabilityContract>();
  const signatureByFeature = new Map<string, string>();
  const featureByCommandKey = new Map<string, string>();
  for (const capability of features) {
    if (byControl.has(capability.controlId)) {
      throw new Error(
        `Duplicate spreadsheet capability Control ID: ${capability.controlId}`,
      );
    }
    byControl.set(capability.controlId, capability);

    const signature = spreadsheetFeatureCommandSemanticContract(capability);
    const existingSignature = signatureByFeature.get(capability.featureId);
    if (existingSignature !== undefined && existingSignature !== signature) {
      throw new Error(
        `Spreadsheet capability alias drift: ${capability.featureId}`,
      );
    }
    signatureByFeature.set(capability.featureId, signature);

    if (
      (capability.command.disposition === "bound" ||
        capability.command.disposition === "parameterized") &&
      capability.command.key !== null
    ) {
      const existingFeature = featureByCommandKey.get(capability.command.key);
      if (
        existingFeature !== undefined &&
        existingFeature !== capability.featureId
      ) {
        throw new Error(
          `Spreadsheet capability command semantic split: ${capability.command.key}`,
        );
      }
      featureByCommandKey.set(capability.command.key, capability.featureId);
    }
  }

  const catalog = Object.freeze({
    schemaVersion: SPREADSHEET_CAPABILITY_CATALOG_SCHEMA_VERSION,
    sourceVersion: SPREADSHEET_CAPABILITY_CATALOG_SCHEMA_VERSION,
    features: Object.freeze(features),
    controlIds: Object.freeze(features.map((item) => item.controlId)),
    featureIds: Object.freeze([
      ...new Set(features.map((item) => item.featureId)),
    ]),
    handlerBindingsVerified: false,
  });
  verifiedSpreadsheetCapabilityCatalogs.add(catalog);
  return catalog;
}

/**
 * Promote a parsed catalog only after the shared Office command authority has
 * authenticated every executable semantic in that catalog. Unsupported and
 * external-only rows intentionally have no binding; they remain visible audit
 * records and can never become executable through this proof.
 */
export function verifySpreadsheetFeatureCapabilityCatalogBindings(
  catalog: SpreadsheetFeatureCapabilityCatalog,
  manifest: FeatureCommandBindingManifest,
  options: VerifySpreadsheetFeatureCapabilityCatalogOptions,
): SpreadsheetFeatureCapabilityCatalog {
  if (
    !verifiedSpreadsheetCapabilityCatalogs.has(catalog) ||
    catalog.sourceVersion !== SPREADSHEET_CAPABILITY_CATALOG_SCHEMA_VERSION ||
    catalog.handlerBindingsVerified
  ) {
    throw new Error("Untrusted spreadsheet capability catalog");
  }

  const trusted = readFeatureCommandBindingManifest(manifest, options);
  const expectedFeatureIds = [
    ...new Set(
      catalog.features
        .filter(
          (capability) =>
            capability.command.disposition === "bound" ||
            capability.command.disposition === "parameterized",
        )
        .map((capability) => capability.featureId),
    ),
  ].sort();
  if (
    trusted.editor !== "spreadsheet" ||
    trusted.boundFeatureIds.length !== expectedFeatureIds.length ||
    trusted.boundFeatureIds.some(
      (featureId, index) => featureId !== expectedFeatureIds[index],
    )
  ) {
    throw new Error("Spreadsheet capability binding identity mismatch");
  }

  const verified = Object.freeze({
    ...catalog,
    handlerBindingsVerified: true,
  });
  verifiedSpreadsheetCapabilityCatalogs.add(verified);
  return verified;
}

function unregisteredCapability(
  controlId: string,
): SpreadsheetFeatureCapabilityContract {
  const empty = Object.freeze([]) as readonly string[];
  return Object.freeze({
    featureId: controlId,
    canonicalId: "unregistered",
    controlId,
    sourceVersion: "unregistered",
    officeBaseline: Object.freeze({
      platform: null,
      version: null,
      evidenceRefs: empty,
    }),
    surface: Object.freeze({
      disposition: "unknown",
      physicalPath: empty,
      group: null,
      order: null,
      controlKind: "unknown",
      availabilityPredicate: "unknown",
    }),
    command: Object.freeze({
      disposition: "unknown",
      key: null,
      handlerRef: null,
      undoPolicy: "unknown",
      availabilityPredicate: "unknown",
    }),
    authoring: Object.freeze({ disposition: "unknown" }),
    semantics: Object.freeze({
      defaults: "unknown",
      limits: "unknown",
      interactions: empty,
      shortcutIds: empty,
    }),
    domain: Object.freeze({
      disposition: "unknown",
      schemaPaths: empty,
      version: null,
      persistence: "unknown",
    }),
    import: Object.freeze({
      disposition: "unknown",
      projectionRefs: empty,
      preservedUnknownParts: "unknown",
      lossCodes: empty,
      flags: empty,
    }),
    fromScratchExport: Object.freeze({
      disposition: "unknown",
      serializerRefs: empty,
      lossCodes: empty,
      flags: empty,
    }),
    patchbackExport: Object.freeze({
      disposition: "unknown",
      serializerRefs: empty,
      preservePolicy: null,
      lossCodes: empty,
      flags: empty,
    }),
    officeOutcome: Object.freeze({ excel: "unknown", wps: "unknown" }),
    acceptanceContractIds: empty,
    owner: null,
  });
}

/** A verified catalog is frozen, so its rows are indexed once: the ribbon
 * resolves hundreds of controls per render, and a linear search per control
 * over hundreds of rows was a visible share of every commit. The first row
 * per control id wins, exactly as the previous linear search did. */
const catalogIndexes = new WeakMap<
  SpreadsheetFeatureCapabilityCatalog,
  { byControlId: Map<string, SpreadsheetFeatureCapabilityContract>; owned: Set<SpreadsheetFeatureCapabilityContract> }
>();
function indexOfCatalog(catalog: SpreadsheetFeatureCapabilityCatalog) {
  let index = catalogIndexes.get(catalog);
  if (!index) {
    const byControlId = new Map<string, SpreadsheetFeatureCapabilityContract>();
    for (const item of catalog.features) {
      if (!byControlId.has(item.controlId)) byControlId.set(item.controlId, item);
    }
    index = { byControlId, owned: new Set(catalog.features) };
    catalogIndexes.set(catalog, index);
  }
  return index;
}

/** Missing Controls resolve to an inert object instead of a plausible default. */
export function resolveSpreadsheetFeatureCapability(
  controlId: string,
  catalog: SpreadsheetFeatureCapabilityCatalog,
): SpreadsheetFeatureCapabilityContract {
  if (!verifiedSpreadsheetCapabilityCatalogs.has(catalog)) {
    return unregisteredCapability(controlId);
  }
  return indexOfCatalog(catalog).byControlId.get(controlId) ?? unregisteredCapability(controlId);
}

function isCatalogOwnedSpreadsheetCapability(
  capability: SpreadsheetFeatureCapabilityContract,
  catalog: SpreadsheetFeatureCapabilityCatalog,
): boolean {
  return (
    verifiedSpreadsheetCapabilityCatalogs.has(catalog) &&
    indexOfCatalog(catalog).owned.has(capability)
  );
}

/** True only for a catalog-owned row with an independently verified handler. */
export function isSpreadsheetFeatureExecutable(
  capability: SpreadsheetFeatureCapabilityContract,
  catalog: SpreadsheetFeatureCapabilityCatalog,
): boolean {
  if (
    !catalog.handlerBindingsVerified ||
    !isCatalogOwnedSpreadsheetCapability(capability, catalog)
  ) {
    return false;
  }
  if (
    capability.sourceVersion !==
      SPREADSHEET_CAPABILITY_CATALOG_SCHEMA_VERSION ||
    capability.owner === null ||
    capability.acceptanceContractIds.length === 0
  ) {
    return false;
  }
  if (capability.surface.disposition !== "cataloged") return false;
  if (
    capability.command.disposition !== "bound" &&
    capability.command.disposition !== "parameterized"
  ) {
    return false;
  }
  if (
    capability.command.key === null ||
    capability.command.handlerRef === null ||
    capability.command.availabilityPredicate === "unknown"
  ) {
    return false;
  }
  if (
    capability.authoring.disposition !== "native" &&
    capability.authoring.disposition !== "projection"
  ) {
    return false;
  }
  return (
    capability.domain.disposition === "typed" ||
    capability.domain.disposition === "view-only" ||
    capability.domain.disposition === "none"
  );
}

export type SpreadsheetNativeWriteRoute = "from-scratch" | "patchback";

/**
 * Where a workbook page's save is written: into its imported original
 * (`patchback`) when the page carries either the import's source marker or a
 * save's patchback marker, else a workbook built from the page.
 */
export function spreadsheetPageWriteRoute(
  page:
    | {
        editable_page_patchback?: unknown;
        editable_page_source?: { adapter?: string | null; resource_id?: string | null } | null;
      }
    | null
    | undefined,
): SpreadsheetNativeWriteRoute {
  if (page?.editable_page_patchback) return "patchback";
  return page?.editable_page_source?.adapter === "xlsx" ? "patchback" : "from-scratch";
}

/**
 * Narrow writer readiness. `declared-loss` remains usable only through a
 * caller that explicitly handles its loss codes; it is never native-ready.
 */
export function isSpreadsheetFeatureNativeWriteReady(
  capability: SpreadsheetFeatureCapabilityContract,
  catalog: SpreadsheetFeatureCapabilityCatalog,
  route: SpreadsheetNativeWriteRoute,
): boolean {
  if (
    !catalog.handlerBindingsVerified ||
    !isCatalogOwnedSpreadsheetCapability(capability, catalog) ||
    capability.sourceVersion !== SPREADSHEET_CAPABILITY_CATALOG_SCHEMA_VERSION
  ) {
    return false;
  }
  const contract =
    route === "from-scratch"
      ? capability.fromScratchExport
      : capability.patchbackExport;
  if (
    contract.disposition !== "native" ||
    contract.serializerRefs.length === 0 ||
    contract.lossCodes.length > 0
  ) {
    return false;
  }
  return (
    route !== "patchback" || capability.patchbackExport.preservePolicy !== null
  );
}

/** Both independent writers plus editable Excel/WPS reopen evidence. */
export function isSpreadsheetFeatureRoundtripReady(
  capability: SpreadsheetFeatureCapabilityContract,
  catalog: SpreadsheetFeatureCapabilityCatalog,
): boolean {
  return (
    isSpreadsheetFeatureExecutable(capability, catalog) &&
    isCatalogOwnedSpreadsheetCapability(capability, catalog) &&
    capability.sourceVersion ===
      SPREADSHEET_CAPABILITY_CATALOG_SCHEMA_VERSION &&
    capability.owner !== null &&
    capability.officeBaseline.evidenceRefs.length > 0 &&
    capability.import.disposition === "native" &&
    capability.import.preservedUnknownParts === true &&
    capability.import.lossCodes.length === 0 &&
    isSpreadsheetFeatureNativeWriteReady(capability, catalog, "from-scratch") &&
    isSpreadsheetFeatureNativeWriteReady(capability, catalog, "patchback") &&
    capability.officeOutcome.excel === "editable" &&
    capability.officeOutcome.wps === "editable" &&
    capability.acceptanceContractIds.length > 0
  );
}
