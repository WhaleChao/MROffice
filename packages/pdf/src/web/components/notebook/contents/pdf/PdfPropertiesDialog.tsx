"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import {
  useRef,
  useState,
  type KeyboardEvent,
  type RefObject,
} from "react";
import { useOpenedOnce } from "@mrnote/core/web/lib/use-opened-once.ts";
import { Loader2 } from "lucide-react";
import {
  Accessibility,
  ChevronRight,
  Download,
  FileText,
  Info,
  ShieldCheck,
  X,
  type HouseIcon,
} from "@mrnote/core/web/components/office-ribbon/icons/houseIcons.tsx";
import { useTranslations } from "@mrnote/core/host/next-intl.tsx";

import {
  Dialog,
  DialogContent,
  DialogDescription,
  DialogFooter,
  DialogHeader,
  DialogTitle,
} from "@mrnote/core/web/components/ui/dialog.tsx";
import {
  officeDialogClassName,
  officeDialogParts,
  useOfficeDialogChrome,
} from "@mrnote/core/web/components/office-ribbon/officeDialogChrome.tsx";

import {
  metadataChanged,
  normalizePdfLanguageTag,
  normalizePdfMetadata,
  PDF_LANGUAGE_MAX_LENGTH,
  type PdfMetadata,
} from "../../../../lib/pdf/document-properties";
import type { PdfTaggedTextTarget } from "@/lib/pdf/native-tagged-text-roundtrip";

export interface PdfTaggedTextProperties {
  readonly sourceSha256: string;
  readonly targets: readonly PdfTaggedTextTarget[];
  readonly busy: boolean;
  readonly disabledReason?: string;
  readonly onApply: (
    target: PdfTaggedTextTarget,
    property: "Alt" | "ActualText",
    value: string,
  ) => Promise<boolean>;
}

function TaggedTextProperties({ model }: { model: PdfTaggedTextProperties }) {
  const translate = useTranslations("console-notebooks");
  const [selected, setSelected] = useState("");
  const [property, setProperty] = useState<"Alt" | "ActualText">("Alt");
  const [drafts, setDrafts] = useState<
    Record<
      string,
      {
        source: string;
        target: PdfTaggedTextTarget;
        value: string;
        explicit: boolean;
      }
    >
  >({});
  const [failed, setFailed] = useState(false);
  const target =
    model.targets.find(
      (t) => `${t.node.objectNumber}:${t.node.generation}` === selected,
    ) ?? model.targets[0];
  const key = target
    ? `${target.node.objectNumber}:${target.node.generation}:${property}`
    : "";
  const nodeValue = target
    ? property === "Alt"
      ? target.alt
      : target.actualText
    : null;
  const inlineValue = target
    ? property === "Alt"
      ? target.inlineAlt
      : target.inlineActualText
    : null;
  const differs = Boolean(
    nodeValue && inlineValue && nodeValue.text !== inlineValue.text,
  );
  const draft = drafts[key];
  const conflict = Boolean(draft && draft.source !== model.sourceSha256);
  const value =
    draft?.value ??
    (differs ? "" : (nodeValue?.text ?? inlineValue?.text ?? ""));
  const choose = (value: string) => {
    if (!target || model.busy) return;
    setFailed(false);
    setDrafts((current) => ({
      ...current,
      [key]: { source: model.sourceSha256, target, value, explicit: true },
    }));
  };
  const reset = () => {
    setDrafts((current) => {
      const next = { ...current };
      delete next[key];
      return next;
    });
    setFailed(false);
  };
  if (!target)
    return (
      <p role="status" className="mt-3 text-sm">
        {model.disabledReason ?? translate("pdf.properties.semanticTextEmpty")}
      </p>
    );
  return (
    <section
      className={`${officeDialogParts.group} mt-5`}
      data-testid="pdf-tagged-text-properties"
    >
      <strong className="text-sm">
        {translate("pdf.properties.semanticTextHeading")}
      </strong>
      <label className={officeDialogParts.field}>
        {translate("pdf.properties.semanticTextContent")}
        <select
          value={`${target.node.objectNumber}:${target.node.generation}`}
          disabled={model.busy}
          onChange={(e) => {
            setSelected(e.target.value);
            setFailed(false);
          }}
          data-testid="pdf-tagged-target"
        >
          {model.targets.map((t) => (
            <option
              key={`${t.node.objectNumber}:${t.node.generation}`}
              value={`${t.node.objectNumber}:${t.node.generation}`}
            >
              {translate("pdf.properties.semanticTextPage", {
                pageNumber: t.pageNumber,
              })}{" "}
              ·{" "}
              {t.role === "/Figure"
                ? translate("pdf.properties.semanticTextFigure")
                : translate("pdf.properties.semanticTextText")}{" "}
              ·{" "}
              {(t.alt?.text ?? t.actualText?.text ?? `${t.mcid + 1}`).slice(
                0,
                70,
              )}
            </option>
          ))}
        </select>
      </label>
      <label className={officeDialogParts.field}>
        {translate("pdf.properties.semanticTextProperty")}
        <select
          value={property}
          disabled={model.busy}
          onChange={(e) => {
            setProperty(e.target.value as "Alt" | "ActualText");
            setFailed(false);
          }}
          data-testid="pdf-tagged-property"
        >
          <option value="Alt">
            {translate("pdf.properties.semanticTextAlt")}
          </option>
          <option value="ActualText">
            {translate("pdf.properties.semanticTextActual")}
          </option>
        </select>
      </label>
      {differs && (
        <div
          className="grid gap-2 text-sm"
          data-testid="pdf-tagged-value-discrepancy"
        >
          <p>{translate("pdf.properties.semanticTextConflict")}</p>
          <button
            type="button"
            disabled={model.busy || Boolean(model.disabledReason) || conflict}
            onClick={() => choose(nodeValue!.text)}
            data-testid="pdf-tagged-use-structure"
          >
            {translate("pdf.properties.semanticTextUseStructure")}:{" "}
            {nodeValue?.text}
          </button>
          <button
            type="button"
            disabled={model.busy || Boolean(model.disabledReason) || conflict}
            onClick={() => choose(inlineValue!.text)}
            data-testid="pdf-tagged-use-page"
          >
            {translate("pdf.properties.semanticTextUsePage")}:{" "}
            {inlineValue?.text}
          </button>
        </div>
      )}
      <textarea
        value={value}
        disabled={model.busy || Boolean(model.disabledReason) || conflict}
        rows={4}
        maxLength={4000}
        onChange={(e) => choose(e.target.value)}
        aria-label={translate("pdf.properties.semanticTextLabel")}
        data-testid="pdf-tagged-value"
        className="w-full"
      />
      {(model.disabledReason || conflict || failed) && (
        <p role="alert" className={officeDialogParts.error}>
          {model.disabledReason ??
            (conflict
              ? translate("pdf.properties.semanticTextSourceChanged")
              : translate("pdf.properties.semanticTextNotSaved"))}
        </p>
      )}
      <div className="flex gap-2">
        <button
          type="button"
          disabled={
            model.busy ||
            Boolean(model.disabledReason) ||
            conflict ||
            !draft?.explicit
          }
          data-testid="pdf-tagged-apply"
          onClick={async () => {
            if (!draft || conflict) return;
            try {
              if (await model.onApply(draft.target, property, draft.value))
                reset();
              else setFailed(true);
            } catch {
              setFailed(true);
            }
          }}
        >
          {model.busy
            ? translate("pdf.properties.semanticTextSaving")
            : translate("pdf.properties.semanticTextApply")}
        </button>
        <button
          type="button"
          disabled={model.busy}
          onClick={reset}
          data-testid="pdf-tagged-reset"
        >
          {conflict
            ? translate("pdf.properties.semanticTextReload")
            : translate("pdf.properties.semanticTextCancel")}
        </button>
      </div>
    </section>
  );
}

export interface PdfPropertiesDialogProps {
  open: boolean;
  /**
   * The document's CURRENT metadata (read from the loaded PDF) used to seed the
   * inputs. The dialog re-seeds whenever this changes while closed→open.
   */
  metadata: PdfMetadata;
  /** Whether an export/flatten is currently in flight (disables Save). */
  busy?: boolean;
  /**
   * Apply the edited metadata: the host writes it into a flattened export and
   * triggers a download. Receives the NORMALIZED draft.
   */
  onSave: (metadata: PdfMetadata) => void;
  /** Close without applying. */
  onClose: () => void;
  /** Exact command trigger to restore after Save, Cancel, Escape, or scrim. */
  returnFocusRef?: RefObject<HTMLElement | null>;
  /** Optional themed portal container. */
  container?: HTMLElement | null;
  /** Source-bound read-only facts; omission keeps legacy metadata-only use. */
  facts?: PdfDocumentPropertiesFacts;
  /** Close the modal and reveal one page-bound accessibility finding. */
  onNavigateToAccessibilityIssue?: (issue: PdfAccessibilityIssue) => void;
  taggedText?:PdfTaggedTextProperties;
}

export type PdfDocumentFactState = "none" | "present" | "unknown";

export type PdfAccessibilityIssueKind =
  | "untagged-page"
  | "incomplete-reading-order"
  | "missing-alt-text";

export interface PdfAccessibilityIssue {
  readonly kind: PdfAccessibilityIssueKind;
  readonly page: number;
  readonly count?: number;
}

export interface PdfDocumentPropertiesFacts {
  formatVersion: string | null;
  pageCount: number;
  pageSizes: ReadonlyArray<{ widthPt: number; heightPt: number }>;
  formFieldCount: number;
  managedAttachmentCount: number;
  optionalContentGroupCount: number;
  encryption: PdfDocumentFactState;
  certificateSignatures: PdfDocumentFactState;
  attachments: PdfDocumentFactState;
  xfa: PdfDocumentFactState;
  taggedStructure: PdfDocumentFactState;
  pdfaClaim: PdfDocumentFactState;
  accessibilityAudit?: Readonly<{
    status: "not-checked" | "checking" | "ready" | "failed";
    pageCount: number;
    untaggedPageCount: number;
    incompleteReadingOrderPageCount: number;
    semanticObjectCount: number;
    missingAltTextCount: number;
    issues: readonly PdfAccessibilityIssue[];
  }>;
}

interface FieldDef {
  key: keyof PdfMetadata;
  labelKey: string;
  placeholderKey: string;
  /** Render as a textarea (multi-line affordance) rather than a single input. */
  multiline?: boolean;
}

type PdfPropertiesSection =
  | "description"
  | "document"
  | "security"
  | "accessibility";

const FIELDS: ReadonlyArray<FieldDef> = [
  {
    key: "title",
    labelKey: "pdf.properties.title",
    placeholderKey: "pdf.properties.titlePlaceholder",
  },
  {
    key: "author",
    labelKey: "pdf.properties.author",
    placeholderKey: "pdf.properties.authorPlaceholder",
  },
  {
    key: "subject",
    labelKey: "pdf.properties.subject",
    placeholderKey: "pdf.properties.subjectPlaceholder",
    multiline: true,
  },
  {
    key: "keywords",
    labelKey: "pdf.properties.keywords",
    placeholderKey: "pdf.properties.keywordsPlaceholder",
  },
  {
    key: "language",
    labelKey: "pdf.properties.language",
    placeholderKey: "pdf.properties.languagePlaceholder",
  },
];

/**
 * Document Properties editor (Acrobat / WPS "File ▸ Properties"): view + edit
 * the standard info fields plus the catalog-level document language, then
 * Save, which the host bakes into a flattened download via pdf-lib.
 * The pure model (normalize / sanitize / change-detection) lives in
 * lib/pdf/document-properties.ts; this is the controlled form + modal chrome.
 */
export default function PdfPropertiesDialog(props: PdfPropertiesDialogProps) {
  if (!useOpenedOnce(props.open)) return null;
  return <PdfPropertiesDialogBody {...props} />;
}

function PdfPropertiesDialogBody({
  open,
  metadata,
  busy: metadataBusy = false,
  onSave,
  onClose,
  returnFocusRef,
  container,
  facts,
  onNavigateToAccessibilityIssue,
  taggedText,
}: PdfPropertiesDialogProps) {
  const busy = metadataBusy || Boolean(taggedText?.busy);
  const t = useTranslations("console-notebooks");
  // Inside the reader this card wears the shared Office dialog chrome, as
  // every other PDF dialog does through OfficeDialogShell.
  const chromeTone = useOfficeDialogChrome();
  const [draft, setDraft] = useState<PdfMetadata>(() =>
    normalizePdfMetadata(metadata),
  );
  const [activeSection, setActiveSection] =
    useState<PdfPropertiesSection>("description");
  const firstFieldRef = useRef<HTMLInputElement | null>(null);
  const capturedReturnFocusRef = useRef<HTMLElement | null>(null);

  // Re-seed the form from the document's current metadata each time the dialog
  // opens, so reopening after a download shows the freshly written values (the
  // host re-reads metadata) rather than a stale draft. Done as a render-phase
  // prev-state comparison (React's "adjust state when a prop changes" pattern)
  // instead of an effect, so it never schedules a cascading post-commit render.
  const [wasOpen, setWasOpen] = useState(open);
  if (open !== wasOpen) {
    setWasOpen(open);
    if (open) {
      setDraft(normalizePdfMetadata(metadata));
      setActiveSection("description");
    }
  }

  const normalized = normalizePdfMetadata(draft);
  const changed = metadataChanged(metadata, normalized);
  const languageValid = normalizePdfLanguageTag(draft.language) !== null;

  const update = (key: keyof PdfMetadata, value: string) =>
    setDraft((prev) => ({ ...prev, [key]: value }));

  const requestSave = () => {
    if (busy || !changed || !languageValid) return;
    onSave(normalized);
  };

  const stateLabel = (state: PdfDocumentFactState) =>
    state === "present"
      ? t("pdf.properties.factPresent")
      : state === "none"
        ? t("pdf.properties.factNone")
        : t("pdf.properties.factUnknown");
  const pageSizeSummary = (() => {
    if (!facts || facts.pageSizes.length === 0)
      return t("pdf.properties.factUnknown");
    const unique = new Map<string, string>();
    for (const size of facts.pageSizes) {
      if (
        !Number.isFinite(size.widthPt) ||
        size.widthPt <= 0 ||
        !Number.isFinite(size.heightPt) ||
        size.heightPt <= 0
      ) {
        continue;
      }
      const widthMm = (size.widthPt * 25.4) / 72;
      const heightMm = (size.heightPt * 25.4) / 72;
      const label = `${widthMm.toFixed(1)} × ${heightMm.toFixed(1)} mm`;
      unique.set(`${widthMm.toFixed(2)}:${heightMm.toFixed(2)}`, label);
    }
    if (unique.size === 0) return t("pdf.properties.factUnknown");
    if (unique.size === 1) return [...unique.values()][0];
    return t("pdf.properties.pageSizeCount", { count: unique.size });
  })();
  const documentFactRows: ReadonlyArray<readonly [string, string]> = facts
    ? [
        [
          t("pdf.properties.factVersion"),
          facts.formatVersion
            ? `PDF ${facts.formatVersion}`
            : stateLabel("unknown"),
        ],
        [t("pdf.properties.factPages"), String(facts.pageCount)],
        [t("pdf.properties.factPageSize"), pageSizeSummary],
        [t("pdf.properties.factFormFields"), String(facts.formFieldCount)],
        [
          t("pdf.properties.factAttachments"),
          facts.managedAttachmentCount > 0
            ? t("pdf.properties.factAttachmentsManaged", {
                state: stateLabel(facts.attachments),
                count: facts.managedAttachmentCount,
              })
            : stateLabel(facts.attachments),
        ],
        [
          t("pdf.properties.factTaggedStructure"),
          stateLabel(facts.taggedStructure),
        ],
        [
          t("pdf.properties.factLayers"),
          String(facts.optionalContentGroupCount),
        ],
      ]
    : [];
  const securityFactRows: ReadonlyArray<readonly [string, string]> = facts
    ? [
        [t("pdf.properties.factEncryption"), stateLabel(facts.encryption)],
        [
          t("pdf.properties.factCertificateSignatures"),
          stateLabel(facts.certificateSignatures),
        ],
        [
          t("pdf.properties.factEmbeddedAttachments"),
          stateLabel(facts.attachments),
        ],
        ["XFA", stateLabel(facts.xfa)],
        [t("pdf.properties.factPdfaClaim"), stateLabel(facts.pdfaClaim)],
      ]
    : [];
  const accessibilityAudit = facts?.accessibilityAudit;
  const accessibilityCountsReady = accessibilityAudit?.status === "ready";
  const accessibilityFactRows: ReadonlyArray<readonly [string, string]> =
    accessibilityAudit
      ? [
          [
            t("pdf.properties.basicCheck"),
            accessibilityAudit.status === "checking"
              ? t("pdf.properties.basicCheckRunning")
              : accessibilityAudit.status === "failed"
                ? t("pdf.properties.basicCheckFailed")
                : accessibilityAudit.status === "not-checked"
                  ? t("pdf.properties.basicCheckNotRun")
                  : accessibilityAudit.untaggedPageCount > 0 ||
                      accessibilityAudit.incompleteReadingOrderPageCount > 0 ||
                      accessibilityAudit.missingAltTextCount > 0
                    ? t("pdf.properties.basicCheckIssues")
                    : t("pdf.properties.basicCheckPassed"),
          ],
          [
            t("pdf.properties.factUntaggedPages"),
            accessibilityCountsReady
              ? `${accessibilityAudit.untaggedPageCount} / ${accessibilityAudit.pageCount}`
              : "—",
          ],
          [
            t("pdf.properties.factIncompleteReadingOrder"),
            accessibilityCountsReady
              ? String(accessibilityAudit.incompleteReadingOrderPageCount)
              : "—",
          ],
          [
            t("pdf.properties.factMissingAltText"),
            accessibilityCountsReady
              ? `${accessibilityAudit.missingAltTextCount} / ${accessibilityAudit.semanticObjectCount}`
              : "—",
          ],
        ]
      : [
          [
            t("pdf.properties.basicCheck"),
            t("pdf.properties.basicCheckNotRun"),
          ],
        ];
  const sections: ReadonlyArray<{
    id: PdfPropertiesSection;
    label: string;
    icon: HouseIcon;
  }> = facts
    ? [
        {
          id: "description",
          label: t("pdf.properties.sectionDescription"),
          icon: FileText,
        },
        {
          id: "document",
          label: t("pdf.properties.sectionDocument"),
          icon: Info,
        },
        {
          id: "security",
          label: t("pdf.properties.sectionSecurity"),
          icon: ShieldCheck,
        },
        {
          id: "accessibility",
          label: t("pdf.properties.sectionAccessibility"),
          icon: Accessibility,
        },
      ]
    : [
        {
          id: "description",
          label: t("pdf.properties.sectionDescription"),
          icon: FileText,
        },
      ];
  const selectedSection = sections.some(
    (section) => section.id === activeSection,
  )
    ? activeSection
    : "description";
  const activeFactRows =
    selectedSection === "document"
      ? documentFactRows
      : selectedSection === "security"
        ? securityFactRows
        : accessibilityFactRows;
  const accessibilityIssueLabel = (issue: PdfAccessibilityIssue) =>
    issue.kind === "untagged-page"
      ? t("pdf.properties.issueUntaggedPage")
      : issue.kind === "incomplete-reading-order"
        ? t("pdf.properties.issueReadingOrder")
        : issue.count
          ? t("pdf.properties.issueMissingAltTextCount", { count: issue.count })
          : t("pdf.properties.issueMissingAltText");

  const selectAdjacentSection = (
    event: KeyboardEvent<HTMLButtonElement>,
    index: number,
  ) => {
    if (
      event.key !== "ArrowLeft" &&
      event.key !== "ArrowRight" &&
      event.key !== "Home" &&
      event.key !== "End"
    ) {
      return;
    }
    event.preventDefault();
    const target =
      event.key === "Home"
        ? 0
        : event.key === "End"
          ? sections.length - 1
          : (index + (event.key === "ArrowRight" ? 1 : -1) + sections.length) %
            sections.length;
    const next = sections[target];
    if (!next) return;
    setActiveSection(next.id);
    event.currentTarget.parentElement
      ?.querySelectorAll<HTMLButtonElement>('[role="tab"]')
      .item(target)
      ?.focus();
  };

  return (
    <Dialog
      open={open}
      onOpenChange={(nextOpen) => {
        if (!nextOpen && !busy) onClose();
      }}
    >
      <DialogContent
        container={container}
        showCloseButton={false}
        skin={chromeTone ? "none" : undefined}
        className={`${
          chromeTone ? `${officeDialogClassName({ tone: chromeTone })} ` : ""
        }max-h-[calc(100vh-64px)] overflow-y-auto sm:max-w-[720px]`}
        aria-busy={busy}
        data-testid="pdf-properties-dialog"
        onEscapeKeyDown={(event) => {
          if (busy) event.preventDefault();
        }}
        onInteractOutside={(event) => {
          if (busy) event.preventDefault();
        }}
        onOpenAutoFocus={(event) => {
          event.preventDefault();
          const activeElement = document.activeElement;
          capturedReturnFocusRef.current =
            returnFocusRef?.current ??
            (activeElement instanceof HTMLElement ? activeElement : null);
          firstFieldRef.current?.focus();
        }}
        onCloseAutoFocus={(event) => {
          const returnTarget = capturedReturnFocusRef.current;
          capturedReturnFocusRef.current = null;
          if (returnTarget?.isConnected) {
            event.preventDefault();
            returnTarget.focus();
          }
        }}
      >
        <form
          onSubmit={(event) => {
            event.preventDefault();
            requestSave();
          }}
        >
          <DialogHeader>
            <div className={officeDialogParts.header}>
              <DialogTitle className={officeDialogParts.title}>
                {t("pdf.properties.heading")}
              </DialogTitle>
              <button
                type="button"
                className={officeDialogParts.close}
                onClick={onClose}
                disabled={busy}
                aria-label={t("common.close")}
                data-testid="pdf-properties-dialog-close"
              >
                <X size={14} aria-hidden="true" />
              </button>
            </div>
            <DialogDescription
              skin={chromeTone ? "none" : undefined}
              className={`${officeDialogParts.description} pdf-properties-dialog__hint`}
            >
              {t("pdf.properties.hint")}
            </DialogDescription>
          </DialogHeader>

          <div
            className={`${officeDialogParts.tabs} mt-4`}
            role="tablist"
            aria-label={t("pdf.properties.categories")}
            data-testid="pdf-properties-tabs"
          >
            {sections.map((section, index) => {
              const Icon = section.icon;
              return (
                <button
                  key={section.id}
                  id={`pdf-properties-tab-${section.id}`}
                  type="button"
                  role="tab"
                  aria-selected={selectedSection === section.id}
                  aria-controls={`pdf-properties-panel-${section.id}`}
                  tabIndex={selectedSection === section.id ? 0 : -1}
                  className={officeDialogParts.tab}
                  onClick={() => setActiveSection(section.id)}
                  onKeyDown={(event) => selectAdjacentSection(event, index)}
                  data-testid={`pdf-properties-tab-${section.id}`}
                >
                  <Icon size={15} aria-hidden />
                  {section.label}
                </button>
              );
            })}
          </div>

          <section
            id={`pdf-properties-panel-${selectedSection}`}
            role="tabpanel"
            aria-labelledby={`pdf-properties-tab-${selectedSection}`}
            className="mt-4 min-h-[310px]"
            data-testid={`pdf-properties-panel-${selectedSection}`}
          >
            {selectedSection === "description" ? (
              <>
                <div className="pdf-properties-dialog__fields">
                  {FIELDS.map((field) => {
                    const id = `pdf-properties-${field.key}`;
                    return (
                      <label
                        key={field.key}
                        htmlFor={id}
                        data-properties-field={field.key}
                      >
                        <span>{t(field.labelKey)}</span>
                        {field.multiline ? (
                          <textarea
                            id={id}
                            data-testid={id}
                            rows={2}
                            value={draft[field.key]}
                            placeholder={t(field.placeholderKey)}
                            onChange={(event) =>
                              update(field.key, event.target.value)
                            }
                          />
                        ) : (
                          <input
                            ref={
                              field.key === "title" ? firstFieldRef : undefined
                            }
                            id={id}
                            data-testid={id}
                            type="text"
                            value={draft[field.key]}
                            placeholder={t(field.placeholderKey)}
                            maxLength={
                              field.key === "language"
                                ? PDF_LANGUAGE_MAX_LENGTH
                                : undefined
                            }
                            aria-invalid={
                              field.key === "language" && !languageValid
                                ? true
                                : undefined
                            }
                            aria-describedby={
                              field.key === "language"
                                ? "pdf-properties-language-feedback"
                                : undefined
                            }
                            onChange={(event) =>
                              update(field.key, event.target.value)
                            }
                          />
                        )}
                      </label>
                    );
                  })}
                </div>
                <p className="pdf-properties-dialog__keywords-hint mt-2">
                  {t("pdf.properties.keywordsHint")}
                </p>
                <p
                  id="pdf-properties-language-feedback"
                  className={`mt-1 ${
                    languageValid
                      ? officeDialogParts.hint
                      : officeDialogParts.error
                  }`}
                  aria-live="polite"
                >
                  {t(
                    languageValid
                      ? "pdf.properties.languageHint"
                      : "pdf.properties.languageInvalid",
                  )}
                </p>
              </>
            ) : (
              <div
                className={officeDialogParts.group}
                aria-label={t("pdf.properties.details")}
                data-testid="pdf-properties-facts"
                data-properties-section={selectedSection}
              >
                <strong className="text-sm">
                  {sections.find((section) => section.id === selectedSection)
                    ?.label ?? t("pdf.properties.details")}
                </strong>
                <dl className="mt-3 grid grid-cols-[minmax(0,1fr)_minmax(0,1.35fr)] gap-x-4 gap-y-2 text-sm">
                  {activeFactRows.map(([label, value]) => (
                    <div
                      className="contents"
                      key={label}
                      data-testid="pdf-properties-fact-row"
                    >
                      <dt className={officeDialogParts.hint}>{label}</dt>
                      <dd className="m-0 break-words text-right font-semibold">
                        {value}
                      </dd>
                    </div>
                  ))}
                </dl>
                {selectedSection === "security" ? (
                  <p className={`${officeDialogParts.message} mt-4 leading-relaxed`}>
                    {t("pdf.properties.factsHint")}
                  </p>
                ) : selectedSection === "accessibility" ? (
                  <>
                    {taggedText && <TaggedTextProperties model={taggedText} />}
                    <p
                      className={`${officeDialogParts.hint} mt-4 leading-relaxed`}
                    >
                      {t("pdf.properties.accessibilityHint")}
                    </p>
                    {accessibilityAudit?.status === "ready" &&
                    accessibilityAudit.issues.length > 0 ? (
                      <div
                        className={`${officeDialogParts.group} mt-5`}
                        data-testid="pdf-properties-accessibility-issues"
                      >
                        <strong className="text-xs">
                          {t("pdf.properties.inspectionResults")}
                        </strong>
                        <ul className="mt-2 grid gap-1.5 p-0">
                          {accessibilityAudit.issues.map((issue) => (
                            <li
                              key={`${issue.kind}:${issue.page}`}
                              className="list-none"
                            >
                              <button
                                type="button"
                                className={`${officeDialogParts.compactButton} flex w-full cursor-pointer items-center gap-2 text-left`}
                                disabled={!onNavigateToAccessibilityIssue}
                                onClick={() => {
                                  if (!onNavigateToAccessibilityIssue) return;
                                  onClose();
                                  onNavigateToAccessibilityIssue(issue);
                                }}
                                data-testid="pdf-properties-accessibility-issue"
                              >
                                <span className="min-w-0 flex-1">
                                  <span className="block font-semibold">
                                    {t("pdf.properties.issuePage", {
                                      page: issue.page,
                                    })}
                                  </span>
                                  <span className={`${officeDialogParts.hint} block`}>
                                    {accessibilityIssueLabel(issue)}
                                  </span>
                                </span>
                                <ChevronRight size={15} aria-hidden />
                              </button>
                            </li>
                          ))}
                        </ul>
                      </div>
                    ) : null}
                  </>
                ) : null}
              </div>
            )}
          </section>

          <DialogFooter
            className={`${officeDialogParts.footer} pdf-properties-dialog__actions mt-6`}
          >
            <button
              type="button"
              className={`${officeDialogParts.button} pdf-properties-dialog__cancel`}
              data-testid="pdf-properties-close"
              onClick={onClose}
              disabled={busy}
            >
              {t("common.cancel")}
            </button>
            <button
              type="submit"
              className={`${officeDialogParts.primaryButton} pdf-properties-dialog__save`}
              data-testid="pdf-properties-save"
              disabled={busy || !changed || !languageValid}
              title={
                !languageValid
                  ? t("pdf.properties.languageInvalid")
                  : changed
                    ? t("pdf.properties.save")
                    : t("pdf.properties.noChanges")
              }
            >
              {busy ? (
                <Loader2 size={14} className="animate-spin" />
              ) : (
                <Download size={14} />
              )}
              <span>
                {busy
                  ? t("pdf.properties.saving")
                  : t("pdf.properties.save")}
              </span>
            </button>
          </DialogFooter>
        </form>
      </DialogContent>
    </Dialog>
  );
}
