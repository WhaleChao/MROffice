"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The contextual Equation tab every MRAI editor shows while an equation is
 * being edited or selected (Word's and PowerPoint's Equation tab): Input
 * Format (Unicode / LaTeX, and Docs' Text), Layout (Professional / Linear),
 * Placement (Change to Inline / Display, where the host has it), Symbols (the
 * basic palette plus the full gallery) and Structures (the eleven families).
 *
 * One component, driven through an ADAPTER (`OfficeEquationRibbonAdapter`):
 * the host says what the selected equation is (its input format, layout,
 * placement) and runs the commands on its own model -- Docs on the selected
 * ProseMirror node and its live MathLive NodeView, the deck on the open
 * in-place draft or the editor it opens. Every control is part of the live
 * draft (`VISUAL_MATH_EDITOR_FOCUS_OWNER_ATTRIBUTE`): pressing one never
 * commits it, and a closed menu hands the keyboard back to it
 * (`refocusDraft`). A host's own groups go first (`leading`).
 */

import { useEffect, useLayoutEffect, useMemo, useRef, useState, type ReactNode } from "react";

import {
  EQUATION_SYMBOL_GROUPS,
  EQUATION_STRUCTURES,
  equationStructureMembersForFamily,
  equationStructurePreviewLatex,
  type EquationInputMode,
  type EquationLayoutMode,
  type EquationPaletteSymbol,
  type EquationStructureId,
} from "../console/editor/extensions/equationCatalog";
import type { EquationPlacement } from "../console/editor/extensions/equationPlacement";
import { renderValidatedKatexMarkup } from "../console/editor/extensions/lazyKatex";
import { VISUAL_MATH_EDITOR_FOCUS_OWNER_ATTRIBUTE } from "../console/editor/extensions/visualMathEditor";
import {
  equationRibbonCopy,
  type EquationRibbonCopy,
} from "../console/editor/native-document/equationRibbonCopy";
import { DocumentEquationSymbolGallery } from "../console/editor/native-document/DocumentEquationSymbolGallery";
import styles from "../console/editor/native-document/DocumentEquationContextRibbon.module.css";
import { OfficeRibbonButton, OfficeRibbonGroup, RibbonHStack } from "../office-ribbon/OfficeRibbonChassis";
import { OfficeHomeMenuTrigger } from "../office-ribbon/OfficeHomeRibbon";
import { OfficeRibbonGallery } from "../office-ribbon/OfficeRibbonGallery";
import { RibbonDropdownMenuContent } from "../office-ribbon/CupertinoRibbonMenu";
import { OfficeSemanticIcon } from "../office-ribbon/icons/OfficeSemanticIcon";
import { DropdownMenu, DropdownMenuItem, DropdownMenuTrigger } from "../ui/dropdown-menu";

const visualMathFocusOwnerProps = {
  [VISUAL_MATH_EDITOR_FOCUS_OWNER_ATTRIBUTE]: "true",
};
const EQUATION_RIBBON_SYMBOLS =
  EQUATION_SYMBOL_GROUPS.find((group) => group.id === "basicMath")?.symbols.slice(0, 40) ?? [];

const NAMED_EQUATION_SYMBOL_IDS: ReadonlySet<string> = new Set([
  "pm",
  "infty",
  "basic-equals",
  "neq",
  "basic-sim",
  "times",
  "div",
  "basic-factorial",
  "basic-propto",
  "basic-less-than",
  "basic-much-less",
  "basic-greater-than",
  "basic-much-greater",
  "leq",
  "geq",
  "basic-mp",
  "basic-simeq",
  "approx",
  "basic-equiv",
  "basic-forall",
  "basic-complement",
  "partial",
  "basic-square-root",
  "basic-cube-root",
  "basic-fourth-root",
  "basic-union",
  "basic-intersection",
  "basic-empty-set",
  "basic-percent",
  "basic-degree",
  "basic-fahrenheit",
  "basic-celsius",
  "basic-capital-delta",
  "nabla",
  "basic-exists",
  "basic-not-exists",
  "basic-element-of",
  "basic-contains",
  "basic-left-arrow",
  "rightarrow",
]);

export function equationSymbolLabel(
  symbol: EquationPaletteSymbol,
  groupLabel: string,
  t: (key: string, values?: Record<string, string>) => string,
): string {
  return NAMED_EQUATION_SYMBOL_IDS.has(symbol.id)
    ? t(`equation.ribbon.symbol.${symbol.id}`)
    : t("equation.ribbon.symbolUnnamed", {
        name: symbol.label,
        group: groupLabel,
        glyph: symbol.glyph,
      });
}

/**
 * A Structures family face: Word's typeset specimen drawn into a fixed
 * 40×28 slot. The specimens differ wildly in natural size (x/y versus a
 * summation with limits or a 2×2 matrix), so the rendered markup is measured
 * and scaled DOWN to the slot — never up, never clipped, never over the label.
 * Re-measured once KaTeX's fonts arrive, which changes the glyph metrics.
 */
export function EquationStructureFamilyFace({
  latex,
  markup,
  fallback,
}: {
  readonly latex: string;
  readonly markup: string | null;
  readonly fallback: string;
}) {
  const faceRef = useRef<HTMLSpanElement>(null);
  const contentRef = useRef<HTMLSpanElement>(null);
  useLayoutEffect(() => {
    const face = faceRef.current;
    const content = contentRef.current;
    if (!face || !content) return;
    let frame = 0;
    const fit = () => {
      content.style.transform = "";
      const slot = face.getBoundingClientRect();
      const natural = content.getBoundingClientRect();
      if (!slot.width || !slot.height || !natural.width || !natural.height) {
        return;
      }
      // Ninety per cent of the slot: a specimen filling it edge to edge read
      
      const scale = Math.min(1, (0.9 * slot.width) / natural.width, (0.9 * slot.height) / natural.height);
      content.style.transform = scale < 1 ? `scale(${scale.toFixed(4)})` : "";
    };
    fit();
    frame = window.requestAnimationFrame(fit);
    const fonts = (document as Document & { fonts?: FontFaceSet }).fonts;
    let cancelled = false;
    fonts?.ready.then(() => {
      if (!cancelled) fit();
    });
    const observer = typeof ResizeObserver === "function" ? new ResizeObserver(fit) : null;
    observer?.observe(face);
    return () => {
      cancelled = true;
      window.cancelAnimationFrame(frame);
      observer?.disconnect();
    };
  }, [markup]);
  return (
    <span
      ref={faceRef}
      className={styles.familyGlyph}
      aria-hidden
      aria-busy={!markup || undefined}
      data-equation-preview={latex}
    >
      {markup ? (
        <span
          ref={contentRef}
          className={`ML__mathlive visual-math-editor__katex-preview ${styles.familyRenderedMath}`}
          data-latex={latex}
          dangerouslySetInnerHTML={{ __html: markup }}
        />
      ) : (
        <span ref={contentRef} className={styles.familyRenderedMath}>
          {fallback}
        </span>
      )}
    </span>
  );
}

/** An Equation-tab command, in the host-neutral words of the shared tab. */
export type OfficeEquationRibbonCommand =
  | { type: "inputMode"; value: EquationInputMode }
  | { type: "layoutMode"; value: EquationLayoutMode }
  | { type: "placement"; value: EquationPlacement }
  | { type: "symbol"; id: string }
  | { type: "structure"; id: EquationStructureId }
  | { type: "structureMember"; id: string };

/** Why a host refused a command (the tab says so in its status line). */
export type OfficeEquationRibbonRefusal =
  | "stale_selection"
  | "readonly"
  | "unknown_structure_member"
  | "unknown_symbol"
  | "unsupported_unicode"
  | "unsupported_text"
  | "unsupported_placement";

/**
 * A command's answer: done, refused (and why), or nothing yet (`void`: the
 * host is opening the editor first and runs the command once it is live).
 */
export type OfficeEquationRibbonResult = { readonly ok: true } | { readonly ok: false; readonly reason: OfficeEquationRibbonRefusal } | void;

export interface OfficeEquationRibbonAdapter {
  /** Whether the equation can be changed now. */
  readonly editable: boolean;
  /** The selected equation's linear input format, and the formats the host offers. */
  readonly inputMode: EquationInputMode;
  readonly inputModes: readonly EquationInputMode[];
  readonly layoutMode: EquationLayoutMode;
  /** Change to Inline / Display, for a host whose equations have both; absent hides the group. */
  readonly placement?: EquationPlacement;
  run(command: OfficeEquationRibbonCommand): OfficeEquationRibbonResult;
  /** Put the keyboard back into the live draft (a menu closed after a pick). */
  refocusDraft(): void;
}

function resultMessage(result: OfficeEquationRibbonResult, copy: EquationRibbonCopy): string {
  if (!result || result.ok) return "";
  if (result.reason === "unsupported_unicode") return copy.unsupportedUnicode;
  if (result.reason === "unsupported_text") return copy.unsupportedText;
  if (result.reason === "unsupported_placement") return copy.unsupportedPlacement;
  if (result.reason === "readonly") return copy.readonly;
  if (result.reason === "unknown_structure_member" || result.reason === "unknown_symbol") {
    return copy.unavailableStructure;
  }
  return "";
}

export interface OfficeEquationContextRibbonProps {
  readonly adapter: OfficeEquationRibbonAdapter;
  readonly t: (key: string, values?: Record<string, string>) => string;
  /** Every test id starts with it (`document-equation`, `presentation-equation`). */
  readonly testIdPrefix: string;
  /** The full Symbols gallery's control id (screen tips). */
  readonly symbolsMoreControlId: string;
  /** The host's own groups, before Input Format. */
  readonly leading?: ReactNode;
}

export function OfficeEquationContextRibbon({
  adapter,
  t,
  testIdPrefix,
  symbolsMoreControlId,
  leading,
}: OfficeEquationContextRibbonProps) {
  const copy = useMemo(() => equationRibbonCopy(t as (key: string) => string), [t]);
  const editable = adapter.editable;
  const inputAxisHint = t("equation.ribbon.inputAxisHint");
  const layoutAxisHint = t("equation.ribbon.layoutAxisHint");
  const placementAxisHint = t("equation.ribbon.placementAxisHint");
  const moreSymbolsLabel = t("equation.ribbon.allSymbols");
  const basicMathGroupLabel = t("equation.ribbon.symbolGroup.basicMath");
  // Word's full palette, mapped onto the shared gallery anatomy once, so this
  // panel and the Insert menu's compact palette cannot drift apart.
  const fullSymbolGalleryGroups = useMemo(
    () =>
      EQUATION_SYMBOL_GROUPS.map((group) => {
        const groupLabel = t(`equation.ribbon.symbolGroup.${group.id}`);
        return {
          id: group.id,
          label: groupLabel,
          cells: group.symbols.map((symbol) => ({
            id: symbol.id,
            glyph: symbol.glyph,
            label: equationSymbolLabel(symbol, groupLabel, t),
          })),
        };
      }),
    [t],
  );
  const [status, setStatus] = useState("");
  const [openStructureFamily, setOpenStructureFamily] = useState<EquationStructureId | null>(null);
  const [symbolsOpen, setSymbolsOpen] = useState(false);
  const restoreStructureTriggerOnCloseRef = useRef<EquationStructureId | null>(null);
  const refocusStructureDraftOnCloseRef = useRef(false);
  const restoreSymbolsTriggerOnCloseRef = useRef(false);
  const refocusSymbolsDraftOnCloseRef = useRef(false);
  // MathLive refocus can unmount the Radix portal during pointerdown. Let that
  // primary pointer own the command, then consume any later select for the
  // same activation; keyboard select remains the independent command path.
  const structurePointerActivationSerialRef = useRef(0);
  const structurePointerActivationRef = useRef<{
    readonly memberId: string;
    readonly token: number;
  } | null>(null);
  const adapterRef = useRef(adapter);
  useLayoutEffect(() => {
    adapterRef.current = adapter;
  });

  useEffect(() => {
    if (editable) return;
    structurePointerActivationRef.current = null;
    const timeout = window.setTimeout(() => {
      setSymbolsOpen(false);
      setOpenStructureFamily(null);
    }, 0);
    return () => window.clearTimeout(timeout);
  }, [editable]);

  useEffect(() => {
    if (openStructureFamily === null && refocusStructureDraftOnCloseRef.current) {
      refocusStructureDraftOnCloseRef.current = false;
      adapterRef.current.refocusDraft();
    }
    if (!symbolsOpen && refocusSymbolsDraftOnCloseRef.current) {
      refocusSymbolsDraftOnCloseRef.current = false;
      adapterRef.current.refocusDraft();
    }
  }, [openStructureFamily, symbolsOpen]);

  const run = (command: OfficeEquationRibbonCommand): OfficeEquationRibbonResult => {
    const result = adapter.run(command);
    setStatus(resultMessage(result, copy));
    return result;
  };
  const succeeded = (result: OfficeEquationRibbonResult) => !result || result.ok;
  const inputModeLabels: Record<EquationInputMode, string> = {
    unicode: copy.unicode,
    latex: copy.latex,
    text: copy.text,
  };

  return (
    <>
      {leading}
      <OfficeRibbonGroup
        ribbon="equation"
        label={copy.inputFormat}
        className={`${styles.ribbonGroup} ${styles.inputFormatGroup}`}
        testId={`${testIdPrefix}-ribbon-input-format`}
      >
        <RibbonHStack className={`${styles.segmentedControl} ${styles.inputModeControl}`}>
          {adapter.inputModes.map((mode) => (
            <OfficeRibbonButton
              {...visualMathFocusOwnerProps}
              key={mode}
              active={adapter.inputMode === mode}
              disabled={!editable}
              className={styles.segmentButton}
              label={inputModeLabels[mode]}
              onClick={() => run({ type: "inputMode", value: mode })}
              testId={`${testIdPrefix}-input-${mode}`}
              data-equation-segment={`input-${mode}`}
            >
              <span className={styles.segmentLabel}>{inputModeLabels[mode]}</span>
            </OfficeRibbonButton>
          ))}
        </RibbonHStack>
        <span className={styles.axisHint}>{inputAxisHint}</span>
      </OfficeRibbonGroup>

      <OfficeRibbonGroup
        ribbon="equation"
        label={copy.layout}
        className={`${styles.ribbonGroup} ${styles.layoutGroup}`}
        testId={`${testIdPrefix}-ribbon-layout`}
      >
        <RibbonHStack className={`${styles.segmentedControl} ${styles.layoutModeControl}`}>
          {(
            [
              ["professional", copy.professional],
              ["linear", copy.linear],
            ] as const
          ).map(([mode, label]) => (
            <OfficeRibbonButton
              {...visualMathFocusOwnerProps}
              key={mode}
              active={adapter.layoutMode === mode}
              disabled={!editable}
              className={styles.segmentButton}
              label={label}
              onClick={() => run({ type: "layoutMode", value: mode })}
              testId={`${testIdPrefix}-layout-${mode}`}
              data-equation-segment={`layout-${mode}`}
            >
              <span className={styles.segmentLabel}>{label}</span>
            </OfficeRibbonButton>
          ))}
        </RibbonHStack>
        <span className={styles.axisHint}>{layoutAxisHint}</span>
      </OfficeRibbonGroup>

      {adapter.placement ? (
        <OfficeRibbonGroup
          ribbon="equation"
          label={copy.placement}
          className={`${styles.ribbonGroup} ${styles.placementGroup}`}
          testId={`${testIdPrefix}-ribbon-placement`}
        >
          <RibbonHStack className={`${styles.segmentedControl} ${styles.placementControl}`}>
            {(
              [
                ["inline", copy.changeToInline],
                ["display", copy.changeToDisplay],
              ] as const
            ).map(([placement, label]) => (
              <OfficeRibbonButton
                {...visualMathFocusOwnerProps}
                key={placement}
                active={adapter.placement === placement}
                disabled={!editable}
                className={styles.segmentButton}
                label={label}
                onClick={() => run({ type: "placement", value: placement })}
                testId={`${testIdPrefix}-placement-${placement}`}
              >
                <span className={styles.segmentLabel}>{label}</span>
              </OfficeRibbonButton>
            ))}
          </RibbonHStack>
          <span className={styles.axisHint}>{placementAxisHint}</span>
        </OfficeRibbonGroup>
      ) : null}

      <OfficeRibbonGroup
        ribbon="equation"
        label={copy.symbols}
        className={`${styles.ribbonGroup} ${styles.symbolGroup}`}
        testId={`${testIdPrefix}-ribbon-symbols`}
      >
        <OfficeRibbonGallery
          ariaLabel={copy.symbols}
          className={styles.symbolPalette}
          testId={`${testIdPrefix}-symbols-inline-gallery`}
          viewport={
            <div className={styles.compactSymbolViewport}>
              <div className={styles.compactSymbolGrid} role="toolbar" aria-label={copy.symbols}>
                {EQUATION_RIBBON_SYMBOLS.map((symbol) => (
                  <OfficeRibbonButton
                    {...visualMathFocusOwnerProps}
                    key={symbol.id}
                    disabled={!editable}
                    className={styles.symbolButton}
                    label={equationSymbolLabel(symbol, basicMathGroupLabel, t)}
                    onClick={() => run({ type: "symbol", id: symbol.id })}
                    testId={`${testIdPrefix}-symbol-${symbol.id}`}
                  >
                    <span className={styles.symbolGlyph} aria-hidden>
                      {symbol.glyph}
                    </span>
                  </OfficeRibbonButton>
                ))}
              </div>
            </div>
          }
          more={
            <DropdownMenu
              modal={false}
              open={editable && symbolsOpen}
              onOpenChange={(nextOpen) => {
                if (nextOpen && !editable) return;
                if (nextOpen) {
                  restoreSymbolsTriggerOnCloseRef.current = false;
                  refocusSymbolsDraftOnCloseRef.current = false;
                }
                setSymbolsOpen(nextOpen);
              }}
            >
              <DropdownMenuTrigger asChild>
                <OfficeHomeMenuTrigger
                  {...visualMathFocusOwnerProps}
                  geometry="gallery"
                  controlId={symbolsMoreControlId}
                  disabled={!editable}
                  icon={<OfficeSemanticIcon semanticId="menu.chevron.down" size={12} />}
                  label={moreSymbolsLabel}
                  onMouseDown={(event) => event.preventDefault()}
                  testId={`${testIdPrefix}-symbols-all`}
                  title={moreSymbolsLabel}
                  variant="icon"
                />
              </DropdownMenuTrigger>
              <RibbonDropdownMenuContent
                {...visualMathFocusOwnerProps}
                align="start"
                sideOffset={5}
                aria-label={moreSymbolsLabel}
                className={`docs-home-gallery-menu ${styles.galleryMenu} ${styles.symbolGalleryMenu}`}
                onEscapeKeyDown={() => {
                  restoreSymbolsTriggerOnCloseRef.current = true;
                  refocusSymbolsDraftOnCloseRef.current = false;
                }}
                onCloseAutoFocus={(event) => {
                  if (!restoreSymbolsTriggerOnCloseRef.current) {
                    event.preventDefault();
                  }
                  restoreSymbolsTriggerOnCloseRef.current = false;
                }}
              >
                <DocumentEquationSymbolGallery
                  groups={fullSymbolGalleryGroups}
                  disabled={!editable}
                  galleryTestId={`${testIdPrefix}-symbol-gallery`}
                  groupTestId={(groupId) => `${testIdPrefix}-symbol-group-${groupId}`}
                  cellTestId={(cellId) => `${testIdPrefix}-symbol-gallery-${cellId}`}
                  headingId={(groupId) => `${testIdPrefix}-symbol-heading-${groupId}`}
                  classNames={{
                    root: styles.symbolGallery,
                    group: styles.symbolCategory,
                    heading: styles.symbolCategoryHeading,
                    grid: `docs-home-gallery-menu__grid ${styles.symbolGalleryGrid}`,
                    cell: `docs-home-gallery-menu__tile ${styles.symbolGalleryTile}`,
                    glyph: styles.symbolGalleryGlyph,
                  }}
                  cellButtonProps={{
                    ...visualMathFocusOwnerProps,
                    onMouseDown: (event) => event.preventDefault(),
                  }}
                  onSelectCell={(cell) => {
                    const result = run({ type: "symbol", id: cell.id });
                    restoreSymbolsTriggerOnCloseRef.current = !succeeded(result);
                    refocusSymbolsDraftOnCloseRef.current = succeeded(result);
                  }}
                />
              </RibbonDropdownMenuContent>
            </DropdownMenu>
          }
        />
      </OfficeRibbonGroup>

      <OfficeRibbonGroup
        ribbon="equation"
        label={copy.structures}
        className={`${styles.ribbonGroup} ${styles.structureGroup}`}
        testId={`${testIdPrefix}-ribbon-structures`}
      >
        <RibbonHStack className={styles.structureStrip}>
          {EQUATION_STRUCTURES.map((family) => {
            const familyLabel = copy.structureLabels[family.id];
            const familyOpen = openStructureFamily === family.id;
            const members = equationStructureMembersForFamily(family.id);
            // The ribbon face is Word's typeset specimen (`iconLatex`), never
            // the insertion template with its hollow placeholder squares.
            const familyPreviewLatex = family.iconLatex;
            const familyPreviewMarkup = renderValidatedKatexMarkup(familyPreviewLatex);
            return (
              <DropdownMenu
                key={family.id}
                modal={false}
                open={editable && familyOpen}
                onOpenChange={(nextOpen) => {
                  if (nextOpen && !editable) return;
                  if (nextOpen) {
                    structurePointerActivationRef.current = null;
                    restoreStructureTriggerOnCloseRef.current = null;
                    refocusStructureDraftOnCloseRef.current = false;
                    setOpenStructureFamily(family.id);
                  } else {
                    setOpenStructureFamily((current) => (current === family.id ? null : current));
                  }
                }}
              >
                <DropdownMenuTrigger asChild>
                  <OfficeHomeMenuTrigger
                    {...visualMathFocusOwnerProps}
                    active={familyOpen}
                    className={styles.familyTrigger}
                    disabled={!editable}
                    hasMenuIndicator
                    icon={
                      <EquationStructureFamilyFace
                        latex={familyPreviewLatex}
                        markup={familyPreviewMarkup}
                        fallback={family.glyph}
                      />
                    }
                    label={familyLabel}
                    onMouseDown={(event) => event.preventDefault()}
                    testId={`${testIdPrefix}-structure-family-${family.id}`}
                    title={familyLabel}
                    variant="large"
                  />
                </DropdownMenuTrigger>
                <RibbonDropdownMenuContent
                  {...visualMathFocusOwnerProps}
                  align="start"
                  sideOffset={5}
                  aria-label={familyLabel}
                  className={`docs-home-gallery-menu ${styles.galleryMenu}`}
                  onEscapeKeyDown={() => {
                    restoreStructureTriggerOnCloseRef.current = family.id;
                    refocusStructureDraftOnCloseRef.current = false;
                  }}
                  onCloseAutoFocus={(event) => {
                    if (restoreStructureTriggerOnCloseRef.current !== family.id) {
                      event.preventDefault();
                    }
                    restoreStructureTriggerOnCloseRef.current = null;
                  }}
                >
                  <div className={`docs-home-gallery-menu__section ${styles.gallerySection}`}>
                    <div className={styles.galleryHeading} data-testid={`${testIdPrefix}-structure-gallery-heading`}>
                      {familyLabel}
                    </div>
                    <div className={`docs-home-gallery-menu__grid ${styles.galleryGrid}`}>
                      {members.map((member, memberIndex) => {
                        const memberLabel =
                          copy.structureMemberLabels[member.id] ?? `${familyLabel} ${memberIndex + 1}`;
                        const previewLatex = equationStructurePreviewLatex(member);
                        const previewMarkup = renderValidatedKatexMarkup(previewLatex);
                        const activateMember = () => {
                          const result = run({ type: "structureMember", id: member.id });
                          restoreStructureTriggerOnCloseRef.current = succeeded(result) ? null : family.id;
                          refocusStructureDraftOnCloseRef.current = succeeded(result);
                        };
                        const activateMemberFromSelect = () => {
                          const pointerActivation = structurePointerActivationRef.current;
                          if (pointerActivation?.memberId === member.id) {
                            structurePointerActivationRef.current = null;
                            return;
                          }
                          activateMember();
                        };
                        return (
                          <DropdownMenuItem
                            asChild
                            key={member.id}
                            disabled={!editable}
                            aria-label={memberLabel}
                            className={`docs-home-gallery-menu__tile ${styles.galleryTile}`}
                            data-testid={`${testIdPrefix}-structure-member-${member.id}`}
                            onKeyDown={(event) => {
                              if (event.key === "Enter" || event.key === " ") {
                                structurePointerActivationRef.current = null;
                              }
                            }}
                            onPointerDown={(event) => {
                              event.preventDefault();
                              const token = structurePointerActivationSerialRef.current + 1;
                              structurePointerActivationSerialRef.current = token;
                              structurePointerActivationRef.current = {
                                memberId: member.id,
                                token,
                              };
                              if (
                                event.button !== 0 ||
                                ((event.pointerType === "touch" || event.pointerType === "pen") && !event.isPrimary)
                              ) {
                                return;
                              }
                              activateMember();
                            }}
                            onSelect={activateMemberFromSelect}
                          >
                            <button
                              {...visualMathFocusOwnerProps}
                              type="button"
                              disabled={!editable}
                              onMouseDown={(event) => event.preventDefault()}
                            >
                              <span className={`docs-home-gallery-menu__preview ${styles.galleryPreview}`}>
                                <span
                                  aria-hidden
                                  aria-busy={!previewMarkup || undefined}
                                  data-equation-preview={previewLatex}
                                >
                                  {previewMarkup ? (
                                    <span
                                      // Preserve the existing gallery sizing
                                      // hook; KaTeX renders the frozen template.
                                      className={`ML__mathlive visual-math-editor__katex-preview ${styles.renderedMath}`}
                                      data-latex={previewLatex}
                                      dangerouslySetInnerHTML={{
                                        __html: previewMarkup,
                                      }}
                                    />
                                  ) : null}
                                </span>
                              </span>
                              <span className={`docs-home-gallery-menu__label ${styles.galleryLabel}`}>
                                {memberLabel}
                              </span>
                            </button>
                          </DropdownMenuItem>
                        );
                      })}
                    </div>
                  </div>
                </RibbonDropdownMenuContent>
              </DropdownMenu>
            );
          })}
        </RibbonHStack>
        {status ? (
          <span className={styles.status} role="status" data-testid={`${testIdPrefix}-conversion-status`}>
            {status}
          </span>
        ) : null}
      </OfficeRibbonGroup>
    </>
  );
}
