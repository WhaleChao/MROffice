"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { useTranslations } from "@mrnote/core/host/next-intl.tsx";
import {
  useRef,
  useState,
  type PointerEvent as ReactPointerEvent,
  type ReactNode,
} from "react";

import { OfficeHomeMenuTrigger } from "@mrnote/core/web/components/office-ribbon/OfficeHomeRibbon.tsx";
import OfficeBordersMenuBody, {
  OfficeBorderLinePreview,
  officeBorderWeightLabel,
  officeBordersMenuStyles,
  type OfficeBorderMenuSection,
} from "@mrnote/core/web/components/office-ribbon/OfficeBordersMenu.tsx";
import { RibbonDropdownMenuSubContent } from "@mrnote/core/web/components/office-ribbon/CupertinoRibbonMenu.tsx";
import {
  DropdownMenu,
  DropdownMenuCheckboxItem,
  DropdownMenuContent,
  DropdownMenuItem,
  DropdownMenuLabel,
  DropdownMenuRadioGroup,
  DropdownMenuRadioItem,
  DropdownMenuSeparator,
  DropdownMenuSub,
  DropdownMenuSubTrigger,
  DropdownMenuTrigger,
} from "@mrnote/core/web/components/ui/dropdown-menu.tsx";
import { DocsDesignIcon } from "../DocumentDesignRibbonIcons";
import edgeStyles from "./DocumentPageBorderMenu.module.css";
import {
  DEFAULT_PAGE_BORDER_EDGE,
  PAGE_BORDER_APPLY_SCOPES,
  PAGE_BORDER_MENU_STYLES,
  PAGE_BORDER_STYLE_PREVIEW,
  PAGE_BORDER_WEIGHT_POINTS,
  pageBorderAllEdgesState,
  pageBorderApplyScope,
  pageBorderColorHex,
  pageBorderColorIsAutomatic,
  pageBorderSides,
  pageBorderStyle,
  pageBorderWeightPt,
  type PageBorderApplyScope,
  type PageBorderMenuStyle,
  type PageBorders,
} from "./documentPageBorder";
import type { SectionBorderEdgeName } from "@mrnote/core/web/components/console/editor/native-document/sectionLayouts.ts";

/**
 * Word's Page Border button menu (Design > Page Borders).
 *
 * Presentation only: every choice is handed straight to one of the commands the
 * host passes in, which compose the section-layout patch (see
 * documentPageBorder.ts). The submenu chrome, the line previews, the weight
 * labels and the colour palette are the SAME shared office-ribbon primitives
 * the Home ribbon's paragraph border menu uses — page borders differ from those
 * only in carrying Word's weight as its own `w:sz`, so the weight list is a
 * separate row instead of being folded into the line style.
 */
export interface DocumentPageBorderMenuCommands {
  /** Toggle one of the four sides on or off. */
  readonly toggleSide: (side: SectionBorderEdgeName) => void;
  /** Toggle the complete four-sided box. */
  readonly toggleAllEdges: () => void;
  /** Set the line weight, in points, on every side that paints. */
  readonly setWeightPt: (points: number) => void;
  /** Set the `ST_Border` line style on every side that paints. */
  readonly setStyle: (style: PageBorderMenuStyle) => void;
  /** Set a literal colour (`#rrggbb`) on every side that paints. */
  readonly setColor: (color: string) => void;
  /** Move every painting side back to Word's Automatic colour. */
  readonly setAutomaticColor: () => void;
  /** Choose which pages/sections the border applies to. */
  readonly setApplyScope: (scope: PageBorderApplyScope) => void;
  /** `w:offsetFrom`: measure `w:space` from the paper edge or the text. */
  readonly setOffsetFrom: (offsetFrom: "page" | "text") => void;
  /** `w:zOrder`: paint the border in front of the text, or behind it. */
  readonly setInFrontOfText: (inFront: boolean) => void;
  /** Remove the page border from the current scope. */
  readonly remove: () => void;
  /** Open Borders and Shading on its Page Border page (Word's Options…). */
  readonly openDialog?: () => void;
}

interface DocumentPageBorderMenuProps {
  readonly borders: PageBorders | null;
  readonly commands: DocumentPageBorderMenuCommands;
  readonly disabled?: boolean;
  /**
   * Word's Options… "Margin" fields (`w:space` per edge). Supplied by the host
   * so the menu keeps using the ribbon's own committed number field rather
   * than growing a second one.
   */
  readonly spacingField?: ReactNode;
  readonly onPointerDown?: (
    event: ReactPointerEvent<HTMLButtonElement>,
  ) => void;
}

/** Word's default Width, in points: `w:sz="4"` is half a point. */
const DEFAULT_PEN_WEIGHT_PT = (DEFAULT_PAGE_BORDER_EDGE.sizeEighths ?? 4) / 8;

const SIDE_ICON = {
  top: "pageBorderTop",
  right: "pageBorderRight",
  bottom: "pageBorderBottom",
  left: "pageBorderLeft",
} as const;

/** Word lists the preview edges bottom, top, left, right after the whole box. */
const SIDE_ORDER = ["bottom", "top", "left", "right"] as const satisfies
  readonly SectionBorderEdgeName[];

function sideLabel(
  t: ReturnType<typeof useTranslations>,
  side: SectionBorderEdgeName,
): string {
  switch (side) {
    case "top":
      return t("toolbar.pageBorderEdgeTop");
    case "right":
      return t("toolbar.pageBorderEdgeRight");
    case "bottom":
      return t("toolbar.pageBorderEdgeBottom");
    case "left":
      return t("toolbar.pageBorderEdgeLeft");
  }
}

function applyScopeLabel(
  t: ReturnType<typeof useTranslations>,
  scope: PageBorderApplyScope,
): string {
  switch (scope) {
    case "wholeDocument":
      return t("toolbar.pageBorderApplyWholeDocument");
    case "thisSection":
      return t("toolbar.pageBorderApplyThisSection");
    case "thisSectionFirstPageOnly":
      return t("toolbar.pageBorderApplyFirstPageOnly");
    case "thisSectionAllExceptFirstPage":
      return t("toolbar.pageBorderApplyExceptFirstPage");
  }
}

/**
 * A submenu row's label: the leading glyph Word shows, then the row name.
 *
 * `swatch` is only ever passed by the Colour row, whose glyph previews the pen
 * the next border will paint with rather than a generic colour bar.
 */
function rowLabel(
  icon: "pageBorderWeight" | "pageBorderStyle" | "pageBorderColor" | "pageBorderApplyTo",
  text: string,
  swatch?: string | null,
) {
  return (
    <span className="flex min-w-0 items-center gap-2">
      <DocsDesignIcon name={icon} size={18} swatch={swatch} />
      <span className="truncate">{text}</span>
    </span>
  );
}

/**
 * A Weight / Style row: the shared line preview, then the name.
 *
 * `OfficeBorderLinePreview` is `width: 100%` by design — the shared borders menu
 * lists line styles as icon-only rows. A named row has to box it to a fixed
 * width or the preview eats the whole row and a CJK label wraps one character
 * per line.
 */
function previewRow(preview: ReactNode, text: string) {
  return (
    <span className="flex min-w-0 items-center gap-2">
      <span
        aria-hidden
        style={{
          alignItems: "center",
          display: "inline-flex",
          flex: "none",
          width: 30,
        }}
      >
        {preview}
      </span>
      <span style={{ whiteSpace: "nowrap" }}>{text}</span>
    </span>
  );
}

export default function DocumentPageBorderMenu({
  borders,
  commands,
  disabled = false,
  spacingField,
  onPointerDown,
}: DocumentPageBorderMenuProps) {
  const t = useTranslations("console-notebooks");
  const triggerRef = useRef<HTMLButtonElement | null>(null);
  const [menuOpen, setMenuOpen] = useState(false);

  const sides = pageBorderSides(borders);
  const hasBorder = sides.length > 0;
  // The pen rows stay live on a borderless section, the way Word for Mac's own
  // Page Border menu shows them: Word has no on/off flag separate from the line
  // style, so choosing a pen property IS how a border is created (see
  // pageBorderPenPatch). The rows then read Word's defaults, which is what such
  // a choice would produce.
  const penDisabled = disabled;
  // Placement is different: `w:offsetFrom`, the margin and `w:zOrder` describe
  // where an existing border sits, and nothing documents them as creating one.
  // They come alive the moment any pen row or Apply to does.
  const placementDisabled = disabled || !hasBorder;
  const weightPt = pageBorderWeightPt(borders) ?? DEFAULT_PEN_WEIGHT_PT;
  const style = pageBorderStyle(borders) ?? DEFAULT_PAGE_BORDER_EDGE.style;
  const colorHex = pageBorderColorHex(borders);

  const sections: OfficeBorderMenuSection[] = [
    {
      id: "weight",
      label: rowLabel("pageBorderWeight", t("toolbar.pageBorderWeight")),
      value: weightPt === null ? "" : String(weightPt),
      triggerTestId: "document-toolbar-page-border-weight-menu",
      disabled: penDisabled,
      options: PAGE_BORDER_WEIGHT_POINTS.map((points) => ({
        id: String(points),
        label: previewRow(
          <OfficeBorderLinePreview sizePt={points} />,
          officeBorderWeightLabel(points),
        ),
        testId: `document-toolbar-page-border-weight-${points}`,
      })),
      onSelect: (value) => {
        const points = Number(value);
        if (Number.isFinite(points) && points > 0) commands.setWeightPt(points);
      },
    },
    {
      id: "style",
      label: rowLabel("pageBorderStyle", t("toolbar.pageBorderStyle")),
      value: style ?? "",
      triggerTestId: "document-toolbar-page-border-style-menu",
      disabled: penDisabled,
      options: PAGE_BORDER_MENU_STYLES.map((id) => ({
        id,
        label: previewRow(
          <OfficeBorderLinePreview lineStyle={PAGE_BORDER_STYLE_PREVIEW[id]} />,
          t(`toolbar.pageBorderStyle_${id}` as "toolbar.pageBorderStyle_single"),
        ),
        testId: `document-toolbar-page-border-style-${id}`,
      })),
      onSelect: (value) => {
        if (PAGE_BORDER_MENU_STYLES.includes(value as PageBorderMenuStyle)) {
          commands.setStyle(value as PageBorderMenuStyle);
        }
      },
    },
    {
      id: "color",
      kind: "color-palette",
      // Word's Automatic pen paints black, so an Automatic (or absent) border
      // previews as black rather than as the ribbon's accent — the swatch has
      // to be the colour the document would get, not a decoration.
      label: rowLabel(
        "pageBorderColor",
        t("toolbar.pageBorderColor"),
        colorHex ?? "#000000",
      ),
      value: colorHex ?? "#000000",
      triggerTestId: "document-toolbar-page-border-color-menu",
      testIdPrefix: "document-toolbar-page-border-color",
      disabled: penDisabled,
      labels: {
        automatic: t("toolbar.textColorAutomatic"),
        themeColors: t("toolbar.themeColors"),
        standardColors: t("toolbar.standardColors"),
        moreColors: t("spreadsheet.moreColors"),
      },
      onAutomatic: () => commands.setAutomaticColor(),
      // Word's default pen is Automatic, so a borderless section reads as
      // Automatic rather than as nothing selected.
      automaticSelected: hasBorder ? pageBorderColorIsAutomatic(borders) : true,
      onSelect: (value) => commands.setColor(value),
    },
  ];

  return (
    <DropdownMenu open={menuOpen} onOpenChange={setMenuOpen}>
      <DropdownMenuTrigger asChild>
        <OfficeHomeMenuTrigger
          ref={triggerRef}
          active={hasBorder}
          disabled={disabled}
          hasMenuIndicator
          icon={<DocsDesignIcon name="pageBorderAllEdges" size={24} />}
          label={t("toolbar.pageBorder")}
          onPointerDown={onPointerDown}
          testId="document-toolbar-ribbon-page-border"
          variant="large"
          aria-haspopup="menu"
          aria-expanded={menuOpen}
        />
      </DropdownMenuTrigger>
      <DropdownMenuContent
        align="start"
        className={`note-editor-toolbar-menu docs-cupertino-menu ${officeBordersMenuStyles.menu}`}
        data-testid="document-toolbar-page-border-menu"
        onCloseAutoFocus={(event) => {
          event.preventDefault();
          triggerRef.current?.focus({ preventScroll: true });
        }}
      >
        <OfficeBordersMenuBody
          title={t("toolbar.pageBorderOptions")}
          presetGroupLabel={t("toolbar.pageBorderEdges")}
          presets={[
            {
              id: "allEdges",
              label: t("toolbar.pageBorderEdgeAll"),
              icon: <DocsDesignIcon name="pageBorderAllEdges" size={20} />,
              // Pressed ONLY when all four sides are present: a partly-on box
              // is not the "all four sides" setting. The tri-state still
              // reaches assistive tech as `mixed`.
              checked: pageBorderAllEdgesState(borders),
              itemProps: { className: edgeStyles.edgeToggle },
              disabled,
              testId: "document-toolbar-page-border-edge-all",
              onSelect: (event) => {
                event.preventDefault();
                commands.toggleAllEdges();
              },
            },
            ...SIDE_ORDER.map((side) => ({
              id: side,
              label: sideLabel(t, side),
              icon: <DocsDesignIcon name={SIDE_ICON[side]} size={20} />,
              checked: sides.includes(side),
              itemProps: { className: edgeStyles.edgeToggle },
              disabled,
              testId: `document-toolbar-page-border-edge-${side}`,
              onSelect: (event: Event) => {
                event.preventDefault();
                commands.toggleSide(side);
              },
            })),
          ]}
          sections={sections}
          disabled={disabled}
          afterSections={
            <>
              {/* "Apply to" carries Word's Options… placement settings under
                  the four scopes: they all answer "where does this border
                  sit", and Word reaches them from the same dialog. */}
              <DropdownMenuSub>
                <DropdownMenuSubTrigger
                  disabled={penDisabled}
                  data-testid="document-toolbar-page-border-apply-menu"
                  data-office-borders-current-value={pageBorderApplyScope(
                    borders,
                  )}
                >
                  {rowLabel("pageBorderApplyTo", t("toolbar.pageBorderApplyTo"))}
                </DropdownMenuSubTrigger>
                <RibbonDropdownMenuSubContent data-office-borders-submenu="applyTo">
                  <DropdownMenuRadioGroup value={pageBorderApplyScope(borders)}>
                    {PAGE_BORDER_APPLY_SCOPES.map((scope) => (
                      <DropdownMenuRadioItem
                        key={scope}
                        value={scope}
                        disabled={penDisabled}
                        data-testid={`document-toolbar-page-border-apply-${scope}`}
                        onSelect={(event) => {
                          event.preventDefault();
                          commands.setApplyScope(scope);
                        }}
                      >
                        {applyScopeLabel(t, scope)}
                      </DropdownMenuRadioItem>
                    ))}
                  </DropdownMenuRadioGroup>
                  <DropdownMenuSeparator />
                  <DropdownMenuLabel>
                    {t("toolbar.pageBorderMeasureFrom")}
                  </DropdownMenuLabel>
                  <DropdownMenuRadioGroup value={borders?.offsetFrom ?? "page"}>
                    {(["page", "text"] as const).map((offsetFrom) => (
                      <DropdownMenuRadioItem
                        key={offsetFrom}
                        value={offsetFrom}
                        disabled={placementDisabled}
                        data-testid={`document-toolbar-page-border-offset-${offsetFrom}`}
                        onSelect={(event) => {
                          event.preventDefault();
                          commands.setOffsetFrom(offsetFrom);
                        }}
                      >
                        {offsetFrom === "page"
                          ? t("toolbar.pageBorderMeasureFromPage")
                          : t("toolbar.pageBorderMeasureFromText")}
                      </DropdownMenuRadioItem>
                    ))}
                  </DropdownMenuRadioGroup>
                  <DropdownMenuSeparator />
                  <DropdownMenuCheckboxItem
                    checked={(borders?.zOrder ?? "front") === "front"}
                    disabled={placementDisabled}
                    data-testid="document-toolbar-page-border-front"
                    onSelect={(event) => {
                      event.preventDefault();
                      commands.setInFrontOfText(
                        (borders?.zOrder ?? "front") !== "front",
                      );
                    }}
                  >
                    {t("toolbar.pageBorderInFrontOfText")}
                  </DropdownMenuCheckboxItem>
                  {spacingField}
                </RibbonDropdownMenuSubContent>
              </DropdownMenuSub>
              <DropdownMenuSeparator />
              <DropdownMenuItem
                disabled={placementDisabled}
                data-testid="document-toolbar-page-border-remove"
                onSelect={() => commands.remove()}
              >
                <span className="flex min-w-0 items-center gap-2">
                  <DocsDesignIcon name="pageBorderRemove" size={18} />
                  <span className="truncate">
                    {t("toolbar.pageBorderRemove")}
                  </span>
                </span>
              </DropdownMenuItem>
              {commands.openDialog ? (
                <DropdownMenuItem
                  disabled={disabled}
                  data-testid="document-toolbar-page-border-dialog"
                  onSelect={() => commands.openDialog?.()}
                >
                  {t("toolbar.bordersShadingEllipsis")}
                </DropdownMenuItem>
              ) : null}
            </>
          }
        />
      </DropdownMenuContent>
    </DropdownMenu>
  );
}
