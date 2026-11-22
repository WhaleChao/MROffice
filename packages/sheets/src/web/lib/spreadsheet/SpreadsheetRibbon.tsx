"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { OfficeSemanticIcon } from "@mrnote/core/web/components/office-ribbon/icons/OfficeSemanticIcon.tsx";
import {
  Fragment,
  type ComponentProps,
  type ReactNode,
  type Ref,
  useEffect,
  useId,
  useRef,
  useState,
} from "react";
import {
  OfficeRibbonGroup,
  type OfficeRibbonGroupLauncher,
  OfficeRibbonTabs,
  RibbonHStack,
} from "@mrnote/core/web/components/office-ribbon/OfficeRibbonChassis.tsx";
import {
  OfficeHomeCommand,
  OfficeHomeMenuTrigger,
  OfficeHomeRibbonGroup,
  OfficeHomeUndoRail,
  type OfficeHomeCommandProps,
  type OfficeHomeGroupSize,
  type OfficeHomeMenuTriggerProps,
} from "@mrnote/core/web/components/office-ribbon/OfficeHomeRibbon.tsx";
import { OfficeRibbonSurfaceStrip } from "@mrnote/core/web/components/office-ribbon/OfficeRibbonSurfaceStrip.tsx";
import { OfficeRibbonQuickStrip } from "@mrnote/core/web/components/office-ribbon/OfficeRibbonQuickStrip.tsx";
import {
  OfficeEditorActionRail,
  type OfficeEditorActionRailProps,
} from "@mrnote/core/web/components/office-ribbon/OfficeEditorActionRail.tsx";
import {
  DropdownMenu,
  DropdownMenuTrigger,
} from "@mrnote/core/web/components/ui/dropdown-menu.tsx";
import { RibbonDropdownMenuContent } from "@mrnote/core/web/components/office-ribbon/CupertinoRibbonMenu.tsx";
import {
  computeRibbonOverflow,
  type RibbonOverflowState,
} from "@mrnote/core/web/components/console/editor/ribbonOverflow.ts";
import spreadsheetHomeStyles from "../../components/notebook/contents/spreadsheet/SpreadsheetHomeRibbon.module.css";
import {
  hideOfficeScreenTip,
  OFFICE_SCREEN_TIP_HOVER_DELAY_MS,
  officeScreenTipTarget,
  registerOfficeScreenTipDescriber,
  showOfficeScreenTip,
} from "@mrnote/core/web/components/office-ribbon/officeScreenTip.ts";

export type SpreadsheetRibbonCoreTabId =
  | "home"
  | "insert"
  | "pageLayout"
  | "formulas"
  | "data"
  | "review"
  | "view"
  | "draw"
  | "automate";

export type SpreadsheetRibbonTabId =
  | SpreadsheetRibbonCoreTabId
  | "tableDesign"
  | "chartDesign"
  | "chartFormat"
  // Office's PivotChart Analyze, in front of Design and Format while a
  // PivotChart is selected.
  | "pivotChartAnalyze"
  | "pivotTable"
  | "sparklineDesign"
  | "slicerDesign"
  | "objectFormat"
  // Office's own contextual pair for a diagram: SmartArt Design (the shared
  // `OfficeSmartArtDesignTab`) alongside the member's Format tab, exactly as
  // chartDesign travels with chartFormat.
  | "smartArtDesign"
  // Office's Picture Format tab for a SmartArt picture slot holding a picture
  // (Crop > Fill / Fit, Picture Styles, Reset Picture), beside SmartArt Design.
  | "smartArtPicture"
  // Office's Equation tab (the shared `OfficeEquationContextRibbon`), after
  // Shape Format while an equation text box is selected or open in place.
  | "equation";

const SPREADSHEET_WEB_CONTEXTUAL_100PX_TABS = new Set<SpreadsheetRibbonTabId>([
  "tableDesign",
  "pivotTable",
  "pivotChartAnalyze",
  "chartDesign",
  "chartFormat",
  "objectFormat",
  "smartArtDesign",
  "smartArtPicture",
  "slicerDesign",
  "equation",
]);

const SPREADSHEET_WEB_PERSISTENT_HIT_LANE_TABS = new Set<SpreadsheetRibbonTabId>([
  "insert",
  "pageLayout",
  "formulas",
  "data",
  "review",
  "view",
  "draw",
  "automate",
]);

export interface SpreadsheetRibbonTab {
  id: SpreadsheetRibbonTabId;
  label: string;
  contextual?:
    | "table"
    | "chart"
    | "pivot"
    | "sparkline"
    | "slicer"
    | "object";
}

export interface SpreadsheetRibbonGroup {
  id: string;
  label: string;
  /** A factory defers element construction until the actual group host mounts. */
  content: ReactNode | (() => ReactNode);
  /** Shared Office groups may provide their own complete group anatomy. */
  unwrapped?: boolean;
  /** Omit a group whose entire command family is permanently Web-unavailable. */
  hidden?: boolean;
  homeSize?: OfficeHomeGroupSize;
  homeWidth?: number;
  /** Mount this command host once in the title-row compact surface, not the Home track. */
  compactOnly?: boolean;
  /** The group's dialog box launcher, in its lower right corner. */
  launcher?: OfficeRibbonGroupLauncher;
  




  collapsedIcon?: ReactNode;
}

/** Keep invocation inside a mounted component. Resolving this in the parent
 * map would eagerly construct compact dialogs even when their flyout is closed. */
function LazyGroupContent({ content, active = true }: { content: SpreadsheetRibbonGroup["content"]; active?: boolean }) {
  return <>{typeof content === "function" ? (active ? content() : null) : content}</>;
}

export type SpreadsheetRibbonGroupSource =
  | readonly SpreadsheetRibbonGroup[]
  | (() => readonly SpreadsheetRibbonGroup[]);

export interface SpreadsheetRibbonProps {
  activeTab: SpreadsheetRibbonTabId;
  onActiveTabChange: (tab: SpreadsheetRibbonTabId) => void;
  groups: Partial<
    Record<SpreadsheetRibbonTabId, SpreadsheetRibbonGroupSource>
  >;
  translate: (key: string) => string;
  tabs?: readonly SpreadsheetRibbonTab[];
  quickActions?: ReactNode;
  quickActionTail?: ReactNode;
  actionRail?: Pick<OfficeEditorActionRailProps, "leading" | "search" | "undo" | "redo" | "save" | "ai" | "primary">;
  quickOverflowLabel?: string;
  ariaLabel?: string;
  compactOverflowOpen?: boolean;
  onCompactOverflowOpenChange?: (open: boolean) => void;
  homeUndoRail?: {
    label: string;
    undo: OfficeHomeCommandProps;
    redo: OfficeHomeCommandProps;
    testId?: string;
  };
  /**
   * XLS-A106: the enhanced ScreenTip of a control, by its test id — its
   * sentence, and the shortcut Excel prints after the name when the control's
   * own title does not name one already.
   */
  screenTip?: (testId: string) => SpreadsheetRibbonScreenTip | string | undefined;
  /** Pen-first drawing (lib/pen-first-draw): the ribbon folds away for the floating pen palette. */
  penCompact?: boolean;
  /** A tap on the Draw tab folds the ribbon (OfficeRibbonTabs `collapse`). */
  collapseRibbon?: { onCollapse: () => void };
}

export interface SpreadsheetRibbonScreenTip {
  description: string;
  shortcut?: string;
}

export interface SpreadsheetRibbonActionProps {
  label: string;
  displayLabel?: string;
  icon: ReactNode;
  onClick: () => void;
  onDoubleClick?: OfficeHomeCommandProps["onDoubleClick"];
  /** Exact Atlas/verified-command identity rendered on the physical control. */
  controlId?: string;
  activation?: OfficeHomeCommandProps["activation"];
  active?: boolean | "mixed";
  armedState?: string;
  disabled?: boolean;
  compact?: boolean;
  homeVariant?: OfficeHomeCommandProps["variant"];
  testId?: string;
  /** Exact rendered trigger for dialog focus return. */
  buttonRef?: Ref<HTMLButtonElement>;
  /** Appended to the accessible name/title when this command is disabled. */
  disabledReason?: string;
  /** Permanent Web-unavailable commands are omitted instead of shown inert. */
  hidden?: boolean;
  /** Live Preview of a format toggle (see OfficeHomeCommand). */
  formatPreview?: OfficeHomeCommandProps["formatPreview"];
  /** The command formats characters (see OfficeHomeCommand). */
  textFormatControl?: boolean;
}

export interface SpreadsheetRibbonMenuActionProps {
  label: string;
  displayLabel?: string;
  icon: ReactNode;
  children: ReactNode;
  /** Exact Atlas/verified-command identity rendered on the physical trigger. */
  controlId?: string;
  active?: boolean;
  disabled?: boolean;
  variant?: "big" | "inline" | "compact";
  homeVariant?: OfficeHomeCommandProps["variant"];
  testId?: string;
  menuClassName?: string;
  onCloseAutoFocus?: ComponentProps<
    typeof RibbonDropdownMenuContent
  >["onCloseAutoFocus"];
  align?: "start" | "center" | "end";
  className?: string;
  triggerRef?: Ref<HTMLButtonElement>;
  /** Appended to the accessible name/title when this menu root is disabled. */
  disabledReason?: string;
  /** Permanent Web-unavailable menu roots are omitted instead of shown inert. */
  hidden?: boolean;
  /**
   * "gallery" draws only the disclosure chevron. Used when the value the menu
   * switches is shown by a sibling — the Sheet View name box — so the trigger
   * does not repeat it.
   */
  geometry?: OfficeHomeMenuTriggerProps["geometry"];
  /** The open editors: the root is a button whose click is the desktop-app notice. */
  onNotice?: () => void;
}

export const SPREADSHEET_RIBBON_TAB_MESSAGE_KEYS = { home: "toolbar.ribbonHome", insert: "toolbar.ribbonInsert" } as const satisfies Record<SpreadsheetRibbonCoreTabId, string>;







export const SPREADSHEET_WEB_BASE_TAB_IDS = ["home", "insert"] as const satisfies readonly SpreadsheetRibbonCoreTabId[];

export function isSpreadsheetWebBaseRibbonTabId(
  tab: SpreadsheetRibbonTabId,
): tab is SpreadsheetRibbonCoreTabId {
  return (SPREADSHEET_WEB_BASE_TAB_IDS as readonly string[]).includes(tab);
}

export function spreadsheetRibbonTabs(
  translate: (key: string) => string,
): readonly SpreadsheetRibbonTab[] {
  return SPREADSHEET_WEB_BASE_TAB_IDS.map((id) => ({
    id,
    label: translate(SPREADSHEET_RIBBON_TAB_MESSAGE_KEYS[id]),
  }));
}

/** Marks the ribbon root whose hover lookup a touch hold may borrow. */
const SCREEN_TIP_ROOT_ATTRIBUTE = "data-spreadsheet-ribbon-screen-tips";

/**
 * XLS-A106 — Excel's enhanced ScreenTip: the command's name and shortcut in
 * bold, then a sentence on what it does. The ribbon shows it on hover (after
 * Office's half-second) or keyboard focus for every control the host gives a
 * description; the control's plain title stays the fallback elsewhere. The
 * card itself is the shared Office one (`officeScreenTip.ts`), which the
 * Presentation and Docs chrome show too.
 */
function useSpreadsheetRibbonScreenTips(
  describe: ((testId: string) => SpreadsheetRibbonScreenTip | string | undefined) | undefined,
) {
  type Found = { element: HTMLElement; description: string; shortcut?: string };
  const targetOf = (node: EventTarget | null, root: EventTarget | null): Found | null => {
    if (!describe || !(node instanceof Element)) return null;
    // Only the ribbon's own controls: a menu it opened portals elsewhere in
    // the DOM (its events still bubble here through React) and keeps the
    // plain tooltips of its items.
    if (root instanceof Node && !root.contains(node)) return null;
    let element: Element | null = node.closest("[data-testid]");
    while (element instanceof HTMLElement) {
      const testId = element.dataset.testid;
      const tip = testId ? describe(testId) : undefined;
      const description = typeof tip === "string" ? tip : tip?.description;
      if (description) {
        return {
          element,
          description,
          shortcut: typeof tip === "string" ? undefined : tip?.shortcut,
        };
      }
      element = element.parentElement?.closest("[data-testid]") ?? null;
    }
    return null;
  };
  const show = (found: Found, delay: number) =>
    showOfficeScreenTip(found, delay, { testId: "spreadsheet-ribbon-screen-tip", owner: "spreadsheet-ribbon" });
  const clear = () => {
    if (officeScreenTipTarget("spreadsheet-ribbon")) hideOfficeScreenTip("spreadsheet-ribbon");
  };
  useEffect(() => clear, []);
  // A finger cannot hover: a hold on a ribbon control (lib/touch-hold) shows
  // the card this hover would.
  const rootId = useId();
  const targetOfRef = useRef(targetOf);
  useEffect(() => {
    targetOfRef.current = targetOf;
  });
  useEffect(
    () =>
      registerOfficeScreenTipDescriber((node) => {
        const root = node.closest(`[${SCREEN_TIP_ROOT_ATTRIBUTE}]`);
        return root?.getAttribute(SCREEN_TIP_ROOT_ATTRIBUTE) === rootId
          ? targetOfRef.current(node, root)
          : null;
      }),
    [rootId],
  );
  const handlers = describe
    ? {
        onPointerOver: (event: { target: EventTarget | null; currentTarget: EventTarget | null }) => {
          const found = targetOf(event.target, event.currentTarget);
          if (!found) {
            clear();
            return;
          }
          show(found, OFFICE_SCREEN_TIP_HOVER_DELAY_MS);
        },
        onPointerLeave: () => clear(),
        onPointerDown: () => clear(),
        onFocus: (event: { target: EventTarget | null; currentTarget: EventTarget | null }) => {
          const found = targetOf(event.target, event.currentTarget);
          if (found && (event.target as Element).matches?.(":focus-visible")) show(found, 0);
        },
        onBlur: () => clear(),
        onKeyDown: (event: { key: string }) => {
          if (event.key === "Escape") clear();
        },
      }
    : {};
  return { handlers, rootProps: { [SCREEN_TIP_ROOT_ATTRIBUTE]: rootId } };
}

/**
 * Which way the ribbon strip can still scroll. Kept as a named export because
 * it is unit-tested; the arithmetic itself is the chassis's, so there is one
 * implementation rather than two that agree by coincidence.
 */
export function spreadsheetRibbonOverflow(input: {
  scrollLeft: number;
  scrollWidth: number;
  clientWidth: number;
}): RibbonOverflowState {
  return computeRibbonOverflow(input);
}

export function SpreadsheetRibbonAction({
  label,
  displayLabel,
  icon,
  onClick,
  onDoubleClick,
  controlId,
  activation,
  // Deliberately NOT defaulted to false. OfficeRibbonButton renders
  // `aria-pressed` for any BOOLEAN active, so a false default made every plain
  // command in the ribbon — Format as Table, Insert Chart, Clear Filters —
  // announce itself to a screen reader as an unpressed TOGGLE. Only a real
  // two-state control passes `active`, and only it should carry aria-pressed.
  active,
  armedState,
  disabled = false,
  compact = false,
  homeVariant,
  testId,
  buttonRef,
  disabledReason,
  hidden = false,
  formatPreview,
  textFormatControl,
}: SpreadsheetRibbonActionProps) {
  if (hidden) return null;
  const accessibleLabel =
    disabled && disabledReason ? `${label}: ${disabledReason}` : label;
  // Legacy callers retain their sizing intent, but all faces now use the same
  // shared command anatomy. The shared chassis preserves argument-free dispatch.
  return (
    <OfficeHomeCommand
      activation={activation}
      active={active}
      armedState={armedState}
      buttonRef={buttonRef}
      controlId={controlId}
      disabled={disabled}
      displayLabel={displayLabel ?? label}
      formatPreview={formatPreview}
      textFormatControl={textFormatControl}
      icon={icon}
      label={accessibleLabel}
      onClick={onClick}
      onDoubleClick={onDoubleClick}
      testId={testId}
      variant={homeVariant ?? (compact ? "icon" : "large")}
    />
  );
}

/**
 * Spreadsheet-owned menu chevron slot. It draws the shared Office disclosure
 * glyph (the same 12px `menu.chevron.down` every integral and split disclosure
 * wears in the Docs and Presentation ribbons) inside a clipped slot so primary
 * 20/24px command glyphs cannot displace it.
 */
export function SpreadsheetRibbonMenuChevron({ size = 12 }: { size?: number }) {
  const resolvedSize = Math.max(12, Math.min(14, size));
  return (
    <span
      aria-hidden
      data-spreadsheet-home-menu-chevron-slot
      style={{
        alignItems: "center",
        display: "inline-flex",
        flex: `0 0 ${resolvedSize}px`,
        height: resolvedSize,
        justifyContent: "center",
        overflow: "hidden",
        width: resolvedSize,
      }}
    >
      <OfficeSemanticIcon semanticId="menu.chevron.down" size={resolvedSize} />
    </span>
  );
}

/** Shared dropdown command anatomy for the spreadsheet Home ribbon. */
/**
 * Excel leaves the keyboard in the worksheet after a menu command that
 * opens nothing (Insert Sheet, Delete Sheet, Insert Sheet Rows ...). The
 * sheet takes it a frame after the command, while the menu returns focus
 * to its trigger on a timer, so either can run first. A menu whose close
 * finds the keyboard already in the grid leaves it there.
 */
export function keepWorksheetKeyboardOnMenuClose(event: Event): void {
  if (event.defaultPrevented || typeof document === "undefined") return;
  const active = document.activeElement;
  if (active instanceof HTMLElement && active.closest('[role="grid"]')) {
    event.preventDefault();
  }
}

export function SpreadsheetRibbonMenuAction({
  label,
  displayLabel,
  icon,
  children,
  controlId,
  active,
  disabled = false,
  variant = "big",
  homeVariant,
  testId,
  menuClassName,
  onCloseAutoFocus,
  align = "start",
  className,
  disabledReason,
  triggerRef,
  hidden = false,
  geometry,
  onNotice,
}: SpreadsheetRibbonMenuActionProps) {
  if (hidden) return null;
  const accessibleLabel =
    disabled && disabledReason ? `${label}: ${disabledReason}` : label;
  if (onNotice) {
    return (
      <OfficeHomeMenuTrigger
        active={active ? true : undefined}
        className={className}
        controlId={controlId}
        disabled={disabled}
        displayLabel={displayLabel ?? label}
        geometry={geometry}
        hasMenuIndicator
        icon={icon}
        label={accessibleLabel}
        onClick={onNotice}
        testId={testId}
        variant={homeVariant ?? (variant === "big" ? "large" : variant === "inline" ? "compact" : "icon")}
        ref={triggerRef}
      />
    );
  }
  // Non-modal, as the split-button menus already are: a modal Radix menu
  // blocks pointer events outside it, so the click that dismisses the menu
  // was swallowed. In Excel the same click closes the menu AND selects the
  // cell under it. Radix then also skips the focus return to the trigger
  // for an outside click, so the focus stays with the clicked cell; Escape
  // and item selection still return it.
  return (
    <DropdownMenu modal={false}>
      <DropdownMenuTrigger asChild>
        <OfficeHomeMenuTrigger
          active={active ? true : undefined}
          className={className}
          controlId={controlId}
          disabled={disabled}
          displayLabel={displayLabel ?? label}
          geometry={geometry}
          hasMenuIndicator
          icon={icon}
          label={accessibleLabel}
          testId={testId}
          variant={homeVariant ?? (variant === "big" ? "large" : variant === "inline" ? "compact" : "icon")}
          ref={triggerRef}
        />
      </DropdownMenuTrigger>
      <RibbonDropdownMenuContent
        align={align}
        className={menuClassName}
        onCloseAutoFocus={(event) => {
          onCloseAutoFocus?.(event);
          keepWorksheetKeyboardOnMenuClose(event);
        }}
      >
        {children}
      </RibbonDropdownMenuContent>
    </DropdownMenu>
  );
}

export function SpreadsheetRibbon({
  activeTab,
  onActiveTabChange,
  groups,
  translate,
  tabs = spreadsheetRibbonTabs(translate),
  quickActions,
  quickActionTail,
  actionRail,
  quickOverflowLabel = translate("toolbar.quickOverflow"),
  ariaLabel = translate("spreadsheet.ribbon.label"),
  homeUndoRail,
  compactOverflowOpen,
  onCompactOverflowOpenChange,
  screenTip,
  penCompact = false,
  collapseRibbon,
}: SpreadsheetRibbonProps) {
  const screenTips = useSpreadsheetRibbonScreenTips(screenTip);
  const panelId = useId();
  const [uncontrolledCompactOpen, setUncontrolledCompactOpen] = useState(false);
  const compactOpen = compactOverflowOpen ?? uncontrolledCompactOpen;
  const setCompactOpen = (open: boolean) => {
    if (compactOverflowOpen === undefined) setUncontrolledCompactOpen(open);
    onCompactOverflowOpenChange?.(open);
  };
  const resolveGroups = (
    source: SpreadsheetRibbonGroupSource | undefined,
  ): readonly SpreadsheetRibbonGroup[] =>
    typeof source === "function" ? source() : (source ?? []);
  // SpreadsheetWindow owns hundreds of command elements across ten base tabs.
  // Resolve only the active tab's factory so editing one cell does not rebuild
  // every hidden ribbon page. Array sources remain supported for adapters and
  // contextual groups that are already small/precomputed.
  const activeGroups = resolveGroups(groups[activeTab]);
  const visibleGroups = activeGroups.filter(
    (group) => !group.compactOnly && !group.hidden,
  );
  const compactGroups = activeGroups.filter(
    (group) => group.compactOnly && !group.hidden,
  );
  const usesCanonicalContextualDeck =
    SPREADSHEET_WEB_CONTEXTUAL_100PX_TABS.has(activeTab);
  const usesPersistentSecondaryHitLane =
    SPREADSHEET_WEB_PERSISTENT_HIT_LANE_TABS.has(activeTab);
  const activeTabLabel =
    tabs.find((tab) => tab.id === activeTab)?.label ?? activeTab;

  return (
    <div
      className="note-editor-toolbar-stack"
      data-testid="spreadsheet-ribbon-stack"
    >
      <div
        {...screenTips.handlers}
        {...screenTips.rootProps}
        aria-label={ariaLabel}
        className={`note-editor-office-toolbar office-ribbon-skin ${spreadsheetHomeStyles.webSurface}`}
        data-office-authoring-ribbon="true"
        data-pen-compact={penCompact ? "true" : undefined}
        data-spreadsheet-ribbon-surface="web"
        data-testid="spreadsheet-ribbon"
        role="toolbar"
      >
        <OfficeRibbonTabs
          collapse={collapseRibbon}
          tabs={tabs.map((tab) => ({
            key: tab.id,
            label: tab.label,
            contextual: tab.contextual,
            className:
              tab.contextual === "table"
                ? "text-emerald-700"
                : tab.contextual === "chart"
                  ? "text-emerald-700"
                : tab.contextual === "sparkline"
                  ? "text-emerald-700"
                  : tab.contextual === "pivot"
                    ? "text-violet-700"
                    : tab.contextual === "object"
                      ? "text-sky-700"
                      : undefined,
          }))}
          activeTab={activeTab}
          onSelect={onActiveTabChange}
          testIdPrefix="spreadsheet"
          listLabel={
            translate("spreadsheet.ribbon.tabs")
          }
          panelId={panelId}
          trailing={
            actionRail ? (
              <OfficeEditorActionRail
                {...actionRail}
                ariaLabel={translate("spreadsheet.ribbon.quickActions")}
                overflowLabel={quickOverflowLabel}
                overflowOpen={compactOpen}
                onOverflowOpenChange={setCompactOpen}
                testIds={{
                  strip: "spreadsheet-ribbon-quick-actions",
                  overflow: "spreadsheet-ribbon-quick-overflow",
                  overflowPanel: "spreadsheet-ribbon-quick-overflow-panel",
                }}
                // Nothing to overflow (no tail, no compact-only group): the
                // rail then draws no "more" button at all.
                overflow={
                  quickActionTail || compactGroups.length > 0 ? (
                    <>
                      {quickActionTail}
                      {compactGroups.length > 0 ? (
                        <span data-testid="spreadsheet-ribbon-compact-overflow" style={{ display: "contents" }}>
                          {compactGroups.map((group) => (
                            <span key={group.id} data-testid={`spreadsheet-ribbon-compact-${group.id}`}>
                              <LazyGroupContent content={group.content} active={compactOpen} />
                            </span>
                          ))}
                        </span>
                      ) : null}
                    </>
                  ) : undefined
                }
              />
            ) : quickActions || quickActionTail || compactGroups.length > 0 ? (
              <OfficeRibbonQuickStrip
                ariaLabel={translate("spreadsheet.ribbon.quickActions")}
                overflowOpen={compactOpen}
                onOverflowOpenChange={setCompactOpen}
                overflowLabel={quickOverflowLabel}
                tail={quickActionTail}
                overflowOnlyTail={
                  compactGroups.length > 0 ? (
                    <span
                      data-testid="spreadsheet-ribbon-compact-overflow"
                      style={{ display: "contents" }}
                    >
                      {compactGroups.map((group) => (
                        <span
                          key={group.id}
                          data-testid={`spreadsheet-ribbon-compact-${group.id}`}
                        >
                          <LazyGroupContent content={group.content} />
                        </span>
                      ))}
                    </span>
                  ) : undefined
                }
                testIds={{
                  strip: "spreadsheet-ribbon-quick-actions",
                  overflow: "spreadsheet-ribbon-quick-overflow",
                  overflowPanel: "spreadsheet-ribbon-quick-overflow-panel",
                }}
              >
                {quickActions}
              </OfficeRibbonQuickStrip>
            ) : null
          }
        />

        {/* One strip for every tab (OfficeRibbonSurfaceStrip): Home and the
            other tabs used to be two strip components, so every crossing
            to or from Home unmounted the whole ribbon strip. */}
        <OfficeRibbonSurfaceStrip
          surface={activeTab === "home" ? "home" : "insert"}
          density="office"
          // Current Excel-Web object carriers use the shared canonical 100px
          // deck. Persistent secondary pages retain Spreadsheet's existing
          // 104px geometry; Sparkline remains gated until Web authoring is
          // proven. Command authority and overflow stay adapter-owned.
          bandHeight={usesCanonicalContextualDeck ? 100 : 104}
          className={
            activeTab === "home"
              ? `${spreadsheetHomeStyles.homeSurface} ${spreadsheetHomeStyles.homePrimaryHitLane}`
              : [
                  usesCanonicalContextualDeck
                    ? spreadsheetHomeStyles.contextualDeck100
                    : usesPersistentSecondaryHitLane
                      ? spreadsheetHomeStyles.persistentSecondaryHitLane
                      : undefined,
                  undefined,
                ].filter(Boolean).join(" ")
          }
          activeTab={activeTab}
          panelId={panelId}
          panelLabel={activeTabLabel}
          panelLabelledBy={`spreadsheet-ribbon-tab-${activeTab}`}
          scrollLeftLabel={
            translate("toolbar.ribbonScrollLeft")
          }
          scrollRightLabel={
            translate("toolbar.ribbonScrollRight")
          }
          testIdPrefix={activeTab === "home" ? "spreadsheet-ribbon" : undefined}
          testIds={{
            body: "spreadsheet-ribbon-panel",
            track: "spreadsheet-ribbon-track",
            scrollLeft: "spreadsheet-ribbon-scroll-left",
            scrollRight: "spreadsheet-ribbon-scroll-right",
          }}
        >
          {activeTab === "home" ? (
            <>
              {homeUndoRail ? <OfficeHomeUndoRail {...homeUndoRail} /> : null}
              {
}
              {visibleGroups.map((group) => (
                <OfficeHomeRibbonGroup
                  key={group.id}
                  className={spreadsheetHomeStyles.homeGroup}
                  label={group.label}
                  ribbon={activeTab}
                  testId={`spreadsheet-ribbon-group-${group.id}`}
                  launcher={group.launcher}
                >
                  <div
                    className={`note-editor-ribbon-hstack ${spreadsheetHomeStyles.homeGroupContent}`}
                    data-testid={`spreadsheet-ribbon-group-content-${group.id}`}
                  >
                    <LazyGroupContent content={group.content} />
                  </div>
                </OfficeHomeRibbonGroup>
              ))}
            </>
          ) : (
            visibleGroups.map((group) => group.unwrapped ? <Fragment key={group.id}><LazyGroupContent content={group.content} /></Fragment> : (
              <OfficeRibbonGroup
                key={group.id}
                className={spreadsheetHomeStyles.secondaryGroup}
                label={group.label}
                ribbon={activeTab}
                testId={`spreadsheet-ribbon-group-${group.id}`}
                launcher={group.launcher}
              >
                <RibbonHStack
                  className={spreadsheetHomeStyles.secondaryGroupContent}
                  testId={`spreadsheet-ribbon-group-content-${group.id}`}
                >
                  <LazyGroupContent content={group.content} />
                </RibbonHStack>
              </OfficeRibbonGroup>
            ))
          )}
        </OfficeRibbonSurfaceStrip>
      </div>
    </div>
  );
}
