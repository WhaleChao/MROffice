"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import {
  useCallback,
  useEffect,
  useMemo,
  useRef,
  useState,
  type CSSProperties,
  type KeyboardEvent as ReactKeyboardEvent,
  type ReactNode,
} from "react";

import {
  createPresentationSlideTransitionAdvanceCommandAuthority,
  createPresentationSlideTransitionAdvanceCommandLineage,
  type PresentationSlideTransitionAdvanceCommandAuthority,
} from "./commandRegistry";
import { RibbonDropdownMenuContent } from "../../../office-ribbon/CupertinoRibbonMenu";
import {
  PRESENTATION_BUILTIN_TRANSITION_SOUNDS,
  PRESENTATION_EXTRA_TRANSITION_SOUNDS,
  builtinTransitionSoundIdForName,
  type PresentationBuiltinTransitionSoundId,
} from "../../../../lib/presentation-transition-builtin-sounds";
import {
  OfficeHomeCheckbox,
  OfficeHomeMenuTrigger,
  OfficeHomeNumberField,
} from "../../../office-ribbon/OfficeHomeRibbon";
import {
  DropdownMenu,
  DropdownMenuItem,
  DropdownMenuRadioGroup,
  DropdownMenuRadioItem,
  DropdownMenuSeparator,
  DropdownMenuSub,
  DropdownMenuSubContent,
  DropdownMenuSubTrigger,
  DropdownMenuTrigger,
} from "../../../ui/dropdown-menu";
import {
  createPresentationSlideTransitionAdvanceCommandAdapter,
  type PresentationSlideTransitionAdvanceCommandPayload,
  type PresentationSlideTransitionCommandAccessors,
  type PresentationSlideTransitionSlide,
} from "../../../../lib/presentation-command-adapters";
import {
  formatTransitionAdvanceAfterSeconds,
  isTransitionAdvanceAfterMs,
  MAX_EDITABLE_TRANSITION_ADVANCE_AFTER_MS,
  parseTransitionAdvanceAfterSeconds,
  TRANSITION_ADVANCE_AFTER_STEP_MS,
  TRANSITION_ADVANCE_AFTER_STEP_SECONDS,
  transitionAdvanceOnClick,
  type PresentationTransitionSound,
} from "../../../../lib/presentation-transitions";

type AuthorityInvokeArgs = Parameters<
  PresentationSlideTransitionAdvanceCommandAuthority["invoke"]
>;

/**
 * Transitions > Timing > Advance Slide: On Mouse Click and After. The command
 * adapter, its registry authority and the native `advClick`/`advTm` round trip
 * already exist; this hook is the one place the ribbon reaches them, so the
 * window only wires the slide accessors in.
 */
export function usePresentationSlideTransitionAdvance<
  T extends PresentationSlideTransitionSlide,
>({
  instanceId,
  accessors,
  capabilityContext,
  runtime,
}: {
  instanceId: string;
  accessors: PresentationSlideTransitionCommandAccessors<T>;
  capabilityContext: AuthorityInvokeArgs[0];
  runtime: AuthorityInvokeArgs[1];
}) {
  const { getSlides, getActiveIndex, commitSlides } = accessors;
  const adapter = useMemo(
    () =>
      createPresentationSlideTransitionAdvanceCommandAdapter<T>({
        getSlides,
        getActiveIndex,
        commitSlides,
      }),
    [commitSlides, getActiveIndex, getSlides],
  );
  const adapterRef = useRef(adapter);
  useEffect(() => {
    adapterRef.current = adapter;
  }, [adapter]);
  const [lineage] = useState(() =>
    createPresentationSlideTransitionAdvanceCommandLineage(instanceId),
  );
  const pendingRef =
    useRef<PresentationSlideTransitionAdvanceCommandPayload | null>(null);
  const authorityRef =
    useRef<PresentationSlideTransitionAdvanceCommandAuthority | null>(null);
  useEffect(() => {
    const authority = createPresentationSlideTransitionAdvanceCommandAuthority({
      lineage,
      execute: () => {
        const payload = pendingRef.current;
        return payload ? adapterRef.current.execute(payload) : false;
      },
      isEnabled: () => adapterRef.current.canExecute(),
    });
    authorityRef.current = authority;
    return () => {
      if (authorityRef.current === authority) authorityRef.current = null;
    };
  }, [lineage]);

  const setAdvance = useCallback(
    (advanceOnClick: boolean, advanceAfterMs: number | null) => {
      if (pendingRef.current) return;
      const active = getSlides()[getActiveIndex()];
      const authority = authorityRef.current;
      if (!active || !authority) return;
      const payload: PresentationSlideTransitionAdvanceCommandPayload = {
        slideModelId: active.id,
        advanceOnClick,
        advanceAfterMs,
      };
      pendingRef.current = payload;
      void authority.invoke(capabilityContext, runtime).finally(() => {
        if (pendingRef.current === payload) pendingRef.current = null;
      });
    },
    [capabilityContext, getActiveIndex, getSlides, runtime],
  );

  return { editable: adapter.canExecute(), setAdvance };
}

export interface PresentationTransitionTimingLabels {
  readonly duration: string;
  readonly sound: string;
  readonly soundNone: string;
  readonly soundOther: string;
  /** The submenu of sounds beyond PowerPoint's list. */
  readonly moreSounds: string;
  /** Each built-in sound's name in the UI language (E037). */
  readonly builtinSound: (id: PresentationBuiltinTransitionSoundId) => string;
  readonly advanceOnClick: string;
  readonly advanceAfter: string;
  readonly advanceAfterSeconds: string;
}

/**
 * The right half of PowerPoint's Timing group: Sound, On Mouse Click and
 * After. Duration and Apply To All stay with the host, which already owns
 * their drafts and commands; they are passed in as `duration` / `applyAll`
 * so the group keeps PowerPoint's two-row reading order.
 */
export function PresentationTransitionTimingControls({
  slideModelId,
  advanceOnClick: rawAdvanceOnClick,
  advanceAfterMs: rawAdvanceAfterMs,
  advanceEditable,
  onAdvanceChange,
  sound,
  preservedSoundName = null,
  soundEditable,
  onSoundNone,
  onSoundBuiltin,
  onSoundOther,
  labels,
  duration,
  applyAll,
  soundIcon,
  gridClassName,
  afterFieldClassName,
}: {
  slideModelId: string;
  advanceOnClick: unknown;
  advanceAfterMs: unknown;
  advanceEditable: boolean;
  onAdvanceChange: (advanceOnClick: boolean, advanceAfterMs: number | null) => void;
  sound: PresentationTransitionSound | null;
  /**
   * The name of an imported sound this editor keeps but cannot re-author
   * (a linked or non-WAV `p:snd`): shown, never replaced by "[No Sound]".
   */
  preservedSoundName?: string | null;
  soundEditable: boolean;
  onSoundNone: () => void;
  /** Desktop PowerPoint's built-in sounds above Other Sound (E037). */
  onSoundBuiltin: (id: PresentationBuiltinTransitionSoundId) => void;
  onSoundOther: (trigger: HTMLButtonElement | null) => void;
  labels: PresentationTransitionTimingLabels;
  duration: ReactNode;
  applyAll: ReactNode;
  /** Sound's glyph, drawn in the glyph column under Duration's. */
  soundIcon?: ReactNode;
  /** The group's grid (label sizes, the field column). */
  gridClassName?: string;
  afterFieldClassName?: string;
}) {
  const onClick = transitionAdvanceOnClick(rawAdvanceOnClick);
  const afterMs = isTransitionAdvanceAfterMs(rawAdvanceAfterMs)
    ? rawAdvanceAfterMs
    : null;
  const [draft, setDraft] = useState<{ slideModelId: string; value: string } | null>(
    null,
  );
  const skipBlurRef = useRef(false);
  const soundTriggerRef = useRef<HTMLButtonElement | null>(null);
  const builtinSoundId = builtinTransitionSoundIdForName(sound?.name);
  const soundLabel = builtinSoundId
    ? labels.builtinSound(builtinSoundId)
    : (sound?.name ?? preservedSoundName ?? labels.soundNone);
  const afterValue =
    draft?.slideModelId === slideModelId
      ? draft.value
      : formatTransitionAdvanceAfterSeconds(afterMs ?? 0);
  // PowerPoint keeps the After time greyed until its check box is on.
  const afterFieldDisabled = !advanceEditable || afterMs === null;

  const commitDraft = () => {
    const current = draft;
    setDraft(null);
    if (!current || current.slideModelId !== slideModelId) return;
    const next = parseTransitionAdvanceAfterSeconds(current.value);
    if (next === undefined || next === afterMs) return;
    onAdvanceChange(onClick, next);
  };
  const step = (direction: "increment" | "decrement") => {
    if (afterFieldDisabled) return;
    const base =
      parseTransitionAdvanceAfterSeconds(afterValue) ?? afterMs ?? 0;
    const next = Math.max(
      0,
      Math.min(
        MAX_EDITABLE_TRANSITION_ADVANCE_AFTER_MS,
        base +
          (direction === "increment"
            ? TRANSITION_ADVANCE_AFTER_STEP_MS
            : -TRANSITION_ADVANCE_AFTER_STEP_MS),
      ),
    );
    setDraft(null);
    if (next !== afterMs) onAdvanceChange(onClick, next);
  };
  const onKeyDown = (event: ReactKeyboardEvent<HTMLInputElement>) => {
    if (event.key === "Escape") {
      event.preventDefault();
      skipBlurRef.current = true;
      setDraft(null);
      event.currentTarget.blur();
    } else if (event.key === "Enter") {
      event.preventDefault();
      skipBlurRef.current = true;
      commitDraft();
      event.currentTarget.blur();
    }
  };

  // One grid for the whole group, as PowerPoint lays it out: glyph | label
  // | field for Duration and Sound (their fields share a left edge), then
  // the advance column (On Mouse Click over After [time]), then Apply To
  // All across both rows. Before, each row kept its own columns: Sound had
  // no glyph and its menu stood far from its label, and the check boxes and
  // the After field were out of line.
  const cell = (column: string, row: number): CSSProperties => ({ gridColumn: column, gridRow: row });
  return (
    <div
      data-testid="presentation-ribbon-transition-timing-controls"
      className={gridClassName}
      style={{
        display: "grid",
        gridTemplateColumns: "20px max-content minmax(78px, max-content) 10px max-content 78px max-content",
        gridTemplateRows: "26px 26px",
        columnGap: 6,
        rowGap: 6,
        alignItems: "center",
      }}
    >
      {/* Row 1: Duration (glyph, label, field) | On Mouse Click | Apply To All. */}
      <div style={{ display: "contents" }}>{duration}</div>
      <div style={cell("5 / 7", 1)}>
        <OfficeHomeCheckbox
          checked={onClick}
          controlId="PPTCTL.RIBBON.TRANSITIONS.ADVANCE.ON_CLICK"
          disabled={!advanceEditable}
          label={labels.advanceOnClick}
          onCheckedChange={(checked) => onAdvanceChange(checked, afterMs)}
          testId="presentation-ribbon-transition-advance-on-click"
        />
      </div>
      <div style={{ gridColumn: 7, gridRow: "1 / span 2", alignSelf: "start", paddingInlineStart: 6 }}>{applyAll}</div>
      {/* Row 2: Sound (glyph, label, menu) | After [time]. */}
      <span style={{ ...cell("1", 2), display: "inline-flex", justifyContent: "center" }}>{soundIcon}</span>
      <span style={cell("2", 2)}>{/[:：]$/.test(labels.sound) ? labels.sound : `${labels.sound}:`}</span>
      <div style={{ ...cell("3", 2), justifySelf: "start" }}>
          <DropdownMenu>
            <DropdownMenuTrigger asChild>
              <OfficeHomeMenuTrigger
                ref={soundTriggerRef}
                data-control-id="PPTCTL.RIBBON.TRANSITIONS.SOUND"
                disabled={!soundEditable}
                displayLabel={soundLabel}
                hasMenuIndicator
                icon={null}
                label={`${labels.sound}: ${soundLabel}`}
                testId="presentation-ribbon-transition-sound"
                variant="compact"
              />
            </DropdownMenuTrigger>
            <RibbonDropdownMenuContent
              align="start"
              data-testid="presentation-ribbon-transition-sound-menu"
            >
              <DropdownMenuRadioGroup
                value={builtinSoundId ? `builtin:${builtinSoundId}` : sound ? "file" : "none"}
              >
                <DropdownMenuRadioItem
                  value="none"
                  onSelect={() => {
                    if (sound) onSoundNone();
                  }}
                  data-testid="presentation-ribbon-transition-sound-none"
                >
                  {labels.soundNone}
                </DropdownMenuRadioItem>
                {PRESENTATION_BUILTIN_TRANSITION_SOUNDS.map((entry) => (
                  <DropdownMenuRadioItem
                    key={entry.id}
                    value={`builtin:${entry.id}`}
                    onSelect={() => {
                      if (builtinSoundId !== entry.id) onSoundBuiltin(entry.id);
                    }}
                    data-testid={`presentation-ribbon-transition-sound-builtin-${entry.id}`}
                  >
                    {labels.builtinSound(entry.id)}
                  </DropdownMenuRadioItem>
                ))}
                {sound && !builtinSoundId ? (
                  <DropdownMenuRadioItem
                    value="file"
                    data-testid="presentation-ribbon-transition-sound-current"
                  >
                    {sound.name}
                  </DropdownMenuRadioItem>
                ) : null}
              </DropdownMenuRadioGroup>
              {/* Sounds beyond PowerPoint's list: kept off its own list, so
                  that list reads as PowerPoint's does. */}
              <DropdownMenuSub>
                <DropdownMenuSubTrigger data-testid="presentation-ribbon-transition-sound-more">
                  {labels.moreSounds}
                </DropdownMenuSubTrigger>
                <DropdownMenuSubContent
                  data-testid="presentation-ribbon-transition-sound-more-menu"
                  onEscapeKeyDown={(event) => event.stopPropagation()}
                  style={{ maxHeight: "min(560px, var(--radix-dropdown-menu-content-available-height, calc(100vh - 24px)))", overflowY: "auto" }}
                >
                  <DropdownMenuRadioGroup value={builtinSoundId ? `builtin:${builtinSoundId}` : ""}>
                    {PRESENTATION_EXTRA_TRANSITION_SOUNDS.map((entry) => (
                      <DropdownMenuRadioItem
                        key={entry.id}
                        value={`builtin:${entry.id}`}
                        onSelect={() => {
                          if (builtinSoundId !== entry.id) onSoundBuiltin(entry.id);
                        }}
                        data-testid={`presentation-ribbon-transition-sound-builtin-${entry.id}`}
                      >
                        {labels.builtinSound(entry.id)}
                      </DropdownMenuRadioItem>
                    ))}
                  </DropdownMenuRadioGroup>
                </DropdownMenuSubContent>
              </DropdownMenuSub>
              <DropdownMenuSeparator />
              <DropdownMenuItem
                onSelect={() => onSoundOther(soundTriggerRef.current)}
                data-testid="presentation-ribbon-transition-sound-other"
              >
                {labels.soundOther}
              </DropdownMenuItem>
            </RibbonDropdownMenuContent>
          </DropdownMenu>
      </div>
      <div style={cell("5", 2)} data-testid="presentation-ribbon-transition-advance">
          <OfficeHomeCheckbox
            checked={afterMs !== null}
            controlId="PPTCTL.RIBBON.TRANSITIONS.ADVANCE.AFTER_ENABLED"
            disabled={!advanceEditable}
            label={labels.advanceAfter}
            onCheckedChange={(checked) => onAdvanceChange(onClick, checked ? 0 : null)}
            testId="presentation-ribbon-transition-advance-after-enabled"
          />
      </div>
      <div style={cell("6", 2)}>
          <OfficeHomeNumberField
            appearance="native-spinner"
            className={afterFieldClassName}
            controlId="PPTCTL.RIBBON.TRANSITIONS.ADVANCE.AFTER_TIME"
            label={labels.advanceAfterSeconds}
            min={0}
            max={MAX_EDITABLE_TRANSITION_ADVANCE_AFTER_MS / 1000}
            step={TRANSITION_ADVANCE_AFTER_STEP_SECONDS}
            inputMode="decimal"
            value={afterValue}
            disabled={afterFieldDisabled}
            onFocus={() => setDraft({ slideModelId, value: afterValue })}
            onChange={(event) => setDraft({ slideModelId, value: event.target.value })}
            onBlur={() => {
              if (skipBlurRef.current) {
                skipBlurRef.current = false;
                return;
              }
              commitDraft();
            }}
            onKeyDown={onKeyDown}
            onDecrement={() => step("decrement")}
            onIncrement={() => step("increment")}
            data-unit="seconds"
            testId="presentation-ribbon-transition-advance-after-time"
          />
      </div>
    </div>
  );
}
