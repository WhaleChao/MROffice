"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The Review tab of the presentation Ribbon (Proofing, Accessibility, Activity,
 * Comments, Ink groups), drawn from the window bindings it reads.
 *
 * Extracted verbatim from the render of
 * `components/notebook/contents/PresentationWindow.tsx`: the JSX is unchanged;
 * the window passes the bindings it used to close over as `ctl`, and the tab
 * condition and key stay at the call site. No hooks run here, so the window's
 * hook order is unchanged.
 */

import {
  OfficeHomeRibbonGroup,
  OfficeHomeCommandRow,
  OfficeHomeSplitLayout,
  OfficeHomeCommand,
  OfficeHomeMenuTrigger,
} from "@mrnote/core/web/components/office-ribbon/OfficeHomeRibbon.tsx";
import presentationChrome from "@mrnote/core/web/lib/presentation/presentation-window-shell.module.css";
import {
  DropdownMenu,
  DropdownMenuTrigger,
  DropdownMenuItem,
  DropdownMenuSeparator,
} from "@mrnote/core/web/components/ui/dropdown-menu.tsx";
import { OfficeSemanticIcon } from "@mrnote/core/web/components/office-ribbon/icons/OfficeSemanticIcon.tsx";
import { RibbonDropdownMenuContent } from "@mrnote/core/web/components/office-ribbon/CupertinoRibbonMenu.tsx";
import { Settings, BookOpen } from "@mrnote/core/web/components/office-ribbon/icons/houseIcons.tsx";
import { IconReviewActivity, IconReviewComment } from "@mrnote/core/web/components/notebook/contents/PresentationRibbonIcons.tsx";
import { canModifyCommentAuthor, type SlideComment } from "@mrnote/core/web/lib/presentation-comments.ts";
import { hasDeletableInk, type InkDeleteScope } from "../../../../lib/presentation/ink-delete";
import type { useTranslations } from "@mrnote/core/host/next-intl.tsx";
import type { Dispatch, RefObject, SetStateAction } from "react";
import type { SlideModel } from "@mrnote/core/web/lib/presentation/deck-model.ts";
import type { embeddedPresentationAutosave } from "../../../../lib/presentation-embedded-session";
import type { PresentationAccessibilityCheckCommandAuthority } from "@mrnote/core/web/components/notebook/contents/ribbon/commandRegistry.ts";

/** The window bindings this tab reads, passed through unchanged. */
export interface PresentationReviewRibbonTabCtl {
  t: ReturnType<typeof useTranslations<"console-notebooks">>;
  reviewAccessibilityTriggerRef: RefObject<HTMLButtonElement | null>;
  accessibilityPanelOpen: boolean;
  proofingLanguageReturnFocusRef: RefObject<HTMLButtonElement | null>;
  reviewInkHidden: boolean;
  setReviewInkHidden: Dispatch<SetStateAction<boolean>>;
  slides: SlideModel[];
  activeIndex: number;
  activeIndexRef: RefObject<number>;
  autosave: ReturnType<typeof embeddedPresentationAutosave>;
  hasPersistedReviewInk: boolean;
  deleteReviewInk: (scope: InkDeleteScope) => void;
  commentsPanelOpen: boolean;
  showProofingErrors: boolean;
  setShowProofingErrors: Dispatch<SetStateAction<boolean>>;
  setSpellingStartIndex: Dispatch<SetStateAction<number | null>>;
  showReviewMarkup: boolean;
  setShowReviewMarkup: Dispatch<SetStateAction<boolean>>;
  showAdvancedReviewMarkup: boolean;
  setShowAdvancedReviewMarkup: Dispatch<SetStateAction<boolean>>;
  currentCommentActorId: string | null;
  canAuthorReviewComment: boolean;
  /** Why comments cannot be written now (presentationCommentAuthoringBlock). */
  commentAuthoringBlockedReason: string;
  canOpenReviewCommentComposer: boolean;
  reviewCommentEntries: { commentId: string; comment: SlideComment; slideId: string; slideIndex: number; }[];
  selectedReviewCommentEntry: { commentId: string; comment: SlideComment; slideId: string; slideIndex: number; } | null;
  openReviewCommentComposer: () => void;
  toggleReviewCommentsPane: () => void;
  navigateReviewComment: (direction: -1 | 1) => void;
  deleteSelectedReviewComment: () => void;
  canDeleteAllReviewCommentsOnSlide: boolean;
  deleteAllReviewCommentsOnSlide: () => void;
  /** Review > Thesaurus: the shared pane on the selected word (E045). */
  openThesaurus: () => void;
  /** Review > Show Changes (E045): slides others changed since your last visit. */
  deckChangeCount: number;
  deckChangeAuthors: readonly string[];
  showChanges: boolean;
  toggleShowChanges: () => void;
  markAllChangesRead: () => void;
  accessibilityCheckCommandAuthority: PresentationAccessibilityCheckCommandAuthority;
  invokeAccessibilityCheck: (trigger: HTMLElement | null) => void;
}

export function PresentationReviewRibbonTab({ ctl }: { ctl: PresentationReviewRibbonTabCtl }) {
  const {
    t,
    reviewAccessibilityTriggerRef,
    accessibilityPanelOpen,
    proofingLanguageReturnFocusRef,
    reviewInkHidden,
    setReviewInkHidden,
    slides,
    activeIndex,
    activeIndexRef,
    autosave,
    hasPersistedReviewInk,
    deleteReviewInk,
    commentsPanelOpen,
    showProofingErrors,
    setShowProofingErrors,
    setSpellingStartIndex,
    showReviewMarkup,
    setShowReviewMarkup,
    showAdvancedReviewMarkup,
    setShowAdvancedReviewMarkup,
    currentCommentActorId,
    canAuthorReviewComment,
    commentAuthoringBlockedReason,
    canOpenReviewCommentComposer,
    reviewCommentEntries,
    selectedReviewCommentEntry,
    openReviewCommentComposer,
    toggleReviewCommentsPane,
    navigateReviewComment,
    deleteSelectedReviewComment,
    canDeleteAllReviewCommentsOnSlide,
    deleteAllReviewCommentsOnSlide,
    openThesaurus,
    deckChangeCount,
    deckChangeAuthors,
    showChanges,
    toggleShowChanges,
    markAllChangesRead,
    accessibilityCheckCommandAuthority,
    invokeAccessibilityCheck,
  } = ctl;
  return (
                  <>
                    {/* PPT-A098: desktop Review ▸ Proofing is Spelling and
                        Thesaurus. Spelling checks the deck with the shared
                        lib/proofing checker (the spreadsheet's) and edits
                        through Find's Replace; its arrow keeps the proofing
                        marks toggle and Editor Settings. Thesaurus opens the
                        shared Thesaurus pane on the selected word (E045). */}
                    <OfficeHomeRibbonGroup
                      className={presentationChrome.reviewProofingGroup}
                      label={t("presentation.review.proofingGroup")}
                      ribbon="review"
                      testId="presentation-ribbon-review-proofing-group"
                    >
                      <OfficeHomeCommandRow>
                        <DropdownMenu>
                          <OfficeHomeSplitLayout
                            className={`${presentationChrome.insertLargeSplit} ${presentationChrome.reviewSplit}`}
                            testId="presentation-ribbon-review-spelling-split"
                            primary={
                              <OfficeHomeCommand
                                variant="large"
                                controlId="PPTCTL.RIBBON.REVIEW.PROOFING.SPELLING"
                                disabled={autosave.readOnly || slides.length === 0}
                                icon={
                                  <OfficeSemanticIcon
                                    semanticId="review.spellcheck"
                                    size={24}
                                  />
                                }
                                label={t("presentation.review.spelling")}
                                onClick={() =>
                                  setSpellingStartIndex(activeIndexRef.current)
                                }
                                testId="presentation-ribbon-review-spelling"
                              />
                            }
                            trigger={
                              <DropdownMenuTrigger asChild>
                                <OfficeHomeMenuTrigger
                                  ref={proofingLanguageReturnFocusRef}
                                  variant="icon"
                                  icon={<OfficeSemanticIcon semanticId="menu.chevron.down" size={12} />}
                                  label={t("presentation.review.proofingGroup")}
                                  testId="presentation-ribbon-review-proofing-menu"
                                />
                              </DropdownMenuTrigger>
                            }
                          />
                          <RibbonDropdownMenuContent align="start">
                            <DropdownMenuItem
                              onSelect={() => setShowProofingErrors((shown) => !shown)}
                              data-testid="presentation-ribbon-review-proofing-errors"
                            >
                              <OfficeSemanticIcon semanticId="review.spellcheck" size={20} />
                              {t(showProofingErrors ? "presentation.review.hideProofingErrors" : "presentation.review.showProofingErrors")}
                            </DropdownMenuItem>
                            {/* Editor Settings: the dialog that configures
                                spelling and grammar rules is not built here,
                                and the reason says so. */}
                            <DropdownMenuItem
                              disabled
                              title={t(
                                "presentation.review.editorSettingsUnavailable",
                              )}
                              data-testid="presentation-ribbon-review-editor-settings"
                            >
                              <Settings size={20} />
                              {t("presentation.review.editorSettings")}
                            </DropdownMenuItem>
                          </RibbonDropdownMenuContent>
                        </DropdownMenu>
                        {null}
                      </OfficeHomeCommandRow>
                    </OfficeHomeRibbonGroup>

                    <OfficeHomeRibbonGroup
                      className={presentationChrome.reviewAccessibilityGroup}
                      label={t("presentation.review.accessibilityGroup")}
                      ribbon="review"
                      testId="presentation-ribbon-review-accessibility-group"
                    >
                      <OfficeHomeCommand
                        variant="large"
                        active={accessibilityPanelOpen}
                        buttonRef={reviewAccessibilityTriggerRef}
                        controlId="PPTCTL.RIBBON.REVIEW.A11Y.CHECK"
                        disabled={
                          !accessibilityCheckCommandAuthority.isEnabled()
                        }
                        icon={
                          <OfficeSemanticIcon
                            semanticId="review.accessibility"
                            size={24}
                          />
                        }
                        label={t("toolbar.accessibilityCheck")}
                        onClick={() =>
                          invokeAccessibilityCheck(
                            reviewAccessibilityTriggerRef.current,
                          )
                        }
                        testId="presentation-ribbon-accessibility-check"
                      />
                    </OfficeHomeRibbonGroup>

                    <OfficeHomeRibbonGroup
                      className={presentationChrome.reviewActivityGroup}
                      label={t("presentation.review.activityGroup")}
                      ribbon="review"
                      testId="presentation-ribbon-review-activity-group"
                    >
                      <OfficeHomeCommandRow>
                        <OfficeHomeCommand
                          variant="large"
                          controlId="PPTCTL.RIBBON.REVIEW.CHANGES.MARK_ALL_READ"
                          icon={
                            <IconReviewActivity kind="read" />
                          }
                          label={t("presentation.review.markAllRead")}
                          displayLabel={t("presentation.review.markAllRead")}
                          onClick={() => void import("@mrnote/core/host/desktop-app.ts").then((module) => module.noteDesktopFeature("Mark All as Read"))}
                          testId="presentation-ribbon-review-mark-all-read"
                        />
                        <OfficeHomeCommand
                          variant="large"
                          controlId="PPTCTL.RIBBON.REVIEW.CHANGES.SHOW"
                          active={showChanges}
                          icon={
                            <IconReviewActivity kind="edits" />
                          }
                          label={
                            deckChangeCount > 0
                              ? t("presentation.review.showChangesSummary", {
                                  count: deckChangeCount,
                                  authors: deckChangeAuthors.join(", ") || t("presentation.review.othersAuthor"),
                                })
                              : t("presentation.review.showChanges")
                          }
                          displayLabel={t("presentation.review.showChanges")}
                          onClick={() => void import("@mrnote/core/host/desktop-app.ts").then((module) => module.noteDesktopFeature("Show Changes"))}
                          testId="presentation-ribbon-review-show-changes"
                        />
                      </OfficeHomeCommandRow>
                    </OfficeHomeRibbonGroup>

                    <OfficeHomeRibbonGroup
                      className={presentationChrome.reviewCommentsGroup}
                      label={t("presentation.review.commentsGroup")}
                      ribbon="review"
                      testId="presentation-ribbon-review-comments-group"
                    >
                      <OfficeHomeCommandRow>
                        <OfficeHomeCommand
                          variant="large"
                          controlId="PPTCTL.RIBBON.REVIEW.COMMENTS.NEW"
                          disabled={!canOpenReviewCommentComposer}
                          icon={
                            <IconReviewComment action="new" />
                          }
                          label={
                            canOpenReviewCommentComposer
                              ? t("presentation.review.newComment")
                              : `${t("presentation.review.newComment")}: ${commentAuthoringBlockedReason}`
                          }
                          displayLabel={t("presentation.review.newComment")}
                          onClick={openReviewCommentComposer}
                          testId="presentation-ribbon-review-new-comment"
                        />
                          <DropdownMenu>
                            <OfficeHomeSplitLayout
                              className={`${presentationChrome.insertLargeSplit} ${presentationChrome.reviewSplit}`}
                              testId="presentation-ribbon-review-delete-split"
                              primary={
                                <OfficeHomeCommand
                                  variant="large"
                                  controlId="PPTCTL.RIBBON.REVIEW.COMMENTS.DELETE"
                                  icon={<IconReviewComment action="delete" />}
                                  label={t("presentation.review.deleteComment")}
                                  disabledReason={
                                    // What stands in the way first: nothing
                                    // to delete, then why comments are
                                    // read-only, then whose comment it is.
                                    !selectedReviewCommentEntry || selectedReviewCommentEntry.slideIndex !== activeIndex
                                      ? t("presentation.comments.empty")
                                      : !canAuthorReviewComment
                                        ? commentAuthoringBlockedReason
                                        : t("presentation.comments.onlyAuthorCanEdit")
                                  }
                                  disabled={
                                    !selectedReviewCommentEntry ||
                                    selectedReviewCommentEntry.slideIndex !== activeIndex ||
                                    !canAuthorReviewComment ||
                                    !canModifyCommentAuthor(selectedReviewCommentEntry.comment.authorId, currentCommentActorId)
                                  }
                                  onClick={deleteSelectedReviewComment}
                                  testId="presentation-ribbon-review-delete-primary"
                                />
                              }
                              trigger={
                            <DropdownMenuTrigger asChild>
                              <OfficeHomeMenuTrigger
                                variant="icon"
                                data-control-id="PPTCTL.RIBBON.REVIEW.COMMENTS.DELETE.MENU"
                                disabled={
                                  !canAuthorReviewComment ||
                                  (!selectedReviewCommentEntry &&
                                    !canDeleteAllReviewCommentsOnSlide)
                                }
                                icon={<OfficeSemanticIcon semanticId="menu.chevron.down" size={12} />}
                                label={
                                  !selectedReviewCommentEntry &&
                                  !canDeleteAllReviewCommentsOnSlide
                                    ? `${t(
                                        "presentation.review.deleteComment",
                                      )}: ${t(
                                        "presentation.comments.empty",
                                      )}`
                                    : !canAuthorReviewComment
                                      ? `${t(
                                          "presentation.review.deleteComment",
                                        )}: ${commentAuthoringBlockedReason}`
                                      : t(
                                          "presentation.review.deleteComment",
                                        )
                                }
                                displayLabel={t(
                                  "presentation.review.deleteComment",
                                )}
                                testId="presentation-ribbon-review-delete-comment"
                              />
                            </DropdownMenuTrigger>
                              }
                            />
                            <RibbonDropdownMenuContent align="start">
                              <DropdownMenuItem
                                disabled={
                                  !selectedReviewCommentEntry ||
                                  selectedReviewCommentEntry.slideIndex !==
                                    activeIndex ||
                                  !canAuthorReviewComment ||
                                  !canModifyCommentAuthor(
                                    selectedReviewCommentEntry.comment.authorId,
                                    currentCommentActorId,
                                  )
                                }
                                onSelect={deleteSelectedReviewComment}
                                data-testid="presentation-ribbon-review-delete-selected-comment"
                              >
                                {t(
                                  "presentation.review.deleteSelectedComment",
                                )}
                              </DropdownMenuItem>
                              <DropdownMenuItem
                                disabled={!canDeleteAllReviewCommentsOnSlide}
                                onSelect={deleteAllReviewCommentsOnSlide}
                                data-testid="presentation-ribbon-review-delete-slide-comments"
                              >
                                {t(
                                  "presentation.review.deleteSlideComments",
                                )}
                              </DropdownMenuItem>
                            </RibbonDropdownMenuContent>
                          </DropdownMenu>
                          <OfficeHomeCommand
                            variant="large"
                            controlId="PPTCTL.RIBBON.REVIEW.COMMENTS.PREVIOUS"
                            disabled={reviewCommentEntries.length === 0}
                            icon={<IconReviewComment action="previous" />}
                            label={
                              reviewCommentEntries.length === 0
                                ? `${t(
                                    "presentation.review.previousComment",
                                  )}: ${t("presentation.comments.empty")}`
                                : t("presentation.review.previousComment")
                            }
                            displayLabel={t(
                              "presentation.review.previousShort",
                            )}
                            onClick={() => navigateReviewComment(-1)}
                            testId="presentation-ribbon-review-previous-comment"
                          />
                          <OfficeHomeCommand
                            variant="large"
                            controlId="PPTCTL.RIBBON.REVIEW.COMMENTS.NEXT"
                            disabled={reviewCommentEntries.length === 0}
                            icon={<IconReviewComment action="next" />}
                            label={
                              reviewCommentEntries.length === 0
                                ? `${t(
                                    "presentation.review.nextComment",
                                  )}: ${t("presentation.comments.empty")}`
                                : t("presentation.review.nextComment")
                            }
                            displayLabel={t(
                              "presentation.review.nextShort",
                            )}
                            onClick={() => navigateReviewComment(1)}
                            testId="presentation-ribbon-review-next-comment"
                          />
                        <DropdownMenu>
                          <OfficeHomeSplitLayout
                            className={`${presentationChrome.insertLargeSplit} ${presentationChrome.reviewSplit}`}
                            testId="presentation-ribbon-review-comments-split"
                            primary={
                              <OfficeHomeCommand
                                variant="large"
                                active={commentsPanelOpen}
                                icon={<IconReviewComment action="show" />}
                                label={t("presentation.review.showComments")}
                                onClick={toggleReviewCommentsPane}
                                testId="presentation-ribbon-comments"
                              />
                            }
                            trigger={
                          <DropdownMenuTrigger asChild>
                            <OfficeHomeMenuTrigger
                              variant="icon"
                              icon={<OfficeSemanticIcon semanticId="menu.chevron.down" size={12} />}
                              label={t("presentation.review.showComments")}
                              testId="presentation-ribbon-review-comments-menu"
                            />
                          </DropdownMenuTrigger>
                            }
                          />
                          <RibbonDropdownMenuContent align="start">
                            <DropdownMenuItem
                              role="menuitemcheckbox"
                              aria-checked={commentsPanelOpen}
                              onSelect={toggleReviewCommentsPane}
                              data-testid="presentation-ribbon-review-comments-pane"
                            >
                              <IconReviewComment action="show" size={20} />
                              {t("presentation.review.commentsPane")}
                            </DropdownMenuItem>
                            <DropdownMenuItem
                              role="menuitemcheckbox"
                              aria-checked={showReviewMarkup}
                              onSelect={() => setShowReviewMarkup((shown) => !shown)}
                              data-testid="presentation-ribbon-review-show-markup"
                            >
                              <IconReviewActivity kind="edits" size={20} />
                              {t("presentation.review.showMarkup")}
                            </DropdownMenuItem>
                            <DropdownMenuItem
                              role="menuitemcheckbox"
                              aria-checked={showAdvancedReviewMarkup}
                              onSelect={() => {
                                setShowAdvancedReviewMarkup((shown) => !shown);
                                setShowReviewMarkup(true);
                              }}
                              data-testid="presentation-ribbon-review-show-advanced-markup"
                            >
                              <IconReviewComment action="show" size={20} />
                              {t("presentation.review.showAdvancedMarkup")}
                            </DropdownMenuItem>
                          </RibbonDropdownMenuContent>
                        </DropdownMenu>
                      </OfficeHomeCommandRow>
                    </OfficeHomeRibbonGroup>

                    {/* PPT-A102: desktop Review keeps Hide Ink ▾ on the tab
                        (greyed without ink): Hide Ink, Delete All Ink on
                        Slide, Delete All Ink in Presentation. */}
                    <OfficeHomeRibbonGroup
                      label={t("presentation.review.inkGroup")}
                      ribbon="review"
                      testId="presentation-ribbon-review-ink-group"
                    >
                      <DropdownMenu>
                        <OfficeHomeSplitLayout
                          className={`${presentationChrome.insertLargeSplit} ${presentationChrome.reviewSplit}`}
                          testId="presentation-ribbon-review-hide-ink-split"
                          primary={
                            <OfficeHomeCommand
                              active={hasPersistedReviewInk && reviewInkHidden}
                              controlId="PPTCTL.RIBBON.REVIEW.INK.HIDE"
                              disabled={!hasPersistedReviewInk}
                              disabledReason={t("presentation.review.noInkUnavailable")}
                              displayLabel={t("presentation.review.hideInk")}
                              icon={
                                <OfficeSemanticIcon
                                  semanticId="draw.pen"
                                  size={24}
                                />
                              }
                              label={t("presentation.review.hideInk")}
                              onClick={() =>
                                setReviewInkHidden((hidden) => !hidden)
                              }
                              testId="presentation-ribbon-review-hide-ink"
                              variant="large"
                            />
                          }
                          trigger={
                            <DropdownMenuTrigger asChild>
                              <OfficeHomeMenuTrigger
                                variant="icon"
                                disabled={!hasPersistedReviewInk}
                                icon={<OfficeSemanticIcon semanticId="menu.chevron.down" size={12} />}
                                label={t("presentation.review.hideInk")}
                                testId="presentation-ribbon-review-hide-ink-menu"
                              />
                            </DropdownMenuTrigger>
                          }
                        />
                        <RibbonDropdownMenuContent align="start">
                          <DropdownMenuItem
                            role="menuitemcheckbox"
                            aria-checked={reviewInkHidden}
                            onSelect={() => setReviewInkHidden((hidden) => !hidden)}
                            data-testid="presentation-ribbon-review-hide-ink-toggle"
                          >
                            {t("presentation.review.hideInk")}
                          </DropdownMenuItem>
                          <DropdownMenuSeparator />
                          <DropdownMenuItem
                            disabled={
                              autosave.readOnly ||
                              !hasDeletableInk(slides, { kind: "slide", index: activeIndex })
                            }
                            onSelect={() => deleteReviewInk({ kind: "slide", index: activeIndexRef.current })}
                            data-testid="presentation-ribbon-review-delete-ink-slide"
                          >
                            {t("presentation.review.deleteInkOnSlide")}
                          </DropdownMenuItem>
                          <DropdownMenuItem
                            disabled={
                              autosave.readOnly ||
                              !hasDeletableInk(slides, { kind: "presentation" })
                            }
                            onSelect={() => deleteReviewInk({ kind: "presentation" })}
                            data-testid="presentation-ribbon-review-delete-ink-presentation"
                          >
                            {t("presentation.review.deleteInkInPresentation")}
                          </DropdownMenuItem>
                        </RibbonDropdownMenuContent>
                      </DropdownMenu>
                    </OfficeHomeRibbonGroup>

                  </>
  );
}
