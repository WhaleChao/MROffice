"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { EditorLoadingState } from "@mrnote/core/web/components/ui/editor-loading-state.tsx";
import { Check } from "@mrnote/core/web/components/office-ribbon/icons/houseIcons.tsx";
import { useTranslations } from "@mrnote/core/host/next-intl.tsx";

import gallery from "@mrnote/core/web/components/office-ribbon/OfficeInsertGallery.module.css";
import {
  OfficeTaskPane,
  officeTaskPaneParts,
  type OfficeTaskPaneTab,
} from "@mrnote/core/web/components/office-ribbon/OfficeTaskPane.tsx";
import type { DesignerEmptyReason } from "../presentation-designer-engine";
import type { PresentationDesignerPreference } from "../presentation-designer-preferences";
import type { PresentationDesignerIdea } from "./usePresentationDesigner";

import styles from "./PresentationDesignerPane.module.css";
import { PRESENTATION_TASK_PANE_TITLE_IDS } from "./presentationTaskPaneDock";

/**
 * PowerPoint's Designer pane (Design > Designer): a column of idea cards at
 * the right of the window. Each card draws the slide as the idea leaves it;
 * clicking one changes the slide and the pane stays, so another can be
 * tried; the idea on the slide is marked. When there is no idea the pane
 * says why, as PowerPoint's does. Its two options (File > Options >
 * PowerPoint Designer in PowerPoint) sit at the foot of the pane.
 */

export interface PresentationDesignerPaneProps {
  readonly ideas: readonly PresentationDesignerIdea[];
  readonly busy: boolean;
  readonly reason: DesignerEmptyReason | null;
  readonly appliedId: string | null;
  readonly renderPreview: (idea: PresentationDesignerIdea) => React.ReactNode;
  readonly onApply: (idea: PresentationDesignerIdea) => void;
  readonly onClose: () => void;
  readonly preferences: Readonly<Record<PresentationDesignerPreference, boolean>>;
  readonly onPreferenceChange: (key: PresentationDesignerPreference, enabled: boolean) => void;
  readonly dock?: {
    readonly tabs?: readonly OfficeTaskPaneTab[];
    readonly tabsLabel?: string;
    readonly hidden?: boolean;
    readonly className?: string;
  };
}

export default function PresentationDesignerPane({
  ideas,
  busy,
  reason,
  appliedId,
  renderPreview,
  onApply,
  onClose,
  preferences,
  onPreferenceChange,
  dock,
}: PresentationDesignerPaneProps): React.JSX.Element {
  const t = useTranslations("console-notebooks");
  const graphic = (idea: PresentationDesignerIdea) =>
    idea.kind === "smartart" || idea.kind === "timeline" || idea.kind === "icons";
  const name = (idea: PresentationDesignerIdea) =>
    idea.kind === "theme"
      ? (idea.layoutName ?? "")
      : idea.kind === "layout"
      ? idea.layoutNameIsKey
        ? t(`presentation.layout.${idea.layoutName}` as "presentation.layout.title")
        : (idea.layoutName ?? "")
      : graphic(idea)
        ? t(`toolbar.smartArtLayoutName_${idea.layoutName}` as "toolbar.smartArtLayoutName_basic-list")
        : t(`presentation.designIdeas.name.${idea.kind}` as "presentation.designIdeas.name.title-band");
  const description = (idea: PresentationDesignerIdea) =>
    idea.kind === "layout" || idea.kind === "theme" || graphic(idea)
      ? t(`presentation.designIdeas.desc.${idea.kind}` as "presentation.designIdeas.desc.layout", { name: name(idea) })
      : t(`presentation.designIdeas.desc.${idea.kind}` as "presentation.designIdeas.desc.title-band");
  return (
    null
  );
}
