/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
// The public types of the npm package (scripts/npm/pack.mjs checks that they
// name every export of the built module).
import type { ReactElement } from "react";
import type { EditorDefinition, EditorProps } from "mroffice-core";

/** The document JSON of a new handwritten note. */
export declare function newHandwritingContent(): Record<string, unknown>;
export declare function HandwritingEditor(props: EditorProps): ReactElement;
export declare const inkEditor: EditorDefinition;
export default inkEditor;
