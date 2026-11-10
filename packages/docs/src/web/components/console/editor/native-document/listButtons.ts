/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { ChainedCommands } from "@tiptap/core";
import type { EditorState, Transaction } from "@tiptap/pm/state";
import { liftTarget } from "@tiptap/pm/transform";
import type { BulletStyleType } from "@mrnote/core/web/components/console/editor/native-document/bulletListStyle.ts";
import type { DocumentListDialogResult } from "./DocumentListDefinitionDialogs";
import { headingListAtSelection, selectedHeadings } from "@mrnote/core/web/components/console/editor/native-document/listDefinitionMarkers.ts";
import {
  normalizeDocumentListStyles,
  presetListDefinition,
  withNumberFormat,
  type ListDefinition,
  type ListLevelFormat,
} from "@mrnote/core/web/components/console/editor/native-document/listLevelDefinition.ts";
import type { ListStyleType } from "@mrnote/core/web/components/console/editor/native-document/listNumbering.ts";
import {
  stripTypedListMarkers,
  type TypedListMarkerKind,
  type TypedListMarkerPlan,
} from "./typedListMarkers";

/**
 * The blocks a list cannot hold: a floating shape, a media frame and a
 * signature line stand beside the paragraphs, anchored to one (Word applies
 * Bullets to the paragraphs, never to the drawing), and the page's save
 * refuses each inside a list ("documentShape nodes are not allowed inside
 * lists") -- for the whole page. A selection dragged across one and given
 * Bullets or Numbering wrapped it with the paragraphs.
 */
const UNLISTABLE_BLOCKS = new Set(["documentShape", "documentMedia", "documentSignatureLine"]);

/** Lift every unlistable block a list command wrapped back out of its list. */
export function liftUnlistableBlocks(tr: Transaction): void {
  for (let guard = 0; guard < 256; guard += 1) {
    let lifted = false;
    tr.doc.descendants((node, pos) => {
      if (lifted) return false;
      if (!UNLISTABLE_BLOCKS.has(node.type.name)) return true;
      const $from = tr.doc.resolve(pos);
      let item = -1;
      for (let depth = $from.depth; depth > 0 && item < 0; depth -= 1) {
        const name = $from.node(depth).type.name;
        if (name === "listItem" || name === "taskItem") item = depth;
      }
      if (item < 0) return false;
      // The block and everything after it in its list item leave together: a
      // list item starts with a paragraph, so the blocks after the shape
      // cannot stay behind as an item of their own.
      const range = $from.blockRange(tr.doc.resolve($from.end(item)));
      const target = range ? liftTarget(range) : null;
      if (range && typeof target === "number") {
        tr.lift(range, target);
        lifted = true;
      }
      return false;
    });
    if (!lifted) return;
  }
}

/** Run a list command's chain, with any unlistable block it took lifted out. */
function runListChain(chain: ChainedCommands): boolean {
  return chain
    .command(({ tr }) => {
      liftUnlistableBlocks(tr);
      return true;
    })
    .run();
}

/** The list-definition number format of each typed marker kind. */
const TYPED_MARKER_LEVEL_FORMAT = {
  decimal: "decimal",
  "lower-alpha": "lowerLetter",
  "upper-alpha": "upperLetter",
  "lower-roman": "lowerRoman",
  "upper-roman": "upperRoman",
} as const satisfies Record<Exclude<TypedListMarkerKind, "bullet">, ListLevelFormat>;

/**
 * Home > Bullets on the chain's selection. Word's Numbering and Bullets on
 * paragraphs that begin with a typed marker ("1.测试", "a.", "一、", "• "):
 * the typed marker goes, in the same Undo step as the list. Pressed again
 * (the selection already in a bullet list) it takes the list off.
 */
export function runBulletsButton(chain: ChainedCommands, inBulletList: boolean): boolean {
  if (!inBulletList) {
    chain.command(({ tr }) => {
      stripTypedListMarkers(tr);
      return true;
    });
  }
  return runListChain(chain.toggleBulletList());
}

/** Home > Checklist on the chain's selection (on, or off again). */
export function runChecklistButton(chain: ChainedCommands): boolean {
  return runListChain(chain.toggleTaskList());
}

/**
 * Home > Numbering on the chain's selection: typed markers go and the list
 * numbers the paragraphs in the markers' own format ("a." stays lower-alpha,
 * "1、" keeps its punctuation after the number).
 */
export function runNumberingButton(chain: ChainedCommands, inOrderedList: boolean): boolean {
  if (inOrderedList) return runListChain(chain.toggleOrderedList());
  let plan: TypedListMarkerPlan | null = null;
  chain
    .command(({ tr }) => {
      plan = stripTypedListMarkers(tr);
      return true;
    })
    .toggleOrderedList()
    .command(({ commands }) => {
      const typed = plan as TypedListMarkerPlan | null;
      if (!typed || typed.kind === "bullet") return true;
      const style = typed.kind;
      if (typed.suffix === "." || typed.suffix === null) {
        return style === "decimal" ? true : commands.setOrderedListStyle(style);
      }
      // "1、" and "1．": the typed punctuation stays after the number.
      const definition = withNumberFormat(presetListDefinition(style), 0, {
        format: TYPED_MARKER_LEVEL_FORMAT[style],
        prefix: "",
        suffix: typed.suffix,
        align: "left",
      });
      return definition ? commands.setOrderedListDefinition(definition) : true;
    });
  return runListChain(chain);
}

/**
 * A Bullet Library tile: the selection becomes a bullet list (typed markers
 * go, as with the Bullets button) and takes the tile's glyph.
 */
export function runBulletLibraryButton(
  chain: ChainedCommands,
  inBulletList: boolean,
  style: BulletStyleType,
): boolean {
  // A library entry replaces typed "1." / "•" markers (Word).
  if (!inBulletList) {
    chain.command(({ tr }) => {
      stripTypedListMarkers(tr);
      return true;
    });
    chain.toggleBulletList();
  }
  return runListChain(chain.setBulletListStyle(style));
}

/**
 * A Numbering Library or Multilevel List tile: the selection becomes a
 * numbered list in the tile's scheme, or takes a document list style
 * (Define New List Style) when `listStyle` is given.
 */
export function runNumberingLibraryButton(
  chain: ChainedCommands,
  inOrderedList: boolean,
  scheme:
    | { style: ListStyleType }
    | { listStyle: { id: string; definition: ListDefinition } },
): boolean {
  if (!inOrderedList) {
    chain.command(({ tr }) => {
      stripTypedListMarkers(tr);
      return true;
    });
    chain.toggleOrderedList();
  }
  if ("style" in scheme) return runListChain(chain.setOrderedListStyle(scheme.style));
  return runListChain(
    chain.setOrderedListDefinition(scheme.listStyle.definition, { styleId: scheme.listStyle.id }),
  );
}

/**
 * OK of the list dialogs (Define New Bullet / Number Format / Multilevel List
 * / List Style, Set Numbering Value) in one Undo step. Word's list commands
 * turn the selection into a list first; on headings they keep the heading
 * and number it.
 */
export function runListDialogResult(
  chain: ChainedCommands,
  state: EditorState,
  result: DocumentListDialogResult,
  inList: { bulletList: boolean; orderedList: boolean },
  newListStyleId: () => string = () => `ls${Date.now().toString(36)}`,
): boolean {
  const headings = selectedHeadings(state);
  const listedHeading = headingListAtSelection(state);
  const inBullet = !headings && inList.bulletList;
  const inOrdered = headings
    ? listedHeading?.node.attrs.listKind === "ordered"
    : inList.orderedList;
  if (result.kind === "defineBullet") {
    if (headings ? listedHeading?.node.attrs.listKind !== "bullet" : !inBullet) chain.toggleBulletList();
    return result.galleryStyle && result.align === "left"
      ? runListChain(chain.setBulletListStyle(result.galleryStyle))
      : runListChain(chain.setBulletListGlyph(result.glyph, result.font, result.align));
  }
  if (result.kind === "defineNumber") {
    if (!inOrdered) chain.toggleOrderedList();
    chain.setOrderedListDefinition(result.definition);
    // The dialog's Start at is written for headings too, as Set
    
    if (headings) chain.setListNumberingValue("new", result.start);
    else chain.setOrderedListStart(result.start);
    return runListChain(chain);
  }
  if (result.kind === "defineMultilevel") {
    if (!inOrdered) chain.toggleOrderedList();
    return runListChain(chain.setOrderedListDefinition(result.definition));
  }
  if (result.kind === "defineListStyle") {
    const existing = normalizeDocumentListStyles(state.doc.attrs.listStyles);
    const id = newListStyleId();
    chain.setDocumentListStyles([
      ...existing.filter((style) => style.name !== result.name),
      { id, name: result.name, definition: result.definition },
    ]);
    if (!inOrdered) chain.toggleOrderedList();
    return runListChain(chain.setOrderedListDefinition(result.definition, { styleId: id }));
  }
  return runListChain(chain.setListNumberingValue(result.mode, result.value));
}
