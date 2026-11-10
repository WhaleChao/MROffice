/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A causal save names its body instead of sending it.
 *
 * The update a save carries already says the whole body: the server projects
 * it and stores that projection. The body was sent beside it only so the
 * server could compare the two (2.5 MB of a long document's request, every
 * save). Both sides now digest the projection the same way
 * (python-canonical-json.ts spells the API's canonical JSON), so the request
 * carries the digest, and the receipt of a save that stored exactly that
 * body carries none either: the author rebuilds it from what it submitted.
 *
 * Only the document model holds anything beyond the shared state: its three
 * layout receipts. They ride the request as values, or as the page version at
 * which this author holds them unchanged from what is stored.
 */
import { documentCausalSemanticContent, markDocumentSaveBody } from "./document-page-causal-model";
import { jsonEquals, type JsonValue } from "./json-crdt";
import {
  compareStringsByCodePoint, pythonCanonicalJson, pythonCanonicalJsonSha256, pythonDefaultSortedJson,
} from "./python-canonical-json";

type Body = Record<string, JsonValue>;

const PRIVATE_LAYOUT = ["pageBlockSizes", "pageBlockLayoutKey", "pageRegionLayoutReceipt"] as const;

/** Page types whose causal save may name its body (the API's own list). */
export const PAGE_TYPES_NAMING_THEIR_BODY: ReadonlySet<string> = new Set(["document", "presentation", "code"]);

export interface PageCausalNamedBody {
  /** The API's canonical digest of the content this author's update projects to. */
  readonly content_sha256: string;
  readonly private_attrs?: Body;
  readonly private_attrs_base_content_version?: number;
  readonly private_attrs_delta?: PrivateAttrsDelta;
  /** The bodies the save may have stored, most likely first, each with its
   * digest when already known (see `storedBodyOfReceipt`). */
  readonly stored: readonly StoredBodyCandidate[];
  /** The layout receipts this author's body carries (none: `{}`). */
  readonly carried: Body;
}

export interface StoredBodyCandidate {
  readonly body: Body;
  readonly sha256?: string;
  /** Read from the room for this one save, and written into by no one. */
  readonly readOnly?: boolean;
}

/**
 * The bodies a checkpoint that projects to `projection` may be stored as: with
 * the layout receipts its author carried, or without (the server drops them
 * when the stored content is no longer the author's own).
 */
export function storedBodyCandidates(model: string, projection: Body, carried: Body): StoredBodyCandidate[] {
  if (model !== "document-canonical-v1") return [{ body: projection }];
  const shared = withApiMarkOrder(projection) as Body;
  return Object.keys(carried).length
    ? [{ body: { ...shared, attrs: { ...(record(shared.attrs) ? shared.attrs : {}), ...carried } } as Body, readOnly: true }, { body: shared, readOnly: true }]
    : [{ body: shared, readOnly: true }];
}

const record = (value: unknown): value is Record<string, unknown> =>
  Boolean(value && typeof value === "object" && !Array.isArray(value));

/** The layout receipts as what changed in them since the stored ones: the
 * API applies `ops` to the receipts stored at `base_content_version` and
 * takes the result only when its canonical digest is `sha256`. */
export interface PrivateAttrsDelta {
  readonly base_content_version: number;
  readonly ops: readonly (readonly [readonly (string | number)[], JsonValue] | readonly [readonly (string | number)[]])[];
  readonly sha256: string;
}

type DeltaOp = [(string | number)[], JsonValue] | [(string | number)[]];

/** `[path, value]` sets, `[path]` removes an object key; arrays are patched
 * in place only while they keep their length. False: over `budget` ops. */
function collectDelta(from: JsonValue, to: JsonValue, path: (string | number)[], ops: DeltaOp[], budget: number): boolean {
  if (from === to) return true;
  if (record(from) && record(to)) {
    for (const key of Object.keys(from)) {
      if (to[key] === undefined) { ops.push([[...path, key]]); if (ops.length > budget) return false; }
    }
    for (const key of Object.keys(to)) {
      const next = to[key] as JsonValue | undefined;
      if (next === undefined) continue;
      const previous = from[key] as JsonValue | undefined;
      if (previous === undefined) { ops.push([[...path, key], next]); if (ops.length > budget) return false; }
      else if (!collectDelta(previous, next, [...path, key], ops, budget)) return false;
    }
    return true;
  }
  if (Array.isArray(from) && Array.isArray(to) && from.length === to.length) {
    for (let index = 0; index < to.length; index += 1) {
      if (!collectDelta(from[index] as JsonValue, to[index] as JsonValue, [...path, index], ops, budget)) return false;
    }
    return true;
  }
  if (path.length === 0) return false;
  ops.push([path, to]);
  return ops.length <= budget;
}

/**
 * A document's layout receipts are a cache of tens of kilobytes that one
 * typed word changes in a dozen numbers; they rode every save whole. Null
 * when the change is not clearly the smaller thing to send.
 */
async function privateAttrsDelta(stored: Body, carried: Body, baseContentVersion: number): Promise<PrivateAttrsDelta | null> {
  const ops: DeltaOp[] = [];
  if (!collectDelta(stored, carried, [], ops, 2_000)) return null;
  if (JSON.stringify(ops).length * 4 > JSON.stringify(carried).length) return null;
  return { base_content_version: baseContentVersion, ops, sha256: await pythonCanonicalJsonSha256(carried) };
}

/** The layout receipts a document body holds beyond its shared state. */
export function privateLayoutOf(content: Record<string, unknown> | null | undefined): Body {
  return content ? privateLayout(content as Body) : {};
}

/** `shared` with these layout receipts put on it (none: `shared` itself). */
export function withPrivateLayout(shared: Body, receipts: Body): Body {
  return Object.keys(receipts).length
    ? { ...shared, attrs: { ...(record(shared.attrs) ? shared.attrs : {}), ...receipts } } as Body
    : shared;
}

function privateLayout(content: Body): Body {
  const attrs = record(content.attrs) ? content.attrs : {};
  return Object.fromEntries(PRIVATE_LAYOUT.filter((key) => attrs[key] !== undefined)
    .map((key) => [key, attrs[key] as JsonValue]));
}

/**
 * A text node's marks in the order the API's projection lists them: by type,
 * then by Python's `json.dumps(attrs, sort_keys=True)` (the web projection
 * orders them by its own key). Arrays are ordered, so the digest needs one.
 */
function withApiMarkOrder(node: JsonValue): JsonValue {
  if (Array.isArray(node)) return node.map(withApiMarkOrder);
  if (!record(node)) return node;
  // A frozen block (a projection's, kept from one save to the next) is put in
  // the API's order once, and the result is kept frozen with it: its
  // canonical text is then kept too (python-canonical-json.ts).
  if (Object.isFrozen(node)) {
    let ordered = apiOrdered.get(node);
    if (ordered === undefined) {
      ordered = orderMarks(node);
      if (ordered !== node) deepFreezeContainers(ordered);
      apiOrdered.set(node, ordered);
    }
    return ordered;
  }
  return orderMarks(node);
}

const apiOrdered = new WeakMap<object, JsonValue>();

/** Keep what a save's digest reads of one frozen projection block (its API
 * mark order and that order's canonical text) ahead of the save
 * (saveTextWarmup): the same kept results namePageCausalBody makes. */
export function warmNamedBodyBlock(block: object): void {
  if (!Object.isFrozen(block)) return;
  pythonCanonicalJson(withApiMarkOrder(block as JsonValue));
}

/** Freeze what withApiMarkOrder built (nodes, content and mark arrays); the
 * nodes and attribute values it carried over are frozen or the room's own. */
function deepFreezeContainers(value: JsonValue): void {
  if (Array.isArray(value)) { value.forEach(deepFreezeContainers); Object.freeze(value); return; }
  if (!record(value) || Object.isFrozen(value)) return;
  const node = value as Record<string, JsonValue>;
  if (Array.isArray(node.content)) deepFreezeContainers(node.content);
  if (Array.isArray(node.marks)) Object.freeze(node.marks);
  Object.freeze(node);
}

function orderMarks(node: Record<string, JsonValue>): JsonValue {
  const marks = node.marks, content = node.content;
  if (!Array.isArray(content) && !(node.type === "text" && Array.isArray(marks) && marks.length > 1)) return node;
  const next: Record<string, JsonValue> = { ...(node as Record<string, JsonValue>) };
  if (Array.isArray(content)) next.content = content.map((child) => withApiMarkOrder(child as JsonValue));
  if (node.type === "text" && Array.isArray(marks) && marks.length > 1) {
    next.marks = marks.map((mark) => ({
      mark: mark as JsonValue,
      type: String((mark as { type?: unknown }).type),
      attrs: pythonDefaultSortedJson((mark as { attrs?: unknown }).attrs ?? {}),
    })).sort((a, b) => compareStringsByCodePoint(a.type, b.type) || compareStringsByCodePoint(a.attrs, b.attrs))
      .map((entry) => entry.mark);
  }
  return next;
}

/**
 * Name the body of one save. `projection` is what the request document
 * projects to (`PreparedPageCausalCheckpoint.apply`), `authored` the editor's
 * own content for it, `durable` the page the save is based on.
 * Null when this body cannot be named (the caller sends it).
 */
export async function namePageCausalBody(
  model: string,
  projection: Body,
  authored: Body,
  durable: { content_version: number; content_json: Body },
  /** `mirrors`: whether `authored` says what `projection` does, when the
   * caller has found out already (the comparison reads both whole). */
  known: { mirrors?: boolean } = {},
): Promise<PageCausalNamedBody | null> {
  try {
    if (model === "page-json-v1" || model === "code-text-v1") {
      const sha256 = await pythonCanonicalJsonSha256(projection);
      return Object.freeze({ content_sha256: sha256, stored: [{ body: projection, sha256 }], carried: {} });
    }
    if (model !== "document-canonical-v1") return null;
    const shared = withApiMarkOrder(projection) as Body;
    const sha256 = await pythonCanonicalJsonSha256(shared);
    // The author's layout receipts describe the body only while its own
    // content is what the update says (retainDocumentCausalSaveProjection).
    const mirrors = known.mirrors ?? jsonEquals(projection, documentCausalSemanticContent(authored));
    const carried = mirrors ? privateLayout(authored) : {};
    const withLayout = Object.keys(carried).length
      ? { ...shared, attrs: { ...(record(shared.attrs) ? shared.attrs : {}), ...carried } } as Body
      : shared;
    const bare = { body: shared, sha256, readOnly: true };
    if (jsonEquals(carried, privateLayout(durable.content_json))) {
      if (!Object.keys(carried).length) return Object.freeze({ content_sha256: sha256, stored: [bare], carried });
      // Stored at this version they stay; at any other the server drops them.
      return Object.freeze({ content_sha256: sha256, private_attrs_base_content_version: durable.content_version,
        stored: [{ body: withLayout, readOnly: true }, bare], carried });
    }
    const stored = withLayout === shared ? [bare] : [{ body: withLayout, readOnly: true }, bare];
    const previous = privateLayout(durable.content_json);
    const delta = Object.keys(carried).length && Object.keys(previous).length
      ? await privateAttrsDelta(previous, carried, durable.content_version) : null;
    return delta
      ? Object.freeze({ content_sha256: sha256, private_attrs_delta: delta, stored, carried })
      : Object.freeze({ content_sha256: sha256, private_attrs: carried, stored, carried });
  } catch {
    // A value the canonical form cannot spell: the body is sent as before.
    return null;
  }
}

/** The request fields that name the body. */
export function namedBodyRequestFields(named: PageCausalNamedBody): {
  content_sha256: string; private_attrs?: Record<string, unknown>; private_attrs_base_content_version?: number;
  private_attrs_delta?: PrivateAttrsDelta; receipt: "compact";
} {
  return {
    content_sha256: named.content_sha256,
    ...(named.private_attrs !== undefined ? { private_attrs: named.private_attrs } : {}),
    ...(named.private_attrs_delta !== undefined ? { private_attrs_delta: named.private_attrs_delta } : {}),
    ...(named.private_attrs_base_content_version !== undefined
      ? { private_attrs_base_content_version: named.private_attrs_base_content_version } : {}),
    receipt: "compact",
  };
}

/**
 * The body a compact receipt left out: the candidate whose canonical digest
 * is the stored body's (`content_sha256` of the receipt). Null when none is,
 * and the caller reads the page.
 */
export async function storedBodyOfReceipt(
  named: Pick<PageCausalNamedBody, "stored">, contentSha256: string,
): Promise<Body | null> {
  for (const candidate of named.stored) {
    try {
      if ((candidate.sha256 ?? await pythonCanonicalJsonSha256(candidate.body)) === contentSha256) {
        if (!candidate.readOnly) return structuredClone(candidate.body);
        // Handed on as a save's own body: on an inked page each save copied
        // megabytes here, and twice more where it was kept.
        markDocumentSaveBody(candidate.body as Record<string, unknown>);
        return candidate.body;
      }
    } catch { /* not this one */ }
  }
  return null;
}
