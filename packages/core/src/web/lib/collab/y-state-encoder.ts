/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The whole state of a Y.Doc as one update, byte for byte what
 * `Y.encodeStateAsUpdate(doc)` writes, and what it holds beyond a state
 * vector, byte for byte what `Y.encodeStateAsUpdate(doc, vector)` writes,
 * with the work proportional to what changed since the last time.
 *
 * Every save of a causal page encoded its whole document (its recovery copy,
 * then its checkpoint, then the delta beyond the stored checkpoint): on an
 * inked page of a few hundred strokes 100-250 ms of the editor's thread at a
 * time, under the pen, and growing with each stroke. The encoding of a
 * struct depends only on that struct: its id, its length, its origins, its
 * parent and its content. Yjs changes a struct in place in three ways only,
 * each of which changes its length or replaces its content: a split (the
 * left part keeps the object and shrinks), a merge (the left part grows) and
 * garbage collection (the content becomes ContentDeleted, or a GC struct
 * replaces the item). So each client's structs are kept with their encoded
 * bytes and the runs of deleted ones, and only those from the first one that
 * differs in identity, length or content are encoded again. Where to start
 * looking is what the transactions since reported (their new structs, their
 * deletions and their splits); a client none of them touched is not looked
 * at, and a document this module has not watched is compared struct by
 * struct.
 *
 * Anything this cannot encode for certain -- a transaction still open,
 * updates still waiting for their predecessors -- goes to Yjs itself.
 */
import * as Y from "yjs";

type Struct = Y.Item | Y.GC;

interface ClientEncoding {
  readonly structs: Struct[];
  readonly lengths: number[];
  readonly contents: unknown[];
  readonly deleted: boolean[];
  /** Byte offset just past each struct's encoding in `bytes`. */
  readonly ends: number[];
  readonly bytes: Uint8Array;
  /** Runs of deleted structs: [clock, length], as createDeleteSetFromStructStore makes them. */
  readonly deletes: ReadonlyArray<readonly [number, number]>;
  /** `bytes` cut into pieces of about STATE_PIECE_BYTES: a piece object
   * stands for the same bytes for as long as it lives (one inside the kept
   * prefix of a re-encoding is carried over), so a consumer keeping
   * something per piece keeps it while those bytes stay. */
  readonly pieces: readonly StatePieceRange[];
}

interface StatePieceRange { readonly start: number; readonly end: number }

/** The size the encoded state of a client is cut into (see yDocStatePieces). */
export const STATE_PIECE_BYTES = 48 * 1024;

/** A piece of an encoded state: `bytes`, and an object that stands for
 * exactly those bytes wherever it appears again (`id`, null for the small
 * framing bytes between clients, which are not kept). */
export interface YDocStatePiece {
  readonly id: object | null;
  readonly bytes: Uint8Array;
}

/** The state an encoding wrote -> its pieces, in order (their bytes
 * concatenated are the state). */
const statePieces = new WeakMap<Uint8Array, readonly YDocStatePiece[]>();

/** The pieces of a state this module encoded (encodeYDocState*), or null. */
export function yDocStatePieces(state: Uint8Array): readonly YDocStatePiece[] | null {
  return statePieces.get(state) ?? null;
}

/** Pieces of `length` bytes: those of `cached` lying wholly in the first
 * `prefix` bytes (which are the same bytes) are kept, a short last one of
 * them is cut again with what follows, and the rest is cut anew. */
function cutPieces(length: number, cached: readonly StatePieceRange[] | null, prefix: number): StatePieceRange[] {
  const pieces: StatePieceRange[] = [];
  if (cached) {
    for (const piece of cached) {
      if (piece.end > prefix) break;
      pieces.push(piece);
    }
    const last = pieces.at(-1);
    if (last && last.end - last.start < STATE_PIECE_BYTES / 2) pieces.pop();
  }
  let at = pieces.at(-1)?.end ?? 0;
  while (at < length) {
    const end = Math.min(length, at + STATE_PIECE_BYTES);
    pieces.push(Object.freeze({ start: at, end }));
    at = end;
  }
  return pieces;
}

interface DocEncoding {
  readonly clients: Map<number, ClientEncoding>;
  /** The last whole state written, and the client encodings it was made of:
   * an unchanged document hands back the very same bytes. */
  last: { readonly state: Uint8Array; readonly from: readonly ClientEncoding[] } | null;
  /** Per client, the lowest clock a transaction touched since the last
   * encoding; a client absent here is untouched. Null: not known (look at
   * every client from its first struct). */
  dirty: Map<number, number> | null;
}

const encodings = new WeakMap<Y.Doc, DocEncoding>();

function writeVarUint(out: number[], value: number) {
  while (value > 127) {
    out.push(128 | (value & 127));
    value = Math.floor(value / 128);
  }
  out.push(value & 127);
}

/** lib0's encoder keeps full chunks in `bufs` and the current one in `cbuf`. */
interface Lib0Encoder { cpos: number; cbuf: Uint8Array; bufs: Uint8Array[] }

function contentOf(struct: Struct): unknown {
  return struct instanceof Y.Item ? struct.content : null;
}

function docEncoding(doc: Y.Doc): DocEncoding {
  let known = encodings.get(doc);
  if (known) return known;
  const created: DocEncoding = { clients: new Map(), dirty: null, last: null };
  doc.on("afterTransactionCleanup", (transaction: Y.Transaction) => {
    const dirty = created.dirty;
    if (!dirty) return;
    const touch = (client: number, clock: number) => {
      const at = dirty.get(client);
      if (at === undefined || clock < at) dirty.set(client, clock);
    };
    transaction.afterState.forEach((clock, client) => {
      const before = transaction.beforeState.get(client) ?? 0;
      if (before !== clock) touch(client, before);
    });
    transaction.deleteSet.clients.forEach((items, client) => {
      for (const item of items) touch(client, item.clock);
    });
    for (const struct of (transaction as unknown as { _mergeStructs: Struct[] })._mergeStructs) {
      touch(struct.id.client, struct.id.clock);
    }
  });
  encodings.set(doc, (known = created));
  return known;
}

/** The first struct whose encoding may differ from the cached one, looking
 * from `from` (structs before it are known unchanged). */
function firstChanged(structs: Struct[], cached: ClientEncoding, from: number): number {
  const limit = Math.min(structs.length, cached.structs.length);
  for (let index = Math.max(0, from); index < limit; index += 1) {
    const struct = structs[index];
    if (struct !== cached.structs[index] || struct.length !== cached.lengths[index] ||
        contentOf(struct) !== cached.contents[index] || struct.deleted !== cached.deleted[index]) return index;
  }
  return limit;
}











function encodeClient(structs: Struct[], cached: ClientEncoding | undefined, from: number): ClientEncoding {
  const keep = cached ? from : 0;
  const prefix = keep > 0 ? cached!.ends[keep - 1] : 0;
  const ends = keep > 0 ? cached!.ends.slice(0, keep) : [];
  const chunks: Uint8Array[] = [];
  if (prefix > 0) chunks.push(cached!.bytes.subarray(0, prefix));
  let total = prefix;
  // A run of cached bytes being carried over, [runStart, runEnd) of cached.bytes.
  let runStart = -1, runEnd = -1;
  const closeRun = () => {
    if (runStart >= 0) chunks.push(cached!.bytes.subarray(runStart, runEnd));
    runStart = runEnd = -1;
  };
  let encoder: InstanceType<typeof Y.UpdateEncoderV1> | null = null;
  let encoderStart = 0;
  const written = () => {
    const rest = encoder!.restEncoder as unknown as Lib0Encoder;
    let size = rest.cpos;
    for (const buf of rest.bufs) size += buf.length;
    return size;
  };
  const closeEncoder = () => {
    if (encoder) chunks.push(encoder.toUint8Array());
    encoder = null;
  };
  let j = cached ? keep : 0;
  const cachedCount = cached ? cached.structs.length : 0;
  for (let index = keep; index < structs.length; index += 1) {
    const struct = structs[index];
    if (cached) {
      const clock = struct.id.clock;
      while (j < cachedCount && cached.structs[j].id.clock < clock) j += 1;
    }
    const same = Boolean(cached) && j < cachedCount && cached!.structs[j] === struct &&
      cached!.lengths[j] === struct.length && cached!.contents[j] === contentOf(struct) &&
      cached!.deleted[j] === struct.deleted;
    if (same) {
      closeEncoder();
      const start = j > 0 ? cached!.ends[j - 1] : 0, end = cached!.ends[j];
      if (runStart >= 0 && runEnd === start) runEnd = end;
      else { closeRun(); runStart = start; runEnd = end; }
      total += end - start;
      j += 1;
    } else {
      closeRun();
      if (!encoder) { encoder = new Y.UpdateEncoderV1(); encoderStart = total; }
      struct.write(encoder, 0);
      total = encoderStart + written();
    }
    ends.push(total);
  }
  closeRun();
  closeEncoder();
  const bytes = chunks.length === 1 && chunks[0].byteOffset === 0 && chunks[0].length === chunks[0].buffer.byteLength
    ? chunks[0] : concat(chunks);
  const deletes: Array<[number, number]> = [];
  let scanFrom = 0;
  if (keep > 0 && keep < structs.length) {
    const boundary = structs[keep].id.clock;
    scanFrom = keep;
    for (const run of cached!.deletes) {
      if (run[0] + run[1] < boundary) { deletes.push([run[0], run[1]]); continue; }
      // The first run reaching the boundary (it may go on into the structs
      // read again) is read again from its own start.
      if (run[0] < boundary) scanFrom = Y.findIndexSS(structs, run[0]);
      break;
    }
  }
  for (let index = scanFrom; index < structs.length; index += 1) {
    const struct = structs[index];
    if (!struct.deleted) continue;
    const last = deletes.at(-1);
    if (last && last[0] + last[1] === struct.id.clock && index > scanFrom && structs[index - 1].deleted) last[1] += struct.length;
    else deletes.push([struct.id.clock, struct.length]);
  }
  return {
    structs: structs.slice(),
    lengths: structs.map((struct) => struct.length),
    contents: structs.map(contentOf),
    deleted: structs.map((struct) => struct.deleted),
    ends,
    bytes,
    deletes,
    pieces: cutPieces(bytes.length, keep > 0 ? cached!.pieces : null, prefix),
  };
}

function concat(parts: Uint8Array[]): Uint8Array {
  let size = 0;
  for (const part of parts) size += part.length;
  const out = new Uint8Array(size);
  let offset = 0;
  for (const part of parts) { out.set(part, offset); offset += part.length; }
  return out;
}

/** Bring each client's kept encoding up to date; the clients, highest first. */
function refresh(doc: Y.Doc): Array<[number, ClientEncoding]> {
  const known = docEncoding(doc);
  const clients = [...doc.store.clients.entries()]
    .filter(([, structs]) => structs.length > 0)
    .sort((a, b) => b[0] - a[0]);
  const present = new Set<number>();
  const result: Array<[number, ClientEncoding]> = [];
  for (const [client, structs] of clients) {
    present.add(client);
    const cached = known.clients.get(client);
    let encoding = cached;
    if (!cached) encoding = encodeClient(structs, undefined, 0);
    else {
      const dirtyClock = known.dirty ? known.dirty.get(client) : 0;
      if (dirtyClock !== undefined) {
        // From the struct before the one holding the first touched clock:
        // a merge grows the struct on its left.
        const holding = dirtyClock >= (structs.at(-1)!.id.clock + structs.at(-1)!.length)
          ? structs.length : Y.findIndexSS(structs, Math.max(dirtyClock, structs[0].id.clock));
        const from = firstChanged(structs, cached, Math.min(holding, cached.structs.length) - 1);
        if (from < structs.length || from < cached.structs.length) encoding = encodeClient(structs, cached, from);
      } else if (structs.length !== cached.structs.length) {
        encoding = encodeClient(structs, cached, firstChanged(structs, cached, 0));
      }
    }
    if (encoding !== cached) known.clients.set(client, encoding!);
    result.push([client, encoding!]);
  }
  for (const client of [...known.clients.keys()]) if (!present.has(client)) known.clients.delete(client);
  known.dirty = new Map();
  return result;
}

function deleteSetBytes(clients: Array<[number, ClientEncoding]>): Uint8Array {
  const tail: number[] = [];
  const withDeletes = clients.filter(([, encoding]) => encoding.deletes.length > 0);
  writeVarUint(tail, withDeletes.length);
  for (const [client, encoding] of withDeletes) {
    writeVarUint(tail, client);
    writeVarUint(tail, encoding.deletes.length);
    for (const [clock, length] of encoding.deletes) {
      writeVarUint(tail, clock);
      writeVarUint(tail, length);
    }
  }
  return Uint8Array.from(tail);
}

function unsettled(doc: Y.Doc): boolean {
  const store = doc.store;
  return Boolean(store.pendingStructs || store.pendingDs || (doc as unknown as { _transaction: unknown })._transaction);
}

/** What the document holds: the clock each client reached and, per client,
 * the [from, to) ranges of its deleted structs (each maximal run). */
export interface YDocExtent {
  readonly clocks: Map<number, number>;
  readonly deleted: Map<number, ReadonlyArray<readonly [number, number]>>;
}

/** Structs encoded per client per step of warmYDocStateEncoding. */
const WARM_STRUCTS = 2048;











export function warmYDocStateEncoding(doc: Y.Doc, budgetMs: number, structsPerStep = WARM_STRUCTS): boolean {
  if (unsettled(doc)) return false;
  const known = docEncoding(doc);
  if (known.dirty === null) {
    // Nothing encoded yet: from now on the transactions are tracked, so a
    // prefix kept here is checked as any kept encoding is.
    if (known.clients.size > 0) return true;
    known.dirty = new Map();
  }
  const end = performance.now() + budgetMs;
  let done = true;
  for (const [client, structs] of doc.store.clients) {
    if (structs.length === 0) continue;
    const cached = known.clients.get(client);
    const have = cached ? cached.structs.length : 0;
    if (have >= structs.length) continue;
    if (performance.now() >= end) return false;
    let from = have;
    const dirtyClock = known.dirty.get(client);
    if (cached && dirtyClock !== undefined) {
      const holding = Y.findIndexSS(structs, Math.min(Math.max(dirtyClock, structs[0].id.clock),
        structs.at(-1)!.id.clock + structs.at(-1)!.length - 1));
      from = firstChanged(structs, cached, Math.min(holding, have) - 1);
    }
    const until = Math.min(structs.length, from + structsPerStep);
    known.clients.set(client, encodeClient(structs.slice(0, until), cached, from));
    if (until < structs.length) done = false;
  }
  return done;
}

/** `Y.encodeStateAsUpdate(doc)`, reusing the encoding of unchanged structs. */
export function encodeYDocState(doc: Y.Doc): Uint8Array {
  return encodeYDocStateAndBeyond(doc, null).state;
}

/**
 * `Y.encodeStateAsUpdate(doc)` and, given a state vector,
 * `Y.encodeStateAsUpdate(doc, Y.encodeStateVector(vector))` of the same
 * moment, sharing one pass.
 */
export function encodeYDocStateAndBeyond(
  doc: Y.Doc,
  vector: ReadonlyMap<number, number> | null,
): { state: Uint8Array; beyond: Uint8Array | null; extent: YDocExtent | null } {
  if (unsettled(doc)) {
    return { state: Y.encodeStateAsUpdate(doc),
      beyond: vector ? Y.encodeStateAsUpdate(doc, Y.encodeStateVector(vector as Map<number, number>)) : null, extent: null };
  }
  const clients = refresh(doc);
  const extent: YDocExtent = { clocks: new Map(), deleted: new Map() };
  for (const [client, encoding] of clients) {
    const last = encoding.structs.at(-1)!;
    extent.clocks.set(client, last.id.clock + last.length);
    if (encoding.deletes.length) extent.deleted.set(client, encoding.deletes.map(([clock, length]) => [clock, clock + length] as const));
  }
  const deletes = deleteSetBytes(clients);
  const head: number[] = [];
  writeVarUint(head, clients.length);
  const parts: Uint8Array[] = [Uint8Array.from(head)];
  for (const [client, encoding] of clients) {
    const header: number[] = [];
    writeVarUint(header, encoding.structs.length);
    writeVarUint(header, client);
    writeVarUint(header, encoding.structs[0].id.clock);
    parts.push(Uint8Array.from(header), encoding.bytes);
  }
  parts.push(deletes);
  const known = docEncoding(doc);
  const from = clients.map(([, encoding]) => encoding);
  const previous = known.last;
  const reused = previous && previous.from.length === from.length && previous.from.every((encoding, index) => encoding === from[index]);
  const state = reused ? previous.state : concat(parts);
  if (!reused) {
    // The same order as `parts`: the count, then each client's header and
    // its bytes (by piece), then the delete set.
    const pieces: YDocStatePiece[] = [{ id: null, bytes: parts[0] }];
    clients.forEach(([, encoding], index) => {
      pieces.push({ id: null, bytes: parts[1 + index * 2] });
      for (const piece of encoding.pieces) pieces.push({ id: piece, bytes: encoding.bytes.subarray(piece.start, piece.end) });
    });
    pieces.push({ id: null, bytes: deletes });
    statePieces.set(state, pieces);
  }
  known.last = { state, from };
  if (!vector) return { state, beyond: null, extent };
  // writeClientsStructs: the clients the vector names that went past it,
  // then every client it does not name, from its first clock; highest first.
  const beyondParts: Uint8Array[] = [];
  let written = 0;
  for (const [client, encoding] of clients) {
    const structs = encoding.structs;
    const last = structs.at(-1)!;
    const state = last.id.clock + last.length;
    const known = vector.get(client);
    if (known !== undefined && state <= known) continue;
    const clock = Math.max(known ?? 0, structs[0].id.clock);
    const start = Y.findIndexSS(structs, clock);
    const header: number[] = [];
    writeVarUint(header, structs.length - start);
    writeVarUint(header, client);
    writeVarUint(header, clock);
    beyondParts.push(Uint8Array.from(header));
    const first = structs[start];
    if (clock > first.id.clock) {
      const encoder = new Y.UpdateEncoderV1();
      first.write(encoder, clock - first.id.clock);
      beyondParts.push(encoder.toUint8Array());
    } else {
      beyondParts.push(encoding.bytes.subarray(start > 0 ? encoding.ends[start - 1] : 0, encoding.ends[start]));
    }
    beyondParts.push(encoding.bytes.subarray(encoding.ends[start]));
    written += 1;
  }
  const count: number[] = [];
  writeVarUint(count, written);
  return { state, beyond: concat([Uint8Array.from(count), ...beyondParts, deletes]), extent };
}
