/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * One Undo history over several collaboration rooms (a page held as parts:
 * the body's room and a room for each later part it has loaded).
 *
 * Each room keeps this member's own steps in its own history (a
 * `createJsonCollaborationHistory`: only this member's local origin, a
 * peer's input never undone). This orders them: every change this member
 * makes -- one command, a run of typing, an insert the server made for them
 * -- is one entry naming how many steps it took in which rooms, and Undo
 * takes the last entry back in every room it touched (a change across two
 * parts is one Undo), Redo brings it again. A room's steps are counted when
 * they appear (`settle`: a room may write a change a moment after the
 * command, as the page room does from its render), and belong to the entry
 * before them.
 *
 * A room whose step cannot be taken back (someone else changed that very
 * place since, or the room was read afresh and its history is gone) is left
 * as it is while the other rooms of the entry are undone: the history moves
 * on in step, and `onPartial` says so.
 */

export interface HistoryRoom {
  readonly undoDepth: number;
  readonly redoDepth: number;
  /** Take one step back; false when there was none to take. */
  undo(): boolean;
  redo(): boolean;
  stopCapturing(): void;
}

/**
 * A step of this member's that no room holds: a server write made for them
 * whose answer carried no room update (its room was not connected), taken
 * back and brought again by the server's own inverse write. `undo`/`redo`
 * run it; they may finish later and report their own refusal.
 */
export interface ExternalHistoryStep {
  undo(): void;
  redo(): void;
}

interface Entry {
  readonly steps: Map<string, number>;
  readonly external?: ExternalHistoryStep;
}

export class MultiRoomHistory {
  private readonly undoStack: Entry[] = [];
  private readonly redoStack: Entry[] = [];
  /** Each room's undo depth as last counted. */
  private readonly counted = new Map<string, number>();
  private readonly listeners = new Set<() => void>();

  constructor(
    private readonly rooms: () => ReadonlyMap<string, HistoryRoom>,
    private readonly onPartial?: () => void,
  ) {}

  subscribe(listener: () => void): () => void {
    this.listeners.add(listener);
    return () => {
      this.listeners.delete(listener);
    };
  }

  private versionValue = 0;

  /** Changes whenever what Undo or Redo would do may have. */
  get version(): number {
    return this.versionValue;
  }

  private changed(): void {
    this.versionValue += 1;
    for (const listener of [...this.listeners]) listener();
  }

  /** A room's own history changed (a step appeared or went): readers look again. */
  touch(): void {
    this.changed();
  }

  /** Whether Undo has something to take back (read without counting). */
  get canUndo(): boolean {
    if (this.undoStack.length > 0) return true;
    for (const [key, room] of this.rooms()) {
      const before = this.counted.get(key);
      if (before !== undefined && room.undoDepth > before) return true;
    }
    return false;
  }

  get canRedo(): boolean {
    return this.redoStack.length > 0;
  }

  /**
   * Count the steps the rooms took since they were last counted, into the
   * last entry (or a new one when `fresh`, or when there is none). A room
   * read afresh (its depth fell) is counted from where it now is.
   */
  settle(fresh = false): void {
    const steps = new Map<string, number>();
    for (const [key, room] of this.rooms()) {
      const depth = room.undoDepth;
      const before = this.counted.get(key);
      this.counted.set(key, depth);
      if (before === undefined || depth <= before) continue;
      steps.set(key, depth - before);
    }
    for (const key of [...this.counted.keys()]) if (!this.rooms().has(key)) this.counted.delete(key);
    if (steps.size === 0) return;
    const top = this.undoStack.at(-1);
    if (top && !fresh) {
      for (const [key, count] of steps) top.steps.set(key, (top.steps.get(key) ?? 0) + count);
    } else {
      this.undoStack.push({ steps });
    }
    // This member's new input: Redo is over, as in one room.
    if (this.redoStack.length > 0) this.redoStack.length = 0;
    this.changed();
  }

  /**
   * A change of this member's starts: what the rooms took before it belongs
   * to the entry before, and every room starts a step of its own. `run` makes
   * the change; the steps it takes at once are a new entry (and the ones a
   * room takes a moment later join it at the next `settle`).
   */
  step<T>(run: () => T): T {
    this.settle();
    for (const room of this.rooms().values()) room.stopCapturing();
    const result = run();
    this.settle(true);
    return result;
  }

  /** A step of this member's made outside the rooms (ExternalHistoryStep):
   * an entry of its own, after what the rooms took before it. */
  external(step: ExternalHistoryStep): void {
    this.settle();
    for (const room of this.rooms().values()) room.stopCapturing();
    this.undoStack.push({ steps: new Map(), external: step });
    if (this.redoStack.length > 0) this.redoStack.length = 0;
    this.changed();
  }

  /**
   * A server write of a step (ExternalHistoryStep) was refused after the entry
   * had moved. `forget`: the step can never be made again (someone changed its
   * slide), so it is out of both stacks and does not use up the next press.
   * `putBack`: nothing happened, so it returns to the stack it came from, ready
   * to be pressed again.
   */
  settleExternal(step: ExternalHistoryStep, outcome: "forget" | "putBack"): void {
    for (const [stack, other] of [[this.redoStack, this.undoStack], [this.undoStack, this.redoStack]] as const) {
      let at = stack.length - 1;
      while (at >= 0 && stack[at].external !== step) at -= 1;
      if (at < 0) continue;
      const [entry] = stack.splice(at, 1);
      if (outcome === "putBack") other.push(entry);
      this.changed();
      return;
    }
  }

  /** Take the last entry back in every room it touched; false when there was
   * none. */
  undo(): boolean {
    return this.replay("undo");
  }

  redo(): boolean {
    return this.replay("redo");
  }

  private replay(direction: "undo" | "redo"): boolean {
    this.settle();
    const from = direction === "undo" ? this.undoStack : this.redoStack;
    const to = direction === "undo" ? this.redoStack : this.undoStack;
    const entry = from.pop();
    if (!entry) return false;
    // A step outside the rooms, with what the rooms took after it: taken
    // back last, made again first.
    if (direction === "redo") entry.external?.redo();
    const rooms = this.rooms();
    let partial = false;
    for (const room of rooms.values()) room.stopCapturing();
    for (const [key, count] of entry.steps) {
      const room = rooms.get(key);
      if (!room) {
        partial = true;
        continue;
      }
      for (let step = 0; step < count; step += 1) {
        if (!(direction === "undo" ? room.undo() : room.redo())) {
          partial = true;
          break;
        }
      }
    }
    if (direction === "undo") entry.external?.undo();
    to.push(entry);
    // What the replay itself did to the rooms is not a new entry.
    for (const [key, room] of rooms) this.counted.set(key, room.undoDepth);
    if (partial) this.onPartial?.();
    this.changed();
    return true;
  }

  /** Forget a room (closed): its steps are no longer taken back. */
  forget(key: string): void {
    this.counted.delete(key);
  }

  clear(): void {
    this.undoStack.length = 0;
    this.redoStack.length = 0;
    this.counted.clear();
    this.changed();
  }
}
