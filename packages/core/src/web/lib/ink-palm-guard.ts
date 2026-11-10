/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { INK_STROKE_SILENCE_MS } from "./ink-pointer-policy";










export const INK_PALM_GUARD_CONTACT_PX = 64;
/** A contact this long in one direction is the side of a hand, however narrow. */
export const INK_PALM_GUARD_EDGE_PX = 120;
/** This many finger contacts down together are a hand. */
export const INK_PALM_CLUSTER_TOUCHES = 3;

export interface InkPalmPoint {
  readonly x: number;
  readonly y: number;
}

export interface InkPalmContactSize {
  readonly width?: number;
  readonly height?: number;
}

export interface InkPalmContact {
  /** `p<pointerId>` for a pointer event's contact, `t<identifier>` for a touch event's. */
  readonly key: string;
  readonly startedAt: number;
  readonly start: InkPalmPoint;
  last: InkPalmPoint;
  palm: boolean;
  /**
   * Judged the hand only after it was down a while (the pen arrived, a third
   * contact landed): the page had already seen its press.
   */
  late: boolean;
}

export type InkPalmGuardListener = (contacts: readonly InkPalmContact[]) => void;

export function isHandSizedContact(size: InkPalmContactSize | undefined): boolean {
  const width = size?.width ?? 0;
  const height = size?.height ?? 0;
  return (
    (width >= INK_PALM_GUARD_CONTACT_PX && height >= INK_PALM_GUARD_CONTACT_PX) ||
    Math.max(width, height) >= INK_PALM_GUARD_EDGE_PX
  );
}

function clock(): number {
  return typeof performance === "undefined" ? 0 : performance.now();
}

export class InkPalmGuard {
  private penDown = false;
  private penInRange = false;
  private penSeenAt = Number.NEGATIVE_INFINITY;
  private readonly contacts = new Map<string, InkPalmContact>();
  private readonly listeners = new Set<InkPalmGuardListener>();

  /** The pen touched down: the contacts on the glass are its hand. */
  penContact(at: number): void {
    this.sawPen(at);
    this.penDown = true;
    this.judgeAll();
  }

  /** The pen moved, in contact or hovering. */
  penMoved(at: number): void {
    this.sawPen(at);
  }

  /**
   * The pen moved with no button down: it hovers, so it is not on the glass
   * whatever the guard last heard (a lift the page never got kept every
   * finger the hand for good).
   */
  penHovered(at: number): void {
    if (this.penDown) this.penLifted(at);
    else this.sawPen(at);
  }

  /** The pen lifted (or its stroke was cancelled): the hand is still there. */
  penLifted(at: number): void {
    this.sawPen(at);
    if (!this.penDown) return;
    this.penDown = false;
  }

  /** The pen left the digitizer's range. */
  penLeft(at: number): void {
    if (this.penDown) this.penLifted(at);
    this.penInRange = false;
  }

  private sawPen(at: number): void {
    this.penInRange = true;
    this.penSeenAt = Math.max(this.penSeenAt, at);
  }

  /**
   * Whether the pen is on the glass: a finger landing now is the writing
   * hand. A pen not heard from for {@link INK_STROKE_SILENCE_MS} has lifted
   * where the page could not hear it.
   */
  penOnGlass(now: number = clock()): boolean {
    if (this.penDown && now - this.penSeenAt >= INK_STROKE_SILENCE_MS) this.penDown = false;
    return this.penDown;
  }

  /** A finger contact landed. True when it is the hand. */
  contactDown(key: string, at: number, point: InkPalmPoint, size?: InkPalmContactSize): boolean {
    const contact: InkPalmContact = {
      key,
      startedAt: at,
      start: point,
      last: point,
      palm: this.penOnGlass(at) || isHandSizedContact(size),
      late: false,
    };
    this.contacts.set(key, contact);
    const kind = key.charAt(0) === "t" ? "t" : "p";
    // Only beside a pen: someone inking with a finger may rest the others.
    if (
      !contact.palm &&
      Number.isFinite(this.penSeenAt) &&
      this.fingersDown(kind) >= INK_PALM_CLUSTER_TOUCHES
    ) {
      // This one was never a finger to anyone; the others were until now.
      contact.palm = true;
      this.judge([...this.contacts.values()].filter((other) => other.key.startsWith(kind)));
    }
    return contact.palm;
  }

  contactMoved(key: string, point: InkPalmPoint): void {
    const contact = this.contacts.get(key);
    if (contact) contact.last = point;
  }

  /**
   * Forget the contacts of one kind whose lift never came: a touch event
   * lists every contact on the glass, so one missing from it is gone.
   */
  forgetContactsExcept(kind: "p" | "t", live: ReadonlySet<string>): void {
    for (const key of [...this.contacts.keys()]) {
      if (key.startsWith(kind) && !live.has(key)) this.contacts.delete(key);
    }
  }

  /** Forget every contact (the page went hidden: the glass is not ours). */
  forgetContacts(): void {
    this.contacts.clear();
  }

  /** The contact lifted. What it was, for whoever cleans up after it. */
  contactUp(key: string): InkPalmContact | null {
    const contact = this.contacts.get(key) ?? null;
    this.contacts.delete(key);
    return contact;
  }

  isPalm(key: string): boolean {
    return this.contacts.get(key)?.palm === true;
  }

  contact(key: string): InkPalmContact | null {
    return this.contacts.get(key) ?? null;
  }

  /** Finger contacts down now that are not the hand, of one kind of event. */
  fingersDown(kind: "p" | "t"): number {
    let count = 0;
    for (const contact of this.contacts.values()) {
      if (contact.key.startsWith(kind) && !contact.palm) count += 1;
    }
    return count;
  }

  /** Finger contacts down now, the hand's included, of one kind of event. */
  contactsDown(kind: "p" | "t"): number {
    let count = 0;
    for (const contact of this.contacts.values()) {
      if (contact.key.startsWith(kind)) count += 1;
    }
    return count;
  }

  /** Someone else decided this contact is the hand (it rested, then slid). */
  markPalm(key: string): void {
    const contact = this.contacts.get(key);
    if (contact) this.judge([contact]);
  }

  /** Called with the contacts that were just judged the hand after landing. */
  subscribe(listener: InkPalmGuardListener): () => void {
    this.listeners.add(listener);
    return () => {
      this.listeners.delete(listener);
    };
  }

  /** Forget everything (the last ink tool was put away; tests). */
  reset(): void {
    this.penDown = false;
    this.penInRange = false;
    this.penSeenAt = Number.NEGATIVE_INFINITY;
    this.contacts.clear();
  }

  private judgeAll(): void {
    this.judge([...this.contacts.values()]);
  }

  private judge(contacts: readonly InkPalmContact[]): void {
    const judged = contacts.filter((contact) => !contact.palm);
    if (judged.length === 0) return;
    for (const contact of judged) {
      contact.palm = true;
      contact.late = true;
    }
    for (const listener of [...this.listeners]) listener(judged);
  }
}

/** The one guard: one hand writes in the product at a time. */
export const inkPalmGuard = new InkPalmGuard();

/** A touch event's contact that is the pen itself (Safari: Apple Pencil). */
export function isStylusTouch(touch: unknown): boolean {
  return (touch as { touchType?: string } | null)?.touchType === "stylus";
}

/** The guard's key for a touch event's contact. */
export function inkPalmTouchKey(identifier: number): string {
  return `t${identifier}`;
}

/** The guard's key for a pointer event's contact. */
export function inkPalmPointerKey(pointerId: number): string {
  return `p${pointerId}`;
}

/**
 * The finger contacts of a touch list a gesture may read: not the pen, not
 * the hand. A heel and a little finger are two contacts, and so are a pen
 * and the hand holding it; neither pair is a pinch.
 */
export function fingerTouches<T extends { readonly identifier: number }>(
  touches: ArrayLike<T>,
  guard: Pick<InkPalmGuard, "isPalm"> = inkPalmGuard,
): T[] {
  const fingers: T[] = [];
  for (let index = 0; index < touches.length; index += 1) {
    const touch = touches[index];
    if (!touch || isStylusTouch(touch) || guard.isPalm(inkPalmTouchKey(touch.identifier))) continue;
    fingers.push(touch);
  }
  return fingers;
}
