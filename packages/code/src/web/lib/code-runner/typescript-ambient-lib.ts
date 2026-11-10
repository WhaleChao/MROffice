/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The built-ins every in-editor TypeScript service compiles against under
 * `noLib: true`: hover/completion/signature help (code-intelligence), the
 * Problems list (lsp-diagnostics) and the open-window project
 * (workspace-intelligence). One declaration, so a page's Problems do not
 * change with how many code windows are open.
 *
 * Deliberately small and permissive instead of the real ~10k-line lib: it
 * declares the surface code pages actually touch, typed where that helps
 * (Math, JSON, Map/Set, the iteration protocol) and `any` elsewhere, with an
 * index-signature fallback on the typed objects so a member it does not spell
 * out is never reported as an error in the person's code.
 */
export const TYPESCRIPT_AMBIENT_LIB = `
interface Boolean {}
interface Function { apply: any; call: any; bind: any; readonly length: number; }
interface CallableFunction extends Function {}
interface NewableFunction extends Function {}
interface IArguments { [i: number]: any; length: number; [Symbol.iterator](): IterableIterator<any>; }
interface Number { toFixed(fractionDigits?: number): string; toString(radix?: number): string; }
interface RegExp { test(string: string): boolean; exec(string: string): any; }
interface Object { hasOwnProperty(v: PropertyKey): boolean; toString(): string; }
type PropertyKey = string | number | symbol;
interface String {
  readonly length: number;
  [i: number]: string;
  charAt(pos: number): string;
  charCodeAt(index: number): number;
  slice(start?: number, end?: number): string;
  substring(start: number, end?: number): string;
  split(separator?: any, limit?: number): string[];
  indexOf(searchString: string, position?: number): number;
  lastIndexOf(searchString: string, position?: number): number;
  toUpperCase(): string;
  toLowerCase(): string;
  trim(): string;
  replace(searchValue: any, replaceValue: any): string;
  includes(searchString: string, position?: number): boolean;
  startsWith(searchString: string, position?: number): boolean;
  endsWith(searchString: string, endPosition?: number): boolean;
  padStart(maxLength: number, fillString?: string): string;
  padEnd(maxLength: number, fillString?: string): string;
  repeat(count: number): string;
  concat(...strings: string[]): string;
  match(regexp: any): any;
  [Symbol.iterator](): IterableIterator<string>;
}
interface Array<T> {
  length: number;
  [i: number]: T;
  map<U>(callbackfn: (value: T, index: number, array: T[]) => U): U[];
  filter(predicate: (value: T, index: number, array: T[]) => any): T[];
  forEach(callbackfn: (value: T, index: number, array: T[]) => void): void;
  reduce(callbackfn: (previousValue: any, currentValue: T, currentIndex: number) => any, initialValue?: any): any;
  push(...items: T[]): number;
  pop(): T | undefined;
  shift(): T | undefined;
  unshift(...items: T[]): number;
  slice(start?: number, end?: number): T[];
  splice(start: number, deleteCount?: number, ...items: T[]): T[];
  concat(...items: any[]): T[];
  join(separator?: string): string;
  indexOf(searchElement: T, fromIndex?: number): number;
  includes(searchElement: T, fromIndex?: number): boolean;
  find(predicate: (value: T, index: number) => any): T | undefined;
  findIndex(predicate: (value: T, index: number) => any): number;
  some(predicate: (value: T, index: number) => any): boolean;
  every(predicate: (value: T, index: number) => any): boolean;
  sort(compareFn?: (a: T, b: T) => number): this;
  reverse(): T[];
  flat(): any[];
  entries(): IterableIterator<[number, T]>;
  keys(): IterableIterator<number>;
  values(): IterableIterator<T>;
  [Symbol.iterator](): IterableIterator<T>;
}
interface ReadonlyArray<T> { readonly length: number; readonly [i: number]: T; [Symbol.iterator](): IterableIterator<T>; }
interface Promise<T> {
  then<U>(onfulfilled?: (value: T) => U, onrejected?: (reason: any) => any): Promise<U>;
  catch(onrejected?: (reason: any) => any): Promise<T>;
  finally(onfinally?: () => void): Promise<T>;
}
interface PromiseLike<T> { then(...a: any[]): any; }
interface PromiseConstructor {
  new <T>(executor: (resolve: (value: T) => void, reject: (reason?: any) => void) => void): Promise<T>;
  resolve<T>(value: T): Promise<T>;
  reject(reason?: any): Promise<never>;
  all(values: any[]): Promise<any[]>;
  race(values: any[]): Promise<any>;
}
declare var Promise: PromiseConstructor;
interface TemplateStringsArray extends ReadonlyArray<string> { readonly raw: readonly string[]; }
interface ImportMeta {}
declare var console: {
  log(...data: any[]): void;
  error(...data: any[]): void;
  warn(...data: any[]): void;
  info(...data: any[]): void;
  debug(...data: any[]): void;
  table(tabularData?: any, properties?: string[]): void;
};
interface Math {
  // Members this list leaves out stay untyped instead of becoming errors.
  [member: string]: any;
  readonly E: number;
  readonly LN10: number;
  readonly LN2: number;
  readonly LOG2E: number;
  readonly LOG10E: number;
  readonly PI: number;
  readonly SQRT1_2: number;
  readonly SQRT2: number;
  abs(x: number): number;
  acos(x: number): number;
  acosh(x: number): number;
  asin(x: number): number;
  asinh(x: number): number;
  atan(x: number): number;
  atanh(x: number): number;
  cbrt(x: number): number;
  ceil(x: number): number;
  clz32(x: number): number;
  cos(x: number): number;
  cosh(x: number): number;
  exp(x: number): number;
  expm1(x: number): number;
  floor(x: number): number;
  fround(x: number): number;
  log(x: number): number;
  log1p(x: number): number;
  log10(x: number): number;
  log2(x: number): number;
  round(x: number): number;
  sign(x: number): number;
  sin(x: number): number;
  sinh(x: number): number;
  sqrt(x: number): number;
  tan(x: number): number;
  tanh(x: number): number;
  trunc(x: number): number;
  atan2(y: number, x: number): number;
  hypot(...values: number[]): number;
  imul(x: number, y: number): number;
  max(...values: number[]): number;
  min(...values: number[]): number;
  pow(x: number, y: number): number;
  random(): number;
}
declare var Math: Math;
interface JSON {
  [member: string]: any;
  parse(text: string, reviver?: (this: any, key: string, value: any) => any): any;
  stringify(value: any, replacer?: ((this: any, key: string, value: any) => any) | (number | string)[] | null, space?: string | number): string;
}
declare var JSON: JSON;
declare var Object: any;
interface ArrayConstructor {
  [member: string]: any;
  new <T = any>(...items: any[]): T[];
  <T = any>(...items: any[]): T[];
  isArray(value: any): value is any[];
  from<T = any>(source: any, map?: (value: any, index: number) => T): T[];
  of<T>(...items: T[]): T[];
}
declare var Array: ArrayConstructor;
declare var String: any;
declare var Number: any;
declare var Boolean: any;
declare var Date: any;
declare var RegExp: any;
interface Map<K, V> {
  [member: string]: any;
  readonly size: number;
  get(key: K): V | undefined;
  set(key: K, value: V): this;
  has(key: K): boolean;
  delete(key: K): boolean;
  clear(): void;
  forEach(cb: (value: V, key: K, map: Map<K, V>) => void): void;
  entries(): IterableIterator<[K, V]>;
  keys(): IterableIterator<K>;
  values(): IterableIterator<V>;
  [Symbol.iterator](): IterableIterator<[K, V]>;
}
interface MapConstructor {
  [member: string]: any;
  new <K = any, V = any>(entries?: Iterable<readonly [K, V]> | null): Map<K, V>;
}
declare var Map: MapConstructor;
interface Set<T> {
  [member: string]: any;
  readonly size: number;
  add(value: T): this;
  has(value: T): boolean;
  delete(value: T): boolean;
  clear(): void;
  forEach(cb: (value: T, key: T, set: Set<T>) => void): void;
  entries(): IterableIterator<[T, T]>;
  keys(): IterableIterator<T>;
  values(): IterableIterator<T>;
  [Symbol.iterator](): IterableIterator<T>;
}
interface SetConstructor {
  [member: string]: any;
  new <T = any>(values?: Iterable<T> | null): Set<T>;
}
declare var Set: SetConstructor;
interface WeakMap<K, V> { [member: string]: any; get(key: K): V | undefined; set(key: K, value: V): this; has(key: K): boolean; delete(key: K): boolean; }
interface WeakMapConstructor { [member: string]: any; new <K = any, V = any>(entries?: Iterable<readonly [K, V]> | null): WeakMap<K, V>; }
declare var WeakMap: WeakMapConstructor;
interface WeakSet<T> { [member: string]: any; add(value: T): this; has(value: T): boolean; delete(value: T): boolean; }
interface WeakSetConstructor { [member: string]: any; new <T = any>(values?: Iterable<T> | null): WeakSet<T>; }
declare var WeakSet: WeakSetConstructor;
interface SymbolConstructor {
  [member: string]: any;
  (description?: string | number): symbol;
  readonly iterator: unique symbol;
  readonly asyncIterator: unique symbol;
}
declare var Symbol: SymbolConstructor;
// The iteration protocol: for...of, spread, array destructuring and
// generators need it, and without it valid code read as an error.
interface IteratorYieldResult<T> { done?: false; value: T; }
interface IteratorReturnResult<T> { done: true; value: T; }
type IteratorResult<T, TReturn = any> = IteratorYieldResult<T> | IteratorReturnResult<TReturn>;
interface Iterator<T, TReturn = any, TNext = any> {
  next(...args: [] | [TNext]): IteratorResult<T, TReturn>;
  return?(value?: TReturn): IteratorResult<T, TReturn>;
  throw?(e?: any): IteratorResult<T, TReturn>;
}
interface Iterable<T, TReturn = any, TNext = any> { [Symbol.iterator](): Iterator<T, TReturn, TNext>; }
interface IterableIterator<T, TReturn = any, TNext = any> extends Iterator<T, TReturn, TNext> {
  [Symbol.iterator](): IterableIterator<T, TReturn, TNext>;
}
interface IteratorObject<T, TReturn = unknown, TNext = unknown> extends Iterator<T, TReturn, TNext> {
  [Symbol.iterator](): IteratorObject<T, TReturn, TNext>;
}
type BuiltinIteratorReturn = any;
interface Generator<T = unknown, TReturn = any, TNext = any> extends IteratorObject<T, TReturn, TNext> {
  next(...args: [] | [TNext]): IteratorResult<T, TReturn>;
  return(value: TReturn): IteratorResult<T, TReturn>;
  throw(e: any): IteratorResult<T, TReturn>;
  [Symbol.iterator](): Generator<T, TReturn, TNext>;
}
interface AsyncIterator<T, TReturn = any, TNext = any> {
  next(...args: [] | [TNext]): Promise<IteratorResult<T, TReturn>>;
  return?(value?: any): Promise<IteratorResult<T, TReturn>>;
  throw?(e?: any): Promise<IteratorResult<T, TReturn>>;
}
interface AsyncIterable<T, TReturn = any, TNext = any> { [Symbol.asyncIterator](): AsyncIterator<T, TReturn, TNext>; }
interface AsyncIterableIterator<T, TReturn = any, TNext = any> extends AsyncIterator<T, TReturn, TNext> {
  [Symbol.asyncIterator](): AsyncIterableIterator<T, TReturn, TNext>;
}
interface AsyncIteratorObject<T, TReturn = unknown, TNext = unknown> extends AsyncIterator<T, TReturn, TNext> {
  [Symbol.asyncIterator](): AsyncIteratorObject<T, TReturn, TNext>;
}
interface AsyncGenerator<T = unknown, TReturn = any, TNext = any> extends AsyncIteratorObject<T, TReturn, TNext> {
  next(...args: [] | [TNext]): Promise<IteratorResult<T, TReturn>>;
  return(value: TReturn | PromiseLike<TReturn>): Promise<IteratorResult<T, TReturn>>;
  throw(e: any): Promise<IteratorResult<T, TReturn>>;
  [Symbol.asyncIterator](): AsyncGenerator<T, TReturn, TNext>;
}
declare var Error: any;
declare var parseInt: (string: string, radix?: number) => number;
declare var parseFloat: (string: string) => number;
declare var isNaN: (number: number) => boolean;
declare var isFinite: (number: number) => boolean;
declare var undefined: undefined;
declare var NaN: number;
declare var Infinity: number;
declare function setTimeout(handler: (...args: any[]) => void, timeout?: number, ...args: any[]): number;
declare function clearTimeout(id?: number): void;
declare function setInterval(handler: (...args: any[]) => void, timeout?: number, ...args: any[]): number;
declare function clearInterval(id?: number): void;
declare function fetch(input: any, init?: any): Promise<any>;
declare function structuredClone<T>(value: T, options?: any): T;
`;
