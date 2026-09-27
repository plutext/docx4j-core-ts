// CR-002 section 22.1: the type of what the shim's proxies hand out.
//
// The shim is more permissive than the views it wraps: a proxy adds `load` / `track` / `untrack`,
// and an array comes back as a `Collection<T>` with `items`, `getFirst` and the rest (`proxy.mts`).
// `Document.body` is typed as the plain `Body`, though, whose `paragraphs` is `Paragraph[]`, so a
// type check of add-in code against this package rejects `paragraphs.items` and
// `paragraphs.load('items/text')` - which the shim accepts at runtime. `Shimmed<T>` is the view the
// proxies really present, so a Node user of `Word.run` and an add-in test can type the same script.
//
// The editor carried this mapped type itself (ED-003 section 11.15 item 1) and deletes its copy now.
import type { ClientResult, Collection } from './proxy.mjs';

/** Office JS's `load` / `track` / `untrack`, which every proxy answers and which return the object. */
export interface Loadable<T> {
  /**
   * Office JS's `load`: a no-op here, since the tree is in memory and nothing is a remote proxy
   * (CR-002 section 3.10). Returns the object, so `body.load('text')` chains as it does in Word.
   */
  load(option?: string | string[] | Record<string, unknown>): T;
  track(): T;
  untrack(): T;
}

/**
 * What the shim presents for a value of type `T`:
 *
 * - a function becomes a function whose result is shimmed in turn (a method call on a proxy hands
 *   back another proxy), except that a promise stays a promise and a `ClientResult` stays one;
 * - an array becomes a `Collection` of shimmed items - Office JS's `items`, `getFirst`,
 *   `getFirstOrNullObject`, `getLast`, `getLastOrNullObject` and `getCount`;
 * - any other object becomes itself with its members shimmed, plus `load`, `track` and `untrack`;
 * - a primitive is itself.
 *
 * Deliberately **not** distributive over unions of objects beyond the array case: a `Range | undefined`
 * shims to `Shimmed<Range> | undefined`, which is what a caller wants, and the conditional's natural
 * distribution gives that.
 */
export type Shimmed<T> =
  T extends ClientResult<unknown> ? T
  : T extends Promise<unknown> ? T
  : T extends (...args: infer A) => infer R ? ((...args: A) => Shimmed<R>) & Loadable<T>
  : T extends readonly (infer I)[] ? ShimmedCollection<I>
  : T extends object ? ShimmedObject<T>
  : T;

/** A collection of shimmed items, as the proxies wrap an array. */
export type ShimmedCollection<I> = Collection<Shimmed<I>> & Loadable<Collection<Shimmed<I>>>;

/**
 * An object with every member shimmed, and Office JS's `load` / `track` / `untrack` beside them.
 *
 * `T &` matters and is not redundant: a mapped type keeps only the public members, so without it a
 * shimmed `Body` is not assignable to `Body` and a consumer cannot pass `context.document.body` to
 * its own function typed on the plain view - which is exactly what an add-in that shares one `edit`
 * between Word and the shim does (`examples/office-addin/taskpane.ts`). Intersecting keeps that
 * assignable while the mapped half still narrows an array member to a `Collection`.
 *
 * The **order** matters too, and is the subtler half: an intersection of two call signatures is an
 * overload set, and TypeScript picks the first that matches. With `T` first, `getRange()` resolves
 * to `T`'s own signature and hands back a plain `Range`, so `paragraph.getRange().font.load(...)`
 * is rejected although the shim allows it. The mapped half therefore comes first, so the shimmed
 * return type wins, and `T` follows to keep assignability.
 */
export type ShimmedObject<T> = { [K in keyof T]: Shimmed<T[K]> } & T & Loadable<T>;
