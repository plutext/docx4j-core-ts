// The proxies, collections and results of the Word shim (CR-002 section 3.10).
//
// The views of `src/model/content/` stay plain objects for direct users (section 3.12); the shim
// is the only thing that wraps them. A proxy does four things: `load` / `track` / `untrack` are
// no-ops returning the object, a member the views implement passes through (its result wrapped in
// turn), an asynchronous member returns a `ClientResult` whose `value` the next `context.sync()`
// fills, and anything else throws `NotSupportedError` naming the class and the member.

import { NotSupportedError, ItemNotFoundError, ValueNotLoadedError } from './errors.mjs';
import type { ExtraMembers } from './extras.mjs';

/** The wrapped object's target, for `unwrap`. */
const RAW = Symbol('docx4j.officeJs.raw');

/** Marks an array that is already a collection, so that it is not wrapped a second time. */
const COLLECTION = Symbol('docx4j.officeJs.collection');

/**
 * The result of an asynchronous call, as Office JS hands one out: `value` after `context.sync()`.
 * It is also awaitable (an extension), so that `await body.insertXml(...)`, which is how the
 * content API and `toApiScript`'s output read, keeps its order through the shim.
 */
export class ClientResult<T = unknown> {
  private loaded = false;
  private result: T | undefined;
  private promise: Promise<T> | undefined;

  constructor(
    /** The call this is the result of ('getOoxml'), for the error message. */
    readonly member: string = 'value',
  ) {}

  /** @internal the underlying call, which `context.sync()` awaits and `then` delegates to. */
  attach(promise: Promise<T>): void {
    this.promise = promise;
  }

  /** Awaiting the result resolves to the value, without waiting for the next sync. */
  then<R1 = T, R2 = never>(
    onFulfilled?: ((value: T) => R1 | PromiseLike<R1>) | null,
    onRejected?: ((reason: unknown) => R2 | PromiseLike<R2>) | null,
  ): Promise<R1 | R2> {
    const promise = this.promise ?? Promise.resolve(this.value);
    return promise.then(onFulfilled, onRejected);
  }

  /** The value; throws until `context.sync()` has resolved it, as Office JS does. */
  get value(): T {
    if (!this.loaded) throw new ValueNotLoadedError(`The result of ${this.member}() is not loaded yet; await context.sync() first`);
    return this.result as T;
  }

  /** @internal filled by the context when the underlying promise settles. */
  resolveWith(value: T): void {
    this.result = value;
    this.loaded = true;
  }

  /** True once `context.sync()` has filled it. */
  get isLoaded(): boolean {
    return this.loaded;
  }

  toJSON(): { value: T | undefined } {
    return { value: this.loaded ? (this.result as T) : undefined };
  }
}

/** A collection, as the subset types describe one: an array that also answers Office JS's collection members. */
export interface Collection<T> extends Array<T> {
  /** Office JS's collection property; the same array. */
  readonly items: T[];
  getFirst(): T;
  getFirstOrNullObject(): T;
  getLast(): T;
  getLastOrNullObject(): T;
  getCount(): ClientResult<number>;
}

/** What a context must offer the wrapper: somewhere to park the promise of an asynchronous call. */
export interface PendingSink {
  /** @internal */
  enqueue(promise: Promise<unknown>): void;
}

/** The item class name behind a member that returns a collection, for empty ones and for messages. */
const COLLECTION_ITEM_CLASS: Readonly<Record<string, string>> = {
  paragraphs: 'Paragraph',
  tables: 'Table',
  rows: 'TableRow',
  cells: 'TableCell',
  contentControls: 'ContentControl',
  inlinePictures: 'InlinePicture',
  lists: 'List',
  fields: 'Field',
  search: 'Range',
  getComments: 'Comment',
  comments: 'Comment',
  replies: 'Comment',
  getTrackedChanges: 'TrackedChange',
  customXmlParts: 'CustomXmlPart',
};

/** Members whose promise is the caller's to await, not a ClientResult: `context.sync()`. */
const RAW_PROMISE_MEMBERS: ReadonlySet<string> = new Set(['sync']);

/** The engine behind the views: handed out as it is, never proxied, so that docx4j code works on it. */
const RAW_VALUE_MEMBERS: ReadonlySet<string> = new Set(['package_', 'part']);

/** Objects that are data, not views: never proxied, handed out as they are. */
const PLAIN_CONSTRUCTORS: ReadonlySet<unknown> = new Set<unknown>([
  Object, Array, Date, RegExp, Map, Set, WeakMap, WeakSet, Promise, Error, Uint8Array, ArrayBuffer, Function,
]);

/** True for an instance of a class of ours (a view, or one of the shim's own objects). */
function isView(value: object): boolean {
  const proto: unknown = Object.getPrototypeOf(value);
  if (proto === null || proto === Object.prototype) return false;
  const ctor = (value as { constructor?: unknown }).constructor;
  return typeof ctor === 'function' && !PLAIN_CONSTRUCTORS.has(ctor);
}

function isThenable(value: unknown): value is Promise<unknown> {
  return typeof value === 'object' && value !== null && typeof (value as { then?: unknown }).then === 'function';
}

/** The target behind a proxy the shim handed out; anything else is returned as it is. */
export function unwrap<T>(value: T): T {
  if (typeof value === 'object' && value !== null) {
    const raw = (value as Record<symbol, unknown>)[RAW];
    if (raw !== undefined) return raw as T;
  }
  return value;
}

/**
 * A null object: `isNullObject` is true and every other member throws, as Office JS's
 * `*OrNullObject` results do.
 */
export function nullObject<T>(className: string): T {
  const fail = (member: string): never => {
    throw new ItemNotFoundError(`Word.${className} is a null object (nothing matched); check isNullObject before using ${member}`);
  };
  const proxy: unknown = new Proxy({ isNullObject: true } as object, {
    get(target, prop): unknown {
      if (typeof prop === 'symbol') return Reflect.get(target, prop);
      if (prop === 'isNullObject') return true;
      if (prop === 'toJSON') return () => ({ isNullObject: true });
      if (prop === 'then') return undefined;
      if (prop === 'load' || prop === 'track' || prop === 'untrack') return () => proxy;
      return fail(prop);
    },
    set(_target, prop): boolean {
      return fail(String(prop));
    },
  });
  return proxy as T;
}

/**
 * Wraps the views of one `Word.run` in proxies. One wrapper per context, so that the same view
 * gives the same proxy and object identity holds within a batch, as it does in Office JS.
 */
export class Wrapper {
  private readonly cache = new WeakMap<object, unknown>();

  constructor(
    private readonly context: PendingSink,
    /** Members the shim adds over a view (`getOoxml`), consulted only when the view has none. */
    private readonly extras: ExtraMembers = {},
  ) {}

  /** Wraps whatever a member returned: a view, an array of views, a promise, or a plain value. */
  value(value: unknown, member = ''): unknown {
    if (value === null || typeof value !== 'object') return value;
    if (value instanceof ClientResult) return value;
    if (Array.isArray(value)) return (value as unknown as Record<symbol, unknown>)[COLLECTION] === true ? value : this.collection(value, member);
    if (!isView(value)) return value;
    return this.object(value);
  }

  /** The proxy over one view. */
  object<T extends object>(view: T, className = view.constructor.name): T {
    const target = unwrap(view);
    const cached = this.cache.get(target);
    if (cached !== undefined) return cached as T;
    const wrapper = this;
    const proxy: T = new Proxy(target, {
      get(obj, prop, receiver): unknown {
        if (typeof prop === 'symbol') return prop === RAW ? obj : Reflect.get(obj, prop, obj);
        switch (prop) {
          case 'load':
          case 'retrieve':
          case 'track':
          case 'untrack':
            if (!(prop in obj)) return () => proxy;
            break;
          case 'isNullObject':
            if (!(prop in obj)) return false;
            break;
          case 'then':
            if (!(prop in obj)) return undefined;
            break;
          default:
            break;
        }
        if (!(prop in obj)) {
          const extra = wrapper.extras[className]?.[prop];
          if (extra === undefined) throw new NotSupportedError(className, prop);
          const call = extra as (target: unknown, ...args: unknown[]) => unknown;
          return (...args: unknown[]): unknown => wrapper.result(call(obj, ...args.map(unwrap)), prop);
        }
        const member: unknown = Reflect.get(obj, prop, obj);
        if (typeof member === 'function') {
          if (RAW_PROMISE_MEMBERS.has(prop)) return (...args: unknown[]): unknown => (member as (...a: unknown[]) => unknown).apply(obj, args.map(unwrap));
          return (...args: unknown[]): unknown => wrapper.result((member as (...a: unknown[]) => unknown).apply(obj, args.map(unwrap)), prop);
        }
        void receiver;
        return RAW_VALUE_MEMBERS.has(prop) ? member : wrapper.value(member, prop);
      },
      set(obj, prop, value): boolean {
        if (typeof prop === 'symbol') return Reflect.set(obj, prop, value, obj);
        if (!(prop in obj)) throw new NotSupportedError(className, prop);
        return Reflect.set(obj, prop, unwrap(value), obj);
      },
    });
    this.cache.set(target, proxy);
    return proxy;
  }

  /** The result of a call: a promise becomes a ClientResult the next sync() fills. */
  result(value: unknown, member: string): unknown {
    if (value instanceof ClientResult || !isThenable(value)) return this.value(value, member);
    const result = new ClientResult(member);
    const settled = value.then((resolved) => {
      const wrapped = this.value(resolved, member);
      result.resolveWith(wrapped);
      return wrapped;
    });
    result.attach(settled);
    this.context.enqueue(settled);
    return result;
  }

  /** An array of views as an Office JS collection: the array, plus `items` and the getters. */
  collection<T>(items: T[], member = ''): T[] {
    if (items.length > 0 && !items.every((i) => typeof i === 'object' && i !== null && isView(i as object))) return items;
    const itemClass = items.length > 0 ? (items[0] as object).constructor.name : (COLLECTION_ITEM_CLASS[member] ?? 'Object');
    const wrapped = items.map((i) => this.value(i, member)) as T[];
    const className = `${itemClass}Collection`;
    const collection = wrapped as unknown as Collection<T>;
    const define = (name: string, value: unknown): void => {
      Object.defineProperty(collection, name, { value, enumerable: false, configurable: true, writable: true });
    };
    define(COLLECTION as unknown as string, true);
    define('items', wrapped);
    define('getFirst', () => {
      if (wrapped.length === 0) throw new ItemNotFoundError(`Word.${className} is empty; getFirst() found nothing (getFirstOrNullObject() does not throw)`);
      return wrapped[0];
    });
    define('getFirstOrNullObject', () => (wrapped.length === 0 ? nullObject<T>(itemClass) : wrapped[0]));
    define('getLast', () => {
      if (wrapped.length === 0) throw new ItemNotFoundError(`Word.${className} is empty; getLast() found nothing (getLastOrNullObject() does not throw)`);
      return wrapped[wrapped.length - 1];
    });
    define('getLastOrNullObject', () => (wrapped.length === 0 ? nullObject<T>(itemClass) : wrapped[wrapped.length - 1]));
    define('getCount', () => {
      const result = new ClientResult<number>('getCount');
      result.resolveWith(wrapped.length);
      return result;
    });
    define('load', () => collectionProxy);
    define('track', () => collectionProxy);
    define('untrack', () => collectionProxy);
    const wrapperSelf = this;
    const collectionProxy = new Proxy(collection, {
      get(obj, prop): unknown {
        if (typeof prop === 'symbol') return prop === RAW ? obj : Reflect.get(obj, prop, obj);
        if (prop === 'then' && !(prop in obj)) return undefined;
        if (prop === 'isNullObject') return false;
        if (!(prop in obj)) throw new NotSupportedError(className, prop);
        const member2: unknown = Reflect.get(obj, prop, obj);
        if (typeof member2 === 'function' && !/^\d+$/.test(prop)) {
          return (...args: unknown[]): unknown => wrapperSelf.result((member2 as (...a: unknown[]) => unknown).apply(obj, args.map(unwrap)), prop);
        }
        return member2;
      },
    }) as unknown as T[];
    return collectionProxy;
  }
}
