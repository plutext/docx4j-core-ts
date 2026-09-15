// `Word.run(pkg, fn)` and its RequestContext (CR-002 section 3.10).
//
// Office JS batches proxy calls and resolves them on `context.sync()`. Here the tree is in memory,
// so a call acts at once (CR-002 section 3.1) and `sync()` only resolves what was asynchronous:
// the marshalling and unmarshalling calls, whose `ClientResult.value` it fills. An add-in script
// written for Word therefore runs unchanged, `load` / `sync` lines and all.

import type { WordprocessingMLPackage } from '../packages/WordprocessingMLPackage.mjs';
import { Document, type RunOptions } from './document.mjs';
import { EXTRA_MEMBERS } from './extras.mjs';
import { Wrapper, unwrap, type PendingSink } from './proxy.mjs';

export type { RunOptions };

/** Office JS's `context.trackedObjects`: object lifetime is the garbage collector's business here. */
export interface TrackedObjects {
  add<T>(object: T): T;
  remove<T>(object: T): T;
}

/** A subset of Office JS `Word.RequestContext` over a package. */
export class RequestContext implements PendingSink {
  readonly document: Document;
  /** Tracking is a no-op: nothing here is a remote proxy. */
  readonly trackedObjects: TrackedObjects = { add: (o) => o, remove: (o) => o };
  private pending: Promise<unknown>[] = [];

  constructor(
    /** The package this context is over (extension: Office JS has a host document). */
    readonly package_: WordprocessingMLPackage,
    options: RunOptions = {},
  ) {
    this.document = new Document(package_, options);
  }

  /** @internal a promise whose result the next `sync()` must resolve. */
  enqueue(promise: Promise<unknown>): void {
    this.pending.push(promise);
  }

  /** Resolves the pending results, filling every `ClientResult.value` handed out since the last sync. */
  async sync<T>(value?: T): Promise<T | RequestContext> {
    while (this.pending.length > 0) {
      const batch = this.pending;
      this.pending = [];
      await Promise.all(batch);
    }
    return value === undefined ? this : value;
  }
}

/** Unmarshals what the document-level members read synchronously, as Office JS's host has already. */
async function prepare(pkg: WordprocessingMLPackage, options: RunOptions): Promise<void> {
  const main = pkg.getMainDocumentPart();
  await main.getContents();
  if (options.unmarshalSideParts === false) return;
  const sideParts = [pkg.docPropsCorePart, pkg.docPropsExtendedPart, main.documentSettingsPart];
  for (const part of sideParts) {
    if (part && !part.isUnmarshalled) {
      try {
        await part.getContents();
      } catch {
        // A side part that will not unmarshal is not worth failing the batch for; reads report nothing.
      }
    }
  }
}

/**
 * Runs an add-in style callback against a package: `Word.run(pkg, async (context) => { ... })`.
 * Office JS's `Word.run(batch)` takes the host document; there is none here, so the package comes
 * first. The callback gets a proxied `RequestContext`; a final `sync()` resolves anything it left
 * pending, and the callback's value is the result.
 */
export async function run<T>(
  pkg: WordprocessingMLPackage,
  batch: (context: RequestContext) => Promise<T> | T,
  options: RunOptions = {},
): Promise<T> {
  await prepare(pkg, options);
  const context = new RequestContext(pkg, options);
  const wrapper = new Wrapper(context, EXTRA_MEMBERS);
  const result = await batch(wrapper.object(context, 'RequestContext'));
  await context.sync();
  return unwrap(result);
}
