// `@docx4j/core-ts/office-js`: the `Word` shim of CR-002 section 3.10.
//
// The content API is a structural subset of Office JS's Word types (section 3.4); this makes that
// executable, so an add-in's `Word.run` callback runs against a package here, in Node or in a
// browser, with the saved docx as the assertion. It is deliberately not part of the root export:
// an add-in bundle that wants only the content API never pulls the proxies in.
//
//   import { Word } from '@docx4j/core-ts/office-js';
//   await Word.run(pkg, async (context) => {
//     const paragraphs = context.document.body.paragraphs;
//     paragraphs.load('text');
//     await context.sync();
//     paragraphs.items[0].font.bold = true;
//   });

import { run, RequestContext, type RunOptions, type TrackedObjects } from './run.mjs';
import { Document, DocumentProperties } from './document.mjs';
import { supported, SUBSET_MEMBERS } from './supported.mjs';
import { ClientResult, unwrap, nullObject, Wrapper, type Collection, type PendingSink } from './proxy.mjs';
import { NotSupportedError, ItemNotFoundError, ValueNotLoadedError } from './errors.mjs';
import { toApiScript, type ApiScriptOptions, type ApiScriptTarget } from './toApiScript.mjs';
import {
  InsertLocation, Alignment, UnderlineType, ChangeTrackingMode, BreakType, ContentControlType, Style, ErrorCodes, SearchOptions,
  type InsertLocationValue, type ChangeTrackingModeValue, type BreakTypeValue, type ContentControlTypeValue, type StyleValue,
  type AlignmentValue, type UnderlineTypeValue,
} from './enums.mjs';

export { run, RequestContext, Document, DocumentProperties, ClientResult, unwrap, nullObject, Wrapper, toApiScript, supported, SUBSET_MEMBERS };
export { NotSupportedError, ItemNotFoundError, ValueNotLoadedError };
export { InsertLocation, Alignment, UnderlineType, ChangeTrackingMode, BreakType, ContentControlType, Style, ErrorCodes, SearchOptions };
export type {
  RunOptions, TrackedObjects, Collection, PendingSink, ApiScriptOptions, ApiScriptTarget,
  InsertLocationValue, ChangeTrackingModeValue, BreakTypeValue, ContentControlTypeValue, StyleValue, AlignmentValue, UnderlineTypeValue,
};

/**
 * The `Word` namespace an add-in script expects: `Word.run`, the enums, and `Word.supported`, the
 * `Class.member` strings this package implements (so a tool can check a script before running it).
 * `Word.run` takes the package first, since there is no host document.
 */
export const Word = Object.freeze({
  run,
  supported,
  InsertLocation,
  Alignment,
  UnderlineType,
  ChangeTrackingMode,
  BreakType,
  ContentControlType,
  Style,
  ErrorCodes,
  SearchOptions,
  ClientResult,
  RequestContext,
  Document,
  DocumentProperties,
  NotSupportedError,
  ItemNotFoundError,
  ValueNotLoadedError,
  unwrap,
});
