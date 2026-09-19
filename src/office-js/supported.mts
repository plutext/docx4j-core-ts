// `Word.supported`: the `Class.member` strings this package implements, so that a tool (CR-003's
// console, an agent writing a script) can check an add-in script before running it. The Office JS
// half is generated from `test/office-js-subset.ts`, which is the compile-time promise that those
// members exist and have Office JS's shape; the shim's own surface is listed here.

import { SUBSET_MEMBERS } from './supported.generated.mjs';

/** What every proxied object answers, whatever its class. */
const UNIVERSAL_MEMBERS = ['load', 'track', 'untrack', 'isNullObject', 'toJSON'];

/** The collection members the shim adds to an array of views. */
const COLLECTION_MEMBERS = ['items', 'getFirst', 'getFirstOrNullObject', 'getLast', 'getLastOrNullObject', 'getCount', 'length', ...UNIVERSAL_MEMBERS];

/** The classes a collection is of; phases C, E and G fill in the ones whose views are not here yet. */
const COLLECTION_CLASSES = ['Paragraph', 'Range', 'Table', 'TableRow', 'TableCell', 'ContentControl', 'Comment', 'InlinePicture', 'List', 'ListItem', 'Field', 'TrackedChange', 'CustomXmlPart'];

/** The shim's own members: the context, the document and what it adds over the views. */
const SHIM_MEMBERS = [
  'RequestContext.document', 'RequestContext.sync', 'RequestContext.trackedObjects', 'RequestContext.package_',
  'Document.body', 'Document.properties', 'Document.changeTrackingMode', 'Document.getSelection', 'Document.getComments',
  'Document.getOoxml', 'Document.save', 'Document.package_',
  'DocumentProperties.title', 'DocumentProperties.subject', 'DocumentProperties.author', 'DocumentProperties.comments',
  'DocumentProperties.keywords', 'DocumentProperties.category', 'DocumentProperties.lastAuthor', 'DocumentProperties.revisionNumber',
  'DocumentProperties.creationDate', 'DocumentProperties.lastSaveTime', 'DocumentProperties.lastPrintDate',
  'DocumentProperties.manager', 'DocumentProperties.company', 'DocumentProperties.applicationName', 'DocumentProperties.template',
  'DocumentProperties.security',
  // over the views (src/office-js/extras.mts) and the extensions the views themselves offer
  'Body.getOoxml', 'Body.getXml', 'Body.insertXml', 'Body.insertElement', 'Body.content', 'Body.tables',
  'Body.paragraphAt', 'Body.elementAt', 'Body.addressOf', 'Body.outline', 'Body.getContent',
  'Paragraph.getOoxml', 'Paragraph.getXml', 'Paragraph.paraId', 'Paragraph.runs', 'Paragraph.element', 'Paragraph.p',
  'Paragraph.parentBody', 'Paragraph.index', 'Paragraph.splitAt', 'Paragraph.segments',
  'Range.getOoxml', 'Range.getXml', 'Range.runs', 'Range.paragraph', 'Range.paragraphs', 'Range.start', 'Range.end',
  // lists (CR-002 phase H): the extensions beside the Office JS members the subset declares
  'Body.listLabels', 'Paragraph.restartList', 'Paragraph.numPr', 'Paragraph.pPr',
  'List.numId', 'List.definition', 'List.element', 'List.abstractElement', 'List.level',
  'List.levelExists', 'List.separate', 'List.restart', 'List.body',
  'ListItem.list', 'ListItem.paragraph',
];

function build(): Set<string> {
  const all = new Set<string>(SUBSET_MEMBERS);
  for (const className of new Set([...SUBSET_MEMBERS].map((m) => m.split('.')[0]!))) {
    for (const member of UNIVERSAL_MEMBERS) all.add(`${className}.${member}`);
  }
  for (const member of SHIM_MEMBERS) all.add(member);
  for (const className of COLLECTION_CLASSES) {
    for (const member of COLLECTION_MEMBERS) all.add(`${className}Collection.${member}`);
  }
  return all;
}

/**
 * Every member this package implements, as `Class.member` (`'Body.insertParagraph'`). A member
 * that is not here throws `NotSupportedError` when an add-in calls it through the shim.
 */
export const supported: ReadonlySet<string> = build();

export { SUBSET_MEMBERS };
