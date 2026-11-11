// SPDX-License-Identifier: Apache-2.0
//
// Which parts a DOCTYPE is a fault in, and what Word 365 does with one elsewhere (CR-033; the portfolio's DTD handling policy,
// sections 3 and 6). ECMA-376 Part 2 §6.2.5 b) forbids DTD declarations in the XML Part 2 defines: the content types, the core
// properties, the digital signature parts and every relationships part. Elsewhere a DOCTYPE is allowed, read by `readDoctype`, and
// said with Word's reading of it, from the bridge rounds dtd-word-check and dtd2-word-check (10 November 2026).
import { ContentTypes } from './ContentTypes.mjs';
import type { DoctypeSummary } from '../xml/doctype.mjs';
import type { DoctypeWarning } from './exceptions.mjs';

/** The parts Word 365 refuses a file for: the main part (each of its four kinds), a header, a footer. */
const WORD_REFUSES: ReadonlySet<string> = new Set([
  ContentTypes.WORDPROCESSINGML_DOCUMENT, ContentTypes.WORDPROCESSINGML_DOCUMENT_MACROENABLED,
  ContentTypes.WORDPROCESSINGML_TEMPLATE, ContentTypes.WORDPROCESSINGML_TEMPLATE_MACROENABLED,
  ContentTypes.WORDPROCESSINGML_HEADER, ContentTypes.WORDPROCESSINGML_FOOTER,
]);

/** The parts Word 365 reads as absent, after its repair prompt. */
const WORD_ABSENT: ReadonlySet<string> = new Set([
  ContentTypes.WORDPROCESSINGML_STYLES, ContentTypes.WORDPROCESSINGML_SETTINGS, ContentTypes.WORDPROCESSINGML_NUMBERING,
  ContentTypes.WORDPROCESSINGML_FOOTNOTES, ContentTypes.WORDPROCESSINGML_ENDNOTES, ContentTypes.WORDPROCESSINGML_COMMENTS,
  ContentTypes.WORDPROCESSINGML_GLOSSARYDOCUMENT, ContentTypes.PACKAGE_COREPROPERTIES,
]);

/** Word 365's reading of a file with a DOCTYPE in a part of this type; a custom XML store (`store`) is read as absent. */
export function wordReadingOf(contentType: string, store = false): DoctypeWarning['word'] {
  if (WORD_REFUSES.has(contentType)) return 'refuses the file';
  if (store || WORD_ABSENT.has(contentType)) return 'reads the part as absent';
  if (contentType === ContentTypes.RELATIONSHIPS_PART) return 'loses every relationship';
  if (contentType === ContentTypes.DIGITAL_SIGNATURE_XML_SIGNATURE_PART) return 'opens it';
  return 'not measured';
}

/** The warning for a part that held a DOCTYPE. */
export function doctypeWarning(partName: string, contentType: string, summary: DoctypeSummary, options: { store?: boolean; repaired?: boolean } = {}): DoctypeWarning {
  const word = wordReadingOf(contentType, options.store === true);
  const repaired = options.repaired === true;
  const reading = repaired
    ? 'read as Word repairs it'
    : `read with ${summary.expanded} entity reference${summary.expanded === 1 ? '' : 's'} expanded and ${summary.defaulted} attribute default${summary.defaulted === 1 ? '' : 's'} applied`;
  const notRead = summary.notRead.length > 0 ? `; not read: ${summary.notRead.join(', ')}` : '';
  return {
    code: 'xml.doctype',
    partName,
    holds: [...summary.holds],
    expanded: summary.expanded,
    defaulted: summary.defaulted,
    notRead: [...summary.notRead],
    word,
    repaired,
    message: `${partName} holds a DOCTYPE (${summary.holds.join(', ')}): ${reading}, nothing outside the package fetched${notRead}. `
      + `Word 365 ${word === 'not measured' ? 'is not measured on one here' : word === 'opens it' ? 'opens such a file' : `${word} for one here`} (CR-033)`,
  };
}

/** A summary for a part whose DOCTYPE wasn't read further than its presence (a part Part 2 defines, read as Word repairs it). */
export function presenceOf(name = ''): DoctypeSummary {
  return { name, holds: ['bare'], expanded: 0, defaulted: 0, notRead: [] };
}
