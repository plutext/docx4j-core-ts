// The outline's reader (CR-012 D3): the text model's, which this module can't import, since the text model
// imports the content API. The text model registers it (`model/text/outline.mts`), and the package classes
// load that module, as the comment parts register themselves (CR-002 phase G).
import { Docx4JException } from '../../opc/exceptions.mjs';
import type { Body, Outline } from './Body.mjs';

/** What `Body.outline` asks of the text model: the outline of a body, or of its top-level blocks at these indices, under these options and context. */
export type OutlineReader = (body: Body, options: unknown, context: unknown, blocks?: readonly number[]) => Promise<Outline>;

let reader: OutlineReader | undefined;

export function setOutlineReader(implementation: OutlineReader): void {
  reader = implementation;
}

export function outlineReader(): OutlineReader {
  if (!reader) throw new Docx4JException('The text model is not loaded: import @docx4j/core-ts, @docx4j/core-ts/model or @docx4j/core-ts/packages');
  return reader;
}
