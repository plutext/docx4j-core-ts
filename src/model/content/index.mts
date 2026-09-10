// The content API (CR-002 phases B and D): Office JS shapes over the docx4j tree.
export { Body, type Address, type BlockElement, type Outline, type OutlineParagraph, type OutlineTable } from './Body.mjs';
export { Paragraph, type Alignment } from './Paragraph.mjs';
export { Range } from './Range.mjs';
export { Font, type UnderlineType, type RPrHolder } from './Font.mjs';
export { type SearchOptions, searchPattern } from './search.mjs';
export { segmentsOf, runsOf, runItemsOf, childrenOf, BLOCK_LEVEL_TYPES, runOf, paragraphOf, type TextSegment } from './tree.mjs';
// The objects package's builders, re-exported so that one import serves the content API and the tree it works on.
export { wml, wmlOne, p, r, t, br, tab, tbl, textOf, walk, find, linkParents, applyRunOptions, readRunOptions, UNDERLINE_TO_WML, isElement, typeNameOf, W_NS, type Element, type Wrapper, type WmlOptions, type Raw, type Interpolated, type RunOptions, type RunFormatting, type ParagraphOptions, type TableOptions } from '@docx4j/generated-objects-ts/builders/wml';
