// The content API (CR-002 phases B, C and D): Office JS shapes over the docx4j tree.
export { Body, type Address, type BlockElement, type Outline, type OutlineParagraph, type OutlineTable } from './Body.mjs';
export { Paragraph, type Alignment } from './Paragraph.mjs';
export { Range } from './Range.mjs';
export { BUILT_IN_STYLES, builtInOf, idOfBuiltIn, displayNameOf, styleNameOf, styleIdOf, type BuiltInStyle } from './styles.mjs';
export { Comment } from './Comment.mjs';
export { type Author, type CommentParts, type CommentPartsAccess, type CommentsExtensible, type CommentMarker, COMMENT_TEXT_STYLE, COMMENT_REFERENCE_STYLE, markersOf, markersOfParagraph, commentIdsOf } from './comments.mjs';
export { Font, type UnderlineType, type RPrHolder } from './Font.mjs';
export { Table, TableRow, TableCell, cellOf } from './Table.mjs';
export { InlinePicture, imageInfoOf, naturalSizeEmu, drawingFor, addImage, writableWidthEmu, type ImageFormat, type ImageInfo, type InlinePictureOptions, type NewPicture } from './InlinePicture.mjs';
export { ContentControl, collectControls, collectRunControls, type ContentControlType, type ContentControlForm } from './ContentControl.mjs';
export { contentOf, isFlatOpc, rewriteRelationshipIds, type OoxmlOptions } from './ooxml.mjs';
export { type SearchOptions, searchPattern } from './search.mjs';
export { segmentsOf, runsOf, runItemsOf, childrenOf, rowsOf, cellsOf, BLOCK_LEVEL_TYPES, SDT_TYPES, runOf, paragraphOf, type TextSegment, type Located } from './tree.mjs';
// The objects package's builders, re-exported so that one import serves the content API and the tree it works on.
export { wml, wmlOne, p, r, t, br, tab, tbl, textOf, walk, find, linkParents, applyRunOptions, readRunOptions, UNDERLINE_TO_WML, isElement, typeNameOf, W_NS, type Element, type Wrapper, type WmlOptions, type Raw, type Interpolated, type RunOptions, type RunFormatting, type ParagraphOptions, type TableOptions } from '@docx4j/generated-objects-ts/builders/wml';
