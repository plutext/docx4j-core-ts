// The content API (CR-002 phases B, C, D, F and G): Office JS shapes over the docx4j tree.
export { Body, type Address, type BlockElement, type Outline, type OutlineParagraph, type OutlineTable } from './Body.mjs';
export { Paragraph, type Alignment } from './Paragraph.mjs';
export { Range } from './Range.mjs';
export { BUILT_IN_STYLES, builtInOf, idOfBuiltIn, displayNameOf, styleNameOf, styleIdOf, type BuiltInStyle } from './styles.mjs';
export { Comment } from './Comment.mjs';
// lists (CR-002 phase H)
export {
  List, ListItem, ListItemNotFoundError, listsOf, listLabelsOf, numberingEmulatorOf, levelOf,
  type ListLevelType, type ListNumbering, type ListBullet, type StartListOptions, type ListLabel,
} from './List.mjs';
export { type Author, type CommentParts, type CommentPartsAccess, type CommentsExtensible, type CommentMarker, COMMENT_TEXT_STYLE, COMMENT_REFERENCE_STYLE, markersOf, markersOfParagraph, commentIdsOf } from './comments.mjs';
export { Font, type UnderlineType, type RPrHolder, type FontTracking } from './Font.mjs';
export { Table, TableRow, TableCell, cellOf } from './Table.mjs';
export { InlinePicture, imageInfoOf, naturalSizeEmu, addImage, writableWidthEmu, type ImageFormat, type ImageInfo, type InlinePictureOptions, type NewPicture } from './InlinePicture.mjs';
export { ContentControl, collectControls, collectRunControls, W14_NS, W15_NS, type ContentControlType, type ContentControlForm, type ContentControlAppearance } from './ContentControl.mjs';
export { contentOf, isFlatOpc, rewriteRelationshipIds, type OoxmlOptions } from './ooxml.mjs';
export { type SearchOptions, searchPattern } from './search.mjs';
export { segmentsOf, runsOf, runItemsOf, childrenOf, rowsOf, cellsOf, BLOCK_LEVEL_TYPES, SDT_TYPES, runOf, paragraphOf, textOfView, revisionKindOf, type TextSegment, type Located, type TextView, type TextViewOptions, type RevisionKind, type RevisionHolder } from './tree.mjs';
// change tracking (CR-002 phase F)
export { TrackedChange, trackedChangesOfParagraph, trackedChangesOfRow, joinWithNext, type TrackedChangeType, type TrackedChangeTarget } from './TrackedChange.mjs';
export {
  ChangeTracker, trackerOf, calendarOf, dateOf, copyRPr, markDeleted, markInserted,
  toDeletedText, toRestoredText, restoreRPr, restorePPr,
  trackInsertedParagraph, trackInsertedTable, wrapNewRuns, paraRPrOf, rowPrOf, pruneParagraphProperties,
  type ChangeTrackingMode, type TrackingHost,
} from './tracking.mjs';
// The objects package's builders, re-exported so that one import serves the content API and the tree it works on.
export { wml, wmlOne, p, r, t, br, tab, tbl, tr, tc, sdt, sdtPr, nextSdtId, sdtProperty, sdtKindOf, inlinePicture, rPrToElements, rPrFromElements, textOf, walk, walkAll, find, linkParents, applyRunOptions, readRunOptions, UNDERLINE_TO_WML, isElement, typeNameOf, W_NS, type Element, type Wrapper, type WmlOptions, type Raw, type Interpolated, type RunOptions, type RunFormatting, type ParagraphOptions, type TableOptions, type CellOptions, type SdtKind, type SdtForm, type SdtOptions } from '@docx4j/generated-objects-ts/builders/wml';
