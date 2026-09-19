// CR-001 Phase B step 2: property resolution (docx4j `org.docx4j.model.PropertyResolver` and
// the catalogue-driven half of `org.docx4j.model.styles`).
//
// The tables' leaf merge and emptiness rules stay inside `catalogue.mts`: what a consumer
// needs is the tables, the generic operations over them, and the resolver.
export {
  type Property, type RunProps, type Merge,
  RUN, PARAGRAPH, TABLE, CELL, TOGGLES, TOGGLE_NAMES, IND,
  RUN_EXCLUDED, PARAGRAPH_EXCLUDED, TABLE_EXCLUDED, CELL_EXCLUDED,
  namedProperty, isEmptyOf, hasDirectFormattingOf, applyCatalogue, applyProperty,
  unsetCatalogue, copyLeaf,
} from './catalogue.mjs';
export * from './styleUtil.mjs';
export { NumberingLevels, type StyleLookup } from './numberingInd.mjs';
export {
  PropertyResolver, headingLevelByName, WORD_DEFAULT_CELL_MARGIN_TWIPS, type ResolverSource,
} from './PropertyResolver.mjs';
export { setLogger, type Logger } from './log.mjs';
