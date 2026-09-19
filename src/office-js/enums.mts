// The Word enums as objects of the string values this API accepts (CR-002 section 3.10), so that
// add-in code written as `paragraph.alignment = Word.Alignment.centered` runs unchanged. Office JS
// declares them as TypeScript enums whose values are the strings; a frozen object of the same
// keys and values is the JavaScript half of that, and the value types below are the API's.

import type { Alignment as AlignmentValue } from '../model/content/Paragraph.mjs';
import type { UnderlineType as UnderlineTypeValue } from '../model/content/Font.mjs';

export type { AlignmentValue, UnderlineTypeValue };

/** Where an insertion goes, as Office JS Word.InsertLocation; a member takes the subset it supports. */
export const InsertLocation = Object.freeze({
  before: 'Before',
  after: 'After',
  start: 'Start',
  end: 'End',
  replace: 'Replace',
} as const);
export type InsertLocationValue = (typeof InsertLocation)[keyof typeof InsertLocation];

/** Word.Alignment; 'Unknown' is what a paragraph without w:jc reports. */
export const Alignment = Object.freeze({
  unknown: 'Unknown',
  left: 'Left',
  centered: 'Centered',
  right: 'Right',
  justified: 'Justified',
} as const);

/** Word.UnderlineType, the subset w:u expresses (the objects package's UNDERLINE_TO_WML). */
export const UnderlineType = Object.freeze({
  mixed: 'Mixed',
  none: 'None',
  single: 'Single',
  word: 'Word',
  double: 'Double',
  thick: 'Thick',
  dotted: 'Dotted',
  dottedHeavy: 'DottedHeavy',
  dashLine: 'DashLine',
  dashLineHeavy: 'DashLineHeavy',
  dashLineLong: 'DashLineLong',
  dashLineLongHeavy: 'DashLineLongHeavy',
  dotDashLine: 'DotDashLine',
  dotDashLineHeavy: 'DotDashLineHeavy',
  twoDotDashLine: 'TwoDotDashLine',
  twoDotDashLineHeavy: 'TwoDotDashLineHeavy',
  wave: 'Wave',
  waveHeavy: 'WaveHeavy',
  waveDouble: 'WaveDouble',
} as const);

/** Word.ChangeTrackingMode; the tracking itself is CR-002 phase F, the setting is read and written here. */
export const ChangeTrackingMode = Object.freeze({
  off: 'Off',
  trackAll: 'TrackAll',
  trackMineOnly: 'TrackMineOnly',
} as const);
export type ChangeTrackingModeValue = (typeof ChangeTrackingMode)[keyof typeof ChangeTrackingMode];

/** Word.BreakType; `insertBreak` here takes 'Page' and 'Line' (CR-002 phase B). */
export const BreakType = Object.freeze({
  page: 'Page',
  next: 'Next',
  sectionNext: 'SectionNext',
  sectionContinuous: 'SectionContinuous',
  sectionEven: 'SectionEven',
  sectionOdd: 'SectionOdd',
  line: 'Line',
} as const);
export type BreakTypeValue = (typeof BreakType)[keyof typeof BreakType];

/** Word.ListLevelType: what a list level paints in front of its paragraphs (CR-002 phase H). */
export const ListLevelType = Object.freeze({
  bullet: 'Bullet',
  number: 'Number',
  picture: 'Picture',
} as const);
export type ListLevelTypeValue = (typeof ListLevelType)[keyof typeof ListLevelType];

/** Word.ListNumbering: the number formats `list.setLevelNumbering` takes. */
export const ListNumbering = Object.freeze({
  none: 'None',
  arabic: 'Arabic',
  upperRoman: 'UpperRoman',
  lowerRoman: 'LowerRoman',
  upperLetter: 'UpperLetter',
  lowerLetter: 'LowerLetter',
} as const);
export type ListNumberingValue = (typeof ListNumbering)[keyof typeof ListNumbering];

/** Word.ListBullet: the bullets `list.setLevelBullet` takes. */
export const ListBullet = Object.freeze({
  custom: 'Custom',
  solid: 'Solid',
  hollow: 'Hollow',
  square: 'Square',
  diamonds: 'Diamonds',
  arrow: 'Arrow',
  checkmark: 'Checkmark',
} as const);
export type ListBulletValue = (typeof ListBullet)[keyof typeof ListBullet];

/** Word.ContentControlType: the kinds w:sdtPr types (CR-002 section 3.5; the views are phase C/E). */
export const ContentControlType = Object.freeze({
  unknown: 'Unknown',
  richText: 'RichText',
  plainText: 'PlainText',
  picture: 'Picture',
  datePicker: 'DatePicker',
  comboBox: 'ComboBox',
  dropDownList: 'DropDownList',
  checkBox: 'CheckBox',
  repeatingSection: 'RepeatingSection',
  repeatingSectionItem: 'RepeatingSectionItem',
  group: 'Group',
  buildingBlockGallery: 'BuildingBlockGallery',
  citation: 'Citation',
  bibliography: 'Bibliography',
  equation: 'Equation',
} as const);
export type ContentControlTypeValue = (typeof ContentControlType)[keyof typeof ContentControlType];

/**
 * Word.Style: the values `styleBuiltIn` reads and accepts ('Heading1'; 'Other' is read for a style
 * that is not built in and cannot be set). `style` is the display name ('Heading 1'); the id is the
 * views' `styleId` extension. The list agrees with the content API's `BUILT_IN_STYLES` (tested).
 */
export const Style = Object.freeze({
  other: 'Other',
  normal: 'Normal',
  heading1: 'Heading1',
  heading2: 'Heading2',
  heading3: 'Heading3',
  heading4: 'Heading4',
  heading5: 'Heading5',
  heading6: 'Heading6',
  heading7: 'Heading7',
  heading8: 'Heading8',
  heading9: 'Heading9',
  toc1: 'Toc1',
  toc2: 'Toc2',
  toc3: 'Toc3',
  toc4: 'Toc4',
  toc5: 'Toc5',
  toc6: 'Toc6',
  toc7: 'Toc7',
  toc8: 'Toc8',
  toc9: 'Toc9',
  footnoteText: 'FootnoteText',
  header: 'Header',
  footer: 'Footer',
  caption: 'Caption',
  footnoteReference: 'FootnoteReference',
  endnoteReference: 'EndnoteReference',
  endnoteText: 'EndnoteText',
  title: 'Title',
  subtitle: 'Subtitle',
  hyperlink: 'Hyperlink',
  strong: 'Strong',
  emphasis: 'Emphasis',
  noSpacing: 'NoSpacing',
  listParagraph: 'ListParagraph',
  quote: 'Quote',
  intenseQuote: 'IntenseQuote',
  subtleEmphasis: 'SubtleEmphasis',
  intenseEmphasis: 'IntenseEmphasis',
  subtleReference: 'SubtleReference',
  intenseReference: 'IntenseReference',
  bookTitle: 'BookTitle',
  bibliography: 'Bibliography',
  tocHeading: 'TocHeading',
  tableGrid: 'TableGrid',
  plainTable1: 'PlainTable1',
  plainTable2: 'PlainTable2',
  plainTable3: 'PlainTable3',
  plainTable4: 'PlainTable4',
  plainTable5: 'PlainTable5',
  gridTable1Light: 'GridTable1Light',
  listTable1Light: 'ListTable1Light',
} as const);
export type StyleValue = (typeof Style)[keyof typeof Style];

/** Word.ErrorCodes, the `code` an error carries. */
export const ErrorCodes = Object.freeze({
  accessDenied: 'AccessDenied',
  generalException: 'GeneralException',
  invalidArgument: 'InvalidArgument',
  itemNotFound: 'ItemNotFound',
  notAllowed: 'NotAllowed',
  notImplemented: 'NotImplemented',
  searchDialogIsOpen: 'SearchDialogIsOpen',
  searchStringInvalidOrTooLong: 'SearchStringInvalidOrTooLong',
} as const);

/** Word.SearchOptions: the options `search` takes (the content API's `SearchOptions`). */
export interface SearchOptions {
  matchCase?: boolean;
  matchWholeWord?: boolean;
  matchWildcards?: boolean;
}

/** Word.SearchOptions.newObject(): a fresh options bag, as add-in code builds one. */
export const SearchOptions = Object.freeze({
  newObject(): SearchOptions {
    return { matchCase: false, matchWholeWord: false, matchWildcards: false };
  },
});
