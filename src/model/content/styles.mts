// Style names, ids and Office JS's built-in style values (CR-002 section 3.1, corrected 2026-09-16).
//
// Office JS: `style` is the style's display name ('Heading 1', 'Table Grid', localised in Word) and
// `styleBuiltIn` a `Word.Style` value ('Heading1', 'TableGrid'; 'Other' for a style that is not
// built in). The tree holds the style id (`w:pStyle`, `w:tblStyle`), which Word derives from the
// English display name by removing spaces, so the three forms map onto each other without the
// styles part for every built-in style; a custom style's name comes from the styles part when it
// is unmarshalled, and is otherwise derived from the id the same way.
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import type { OpcPackage } from '../../packages/OpcPackage.mjs';
import type { XmlPart } from '../../parts/XmlPart.mjs';

/** Office JS `Word.Style` values other than 'Other', in Office JS's spelling. */
export const BUILT_IN_STYLES = Object.freeze([
  'Normal', 'Heading1', 'Heading2', 'Heading3', 'Heading4', 'Heading5', 'Heading6', 'Heading7', 'Heading8', 'Heading9',
  'Toc1', 'Toc2', 'Toc3', 'Toc4', 'Toc5', 'Toc6', 'Toc7', 'Toc8', 'Toc9',
  'FootnoteText', 'Header', 'Footer', 'Caption', 'FootnoteReference', 'EndnoteReference', 'EndnoteText', 'Title', 'Subtitle',
  'Hyperlink', 'Strong', 'Emphasis', 'NoSpacing', 'ListParagraph', 'Quote', 'IntenseQuote', 'SubtleEmphasis', 'IntenseEmphasis',
  'SubtleReference', 'IntenseReference', 'BookTitle', 'Bibliography', 'TocHeading', 'TableGrid',
  'PlainTable1', 'PlainTable2', 'PlainTable3', 'PlainTable4', 'PlainTable5', 'GridTable1Light', 'ListTable1Light',
] as const);
export type BuiltInStyle = (typeof BUILT_IN_STYLES)[number] | 'Other';

/**
 * Word's ids and display names where the id is not the display name without spaces, or the name
 * Word stores in `w:name` is not the display name (docx4j's KnownStyles.xml has the stored names).
 */
const DISPLAY_NAMES: Readonly<Record<string, string>> = Object.freeze({
  TableNormal: 'Normal Table', NormalWeb: 'Normal (Web)', MacroText: 'Macro Text',
  TableofAuthorities: 'Table of Authorities', TableofFigures: 'Table of Figures', TOAHeading: 'TOA Heading',
  CommentText: 'Comment Text', CommentSubject: 'Comment Subject', CommentReference: 'Comment Reference',
  '1ai': 'Outline List 1', '111111': 'Outline List 2', ArticleSection: 'Outline List 3',
});
/** Word's spelling of an id where it is not Office JS's spelling of the built-in value. */
const WORD_IDS: Readonly<Record<string, string>> = Object.freeze({
  toc1: 'TOC1', toc2: 'TOC2', toc3: 'TOC3', toc4: 'TOC4', toc5: 'TOC5', toc6: 'TOC6', toc7: 'TOC7', toc8: 'TOC8', toc9: 'TOC9',
});
const BUILT_IN_BY_LOWER = new Map(BUILT_IN_STYLES.map((s) => [s.toLowerCase(), s] as const));

/** Office JS's `styleBuiltIn` for a style id: the `Word.Style` value, or 'Other'. */
export function builtInOf(styleId: string): BuiltInStyle {
  return BUILT_IN_BY_LOWER.get(styleId.replace(/\s+/g, '').toLowerCase()) ?? 'Other';
}

/**
 * The style id for a `Word.Style` value ('Heading1'; leniently also the display name 'Heading 1'),
 * in Word's spelling ('TOC1' for 'Toc1'). 'Other' cannot be set, as in Office JS.
 */
export function idOfBuiltIn(value: string): string {
  if (value === 'Other') throw new RangeError("styleBuiltIn cannot be set to 'Other'");
  const compact = value.replace(/\s+/g, '');
  const lower = compact.toLowerCase();
  return WORD_IDS[lower] ?? BUILT_IN_BY_LOWER.get(lower) ?? compact;
}

/** The display name of a style id, from its stored `w:name` when given ('heading 1' to 'Heading 1'). */
export function displayNameOf(styleId: string, storedName?: string): string {
  const known = DISPLAY_NAMES[styleId];
  if (known) return known;
  if (storedName !== undefined) {
    // Word stores its built-in names in lower case ('heading 1', 'toc 1', 'footnote text'); a
    // custom style keeps the case its author gave it. Built-ins are the ids Word derives from them.
    if (storedName.replace(/\s+/g, '').toLowerCase() !== styleId.toLowerCase()) return storedName;
    return storedName.replace(/(^|\s)([a-z])/g, (_, sp: string, c: string) => sp + c.toUpperCase()).replace(/^Toc\b/, 'TOC').replace(/^Toa\b/, 'TOA');
  }
  return styleId.replace(/(?<=[a-z0-9])(?=[A-Z])|(?<=[A-Za-z])(?=[0-9])/g, ' ');
}

const stylesPartOf = (pkg: OpcPackage | undefined): XmlPart<wml.Styles> | undefined =>
  (pkg?.getMainPart() as { styleDefinitionsPart?: XmlPart<wml.Styles> } | undefined)?.styleDefinitionsPart;

/** The styles of the package's styles part when it is unmarshalled; nothing is unmarshalled here. */
function stylesOf(pkg: OpcPackage | undefined): wml.Style[] {
  const part = stylesPartOf(pkg);
  return part?.isUnmarshalled ? part.contents.style ?? [] : [];
}

/** Office JS's `style`: the display name of a style id. */
export function styleNameOf(pkg: OpcPackage | undefined, styleId: string): string {
  const style = stylesOf(pkg).find((s) => s.styleId === styleId);
  return displayNameOf(styleId, style?.name?.val);
}

/**
 * The style id for what Office JS's `style` accepts: a display name ('Heading 1', 'My Style'), the
 * stored name ('heading 1'), or, as an extension, the id itself ('Heading1'). Resolved against the
 * styles part when it is unmarshalled, else by removing spaces.
 */
export function styleIdOf(pkg: OpcPackage | undefined, name: string): string {
  const lower = name.toLowerCase();
  for (const s of stylesOf(pkg)) {
    if (s.styleId?.toLowerCase() === lower || s.name?.val?.toLowerCase() === lower || (s.styleId && displayNameOf(s.styleId, s.name?.val).toLowerCase() === lower)) return s.styleId ?? name;
  }
  const compact = name.replace(/\s+/g, '');
  for (const [id, display] of Object.entries(DISPLAY_NAMES)) if (display.toLowerCase() === lower) return id;
  return WORD_IDS[compact.toLowerCase()] ?? compact;
}
