// docx4j `MainDocumentPart.fontsInUse()` and `getStylesInUse()`, as CR-016 phase 4 made them:
// **a walk for names**, not a selection.
//
// Until docx4j 17.1.1 discovery ran a `RunFontSelector` in a DISCOVERY mode over every run,
// deciding among the run's fonts by glyph checks it could not yet answer - the mapper is
// populated from this list, so nothing was mapped - and took `w:ascii` alone from the numbering
// levels.  What discovery needs is every name the document could ask for, so that the mapper
// can map it: the four slots of every `w:rFonts` on the runs, the paragraph marks and the
// content controls, over the body, headers, footers, footnotes, endnotes and comments; the
// styles in use with what they are based on (a table style's `w:tblStylePr` run properties
// included); the numbering levels; the document defaults; and the default font.  Each theme
// reference is resolved through the theme part for the document's `themeFontLang`, and a CJK
// font name is collected by its English name.
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import { walk } from '@docx4j/generated-objects-ts/builders/wml';
import type { RunFontSelector } from './RunFontSelector.mjs';
import { toEnglish } from './CJKToEnglish.mjs';

/** The stories a walk covers, in docx4j's order: the body, then headers and footers, endnotes,
 *  footnotes and comments.  Each value is whatever holds block-level content. */
export type Stories = Iterable<unknown>;

/** The styles part, as much of it as the walk needs. */
export interface StyleSource {
  styles?: wml.Styles | undefined;
  /** The `w:default="1"` paragraph, character and table style ids. */
  defaultParagraphStyleId?: string | undefined;
  defaultCharacterStyleId?: string | undefined;
  defaultTableStyleId?: string | undefined;
}

/** What one walk found. */
export interface FontsAndStyles {
  /** The document font names, theme references resolved, CJK names in English. */
  fonts: Set<string>;
  /** The style ids used directly (not those merely based on). */
  styles: Set<string>;
}

interface Collector {
  rFonts(rFonts: wml.RFonts | undefined): void;
  name(fontName: string | undefined): void;
}

function collector(selector: RunFontSelector | undefined, into: Set<string>): Collector {
  const add = (fontName: string | undefined): void => {
    if (fontName === undefined || fontName === null) return;
    const name = fontName.trim();
    if (name.length === 0) return;
    // where there is an English name, no point adding the original CJK name
    into.add(toEnglish(name) ?? name);
  };
  return {
    name: add,
    rFonts(rFonts) {
      if (rFonts === undefined || selector === undefined) return;
      for (const name of selector.documentFontsOf(rFonts)) add(name);
    },
  };
}

/** The `w:rPr` a content control's `w:sdtPr` carries, if any. */
function sdtRPr(sdtPr: wml.SdtPr): wml.RPr | undefined {
  for (const entry of sdtPr.rPrOrAliasOrLock ?? []) {
    const value = (entry as { value?: { TYPE_NAME?: string } }).value;
    if (value?.TYPE_NAME === 'org_docx4j_wml.RPr') return value as wml.RPr;
  }
  return undefined;
}

/**
 * One traversal of every story, collecting font names and the style ids used directly
 * (docx4j's `FontAndStyleFinder`).  The default paragraph, character and table styles count as
 * used where a paragraph, run or table names none of its own.
 *
 * @param selector resolves theme references; omit to collect styles only
 */
export function walkStories(stories: Stories, styleSource: StyleSource | undefined,
  selector: RunFontSelector | undefined, collectFonts: boolean): FontsAndStyles {

  const fonts = new Set<string>();
  const styles = new Set<string>();
  const add = collector(selector, fonts);
  let defaultParagraphStyleUsed = false;
  let defaultCharacterStyleUsed = false;
  let defaultTableStyleUsed = false;

  for (const story of stories) {
    walk(story, (value) => {
      switch ((value as { TYPE_NAME?: string }).TYPE_NAME) {
        case 'org_docx4j_wml.P': {
          const pPr = (value as wml.P).pPr;
          if (collectFonts && pPr?.rPr !== undefined) add.rFonts(pPr.rPr.rFonts); // the paragraph mark's
          let customPStyle = false;
          if (pPr !== undefined) {
            if (pPr.pStyle?.val !== undefined) { customPStyle = true; styles.add(pPr.pStyle.val); }
            if (pPr.rPr?.rStyle?.val !== undefined) styles.add(pPr.rPr.rStyle.val);
          }
          defaultParagraphStyleUsed ||= !customPStyle;
          break;
        }
        case 'org_docx4j_wml.R': {
          const rPr = (value as wml.R).rPr;
          if (collectFonts && rPr !== undefined) add.rFonts(rPr.rFonts);
          if (rPr !== undefined) {
            if (rPr.rStyle?.val === undefined) defaultCharacterStyleUsed = true;
            else styles.add(rPr.rStyle.val);
          }
          break;
        }
        case 'org_docx4j_wml.R.Sym':
          if (collectFonts) add.name((value as wml.R.Sym).font);
          break;
        case 'org_docx4j_wml.Tbl': {
          const tblStyle = (value as wml.Tbl).tblPr?.tblStyle?.val;
          if (tblStyle !== undefined) styles.add(tblStyle);
          defaultTableStyleUsed ||= tblStyle === undefined;
          break;
        }
        case 'org_docx4j_wml.SdtPr': {
          const rPr = sdtRPr(value as wml.SdtPr);
          if (rPr === undefined) break;
          if (collectFonts) add.rFonts(rPr.rFonts);
          const rStyle = rPr.rStyle?.val;
          if (rStyle !== undefined) {
            styles.add(rStyle);
            // and the linked paragraph style, if any
            const linked = linkedStyleOf(styleSource?.styles, rStyle);
            if (linked !== undefined) styles.add(linked);
          }
          break;
        }
        default:
          break;
      }
      return undefined;
    });
  }

  if (defaultParagraphStyleUsed && styleSource?.defaultParagraphStyleId !== undefined) {
    styles.add(styleSource.defaultParagraphStyleId);
    const style = styleById(styleSource.styles, styleSource.defaultParagraphStyleId);
    if (style?.rPr?.rStyle?.val !== undefined) styles.add(style.rPr.rStyle.val);
  }
  if (defaultCharacterStyleUsed && styleSource?.defaultCharacterStyleId !== undefined) {
    styles.add(styleSource.defaultCharacterStyleId);
  }
  if (defaultTableStyleUsed && styleSource?.defaultTableStyleId !== undefined) {
    styles.add(styleSource.defaultTableStyleId);
  }
  return { fonts, styles };
}

function styleById(styles: wml.Styles | undefined, styleId: string): wml.Style | undefined {
  return styles?.style?.find((s) => s.styleId === styleId);
}

/** The paragraph style a character style is linked to (`w:link`), by style id. */
function linkedStyleOf(styles: wml.Styles | undefined, styleId: string): string | undefined {
  const link = styleById(styles, styleId)?.link?.val;
  if (link === undefined) return undefined;
  return styleById(styles, link)?.styleId;
}

/**
 * The font names a document uses: the walk above, then the styles in use with their `w:basedOn`
 * chains, the numbering levels, the document defaults, and the default font.
 *
 * @param stories the body and the other story parts' content
 * @param styleSource the styles part
 * @param numbering the numbering part's content, or undefined
 * @param selector the package's selector (it resolves the theme references and answers the
 *   default font)
 */
export function fontsInUse(stories: Stories, styleSource: StyleSource | undefined,
  numbering: wml.Numbering | undefined, selector: RunFontSelector): Set<string> {

  const { fonts, styles } = walkStories(stories, styleSource, selector, true);
  const add = collector(selector, fonts);
  const allStyles = styleSource?.styles;

  const docDefaultsRFonts = allStyles?.docDefaults?.rPrDefault?.rPr?.rFonts;
  if (allStyles !== undefined) {
    add.rFonts(docDefaultsRFonts);
    // the styles in use, each with the chain it is based on
    for (const styleId of styles) {
      let style = styleById(allStyles, styleId);
      for (let guard = 0; style !== undefined && guard < 64; guard++) {
        if (style.rPr !== undefined) add.rFonts(style.rPr.rFonts);
        for (const tblStylePr of style.tblStylePr ?? []) {
          if (tblStylePr.rPr !== undefined) add.rFonts(tblStylePr.rPr.rFonts);
        }
        const basedOn = style.basedOn?.val;
        style = basedOn === undefined ? undefined : styleById(allStyles, basedOn);
      }
    }
  }

  // Fonts can also be used in the numbering part.  For now, treat any font mentioned in that
  // part as in use (docx4j's comment: ideally only the levels actually used).
  for (const abstractNum of numbering?.abstractNum ?? []) {
    for (const lvl of abstractNum.lvl ?? []) {
      if (lvl.rPr !== undefined) add.rFonts(lvl.rPr.rFonts);
    }
  }

  fonts.add(selector.defaultFontOf(docDefaultsRFonts));
  return fonts;
}
