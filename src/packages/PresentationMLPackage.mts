import type * as pml from '@docx4j/generated-objects-ts/modules/org_pptx4j_pml';
import { createPresentation, createPresentationSldSz } from '@docx4j/generated-objects-ts/factory/org_pptx4j_pml';
import { OpcPackage, type PackageSource } from './OpcPackage.mjs';
import { registerPackageClass } from './registry.mjs';
import type { LoadOptions } from '../opc/Load.mjs';
import { ContentTypes } from '../opc/ContentTypes.mjs';
import { Docx4JException } from '../opc/exceptions.mjs';
import { Namespaces } from '../parts/Namespaces.mjs';
import type { Part } from '../parts/Part.mjs';
import {
  MainPresentationPart, SlidePart, SlideLayoutPart, SlideMasterPart, NotesMasterPart,
  PresentationPropertiesPart, ViewPropertiesPart, TableStylesPart, CommentAuthorsPart,
} from '../parts/pml/index.mjs';
import { DEFAULT_SLIDE_MASTER_XML, DEFAULT_SLIDE_LAYOUT_XML, DEFAULT_SLIDE_XML } from '../parts/pml/defaults.mjs';
import { ThemePart } from '../parts/dml/index.mjs';
import { newFontSettings, themeOfSettings, type FontSettings, type DefaultThemeValue } from '../model/fonts/defaultTheme.mjs';

const MAIN_CONTENT_TYPES = [
  ContentTypes.PRESENTATIONML_MAIN,
  ContentTypes.PRESENTATIONML_TEMPLATE,
  ContentTypes.PRESENTATIONML_MACROENABLED,
  ContentTypes.PRESENTATIONML_TEMPLATE_MACROENABLED,
  ContentTypes.PRESENTATIONML_SLIDESHOW,
];

/** docx4j / pptx4j `org.pptx4j.model.SlideSizesWellKnown`. */
export type SlideSizesWellKnown =
  | 'LETTER' | 'A3' | 'A4' | 'B4JIS' | 'SCREEN4x3' | 'SCREEN16x9' | 'SCREEN16x10'
  | 'LEDGER' | 'B4ISO' | 'B5ISO' | 'MM35' | 'OVERHEAD' | 'BANNER';

/**
 * The landscape `cx` and `cy` in EMU, and the `p:sldSz/@type`, as docx4j's
 * `MainPresentationPart.createSlideSize` has them. Portrait swaps the two.
 * `B4JIS` is in the enumeration but has no size in docx4j, which throws for it; so does this.
 */
const SLIDE_SIZES: Record<Exclude<SlideSizesWellKnown, 'B4JIS'>, [cx: number, cy: number, type: string]> = {
  LETTER: [9144000, 6858000, 'letter'],
  A3: [12801600, 9601200, 'A3'],
  A4: [9906000, 6858000, 'A4'],
  SCREEN4x3: [9144000, 6858000, 'screen4x3'],
  SCREEN16x9: [9144000, 5143500, 'screen16x9'],
  SCREEN16x10: [9144000, 5715000, 'screen16x10'],
  LEDGER: [12179300, 9134475, 'ledger'],
  B4ISO: [10826750, 8120063, 'B4ISO'],
  B5ISO: [7169150, 5376863, 'B5ISO'],
  MM35: [10287000, 6858000, '35mm'],
  OVERHEAD: [9144000, 6858000, 'overhead'],
  BANNER: [7315200, 914400, 'banner'],
};

/** docx4j `MainPresentationPart.createSlideSize(sz, landscape)`. */
export function createSlideSize(size: SlideSizesWellKnown, landscape = true): pml.Presentation.SldSz {
  const dims = SLIDE_SIZES[size as Exclude<SlideSizesWellKnown, 'B4JIS'>];
  if (!dims) throw new Docx4JException(`No support for slide size ${size}`);
  const [cx, cy, type] = dims;
  return createPresentationSldSz(landscape ? { cx, cy, type } : { cx: cy, cy: cx, type });
}

export interface CreatePresentationOptions {
  /** docx4j's `pptx4j.PageSize`; `'A4'` by default, as docx4j's property does. */
  slideSize?: SlideSizesWellKnown;
  /** docx4j's `pptx4j.PageOrientationLandscape`; true by default. */
  landscape?: boolean;
  /** Which Office theme the new presentation's theme part carries; see `pkg.fonts.defaultTheme`. */
  defaultTheme?: DefaultThemeValue;
}

/** A pptx (docx4j PresentationMLPackage). */
export class PresentationMLPackage extends OpcPackage {
  mainPresentationPart: MainPresentationPart | undefined;

  /** The package's font settings, as `WordprocessingMLPackage.fonts`: which Office theme a new
   *  presentation's theme part carries. */
  readonly fonts: FontSettings = newFontSettings();

  static override async load(source: PackageSource, options?: LoadOptions): Promise<PresentationMLPackage> {
    const pkg = await OpcPackage.load(source, options);
    if (!(pkg instanceof PresentationMLPackage)) {
      throw new Docx4JException(`Not a PresentationML package: main part is ${pkg.getMainPart()?.contentType ?? 'missing'}`);
    }
    return pkg;
  }

  /**
   * A new presentation, as docx4j's `PresentationMLPackage.createPackage(sz, landscape)` builds
   * one: the presentation part with the slide size, a slide master, a slide layout, a theme part
   * in both the master's and the presentation's relationships, and — unlike docx4j, whose slide
   * creation is commented out — one empty slide, so the file opens as a presentation rather than
   * as an empty deck (CR-001 section 16).
   *
   * The master, layout and slide are docx4j's own template markup
   * (`src/parts/pml/defaults.mts`); the theme is the Office theme `pkg.fonts.defaultTheme` names,
   * the same one `WordprocessingMLPackage.createPackage` uses.
   */
  static async createPackage(options: CreatePresentationOptions = {}): Promise<PresentationMLPackage> {
    const pkg = new PresentationMLPackage();
    if (options.defaultTheme !== undefined) pkg.fonts.defaultTheme = options.defaultTheme;

    // The presentation part (docx4j MainPresentationPart.createJaxbPresentationElement)
    const pp = new MainPresentationPart();
    pp.setContents(createPresentation({
      sldMasterIdLst: { sldMasterId: [] },
      sldIdLst: { sldId: [] },
      sldSz: createSlideSize(options.slideSize ?? 'A4', options.landscape ?? true),
      notesSz: { cx: 6858000, cy: 9144000 },
    }));
    pkg.addTargetPart(pp);

    // The layout, and the master that lists it
    const layoutPart = new SlideLayoutPart();
    layoutPart.setXml(DEFAULT_SLIDE_LAYOUT_XML);

    const masterPart = new SlideMasterPart();
    pp.addSlideMasterIdListEntry(masterPart);
    masterPart.setXml(DEFAULT_SLIDE_MASTER_XML);
    await masterPart.getContents();
    masterPart.addSlideLayoutIdListEntry(layoutPart);
    layoutPart.addTargetPart(masterPart);

    // The theme, in both places docx4j puts it (PowerPoint's own files relate it from the
    // presentation part as well as from the master)
    const theme = new ThemePart('/ppt/theme/theme1.xml');
    theme.setXml(themeOfSettings(pkg.fonts).xml);
    masterPart.addTargetPart(theme);
    pp.addTargetPart(theme);

    // One slide on the layout
    const slidePart = new SlidePart();
    slidePart.setXml(DEFAULT_SLIDE_XML);
    pp.addSlideIdListEntry(slidePart);
    slidePart.addTargetPart(layoutPart);

    return pkg;
  }

  /**
   * Adds a slide part to the presentation (docx4j `MainPresentationPart.addSlide`), on the
   * layout given or on the presentation's first layout, with docx4j's empty shape tree when the
   * part has no contents of its own. Appends, or inserts at `index`.
   */
  async addSlide(options: { index?: number; layoutPart?: SlideLayoutPart; partName?: string } = {}): Promise<SlidePart> {
    const pp = this.getMainPresentationPart();
    await pp.getContents();
    const layoutPart = options.layoutPart ?? this.slideMasterParts[0]?.slideLayoutParts[0];
    if (!layoutPart) throw new Docx4JException('The presentation has no slide layout to put a slide on');
    const slidePart = new SlidePart(options.partName ?? this.nextSlideName());
    slidePart.setXml(DEFAULT_SLIDE_XML);
    pp.addSlide(slidePart, options.index, layoutPart);
    return slidePart;
  }

  /** `/ppt/slides/slideN.xml` for the lowest N the package has no part for. */
  private nextSlideName(): string {
    for (let n = 1; ; n++) {
      const name = `/ppt/slides/slide${n}.xml`;
      if (!this.getPart(name)) return name;
    }
  }

  override setPartShortcut(part: Part, relationshipType: string): boolean {
    if (relationshipType === Namespaces.DOCUMENT || relationshipType === Namespaces.DOCUMENT_STRICT) {
      this.mainPresentationPart = part as MainPresentationPart;
      return true;
    }
    return super.setPartShortcut(part, relationshipType);
  }

  getMainPresentationPart(): MainPresentationPart {
    if (!this.mainPresentationPart) throw new Docx4JException('No main presentation part');
    return this.mainPresentationPart;
  }

  /** The slides, in `p:sldIdLst` order (docx4j `MainPresentationPart.getSlideParts()`). */
  get slideParts(): SlidePart[] {
    return this.mainPresentationPart?.slideParts ?? [];
  }

  /** The slides in `p:sldIdLst` order, unmarshalling the presentation part first. */
  async getSlideParts(): Promise<SlidePart[]> {
    return this.getMainPresentationPart().getSlideParts();
  }

  /** The slide masters, in `p:sldMasterIdLst` order. */
  get slideMasterParts(): SlideMasterPart[] {
    return this.mainPresentationPart?.slideMasterParts ?? [];
  }

  /** Every slide layout of every master, masters in order and layouts in `p:sldLayoutIdLst` order. */
  get slideLayoutParts(): SlideLayoutPart[] {
    return this.slideMasterParts.flatMap((m) => m.slideLayoutParts);
  }

  get notesMasterPart(): NotesMasterPart | undefined {
    return this.mainPresentationPart?.notesMasterPart;
  }

  get presentationPropertiesPart(): PresentationPropertiesPart | undefined {
    return this.mainPresentationPart?.presentationPropertiesPart;
  }

  get viewPropertiesPart(): ViewPropertiesPart | undefined {
    return this.mainPresentationPart?.viewPropertiesPart;
  }

  get tableStylesPart(): TableStylesPart | undefined {
    return this.mainPresentationPart?.tableStylesPart;
  }

  get commentAuthorsPart(): CommentAuthorsPart | undefined {
    return this.mainPresentationPart?.commentAuthorsPart;
  }

  /** The theme related from the presentation part (docx4j `MainPresentationPart.getThemePart()`). */
  get themePart(): ThemePart | undefined {
    return this.mainPresentationPart?.themePart;
  }

  /** A clone carries the theme setting. */
  protected override copyPackageSettingsTo(target: OpcPackage): void {
    if (target instanceof PresentationMLPackage) target.fonts.defaultTheme = this.fonts.defaultTheme;
  }

  protected override get progId(): string {
    return 'PowerPoint.Show';
  }
}

registerPackageClass(MAIN_CONTENT_TYPES, PresentationMLPackage);
