import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import { OpcPackage, type PackageSource } from './OpcPackage.mjs';
import { registerPackageClass } from './registry.mjs';
import type { LoadOptions } from '../opc/Load.mjs';
import { ContentTypes } from '../opc/ContentTypes.mjs';
import { Docx4JException } from '../opc/exceptions.mjs';
import { Namespaces } from '../parts/Namespaces.mjs';
import type { Part } from '../parts/Part.mjs';
import { MainDocumentPart, StyleDefinitionsPart, DocumentSettingsPart } from '../parts/wml/index.mjs';
import { DocPropsCorePart, DocPropsExtendedPart } from '../parts/docProps/index.mjs';
import { DEFAULT_STYLES_XML } from '../parts/wml/defaultStyles.mjs';
import { HeaderPart, FooterPart } from '../parts/wml/index.mjs';
import type { Body, Address, Outline, OutlineParagraph, OutlineTable } from '../model/content/Body.mjs';
import type { Paragraph } from '../model/content/Paragraph.mjs';
import type { Author } from '../model/content/comments.mjs';
// Registers the comment parts with the content API (CR-002 phase G); no cycle: the model half
// never imports a part, as packages/registry.mts does for the package classes.
import '../parts/wml/comments.mjs';

/** The outline of a document: the body's paragraphs and tables, then each header's and footer's (loaded ones only). */
export interface DocumentOutline extends Outline {
  headers: { relId: string; partName: string; paragraphs: OutlineParagraph[]; tables: OutlineTable[] }[];
  footers: { relId: string; partName: string; paragraphs: OutlineParagraph[]; tables: OutlineTable[] }[];
}

const MAIN_CONTENT_TYPES = [
  ContentTypes.WORDPROCESSINGML_DOCUMENT,
  ContentTypes.WORDPROCESSINGML_DOCUMENT_MACROENABLED,
  ContentTypes.WORDPROCESSINGML_TEMPLATE,
  ContentTypes.WORDPROCESSINGML_TEMPLATE_MACROENABLED,
];

/** docx4j PageSizePaper. */
export type PageSizePaper = 'LETTER' | 'LEGAL' | 'A3' | 'A4' | 'A5' | 'B4JIS';

/** Portrait width and height in twips, and the `w:code`, as docx4j PageDimensions. */
const PAGE_SIZES: Record<PageSizePaper, [w: number, h: number, code: number]> = {
  LETTER: [12240, 15840, 1],
  LEGAL: [12240, 20160, 5],
  A3: [16839, 23814, 8],
  A4: [11907, 16839, 9],
  A5: [8391, 11907, 11],
  B4JIS: [14572, 20639, 12],
};

export interface CreatePackageOptions {
  pageSize?: PageSizePaper;
  landscape?: boolean;
}

/** A docx (docx4j WordprocessingMLPackage). */
export class WordprocessingMLPackage extends OpcPackage {
  mainDocumentPart: MainDocumentPart | undefined;

  /**
   * Who this package's comments (CR-002 phase G) and, with phase F, its tracked changes are by.
   * There is no signed-in user here, so the package carries the identity; the initials default to
   * the first letter of each word of the name and the email, when given, is written to `w:people`.
   */
  author: Author = { name: 'docx4j' };

  static override async load(source: PackageSource, options?: LoadOptions): Promise<WordprocessingMLPackage> {
    const pkg = await OpcPackage.load(source, options);
    if (!(pkg instanceof WordprocessingMLPackage)) {
      throw new Docx4JException(`Not a WordprocessingML package: main part is ${pkg.getMainPart()?.contentType ?? 'missing'}`);
    }
    return pkg;
  }

  /**
   * A new document with a main document part (one section, the page size and 2.54 cm margins),
   * docx4j's default styles, settings, and core and extended properties (docx4j createPackage).
   */
  static async createPackage(options: CreatePackageOptions = {}): Promise<WordprocessingMLPackage> {
    const pkg = new WordprocessingMLPackage();
    const [pw, ph, code] = PAGE_SIZES[options.pageSize ?? 'A4'];
    const landscape = options.landscape === true;
    const pgSz: wml.SectPr.PgSz = { w: landscape ? ph : pw, h: landscape ? pw : ph, code };
    if (landscape) pgSz.orient = 'landscape';
    const pgMar: wml.SectPr.PgMar = { top: 1440, right: 1440, bottom: 1440, left: 1440, header: 708, footer: 708, gutter: 0 };
    const document: wml.Document = { body: { content: [], sectPr: { pgSz, pgMar } } };
    const main = new MainDocumentPart();
    main.setContents(document);
    pkg.addTargetPart(main);

    const styles = new StyleDefinitionsPart();
    styles.setXml(DEFAULT_STYLES_XML);
    main.addTargetPart(styles);

    const settings = new DocumentSettingsPart();
    settings.setContents({});
    main.addTargetPart(settings);

    const core = new DocPropsCorePart();
    core.setContents({});
    pkg.addTargetPart(core);
    const app = new DocPropsExtendedPart();
    app.setContents({});
    pkg.addTargetPart(app);
    return pkg;
  }

  override setPartShortcut(part: Part, relationshipType: string): boolean {
    if (relationshipType === Namespaces.DOCUMENT || relationshipType === Namespaces.DOCUMENT_STRICT) {
      this.mainDocumentPart = part as MainDocumentPart;
      return true;
    }
    return super.setPartShortcut(part, relationshipType);
  }

  /** The main document part; throws if the package has none. */
  getMainDocumentPart(): MainDocumentPart {
    if (!this.mainDocumentPart) throw new Docx4JException('No main document part');
    return this.mainDocumentPart;
  }

  /** Office JS `document.body`: the main document part's body; the part must be unmarshalled (`getBody()` does that). */
  get body(): Body {
    return this.getMainDocumentPart().body;
  }

  async getBody(): Promise<Body> {
    return this.getMainDocumentPart().getBody();
  }

  /** The paragraph at an address anywhere in the document: 'body/3', 'header:rId5/0', a paraId, or a text match. */
  async paragraphAt(address: Address): Promise<Paragraph | undefined> {
    for (const body of await this.bodies(address)) {
      const p = body.paragraphAt(address);
      if (p) return p;
    }
    return undefined;
  }

  /** Every paragraph and table with its address, style and text (CR-002 section 3.3). Unmarshals the main part, headers and footers. */
  async outline(): Promise<DocumentOutline> {
    const main = await this.getBody();
    const out: DocumentOutline = { ...main.outline(), headers: [], footers: [] };
    for (const part of this.getMainDocumentPart().headerParts) {
      const body = await part.getBody();
      out.headers.push({ relId: part.sourceRelationship?.id ?? '', partName: part.partName.name, ...body.outline() });
    }
    for (const part of this.getMainDocumentPart().footerParts) {
      const body = await part.getBody();
      out.footers.push({ relId: part.sourceRelationship?.id ?? '', partName: part.partName.name, ...body.outline() });
    }
    return out;
  }

  private async bodies(address: Address): Promise<Body[]> {
    const main = this.getMainDocumentPart();
    if (typeof address === 'string' && !address.startsWith('w14:')) {
      const prefix = address.split('/')[0]!;
      if (prefix === 'body') return [await main.getBody()];
      const [kind, relId] = prefix.split(':');
      const parts = kind === 'header' ? main.headerParts : kind === 'footer' ? main.footerParts : [];
      const part = parts.find((p) => p.sourceRelationship?.id === relId || p.partName.fileName === relId);
      return part instanceof HeaderPart || part instanceof FooterPart ? [await part.getBody()] : [];
    }
    const all: Body[] = [await main.getBody()];
    for (const part of [...main.headerParts, ...main.footerParts]) all.push(await part.getBody());
    return all;
  }

  protected override get progId(): string {
    return 'Word.Document';
  }
}

registerPackageClass(MAIN_CONTENT_TYPES, WordprocessingMLPackage);
