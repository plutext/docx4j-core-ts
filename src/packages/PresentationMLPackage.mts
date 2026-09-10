import { OpcPackage, type PackageSource } from './OpcPackage.mjs';
import { registerPackageClass } from './registry.mjs';
import type { LoadOptions } from '../opc/Load.mjs';
import { ContentTypes } from '../opc/ContentTypes.mjs';
import { Docx4JException } from '../opc/exceptions.mjs';
import { Namespaces } from '../parts/Namespaces.mjs';
import type { Part } from '../parts/Part.mjs';
import type { MainPresentationPart } from '../parts/pml/index.mjs';

const MAIN_CONTENT_TYPES = [
  ContentTypes.PRESENTATIONML_MAIN,
  ContentTypes.PRESENTATIONML_TEMPLATE,
  ContentTypes.PRESENTATIONML_MACROENABLED,
  ContentTypes.PRESENTATIONML_TEMPLATE_MACROENABLED,
  ContentTypes.PRESENTATIONML_SLIDESHOW,
];

/** A pptx (docx4j PresentationMLPackage). */
export class PresentationMLPackage extends OpcPackage {
  mainPresentationPart: MainPresentationPart | undefined;

  static override async load(source: PackageSource, options?: LoadOptions): Promise<PresentationMLPackage> {
    const pkg = await OpcPackage.load(source, options);
    if (!(pkg instanceof PresentationMLPackage)) {
      throw new Docx4JException(`Not a PresentationML package: main part is ${pkg.getMainPart()?.contentType ?? 'missing'}`);
    }
    return pkg;
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

  protected override get progId(): string {
    return 'PowerPoint.Show';
  }
}

registerPackageClass(MAIN_CONTENT_TYPES, PresentationMLPackage);
