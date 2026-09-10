import { OpcPackage, type PackageSource } from './OpcPackage.mjs';
import { registerPackageClass } from './registry.mjs';
import type { LoadOptions } from '../opc/Load.mjs';
import { ContentTypes } from '../opc/ContentTypes.mjs';
import { Docx4JException } from '../opc/exceptions.mjs';
import { Namespaces } from '../parts/Namespaces.mjs';
import type { Part } from '../parts/Part.mjs';
import type { WorkbookPart } from '../parts/sml/index.mjs';

const MAIN_CONTENT_TYPES = [
  ContentTypes.SPREADSHEETML_WORKBOOK,
  ContentTypes.SPREADSHEETML_WORKBOOK_MACROENABLED,
  ContentTypes.SPREADSHEETML_TEMPLATE,
  ContentTypes.SPREADSHEETML_TEMPLATE_MACROENABLED,
];

/** An xlsx (docx4j SpreadsheetMLPackage). */
export class SpreadsheetMLPackage extends OpcPackage {
  workbookPart: WorkbookPart | undefined;

  static override async load(source: PackageSource, options?: LoadOptions): Promise<SpreadsheetMLPackage> {
    const pkg = await OpcPackage.load(source, options);
    if (!(pkg instanceof SpreadsheetMLPackage)) {
      throw new Docx4JException(`Not a SpreadsheetML package: main part is ${pkg.getMainPart()?.contentType ?? 'missing'}`);
    }
    return pkg;
  }

  override setPartShortcut(part: Part, relationshipType: string): boolean {
    if (relationshipType === Namespaces.DOCUMENT || relationshipType === Namespaces.DOCUMENT_STRICT) {
      this.workbookPart = part as WorkbookPart;
      return true;
    }
    return super.setPartShortcut(part, relationshipType);
  }

  getWorkbookPart(): WorkbookPart {
    if (!this.workbookPart) throw new Docx4JException('No workbook part');
    return this.workbookPart;
  }

  protected override get progId(): string {
    return 'Excel.Sheet';
  }
}

registerPackageClass(MAIN_CONTENT_TYPES, SpreadsheetMLPackage);
