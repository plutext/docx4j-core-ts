import { createWorkbook, createSheet, createWorksheet } from '@docx4j/generated-objects-ts/factory/org_xlsx4j_sml';
import { OpcPackage, type PackageSource } from './OpcPackage.mjs';
import { registerPackageClass } from './registry.mjs';
import type { LoadOptions } from '../opc/Load.mjs';
import { ContentTypes } from '../opc/ContentTypes.mjs';
import { Docx4JException } from '../opc/exceptions.mjs';
import { Namespaces } from '../parts/Namespaces.mjs';
import type { Part } from '../parts/Part.mjs';
import {
  WorkbookPart, WorksheetPart, SharedStringsPart, StylesPart, CalcChainPart,
} from '../parts/sml/index.mjs';
import type { ThemePart } from '../parts/dml/index.mjs';

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

  /**
   * A new workbook with no sheets (docx4j `SpreadsheetMLPackage.createPackage()`): the workbook
   * part, an empty `sheets`, and the one `bookViews/workbookView` docx4j adds because without it
   * Excel 2010 could crash on print. Add sheets with `createWorksheetPart`.
   */
  static async createPackage(): Promise<SpreadsheetMLPackage> {
    const pkg = new SpreadsheetMLPackage();
    const wb = new WorkbookPart();
    wb.setContents(createWorkbook({
      bookViews: { workbookView: [{}] },
      sheets: { sheet: [] },
    }));
    pkg.addTargetPart(wb);
    return pkg;
  }

  /**
   * Creates a worksheet part, relates it from the workbook and adds the `sheet` entry that names
   * it (docx4j `createWorksheetPart(partName, sheetName, sheetId)`). The part holds an empty
   * `sheetData`, which is the least Excel will open.
   *
   * `index` inserts the tab at that position in `sheets` instead of appending. The part name
   * defaults to the first free `/xl/worksheets/sheetN.xml` and the `sheetId` to the next free one.
   */
  createWorksheetPart(name: string, index?: number, options: { partName?: string; sheetId?: number } = {}): WorksheetPart {
    const wb = this.getWorkbookPart();
    const workbook = wb.contents;
    const sheets = (workbook.sheets ??= { sheet: [] });
    const list = (sheets.sheet ??= []);
    if (index !== undefined && (index < 0 || index > list.length)) {
      throw new Docx4JException(`Can't add a sheet at index ${index}. (There are ${list.length} sheets) `);
    }
    const part = new WorksheetPart(options.partName ?? this.nextWorksheetName());
    const rel = wb.addTargetPart(part);
    const sheetId = options.sheetId ?? list.reduce((max, s) => Math.max(max, s.sheetId), 0) + 1;
    const entry = createSheet({ name, sheetId, id: rel.id });
    if (index === undefined) list.push(entry); else list.splice(index, 0, entry);
    part.setContents(createWorksheet({ sheetData: {} }));
    return part;
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

  /** The worksheets, in `sheets` order (the tab order Excel shows). */
  get worksheetParts(): WorksheetPart[] {
    return this.workbookPart?.worksheetParts ?? [];
  }

  /** The worksheets in `sheets` order, unmarshalling the workbook part first. */
  async getWorksheetParts(): Promise<WorksheetPart[]> {
    return this.getWorkbookPart().getWorksheetParts();
  }

  get sharedStringsPart(): SharedStringsPart | undefined {
    return this.workbookPart?.sharedStringsPart;
  }

  get stylesPart(): StylesPart | undefined {
    return this.workbookPart?.stylesPart;
  }

  get calcChainPart(): CalcChainPart | undefined {
    return this.workbookPart?.calcChainPart;
  }

  get themePart(): ThemePart | undefined {
    return this.workbookPart?.themePart;
  }

  /** `/xl/worksheets/sheetN.xml` for the lowest N the package has no part for. */
  private nextWorksheetName(): string {
    for (let n = 1; ; n++) {
      const name = `/xl/worksheets/sheet${n}.xml`;
      if (!this.getPart(name)) return name;
    }
  }

  protected override get progId(): string {
    return 'Excel.Sheet';
  }
}

registerPackageClass(MAIN_CONTENT_TYPES, SpreadsheetMLPackage);
