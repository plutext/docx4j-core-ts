// Compile-only: the built package as a Node ES module consumer sees it (module/moduleResolution nodenext),
// through the package's own name and exports map. The repository's tsconfigs use moduleResolution bundler,
// which accepts extensionless relative imports in declarations; nodenext rejects them (TS2835), and with
// skipLibCheck they silently degrade to any (@docx4j/generated-objects-ts 0.1.0 shipped that). Run by
// npm test after the build.
import { Part, unwrap } from '@docx4j/core-ts';
import { PartName } from '@docx4j/core-ts/opc';
import { defaultPartRegistry } from '@docx4j/core-ts/parts';
import { WordprocessingMLPackage } from '@docx4j/core-ts/packages';
import { DirectoryPartStore, DirectoryPartSink } from '@docx4j/core-ts/node';
import * as model from '@docx4j/core-ts/model';
import { Word, NotSupportedError, toApiScript, type ShimmedDocument } from '@docx4j/core-ts/office-js';
import { FontoXPathEngine } from '@docx4j/core-ts/xpath-fonto';

export async function consume(): Promise<string> {
  const pkg: WordprocessingMLPackage = await WordprocessingMLPackage.createPackage();
  const name: PartName = PartName.of('/word/document.xml');
  void [Part, unwrap, defaultPartRegistry, model, pkg, DirectoryPartSink];
  pkg.xpathEngine = new FontoXPathEngine();
  // the browser form: the application imports fontoxpath itself and hands it over
  pkg.xpathEngine = new FontoXPathEngine(await import('fontoxpath'));
  pkg.xpathEngine = new FontoXPathEngine();
  // the browser form: the application imports fontoxpath itself and hands it over
  pkg.xpathEngine = new FontoXPathEngine(await import('fontoxpath'));
  const script: string = await toApiScript(pkg.body);
  await Word.run(pkg, (context) => context.document.body.insertParagraph(script, 'End'));
  void NotSupportedError;
  return name.extension;
}


/**
 * CR-002 section 22.1: the shim accepts Office JS's collection members at runtime, so its types must
 * accept them too - an add-in script that type-checks against Office JS must type-check against this
 * package. Before `ShimmedDocument`, `context.document.body.paragraphs` was `Paragraph[]` and every
 * line below was an error while the shim ran them happily.
 */
export async function addInScript(pkg: WordprocessingMLPackage): Promise<number> {
  return Word.run(pkg, async (context) => {
    const document: ShimmedDocument = context.document;
    const body = document.body;
    body.load('text');
    const paragraphs = body.paragraphs;
    paragraphs.load('items/text');
    paragraphs.getFirst().load('text');
    paragraphs.getFirstOrNullObject().track();
    for (const paragraph of paragraphs.items) {
      const font = paragraph.getRange().font;
      font.load('bold');
      font.bold = true;
    }
    const count = paragraphs.getCount();
    await context.sync();
    const text: string = body.text;
    return count.value + text.length;
  });
}

// Each entry point's types must be real, not any: each line below is an error only while its import
// resolves, so a degraded import leaves the directive unused and the check fails.
// @ts-expect-error FontoXPathEngine is a class, not a number
export const fonto: number = FontoXPathEngine;
// @ts-expect-error Part is a class, not a number
export const root: number = Part;
// @ts-expect-error PartName is a class, not a number
export const opc: number = PartName;
// @ts-expect-error defaultPartRegistry is a PartRegistry, not a number
export const parts: number = defaultPartRegistry;
// @ts-expect-error WordprocessingMLPackage is a class, not a number
export const packages: number = WordprocessingMLPackage;
// @ts-expect-error the model entry point is a module namespace, not a number
export const modelNs: number = model;
// @ts-expect-error Word is the shim object, not a number
export const officeJs: number = Word;
// @ts-expect-error DirectoryPartStore is a class, not a number
export const nodeOnly: number = DirectoryPartStore;
