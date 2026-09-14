// Compile-only: the built package as a Node ES module consumer sees it (module/moduleResolution nodenext),
// through the package's own name and exports map. The repository's tsconfigs use moduleResolution bundler,
// which accepts extensionless relative imports in declarations; nodenext rejects them (TS2835), and with
// skipLibCheck they silently degrade to any (@docx4j/generated-objects-ts 0.1.0 shipped that). Run by
// npm test after the build.
import { Part, unwrap } from '@docx4j/core-ts';
import { PartName } from '@docx4j/core-ts/opc';
import { defaultPartRegistry } from '@docx4j/core-ts/parts';
import { WordprocessingMLPackage } from '@docx4j/core-ts/packages';
import * as model from '@docx4j/core-ts/model';

export async function consume(): Promise<string> {
  const pkg: WordprocessingMLPackage = await WordprocessingMLPackage.createPackage();
  const name: PartName = PartName.of('/word/document.xml');
  void [Part, unwrap, defaultPartRegistry, model, pkg];
  return name.extension;
}

// Each entry point's types must be real, not any: each line below is an error only while its import
// resolves, so a degraded import leaves the directive unused and the check fails.
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
