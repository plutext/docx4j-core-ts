// A presentation from nothing: create, add a second slide, save.
//
//   node examples/node/pptx.mjs [out.pptx]
//
// PresentationML has no content API (CR-002 is WordprocessingML); what you get is the typed
// parts, so a slide's shape tree is `slide.contents.cSld.spTree` from the object model. A created
// slide carries the title and body placeholders of its layout, which `{ title, body }` fill in.
//
// Published, the import is `from '@docx4j/core-ts'`; here it is the built dist of this checkout.
import { writeFile } from 'node:fs/promises';
import { PresentationMLPackage } from '../../dist/index.mjs';

const out = process.argv[2] ?? 'hello.pptx';

const pkg = await PresentationMLPackage.createPackage({
  slideSize: 'SCREEN16x9',
  title: 'Hello from docx4j',
  body: ['Created with @docx4j/core-ts', 'No PowerPoint needed'],
});
await pkg.addSlide({ title: 'A second slide', body: 'One body line\nAnd another' });

const presentation = pkg.getMainPresentationPart().contents;
console.log(`slide size: ${presentation.sldSz.cx} x ${presentation.sldSz.cy} (${presentation.sldSz.type})`);
console.log(`slides: ${pkg.slideParts.length}, masters: ${pkg.slideMasterParts.length}, layouts: ${pkg.slideLayoutParts.length}`);
for (const slide of pkg.slideParts) {
  console.log(`  ${slide.partName.name} on ${slide.slideLayoutPart.partName.name}`);
}

await writeFile(out, await pkg.save());
console.log(`wrote ${out}`);
