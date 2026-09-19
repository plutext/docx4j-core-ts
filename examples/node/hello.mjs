// A document from nothing: create, add a heading and a paragraph, save.
//
//   node examples/node/hello.mjs [out.docx]
//
// Published, the import is `from '@docx4j/core-ts'`; here it is the built dist of this checkout.
import { writeFile } from 'node:fs/promises';
import { WordprocessingMLPackage } from '../../dist/index.mjs';

const out = process.argv[2] ?? 'hello.docx';

const pkg = await WordprocessingMLPackage.createPackage({ pageSize: 'A4' });
const body = pkg.body;

const heading = body.insertParagraph('Hello World', 'End');
heading.styleBuiltIn = 'Heading1';

const p = body.insertParagraph('Written by @docx4j/core-ts, with no Word in sight.', 'End');
p.alignment = 'Justified';
const run = p.insertText(' Emphasised.', 'End');
run.font.italic = true;
run.font.color = '#B22222';

await writeFile(out, await pkg.save());

console.log(`paragraphs: ${body.paragraphs.length}`);
console.log(`first paragraph style: ${body.paragraphs[0].style}`);
console.log(`wrote ${out}`);
