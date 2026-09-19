// The add-in's own edit, run in Node against a fixture: the taskpane's "without Word" button
// with no browser. This is what makes an add-in's logic testable (test/examples.test.mjs runs it).
//
//   node examples/office-addin/run-in-node.mjs [file.docx]
//
// In a real add-in the imports are `from '@docx4j/core-ts'` and `from '@docx4j/core-ts/office-js'`.
import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import { WordprocessingMLPackage } from '../../dist/index.mjs';
import { Word } from '../../dist/office-js/index.mjs';
import { edit, applyEdit } from './edit.mjs';

const here = dirname(fileURLToPath(import.meta.url));
const file = process.argv[2] ?? join(here, '..', '..', 'test', 'fixtures', 'tracked-changes.docx');
// The word to highlight. The taskpane's default is 'draft'; this fixture has no draft in it.
const term = process.argv[3] ?? 'document';

// 1. through the shim, as the taskpane's second button does
const pkg = await WordprocessingMLPackage.load(new Uint8Array(await readFile(file)));
const highlighted = await Word.run(pkg, (context) => edit(context.document.body, term));
console.log(`highlighted ${highlighted} occurrences`);
console.log(`first paragraph style: ${pkg.body.paragraphs[0].style}`);

// 2. through the flat OPC round trip, as the taskpane's first button does: `getOoxml()` of the
// whole document in, the edited package out, which is what `insertOoxml()` would take.
const ooxml = await (await WordprocessingMLPackage.load(new Uint8Array(await readFile(file)))).saveFlatOpc();
const edited = await applyEdit(ooxml, term);
console.log(`flat OPC in ${ooxml.length} bytes, out ${edited.length} bytes`);
const back = await WordprocessingMLPackage.load(edited);
console.log(`round trip style: ${(await back.getBody()).paragraphs[0].style}`);
