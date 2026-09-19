// Report on a .docx: its outline, its tracked changes, its comments and the list labels Word
// paints — the reads that need docx4j's resolution, not just the markup.
//
//   node examples/node/report.mjs [file.docx]
//
// Published, the import is `from '@docx4j/core-ts'`; here it is the built dist of this checkout.
import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import { WordprocessingMLPackage } from '../../dist/index.mjs';

const here = dirname(fileURLToPath(import.meta.url));
const file = process.argv[2] ?? join(here, '..', '..', 'test', 'fixtures', 'loadAndSave.docx');

const pkg = await WordprocessingMLPackage.load(new Uint8Array(await readFile(file)));

// The outline: every paragraph and table of the body, then of each header and footer, with the
// address you can hand back to pkg.paragraphAt().
const outline = await pkg.outline();
console.log(`outline: ${outline.paragraphs.length} paragraphs, ${outline.tables.length} tables, ` +
  `${outline.headers.length} headers, ${outline.footers.length} footers`);
for (const p of outline.paragraphs.slice(0, 5)) {
  console.log(`  ${p.address}  [${p.style}]  ${JSON.stringify(p.text.slice(0, 60))}`);
}

// Tracked changes: Word's own revision markup, as Word's review pane shows it.
const changes = pkg.getTrackedChanges();
console.log(`tracked changes: ${changes.length}`);
for (const c of changes.slice(0, 5)) {
  console.log(`  ${c.type} by ${c.author}${c.date ? ` on ${c.date.toISOString().slice(0, 10)}` : ''}: ${JSON.stringify(c.text.slice(0, 40))}`);
}

// Comments, threads and all.
const comments = await pkg.body.getComments();
console.log(`comments: ${comments.length}`);
for (const c of comments) {
  console.log(`  ${c.authorName}${c.resolved ? ' (resolved)' : ''}: ${JSON.stringify(c.content.slice(0, 60))}` +
    (c.replies.length ? ` (+${c.replies.length} replies)` : ''));
}

// List labels: the numbering emulator counts the document in order, so these are the labels Word
// paints, not the numbering definitions.
const items = pkg.body.paragraphs.filter((p) => p.isListItem);
console.log(`list items: ${items.length}`);
for (const p of items.slice(0, 8)) {
  console.log(`  ${p.listItem.listString}\t${JSON.stringify(p.text.slice(0, 50))}`);
}
