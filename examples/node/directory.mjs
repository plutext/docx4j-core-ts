// The unzipped container (docx4j's UnzippedPartStore): write a package out as a directory of
// files, read it back, and check that every untouched part came through byte for byte — which is
// what makes a directory useful for diffing two documents, or for putting one under version
// control.
//
//   node examples/node/directory.mjs [file.docx] [directory]
//
// `@docx4j/core-ts/node` is the only subpath that imports `node:` builtins; everything else runs
// in a browser or a Word add-in. Published, the imports are `from '@docx4j/core-ts'` and
// `from '@docx4j/core-ts/node'`; here they are the built dist of this checkout.
import { readFile, rm } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import { tmpdir } from 'node:os';
import { OpcPackage, WordprocessingMLPackage, ZipPartStore } from '../../dist/index.mjs';
import { DirectoryPartStore, DirectoryPartSink } from '../../dist/node/index.mjs';

const here = dirname(fileURLToPath(import.meta.url));
const file = process.argv[2] ?? join(here, '..', '..', 'test', 'fixtures', 'loadAndSave.docx');
const dir = process.argv[3] ?? join(tmpdir(), 'core-ts-unzipped');

const bytes = new Uint8Array(await readFile(file));

// zip -> directory
await rm(dir, { recursive: true, force: true });
const pkg = await OpcPackage.load(bytes);
await pkg.saveTo(new DirectoryPartSink(dir));
console.log(`unzipped ${pkg.parts.size} parts into ${dir}`);

// directory -> package -> zip
const store = await DirectoryPartStore.open(dir);
const back = await WordprocessingMLPackage.load(store);
console.log(`loaded back: ${[...store.partNames()].length} files, ${back.parts.size} parts`);
const out = new ZipPartStore(await back.save());

// Part by part. The relationships parts and [Content_Types].xml are written afresh by every
// save, a zip round trip included; everything else must be the bytes it was loaded with.
const source = new ZipPartStore(bytes);
let same = 0;
let regenerated = 0;
for (const name of source.partNames()) {
  if (name === '[Content_Types].xml' || name.endsWith('.rels')) { regenerated++; continue; }
  const a = source.loadSync(name);
  const b = out.loadSync(name);
  if (b && a.length === b.length && a.every((v, i) => v === b[i])) same++;
  else console.log(`  DIFFERS: ${name}`);
}
console.log(`byte-identical: ${same} parts (${regenerated} regenerated: the rels and the content types)`);
