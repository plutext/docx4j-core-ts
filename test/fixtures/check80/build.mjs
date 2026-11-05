// The jobs for Word check 80 (CR-002 section 66): does a file this package edited keep its ids in Word 365? Run from the
// repository root after `npm run build`:
//
//   node test/fixtures/check80/build.mjs <directory>
//
// writes one folder a job into that directory, each with its input and job.json. Each input is Word's own save
// (test/check79/79c-save-1.word.docx) edited here: 80a rows, 80b a table, 80c columns, 80d all three. 80e is 80d with one new
// row's id taken off, the control: check 79 says that file is renumbered. 80d also goes through a save.
import { copyFile, mkdir, readFile, writeFile } from 'node:fs/promises';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import { WordprocessingMLPackage } from '../../../dist/index.mjs';

const here = dirname(fileURLToPath(import.meta.url));

async function edited(edit) {
  const pkg = await WordprocessingMLPackage.load(new Uint8Array(await readFile(join(here, '..', '..', 'check79', '79c-save-1.word.docx'))));
  await pkg.getBody();
  edit(pkg);
  return pkg.save();
}

const rows = (pkg) => pkg.body.tables[0].addRows('End', 2, [['R1a', 'R1b'], ['R2a', 'R2b']]);
const table = (pkg) => pkg.body.insertTable(2, 2, 'End', [['T1a', 'T1b'], ['T2a', 'T2b']]);
const columns = (pkg) => pkg.body.tables[0].addColumns('End', 1, [['C1']]);

/** `bytes` with the id off the last row: the control. */
function withoutARowId(bytes) {
  const parts = unzipSync(bytes);
  const xml = strFromU8(parts['word/document.xml']);
  const rowsHeld = [...xml.matchAll(/<w:tr\b[^>]*>/g)];
  const last = rowsHeld[rowsHeld.length - 1];
  const bare = last[0].replace(/ w14:paraId="[^"]*"/, '').replace(/ w14:textId="[^"]*"/, '');
  if (bare === last[0]) throw new Error('that row has no id to take off');
  const taken = xml.slice(0, last.index) + bare + xml.slice(last.index + last[0].length);
  return zipSync({ ...parts, 'word/document.xml': strToU8(taken) }, { mtime: new Date(Date.UTC(2026, 10, 6)) });
}

export async function check80Inputs() {
  const all = await edited((pkg) => { rows(pkg); columns(pkg); table(pkg); });
  return {
    '80a-rows.docx': await edited(rows),
    '80b-table.docx': await edited(table),
    '80c-columns.docx': await edited(columns),
    '80d-all.docx': all,
    '80e-control.docx': withoutARowId(all),
  };
}

const check = (folder, input) => ({ folder, input, spec: { verb: 'check', input, checkHtml: 'check80.html', checkScript: 'check80-script.js', timeoutSeconds: 120 } });
export const check80Jobs = [
  check('80a-rows', '80a-rows.docx'), check('80b-table', '80b-table.docx'), check('80c-columns', '80c-columns.docx'),
  check('80d-all', '80d-all.docx'),
  { folder: '80d-save', input: '80d-all.docx', spec: { verb: 'save', input: '80d-all.docx', formats: ['docx'], timeoutSeconds: 120 } },
  check('80e-control', '80e-control.docx'),
];

if (process.argv[1] && fileURLToPath(import.meta.url) === process.argv[1]) {
  const dir = process.argv[2];
  if (!dir) throw new Error('Give the directory to write into');
  const inputs = await check80Inputs();
  for (const job of check80Jobs) {
    const folder = join(dir, job.folder);
    await mkdir(folder, { recursive: true });
    await writeFile(join(folder, job.input), inputs[job.input]);
    if (job.spec.verb === 'check') for (const f of ['check80.html', 'check80-script.js']) await copyFile(join(here, f), join(folder, f));
    await writeFile(join(folder, 'job.json'), `${JSON.stringify(job.spec, null, 2)}\n`);
  }
  console.log(`${check80Jobs.length} jobs into ${dir}`);
}
