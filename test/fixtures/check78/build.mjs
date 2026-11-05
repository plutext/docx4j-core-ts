// The jobs for Word check 78 (getFileAsync against SaveAs). Run from the repository root after `npm run build`:
//
//   node test/fixtures/check78/build.mjs <directory>
//
// writes one folder a job into that directory, each with its input and job.json, run in the order of `check78Jobs`:
//   78a  a Word-made file holding tracked changes, comments, a footnote, a table and a field: every format
//   78b  WD-001's pair, a character outside the code page at offset 200: Save As Plain Text succeeds there
//   78c  the same at offset 300, where Save As Plain Text fails with no message, so `txt` is a job of its own
// Each input goes through the `check` verb and the `save` verb, for the comparison. The word_defects.md rule is the boundary 256.
import { copyFile, mkdir, readFile, writeFile } from 'node:fs/promises';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { WordprocessingMLPackage } from '../../../dist/index.mjs';

const here = dirname(fileURLToPath(import.meta.url));
const STORED = '₂';

/** One paragraph, the character at the 0-based offset, on padding that stays inside the code page. */
async function wd001(offset) {
  const pkg = await WordprocessingMLPackage.createPackage();
  const padding = (n) => 'x'.repeat(n);
  pkg.body.insertParagraph(`${padding(offset)}${STORED}${padding(20)}`, 'End');
  return pkg.save();
}

export async function check78Inputs() {
  return {
    '78a-file-output.docx': await readFile(join(here, '..', 'revisions', 'review', 'review-for-word-word15.docx')),
    '78b-wd001-at-200.docx': await wd001(200),
    '78c-wd001-at-300.docx': await wd001(300),
  };
}

export const check78Jobs = [
  { folder: '78a-check', input: '78a-file-output.docx', verb: 'check' },
  { folder: '78a-save', input: '78a-file-output.docx', verb: 'save', formats: ['docx', 'txt', 'pdf'] },
  { folder: '78b-check', input: '78b-wd001-at-200.docx', verb: 'check' },
  { folder: '78b-save', input: '78b-wd001-at-200.docx', verb: 'save', formats: ['txt'] },
  { folder: '78c-check', input: '78c-wd001-at-300.docx', verb: 'check' },
  { folder: '78c-save', input: '78c-wd001-at-300.docx', verb: 'save', formats: ['txt'] },
];

if (process.argv[1] && fileURLToPath(import.meta.url) === process.argv[1]) {
  const dir = process.argv[2];
  if (!dir) throw new Error('Give the directory to write into');
  const inputs = await check78Inputs();
  for (const job of check78Jobs) {
    const folder = join(dir, job.folder);
    await mkdir(folder, { recursive: true });
    await writeFile(join(folder, job.input), inputs[job.input]);
    const spec = job.verb === 'check'
      ? { verb: 'check', input: job.input, checkHtml: 'check78.html', checkScript: 'check78-script.js', timeoutSeconds: 180 }
      : { verb: 'save', input: job.input, formats: job.formats, timeoutSeconds: 120 };
    if (job.verb === 'check') for (const f of ['check78.html', 'check78-script.js']) await copyFile(join(here, f), join(folder, f));
    await writeFile(join(folder, 'job.json'), `${JSON.stringify(spec, null, 2)}\n`);
  }
  console.log(`${check78Jobs.length} jobs into ${dir}`);
}
