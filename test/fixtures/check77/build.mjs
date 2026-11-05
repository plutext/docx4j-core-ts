// The files for Word check 77 (tracked changes through insertOoxml; section 14 of the portfolio's
// Word-LLM-interaction-models.docx). Run from the repository root after `npm run build`:
//
//   node test/fixtures/check77/build.mjs <directory>
//
// writes into that directory the bridge's job: 77-tracked-insertooxml.docx, check77.html, check77-script.js and job.json.
// Each case is three paragraphs: "BEGIN Cnn mode target payload", the target, "END Cnn". The markers show a replacement
// reaching past its paragraph. The script reads the cases from the file, so the two can't disagree.
import { copyFile, writeFile } from 'node:fs/promises';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { WordprocessingMLPackage } from '../../../dist/index.mjs';

const here = dirname(fileURLToPath(import.meta.url));
const PAYLOADS = ['plain', 'ins', 'del', 'redline', 'twoparas'];
const TARGETS = ['word', 'content', 'whole', 'para'];

/** Tracking off and on over every target and payload, then what TrackMineOnly does, then Word's own marks. */
export function check77Cases() {
  const cases = [];
  for (const mode of ['off', 'on']) for (const target of TARGETS) for (const payload of PAYLOADS) cases.push({ mode, target, payload });
  cases.push({ mode: 'mine', target: 'word', payload: 'redline' }, { mode: 'on', target: 'word', payload: 'self' }, { mode: 'on', target: 'word', payload: 'selfdelete' });
  return cases.map((c, i) => ({ ...c, id: `C${String(i + 1).padStart(2, '0')}` }));
}

export async function check77Bytes() {
  const pkg = await WordprocessingMLPackage.createPackage();
  const body = pkg.body;
  body.insertParagraph('Check 77: tracked changes through insertOoxml', 'End');
  for (const c of check77Cases()) {
    body.insertParagraph(`BEGIN ${c.id} ${c.mode} ${c.target} ${c.payload}`, 'End');
    body.insertParagraph(`${c.id} before TARGET after`, 'End');
    body.insertParagraph(`END ${c.id}`, 'End');
  }
  body.insertParagraph('LAST LINE', 'End');
  return pkg.save();
}

if (process.argv[1] && fileURLToPath(import.meta.url) === process.argv[1]) {
  const dir = process.argv[2];
  if (!dir) throw new Error('Give the directory to write into');
  const input = '77-tracked-insertooxml.docx';
  await writeFile(join(dir, input), await check77Bytes());
  for (const f of ['check77.html', 'check77-script.js']) await copyFile(join(here, f), join(dir, f));
  await writeFile(join(dir, 'job.json'), `${JSON.stringify({ verb: 'check', input, checkHtml: 'check77.html', checkScript: 'check77-script.js', timeoutSeconds: 180 }, null, 2)}\n`);
  console.log(`${check77Cases().length} cases into ${dir}`);
}
