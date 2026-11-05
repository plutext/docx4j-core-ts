// The jobs for Word check 79 (granularity: reaching a paragraph, and what a replacement costs). Run from the repository root
// after `npm run build`:
//
//   node test/fixtures/check79/build.mjs <directory>
//
// writes one folder a job into that directory, each with its input and job.json, run in the order of `check79Jobs`:
//   79a  every kind of paragraph with a w14:paraId of its own: check79a twice (a session's ids against the next's), and two saves
//   79n  the same without ids: check79a, and two saves (phase 1 saw ids differ between saves; did those sources carry any?)
//   79b  long documents built for the purpose, at each size in `SIZES`: check79b
// The inputs are built here, never a customer's. The follow-up is apart, since it needs Word's save of 79a:
//
//   node test/fixtures/check79/build.mjs <directory> --follow-up <79a-ids.word.docx>
//
//   79c  Word's save of 79a, through check79c and two saves: does a file Word wrote keep its ids?
//   79d  79a's ids with one kind of content each, to find what renumbers them at open: 79a's all changed, 79b's didn't
//
//   node test/fixtures/check79/build.mjs <directory> --second-follow-up <79c-word-saved.docx>
//
//   79e  Word's save again, for the calls 79c's run couldn't reach; 79f, a table whose row has an id, and a plain file
//        missing one id, since a built table's row had none and all 14 changed
import { copyFile, mkdir, readFile, writeFile } from 'node:fs/promises';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { unzipSync, zipSync, strToU8 } from 'fflate';
import { WordprocessingMLPackage } from '../../../dist/index.mjs';

const here = dirname(fileURLToPath(import.meta.url));
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
const V = 'urn:schemas-microsoft-com:vml';
const declaration = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
const SECT = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>';
export const SIZES = [{ paragraphs: 100, budget: 100 }, { paragraphs: 1000, budget: 200 }, { paragraphs: 5000, budget: 500 }, { paragraphs: 20000, budget: 1200 }];

const documentXml = (blocks) => `${declaration}<w:document xmlns:w="${W}" xmlns:w14="${W14}" xmlns:mc="${MC}" xmlns:v="${V}" mc:Ignorable="w14"><w:body>${blocks}${SECT}</w:body></w:document>`;

async function withDocument(xml) {
  const base = unzipSync(await (await WordprocessingMLPackage.createPackage()).save());
  // a fixed time makes the file the same on every build
  return zipSync({ ...base, 'word/document.xml': strToU8(xml) }, { mtime: new Date(Date.UTC(2026, 10, 6)) });
}

/** The paragraphs check79a edits, then a table, a control and a text box, whose paragraphs Office JS and the XML may count apart. */
function idsDocument(withIds, only, { rowIds = false, gap = false } = {}) {
  const has = (kind) => only === undefined || only === kind;
  let n = 0;
  // `gap`: the fifth paragraph alone has none. `rowIds`: a row has one, as Word writes a w:tr
  const attrs = () => { n += 1; return withIds && !(gap && n === 5) ? ` w14:paraId="${(0x10000000 + n).toString(16).toUpperCase()}" w14:textId="77777777"` : ''; };
  const p = (text) => `<w:p${attrs()}><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
  const cell = (text) => `<w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>${p(text)}</w:tc>`;
  return documentXml([
    p('P01 left alone'), p('P02 text at its end'), p('P03 replaced, carrying its own id'), p('P04 replaced, carrying none'),
    p('P05 content replaced, its own id'), p('P06 read and put back'), p('P07 a paragraph after it'), p('P08 before SPLIT after'),
    p('P09 deleted'), p('P10 made bold'), p('P11 edited under tracking'),
    has('table') ? `<w:tbl><w:tblPr><w:tblW w:w="8000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid><w:tr${rowIds ? ' w14:paraId="2000000A" w14:textId="77777777"' : ''}>${cell('P12 in a cell')}${cell('P13 in a cell')}</w:tr></w:tbl>` : '',
    has('control') ? `<w:sdt><w:sdtPr><w:id w:val="791"/></w:sdtPr><w:sdtContent>${p('P14 in a control')}</w:sdtContent></w:sdt>` : '',
    has('textbox') ? `<w:p${attrs()}><w:r><w:t xml:space="preserve">P15 holding a text box </w:t></w:r><w:r><w:pict><v:rect style="width:200pt;height:40pt"><v:textbox><w:txbxContent>${p('T01 in a text box')}</w:txbxContent></v:textbox></v:rect></w:pict></w:r></w:p>` : '',
    p('P16 left alone at the end'),
  ].join(''));
}

const FILLER = 'The quick brown fox jumps over the lazy dog and keeps running across the long field until the light fades behind the hills.';

/** A long document: each paragraph a label in bold and a sentence, each with an id. */
function longDocument({ paragraphs, budget }) {
  const blocks = [`<w:p w14:paraId="0FFFFFFF" w14:textId="77777777"><w:r><w:t>Check 79b: ${paragraphs} paragraphs, budget ${budget}</w:t></w:r></w:p>`];
  for (let i = 1; i < paragraphs; i++) {
    const id = (0x10000000 + i).toString(16).toUpperCase();
    blocks.push(`<w:p w14:paraId="${id}" w14:textId="77777777"><w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">N${String(i).padStart(5, '0')} </w:t></w:r><w:r><w:t>${FILLER}</w:t></w:r></w:p>`);
  }
  return documentXml(blocks.join(''));
}

export async function check79Inputs() {
  const inputs = { '79a-ids.docx': await withDocument(idsDocument(true)), '79n-no-ids.docx': await withDocument(idsDocument(false)) };
  for (const size of SIZES) inputs[`79b-${size.paragraphs}.docx`] = await withDocument(longDocument(size));
  return inputs;
}

const check = (folder, input, script, timeoutSeconds) => ({ folder, input, spec: { verb: 'check', input, checkHtml: 'check79.html', checkScript: script, timeoutSeconds } });
const save = (folder, input) => ({ folder, input, spec: { verb: 'save', input, formats: ['docx'], timeoutSeconds: 120 } });
export const check79Jobs = [
  check('79a-check-1', '79a-ids.docx', 'check79a-script.js', 180), check('79a-check-2', '79a-ids.docx', 'check79a-script.js', 180),
  save('79a-save-1', '79a-ids.docx'), save('79a-save-2', '79a-ids.docx'),
  check('79n-check', '79n-no-ids.docx', 'check79a-script.js', 180),
  save('79n-save-1', '79n-no-ids.docx'), save('79n-save-2', '79n-no-ids.docx'),
  ...SIZES.map((s) => check(`79b-${s.paragraphs}`, `79b-${s.paragraphs}.docx`, 'check79b-script.js', s.budget + 120)),
];

export async function check79FollowUpInputs(wordSave) {
  const inputs = { '79c-word-saved.docx': wordSave };
  for (const kind of ['plain', 'table', 'control', 'textbox']) inputs[`79d-${kind}.docx`] = await withDocument(idsDocument(true, kind));
  return inputs;
}

export async function check79SecondFollowUpInputs(wordSave) {
  return {
    '79e-word-saved.docx': wordSave,
    '79f-row-ids.docx': await withDocument(idsDocument(true, 'table', { rowIds: true })),
    '79f-one-missing.docx': await withDocument(idsDocument(true, 'plain', { gap: true })),
  };
}

export const check79SecondFollowUpJobs = [
  check('79e-reach', '79e-word-saved.docx', 'check79e-script.js', 120),
  check('79f-row-ids', '79f-row-ids.docx', 'check79e-script.js', 120),
  check('79f-one-missing', '79f-one-missing.docx', 'check79e-script.js', 120),
];

export const check79FollowUpJobs = [
  check('79c-check', '79c-word-saved.docx', 'check79c-script.js', 180),
  save('79c-save-1', '79c-word-saved.docx'), save('79c-save-2', '79c-word-saved.docx'),
  check('79d-plain', '79d-plain.docx', 'check79c-script.js', 180),
  save('79d-plain-save-1', '79d-plain.docx'), save('79d-plain-save-2', '79d-plain.docx'),
  check('79d-table', '79d-table.docx', 'check79c-script.js', 180),
  check('79d-control', '79d-control.docx', 'check79c-script.js', 180),
  check('79d-textbox', '79d-textbox.docx', 'check79c-script.js', 180),
];

if (process.argv[1] && fileURLToPath(import.meta.url) === process.argv[1]) {
  const dir = process.argv[2];
  if (!dir) throw new Error('Give the directory to write into');
  const followUp = process.argv[3] === '--follow-up';
  const second = process.argv[3] === '--second-follow-up';
  const given = followUp || second ? await readFile(process.argv[4]) : undefined;
  const inputs = second ? await check79SecondFollowUpInputs(given) : followUp ? await check79FollowUpInputs(given) : await check79Inputs();
  const jobs = second ? check79SecondFollowUpJobs : followUp ? check79FollowUpJobs : check79Jobs;
  for (const job of jobs) {
    const folder = join(dir, job.folder);
    await mkdir(folder, { recursive: true });
    await writeFile(join(folder, job.input), inputs[job.input]);
    if (job.spec.verb === 'check') for (const f of ['check79.html', job.spec.checkScript]) await copyFile(join(here, f), join(folder, f));
    await writeFile(join(folder, 'job.json'), `${JSON.stringify(job.spec, null, 2)}\n`);
  }
  console.log(`${jobs.length} jobs into ${dir}`);
}
