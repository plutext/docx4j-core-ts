// The files for Word check 76 (CR-026 section 4, 2026-11-05; Jason: "walk me through word check 76"). Run from the repository
// root after `npm run build`:
//
//   node test/fixtures/check76/build.mjs <directory>
//
// writes the files into that directory, which has to be given: not this one, since most of them fail a part here until
// CR-026 is built, and CR-016's gate walks everything under test/fixtures/. `check76Files` gives a test the same bytes. Each
// is a created package with its main document replaced: what does Word 365 do with an element or attribute it has no way to
// know, in a namespace listed in mc:Ignorable, in one that isn't, and as an unknown name in the w: namespace? Each file says
// on its page what it holds; its middle paragraph reads TARGET TEXT and its last LAST LINE, so a loss shows.
import { writeFile } from 'node:fs/promises';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { unzipSync, zipSync, strToU8 } from 'fflate';
import { WordprocessingMLPackage } from '../../../dist/index.mjs';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
const EXT = 'urn:docx4j:check76';
const declaration = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

const run = (text) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
const p = (text) => `<w:p>${run(text)}</w:p>`;
const TARGET = run('TARGET TEXT');
const OPAQUE = '<ext:opaque ext:note="check 76"/>';
const MADE_UP = '<w:madeUp w:val="1"/>';

/** The files: `name`, what it holds (`asks`), whether `ext` is in `mc:Ignorable`, and the content between the first and last paragraphs. */
export const FILES = [
  { name: '76-control', asks: 'nothing unknown: ext declared and ignorable, and used nowhere', ignorable: true,
    content: `<w:p>${TARGET}</w:p>` },
  { name: '76a-ignorable-in-sdtPr', asks: 'an ignorable element in a w:sdtPr', ignorable: true,
    content: `<w:sdt><w:sdtPr><w:alias w:val="Check 76"/><w:id w:val="760001"/>${OPAQUE}</w:sdtPr><w:sdtContent><w:p>${TARGET}</w:p></w:sdtContent></w:sdt>` },
  { name: '76b-ignorable-in-p', asks: 'an ignorable element in a w:p', ignorable: true,
    content: `<w:p>${OPAQUE}${TARGET}</w:p>` },
  { name: '76c-ignorable-in-r', asks: 'an ignorable element in a w:r', ignorable: true,
    content: `<w:p><w:r>${OPAQUE}<w:t>TARGET TEXT</w:t></w:r></w:p>` },
  { name: '76d-ignorable-in-tc', asks: 'an ignorable element in a w:tc', ignorable: true,
    content: '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc>'
      + `<w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>${OPAQUE}<w:p>${TARGET}</w:p></w:tc></w:tr></w:tbl>` },
  { name: '76e-not-ignorable-in-p', asks: 'an element in a w:p, its namespace declared and not ignorable', ignorable: false,
    content: `<w:p>${OPAQUE}${TARGET}</w:p>` },
  { name: '76f-madeup-in-p', asks: 'an unknown w: element in a w:p', ignorable: true,
    content: `<w:p>${MADE_UP}${TARGET}</w:p>` },
  { name: '76g-madeup-in-pPr', asks: 'an unknown w: element in a w:pPr', ignorable: true,
    content: `<w:p><w:pPr>${MADE_UP}</w:pPr>${TARGET}</w:p>` },
  { name: '76h-ignorable-attribute', asks: 'an ignorable attribute on a w:p', ignorable: true,
    content: `<w:p ext:note="check 76">${TARGET}</w:p>` },
  { name: '76i-not-ignorable-attribute', asks: 'an attribute on a w:p, its namespace declared and not ignorable', ignorable: false,
    content: `<w:p ext:note="check 76">${TARGET}</w:p>` },
  { name: '76j-madeup-attribute', asks: 'an unknown w: attribute on a w:p', ignorable: true,
    content: `<w:p w:madeUp="1">${TARGET}</w:p>` },
  { name: '76k-ignorable-holding-a-run', asks: 'an ignorable element in a w:p holding a run: do the words HIDDEN WORDS show?', ignorable: true,
    content: `<w:p>${run('TARGET TEXT, then ')}<ext:opaque>${run('HIDDEN WORDS')}</ext:opaque>${run(' and the rest.')}</w:p>` },
  // Added after jsonix-CR-015 (2026-11-05): the hosts its branch doesn't reach, CR-026 section 3.4.
  { name: '76l-ignorable-in-b', asks: 'an ignorable element in a w:b, on bold text', ignorable: true,
    content: `<w:p><w:r><w:rPr><w:b>${OPAQUE}</w:b></w:rPr><w:t>TARGET TEXT</w:t></w:r></w:p>` },
  { name: '76m-ignorable-in-t', asks: 'an empty ignorable element inside a w:t, between TARGET and TEXT', ignorable: true,
    content: `<w:p><w:r><w:t xml:space="preserve">TARGET ${OPAQUE}TEXT</w:t></w:r></w:p>` },
  { name: '76n-ignorable-with-text-in-t', asks: 'an ignorable element inside a w:t holding text: do the words HIDDEN WORDS show?', ignorable: true,
    content: '<w:p><w:r><w:t xml:space="preserve">TARGET <ext:opaque>HIDDEN WORDS</ext:opaque>TEXT</w:t></w:r></w:p>' },
  // Round 2 (2026-11-05; Jason: "let's do the second round of files"; CR-026 section 4.3). What 76e refused, in other hosts:
  { name: '76o-not-ignorable-in-sdtPr', asks: 'an element in a w:sdtPr, its namespace declared and not ignorable', ignorable: false,
    content: `<w:sdt><w:sdtPr><w:alias w:val="Check 76"/><w:id w:val="760002"/>${OPAQUE}</w:sdtPr><w:sdtContent><w:p>${TARGET}</w:p></w:sdtContent></w:sdt>` },
  { name: '76p-not-ignorable-in-pPr', asks: 'an element in a w:pPr, its namespace declared and not ignorable', ignorable: false,
    content: `<w:p><w:pPr>${OPAQUE}</w:pPr>${TARGET}</w:p>` },
  { name: '76q-not-ignorable-in-r', asks: 'an element in a w:r, its namespace declared and not ignorable', ignorable: false,
    content: `<w:p><w:r>${OPAQUE}<w:t>TARGET TEXT</w:t></w:r></w:p>` },
  { name: '76r-not-ignorable-in-tc', asks: 'an element in a w:tc, its namespace declared and not ignorable', ignorable: false,
    content: '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc>'
      + `<w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>${OPAQUE}<w:p>${TARGET}</w:p></w:tc></w:tr></w:tbl>` },
  { name: '76s-not-ignorable-in-b', asks: 'an element in a w:b, on bold text, its namespace declared and not ignorable', ignorable: false,
    content: `<w:p><w:r><w:rPr><w:b>${OPAQUE}</w:b></w:rPr><w:t>TARGET TEXT</w:t></w:r></w:p>` },
  // An unknown w: name holding something: 76k's question, which 76f's empty w:madeUp doesn't ask.
  { name: '76t-madeup-holding-a-run', asks: 'an unknown w: element in a w:p holding a run: do the words HIDDEN WORDS show?', ignorable: true,
    content: `<w:p>${run('TARGET TEXT, then ')}<w:madeUp>${run('HIDDEN WORDS')}</w:madeUp>${run(' and the rest.')}</w:p>` },
  // Inside a w:t, for what CR-021's rule 2 does with each (CR-021 section 11). A br in no namespace is the shape of a real file,
  // docx4j's survey sample 01af91808eac, and a span the same with text in it.
  { name: '76u-madeup-with-text-in-t', asks: 'an unknown w: element inside a w:t holding text: do the words HIDDEN WORDS show?', ignorable: true,
    content: '<w:p><w:r><w:t xml:space="preserve">TARGET <w:madeUp>HIDDEN WORDS</w:madeUp>TEXT</w:t></w:r></w:p>' },
  { name: '76v-not-ignorable-with-text-in-t', asks: 'an element inside a w:t holding text, its namespace not ignorable: do the words HIDDEN WORDS show?', ignorable: false,
    content: '<w:p><w:r><w:t xml:space="preserve">TARGET <ext:opaque>HIDDEN WORDS</ext:opaque>TEXT</w:t></w:r></w:p>' },
  { name: '76w-ignorable-holding-a-break-in-t', asks: 'an ignorable element inside a w:t holding a w:br: does TEXT start a new line?', ignorable: true,
    content: '<w:p><w:r><w:t xml:space="preserve">TARGET <ext:opaque><w:br/></ext:opaque>TEXT</w:t></w:r></w:p>' },
  { name: '76x-no-namespace-with-text-in-t', asks: 'an element in no namespace inside a w:t holding text: do the words HIDDEN WORDS show?', ignorable: true,
    content: '<w:p><w:r><w:t xml:space="preserve">TARGET <span>HIDDEN WORDS</span>TEXT</w:t></w:r></w:p>' },
  { name: '76y-no-namespace-break-in-t', asks: 'a br in no namespace inside a w:t: does TEXT start a new line?', ignorable: true,
    content: '<w:p><w:r><w:t xml:space="preserve">TARGET <br/>TEXT</w:t></w:r></w:p>' },
  { name: '76z-tab-in-t', asks: 'a w:tab inside a w:t: is there a tab between TARGET and TEXT?', ignorable: true,
    content: '<w:p><w:r><w:t xml:space="preserve">TARGET<w:tab/>TEXT</w:t></w:r></w:p>' },
];

/** The files, { name, asks, ignorable, document, bytes }, over a created package's parts. */
export async function check76Files() {
  const base = unzipSync(await (await WordprocessingMLPackage.createPackage()).save());
  return FILES.map((f) => {
    const document = `${declaration}<w:document xmlns:w="${W}" xmlns:mc="${MC}" xmlns:ext="${EXT}"${f.ignorable ? ' mc:Ignorable="ext"' : ''}><w:body>`
      + p(`Check 76, ${f.name}: ${f.asks}.`) + f.content + p('LAST LINE.')
      + '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>'
      + '</w:body></w:document>';
    // a fixed time makes the file the same on every build
    const bytes = zipSync({ ...base, 'word/document.xml': strToU8(document) }, { mtime: new Date(Date.UTC(2026, 10, 5)) });
    return { name: f.name, asks: f.asks, ignorable: f.ignorable, document, bytes };
  });
}

if (process.argv[1] && fileURLToPath(import.meta.url) === process.argv[1]) {
  const dir = process.argv[2];
  if (!dir) throw new Error('Give a directory: not test/fixtures/, whose gate these files would fail until CR-026 is built.');
  const files = await check76Files();
  for (const f of files) await writeFile(join(dir, `${f.name}.docx`), f.bytes);
  console.log(`wrote ${files.length} files to ${dir}`);
}
