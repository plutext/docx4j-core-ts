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
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import { WordprocessingMLPackage } from '../../../dist/index.mjs';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const EXT = 'urn:docx4j:check76';
const declaration = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

const run = (text) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
const p = (text) => `<w:p>${run(text)}</w:p>`;
const TARGET = run('TARGET TEXT');
const OPAQUE = '<ext:opaque ext:note="check 76"/>';
const MADE_UP = '<w:madeUp w:val="1"/>';

// Round 4's hosts, each given what to do with what w:madeUp holds: `made` wraps it, `bare` leaves it standing, for the control.
const made = (inner) => `<w:madeUp>${inner}</w:madeUp>`;
const bare = (inner) => inner;
const cell = (inner) => `<w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>${inner}</w:tc>`;
const table = (columns, rows) => `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid>${'<w:gridCol w:w="4000"/>'.repeat(columns)}</w:tblGrid>${rows}</w:tbl>`;
const control = (id, inner) => `<w:sdt><w:sdtPr><w:id w:val="${id}"/></w:sdtPr><w:sdtContent>${inner}</w:sdtContent></w:sdt>`;
const inRuns = (hold) => `${run('TARGET TEXT, then ')}${hold(run('HIDDEN WORDS'))}${run(' and the rest.')}`;
const inBlocks = (hold) => `${p('TARGET TEXT')}${hold(p('HIDDEN WORDS'))}`;
const HOSTS = {
  tr: (hold) => table(2, `<w:tr>${cell(p('TARGET TEXT'))}${hold(cell(p('HIDDEN WORDS')))}</w:tr>`),
  blockControl: (hold) => control(7601, inBlocks(hold)),
  runControl: (hold) => `<w:p>${control(7602, inRuns(hold))}</w:p>`,
  rowControl: (hold) => table(1, control(7603, `<w:tr>${cell(p('TARGET TEXT'))}</w:tr>${hold(`<w:tr>${cell(p('HIDDEN WORDS'))}</w:tr>`)}`)),
  cellControl: (hold) => table(2, `<w:tr>${control(7604, `${cell(p('TARGET TEXT'))}${hold(cell(p('HIDDEN WORDS')))}`)}</w:tr>`),
  hyperlink: (hold) => `<w:p><w:hyperlink w:anchor="check76">${inRuns(hold)}</w:hyperlink></w:p>`,
  ins: (hold) => `<w:p><w:ins w:id="7605" w:author="Check 76" w:date="2026-11-05T00:00:00Z">${inRuns(hold)}</w:ins></w:p>`,
};
const COMMENTED = `<w:p><w:commentRangeStart w:id="0"/>${TARGET}<w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r></w:p>`;

// Round 5's, apart from HOSTS so that 76t18 stays as Word ran it. The URI isn't the ext namespace's, which the saves are searched for.
const tagged = (name, inner) => `<w:${name} w:uri="urn:check76:host" w:element="check">${inner}</w:${name}>`;
const MORE_HOSTS = {
  fldSimple: (hold) => `<w:p><w:fldSimple w:instr=" MERGEFIELD Check76 ">${inRuns(hold)}</w:fldSimple></w:p>`,
  smartTag: (hold) => `<w:p>${tagged('smartTag', inRuns(hold))}</w:p>`,
  blockCustomXml: (hold) => tagged('customXml', inBlocks(hold)),
  runCustomXml: (hold) => `<w:p>${tagged('customXml', inRuns(hold))}</w:p>`,
  rowCustomXml: (hold) => table(1, tagged('customXml', `<w:tr>${cell(p('TARGET TEXT'))}</w:tr>${hold(`<w:tr>${cell(p('HIDDEN WORDS'))}</w:tr>`)}`)),
  cellCustomXml: (hold) => table(2, `<w:tr>${tagged('customXml', `${cell(p('TARGET TEXT'))}${hold(cell(p('HIDDEN WORDS')))}`)}</w:tr>`),
};
const FOOTNOTED = `<w:p>${TARGET}<w:r><w:footnoteReference w:id="1"/></w:r></w:p>`;

/** The files: `name`, what it holds (`asks`), whether `ext` is in `mc:Ignorable`, the content between the first and last paragraphs, and any `header`, `comment` or `footnote` content, in a part of its own. */
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
  // Round 3 (2026-11-05; Jason: "commit, then let's do the third round of files"; CR-026 section 4.4, D7): 76t has Word read what an
  // unknown w: element in a w:p holds. Where else? Each holds the words HIDDEN WORDS, or a property, in what its host admits, the
  // host valid with or without it, so that only the element is in question.
  { name: '76t1-madeup-holding-p-in-body', asks: 'an unknown w: element in the w:body holding a w:p: do the words HIDDEN WORDS show?', ignorable: true,
    content: `<w:p>${TARGET}</w:p><w:madeUp>${p('HIDDEN WORDS')}</w:madeUp>` },
  { name: '76t2-madeup-holding-p-in-tc', asks: 'an unknown w: element in a w:tc holding a w:p: do the words HIDDEN WORDS show?', ignorable: true,
    content: '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc>'
      + `<w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr><w:madeUp>${p('HIDDEN WORDS')}</w:madeUp><w:p>${TARGET}</w:p></w:tc></w:tr></w:tbl>` },
  { name: '76t3-madeup-holding-t-in-r', asks: 'an unknown w: element in a w:r holding a w:t: do the words HIDDEN WORDS show?', ignorable: true,
    content: '<w:p><w:r><w:t xml:space="preserve">TARGET TEXT, then </w:t><w:madeUp><w:t>HIDDEN WORDS</w:t></w:madeUp><w:t xml:space="preserve"> and the rest.</w:t></w:r></w:p>' },
  { name: '76t4-madeup-holding-tr-in-tbl', asks: 'an unknown w: element in a w:tbl holding a w:tr: is there a row reading HIDDEN WORDS?', ignorable: true,
    content: '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>'
      + `<w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr><w:p>${TARGET}</w:p></w:tc></w:tr>`
      + `<w:madeUp><w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>${p('HIDDEN WORDS')}</w:tc></w:tr></w:madeUp></w:tbl>` },
  { name: '76t5-madeup-holding-jc-in-pPr', asks: 'an unknown w: element in a w:pPr holding a w:jc: is the next paragraph centred?', ignorable: true,
    content: `<w:p><w:pPr><w:madeUp><w:jc w:val="center"/></w:madeUp></w:pPr>${TARGET}</w:p>` },
  { name: '76t6-madeup-holding-b-in-rPr', asks: 'an unknown w: element in a w:rPr holding a w:b: is the next paragraph bold?', ignorable: true,
    content: '<w:p><w:r><w:rPr><w:madeUp><w:b/></w:madeUp></w:rPr><w:t>TARGET TEXT</w:t></w:r></w:p>' },
  { name: '76t7-madeup-in-madeup', asks: 'an unknown w: element in another, in a w:p, holding a run: do the words HIDDEN WORDS show?', ignorable: true,
    content: `<w:p>${run('TARGET TEXT, then ')}<w:madeUp><w:madeUp>${run('HIDDEN WORDS')}</w:madeUp></w:madeUp>${run(' and the rest.')}</w:p>` },
  // 76t's content with no mc:Ignorable on the root: every other w:madeUp file has it, and the foreign elements Word refused don't.
  { name: '76t8-madeup-with-no-ignorable', asks: 'an unknown w: element in a w:p holding a run, with no mc:Ignorable on the root: do the words HIDDEN WORDS show?', ignorable: false,
    content: `<w:p>${run('TARGET TEXT, then ')}<w:madeUp>${run('HIDDEN WORDS')}</w:madeUp>${run(' and the rest.')}</w:p>` },
  // Round 4 (2026-11-05; Jason: "CR-026 D7 we need to measure w:tr, w:sdtContent, w:hyperlink, w:ins, w:hdr and w:comment before making
  // any decision"; CR-026 section 4.4): the hosts round 3 left. w:sdtContent four times, since block, run, row and cell are four types in the
  // model. No mc:Ignorable, as 76t8. 76t18 holds every one with nothing unknown, so a refusal of the others is the element's and not these parts'.
  { name: '76t9-madeup-holding-tc-in-tr', asks: 'an unknown w: element in a w:tr holding a w:tc: is there a second cell reading HIDDEN WORDS?', ignorable: false,
    content: HOSTS.tr(made) },
  { name: '76t10-madeup-in-block-control', asks: 'an unknown w: element in a block control holding a w:p: do the words HIDDEN WORDS show, and inside the control?', ignorable: false,
    content: HOSTS.blockControl(made) },
  { name: '76t11-madeup-in-run-control', asks: 'an unknown w: element in a run-level control holding a run: do the words HIDDEN WORDS show, and inside the control?', ignorable: false,
    content: HOSTS.runControl(made) },
  { name: '76t12-madeup-in-row-control', asks: 'an unknown w: element in a row-level control holding a w:tr: is there a row reading HIDDEN WORDS, and inside the control?', ignorable: false,
    content: HOSTS.rowControl(made) },
  { name: '76t13-madeup-in-cell-control', asks: 'an unknown w: element in a cell-level control holding a w:tc: is there a cell reading HIDDEN WORDS, and inside the control?', ignorable: false,
    content: HOSTS.cellControl(made) },
  { name: '76t14-madeup-in-hyperlink', asks: 'an unknown w: element in a w:hyperlink holding a run: do the words HIDDEN WORDS show, and inside the link?', ignorable: false,
    content: HOSTS.hyperlink(made) },
  { name: '76t15-madeup-in-ins', asks: 'an unknown w: element in a w:ins holding a run: do the words HIDDEN WORDS show, and still tracked as inserted?', ignorable: false,
    content: HOSTS.ins(made) },
  { name: '76t16-madeup-in-hdr', asks: 'an unknown w: element in a w:hdr holding a w:p: does the header read HIDDEN WORDS?', ignorable: false,
    content: p('TARGET TEXT'), header: inBlocks(made) },
  { name: '76t17-madeup-in-comment', asks: 'an unknown w: element in a w:comment holding a w:p: does the comment read HIDDEN WORDS?', ignorable: false,
    content: COMMENTED, comment: inBlocks(made) },
  { name: '76t18-hosts-control', asks: 'nothing unknown: every host of 76t9 to 76t17, with HIDDEN WORDS standing bare in each', ignorable: false,
    content: Object.values(HOSTS).map((host) => host(bare)).join(p('NEXT HOST.')) + COMMENTED, header: inBlocks(bare), comment: inBlocks(bare) },
  // Round 5 (2026-11-05; Jason: "measure w:fldSimple, w:smartTag, w:customXml and w:footnote"; CR-026 section 4.4): the content hosts of section
  // 1's table that round 4 left. w:customXml four times, block, run, row and cell being four types in the model. 76t26 is their control.
  { name: '76t19-madeup-in-fldSimple', asks: 'an unknown w: element in a w:fldSimple holding a run: do the words HIDDEN WORDS show, and in the field result?', ignorable: false,
    content: MORE_HOSTS.fldSimple(made) },
  { name: '76t20-madeup-in-smartTag', asks: 'an unknown w: element in a w:smartTag holding a run: do the words HIDDEN WORDS show?', ignorable: false,
    content: MORE_HOSTS.smartTag(made) },
  { name: '76t21-madeup-in-block-customXml', asks: 'an unknown w: element in a block w:customXml holding a w:p: do the words HIDDEN WORDS show?', ignorable: false,
    content: MORE_HOSTS.blockCustomXml(made) },
  { name: '76t22-madeup-in-run-customXml', asks: 'an unknown w: element in a run-level w:customXml holding a run: do the words HIDDEN WORDS show?', ignorable: false,
    content: MORE_HOSTS.runCustomXml(made) },
  { name: '76t23-madeup-in-row-customXml', asks: 'an unknown w: element in a row-level w:customXml holding a w:tr: is there a row reading HIDDEN WORDS?', ignorable: false,
    content: MORE_HOSTS.rowCustomXml(made) },
  { name: '76t24-madeup-in-cell-customXml', asks: 'an unknown w: element in a cell-level w:customXml holding a w:tc: is there a cell reading HIDDEN WORDS?', ignorable: false,
    content: MORE_HOSTS.cellCustomXml(made) },
  { name: '76t25-madeup-in-footnote', asks: 'an unknown w: element in a w:footnote holding a w:p: does the footnote read HIDDEN WORDS?', ignorable: false,
    content: FOOTNOTED, footnote: inBlocks(made) },
  { name: '76t26-more-hosts-control', asks: 'nothing unknown: every host of 76t19 to 76t25, with HIDDEN WORDS standing bare in each', ignorable: false,
    content: Object.values(MORE_HOSTS).map((host) => host(bare)).join(p('NEXT HOST.')) + FOOTNOTED, footnote: inBlocks(bare) },
];

/** A file's header, comments and footnotes parts, each in the content types and the main document's relationships; nothing for a file with none. */
function ownParts(f, base) {
  const parts = [];
  if (f.header) parts.push(['header1.xml', 'header', `<w:hdr xmlns:w="${W}">${f.header}</w:hdr>`]);
  if (f.comment) parts.push(['comments.xml', 'comments', `<w:comments xmlns:w="${W}"><w:comment w:id="0" w:author="Check 76" w:date="2026-11-05T00:00:00Z" w:initials="C">${f.comment}</w:comment></w:comments>`]);
  if (f.footnote) parts.push(['footnotes.xml', 'footnotes', `<w:footnotes xmlns:w="${W}"><w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>`
    + `<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote><w:footnote w:id="1">${f.footnote}</w:footnote></w:footnotes>`]);
  if (parts.length === 0) return {};
  const added = (bytes, end, line) => strToU8(strFromU8(bytes).replace(end, parts.map(line).join('') + end));
  return {
    '[Content_Types].xml': added(base['[Content_Types].xml'], '</Types>',
      ([name, kind]) => `<Override PartName="/word/${name}" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${kind}+xml"/>`),
    'word/_rels/document.xml.rels': added(base['word/_rels/document.xml.rels'], '</Relationships>',
      ([name, kind]) => `<Relationship Id="rId76${kind}" Target="${name}" Type="${R}/${kind}"/>`),
    ...Object.fromEntries(parts.map(([name, , xml]) => [`word/${name}`, strToU8(declaration + xml)])),
  };
}

/** The files, { name, asks, ignorable, document, bytes }, over a created package's parts. */
export async function check76Files() {
  const base = unzipSync(await (await WordprocessingMLPackage.createPackage()).save());
  return FILES.map((f) => {
    const document = `${declaration}<w:document xmlns:w="${W}" xmlns:mc="${MC}" xmlns:ext="${EXT}"${f.ignorable ? ' mc:Ignorable="ext"' : ''}><w:body>`
      + p(`Check 76, ${f.name}: ${f.asks}.`) + f.content + p('LAST LINE.')
      + `<w:sectPr>${f.header ? `<w:headerReference xmlns:r="${R}" w:type="default" r:id="rId76header"/>` : ''}`
      + '<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>'
      + '</w:body></w:document>';
    // a fixed time makes the file the same on every build
    const bytes = zipSync({ ...base, 'word/document.xml': strToU8(document), ...ownParts(f, base) }, { mtime: new Date(Date.UTC(2026, 10, 5)) });
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
