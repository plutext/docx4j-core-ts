// CR-012 section 24 (docx4j CR-069, released in 17.3.5), and section 23.1's profile: w:caps in capitals by the run's
// language (REQ-607), the block model's case, alignment and language members, a drop cap joined to its word (REQ-521),
// strike taking w:dstrike, and the options small_caps, drop_caps and profile. docx4j's CaseMappingTest, DropCapJoinTest
// and TextOptionsProfileTest, mirrored with docx4j's strings; Markdown's <sup> and <sub> are in markdown.test.mjs.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { execFileSync } from 'node:child_process';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import {
  WordprocessingMLPackage, Texts, PlainText, TextOptions, TextOptionError, TEXT_OPTION_NAMES, CaseMapping,
} from '../dist/index.mjs';
import { verifyWordSave } from './text-extraction/conformance.mjs';

const here = dirname(fileURLToPath(import.meta.url));
const suite = join(here, 'text-extraction', 'suite');
const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const load = (path) => WordprocessingMLPackage.load(readFileSync(join(suite, path)));

/** docx4j's test `document(body)`: createPackage, its body this XML. */
async function document(body, ...styles) {
  const pkg = await WordprocessingMLPackage.createPackage();
  const mdp = pkg.getMainDocumentPart();
  mdp.setXml(`<w:document ${W}><w:body>${body}</w:body></w:document>`);
  if (styles.length > 0) {
    const part = mdp.styleDefinitionsPart;
    part.setXml((await part.getXml()).replace('</w:styles>', `${styles.map((s) => s.replace('<w:style ', `<w:style ${W} `)).join('')}</w:styles>`));
    await pkg.refresh();
  }
  return pkg;
}

const runOf = (r, text) => {
  for (const b of r.blocks) for (const t of b.runs) if (t.text === text) return t;
  throw new Error(`No run '${text}'`);
};
const addresses = (r) => r.blocks.map((b) => b.address);
const texts = (runs) => runs.map((t) => t.text);

// ---------------------------------------------------------------- the table (section 24.1)

test('the case table is Unicode 15.0\'s simple uppercase mapping, generated, not the runtime\'s', () => {
  assert.equal(CaseMapping.UNICODE_VERSION, '15.0.0');
  // the generator's own header: the file it read, by its SHA-256, and the count Java 21's Character.toUpperCase gives
  const generated = readFileSync(join(here, '..', 'src', 'model', 'text', 'caseMapping.generated.mts'), 'utf8');
  assert.match(generated, /Public\/15\.0\.0\/ucd\/UnicodeData\.txt\n\/\/ \(SHA-256 806e9aed65037197f1ec85e12be6e8cd870fc5608b4de0fffd990f689f376a73\): 1450 code points in 200 runs\./);
  // a pair Unicode 16.0 added stays unmapped (WD 0.52), where the runtime's ICU may map it
  assert.equal(CaseMapping.upper('ᲊ'), 'ᲊ');
  assert.equal(CaseMapping.upper('ɤ'), 'ɤ');
  // the generator is reproducible from its file: run it with --check where the file is at hand (not in CI)
  assert.throws(() => execFileSync(process.execPath, [join(here, '..', 'scripts', 'generate-case-mapping.mjs')], { stdio: 'pipe' }), /usage/);
});

// ---------------------------------------------------------------- docx4j CaseMappingTest

test('Word\'s capitals: the run\'s language, one code point to one', () => {
  assert.equal(CaseMapping.upper('istanbul film', 'tr-TR'), 'İSTANBUL FİLM'); // caps-lang7c, 6a
  assert.equal(CaseMapping.upper('istanbul film', 'en-US'), 'ISTANBUL FILM'); // caps-lang6a
  assert.equal(CaseMapping.upper('istanbul film', undefined), 'ISTANBUL FILM'); // docx4j's default, en-US
  assert.equal(CaseMapping.upper('straße', 'de-DE'), 'STRAßE'); // caps-lang6b, 7b: no SS, no U+1E9E
  assert.equal(CaseMapping.upper('café', 'en-US'), 'CAFÉ'); // caps-lang6b
  assert.equal(CaseMapping.upper('ﬁne café ǆ', 'en-US'), 'ﬁNE CAFÉ Ǆ'); // caps-lang7d
});

test('the dotted I is Turkish\'s and Azeri\'s, by the primary subtag in any case; every other language maps i to I', () => {
  for (const lang of ['tr', 'TR-tr', 'tr_TR', 'az', 'az-Latn-AZ']) assert.equal(CaseMapping.upper('i', lang), 'İ', lang);
  for (const lang of ['en-US', 'de-DE', 'lt-LT', '', 'trk', 'x-tr']) assert.equal(CaseMapping.upper('i', lang), 'I', lang);
  // dotless i is lowercase-only: I in every language
  assert.equal(CaseMapping.upper('ı', 'tr-TR'), 'I');
  assert.equal(CaseMapping.upper('ı', 'en-US'), 'I');
});

test('Unicode\'s simple mapping, never the full mapping toUpperCase() is: the differences pinned', () => {
  assert.equal('straße'.toUpperCase(), 'STRASSE');
  assert.equal('ﬁne'.toUpperCase(), 'FINE');
  assert.equal(CaseMapping.upper('ﬁne', 'en-US'), 'ﬁNE');
  // U+1FB3: simple uppercase U+1FBC, the full mapping two letters
  assert.equal('ᾳ'.toUpperCase(), 'ΑΙ');
  assert.equal(CaseMapping.upper('ᾳ', 'el-GR'), 'ᾼ');
  // the titlecase digraph to its uppercase
  assert.equal(CaseMapping.upper('ǅ', 'en-US'), 'Ǆ');
  // a supplementary letter (Deseret)
  assert.equal(CaseMapping.upper(String.fromCodePoint(0x10428), 'en-US'), String.fromCodePoint(0x10400));
});

test('the mapping keeps the length in code points and in UTF-16 units, for every code point, so offsets hold', () => {
  let all = '';
  for (let cp = 0; cp <= 0x10FFFF; cp++) if (cp < 0xD800 || cp > 0xDFFF) all += String.fromCodePoint(cp);
  for (const lang of ['en-US', 'tr-TR']) {
    const upper = CaseMapping.upper(all, lang);
    assert.equal(upper.length, all.length);
    assert.equal([...upper].length, [...all].length);
  }
  assert.equal(CaseMapping.upper('', 'tr-TR'), '');
});

test('the block model keeps the stored text and carries the flags (Word\'s probe, dpp-forks-word365.docx)', async () => {
  const pkg = await load('dpp/dpp-forks-word365.docx');
  const r = await Texts.of(pkg);
  const jackson = runOf(r, 'Jackson');
  const roberts = runOf(r, 'Roberts');
  assert.ok(jackson.smallCaps);
  assert.ok(!jackson.caps);
  assert.equal(jackson.language, 'en-US'); // no w:lang: the document defaults'
  assert.ok(roberts.caps);
  assert.ok(!roberts.smallCaps);
  const plain = runOf(r, ', J., dissenting; ');
  assert.ok(!plain.caps && !plain.smallCaps);
  assert.equal(plain.language, undefined);
  assert.equal(plain.verticalAlignment, 'BASELINE');
  const raised = [];
  for (const b of r.blocks) for (const t of b.runs) if (t.verticalAlignment !== 'BASELINE') raised.push(`${t.text} ${t.verticalAlignment}`);
  assert.deepEqual(raised, ['2 SUBSCRIPT', '2 SUPERSCRIPT', 'st SUPERSCRIPT', '3 SUPERSCRIPT']);
  // the writers' capitals, not the block model's
  const text = PlainText.of(r);
  assert.ok(text.includes('C05_ Jackson, J., dissenting; ROBERTS, C.J.'), text);
  assert.ok(PlainText.of(await Texts.of(pkg, { small_caps: 'upper' })).includes('C05_ JACKSON, J., dissenting; ROBERTS, C.J.'));
  // the block model's JSON carries the four members (REQ-706)
  assert.deepEqual(JSON.parse(JSON.stringify(roberts)), { ...JSON.parse(JSON.stringify(roberts)), caps: true, smallCaps: false, verticalAlignment: 'BASELINE', language: 'en-US' });
});

test('the run\'s language: its own w:lang, else the document defaults\' (caps-lang6a and 6b)', async () => {
  const a = await Texts.of(await load('dpp/caps/caps-lang6a.docx'));
  const langs = [];
  for (const b of a.blocks) for (const t of b.runs) if (t.language !== undefined) langs.push(`${t.language} ${t.caps ? 'caps' : 'smallCaps'}`);
  assert.deepEqual(langs, ['tr-TR caps', 'en-US caps', 'tr-TR smallCaps']);
  const text = PlainText.of(a);
  assert.ok(text.includes('L02_ tr-TR caps: İSTANBUL FİLM (dotted'), text);
  assert.ok(text.includes('L03_ en-US caps: ISTANBUL FILM\n'), text);
  assert.ok(text.includes('L04_ tr-TR smallCaps: istanbul film\n'), text);
  const b = await Texts.of(await load('dpp/caps/caps-lang6b.docx'));
  assert.ok(PlainText.of(b).includes('L07_ no w:lang, caps: ISTANBUL FILM\n'));
});

test('the case is the text\'s whatever the paragraph style; caps over small caps; runs merge where case and language agree', async () => {
  const pkg = await document('<w:p><w:pPr><w:pStyle w:val="Shout"/></w:pPr><w:r><w:t>loud</w:t></w:r></w:p>'
    + '<w:p><w:r><w:rPr><w:caps/><w:smallCaps/></w:rPr><w:t>both</w:t></w:r>'
    + '<w:r><w:rPr><w:caps/><w:lang w:val="en-US"/></w:rPr><w:t xml:space="preserve"> i</w:t></w:r>'
    + '<w:r><w:rPr><w:caps/><w:lang w:val="tr-TR"/></w:rPr><w:t xml:space="preserve"> i</w:t></w:r>'
    + '<w:r><w:rPr><w:caps w:val="0"/></w:rPr><w:t xml:space="preserve"> off</w:t></w:r></w:p>',
  '<w:style w:type="paragraph" w:styleId="Shout"><w:name w:val="Shout"/><w:rPr><w:b/><w:caps/></w:rPr></w:style>');
  const r = await Texts.of(pkg);
  const loud = r.blocks[0].runs[0];
  assert.ok(loud.caps);
  assert.ok(!loud.bold); // the style's bold is the heading's, not emphasis
  const second = r.blocks[1].runs;
  // "both" takes the document defaults' en-US, so it merges with the en-US caps run, not the tr-TR one
  assert.deepEqual(second.map((t) => (t.text.startsWith(' ') ? `'${t.text}'${t.language === undefined ? '' : ` ${t.language}`}` : t.text)),
    ['both i', '\' i\' tr-TR', '\' off\'']);
  assert.ok(second[0].caps);
  assert.ok(!second[0].smallCaps);
  assert.ok(!second[2].caps);
  assert.equal(PlainText.of(r), 'LOUD\nBOTH I İ off\n');
});

test('the language of last resort: w:themeFontLang, else en-US', async () => {
  const pkg = await document('<w:p><w:r><w:rPr><w:caps/></w:rPr><w:t>istanbul</w:t></w:r></w:p>');
  const mdp = pkg.getMainDocumentPart();
  const styles = mdp.styleDefinitionsPart;
  styles.setXml((await styles.getXml()).replace(/<w:lang [^>]*\/>/, ''));
  await pkg.refresh();
  const settings = await mdp.documentSettingsPart.getContents();
  delete settings.themeFontLang;
  assert.equal(PlainText.of(await Texts.of(pkg)), 'ISTANBUL\n');
  assert.equal((await Texts.of(pkg)).blocks[0].runs[0].language, CaseMapping.DEFAULT_LANGUAGE);
  settings.themeFontLang = { TYPE_NAME: 'org_docx4j_wml.CTLanguage', val: 'tr-TR' };
  // a tr-TR default, unmeasured in Word (docx4j CR-069 section 7): predicted as the run-level case
  assert.equal(PlainText.of(await Texts.of(pkg)), 'İSTANBUL\n');
});

test('a note reference is a reference, not a superscript (REQ-710), whatever its style raises', async () => {
  const r = await Texts.of(await load('phase0/extraction-probe.docx'));
  let references = 0;
  for (const b of r.blocks) for (const t of b.runs) {
    if (t.kind !== 'NOTE_REFERENCE') continue;
    references++;
    assert.equal(t.verticalAlignment, 'BASELINE');
  }
  assert.ok(references > 0);
});

test('small_caps by name, its default and the word-text preset\'s value', () => {
  assert.equal(new TextOptions().smallCaps, 'asis');
  assert.equal(new TextOptions().toMap().small_caps, 'asis');
  assert.equal(TextOptions.of({ small_caps: 'upper' }).smallCaps, 'upper');
  assert.equal(new TextOptions().profile('word-text').smallCaps, 'upper');
  assert.equal(new TextOptions().set('small_caps', 'asis').profile('word-text').smallCaps, 'asis');
  assert.equal(TextOptions.of({ small_caps: 'upper' }).copy().smallCaps, 'upper');
  assert.equal(TEXT_OPTION_NAMES.indexOf('symbols') + 1, TEXT_OPTION_NAMES.indexOf('small_caps')); // WD section 10's order
});

// ---------------------------------------------------------------- strike takes w:dstrike (section 24.5)

test('strike takes w:dstrike, the run\'s own or its styles\', and the paragraph style\'s either cancels it', async () => {
  const pkg = await document('<w:p><w:r><w:rPr><w:dstrike/></w:rPr><w:t>twice</w:t></w:r><w:r><w:t xml:space="preserve"> plain</w:t></w:r></w:p>'
    + '<w:p><w:pPr><w:pStyle w:val="Struck"/></w:pPr><w:r><w:rPr><w:dstrike/></w:rPr><w:t>styled</w:t></w:r></w:p>',
  '<w:style w:type="paragraph" w:styleId="Struck"><w:name w:val="Struck"/><w:rPr><w:strike/></w:rPr></w:style>');
  const r = await Texts.of(pkg);
  assert.deepEqual(r.blocks[0].runs.map((t) => `${t.text}${t.strike ? ' struck' : ''}`), ['twice struck', ' plain']);
  assert.equal(r.blocks[1].runs[0].strike, false);
});

// ---------------------------------------------------------------- docx4j DropCapJoinTest

const FRAME = '<w:pPr><w:framePr w:dropCap="drop" w:lines="3" w:wrap="around" w:vAnchor="text" w:hAnchor="text"/></w:pPr>';

test('Word\'s drop cap is joined to its word: one block, the second paragraph\'s address; the warning at the frame\'s', async () => {
  const pkg = await load('dpp/bisect/C02.docx');
  const r = await Texts.of(pkg);
  assert.deepEqual(addresses(r), ['body/0', 'body/2', 'body/3']);
  const joined = r.blocks[1];
  assert.ok(joined.text.startsWith('The C02_ modern Recreational Boating'), joined.text);
  assert.deepEqual(texts(joined.runs), ['The C02_ modern Recreational Boating Safety Program traces its inception to the Federal Boat Safety Act of 1971.']);
  // D6: the letter's source, no officeOffset (the frame is its own paragraph in Office JS)
  assert.equal(joined.runs[0].officeOffset, undefined);
  assert.equal(r.warnings.length, 1);
  const w = r.warnings[0];
  assert.equal(w.kind, 'paragraph.drop_cap_joined');
  assert.equal(w.address, 'body/1');
  assert.equal(w.selector, '/word/document.xml');
  assert.ok(PlainText.of(r).includes('\nThe C02_ modern Recreational'));
  for (const view of ['accepted', 'original', 'markup']) assert.deepEqual(addresses(await Texts.of(pkg, { view })), ['body/0', 'body/2', 'body/3'], view);
});

test('drop_caps: asis, and the word-text preset, which sets it: the two paragraphs as the package holds them, no warning', async () => {
  const pkg = await load('dpp/bisect/C02.docx');
  for (const o of [{ drop_caps: 'asis' }, TextOptions.of({ drop_caps: 'ASIS' }), new TextOptions().profile('word-text')]) {
    const r = await Texts.of(pkg, o);
    assert.deepEqual(addresses(r), ['body/0', 'body/1', 'body/2', 'body/3']);
    assert.equal(r.blocks[1].text, 'T');
    assert.deepEqual(r.warnings, []);
    assert.ok(PlainText.of(r).includes('\nT\nhe C02_ modern'));
  }
  // a caller's drop_caps overrides the preset's, before or after it
  const join = new TextOptions().set('drop_caps', 'join').profile('word-text');
  assert.equal(join.dropCaps, 'join');
  assert.deepEqual(addresses(await Texts.of(pkg, join)), ['body/0', 'body/2', 'body/3']);
});

test('the shapes REQ-521 leaves as they are: more than a word, no paragraph next, a frame that is no drop cap, a section\'s end', async () => {
  const next = '<w:p><w:r><w:t>next</w:t></w:r></w:p>';
  const bodies = [
    // two words
    `<w:p>${FRAME}<w:r><w:t>Once upon</w:t></w:r></w:p>${next}`,
    // a table next
    `<w:p>${FRAME}<w:r><w:t>T</w:t></w:r></w:p><w:tbl><w:tr><w:tc><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>${next}`,
    // the end of the story
    `${next}<w:p>${FRAME}<w:r><w:t>T</w:t></w:r></w:p>`,
    // a frame with no drop cap
    `<w:p><w:pPr><w:framePr w:dropCap="none" w:wrap="around"/></w:pPr><w:r><w:t>T</w:t></w:r></w:p>${next}`,
    `<w:p><w:pPr><w:framePr w:wrap="around"/></w:pPr><w:r><w:t>T</w:t></w:r></w:p>${next}`,
    // the frame ends a section: the paragraph next is in another
    `<w:p><w:pPr><w:framePr w:dropCap="drop"/><w:sectPr/></w:pPr><w:r><w:t>T</w:t></w:r></w:p>${next}`,
    // a block-level control next
    `<w:p>${FRAME}<w:r><w:t>T</w:t></w:r></w:p><w:sdt><w:sdtContent>${next}</w:sdtContent></w:sdt>`,
    // a no-break space is a space (Java's isSpaceChar)
    `<w:p>${FRAME}<w:r><w:t>T </w:t></w:r></w:p>${next}`,
  ];
  for (const body of bodies) {
    const r = await Texts.of(await document(body));
    assert.equal(r.blocks.length, 2 + (body.includes('<w:tbl>') ? 1 : 0), body);
    assert.deepEqual(r.warnings, [], body);
  }
  // the end of a cell
  const cell = await Texts.of(await document(`<w:tbl><w:tr><w:tc><w:p>${FRAME}<w:r><w:t>T</w:t></w:r></w:p></w:tc></w:tr></w:tbl>${next}`));
  assert.deepEqual(cell.warnings, []);
});

test('dropCap margin, a style\'s frame, a bookmark between, a cell\'s paragraphs, an empty frame: each joined', async () => {
  const word = '<w:p><w:r><w:t xml:space="preserve">he end</w:t></w:r></w:p>';
  const bodies = [
    `<w:p><w:pPr><w:framePr w:dropCap="margin" w:lines="3"/></w:pPr><w:r><w:t>T</w:t></w:r></w:p>${word}`,
    `<w:p><w:pPr><w:pStyle w:val="DropCap"/></w:pPr><w:r><w:t>T</w:t></w:r></w:p>${word}`,
    `<w:p>${FRAME}<w:r><w:t>T</w:t></w:r></w:p><w:bookmarkStart w:id="0" w:name="b"/><w:bookmarkEnd w:id="0"/>${word}`,
    `<w:p>${FRAME}<w:r><w:t>T</w:t></w:r><w:bookmarkStart w:id="1" w:name="c"/></w:p>${word}`,
  ];
  for (const body of bodies) {
    const pkg = body.includes('DropCap')
      ? await document(body, '<w:style w:type="paragraph" w:styleId="DropCap"><w:name w:val="Drop Cap"/><w:pPr><w:framePr w:dropCap="drop" w:lines="3"/></w:pPr></w:style>')
      : await document(body);
    const r = await Texts.of(pkg);
    assert.equal(r.blocks.length, 1, body);
    assert.equal(r.blocks[0].text, 'The end', body);
    assert.equal(r.blocks[0].address, `body/${body.includes('w:bookmarkStart w:id="0"') ? 3 : 1}`, body);
    assert.equal(r.warnings.length, 1, body);
    assert.equal(r.warnings[0].address, 'body/0', body);
  }
  // in a cell, the paragraph next in the cell takes it; the cell's address
  const cell = await Texts.of(await document(`<w:tbl><w:tr><w:tc><w:p>${FRAME}<w:r><w:t>T</w:t></w:r></w:p>${word}</w:tc></w:tr></w:tbl>`));
  assert.equal(cell.blocks[0].text, 'The end');
  assert.equal(cell.blocks[0].rows[0].cells[0].blocks[0].address, 'body/0/0/0/1');
  assert.equal(cell.warnings[0].address, 'body/0/0/0/0');
  // a frame with no text (a picture's, say): at most one word, so joined
  const empty = await Texts.of(await document(`<w:p>${FRAME}</w:p>${word}`));
  assert.deepEqual(addresses(empty), ['body/1']);
  assert.equal(empty.warnings.length, 1);
});

test('a frame whose mark the view deletes is one paragraph with the next already: no drop-cap join and no warning', async () => {
  const body = '<w:p><w:pPr><w:framePr w:dropCap="drop" w:lines="3"/><w:rPr><w:del w:id="1" w:author="A"'
    + ' w:date="2026-10-31T00:00:00Z"/></w:rPr></w:pPr><w:r><w:t>T</w:t></w:r></w:p>'
    + '<w:p><w:r><w:t xml:space="preserve">he end</w:t></w:r></w:p>';
  const accepted = await Texts.of(await document(body), { view: 'accepted' });
  assert.deepEqual(addresses(accepted), ['body/0']);
  assert.equal(accepted.blocks[0].text, 'The end');
  assert.deepEqual(accepted.warnings, []);
  const original = await Texts.of(await document(body), { view: 'original' });
  assert.deepEqual(addresses(original), ['body/1']);
  assert.equal(original.blocks[0].text, 'The end');
  assert.equal(original.warnings.length, 1);
});

test('the letter merges into its word where the block model\'s formatting agrees; a bold or a cased letter stays its own run (D6)', async () => {
  const word = '<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t xml:space="preserve">he end</w:t></w:r></w:p>';
  const plain = await Texts.of(await document(`<w:p>${FRAME}<w:r><w:rPr><w:sz w:val="80"/></w:rPr><w:t>T</w:t></w:r></w:p>${word}`));
  assert.deepEqual(texts(plain.blocks[0].runs), ['The end']);
  assert.equal(plain.blocks[0].alignment, 'center');
  assert.equal(plain.blocks[0].runs[0].officeOffset, undefined);
  const bold = await Texts.of(await document(`<w:p>${FRAME}<w:r><w:rPr><w:b/></w:rPr><w:t>T</w:t></w:r></w:p>${word}`));
  const runs = bold.blocks[0].runs;
  assert.deepEqual(texts(runs), ['T', 'he end']);
  assert.ok(runs[0].bold);
  assert.ok(!runs[1].bold);
  // each keeps its offset in its own paragraph's Paragraph.text
  assert.deepEqual(runs.map((t) => t.officeOffset), [0, 0]);
  const caps = await Texts.of(await document(`<w:p>${FRAME}<w:r><w:rPr><w:caps/></w:rPr><w:t>t</w:t></w:r></w:p>${word}`));
  assert.equal(PlainText.of(caps), 'The end\n');
});

test('drop_caps by name, its default, its place in WD section 10\'s order and the word-text preset\'s value', () => {
  assert.equal(new TextOptions().dropCaps, 'join');
  assert.equal(new TextOptions().toMap().drop_caps, 'join');
  assert.equal(TextOptions.of({ drop_caps: 'ASIS' }).dropCaps, 'asis');
  assert.equal(new TextOptions().profile('word-text').dropCaps, 'asis');
  assert.equal(TextOptions.of({ drop_caps: 'asis' }).copy().dropCaps, 'asis');
  assert.equal(TEXT_OPTION_NAMES.indexOf('toc') + 1, TEXT_OPTION_NAMES.indexOf('drop_caps'));
  assert.throws(() => TextOptions.of({ drop_caps: 'split' }), (e) => e instanceof TextOptionError && e.message.includes('drop_caps'));
});

// ---------------------------------------------------------------- docx4j TextOptionsProfileTest (section 23.1)

test('profile: word-text is Word\'s plain text as the options reach it', () => {
  const o = new TextOptions().profile('word-text');
  assert.equal(o.view, 'accepted');
  assert.ok(o.hidden);
  assert.ok(o.includes.has('headers') && o.includes.has('footers'));
  assert.ok(o.comments);
  assert.ok(!o.textBoxes);
  assert.ok(o.notes && o.noteReferences);
  assert.ok(!o.listLabels);
  assert.equal(o.fields, 'cached');
  assert.ok(!o.fieldCodes);
  assert.ok(o.toc);
  assert.ok(o.placeholders);
  assert.equal(o.bindings, 'bound');
  assert.ok(o.altchunks);
  assert.equal(o.symbols, 'asis');
  assert.equal(o.tables, 'tsv');
  assert.ok(!o.pageMarkers);
  assert.equal(o.layout, 'none');
  assert.ok(!(o.altText || o.metadata || o.charts || o.smartart || o.embedded));
  // Word's linear math is not produced: math: unicodemath is refused here as in docx4j, so the preset leaves LaTeX
  assert.equal(o.math, 'latex');
  assert.equal(o.toMap().profile, 'word-text');
  assert.equal(Object.keys(o.toMap()).at(-1), 'profile');
  assert.equal(new TextOptions().toMap().profile, null);
  assert.equal(o.profileApplied, 'word-text');
});

test('the caller\'s option overrides the profile, whichever came first', () => {
  const after = new TextOptions().profile('word-text').set('fields', 'update').set('tables', 'grid');
  assert.equal(after.fields, 'update');
  assert.equal(after.tables, 'grid');
  const before = new TextOptions().set('fields', 'update').set('tables', 'grid').profile('word-text');
  assert.equal(before.fields, 'update');
  assert.equal(before.tables, 'grid');
  assert.ok(before.hidden); // the rest of the preset applied
  // by WD name, in either order, as a tool passes them
  const fromMap = TextOptions.of({ fields: 'update', profile: 'word-text' });
  assert.equal(fromMap.fields, 'update');
  assert.ok(fromMap.hidden && fromMap.toc);
  const m2 = TextOptions.of({ profile: 'WORD_TEXT', toc: false });
  assert.ok(!m2.toc);
  assert.ok(m2.hidden);
  // include() is the caller's too
  assert.deepEqual([...new TextOptions().include('glossary').profile('word-text').includes], ['glossary', 'comments']);
});

test('an unknown profile is refused, naming REQ-1002', () => {
  assert.throws(() => TextOptions.of({ profile: 'libreoffice' }), (e) => e instanceof TextOptionError && e.statement === 'REQ-1002');
});

test('under the preset, a document reads as Word\'s text save of it (the dpp suite\'s control save)', async () => {
  // docx4j's own test reads its field memo, which this repository does not hold; the suite's 22 saves are the measure
  const r = await Texts.of(await load('dpp/bisect/control.docx'), { profile: 'word-text' });
  verifyWordSave(join(suite, 'dpp', 'bisect', 'control.txt'), 'windows-1252', PlainText.of(r));
  // drop_caps: join over the preset joins the drop cap, where Word's save writes T, then he C02_ ...
  const c02 = join(suite, 'dpp', 'bisect', 'C02.txt');
  const pkg = await load('dpp/bisect/C02.docx');
  verifyWordSave(c02, 'windows-1252', PlainText.of(await Texts.of(pkg, { profile: 'word-text' })));
  const joined = PlainText.of(await Texts.of(pkg, { profile: 'word-text', drop_caps: 'join' }));
  assert.throws(() => verifyWordSave(c02, 'windows-1252', joined), /: line \d+: Word wrote \[T\]/);
});
