// CR-025: what a read drops, reported as `part.skipped` on the part and on pkg.warnings (phase 1); a part read and not changed
// copied, and a save that would lose such content refused (phase 2).
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readdirSync, readFileSync, statSync } from 'node:fs';
import { join } from 'node:path';
import { unzipSync, zipSync, strToU8 } from 'fflate';
import { MemoryPartSink, OpcPackage, PlainText, SaveLossException, Texts, WordprocessingMLPackage } from '../dist/index.mjs';
import { bytesEqual, fixturesDir } from './helpers.mjs';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
const EXT = 'urn:docx4j:cr025';
const roots = `xmlns:w="${W}" xmlns:mc="${MC}" xmlns:ext="${EXT}" mc:Ignorable="ext"`;
const X = '@@';
const ELEMENTS = { 'ext:opaque': '<ext:opaque ext:a="1"><ext:inner/>text</ext:opaque>', 'w:madeUp': '<w:madeUp w:val="1"/>' };

/** A created package with its named parts replaced, read through the real load path. */
async function loaded(parts) {
  const base = unzipSync(await (await WordprocessingMLPackage.createPackage()).save());
  const replaced = Object.fromEntries(Object.entries(parts).map(([name, xml]) => [name, strToU8(xml)]));
  return WordprocessingMLPackage.load(zipSync({ ...base, ...replaced }));
}
const inBody = (blocks) => `<w:document ${roots}><w:body>${blocks}</w:body></w:document>`;
const documentOf = (blocks) => loaded({ 'word/document.xml': inBody(blocks) });
const skippedIn = (part) => part.warnings.filter((w) => w.code === 'part.skipped');
const partOf = (pkg, name) => [...pkg.parts].find((p) => p.partName.name === name);

/** CR-025 section 1's hosts, each dropping what it doesn't know: its part, its class and its path. The twelfth, `w:lvl`, needs a part of its own, below. */
const HOSTS = [
  ['w:document', '/word/document.xml', `<w:document ${roots}>${X}<w:body><w:p/></w:body></w:document>`, 'org_docx4j_wml.Document', '/w:document'],
  ['w:pPr', '/word/document.xml', inBody(`<w:p><w:pPr>${X}</w:pPr></w:p>`), 'org_docx4j_wml.PPr', '/w:document/w:body/w:p/w:pPr'],
  ['w:rPr', '/word/document.xml', inBody(`<w:p><w:r><w:rPr>${X}</w:rPr><w:t>a</w:t></w:r></w:p>`), 'org_docx4j_wml.RPr', '/w:document/w:body/w:p/w:r/w:rPr'],
  ['w:tblPr', '/word/document.xml', inBody(`<w:tbl><w:tblPr>${X}</w:tblPr><w:tblGrid/><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl><w:p/>`), 'org_docx4j_wml.TblPr', '/w:document/w:body/w:tbl/w:tblPr'],
  ['w:tblGrid', '/word/document.xml', inBody(`<w:tbl><w:tblPr/><w:tblGrid>${X}</w:tblGrid><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl><w:p/>`), 'org_docx4j_wml.TblGrid', '/w:document/w:body/w:tbl/w:tblGrid'],
  ['w:tcPr', '/word/document.xml', inBody(`<w:tbl><w:tblPr/><w:tblGrid/><w:tr><w:tc><w:tcPr>${X}</w:tcPr><w:p/></w:tc></w:tr></w:tbl><w:p/>`), 'org_docx4j_wml.TcPr', '/w:document/w:body/w:tbl/w:tr/w:tc/w:tcPr'],
  ['w:sdt', '/word/document.xml', inBody(`<w:sdt>${X}<w:sdtPr/><w:sdtContent><w:p/></w:sdtContent></w:sdt>`), null, '/w:document/w:body/w:sdt'],
  ['w:sdtEndPr', '/word/document.xml', inBody(`<w:sdt><w:sdtPr/><w:sdtEndPr>${X}</w:sdtEndPr><w:sdtContent><w:p/></w:sdtContent></w:sdt>`), null, '/w:document/w:body/w:sdt/w:sdtEndPr'],
  ['w:sectPr', '/word/document.xml', inBody(`<w:p/><w:sectPr>${X}</w:sectPr>`), 'org_docx4j_wml.SectPr', '/w:document/w:body/w:sectPr'],
  ['w:style', '/word/styles.xml', `<w:styles ${roots}><w:style w:styleId="a">${X}</w:style></w:styles>`, null, '/w:styles/w:style'],
  ['w:settings', '/word/settings.xml', `<w:settings ${roots}>${X}</w:settings>`, null, '/w:settings'],
];

for (const [host, partName, xml, parentClass, path] of HOSTS) {
  for (const [name, element] of Object.entries(ELEMENTS)) {
    test(`CR-025 section 2.1: ${name} dropped from ${host} is reported, its fields held to the source`, async () => {
      const pkg = await loaded({ [partName.slice(1)]: xml.replace(X, element) });
      const part = partOf(pkg, partName);
      await part.getContents();
      const [w, ...rest] = skippedIn(part);
      assert.deepEqual(rest, [], 'the element once, not what it holds');
      assert.equal(w.partName, partName);
      assert.equal(w.kind, 'element');
      assert.equal(w.name, name);
      assert.equal(w.parentName, host);
      if (parentClass) assert.equal(w.parentClass, parentClass);
      else assert.match(w.parentClass, /^org_docx4j_wml\./);
      assert.equal(w.path, `${path}/${name}`);
      assert.equal(w.ignorable, name === 'ext:opaque', 'listed in mc:Ignorable, or not');
      assert.equal(typeof w.line, 'number');
      assert.equal(typeof w.column, 'number');
      assert.ok(pkg.warnings.includes(w), 'gathered by pkg.warnings');
    });
  }
}

test('CR-025 section 2.1: an element dropped from w:lvl is reported, in a numbering part', async () => {
  const base = unzipSync(await (await WordprocessingMLPackage.createPackage()).save());
  const text = (name) => new TextDecoder().decode(base[name]);
  const NUMBERING = 'application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml';
  const pkg = await WordprocessingMLPackage.load(zipSync({
    ...base,
    '[Content_Types].xml': strToU8(text('[Content_Types].xml').replace('</Types>', `<Override PartName="/word/numbering.xml" ContentType="${NUMBERING}"/></Types>`)),
    'word/_rels/document.xml.rels': strToU8(text('word/_rels/document.xml.rels').replace('</Relationships>', '<Relationship Id="rId25" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/></Relationships>')),
    'word/numbering.xml': strToU8(`<w:numbering ${roots}><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><ext:opaque/></w:lvl></w:abstractNum></w:numbering>`),
  }));
  const part = partOf(pkg, '/word/numbering.xml');
  await part.getContents();
  assert.deepEqual(skippedIn(part).map((w) => [w.name, w.parentName, w.path, w.ignorable]),
    [['ext:opaque', 'w:lvl', '/w:numbering/w:abstractNum/w:lvl/ext:opaque', true]]);
});

test('CR-025 section 1: the six attributes, on a paragraph and a run, each named as the part writes it', async () => {
  const OTHER = 'urn:docx4j:cr025:other';
  const pkg = await loaded({ 'word/document.xml': `<w:document ${roots} xmlns:o="${OTHER}"><w:body><w:p ext:a="1" o:b="2" w:madeUp="3"><w:r ext:a="1" o:b="2" w:madeUp="3"><w:t>a</w:t></w:r></w:p></w:body></w:document>` });
  const main = pkg.getMainDocumentPart();
  await main.getContents();
  const found = skippedIn(main).map((w) => [w.kind, w.name, w.parentName, w.parentClass, w.path, w.ignorable]);
  const onP = '/w:document/w:body/w:p';
  assert.deepEqual(found.sort(), [
    ['attribute', 'ext:a', 'w:p', 'org_docx4j_wml.P', `${onP}/@ext:a`, true],
    ['attribute', 'o:b', 'w:p', 'org_docx4j_wml.P', `${onP}/@o:b`, false],
    ['attribute', 'w:madeUp', 'w:p', 'org_docx4j_wml.P', `${onP}/@w:madeUp`, false],
    ['attribute', 'ext:a', 'w:r', 'org_docx4j_wml.R', `${onP}/w:r/@ext:a`, true],
    ['attribute', 'o:b', 'w:r', 'org_docx4j_wml.R', `${onP}/w:r/@o:b`, false],
    ['attribute', 'w:madeUp', 'w:r', 'org_docx4j_wml.R', `${onP}/w:r/@w:madeUp`, false],
  ].sort());
});

test('CR-025 section 2.1: mc:Ignorable counts where it stands, on the element itself and on an ancestor', async () => {
  const pkg = await loaded({ 'word/document.xml': `<w:document xmlns:w="${W}" xmlns:mc="${MC}"><w:body><w:p><w:pPr xmlns:e="${EXT}" mc:Ignorable="e"><e:here/></w:pPr></w:p><w:p><w:pPr><e:there xmlns:e="${EXT}"/></w:pPr></w:p></w:body></w:document>` });
  const main = pkg.getMainDocumentPart();
  await main.getContents();
  // mc:Ignorable off the root has no place in w:pPr's class either, so it's dropped and reported like any other attribute
  assert.deepEqual(skippedIn(main).filter((w) => w.kind === 'element').map((w) => [w.name, w.ignorable]), [['e:here', true], ['e:there', false]]);
  assert.deepEqual(skippedIn(main).filter((w) => w.kind === 'attribute').map((w) => w.name), ['mc:Ignorable']);
});

test('CR-025 section 2.1: text a class has no place for is reported, its first 80 characters as the name', async () => {
  const long = 'x'.repeat(100);
  const pkg = await documentOf(`<w:p><w:pPr>${long}</w:pPr></w:p>`);
  const main = pkg.getMainDocumentPart();
  await main.getContents();
  const [w, ...rest] = skippedIn(main);
  assert.deepEqual(rest, []);
  assert.equal(w.kind, 'text');
  assert.equal(w.name, 'x'.repeat(80));
  assert.equal(w.parentName, 'w:pPr');
  assert.equal(w.parentClass, 'org_docx4j_wml.PPr');
  assert.equal(w.path, '/w:document/w:body/w:p/w:pPr/text()');
  assert.equal(w.ignorable, false);
});

test('CR-025 section 2.1: a string CR-016 reports as misplaced isn\'t reported twice', async () => {
  const pkg = await documentOf('<w:p><w:r>stray<w:t>a</w:t></w:r></w:p>');
  const main = pkg.getMainDocumentPart();
  await main.getContents();
  assert.deepEqual(skippedIn(main), []);
  assert.deepEqual(main.warnings.filter((w) => w.code === 'target.misplaced').map((w) => w.text), ['stray']);
});

test('CR-025 D9: an element in a w:t that goes with all it holds is reported by CR-021\'s pass, one that gives its text isn\'t', async () => {
  const pkg = await documentOf(`<w:p><w:r><w:t>a<ext:opaque>lost</ext:opaque>b</w:t><w:t>c<w:madeUp>lost</w:madeUp>d</w:t><w:t>e<w:t>kept</w:t>f</w:t><w:t>g<plain xmlns="">kept</plain>h</w:t></w:r></w:p>`);
  const main = pkg.getMainDocumentPart();
  await main.getContents();
  const skipped = skippedIn(main);
  assert.deepEqual(skipped.map((w) => [w.kind, w.name, w.parentName, w.path, w.ignorable]), [
    ['element', 'ext:opaque', 'w:t', '/w:document/w:body/w:p/w:r/w:t/ext:opaque', true],
    ['element', 'w:madeUp', 'w:t', '/w:document/w:body/w:p/w:r/w:t/w:madeUp', false],
  ]);
  assert.ok(skipped.every((w) => !('parentClass' in w)), 'no class has met it');
  assert.equal(main.warnings.filter((w) => w.code === 'br.outside_run').length, 1, 'the pass\'s own count stands');
});

test('CR-025 section 2.1: a Choice taken reports what its typed content drops, where a kept load has lost nothing, a prefix the Choice declared still counting as ignorable', async () => {
  const bytes = new Uint8Array(readFileSync(join(fixturesDir, 'cr022-checkbox-linked.xlsx')));
  const read = async (options) => {
    const pkg = await OpcPackage.load(bytes, options);
    for (const part of pkg.parts) if (typeof part.readContents === 'function') await part.readContents();
    return pkg.warnings.filter((w) => w.code === 'part.skipped');
  };
  assert.deepEqual(await read({ mcePreprocess: false }), []);
  const resolved = await read({ mcePreprocess: true });
  const colour = resolved.find((w) => w.name === 'a14:legacySpreadsheetColorIndex');
  assert.equal(colour.parentName, 'a:srgbClr');
  assert.equal(colour.ignorable, true, 'mc:Ignorable lists a14, whose declaration went with the mc:Choice');
});

test('CR-025 section 2.1: a part never read reports nothing, and has lost nothing', async () => {
  const bytesIn = await (await documentOf('<w:p><w:pPr><ext:opaque/></w:pPr></w:p>')).save();
  const pkg = await WordprocessingMLPackage.load(bytesIn);
  assert.deepEqual(pkg.warnings, []);
  const saved = unzipSync(await pkg.save());
  assert.ok(bytesEqual(saved['word/document.xml'], unzipSync(bytesIn)['word/document.xml']), 'untouched, byte for byte');
});

test('CR-025 section 2.1: a deep copy reports nothing: the callbacks stay at the part\'s own read (objects CR-007)', async () => {
  const pkg = await documentOf('<w:p><w:pPr><w:jc w:val="left"/></w:pPr><w:r><w:t>a</w:t></w:r></w:p>');
  const main = pkg.getMainDocumentPart();
  await main.getContents();
  assert.deepEqual(skippedIn(main), []);
  pkg.changeTrackingMode = 'TrackAll';
  pkg.body.paragraphs[0].alignment = 'Centered'; // recordPPrChange's deepCopyAsSync drops what CT_PPrBase doesn't declare
  assert.deepEqual(skippedIn(main), []);
});

test('CR-025 section 2.1: setting a part\'s contents clears what its read reported', async () => {
  const pkg = await documentOf('<w:p><w:pPr><ext:opaque/></w:pPr></w:p>');
  const main = pkg.getMainDocumentPart();
  await main.getContents();
  assert.equal(skippedIn(main).length, 1);
  main.setContents(main.contents);
  assert.deepEqual(skippedIn(main), []);
});

test('CR-025 section 1: no fixture reports any, the objects package\'s CR-019 having kept NumberingImplicitNumId.docx\'s 15', async () => {
  const files = [];
  // not the git-ignored files checks 72 and 75 build, which hold what they hold by design
  const walk = (dir) => {
    for (const name of readdirSync(dir)) {
      const path = join(dir, name);
      if (statSync(path).isDirectory()) { if (name !== 'check72' && name !== 'check75') walk(path); } else if (/\.(docx|pptx|xlsx)$/i.test(name)) files.push(path);
    }
  };
  walk(fixturesDir);
  walk(join(fixturesDir, '..', 'mc-fun'));
  // what a clean checkout holds, as CR-016's gate: the built files aren't on CI
  assert.ok(files.length > 280, `${files.length} files`);
  const found = new Map();
  for (const file of files) {
    const pkg = await OpcPackage.load(new Uint8Array(readFileSync(file)));
    await pkg.unmarshalAll();
    const skipped = pkg.warnings.filter((w) => w.code === 'part.skipped');
    if (skipped.length > 0) found.set(file.slice(fixturesDir.length + 1), skipped);
  }
  assert.deepEqual([...found.keys()], []);
});

// Phase 2. The element is one the read drops today, in a host that reads: w:pPr.
const DROPPED = '<w:p><w:pPr><ext:opaque/></w:pPr><w:r><w:t>a</w:t></w:r></w:p>';
const bytesOfDocument = (zipped) => unzipSync(zipped)['word/document.xml'];

test('CR-025 section 2.2: a part read and not changed is copied byte for byte, what the read dropped still in it', async () => {
  const bytes = await (await documentOf(DROPPED)).save();
  const pkg = await WordprocessingMLPackage.load(bytes);
  await pkg.getMainDocumentPart().getContents();
  assert.equal(skippedIn(pkg.getMainDocumentPart()).length, 1);
  const saved = await pkg.save();
  assert.ok(bytesEqual(bytesOfDocument(saved), bytesOfDocument(bytes)));
  assert.match(new TextDecoder().decode(bytesOfDocument(saved)), /ext:opaque/);
  assert.deepEqual(pkg.lastSave.lost, []);
  assert.deepEqual(pkg.lastSave.normalised, [], 'copied, so no marshal to report on');
});

test('CR-025 section 2.3: a part changed is refused, nothing written, with python\'s message and this package\'s way out', async () => {
  const pkg = await WordprocessingMLPackage.load(await (await documentOf(DROPPED)).save());
  await pkg.getBody();
  pkg.body.paragraphs[0].insertText(' more', 'End');
  const puts = [];
  const sink = { put: (name) => { puts.push(name); }, finish: async () => { puts.push('finished'); } };
  await assert.rejects(pkg.saveTo(sink), (e) => {
    assert.ok(e instanceof SaveLossException);
    assert.equal(e.code, 'save.content_lost');
    assert.deepEqual(e.losses, [{ partName: '/word/document.xml', kind: 'skipped', count: 1 }]);
    assert.equal(e.message, 'Saving would lose content: /word/document.xml holds 1 item the load skipped (part.skipped). Set pkg.checkLossOnSave = false to write the part without it.');
    return true;
  });
  assert.deepEqual(puts, [], 'nothing written, whatever the sink: a directory sink writes as it goes');
});

test('CR-025 section 2.3: every part that would lose content is listed at once', async () => {
  const pkg = await loaded({
    'word/document.xml': inBody(DROPPED),
    'word/settings.xml': `<w:settings ${roots}><ext:one/><ext:two/></w:settings>`,
  });
  await pkg.getBody();
  const settings = partOf(pkg, '/word/settings.xml');
  await settings.getContents();
  pkg.body.paragraphs[0].insertText(' more', 'End');
  settings.contents.zoom = { percent: 150 };
  await assert.rejects(pkg.save(), (e) => {
    assert.deepEqual([...e.losses].sort((a, b) => a.partName.localeCompare(b.partName)), [
      { partName: '/word/document.xml', kind: 'skipped', count: 1 },
      { partName: '/word/settings.xml', kind: 'skipped', count: 2 },
    ]);
    assert.match(e.message, /write the parts without them\.$/);
    return true;
  });
});

test('CR-025 section 2.4: with checkLossOnSave off the part is written without the items, each on lastSave.lost, through every route', async () => {
  const changed = async () => {
    const pkg = await WordprocessingMLPackage.load(await (await documentOf(DROPPED)).save());
    await pkg.getBody();
    pkg.body.paragraphs[0].insertText(' more', 'End');
    pkg.checkLossOnSave = false;
    return pkg;
  };
  const viaZip = await changed();
  assert.doesNotMatch(new TextDecoder().decode(bytesOfDocument(await viaZip.save())), /ext:opaque/);
  assert.deepEqual(viaZip.lastSave.lost.map((w) => [w.code, w.name]), [['part.skipped', 'ext:opaque']]);
  const viaFlat = await changed();
  assert.doesNotMatch(await viaFlat.saveFlatOpc(), /ext:opaque/);
  assert.equal(viaFlat.lastSave.lost.length, 1);
  const viaSink = await changed();
  await viaSink.saveTo(new MemoryPartSink());
  assert.equal(viaSink.lastSave.lost.length, 1);
});

test('CR-025 section 2.4: clone() carries the setting, and a changed part refuses a clone as it does a save', async () => {
  const pkg = await WordprocessingMLPackage.load(await (await documentOf(DROPPED)).save());
  await pkg.getBody();
  pkg.body.paragraphs[0].insertText(' more', 'End');
  await assert.rejects(pkg.clone(), SaveLossException);
  pkg.checkLossOnSave = false;
  assert.equal((await pkg.clone()).checkLossOnSave, false);
});

test('CR-025 section 2.4: a part read and not changed is copied with the setting off too', async () => {
  const bytes = await (await documentOf(DROPPED)).save();
  const pkg = await WordprocessingMLPackage.load(bytes);
  await pkg.getMainDocumentPart().getContents();
  pkg.checkLossOnSave = false;
  assert.ok(bytesEqual(bytesOfDocument(await pkg.save()), bytesOfDocument(bytes)));
  assert.deepEqual(pkg.lastSave.lost, []);
});

test('CR-025 section 2.3: a part given new contents is never refused, and is written without the element: the way to mend a file', async () => {
  const pkg = await WordprocessingMLPackage.load(await (await documentOf(DROPPED)).save());
  const main = pkg.getMainDocumentPart();
  main.setContents(await main.getContents());
  const saved = new TextDecoder().decode(bytesOfDocument(await pkg.save()));
  assert.doesNotMatch(saved, /ext:opaque/);
  assert.deepEqual(pkg.lastSave.lost, []);
});

test('CR-025 section 2.2: a part given bytes is held to those bytes, not to its source', async () => {
  const pkg = await WordprocessingMLPackage.load(await (await documentOf('<w:p/>')).save());
  const main = pkg.getMainDocumentPart();
  const given = inBody(DROPPED);
  main.setXml(given);
  await main.getContents();
  assert.equal(skippedIn(main).length, 1);
  assert.match(new TextDecoder().decode(bytesOfDocument(await pkg.save())), /ext:opaque/);
});

test('CR-025 D9: what CR-021\'s pass takes counts: read and not changed the part is copied, changed it\'s refused', async () => {
  const bytes = await (await documentOf('<w:p><w:r><w:t>a<ext:opaque>lost</ext:opaque>b</w:t></w:r></w:p>')).save();
  const read = await WordprocessingMLPackage.load(bytes);
  await read.getMainDocumentPart().getContents();
  assert.ok(bytesEqual(bytesOfDocument(await read.save()), bytesOfDocument(bytes)));
  const changed = await WordprocessingMLPackage.load(bytes);
  await changed.getBody();
  changed.body.paragraphs[0].insertText(' more', 'End');
  await assert.rejects(changed.save(), SaveLossException);
});

test('CR-025 section 1: the real file: nothing is lost from its style pane filter, read or changed, since the objects package\'s CR-019', async () => {
  const bytes = new Uint8Array(readFileSync(join(fixturesDir, 'parity', 'NumberingImplicitNumId.docx')));
  const filter = (settings) => new TextDecoder().decode(settings).match(/<w:stylePaneFormatFilter[^>]*>/)[0];
  const attributes = (element) => [...element.matchAll(/w:(\w+)="/g)].map((m) => m[1]).sort();
  const writer = await WordprocessingMLPackage.load(bytes);
  await writer.getBody();
  writer.changeTrackingMode = 'TrackAll';
  const saved = unzipSync(await writer.save());
  assert.deepEqual(writer.lastSave.lost, []);
  assert.equal(attributes(filter(saved['word/settings.xml'])).length, 16);
  assert.deepEqual(attributes(filter(saved['word/settings.xml'])), attributes(filter(unzipSync(bytes)['word/settings.xml'])));
});

test('CR-025 section 2.2 over every fixture: skipped content injected, the part read every way a reader does, and copied byte for byte', async () => {
  const files = [];
  const walk = (dir) => {
    for (const name of readdirSync(dir)) {
      const path = join(dir, name);
      if (statSync(path).isDirectory()) { if (name !== 'check72' && name !== 'check75') walk(path); } else if (/\.docx$/i.test(name)) files.push(path);
    }
  };
  walk(fixturesDir);
  walk(join(fixturesDir, '..', 'mc-fun'));
  assert.ok(files.length > 250, `${files.length} files`);
  const refused = [];
  let held = 0;
  for (const file of files) {
    const parts = unzipSync(new Uint8Array(readFileSync(file)));
    const name = 'word/document.xml';
    if (!parts[name]) continue;
    const xml = new TextDecoder().decode(parts[name]);
    const at = /<w:p[ >]/.exec(xml);
    if (!at) continue;
    const injected = `${xml.slice(0, at.index + 4)} xmlns:zz="urn:docx4j:cr025:zz" zz:mark="1"${xml.slice(at.index + 4)}`;
    const bytes = zipSync({ ...parts, [name]: strToU8(injected) });
    const pkg = await WordprocessingMLPackage.load(bytes);
    await pkg.unmarshalAll();
    const main = pkg.getMainDocumentPart();
    if (!skippedIn(main).some((w) => w.name === 'zz:mark')) continue;
    try {
      const body = await pkg.getBody();
      for (const p of body.paragraphs) void p.text;
      PlainText.of(await Texts.of(pkg));
      const saved = await pkg.save();
      if (!bytesEqual(bytesOfDocument(saved), strToU8(injected))) refused.push(`${file}: not copied`);
      held++;
    } catch (e) { refused.push(`${file}: ${e.message.split('\n')[0]}`); }
  }
  assert.deepEqual(refused, [], 'a read changes nothing a marshal shows');
  assert.ok(held > 200, `${held} files held`);
});
