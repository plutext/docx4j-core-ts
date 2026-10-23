// CR-007 phase 1: the table context in the PropertyResolver (docx4j CR-030). What docx4j answers
// for every table paragraph of the fixtures is test/parity.test.mjs's business; here is what a
// golden cannot say: the overloads without a context are untouched, the caches follow refresh(),
// the settings part is read without being marked, and the edges of the helpers.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import {
  WordprocessingMLPackage, PropertyResolver, ZipPartStore, TableContextTracker, TableStyleConditions,
  CellContext, enclosingCellOf, parseXml, unmarshalNode,
} from '../dist/index.mjs';
import { fixture, bytesEqual, plain } from './helpers.mjs';

async function load(name) {
  const pkg = await WordprocessingMLPackage.load(await fixture(`parity/${name}.docx`));
  const resolver = await pkg.getPropertyResolver();
  const body = (await pkg.getMainDocumentPart().getContents()).body;
  const paragraphs = [];
  new TableContextTracker(resolver).walk(body, (p, ctx) => paragraphs.push({ p, ctx, text: textOf(p) }));
  return { pkg, resolver, body, paragraphs, find: (text) => paragraphs.find((x) => x.text === text) };
}

function textOf(p) {
  return (p.content ?? []).map((e) => (e.value?.content ?? []).map((t) => (typeof t.value?.value === 'string' ? t.value.value : '')).join('')).join('');
}

const firstRun = (p) => p.content.find((e) => e.value?.TYPE_NAME === 'org_docx4j_wml.R').value;

test('a paragraph in a header row reads the table style\'s formatting only with its context', async () => {
  const { resolver, find } = await load('tables-conditions');
  // T03: no w:tblLook, so first row and first column on; r0c1 is under firstRow and band-less
  const { p, ctx } = find('c r0c1');
  assert.equal(TableStyleConditions.key(ctx.conditions), 'firstRow');
  assert.deepEqual(ctx.textConditions.map((pr) => pr.type), ['wholeTable', 'firstRow']);
  assert.equal(ctx.tableStyleId, 'Full');
  assert.equal(ctx.formatsText, true);

  const plainRPr = resolver.getEffectiveRPr(firstRun(p).rPr, p.pPr);
  assert.equal(plainRPr.b, undefined, 'without a context the resolver knows nothing of the table');
  assert.equal(plainRPr.sz.val, 24);
  const inCell = resolver.getEffectiveRPr(firstRun(p).rPr, p.pPr, ctx);
  assert.equal(inCell.b.val ?? true, true, 'bold because the header row is bold');
  assert.equal(inCell.smallCaps.val ?? true, true, 'Full restates firstRow: merged per condition with Base\'s');
  assert.equal(inCell.color.val, '0000F1');
  assert.equal(inCell.sz.val, 20, 'the table style\'s own size, mode 15: Normal states none');
  assert.equal(inCell.rFonts.ascii, 'Liberation Sans');
  assert.equal(resolver.getEffectivePPr(p.pPr, ctx).jc.val, 'left', 'Normal\'s w:jc over the first row\'s centre, in mode 15');
  assert.equal(resolver.getEffectiveParagraphMarkRPr(p.pPr, ctx).color.val, '0000F1');

  // the same context through the PARENT pointers, one paragraph at a time
  const viaParents = resolver.cellContextOf(p);
  assert.ok(viaParents.key.equals(ctx.key));
  assert.equal(String(viaParents), 'CellContext[Full firstRow]');
  assert.ok(enclosingCellOf(p).tbl.TYPE_NAME === 'org_docx4j_wml.Tbl');
});

test('an undefined context, and a table style which gives text nothing, resolve as without one', async () => {
  const { resolver, find, paragraphs } = await load('tables-conditions');
  const outside = paragraphs.find((x) => x.text.startsWith('T01'));
  assert.equal(outside.ctx, undefined);
  assert.equal(resolver.cellContextOf(outside.p), undefined);
  assert.deepEqual(plain(resolver.getEffectivePPr(outside.p.pPr, undefined)), plain(resolver.getEffectivePPr(outside.p.pPr)));
  assert.equal(resolver.getEffectivePPr(outside.p.pPr, undefined), resolver.getEffectivePPr(outside.p.pPr), 'the same cached object');

  // T13: NoText's conditions format cells, not text; T11: Normal Table gives text nothing
  for (const text of ['m r0c0', 'k r0c0', 'l r0c0']) {
    const { p, ctx } = find(text);
    assert.ok(ctx instanceof CellContext, text);
    assert.equal(ctx.formatsText, false, text);
    assert.deepEqual(plain(resolver.getEffectiveRPr(firstRun(p).rPr, p.pPr, ctx)), plain(resolver.getEffectiveRPr(firstRun(p).rPr, p.pPr)), text);
    assert.equal(resolver.getEffectivePPr(p.pPr, ctx), resolver.getEffectivePPr(p.pPr), text);
  }
  assert.equal(find('k r0c0').ctx.tableStyleId, 'TableNormal', 'a table naming no style takes the default table style');
  assert.equal(find('l r0c0').ctx.tableStyleId, 'Missing');

  // a paragraph built by hand has no PARENT pointers: no context, and no error
  assert.equal(resolver.cellContextOf({ TYPE_NAME: 'org_docx4j_wml.P' }), undefined);
  assert.equal(enclosingCellOf({ TYPE_NAME: 'org_docx4j_wml.P' }), undefined);
});

test('a nested table\'s paragraphs are in the nested table, and a tracker with no resolver tracks nothing', async () => {
  const { find, body } = await load('tables-conditions');
  assert.equal(find('j nested r0c0').ctx.tableStyleId, 'NoText');
  assert.equal(find('j r0c1 after it').ctx.tableStyleId, 'Full', 'the cell\'s own table again after the nested one');
  assert.equal(TableStyleConditions.key(find('j r1c1 in a block control').ctx.conditions), 'band1Horz');
  assert.equal(TableStyleConditions.key(find('j r2c0 in a cell control').ctx.conditions), 'band2Horz-firstCol');
  let seen = 0;
  new TableContextTracker(undefined).walk(body, (p, ctx) => { seen++; assert.equal(ctx, undefined); });
  assert.ok(seen > 190);
});

test('the size and justification exception follows the settings part, read without marking it', async () => {
  const mode15 = await load('tables-conditions');
  const mode12 = await load('tables-conditions-mode12');
  assert.equal(mode15.resolver.appliesTableStyleSizeJcException(), false);
  assert.equal(mode12.resolver.appliesTableStyleSizeJcException(), true);
  assert.equal(PropertyResolver.fromSource({}).appliesTableStyleSizeJcException(), true, 'no settings is mode 12');

  // c r0c1 is in Normal, which states left and no size: the table style's 10pt stands over the
  // document default in both (the table level is above the defaults), and in mode 12 Normal's
  // left gives way to the first row's centre
  const at = (doc, text) => {
    const { p, ctx } = doc.find(text);
    return { sz: doc.resolver.getEffectiveRPr(firstRun(p).rPr, p.pPr, ctx).sz.val, jc: doc.resolver.getEffectivePPr(p.pPr, ctx).jc.val };
  };
  assert.deepEqual(at(mode15, 'c r0c1'), { sz: 20, jc: 'left' });
  assert.deepEqual(at(mode12, 'c r0c1'), { sz: 20, jc: 'center' });
  // BoldPara states a size and a justification of its own, and keeps them in both
  assert.deepEqual(at(mode15, 'g r0c0'), { sz: 30, jc: 'right' });
  assert.deepEqual(at(mode12, 'g r0c0'), { sz: 30, jc: 'right' });

  // the settings part was read privately: it, and the styles part, are saved from the source
  const bytes = await fixture('parity/tables-conditions.docx');
  const saved = await mode15.pkg.save();
  for (const name of ['word/settings.xml', 'word/styles.xml']) {
    assert.ok(bytesEqual(await new ZipPartStore(bytes).load(name), await new ZipPartStore(saved).load(name)), `${name} changed`);
  }

  // once the part is unmarshalled, refresh() reads the live tree
  const settings = await mode15.pkg.getMainDocumentPart().documentSettingsPart.getContents();
  settings.compat.compatSetting.find((cs) => cs.name === 'compatibilityMode').val = '14';
  settings.compat.compatSetting.find((cs) => cs.name === 'overrideTableStyleFontSizeAndJustification').val = '0';
  assert.equal(mode15.resolver.appliesTableStyleSizeJcException(), false, 'read once per refresh');
  mode15.resolver.refresh();
  assert.equal(mode15.resolver.appliesTableStyleSizeJcException(), true);
  assert.deepEqual(at(mode15, 'c r0c1'), { sz: 20, jc: 'center' });
});

test('refresh() discards the composed table levels', async () => {
  const { pkg, resolver, find } = await load('tables-conditions');
  const { p, ctx } = find('c r0c1');
  assert.equal(resolver.getEffectiveRPr(firstRun(p).rPr, p.pPr, ctx).color.val, '0000F1');
  const styles = await pkg.getMainDocumentPart().styleDefinitionsPart.getContents();
  resolver.refresh();                                    // the live styles tree from here on
  assert.equal(resolver.getEffectiveRPr(firstRun(p).rPr, p.pPr, resolver.cellContextOf(p)).color.val, '0000F1');
  const base = styles.style.find((s) => s.styleId === 'Base');
  base.tblStylePr.find((pr) => pr.type === 'firstRow').rPr.color.val = 'ABCDEF';
  const stale = resolver.cellContextOf(p);
  assert.equal(resolver.getEffectiveRPr(firstRun(p).rPr, p.pPr, stale).color.val, '0000F1', 'cached until refresh');
  resolver.refresh();
  const fresh = resolver.cellContextOf(p);
  assert.equal(resolver.getEffectiveRPr(firstRun(p).rPr, p.pPr, fresh).color.val, 'ABCDEF');
});

test('the table style is decided by name: Normal Table is Word\'s built-in wherever it is', async () => {
  // docx4j CR-030 probe T2: the default table style renamed My Default, and a non-default style
  // named "Normal Table"; Word gave (a) the built-in's margins and no text formatting
  const { resolver, find } = await load('tables-named-normal-table');
  const a = find('(a) cell text');
  assert.equal(a.ctx.tableStyleId, 'OtherNormal');
  assert.equal(a.ctx.formatsText, false);
  assert.equal(resolver.getTableStyleChain('OtherNormal').rPr, undefined);
  assert.equal(resolver.reachesDefaultTableStyle({ tblStyle: { val: 'OtherNormal' } }), true);
  assert.equal(resolver.getEffectiveTableStyle({ tblStyle: { val: 'OtherNormal' } }).tblPr.tblCellMar.left.w, 108);
  // (b) names no style: the default, My Default, applies as written, text formatting included
  const b = find('(b) cell text');
  assert.equal(b.ctx.tableStyleId, resolver.getTableStyleIdOf(undefined));
  assert.equal(resolver.getEffectiveRPr(firstRun(b.p).rPr, b.p.pPr, b.ctx).sz.val, 40);
  assert.equal(resolver.reachesDefaultTableStyle(undefined), false);
  assert.equal(resolver.getTableStyleChain(undefined).TYPE_NAME, 'org_docx4j_wml.Style');
});

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

/**
 * docx4j's table-style-size-trigger probes' document (its CR-001 batch 53, corpus document
 * 1912): docDefaults 11pt, Normal 12pt, one two-row table per table style.
 */
const TRIGGER_STYLES = `<w:styles ${W}>`
  + '<w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults>'
  + '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/>'
  + '<w:rPr><w:sz w:val="24"/></w:rPr></w:style>'
  + '<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/></w:style>'
  + '<w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/><w:basedOn w:val="TableNormal"/>'
  + '<w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:style>'
  + '<w:style w:type="table" w:styleId="PlainGrid"><w:name w:val="PlainGrid"/><w:basedOn w:val="TableNormal"/>'
  + '<w:tblPr><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders></w:tblPr></w:style>'
  + '<w:style w:type="table" w:styleId="RPrColour"><w:name w:val="RPrColour"/><w:basedOn w:val="TableNormal"/>'
  + '<w:rPr><w:color w:val="1F3864"/></w:rPr></w:style>'
  + '<w:style w:type="table" w:styleId="FirstRowOnly"><w:name w:val="FirstRowOnly"/><w:basedOn w:val="TableNormal"/>'
  + '<w:tblStylePr w:type="firstRow"><w:rPr><w:b/></w:rPr></w:tblStylePr></w:style>'
  + '<w:style w:type="table" w:styleId="Nothing"><w:name w:val="Nothing"/><w:basedOn w:val="TableNormal"/></w:style>'
  + '</w:styles>';

const TRIGGER_TABLES = ['TableGrid', 'PlainGrid', 'RPrColour', 'FirstRowOnly', 'Nothing'];

function triggerDocument() {
  const p = (text) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
  const tc = (content) => `<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>${content}</w:tc>`;
  let out = `<w:document ${W}><w:body>`;
  for (const style of TRIGGER_TABLES) {
    out += `<w:tbl><w:tblPr><w:tblStyle w:val="${style}"/>`
      + '<w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="0" w:lastColumn="0" w:noHBand="1" w:noVBand="1"/>'
      + '</w:tblPr><w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid>'
      + `<w:tr>${tc(p(`${style} row 1`))}</w:tr><w:tr>${tc(p(`${style} row 2`))}</w:tr></w:tbl>${p(`after ${style}`)}`;
  }
  return `${out}</w:body></w:document>`;
}

/**
 * docx4j `TableContextResolutionTest.theExceptionWhereTheTableStyleFormatsTextWithoutASize`
 * (900b6a212, 17.3.2): below mode 15 the size exception applies wherever the table style formats
 * text at all - a w:pPr, an rPr without a size, a conditional format, whether or not it applies
 * to the row - and gives the table level's size, the docDefaults' where the style states none. A
 * style which formats no text leaves Normal's 12pt. In mode 15 every table takes Normal's.
 * Measured in Word by docx4j: the table-style-size-trigger probes. And where the docDefaults state
 * no size either, Normal keeps its 12pt (docx4j's corpus documents 10244 and 12723).
 */
test('the size exception wherever the table style formats text, without stating a size', async () => {
  for (const docDefaultsSize of [true, false]) {
    for (const mode of [12, 15]) {
      const stylesXml = docDefaultsSize ? TRIGGER_STYLES
        : TRIGGER_STYLES.replace('<w:rPrDefault><w:rPr><w:sz w:val="22"/></w:rPr></w:rPrDefault>', '');
      const styles = (await unmarshalNode(parseXml(stylesXml))).value;
      const settings = (await unmarshalNode(parseXml(`<w:settings ${W}><w:compat>`
        + `<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="${mode}"/>`
        + '</w:compat></w:settings>'))).value;
      const resolver = PropertyResolver.fromSource({ styles, settings });
      const body = (await unmarshalNode(parseXml(triggerDocument()))).value.body;
      const paragraphs = new Map();
      new TableContextTracker(resolver).walk(body, (p, ctx) => paragraphs.set(textOf(p), { p, ctx }));

      for (const style of TRIGGER_TABLES) {
        const formats = style !== 'PlainGrid' && style !== 'Nothing';
        const expected = mode === 12 && formats && docDefaultsSize ? 22 : 24;
        for (const row of [' row 1', ' row 2']) {
          const { p, ctx } = paragraphs.get(style + row);
          const where = `mode ${mode}${docDefaultsSize ? '' : ', no docDefaults size'}, ${style}${row}`;
          assert.equal(ctx.styleFormatsText, formats, where);
          assert.equal(resolver.getEffectiveRPr(firstRun(p).rPr, p.pPr, ctx).sz.val, expected, where);
          // docx4j's paragraph mark overload composes the same way (not measured on its own)
          assert.equal(resolver.getEffectiveParagraphMarkRPr(p.pPr, ctx).sz.val, expected, `${where}, the mark`);
        }
      }
      // FirstRowOnly's second row is under no condition with text formatting, and still comes in
      assert.equal(paragraphs.get('FirstRowOnly row 2').ctx.formatsText, false);

      const after = paragraphs.get('after TableGrid');
      assert.equal(after.ctx, undefined);
      assert.equal(resolver.getEffectiveRPr(firstRun(after.p).rPr, after.p.pPr, after.ctx).sz.val, 24, 'outside a table, Normal\'s');
    }
  }
});

test('TableStyleConditions: the look, the cnfStyle bits and the key', () => {
  const { lookOf, DEFAULT_LOOK, fromCnf, key, atPosition, resolve, rowConditions, gate } = TableStyleConditions;
  assert.deepEqual(lookOf(undefined), DEFAULT_LOOK);
  assert.deepEqual(lookOf({ val: '04A0' }), DEFAULT_LOOK);
  assert.deepEqual(lookOf({ val: ' ' }), DEFAULT_LOOK);
  assert.deepEqual(lookOf({ val: 'zz' }), DEFAULT_LOOK, 'unreadable: Word\'s default');
  assert.deepEqual(lookOf({ val: '0000' }), { firstRow: false, lastRow: false, firstColumn: false, lastColumn: false, hBand: true, vBand: true });
  // any of the six attributes wins over the bitmask
  assert.deepEqual(lookOf({ val: '01E0', lastRow: '1' }), { firstRow: false, lastRow: true, firstColumn: false, lastColumn: false, hBand: true, vBand: true });
  assert.deepEqual(lookOf({ firstRow: 'true', noHBand: 'on', noVBand: '0' }), { firstRow: true, lastRow: false, firstColumn: false, lastColumn: false, hBand: false, vBand: true });

  assert.deepEqual([...fromCnf('100000000000')], ['firstRow']);
  assert.deepEqual([...fromCnf('0000101')], ['band1Vert', 'band1Horz'], 'a short string is read as far as it goes');
  assert.equal(fromCnf(undefined).size, 0);
  assert.equal(key(new Set(['nwCell', 'firstRow', 'firstCol', 'wholeTable'])), 'firstCol-firstRow-nwCell');
  assert.equal(key(new Set()), '');

  const all = { firstRow: true, lastRow: true, firstColumn: true, lastColumn: true, hBand: true, vBand: true };
  assert.equal(key(atPosition(all, 1, 1, 0, 3, 0, 1, 3)), 'firstCol-firstRow-nwCell');
  assert.equal(key(atPosition(all, 1, 1, 2, 3, 1, 2, 3)), 'lastCol-lastRow-seCell', 'a span reaching the last column is in it');
  assert.equal(key(atPosition(all, 2, 2, 3, 9, 3, 1, 9)), 'band2Vert-band2Horz', 'bands two deep, counted after the first row and column');
  // a row cache decides the row axis, the look gates it, and the corners are derived
  assert.equal(key(resolve(all, 1, 1, 1, 3, 0, 1, 3, { val: '100000000000' }, undefined, undefined)), 'firstCol-firstRow-nwCell');
  assert.equal(key(resolve(DEFAULT_LOOK, 1, 1, 1, 3, 1, 1, 3, { val: '010000000000' }, undefined, undefined)), '', 'lastRow is off in the look');
  assert.equal(key(rowConditions(all, 1, 2, 3, undefined)), 'lastRow');
  assert.equal(key(gate(new Set(['lastCol', 'band1Horz']), DEFAULT_LOOK)), 'band1Horz');
});

// ---------------------------------------------------------------- phase 2: the content API reads in context

/** Every Paragraph view of a body, those in table cells (and nested tables) included. */
function viewsOf(body) {
  const out = [...body.paragraphs];
  const visit = (table) => {
    for (const row of table.rows) {
      for (const cell of row.cells) {
        out.push(...cell.body.paragraphs);
        for (const nested of cell.body.tables) visit(nested);
      }
    }
  };
  for (const table of body.tables) visit(table);
  return out;
}

test('Font and Paragraph read the table style\'s formatting in a cell; direct reads do not', async () => {
  const pkg = await WordprocessingMLPackage.load(await fixture('parity/tables-conditions.docx'));
  const views = viewsOf(await pkg.getBody());
  const view = (text) => views.find((p) => p.text === text);

  const header = view('c r0c1');
  assert.equal(String(header.cellContext), 'CellContext[Full firstRow]');
  assert.equal(header.font.bold, true, 'bold because the header row is bold');
  assert.equal(header.getFont({ direct: true }).bold, false);
  assert.equal(header.font.color, '#0000F1');
  assert.equal(header.font.size, 10);
  assert.equal(header.font.name, 'Liberation Sans', 'the document font follows the table style\'s w:rFonts');
  assert.equal(header.effectiveParagraphMarkRPr.color.val, '0000F1');
  assert.equal(header.effectivePPr.keepNext.val ?? true, true, 'the whole-table condition\'s w:keepNext');
  assert.equal(header.formatting().spaceBefore, 4, 'firstRow\'s w:spacing w:before 80, merged over Base\'s');
  assert.equal(header.formatting({ direct: true }).spaceBefore, 0);
  // a range of it the same
  assert.equal(header.search('r0c1')[0].font.bold, true);

  // a body row: banded, not bold; a paragraph style over the table level; outside any table
  assert.equal(view('c r1c1').font.bold, false);
  assert.equal(view('c r1c1').font.color, '#0000B1');
  assert.equal(view('g r0c0').font.bold, false, 'BoldPara over a bold first row and first column: the toggles combine');
  assert.equal(view('g r0c0').font.size, 15);
  const outside = views.find((p) => p.text.startsWith('T01'));
  assert.equal(outside.cellContext, undefined);
  assert.equal(outside.font.bold, false);
  assert.equal(outside.font.size, 12);
  // a nested table's paragraph is in the nested table, whose style gives text nothing
  assert.equal(view('j nested r0c0').font.color, '#000000', 'automatic, as Office JS reads it (CR-002 section 59), not the enclosing table\'s colour');
  assert.equal(view('j nested r0c0').cellContext.tableStyleId, 'NoText');
});

test('an alignment a cell\'s paragraph inherits from its table style is not written', async () => {
  // mode 12: Normal's left gives way to the first row's centre, so "Centered" is what the header
  // paragraph inherits already and the setter writes no w:jc; "Left" is a value of its own
  const pkg = await WordprocessingMLPackage.load(await fixture('parity/tables-conditions-mode12.docx'));
  const views = viewsOf(await pkg.getBody());
  const header = views.find((p) => p.text === 'c r0c1');
  assert.equal(header.alignment, 'Centered');
  header.alignment = 'Centered';
  assert.equal(header.p.pPr?.jc, undefined);
  header.alignment = 'Left';
  assert.equal(header.p.pPr.jc.val, 'left');
  assert.equal(header.alignment, 'Left');
  header.alignment = 'Centered';
  assert.equal(header.p.pPr?.jc, undefined, 'back to what the table style gives');
});
