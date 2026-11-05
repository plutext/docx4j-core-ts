// CR-027: w:numId 0 clears every inherited left, hanging and first-line indent, the right one staying, and only a w:ind stated
// with the w:numId 0 or below it counts. docx4j 9e3100351's rule (StyleUtil; docx4j-export-fo/docs/word-layout-rules.md at s28numid0,
// probe numbering-off-style-ind, Word's goldens of 5 November 2026). docx4j's NumIdZeroIndentTest mirrored, its styles and numbering
// as docx4j's, its indents here in twips through getEffectivePPr and in points through Paragraph. No golden has this shape, so
// parity can't hold it: docx4j's answers are its test's, until the pom moves to the release holding the rule.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import { WordprocessingMLPackage } from '../dist/index.mjs';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

const level = (ind) => `<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/>${ind}</w:lvl>`;
const NUMBERING = `<w:numbering ${W}>`
  + `<w:abstractNum w:abstractNumId="0">${level('<w:pPr><w:ind w:left="1440" w:hanging="1152"/></w:pPr>')}</w:abstractNum>`
  + `<w:abstractNum w:abstractNumId="1">${level('')}</w:abstractNum>`
  + `<w:abstractNum w:abstractNumId="2">${level('<w:pPr><w:ind w:left="720" w:right="57"/></w:pPr>')}</w:abstractNum>`
  + '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>'
  + '<w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>'
  + '<w:num w:numId="3"><w:abstractNumId w:val="2"/></w:num>'
  + '</w:numbering>';

const style = (id, inner) => `<w:style w:type="paragraph" w:styleId="${id}"><w:name w:val="${id}"/>${inner}</w:style>`;
const NUM1 = '<w:numPr><w:numId w:val="1"/></w:numPr>';
const NUM_OFF = '<w:numPr><w:numId w:val="0"/></w:numPr>';
const STYLES = `<w:styles ${W}>`
  + '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>'
  // docx4j's
  + style('Numbered', '<w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>')
  + style('OwnInd', `<w:pPr>${NUM1}<w:ind w:left="454" w:right="57" w:hanging="454"/></w:pPr>`)
  + style('IndBase', '<w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>')
  + style('OnBase', `<w:basedOn w:val="IndBase"/><w:pPr>${NUM1}</w:pPr>`)
  + style('Switch', `<w:basedOn w:val="OwnInd"/><w:pPr>${NUM_OFF}</w:pPr>`)
  + style('SwitchInd', `<w:basedOn w:val="OwnInd"/><w:pPr>${NUM_OFF}<w:ind w:left="300" w:hanging="100"/></w:pPr>`)
  // section 3's
  + style('NoList', '<w:pPr><w:numPr><w:numId w:val="9"/></w:numPr><w:ind w:left="454" w:hanging="454"/></w:pPr>')
  + style('Chars', `<w:pPr>${NUM1}<w:ind w:leftChars="100" w:rightChars="50" w:hangingChars="25"/></w:pPr>`)
  + style('FirstLine', `<w:pPr>${NUM1}<w:ind w:left="454" w:firstLine="227"/></w:pPr>`)
  + style('StartEnd', `<w:pPr>${NUM1}<w:ind w:start="454" w:startChars="100" w:end="57" w:endChars="50" w:firstLineChars="25"/></w:pPr>`)
  + style('NoLevelInd', '<w:pPr><w:numPr><w:numId w:val="2"/></w:numPr><w:ind w:left="454" w:hanging="454"/></w:pPr>')
  + style('LevelRight', '<w:pPr><w:numPr><w:numId w:val="3"/></w:numPr></w:pPr>')
  + style('DecimalList', '<w:pPr><w:numPr><w:numId w:val="1.0"/></w:numPr><w:ind w:left="454" w:right="57" w:hanging="454"/></w:pPr>')
  + '</w:styles>';

const OFF = `<w:numPr><w:ilvl w:val="0"/><w:numId w:val="0"/></w:numPr>`;

/** The effective w:ind of one paragraph per [style, direct pPr], and the paragraphs. */
async function read(cases) {
  const base = unzipSync(await (await WordprocessingMLPackage.createPackage()).save());
  const body = cases.map(([s, pPr]) => `<w:p><w:pPr><w:pStyle w:val="${s}"/>${pPr}</w:pPr><w:r><w:t>heading</w:t></w:r></w:p>`).join('');
  const pkg = await WordprocessingMLPackage.load(zipSync({
    ...base,
    '[Content_Types].xml': strToU8(strFromU8(base['[Content_Types].xml']).replace('</Types>',
      '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/></Types>')),
    'word/_rels/document.xml.rels': strToU8(strFromU8(base['word/_rels/document.xml.rels']).replace('</Relationships>',
      '<Relationship Id="rIdN" Target="numbering.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering"/></Relationships>')),
    'word/styles.xml': strToU8(DECL + STYLES),
    'word/numbering.xml': strToU8(DECL + NUMBERING),
    'word/document.xml': strToU8(`${DECL}<w:document ${W}><w:body>${body}</w:body></w:document>`),
  }));
  await pkg.getMainDocumentPart().getContents();
  const resolver = await pkg.getPropertyResolver();
  const paragraphs = pkg.body.paragraphs;
  return { paragraphs, inds: paragraphs.map((p) => resolver.getEffectivePPr(p.p.pPr).ind ?? {}) };
}

const KEYS = ['left', 'leftChars', 'start', 'startChars', 'right', 'rightChars', 'end', 'endChars', 'hanging', 'hangingChars', 'firstLine', 'firstLineChars'];
const plainInd = (ind) => Object.fromEntries(KEYS.filter((k) => ind[k] !== undefined).map((k) => [k, Number(ind[k])]));

// docx4j's cases, its pt in twips: 2.85pt is 57, 15pt 300, -5pt a hanging 100, -10pt a hanging 200, 7.1pt 142
const DOCX4J = [
  ['a numbered paragraph, the level\'s', 'Numbered', '', { left: 1440, hanging: 1152 }],
  ['numId 0 leaves only the paragraph\'s own w:ind', 'Numbered', `${OFF}<w:ind w:left="720"/>`, { left: 720 }],
  ['N1, the style\'s own left and hanging go, its right stays', 'OwnInd', OFF, { right: 57 }],
  ['N4, and the w:ind of the style it is based on', 'OnBase', OFF, {}],
  ['N7, w:numId 0 in a style', 'Switch', '', { right: 57 }],
  ['N8, a w:ind in the style that switches off stands', 'SwitchInd', '', { left: 300, right: 57, hanging: 100 }],
  ['N10, a direct left, no hanging', 'OwnInd', `${OFF}<w:ind w:left="300"/>`, { left: 300, right: 57 }],
  ['N11, a direct hanging, no left', 'OwnInd', `${OFF}<w:ind w:hanging="200"/>`, { right: 57, hanging: 200 }],
  ['N14, w:numId 0 in a style, a direct left', 'Switch', '<w:ind w:left="142"/>', { left: 142, right: 57 }],
];

test('docx4j\'s NumIdZeroIndentTest, its nine cases through getEffectivePPr', async () => {
  const { inds } = await read(DOCX4J.map(([, s, pPr]) => [s, pPr]));
  DOCX4J.forEach(([name, , , want], i) => assert.deepEqual(plainInd(inds[i]), want, name));
});

test('Paragraph reads them in points', async () => {
  const { paragraphs } = await read([['OwnInd', OFF], ['OwnInd', `${OFF}<w:ind w:left="300"/>`], ['OwnInd', `${OFF}<w:ind w:hanging="200"/>`], ['OwnInd', '']]);
  const indents = paragraphs.map((p) => [p.leftIndent, p.firstLineIndent, p.rightIndent]);
  assert.deepEqual(indents[0], [0, 0, 2.85]);
  assert.deepEqual(indents[1], [15, 0, 2.85]);
  assert.deepEqual(indents[2], [0, -10, 2.85]);
  assert.deepEqual(indents[3], [22.7, -22.7, 2.85]); // no numId 0: the style's own, stated over the level's
});

test('an inherited w:numPr naming no list leaves the indent alone', async () => {
  const { inds } = await read([['NoList', OFF]]);
  assert.deepEqual(plainInd(inds[0]), { left: 454, hanging: 454 });
});

test('every left and first-line form clears; every right form stays', async () => {
  const { inds } = await read([['Chars', OFF], ['FirstLine', OFF], ['StartEnd', OFF]]);
  assert.deepEqual(plainInd(inds[0]), { rightChars: 50 });
  assert.deepEqual(plainInd(inds[1]), {});
  assert.deepEqual(plainInd(inds[2]), { end: 57, endChars: 50 });
});

test('a level with no w:ind clears as docx4j\'s does, which Word\'s goldens don\'t measure, and a right equal to the level\'s stays', async () => {
  const { inds } = await read([['NoLevelInd', OFF], ['LevelRight', OFF]]);
  assert.deepEqual(plainInd(inds[0]), {});
  assert.deepEqual(plainInd(inds[1]), { right: 57 });
});

// python's CR-036 found its numbering_off comparing a raw "0.0" with 0; here the objects package truncates a decimal before the model sees it
test('a w:numId written as a decimal reads as its integer, on either side', async () => {
  const { inds } = await read([['OwnInd', '<w:numPr><w:numId w:val="0.0"/></w:numPr>'], ['DecimalList', OFF]]);
  assert.deepEqual(plainInd(inds[0]), { right: 57 });
  assert.deepEqual(plainInd(inds[1]), { right: 57 });
});

test('w:numId 0 on a paragraph with nothing inherited changes nothing', async () => {
  const { inds } = await read([['IndBase', OFF]]);
  assert.deepEqual(plainInd(inds[0]), { left: 720, hanging: 360 });
});
