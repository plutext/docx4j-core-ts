// CR-012 phase 3: Markdown from the text model (src/model/text/MarkdownExporter.mts). docx4j-markdown's
// MarkdownTextModelTest, MarkdownExportTest, MarkdownExportExtensionsTest and MarkdownAlternateContentTest,
// ported with docx4j's expected strings, where they don't go through docx4j's Markdown importer (this
// package has none: those build their document here instead); and this package's own: blocks(), the image
// destination function, the package unchanged (REQ-204), block quotes with and without the package.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { unzipSync } from 'fflate';
import {
  WordprocessingMLPackage, Texts, PlainText, TextOptions, MarkdownExporter, MarkdownExportOptions,
  NumberingDefinitionsPart, EndnotesPart, CommentsPart, ImagePart, XmlPart, LinkTargets, Namespaces, setLogger,
} from '../dist/index.mjs';

const here = dirname(fileURLToPath(import.meta.url));
const phase0 = join(here, 'text-extraction', 'suite', 'phase0');

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const DATE = ' w:author="a" w:date="2026-01-01T00:00:00Z"';

/**
 * docx4j's `createPackage()` with what a test adds: styles (fragments appended to the styles part), a
 * numbering part (a `w:numbering` body), then the body's content as XML.
 */
async function docx({ styles = [], numbering, docDefaultsFont, body = '' } = {}) {
  const pkg = await WordprocessingMLPackage.createPackage();
  const mdp = pkg.getMainDocumentPart();
  if (styles.length > 0 || docDefaultsFont) {
    const part = mdp.styleDefinitionsPart;
    let xml = await part.getXml();
    if (docDefaultsFont) {
      // the document defaults' run fonts, as docx4j's test sets them on rPrDefault
      xml = xml.replace(/<w:rPrDefault>\s*<w:rPr>/, `$&<w:rFonts w:ascii="${docDefaultsFont}" w:hAnsi="${docDefaultsFont}"/>`);
      xml = xml.replace(/(<w:rPrDefault>\s*<w:rPr>(?:(?!<\/w:rPr>).)*?<w:rFonts [^>]*\/>)\s*<w:rFonts [^>]*\/>/s, '$1');
    }
    part.setXml(xml.replace('</w:styles>', `${styles.join('')}</w:styles>`));
  }
  if (numbering !== undefined) {
    const ndp = new NumberingDefinitionsPart();
    ndp.setXml(`<w:numbering ${W}>${numbering}</w:numbering>`);
    mdp.addTargetPart(ndp);
  }
  // the resolver was built by createPackage, before these parts existed
  await pkg.refresh();
  if (body !== '') await pkg.body.insertXml(body, 'End');
  return pkg;
}

const md = (pkg, options) => new MarkdownExporter(options).export(pkg);
const style = (xml) => xml.replace('<w:style ', `<w:style ${W} `);

// ---------------------------------------------------------------- MarkdownTextModelTest

test('a HYPERLINK field\'s result is a link to its instruction\'s address, whatever field_codes says (REQ-703)', async () => {
  const pkg = await docx({ body: '<w:p><w:r><w:t xml:space="preserve">See </w:t></w:r>'
    + '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
    + '<w:r><w:instrText xml:space="preserve"> HYPERLINK "https://example.org/a" \\l "part2" \\o "tip" </w:instrText></w:r>'
    + '<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>the example</w:t></w:r>'
    + '<w:r><w:fldChar w:fldCharType="end"/></w:r><w:r><w:t>.</w:t></w:r></w:p>' });
  assert.equal(await md(pkg), 'See [the example](https://example.org/a#part2).\n');
  assert.equal(await md(pkg, { textOptions: { field_codes: true } }),
    'See  HYPERLINK "https://example.org/a" \\\\l "part2" \\\\o "tip" [the example](https://example.org/a#part2).\n');
});

test('a simple field\'s HYPERLINK, and a w:hyperlink\'s anchor', async () => {
  const pkg = await docx({ body: '<w:p><w:fldSimple w:instr=" HYPERLINK \\l &quot;intro&quot; "><w:r><w:t>back</w:t></w:r></w:fldSimple>'
    + '<w:r><w:t xml:space="preserve"> and </w:t></w:r><w:hyperlink w:anchor="end"><w:r><w:t>on</w:t></w:r></w:hyperlink></w:p>' });
  assert.equal(await md(pkg), '[back](#intro) and [on](#end)\n');
});

test('hidden text is the text model\'s: left out by default, nothing trimmed (REQ-507, REQ-602)', async () => {
  const pkg = await docx({ body: '<w:p><w:r><w:t xml:space="preserve">shown </w:t></w:r><w:r><w:rPr><w:vanish/></w:rPr><w:t>hidden</w:t></w:r></w:p>' });
  assert.equal(await md(pkg), 'shown \n');
  assert.equal(await md(pkg, { textOptions: { hidden: true } }), 'shown hidden\n');
});

test('footnotes [^n], endnotes [^en]: numbered apart, so labelled apart (REQ-503)', async () => {
  const pkg = await docx();
  const endnotes = new EndnotesPart();
  endnotes.setXml(`<w:endnotes ${W}>`
    + '<w:endnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:endnote>'
    + '<w:endnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:endnote>'
    + '<w:endnote w:id="1"><w:p><w:r><w:endnoteRef/></w:r><w:r><w:t xml:space="preserve"> The end note.</w:t></w:r></w:p></w:endnote></w:endnotes>');
  pkg.getMainDocumentPart().addTargetPart(endnotes);
  await pkg.body.insertXml('<w:p><w:r><w:t>Body</w:t></w:r><w:r><w:endnoteReference w:id="1"/></w:r></w:p>', 'End');
  assert.equal(await md(pkg), 'Body[^e1]\n\n[^e1]: The end note.\n');
});

test('under metadata: true, a YAML front-matter block of the document properties (REQ-704)', async () => {
  const pkg = await docx({ body: '<w:p><w:r><w:t>Body</w:t></w:r></w:p>' });
  pkg.docPropsCorePart.setXml('<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"'
    + ' xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>A "quoted" title</dc:title></cp:coreProperties>');
  const out = await md(pkg, { textOptions: { metadata: true } });
  assert.ok(out.startsWith('---\ncore:\n  title: "A \\"quoted\\" title"\n'), out);
  assert.ok(out.endsWith('---\n\nBody\n'), out);
  assert.equal(await md(pkg), 'Body\n');
});

test('in markup, a comment at its reference: {>>text<<} (REQ-504, REQ-703)', async () => {
  const pkg = await docx();
  const comments = new CommentsPart();
  comments.setXml(`<w:comments ${W}><w:comment w:id="3" w:author="Ann"><w:p><w:r><w:t>Check this</w:t></w:r></w:p></w:comment></w:comments>`);
  pkg.getMainDocumentPart().addTargetPart(comments);
  await pkg.body.insertXml('<w:p><w:commentRangeStart w:id="3"/><w:r><w:t>Claim</w:t></w:r><w:commentRangeEnd w:id="3"/>'
    + '<w:r><w:commentReference w:id="3"/></w:r></w:p>', 'End');
  assert.equal(await md(pkg), 'Claim\n');
  assert.equal(await md(pkg, { view: 'markup' }), 'Claim{>>Check this<<}\n');
});

test('a text box\'s paragraphs once, after its anchoring paragraph, from the branch read for text (REQ-518)', async () => {
  const pkg = await docx({ body: '<w:p><w:r><w:t>Anchor</w:t></w:r><w:r><mc:AlternateContent>'
    + '<mc:Choice Requires="wps"><w:drawing><wp:inline><wp:extent cx="1" cy="1"/><wp:docPr id="1" name="Box"/>'
    + '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">'
    + '<wps:wsp><wps:txbx><w:txbxContent><w:p><w:r><w:t>In the box</w:t></w:r></w:p></w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp>'
    + '</a:graphicData></a:graphic></wp:inline></w:drawing></mc:Choice>'
    + '<mc:Fallback><w:pict><v:shape><v:textbox><w:txbxContent><w:p><w:r><w:t>In the box</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape></w:pict></mc:Fallback>'
    + '</mc:AlternateContent></w:r></w:p><w:p><w:r><w:t>After</w:t></w:r></w:p>' });
  assert.equal(await md(pkg), 'Anchor\n\nIn the box\n\nAfter\n');
});

const tc = (jc, text) => `<w:tc><w:p>${jc === undefined ? '' : `<w:pPr><w:jc w:val="${jc}"/></w:pPr>`}${text === undefined ? '' : `<w:r><w:t>${text}</w:t></w:r>`}</w:p></w:tc>`;

test('the GFM header is the w:tblHeader rows, several of them one row " / " apart; a column aligns as its topmost header cell (REQ-703)', async () => {
  const header = '<w:trPr><w:tblHeader/></w:trPr>';
  let pkg = await docx({ body: `<w:tbl><w:tr>${header}${tc('right', 'A')}${tc(undefined, 'B')}</w:tr>`
    + `<w:tr>${header}${tc(undefined, 'a2')}${tc('center', 'b2')}</w:tr><w:tr>${tc(undefined, '1')}${tc('center', '2')}</w:tr></w:tbl>` });
  assert.equal(await md(pkg), '|A / a2|B / b2|\n|---:|---|\n|1|2|\n');
  pkg = await docx({ body: `<w:tbl><w:tr>${tc('right', undefined)}${tc('center', 'H')}</w:tr><w:tr>${tc(undefined, 'x')}${tc(undefined, '1')}</w:tr></w:tbl>` });
  assert.equal(await md(pkg), '||H|\n|---:|:---:|\n|x|1|\n');
});

const SVG = '<svg xmlns="http://www.w3.org/2000/svg" width="40" height="20"><rect width="40" height="20" fill="red"/></svg>';
const PNG = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';

/** A picture as docx4j's ImageAddSVG sample writes it. */
const picture = (blip, descr = 'A red box') => '<w:p><w:r><w:drawing><wp:inline><wp:extent cx="381000" cy="190500"/>'
  + `<wp:docPr id="2" name="Picture 2" descr="${descr}"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic>`
  + '<pic:nvPicPr><pic:cNvPr id="1" name=""/><pic:cNvPicPr/></pic:nvPicPr>'
  + `<pic:blipFill>${blip}<a:stretch><a:fillRect/></a:stretch></pic:blipFill>`
  + '<pic:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>';

/** A package with one picture of the given image, saved and loaded again (the picture as loaded). */
async function pictured(name, contentType, bytes, svgBlip = false, descr = undefined) {
  const pkg = await docx();
  const image = new ImagePart(`/word/media/${name}`, contentType);
  image.setBytes(bytes);
  const rId = pkg.getMainDocumentPart().addTargetPart(image).id;
  await pkg.body.insertXml(picture(svgBlip
    ? '<a:blip><a:extLst><a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}">'
      + `<asvg:svgBlip xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main" r:embed="${rId}"/></a:ext></a:extLst></a:blip>`
    : `<a:blip r:embed="${rId}"/>`, descr), 'End');
  return WordprocessingMLPackage.load(await pkg.save());
}

test('an SVG picture with no PNG copy is written with its SVG, not dropped (docx4j #699)', async () => {
  const pkg = await pictured('image1.svg', 'image/svg+xml', new TextEncoder().encode(SVG), true);
  const result = await Texts.of(pkg);
  const run = result.blocks.flatMap((b) => b.runs).find((r) => r.kind === 'IMAGE');
  assert.equal(run.image.part.contentType, 'image/svg+xml');
  const out = await md(pkg);
  const prefix = '![A red box](data:image/svg+xml;base64,';
  assert.ok(out.startsWith(prefix), out);
  // the part's bytes as stored (docx4j writes its SVG part re-serialized: CR-012 section 16.4)
  assert.equal(new TextDecoder().decode(Uint8Array.from(atob(out.slice(prefix.length, out.indexOf(')'))), (c) => c.charCodeAt(0))), SVG);
  // docx4j's imageDirPath is a function here: given the bytes and docx4j's file name, it says where the picture points
  const seen = [];
  const linked = await md(pkg, { imageDestination: (image) => { seen.push(image); return image.fileName; } });
  assert.equal(linked, '![A red box](image1.svg)\n');
  assert.equal(seen.length, 1);
  assert.equal(seen[0].index, 1);
  assert.equal(seen[0].contentType, 'image/svg+xml');
  assert.equal(new TextDecoder().decode(seen[0].bytes), SVG);
});

test('a picture whose image can\'t be had (a missing relationship) keeps its alt text', async () => {
  const pkg = await docx({ body: picture('<a:blip r:embed="rIdMissing"/>') });
  assert.equal(await md(pkg), '![A red box]()\n');
});

test('a list\'s numbering is its num instance\'s, w:lvlOverride applied: a start override, an overriding level', async () => {
  const pkg = await docx({ numbering: '<w:abstractNum w:abstractNumId="1"><w:multiLevelType w:val="singleLevel"/>'
    + '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum>'
    + '<w:num w:numId="1"><w:abstractNumId w:val="1"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="5"/></w:lvlOverride></w:num>'
    + '<w:num w:numId="2"><w:abstractNumId w:val="1"/><w:lvlOverride w:ilvl="0">'
    + '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="o"/></w:lvl></w:lvlOverride></w:num>',
  body: '<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>Fifth</w:t></w:r></w:p><w:p/>'
    + '<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr></w:pPr><w:r><w:t>Bulleted</w:t></w:r></w:p>' });
  assert.equal(await md(pkg), '5. Fifth\n\n- Bulleted\n');
});

test('numbering through styles: w:numStyleLink followed, a style\'s w:numPr counts, a level linked to another style numbers no paragraph merely carrying its w:numPr', async () => {
  const pkg = await docx({
    numbering: '<w:abstractNum w:abstractNumId="1"><w:styleLink w:val="MyList"/>'
      + '<w:lvl w:ilvl="0"><w:start w:val="3"/><w:numFmt w:val="upperRoman"/><w:pStyle w:val="Numbered"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum>'
      + '<w:abstractNum w:abstractNumId="2"><w:numStyleLink w:val="MyList"/></w:abstractNum>'
      + '<w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num><w:num w:numId="2"><w:abstractNumId w:val="2"/></w:num>',
    styles: [
      '<w:style w:type="numbering" w:styleId="MyList"><w:name w:val="My List"/><w:pPr><w:numPr><w:numId w:val="1"/></w:numPr></w:pPr></w:style>',
      '<w:style w:type="paragraph" w:styleId="Numbered"><w:name w:val="Numbered"/><w:pPr><w:numPr><w:numId w:val="1"/></w:numPr></w:pPr></w:style>',
      '<w:style w:type="paragraph" w:styleId="Other"><w:name w:val="Other"/><w:pPr><w:numPr><w:numId w:val="1"/></w:numPr></w:pPr></w:style>'],
    body: '<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr></w:pPr><w:r><w:t>Linked</w:t></w:r></w:p>'
      + '<w:p><w:pPr><w:pStyle w:val="Numbered"/></w:pPr><w:r><w:t>Styled</w:t></w:r></w:p>'
      + '<w:p><w:pPr><w:pStyle w:val="Other"/></w:pPr><w:r><w:t>Unnumbered</w:t></w:r></w:p>' });
  // two lists, III. and III.: the second takes the other delimiter, or Markdown would run them together
  assert.equal(await md(pkg), '3. Linked\n\n3) Styled\n\nUnnumbered\n');
  // the list labels agree
  const r = await Texts.of(pkg, { list_labels: true });
  assert.equal(r.blocks[0].listLabel, 'III.\t');
  assert.equal(r.blocks[1].listLabel, 'III.\t');
  assert.equal(r.blocks[2].listLabel, undefined);
  assert.equal(r.blocks[2].numbering, undefined);
});

/** A decimal list, numId 1, levels 0 and 1 at 720 and 1440 (hanging 360); and numId 2, another list. */
function twoLists() {
  let x = '';
  for (let a = 1; a <= 2; a++) {
    x += `<w:abstractNum w:abstractNumId="${a}">`;
    for (let l = 0; l <= 1; l++) {
      x += `<w:lvl w:ilvl="${l}"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%${l + 1}."/><w:lvlJc w:val="left"/>`
        + `<w:pPr><w:ind w:left="${720 + 720 * l}" w:hanging="360"/></w:pPr></w:lvl>`;
    }
    x += '</w:abstractNum>';
  }
  return `${x}<w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num><w:num w:numId="2"><w:abstractNumId w:val="2"/></w:num>`;
}
const item = (numId, ilvl, text) => `<w:p><w:pPr><w:numPr><w:ilvl w:val="${ilvl}"/><w:numId w:val="${numId}"/></w:numPr></w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`;
const para = (left, text) => `<w:p>${left === undefined ? '' : `<w:pPr><w:ind w:left="${left}"/></w:pPr>`}${text === '' ? '' : `<w:r><w:t>${text}</w:t></w:r>`}</w:p>`;

test('a paragraph at the margin between two items ends the list, and the list goes on at Word\'s count where it resumes (REQ-703)', async () => {
  const pkg = await docx({ numbering: twoLists(), body: item(1, 0, 'Apples') + para(undefined, 'A note') + item(1, 0, 'Pears') + item(1, 0, 'Plums') });
  assert.equal(await md(pkg), '1. Apples\n\nA note\n\n2. Pears\n3. Plums\n');
  const r = await Texts.of(pkg);
  assert.equal(r.blocks[0].numbering.value, 1);
  assert.equal(r.blocks[2].numbering.value, 2);
});

test('an unlabelled paragraph indented to the item\'s text continues the item: one list, the count unbroken', async () => {
  const pkg = await docx({ numbering: twoLists(), body: item(1, 0, 'Apples') + para(720, 'A note') + item(1, 0, 'Pears') });
  assert.equal(await md(pkg), '1. Apples\n   \n   A note\n\n2. Pears\n');
});

test('under a nested item, the paragraph continues the deepest item whose text it reaches', async () => {
  const pkg = await docx({ numbering: twoLists(), body: item(1, 0, 'A') + item(1, 1, 'A.1') + para(720, 'about A') + item(1, 1, 'A.2') + item(1, 0, 'B') });
  assert.equal(await md(pkg), '1. A\n   \n   1. A.1\n   \n   about A\n   \n   2. A.2\n\n2. B\n');
});

test('Markdown has no empty paragraphs: an empty one between items neither ends the list nor joins it', async () => {
  const pkg = await docx({ numbering: twoLists(), body: item(1, 0, 'Apples') + para(undefined, '') + item(1, 0, 'Pears') });
  assert.equal(await md(pkg), '1. Apples\n2. Pears\n');
});

test('two lists one after the other stay two: the second takes the other delimiter', async () => {
  const pkg = await docx({ numbering: twoLists(), body: item(1, 0, 'X') + item(2, 0, 'Y') + item(1, 0, 'X again') });
  assert.equal(await md(pkg), '1. X\n\n1) Y\n\n2. X again\n');
});

test('a margin paragraph inside a nested list ends it all; the nested list resumes under an empty item showing its parent\'s count', async () => {
  const pkg = await docx({ numbering: twoLists(), body: item(1, 0, 'A') + item(1, 0, 'B') + item(1, 1, 'B.1') + para(undefined, 'A note') + item(1, 1, 'B.2') });
  assert.equal(await md(pkg), '1. A\n2. B\n   1. B.1\n\nA note\n\n2. 2. B.2\n');
});

test('export(result) writes the reading it is given: the Markdown and the plain text of one reading', async () => {
  const pkg = await docx();
  pkg.body.addStyledParagraphOfText('Heading1', 'Title');
  pkg.body.addParagraphOfText('Body');
  const result = await Texts.of(pkg, { view: 'accepted' });
  assert.equal(await MarkdownExporter.export(result), await md(pkg));
  assert.equal(await MarkdownExporter.export(result), '# Title\n\nBody\n');
  assert.equal(PlainText.of(result), 'Title\nBody\n');
});

// ---------------------------------------------------------------- MarkdownExportTest

test('a heading by its effective outlineLvl, not its style\'s name', async () => {
  const pkg = await docx({ styles: ['<w:style w:type="paragraph" w:styleId="Overskrift2"><w:name w:val="overskrift 2"/><w:pPr><w:outlineLvl w:val="1"/></w:pPr></w:style>'] });
  pkg.body.addStyledParagraphOfText('Overskrift2', 'Localized heading');
  assert.equal(await md(pkg), '## Localized heading\n');
});

test('a heading\'s own bold is its style\'s, not **markers**', async () => {
  const pkg = await docx();
  pkg.body.addStyledParagraphOfText('Heading1', 'The Title');
  assert.equal(await md(pkg), '# The Title\n');
});

test('a field\'s cached result is its text', async () => {
  const pkg = await docx({ body: '<w:p><w:r><w:t xml:space="preserve">Page </w:t></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r>'
    + '<w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>'
    + '<w:r><w:t>7</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r><w:r><w:t xml:space="preserve"> of many.</w:t></w:r></w:p>' });
  assert.equal(await md(pkg), 'Page 7 of many.\n');
});

test('a content control contributes its content', async () => {
  const pkg = await docx({ body: '<w:sdt><w:sdtPr><w:tag w:val="x"/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>Inside the control</w:t></w:r></w:p></w:sdtContent></w:sdt>' });
  assert.equal(await md(pkg), 'Inside the control\n');
});

test('a run in a mono font is inline code, by its effective rPr', async () => {
  const pkg = await docx({ body: '<w:p><w:r><w:t xml:space="preserve">run </w:t></w:r>'
    + '<w:r><w:rPr><w:rFonts w:ascii="Courier New" w:hAnsi="Courier New"/></w:rPr><w:t>mono()</w:t></w:r></w:p>' });
  assert.equal(await md(pkg), 'run `mono()`\n');
});

test('empty paragraphs are dropped', async () => {
  const pkg = await docx();
  pkg.body.addParagraphOfText('first');
  await pkg.body.insertXml('<w:p/>', 'End');
  pkg.body.addParagraphOfText('second');
  assert.equal(await md(pkg), 'first\n\nsecond\n');
});

test('Markdown-significant text is escaped, as commonmark-java escapes it', async () => {
  const pkg = await docx();
  pkg.body.addParagraphOfText('2 * 3 = 6 and #tag [x]');
  // docx4j's test reads it back through its importer; this is what docx4j writes (its MarkdownExporter at 198d015a3)
  assert.equal(await md(pkg), '2 \\* 3 = 6 and #tag \\[x\\]\n');
});

test('a paragraph style whose own font is mono is a fenced code block: one fence, a line each, a w:br a newline, nothing escaped', async () => {
  const pkg = await docx({
    styles: ['<w:style w:type="paragraph" w:styleId="Code"><w:name w:val="Code"/><w:rPr><w:rFonts w:ascii="Consolas" w:hAnsi="Consolas"/><w:sz w:val="18"/></w:rPr></w:style>'],
    body: '<w:p><w:pPr><w:pStyle w:val="Code"/></w:pPr><w:r><w:t xml:space="preserve">&lt;xpaths xmlns="http://opendope.org/xpaths"&gt;</w:t></w:r></w:p>'
      + '<w:p><w:pPr><w:pStyle w:val="Code"/></w:pPr><w:r><w:t xml:space="preserve">  &lt;xpath id="x1"&gt;</w:t><w:br/>'
      + '<w:t xml:space="preserve">    *not emphasis*</w:t></w:r></w:p><w:p><w:r><w:t>after</w:t></w:r></w:p>' });
  assert.equal(await md(pkg), '```\n<xpaths xmlns="http://opendope.org/xpaths">\n  <xpath id="x1">\n    *not emphasis*\n```\n\nafter\n');
});

test('a mono document default is not a code block: a document set in Courier New throughout is prose', async () => {
  const pkg = await docx({ docDefaultsFont: 'Courier New' });
  pkg.body.addParagraphOfText('a * b');
  assert.equal(await md(pkg), 'a \\* b\n');
});

test('a numbering level with numFmt none and a lvlText with no %n is a label, not a list: it leads the paragraph in bold', async () => {
  const pkg = await docx({
    numbering: '<w:abstractNum w:abstractNumId="7"><w:multiLevelType w:val="singleLevel"/>'
      + '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="none"/><w:suff w:val="tab"/>'
      + '<w:lvlText w:val="NOTE"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="1134" w:hanging="1134"/></w:pPr></w:lvl></w:abstractNum>'
      + '<w:abstractNum w:abstractNumId="8"><w:multiLevelType w:val="singleLevel"/>'
      + '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/></w:lvl></w:abstractNum>'
      + '<w:num w:numId="7"><w:abstractNumId w:val="7"/></w:num><w:num w:numId="8"><w:abstractNumId w:val="8"/></w:num>',
    body: '<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="7"/></w:numPr></w:pPr><w:r><w:t>The label is not a number.</w:t></w:r></w:p>'
      + '<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="8"/></w:numPr></w:pPr><w:r><w:t>But this is a list.</w:t></w:r></w:p>' });
  assert.equal(await md(pkg), '**NOTE** The label is not a number.\n\n1. But this is a list.\n');
});

test('a table of contents is left out under toc: false, the default; toc: true keeps it', async () => {
  const pkg = await docx({
    styles: ['<w:style w:type="paragraph" w:styleId="TOC1"><w:name w:val="toc 1"/></w:style>'],
    body: '<w:p><w:pPr><w:pStyle w:val="TOC1"/></w:pPr>'
      + '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h </w:instrText></w:r>'
      + '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
      + '<w:hyperlink w:anchor="_Toc1004"><w:r><w:t>1.1</w:t></w:r><w:r><w:tab/></w:r>'
      + '<w:r><w:t>Overview</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>3</w:t></w:r></w:hyperlink></w:p>'
      + '<w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>' });
  pkg.body.addStyledParagraphOfText('Heading1', 'Overview');
  assert.equal(await md(pkg), '# Overview\n');
  assert.equal(await md(pkg, { textOptions: { toc: true } }), '[1.1\tOverview\t3](#_Toc1004)\n\n# Overview\n');
});

// ---------------------------------------------------------------- MarkdownExportExtensionsTest

test('gridSpan pads the row with empty cells, so the columns stay aligned', async () => {
  const pkg = await docx({ body: '<w:tbl><w:tr><w:tc><w:tcPr><w:gridSpan w:val="2"/></w:tcPr><w:p><w:r><w:t>span</w:t></w:r></w:p></w:tc></w:tr>'
    + `<w:tr>${tc(undefined, 'a')}${tc(undefined, 'b')}</w:tr></w:tbl>` });
  assert.equal(await md(pkg), '|span||\n|---|---|\n|a|b|\n');
});

test('a vMerge continuation is an empty cell', async () => {
  const pkg = await docx({ body: `<w:tbl><w:tr>${tc(undefined, 'h1')}${tc(undefined, 'h2')}</w:tr>`
    + `<w:tr><w:tc><w:tcPr><w:vMerge w:val="restart"/></w:tcPr><w:p><w:r><w:t>top</w:t></w:r></w:p></w:tc>${tc(undefined, 'r1')}</w:tr>`
    + `<w:tr><w:tc><w:tcPr><w:vMerge/></w:tcPr><w:p/></w:tc>${tc(undefined, 'r2')}</w:tr></w:tbl>` });
  assert.equal(await md(pkg), '|h1|h2|\n|---|---|\n|top|r1|\n||r2|\n');
});

test('a cell\'s paragraphs flatten, <br> apart', async () => {
  const pkg = await docx({ body: '<w:tbl><w:tr><w:tc><w:p><w:r><w:t>one</w:t></w:r></w:p><w:p><w:r><w:t>two</w:t></w:r></w:p></w:tc></w:tr></w:tbl>' });
  assert.equal(await md(pkg), '|one<br>two|\n|---|\n');
});

const tracked = `<w:p><w:r><w:t xml:space="preserve">keep </w:t></w:r><w:ins w:id="1"${DATE}><w:r><w:t>added</w:t></w:r></w:ins>`
  + `<w:del w:id="2"${DATE}><w:r><w:delText xml:space="preserve"> removed</w:delText></w:r></w:del></w:p>`;

test('tracked changes: the accepted view by default; in markup CriticMarkup, not ~~ (REQ-703); the original view', async () => {
  const pkg = await docx({ body: tracked });
  assert.equal(await md(pkg), 'keep added\n');
  assert.equal(await md(pkg, { view: 'markup' }), 'keep {++added++}{-- removed--}\n');
  assert.equal(await md(pkg, MarkdownExportOptions.of({ textOptions: { view: 'markup' } })), 'keep {++added++}{-- removed--}\n');
  assert.equal(await md(pkg, { view: 'original' }), 'keep  removed\n');
  assert.equal(await md(pkg, { view: 'markup', textOptions: { markup_detail: 'author' } }), 'keep {++a: added++}{--a:  removed--}\n');
});

test('a picture is a data URI by default, or where the destination function says', async () => {
  const bytes = Uint8Array.from(atob(PNG), (c) => c.charCodeAt(0));
  const pkg = await pictured('image1.png', 'image/png', bytes, false, 'tiny');
  assert.equal(await md(pkg), `![tiny](data:image/png;base64,${PNG})\n`);
  const linked = await md(pkg, { imageDestination: (image) => `images/${image.fileName}` });
  assert.equal(linked, '![tiny](images/image1.png)\n');
});

// ---------------------------------------------------------------- MarkdownAlternateContentTest

test('a block-level mc:AlternateContent gives the branch the text model reads, as built and as loaded either way (docx4j CR-051)', async () => {
  const p = (text) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
  const built = await docx();
  built.getMainDocumentPart().setXml(`<w:document ${W} xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"`
    + ` xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body>${p('before')}`
    + `<mc:AlternateContent><mc:Choice Requires="w14">${p('CHOICEW14')}</mc:Choice><mc:Fallback>${p('FALLBACKTEXT')}</mc:Fallback></mc:AlternateContent>`
    + `${p('after')}</w:body></w:document>`);
  const bytes = await built.save();
  for (const pkg of [built, await WordprocessingMLPackage.load(bytes), await WordprocessingMLPackage.load(bytes, { mcePreprocess: true })]) {
    const out = await md(pkg);
    assert.ok(out.includes('CHOICEW14'), out);
    assert.ok(!out.includes('FALLBACKTEXT'), out);
  }
});

// ---------------------------------------------------------------- this package's own

test('block quotes: the Quote styles, nested by the indent beyond the quote style\'s own; by the indent alone with no package to ask', async () => {
  const pkg = await docx({
    styles: ['<w:style w:type="paragraph" w:styleId="Quote"><w:name w:val="Quote"/><w:pPr><w:ind w:left="864" w:right="864"/></w:pPr></w:style>'],
    body: '<w:p><w:pPr><w:pStyle w:val="Quote"/></w:pPr><w:r><w:t>Said once.</w:t></w:r></w:p>'
      + '<w:p><w:pPr><w:pStyle w:val="Quote"/></w:pPr><w:r><w:t>And again.</w:t></w:r></w:p>'
      + '<w:p><w:pPr><w:pStyle w:val="Quote"/><w:ind w:left="1584"/></w:pPr><w:r><w:t>Deeper.</w:t></w:r></w:p>'
      + '<w:p><w:r><w:t>After.</w:t></w:r></w:p>' });
  // docx4j's MarkdownExporter at 198d015a3 writes the same, with and without the package
  assert.equal(await md(pkg), '> Said once.\n> \n> And again.\n\n> > Deeper.\n\nAfter.\n');
  const result = await Texts.of(pkg);
  assert.equal(await MarkdownExporter.export(result, pkg), await md(pkg));
  assert.equal(await MarkdownExporter.export(result), '> > Said once.\n> > \n> > And again.\n\n> > > Deeper.\n\nAfter.\n');
});

test('a thematic break: an empty paragraph with a bottom border of its own', async () => {
  const pkg = await docx({ body: '<w:p><w:r><w:t>Above</w:t></w:r></w:p><w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="auto"/></w:pBdr></w:pPr></w:p>'
    + '<w:p><w:r><w:t>Below</w:t></w:r></w:p>' });
  assert.equal(await md(pkg), 'Above\n\n---\n\nBelow\n');
});

test('emphasis: bold, italic, both, struck through; a line break is a hard break', async () => {
  const pkg = await docx({ body: '<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>bold</w:t></w:r><w:r><w:t xml:space="preserve"> </w:t></w:r>'
    + '<w:r><w:rPr><w:i/></w:rPr><w:t>italic</w:t></w:r><w:r><w:t xml:space="preserve"> </w:t></w:r>'
    + '<w:r><w:rPr><w:b/><w:i/></w:rPr><w:t>both</w:t></w:r><w:r><w:t xml:space="preserve"> </w:t></w:r>'
    + '<w:r><w:rPr><w:strike/></w:rPr><w:t>gone</w:t></w:r><w:r><w:br/><w:t>next line</w:t></w:r></w:p>' });
  assert.equal(await md(pkg), '**bold** *italic* ***both*** ~~gone~~  \nnext line\n');
});

test('blocks(): each block\'s Markdown, whole lines with the blank line after it; joined, they are export\'s; the front matter apart', async () => {
  const pkg = await docx({ numbering: twoLists(), body: item(1, 0, 'Apples') + para(720, 'A note') + para(undefined, '') + item(1, 0, 'Pears')
    + '<w:tbl><w:tr><w:tc><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>' });
  pkg.body.addStyledParagraphOfText('Heading1', 'Last');
  pkg.docPropsCorePart.setXml('<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"'
    + ' xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>Fruit</dc:title></cp:coreProperties>');
  const options = { textOptions: { metadata: true } };
  const { frontMatter, blocks, markdown } = await new MarkdownExporter(options).blocks(pkg);
  assert.equal(markdown, await md(pkg, options));
  assert.equal(frontMatter, '---\ncore:\n  title: "Fruit"\n---\n\n');
  assert.equal(frontMatter + blocks.map((b) => b.markdown).join(''), markdown);
  assert.deepEqual(blocks.map((b) => [b.block.address, b.markdown]), [
    ['body/0', '1. Apples\n   \n'],
    ['body/1', '   A note\n\n'],
    ['body/2', ''],                                                     // an empty paragraph writes nothing
    ['body/3', '2. Pears\n\n'],
    ['body/4', '|cell|\n|---|\n\n'],
    ['body/5', '# Last\n'],
  ]);
  // a result, as export takes one; the static form has the default options
  const result = await Texts.of(pkg);
  assert.equal((await MarkdownExporter.blocks(result, pkg)).frontMatter, '');
  assert.equal((await MarkdownExporter.blocks(result, pkg)).markdown, await md(pkg));
});

test('blocks(): a code fence is its first paragraph\'s, a note\'s definition its first block\'s', async () => {
  const pkg = await docx({
    styles: ['<w:style w:type="paragraph" w:styleId="SourceCode"><w:name w:val="Source Code"/></w:style>'],
    body: '<w:p><w:pPr><w:pStyle w:val="SourceCode"/></w:pPr><w:r><w:t>one();</w:t></w:r></w:p>'
      + '<w:p><w:pPr><w:pStyle w:val="SourceCode"/></w:pPr><w:r><w:t>two();</w:t></w:r></w:p><w:p><w:r><w:t>Prose</w:t></w:r></w:p>' });
  const { blocks } = await MarkdownExporter.blocks(pkg);
  assert.deepEqual(blocks.map((b) => b.markdown), ['```\none();\ntwo();\n```\n\n', '', 'Prose\n']);
});

test('REQ-204: a Markdown export changes nothing, and the probes save byte for byte as before', async () => {
  for (const name of ['extraction-probe.docx', 'extraction-probe-structure.docx', 'extraction-probe-external.docm']) {
    const bytes = readFileSync(join(phase0, name));
    const before = unzipSync(await (await WordprocessingMLPackage.load(bytes)).save());
    const pkg = await WordprocessingMLPackage.load(bytes);
    for (const view of ['accepted', 'original', 'markup']) {
      await md(pkg, { view, textOptions: { metadata: true, include: ['headers', 'footers', 'comments', 'glossary'], hidden: true, list_labels: true, field_codes: true } });
    }
    for (const part of pkg.parts) if (part instanceof XmlPart && !part.partName.name.endsWith('.rels')) assert.equal(part.isUnmarshalled, false, `${name} ${part.partName.name}`);
    const after = unzipSync(await pkg.save());
    assert.deepEqual(Object.keys(after).sort(), Object.keys(before).sort());
    for (const entryName of Object.keys(before)) assert.deepEqual(after[entryName], before[entryName], `${name} ${entryName}`);
  }
});

test('MarkdownExportOptions: the view is the text options\'; options are taken as given or made from their init', () => {
  const o = MarkdownExportOptions.of({ view: 'markup', textOptions: TextOptions.of({ hidden: true }) });
  assert.equal(o.view, 'markup');
  assert.equal(o.textOptions.view, 'markup');
  assert.equal(o.textOptions.hidden, true);
  o.view = 'original';
  assert.equal(o.textOptions.view, 'original');
  assert.equal(MarkdownExportOptions.of(o), o);
  assert.equal(new MarkdownExporter(o).options, o);
  assert.equal(new MarkdownExportOptions().textOptions.view, 'accepted');
});

// ------------------------------------------------------------------ CR-012 section 21: docx4j CR-056's links

// docx4j-core-tests LinkTargetsTest (84ee9741c), its property an argument here
test('LinkTargets: the allowlist, the scheme read as a browser reads it, schemes added, the excerpt (docx4j LinkTargetsTest)', () => {
  for (const t of [undefined, '', 'http://example.org/a', 'HTTPS://example.org/b', 'mailto:a@example.org',
    'ftp://example.org/c', 'file:///C:/docs/d.docx', 'tel:+61000000', 'news:comp.lang.java', 'sub/dir/page.html',
    '#bookmark1', 'C:\\docs\\e.docx', 'c:/docs/e.docx', '\\\\server\\share\\f.docx', 'page.html?x=javascript:1',
    'page.html#javascript:1', '/root/path:with:colons', '  http://example.org/leading-spaces']) {
    assert.ok(LinkTargets.isSafe(t), String(t));
  }
  for (const t of ['javascript:alert(1)', 'JavaScript:alert(1)', ' \u0001java\tscri\npt:alert(2)',
    '\r\njavascript:alert(3)', 'vbscript:MsgBox(1)', 'data:text/html,<script>alert(1)</script>', 'onenote:///x',
    'ms-word:ofe|u|https://example.org/x.docx', 'C:', 'java script:alert(1)', 'about:blank', 'blob:x']) {
    assert.ok(!LinkTargets.isSafe(t), t);
  }
  assert.equal(LinkTargets.scheme(' \u0001java\tscri\npt:alert(2)'), 'javascript');
  assert.equal(LinkTargets.scheme('HTTP://x'), 'http');
  assert.equal(LinkTargets.scheme('sub/dir:x'), undefined);
  assert.equal(LinkTargets.scheme('#frag:x'), undefined);
  assert.equal(LinkTargets.scheme(':x'), undefined);
  assert.equal(LinkTargets.scheme('C:\\x'), 'c');
  // the property's test: docx4j's comma-separated value, as the list it splits to
  const extra = ' OneNote:, ms-word ,,notes'.split(',');
  for (const scheme of LinkTargets.defaultSchemes()) assert.ok(LinkTargets.schemes(extra).includes(scheme));
  assert.ok(LinkTargets.schemes(extra).includes('onenote'));
  assert.ok(LinkTargets.isSafe('onenote:///x', extra));
  assert.ok(LinkTargets.isSafe('ms-word:ofe|u|https://example.org/x.docx', extra));
  assert.ok(LinkTargets.isSafe('Notes://server/db.nsf', extra));
  assert.ok(!LinkTargets.isSafe('javascript:alert(1)', extra), 'a default is never removed');
  assert.ok(!LinkTargets.isSafe('onenote:///x'), 'only where given');
  assert.equal(LinkTargets.excerpt('java\tscri\npt:alert(1)'), 'java scri pt:alert(1)');
  assert.equal(LinkTargets.excerpt('x'.repeat(200)).length, 83);
});

/** docx4j-markdown MarkdownHostileLinksTest's document: a paragraph per target, a w:hyperlink to it. */
async function hostileLinks() {
  const pkg = await WordprocessingMLPackage.createPackage();
  const body = await pkg.getBody();
  const main = pkg.getMainDocumentPart();
  const rels = main.relationshipsPart ?? main.getRelationshipsPart(true);
  const escape = (text) => text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  // Java's String.trim drops every character up to U+0020 at either end, U+0001 too
  const javaTrim = (text) => text.replace(/^[\u0000-\u0020]+|[\u0000-\u0020]+$/g, '');
  for (const target of ['javascript:alert(1)', ' \u0001java\tscri\npt:alert(2)', 'vbscript:MsgBox(3)', 'https://example.org/ok', '#bookmark1', 'sub/page.html']) {
    const rel = rels.addExternalRelationship(Namespaces.HYPERLINK, target);
    await body.insertXml(`<w:p><w:hyperlink r:id="${rel.id}"><w:r><w:t xml:space="preserve">${escape(`text for ${javaTrim(target)}`)}</w:t></w:r></w:hyperlink></w:p>`, 'End');
  }
  return pkg;
}

test('the Markdown writes no link whose target is a script scheme: the text, once logged; a benign link kept (docx4j MarkdownHostileLinksTest)', async () => {
  const logged = [];
  setLogger({ warn: (message) => logged.push(message) });
  let md;
  let allowed;
  try {
    md = await MarkdownExporter.export(await hostileLinks());
    allowed = await new MarkdownExporter({ linkSchemes: ['vbscript'] }).export(await hostileLinks());
  } finally {
    setLogger();
  }
  assert.ok(!(md.includes('(javascript:') || md.includes('(<javascript:')), `javascript: destination\n${md}`);
  assert.ok(!(md.includes('](') && md.includes('scri\npt:alert(2)>)')), `sneaky javascript: destination\n${md}`);
  assert.ok(!(md.includes('(vbscript:') || md.includes('(<vbscript:')), `vbscript: destination\n${md}`);
  assert.ok(md.includes('text for javascript:alert(1)'), `the text kept\n${md}`);
  assert.ok(!md.includes('[text for javascript:alert(1)]'), `a link around the refused text\n${md}`);
  // benignLinksKept
  assert.ok(md.includes('(https://example.org/ok)') || md.includes('(<https://example.org/ok>)'), md);
  assert.ok(md.includes('(#bookmark1)') || md.includes('(<#bookmark1>)'), md);
  assert.ok(md.includes('(sub/page.html)') || md.includes('(<sub/page.html>)'), md);
  // one warning for each document, naming its first refused target (docx4j WARNs once, DEBUG after)
  assert.deepEqual(logged, Array(2).fill('Link target refused (scheme not in MarkdownExportOptions.linkSchemes): the text is written without a link: javascript:alert(1)'));
  // a scheme the caller adds is written as a link; a default refusal stays
  assert.ok(!allowed.includes('(javascript:') && !allowed.includes('(<javascript:'), allowed);
  assert.ok(allowed.includes('(vbscript:MsgBox(3))') || allowed.includes('(<vbscript:MsgBox(3)>)'), allowed);
});

