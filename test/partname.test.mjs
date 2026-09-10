import { test } from 'node:test';
import assert from 'node:assert/strict';
import { PartName, ContentTypeManager, ContentTypes, InvalidFormatException } from '../dist/index.mjs';

test('PartName: validation and accessors', () => {
  const pn = PartName.of('/word/document.xml');
  assert.equal(pn.name, '/word/document.xml');
  assert.equal(pn.extension, 'xml');
  assert.equal(pn.directory, '/word');
  assert.equal(pn.fileName, 'document.xml');
  assert.equal(pn.storeName, 'word/document.xml');
  assert.equal(PartName.of('word/document.xml').name, '/word/document.xml');
  assert.ok(pn.equals('/WORD/Document.XML'));
  assert.equal(PartName.of('/[Content_Types].xml').directory, '/');
  assert.ok(PartName.of('/_rels/.rels').isRelationshipsPart);
  assert.ok(PartName.of('/word/_rels/document.xml.rels').isRelationshipsPart);
  assert.ok(!pn.isRelationshipsPart);
  assert.throws(() => PartName.of('/word/'), InvalidFormatException);
  assert.throws(() => PartName.of('/word//document.xml'), InvalidFormatException);
  assert.throws(() => PartName.of('/word/document.'), InvalidFormatException);
  assert.throws(() => PartName.of('http://example.com/a.xml'), InvalidFormatException);
  assert.throws(() => PartName.of('/word/my doc.xml'), InvalidFormatException);
  assert.equal(PartName.of('/word/media/image%201.png').name, '/word/media/image%201.png');
});

test('PartName: rels names, resolve and relativize', () => {
  assert.equal(PartName.relsFor('/word/document.xml').name, '/word/_rels/document.xml.rels');
  assert.equal(PartName.relsFor(PartName.ROOT).name, '/_rels/.rels');
  assert.equal(PartName.relsFor('/[Content_Types].xml').name, '/_rels/[Content_Types].xml.rels');
  assert.equal(PartName.sourceOfRels('/word/_rels/document.xml.rels').name, '/word/document.xml');
  assert.equal(PartName.sourceOfRels('/_rels/.rels'), PartName.ROOT);

  assert.equal(PartName.resolve(PartName.ROOT, 'word/document.xml').name, '/word/document.xml');
  assert.equal(PartName.resolve('/word/document.xml', 'media/image1.png').name, '/word/media/image1.png');
  assert.equal(PartName.resolve('/word/document.xml', '../customXml/item1.xml').name, '/customXml/item1.xml');
  assert.equal(PartName.resolve('/word/document.xml', '/word/styles.xml').name, '/word/styles.xml');
  assert.equal(PartName.resolve('/ppt/slides/slide1.xml', '../slideLayouts/slideLayout1.xml').name, '/ppt/slideLayouts/slideLayout1.xml');
  assert.equal(PartName.resolve('/word/document.xml', './settings.xml').name, '/word/settings.xml');

  assert.equal(PartName.relativize(PartName.ROOT, '/word/document.xml'), 'word/document.xml');
  assert.equal(PartName.relativize('/word/document.xml', '/word/media/image1.png'), 'media/image1.png');
  assert.equal(PartName.relativize('/word/document.xml', '/customXml/item1.xml'), '../customXml/item1.xml');
  assert.equal(PartName.relativize('/word/document.xml', '/word/styles.xml'), 'styles.xml');
  assert.equal(PartName.relativize('/ppt/slides/slide1.xml', '/ppt/slideLayouts/slideLayout1.xml'), '../slideLayouts/slideLayout1.xml');
  // round trip
  for (const [s, t] of [['/word/document.xml', '/word/media/image1.png'], ['/a/b/c.xml', '/d/e.xml'], ['/', '/word/document.xml']]) {
    assert.equal(PartName.resolve(s, PartName.relativize(s, t)).name, t);
  }
});

test('ContentTypeManager: parse, lookup, write', () => {
  const xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="PNG" ContentType="image/png"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>';
  const ctm = ContentTypeManager.parse(xml);
  assert.equal(ctm.getContentType('/word/document.xml'), ContentTypes.WORDPROCESSINGML_DOCUMENT);
  assert.equal(ctm.getContentType('/Word/Document.xml'), ContentTypes.WORDPROCESSINGML_DOCUMENT);
  assert.equal(ctm.getContentType('/word/styles.xml'), ContentTypes.APPLICATION_XML);
  assert.equal(ctm.getContentType('/word/media/image1.png'), ContentTypes.IMAGE_PNG);
  assert.equal(ctm.getContentType('/word/media/image1.jpeg'), undefined);
  ctm.addContentType('/word/media/image2.png', ContentTypes.IMAGE_PNG); // covered by the default: no override
  assert.equal(ctm.getOverrideContentType('/word/media/image2.png'), undefined);
  ctm.addContentType('/word/styles.xml', ContentTypes.WORDPROCESSINGML_STYLES);
  const out = ctm.toXml();
  assert.ok(out.startsWith('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns='));
  assert.ok(out.includes('<Default Extension="png" ContentType="image/png"/>'));
  assert.ok(out.includes('<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'));
  assert.deepEqual(ContentTypeManager.parse(out).getContentType('/word/styles.xml'), ContentTypes.WORDPROCESSINGML_STYLES);
  ctm.removeContentType('/word/styles.xml');
  assert.equal(ctm.getContentType('/word/styles.xml'), ContentTypes.APPLICATION_XML);
});
