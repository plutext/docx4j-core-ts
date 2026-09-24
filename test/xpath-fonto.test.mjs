// CR-005 phase A: the FontoXPath engine and the three boolean conversion modes of the OpenDoPE
// specification v3, section 7.2 table 7. The table's own examples are the first test; the second
// is docx4j's invoice_Saxon_XPath2.docx, whose conditions are the discriminating cases (a data
// element holding "false", and an expression in XPath 2.0 syntax that XPath 1.0 cannot evaluate).
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { booleanValue, DefaultXPathEngine, parseXml, parsePrefixMappings, WordprocessingMLPackage } from '../dist/index.mjs';
import { FontoXPathEngine } from '../dist/model/customxml/xpath-fonto.mjs';
import { fixture } from './helpers.mjs';

const DATA = parseXml(
  '<invoice><t>true</t><T>True</T><f>false</f><one>1</one><zero>0</zero><ws> false </ws>'
  + '<yes>yes</yes><empty/><two>true</two><two>false</two><amt>1500</amt></invoice>',
);

async function engines() {
  const fonto = new FontoXPathEngine();
  const dflt = new DefaultXPathEngine();
  await fonto.ready();
  await dflt.ready();
  return { fonto, dflt };
}

const verdict = (engine, expression, mode) => {
  try {
    return booleanValue(engine, expression, DATA, {}, mode);
  } catch {
    return 'error';
  }
};

test('table 7: the three modes over the specification\'s own examples', async () => {
  const { fonto, dflt } = await engines();
  // | expression        | java  | xpath1 | xpath2  |
  //
  // The xpath2 column is docx4j's, measured on Saxon-HE 9.9.0-2 with
  // opendope.conditions.Xpathref.XPathBoolean=cast2 (the docx4j session, 2026-09-25): this engine
  // wraps the expression in xs:boolean(...) exactly as XmlPart.cachedXPathGetBoolean does, so the
  // two answer alike by construction rather than by agreement.
  const table = [
    ['/invoice/t', true, true, true],
    ['/invoice/T', true, true, 'error'],      // java ignores case; xs:boolean's lexical space does not
    ['/invoice/f', false, true, false],       // "false" is true in XPath 1.0, false in 2.0
    ['/invoice/one', false, true, true],      // "1" is false in the Java mode
    ['/invoice/zero', false, true, false],
    ['/invoice/ws', false, true, false],      // whitespace is collapsed before the cast
    ['/invoice/yes', false, true, 'error'],   // "yes" cannot be cast to xs:boolean
    ['/invoice/empty', false, true, 'error'], // "" cannot be cast either: an empty element is not false
    ['/invoice/missing', false, false, false], // but an expression selecting nothing is
    ['/invoice/two', true, true, 'error'],    // the java mode takes the first item's string value; the cast will not
    ['/invoice/amt > 1000', true, true, true], // already boolean: every mode agrees
  ];
  for (const [expression, java, xpath1, xpath2] of table) {
    assert.equal(verdict(fonto, expression, 'java'), java, `${expression} java`);
    assert.equal(verdict(fonto, expression, 'xpath1'), xpath1, `${expression} xpath1`);
    assert.equal(verdict(fonto, expression, 'xpath2'), xpath2, `${expression} xpath2`);
    // the default engine answers the two modes XPath 1.0 can, and the same way
    assert.equal(verdict(dflt, expression, 'java'), java, `${expression} java, default engine`);
    assert.equal(verdict(dflt, expression, 'xpath1'), xpath1, `${expression} xpath1, default engine`);
  }
});

test('the default engine refuses xpath2 by naming the engine to install (REQ-032)', async () => {
  const { dflt } = await engines();
  assert.equal(dflt.booleanValue, undefined, 'the optional member is absent, so an older implementation still compiles');
  assert.throws(() => booleanValue(dflt, '/invoice/f', DATA, {}, 'xpath2'), (e) => {
    assert.match(e.message, /xpath2/);
    assert.match(e.message, /@docx4j\/core-ts\/xpath-fonto/);
    assert.match(e.message, /fontoxpath/);
    return true;
  });
});

test('booleanValue defaults to the java mode, as a template that declares none does', async () => {
  const { fonto, dflt } = await engines();
  assert.equal(booleanValue(dflt, '/invoice/f', DATA), false);
  assert.equal(booleanValue(fonto, '/invoice/f', DATA), false);
});

test('select and selectValue answer as the default engine does', async () => {
  const { fonto, dflt } = await engines();
  for (const expression of ['/invoice/f', '/invoice/items']) {
    assert.deepEqual(fonto.select(expression, DATA).map((n) => n.nodeName),
      dflt.select(expression, DATA).map((n) => n.nodeName), expression);
  }
  assert.equal(fonto.selectValue('string(/invoice/T)', DATA), 'True');
  assert.equal(fonto.selectValue('count(/invoice/*)', DATA), 11);
  assert.equal(fonto.selectValue('/invoice/amt > 1000', DATA), true);
  assert.equal(fonto.selectValue('/invoice/missing', DATA), undefined);
  // a selected element reports its text, as the default engine does
  assert.equal(fonto.selectValue('/invoice/t', DATA), dflt.selectValue('/invoice/t', DATA));
});

test('an engine that is not ready says how to ready it', () => {
  const fonto = new FontoXPathEngine();
  assert.equal(fonto.isReady, false);
  assert.throws(() => fonto.select('/invoice', DATA), /not ready/);
});

test('docx4j\'s invoice_Saxon_XPath2.docx: the conditions an XPath 1.0 engine cannot answer', async () => {
  const pkg = await WordprocessingMLPackage.load(await fixture('invoice_Saxon_XPath2.docx'));
  const fonto = new FontoXPathEngine();
  pkg.xpathEngine = fonto;
  await pkg.customXmlParts.load();

  const data = pkg.customXmlParts.getItem('{8b049945-9dfe-4726-9de9-cf5691e53858}');
  assert.ok(data, 'the invoice data part, by the item id its XPaths name');
  const doc = data.document;

  // x5 and x6 are the template's two conditions over plain data elements. wantspam holds "false",
  // which is the case where the modes disagree: XPath 1.0 says a non-empty string is true.
  assert.equal(booleanValue(fonto, '/invoice[1]/misc/includeBankDetails', doc, {}, 'xpath2'), true);
  assert.equal(booleanValue(fonto, '/invoice[1]/misc/wantspam', doc, {}, 'xpath2'), false);
  assert.equal(booleanValue(fonto, '/invoice[1]/misc/wantspam', doc, {}, 'xpath1'), true);
  assert.equal(booleanValue(fonto, '/invoice[1]/misc/wantspam', doc, {}, 'java'), false);

  // dateGt is the reason the template is named for Saxon: XPath 2.0 syntax, which an XPath 1.0
  // engine cannot evaluate at all, in any mode.
  const dateGt = "xs:date(/invoice/date) > xs:date('2018-12-31')";
  const namespaces = parsePrefixMappings("xmlns:xs='http://www.w3.org/2001/XMLSchema'");
  assert.deepEqual(namespaces, { xs: 'http://www.w3.org/2001/XMLSchema' });
  assert.equal(booleanValue(fonto, dateGt, doc, namespaces, 'xpath2'), true);

  const dflt = new DefaultXPathEngine();
  await dflt.ready();
  assert.throws(() => booleanValue(dflt, dateGt, doc, namespaces, 'xpath1'),
    'XPath 1.0 cannot evaluate an xs:date comparison');
});

// A defect in the optional peer `xpath` 0.0.34, recorded as an assertion of the broken behaviour
// rather than a skip, so its fix announces itself (CR-001 section 19). In Node the default engine
// is that package, and it matches element names **case-insensitively**: `/invoice/t` answers both
// <t> and <T>, where the DOM holds them as distinct elements. A browser's document.evaluate is
// correct, and so is FontoXPath, so the same template binds differently in Node and in an add-in.
// Reported for CR-005; when the package is fixed this test fails and the assertion becomes
// equality with FontoXPath's answer.
test('KNOWN, xpath 0.0.34: the default engine matches element names case-insensitively in Node', async () => {
  const { fonto, dflt } = await engines();
  assert.deepEqual(Array.from(DATA.documentElement.childNodes).filter((n) => n.nodeType === 1)
    .map((n) => n.nodeName).filter((n) => n.toLowerCase() === 't'), ['t', 'T'],
    'the DOM holds the two elements distinctly');

  assert.deepEqual(fonto.select('/invoice/t', DATA).map((n) => n.textContent), ['true'], 'FontoXPath is right');
  if (dflt.isReady && typeof document === 'undefined') {
    assert.deepEqual(dflt.select('/invoice/t', DATA).map((n) => n.textContent), ['true', 'True'],
      'the xpath package answers both: when this fails, the package is fixed - assert equality with FontoXPath instead');
  }
});
