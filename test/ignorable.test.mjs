// Every prefix an mc:Ignorable names must be declared on that root: Word and Excel repair a
// document otherwise (CR-001 section 17.3; docx4j CR-023 found the same with `cr` on
// commentsExtensible.xml, which nothing but Word or this check catches). Every XML part of every
// fixture is unmarshalled and re-marshalled, and each root's list is checked prefix by prefix.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readdir } from 'node:fs/promises';
import { join } from 'node:path';
import { OpcPackage, ZipPartStore } from '../dist/index.mjs';
import { fixture, fixturesDir } from './helpers.mjs';

const names = [
  ...(await readdir(fixturesDir)).filter((n) => /\.(docx|pptx|xlsx)$/.test(n)),
  ...(await readdir(join(fixturesDir, 'parity'))).filter((n) => n.endsWith('.docx')).map((n) => join('parity', n)),
].sort();

function undeclared(xml) {
  const root = xml.match(/<[^?!][^>]*>/)?.[0] ?? '';
  const ignorable = root.match(/mc:Ignorable="([^"]*)"/)?.[1];
  if (!ignorable) return [];
  const declared = new Set([...root.matchAll(/xmlns:([\w.-]+)=/g)].map((m) => m[1]));
  return ignorable.split(/\s+/).filter((p) => p && !declared.has(p));
}

/**
 * Parts that cannot be unmarshalled at all, with the reason, so that forcing a read of every part
 * does not hide one behind another. Unmarshalling is not what these fixtures are for; a part here
 * still round-trips from its source bytes, which `roundtrip.test.mjs` asserts.
 */
const UNREADABLE = {
  // The a14 Choice is taken (a14 is understood), and inside it a:graphicData holds the slicer's
  // sle:slicer, for which the model has no module. CT_GraphicalObjectData's wildcard is
  // allowDom: false, so an unknown graphic is fatal rather than kept as DOM, where JAXB's
  // @XmlAnyElement(lax=true) keeps it. Reported upstream (CR-004 section 5); until the schema is
  // lax (docx4j 8e8f6ea83 on VERSION_17_2_1 does exactly that, awaiting an objects release; when
  // it arrives this entry goes and the part must read) a workbook's drawing cannot be
  // unmarshalled if it frames a slicer through an understood
  // Choice. drawing2.xml frames one too but its Choice requires sle15, which is not understood,
  // so the preprocessor takes its Fallback picture and the part reads: what saves it is the
  // branch being given up, not anything about the slicer.
  'cr022-slicers-timelines.xlsx': ['/xl/drawings/drawing1.xml'],
};

for (const name of names) {
  test(`mc:Ignorable prefixes declared after re-marshalling every part: ${name}`, async () => {
    const pkg = await OpcPackage.load(await fixture(name));
    const unreadable = UNREADABLE[name] ?? [];
    for (const part of pkg.parts.values()) {
      if (typeof part.getContents !== 'function') continue;
      if (unreadable.includes(part.partName.name)) {
        await assert.rejects(() => part.getContents(), `${part.partName.name} is listed as unreadable but unmarshalled`);
        continue;
      }
      await part.getContents();
    }
    const out = new ZipPartStore(await pkg.save());
    const failures = [];
    for (const part of out.partNames()) {
      if (!part.endsWith('.xml') || part.endsWith('.rels') || part === '[Content_Types].xml') continue;
      const missing = undeclared(new TextDecoder().decode(await out.load(part)));
      if (missing.length) failures.push(`${part}: ${missing.join(' ')}`);
    }
    assert.deepEqual(failures, [], failures.join('\n'));
  });
}
