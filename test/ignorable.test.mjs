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

for (const name of names) {
  test(`mc:Ignorable prefixes declared after re-marshalling every part: ${name}`, async () => {
    const pkg = await OpcPackage.load(await fixture(name));
    await pkg.unmarshalAll();
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
