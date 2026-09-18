// CR-001 Phase B, step 0: every parity fixture loads, unmarshals fully and round-trips untouched
// byte for byte. The parity test proper (test/parity.test.mjs, step 1) compares against the Java
// goldens under test/golden/.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readdir } from 'node:fs/promises';
import { join } from 'node:path';
import { WordprocessingMLPackage, ZipPartStore } from '../dist/index.mjs';
import { fixture, fixturesDir, bytesEqual } from './helpers.mjs';

const dir = join(fixturesDir, 'parity');
const names = (await readdir(dir)).filter((n) => n.endsWith('.docx')).sort();

test('parity fixtures: present', () => {
  assert.ok(names.length >= 37, `${names.length} fixtures`);
});

for (const name of names) {
  test(`parity fixture ${name}: loads, unmarshals, round-trips untouched`, async () => {
    const bytes = await fixture(join('parity', name));
    const pkg = await WordprocessingMLPackage.load(bytes);
    await pkg.unmarshalAll();
    const body = await pkg.getBody();
    assert.ok(body.paragraphs.length >= 0);
    const again = await WordprocessingMLPackage.load(bytes);   // untouched: no part unmarshalled
    const saved = await again.save();
    const src = new ZipPartStore(bytes), out = new ZipPartStore(saved);
    for (const part of src.partNames()) {
      if (part === '[Content_Types].xml' || part.endsWith('.rels')) continue;   // regenerated and re-marshalled by design (CR-001 section 12)
      assert.ok(out.has(part), `${part} saved`);
      assert.ok(bytesEqual(await src.load(part), await out.load(part)), `${part} byte-identical`);
    }
  });
}
