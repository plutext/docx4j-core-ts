// CR-011 section 4: every story's tracked changes, and what each resolve-all leaves, against docx4j's
// answers. test/golden/revisions/ is written by test/java's RevisionsHarness (docx4j CR-043's
// TrackedChanges.in, acceptAll and rejectAll, with and without a TrackedChangeFilter) for every fixture
// whose word/ parts hold revision markup, a golden's path mirroring the fixture's. Parity is zero
// differences (CLAUDE.md): a difference is investigated, and fixed in the port unless docx4j is
// plainly wrong.
//
// A w:date without w16du:dateUtc is local wall-clock time (CR-002 section 29), read in the process's
// time zone on both sides; the harness ran in UTC (each header's timeZone), and so does this.
process.env.TZ = 'UTC';

import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFile, readdir } from 'node:fs/promises';
import { dirname, join, relative } from 'node:path';
import { fileURLToPath } from 'node:url';
import { WordprocessingMLPackage } from '../dist/index.mjs';

const here = dirname(fileURLToPath(import.meta.url));
const goldenDir = join(here, 'golden', 'revisions');
const fixturesDir = join(here, 'fixtures');

async function goldensIn(dir) {
  const out = [];
  for (const entry of await readdir(dir, { withFileTypes: true })) {
    if (entry.isDirectory()) out.push(...(await goldensIn(join(dir, entry.name))));
    else if (entry.name.endsWith('.json')) out.push(join(dir, entry.name));
  }
  return out;
}
const goldens = (await goldensIn(goldenDir)).map((file) => relative(goldenDir, file)).sort();

/** A change as the harness writes one: docx4j's type and kind names, the date as an Instant prints, the part's name, its pieces' kinds. */
const shape = (change) => ({
  type: change.type.toUpperCase(),
  kind: change.kind,
  author: change.author ?? null,
  date: change.date ? change.date.toISOString().replace(/\.000Z$/, 'Z') : null,
  id: change.id ?? null,
  text: change.text,
  part: change.part?.partName.name ?? null,
  pieces: change.pieces.map((piece) => piece.kind),
});

test('revisions goldens: one per fixture with revision markup, from RevisionsHarness version 1 in UTC', () => {
  assert.equal(goldens.length, 90, 'the fixtures whose word/ parts hold revision markup (CR-011 section 4)');
  return Promise.all(goldens.map(async (name) => {
    const golden = JSON.parse(await readFile(join(goldenDir, name), 'utf8'));
    assert.equal(golden.header.harnessVersion, '1', name);
    assert.equal(golden.header.timeZone, 'UTC', name);
    assert.equal(golden.header.fixture, name.replace(/\.json$/, '.docx'), name);
  }));
});

for (const name of goldens) {
  const fixture = name.replace(/\.json$/, '.docx');
  test(`revisions parity: ${fixture}`, async () => {
    const golden = JSON.parse(await readFile(join(goldenDir, name), 'utf8'));
    const bytes = await readFile(join(fixturesDir, fixture));
    const pkg = await WordprocessingMLPackage.load(bytes);
    assert.deepEqual((await pkg.getAllTrackedChanges()).map(shape), golden.listing, 'getAllTrackedChanges is TrackedChanges.in');
    for (const [scenario, expected] of Object.entries(golden.resolved)) {
      const resolved = await WordprocessingMLPackage.load(bytes);
      const verb = scenario.startsWith('reject') ? 'rejectAllRevisions' : 'acceptAllRevisions';
      const count = await resolved[verb](expected.filter);
      assert.equal(count, expected.count, `${scenario} ${JSON.stringify(expected.filter ?? {})}: the count`);
      assert.deepEqual((await resolved.getAllTrackedChanges()).map(shape), expected.left, `${scenario} ${JSON.stringify(expected.filter ?? {})}: what is left`);
    }
  });
}
