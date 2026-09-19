// The examples under examples/ are run as scripts, against the fixtures, so they cannot rot:
// each must exit 0 and print the line the guide quotes.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { execFile } from 'node:child_process';
import { promisify } from 'node:util';
import { mkdtemp, rm, stat } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import { fixturesDir } from './helpers.mjs';

const run = promisify(execFile);
const root = dirname(dirname(fileURLToPath(import.meta.url)));

async function example(script, args = []) {
  const { stdout } = await run(process.execPath, [join(root, script), ...args], { cwd: root });
  return stdout;
}

test('examples/node/hello.mjs creates a document', async () => {
  const dir = await mkdtemp(join(tmpdir(), 'core-ts-examples-'));
  try {
    const out = join(dir, 'hello.docx');
    const stdout = await example('examples/node/hello.mjs', [out]);
    assert.match(stdout, /^paragraphs: 2$/m);
    assert.match(stdout, /^first paragraph style: Heading 1$/m);
    assert.ok((await stat(out)).size > 0);
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
});

test('examples/node/report.mjs reports on a fixture', async () => {
  const stdout = await example('examples/node/report.mjs', [join(fixturesDir, 'tracked-changes.docx')]);
  assert.match(stdout, /^outline: 74 paragraphs, 1 tables, 1 headers, 0 footers$/m);
  assert.match(stdout, /^tracked changes: 2$/m);
  assert.match(stdout, /Added by Jason Harrop/m);
  assert.match(stdout, /^list items: 5$/m);

  const withComments = await example('examples/node/report.mjs', [join(fixturesDir, 'loadAndSave.docx')]);
  assert.match(withComments, /^comments: 1$/m);
});

test('examples/node/pptx.mjs creates a presentation', async () => {
  const dir = await mkdtemp(join(tmpdir(), 'core-ts-examples-'));
  try {
    const out = join(dir, 'hello.pptx');
    const stdout = await example('examples/node/pptx.mjs', [out]);
    assert.match(stdout, /^slide size: 9144000 x 5143500 \(screen16x9\)$/m);
    assert.match(stdout, /^slides: 2, masters: 1, layouts: 1$/m);
    assert.ok((await stat(out)).size > 0);
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
});

test('examples/node/xlsx.mjs creates a workbook', async () => {
  const dir = await mkdtemp(join(tmpdir(), 'core-ts-examples-'));
  try {
    const out = join(dir, 'hello.xlsx');
    const stdout = await example('examples/node/xlsx.mjs', [out]);
    assert.match(stdout, /^sheets: Sales, Notes$/m);
    assert.match(stdout, /^rows in \/xl\/worksheets\/sheet1\.xml: 3$/m);
    assert.ok((await stat(out)).size > 0);
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
});

test('examples/node/directory.mjs round-trips through a directory', async () => {
  const dir = await mkdtemp(join(tmpdir(), 'core-ts-examples-'));
  try {
    const stdout = await example('examples/node/directory.mjs', [join(fixturesDir, 'loadAndSave.docx'), join(dir, 'unzipped')]);
    assert.match(stdout, /^unzipped 30 parts into /m);
    assert.match(stdout, /^byte-identical: 30 parts \(5 regenerated: the rels and the content types\)$/m);
    assert.ok(!stdout.includes('DIFFERS'));
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
});

test('examples/office-addin/run-in-node.mjs runs the add-in edit without Word', async () => {
  const stdout = await example('examples/office-addin/run-in-node.mjs', [join(fixturesDir, 'tracked-changes.docx')]);
  assert.match(stdout, /^highlighted 3 occurrences$/m);
  assert.match(stdout, /^first paragraph style: Heading 1$/m);
  assert.match(stdout, /^round trip style: Heading 1$/m);
});
