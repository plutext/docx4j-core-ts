#!/usr/bin/env node
// Copies the text extraction conformance suite from a docx4j checkout at a named commit into
// test/text-extraction/, and writes its manifest (CR-012 phase 0, D8 a and b).
//
//   node scripts/copy-text-extraction-suite.mjs ../docx4j <commit>
//
// What is copied, from the commit (git archive), never from the working tree:
//   - docx4j-core-tests/src/test/resources/text-extraction/ -> test/text-extraction/suite/, whole and
//     unedited: the fixtures, the TSVs, the goldens, GOLDENS.md and INSPECTION.md (the suite);
//   - docx4j-core/src/main/resources/org/docx4j/jaxb/understood-namespaces.txt (REQ-306's list);
//   - the .docx under docx4j-core-tests/src/test/resources/revisions/ (CR-043's fixtures, which
//     ViewAgreementTest reads with the suite's): one whose bytes this repository already holds under
//     test/fixtures/ is not copied again, the manifest naming that copy; any other goes to
//     test/text-extraction/revisions/, at its path under docx4j's revisions/.
//
// test/text-extraction/manifest.json records the commit and, for each file, docx4j's path, its
// SHA-1 and the path here; test/text-extraction.test.mjs checks every file against it. A recopy
// replaces suite/ and revisions/ whole: review the difference (git diff) before committing it.
import { execFileSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import { mkdtempSync, mkdirSync, readFileSync, readdirSync, rmSync, writeFileSync, copyFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, join, relative, sep } from 'node:path';
import { fileURLToPath } from 'node:url';

const root = dirname(dirname(fileURLToPath(import.meta.url)));
const [docx4j, commit] = process.argv.slice(2);
if (!docx4j || !commit) {
  console.error('usage: node scripts/copy-text-extraction-suite.mjs <docx4j checkout> <commit>');
  process.exit(2);
}

const SUITE = 'docx4j-core-tests/src/test/resources/text-extraction';
const REVISIONS = 'docx4j-core-tests/src/test/resources/revisions';
const UNDERSTOOD = 'docx4j-core/src/main/resources/org/docx4j/jaxb/understood-namespaces.txt';

const git = (...args) => execFileSync('git', ['-C', docx4j, ...args], { maxBuffer: 1 << 30 });
const full = git('rev-parse', '--verify', `${commit}^{commit}`).toString().trim();
const date = git('show', '-s', '--format=%cs', full).toString().trim();
const subject = git('show', '-s', '--format=%s', full).toString().trim();

const scratch = mkdtempSync(join(tmpdir(), 'text-extraction-'));
try {
  execFileSync('tar', ['-x', '-C', scratch], { input: git('archive', '--format=tar', full, SUITE, REVISIONS, UNDERSTOOD) });

  const sha1 = (file) => createHash('sha1').update(readFileSync(file)).digest('hex');
  /** Every file under dir, as a '/'-separated path relative to it, sorted. */
  const filesUnder = (dir) => {
    const out = [];
    const walk = (d) => {
      for (const e of readdirSync(d, { withFileTypes: true })) {
        if (e.isDirectory()) walk(join(d, e.name));
        else if (e.isFile()) out.push(relative(dir, join(d, e.name)).split(sep).join('/'));
      }
    };
    walk(dir);
    return out.sort();
  };

  const target = join(root, 'test', 'text-extraction');
  rmSync(join(target, 'suite'), { recursive: true, force: true });
  rmSync(join(target, 'revisions'), { recursive: true, force: true });

  const files = [];
  const copy = (from, docx4jPath, to) => {
    mkdirSync(dirname(join(root, to)), { recursive: true });
    copyFileSync(from, join(root, to));
    files.push({ docx4j: docx4jPath, sha1: sha1(from), path: to });
  };

  for (const name of filesUnder(join(scratch, SUITE))) {
    copy(join(scratch, SUITE, name), `${SUITE}/${name}`, `test/text-extraction/suite/${name}`);
  }
  copy(join(scratch, UNDERSTOOD), UNDERSTOOD, 'test/text-extraction/understood-namespaces.txt');

  // what test/fixtures already holds, by content
  const held = new Map();
  for (const name of filesUnder(join(root, 'test', 'fixtures'))) {
    if (name.endsWith('.docx')) held.set(sha1(join(root, 'test', 'fixtures', name)), `test/fixtures/${name}`);
  }
  let reused = 0;
  for (const name of filesUnder(join(scratch, REVISIONS)).filter((n) => n.endsWith('.docx'))) {
    const from = join(scratch, REVISIONS, name);
    const hash = sha1(from);
    if (held.has(hash)) {
      files.push({ docx4j: `${REVISIONS}/${name}`, sha1: hash, path: held.get(hash) });
      reused++;
    } else {
      copy(from, `${REVISIONS}/${name}`, `test/text-extraction/revisions/${name}`);
    }
  }

  const manifest = { docx4j: { commit: full, date, subject }, files };
  writeFileSync(join(target, 'manifest.json'), `${JSON.stringify(manifest, null, 1)}\n`);
  console.log(`${files.length} files from docx4j ${full.slice(0, 9)} (${date}); ${reused} revisions fixtures already under test/fixtures`);
} finally {
  rmSync(scratch, { recursive: true, force: true });
}
