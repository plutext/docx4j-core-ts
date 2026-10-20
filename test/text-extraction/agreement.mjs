// CR-012 phase 0 (D8 b): the agreement test's runner, a mirror of docx4j's ViewAgreementTest (WD
// REQ-406 and REQ-1202). For each fixture, in the accepted and original views, in the default scope
// and with every story and content class included, the view walk's plain text must equal the plain
// text of a copy resolved by acceptAllRevisions or rejectAllRevisions and read with no view logic.
// The copy is the reference: where they differ, the walk is wrong.
//
// The fixtures are docx4j's: every .docx under its revisions/ (CR-043's) and text-extraction/, as
// manifest.json lists them, in docx4j's order. As in docx4j the copy is the fixture loaded a second
// time, not clone().
//
// An engine is the text model's (docx4j `PlainText.of(Texts.of(...))` and `Texts.ofResolved`):
//   walk(pkg, view, all)      the walk's plain text, view 'accepted' or 'original'
//   resolved(copy, view, all) the resolved copy's plain text, read with no view logic
// `all` asks for headers and footers, comments, hidden text and field codes. With no engine (null)
// every comparison is pending.
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const here = dirname(fileURLToPath(import.meta.url));
const root = dirname(dirname(here));

const REVISIONS = 'docx4j-core-tests/src/test/resources/revisions/';
const SUITE = 'docx4j-core-tests/src/test/resources/text-extraction/';

/** docx4j's fixture list: revisions/ then text-extraction/, each sorted, .docx only; with the path here. */
export function agreementFixtures() {
  const { files } = JSON.parse(readFileSync(join(here, 'manifest.json'), 'utf8'));
  const under = (prefix) => files
    .filter((f) => f.docx4j.startsWith(prefix) && f.docx4j.endsWith('.docx'))
    .sort((a, b) => (a.docx4j < b.docx4j ? -1 : a.docx4j > b.docx4j ? 1 : 0));
  return [...under(REVISIONS), ...under(SUITE)].map((f) => ({
    name: f.docx4j.startsWith(REVISIONS) ? `revisions/${f.docx4j.slice(REVISIONS.length)}` : `text-extraction/${f.docx4j.slice(SUITE.length)}`,
    path: join(root, f.path),
  }));
}

/** The first line on which two texts differ, as ViewAgreementTest reports it. */
export function firstDifference(walk, reference) {
  const w = walk.split('\n');
  const r = reference.split('\n');
  let i = 0;
  while (i < Math.min(w.length, r.length) && w[i] === r[i]) i++;
  return `: line ${i + 1}\n  walk: ${i < w.length ? w[i] : '<end>'}\n  copy: ${i < r.length ? r[i] : '<end>'}`;
}

/**
 * Runs the comparison over the fixtures. `load(bytes)` loads a package (WordprocessingMLPackage.load).
 * Returns the differences, each naming the fixture, view and scope; and how many comparisons were
 * made and how many are pending.
 */
export async function agreement(fixtures, engine, load) {
  const differences = [];
  let compared = 0;
  let pending = 0;
  for (const f of fixtures) {
    const bytes = readFileSync(f.path);
    for (const view of ['accepted', 'original']) {
      for (const all of [false, true]) {
        if (engine === null) { pending++; continue; }
        const walk = await engine.walk(await load(bytes), view, all);
        const copy = await load(bytes);
        if (view === 'accepted') await copy.acceptAllRevisions();
        else await copy.rejectAllRevisions();
        const reference = await engine.resolved(copy, view, all);
        compared++;
        if (walk !== reference) differences.push(`${f.name} ${view}${all ? ' (all stories)' : ''}${firstDifference(walk, reference)}`);
      }
    }
  }
  return { compared, pending, differences };
}
