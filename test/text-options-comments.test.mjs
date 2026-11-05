// CR-012 section 29 (WD 0.91, docx4j-python CR-035): one setting for comments. `comments` is the only state, `include` is
// headers, footers and glossary, and a caller's `comments: false` holds under `profile: word-text`. `include: comments`,
// released in 0.6.0, stays as a deprecated spelling of `comments: true`. Until docx4j moves, the preset's oracle is the WD's
// REQ-1002, not docx4j's WORD_TEXT_VALUES, which still names comments under include.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import {
  WordprocessingMLPackage, Texts, PlainText, TextOptions, TextOptionError, setLogger,
} from '../dist/index.mjs';

const here = dirname(fileURLToPath(import.meta.url));
const probe = await WordprocessingMLPackage.load(readFileSync(join(here, 'text-extraction', 'suite', 'phase0', 'extraction-probe.docx')));
const COMMENT = 'CMT28_COMMENT_TEXT';
const textOf = async (options) => PlainText.of(await Texts.of(probe, options));

// first in the file: the notice is once a process
test('include: comments writes its notice once a process, and comments: true writes none', () => {
  const said = [];
  setLogger({ warn: (message) => said.push(message) });
  try {
    TextOptions.of({ comments: true, include: ['headers'] });
    assert.deepEqual(said, []);
    TextOptions.of({ include: 'comments' });
    TextOptions.of({ include: ['comments'] });
    new TextOptions().include('comments');
    assert.equal(said.length, 1);
    assert.match(said[0], /include: comments is deprecated: use comments: true/);
  } finally {
    setLogger();
  }
});

test('the contract: what a caller writes under word-text, as the options and as the text', async () => {
  for (const [init, comments] of [
    [{ profile: 'word-text' }, true],
    [{ profile: 'word-text', comments: false }, false],
    [{ comments: false, profile: 'word-text' }, false],
    [{ profile: 'word-text', include: ['headers', 'footers'] }, true],
    [{ profile: 'word-text', comments: false, include: ['headers', 'footers'] }, false],
  ]) {
    const name = JSON.stringify(init);
    assert.equal(TextOptions.of(init).comments, comments, name);
    assert.equal((await textOf(init)).includes(COMMENT), comments, name);
  }
});

test('a caller\'s comments: false holds by every route; a direct assignment before a profile is the profile\'s to override', () => {
  assert.equal(new TextOptions().profile('word-text').set('comments', false).comments, false);
  assert.equal(new TextOptions().set('comments', false).profile('word-text').comments, false);
  assert.equal(new TextOptions().set('comments', 'false').profile('word-text').comments, false);
  const after = new TextOptions().profile('word-text');
  after.comments = false;
  assert.equal(after.comments, false);
  const before = new TextOptions();
  before.comments = false;
  assert.equal(before.profile('word-text').comments, true);
});

test('the old spelling in every form sets comments and stays out of include', () => {
  for (const o of [
    TextOptions.of({ include: 'comments' }), TextOptions.of({ include: ['comments'] }), TextOptions.of({ include: 'COMMENTS' }),
    new TextOptions().set('include', 'comments'), new TextOptions().include('comments'),
  ]) {
    assert.equal(o.comments, true);
    assert.deepEqual(o.toMap().include, []);
    assert.equal(o.toMap().comments, true);
    assert.ok(o.includes.has('comments'));
  }
  const both = TextOptions.of({ include: 'headers, comments footers' });
  assert.deepEqual(both.toMap().include, ['headers', 'footers']);
  assert.equal(both.comments, true);
});

test('include no longer speaks for comments: a later include leaves them, and includes still says what is in scope', () => {
  assert.equal(new TextOptions().set('include', 'comments').set('include', 'headers').comments, true);
  assert.equal(TextOptions.of({ comments: true, include: 'headers' }).comments, true);
  assert.equal(new TextOptions().comments, false);
  assert.ok(!new TextOptions().includes.has('comments'));
  assert.deepEqual([...TextOptions.of({ comments: true, include: 'glossary' }).includes].sort(), ['comments', 'glossary']);
});

test('one call saying comments: false and an include naming comments is refused, either way round; sequential calls, the last stands', () => {
  for (const init of [
    { comments: false, include: 'comments' }, { include: ['headers', 'comments'], comments: false },
    { comments: 'false', include: 'COMMENTS' }, { profile: 'word-text', comments: false, include: 'comments' },
  ]) {
    assert.throws(() => TextOptions.of(init), (e) => e instanceof TextOptionError && e.option === 'comments'
      && e.statement === 'REQ-504' && e.reason === 'invalid' && e.message.includes('deprecated spelling'), JSON.stringify(init));
  }
  assert.equal(TextOptions.of({ comments: true, include: 'comments' }).comments, true);
  assert.equal(new TextOptions().set('comments', false).set('include', 'comments').comments, true);
  assert.equal(new TextOptions().set('include', 'comments').set('comments', false).comments, false);
});

test('word-text is REQ-1002\'s, and reads as it did', async () => {
  const o = new TextOptions().profile('word-text');
  assert.deepEqual(o.toMap().include, ['headers', 'footers']);
  assert.ok(o.comments && o.includes.has('headers') && o.includes.has('footers') && o.includes.has('comments'));
  const preset = await textOf({ profile: 'word-text' });
  assert.ok(preset.includes(COMMENT));
  assert.equal(await textOf({ profile: 'word-text', include: 'headers,footers,comments' }), preset); // the preset's old form
});

test('options round-trip by name; a mapping in the old form loads, but for the one that contradicts itself', () => {
  for (const o of [new TextOptions(), new TextOptions().profile('word-text'), TextOptions.of({ profile: 'word-text', comments: false })]) {
    assert.deepEqual(TextOptions.of(o.toMap()).toMap(), o.toMap());
  }
  const old = TextOptions.of({ include: ['headers', 'footers', 'comments'], comments: true }); // word-text's map before WD 0.91
  assert.deepEqual(old.toMap().include, ['headers', 'footers']);
  assert.equal(old.comments, true);
  // 0.6.0's toMap() after include: comments reported the two apart, so such a saved map is refused, naming the fix
  assert.throws(() => TextOptions.of({ include: ['comments'], comments: false }), TextOptionError);
});
