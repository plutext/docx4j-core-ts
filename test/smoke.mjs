import assert from 'node:assert/strict';
import * as core from '../dist/index.mjs';

// The facade of @docx4j/generated-objects-ts is re-exported unchanged.
for (const name of ['getContext', 'unmarshalString', 'marshalString', 'unmarshalPackage', 'marshalPackage', 'unwrap', 'deepCopy']) {
  assert.equal(typeof core[name], 'function', name);
}
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const xml = `<w:p xmlns:w="${W}"><w:r><w:t>Hello, core-ts</w:t></w:r></w:p>`;
const p = core.unwrap(await core.unmarshalString(xml));
assert.equal(p.TYPE_NAME, 'org_docx4j_wml.P');
assert.equal(core.unwrap(core.unwrap(p.content[0]).content[0]).value, 'Hello, core-ts');
console.log('core-ts smoke: facade re-export and unmarshal OK');
