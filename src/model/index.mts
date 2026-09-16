// src/model: the content API now; PropertyResolver, list numbering and fonts with CR-001 Phase B.
export * from './content/index.mjs';
// CR-002 phase E: custom XML parts, XML mapping, the XPath engine, the typed content controls.
export * from './customxml/index.mjs';
// The source generator of CR-002 phase I: the content-API calls that reproduce a paragraph, a
// range or a table. The `Word` shim it lives next to stays on the ./office-js subpath.
export { toApiScript, type ApiScriptOptions, type ApiScriptTarget } from '../office-js/toApiScript.mjs';
