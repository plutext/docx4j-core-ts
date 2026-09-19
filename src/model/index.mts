// src/model: the content API, property resolution and list numbering; fonts follow with the
// rest of CR-001 Phase B.
export * from './content/index.mjs';
// CR-001 Phase B step 2: PropertyResolver, the property catalogue and StyleUtil.
export * from './properties/index.mjs';
// CR-001 Phase B step 3: the list definitions, the counters and the numbering Emulator.
export * from './listnumbering/index.mjs';
// CR-001 Phase B step 4: RunFontSelector, the Mapper and the default theme.
export * from './fonts/index.mjs';
// CR-002 phase E: custom XML parts, XML mapping, the XPath engine, the typed content controls.
export * from './customxml/index.mjs';
// The source generator of CR-002 phase I: the content-API calls that reproduce a paragraph, a
// range or a table. The `Word` shim it lives next to stays on the ./office-js subpath.
export { toApiScript, type ApiScriptOptions, type ApiScriptTarget } from '../office-js/toApiScript.mjs';
