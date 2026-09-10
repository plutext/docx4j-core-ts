// @docx4j/core-ts: the counterpart of docx4j-core. The Open Packaging layer, the typed parts and
// the packages are specified in docs/change-requests/CR-001-engine.md. The object model's facade
// (docx4j's XmlUtils names) is re-exported so that code written against
// @docx4j/generated-objects-ts reads the same here.
export * from '@docx4j/generated-objects-ts';
// The facade's flat OPC `Part` type (pkg:part) is renamed: `Part` here is the part class, as in docx4j.
export type { Part as FlatOpcPart } from '@docx4j/generated-objects-ts';
export { Part } from './parts/Part.mjs';
export * from './opc/index.mjs';
export * from './parts/index.mjs';
export * from './packages/index.mjs';
export * from './model/index.mjs';
