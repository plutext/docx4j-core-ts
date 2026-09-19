// `@docx4j/core-ts/node`: the parts of the engine that need Node's own APIs.
//
// Everything else in this package runs unchanged in a browser or a Word add-in, which is why the
// `node:` imports are fenced off here: nothing under `.`, `./opc`, `./parts`, `./packages`,
// `./model` or `./office-js` reaches a Node builtin, so a bundle for the web never has to shim
// `node:fs`. Import this subpath only from code you know runs in Node.
//
//   import { DirectoryPartStore, DirectoryPartSink } from '@docx4j/core-ts/node';
//
//   const pkg = await WordprocessingMLPackage.load(await DirectoryPartStore.open('unzipped'));
//   await pkg.saveTo(new DirectoryPartSink('unzipped-out'));
export { DirectoryPartStore, DirectoryPartSink } from '../opc/DirectoryPartStore.mjs';
