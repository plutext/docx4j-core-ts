# In a Word add-in

Office JS gives an add-in the document as markup, through `getOoxml()` and `insertOoxml()`, and
otherwise only what its own API can say. `@docx4j/core-ts` runs **inside the add-in**: the flat
OPC package Word hands over is loaded here, edited through the same content API you would use in
Node, and handed back. Nothing goes to a server, and everything the object model can express is
reachable, not just Office JS's verbs.

A complete, sideloadable example is [`examples/office-addin/`](../../examples/office-addin).

## Flat OPC in and out

`getOoxml()` returns a `pkg:package` string; `WordprocessingMLPackage.load` takes one directly,
and `saveFlatOpc()` produces one for `insertOoxml()`.

```ts
import { WordprocessingMLPackage } from '@docx4j/core-ts';

await Word.run(async (context) => {
  const selection = context.document.getSelection();
  const ooxml = selection.getOoxml();
  await context.sync();

  const pkg = await WordprocessingMLPackage.load(ooxml.value);
  const body = await pkg.getBody();
  for (const range of body.search('draft', { matchCase: false })) {
    range.font.highlightColor = '#FFFF00';
  }

  selection.insertOoxml(await pkg.saveFlatOpc(), 'Replace');
  await context.sync();
});
```

The selection's package is a whole document: `word/document.xml` with the styles, the numbering
and the theme parts Word thought relevant. Everything the engine does works on it — the resolver,
the numbering emulator, the font selector, the custom XML parts.

## Testing add-in code in Node

`@docx4j/core-ts/office-js` is a `Word` shim: `Word.run(pkg, fn)` gives the callback a `context`
whose `document.body` is the package's body, `load()` is a no-op and `context.sync()` resolves
what was asynchronous. An add-in's batch therefore runs unchanged in a test or in CI, with the
saved `.docx` as the assertion.

```ts
import { readFile } from 'node:fs/promises';
import { WordprocessingMLPackage } from '@docx4j/core-ts';
import { Word } from '@docx4j/core-ts/office-js';

const pkg = await WordprocessingMLPackage.load(new Uint8Array(await readFile('fixture.docx')));
await Word.run(pkg, async (context) => {
  const paragraphs = context.document.body.paragraphs;
  paragraphs.load('text');
  await context.sync();
  paragraphs.items[0].styleBuiltIn = Word.Style.heading1;
});
```

Write the edit once, against a `Word.Body`, and call it from both places — that is what
`examples/office-addin/edit.mjs` is, and `run-in-node.mjs` beside it is the Node half, run by
this repository's own test suite so it cannot rot.

A member this package does not implement throws `NotSupportedError` naming it, rather than doing
nothing quietly. `Word.supported` is the set of `Class.member` strings that are implemented
(`Word.supported.has('Body.insertParagraph')`), so a tool can check a script before running it.

## Bundling

- **Import from the `exports` subpaths only**: `@docx4j/core-ts` (everything), `/opc`, `/parts`,
  `/packages`, `/model`, `/office-js`. They are the public surface and are stable; reaching into
  `dist/` is not.
- **Never import `@docx4j/core-ts/node`.** It is the one subpath that touches `node:fs` (the
  unzipped-directory container), and it exists so that nothing else does. No other module in the
  package imports a `node:` builtin, so an add-in bundle needs no polyfills and no `browser`
  field.
- **`xpath` is optional**, and Node-only: it backs XPath over custom XML parts where there is no
  DOM. A browser and an add-in use `document.evaluate`, which is the default there, so leave
  `xpath` out of the bundle. `pkg.xpathEngine` takes an engine of your own.
- **ESM only.** Bundle with `--format=esm` (esbuild) or let webpack/rollup transpile:

  ```
  npx esbuild taskpane.ts --bundle --format=esm --target=es2020 --outfile=taskpane.js
  ```

- The zip container imports `fflate` statically (about 8 KB gzipped) and `OpcPackage.load`
  reaches it, so it is in the bundle even when only flat OPC is used (CR-001 section 12).
- Office requires https for the task pane; `npx office-addin-dev-certs install` makes a
  certificate your machine trusts. The example's README has the sideload steps for Windows,
  macOS and Word on the web.
