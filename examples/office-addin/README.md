# A Word add-in that runs the engine in the task pane

The add-in does the editing itself. Word hands over the selection as a flat OPC package
(`getOoxml()`), `@docx4j/core-ts` loads it, the content API makes the edit, and the package goes
back through `insertOoxml()`. Nothing round-trips to a server, and the same `edit()` function
runs in Node under the `Word` shim, so the add-in's logic is testable without Word.

| File | |
|---|---|
| `edit.mjs` | the edit itself, on any `Word.Body` — Office JS's or this package's |
| `taskpane.ts` | the two buttons: through Word, and through the shim |
| `taskpane.html` | the task pane |
| `manifest.xml` | the add-in manifest (replace the `Id` and the URLs) |
| `office.d.ts` | the handful of Office JS globals used, so `@types/office-js` is not needed here |
| `run-in-node.mjs` | the "without Word" path as a script; `test/examples.test.mjs` runs it |

## Run the Node half

```
npm run build                              # from the repository root
node examples/office-addin/run-in-node.mjs
```

## Build and sideload

```
npx esbuild taskpane.ts --bundle --format=esm --target=es2020 --outfile=taskpane.js
npx http-server -S -C cert.pem -K key.pem -p 3000 .    # any https server on port 3000 will do
```

Office requires https, so the task pane needs a certificate your machine trusts;
`npx office-addin-dev-certs install` makes one.

Then sideload `manifest.xml`:

- **Windows**: share a folder, add it under File > Options > Trust Center > Trust Center Settings >
  Trusted Add-in Catalogs, restart Word, and pick the add-in from Insert > My Add-ins > Shared Folder.
- **macOS**: copy `manifest.xml` into
  `~/Library/Containers/com.microsoft.Word/Data/Documents/wef`, restart Word, Insert > My Add-ins.
- **Word on the web**: Insert > Add-ins > Upload My Add-in.

Replace the manifest's `Id` with a GUID of your own (`node -e "console.log(crypto.randomUUID())"`)
before sideloading; two add-ins with the same id cannot be installed together.

## Bundling notes

- Import from the package's subpaths, not from its internals: `@docx4j/core-ts` (everything),
  `/opc`, `/parts`, `/packages`, `/model`, `/office-js`. They are the `exports` map and are stable.
- **Never import `@docx4j/core-ts/node`** in an add-in: it is the one subpath that touches
  `node:fs`, and it exists so that nothing else does. No other module in the package imports a
  `node:` builtin, so no polyfill or `browser` field is needed.
- `xpath` is an optional peer dependency, for XPath over custom XML parts in Node. A browser and
  an add-in use the DOM's own `document.evaluate`, so leave it out of the bundle; set
  `pkg.xpathEngine` if you want another one.
- The zip container imports `fflate` statically (about 8 KB gzipped) and `OpcPackage.load` reaches
  it, so it comes along even in an add-in that only ever sees flat OPC (CR-001 section 12).
- ESM only: build with `--format=esm`, or let your bundler transpile.
