# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

`@docx4j/core-ts`: the TypeScript counterpart of docx4j's `docx4j-core` module. It provides the Open
Packaging layer (packages, parts, relationships, content types; zip and flat OPC containers), the
typed parts (MainDocumentPart, StyleDefinitionsPart, ...) and the resolution utilities
(PropertyResolver, list numbering `Emulator`, font selection) over `@docx4j/generated-objects-ts`,
the generated Office Open XML object model (the counterpart of `docx4j-generated-objects`), and
the `@docx4j/jsonix` runtime. Design and scope live in `docs/change-requests/`; CR-001 is the
engine and is the spec for everything below.

**Status:** CR-001 Phase A (packaging, parts, packages, MCE), CR-001 Phase B (`PropertyResolver`,
list numbering `Emulator`, `RunFontSelector` and `IdentityPlusMapper`, held to docx4j's own
answers by 45 parity goldens) and CR-002 phases B to I, the whole content API (`Body`,
`Paragraph`, `Range`, `Font`, `Table`, `InlinePicture`, `ContentControl` with `XmlMapping` and
the typed kinds, `Comment`, `TrackedChange`, `List` and `ListItem`, custom XML parts with XPath,
search, `replaceText`, `insertOoxml`, addresses, `outline`; the `Word` shim on `./office-js`) are
implemented; CR-002 phase A is the objects package's `builders/wml`. Not yet: CR-001 Phase C.
Each CR's last sections record decisions and departures.

The dividing rule with the objects package: anything that needs only an object tree (helpers,
the `XmlUtils`-style facade, flat OPC typing) lives there; anything that needs parts or
relationships lives here. This package re-exports the objects facade so both read the same
(its flat OPC `Part` type is renamed `FlatOpcPart`; `Part` here is the part class).

## Commands

```
npm ci              # fflate, typescript, xpath (dev; an optional peer for consumers), @docx4j/generated-objects-ts and @docx4j/jsonix, as locked in package-lock.json

npm run build       # tsc -p tsconfig.build.json: src/ -> dist/ (git-ignored)
npm run typecheck   # tsc --strict over src/ and test/*.ts (skipLibCheck: fflate 0.8.3's typings need TS 5.7)
npm test            # pretest regenerates src/office-js/supported.generated.mts; then build, the nodenext consumer check (test/nodenext), node --test test/*.test.mjs
                    # (a shell glob: Node 22 does not accept a directory, and Node 18 would also run test/helpers.mjs)
node --test test/roundtrip.test.mjs   # one file, after npm run build

npm run generate:fonts   # src/model/fonts/*.generated.mts from ../docx4j's font tables and themes; committed output
```

Dependencies come from npm; `package-lock.json` is committed, and `npm update` then a lockfile commit
picks up newer versions. To try an unreleased change of the objects package, build it in
`../docx4j-generated-objects-ts` and run `npm install --no-save ../docx4j-generated-objects-ts` here (a
symlink; `npm ci` restores the locked version), but never release against it. CI
(`.github/workflows/test.yml`) runs `npm ci`, typecheck and test on Node 18, 20, 22. Releases publish to
npm from `.github/workflows/push-to-npm.yml` on a GitHub release (trusted publishing, tag = `package.json`
version); see `RELEASING.md`.
Tests import from `../dist/`, so build before running them by hand.
`test/nodenext/consumer.mts` (compile-only) imports each public entry point by the package's own name
under `module`/`moduleResolution: nodenext`, as a Node ES module consumer does; the repository's
tsconfigs use `bundler`, which accepts extensionless relative imports in the emitted declarations. Import
with the `.mjs` extension in `src/`. Its `@ts-expect-error` lines go unused, and fail, if an entry
point's types degrade to `any`.

## Architecture

Names follow `org.docx4j.openpackaging` and friends, so the docx4j Java source is the reference
for behaviour. Layout and `exports` subpaths:

```
src/xml/dom.mts   parse/serialize over Jsonix.DOM (typed since @docx4j/jsonix 3.2.1), text and base64 helpers
src/opc/          PartName, ContentTypes, ContentTypeManager, PartStore/PartSink + Memory/Zip/FlatOpc, Load, Save, mce/, exceptions
src/parts/        Part, BinaryPart (+ImagePart, ...), XmlPart<T>, DefaultXmlPart (+CustomXmlDataStoragePart), RelationshipsPart,
                  Parts, Namespaces, PartRegistry; wml/ dml/ pml/ sml/ docProps/ customXml/ typed parts
src/packages/     OpcPackage, WordprocessingMLPackage (createPackage, default styles and theme, body, outline, paragraphAt, getPropertyResolver/getNumberingEmulator/refresh), PresentationMLPackage, SpreadsheetMLPackage, registry
src/model/properties/ PropertyResolver, the property catalogue (one entry per schema member) and the catalogue-driven half of StyleUtil (apply, applyStyleLevel and the toggle XOR, isEmpty, unset, hasDirectFormatting)
src/model/listnumbering/ definitions.mts (LevelDefinition, ListDefinition, NumberingDefinitions), state.mts (Counter, NumberingState, NumberingStates: one per story), formats.mts (the label formatters), Emulator.mts
src/model/fonts/  RunFontSelector (the document font per code point), ThemeFonts, Mapper/IdentityPlusMapper/PhysicalFont/FontRegistry, FontFallback, fontsInUse; *.generated.mts from npm run generate:fonts
src/model/content/ the content API in Office JS shapes: Body, Paragraph, Range, Font, Table (+TableRow, TableCell), InlinePicture, ContentControl, Comment, TrackedChange, List (+ListItem; labels from the Emulator, one story walk per read or Body.listLabels() for all), search;
                  ooxml.mts is insertOoxml/insertXml (flat OPC in, referenced parts copied); comments.mts the comment plumbing (parts side in parts/wml/comments.mts);
                  tree.mts is the paragraph text model (segmentsOf, childrenOf; runItemsOf is re-exported from builders/wml); fragments, run mapping and traversal come from the objects package's builders/wml
src/model/customxml/ CustomXmlPart/CustomXmlNode over the custom XML DOM parts, XPathEngine (native document.evaluate, else the optional xpath package; await pkg.customXmlParts.load() once),
                  XmlMapping over w:dataBinding, the typed content-control kinds, insertContentControl, applyBindings/updateFromContentControls (docx4j BindingHandler)
src/model/content/tracking.mts, TrackedChange.mts: change tracking (pkg.changeTrackingMode; revision markup written by the paragraph primitives so every caller inherits it)
src/office-js/    the Word shim (Word.run(pkg, fn), context.sync, proxies throwing NotSupportedError, enums, Word.supported) and toApiScript; exported only from ./office-js.
                  supported.generated.mts is written by scripts/generate-supported.mjs from test/office-js-subset.ts (npm run generate; pretest runs it): commit it with every subset change
src/index.mts     re-exports all of the above plus the objects facade
```

Key mechanics:

- **Untouched parts round-trip byte for byte.** A package keeps its `sourcePartStore`; on save,
  only relationships parts and `XmlPart`s that were unmarshalled (or had contents/bytes set) are
  re-marshalled, everything else is copied from the source. Do not call `getContents()` on a
  part unless the caller asked for its contents. `[Content_Types].xml` is always regenerated.
- **Container is separate from package.** `PartStore` (load) and `PartSink<R>` (save) abstract
  zip (`fflate`; the central directory is parsed here, entries inflate lazily), flat OPC
  (`pkg:package`; content types are per part, so the store synthesises `[Content_Types].xml`),
  and memory. `OpcPackage.load` sniffs: string is flat OPC, bytes are a zip, else a store.
- **`XmlPart<T>.getContents()` is async** because the facade builds its Jsonix context lazily;
  afterwards `contents` is synchronous. `OpcPackage.unmarshalAll()` does them all.
- **Load follows relationships** from `/_rels/.rels`; unreachable parts are not loaded. The
  registry maps content type (relationship type first, for altChunk/embedded/OLE/custom XML)
  to a part class; unknown XML becomes `DefaultXmlPart`, anything else `BinaryPart`. Custom XML
  parts are indexed by the `ds:itemID` of their properties part (read from the DOM, so that
  part stays untouched).
- **MCE preprocessing** resolves `mc:AlternateContent` (first `mc:Choice` whose `Requires`
  prefixes are all in `UNDERSTOOD_NAMESPACES`, else `mc:Fallback`) on the DOM before
  unmarshalling. It is required, not cosmetic: a `w:drawing` inside `mc:Choice` cannot be typed
  by the model. `UNDERSTOOD_NAMESPACES` is a hand-kept copy of the generated modules' namespaces.
- **Namespace prefixes** are the facade's job (objects CR-001: `NAMESPACE_PREFIXES` default,
  `mc:Ignorable` declarations kept, `xml` never declared). Nothing here touches prefixes.
- **The content API is views, not a model.** `Body`/`Paragraph`/`Range` hold a reference to
  the element and the array containing it; nothing is cached but the `Body` per container
  (a WeakMap in `parts/wml`). Text edits work on the paragraph's text segments (`segmentsOf`)
  and edit `w:t` values in place, split runs only when formatting a span, and link `PARENT`
  on everything inserted (`linkParents`). `Font` and `Paragraph` reads report **effective**
  formatting through the `PropertyResolver` (and, for `Font.name`, the `RunFontSelector`);
  `{ direct: true }` on `getFont` / `formatting` reads the direct values, and writes are always
  direct. `test/office-js-subset.ts` must stay assignable: it is the Office JS promise.
  Fragments (`wml`), the element builders (`p`, `r`, `t`, `tbl`, `tr`, `tc`, `sdt`, `sdtPr`,
  `inlinePicture`), the run mapping (`applyRunOptions` / `readRunOptions`, `rPrToElements` /
  `rPrFromElements`), the `w:sdt` accessors (`sdtProperty`, `sdtKindOf`, `nextSdtId`) and
  traversal (`walk`, `walkAll`, `find`, `linkParents`, `textOf`) are imported from
  `@docx4j/generated-objects-ts/builders/wml`, and `deepCopyAs` / `deepCopyAsSync` from the
  facade (objects CR-003 phase A, in 0.1.4); do not re-create them here. Run holders (`w:ins`, `w:del`, `w:moveFrom`, `w:moveTo`) keep their runs under
  `customXmlOrSmartTagOrSdt`, not `content`: always go through `runItemsOf`.
  objects CR-002 specifies and move there when it lands.
- **No runtime import cycles:** `Part` gets its relationships-part factory through
  `setRelationshipsPartFactory`, and `OpcPackage.load` picks subclasses through
  `packages/registry.mts`; each package module registers itself, `packages/index.mts` imports all.

## Tests

`test/*.test.mjs` on Node's runner, against `dist/` (`content.test.mjs` covers the content API); helpers in `test/helpers.mjs` (`plain()`
strips `PARENT` for deep equality). Fixtures under `test/fixtures/` are from docx4j's
`docx4j-core-tests` resources; `test/README.md` lists them and holds the manual Word acceptance
checklist. The contract: untouched parts byte-identical after a round trip, re-marshalled parts
deep-equal after reload, flat OPC through the objects package's `unmarshalPackage` and back.
`test/golden/` holds 45 JSON goldens - what **docx4j itself** answers for each fixture
(effective properties, style resolutions, table styles, list labels and counters, the document
font of every character), written by the Maven harness in `test/java/` and compared by
`test/parity.test.mjs`, which unmarshals both sides and deep-equals object trees, never text.
Regenerate by hand (build docx4j, then `mvn -q -o compile exec:java -Dfixtures=../fixtures
-Dout=../golden -Ddocx4j.commit=<hash>` in `test/java`; each README says more);
`.github/workflows/parity.yml` does it weekly against docx4j's head and opens a pull request
when an answer moves. Parity is zero differences.

## Rules

- ES modules only (`"type": "module"`, `.mts` sources built to `dist/*.mjs` with `.d.mts`). Unlike
  the objects package there are no UMD files here, so `"type": "module"` is fine.
- Public paths are the `exports` map only (`.`, `./opc`, `./parts`, `./packages`, `./model`,
  `./office-js`). Keep them stable; add subpaths deliberately.
- Names follow docx4j (`OpcPackage`, `WordprocessingMLPackage`, `MainDocumentPart`,
  `RelationshipsPart`, `PropertyResolver`, `Emulator`) so docx4j Java code and documentation
  transfer; where Java conventions read badly in TypeScript use camelCase and note the mapping in
  the class doc comment. Document deliberate departures in the CR that introduces them
  (CR-001 section 12 has Phase A's, sections 15.1 to 15.4 Phase B's).
- The parity goldens are docx4j's answers, not ours: a difference is investigated, and fixed in
  the port unless the Java is plainly wrong, in which case it is reported to docx4j and the
  golden stays.
- Never edit `../docx4j-generated-objects-ts/modules/`; it is generated. A gap in the objects
  facade is fixed there (its own CR and release), then the dependency range is raised here.
  Runtime gaps go to `../jsonix` the same way (CR-001 section 9 lists the known ones).
- Work is proposed as numbered change requests in `docs/change-requests/` (index in its
  `README.md`); record decisions and implementation notes inside the CR. Commit only when asked.

## Related repositories (checked out side by side)

- `../docx4j-generated-objects-ts`: the object model and facade this package builds on; its
  `CLAUDE.md` explains what the declarations promise, its `docs/change-requests/` holds the
  facade CRs (CR-001: the namespace prefix table). Content is built with its generated
  factories: `el/org_docx4j_wml` (`el.p`, `el.r`, `el.t`: `{ name, value }` pairs with
  `TYPE_NAME`) and `factory/org_docx4j_wml` (`createP(init?)`, `createRT`, `createPElement`);
  never write qualified names by hand.
- `../jsonix`: the `@docx4j/jsonix` runtime; the npm package is `nodejs/scripts`, typings in
  `nodejs/scripts/types/main.d.ts`.
- `../jsonix-schema-compiler`: generates the objects package's `modules/`; only relevant when a
  gap turns out to be in the generated declarations rather than the facade.
- `../docx4j`: the Java original. Releases live on `VERSION_x_y_z` branches (`master` is old);
  CR-001 was written against `VERSION_17_1_1`, which is the checked-out branch. The classes being
  ported are under `docx4j-core/src/main/java/org/docx4j/`: `openpackaging/` (`io3/` for
  Load3/Save and `io3/stores/` for `PartStore`, `parts/relationships/` for `RelationshipsPart`
  and `Namespaces`, `contenttype/`, `packages/`), `model/PropertyResolver.java`,
  `model/listnumbering/`, `model/styles/` (`StyleUtil`, `PropertyCatalogue`), `fonts/`;
  `jaxb/McSelection.java` and `McMode.java` are the `mc:AlternateContent` reference (CR-021),
  `jaxb/mc-preprocessor.xslt` the load-time one. The schemas are `xsd/ROOT.xsd`. Phase B was
  ported against `7fba7a150`, which is what the goldens record; build docx4j from that commit
  before regenerating them, and never modify that checkout.

## Portfolio task registry

This repository's change requests are indexed, with their dependencies on work in the other
docx4j repositories, in `../docx4j-portfolio/tasks.yaml` (ids `<repo>/<CR>[.<phase>]`; this
repository's key is `core-ts`).

- When a CR's status changes (a phase lands; a CR is proposed, deferred or abandoned) or its
  dependencies change, update the matching entry in `tasks.yaml` in the same session (`status`,
  `depends_on`; add an entry for a new CR or phase).
- Then run `python3 ../docx4j-portfolio/scripts/tasks.py check`. It reports `CHANGED` for each CR
  whose Status line was edited; once the registry entry agrees, run `tasks.py accept` (and
  `tasks.py graph` if dependencies changed).
- Before starting a CR or phase, check `python3 ../docx4j-portfolio/scripts/tasks.py blocked`: it
  may be waiting on work in another repository. The portfolio rule of 2026-09-15 (XSL-FO fidelity
  work in `../docx4j` before porting docx4j code) was lifted for CR-001 Phase B on 2026-09-19
  (section 14.5); the editor (`../docx4j-ts-editor`) waits on phases here.
