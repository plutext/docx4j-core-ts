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

**Status:** CR-001 Phase A (packaging, parts, packages, MCE) and CR-002 phases B and D (the
content API: `Body`, `Paragraph`, `Range`, `Font`, search, addresses, `outline`) are
implemented. Not yet: CR-001 Phase B (`PropertyResolver`, numbering, fonts) and C; CR-002 A
(the objects package's `builders/wml`), C (tables, pictures, `insertOoxml`) and E (custom XML,
XML mapping, content controls). Each CR's last sections record decisions and departures.

The dividing rule with the objects package: anything that needs only an object tree (helpers,
the `XmlUtils`-style facade, flat OPC typing) lives there; anything that needs parts or
relationships lives here. This package re-exports the objects facade so both read the same
(its flat OPC `Part` type is renamed `FlatOpcPart`; `Part` here is the part class).

## Commands

```
npm install         # fflate, typescript, @docx4j/generated-objects-ts and @docx4j/jsonix from npm

npm run build       # tsc -p tsconfig.build.json: src/ -> dist/ (git-ignored)
npm run typecheck   # tsc --strict over src/ and test/*.ts (skipLibCheck: fflate 0.8.3's typings need TS 5.7)
npm test            # build, then node --test test/  (every test/*.test.mjs)
node --test test/roundtrip.test.mjs   # one file, after npm run build
```

Dependencies come from npm. To try an unreleased change of the objects package, build it in
`../docx4j-generated-objects-ts` and run `npm install --no-save ../docx4j-generated-objects-ts` here (a
symlink; `npm install` restores the registry version), but never release against it. CI
(`.github/workflows/test.yml`) runs `npm install`, typecheck and test on Node 18, 20, 22.
Tests import from `../dist/`, so build before running them by hand.

## Architecture

Names follow `org.docx4j.openpackaging` and friends, so the docx4j Java source is the reference
for behaviour. Layout and `exports` subpaths:

```
src/xml/dom.mts   parse/serialize (Jsonix.DOM, cast: not in the runtime typings), text and base64 helpers
src/opc/          PartName, ContentTypes, ContentTypeManager, PartStore/PartSink + Memory/Zip/FlatOpc, Load, Save, mce/, exceptions
src/parts/        Part, BinaryPart (+ImagePart, ...), XmlPart<T>, DefaultXmlPart (+CustomXmlDataStoragePart), RelationshipsPart,
                  Parts, Namespaces, PartRegistry; wml/ dml/ pml/ sml/ docProps/ customXml/ typed parts
src/packages/     OpcPackage, WordprocessingMLPackage (createPackage, default styles, body, outline, paragraphAt), PresentationMLPackage, SpreadsheetMLPackage, registry
src/model/content/ the content API in Office JS shapes: Body, Paragraph, Range, Font, search; tree.mts is the paragraph text model (segmentsOf, runItemsOf, childrenOf); fragments, run mapping and traversal come from the objects package's builders/wml
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
  on everything inserted (`linkParents`). Font reads are direct formatting only until CR-001
  Phase B. `test/office-js-subset.ts` must stay assignable: it is the Office JS promise.
  Fragments (`wml`), the element builders (`p`, `r`, `t`, `tbl`), the run mapping
  (`applyRunOptions` / `readRunOptions`) and traversal (`walk`, `find`, `linkParents`,
  `textOf`) are imported from `@docx4j/generated-objects-ts/builders/wml`; do not re-create
  them here. Run holders (`w:ins`, `w:del`, `w:moveFrom`, `w:moveTo`) keep their runs under
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
Phase B adds golden files from a Java harness under `test/java/`.

## Rules

- ES modules only (`"type": "module"`, `.mts` sources built to `dist/*.mjs` with `.d.mts`). Unlike
  the objects package there are no UMD files here, so `"type": "module"` is fine.
- Public paths are the `exports` map only (`.`, `./opc`, `./parts`, `./packages`). Keep them
  stable; add subpaths deliberately.
- Names follow docx4j (`OpcPackage`, `WordprocessingMLPackage`, `MainDocumentPart`,
  `RelationshipsPart`, `PropertyResolver`, `Emulator`) so docx4j Java code and documentation
  transfer; where Java conventions read badly in TypeScript use camelCase and note the mapping in
  the class doc comment. Document deliberate departures in the CR that introduces them
  (CR-001 section 12 has Phase A's).
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
  `model/listnumbering/`, `fonts/`; `jaxb/mc-preprocessor.xslt` is the MCE reference. The
  schemas are `xsd/ROOT.xsd`.
