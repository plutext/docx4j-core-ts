# CR-001: The engine: container interface, Open Packaging layer, typed parts, resolution utilities

**Status:** Phase A implemented 2026-09-10 (both steps); Phases B and C proposed
**Depends on:** `@docx4j/generated-objects-ts` 0.1.0 (the object model and its facade),
`@docx4j/jsonix` 3.2.0 (`parentPointers`, `deepCopy`); one small runtime addition is listed in
section 9.
**Counterpart:** docx4j-core (`org.docx4j.openpackaging.*`, `org.docx4j.model.PropertyResolver`,
`org.docx4j.model.listnumbering.*`, `org.docx4j.fonts.*`), docx4j VERSION_17_1_1.

## 1. Summary

`@docx4j/core-ts` gives the generated object model a document to live in. It reads and writes
Office Open XML packages, as zip files and as the flat OPC (`pkg:package`) strings that Office JS
`getOoxml()` / `insertOoxml()` exchange, exposes their parts as typed objects related by
relationships exactly as docx4j does, and ports the three utilities every docx4j consumer ends up
needing: effective paragraph and run properties (`PropertyResolver`), list numbering
(`Emulator`), and font selection (`RunFontSelector` with a pluggable `Mapper`).

Design rule, from docx4j: **a part that is never touched is written back byte for byte.** Only
parts whose content was unmarshalled are re-marshalled on save. This is what makes round trips
safe with content the object model does not know, and it is what keeps saving fast.

Scope of this CR is WordprocessingML end to end. PresentationML and SpreadsheetML get their
package classes and main parts; their remaining parts load as generic parts through the same
registry and are typed in later CRs.

## 2. Non-goals

Encryption (password-protected packages), digital signatures, VBA, the XSLT/FO/PDF export line,
mail merge, OpenDoPE binding, `docx4j-markdown`, and the `docx4j-diffx` comparison are separate
products in Java and are separate CRs or packages here. Validation against the schemas is not
part of the engine (the declarations already do the static half).

## 3. Runtime and dependencies

- **Runtime targets:** Node 18+, browsers, Office add-ins (webpack-bundled). ES modules only.
- **Zip:** `fflate`. It is small (about 8 KB gzipped for the parts we use), synchronous and
  asynchronous, and identical in Node and browsers. JSZip is the familiar alternative; it is
  larger and promise-only, and its API adds nothing we need. The zip library is confined to one
  `PartStore` implementation, so the choice is reversible.
- **XML:** the `@docx4j/jsonix` runtime through the objects facade, and the DOM for what the
  model does not type (`DefaultXmlPart`, custom XML data storage parts, MCE preprocessing). In Node
  the runtime brings its own DOM implementation; the engine uses whatever the objects facade uses
  (`unmarshalNode` / `marshalNode`).
- **Content types:** `[Content_Types].xml` is not in the generated set (docx4j generates it
  separately from `xsd/contentTypes/opc-contentTypes.xsd`). It has two element types, `Default`
  and `Override`; the engine parses and writes it by hand (section 5.3) rather than adding a
  module to the objects package.

## 4. The container interface

docx4j's `org.docx4j.openpackaging.io3.stores.PartStore` separates *where bytes come from and go
to* from *what the package is*. The same split is the point here, because there are three
containers from day one: a zip, a flat OPC string, and memory.

```ts
/** Where a package's bytes come from. docx4j: io3.stores.PartStore (the load half). */
export interface PartStore {
  /** Part names as stored, without the leading '/', in container order. */
  partNames(): Iterable<string>;
  has(partName: string): boolean;
  /** Raw bytes of a part. Implementations may inflate lazily on first access. */
  load(partName: string): Promise<Uint8Array>;
  /** Size if known without loading (zip central directory), else undefined. */
  size?(partName: string): number | undefined;
}

/** Where a package's bytes go. docx4j: PartStore (the save half: setOutputStream/save*/finishSave). */
export interface PartSink {
  put(partName: string, bytes: Uint8Array, options?: { compress?: boolean }): void;
  /** Completes the container. */
  finish(): Promise<Uint8Array | string>;
}
```

Implementations:

| Class | Source | Notes |
|---|---|---|
| `ZipPartStore` | `Uint8Array` / `ArrayBuffer` | Reads the central directory once; inflates an entry on first `load`. |
| `FlatOpcPartStore` | `pkg:package` string, or a `PackageElement` from the objects facade | `pkg:xmlData` parts are serialised back to bytes (the model leaves them as DOM); `pkg:binaryData` is base64-decoded. |
| `MemoryPartStore` | `Map<string, Uint8Array>` | New packages, tests, and `PartSink` results. |
| `DirectoryPartStore` (Node only, later) | an unzipped directory | docx4j's `UnzippedPartStore`; useful for diffing. |
| `ZipPartSink` | | `[Content_Types].xml` is written first, as Word does; entries deflated except for already-compressed media (png, jpeg, ...) which are stored. |
| `FlatOpcPartSink` | | Produces the `pkg:package` string for `insertOoxml`. XML parts are embedded as XML, others as base64. |

The package keeps a reference to its **source** store. On save, a part that was never unmarshalled
(or whose bytes were never replaced) is copied from the source store to the sink unchanged
(docx4j `Save.saveRawXmlPart`, `OpcPackage.getSourcePartStore()`).

## 5. The Open Packaging layer

Names follow `org.docx4j.openpackaging`.

### 5.1 `PartName`

Value object for an OPC part name: leading `/`, no trailing `/`, segments percent-encoded per
OPC Annex A. `PartName.resolve(source, target)` resolves a relationship target relative to the
source part's directory (`../` handling, `TargetMode="External"` left untouched);
`PartName.relsFor(name)` gives `/word/_rels/document.xml.rels`; `extension`, `directory`,
`equals`. Part names are case-insensitive for equality (Word writes mixed case).

### 5.2 `RelationshipsPart`

An `XmlPart<Relationships>` over `org_docx4j_relationships.Relationships`. Package relationships
live at `/_rels/.rels`; every part with relationships owns one, linked from `Part.relationshipsPart`.
API as docx4j: `getRelationshipByType`, `getRelationshipsByType`, `getRelationshipById`,
`getPart(rel)`, `getRel(partName)`, `isATarget(partName)`, `addPart(part, mode)` (allocates the
next free `rIdN`; `mode` is docx4j's `AddPartBehaviour`: `OVERWRITE_IF_NAME_EXISTS`,
`RENAME_IF_NAME_EXISTS`, `REUSE_EXISTING`), `addRelationship`, `removePart`, `removeRelationship`.
`Namespaces` holds the relationship type constants (`OFFICE_DOCUMENT`, `STYLES`, `NUMBERING`,
`FONT_TABLE`, `HEADER`, `FOOTER`, `IMAGE`, `HYPERLINK`, `CUSTOM_XML`, ...), as docx4j's
`org.docx4j.openpackaging.parts.relationships.Namespaces`.

### 5.3 `ContentTypeManager`

Parses and writes `[Content_Types].xml`: `Default` (extension to content type) and `Override`
(part name to content type). `getContentType(partName)`, `addDefaultContentType`,
`addOverrideContentType`, `removeContentType`. `ContentTypes` holds the constants
(`WORDPROCESSINGML_DOCUMENT`, `WORDPROCESSINGML_STYLES`, `IMAGE_PNG`, ...), as docx4j's
`org.docx4j.openpackaging.contenttype.ContentTypes`.

### 5.4 Parts

```
Part                          partName, contentType, relationshipType, package,
                              relationshipsPart?, owningRelationshipPart, sourceRelationships[]
├── BinaryPart                getBytes(): Promise<Uint8Array>, setBytes()           docx4j BinaryPart
│   ├── ImagePart             one class; the image kind is the content type            docx4j BinaryPartAbstractImage + Image*Part
│   ├── EmbeddedPackagePart, OleObjectBinaryPart, ObfuscatedFontPart, ...
├── XmlPart<T>                getContents(): Promise<T>, contents (sync, once unmarshalled),
│   │                         setContents(T), isUnmarshalled, getXml(), rootName             docx4j JaxbXmlPart<E>
│   ├── RelationshipsPart<Relationships>
│   ├── MainDocumentPart<Document>, StyleDefinitionsPart<Styles>, NumberingDefinitionsPart<Numbering>,
│   │   FontTablePart<Fonts>, DocumentSettingsPart<CTSettings>, WebSettingsPart<CTWebSettings>,
│   │   HeaderPart<Hdr>, FooterPart<Ftr>, FootnotesPart<CTFootnotes>, EndnotesPart<CTEndnotes>,
│   │   CommentsPart<Comments>, CommentsExtendedPart, CommentsIdsPart, PeoplePart,
│   │   GlossaryDocumentPart<GlossaryDocument>, ThemePart<Theme>, BibliographyPart,
│   │   DocPropsCorePart<CoreProperties>, DocPropsExtendedPart<Properties>, DocPropsCustomPart<Properties>,
│   │   CustomXmlDataStoragePropertiesPart<DatastoreItem>, VMLPart, AlternativeFormatInputPart (altChunk),
│   │   Chart / ChartStyle / ChartColorStyle parts, diagram Data / Layout / Style / Colors parts,
│   │   PML: MainPresentationPart<Presentation>, SlidePart<Sld>, SlideLayoutPart, SlideMasterPart,
│   │        NotesSlidePart, NotesMasterPart, HandoutMasterPart, ViewPropertiesPart, PresentationPropertiesPart, TableStylesPart
│   │   SML: WorkbookPart<Workbook>, WorksheetPart<Worksheet>, SharedStringsPart<Sst>, StylesPart<CTStylesheet>, CalcChain, ...
├── CustomXmlDataStoragePart  the DOM (a customer's own schema; typed by the consumer's own mappings) docx4j CustomXmlDataStoragePart
└── DefaultXmlPart            the DOM, for XML the registry does not know                              docx4j DefaultXmlPart
```

`XmlPart<T>` is the JAXB-part analogue. `T` is the root element's *value* type from the objects
package (`Document`, `Styles`, ...); the part stores the `TypedNamedValue` it unmarshalled so the
root name round-trips, and exposes `contents` as its value, as docx4j's `getJaxbElement()` returns
the content and not the `JAXBElement` wrapper. `getContents()` is asynchronous because the objects
facade builds its context lazily with dynamic `import()`; once a part is unmarshalled, `contents`
is a plain synchronous property (docx4j: `getContents()` unmarshals on demand, `getJaxbElement()`
returns what is there). A consumer that needs everything synchronous can `await pkg.unmarshalAll()`
after load.

`PartRegistry` maps a content type to a part class, with a fallback on relationship type for the
generic content types (`application/xml` custom XML parts are recognised by their relationship,
as docx4j's `ContentTypeManager.getPart(partName, rel)` does). It ships with docx4j's table
(`ContentTypeManager.newPartForContentType`) and is extensible: `registry.register(contentType,
PartClass)`. Unknown XML becomes `DefaultXmlPart`, anything else `BinaryPart`; nothing is dropped.

### 5.5 Packages

```ts
class OpcPackage {                                     // docx4j OpcPackage
  static load(source: Uint8Array | ArrayBuffer | string | PartStore, options?: LoadOptions): Promise<OpcPackage>;
  contentTypeManager: ContentTypeManager;
  relationshipsPart: RelationshipsPart;                // /_rels/.rels
  parts: Parts;                                        // Map<PartName, Part>, docx4j Parts
  getPart(partName: PartName | string): Part | undefined;
  externalResources: Map<string, ExternalTarget>;
  customXmlDataStorageParts: Map<string, CustomXmlDataStoragePart>;   // by itemId, docx4j getCustomXmlDataStorageParts
  sourcePartStore?: PartStore;
  unmarshalAll(): Promise<void>;
  save(): Promise<Uint8Array>;                         // zip
  saveFlatOpc(): Promise<string>;                      // pkg:package
  saveTo(sink: PartSink): Promise<Uint8Array | string>;
}
class WordprocessingMLPackage extends OpcPackage {     // docx4j WordprocessingMLPackage
  static load(...): Promise<WordprocessingMLPackage>;  // throws if the main part is not a w:document
  static createPackage(options?: { pageSize?: PageSizePaper; landscape?: boolean }): Promise<WordprocessingMLPackage>;
  mainDocumentPart: MainDocumentPart;
  // shortcuts as docx4j: styles, numbering, fontTable, settings, theme, headers/footers by rel
  getPropertyResolver(): Promise<PropertyResolver>;    // docx4j: on MainDocumentPart; offered here too
  fontMapper?: Mapper;
}
class PresentationMLPackage extends OpcPackage { mainPresentationPart: MainPresentationPart; ... }
class SpreadsheetMLPackage extends OpcPackage { workbookPart: WorkbookPart; ... }
```

`OpcPackage.load` sniffs the kind: it reads `[Content_Types].xml` and `/_rels/.rels`, follows the
`officeDocument` relationship, and picks the package class from that part's content type. A
`string` source is a flat OPC package; bytes are a zip; a `PartStore` is used as is.

**Load** (docx4j `Load3`): content types, then package relationships, then for every internal
relationship resolve the target part name, create the part through the registry (or link the
already-created part: a part can be the target of several relationships), attach its own `.rels`
part if present, and recurse. Parts not reachable through relationships are not loaded, as in
docx4j and as Word treats them. Part contents are not unmarshalled during load.

**Save** (docx4j `Save`): write `[Content_Types].xml` from the manager; walk the relationship graph
from the package relationships writing each part once: relationship parts and unmarshalled
`XmlPart`s are marshalled, everything else is copied from the source store (or from the bytes set
on the part). `MemoryPartStore` backs new packages.

### 5.6 Markup compatibility (`mc:AlternateContent`, `mc:Ignorable`)

On unmarshal, an `XmlPart` runs an MCE preprocessor over the DOM before handing it to the
objects context: each `mc:AlternateContent` is replaced by the content of the first `mc:Choice`
whose `Requires` namespaces the model understands (`w14`, `w15`, `wp14`, `w16*`, ...), else by
its `mc:Fallback`. This is docx4j's behaviour (its preprocessor) and is what Word itself does on
open; the untaken branch is lost for that part, which is why it happens only for parts that are
unmarshalled. Off switch: `LoadOptions.mcePreprocess = false` (the content then unmarshals as
`org_docx4j_mce` objects where the schema allows it).

On marshal, the root element must declare every prefix named in `mc:Ignorable` even if unused in
the tree, and Word expects the conventional prefixes (`w`, `r`, `wp`, `a`, `pic`, `v`, `o`,
`w10`, `w14`, `w15`, `mc`, ...). `XmlPart` marshals with docx4j's prefix table
(`NamespacePrefixMapper`) and re-declares the ignorable prefixes on the root; see section 9 for
the runtime hook this needs.

## 6. WordprocessingML resolution utilities

### 6.1 `PropertyResolver` (docx4j `org.docx4j.model.PropertyResolver`)

Built from a `WordprocessingMLPackage` (styles, numbering, theme fonts are read once; `refresh()`
after styles change). API and semantics as docx4j:

- `getDocumentDefaultPPr()`, `getDocumentDefaultRPr()` from `w:docDefaults`.
- `getEffectivePPr(pPr)` for a paragraph's direct `pPr`: document defaults, then the default
  paragraph style, then the `basedOn` chain of `pStyle` (cycle detection throws
  `CyclicStylesException`), then the numbering level's `pPr` (`numPr` from the style chain or
  direct; `lvl/pPr` from the abstract definition with `lvlOverride`), then direct formatting.
  `getEffectivePPr(styleId)` for a style alone.
- `getEffectiveRPr(rPr, pPr)`: document defaults, paragraph style chain `rPr`, the run's
  character style chain (`rStyle`), direct `rPr`. Toggle properties (`b`, `i`, `caps`, `strike`,
  ...) follow ECMA-376 17.7.3: applied at the style level they toggle, at direct-formatting level
  they set. `getEffectiveRPr(styleId)`, `getEffectiveRPrUsingPStyleRPr`.
- `getEffectiveTableStyle(tblPr)`: the table style chain; conditional formatting (`tblStylePr`
  per `cnfStyle`) is a later CR.
- `getLvlFromHeadingStyle(styleId)`, `activateStyle(styleId)`, `hasDirectRPrFormatting`.

Property overlay is a `StyleUtil.apply` analogue: property-wise, deep-copying (`deepCopy` from
the runtime) so results never alias the style definitions. Results are plain `PPr`/`RPr` objects
from the object model.

### 6.2 List numbering (docx4j `org.docx4j.model.listnumbering`)

`NumberingDefinitionsPart` exposes `getEmulator()` and the definitions: `ListNumberingDefinition`
per `w:num` (with `lvlOverride` / `startOverride` applied), `AbstractListNumberingDefinition` per
`w:abstractNum`, `ListLevel` per `w:lvl` (`numFmt`, `lvlText`, `start`, `lvlRestart`, `isLgl`,
`rPr` font, `pPr/ind`). `Emulator.getNumber(pkg, pPr)` (and the `(pkg, pStyleVal, numId, ilvl)`
form) returns docx4j's `ResultTriple`: `numString`, `numFont`, `isBullet`, plus the level's `ind`.
Counters are stateful: callers walk paragraphs in document order, as docx4j's consumers do.
Number formats: `decimal`, `decimalZero`, `lowerLetter`, `upperLetter`, `lowerRoman`,
`upperRoman`, `bullet`, `none` in this CR; `decimalEnclosedCircle` and the Chinese formats as in
docx4j later. `Emulator.getInd(pkg, pStyleVal, numId, ilvl)` resolves the indent the same way
`PropertyResolver` does.

### 6.3 Fonts (docx4j `org.docx4j.fonts`)

Two halves, as in docx4j:

- **`RunFontSelector`** decides which *document* font a run (or a character range within it) uses:
  effective `rFonts` (`ascii`, `hAnsi`, `eastAsia`, `cs`) and `hint`, theme references
  (`asciiTheme` and friends resolved through `ThemePart`'s `majorFont` / `minorFont` by script),
  and the script of the characters (Unicode ranges: Latin to `ascii`/`hAnsi`, CJK to `eastAsia`,
  Arabic/Hebrew to `cs` together with `rtl`/`cs` flags). It yields runs of text with a font name
  and bold/italic flags, which is what a renderer, a text extractor or a measurement routine needs.
- **`Mapper`** maps a document font name (plus bold/italic) to a `PhysicalFont`. This CR ships
  `IdentityPlusMapper`: the name itself, with docx4j's metric-compatible substitutes (Arial and
  Liberation Sans, Times New Roman and Liberation Serif, Calibri and Carlito, Cambria and
  Caladea, Courier New and Liberation Mono) and `FontTablePart` `altName`s, which is what a
  browser or CSS consumer wants (`font-family` stacks). A `BestMatchingMapper` over the system's
  installed fonts (Node, `fontkit`) and embedded (`ObfuscatedFontPart`) fonts is a later CR; the
  interface is fixed here so consumers can plug their own. `MainDocumentPart.fontsInUse()` and
  `stylesInUse()` as in docx4j.

## 7. Package layout and exports

```
src/
  opc/        PartName, ContentTypeManager, ContentTypes, PartStore + implementations, PartSink + implementations, Load, Save, mce/
  parts/      Part, BinaryPart, XmlPart, DefaultXmlPart, RelationshipsPart, Namespaces, PartRegistry,
              wml/  pml/  sml/  dml/  docProps/  customXml/
  packages/   OpcPackage, WordprocessingMLPackage, PresentationMLPackage, SpreadsheetMLPackage
  model/      PropertyResolver, StyleUtil, listnumbering/, fonts/
  index.mts   re-exports everything above plus the objects facade
```

`exports`: `.` (all), `./opc`, `./parts`, `./packages`, `./model`, so a bundle for an add-in that
only needs flat OPC and parts does not pull in the numbering and font code. ESM only. The zip
store imports `fflate` statically; the flat OPC store does not, so an add-in bundle that never
touches zip tree-shakes it away.

## 8. Tests

- **Fixtures** under `test/fixtures/`: a handful of `.docx` from `docx4j-samples-resources`
  (Apache-2.0), one saved by a current Word with `w14`/`w15` content and `mc:AlternateContent`, a
  flat OPC capture from `getOoxml()`, one `.pptx`, one `.xlsx`.
- **Round trip:** load, save, reload. Untouched parts must be byte-identical (they come from the
  source store). For a part that was unmarshalled and marshalled unchanged, the re-parsed object
  tree must deep-equal the original (attribute order and namespace prefixes are not part of the
  contract, content is).
- **Flat OPC:** `getOoxml()` capture to package and back; the objects package's
  `unmarshalPackage` remains the typing seam.
- **Parity with docx4j:** golden files produced by a small Java harness (kept under
  `test/java/`, run by hand, output committed): for each fixture and paragraph, docx4j's effective
  `PPr` / `RPr` as XML and the numbering string; the TypeScript resolver must produce the same.
  This is the contract that "docx4j for TypeScript" makes.
- **Word acceptance** is manual: saved output opens in Word without repair prompts. A short
  checklist lives in `test/README.md`.
- CI: Node 18, 20, 22, as the objects repository.

## 9. What the runtime and the objects package need from this

- **Objects facade (done, objects CR-001, 2026-09-10):** the review on 2026-09-10 found that the
  facade built its context without a namespace prefix table, so marshalled output carried
  generated prefixes (`p1:Ignorable="w14"` with no `xmlns:w14`, `p2:space` bound to the `xml`
  namespace), which Word rejects. The engine cannot pass the table itself: the context is built
  by whoever calls `getContext` first. The facade now ships docx4j's table as its default
  (`NAMESPACE_PREFIXES`), makes the relationships namespace the default namespace on a
  `Relationships` root, and strips unused root declarations while keeping those named by
  `mc:Ignorable`. This package therefore has no namespace handling of its own; `XmlPart`
  marshals through `marshalNode`.
- **Runtime (jsonix-CR-003 candidate):** declare-on-root as a marshaller option, so the facade's
  stripping pass can go; `Jsonix.DOM` in the typings (this package and the facade both cast).
- **Objects facade, later:** nothing else; `getContext`, `unmarshalNode`, `marshalNode`,
  `unmarshalPackage`, `deepCopy`, `unwrap` suffice.

## 10. Phasing and effort

| Phase | Content | Effort |
|---|---|---|
| A | `opc/`, `parts/` core, `packages/`, WML typed parts, registry, load/save for zip and flat OPC, MCE preprocessing, round-trip tests | 5 days |
| B | `PropertyResolver` and `StyleUtil`; numbering `Emulator`; `RunFontSelector` and `IdentityPlusMapper`; parity harness and golden files | 6 days |
| C | PML and SML packages with main parts; `DirectoryPartStore`; docs and examples (Node, add-in) | 3 days |

Phase A alone is useful (a typed docx round trip in Node and an add-in package round trip);
each phase ships as a minor version.

## 11. Open questions (decided 2026-09-10)

1. `fflate` versus JSZip: **fflate**. Only `inflateSync`, `zipSync` are used; the zip container
   parses the central directory itself so an entry inflates on first load.
2. Asynchronous `getContents()` plus synchronous `contents`: **as specified**. `unmarshalAll()`
   removes the asynchrony for callers that prefer.
3. MCE preprocessing on by default: **on**. It turned out to be required, not merely
   Word-like: the model types `mc:AlternateContent` only where a global element can follow, so a
   `w:drawing` (a local element) inside `mc:Choice` cannot be unmarshalled at all. Every
   Word-saved document with a shape hits this. With `mcePreprocess: false` such a part throws on
   `getContents()`; the option exists for parts whose branches hold global elements only.
4. Hand-written `[Content_Types].xml`: **hand-written** (`ContentTypeManager`).
5. `PropertyResolver` on the part or the package: **both**, package delegating (Phase B).
6. Name departures: camelCase, noted in the class doc comment. The departures made in Phase A
   are listed in section 12.

## 12. Phase A implementation notes (2026-09-10)

Layout as section 7 (`src/opc`, `src/parts`, `src/packages`; `src/model` comes with Phase B);
`exports` `.`, `./opc`, `./parts`, `./packages`. 22 tests under `test/*.test.mjs` on Node's
runner: part names and content types; loading the docx, pptx and xlsx fixtures; byte-identical
round trips of untouched parts; deep-equal round trips of re-marshalled parts; flat OPC out,
through the objects package's `unmarshalPackage`, and back; a new package; adding, renaming and
removing parts; MCE resolution; prefix and `mc:Ignorable` declarations on re-marshalled output.
The Word acceptance checklist is `test/README.md`; it passed on 2026-09-10 in Word 2016 (untouched
round trip, re-marshalled round trip with a content API edit, a created document, and the
flat OPC opened as a `pkg:package` file), which confirms the prefix table, the `mc:Ignorable`
declarations and the MCE resolution in Word itself.

Departures from the text above, and from docx4j, all deliberate:

- `PartSink<R>` is generic in what `finish()` returns: `ZipPartSink` gives bytes,
  `FlatOpcPartSink` a string, `MemoryPartSink` a `MemoryPartStore`; `OpcPackage.saveTo(sink)`
  returns that. `put` takes the part's content type, which the flat OPC sink needs per part.
- `FlatOpcPartStore` synthesises `[Content_Types].xml` from the per-part `pkg:contentType`
  (a default for `rels`, an override for everything else), so the loader sees one container
  shape. XML parts come back as re-serialised bytes, so a flat OPC load is not byte-identical
  in the way a zip load is.
- `CustomXmlDataStoragePart` extends `DefaultXmlPart` (both hold a DOM). `VMLPart` is a
  `DefaultXmlPart`: docx4j's `org.docx4j.vml.root.Xml` wrapper has a synthetic namespace in the
  generated model and does not match the unqualified `<xml>` root Word writes. Chart style,
  chart colour style, chartEx and diagram drawing parts are DOM parts until typed.
- `ImagePart` is one class; the image kind is the content type. `ChartPart` and `DrawingPart`
  are docx4j's `Chart` and `Drawing`. The PresentationML and SpreadsheetML comments parts are
  `PresentationCommentsPart` and `SpreadsheetCommentsPart`; docx4j's `Styles` and `CalcChain`
  are `StylesPart` and `CalcChainPart`.
- The facade's flat OPC `Part` type (`pkg:part`) is re-exported as `FlatOpcPart`; `Part` is the
  part class, as in docx4j.
- `AddPartBehaviour` is a string union. `RENAME_IF_NAME_EXISTS` appends docx4j's counter to the
  proposed name (`image1.png` becomes `image12.png`), as docx4j does.
- `externalResources` is not implemented: external targets (hyperlinks, linked images) stay in
  the relationships and `getPart(rel)` returns undefined for them. Strict (`purl.oclc.org`)
  packages are not converted on load. ZIP64 archives are rejected.
- The MCE preprocessor's "understood" set is `UNDERSTOOD_NAMESPACES`, a hand-kept copy of the
  generated modules' namespaces (the context does not expose them); `createMcePreprocessor`
  takes another set. Unlike docx4j 3.3.8+, `mc:AlternateContent` inside `w:r` is resolved too
  (question 3 says why).
- `fflate` is imported statically by the zip container and `OpcPackage.load` reaches it, so an
  add-in bundle carries it (about 8 KB); the sentence in section 7 about tree-shaking it away
  does not hold. `fflate` 0.8.3's typings need TypeScript 5.7, hence `skipLibCheck` in the
  typecheck configuration.
- The relationship-source hook (`setRelationshipsPartFactory`) and the package registry
  (`registerPackageClass`) exist to avoid ES module cycles between `Part`, `RelationshipsPart`,
  `OpcPackage` and its subclasses; `packages/index.mts` registers the three Office kinds.

## 13. Phase B execution plan (agreed 2026-09-13)

Phase B is ported by Opus agents, orchestrated from a fresh session, against the **current**
docx4j code on branch `VERSION_17_1_1` (head `325461031` on 2026-09-13; the harness records the
hash it ran against in every golden file's header). That code is not the code section 6
was written against: docx4j reworked all three areas this week, and those CRs are the
reference now, above section 6 where they differ:

- `../docx4j/docs/developer/change-requests/CR-014-list-numbering-model.md` (done
  2026-09-12): `org.docx4j.model.listnumbering`, definitions separated from counter state,
  a `LabelFormatter` registry, `w:lvlRestart`, probes P1 to P8 verified in Word.
- `CR-015-property-resolution.md` (done 2026-09-12): `PropertyResolver` and
  `StyleUtil.apply`, one property catalogue with merge rules, the resolution order, the
  default paragraph style, no mutation or aliasing; five `styles-*` probes with goldens.
- `CR-016-font-selection-and-mapping.md` (done 2026-09-13): `RunFontSelector` and the
  `Mapper`s, one resolution, `w:cs` by value, the theme language, the character-range
  dispatch as a function of the code point, the mapping order, embedded fonts; eight probes
  with goldens and a mapper matrix.

Each of those CRs carries a verification table of claims about Word with the golden that
evidences each; the TypeScript port reproduces docx4j's behaviour as those CRs settled it,
quirks included, and records in this CR's implementation notes any place it knowingly
differs. Where docx4j's own tests and probe documents exercise a rule, they become this
package's fixtures (copied with their provenance).

Sequence, each step a separate agent run with a reviewable diff, nothing committed by the
agent:

1. **The harness first** (`test/java/`, a small Maven project against the docx4j checkout):
   for every fixture and every paragraph, docx4j's effective `PPr` and `RPr` as canonical
   XML (attributes sorted, docx4j's prefixes) and the numbering label, the resolved font
   per run span; output committed under `test/golden/` with the docx4j commit in each
   header. The TypeScript comparison unmarshals the golden XML and compares object trees,
   never text. Option C of 2026-09-12: run by hand, plus a scheduled workflow that reruns it
   against docx4j's head and opens a pull request on a difference. The harness's first
   output is reviewed by a person before anything is measured against it.
2. **`PropertyResolver` and `StyleUtil`** (section 6.1 as revised by CR-015).
3. **`NumberingDefinitionsPart`'s definitions and the `Emulator`** (section 6.2 as revised
   by CR-014).
4. **`RunFontSelector` and `IdentityPlusMapper`** (section 6.3 as revised by CR-016; the
   `BestMatchingMapper` over installed fonts stays a later CR).
5. A review pass, line by line against the Java, on two files: the toggle-property overlay
   in the resolver and the numbering counters. These are where parity failures hide when
   the fixtures do not exercise a branch.

Each agent's prompt carries: this CR (sections 6, 11, 12 and 13), the three docx4j CRs,
`CLAUDE.md`, the Java paths as the reference, the rule that docx4j's quirks are reproduced
and noted rather than corrected, and the rule that decisions and departures go into this
CR's implementation notes. `Font` reads in the content API (CR-002 section 7) switch to
effective values when step 2 lands; `Paragraph.alignment` loses its `'Unknown'`.
