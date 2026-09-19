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
- **Runtime (jsonix-CR-003):** `Jsonix.DOM` and `Context.namespacePrefixes` are typed since
  `@docx4j/jsonix` 3.2.1 (2026-09-15, typings only); `src/xml/dom.mts` uses `Jsonix.DOM` directly
  and its cast is gone. Declare-on-root as a marshaller option (`declareNamespaces`, per-marshal
  prefixes, declaring only where first used), which would let the facade's stripping pass go, is
  deferred to 3.3.0; it is revisited if Phase B's re-marshalling of large parts makes marshal
  time matter. The dependency range is `^0.1.2` of the objects package (2026-09-16), which requires
  3.2.1.
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

## 14. Phase B implementation plan (2026-09-19; refines section 13, nothing implemented)

Section 13 fixed the approach: harness first, then a port against docx4j `VERSION_17_1_1` as
its CRs 014, 015 and 016 settled it, Opus agents with reviewable diffs. This section turns it
into files, APIs, steps and tests, and records what has moved since 2026-09-13. It does not lift
the portfolio rule of 2026-09-15 that gates the work (`docx4j/fo-fidelity-next` first); it is
ready to run the day that is lifted or narrowed.

### 14.1 What changed since section 13

- **docx4j moved.** `VERSION_17_1_1` is at `0077d749c` (2026-09-19). CR-015 gained a toggle-property
  section on 2026-09-17 (ECMA-376 §17.7.3: twelve `w:rPr` booleans XOR across style levels,
  `PropertyCatalogue.TOGGLES`, `StyleUtil.applyStyleLevel`); section 6.1's one-line description of
  toggles is superseded by it. The harness pins the hash it ran against.
- **Python ported the numbering counting core early** (docx4j-python CR-002 section 12.11,
  2026-09-17), because its content API's lists phase needed labels: `definitions.py`,
  `emulator.py`, `formats.py` (1,725 lines), six fixtures (`numbering-*.docx`,
  `styles-numpr-ilvl-only.docx`) cut from docx4j's CR-014 and CR-015 probes, and every label equal
  to Word's. It left for Phase B: resolution through the resolver (it walks `w:basedOn` itself
  meanwhile), the exotic formats, `getInd`'s full form and `labelRPr`. Neither repository has
  built the Java harness; the Python CR says "build it here and share" and this plan builds it
  here, since the fixtures and the comparison test already have a home in `test/`. The Python
  port's module split and its fixtures are reused, so the two ports stay comparable.
- **The consumers now exist.** CR-002 phases B to G and I are implemented and wait on this phase
  at four points: `Font` reads report direct formatting only (section 7 of CR-002),
  `Paragraph.alignment` and the indents report direct values (`'Unknown'` when absent),
  `toApiScript` emits direct formatting only, and phase H (lists) needs `listString`,
  `siblingIndex` and style-derived list membership. The editor's E1 needs the resolver for CSS.
- **The objects package is at 0.1.4** with `deepCopy`, `deepCopyAsSync`, `walk`, `walkAll` and the
  `rPr` element-list helpers. Nothing else is needed from it for this phase; the property
  catalogue is hand-written here as in Java (section 14.3).
- **The Java toolchain is on this machine**: JDK 21, Maven 3.9, and `docx4j-core-17.1.1-SNAPSHOT`
  built in `../docx4j/docx4j-core/target`, so the harness can run against the checkout without a
  Maven Central release.

### 14.2 Scope, file by file

What is ported, into which module, and what is deliberately not. Line counts are the Java at
`0077d749c`.

**Property resolution** (`src/model/properties/`): docx4j `PropertyResolver.java` (927),
`PropertyCatalogue.java` (411), and the catalogue-driven half of `StyleUtil.java` (the `apply`,
`isEmpty`, `unset`, `hasDirectFormatting` family and the per-member merge rules; not its 73
`areEqual` overloads, not `StyleTree`, `Node`, `Tree` and `BrokenStyleRemediator`, which serve
the HTML and CSS export and are a later CR).

- `catalogue.mts`: the tables CR-015 phase 1 introduced, one entry per schema member in schema
  order, `(name, get, set, merge, isEmpty, isFormatting, copyOf)` for `RPr` and `ParaRPr` (one
  table, generic accessors), `PPrBase`, `CTTblPrBase`, `TcPr`; `TOGGLES` named from the run
  table. The leaf merge rules are copied one for one (`spacing` takes `lineRule` only when the
  source states `line`; `lang`, `numPr`, `u`, `highlight` inherit per attribute; enums only when
  stated; every carried leaf is a copy, by a per-type `copyOf`, never `deepCopy`).
- `styleUtil.mts`: `apply`, `applyStyleLevel` (non-toggles by override, toggles by XOR),
  `isEmpty`, `unset`, `hasDirectFormatting`, derived from the tables by iteration.
- `PropertyResolver.mts`: the resolution order of CR-015 phase 2 verbatim
  (`docDefaults ⊕ chain(styleOf(pPr)) ⊕ chain(rStyle) ⊕ direct`; `styleOf` falls back to the
  `w:default` paragraph style; chains merged root-first and cached per style id, without document
  defaults; heading outline level computed into the resolved `pPr`; the 10 pt default size on a
  private copy of the defaults; missing-style lookups logged once; `ancestry(styleId)` with cycle
  detection throwing `CyclicStylesException`). `getEffectiveTableStyle` reports whether the chain
  reaches the default table style, and `WORD_DEFAULT_CELL_MARGIN_TWIPS` is the constant the
  consumer applies, as CR-015 phase 4 decided. Table conditional formatting stays out
  (`docx4j/table-conditions`).

**List numbering** (`src/model/listnumbering/`): the whole package (3,156 lines, most of it the
formatters), in the split Python used so the two ports read alike: `definitions.mts`
(`LevelDefinition` merged from abstract and override, `AbstractListDefinition`, `ListDefinition`,
`NumberingDefinitions` built once from the part with `w:numStyleLink` resolved as
`initialiseMaps` does), `state.mts` (`NumberingState` per story with counters keyed by the
referencing abstract id and level, the first-use flag a reset leaves, `NumberingStates` keyed by
part with headers and footers folded together, a fresh state per text box), `formats.mts`
(`LabelFormatter` registry with docx4j's fail-soft rule; all of docx4j's formatters, since none
is over seventy lines: decimal, decimalZero, lowerLetter, upperLetter, lowerRoman, upperRoman,
bullet, none, ordinal, cardinalText, ordinalText, hex, chicago, numberInDash,
decimalEnclosedCircle, the two Chinese, hebrew1, and the alphabet formats), `Emulator.mts`
(`getNumber(pkg, pPr, state?)`, `peek`, the `(pStyleVal, numId, ilvl)` form, `getInd`, one
`resolve` through the resolver's effective `pPr` with `styleLinkedElsewhere`, `NumberingResult`
with `numString`, `numFont`, `isBullet`, `ind`, `rPr`, `labelRPr`, `lvl`, `notNumbered`).
`NumberingDefinitionsPart` gains `definitions`, `getEmulator()` and `state` (section 6.2).

**Fonts** (`src/model/fonts/`): the selection and mapping core, not the rendering half.
docx4j's `RunFontSelector.java` (2,548) is mostly FO output: element creation, kerning and
ligature suffixes, character scaling, small caps, line metrics, the glyph-coverage pass over
physical fonts. What section 6.3 promised and what a browser, an extractor or a measurer needs is
the decision per character range, so the port takes lines 173 to 320 and 1,695 to 2,240 of it
(about 700 lines): `documentFontsOf`, `defaultFontOf`, the theme lookup by `themeFontLang`,
`resolvedSlots`, `complexScriptFont` (`w:cs` and `w:rtl` by value), `preambleRule`, `spanScript`
(the code point to slot table, ported verbatim), `isEmoji` by code point, `symbolFontName` and
the symbol segments, the `hint` rules, `LanguageTagToScriptMapping` (122), `CJKToEnglish` (41).
Its result is `FontSpan[]` (`text`, `documentFont`, `bold`, `italic`, `complexScript`, `rtl`,
`script`), plus `documentFontFor(pPr, rPr, codePoint)`. The mapper side: `Mapper.mts` (the
precedence template of CR-016 phase 3: installed or embedded first, the mapper's own step,
metrically compatible substitutes, the `altName` chain to its class, class-based substitutes from
`w:family` and `w:panose1`, the unresolvable face), `IdentityPlusMapper.mts`, `PhysicalFont`
as a name with its family and the source of the decision (no files, no FOP `EmbedFontInfo`),
`substitutions.generated.mts` from `font-substitutes.xml` (494 lines, converted by a script
under `scripts/`) and the no-bold-face family list from `fonts/microsoft`. `BestMatchingMapper`,
glyph checks, embedded font metrics and `WordLineMetrics` stay a later CR, as section 6.3 says;
the `Mapper` interface is what a consumer with `fontkit` plugs into. `MainDocumentPart.fontsInUse()`
(the walk for names of CR-016 phase 4: all four slots on runs, marks, styles in use, headers,
footers, notes, comments, text boxes, numbering levels, theme references resolved, `w:sym`
fonts, the default) and `stylesInUse()`.

**Wiring into what exists.** `WordprocessingMLPackage.getPropertyResolver()` (asynchronous,
because it unmarshals the styles, numbering and theme parts once; `propertyResolver` synchronous
afterwards, the pattern of `getBody()` / `body`), `refresh()` called by the content API's style
mutations (`Comment`'s style creation, phase H's `startNewList`, `insertOoxml`). In CR-002:
`Font` reads switch to effective values through `getEffectiveRPr(rPr, pPr)` with a
`{ direct: true }` option keeping today's behaviour; `Paragraph.alignment`, the indents and
spacing read effective values and `'Unknown'` goes; `toApiScript` stays on direct formatting
(a script reproduces markup, not appearance) and says so; phase H is unblocked.

### 14.3 The harness

`test/java/`, a Maven project (`pom.xml`, one `Harness.java`, Java 21) depending on
`org.docx4j:docx4j-core:17.1.1-SNAPSHOT` from the local repository, so it runs against the
checkout at any commit; the commit hash, the date and the harness version go into every golden's
header. Run by hand with `mvn -q exec:java -Dfixtures=../fixtures -Dout=../golden`.

Input: every `.docx` under `test/fixtures/` plus the probe documents, copied in with provenance:
the five `styles-*` and the fifteen `fonts-*` probes that `docx4j-layout-fidelity`'s `Corpus`
generates (`Fidelity generate`, then copy from `target/corpus`), the six numbering fixtures
Python already cut, and from `docx4j-core-tests`: `numbering_indentation*.docx`,
`NumberingImplicitNumId.docx`, `article-section-*.docx`, `startOverride.docx`,
`numbering-stories.docx`, `NumberingIndents.docx`, `Mac_OSX_Fonts.docx`,
`fonts-modesOfApplication.docx`. The Word PDF goldens are not needed: the contract here is
docx4j's answer, which the three CRs verified against Word.

Output, one `test/golden/<fixture>.json` per document (JSON rather than the XML files section 8
imagined, so that one file carries the mixed content; the property values inside it are XML
strings marshalled by docx4j with its prefixes, which the TypeScript side unmarshals through the
facade and compares as object trees, never as text):

- header: docx4j commit, date, fixture name and size;
- `styles`: for every style id in the part, `effectivePPr(styleId)` and `effectiveRPr(styleId)`;
  `defaultParagraphStyleId`; the document defaults;
- `stories`: for the main document, each header and footer, footnotes, endnotes, comments: per
  block-level paragraph in document order (ordinal address, `w14:paraId` when present,
  `pStyle`): `effectivePPr`, `paragraphMarkRPr`, `numbering` (`numString`, `isBullet`,
  `numFont`, `ind`, `ilvl`, `labelRPr`, `notNumbered` and its reason, plus, at Python's
  request of 2026-09-19, docx4j's `NumRef` (`numId`, `ilvl`, whether the `w:numPr` was direct
  or came from the style) and the story's `NumberingState` after the paragraph (each counter's
  value per abstract id and level, and which start overrides have been spent)) with one
  `NumberingState` per story as CR-014 defines them; per run: `effectiveRPr`, and
  `fontSpans` from `RunFontSelector.documentFontFor` per code point folded into spans
  (`documentFont`, `bold`, `italic`, `cs`, `rtl`, and the `IdentityPlusMapper` result with the
  document's font table);
- `tables`: per table, `effectiveTableStyle` and whether the chain reaches the default.

`test/parity.test.mjs` loads each golden, runs this package's resolver, emulator and selector
over the same fixture, and deep-equals per paragraph with `plain()`; a failure prints the
paragraph address and the first differing property. The first harness output is read by a
person before anything is measured against it (section 13). Option C: `.github/workflows/parity.yml`
on a weekly schedule checks out `plutext/docx4j` `VERSION_17_1_1`, builds `docx4j-core`
(`mvn -q -pl docx4j-core -am -DskipTests -Dgpg.skip`), runs the harness into a temporary
directory, and opens a pull request when a golden differs, with the docx4j commits since the
recorded hash in its body. Python consumes the committed goldens by path or copy; its CR-002
section 8 is amended to say the harness lives here.

### 14.4 Steps

Each step is one agent run in a worktree (the workflow of CR-002 phases C to I: symlinked
`node_modules`, an integration branch, a squash into main uncommitted, file ownership per step),
reviewed and committed before the next depends on it. Step 1 must precede everything; step 2
precedes 3 and 4, which run in parallel; step 5 closes.

| Step | Deliverable | Owns | Tests | Effort |
|---|---|---|---|---|
| 0 | Fixture inventory: the probe documents generated and copied with provenance into `test/fixtures/`; `test/README.md` lists them | `test/fixtures/`, `test/README.md` | none | 0.5 day |
| 1 | The harness and the first goldens (14.3), the parity test skeleton that loads goldens and reports "not implemented" per area, the weekly workflow | `test/java/`, `test/golden/`, `test/parity.test.mjs`, `.github/workflows/parity.yml` | the goldens read by a person | 1 day |
| 2 | Property resolution (14.2), `getPropertyResolver()`, the `Font` and `Paragraph` switch in CR-002 with the `direct` option, `CyclicStylesException` | `src/model/properties/`, `Font.mts`, the paragraph property getters, `WordprocessingMLPackage.mts` | parity on `effectivePPr`, `effectiveRPr`, `paragraphMarkRPr`, `styles`; a table-driven catalogue test as CR-015's `PropertyCatalogueTest`; the order test (its S1 to S15 cases); no-mutation (styles part byte-identical after resolving everything); the toggle cases of CR-015 | 2 days |
| 3 | List numbering (14.2) on the resolver; `NumberingDefinitionsPart` accessors; the default numbering resource for phase H | `src/model/listnumbering/`, `NumberingDefinitionsPart` | parity on `numbering` per story; the formatter table test; `w:lvlRestart`, `startOverride`, stories, `isLgl`, `numStyleLink` from the docx4j tests; Python's six fixtures give the same labels | 1.5 days |
| 4 | Fonts (14.2): selector core, `Mapper`, `IdentityPlusMapper`, the substitution data, `fontsInUse`, `stylesInUse` | `src/model/fonts/`, `scripts/generate-substitutions.mjs`, `MainDocumentPart` | parity on `fontSpans`; the CR-016 phase 1 and 2 unit cases (`cs` by value, theme language, no `rFonts`, the range dispatch per script, symbol fonts, emoji, `hint`), `MapperPrecedenceTest`, `AltNameChainTest`, `NoBoldFaceTest` ported | 2 days |
| 5 | Line-by-line review of the toggle overlay and the numbering counters against the Java (section 13 step 5); implementation notes as section 15; CLAUDE.md, README, registry; CR-002's status line for the switched reads | docs | the whole suite; parity zero-difference | 0.5 day |

Seven and a half days against section 10's six: the toggle rule, the harness's stories and font
spans, and the formatter set account for the difference.

### 14.5 Decisions to take before step 1 (all five decided by Jason 2026-09-19, as recommended)

1. **The gate.** The portfolio rule blocks the port. The numbering core was ported to Python on
   2026-09-17 under it, for the same reason phase H needs it here. Three options: wait for
   `docx4j/fo-fidelity-next` to be scoped and done; lift the rule for this phase (its Java
   sources were reworked and Word-verified within the last week, which is the state the rule
   wants); or lift it for step 3 only, as Python did, and run steps 0 and 1 now since a harness
   is not a port. The recommendation is the second: the three CRs are the settled reference and
   the harness catches drift.
2. **Golden format**: JSON with XML strings inside (14.3), not the XML files of section 8.
   Recommended.
3. **Where the harness lives**: here, Python consuming (14.3). Recommended; it reverses the
   Python CR's sentence and needs a line there.
4. **`Font` and `Paragraph` reads become effective by default**, with `{ direct: true }` for the
   old behaviour, and `toApiScript` stays direct. This is what Office JS reports and what CR-002
   section 7 promised; it changes the values existing callers see, so it is a minor-version note.
5. **Fonts scope**: the selection core and `IdentityPlusMapper` only (14.2); the glyph-coverage
   pass, `BestMatchingMapper` and metrics are a later CR with `fontkit`. As section 6.3, restated
   against the Java as it now is.

Decided 2026-09-19: the gate is lifted for this phase (`core-ts/CR-001.B` no longer depends on
`docx4j/fo-fidelity-next` in the portfolio registry; the rule stands for other ports, Python's
Phase B included until it is decided there); goldens are JSON with XML strings; the harness
lives here and Python consumes the committed goldens (docx4j-python CR-002 section 8 to be
amended in that repository); `Font` and `Paragraph` reads become effective by default with
`{ direct: true }`, `toApiScript` stays direct; fonts are the selection core and
`IdentityPlusMapper`. Step 0 can start.

### 14.6 Implementation notes: step 1 (2026-09-19)

The harness (`test/java/`), the first 45 goldens (`test/golden/`), the parity test skeleton
(`test/parity.test.mjs`) and the weekly workflow (`.github/workflows/parity.yml`) are in.
`test/java/README.md` documents the harness and `test/golden/README.md` reviews the goldens;
what follows is what was decided while building them.

- **Input.** Every `.docx` directly under `test/fixtures/` (the eight real Word documents) as
  well as the 37 probes under `test/fixtures/parity/`: 45 goldens, 2.6 MB, three seconds.
  618 paragraphs, 1,113 runs, 1,112 font spans, 984 style resolutions, 8 tables, 185 numbered
  paragraphs. Nothing is truncated and nothing threw (every `notes` array is empty).
- **The golden's shape** is 14.3's, with four additions the port will want: `docDefaults` (the
  element as the styles part states it), `indResolved`
  (`NumberingDefinitionsPart.getInd`, which follows a level's linked style, beside the
  deprecated `NumberingResult.ind`), `fonts.themePart` / `fonts.defaultFont` (a theme
  reference that cannot be answered changes every font span in a document, so it is recorded
  rather than implied), and `header.docx4jCoreJarSha256`.
- **`docx4jCoreJarSha256`** exists because one `~/.m2` holds one `17.1.1-SNAPSHOT`: the hash
  says which build a golden actually came from, beside the commit the harness was *told*
  about. The weekly workflow ignores it, and the header's date, when it diffs.
- **The XML is trimmed of unused namespace declarations.** docx4j's prefix mapper
  pre-declares all ninety-odd Office namespaces on whatever it marshals: 3 kB of `xmlns:` on
  each of tens of thousands of fragments, ten times the size of the goldens (26 MB before,
  2.6 MB after). Nothing that is compared changes, since the comparison unmarshals.
- **Determinism.** `docx4j.fonts.discoverPhysicalFonts.enabled=false` before anything builds a
  `Mapper`, jar discovery left on, and exactly the symbol, croscore and crosextra font jars on
  the classpath: CR-016 phase 0c's `-Dfidelity.fonts=jars` environment. An empty registry would
  also be deterministic but would make every mapping `UNMAPPED`; this way the goldens exercise
  every pass of CR-016 phase 3's precedence template (metric clone, `w:altName`, class, Word's
  default, unresolvable), which is what step 4 ports. The font cache goes to a fresh temporary
  directory per run. Verified: two runs into two directories differ only in the header's date.
- **`mc:AlternateContent` is resolved to its first `mc:Choice`** in the harness's walk.
  `TraversalUtil` walks the choices *and* the fallback, which put a text box's paragraphs in a
  golden twice; docx4j's own `mc-preprocessor.xslt` does not help, since it *keeps* an
  `mc:AlternateContent` whose parent is a `w:r` — which is where Word puts a text box — and
  prefers the fallback elsewhere. This package resolves every `mc:AlternateContent` on the DOM
  before unmarshalling, taking the first choice whose `Requires` namespaces are in
  `UNDERSTOOD_NAMESPACES` (`wps` among them), so the harness's rule is the one that makes the
  two walk the same document. A choice requiring something neither side knows would diverge;
  no fixture has one, and the unmarshalled `Requires` no longer carries the prefix bindings
  needed to do better.
- **The story rule**, read off the Java rather than guessed: `NumberingStates.forPart(part)`
  (body; one state shared by every header and footer; one per footnotes, endnotes and comments
  part) and `NumberingStates.newStory()` on entering a text box, which is what
  `AbstractWmlConversionContext.enterTextBox` does. A text box's paragraphs stay in the
  containing part's list, in document order, and count on their own; the containing story runs
  past them untouched. `numbering-stories.json` shows body 1 2 3, text box 1 2 3, body 4 5 6,
  header 1 2 3, footer 4 5 6 — docx4j's own `NumberingStoriesTest`, CR-014 probe P7.
- **`reachesDefaultTableStyle` has no API.** It is `getEffectiveTableStyle`'s `builtIn` local
  (the chain is empty, or contains the default table style id); its only public sign is the
  cell margins in the resulting style, which is what `AbstractTableWriter` reads. The harness
  recomputes it from `PropertyResolver.ancestry`, which is private.
- **Three reflections** in harness version 1 — `Emulator.resolve` for `NumRef`,
  `NumberingState`'s private maps and `ListLevel.Counter`'s fields for `stateAfter`, and
  `PropertyResolver.ancestry` for `reachesDefaultTableStyle` — plus a no-op
  `RunFontCharacterVisitor` of its own. All four were requested of docx4j and granted; see
  the regeneration note below.
- **`w:delText` is its own class**, not a `w:t`: a walk that only reads `org.docx4j.wml.Text`
  silently loses a deleted run's text. Worth remembering for the port's own traversal.
- **Regenerated from the merged branch (2026-09-19, harness version 2).** docx4j merged
  CR-001 batch 49 into `VERSION_17_1_1` at `d5809a1d8`, which makes all four parity
  accessors public (`Emulator.numRefFor`, `NumberingState.counters()` /
  `startOverridesApplied()` with a public `ListLevel.Counter`,
  `PropertyResolver.reachesDefaultTableStyle`, `FontsAnalysis.NO_OP_VISITOR`, all covered by
  its `ParityAccessorsTest`). The harness calls them, is 100 lines shorter and holds no
  reflection; run side by side, version 2's goldens are **identical to version 1's on all 45
  fixtures**, so the accessors and the reflection agree on `numRef`, `stateAfter` and
  `reachesDefaultTableStyle` everywhere. The batch also settles the themeless default —
  `docx4j.fonts.defaultTheme` = "2023", Word 365's Aptos — which is the behaviour the
  provisional goldens had picked up from an unmerged build, so the caveat on those sixteen
  goldens is discharged rather than encoded. Against the provisional set the only substantive
  change is `fonts.mapping` in 18 goldens: Aptos and Aptos Display were `UNMAPPED` and are now
  `METRIC_CLONE` to Akasia Regular and Intos Display Regular, from the new
  `docx4j-export-fo-fonts-theme2023` jar, which joins symbol, croscore and crosextra in the
  harness's font environment and in the weekly workflow's build list. No effective property,
  numbering label, counter, table flag or font span moved.
- **Step 2 and after** read the goldens through `test/parity.test.mjs`, which today proves
  only that every recorded fragment unmarshals through the facade (each inside the container
  the schema puts it in — `w:pPr` in a `w:p`, `w:lvl` in a `w:abstractNum`, and so on, since
  the runtime resolves global elements only) and that each header names a fixture of the
  recorded size. The comparisons are `test.todo`s naming the step that turns them on.

**The themeless default, and what it means for steps 4 and Phase A (2026-09-19).** The docx4j
session confirmed that the Aptos behaviour the 16 goldens carry is docx4j CR-001 batch 49 item 6,
mid-implementation. Names final as of 2026-09-19 (docx4j branch commits a309177d6, 5e8b2347b,
04e6e6936, 1316f21c2, merging as a merge commit a day or two later): property
`docx4j.fonts.defaultTheme` with values `"2023"` (the default: Aptos Display / Aptos, what Word 365
supplies to a package with no theme part), `"2013"` (Calibri Light / Calibri) and `"2007"`
(Cambria / Calibri, the answer at branch HEAD before the merge), a `Docx4jProperties.DefaultTheme`
enum `THEME_2023` / `THEME_2013` / `THEME_2007` carrying the two face names, theme resources
`org/docx4j/openpackaging/parts/WordprocessingML/theme-2023.xml`, `theme-2013.xml`,
`theme-2007.xml`, `RunFontSelector.themeFont` answering the themeless case from the property, and
`WordprocessingMLPackage.createPackage` adding a `ThemePart` from the matching resource. Font
substitutions added to `font-substitutes.xml`: Aptos to Akasia Regular (metric), Aptos Display to
Intos Display Regular (metric), Aptos Light to Akasia Light, Aptos Narrow to Arimo (class), Carlito
behind each; a new module `docx4j-export-fo-fonts-theme2023` carries Akasia and Intos Display
(the harness classpath adds it at regeneration so those substitutes resolve).
Two consequences here: step 4's selector implements the same default under the same names (a
loaded document with no theme part resolves to the 2023 faces; a golden recording themeless
behaviour must therefore be a loaded document, never a created one), and `createPackage` here
gains the same three theme resources and the same default so that a document created by either
library resolves alike (a Phase A departure to close in step 4, recorded in section 12 then). The goldens
are regenerated at the merge, from the commit range the docx4j session sends.

**`w:numId` 0 (for step 3; from docx4j-python, 2026-09-19).** docx4j's `Emulator.resolve` does not
special-case a direct `w:numPr` whose `w:numId` is 0 (ECMA-376: numbering turned off): `numRefFor`
answers `numId "0"`, `notNumbered false`, and `getNumber` then finds no list 0 and returns an empty
result, so the paragraph is unnumbered in effect but the `NumRef` says otherwise.
`numbering-label-ilvl0.json` shows it. Python answers `notNumbered` with reason "numbering turned
off" inside its resolve, because its `isListItem` and `detachFromList` are written over that. The
TypeScript port takes Python's reading (the two ports must agree, and CR-002 phase H has the same
two callers). Reported to docx4j and fixed the same day (`01d661547`: `numRefFor` answers
`notNumbered` with the reason "the paragraph's w:numId 0 turns numbering off", or the style's
when the 0 comes from the chain, and `getNumber` returns null); the goldens were regenerated
from that commit, and only `numbering-label-ilvl0.json`'s two such paragraphs changed (their
`numbering` is now null). No exception is needed in either port.

## 15. Phase B implementation notes

### 15.1 Step 2: property resolution (2026-09-19)

`src/model/properties/` (`catalogue.mts`, `styleUtil.mts`, `PropertyResolver.mts`,
`numberingInd.mts`, `log.mts`, `index.mts`), `WordprocessingMLPackage.getPropertyResolver()`,
the `Font` and `Paragraph` switch of decision 4 in section 14.5, the comparisons in
`test/parity.test.mjs` and `test/properties.test.mjs` are in. Ported against docx4j
`VERSION_17_1_1` at `01d661547` (the commit the goldens carry): `PropertyResolver.java` (942),
`PropertyCatalogue.java` (411) and the `apply` / `applyStyleLevel` / `isEmpty` / `unset` /
`hasDirectFormatting` half of `StyleUtil.java`. **Parity is zero differences** on all 45
goldens, over 1,647 `w:pPr` and 2,760 `w:rPr` comparisons and the 8 table styles.

**What is in.**

- `catalogue.mts`: `RUN` (52 members of `w:rPr` and `w:pPr/w:rPr` through one set of
  accessors, since the two generated interfaces have the same member names), `PARAGRAPH` (34
  of `PPrBase`), `TABLE` (17 of `CTTblPrBase`) and `CELL` (13 of `TcPr`), each entry
  `(name, isFormatting, countsEmpty, get, set, merge, isEmpty, copyOf)` in schema order;
  `TOGGLE_NAMES` / `TOGGLES`; every leaf merge rule of `StyleUtil`, copied one for one
  (`w:spacing` takes `w:lineRule` only beside a `w:line` the source states; `w:lang`, `w:numPr`,
  `w:u`, `w:highlight`, `w:tblpPr`, `w:framePr`, `w:shd`, `w:bdr`, `w:color` per attribute;
  `w:ind`'s firstLine/hanging as one property; `w:tabs` cumulative with `clear` removing;
  enums only when stated). A carried leaf is always a copy, by a per-entry `copyOf`
  (`copyLeaf`, a small structural clone that drops `PARENT`), never the facade's `deepCopy`.
- `styleUtil.mts`: `applyRPr`, `applyPPrBase`, `applyPPr`, `applyTblPr`, `applyTcPr`,
  `applyTrPr`, `applyTblStylePrList`, `applyStyle`, `applyStyleLevel`, `applyToggles`,
  `toggle`, the `isEmpty*` family, `unsetRPr` / `unsetPPrBase`, `hasDirectFormattingRPr` /
  `hasDirectFormattingPPr`, `isCyclic`. Not ported: the `areEqual` family, `StyleTree`,
  `Node`, `Tree`, `BrokenStyleRemediator` (docx4j's HTML/CSS export; a later CR).
- `PropertyResolver.mts`: the resolution order of CR-015 phase 2 verbatim, `getChainPPr` /
  `getChainRPr` cached per style id without document defaults, `ancestry`, the heading outline
  level by `w:name` computed into the resolved `w:pPr`, the 10 pt default on a private copy of
  the document defaults, missing styles logged once per id, `getEffectiveTableStyle` /
  `reachesDefaultTableStyle` / `WORD_DEFAULT_CELL_MARGIN_TWIPS`, `refresh()`,
  `headingLevelByName`, `getLvlFromHeadingStyle`, `getStyle`, `activateStyle`.
  Table conditional formatting stays out (docx4j's `table-conditions` CR).

**Departures from the Java, and why.**

1. **`CyclicStylesException` throws only when asked**, as docx4j does: its `isCyclic` reads
   `docx4j.openpackaging.exceptions.CyclicStylesException.throw`, false by default, and
   otherwise stops the walk and resolves with as much of the hierarchy as it has. Here that
   property is the function `throwOnCyclicStyles(true)`. Section 6.1's "cycle detection throws"
   is the property's `true` branch, not the default.
2. **`activateStyle(styleId)` cannot activate a style this package does not hold.** docx4j
   reads `KnownStyles.xml` from its jar; there is no such resource here yet, so the string form
   answers `true` for a style already live and `false` (logged) otherwise. `activateStyle(style)`
   with a `w:style` object works as docx4j's does, recursively for `w:basedOn` and `w:link`.
   A `KnownStyles.xml` equivalent belongs with the content API's style creation, not here.
3. **The deprecated overloads are not ported**: the four-flag `getEffectiveRPr(styleId, ...)`
   and `getEffectiveRPrUsingPStyleRPr`, which CR-015 phase 5 deprecated and which exist only
   for callers older than the fix. `getEffectiveRPr(styleId)` and `getChainRPr(styleId)` are
   what they were kept for.
4. **`ImmutablePropertyResolver` is not ported** (deprecated, no caller).
5. **Concurrency is not a concern here.** CR-015 phase 3 made the caches `ConcurrentHashMap`s;
   JavaScript has one thread per realm, so `Map` is the same thing. What that phase also
   removed - the mutations - is removed here, and the no-mutation test is ported.
6. **The resolver reads the styles and numbering parts with a new `XmlPart.readContents()`**,
   which unmarshals into a *private* tree without marking the part unmarshalled. Without it,
   building a resolver would cost every document its byte-for-byte round trip of
   `word/styles.xml` and `word/numbering.xml`, since this package re-marshals any part that was
   unmarshalled (CLAUDE.md, "untouched parts round-trip byte for byte") - and `getBody()` now
   builds one. When something else unmarshals the part (the comment styles, a caller's
   `getContents()`), `refresh()` switches to that live tree, so a change made through the part
   is seen and the private copy is dropped. `activateStyle(style)` throws while the styles part
   is not unmarshalled, since it would otherwise write into the private copy.
7. **`NumberingLevels` (`numberingInd.mts`) is step 3's work done narrowly.** The paragraph
   merge folds in the numbering level's `w:ind` per layer (`StyleUtil.apply(PPrBase, PPrBase,
   NumberingDefinitionsPart)`), so parity on `effectivePPr` is impossible without
   `NumberingDefinitionsPart.getInd`. What is ported is that walk alone - the instance
   definitions, `w:lvlOverride`, `w:numStyleLink`'s second pass, and `getIndFromLvl`'s rule
   that the level's own `w:ind` comes before the linked `w:pStyle`'s - read off the Java.
   Step 3 replaces it with the real `definitions.mts` and moves `getInd` onto
   `NumberingDefinitionsPart`.
8. **Three flags on catalogue entries record docx4j quirks rather than hiding them**:
   `countsEmpty: false` for `w:bidiVisual` (applied by `apply(CTTblPrBase)` but not counted by
   `isEmpty(CTTblPrBase)`), and both that and a "keep the destination" merge for
   `w:tblCaption` and `w:tblDescription`, which docx4j's hand-written pair neither counts nor
   carries. Likewise `w:start`/`w:end` on `w:tblCellMar` and `w:tl2Br`-less `w:tcBorders`
   members follow the Java's lists, not the schema's.

**Quirks reproduced deliberately.**

- **`w:framePr` gains an explicit `w:anchorLock`.** docx4j reads it through `isAnchorLock()`,
  which answers `true` when the attribute is absent, and then applies that answer, so any
  merged `w:framePr` carries `w:anchorLock="true"` even when no layer stated it. The same
  reading makes `isEmpty(CTFramePr)` treat an explicit `w:anchorLock="0"` as the only
  attribute that makes a frame non-empty on its own. Reproduced; the catalogue test allows for
  it by name.
- **A `w:rFonts` that states only a `w:hint`, or nothing at all, still creates the element** in
  the destination (docx4j's `apply(RFonts)` comment cites `RunFontSelectorChinese2Test`); a
  *null* source leaves the destination alone, which is the CR-015 phase 1 fix.
- **`isEmpty(SectPr)` is "anything non-null is non-empty"**, which is all `isEmptyPPr` uses it
  for; `w:sectPr` is not merged (CR-015 phase 4).
- **`getEffectivePPr(styleId)` of a *character* style** answers the document defaults plus that
  style's (empty) paragraph chain, not `undefined`: the golden records it for every style in
  the part and we match it.
- **`getEffectiveRPr(styleId)` of a missing style is `undefined`**, where
  `getEffectivePPr(styleId)` of a missing style is the default paragraph style's.

**Goldens questioned: none.** Every recorded value is reproduced. Nothing was left failing.

**The content API switch (section 14.5 decision 4).** `Font` reads go through
`getEffectiveRPr(rPr, pPr)` of the first run in scope; writes are unchanged and still direct.
`Paragraph.alignment`, `leftIndent`, `rightIndent`, `firstLineIndent`, `spaceBefore`,
`spaceAfter`, `lineSpacing` and `outlineLevel` read `getEffectivePPr(pPr)`. The `direct`
option's final shape:

| read | effective | direct |
|---|---|---|
| a run's font | `paragraph.font`, `range.font` | `paragraph.getFont({ direct: true })`, `range.getFont({ direct: true })` |
| a paragraph's properties | the eight getters, `paragraph.formatting()` | `paragraph.formatting({ direct: true })` |
| the resolved objects themselves | `paragraph.effectivePPr`, `paragraph.effectiveParagraphMarkRPr` | `paragraph.p.pPr` |

`Alignment` loses `'Unknown'`: an absent `w:jc` after resolution reads `'Left'`, as Word lays
it out and as Office JS answers. `'Unknown'` remains in `AlignmentOrUnknown`, which is what the
direct read returns and what the setter accepts (where it removes the element). `Font.name` is
the one read that is not yet fully effective: it is `w:rFonts/@w:ascii` of the resolved
properties, so a run whose face comes from the theme still reads `''` until step 4's
`RunFontSelector`. `toApiScript` keeps reading the tree directly and says so at
`paragraphProperties`: a script reproduces markup, not appearance, and emitting effective
values would write a style's alignment, indents and size onto every paragraph and run of the
generated document.

`Body.propertyResolver` throws `PropertyResolverNotCreatedException` - naming
`getPropertyResolver()` and the `{ direct: true }` option - when the body has no package or the
package has not built a resolver. `MainDocumentPart.getBody()`, `HeaderPart.getBody()` and
`FooterPart.getBody()` build one (structurally, through `part.package`, so no import cycle),
and `createPackage()` awaits one before returning, since its styles part is set as XML and a
synchronous build is not possible. `ensureCommentStyles` calls `refreshPropertyResolver()`
after adding a style; `insertOoxml` does not merge styles (CR-002 open question 5) and needs no
refresh.

**Tests.** `test/properties.test.mjs` (32): catalogue completeness against the runtime's own
mapping model for all five types, the table-driven per-member case for all four tables, the
merge rules, the toggle cases of CR-015's 2026-09-17 section and its golden, the resolution
order (docx4j's `PropertyResolverOrderTest` in full), cycles, headings, `refresh()`,
no mutation (the styles part marshalled before and after resolving everything, byte-equal), the
table-style probe and the wiring. `test/parity.test.mjs` compares every recorded property of
all 45 goldens by marshalling our object through the facade and unmarshalling *both* strings,
one batch per kind per fixture, so neither prefixes nor attribute order nor `TYPE_NAME` can
differ. Suite: 256 tests, 253 pass, 0 fail, 3 `todo` (steps 3 and 4).

**For the objects package and the runtime: nothing.** The facade's `deepCopy`, `marshalString`,
`unmarshalNode` and the generated factories covered everything this step needed. Two
observations for a later CR there, neither blocking: `Jsonix.Util.deepCopy` allocates a `Map`
per call, which is why leaf copies here are a local `copyLeaf`; and the mapping model's
`getTypeInfoByName(...).properties` (used by the catalogue-completeness test) is not in the
runtime's typings, as `Jsonix.DOM` was not.
