# CR-001: The engine: container interface, Open Packaging layer, typed parts, resolution utilities

**Status:** Phase A implemented 2026-09-10 (both steps); Phase B implemented 2026-09-19 (plan in
section 14, notes in section 15; parity with docx4j VERSION_17_1_1 7fba7a150 on 45 goldens);
Phase C implemented 2026-09-19 (section 16)
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
`org_docx4j_mce` objects with each branch kept as DOM and written back on save; since objects
0.1.5 the schema admits the element wherever Word writes it, section 17.5).

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
- **Runtime (jsonix CR candidate, found by Phase B step 4 on 2026-09-19):** `@xmldom/xmldom`
  0.9.12, the DOM the runtime injects in Node, applies XML 1.1's end-of-line normalisation to an
  XML 1.0 document: a raw U+0085 (NEL) or U+2028 (LINE SEPARATOR) in text content is parsed as
  U+000A (a bare U+000D and CRLF become U+000A too, which XML 1.0 does prescribe; a character
  reference `&#x85;` survives). Xerces keeps both, so docx4j's goldens hold them and the parity
  test normalises around the one occurrence (`tracked-changes.docx`). The consequence outside
  tests is that a part containing either character loses it on unmarshal, so a re-marshalled part
  is not what Word wrote. Browsers' `DOMParser` is correct. The fix belongs in the runtime: an
  upstream xmldom fix with the version pinned, or `Jsonix.DOM.parse` escaping the two characters
  as references when the declaration is 1.0. Sent to the jsonix session on 2026-09-19.
- **Objects facade, later:** nothing else; `getContext`, `unmarshalNode`, `marshalNode`,
  `unmarshalPackage`, `deepCopy`, `unwrap` suffice.

## 10. Phasing and effort

| Phase | Content | Effort | Actual |
|---|---|---|---|
| A | `opc/`, `parts/` core, `packages/`, WML typed parts, registry, load/save for zip and flat OPC, MCE preprocessing, round-trip tests | 5 days | implemented 2026-09-10 |
| B | `PropertyResolver` and `StyleUtil`; numbering `Emulator`; `RunFontSelector` and `IdentityPlusMapper`; parity harness and golden files | 6 days | 7.5 days as re-planned in 14.4 (steps 0 to 5, none of which recorded its own); implemented 2026-09-19 |
| C | PML and SML packages with main parts; `DirectoryPartStore`; docs and examples (Node, add-in) | 3 days | implemented 2026-09-19 (notes in section 16; `createPackage()` for both, `clone()`, the `./node` subpath, four guides and six examples) |

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
   Word-saved document with a shape hits this. With `mcePreprocess: false` such a part threw on
   `getContents()` until objects 0.1.5 admitted the element in the WordprocessingML hosts
   (section 17.5); it now unmarshals with both branches kept as DOM, which the content API
   cannot see into, so resolving on load stays the default.
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
- Three more departures came out of the acceptance run of 2026-09-19 and are recorded in
  section 17: `createPackage()` writes `w:compat/w:compatSetting compatibilityMode` 15, which
  docx4j does not; a created presentation's layout and slides carry placeholder shapes, which
  docx4j's template markup does not; and `XmlPart` re-declares `mc:Ignorable` prefixes on
  marshal, which section 5.6 assumed the facade would do alone.

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

**docx4j CR-021 phase 1 (2026-09-19, docx4j `c9612931c`, for the harness and step 3).** docx4j's
`TraversalUtil` now visits one branch of each `mc:AlternateContent` (`McMode.READ`, the default;
`ALL` for mutators), chosen by `org.docx4j.jaxb.McSelection`: the first `mc:Choice` whose
`Requires` prefixes are all named in the property `docx4j.jaxb.mc.preferChoice`, else the
`mc:Fallback`. The property's default is empty, so docx4j's default reads the Fallback (the VML
branch of a text box), where this package's DOM preprocessor (section 5.6) takes the first
`mc:Choice` whose prefixes are in `UNDERSTOOD_NAMESPACES`, which is what Word draws. The two
branches differ in real documents (docx4j's own probes `mc-textbox-branches-*` show it). Decision:
this package keeps the Choice (it types the drawing content, and it is what Word renders), and
the harness sets `docx4j.jaxb.mc.preferChoice` to the prefixes of `UNDERSTOOD_NAMESPACES` so
docx4j walks the same branch; the harness's own branch selection (14.6, version 2) is then
replaced by `McSelection` at the next regeneration, and the goldens should not change. Word facts
recorded there for step 3: a text box is its own numbering story (already the rule here); an
untouched Fallback survives re-save verbatim, so a Fallback's `w:numId` may name a `w:num` that
no longer exists, and a dangling `numId` must resolve as not numbered rather than throw. Phase 2
(schema: `mc:AlternateContent` admitted in `EG_PContent` and `CT_NumPicBullet`) is a later
objects-package regeneration.

**`numRefFor`'s reason strings at docx4j `7fba7a150` (2026-09-19), adopted verbatim at the
regeneration after step 4.** `getNumber` there is a number or null, never an empty result; the
checks run in this order: no pPr / no numbering part / no resolver, the style chain, `numId` 0,
no `w:num`, `ilvl` defaulted to 0 when absent, no `w:lvl`, style-linked-elsewhere, then numbered.
The strings (X is the `w:styleId` the `w:numPr` came from; the suffix says whether the `w:numPr`
was the paragraph's own): "the paragraph's w:numId 0 turns numbering off" / "style 'X's w:numId 0
turns numbering off"; "no w:num for numId N (the paragraph's own)" / "... (from style 'X')";
"no w:lvl L in w:num N (the paragraph's own)" / "... (from style 'X')"; and the pre-existing "no
pPr", "no numbering part", "no numId, no paragraph style and no default paragraph style", "cyclic
styles at X", "style 'X' has no pPr", "no numId, and style 'X' is not numbered", "style 'X' has a
w:numPr without a w:numId val", "level L of numId N is linked to a paragraph style other than
'X'". docx4j's tests: `ParityAccessorsTest` (`numIdZeroTurnsNumberingOff`,
`danglingNumIdIsNotNumberedAndSaysWhy`, `missingLevelIsNotNumberedAndSaysWhy`). Branch commits
since the goldens' `01d661547`: `546b18b56`, `c9612931c` (CR-021 phase 1), `d34852bdd`, `7fba7a150`.

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

### 15.2 Step 3: list numbering (2026-09-19)

`src/model/listnumbering/` (`definitions.mts`, `state.mts`, `formats.mts`, `Emulator.mts`,
`index.mts`), the `NumberingDefinitionsPart` accessors, `src/parts/wml/defaultNumbering.mts`,
`WordprocessingMLPackage.refresh()` / `getNumberingEmulator()` / `numberingDefinitionsPart`,
the numbering comparison in `test/parity.test.mjs` and `test/numbering.test.mjs` are in. Ported
against docx4j `VERSION_17_1_1` at `01d661547` (the commit the goldens carry): `Emulator.java`
(645), `ListNumberingDefinition.java` (558), `ListLevel.java` (591),
`AbstractListNumberingDefinition.java` (223), `NumberingState.java`, `NumberingStates.java`,
`NumberFormatter.java` and every `NumberFormat*`, and the `initialiseMaps` / `getInd` /
`getIndFromLvl` / `resolveLinkedAbstractNum` / `getLinkedStyleId` / `getEmulator` /
`unmarshalDefaultNumbering` half of `NumberingDefinitionsPart.java`. **Parity is zero
differences** on all 45 goldens: 183 numbered paragraphs across 21 documents and every story
kind, each compared on `numString`, `isBullet`, `numFont`, `ilvl`, `numId`, all five `NumRef`
members, the whole `stateAfter` (every counter's value, `encounteredAlready` and `resetPending`,
and the spent start overrides), and `ind`, `indResolved`, `lvl` and `labelRPr` as object trees.

**What is in.**

- `definitions.mts`: `LevelDefinition` (one `w:lvl`, the abstract level with the instance's
  `w:lvlOverride/w:lvl` over it), `AbstractListDefinition`, `ListDefinition` (the counter walk
  and the `w:lvlText` expansion, `w:isLgl` included) and `NumberingDefinitions` - docx4j's
  `initialiseMaps`, its two passes, `w:numStyleLink` resolved through the numbering style, and
  `getInd` / `indOf` / `getLinkedStyleId`. `NumberingDefinitions implements NumberingIndents`,
  which is the interface step 2 left as the seam.
- `state.mts`: `Counter` (`value`, `encounteredAlready`, `resetPending`), `NumberingState`
  (counters keyed by the **referencing** `w:abstractNumId` and level, `startOverridesApplied`
  keyed `numId/ilvl`, `counter`, `reset`, `isEmpty`, `copy`, and the two read-only views the
  parity harness records), and `NumberingStates` (`main()`, `forPart(part)` with every header
  and footer in one story, `newStory()` for a text box).
- `formats.mts`: the registry keyed by `w:numFmt` and every formatter docx4j ships - decimal,
  decimalHalfWidth, decimalZero, lowerLetter and upperLetter with Word's repeated letter past z,
  lowerRoman, upperRoman, bullet, none, ordinal, cardinalText, ordinalText, hex, chicago,
  numberInDash, decimalFullWidth (and `2`), thaiNumbers, hindiNumbers, russianLower/Upper,
  arabicAlpha, thaiLetters, hebrew1, the two Chinese counting sets, decimalEnclosedCircle (and
  its Chinese alias) - with `register`, `formatterFor` and the fail-soft `formatValue`.
- `Emulator.mts`: `NumRef`, `NumberingResult`, the one `resolve`, `styleLinkedElsewhere`,
  `numRefFor`, `getNumber`, `getNumberOf`, `peek` and `getInd` as instance methods, and docx4j's
  static forms over a package (`Emulator.getNumber`, `.peek`, `.numRefFor`, `.getInd`, `.of`).
- `NumberingDefinitionsPart`: `definitions`, `getDefinitions()`, `abstractListDefinitions`,
  `instanceListDefinitions`, `getInd(numPr)`, `getIndOf(numId, ilvl)`, `getLinkedStyleId`,
  `getEmulator([reset])`, `numberingState`, `unmarshalDefaultNumbering()` and
  `refreshDefinitions()`.

**Departures from the Java, and why.**

1. **`resolveLinkedAbstractNum` does not write into the tree.** docx4j copies the linked
   definition's `w:lvl` list into the referencing `w:abstractNum` object as well as into its
   `AbstractListNumberingDefinition`. Here that would cost a document with a `w:numStyleLink`
   its byte-for-byte round trip of `word/numbering.xml` the moment anything asked for a label.
   The definitions carry the levels; nothing else needs the tree to carry them.
2. **Only the door is asynchronous.** docx4j's statics are synchronous because JAXB is. Here
   the definitions come from `readContents()`, so the part's `getDefinitions()` and the
   package's `getNumberingEmulator()` are `async`; everything after that -
   `emulator.getNumber(pPr, state)`, `Emulator.getNumber(pkg, ...)`, `peek`, `numRefFor` - is
   synchronous, as the `PropertyResolver` is after `getPropertyResolver()`.
3. **The two `getNumber` overloads are told apart by shape.** docx4j has
   `getNumber(pkg, pPr, state)` and
   `getNumber(pkg, pStyleVal, numId, levelId, directNumPr, state)`; the TypeScript overload
   dispatches on "a `string` second argument, or more than three arguments", so
   `Emulator.getNumber(pkg, undefined, '80', '0')` is the style form and
   `Emulator.getNumber(pkg, pPr, state)` the `w:pPr` one. The instance methods are
   `getNumber(pPr, state?)` and `getNumberOf(pStyleVal, numId, ilvl, direct?, state?)`, the
   split docx4j-python made for the same reason.
4. **`getNumber` returns `undefined` where docx4j returns an *empty* `ResultTriple`.** For a
   `w:numId` with no definition, or a level that does not exist, docx4j returned a triple whose
   every member was null; here the result carries `notNumbered: true` and a `reason` instead, so
   a caller is not handed an object whose label is silently absent. No golden reaches that
   branch (every `numbering` the goldens record has a `numString`), so parity does not depend on
   the choice; CR-002 phase H does. This is also where the CR-021 phase 1 note of section 14.6
   lands: a **dangling `w:numId`** - the one an untouched `mc:Fallback` can keep pointing at
   after the `w:num` it named has gone - resolves as not numbered and throws nothing.
   *Superseded in step 5:* docx4j closed the gap itself at `d34852bdd` and `7fba7a150`, so both
   cases are now answered by `resolve` with docx4j's own reasons ("no w:num for numId N ...",
   "no w:lvl L in w:num N ..."), `getNumber` returns null there as it does here, and the empty
   result is unreachable through `resolve` in either library. See section 15.4.
5. **The definitions are built by the `PropertyResolver` and shared with the part.** Step 2's
   paragraph merge already needs `NumberingDefinitionsPart.getInd` per layer, so the resolver
   builds a `NumberingDefinitions` in `init()`; `NumberingDefinitionsPart.definitions` hands out
   that same object, and builds its own only when the part has been unmarshalled and edited
   since the resolver read it, or when there is no resolver at all. One set of definitions, one
   `getInd`, and `refresh()` rebuilds both.
6. **`WordprocessingMLPackage.refresh()` is new**, and is docx4j's `getPropertyResolver(true)`:
   `refreshPropertyResolver()` re-reads the parts the resolver already knows about and cannot
   find a styles or numbering part *added* since - which is exactly what docx4j's
   `StyleLinkedLevelTest` and `DefaultStyleNumberedTest` do.
7. **A malformed `w:lvlText` does not throw.** docx4j's expansion loop calls `Integer.parseInt`
   on the character after a `%` (a `NumberFormatException` for anything but a digit) and
   dereferences the level it names (a `NullPointerException` when Word's referential-integrity
   bugs leave it missing). Here a non-digit is written out literally and a missing level
   contributes nothing; a trailing `%` is dropped, which is what docx4j's loop does.
   `LevelExists`'s own comment says such documents exist, so tolerating them is the point.
8. **`NumberingStates.forPart` keys on the part's content type**, not on its class, so that
   `state.mts` imports no part and `parts/wml` can import it. The rule is docx4j's: every header
   and footer in one story, the footnotes, endnotes and comments parts each their own,
   everything else the main story.

**Quirks reproduced deliberately.**

- **`w:isLgl`, `w:pStyle`, `w:suff`, `w:lvlJc` and `w:lvlPicBulletId` are read off *one*
  `w:lvl`** - the `w:lvlOverride/w:lvl` where the instance has one, the abstract level
  otherwise - while `w:start`, `w:lvlRestart`, `w:lvlText`, `w:numFmt` and the level font are
  merged attribute by attribute. That is not a uniform rule: docx4j's `setOverrides` merges the
  five, and `getCurrentNumberString` reads `w:isLgl` off the controlling element, so an override
  level stating no `w:isLgl` turns legal numbering **off** for a level whose abstract definition
  has it. Reproduced (`LevelDefinition.controllingLvl`). docx4j-python merged `isLgl` like the
  other five when this was written and aligned to docx4j's reading the same day (its CR-003
  section 18.10), as it did the trailing-`%` rule of item 7 below; no fixture has either case,
  and a Word-checked probe in docx4j's harness (an abstract level with `w:isLgl` and an
  override level without it) would settle which reading is Word's, both ports following.
- **A `w:num`'s first use at a level is recorded in `startOverridesApplied` whether or not it
  carries a `w:startOverride`.** docx4j sets the flag inside the branch that takes the start
  value, which fires on a fresh counter too; the goldens show `["70/0"]` for a `w:num` with no
  override at all. Reproduced, and the parity test compares the set.
- **Every level of an abstract list gets a counter on the list's first use**, not only the level
  used: incrementing level 0 resets each deeper level that restarts after it, and a reset
  creates the counter. So `numbering-stories`' first label leaves three counters standing, two
  of them `resetPending` at their start value. Reproduced; this is `w:lvlRestart`'s doing
  (CR-014 phase 2), not an initialisation pass.
- **A `w:startOverride` beside an override `w:lvl` that states its own `w:start`.** docx4j
  applies the `w:startOverride` first and `setOverrides` after, so the `w:lvl`'s `w:start` wins
  the value while `hasStartOverride` stays set and the *first use of that `w:num`* still resets
  the shared counter. Reproduced in `ListDefinition`'s constructor, with the ordering commented.
- **`getInd` prefers the level's own `w:ind` to its linked `w:pStyle`'s** (docx4j 17.1.0,
  measured; CR-014 claim 7 and probe P5), and follows `w:basedOn` from that style when the level
  states none (17.1.1). Both halves are `ListNumberIndTest`'s seven cases, ported.
- **`w:numFmt="bullet"` formats as `"*"`.** Reachable only through a `w:lvlText` that asks for a
  bullet level's counter; docx4j's own TODO says the bullet handling overlaps with this.

**Goldens questioned: none.** Every recorded value is reproduced; nothing is left failing.

**What step 2's `numberingInd.mts` became.** It is deleted. `NumberingLevels` was the indent
walk alone - the instance definitions, `w:lvlOverride`, `w:numStyleLink`'s second pass and
`getIndFromLvl` - done narrowly so that `effectivePPr` could be measured. `NumberingDefinitions`
is the same walk inside the real model and implements the same `NumberingIndents` interface, so
`PropertyResolver` changed in two lines (the import and the field) and `effectivePPr` parity is
unchanged. `styleUtil.mts` keeps `NumberingIndents` as the seam, which keeps it free of any
numbering import.

**Tests.** `test/numbering.test.mjs` (26): the formatter table and the fail-soft rules
(`LabelFormatterTest`); `NumberingRestartTest`'s P8 walk and its `restartsAfter` table;
`NumberingStoriesTest`'s P7 stories, `peek` and `reset`; `StartOverrideTest` and `IsLglTest`
over the `[expect]...[/expect]` documents; `ListNumberIndTest`'s seven flat OPC cases;
`StyleLinkedLevelTest`'s seven; `DefaultStyleNumberedTest`'s three; `numRefFor` for no numbering
part, no `w:pPr` and `w:numId` 0 both ways; `unmarshalDefaultNumbering`; and a document whose
numbering is only *read* saving `word/numbering.xml` byte for byte. `test/fixtures/ind/` is the
one fixture directory added (docx4j's `ListNumberIndTest` inputs, copied unchanged).
`NumberingConcurrencyTest` is not ported: JavaScript has one thread per realm, and what that
test proves - that two traversals with their own `NumberingState` do not interleave - is what
the parity test's per-story states exercise on every fixture.

### 15.3 Step 4: fonts (2026-09-19)

`src/model/fonts/` (`RunFontSelector.mts`, `ThemeFonts.mts`, `Mapper.mts`,
`IdentityPlusMapper.mts`, `PhysicalFont.mts`, `registry.mts`, `FontFallback.mts`,
`scripts.mts`, `LanguageTagToScriptMapping.mts`, `CJKToEnglish.mts`, `defaultTheme.mts`,
`fontsInUse.mts`, `lookup.mts`, `index.mts` and three generated modules),
`MainDocumentPart.getRunFontSelector()` / `getFontMapper()` / `fontsInUse()` /
`getStylesInUse()`, `WordprocessingMLPackage.fonts` and `createPackage`'s theme part,
`Font.name`, `scripts/generate-substitutions.mjs`, and the comparisons in
`test/parity.test.mjs` and `test/fonts.test.mjs` are in. Ported against docx4j
`VERSION_17_1_1` as CR-016 settled it (the goldens are from `01d661547`; the generated data
names the commit the generator last ran against). **Parity is zero differences** on all 45
goldens: 1,112 font spans, `fontsInUse` and `stylesInUse` as sorted lists, the default font
and theme part, and 42 distinct `IdentityPlusMapper` decisions.

**What is in.**

- `RunFontSelector.mts`: `documentFontsOf(rFonts)`, `defaultFontOf(rFonts)` / `defaultFont`,
  `themeFont(type)`, `resolvedSlots`, `complexScriptFont` (`w:cs` and `w:rtl` by value through
  `isOn`), `preambleRule`, `asciiFontName`, `fontFor` (the [MS-OI29500] 17.3.2.26 table,
  copied branch for branch), `spanScript`, `isEmoji`, `symbolFontName`, `runLocale`,
  `documentFontFor(pPr, rPr, cp, rPrIsEffective?)` and `spans(pPr, rPr, text, opts)` yielding
  `FontSpan[]` (`text`, `documentFont`, `bold`, `italic`, `cs`, `rtl`, `script`). One
  resolution, as CR-016 phase 1 decided: the selector resolves the run's effective `w:rPr`
  through the `PropertyResolver` unless the caller passes `{ rPrIsEffective: true }`.
- `ThemeFonts.mts`: `themeFontOf(theme, type, themeFontLang)`, docx4j's
  `ThemePart.getFont(STTheme, CTLanguage)` over an `a:fontScheme` rather than a part, since the
  selector, font discovery and the content API all resolve references and none of them should
  have to hold a `ThemePart`.
- `Mapper.mts`: the one precedence of CR-016 phase 3 as `populateFontMappings` plus the shared
  passes (`installedOrEmbedded`, `addMetricallyCompatibleSubstitutes`, `addAltNameSubstitutes`
  with the chain and its cycle guard, `addClassBasedSubstitutes`, `addMapperSubstitutes` as the
  empty hook, `addWordDefaultSubstitutes`, `addNoBoldFaceAliases`), `populate()` running them
  in `setFontMapper`'s order, `isKnownFamily`, `hasBoldFace`, `wordDefaultFor`,
  `registerLineMetricsAlias` / `lineMetricsFamily`, and `FontDecision` (`source`, `via`,
  `widthError`, `physicalFont`), which is what the goldens record.
- `IdentityPlusMapper.mts`: `resolveDocumentFont` alone - the name variants in the order
  regular, bold, italic, bold italic.
- `FontFallback.mts`: `classOf`, `substitutionClass`, `classFromName`, `isCondensed`,
  `leftToTheDocumentDefault`, `selectByClass`, and the coverage-group predicates `spanScript`
  needs (`isSymbol`, `isEmoji`, `isEastAsianForm`, `coverageGroupOf`).
- `fontsInUse.mts`: CR-016 phase 4's names walk over the body, headers, footers, notes and
  comments - the four slots of every `w:rFonts` on runs, paragraph marks and `w:sdtPr`, `w:sym`,
  the styles in use with their `w:basedOn` chains and `w:tblStylePr` run properties, the
  numbering levels, the document defaults and the default font - and the styles-in-use half of
  the same traversal.
- `defaultTheme.mts` and the three theme resources; see below.

**What of the Java is deliberately not ported, and why.** docx4j's `RunFontSelector` is 2,572
lines, of which this takes about 700: the decision per character. Left out, with the reason
each is a later CR (section 14.2, decision 5 of 14.5):

1. **Everything that makes XSL-FO**: `createElement` / `setAttribute` / `symbolSetAttribute`,
   the `RunFontCharacterVisitor` and the three output modes, `finish()`'s post-processes -
   `kernSpaces`, `characterScaling`, `noLigatures`, `smallCaps`, `applyLineHeight`,
   `markWidthFactor` - and `getCssProperty`. The port's answer is a `FontSpan[]`; how a span is
   represented is the consumer's.
2. **The glyph-coverage pass** (`glyphFallback`, `FontFallback.selectCovering` / `covers` /
   `needsCoverage`, `GlyphCheck`, `GlyphAdvances`, `TextMeasurer`) and with it
   `BestMatchingMapper`'s panose matching: all of it reads a font file's cmap and metrics, which
   needs `fontkit`. `addMapperSubstitutes` is the empty hook docx4j's base class has, so the
   later CR adds a mapper rather than changing this one.
3. **`WordLineMetrics` and `WidthFactors`**: line boxes and advance corrections are measurement,
   not selection. `registerLineMetricsAlias` / `lineMetricsFamily` are kept, because the
   `w:altName` and Word-default passes are written over them and a consumer that does measure
   needs to know whose metrics Word used, but nothing here reads a metric.
4. **`arabicNumbering` and `capsAndSoftHyphens`**: both rewrite the run's *text* before the
   dispatch (Arabic-Indic digit shaping; `w:caps` / `w:smallCaps`). They belong with an output
   pathway - the caller passes the text it means to draw - and neither changes which font a
   character gets.
5. **`symbolRun`'s Unicode replacement** (`SymbolMapper`, `translateUnicode2SingleByte`,
   `symbolSegments`): the selector answers "Symbol" or "Wingdings" for such a run, which is the
   document font and what the golden records; turning a private-use code point into the
   replacement character a substitute face can draw is output, and it needs the symbol jar's
   glyph tables.
6. **`registerUsedFont`, `ownFont`, `warnedOnce`, the FOP configuration**: no FOP here.

**Departures from the Java, and why.**

1. **A `FontRegistry` is a parameter, not a static map.** docx4j's `PhysicalFonts` discovers the
   machine's fonts once, in the `Mapper` static initialiser, by walking font directories and its
   own jars. None of that is portable - a browser cannot enumerate fonts, an add-in has no file
   system, Node needs `fontkit` to read a name table - so a `Mapper` is constructed over a
   `FontRegistry` the caller supplies (`get(name)` and `all()`, case-insensitive and stripping
   the twin suffixes, as `PhysicalFonts.get` does). `DEFAULT_FONT_REGISTRY` is the 43 faces
   docx4j's four font jars carry, read off `PhysicalFonts.getPhysicalFonts()` with system
   discovery off - the environment CR-016 phase 0c measured and the goldens record - so a
   mapping computed here reproduces a golden on any machine. `PhysicalFonts.discover` over
   installed fonts is a later CR, as section 6.3 says.
2. **`PhysicalFont` is a name, a family and the no-bold-face flag**, with no `EmbedFontInfo`,
   no panose, no file URI and no `Typeface`. `getFamilyName`'s triplet walk becomes the family
   the registry was told, falling back to the name with its suffixes stripped.
3. **`spans()` folds by font alone**, which is what `documentFontFor` per code point answers and
   what the harness records. docx4j's rendering walk additionally cuts a span where the
   *script* changes between two non-shared characters, so that the coverage pass can substitute
   a Greek stretch as a sibling span rather than nesting it inside the Latin one (measured:
   nesting cost a corpus document a page). That cut serves the coverage pass, which is not
   ported; `spanScript` is exported and each span reports its script, so a renderer can apply it.
4. **`U+2190-U+2BFF` answers the hAnsi font.** docx4j asks whether the hAnsi font has the glyph
   and names Segoe UI Symbol where it has that face. Both are glyph checks; without them the
   branch answers hAnsi, which is also docx4j's answer in the goldens' font environment (Segoe
   UI Symbol is not among the jars' faces), so parity holds and the divergence shows only on a
   machine that has that face.
5. **The emoji font is an option, not a property.** `docx4j.fonts.RunFontSelector.EmojiFont` is
   `RunFontSelectorSource.emojiFont`; unset by default, as docx4j's property is.
6. **`BestMatchingMapper`, `FontReport`, `FontsAnalysis.usage`, `MetricsOnlyFonts` and
   `FontEnvironment` are not ported** (no callers here, and each needs font files).
7. **The `Mapper`'s per-script choices are not recorded.** `FontDecision` here carries `source`,
   `via`, `widthError` and `physicalFont` - what a golden records - and not
   `recordScriptChoice` / `recordSymbolFace` / `getBoldFace` / `getLineBox` / `getWidthFactor`,
   all of which are the conversion's report of what the coverage pass and the metrics did.
8. **Asynchronous accessors.** `getRunFontSelector()` and `getFontMapper()` are `async` because
   they read the theme, settings, styles, numbering and font table parts;
   `runFontSelector` / `runFontSelectorOrUndefined` are the synchronous accessors afterwards,
   the pattern `getPropertyResolver()` set in step 2. Every one of those parts is read with
   `readContents()`, so resolving a document's fonts costs it no byte of its round trip
   (`test/roundtrip.test.mjs` proves it: after `getFontMapper()` only the main document part is
   unmarshalled).
9. **`Character.UnicodeScript.of(cp).name()` is `scripts.mts`**, a memoised sweep of
   `\p{Script=...}` regular expressions over the Unicode script long aliases. A script this
   engine's Unicode version does not know is dropped when its pattern fails to compile, so a
   newer script answers `UNKNOWN` on an older Node as it would on an older JDK.

**Quirks reproduced deliberately.**

- **`w:cs w:val="0"` turns an inherited complex-script flag off, and `w:rtl` alone on Latin text
  still takes the cs font.** Both are `isOn`, the `ST_OnOff` value; docx4j tested the elements'
  presence until 17.1.1.
- **A theme reference beats the explicit attribute beside it**, at every level: `resolvedSlots`
  takes `themeFont(ref) ?? explicit`, and where the package has no theme part the reference
  resolves to the Office theme's Latin face - so the explicit name is never used (Word's own
  answer, probe `fonts-missing-slots` (b)).
- **A theme reference the theme *part* cannot answer gives Calibri in the document default
  only** (`defaultFontOf`'s third branch), where the same reference on a run resolves to nothing
  and the explicit attribute stands. Asymmetric, and docx4j's.
- **An `a:font` entry whose `typeface` is empty answers the empty string**, where a script the
  list does not carry falls back to the collection's own `a:latin`. `ThemePart.getFont`'s map
  lookup distinguishes the two and this does too.
- **A run with no `w:rFonts` at all gets the default font in the ascii and hAnsi slots only**,
  never the eastAsia slot: a Times New Roman there would fire the preamble rule and set the
  whole run in one span.
- **`w:cs=""` (LibreOffice writes it in docDefaults) is read as no complex-script font**, but
  only in `resolvedSlots`; `complexScriptFont` returns it as it is, as docx4j's does.
- **`isKnownFamily` asks `substitutionClass`, not `classOf`**: a name that merely ends in "Sans"
  is *not* a known family, so the Word-default pass acts on it - the two disagreed until
  docx4j 17.1.1 and a corporate face fell between them, drawn in the document default's serif
  throughout.
- **`addNoBoldFaceAliases` skips a Word-defaulted font**: a font Word itself could not find is
  substituted whole, its real bold included ("EnBW DIN Pro Light" is Calibri Bold in Word).

**Goldens questioned: none.** One comparison normalises the golden before matching it, and it
is not a font difference: `@xmldom/xmldom` applies XML 1.1's line-ending normalisation to an
XML 1.0 document, so a `U+0085` (NEL) inside a `w:t` arrives here as `U+000A` where Xerces keeps
it (XML 1.0 section 2.11 normalises only `#xD` and `#xD#xA`). One fixture has one such character
(`tracked-changes.docx`, "and here it continues"), and the span it falls in is otherwise
identical - same font, same flags. `asParsed()` in `test/parity.test.mjs` names the cause. This
is a defect of the XML layer, not of this step: a *re-marshalled* part would lose the character,
so it belongs in a runtime or objects CR (see the last paragraph).

**The `defaultTheme` setting, and `createPackage`'s theme part (section 14.6, the Phase A
departure closed).** docx4j's `docx4j.fonts.defaultTheme` is one JVM-wide property; this package
has no properties file, and a browser or an add-in may hold several packages at once, so the
setting is per package: `pkg.fonts.defaultTheme` (a `FontSettings` object on the package),
`'2023'` | `'2013'` | `'2007'`, defaulting to the process-wide `defaultThemeSetting()`, which
is `'2023'`. `createPackage({ defaultTheme })` sets it and then adds a `ThemePart` from that
theme's bundled resource, as docx4j's `addDefaultThemePart` does, so a package created by either
library resolves its `minorHAnsi` defaults to the same face. The three resources are embedded as
`themes.generated.mts` (docx4j's `theme-2023.xml`, `theme-2013.xml`, `theme-2007.xml`), beside
`defaultStyles.mts`. A created package therefore has nine zip entries rather than eight
(`test/create.test.mjs` and the acceptance note in `test/README.md` say so), and a *loaded*
package with no theme part still resolves its references from the setting, which is where the
golden `fonts-missing-slots` gets its Aptos.

**The generated data** (`scripts/generate-substitutions.mjs`, `npm run generate:fonts`, output
committed): `substitutions.generated.mts` from `font-substitutes.xml`'s `<substitutes>` rows
(30 document fonts) and `FontSubstitutions.xml` (419 OpenOffice VCL entries: the font classes
and candidate lists `FontFallback` reads); `families.generated.mts` from `MicrosoftFonts.xml`
(137 families, 41 with a bold face of their own - which is what `hasBoldFace` answers on, and
the no-bold-face list) and `word-line-metrics.properties` (512 family names, 46 flagged East
Asian - `isKnownFamily` and the alt-name pass's East Asian hop; no metric is read);
`themes.generated.mts`. Each header names the docx4j commit the generator ran against. The
table's `<scriptSubstitutes>` and `<widthFactors>` blocks are deliberately not generated: they
serve the coverage and width passes, which this CR does not port.

**`Font.name` (step 2's one remaining direct read, closed).** `Font` gains a fourth supplier,
the resolved ASCII font name, which `Paragraph.getFont()` and `Range.getFont()` fill with
`RunFontSelector.asciiFontName` of the effective `w:rPr` - `w:rFonts/@w:ascii`, or the face
`w:asciiTheme` names, else `w:hAnsi`, else the document default. So a run whose face comes from
the theme now reads "Aptos" (or "Calibri") rather than `''`, and one that names no font anywhere
reads the document default rather than `''`. `getFont({ direct: true }).name` is unchanged, and
so is the read where the package's selector has not been built yet. The lookup is structural
(`runFontSelectorOf`), as `Body.propertyResolver` is, so the content API keeps no import of a
part class.

**Tests.** `test/fonts.test.mjs` (23): the CR-016 phase 1 and 2 cases
(`RunFontSelectorCsValueTest`, `RunFontSelectorThemeLangTest` with Estonian and the default
font's theme language, `RunFontSelectorNoRFontsTest`, the theme reference with no theme part,
the range table of `RunFontSelectorDispatchTest` including the Latin-1 exceptions and the Indic
ranges, the Latin-1 range reset, span joining, the preamble rule, the symbol fonts and emoji),
`LanguageTagToScriptMappingTest`, `MapperPrecedenceTest` in five parts (installed over embedded,
the embedded form, the face order, the altName chain and its cycle, Word's default and the
guessed-class case), `NoBoldFaceTest` with the Word-defaulted exception,
`MetricallyCompatibleSubstituteTest`, `ClassBasedSubstituteTest`, `FontsInUseTest` with
`getStylesInUse`, the themeless default for all three values, and `createPackage`'s theme part
for all three. `test/parity.test.mjs` adds a second per-golden test (font spans, `fontsInUse`,
`stylesInUse`, the default font, the theme part and the mapping), and `test/roundtrip.test.mjs`
the byte-identity check. Suite: 323 tests, 322 pass, 0 fail, 1 `todo` (step 3).

**For the objects package and the runtime: one thing, not blocking.** `@xmldom/xmldom`
normalises `U+0085` and `U+2028` to `U+000A` in an XML 1.0 document, which is XML 1.1's rule
(XML 1.0 section 2.11 normalises only `#xD` and `#xD#xA`). Character data therefore differs from
Java's, and a re-marshalled part loses the character. Found through the font spans of
`tracked-changes.docx`; it belongs to `src/xml/dom.mts` or the runtime, not here. Everything
else this step needed - `walk`, `readContents`, the generated declarations, the facade - was
already there.

### 15.4 Step 5: review, regeneration and close-out (2026-09-19)

The review of section 13 step 5, the regeneration of all 45 goldens from docx4j's current head,
the reason-string alignment, and the documents. Phase B is implemented.

**Harness version 3, and the regeneration.** docx4j `VERSION_17_1_1` at `7fba7a150`, built
`-pl docx4j-core,docx4j-JAXB-ReferenceImpl,docx4j-export-fo-fonts-{symbol,croscore,crosextra,theme2023}
-am -DskipTests -Dgpg.skip install`; `docx4j-core-17.1.1-SNAPSHOT.jar` SHA-256
`1c2e2e42a2f50e701e7193a6356fa6c9f535edec45d88db91f9815e56db92cb9`. The branch had already
moved past that commit while this step ran (`1ca86c3c3`, CR-021 phase 2, the schema change), so
the build was made from a read-only `git archive` export of `7fba7a150` into a scratch directory
rather than from the checkout - the rule that nothing in `../docx4j` is touched, kept literally.

**A harness build must not share `~/.m2` with another session**, which this step learned the
hard way: `17.1.1-SNAPSHOT` is one coordinate, the install of it here and the docx4j session's
own install of `1ca86c3c3` overwrote each other within the same minute, and a determinism check
that straddled the two compared goldens made from two different jars. (They agreed, which says
something reassuring about CR-021 phase 2, but it was not the check being run.) The rule, now in
`test/java/README.md` and in the workflow's comment: when the docx4j checkout or the machine is
shared, give the docx4j build *and* the harness run a repository of their own with
`-Dmaven.repo.local=<dir>`, the same directory for both and no `-o` on its first use. The
scheduled workflow has one by construction. The goldens committed here are from the build made
before that clash, whose jar the headers name (`1c2e2e42a2f50e70`), verified by two runs against
it; nothing was rebuilt afterwards.

The harness change is what the CR-021 phase 1 paragraph of section 14.6 said it would be. Its
own `mc:AlternateContent` branch selection is gone; `childrenOf` is `TraversalUtil.getChildrenImpl`
in its default `McMode.READ`, which gives up the one branch `McSelection` selects, and `main`
sets `docx4j.jaxb.mc.preferChoice` before any walk to the prefixes of this package's
`UNDERSTOOD_NAMESPACES`, mapped through the objects package's `NAMESPACE_PREFIXES`:

```
a a13cmd a14 a15 a16 a1611 a16svg a18hc adec am3d an18 anam3d b c c14 c15 c16 c16ac c173 cdr
cdr14 comp cp cppr cs cx dc dcterms dgm dgm14 dgm1612 ds dsp iact ink16 lc m mc msink o p
p13cmd p14 p15 p1510 p159 p16 p166 p1710 p173 p184 pic pic14 pkg prop properties psez pslz
psuz pvml r rel sl thm15 v vt w w10 w14 w15 w16cid w16se we wetp wne wp wp14 wp15 wpc wpg wps
xdr xdr14 xvml
```

Eighty-four prefixes, recorded in every golden's header as `mcPreferChoice` and documented in
`test/java/README.md`. Nine understood namespaces have no entry in that table (MathML, InkML,
the two Excel mains, the three encryption ones) and SpreadsheetML's entry is the default
namespace, so none is nameable; a `Requires` naming one would diverge and no fixture has one.
`v` is in the list, where docx4j's own javadoc advises against it: the rule here is parity with
this package's preprocessor, not docx4j's rendering advice. The property is set through
`Docx4jProperties`, which wins over anything the environment supplies, so the weekly workflow
makes the same choice without setting anything of its own - one copy of the list, which is why
`.github/workflows/parity.yml` gained a comment rather than a duplicate of it. `HARNESS_VERSION`
is `"3"`.

**The regeneration: no difference anywhere but the header.** All 45 goldens compare equal, field
for field, to the committed version-2 set once the header is removed - no paragraph set moved
(the text boxes still appear once, from the `mc:Choice`, now by docx4j's rule rather than the
harness's), no effective property, no label, no counter, no table flag, no font span. The only
`numRef.reason` strings a regeneration could have moved are the two new ones, and they arise
only where a paragraph hits a dangling `w:numId` or a missing level, which no golden reaches -
`numbering` is null for an unnumbered paragraph, so the reason is not recorded at all there. The
headers changed in four fields and gained one: `harnessVersion` 2 to 3, `docx4jCommit`
`01d661547` to `7fba7a150`, `docx4jCoreJarSha256` `d0bd889c3a135e27` to `1c2e2e42a2f50e70`,
`date`, and the new `mcPreferChoice`. Determinism re-checked at version 3: two runs into two
directories, against the same jar, differ only in `date`.

**Reason strings.** `numRefFor`'s reasons are docx4j's at `7fba7a150` verbatim, in its check
order (no pPr / no numbering part / no resolver, the style chain, `w:numId` 0, no `w:num`, the
`ilvl` default to 0, no `w:lvl`, style-linked-elsewhere, then numbered). The port was missing
the two checks docx4j added at `d34852bdd` and `7fba7a150` - a `w:numId` naming no `w:num`, and
a `w:num` whose definition has no `w:lvl` for the level - which it had been answering later, as
an empty result from the counting half with reasons of its own (`Couldn't find list N`).
`Emulator.resolve` now makes both, in docx4j's order and with its strings, including the
parenthesised source suffix that `sourceSuffix(direct, styleId)` writes: "(the paragraph's own)"
or "(from style 'X')". `number()`'s empty-result branch is consequently unreachable through
`resolve` and says so. The parity comparison of `numRef.reason` in `test/parity.test.mjs` is
exact equality and always was - no prefix matching was ever there, and there is none now.
`test/numbering.test.mjs` gained docx4j's `ParityAccessorsTest` cases with its exact strings:
`numIdZeroTurnsNumberingOff`, `danglingNumIdIsNotNumberedAndSaysWhy` and
`missingLevelIsNotNumberedAndSaysWhy` over a package whose Normal is numbered, plus the same
two with the `w:numPr` coming from a style so that the "(from style 'X')" half is covered, and
an assertion that none of the five touched a counter.

**Review 1: the toggle-property overlay.** `styleUtil.mts` (`applyStyleLevel`, `applyToggles`,
`toggle`) and `catalogue.mts` (`TOGGLE_NAMES`, `TOGGLES`) against `StyleUtil.applyStyleLevel` /
`applyToggles` / `toggle` and `PropertyCatalogue`'s static block, with every call site in
`PropertyResolver`.

| # | Rule | Java | Port | Agree |
|---|---|---|---|---|
| 1 | The twelve toggles in §17.7.3's order (`b bCs caps emboss i iCs imprint outline shadow smallCaps strike vanish`); `w:dstrike`, `w:noProof`, `w:snapToGrid`, `w:webHidden`, `w:rtl`, `w:cs`, `w:specVanish`, `w:oMath` are Boolean but not toggles | `PropertyCatalogue` static block, `TOGGLE_NAMES` | `catalogue.mts` `TOGGLE_NAMES` | yes |
| 2 | `TOGGLES` is those names looked up in the run table; construction fails if one is missing | same (and checks the member's type is `BooleanDefaultTrue`) | `catalogue.mts` `TOGGLES` (throws on a missing name; no runtime type check, which TypeScript has no equivalent of - the catalogue-completeness test covers it) | yes |
| 3 | `applyStyleLevel`: a null source or destination is a no-op | `StyleUtil` 2044 | `styleUtil.mts` 139 | yes |
| 4 | it applies the non-toggles by ordinary override, excepting the toggle names | `PropertyCatalogue.apply(RUN, …, TOGGLE_NAMES)` | `applyCatalogue(RUN, …, TOGGLE_NAMES)` | yes |
| 5 | it does **not** call `skipRun`, so a source with nothing to say still goes through the toggles | `StyleUtil` 2044-2049 | `styleUtil.mts` 139-144 | yes |
| 6 | `applyToggles` writes only the twelve; a null destination is a no-op | `StyleUtil` 2063 | `styleUtil.mts` 150 | yes |
| 7 | writing null clears the member | `Property.set(destination, null)` | `prop().set` deletes the key | yes |
| 8 | `toggle`: a silent upper level is no boundary - the lower value stands and the document-defaults-true rule does not arise | `if (upper == null) return lower;` | `if (upper === undefined) return lower;` | yes |
| 9 | `toggle`: the document defaults are a base value, never an XOR term - where they say true and the upper level states the property, a *copy* of the defaults' value wins | `apply(documentDefault, lower)`, which is `XmlUtils.deepCopy(documentDefault)` | `copyLeaf(documentDefault)` | yes |
| 10 | `toggle`: an explicit false XORs like any other value - false XOR lower = lower; with nothing beneath, the level's own false stands rather than the property going absent | `if (!upper.isVal()) return lower != null ? lower : upper;` | `if (!isTrue(upper)) return lower !== undefined ? lower : upper;` | yes |
| 11 | `toggle`: an explicit true inverts, into a new element that always states `w:val` | `out.setVal(!(lower != null && lower.isVal()))` | `out.val = !isTrue(lower)` | yes |
| 12 | `BooleanDefaultTrue`: an absent `w:val` is true | `isVal()` | `isTrue(v)`: `v.val !== false` | yes |
| 13 | The only level boundary the resolver applies is paragraph style to character style | `applyCharacterStyleAndDirect`, the sole `applyStyleLevel` call | `PropertyResolver.mts` `applyCharacterStyleAndDirect`, sole call | yes |
| 14 | Document defaults are the **base**, applied with `applyRPr`, not a level | `getEffectiveRPr(RPr, PPr)` | same | yes |
| 15 | A style's `w:basedOn` chain is one level, merged root-first with `applyRPr`, cached per style id without the defaults | `chainRPr` | `getChainRPr` | yes |
| 16 | Direct formatting is not a level: applied as it stands with `applyRPr`, gated on `hasDirectRPrFormatting` | `applyCharacterStyleAndDirect` | same | yes |
| 17 | A character style that does not exist is logged once and contributes no level | `getLiveStyle(runStyleId) == null` | same | yes |
| 18 | `getEffectiveParagraphMarkRPr` has no level boundary: `w:pPr/w:rPr` goes through `applyRPr` | `PropertyResolver` 456 | same | yes |
| 19 | `getEffectiveRPr(styleId)` has no level boundary either | `PropertyResolver` 521 | `effectiveRPrOfStyle` | yes |
| 20 | Table styles are not a level: there is no `applyStyleLevel` for them (table conditional formatting is docx4j's `table-conditions` CR) | no such call | no such call | yes |

**No disagreement.** One cosmetic difference, not behaviour: docx4j logs a warning when a
`w:pStyle` carries no `w:val` before falling back to the default paragraph style, where the port
falls back silently.

**Review 2: the numbering counters.** `state.mts` (`Counter`, `NumberingState`,
`NumberingStates`) and the counter half of `definitions.mts` against `NumberingState.java`,
`ListLevel` (`Counter`, `counter`, `incrementCounter`, `resetCounter`, `restartsAfter`,
`setStartValue`), `ListNumberingDefinition.incrementCounter` and its `w:lvlOverride` pass, and
`NumberingStates.java`.

| # | Rule | Java | Port | Agree |
|---|---|---|---|---|
| 1 | A counter is `value`, `encounteredAlready`, `resetPending`; `copy()` carries all three | `ListLevel.Counter` | `state.mts` `Counter` | yes |
| 2 | Counters are keyed by the **referencing** `w:abstractNumId` and the level, and created on first use at the level's start value (0 where none) | `NumberingState.counter` | `NumberingState.counter` | yes |
| 3 | Spent start overrides are keyed `<numId>/<ilvl>` | `startOverridesApplied` | same | yes |
| 4 | `reset()` clears both, `isEmpty` is both empty, `copy()` is independent | `NumberingState` | same | yes |
| 5 | `counters()` / `startOverridesApplied()` are read-only views **of the live maps** | unmodifiable wrappers | `ReadonlyMap` / `ReadonlySet` over the live maps | yes |
| 6 | `ListLevel.incrementCounter`: the counter takes the level's start value when the level has not been encountered **or** when this `w:num`'s `w:startOverride` has not been spent in this story | `ListLevel` 386 | `definitions.mts` 224 | yes |
| 7 | That branch sets `encounteredAlready`, clears `resetPending` and marks the override applied - for any `w:num` with an owner, whether or not it carries an override (which is why the goldens show `["70/0"]` for a `w:num` with none) | `ListLevel` 392-396 | same | yes |
| 8 | A pending reset consumes itself and does **not** increment | `ListLevel` 398-401 | same | yes |
| 9 | Otherwise increment by one | `counter.increment()` | same | yes |
| 10 | `resetCounter` puts the counter at start + 1 and sets `resetPending`, so a deeper label reads "2.1.1" and not "2.0.1" (probe P8) | `ListLevel` 427 | `definitions.mts` 247 | yes |
| 11 | `restartsAfter`: no `w:lvlRestart` means any shallower level restarts; `0` means none does; `n` means ilvl 0..n-1 do | `ListLevel` 458 | `definitions.mts` 258 | yes |
| 12 | The list walk: where this level is unencountered, walk shallower levels down from `levelInt - 1`, stopping at the first that is missing or already encountered, incrementing each | `ListNumberingDefinition` 314 | `definitions.mts` 426 | yes |
| 13 | Then increment this level | same | same | yes |
| 14 | Then reset every deeper level that `restartsAfter(levelInt)` - which creates its counter, so a list's first use leaves a counter standing for every level, most of them `resetPending` at their start | `ListNumberingDefinition` 337 | `definitions.mts` 443 | yes |
| 15 | `w:startOverride` is applied at construction as `setStartValue(val - 1)`, which also raises `hasStartOverride`; an override `w:lvl`'s own `w:start` is applied after it, so it wins the value while the flag stands and the `w:num`'s first use still resets the shared counter | `ListNumberingDefinition` 239, then `setOverrides` | `ListDefinition`'s constructor, with the ordering commented | yes |
| 16 | `NumberingStates`: one main state; one shared by every header and footer; one per footnotes, endnotes and comments part; `newStory()` for a text box | `NumberingStates` (by part class) | `state.mts` (by content type: the recorded departure, same rule) | yes |
| 17 | A level the definition does not have | `getLevel` would dereference null; unreachable, since `resolve` checks `levelExists` first | returns early, and `resolve` checks `levelExists` too (departure 7, fail-soft) | yes |

**No disagreement.** Two differences of kind rather than of behaviour, both already recorded:
counters are `number` here where Java uses `BigInteger` (no fixture, and no format this CR
ports, reaches 2^53), and `NumberingStates.forPart` keys on the part's content type rather than
its class so that `state.mts` imports no part.

**What Phase B leaves open.**

- **`BestMatchingMapper` and the glyph-coverage pass** (with `GlyphCheck`, `GlyphAdvances`,
  `TextMeasurer`, `WordLineMetrics`, `WidthFactors`, panose matching, `PhysicalFonts.discover`):
  everything that reads a font file. A later CR with `fontkit`; `addMapperSubstitutes` is the
  empty hook it plugs into, and `FontRegistry` the seam (section 15.3, departures 1 and 6).
- **Table conditional formatting** (`w:tblStylePr` per `w:cnfStyle`): docx4j's own
  `table-conditions` CR has not landed, so neither side has it (section 15.1).
- **`activateStyle(styleId)` without a `KnownStyles.xml`**: the string form can only activate a
  style the package already holds (section 15.1, departure 2). The resource equivalent belongs
  with the content API's style creation.
- **The jsonix U+0085 defect**: `@xmldom/xmldom` applies XML 1.1's end-of-line normalisation to
  an XML 1.0 document, so a raw U+0085 or U+2028 in character data becomes U+000A on unmarshal
  and a re-marshalled part loses it. One parity comparison normalises around the single
  occurrence in `tracked-changes.docx`. Sent to the jsonix session on 2026-09-19; section 9 has
  the detail. It is an XML-layer defect, not a Phase B one, and it is the only place a golden is
  touched before it is matched.
- **The CR-021 phase 2 regeneration.** docx4j committed phase 2 (`1ca86c3c3`) while this step
  ran: the schema admits `mc:AlternateContent` in `EG_PContent` and `CT_NumPicBullet`, so load
  can keep both branches instead of resolving the element to one. That needs an objects-package
  regeneration before this package can follow, and until then section 5.6's DOM preprocessor
  stays as it is. The goldens are deliberately pinned at `7fba7a150`, before it. The weekly
  workflow will raise it as a pull request if docx4j's answers move; the two probes
  `mc-textbox-branches-*` are the fixtures to add when they do.

**Everything else the phase leaves.** `test/README.md`'s parity-golden section names harness
version 3 and the commit; its acceptance checklist records that **check 3 is outstanding** -
`createPackage()` gained a theme part in step 4 (nine zip entries instead of eight) and has not
been opened in Word since - and its script for checks 1 to 3 now creates a document for each of
the three `defaultTheme` values. `README.md`'s status paragraph says Phase B is in and names
the goldens; its Development section gains `npm run generate:fonts` and the harness. No script
was needed beyond `generate:fonts`, which exists. `npm run generate` produces no diff.
`npm run generate:fonts` reproduces all three generated modules byte for byte except the
docx4j commit in each header, which follows `../docx4j`'s HEAD - now past the pinned commit, so
the regenerated headers were reverted and the modules still name `7fba7a150`.

**Tests.** 395 tests, 395 pass, 0 fail, 0 `todo` (step 4 left one, which step 3's landing
closed); `npm run typecheck` clean. `test/numbering.test.mjs` is 28.

**For the objects package and the runtime: nothing new.** The U+0085 defect above is the one
open item, already sent.

**CR-021 phases 2 and 3 (docx4j `1ca86c3c3`, `e64780db2`, `e864468a4`, 2026-09-19; for the section 5.6
decision and Phase C).** docx4j's schema now admits `mc:AlternateContent` in `w:p`, `w:numPicBullet`
and the PresentationML places, `a:p`'s run list admits `a14:m`, and `a14:m` holds its math as a lax
wildcard; all reach this package through an objects-package regeneration. When it lands, section
5.6's DOM preprocessor becomes a choice (resolve on load, as the content API and the editor want;
or keep both branches for a byte-faithful re-marshal). The rule docx4j recorded in CR-021 §8.6 item
10, from a PowerPoint check that caught a lossy kept branch: **a kept `mc:Choice` is lossless only
if the model binds its whole content**, so any kept branch is verified in the producing
application before the preprocessor's default changes. Nothing in WordprocessingML traversal or
numbering moved; the goldens stand.

**Objects regeneration for CR-021 (2026-09-19; objects commits `fc6d851` from docx4j `e864468a4`,
then `9afaba8` from docx4j `a58cf10b8`; unreleased, 0.1.5 proposed).** Shape changes for this
package when it upgrades: `P.content` and the other run-content unions (`Hyperlink`, `P.Dir`,
`P.Bdo`, `CTSimpleField`, `CTSmartTagRun`, `CTCustomXmlRun`, `CTSdtContentRun`) gain an
`mc:AlternateContent` member (additive; an exhaustive narrowing needs a new arm);
`NumPicBullet.alternateContent`; DrawingML's `CTTextParagraph` admits `a14` math;
`Workbook.alternateContent` (SpreadsheetML) becomes optional, a compile error for any reader (none
here). **The constraint on section 5.6, and its removal, in one day.** The first regeneration
measured that a kept `mc:Choice` loaded only when every child was a global element the context
types (`wps:wsp`, `w:r`: yes; `w:tbl` or an untyped namespace: the whole unmarshal threw): the
objects session traced it to `processContents="strict"` on the two wildcards in docx4j's
`xsd/mce/markup-compatibility-2006-MINIMAL.xsd`, docx4j changed them to `lax` (`a58cf10b8`), and
the second regeneration's diff was exactly the two `allowDom: false` flags disappearing. Measured
after: a Choice holding `w:r` is typed as before; one holding `w:tbl` or a vendor-namespace element
loads as a DOM element and round-trips unchanged, attributes and text intact. So keeping
`mc:AlternateContent` unresolved is now lossless. This package's default nonetheless stays
resolve-on-load, for the two reasons the objects session weighed: a DOM branch is opaque to the
typed views (`textOf` reads nothing from it, `find` and `walk` do not enter it; `walkAll` does), and
"typed inside a branch" means the mapping's global elements, which is not the set docx4j's Java
binds (`w:tbl` is bound there through an `@XmlRootElement`, not here; a compiler change, filed
nowhere while nothing needs it). Section 5.6's default becomes a decision to revisit when a
consumer needs a byte-faithful re-marshal of both branches. The objects package's `textOf` follows
docx4j's `McSelection` through `mcBranchOf`.

## 16. Phase C implementation notes (2026-09-19)

Phase C is the CR's last: the PresentationML and SpreadsheetML packages get `createPackage()` and
the shortcuts docx4j has, the unzipped-directory container arrives as a Node-only subpath,
`OpcPackage.clone()` lands, and the package grows guides and runnable examples. 428 tests (411
before), `npm run typecheck` clean, no new runtime dependency.

### 16.1 `PresentationMLPackage.createPackage()`

Ported from docx4j `PresentationMLPackage.createPackage(SlideSizesWellKnown, boolean)`,
`MainPresentationPart`, `SlideMasterPart`, `SlideLayoutPart` and `SlidePart` at
`VERSION_17_1_1` `a58cf10b8`. A created presentation is eleven zip entries:

```
[Content_Types].xml   _rels/.rels
ppt/presentation.xml + _rels          ppt/slideMasters/slideMaster1.xml + _rels
ppt/slideLayouts/slideLayout1.xml + _rels   ppt/theme/theme1.xml
ppt/slides/slide1.xml + _rels
```

Relationships exactly as docx4j builds them: the presentation relates the master, the theme and
each slide; the master relates the layout and the theme; the layout relates the master back; the
slide relates the layout. PowerPoint's own files do relate the theme from the presentation part as
well as from the master, so docx4j's "add it in 2 places" is right, not a quirk.

- **The slide is ours.** docx4j's `createPackage` has its slide creation commented out, so its
  presentation has an empty `p:sldIdLst`. Phase C adds one empty slide on the layout, because a
  deck with no slides is not what a caller asking for a presentation means, and
  `MainPresentationPart.addSlide` / `addSlideIdListEntry` are ported anyway.
- **The theme is ours too.** docx4j reads one fixed
  `org/docx4j/openpackaging/parts/PresentationML/theme.xml`; here the presentation uses the same
  embedded Office themes a new `.docx` gets (`src/model/fonts/themes.generated.mts`), selected by
  `pkg.fonts.defaultTheme` — `PresentationMLPackage` gains the `fonts` setting
  `WordprocessingMLPackage` has. One theme resource for both formats, and a created deck's fonts
  then match a created document's.
- **`presProps`, `viewProps`, `tableStyles`, `docProps`: not written.** docx4j writes none of
  them and they are optional; PowerPoint supplies its own. The *shortcuts* for them exist
  (`presentationPropertiesPart`, `viewPropertiesPart`, `tableStylesPart`, `commentAuthorsPart`,
  `notesMasterPart`) and are set on load, as docx4j's `setPartShortcut` does for the three it has.
- **The template markup is `src/parts/pml/defaults.mts`**, the way `defaultStyles.mts` is: the
  master, layout and slide are docx4j's `COMMON_SLIDE_DATA` and `COLOR_MAPPING` written out as
  whole parts, with a header naming the docx4j commit and the Java files. They are set with
  `setXml`, so a created layout or slide nobody touches is written byte for byte as it is there;
  the master is unmarshalled during creation because `addSlideLayoutIdListEntry` has to append to
  `p:sldLayoutIdLst`, exactly as docx4j marshals its own tree. (**Superseded for the layout and
  the slide by section 17.2**: both are unmarshalled during creation now, to take the placeholder
  shapes PowerPoint needs before it will offer "Click to add title". The master is still written
  exactly as it is here.)
- **Ids** follow docx4j: `p:sldId/@id` random in 256 to 2147483647, `p:sldLayoutId/@id` and
  `p:sldMasterId/@id` random above 2147483648 (ECMA-376 4.8.17, 4.8.18, 4.8.20), exported as
  `nextSlideId()` and `nextSlideLayoutOrMasterId()`. Random means a created package is not
  byte-reproducible; docx4j is the same and the tests assert ranges, not values.
- **Slide sizes** are docx4j's `SlideSizesWellKnown` as a string union, with docx4j's EMU table
  (`createSlideSize(size, landscape)`); `B4JIS` is in the enumeration and has no size in docx4j,
  which throws, and so does this. The default is docx4j's: `A4`, landscape — not PowerPoint's own
  16:9, so that a package created here and one created by docx4j are the same package.
- **`slideParts`, `slideMasterParts`, `slideLayoutParts`** read the id lists when the owning part
  is unmarshalled (docx4j's `getSlideParts()`, which requires it) and fall back to relationship
  order when it is not, so a synchronous getter never throws on a freshly loaded package; the
  asynchronous `getSlideParts()` unmarshals first. `MainPresentationPart.addSlide(part, index?,
  layout?)` renames the part when the name is taken, as docx4j does, and
  `PresentationMLPackage.addSlide()` picks the next free `/ppt/slides/slideN.xml` instead of
  docx4j's counter-appending rename (`slide1.xml` would become `slide11.xml`).

### 16.2 `SpreadsheetMLPackage.createPackage()` and the workbook's `mc:AlternateContent`

`createPackage()` is docx4j's, part for part: `/xl/workbook.xml` with one
`bookViews/workbookView` (docx4j adds it because without it Excel 2010 could crash on print) and
an empty `sheets`. `createWorksheetPart(name, index?, { partName, sheetId })` is docx4j's
`createWorksheetPart(PartName, String, long)` with the part name and the sheet id defaulted (the
first free `/xl/worksheets/sheetN.xml`, the next free id) and an `index` that inserts the tab
rather than appending. Five zip entries for a one-sheet workbook. `worksheetParts` is in `sheets`
order — the tab order — when the workbook is unmarshalled; `sharedStringsPart`, `stylesPart`,
`calcChainPart`, `themePart` and a worksheet's `drawingPart`, `commentsPart`, `tableParts`,
`workbookPart` are the shortcuts.

**The workbook `mc:AlternateContent` finding (docx4j CR-021).** Measured on
`test/fixtures/loadAndSave.xlsx`, and pinned by a test in `test/create.test.mjs`: Excel writes

```xml
<mc:AlternateContent><mc:Choice Requires="x15">
  <x15ac:absPath url="/Users/bcronk/Downloads/" .../>
</mc:Choice></mc:AlternateContent>
```

at the top of `xl/workbook.xml` — **a single Choice, no Fallback**. `x15` is
`http://schemas.microsoft.com/office/spreadsheetml/2010/11/main`, which is not in
`UNDERSTOOD_NAMESPACES`, so section 5.6's rule finds no understood Choice and no Fallback. The
preprocessor's answer in that case is to **drop the `mc:AlternateContent` entirely**, which is
what ECMA-376 Part 3 10.2.1 prescribes (the element is removed; with no selected branch, nothing
replaces it) and what Excel itself does with a Choice it does not understand. Consequences:

- a workbook nobody unmarshals round-trips **byte for byte** (asserted);
- a workbook that *is* unmarshalled comes back without the element, so a re-marshalled
  `xl/workbook.xml` has no `x15ac:absPath`. The lost content is the author's local folder, which
  Excel rewrites on its next save, so nothing of the document is lost; the general rule stands
  that unmarshalling a part costs its untaken MCE branches.
- **`Workbook.alternateContent` is declared required** in `@docx4j/generated-objects-ts` 0.1.4
  (`alternateContent: AlternateContent;`) and is absent on every workbook loaded here, so the
  declaration lies for a reader. The CR-021 regeneration (section 15.4, objects 0.1.5 proposed)
  makes it optional, which fixes it; nothing here reads the member, and `createPackage` builds
  the workbook through the generated factory (`createWorkbook`, whose `init` is a `Partial`), so
  the required member never had to be supplied. **For the coordinating session: this is one more
  reason to land objects 0.1.5.**

### 16.3 `DirectoryPartStore`, `DirectoryPartSink` and the `./node` subpath

docx4j's `io3.stores.UnzippedPartStore`, in `src/opc/DirectoryPartStore.mts`:
`DirectoryPartStore.open(dir)` scans the directory once (a `PartStore` lists its names
synchronously, which a `readdir` per call cannot do) and maps a part name to a file path, with
`size()` from the file; `new DirectoryPartSink(dir)` writes one file per part, creating
directories, `[Content_Types].xml` first, and refuses a part name that climbs out of the
directory (docx4j's "Zip Slip" check). `finish()` resolves to the directory path. Nothing already
in the directory is removed.

**Fidelity, measured**: zip → directory → zip reproduces every part byte for byte except the
relationships parts and `[Content_Types].xml`, which every save writes afresh — a plain zip → zip
round trip of the same fixture differs in exactly the same four `.rels` and no other part. So the
directory container costs nothing.

**The `./node` subpath.** It is the one module in the package that imports `node:` builtins, so it
is exported only from a new `./node` entry point (`src/node/index.mts`), never from `.`, `./opc`,
`./parts`, `./packages`, `./model` or `./office-js`. A browser or add-in bundle therefore never
sees `node:fs`, needs no polyfill and no `browser` field — which is the property the add-in guide
can promise. The rule is now in `CLAUDE.md`'s Public paths.

Two consequences worth recording:

- **No `@types/node`.** The repository's tsconfigs have `"types": []` and `lib` es2019 + dom
  because everything else is platform-neutral; rather than add a types-only dependency for one
  module, the handful of Node functions used are declared in `src/node/node-builtins.d.mts`
  (`readFile`, `readdir`, `stat`, `mkdir`, `writeFile`, `join`, `dirname`, `sep`). It is not
  emitted, and the public declarations of `./node` name no Node type, so a consumer with or
  without `@types/node` sees no conflict. If more of Node is ever needed here, the honest move is
  the dependency.
- `test/nodenext/consumer.mts` imports `./node` like the other subpaths, so the entry point is
  checked under `moduleResolution: nodenext` as well.

### 16.4 `OpcPackage.clone()`

docx4j's: **save and reload**, into a `MemoryPartSink` rather than a zip so nothing is deflated.
Chosen over a part-by-part copy because it is the copy whose cost matches what was touched — a
part nobody unmarshalled is copied as the bytes it was loaded with and never parsed, an
unmarshalled part is marshalled once and comes back as its own tree, unmarshalled again on
demand — and because a part-by-part copy would have to be lazy in the same way, with `deepCopy`
on every unmarshalled tree, and would not be simpler. The clone's `sourcePartStore` is the memory
store, so it no longer depends on the original's container; its load options are the original's;
and it is of the same class, because `OpcPackage.load` picks the class from the main part's
content type (a mismatch throws rather than returning the wrong class).

Settings that are not in the package travel by a `protected copyPackageSettingsTo(target)` hook:
`author`, `trackedChangeDate`, `fonts.defaultTheme` and `xpathEngine` for WordprocessingML,
`fonts.defaultTheme` for PresentationML. docx4j carries only its `name`, which has no counterpart
here. `clone()` is asynchronous, as everything that marshals here is; docx4j's is synchronous.

### 16.5 Docs and examples

`docs/guides/`: `getting-started-node.md` (install, load, the content API, effective reads, the
parts underneath, the parity promise), `office-addin.md` (flat OPC in and out, the `./office-js`
shim for testing add-in code in Node, the bundling rules above), `presentationml-spreadsheetml.md`
(what the two packages give, and what they do not: no content API), `parity.md` (what the goldens
are and how a consumer runs the harness against a newer docx4j). Linked from `README.md`, whose
status paragraph now says the CR is complete and whose add-in section defers the bundling detail
to the guide.

`examples/node/`: `hello.mjs`, `report.mjs`, `pptx.mjs`, `xlsx.mjs`, `directory.mjs`, each run as
`node examples/node/<name>.mjs [file]` against the built `dist/` with a comment saying what the
published import is. `examples/office-addin/`: `manifest.xml`, `taskpane.html`, `taskpane.ts`,
`edit.mjs` (the edit itself, written against a `Word.Body` and so runnable on either side),
`edit.d.mts`, `office.d.ts` (the few Office globals used, so `@types/office-js` is not a
dependency of this repository), `run-in-node.mjs` and a `README.md` with the sideload steps and
the esbuild command. `tsconfig.json` there is run by a new `npm run typecheck:examples`, which
`npm run typecheck` chains.

`test/examples.test.mjs` runs every example in a child process against `test/fixtures/` and
asserts exit 0 and an expected line of output, so an example cannot rot; the add-in's Node half is
one of them. Acceptance checks **12** (a created pptx opens in PowerPoint) and **13** (a created
xlsx opens in Excel) are added to `test/README.md` with a script, and are outstanding, as check 3
still is.

### 16.6 What stays out

Deliberately not in this CR, and not in this package yet (docx4j-python's CR-002 12.9 lists the
same set):

- **External resources.** External relationship targets (hyperlinks, linked images) stay in the
  relationships and `getPart(rel)` is undefined for them; `OpcPackage.externalResources` of
  section 5.5 is not implemented (section 12 already recorded this).
- **Digital signatures**, encryption, VBA, and the strict (`purl.oclc.org`) conversion on load.
- **`DrawingPropsIdTracker`** and the `docPr`/`cNvPr` id uniqueness docx4j maintains when parts
  are combined; `insertOoxml` does not renumber drawing ids.
- **PresentationML and SpreadsheetML beyond the parts**: `ResolvedLayout`, `ShapeWrapper` and the
  placeholder resolution, notes and handout creation, cell values as values, formulas, the shared
  strings table as an index.
- **`BestMatchingMapper`** over installed or embedded fonts, glyph coverage and font metrics
  (section 6.3 and 14.2 already said so); the `Mapper` interface is the seam.
- **Table conditional formatting** in `getEffectiveTableStyle` (`docx4j/table-conditions`).
- **ZIP64** archives, which the zip container rejects.

**Decided 2026-09-19 (Jason):** `docs/` and `examples/` stay GitHub-only; the npm package ships `dist`, the README (which links the guides), LICENSE and NOTICE, as `package.json`'s `files` already says.

## 17. Acceptance findings of 2026-09-19

The manual acceptance run of `test/README.md` (Word 365, PowerPoint and Excel, checks 3 and 9 to
13, the first time a created pptx and xlsx were opened in Office) found five things. One, check
13's "values only on the second tab", was the check's expectation being wrong and not the output:
the script writes the values into `Sales` and then inserts `Cover` at index 0, so `sheets` order
is Cover, Sales, Notes and the values are on the second tab. The check text now says so. The
other four are real, and are fixed here; finding 2 (a tracked row deletion struck nothing out) is
the content API's and is written up in CR-002 section 13. The three that belong to this CR:

### 17.1 A created document opened in compatibility mode

Word 365 opened every `createPackage()` document with **Compatibility Mode** in the title bar,
because the settings part was empty: with no `w:compat/w:compatSetting compatibilityMode`, Word
reads the document as pre-2013 and lays it out by the old rules. `createPackage` now writes, in
Word's own order and all with `w:uri="http://schemas.microsoft.com/office/word"`:

```xml
<w:compat>
  <w:compatSetting w:name="compatibilityMode" w:val="15"/>
  <w:compatSetting w:name="overrideTableStyleFontSizeAndJustification" w:val="1"/>
  <w:compatSetting w:name="enableOpenTypeFeatures" w:val="1"/>
  <w:compatSetting w:name="doNotFlipMirrorIndents" w:val="1"/>
  <w:compatSetting w:name="differentiateMultirowTableHeaders" w:val="1"/>
  <w:compatSetting w:name="useWord2013TrackBottomHyphenation" w:val="1"/>
</w:compat>
```

- These are the six settings Word 365 writes for a new document, measured by the docx4j session
  on a re-save (2026-09-19); Word also writes a locale-dependent `w:themeFontLang`, which neither
  library writes. The shape is docx4j's `DocumentSettingsPart.setWordCompatSetting(name, val)`.
- **Parity, not a departure.** docx4j wrote no mode at `a58cf10b8` (its `getCompatibilityMode()`
  answered 12 when absent, and its javadoc said Word opens such a document in compatibility mode);
  this was first fixed here as a departure, then Jason decided it for docx4j the same day:
  `9de10aac9` adds `DocumentSettingsPart.setCompatSettingsAsWord365` and `createPackage` writes the
  same six in the same order, so a created settings part matches docx4j's byte for byte. Any
  theme year is consistent with mode 15 (the mode is the layout rules, the theme the fonts).
- `createPackage({ compatibilityMode })` sets the value (`'15'` by default) and `''` leaves the
  setting out, which is docx4j's behaviour back. Nothing else in the package reads it, and a
  **loaded** document's settings part is untouched, as every untouched part is (a test asserts
  `word/settings.xml` of `loadAndSave.docx` is byte-identical after a round trip).

### 17.2 A created presentation's slides were blank

PowerPoint opened `check12-*.pptx` with slides that had nothing on them: no "Click to add title",
nothing to type into. It offers those only where the **slide** holds placeholder shapes — a
`p:sp` whose `p:nvSpPr/p:nvPr/p:ph` names a type or an index — matching ones the layout defines;
docx4j's `COMMON_SLIDE_DATA`, which section 16.1 took as the slide and the layout template, is an
empty `p:spTree`, and docx4j's own `createPackage` never gets as far as a slide.

`src/parts/pml/placeholders.mts` now builds them with the generated factories (`createShape`,
`createShapeNvSpPr`, `createNvPr`, `createCTPlaceholder` and the `org_docx4j_dml` creators), not
markup, so the shapes are the object model's and a caller can go on editing them:

- **The layout** (`createPackage`) becomes a Title and Content layout: `type="obj"`,
  `preserve="true"`, `p:cSld/@name` "Title and Content", and two placeholder shapes — the title
  (`p:ph type="title"`) and the body (`p:ph idx="1"`) — each with an `a:xfrm` and PowerPoint's
  prompt text. The geometry is PowerPoint's own Title and Content layout **as fractions of the
  slide size** (`placeholderGeometry`), so it holds at every `slideSize` rather than only at 4:3.
  This is where docx4j's template markup stops being written byte for byte: the layout part is
  unmarshalled during creation, as the master already was.
- **Every slide** carries one shape per placeholder its layout defines, bar the date, footer and
  slide-number ones (`slidePlaceholdersFor`, which reads the layout with `readContents()` so that
  a loaded package's layout keeps its byte-for-byte round trip), with the layout's `type` and
  `idx` and **no geometry of its own** — inheriting the layout's is what a placeholder is for.
  Empty, so PowerPoint shows the prompts.
- **`addSlide({ title, body })`** and `createPackage({ title, body })` fill them: the title in the
  first title placeholder, `body` in the first body one as `a:p/a:r/a:t`, one paragraph per line
  (a string is split on `\n`; an array is one entry per paragraph).
- Still out, as section 16.6 has it: `ResolvedLayout`, `ShapeWrapper` and real placeholder
  resolution, notes and handouts, and the master's `p:txStyles` (the layout's own `a:lstStyle`
  would be the place for per-level text styles, and neither is written).

### 17.3 `mc:Ignorable` naming a prefix nothing declared

`check13-remarshalled.xlsx` was the only file any Office application offered to repair. Measured:
a re-marshalled `xl/workbook.xml` keeps `mc:Ignorable="x15 xr xr6 xr10 xr2"` — the model binds
`mc:Ignorable` on `CT_Workbook`, so the attribute survives — but declares only `mc`, `r`, `x15`
and `xr`. `xr:revisionPtr` (with its `xr6:` and `xr10:` attributes) and `workbookView/@xr2:uid`
are **not in the object model** and are dropped on unmarshal, so nothing in the tree resolves to
`xr6`, `xr10` or `xr2`; the runtime declares one prefix per entry of the objects package's
`NAMESPACE_PREFIXES`, which has none of those three, and the facade's
`stripUnusedNamespaceDeclarations` only ever *removes* declarations the marshaller made — it
keeps the ones `mc:Ignorable` names, but it cannot add a declaration that was never made. An
`mc:Ignorable` naming an undeclared prefix is invalid (ECMA-376 Part 3 10.1.2) and is what Excel
repairs.

The fix is in `src/parts/XmlPart.mts`, so it covers **every** XML part of every format — the same
class of defect the comments part had before objects 0.1.3:

- At unmarshal (`getContents` and `readContents`), the source root element's namespace
  declarations are recorded, prefix to URI, from the DOM as parsed — before the MCE preprocessor
  touches it.
- At marshal (`marshalToNode`, and so `getBytes` and `getXml` too), every prefix the marshalled
  root's `mc:Ignorable` names that the root does not declare is declared: from the recorded map,
  else from `NAMESPACE_PREFIXES`. A prefix neither knows is **dropped from `mc:Ignorable`** and
  warned about once — an undeclared prefix is the one thing that must not be written.
- `test/roundtrip.test.mjs` asserts it for `loadAndSave.xlsx`'s workbook (all five prefixes) and
  worksheet and for `loadAndSave.docx`'s document and settings parts.

`CLAUDE.md`'s "Nothing here touches prefixes" is now narrower than it was: the facade still owns
the prefix table and the `xml` rule, and this is the one thing a part adds on top of it.

**Closed in the objects package the same day** (`d8bc453`, unreleased): the facade's root pass now
declares every prefix `mc:Ignorable` names from its table (grown by `xr2`, `xr3`, `xr6`, `xr10`) and
drops a prefix it cannot resolve with one warning, docx4j's `McIgnorableNamespaceDeclarator` in
both halves. The `XmlPart` handling here stays: it re-declares from the source root's own
declarations, so it covers a prefix bound to a namespace the table does not know, or bound
unconventionally, where the facade can only drop it. A facade that carried the source's
declarations through the unmarshal would be a larger change and is not asked for.


### 17.4 To relay

- **To the objects package** (its own CR, not fixable here): the facade cannot *add* a namespace
  declaration, only keep or remove one, so an `mc:Ignorable` prefix the model does not bind loses
  its declaration. `stripUnusedNamespaceDeclarations` could declare an ignorable prefix it knows
  from `NAMESPACE_PREFIXES`, and `NAMESPACE_PREFIXES` could gain `xr2`
  (`.../office/spreadsheetml/2015/revision2`), `xr6` (`.../2016/revision6`) and `xr10`
  (`.../2016/revision10`), which docx4j's own prefix table also lacks. This package no longer
  depends on that, but it is the right place for the general rule.
- **To docx4j**: `CT_Workbook` drops `xr:revisionPtr`, and `CT_BookView` drops
  `workbookView/@xr2:uid`, because the schema itself does not allow them — checked in
  `xsd/sml/sml_ECMA376_4ed_transitional.xsd`, where `CT_Workbook` (line 4286) and `CT_BookView`
  (4368) have **no `xs:any` and no `xs:anyAttribute`**; docx4j's own comment in that file quotes
  the very markup and the four `xr*` namespaces. The loss is Excel's revision bookkeeping, which
  Excel rewrites on its next save, so it is lossy but not a repair trigger now that the
  declarations are written. `CT_Worksheet` has no `mc:Ignorable` attribute at all, so a
  re-marshalled worksheet loses the attribute outright (nothing is left undeclared, so again no
  repair). Whether to give those types the wildcards is docx4j's call, as the `mc:AlternateContent`
  laxity of section 15.4 was.
- **A CR candidate here**: this package reports nothing about content it drops on unmarshal.
  docx4j-python has a skipped-content report; with one, `xr:revisionPtr` and `xr2:uid` would have
  been visible the moment the workbook was unmarshalled instead of at a repair prompt in Excel.
  Worth a small CR: a per-part list of elements and attributes the model did not bind, collected
  during unmarshal behind a load option.

**Closed in docx4j the next day** (CR-022 phase 1, `16844ff03`, 2026-09-20, unpushed): the sml
binding keeps `mc:Ignorable` on the worksheet, styleSheet, table, comments, pivot, connections and
queryTable roots and re-declares every prefix it names; `x14ac:dyDescent` and `x14ac:knownFonts`
are typed; `xr:revisionPtr` with its `xr6`/`xr10` attributes is kept; a worksheet's
`mc:AlternateContent` for controls is kept whole; `x14`/`x15` `extLst` content unmarshals typed;
`x12ac`, `x16`, `xr16` join the prefix table. For this package that is an objects regeneration
item (the schema and bindings changed); the parity goldens are unaffected (the harness records
WordprocessingML only), and once the objects package regenerates from that commit a re-marshalled
workbook here keeps what docx4j now keeps. Still `DefaultPart` there until phase 2: slicer,
slicerCache, timeline, timelineCache and ctrlProp parts; still Fallback-resolved: the DrawingML
`mc:AlternateContent` in drawings. Oracle test to mirror when it lands here:
`docx4j-core-tests org.xlsx4j.ExcelExtensionsTest` on `cr022-slicers-timelines.xlsx`.

**docx4j CR-023 (2026-09-20, `cdb44df87` and `0b1dd27a4`, unpushed): Word's extension attributes
kept.** The oracle now binds what its schema had no reference for where Word writes it:
`w16du:dateUtc` on every tracked change and comment, `w15:restartNumberingAfterBreak` on
`w:abstractNum`, `w16cid:durableId` on `w:num`, `w14:noSpellErr` on `w:p`,
`w16sdtdh:storeItemChecksum` on `w:dataBinding`, `w16sdtfl:formattingAllowed` on `w:sdtPr`,
`w16se:symEx` in run content, and the `cei` namespace inside commentsExtensible's `extLst`
(prefix table: `cei`, `cr`). For this package an objects regeneration item, after which a
re-marshalled part here keeps them too; the parity goldens record effective properties and labels,
none of which these attributes enter, so the weekly workflow is expected to show no change. Two
writer findings from docx4j's Word check: an `mc:Ignorable` naming an undeclared prefix (`cr` on
commentsExtensible.xml, undeclared by docx4j's save until CR-023) is repaired by Word, which
section 17.3's re-declaration already prevents here and `test/ignorable.test.mjs` now checks for
every XML part of every fixture; and docx4j's settings part used to rebuild its `mc:Ignorable` on
save, discarding Word's list, which this package never did (the `ignorable` property round-trips).

**docx4j CR-024 (2026-09-20, `bcb4c7f57`, `85f37d05e`, unpushed): `mc:AlternateContent` kept in
DrawingML hosts.** The oracle keeps both branches of the `mc:AlternateContent` a spreadsheet
drawing holds (a check box's `a14` shape under `xdr:wsDr`, whose Fallback is empty; a slicer's or
timeline's `graphicFrame` in an anchor, Choice `Requires` `a14` / `tsle` / `sle15` with an
"Excel 2010 or higher" box as Fallback) and the one every chart's `c:style` position holds
(Choice `c14:style` 102, Fallback `c:style` 2). What this package's resolve-on-load preprocessor
(section 5.6) does with the same three, by `UNDERSTOOD_NAMESPACES`: the check box's Choice
(`a14`, understood) is taken, so the drawing survives where resolving to the empty Fallback would
delete it; the chart's Choice (`c14`, understood) is taken, so a re-marshalled chart part carries
`c14:style` and loses `c:style` 2 (Excel, Word and PowerPoint read `c14`; a consumer that knows
only `c:style` sees none); the slicer's and timeline's Choices (`tsle`, `sle15`, `sle`: not
understood, no modules for them) fall to the "Excel 2010 or higher" box, so **a re-marshalled
spreadsheet drawing loses its slicers and timelines** (an untouched drawing part keeps its bytes).
CR-004 Phase B adds those namespaces to `UNDERSTOOD_NAMESPACES` once the objects package carries
their modules, which it cannot yet: docx4j's `xsd/` has the timeslicer and 2010 slicer schemas but
`ROOT.xsd` imports neither, and the 2012 slicer (`sle15`) schema is absent, so no regeneration can
type that content today (objects session, 2026-09-20; the docx4j session is asked to bind them).
Until then the Fallback behaviour stands, documented. The same regeneration (objects `529a957`)
does admit `mc:AlternateContent` in the DrawingML hosts, which is what lets a spreadsheet drawing
survive a round trip with `mcePreprocess: false` at all. Two writer findings
from docx4j's Excel opens, both covered here: every root's `mc:Ignorable` prefixes declared
(section 17.3, and `test/ignorable.test.mjs` re-marshals every part of every fixture, the forced
re-save docx4j recommends, since an untouched part is written from bytes and hides the fault);
and a Choice's `Requires` prefix declared, which does not arise here because the preprocessor
removes the wrapper before unmarshalling. docx4j's `x` alias for the SpreadsheetML main namespace
on `x14`/`x15` parts, declared beside the default binding, is what the source-root map here
re-declares when those parts are typed (Phase B).

### 17.5 Objects 0.1.5 tried and held back: `a14:m` equations in DrawingML text (2026-09-20)

`@docx4j/generated-objects-ts` 0.1.5 (npm, tag `d21e144`; the CR-021 to CR-024 regenerations,
the `x14`/`x15` modules CR-004 Phase B needs, `mc:Ignorable` prefixes declared by the facade,
CR-003 phase A) was tried here with `^0.1.5`: typecheck clean, 474 of 478 tests pass. The four
failures are two findings.

**1. A regression: `loadAndSave.pptx` slide 2 and `loadAndSave.xlsx` `drawing1.xml` no longer
unmarshal** (`load: pptx and xlsx`, the two `ignorable` cases): `Element a:rPr could not be
unmarshalled as is not known in this context and the property does not allow DOM content`. The
slide's placeholder holds an equation the way PowerPoint writes one: `mc:AlternateContent`,
Choice `Requires="a14"` with the shape whose `a:p` carries `a14:m` holding `m:oMathPara`,
Fallback a picture. `a14` is understood, so the Choice is taken. Under 0.1.4 `a14:m` was not
admitted by `EG_TextRun` and stayed DOM, written back as it came: lossless. docx4j CR-021
section 8.9 admitted `a14:m` and made `CT_TextMath` a **lax** wildcard, reasoning that the
PresentationML and SpreadsheetML JAXB contexts do not carry the OMML classes, so the equation
stays DOM there and is typed only in the WordprocessingML context. Jsonix has one context with
every module, so the lax wildcard unmarshals `m:oMathPara` typed through `org_docx4j_math`, and
inside a DrawingML text body **`m:r` and `m:ctrlPr` hold `a:rPr`**, where `shared-math-2ed.xsd`'s
`CT_R` and `CT_CtrlPr` reference only `w:EG_RPr` / `w:EG_RPrMath`. JAXB would drop the `a:rPr`
silently; Jsonix throws. Every pptx or xlsx with an equation in a text body is affected once
its part is unmarshalled (an untouched part still round-trips from bytes).

Two fixes, both upstream of this package, relayed to the objects session (which owns the
regeneration and the release) for docx4j's decision:

- *Minimal, mirrors Java's pptx/xlsx outcome exactly:* `CT_TextMath`'s wildcard DOM-only
  (`processContents="skip"` in `oart14docprop.xsd`, `@XmlAnyElement` without `lax`, Jsonix
  `allowTypedObject: false` as `CT_XmlData` and the VML text box already have). Costs Java the
  typed equation in the WordprocessingML context, where an `a14:m` in a DrawingML text body is
  rare (Word's text boxes are `w:p`).
- *Complete, "admit what Office writes" one level down:* `CT_R` and `CT_CtrlPr` in
  `shared-math-2ed.xsd` admit `a:rPr` beside the `w:` group ([MS-ODRAWXML] 2.3.1 puts OMML in
  DrawingML text, where the run properties are DrawingML's). Then the equation is typed in every
  context, here and in Java. More schema surface; the math module gains a reference to `dml`,
  which every context that holds math already loads.

Until an objects release carries one of them this package stays on `^0.1.4` with the lockfile at
0.1.4, and CR-004 Phase B remains blocked on the objects package after all (the registry entry
`objects-ts/cr022-regeneration` is done, but the release that carries it is unusable here).
**A consumer of the published core-ts 0.1.0 is exposed today:** its range is `^0.1.1`, so a fresh
install resolves the objects package to 0.1.5 and a pptx or xlsx with an equation fails to load.
The next core-ts release should pin below 0.1.5 or above the fix, whichever comes first.

**2. A promise that moved, to update when the upgrade lands:** `mce.test.mjs`'s
"preprocessing off" case asserted that `loadAndSave.docx`'s main document part rejects on
`getContents()` with `mcePreprocess: false` (section 12 item 3: a `w:drawing` inside `mc:Choice`
could not be unmarshalled). With the CR-021 regeneration it unmarshals: the element is an
`org_docx4j_mce.AlternateContent` whose Choice (`Requires="wps"`) and Fallback keep their
`w:drawing` and `w:pict` as DOM, a save writes both branches back (measured: one
`mc:AlternateContent`, one Choice, one Fallback, three `w:drawing`, one `w:pict`, as in the
source), and reloading the saved bytes with preprocessing on takes the Choice. So with the
option off a part keeps every branch, which is what section 15.4 called a lossless kept branch.
The test becomes: contents unmarshal, the `mc:AlternateContent` is typed with both branches, the
re-marshalled part carries both, and the section 5.6 default (resolve on load) stays, for the
reason section 12 item 3 gives in its second half: the content API cannot see a `w:drawing`
inside a DOM branch. Section 12 item 3's first half is then history.

**Relay (2026-09-20):** the objects session reproduced the failure on the slide with the `a14`
Choice inlined (its own round-trip tests read the Choice content as DOM straight from the zip,
which is why they passed) and passed both fixes to the docx4j session recommending the second
(`a:rPr` admitted in `CT_R` and `CT_CtrlPr`), the first if the second needs thought; objects
0.1.6 follows the docx4j commit and a regeneration, and the two fixture parts join objects
CR-004's fidelity set. The published-consumer exposure is with Jason for an `npm deprecate` of
objects 0.1.5.

**docx4j CR-025 (2026-09-20, `34948f8e9` on `VERSION_17_1_1`, unpushed; the branch becomes
`VERSION_17_2_0` for the release): `a:rPr` in OMML.** docx4j took the second fix. Once objects
0.1.6 carries the regeneration, an equation in a DrawingML text body typed through the WML
context has, on every `m:r`, the `a:rPr` as the first element of its content (Java: a
`JAXBElement<CTTextCharacterProperties>` in `CTR.content`, factory `createCTRRPrDml`), and on
every `m:ctrlPr` a property `rPrDml` (`CTTextCharacterProperties`) beside the `w:` ones; docx4j's
typed round trip of `loadAndSave.pptx` slide 2's fragment equals PowerPoint's bytes canonically
except booleans' lexical form (`i="1"` to `i="true"`). In its PresentationML and SpreadsheetML
contexts the equation stays DOM (CR-021 §8.9), so this package, one context, types more than the
oracle does there. Test to mirror when 0.1.6 lands: `docx4j-core-tests
org.docx4j.jaxb.OmmlInDrawingMLTextTest` (three tests); the CR is
`docs/developer/change-requests/CR-025-omml-in-drawingml-text.md`.

**Upgraded to objects 0.1.6 (2026-09-20, npm tag `f71ab46`; `^0.1.6`).** Tested twice before the
release at the objects session's request (main `daf9395`, then `02c1d2a` with the `c16r3` chart
attributes unqualified and that prefix in place of `c173`): typecheck clean, every test but the
moved promise passing, the 45 goldens unchanged, each fixture's `chart1.xml` now keeping
`c16r3:dispNaAsBlank`'s `val` through a typed round trip. Landed here: `mce.test.mjs`'s
"preprocessing off" case asserts the new promise (one typed `mc:AlternateContent`, Choice
`wps` holding a DOM `w:drawing`, Fallback a DOM `w:pict`, the saved part with every branch tag
counted equal to the source); `test/omml.test.mjs` mirrors `OmmlInDrawingMLTextTest` on both
fixture parts (16 typed `m:r` whose first content item is the `a:rPr` named value of
`CTTextCharacterProperties`, 7 `m:ctrlPr` with `rPrDml`, all Cambria Math; the re-marshalled part
with the same counts and its one `a14:m`). 480 tests. The harness's `c173` becomes `c16r3` at the
next golden regeneration. Objects 0.1.5 stays on npm undeprecated (no external users yet, Jason's
call); this package's `^0.1.6` keeps a consumer of the next core-ts release off it. CR-004 Phase B
now waits only on Phase A.

### 17.6 Strict OOXML: no conversion here, and what a port must get right (2026-09-21)

Asked by the docx4j session whether this package shares two defects fixed that day in
`org/docx4j/jaxb/mc-preprocessor.xslt`'s strict path (`VERSION_17_2_0` `7892cb501`: the
points-to-twips conversion had been commented out, so every Word-saved strict document failed
on `w:spacing w:after="8pt" w:line="12.95pt"`; `3f4776c35`: strict keeps the 2010
shape/group/canvas elements in the wordprocessingDrawing namespace, and mapping it to `wp:`
left `wp:wsp`, `wp:wgp`, `wp:wpc` and their children unknown, so Word refused the result). It
does not: there is no strict-to-transitional path. Measured on docx4j's
`docx4j-core-tests/src/test/resources/strict/strict-chart.docx`: the package loads (the
`purl.oclc.org` officeDocument relationship type is recognised, `MAIN_PART_RELATIONSHIPS`, and
every part gets its typed class), an untouched package round-trips, and the first
`getContents()` throws `{http://purl.oclc.org/ooxml/wordprocessingml/main}w:document ... not
known in this context`. Section 12 and section 16 already record strict conversion as out of
scope.

When it is ported (a CR of its own; the per-part DOM preprocessor hook that runs the section 5.6
MCE resolution is where a strict rewriter goes, before it), the oracle is `StrictLoadTest` and the
four Word-saved strict samples under `strict/`, and the two fixes are the specification: for the
strict WordprocessingML namespace, x2 for the half-point family (`sz`, `szCs`, `kern`,
`position`, `hps`, `hpsRaise`, `hpsBaseText`) and x20 for every other point value (spacing,
indents, table widths, tabs, `defaultTabStop`, drawing grid; `w:line` under `lineRule="auto"`
too, Word writing the 240ths as points); and the wordprocessingDrawing 2010 elements mapped to
`wps:`/`wpg:`/`wpc:` by root name, children by nearest such ancestor.

**The repair side of the same stylesheet (asked the same day, Jason's actual question).** docx4j's
`JaxbXmlPart.unmarshal` tries plain JAXB, and on failure transforms the part through
`mc-preprocessor.xslt` and unmarshals again with the validation handler set to continue; the
stylesheet's repairs, apart from strict and `mc:AlternateContent`, are (1) decimals in twips
attributes rounded (Google Docs 2014/2015, pandoc 2.2.1), (2) SSRS 2012's empty `rsid*`
attributes dropped and empty `pgMar` header/footer made 0, (3) `wordml201011:*` attributes
dropped, (4) since 17.2.0 malformed nesting repaired: a `w:r` inside a `w:r` hoisted, a `w:p`
inside a `w:r` or `w:hyperlink` split out as a paragraph of its own (`MalformedNestingTest`, 17
shapes). This package has none of it: the section 5.6 preprocessor is the only load-time DOM
pass, run always and once, and Jsonix has no continue mode, so what the model rejects is fatal for
that part, never silently dropped. Measured against objects 0.1.6: (1) the part's `getContents()`
throws `Argument [259.2] must be an integer` for every attribute the model types as Int
(`w:spacing`, `w:ind`, `w:gridCol`, `w:trHeight`, `w:pgSz`, `w:sz`); `w:tblW/@w:w` is a string
type and keeps the decimal; (2) empty `rsid*` kept as `""` and written back, empty `pgMar`
header, footer and gutter all become 0; (3) dropped, as any unknown attribute; (4) kept typed
(`w:p` and `w:r` are global elements and the CR-021 lax wildcard admits them inside run
content), marshalled back byte-equivalent, and invisible to the content API: `textOf` of
`<w:r><w:t>a</w:t><w:p>…</w:p><w:t>b</w:t></w:r>` is `ab`, so search, `replaceText`,
`Paragraph.text` and the Emulator walk pass over the nested subtree without reporting it. The one
repair here that docx4j lacks is on the marshal side: an `mc:Ignorable` prefix nothing declares is
dropped from the list with a warning (`XmlPart.declareIgnorablePrefixes`). A repair CR, when a
corpus asks for it, goes in the same preprocessor hook, with the stylesheet's rounding table and
`MalformedNestingTest`'s shapes as its oracle; item (4) is the one that matters most, since it
silently hides text from every reader here.

### 17.7 A vmlDrawing part could never unmarshal, and why nothing here noticed (2026-09-23)

The objects session found that no `vmlDrawing` part had ever unmarshalled through
`@docx4j/generated-objects-ts`: the root of such a part is `<xml>` in **no** namespace, and
docx4j's wrapper schema declared that global element in an invented `urn:docx4j:vml:root`, so
every real part threw `Element [xml] could not be unmarshalled`. docx4j CR-026 (`b4ca0d98a`,
`VERSION_17_2_1`, shipping in 17.2.1) drops the invented target namespace and widens the wildcard
to `##any`; objects `f9996c3` is the regeneration, proposed as 0.1.7.

Nothing here noticed because `VMLPart` extends `DefaultXmlPart`: a DOM part that never goes
through Jsonix. Measured on `loadAndSave.xlsx`: the part loads as `VMLPart`, `unmarshalAll()`
passes over it (it walks `XmlPart` only), and after `unmarshalAll()` and a save the part is
byte-identical to the source.

**docx4j's companion defect does not exist here.** There, save converts every not-yet-read part,
so a *strict* package holding a vmlDrawing could not be saved at all — a strict workbook with a
comment was unsaveable. This package re-marshals only the relationships parts,
`[Content_Types].xml` and `XmlPart`s that were unmarshalled or had contents or bytes set;
everything else is copied from `sourcePartStore` as bytes. A part nothing read is never
converted, so an unreadable part cannot cost a package its save. That is the section 3 contract
the round-trip suite asserts, not a side effect of the vml part staying DOM.

Tested for the proposed 0.1.7: typecheck clean, 480 tests, no golden moves. With the fix the root
does unmarshal typed (`{}xml`, `org_docx4j_vml_root.Xml`, holding `o:shapelayout`, `v:shapetype`
and `v:shape`) and re-marshals with every element of the source present, so typing `VMLPart` as an
`XmlPart<Xml>` is now possible. Not done: the DOM costs nothing, keeps the byte-for-byte round
trip, and the comment and form-control paths read their VML through the part's DOM today; a CR
if a caller wants the tree.
