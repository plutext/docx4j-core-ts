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
answers by the parity goldens, 62 since CR-007) and CR-002 phases B to I, the whole content API (`Body`,
`Paragraph`, `Range`, `Font`, `Table`, `InlinePicture`, `ContentControl` with `XmlMapping` and
the typed kinds, `Comment`, `TrackedChange`, `List` and `ListItem`, custom XML parts with XPath,
search, `replaceText`, `insertOoxml`, addresses, `outline`; the `Word` shim on `./office-js`) are
implemented; CR-002 phase A is the objects package's `builders/wml`. CR-007 (table conditions in
the `PropertyResolver`, the port of docx4j CR-030: `TableContext`, `CellContext`, the overloads
taking one, 62 parity goldens; and the content API's effective reads passing `Paragraph.cellContext`)
is implemented, released in 0.3.2. CR-002 section 41 (`styles.ensure` falling back to Word's built-ins
from docx4j's `KnownStyles.xml`) and section 42 (columns added and removed in Word's three table layout
modes, held to Word 365's saves, checks 39 to 41) are implemented, released in 0.4.0 (2026-10-09).
So are CR-001 Phases C and D
(pptx and xlsx creation, `./node`, `clone()`; font discovery over private reads), CR-004 (the Excel
2010/2013 extension parts), CR-005 phase A (the FontoXPath engine) and its sections 8 and 9 (escaped
XHTML through the `html` module, and picture bindings; implemented 2026-10-08, released in 0.4.0) and its section 11 (Word fragments, the editor's
ED-005 proposal 46 parts 1 and 2: a control tagged `od:progid=Word.Document` and Word's `w15`-mapped rich text over Flat OPC filled
as Word 365 resolves docx4j's altChunk, styles, lists, notes and ids reconciled by Word's rule, or with that altChunk; implemented
2026-11-11, unreleased; its section 11.6, `insertOoxml` of a package reconciled by the same rule, held to Word's own insertOoxml), and CR-006 (`selectObjects` and
`Body.select`, XPath over the tree, held to docx4j by 62 XPath goldens, one per fixture). CR-008 phase 1 (docx4j CR-034's colour merge rules in the
catalogue, 2026-10-12) and CR-010 (every branch of `mc:AlternateContent` kept on load) are released in
0.5.0 (2026-10-19), with the rest the CR index (`docs/change-requests/README.md`) lists for it.
CR-011 (every story's tracked changes in one listing, and filters on it and the resolve-all verbs,
docx4j CR-043's, held to its answers by 96 revisions goldens) is implemented, released in 0.6.0 (2026-10-25). So
are CR-002 section 54 (docx4j CR-043 section 15: a comment whose range a resolution empties goes
whole wherever its markers were, and the comment parts with the last, as Word 365 does; the
revisions goldens record each resolution's comments since) and CR-012's phases 0 to 3 (the text model's conformance suite and agreement test, copied from
docx4j into `test/text-extraction/` with their runners; REQ-510's level owner in the `Emulator`; the
plain-text model, `Texts.of` and the block model in `src/model/text/`; Markdown, `MarkdownExporter`, docx4j's
byte for byte; 786 suite rows passing and the agreement test's 588 comparisons agreeing; phase 4, the readers:
`Paragraph.text` and `search` in Office JS's characters and rules in Word's default view, held to Word 365's checks 48,
48b and 51, a kept branch's text read and edited (ED-009), text boxes at `textbox:` addresses read-only, `outline()` on the
model and async, and per block under a context, section 17.13's `{ blocks }`; phase 6, the block-model parity goldens:
282 text goldens from docx4j 17.3.4's `Texts.of`, four readings each, at zero differences outside what phase 5 and REQ-303
govern), and CR-013, all four phases
(`mc:AlternateContent` at block, row and cell level: docx4j's `mc-fun` probes in `test/mc-fun/`, the walker audit, the
text model's rows and cells in the selected branch, `mcePreprocess: true` over the readers' set; the content API's readers and
the tracked-change listing in the selected branch, a view in a branch read-only; the mutators in every branch, a listed
change's twins resolved with it, as docx4j's), and CR-014's phases 0 to 2 (ranges across paragraphs, as Office JS's: checks
52 and 53 run in Word 365; one `Range` with two ends, its reads and its writes held to Word's runs and saves through the
Word shim; phase 3's comments, `^p` and controls too, held to checks 52 and 54), and CR-008 phase 2 (docx4j's colour resolver,
`org.docx4j.model.color`, in `src/model/color/`: every probe held to docx4j's own tallies; `Font.color` reads a theme colour
as Word draws it). CR-015 (run properties read as Word 365 reads them: on/off, enumerations, colours; docx4j CR-060 and
CR-061) is implemented, released in 0.6.0, on objects 0.5.0 (its CR-013 A to C), and held to Word 365's check 60, where Office JS
differed only in Word's green highlight until objects 0.6.0 took Word's palette (its CR-014: checks 60 to 62 now Office JS's throughout). CR-017 (an integer
attribute value Word 365 reads differently reported on load and on save, never rewritten; docx4j CR-065 and CR-066, the objects package's CR-015 and CR-016) is
implemented, P1 to P3, released in 0.7.0 (2026-11-09), with CR-002 section 65 (the measurement setters refusing past Word's storage, check 71; the spacing, line, list and font setters at Word's own
limits below it, and `lineSpacing` writing Word's auto rule, check 71b, section 65.9) and its section 65.3 (the setters' table the objects package's). So is CR-002 section 63
(writing into a control showing its placeholder as Word does, the placeholder no content, in Word's form, `w:placeholder` naming a glossary
entry, and in the bare form alike; placeholders written in Word's form, their entries in a glossary made on save; checks 63 to 67, check 63
run again with Design Mode off, since Design Mode changes what a write into a placeholder does). CR-016 (misplaced content refused before it reaches a saved file, docx4j CR-064's rules
in the objects package's CR-017: a save throws `MisplacedContentException` for a run in the body, a table in a paragraph, a string in a run, ..., unless `pkg.checkContentOnSave`
is off; `checkContent()` on request; a `target.misplaced` part warning per child read; check 72 for Word's repairs) is implemented, released in 0.7.0. CR-018 (dates in the package's zone, the portfolio's
date rule of 2026-10-28: `pkg.dateTimeZone` for every instant's wall clock, a comment stamped to the minute, Word's bound-date conversion and the
save's `dcterms` times as opt-ins) is implemented, released in 0.7.0. So is CR-020 (bound controls as Word repairs and formats
them, docx4j CR-058: a `w:storeItemID` naming no part bound to the part its XPath fits, docx4j's `StoreItemResolver`, by every
reader and the bind, which writes the part's ID back as Word's save does; `XmlMapping.isMapped` as Office JS reports it; a bound
date's `am/pm` and the document's default language; held to Word 365's saves and checks 73 and 73b). So is CR-021 (a `w:br` outside a
`w:r` read as Word reads it, docx4j CR-073's three rules as a DOM pass ahead of the package's hook, `br.outside_run` on the part; held
to Word 365's saves of docx4j's thirteen probes; check 72's 72e and 72j now read and saved as Word saves them). So is CR-024 (a package holding two entries of one name, or more than one officeDocument relationship, refused, WD REQ-206; `repair: 'word'` as Word reads each, phase 2 in its section 9). So is CR-031 (the `repair` load option, WD REQ-207 and REQ-208: `refuse`, the default, `word` following Word's repairs, `all` adding this package's own; `duplicateEntries` folded into it, an `External` officeDocument relationship beside an internal one refused but at `all`, one alone refused as no main part, a relationships part in a wrong namespace read by local name under `word`, `repaired` on the package and the text model's metadata; README's "Repairs"). CR-032 (a relationships part that isn't well formed read as Word's repair reads it at every value of `repair`, docx4j CR-075 D5 and D6, released in its 17.3.6: the relationships before the break, `part.not_read` with the count kept, a part only a lost one named absent, `_rels/.rels` that lost the main part's relationship refused; a save refused unless `repair` is `word` or `all` or `checkLossOnSave` is off) is implemented, unreleased. So is CR-033 (a DOCTYPE read by this package, never by the platform's parser, per the portfolio's DTD handling policy: refused where ECMA-376 Part 2 forbids one, the content types and `_rels/.rels` at every value, a part's relationships part and the core properties read as Word repairs them under `repair: 'word'`; elsewhere internal entities expanded and attribute defaults applied within limits, nothing fetched, `xml.doctype`; an edited save writing every such part without it). CR-029 (an `mc:AlternateContent` that selects nothing, WD REQ-305 and REQ-306: `Requires` by URI over the 114 Word 365 understands, `mce.nothing_selected`, `mcePreprocess: true` removing such an element) is implemented, released in 0.7.0, on objects 0.8.0. So is CR-026 (an element the model has no place for no longer failing its part, WD REQ-308 to REQ-316, docx4j CR-076; held to Word 365's check 76, eleven rounds). So are CR-028's D10 and D11 (a main document part whose root isn't `w:document` read as a document, refused, or read as an empty document with `part.unexpected_root` and its writes refused, as Word 365 reads 39 files, WD REQ-318 to REQ-320; the rest of CR-028 proposed) and CR-030 (a prefix in `mc:Ignorable` with no declaration read as naming nothing and said, `mce.undeclared_prefix`, WD REQ-323). So is CR-027 (`w:numId 0` clearing every inherited left, hanging and first-line indent, docx4j `9e3100351`). So is CR-025 (what a read drops reported as `part.skipped`, a part read and not changed copied, a save that would lose such content refused unless `pkg.checkLossOnSave` is off), and CR-002 section 66 (a `w14:paraId` on every paragraph and row the API adds). So is CR-012
section 24 (docx4j CR-069, released in 17.3.5: `w:caps` in capitals by the run's language through `CaseMapping`'s generated
Unicode 15.0 table, the block model's `caps`, `smallCaps`, `verticalAlignment` and `language`, a drop cap joined to its word,
`<sup>` and `<sub>` in Markdown, `strike` taking `w:dstrike`; the options `small_caps`, `drop_caps` and section 23.1's `profile`, its `word-text` writing list labels since WD 0.75, its `comments` the one setting since WD 0.91, section 29, `include: comments` a deprecated spelling; section 24.8's `default_language`, WD 0.80's language of last resort for the capitals, `w:themeFontLang` no longer read for them), and section 28 (a styles, numbering, settings or properties part, or a custom XML store, that cannot be read read as absent, with `part.not_read` where the reading needs it, through `PropertyResolver.create`'s `onUnreadable`; Word's default list, with `part.absent`, where the numbering part is absent or unreadable; the content API's refusal kept), and section 23.7 (REQ-702's date form, docx4j's `RevisionDates`: a revision's or comment's date as the author's local time with the offset `dateUtc` gives, its stored `Z` never written),
with the suite copied at the 17.3.5 release (its `dpp/` rows, Word's 22 text saves among them) and the text goldens recut from it (from 17.3.6 since CR-001 section 27). The editor's release requests are
CR-002 sections 22 to 30, reviewed in section 31. Each CR's last sections record decisions and
departures.

The dividing rule with the objects package: anything that needs only an object tree (helpers,
the `XmlUtils`-style facade, flat OPC typing) lives there; anything that needs parts or
relationships lives here. This package re-exports the objects facade so both read the same
(its flat OPC `Part` type is renamed `FlatOpcPart`; `Part` here is the part class).

## Commands

```
npm ci              # fflate, typescript, xpath (dev; an optional peer for consumers), @docx4j/generated-objects-ts and @docx4j/jsonix, as locked in package-lock.json

npm run build       # tsc -p tsconfig.build.json: src/ -> dist/ (git-ignored)
npm run typecheck   # tsc --strict over src/ and test/*.ts, then typecheck:examples (skipLibCheck: fflate 0.8.3's typings need TS 5.7)
npm run typecheck:examples   # builds, then examples/office-addin's own tsconfig: the taskpane imports the package by name, resolved
                    # through exports to dist/*.d.mts (the Office globals are a local .d.ts, not @types/office-js)
npm test            # pretest regenerates src/office-js/supported.generated.mts; then build, the nodenext consumer check (test/nodenext), node --test test/*.test.mjs
                    # (a shell glob: Node 22 does not accept a directory, and Node 18 would also run test/helpers.mjs)
node --test test/roundtrip.test.mjs   # one file, after npm run build
node examples/node/report.mjs test/fixtures/loadAndSave.docx   # the examples run against dist/

npm run generate:fonts   # src/model/fonts/*.generated.mts from ../docx4j's font tables and themes; committed output
node scripts/generate-wml-names.mjs ../docx4j/xsd/wml/wml.xsd   # src/opc/repair/wmlNames.generated.mts, the schema's element names (CR-026); --check compares
node scripts/generate-case-mapping.mjs <UnicodeData.txt> 15.0.0   # src/model/text/caseMapping.generated.mts (CR-012 section 24.1);
                    # the file from unicode.org/Public/15.0.0/ucd/, its SHA-256 in the output's header; --check compares
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
src/xml/dom.mts   parse/serialize over Jsonix.DOM (typed since @docx4j/jsonix 3.2.1), text and base64 helpers; parseXml reads a DOCTYPE through doctype.mts first (CR-033)
src/xml/doctype.mts readDoctype: the DOCTYPE out of the text before any parser sees it, internal entities expanded and defaults applied within DOCTYPE_LIMITS, nothing fetched (CR-033)
src/opc/          PartName, ContentTypes, ContentTypeManager, PartStore/PartSink + Memory/Zip/FlatOpc, Load, Save, mce/, exceptions, skipped.mts (`part.skipped`, CR-025);
                  repair/breaks.mts: a w:br outside a w:r put in a run before unmarshalling, as Word reads it (CR-021);
                  repair/unknownElements.mts: an unknown w: element in a content host read through, ahead of that pass, by the schema's names in wmlNames.generated.mts (CR-026);
                  repair/mainRoot.mts: a main document part's root other than w:document, read, refused or read as empty, ahead of both (CR-028 D10 and D11)
src/parts/        Part, BinaryPart (+ImagePart, ...), XmlPart<T>, DefaultXmlPart (+CustomXmlDataStoragePart), RelationshipsPart,
                  Parts, Namespaces, PartRegistry; wml/ dml/ pml/ sml/ docProps/ customXml/ typed parts
src/packages/     OpcPackage, WordprocessingMLPackage (createPackage, default styles and theme, body, outline, paragraphAt, getPropertyResolver/getNumberingEmulator/refresh), PresentationMLPackage, SpreadsheetMLPackage, registry
src/model/properties/ PropertyResolver, the property catalogue (one entry per schema member) and the catalogue-driven half of StyleUtil (apply, applyStyleLevel and the toggle XOR, isEmpty, unset, hasDirectFormatting);
                  tableStyleConditions.mts (namespace TableStyleConditions) and table.mts (TableContext, CellContext, TableContextTracker): a table style reaches a paragraph only through its CellContext (CR-007)
src/model/listnumbering/ definitions.mts (LevelDefinition, ListDefinition, NumberingDefinitions), state.mts (Counter, NumberingState, NumberingStates: one per story), formats.mts (the label formatters), Emulator.mts
src/model/fonts/  RunFontSelector (the document font per code point), ThemeFonts, Mapper/IdentityPlusMapper/PhysicalFont/FontRegistry, FontFallback, fontsInUse; *.generated.mts from npm run generate:fonts
src/model/dateZone.mts  the package's time zone (CR-018): wallClockIn and instantOf over Intl, for a revision's or comment's w:date (Word's local wall clock labelled Z)
                  written and read in pkg.dateTimeZone, else the runtime's zone; assertTimeZone refuses a zone Intl doesn't know
src/model/wordIntegers.mts  the range Word 365 keeps as written in an integer attribute, from the objects table (keptRange; rangeOf from the entry a report names,
                  since objects 0.7.0): one source for the setters' guards and the value.out_of_range warnings (CR-017, CR-002 section 65.3)
src/model/color/  docx4j's org.docx4j.model.color (CR-008 phase 2): ResolvedColor, ThemeColors (ofTheme sync; `await ThemeColors.of(pkg or part)`,
                  forChart; a themeless package the bundled theme pkg.fonts.defaultTheme names), WordColors (of, ofBorder, ofUnderline, fill, pattern;
                  normalise and stated, docx4j CR-060's, over the objects package's wordColor: a value read as Word reads it, CR-015),
                  DrawingMLColors (transforms in document order, style references), SpreadsheetColors, StyleFill; ColorMath is docx4j's arithmetic
                  in its order (the luminance rounding depends on it); PropertyResolver.getThemeColors() and resolveColor/resolveFill/... use them
src/model/content/ the content API in Office JS shapes: Body, Paragraph, Range (+hyperlink; two ends since CR-014 phase 1, story.mts its story's layout, spanEdits.mts its writes across paragraphs), Font, Table (+TableRow, TableCell; columns added and removed in Word's three layout modes, tableColumns.mts, CR-002 section 42), InlinePicture, ContentControl, Comment, TrackedChange, List (+ListItem; labels from the Emulator, one story walk per read or Body.listLabels() for all; NumberingFacade is pkg.numbering, the definition verbs without a paragraph), search;
                  ooxml.mts is insertOoxml/insertXml (flat OPC in, referenced parts copied; a package's styles, lists and notes reconciled as Word 365's insertOoxml does, through fragments.mts's insertedContentOf, CR-005 section 11.6); comments.mts the comment plumbing (parts side in parts/wml/comments.mts);
                  tree.mts is the paragraph text model (segmentsOf in Office JS's characters, childrenOf; runItemsOf is re-exported from builders/wml); fragments, run mapping and traversal come from the objects package's builders/wml;
                  officeText.mts writes Office JS's characters back as the elements they read as (CR-002 section 60: officeItems, dropExcluded, forWriter);
                  tableRules.mts is docx4j's TableRules (endsWithParagraph, endWithParagraph: a cell ends with a paragraph as Word reads it), which
                  Body.insertElement, clear, Paragraph.delete, ContentControl.delete and the XHTML pass keep, as Word keeps a cell's last paragraph (CR-002 section 55);
                  textBoxes.mts numbers a paragraph's text boxes as the text model does (`textbox:N`); outline.mts is the registry the text model's outline reader joins (model/text/outline.mts; Body can't import the text model);
                  placeholders.mts is a placeholder's form (`w:placeholder`) and the glossary's entries, `pkg.glossaryEntries`, read privately and written on save
                  (CR-002 section 63, D8 and D9); a text write into a control showing its placeholder, by any route, goes through `ContentControl.writePlaceholder`;
                  wordRange.mts holds every measurement setter to the range Word 365 stores its attribute in (`withinWord`, `WORD_STORAGE`:
                  16 bits, signed or unsigned per attribute, each row the objects package's `WORD_INTEGER_STORAGE` entry through model/wordIntegers.mts's
                  `keptRange`, narrowed only by the setter's stated policy, CR-002 section 65.3: Word's own limits for spacing, line, list indent and
                  font size, check 71b, section 65.9; past it `InvalidArgument`, CR-002 section 65);
                  binder.mts records which object each element of a marshalled snapshot came from (CR-006: XmlPart.selectObjects, pkg.selectObjects, Body.select); stylesFacade.mts is pkg.styles
src/model/customxml/ CustomXmlPart/CustomXmlNode over the custom XML DOM parts, XPathEngine (native document.evaluate, else the optional xpath package; await pkg.customXmlParts.load() once),
                  FontoXPathEngine on ./xpath-fonto (optional peer fontoxpath) for the OpenDoPE xpath2 boolean mode, with booleanValue() over any engine (CR-005 phase A),
                  XmlMapping over w:dataBinding, the typed content-control kinds, insertContentControl, applyBindings/updateFromContentControls (docx4j BindingHandler);
                  storeItems.mts: the part a binding is evaluated against, docx4j's StoreItemResolver (CR-020): the part its w:storeItemID names, else the one
                  its XPath fits, as Word repairs it; every reader (the text model's boundValue, XmlMapping, the od:xpath passes) and the bind go through it;
                  opendope.mts (the tag's parameters, the XPaths part's entries), xhtml.mts and pictures.mts: the tag-bound passes, od:ContentType=application/xhtml+xml and od:Handler=picture (CR-005 sections 8 and 9);
                  fragments.mts: Word fragments, od:progid=Word.Document and Word's w15-mapped rich text over Flat OPC, last of the passes (CR-005 section 11), over content/fragments.mts's importFlatOpc (Word's reconciliation: styles by id, each list a new definition, notes copied, ids fresh where they collide, keepSourceFormatting); picturePlaceholder.mts: Word's form for a picture control whose node is empty, the flag and docx4j's placeholder picture 1270000 EMU square (CR-005 section 10)
src/model/content/tracking.mts, TrackedChange.mts: change tracking (pkg.changeTrackingMode; revision markup written by the paragraph primitives so every caller inherits it);
                  pkg.getAllTrackedChanges(filter?) lists every story in docx4j's order, TrackedChange.kind/part/pieces are docx4j's, and trackedChangeFilter.mts
                  is the filter (docx4j's TrackedChangeFilter, or a function) the listing, acceptAllRevisions/rejectAllRevisions and Body.acceptAll/rejectAll take (CR-011);
                  tableRevisions.mts: a table's property revisions, listed as one change and resolved as Word does (CR-002 section 35)
src/model/text/   the text model (CR-012, docx4j's org.docx4j.model.text): Texts.of (a package, a story part, content in a package; ofResolved, the copy route),
                  TextContext (note numbers, numbering states, plans, made once by Texts.contextOf over PackageTrees, every story read with readContents),
                  TextOptions keyed by the specification's option names, the block model (TextResult, TextBlock, TextRun, ...: docx4j's enum names as strings),
                  Extractor (the walk: selectedBranchOf's branch, the view through content/ResolvedView.mts, scope, characters; a drop cap's frame
                  joined to the paragraph next, section 24.3), PlainText; each TextRun's officeOffset is where its text lies in Paragraph.text (REQ-202's
                  emulated-API exemption; a joined drop cap's merged run has none). The Office JS readers keep Office JS's text (D2).
                  CaseMapping (section 24): w:caps written in capitals by the run's language, Unicode 15.0's simple mapping from caseMapping.generated.mts,
                  never the runtime's toUpperCase; the block model keeps the stored text, the writers (PlainText, WmlToMarkdown) map it
                  MarkdownExporter (export, blocks: a block's Markdown each; a link whose scheme LinkTargets.isSafe refuses, its character references decoded too, written as its text, CR-012 sections 21 and 21.2) over WmlToMarkdown, docx4j-markdown's, which builds a Markdown tree that
                  commonmark.mts, a port of commonmark-java's Markdown renderer (its BSD notice kept), writes as docx4j writes it (CR-012 phase 3)
src/office-js/    the Word shim (Word.run(pkg, fn), context.sync, proxies throwing NotSupportedError, enums, Word.supported) and toApiScript; exported only from ./office-js.
                  supported.generated.mts is written by scripts/generate-supported.mjs from test/office-js-subset.ts (npm run generate; pretest runs it): commit it with every subset change
src/index.mts     re-exports all of the above plus the objects facade
```

Key mechanics:

- **Untouched parts round-trip byte for byte.** A package keeps its `sourcePartStore`; on save,
  only relationships parts and `XmlPart`s that were unmarshalled (or had contents/bytes set) are
  re-marshalled, everything else is copied from the source. Do not call `getContents()` on a
  part unless the caller asked for its contents. `[Content_Types].xml` is always regenerated.
- **Run properties are read and written as Word 365 reads them** (CR-015). The objects facade reads an attribute value
  its type refuses by Word's rule rather than failing the part (on/off on unless exactly `false`, `off` or `0`; an
  enumeration value outside its list dropped; `w:u`'s on/off words `none`), and each such value is a `ValueKeptWarning`
  on the part's `warnings`. A marshal writes run properties in their normal form (objects CR-013 phase C, docx4j CR-061's
  rule: colours upper-case `RRGGBB`, a rejected text colour left out, what Word repairs left out), unless
  `pkg.normaliseOnSave` (else the objects package's `setNormaliseOnSave`) is false; `pkg.lastSave.normalised` lists what
  the last save rewrote. Only an unmarshalled part is marshalled, so an untouched part is still copied byte for byte.
- **An integer attribute value past what Word 365 stores is reported, never rewritten** (CR-017; the objects package's
  `onOutOfRange`, its CR-016, over docx4j CR-066's rounds 5 and 6). Word keeps each such attribute in a field of its binary
  format's width, 16 bits for most, and reads a value past it wrapped, sometimes then clamped, with no error: `w:ind
  w:left="7205001"` opens as -3959. On load each one is a `ValueOutOfRangeWarning` on the part (`value.out_of_range`, the
  code and fields shared with docx4j-python: `element`, `attribute`, `value`, `wordReads`, `storage`, `range`, and `readAs`
  or `rule` where the shared field's sign or a line rule moved it) and on save on `pkg.lastSave.outOfRange`, with its path;
  the tree, the readers and the save keep the value as written (E2), and a decimal reads truncated as Word reads it
  (`value.truncated`) or, in a form Word refuses a document for, dropped (`value.refused`; P1, the objects package's CR-015).
  The setters refuse such a value before it is written (CR-002 section 65), from the same table (`keptRange`).
- **A child its container never admits is refused at save** (CR-016; docx4j CR-064's rules in the objects facade, its CR-017,
  `onMisplaced` on read and on save: a run straight in the body, a table or a paragraph in a paragraph, a cell beside a
  table's rows, a string in a run, a name and value that disagree; never an allow-list). Only a part being marshalled is
  checked, so a document read and saved untouched is never refused. The save throws `MisplacedContentException` with every
  such child in the part (`findings`: `target.misplaced`, docx4j-python's code and fields, `path`, `holder`, `item`,
  `element`, `index`, `rule`, docx4j's message naming the fix) and writes nothing, unless `pkg.checkContentOnSave` is false,
  when the part is written as held and they go to `pkg.lastSave.misplaced`. `part.checkContent()` and `pkg.checkContent()`
  are the same walk on request, changing nothing (an untouched part read privately). On load each one is the same warning
  on the part (`line` and `column` in place of `path`; a string in a run, which the runtime drops, has `text`, and is the only
  notice of it). The builders (`p`, `r`, `tbl`, ...) refuse first. The facade's walk keys on the container's `TYPE_NAME`:
  `createPackage`'s document, body and `sectPr` carry theirs; since objects 0.8.0 a container built by hand without one is
  walked too. The gate: 806 corpus files save under the check unrefused.
- **What a read drops is reported, and a save doesn't lose it unasked** (CR-025, the repair policy's P7). `XmlPart.unmarshal` passes the runtime's three `onUnexpected*` callbacks, and each item is a `ContentSkippedWarning` (`part.skipped`, `src/opc/skipped.mts`) on the part and on `pkg.warnings`: `kind`, `name`, `parentName`, `parentClass`, `path`, `ignorable`, `line`, `column`. The callbacks stay at that one call, since a drop is sometimes the mechanism (`recordPPrChange`'s deep copy), and CR-021's pass reports what it removes with all it holds. A string CR-016 reports as misplaced isn't reported twice. On save, before anything is written, `Saver.settleLosses` walks the parts the save reaches: one holding skipped content that `isUnchangedSinceRead` finds unchanged (its source read again and the two marshalled alike) is copied byte for byte, with no normalising or checks (its source read again with the read's callbacks, section 9). One that changed is written without what Word's own save drops too, each item on `lastSave.dropped`, never refused; for the rest it throws `SaveLossException` (`save.content_lost`, `losses` naming every part and counting those), unless `pkg.checkLossOnSave` is false, when it's written without them and each goes on `lastSave.lost` (section 10). Each item's class is on it, `onSave` (`dropped` or `lost`), `measured` (the Word file or statement that decides it) and `namespace`: by its shape against check 76's files in `skipped.mts`, the read-through pass and CR-021's and CR-028's giving their own items' class; a shape nobody measured is lost (D15). `setContents` clears the part's warnings, so it's never refused, and is how a caller mends a file. A relationships part in another namespace, or in none, on its root or on any `Relationship`, is refused at the read (`RelationshipsNamespaceException`, `package.relationships_namespace`, CR-025 section 8), since the runtime would drop each such `Relationship` and the part it names; under `repair: 'word'` it's read by local name instead, as Word's repair reads it, a `package.relationships_namespace` warning on the part (`pkg.relationshipsNamespaces`) and a save in the right namespace (`RelationshipsPart.readsByLocalName`, CR-031 section 2.3); what else a relationships part's read drops is on `pkg.warnings`, outside the gate. `clone()` goes through a save, so it's refused too. A single-valued child given twice keeps the last with no callback, so neither the report nor the gate sees it (section 2.6).
- **An element the model has no place for doesn't fail its part** (CR-026, docx4j CR-076, WD REQ-308 to REQ-316, check 76). Where it used to fail
  the part, in 20 of the 32 hosts tried, the runtime skips it and reports it, `part.skipped` through `onUnexpectedElement` (route B, jsonix-CR-015). An unknown `w:`
  element in one of the twenty hosts measured is read through first, as Word reads it: `readThroughUnknownElements`, a DOM pass in `XmlPart.preprocess` ahead of
  CR-021's, takes its tags and keeps what it holds, a run in one in a `w:pPr` going to the head of the paragraph and a `w:p` in one, last in a `w:p` with a run before it,
  ending that paragraph. `insertOoxml` refuses a fragment holding such an element, on both routes (D12). 78 of check 76's 79 saved sources read with Word's text;
  `76t71`, in a `w:hyperlink`, is the known difference.
- **A main document part's root is read as Word 365 reads it** (CR-028 D10 and D11, WD REQ-318 to REQ-320, docx4j CR-079). `readMainPartRoot`,
  run by `MainDocumentPart` through `XmlPart.readRoot` ahead of every other pass: a root Word reads (`w:body`, `w:p`, `w:sdtContent`,
  a block `w:sdt`, an unknown `w:` element holding block content) becomes a `w:document` holding a `w:body`, the root's declarations
  and `mc:` attributes on it, a `w:sectPr` standing in an unknown root going as `part.skipped`; a root Word refuses (`w:footnotes`,
  `w:endnotes`, `w:tc`, `w:txbxContent`, an unknown root holding a `w:document` or a run, ...) throws `PartUnmarshalException`
  at the read; any other (`w:hdr`, `w:styles`, a root in a namespace its `mc:Ignorable` names, ...) reads as an empty document with
  `part.unexpected_root` on the part and in the text model, the root a `part.skipped` so a save copies the part as written,
  `MainDocumentPart.unexpectedRoot` naming it and `body` refusing it. A `w:document` passes untouched. An `mc:AlternateContent` root (REQ-317,
  CR-028's rest) is left whole, by a resolving load too, and `body` refuses it, since its branch isn't read as the document and a write would go nowhere.
- **A relationships part that isn't well formed is read short, as Word's repair reads it** (CR-032, docx4j CR-075 D5 and D6), at
  every value of `repair`: `XmlPart.parse` hands a part that won't parse to `readMalformed`, which only `RelationshipsPart`
  overrides, through `src/opc/repair/relationshipsBeforeTheBreak.mts` (the longest cut after a child of the root that parses, the
  parser deciding, never the scan). Each whole `Relationship` before the break is kept, in the right namespace; `part.not_read` on
  the part (`pkg.relationshipsNotRead`, the text model's warning, the count its detail); the lost parts' Overrides go. `_rels/.rels`
  without the main part's relationship among those kept is refused at the load. `Saver.settleLosses` refuses the save
  (`SaveLossException`, a `relationships` loss) unless the load's `repair` is `'word'` or `'all'` or `checkLossOnSave` is off;
  `lastSave.relationshipsNotRead` lists what it wrote so. `clone()` is refused with it. What breaks a part is xmldom's: it reads
  through an undefined entity, an unquoted value or a lone `&`, where docx4j's StAX stops.
- **A DOCTYPE is read by this package, never by the platform's parser** (CR-033, the portfolio's `docs/dtd_handling_policy.md`).
  `parseXml` hands every text to `readDoctype` first: the DOCTYPE comes out, its internal entities are expanded (content and
  attribute values, not comments, PIs or CDATA) and its attribute defaults applied after the parse, within `DOCTYPE_LIMITS`
  (nesting 8, 10,000 references, 1 MiB of replacement text a part; past one, or a recursive entity, `Malformed XML`), and
  nothing outside the text is ever read (an external subset named and left, an external entity read as nothing, declarations
  after a parameter entity reference not processed). So Node and a browser read alike. Where ECMA-376 Part 2 §6.2.5 forbids one
  (`[Content_Types].xml` and `/_rels/.rels` at every value; a part's relationships part, the core properties and a signature
  part unless `repair` is `'word'` or `'all'`) the load throws `DoctypeException` (`package.doctype`); under `'word'` the
  relationships part reads as none kept (CR-032's `part.not_read`, count 0), the core properties as absent, a signature as it is,
  and `pkg.repaired` is `'word'`. Elsewhere `xml.doctype` (`DoctypeWarning`, with Word's reading) is on the part. A save after an
  edit (`Saver.dropDoctypes`) writes every part whose source holds one afresh without it, a signature part excepted;
  `lastSave.doctypes` lists them. The text model's `doctype: 'word'` waits for the WD.
- **A `w:br` outside a `w:r` is read as Word 365 reads it** (CR-021, docx4j CR-073's rules from its `mc-preprocessor.xslt`).
  `repairBreaksOutsideRuns` runs in `XmlPart.preprocess` on every parse, ahead of the package's hook (and before it on
  `insertOoxml`'s fragment route): a break in a run container gets a run of its own, with the mark's run properties while no
  run precedes it in its paragraph; a `w:t`, `w:instrText`, `w:delText` or `w:delInstrText` holding elements (which the runtime
  refuses, failing the part) becomes its text and its breaks, a nested text element giving its text and any other `w:` element or one in a namespace Word doesn't know going with all it holds (check 76, CR-021 section 11, where docx4j gives its text); a break in `w:body` goes into the next paragraph. Each is decided
  on the tree as read. One `br.outside_run` warning a part, docx4j's counts; such a break is then no misplaced child. A clean
  part pays one walk (0.4% of a load); an untouched part is still copied byte for byte.
- **Dates are in the package's zone** (CR-018, the portfolio's date rule, docx4j CR-067 D12). A revision's or a comment's `w:date`
  is Word's local wall-clock time labelled `Z`, the true UTC beside it in `dateUtc` (CR-002 section 29); "local" is
  `pkg.dateTimeZone` (an IANA id, the client's for this package, never saved; copied by `clone()`), else the runtime's zone,
  for every stamp written and every `w:date` without `dateUtc` read (`TrackedChange.date`, `Comment.creationDate`). Stamps
  are to the minute. A bound date is nominal (section 56) unless `pkg.boundDatesConvertedAsWord` asks for Word's conversion
  in that zone. A save leaves `dcterms` alone unless `pkg.writeTimesOnSave` (docx4j's `docx4j.dcterms.write`). The shim
  reads a date-only or offset-free `dcterms` value as UTC, as Word does (docx4j CR-067 section 12.2's probe).
- **Container is separate from package.** `PartStore` (load) and `PartSink<R>` (save) abstract
  zip (`fflate`; the central directory is parsed here, entries inflate lazily), flat OPC
  (`pkg:package`; content types are per part, so the store synthesises `[Content_Types].xml`),
  and memory. `OpcPackage.load` sniffs: string is flat OPC, bytes are a zip, else a store.
  Each store keys its entries as OPC part names, without case (`EntryIndex`): a container holding
  two entries of one name is refused at the load (`DuplicateEntryException`, WD REQ-206, CR-024),
  unless `LoadOptions.repair` is `'word'` or `'all'` (CR-031, WD REQ-207), which reads the first of each, as Word does
  after its repair, with `package.duplicate_entry` on `pkg.warnings` and in the text model's. The same
  switch decides more than one officeDocument relationship (CR-024 section 9): refused
  (`MultipleMainPartsException`), or under `'word'` the **last** relationship's target read, as a fresh
  Word reads it, the others dropped at the load so every route finds one main part and a save writes one,
  with `package.multiple_main_parts` (`pkg.multipleMainParts`). An `External` officeDocument relationship
  beside an internal one is refused at `'refuse'` and `'word'` (`ExternalMainPartException`, REQ-206's NOTE),
  checked first; under `'all'`, this package's own repair, every such relationship goes at the load, the URL
  never fetched, with `package.external_main_part` (`pkg.externalMainPart`). One with no internal one beside it
  is refused at every value, as no main part (REQ-321). `pkg.repaired` (and the text model's `repaired`
  metadata) says `own`, `word` or `none` from what the load recorded. `repair` never touches `checkLossOnSave` (D6).
- **`XmlPart<T>.getContents()` is async** because the facade builds its Jsonix context lazily;
  afterwards `contents` is synchronous. `OpcPackage.unmarshalAll()` does them all.
- **Load follows relationships** from `/_rels/.rels`; unreachable parts are not loaded. The
  registry maps content type (relationship type first, for altChunk/embedded/OLE/custom XML)
  to a part class; unknown XML becomes `DefaultXmlPart`, anything else `BinaryPart`. Custom XML
  parts are indexed by the `ds:itemID` of their properties part (read from the DOM, so that
  part stays untouched).
- **`mc:AlternateContent` is kept on load** with every branch, as docx4j keeps it, since CR-010
  phase 4 (`mcePreprocess` defaults to false). Readers take `selectedBranchOf`'s branch (the objects package's
  `mcBranchOf`: the first `mc:Choice` whose `Requires` prefixes all resolve, by URI, to `UNDERSTOOD_NAMESPACES`, the 114
  Word 365 was measured to understand, WD REQ-306; else `mc:Fallback`; else nothing, REQ-305, CR-029), mutators every
  branch, a run's element is one object segment, and a save writes every branch. The default pass (`mceKeep`) changes nothing and reports
  `mce.branches_differ` on `XmlPart.warnings` where the branches' text differs, and `mce.nothing_selected`
  where an element in a word-processing part selects nothing, and `mce.undeclared_prefix` once a part where an `mc:Ignorable` lists a prefix no
  declaration in scope binds, read as naming nothing (WD REQ-323, CR-030; a list in a branch read where it stood before a resolving load moves it) (`OpcPackage.warnings` gathers them; the text model
  raises its own at the element's address, and carries a record with no address for what its walk can't meet, an
  element a resolving load took out or one at a root). `mcePreprocess: true` resolves on the DOM before unmarshalling
  instead, as Word does on open, by the same set: it removes an element that selects nothing, a `w:tc` left with no block
  getting the paragraph Word's repair gives one, single-spaced with no space after (CR-029 D9), which Word opens silently, but keeps the element under `KEPT_HOSTS` (`c:chartSpace`, the
  SpreadsheetML workbook) and in a part that isn't word-processing (CR-029 D3), and leaves any `mc:AlternateContent` that is a part's root whole (CR-028 D10). Since objects 0.4.0 a kept branch's `w:drawing` and
  `w:pict` are typed (objects CR-009, CR-010 phase 2); the `xdr:` and `p:` drawing content, a
  worksheet's `x:controls` and `c:style` inside one stay DOM, so `worksheet.controls` is typed only
  when resolved. A code path that walks the tree itself has to choose: `selectedBranchOf` to read,
  every branch to write. At block, row and cell level the walks take a `BranchMode` (CR-013 phase 2):
  `childrenOf(v, 'read')`, `rowsOf(tbl, 'read')`, `cellsOf(tr, 'read')` for readers, `'every'` for mutators
  (`branchListsOf`; `cellLayoutsOf` gives a row's cells along each reading, for the grid); no mode steps over
  the element. A view in a branch (`inBranch`) refuses its writes (`assertOutsideBranch`, D4); a bulk edit
  writes there inside `writingEveryBranch`, and counts what a reader sees (`inUnreadBranch`). The tracked-change
  listing pairs each other branch's changes with the selected branch's as twins, resolved with them (`Listing`,
  `pairTwins`; resolve-all lists every branch's on their own, CR-013 phase 3). Kept, a Word text box's VML Fallback is unmarshalled, so jsonix decodes its
  `o:gfxdata`, which needs `@docx4j/jsonix` 3.4.2 under strict mode; objects 0.4.1's range brings
  it (CR-010 section 14.4). The floor is objects 0.8.0, over `@docx4j/jsonix` 3.12.0, whose marshal declares the
  prefixes a kept `mc:Choice`'s `Requires` names (0.4.2, CR-010 section 14.8), from the source where
  only the source bound them (its CR-011, section 14.9), on the `mc:Choice` itself, as Word does (0.5.0, its CR-012);
  which reads and writes run properties as Word 365 does (0.5.0, its CR-013; CR-015); and which reads a highlight in
  Word's palette (0.6.0, its CR-014) and an integer attribute's decimal as Word does, `value.truncated` or
  `value.refused` on a part's `ValueKeptWarning`s (0.6.0, its CR-015; CR-017 P1), and reports a value past what Word stores
  (its CR-016, `onOutOfRange`, each report naming its table entry since 0.7.0; CR-017 P2 and P3); and which reports a child its
  container never admits on read and on save (0.7.0, its CR-017, `onMisplaced`; CR-016 here), types a kept branch's `w:sdt` at its
  host's level (0.7.0, its CR-018), and keeps `w:legacy` out of its integer table (0.7.0, its CR-016 addendum: 49 entries).
  Since 0.8.0 it selects a branch by namespace over Word's set (its CR-021; CR-029 here), keeps `w:stylePaneFormatFilter`'s attributes on a save (its CR-019),
  refuses on a marshal a character XML 1.0 excludes (its CR-020), and walks a container built by hand without its `TYPE_NAME`; and `@docx4j/jsonix` 3.12.0 skips
  and reports an unknown element where it used to fail the part (its CR-015; CR-026 here).
  WD REQ-306's two sets are on `./opc`: `UNDERSTOOD_NAMESPACES` is the objects package's, docx4j's published 114, and decides
  every markup compatibility question; `TRAVERSAL_NAMESPACES`, docx4j's 118, is what a reader here can traverse and decides
  nothing about selection. Its first 107 are a hand-kept copy of the generated modules' namespaces, plus the slicer namespaces
  nothing binds (CR-004 section 6) and the chartex `cx1` and `cx2` (this package's CR-009, section 8). Taking a Choice
  (`mcePreprocess: true`) **converts DOM into typed content**, so it can expose a binding gap the unresolved form hid: content
  the model binds only partly loses the unbound part on a re-marshal, where leaving the Choice alone would have kept
  all of it (measured, CR-004 section 7). The set follows what Word understands, so such a gap goes to the objects package (CR-029 D6).
- **Namespace prefixes** are the facade's job (objects CR-001: `NAMESPACE_PREFIXES` default,
  `mc:Ignorable` declarations kept, a kept `mc:Choice`'s `Requires` prefixes declared since 0.4.2,
  and since 0.4.3 a prefix only the source bound declared from the source,
  `xml` never declared). Nothing here touches prefixes.
- **The content API is views, not a model.** `Body`/`Paragraph`/`Range` hold a reference to
  the element and the array containing it; nothing is cached but the `Body` per container
  (a WeakMap in `parts/wml`). Text edits work on the paragraph's text segments (`segmentsOf`)
  and edit `w:t` values in place, split runs only when formatting a span, and link `PARENT`
  on everything inserted (`linkParents`). `Paragraph.text` is those segments' text, Office JS's
  (CR-012 phase 4): check 48's characters (`officeItemText`: a note reference U+0002, a break
  U+000B, a symbol `(`, an equation as Word draws it, `mathText.mts`), hidden runs left out
  (`hiddenRunsOf`, zero-width segments listed with the objects), and so a control's placeholder (`inShowingPlaceholder`, CR-012
  section 18), a mark tracked as deleted written as nothing (`storyLayout`, `markText`, CR-012 section 19), a kept `mc:AlternateContent`'s
  selected branch read and edited in place (`alternate` on a segment; a run split inside a run-level
  branch writes the editor's ED-009 P2 pieces), so text, offsets, `Range`, `search` (`officeMatches`,
  Word's Find codes) and every edit share one reader. A tracked change's `text` reads
  `characters: 'textOf'`, the old reader, which the revisions goldens hold to docx4j's. A paragraph in
  a text box (`textbox:` addresses, `textBoxes.mts`) is read-only: mutators call `assertWritable()`. A `Range` runs
  from `start` in `paragraph` to `end` in `endParagraph`, the same paragraph unless it spans several (CR-014 phase 1):
  across paragraphs it reads a slice of `story.mts`'s layout (`body.text`'s form, which `Body.text` is too), lists
  the story's tracked changes that overlap it, and reads `font` `null` where runs differ (`MixedFont`). Its writes
  (phase 2, `spanEdits.mts`) are Word's from check 52: `delete` and `Replace` join the paragraphs in the last one's
  properties (`joinWithNext`, Accept All's), a range holding its last mark (a 'Whole' range, even within one paragraph)
  takes it and joins the next paragraph, tracked every mark in it but a last one is deleted with no `w:pPrChange`;
  `font` writes runs and the marks it holds, a toggle only where it doesn't read so already; `insertParagraph` splits
  the paragraph at the range's start or end; a range starting or ending inside a table is refused with Word's
  `GeneralException`, one holding a whole control or `w:altChunk` refused under tracking. Phase 3: `insertComment`
  puts its markers in the first and last paragraphs, and a plain `search` query holding `^p` matches paragraph marks
  across paragraphs, never a cell's or row's end or across a table's edge (`searchMarks`; `replaceText` with `^p`
  refuses), and so do wildcards (`?`, `*`, a class, `^13`; `searchWildcards` over `wildcardMatchesAcross`, check 59,
  CR-002 section 32 K7; `replaceText` refuses where a hit holds a mark). Check 54's answers (section 12): `insertParagraph` splits within one paragraph too; a mark the range
  holds goes (`dropMark`: joins the next paragraph, else its emptied paragraph goes, a cell's last kept); delete
  crosses a control's edge one level down without joining; a font property goes only where it doesn't read so already;
  a comment goes with its reference (every deletion: `Paragraph.splice` and `deleteText` take a reference run in the
  span, and the resolution removes a comment left without one, docx4j's `Removals` rule; section 12.6); `style` and
  `insertOoxml` work across paragraphs; `insertContentControl` wraps
  paragraphs held whole, else `GeneralException`; a tracked delete holding a whole control writes Word's
  `w:customXmlDelRange*` markers (`markHolderDeleted`), which the `holder` tracked change lists and resolves beside the
  Ins form (`holderChangeAt`). `hyperlink` across paragraphs still refuses. `expandTo`, `Body.getRange`, `Table.getRange`, a block
  control's and a group change's range and `Comment.getRange()` (one `Range`; `getRanges()` the per-paragraph array)
  make one. Code that takes a `Range` checks `spansParagraphs`. **The writers write Office JS's characters back** (CR-002
  section 60, `officeText.mts`): `splice` (every `insertText`, `replaceText`) inserts its text, then writes the span's
  characters that are elements as them (`writeElements`: U+000B a `w:br`, U+000C and U+000E page and column breaks, `\r` a
  `w:cr`, U+001E and U+001F the hyphens, `\t` a `w:tab`); a `\n` or `\r` is a paragraph break (E1, `breakParagraphAt`, the
  returned `Range` across them), refused before anything changes in an equation, a kept branch or a run-level control
  (F4), and inside a field's result splitting it, a `w:fldSimple` made a complex field first, as Word does (section 62.3,
  `unfoldSimpleFieldsAround`); `insertParagraph` and `Paragraph.text =` keep one paragraph (E2: `\n` a `w:br`, `\r` a `w:cr`); a
  plain-text control takes a break as its kind does (F3); the characters XML 1.0 excludes are dropped (E3), and
  `XmlPart.getBytes` refuses a part still holding one (D3). **A write at a paragraph's End** (`insertText('End')`, its End range at any location, its Content range's End and After, so a body's or a cell's End) goes through `Paragraph.writeAtEnd`, after everything the paragraph holds and into none of it (a run-level control, a link, a revision, a field, markers, a text box), in the last item's run properties only where that item is a run holding text, the End range's After in none (CR-002 sections 63.24 to 63.27, checks 69 and 70); a range that merely ends there writes as `splice` does. A block-level control bound by `w:dataBinding` (untyped,
  `w:text`, `w:richText`; not a flat OPC node) holds one paragraph: every write that would give it another block is
  refused (CR-002 section 61, `assertControlHolds`), it reads as `PlainText`, and `getContentControls({ types })` finds a
  kind as Word's does. `Font` and `Paragraph` reads report **effective**
  formatting through the `PropertyResolver` (and, for `Font.name`, the `RunFontSelector`); `Font` reads every run in scope
  (a paragraph's: the runs of its text), `null`, `''` or `Mixed` where they differ, a paragraph with no text its mark, an
  automatic colour `#000000`, as Office JS (CR-002 section 59), a theme colour the colour Word draws (CR-008 phase 2; the
  direct view keeps `w:val`);
  `{ direct: true }` on `getFont` / `formatting` reads the direct values, and writes are always
  direct. In a table cell the effective reads pass the paragraph's `cellContext`, so they include
  the table style and its conditional formats (CR-007). `test/office-js-subset.ts` must stay assignable: it is the Office JS promise.
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

`test/*.test.mjs` on Node's runner, against `dist/` (`content.test.mjs` covers the content API,
`create.test.mjs` the created docx, pptx and xlsx, `directory.test.mjs` the `./node` container and
`clone()`, `omml.test.mjs` an equation in DrawingML text (docx4j `OmmlInDrawingMLTextTest`),
`sml-extensions.test.mjs` the Excel 2010/2013 extension parts (CR-004),
`kept-branches.test.mjs` the content API over kept `mc:AlternateContent`, resolved and kept loads compared (CR-010 phase 3),
`text-extraction.test.mjs` the text extraction suite and agreement test (CR-012 phase 0: `test/text-extraction/`, docx4j's
suite copied unedited by `scripts/copy-text-extraction-suite.mjs` at the commit its `manifest.json` names, never edited here,
the 17.3.5 release since section 24; the engine since phase 2 is `test/text-extraction/adapter.mjs`, and a claimed row that fails,
or a count that moves, fails the test; `phase0/` and docx4j CR-069's `dpp/`, whose Word text saves `verifyWordSave` reads by encoding),
`content-skipped.test.mjs` CR-025's: each host and attribute's `part.skipped` held to the source, a part read and not changed copied, over every fixture with skipped content injected, the gate, `checkLossOnSave`, `clone()`, CR-021's pass and the real file,
`fragments.test.mjs` CR-005 section 11's: the editor's proposal 46 check (`test/fragments/`, Word 365's saves of 46a, 46b and 46c in compatibility modes 12 and 15, copied unedited from the editor's `c71ebd6` with a SHA-1 manifest): the `od:progid` control held to Word's resolution of docx4j's chunk paragraph for paragraph, the `w15` control to Word's fill on open, `keepSourceFormatting` to Word's `w:matchSrc`, the reconciliation (styles, lists by their labels, the section), the altChunk route to docx4j's form, `w15Mapping` and `w15FlatOpc`, placement, what is dropped, ids, a header, and what is left with a note; the text model's reading of a plain text control mapped to a Flat OPC node pinned as a known difference,
`relationships-namespace.test.mjs` CR-025 section 8's: docx4j's `rn-*` round (`test/rels-namespace/`, with a SHA-1 manifest), each form refused at the load naming its part, Word's saves read as their controls, flat OPC, and a relationships part's other dropped content on `pkg.warnings`; and under `repair: 'word'` (CR-031 section 2.3) each form read as its control, said once, saved in the right namespace with every part kept,
`rels-not-read.test.mjs` CR-032's: docx4j's broken-part files 100, 119 and 120 at the 17.3.6 release with Word's saves and texts and the family's two tables (`test/rels-not-read/`, copied unedited with a SHA-1 manifest), each file at each value of `repair` read short as docx4j reads it, the tables' tokens in the three views by the suite's `verify`, `with_option` held to Word's text, the save refused at the default and written under `word`, `all` and `checkLossOnSave` off, `_rels/.rels` without its main part refused, and the shapes (a cut in the root's start tag, a fault mid-part, a `Relationship` without a `Target`, a comment and a `>` in a value, a wrong namespace before a break),
`doctype.test.mjs` CR-033's: Word's two DTD rounds (`test/doctype/`, docx4j-python's dtd-word-check and dtd2-word-check through the portfolio's bridge, copied unedited with a SHA-1 manifest): the forbidden parts refused or read as Word repairs them, dtd-18's text and save Word's, every other case read with `xml.doctype` and Word's reading, the subset as XML 1.0 has it (an entity, a default, an external subset), the limits and nothing fetched, and the edited save without each DOCTYPE,
`repair-option.test.mjs` CR-031's: the share's `x-ext-*` files (`test/repair/`, copied unedited with a SHA-1 manifest that records Word 365's refusal of each open), each refused at `refuse` and `word` and read at `all` with `package.external_main_part`, `repaired: own` and a save holding no `External` relationship, held to D7's Word round of those saves (`ctra-*`: Word opened all eight silently, its text and its save held here); `repaired` at every value, `checkLossOnSave` left alone (D6), and the text model's `repair`,
`save-classes.test.mjs` CR-025 section 10's: each skipped item's class held to Word check 76's own verdicts (every file Word opened silently holds no lost item but `76t71`, every one it refused that holds an item holds a lost one, the refusals listed held to the bridge's results), each row of section 10.3's table in memory, a part holding both classes, and the copy,
`unknown-elements.test.mjs` CR-026's: every check 76 source Word saved held to Word's text, section 3.7's pass by shape (D7 to D11), the hosts by parent, a fragment refused (D12), and `76t71`, which Word reads differently, pinned as a known difference,
`nothing-selected.test.mjs` CR-029's: Word's rounds (`test/nothing-selected/`, copied unedited with a SHA-1 manifest, outside `fixtures/`): the 142 files read as Word's text by both load routes and `mce.nothing_selected` exactly where Word shows neither branch, `rw-24` at a root, the `tc-empty` round, each level's warning and its `detail`, the pass's removal and where it keeps the element, the default load keeping it, and `wpi`'s resolved round trip,
`main-part-root.test.mjs` CR-028 D10 and D11's: Word's 40 files of a main part's root (`test/main-part-root/`, docx4j's at `4b26e7422` and the bridge's `rq-*`, each named, with a SHA-1 manifest) held to a table of which root reads as what, the roots read held to Word's text by both load routes, those refused refused, the rest empty with `part.unexpected_root`, their writes refused and their saves copied, and the shapes no file holds,
`undeclared-ignorable.test.mjs` CR-030's: the portfolio's `ui-*` round (`test/undeclared-ignorable/`, with a SHA-1 manifest), `mce.undeclared_prefix` per part by both load routes and in the text model, Word's text, `part.skipped`'s `ignorable`, and the saves to REQ-323's outcome,
`para-ids.test.mjs` CR-002 section 66's: every route that adds a paragraph or a row, on Word's own save, leaving each with a `w14:paraId`, as Word needs to keep a file's ids (check 79),
`numid-zero-indent.test.mjs` CR-027's: docx4j's `NumIdZeroIndentTest` mirrored, `w:numId 0` through `getEffectivePPr` and `Paragraph`, no golden having the shape,
`text-options-comments.test.mjs` CR-012 section 29's: one setting for comments (WD 0.91, python's CR-035), held as options and as text, the old `include: comments` spelling and the refusal of a call contradicting itself,
`text-caps.test.mjs` CR-012 section 24's: docx4j's `CaseMappingTest`, `DropCapJoinTest` and `TextOptionsProfileTest` mirrored, the
generated table's version and source, `w:dstrike`, and a Word save under `profile: word-text`,
`text-model.test.mjs` the text model itself (docx4j's `TextsTest`, `TextsTocTest` and `TextsPhase2Test` where phase 2 reaches, and
this package's own: a byte-for-byte save after a reading, `officeOffset`, the part overloads under one `TextContext`;
a `SET`'s result never shown, on docx4j's field probe and Word's save of it in `test/fixtures/fields/`, CR-012 section 26; a part that cannot be read, on section 28.1's probe built in memory and CR-020's store probe, CR-012 section 28),
`markdown.test.mjs` the Markdown (docx4j-markdown's `MarkdownTextModelTest`, `MarkdownExportTest`, `MarkdownExportExtensionsTest`,
`MarkdownAlternateContentTest`, `MarkdownHostileLinksTest`, `MarkdownHostileMarkupTest` and `MarkdownVerticalAlignTest`, with docx4j's
strings, docx4j-core's `LinkTargetsTest`, and `blocks()`),
`mc-fun.test.mjs` CR-013's: docx4j's `mc-fun` probes (`test/mc-fun/`, copied with a manifest, outside `fixtures/`) read by
every reader against Word 365's saves, docx4j's `TrackedChangesAlternateContentTest` (twins, resolve-all, rows, cells, records), D4's
refusals, and the mutators reaching every branch,
`readers.test.mjs` CR-012 phase 4's: `Paragraph.text`, `Body.text`, whole ranges and `search` against Word 365's JSON on check
48's, 48b's and 51's files (`check48Bytes`, `check51Bytes` build them in memory; never commit the `.docx`), cells, controls,
comments and notes of several paragraphs, equations, hidden text, the `textbox:`
addresses agreeing with the text model over every fixture holding a text box, the read-only boxes, the notes' and comments'
addresses, and `outline()`; and check 58's file and snippet (`test/fixtures/check58/`, `check58Bytes`) held to each other and
run through the shim in both modes and held to Word's runs (K1 to K6), and check 59's (`test/fixtures/check59/`, where a
wildcard matches a mark) through the shim until Word's run is back,
`ranges.test.mjs` CR-014's (and docx4j CR-043 section 17's text-left probe with Word's saves in `test/s17/`, with a SHA-1
manifest, Accept All and Reject All held to Word's): check 52's file (`check52Bytes`, `test/fixtures/check52/`) and its Script Lab snippet held to each
other, every case parsed by the snippet's own parser and every word it searches for in the file once, Word 365's run and its
report after Accept All of the same cases, and `acceptAllRevisions()` on Word's save after the run held to Word's own Accept All,
region for region (`test/check52/`, Word's saves, outside `fixtures/` like `mc-fun/`, with a SHA-1 manifest; phase 0; cases 86
to 95 are CR-002 section 55's); and phase 1's reads, held to Word's by running the check's own snippet (`CHECK52.runCheck`)
through the Word shim on the same file and on Word's saves, and check 53's (`test/fixtures/check53/`, phase 1's unmeasured
choices, read only, Word's runs in Simple and All Markup) the same way, every difference left listed; and phase 2's writes,
the snippet's `run` mode through the shim held to Word's run, to each region's markup in Word's save after the run, and,
accepted, to Word's Accept All; and check 54's (`test/fixtures/check54/`, phase 2's unmeasured choices and phase 3's
controls; Word's Accept All save in `test/check54/`) the same way, its run and, accepted, its Accept All,
`characters.test.mjs` CR-002 section 60's: check 55's and 55b's files (`check55Bytes`, `check55bBytes`, `test/fixtures/check55/`) and their Script Lab snippets
held to each other, every case parsed by the snippet's own parser, and the snippet run through the Word shim; Word 365's
run (its JSONs beside the build, its save in `test/check55/`, outside `fixtures/`, with a SHA-1 manifest), the shim's run of
both held to Word's case for case, the decided differences listed, and each section 60 rule,
`bound-controls.test.mjs` CR-002 section 61's: check 56's snippet file and snippet held to each other and the snippet run through the Word
shim, and its four files to open holding their shapes (`check56Bytes`, `check56OpenFiles`, `test/fixtures/check56/`); Word 365's
run (its JSON beside the build, its saves and repairs in `test/check56/`, with a SHA-1 manifest), the shim's run held to
Word's, and each section 61 rule,
`fields-equations.test.mjs` CR-002 section 62's: check 57's file (`check57Bytes`, `test/fixtures/check57/`) and its Script Lab snippet held
to each other; Word 365's run (its JSON beside the build, its saves in `test/check57/`, with a SHA-1 manifest), the shim's run held to
Word's case for case, this package's save after the run and Word's, each accepted and rejected here, held to Word's saves region for
region, the readers on Word's saves, and section 62.3's rules (a `w:fldSimple` split, an equation deleted under tracking),
`theme-colors.test.mjs` CR-008 phase 2's: docx4j's colour vectors (`test/theme-colors/`, copied unedited with a SHA-1 manifest,
outside `fixtures/`), every probe case resolved through its host's `ThemeColors` and held to Office's reading at docx4j's bar and to
docx4j's own tallies family by family, docx4j's `ThemeColorsTest` and `ThemeColourOutputTest`'s resolver cases, and `Font.color` on the
probes against what Word drew,
`runprops-word365.test.mjs` CR-015's: docx4j's run property probes (`test/runprops-word365/`, copied unedited with a SHA-1 manifest,
outside `fixtures/`), each of the 376 markers held to Word's save of its probe through the resolver, the text model, `Font` and a round
trip; docx4j's `RunPropertiesWord365Test`, `OnOffAttributesTest` (every `Boolean` attribute of the wml mapping), `ColourNormalFormTest`
and `ColourNormalFormOptOutTest`; D8's warnings and save report; and check 60's file (`check60Bytes`, `test/fixtures/check60/`, built
from 25 of the probes) and snippet, run through the shim and held to Word 365's run (`60-runprops-word365.json`), the WordApiDesktop
properties the shim lacks held to the resolver; and check 61's (`check61Bytes`, `test/fixtures/check61/`: the sixteen highlight names
read and written) through the shim, held to Word 365's run with the objects table's 39 differences listed; and check 62's
(`check62Bytes`, `test/fixtures/check62/`: 7,225 colours written to a highlight, for Word's snapping rule) through the shim, and
Word's run (`62-highlight-snap-word365.json`) pinned where it speaks plainly (the grey and channel ramps),
`placeholders.test.mjs` CR-002 section 63's: check 63's file (`check63Bytes`, `test/fixtures/check63/`: Office JS's writes into a
control showing its placeholder) and snippet held to each other, Word 365's run (its JSONs beside the build, its saves in `test/check63/`, with a SHA-1
manifest) pinned case by case: run 1, in Design Mode, and run 2, with it off, which the shim's run is held to (section 63.17); and check 64's
(`test/fixtures/check64/`: the same writes on controls in Word's form, with a glossary Word made, and on controls Word makes) through the shim,
and Word's run (its saves in `test/check64/`) pinned the same way; and check 65's two files (`test/fixtures/check65/`: D8's form with no glossary part,
and with a glossary of one part) through the shim, and Word's runs (its saves in `test/check65/`) pinned: the reference gives Word's form, only an entry keeps it through Word's save, and a glossary without its styles loses the Placeholder Text style, and check 66's two (`test/fixtures/check66/`: D9's glossaries, the document and styles parts, and Word's six) through the shim, and Word's runs (its saves in `test/check66/`) pinned: the two alike;
and section 63.12's: a tracked Replace over a control's markup, accepted and rejected (D3), the level beside a block control (D6), bound nodes (D7),
placeholders in Word's form (D8) and the glossary written on save (D9); and check 67's three files (`test/fixtures/check67/`: check 58's writes into
the bare form in four phases, the view and the tracking mode, through two lookups, for why checks 58 and 63 disagree) and snippet held to each other and run through the shim,
and Word's run (its saves in `test/check67/`) pinned and the shim held to it: every write a plain run, as check 58's, check 58's save pinned too, and only check 63's run 1 saves holding `w:formsDesign`;
and check 68's (`test/fixtures/check68/`: the run properties a write into a placeholder takes, and an entry under a control's own) file and snippet held to each other, and Word's run (its saves in `test/check68/`) pinned and the shim held to it;
and check 69's (`test/fixtures/check69/`: a write at a paragraph's End past a placeholder control, a hidden run or a control holding text, by every route) file and snippet held to each other and run through the shim, and Word's run (its saves in `test/check69/`) pinned and the shim held to it (section 63.27);
and check 70's (`test/fixtures/check70/`: the same at the endings check 69 left, a bookmark, comment markers, a tracked deletion or insertion, a field, a text box, and D13's runs before a control) file and snippet held to each other and run through the shim, and Word's run (its saves in `test/check70/`) pinned and the shim held to it, its three non-placement differences listed (section 63.27),
`measurements.test.mjs` CR-002 section 65's: check 71's file (`check71Bytes`, `test/fixtures/check71/`: Office JS's measurement setters past
Word's 16-bit ranges) and snippet held to each other and run through the Word shim; Word 365's run (its JSON beside the build, its save in `test/check71/`,
with a SHA-1 manifest) pinned, and every case where the shim differs listed (section 65.5's D1 and D2); and each setter's limits, message and refusal before any change;
and check 71b's file (`check71bBytes`, `test/fixtures/check71b/`: the same setters at and either side of Word's own limits, below storage, for D2) and check 71's
snippet, a copy held equal; Word 365's run (its JSONs beside the build, its save in `test/check71b/`, with a SHA-1 manifest) pinned, and the shim's run held to
it case for case, its one difference listed (section 65.9),
`check72.test.mjs` CR-016 D5's: check 72's 18 files (`check72Files`, `test/fixtures/check72/`: one misplaced child each, per rule of docx4j CR-064's five sets, for
Word's repairs), and each file's fate here (a `target.misplaced` warning at the load, the readers see two, the save refused but 72m's, whose string the runtime dropped at the read, and under the opt-out fifteen written back as held; CR-016 sections 7 and 8;
the bare breaks 72e and 72j, since CR-021, `br.outside_run`, nothing misplaced, and the save Word's save's structure), and Word 365's run (its saves and screenshots in `test/check72/`, with a SHA-1 manifest): 11 files don't open (no repair offered), 7 are repaired silently, the child moved to the nearest place it fits, a string in a run dropped (section 7.1),
`misplaced-content.test.mjs` CR-016 phase 2's: section 2's five direct edits refused at save with the part, the path and docx4j's fix, nothing written; the opt-out writing as held
onto `pkg.lastSave.misplaced`; `checkContent()` on check 72's files leaving the part untouched and the untouched save unrefused; D7's warning fields per file; the gate over every
fixture and `mc-fun` file; the facade's walk reaching a container built without its `TYPE_NAME`, since objects 0.8.0,
`integer-attributes.test.mjs` CR-017's and CR-002 section 65.3's: docx4j CR-065's and CR-066's probes and Word 365's saves (`test/integer-attributes/`, copied
unedited from the share with a SHA-1 manifest, outside `fixtures/`): rounds 5 and 6's 660 integer cases, a `value.out_of_range` warning on load and on
`pkg.lastSave` exactly where Word's save changed the value, its rule or its attribute (432, the objects package's own count), `wordReads` what Word wrote,
every value written as held, Word's saves reporting nothing; round 2's 64 decimals and odd forms read as Word's saves have them (the six known
differences stated, strict both ways; the refused forms dropped and reported); `keptRange` held to the objects package's `wordIntegerReading` at every
boundary of its 49 entries; `WORD_STORAGE`'s rows the objects table's, narrowed only where a row says why; a value set by code reported on save with its
path; no fixture holding one,
`store-items.test.mjs` CR-020's: docx4j CR-058's five probes and Word 365's saves (`test/store-items/`, copied unedited with a SHA-1 manifest,
outside `fixtures/`), docx4j's `StoreItemRepairTest` and `BoundDateFormatTest` mirrored, every control after `applyBindings` held to Word's save (the
dates in Word's zone), the readers writing nothing, and check 73b's snippet through the shim held to Word's runs (checks 73 and 73b's JSONs),
`br-placement.test.mjs` CR-021's: docx4j CR-073's br-placement probe and Word 365's saves (`test/br-placement/`, copied unedited with a SHA-1 manifest,
outside `fixtures/`): each probe's `br.outside_run`, `Paragraph.text` and the text model held to Word's save (`br2-3`'s bare field code, D5, pinned), the save
after unmarshalling unrefused and each story part's tokens (docx4j `BrOutsideARunTest`'s comparison, `breakTokens` in `helpers.mjs`) Word's save's, an untouched
save byte for byte, docx4j's edge tests mirrored, `insertOoxml`'s two routes and a caller's own hook, and check 76's `76m`, `76n`, `76u`, `76w` and `76z` (`check76Files`, `test/fixtures/check76/`) held to Word's saves (`test/check76/`, with a SHA-1 manifest), the shapes Word refuses or nothing measures pinned,
`dates.test.mjs` CR-018's: the zone option and its refusals, a revision stamped and a `w:date` without `dateUtc` read in five zones (docx4j's `PackageDateTimeZoneTest`
mirrored), a comment to the minute, the bound-date opt-in off and on, `writeTimesOnSave` on a created, a loaded and a stripped package, `clone()`,
`duplicate-entries.test.mjs` CR-024's: a package holding two entries of one name, each store, the option (`repair: 'word'` since CR-031, `duplicate_entries` refused on `Texts.of`), the warning and the save, and Word check 75's saves (`test/check75/`, with a SHA-1 manifest; `test/fixtures/check75/build.mjs` builds the files, git-ignored); and section 9's, more than one officeDocument relationship, held to Word 365's first-open saves of docx4j's probes (`test/package-wiring/`, copied unedited with a SHA-1 manifest, outside `fixtures/`),
`table-rules.test.mjs` CR-002 section 55's: docx4j's `TableRulesTest` ported, and each edit that could leave a cell without a
last paragraph, untracked and tracked, accepted and rejected (check 52's cases 86 to 95 are held to Word's in `ranges.test.mjs`),
`examples.test.mjs` runs every example in a child process and asserts a line of its
output); helpers in `test/helpers.mjs` (`plain()` strips `PARENT` for deep equality). Fixtures under `test/fixtures/` are from docx4j's
`docx4j-core-tests` resources; `test/README.md` lists them and holds the manual Word acceptance
checklist. The contract: untouched parts byte-identical after a round trip, re-marshalled parts
deep-equal after reload, flat OPC through the objects package's `unmarshalPackage` and back.
A defect in a dependency is recorded as a test that **asserts the broken behaviour**, with the
upstream commit in a comment, never as a skip: the arriving fix then announces itself by failing
that assertion (CR-001 section 19).
`test/golden/` holds 62 JSON goldens - what **docx4j itself** answers for each fixture
(effective properties, style resolutions, table styles and contexts, each table paragraph's
in-context properties, list labels and counters, the document font of every character), written by the Maven harness in `test/java/` and compared by
`test/parity.test.mjs`, which unmarshals both sides and deep-equals object trees, never text.
Regenerate by hand (build docx4j, then `mvn -q -o compile exec:java -Dfixtures=../fixtures
-Dout=../golden -Ddocx4j.commit=<hash>` in `test/java`; each README says more);
`.github/workflows/parity.yml` does it weekly against docx4j's head and opens a pull request
when an answer moves. Parity is zero differences. `test/golden/xpath/` is CR-006's oracle: what
docx4j's `getJAXBNodesViaXPath` selects for 47 expressions per fixture, made by the same Maven
project's `XPathHarness` (`-Dexec.mainClass=org.docx4j.parity.XPathHarness`) and compared by
`test/xpath-parity.test.mjs`. `test/golden/revisions/` is CR-011's: what docx4j's
`org.docx4j.model.revisions` (its CR-043) lists of each of the 96 fixtures with revision markup, at
any depth, and what each resolve-all leaves, with and without a filter, tracked changes and (harness
version 2, CR-002 section 54) comments (`RevisionsHarness`, in UTC; a golden's path mirrors its
fixture's; recorded from the docx4j 17.3.6 release, `4afcf4be8`), compared by
`test/revisions-parity.test.mjs`. `test/golden/text/` is CR-012 phase 6's: what docx4j's `org.docx4j.model.text`
(`Texts.of`, `PlainText.of`) answers for each of the 338 committed documents under `test/fixtures/` and `test/text-extraction/`,
in four readings (the three views, and every story and optional item included), as the block model and the plain text
(`TextHarness` version 2, cut from the 17.3.6 release, CR-001 section 27, from 17.3.5 before, CR-012 section 24.7; `-Dfixtures=../fixtures,../text-extraction`; a golden's
path is the directory's name and the document's), compared by `test/text-parity.test.mjs` with the runs joined by docx4j's rule
(CR-012 section 15.5), its `PENDING` table setting aside what phase 5 and REQ-303 govern, each required to still differ;
REQ-702's date form, masked until CR-012 section 23.7, now agrees with nothing masked.

## Rules

- ES modules only (`"type": "module"`, `.mts` sources built to `dist/*.mjs` with `.d.mts`). Unlike
  the objects package there are no UMD files here, so `"type": "module"` is fine.
- Public paths are the `exports` map only (`.`, `./opc`, `./parts`, `./packages`, `./model`,
  `./office-js`, `./node`, `./xpath-fonto`, `./html`). Keep them stable; add subpaths deliberately. **`./node` is the only
  one that may import a `node:` builtin** (`DirectoryPartStore`); nothing else under `src/` may,
  so a browser or add-in bundle never sees one. There is no `@types/node`: what that module uses
  is declared in `src/node/node-builtins.d.mts`.
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
  CR-001 was written against `VERSION_17_1_1`, renamed `VERSION_17_2_0` on 2026-09-21 when 17.1.1
  became the 17.2.0 release; development moved to `VERSION_17_3_0` on 2026-09-27 when 17.2.1 was
  released, to `VERSION_17_3_1` on 2026-10-02 when 17.3.0 was released (tag `docx4j-17.3.0`,
  `883a08e07`), to `VERSION_17_3_2` on 2026-10-07 when 17.3.1 was released (tag `docx4j-17.3.1`,
  `5cf47d752`; 17.3.1 ships CR-030, the table context, and otherwise PDF-side work that changes no
  XML), to `VERSION_17_3_3` on 2026-10-09 when 17.3.2 was released (tag `docx4j-17.3.2`,
  `856afc5da`), to `VERSION_17_3_4` on 2026-10-17 when 17.3.3 was released (tag `docx4j-17.3.3`,
  `c95b55be6`; 17.3.3 ships CR-043, the tracked-changes API this package's CR-011 follows), to
  `VERSION_17_3_5` on 2026-10-22 when 17.3.4 was released (tag `docx4j-17.3.4`, `8f455ebab`; 17.3.4 ships
  CR-046, the text model this package's CR-012 follows), to `VERSION_17_3_6` on 2026-11-01 when 17.3.5
  was released (tag `docx4j-17.3.5`, `64fae4cd0`), and to `VERSION_17_3_7` on 2026-11-08 when 17.3.6 was
  released (tag `docx4j-17.3.6`, `4afcf4be8`; 17.3.6 ships `RepairMode`, docx4j CR-075, which this package's
  CR-031 follows). The parity workflow checks out `VERSION_17_3_7` (since 2026-11-10), and
  `test/java/pom.xml` resolves the 17.3.6 release for a hand-run (the workflow reads docx4j's
  `<revision>` instead), the release every golden is cut from (CR-001 section 27). The classes being
  ported are under `docx4j-core/src/main/java/org/docx4j/`: `openpackaging/` (`io3/` for
  Load3/Save and `io3/stores/` for `PartStore`, `parts/relationships/` for `RelationshipsPart`
  and `Namespaces`, `contenttype/`, `packages/`), `model/PropertyResolver.java`,
  `model/listnumbering/`, `model/styles/` (`StyleUtil`, `PropertyCatalogue`), `fonts/`;
  `jaxb/McSelection.java` and `McMode.java` are the `mc:AlternateContent` reference (CR-021),
  `jaxb/mc-preprocessor.xslt` the load-time one. The schemas are `xsd/ROOT.xsd`. Phase B was
  ported against `7fba7a150`; the goldens recorded `a8d20c1cc` (then `VERSION_17_3_0`, refreshed by
  the weekly workflow on 2026-09-28 with every answer unchanged) until CR-007, then `0349796f9`
  (`VERSION_17_3_1`, CR-030 D11), then `856afc5da`, the 17.3.2 release from Maven Central (CR-001
  section 22), then `2e381ef26` (`VERSION_17_3_3`, CR-034's colour merge, CR-008 section 6), then
  `c95b55be6`, the 17.3.3 release from Maven Central (CR-011 phase 0: no answer moved), then
  `8f455ebab`, the 17.3.4 release (CR-012 phase 6: no answer moved), then `5163c2986`
  (`VERSION_17_3_5`, pushed; CR-015 section 11: docx4j `82c64bf14` reads `w:framePr`'s `anchorLock` as
  stated, which moved eight answers, ported; CR-060 and CR-061 moved none), and now record `4afcf4be8`,
  the 17.3.6 release, every golden alike (CR-001 section 27: no property, XPath or revisions answer moved;
  the text goldens gained `repaired`, and one ink Choice now read as this package already read it),
  harness version 4, which needs docx4j 17.3.1's table context; the project does not compile against
  anything before 17.3.5, whose CR-069 run members `TextHarness` version 2 calls. Build docx4j from the commit you mean to record before regenerating them
  by hand (or resolve a release from Central), and never modify that checkout.

## Portfolio task registry

This repository's change requests are indexed, with their dependencies on work in the other
docx4j repositories, in `../docx4j-portfolio/tasks.yaml` (ids `<repo>/<CR>[.<phase>]`; this
repository's key is `core-ts`).

- When a CR's status changes (a phase lands; a CR is proposed, deferred or abandoned) or its
  dependencies change, the matching entry in `tasks.yaml` changes in the same session (`status`,
  `depends_on`; a new entry for a new CR or phase).
- Since 2026-10-17 (Jason) this repository does not edit `../docx4j-portfolio/tasks.yaml`,
  `scripts/status-snapshots.yaml` or `docs/dependencies.md` itself. Send the change to the session
  working in `../docx4j-portfolio` (find it with ListAgents by that directory, not by its name's
  suffix, which changes): the entry id, the new `status`, `depends_on` and `notes` (or the whole
  entry), and any CR Status line to re-baseline. That session applies it, runs `tasks.py check`,
  `graph` and `accept`, and commits.
- Only when no session is running in `../docx4j-portfolio`, make the change from this session:
  edit the files, run `python3 ../docx4j-portfolio/scripts/tasks.py check` (it reports `CHANGED`
  for each CR whose Status line was edited; once the registry entry agrees, `tasks.py accept`,
  which baselines every repository's changed lines, and `tasks.py graph` if dependencies changed),
  and commit there only your own hunks, through a temporary index from `HEAD`, never the shared
  index.
- Before starting a CR or phase, check `python3 ../docx4j-portfolio/scripts/tasks.py blocked`: it
  may be waiting on work in another repository. The portfolio rule of 2026-09-15 (XSL-FO fidelity
  work in `../docx4j` before porting docx4j code) was lifted for CR-001 Phase B on 2026-09-19
  (section 14.5); the editor (`../docx4j-ts-editor`) waits on phases here.
