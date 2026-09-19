# CR-002: A content API in the shape of Office JS, over the docx4j tree

**Status:** Phases B and D implemented 2026-09-10 (section 7); phase A implemented 2026-09-10 as objects CR-002; phases C, G and I implemented 2026-09-15 (sections 8, 9 and 10); phases E and F implemented 2026-09-16 (sections 12 and 13; section 11 corrects `style` / `styleBuiltIn`); phase H implemented 2026-09-19 (section 17); section 16: `Font` and `Paragraph` reads effective since 2026-09-19 (CR-001 Phase B).
**Depends on:** CR-001 Phase A (parts and packages; implemented). The tree-level half depends on
an objects-package CR (its CR-002, proposed below) because it needs only the object model.
**Counterpart:** docx4j `MainDocumentPart.addParagraphOfText` / `addStyledParagraphOfText` /
`addObject` / `getContent`, `XmlUtils.unmarshalString` and `W_NAMESPACE_DECLARATION`,
`TraversalUtil`, `ClassFinder`, `TextUtils`, `org.docx4j.wml.ObjectFactory`; and the tool surface
of `plutext/docx4j-mcp`.
**Editor references:** "CR-003" below is the web editor design, which moved to
`plutext/docx4j-ts-editor` as `docs/change-requests/ED-001-web-editor.md` with its section numbers unchanged.

## 1. Why

The README's Hello World is honest and low-level:

```ts
body.content.push(w('p', { content: [w('r', { content: [w('t', { value: 'Hello World' })] })] }));
```

Three things make it heavier than it should be. Every element is a `{ name, value }` pair, so
the user writes the qualified name of `w:p`, `w:r` and `w:t` by hand. The tree is the whole
API: there is no "add a paragraph" and no "find the paragraph that says X". And an XML snippet,
the most natural input for anyone who has seen a docx, needs its `xmlns:w` declaration and,
once unmarshalled, a `deepCopy(value, parent)` to get its parent pointer right.

This CR also covers custom XML parts, XML mapping and typed content controls, which
WordApiDesktop 1.3 (August 2026) brought to add-ins and which docx4j has always had; see 3.5.

Two kinds of consumer are in view. A **developer** who knows Word documents and wants the
common operations to read the way they do in docx4j, python-docx or Office JS. An **agent**
(Claude Code through an MCP server, or code an LLM writes against this package) that wants
few operations, string in and string out, stable addresses for "the third paragraph", and no
way to produce an invalid tree by accident.

## 2. What others do

| Library | Building | Reading and editing | Shape |
|---|---|---|---|
| docx4j (Java) | `mdp.addParagraphOfText("x")`, `addStyledParagraphOfText("Heading1","x")`, `addObject(o)`; `XmlUtils.unmarshalString(xml)` with `W_NAMESPACE_DECLARATION` | `mdp.getContent()`, `TraversalUtil` / `ClassFinder`, XPath via `getJAXBNodesViaXPath`, `TextUtils.extractText` | methods on the part; the tree everywhere else |
| Office JS (Word add-ins) | `body.insertParagraph(text, "End")`, `insertText`, `insertOoxml(ooxml, "Replace")`, `insertHtml`, `insertTable(r, c, "Start", values)` | `body.paragraphs`, `body.text`, `body.search(text, options)` returning ranges, `getOoxml()` | fluent proxies with insert locations `Start` / `End` / `Replace` / `Before` / `After` |
| docx (docx.js) | `new Document({ sections: [{ children: [new Paragraph({ children: [new TextRun("Hello")] })] }] })`, then `Packer` | none: write-only | constructor tree, options objects |
| python-docx | `document.add_paragraph("x", style="Heading 1")`, `paragraph.add_run("bold").bold = True` | `document.paragraphs`, `paragraph.text`, `paragraph.runs` | methods on document and paragraph, properties for formatting |
| Open XML SDK (.NET) | `body.AppendChild(new Paragraph(new Run(new Text("x"))))` | LINQ over `Descendants<Paragraph>()`, `InnerText` | typed DOM |
| Office-Word-MCP-Server (GongRzhe) | `add_paragraph(file, text, style)`, `add_heading(file, text, level)`, `add_table(file, rows, cols, data)`, `add_picture`, `insert_line_or_paragraph_near_text(file, target_text, line_text, position)` | `get_document_text`, `get_document_outline`, `get_paragraph_text_from_document(file, index)`, `find_text_in_document`, `search_and_replace`, `format_text(file, paragraph_index, start, end, bold, ...)`, `delete_paragraph(file, index)` | flat tools: text in, text out, paragraph index as the address |
| docx4j-mcp (this project) | `markdown_to_docx`, `html_to_docx`, `fill_template` | `docx_to_markdown`, `extract_text`, `describe_template` | whole-document conversions; no in-place editing yet |

Two observations. Every developer-facing library converges on the same handful of verbs on the
body or the part: add a paragraph of text (with a style), add a table, add an image, insert at
start or end or relative to something, list paragraphs, get text, search. The agent-facing
servers converge on the same verbs plus **addresses** (a paragraph index, or "near this
text") and **string payloads** (text, markdown, HTML, OOXML). Nothing here needs to be
imitated; the convergence is what matters, and it maps almost one to one onto docx4j's own
`MainDocumentPart` helpers plus Office JS's insert locations.

## 3. Design: the shape of Office JS, over the docx4j tree

Decision (2026-09-10): where Office JS has a name or a shape for an operation, this API uses
it. The package is meant to be used next to Word add-ins, so a body, a paragraph, a range, a
table and an insert location should mean the same thing on both sides of `getOoxml()`. Two
further reasons: Office JS is the most documented Word content API there is, so an agent
writing code against this package already knows the vocabulary; and a **structural subset** of
`Word.Body`, `Word.Paragraph`, `Word.Range` and `Word.Table` lets one function run unchanged
against a live document in Word and against a package here. docx4j's own helper names
(`addParagraphOfText`, `addStyledParagraphOfText`, `addObject`, `getContent`) stay as thin
aliases, since docx4j Java users expect them and they cost nothing. This is a deliberate
departure from "names follow docx4j": docx4j-core has no content API to follow beyond those
four methods, and CR-001 section 11.6 allows departures when the Java reads badly here.

What is mimicked is the **object model and the method surface**, not the execution model.
Office JS proxies are batched and resolved by `context.sync()`, with `load()` before a
property is readable. Here everything is in memory: properties are plain, methods act at once,
and only what marshals or unmarshals (`getOoxml`, `insertOoxml`) is asynchronous. Office JS
code ported here drops its `load` / `sync` lines and gains nothing else to learn.

### 3.1 The surface (this package, `src/model/content/`)

```ts
// Word.Body (a subset): on WordprocessingMLPackage.body, and on any part whose root has a
// content list (main document, headers, footers, footnotes, endnotes, comments, glossary).
interface Body {
  readonly paragraphs: Paragraph[];          // document order; tables descended
  readonly tables: Table[];
  readonly contentControls: ContentControl[];
  readonly text: string;                     // paragraph per line
  insertParagraph(text: string, location: 'Start' | 'End'): Paragraph;
  insertText(text: string, location: 'Start' | 'End' | 'Replace'): Range;
  insertTable(rowCount: number, columnCount: number, location: 'Start' | 'End', values?: string[][]): Table;
  insertBreak(type: 'Page' | 'Line' | 'SectionNext' | 'SectionContinuous', location: 'Start' | 'End'): void;
  insertInlinePictureFromBase64(base64: string, location: 'Start' | 'End'): InlinePicture;  // adds the ImagePart and the relationship
  insertOoxml(ooxml: string, location: 'Start' | 'End' | 'Replace'): Promise<Range>;         // pkg:package, as Word; or a bare w:p / w:tbl fragment (extension)
  search(text: string, options?: { matchCase?: boolean; matchWholeWord?: boolean; matchWildcards?: boolean }): Range[];
  clear(): void;
  getRange(location?: 'Whole' | 'Start' | 'End' | 'Content'): Range;
  getOoxml(): Promise<string>;
  // extensions, not in Office JS
  readonly content: TypedNamedValue[];        // the live tree (docx4j getContent()); sectPr excluded
  insertElement(element: TypedNamedValue, location: 'Start' | 'End' | 'Before' | 'After', target?: Paragraph | Table): void;  // docx4j addObject
  insertXml(xml: string, location: ...): Promise<Range>;   // a w:p / w:tbl fragment, namespaces added for you
}

// Word.Paragraph (a subset)
interface Paragraph {
  text: string;                               // read: runs joined (w:t, w:tab, w:br, w:sym); write: replaces the runs with one
  style: string;                              // the display name ('Heading 1'), as Office JS; setting accepts a name or an id
  styleBuiltIn: string;                       // the Word.Style value ('Heading1'), 'Other' when not built in, as Office JS
  readonly styleId: string;                   // extension: w:pStyle (docx4j's name); also settable
  alignment: 'Left' | 'Centered' | 'Right' | 'Justified' | 'Unknown';
  readonly font: Font;                        // paragraph mark run properties; setting applies to every run
  leftIndent: number; firstLineIndent: number; spaceAfter: number; spaceBefore: number; lineSpacing: number;  // points, as Office JS
  outlineLevel: number;
  readonly parentBody: Body;
  readonly parentTableCell: TableCell | undefined;
  insertText(text: string, location: 'Start' | 'End' | 'Replace'): Range;
  insertParagraph(text: string, location: 'Before' | 'After'): Paragraph;
  insertBreak(type: 'Page' | 'Line', location: 'Before' | 'After' | 'Start' | 'End'): void;
  insertInlinePictureFromBase64(base64: string, location: 'Start' | 'End' | 'Replace'): InlinePicture;
  insertOoxml(ooxml: string, location: 'Before' | 'After' | 'Start' | 'End' | 'Replace'): Promise<Range>;
  search(text: string, options?: SearchOptions): Range[];
  getRange(location?: 'Whole' | 'Start' | 'End' | 'Content'): Range;
  getOoxml(): Promise<string>;
  delete(): void;
  // extensions
  readonly element: TypedNamedValue<P>;       // the tree
  readonly paraId: string | undefined;        // w14:paraId, the stable address (3.3)
}

// Word.Range (a subset): a run of text within one paragraph, or a whole paragraph
interface Range {
  text: string;
  readonly font: Font;
  readonly paragraphs: Paragraph[];
  style: string;
  insertText(text: string, location: 'Before' | 'After' | 'Start' | 'End' | 'Replace'): Range;
  insertParagraph(text: string, location: 'Before' | 'After'): Paragraph;
  insertOoxml(ooxml: string, location: 'Before' | 'After' | 'Replace'): Promise<Range>;
  search(text: string, options?: SearchOptions): Range[];
  delete(): void;
  getOoxml(): Promise<string>;
  // extension: the runs it spans, split at the boundaries when set through font or text
  readonly runs: TypedNamedValue<R>[];
}

// Word.Font (a subset), backed by w:rPr
interface Font {
  bold: boolean; italic: boolean; underline: 'None' | 'Single' | 'Double' | 'Dotted' | 'Wave' | ...; strikeThrough: boolean;
  subscript: boolean; superscript: boolean;
  name: string;          // w:rFonts ascii/hAnsi
  size: number;          // points (w:sz is half-points)
  color: string;         // '#RRGGBB'
  highlightColor: string | null;   // a highlight name or '#RRGGBB' mapped through helpers/wml
}

// Word.Table (a subset), backed by w:tbl
interface Table {
  readonly rowCount: number; readonly rows: TableRow[]; values: string[][];
  style: string; styleBuiltIn: string; headerRowCount: number;
  getCell(rowIndex: number, cellIndex: number): TableCell;
  addRows(location: 'Start' | 'End', rowCount: number, values?: string[][]): TableRow[];
  deleteRows(rowIndex: number, rowCount?: number): void;
  delete(): void;
  readonly element: TypedNamedValue<Tbl>;
}

// Word.ContentControl (a subset), backed by w:sdt: tag, title, text, paragraphs, insertText, delete(keepContent)
```

`insertOoxml` takes what Word's does, a flat OPC `pkg:package` string (the CR-001 container
reads it; the parts it brings, an image say, are added to this package with new relationship
ids and the content is re-targeted). As an extension it also accepts a bare fragment, which
is what `insertXml` is; an agent that already produces OOXML for `insertOoxml` in an add-in
can send the same string here.

The Hello World:

```ts
const pkg = await WordprocessingMLPackage.createPackage({ pageSize: 'A4' });
pkg.body.insertParagraph('Hello World', 'End');
await writeFile('hello.docx', await pkg.save());
```

And the same code against a live document, for comparison:

```ts
await Word.run(async (context) => { context.document.body.insertParagraph('Hello World', 'End'); await context.sync(); });
```

### 3.2 Tree layer (objects package, its CR-002): fragments and constructors

Needs only the object model, so it lives in `@docx4j/generated-objects-ts` (`builders/wml`).
It is what section 3.1 is built on and what an add-in that never loads a package uses on its
own.

```ts
import { wml, p, r, t, tbl, textOf, walk, find } from '@docx4j/generated-objects-ts/builders/wml';

const para = await wml`<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Hello</w:t></w:r></w:p>`;
// the standard declarations (the facade's NAMESPACE_PREFIXES) are added, so snippets read as
// they do inside document.xml; docx4j: W_NAMESPACE_DECLARATION. `wml.all` for several siblings.
p([r('Hello, ', { bold: true }), r('world')], { style: 'Heading1' })   // the ObjectFactory the model lacks
t('  two spaces  ')                                                     // xml:space="preserve" for you
textOf(para)                                                            // docx4j TextUtils
walk(root, visitor); find(root, 'org_docx4j_wml.P')                     // TraversalUtil, ClassFinder
```

#### An ObjectFactory for the model? (decided 2026-09-10; landed the same day)

docx4j's `org.docx4j.wml.ObjectFactory` (11.5.15: 10,821 lines) is two things. **375 no-argument
creators** (`createP()`, `createR()`, `createText()`) exist because Java beans need a
constructor call; in TypeScript a typed literal (`const p: P = { ... }`) does that with the
compiler checking the shape, so that half is nearly redundant. The one thing a creator adds is
`TYPE_NAME`, which the unmarshaller sets and a literal does not; new content therefore does
not answer `find(root, 'org_docx4j_wml.P')` unless the author writes `TYPE_NAME` by hand.

The **834 `JAXBElement` wrappers** (`createRT(Text)` for `w:t` in a run, `createBodyPermStart`,
`createSettings(CTSettings)` for the root) are not redundant at all: they are exactly the
`{ name, value }` pairs this CR's section 1 complains about, and in the TypeScript model they
are needed for *more* elements than in Java, because every item of a content array is a
`TypedNamedValue` here (Java put `P` into `Body` directly through `@XmlElements`; here it is
`{ name: w:p, value }`). Nobody should type qualified names by hand, and the mappings already
know every element name and the type it takes in each scope.

Decision: generate it, do not hand-write it. Compiler CR-010 (implemented 2026-09-10 with the
modifications recorded in its sections 8 and 9) emits per module `factory/<module>` and
`el/<module>` in the objects package:

- `factory/org_docx4j_wml`: creators per non-abstract class, `createP(init?)`, setting
  `TYPE_NAME`; wrappers per element declaration under XJC's own squeezed names, so scoped
  ones read as in docx4j (`createRT(value)` for `w:t` in a run) and global ones take an
  `Element` suffix (`createPElement`, `createDocumentElement`), since 28 of WML's 37 global
  element names are also class names.
- `el/org_docx4j_wml`: one wrapper per element name of the module's namespace (170 in WML),
  `el.p(value)`, `el.r(value)`, `el.t(value)`, setting `TYPE_NAME` where the name determines
  the type; `el.sdt` and `el.customXml` (one name, four types) are union-typed and leave it.

The README's Hello World now uses `el`:
`body.content.push(el.p({ content: [el.r({ content: [el.t({ value: 'Hello World' })] })] }))`.
The hand-written `p()`, `r()`, `t()`, `tbl()` of 3.2 are therefore the *text-in* sugar only
(`p('Hello')` makes the run and the text, `t('  x ')` sets `xml:space`), built on `el`, and
`wml\`...\`` stays the string route.

### 3.3 Addresses and the agent surface

Office JS addresses content by object identity within one `context.sync()`; an agent across
tool calls needs something serialisable. Three address forms, accepted wherever a
`Paragraph | Table` target is, and returned by `outline()`:

- **Ordinal:** `'body/3'` (the fourth block-level child of the body), `'body/3/1'` (a row, or a
  run), `'header:rId3/0'`. Stable until the content changes.
- **paraId:** `'w14:5A2B1C3D'` where Word wrote one; this package writes `w14:paraId` on
  paragraphs it creates when the document already uses them. Stable across edits.
- **Text:** `{ contains: 'Chapter 1' }`, first match, as the MCP servers' `near_text` tools.

```ts
await pkg.outline();
// { paragraphs: [{ address: 'body/0', paraId: '5A2B1C3D', style: 'Heading1', text: 'Chapter 1' }, ...],
//   tables: [{ address: 'body/4', rows: 3, cols: 2 }], headers: [...], footers: [...] }
pkg.body.paragraphAt('body/3').insertParagraph('New', 'After');
```

The MCP tool set of section 2 then maps onto one-liners: `add_paragraph` is
`body.insertParagraph(text, 'End')` plus `style`; `insert_line_or_paragraph_near_text` is
`body.search(text)[0].paragraphs[0].insertParagraph(line, 'After')`; `delete_paragraph(index)`
is `body.paragraphAt('body/3').delete()`; `get_document_outline` is `outline()`;
`search_and_replace` is `body.search(a).forEach(r => r.insertText(b, 'Replace'))`, which is
exactly the Office JS idiom. `docx4j-mcp` can grow in-place editing tools over this package
for the add-in and Node cases.

### 3.4 Compatibility with Office JS, enforced

A compile-time test (`test/office-js-subset.ts`, as the objects package's
`test/readme-examples.ts` does with Office JS stubs) declares the intended subset as
TypeScript interfaces copied from `@types/office-js` and checks that `Body`, `Paragraph`,
`Range`, `Table` and `Font` here are assignable to them, minus the `load` / `sync` / `context`
members and with `Promise<T>` where this package marshals. A function written as
`(body: OfficeSubset.Body) => ...` then type-checks against both. Where Office JS returns a
collection proxy (`ParagraphCollection`) this package returns an array; the subset types say
`ArrayLike<Paragraph> & { items: Paragraph[] }` so `.items` works on both sides.

### 3.5 Custom XML parts, XML mapping and typed content controls (WordApiDesktop 1.3)

WordApiDesktop 1.3 (August 2026) gives add-ins a native XML mapping model: `Word.XmlMapping`
on a content control (`setMapping(xpath)`, `setMappingByNode(node)`, `prefixMappings`,
`isMapped`, `customXmlPart`, `customXmlNode`, `delete()`), a `CustomXmlPart` /
`CustomXmlNode` model with XPath queries, node insertion and prefix mappings, and typed
content controls. That is docx4j's home ground (`w:dataBinding`, `CustomXmlDataStoragePart`,
OpenDoPE's `BindingHandler` and `XPathsPart`), so the same surface is offered here, over the
same tree. Building blocks and `Template` objects are out of scope.

```ts
// Word.CustomXmlPartCollection (subset): pkg.customXmlParts
pkg.customXmlParts.items                       // CustomXmlPart[], as pkg.customXmlDataStorageParts (CR-001) by itemId
pkg.customXmlParts.getByNamespace(ns)          // the parts whose document element is in that namespace
pkg.customXmlParts.getItem(id)                 // by itemId ('{...}', case-insensitive)
pkg.customXmlParts.add(xml)                    // new CustomXmlDataStoragePart + properties part with a fresh itemID, related from the main part

// Word.CustomXmlPart (subset), backed by CustomXmlDataStoragePart (a DOM) and its properties part
interface CustomXmlPart {
  readonly id: string;                         // itemId
  readonly namespaceUri: string;               // of the document element
  readonly builtIn: boolean;                   // core/extended/cover-page properties stores
  readonly documentElement: CustomXmlNode;
  readonly namespaceManager: CustomXmlPrefixMappingCollection;   // addNamespace(prefix, uri), lookupNamespace(prefix), lookupPrefix(uri), items
  readonly schemaCollection: string[];         // ds:schemaRef URIs from the properties part
  getXml(): string;
  setXml(xml: string): void;
  selectNodes(xpath: string, namespaceMappings?: string): CustomXmlNode[];   // 'xmlns:ns0=...' as Word's prefixMappings, or the manager's
  selectSingleNode(xpath: string, namespaceMappings?: string): CustomXmlNode | undefined;
  insertElement(xpath: string, xml: string, namespaceMappings?: string, index?: number): CustomXmlNode;
  updateElement(xpath: string, xml: string, namespaceMappings?: string): void;
  deleteElement(xpath: string, namespaceMappings?: string): void;
  insertAttribute(xpath: string, name: string, value: string, namespaceMappings?: string): void;
  updateAttribute(...); deleteAttribute(...);
  delete(): void;                              // removes the part, its properties part and the relationships; unlinks mappings
  readonly part: CustomXmlDataStoragePart;     // extension: the part
}

// Word.CustomXmlNode (subset), a view over a DOM node
interface CustomXmlNode {
  readonly baseName: string; readonly namespaceUri: string; readonly nodeType: 'Element' | 'Attribute' | 'Text' | ...;
  nodeValue: string; text: string; readonly xml: string; readonly xpath: string;   // xpath: the canonical /ns0:a[1]/ns0:b[2] form Word writes
  readonly attributes: CustomXmlNode[]; readonly childNodes: CustomXmlNode[]; readonly parentNode: CustomXmlNode | undefined;
  readonly firstChild: CustomXmlNode | undefined; readonly lastChild: ...; readonly nextSibling: ...; readonly previousSibling: ...;
  readonly ownerPart: CustomXmlPart;
  hasChildNodes(): boolean;
  selectNodes(xpath: string, namespaceMappings?: string): CustomXmlNode[];
  selectSingleNode(xpath: string, namespaceMappings?: string): CustomXmlNode | undefined;
  appendChildNode(xml: string): CustomXmlNode;   // or (name, namespaceUri, nodeType, nodeValue) as Office JS
  insertNodeBefore(xml: string, nextSibling: CustomXmlNode): CustomXmlNode;
  removeChild(child: CustomXmlNode): void;
  replaceChildNode(oldNode: CustomXmlNode, xml: string): CustomXmlNode;
  delete(): void;
  readonly node: Node;                          // extension: the DOM node
}

// Word.XmlMapping (subset): contentControl.xmlMapping, backed by w:dataBinding
interface XmlMapping {
  readonly isMapped: boolean;
  readonly xpath: string;                       // w:xpath
  readonly prefixMappings: string;              // w:prefixMappings, "xmlns:ns0='...' xmlns:ns1='...'"
  readonly customXmlPart: CustomXmlPart | undefined;      // by w:storeItemID
  readonly customXmlNode: CustomXmlNode | undefined;      // the XPath evaluated now
  setMapping(xpath: string, prefixMappings?: string, part?: CustomXmlPart): boolean;   // false when the XPath selects nothing, as Word
  setMappingByNode(node: CustomXmlNode): boolean;          // writes node.xpath and the node's prefixes
  delete(): void;                               // removes w:dataBinding
}

// Word.ContentControl (subset), typed by w:sdtPr: type is the kind Word reports
interface ContentControl {
  readonly type: 'RichText' | 'PlainText' | 'Picture' | 'DatePicker' | 'ComboBox' | 'DropDownList' | 'CheckBox' | 'RepeatingSection' | 'RepeatingSectionItem' | 'Group' | 'BuildingBlockGallery' | 'Citation' | 'Bibliography' | 'Equation';
  tag: string; title: string; id: number; text: string; placeholderText: string;
  appearance: 'BoundingBox' | 'Tags' | 'Hidden'; color: string; cannotDelete: boolean; cannotEdit: boolean; removeWhenEdited: boolean;
  readonly xmlMapping: XmlMapping;
  readonly paragraphs: Paragraph[]; readonly tables: Table[]; readonly contentControls: ContentControl[];
  // the kind-specific properties, undefined for the other kinds (Office JS: checkboxContentControl, ...)
  readonly checkboxContentControl?: { isChecked: boolean; checkedSymbol; uncheckedSymbol };
  readonly datePickerContentControl?: { dateFormat: string; dateDisplayLocale: string; dateCalendarType; dateStorageFormat };
  readonly dropDownListContentControl?: { listItems: { displayText: string; value: string }[]; addListItem(...); deleteAllListItems() };
  readonly comboBoxContentControl?: ...;
  readonly pictureContentControl?: { readonly inlinePicture: InlinePicture | undefined };
  readonly repeatingSectionContentControl?: { readonly items: ContentControl[]; allowInsertDeleteSection: boolean; sectionTitle: string; insertItemAfter(index)... };
  readonly groupContentControl?: {};
  insertText(text: string, location: 'Start' | 'End' | 'Replace'): Range;
  insertParagraph(...); insertOoxml(...); search(...); getRange(...); getOoxml(); delete(keepContent: boolean): void;
  readonly element: TypedNamedValue<SdtBlock | SdtRun | CTSdtRow | CTSdtCell>;   // extension
}
body.insertContentControl(kind?: ContentControl['type']): ContentControl        // wraps the body's content, as Office JS
paragraph.insertContentControl(kind?): ContentControl                           // wraps the paragraph
range.insertContentControl(kind?): ContentControl                               // wraps the run(s)
```

The model already types all of it: `CTDataBinding` (`xpath`, `prefixMappings`,
`storeItemID`), `SdtPr`'s union of kinds (`CTSdtText`, `CTSdtDate`, `CTSdtDropDownList`,
`CTSdtComboBox`, `SdtPr.Picture`, `SdtPr.Group`, `SdtPr.RichText`, `CTSdtDocPart`, w14
`CTSdtCheckbox`, w15 `CTSdtRepeatedSection`, `CTSdtAppearance`), and the custom XML part's
`DatastoreItem`. The views are a translation, not a new model. Two pieces of docx4j come with
them and are worth exposing under docx4j's names: `BindingHandler.applyBindings()` (push the
custom XML values into the bound controls, what Word does on open) and the reverse
`XmlDataStorage.updateFromControls()`; they are `pkg.customXmlParts.applyBindings()` and
`updateFromContentControls()` here, and are the Phase E of this CR because
OpenDoPE's repeats and conditions build on them.

### 3.6 XPath: how doable

Very. Custom XML parts are a DOM here already (CR-001 `CustomXmlDataStoragePart`), and XPath
1.0 over a DOM is a solved problem in both runtimes:

- **Browsers and Office add-ins:** `document.evaluate(expr, node, resolver, type)` is native,
  XPath 1.0 complete, and takes a namespace resolver built from the prefix mappings.
- **Node:** the `xpath` package (goto100, MIT, 0.0.34, 180 KB unpacked, no dependencies) is
  the standard companion of `@xmldom/xmldom`, the DOM the runtime uses in Node; `xpath.useNamespaces(map)`
  gives a selector with prefixes.

Measured on 2026-09-10 with the docx4j `OpenDoPE/invoice.docx` fixture, through this
package: the three `w:dataBinding` XPaths of the document (`/invoice[1]/customer[1]/name[1]`
and two `items/item[1]/...` paths) each select their node from the custom XML DOM, with the
prefix map parsed from `w:prefixMappings`, in 4.2 ms for the three including the part parse;
`count(//*)` and `name(/*)` work too. Word's own binding XPaths are simple location paths with
positional predicates, well inside XPath 1.0.

Design: an `XPathEngine` interface with one method,
`select(expr: string, context: Node, namespaces: Record<string, string>): Node[]` (plus
`selectValue` for string, number and boolean results). The default engine uses
`document.evaluate` when it exists and otherwise `import('xpath')`; `xpath` is an **optional
peer dependency** so add-in bundles do not carry it, and `pkg.xpathEngine` can be set to
anything else (a consumer with a different DOM, or one who wants no dynamic import). The
canonical `xpath` a `CustomXmlNode` reports is computed here (walk to the root counting
same-name siblings), which is what `setMappingByNode` and Word's `XmlMapping.xpath` write.

XPath over the **WordprocessingML tree** (the typed objects, not a DOM) is a different
matter and not proposed: docx4j does it by marshalling to a DOM and mapping nodes back through
a JAXB `Binder`, which this runtime does not have. The typed tree has `walk` / `find` (3.2) and
`PARENT`, and the agent surface has addresses; that covers the uses. If demand appears, the
route is a DOM view over the typed tree, its own CR.

### 3.7 Change tracking and find-and-replace (phase F; requested by CR-003 section 4.2)

Office JS has `document.changeTrackingMode` (`Off` | `TrackAll` | `TrackMineOnly`),
`Word.TrackedChange` (`type`: `Added` | `Deleted` | `Formatted` | `None`; `author`, `date`,
`text`; `accept()`, `reject()`, `getRange()`) and `getTrackedChanges()` on body, paragraph
and range. The same here:

- `pkg.changeTrackingMode`, with `trackedChangeAuthor` and an optional fixed date (the
  package's, since there is no signed-in user). When on, every mutation of the content API
  writes Word's revision markup instead of editing in place: `insertText` wraps the new
  run(s) in `w:ins`; a deletion moves the runs into `w:del` with `w:t` turned into
  `w:delText`; a replacement is a `w:del` followed by a `w:ins`; `insertParagraph` marks the
  new paragraph mark inserted (`w:pPr/w:rPr/w:ins`) and `delete()` marks it deleted;
  `Font` and paragraph property writes keep the old properties in `w:rPrChange` /
  `w:pPrChange`; row insertions and deletions go on `w:trPr`. Ids come from a per-package
  counter above the highest in use; author and date from the settings.
- The text model already reads the accepted view (`w:del` excluded, `w:ins` included), so
  `text`, `search` and addresses see the document as it will be once accepted, which is
  what a reviewer and an agent want; a `{ view: 'original' }` option on `text` and `search`
  gives the other view when needed.
- `TrackedChange` views over `w:ins`, `w:del`, `w:rPrChange`, `w:pPrChange` and the
  paragraph-mark and row forms: `accept()` unwraps or removes, `reject()` restores;
  `getTrackedChanges()` on `Body`, `Paragraph` and `Range`, and `acceptAll()` /
  `rejectAll()` on `Body` (docx4j's accept-all in the Java `docx4j-mcp` is the
  reference). The editor renders and lists them (CR-003 section 4.2).
- `replaceText(find, replace, options?)` on `Body`, `Paragraph` and `Range`: `search`, then
  `insertText(replace, 'Replace')` on each match from last to first, so offsets stay valid;
  returns the count. Tracked when the mode is on. This is the verb agents and the UI's
  Find and Replace call; the Office JS idiom remains available.

### 3.8 Comments (phase G; requested by CR-003 section 3.3)

Office JS has `Word.Comment` (`authorName`, `authorEmail`, `content`, `creationDate`,
`resolved`, `replies`, `reply(text)`, `delete()`, `getRange()`), `Range.insertComment(text)`
and `getComments()` on body, paragraph and range. The same here, over the four parts a
current Word writes, all typed and loaded by CR-001 already:

- `w:comments` (`CommentsPart`): the comment's id, author, initials, date and its content
  paragraphs (the first opening with a `w:annotationRef` run in `CommentReference` style).
- `w:commentRangeStart` / `End` in the body around the commented range, and a run holding
  `w:commentReference` after it.
- `w:commentsExtended` (`CommentsExtendedPart`, w15): per comment, keyed by the
  `w14:paraId` of its first paragraph, the `w15:done` flag and `w15:paraIdParent` for a
  reply, which is how threads are formed.
- `w:commentsIds` (w16cid `durableId`) and `w:people` (authors' presence data), and, when
  present, `w:commentsExtensible` (w16cex, a UTC date), kept in step.

Operations: `getComments()` returns `Comment` views in document order, replies nested under
their parent; `content` reads the comment's text (the same text model), and setting it
replaces the paragraphs; `resolved` reads and writes `w15:done`; `reply(text)` adds a
comment with the same range whose first paragraph's `paraId` is linked by `paraIdParent`;
`delete()` removes the comment, its replies, the range markers, the reference run and the
entries in the three side parts; `Range.insertComment(text)` allocates the next id and
`paraId`, writes the markers and reference run, the comment with one paragraph in
`CommentText` style, and the side-part entries, creating any of the four parts that are
absent (with their relationships and content types, through `addTargetPart`). Author,
initials and email come from package-level settings, the same as `trackedChangeAuthor`; an
agent's comments carry its name. A comment's range may span paragraphs; `getRange()`
returns a `Range` per paragraph it touches (extension: Office JS returns one range).

### 3.9 Lists (phase H; requested by CR-003 section 3.4)

Office JS has `Word.List` (`id`, `levelTypes`, `levelExists`, `getLevelParagraphs`,
`setLevelNumbering`, `setLevelBullet`, `setLevelIndents`), `Word.ListItem` (`level`,
`listString`, `siblingIndex`), and on `Paragraph`: `isListItem`, `list`, `listItem`,
`startNewList()`, `attachToList(listId, level)`, `detachFromList()`. The same here over
`w:numPr` and the numbering part:

- `List` is a view over a `w:num` and the `w:abstractNum` it points to (through
  `NumberingDefinitionsPart`), `id` being the `w:numId`; `levelTypes` reads each level's
  `w:numFmt` as `Bullet` or `Number` (Office JS's two kinds); `setLevelNumbering` /
  `setLevelBullet` / `setLevelIndents` write the level's `w:numFmt`, `w:lvlText`, `w:start`
  and `w:ind`, copying the abstract definition first when other `w:num`s share it, so a
  change stays local to this list.
- `Paragraph.startNewList()` creates a `w:abstractNum` from docx4j's default definitions
  (`NumberingDefinitionsPart.unmarshalDefaultNumbering`: a bullet set and a decimal set;
  the resource is embedded here as `DEFAULT_STYLES_XML` is) and a `w:num` for it, adds the
  numbering part and its relationship when the document has none, and sets the paragraph's
  `w:numPr`; `attachToList(listId, level)` sets `w:numPr`; `detachFromList()` removes it;
  `listItem.level` reads and writes `w:ilvl`. A restart is a new `w:num` on the same
  abstract definition with a `w:lvlOverride` / `w:startOverride` for level 0.
- `listItem.listString` (the rendered label) and `siblingIndex` need CR-001 Phase B's
  `Emulator`. **Phase B landed on 2026-09-19** (`Emulator`, `NumberingState`, `NumberingStates`,
  `pkg.getNumberingEmulator()`), so phase H is unblocked: `listString` is the label
  `emulator.getNumber(p.pPr, state)` answers, counted in the state of the paragraph's story,
  and the editor's list rendering (CR-003 section 3.4) no longer waits.
- A paragraph whose style carries `w:numPr` (a "List Number" style) is a list item through
  the style; `isListItem` is true and `list` resolves through the style's `w:numPr` (Phase
  B's `PropertyResolver` gives the effective `numPr`; before it, the direct one only).

### 3.10 The `Word` shim and the source generators (phase I; requested by CR-003 section 3.5)

Section 3.4 promises that the content API is a structural subset of `Word.*`. Phase I
makes that executable: a `Word` object (`@docx4j/core-ts/office-js`) whose `Word.run(fn)`
calls `fn` with a `context` over a package (`Word.run(pkg, fn)` here, since there is no
host document), where:

- `context.document.body` is the package's `Body`; `context.document.getSelection()` is a
  `Range` the caller supplies (the editor's selection; none in Node); `context.document.
  changeTrackingMode`, `properties` (core properties through `DocPropsCorePart`),
  `getComments()`.
- `load(...)` on any object is a no-op returning the object; `context.sync()` resolves the
  pending results: `getOoxml()` and other asynchronous calls return `ClientResult`-shaped
  objects whose `value` is filled when `sync()` resolves, as in Office JS.
- Collections are arrays with an `items` property (the subset types already say
  `ArrayLike & { items }`); `getFirst()`, `getFirstOrNullObject()` (with `isNullObject`)
  are provided since add-in code uses them constantly.
- The enums are objects of the string values the API accepts: `Word.InsertLocation`,
  `Word.Alignment`, `Word.UnderlineType`, `Word.ChangeTrackingMode`, `Word.BreakType`,
  `Word.ContentControlType`, `Word.SearchOptions` shape.
- Every object handed out is a proxy: a member that the subset does not implement throws
  `NotSupportedError` naming the member, so an add-in's unsupported call fails at the call
  with a clear message rather than silently. A `Word.supported` set lists the implemented
  members, generated from the subset declarations, for tools that want to check a script
  before running it (CR-003's console).

Use in Node: run an add-in's `Word.run` callbacks against a package in a test, with the
saved docx as the assertion; this is a way to test add-in logic in CI without Word.

Two generators, for reveal codes and for agents learning the API from examples:

- `toSource(element)` in the objects package's `builders/wml` (its next CR): the factory
  call that builds the element.
- `toApiScript(element | Paragraph | Range)` here: the content-API calls that produce it
  where verbs exist (`insertParagraph`, `style`, `font`, `insertTable`, `insertContentControl`,
  ...), falling back to `insertXml` with the marshalled fragment for the rest; emitted as
  TypeScript against `body`.

### 3.11 Scripts and directionality (foundations; CR-003 section 3.7)

The content API targets left-to-right text first and must not assume it. Commitments,
with what phases B and D already do and what is still to be done:

- Offsets in `Range`, `search`, addresses and `TextSegment` are UTF-16 code units, as in
  Office JS; the split in `Paragraph.splitAt` must never fall inside a surrogate pair or
  a grapheme cluster: **to do**, `Intl.Segmenter` (grapheme) to move the split to the
  cluster boundary, with a test on an emoji sequence and on Devanagari.
- `matchWholeWord` uses `(?<![\w])...(?![\w])`, which is ASCII-minded: **to do**,
  `Intl.Segmenter` word boundaries when available (all current browsers and Node 16+),
  the regex as the fallback; `matchCase: false` uses the `i` flag, which is Unicode
  case-insensitive under the `u` flag: **to do**, add `u`.
- The run mapping writes `w:bCs`, `w:iCs` and `w:szCs` with their Latin twins (done), so
  complex-script text is formatted by the same calls; `Font` gains `nameBidi`,
  `nameEastAsia` and `sizeBidi` (WordApiDesktop) with Phase B's font work: **later**.
- `Paragraph.alignment` maps `start`/`end` as well as `left`/`right` (done); the indent
  properties read `w:ind/@w:left` and `@w:right` and should also read the strict-form
  `@w:start` and `@w:end` and write whichever the paragraph already uses: **to do**.
- A `bidi` property on `Paragraph` (`w:bidi`) and `rtl` on `Font` (`w:rtl`), so the editor
  can set direction without the tree: **to do**, small.
- Nothing in the text model depends on direction: `w:bidi` and `w:rtl` are display, and
  the logical order of runs is the document order, which is what the model keeps.

### 3.12 What stays as it is

- `contents` remains the typed tree and `content` arrays remain plain arrays. Nothing is
  wrapped in proxies; `Paragraph` and friends are light views holding a reference to the
  element and its container, created on access and cheap to discard. A user who pushes to
  `body.content` directly keeps working (with the `deepCopy` idiom for `PARENT`).
- The packaging layer keeps docx4j's names; only the content API takes Office JS's. Where
  Office JS has no name (the tree, addresses, `insertXml`, `outline`), the extension is marked
  as such in the doc comment.
- Everything mutates the tree in place, marks the part for re-marshalling (CR-001's rule), and
  links `PARENT`. Inserts reject an element the container cannot hold, naming what was passed:
  the guard an agent needs so that a wrong tree fails at insert time, not in Word.

## 4. Where each piece lives

| Piece | Package | Why |
|---|---|---|
| `wml` fragments, constructors, `textOf`, `walk`, `find` | objects (`builders/wml`, its CR-002) | tree only; add-ins that never load a package want them |
| generated element wrappers and creators (docx4j `ObjectFactory` names: `createRT`, `createPElement`, `createP`; `el.p`) | compiler CR-010 (done); objects package ships `factory/*` and `el/*` | mechanical, complete, one source for QNames |
| `Body`, `Paragraph`, `Range`, `Table`, `Font`, `ContentControl` views; docx4j aliases; `insertOoxml` / `insertXml` | this package (`src/model/content/`) | needs parts (image parts and relationships, headers and footers by relationship) |
| addresses, `outline()`, `paragraphAt` | this package | needs the package (headers, footers, footnotes) |
| `CustomXmlPart`, `CustomXmlNode`, `XmlMapping`, typed `ContentControl`, `XPathEngine` (native `evaluate` or the optional `xpath` package) | this package (`src/model/customxml/`) | needs parts (the custom XML part, its properties part, `storeItemID`) |
| the Office JS subset types and the assignability test | this package (`test/office-js-subset.ts`) | the compatibility promise lives with the implementation |
| MCP tools over the above | `docx4j-mcp` (or a new `docx4j-mcp-ts`) | product, not library |

## 5. Phasing

| Phase | Content | Effort |
|---|---|---|
| A | objects `builders/wml`: `wml` template with the declarations, `p`/`r`/`t`/`tbl`, `textOf`, `walk`, `find`; tests against marshalled XML | 2 days |
| B | `Body` and `Paragraph` with `insertParagraph`, `insertText`, `text`, `style`, `font`, `search` (within runs), `delete`, `getOoxml`, `insertXml`; `Range` over runs; the Office JS subset test; README rewritten around `pkg.body` | 4 days |
| C | `Table`, `TableCell`, `insertTable`, `insertInlinePictureFromBase64` (ImagePart + relationship + `wp:inline`), `insertOoxml` from a `pkg:package`, `ContentControl` | 4 days |
| D | addresses (ordinal, paraId, text), `outline()`, `paragraphAt`; a note in `docx4j-mcp`'s CR | 2 days |
| E | `CustomXmlPart` / `CustomXmlNode` with the `XPathEngine`, `XmlMapping`, typed `ContentControl` properties, `insertContentControl`; `applyBindings` / `updateFromContentControls` (docx4j `BindingHandler`) | 5 days |
| I | The `Word` shim (3.10): `Word.run(pkg, fn)`, `load` / `sync`, collections with `items`, enums, proxies throwing `NotSupportedError`, `Word.supported`; `toApiScript`; usable in Node to test add-in code against a package | 4 days |
| H | Lists (3.9): `List` and `ListItem` views, `startNewList` from docx4j's default definitions, `attachToList`, `detachFromList`, level, restart; labels from Phase B | 4 days |
| G | Comments (3.8): `getComments()`, `Comment` views with `content`, `resolved`, `reply`, `delete`, `getRange`; `Range.insertComment`; the four comment parts created and kept in step | 4 days |
| F | Change tracking (3.7): `changeTrackingMode` in Office JS's shape, every mutation writing `w:ins` / `w:del` / `w:rPrChange` / `w:pPrChange` / row revisions when on, `TrackedChange` views with `accept` / `reject`, `getTrackedChanges()`; `replaceText(find, replace, options)` as a verb, tracked or not; the text model's accepted view already excludes `w:del` | 5 days |

## 6. Open questions

1. How far to take the Office JS surface. Recommendation: the members listed in 3.1, which
   are the ones add-in code uses for content; not `load`, `track`, `select`, `getComments`,
   fields or the collection proxies. Add members on demand, keeping assignability.
2. Units. Office JS reports indents and spacing in points; the tree holds twips and
   half-points. Recommendation: points on the views, as Office JS, converted on read and write;
   the tree stays as it is.
3. `search` across runs (Word splits words across runs freely). Recommendation: within runs in
   Phase B; cross-run matching with run splitting and merging in its own CR, shared with the
   Java implementation's approach.
4. Ordinal addresses count block-level children of the body (paragraphs, tables, sdt blocks,
   altChunks) or paragraphs only? Recommendation: block-level children, so a table is
   addressable; `outline()` lists both so the agent never counts.
5. Whether `insertOoxml` with a `pkg:package` merges styles and numbering from the incoming
   package or only its content. Recommendation: content plus the parts the content references
   (images, embedded objects) in this CR; style merging is what docx4j's MergeDocx does and is
   a separate product.
6. Tagged template versus function for `wml`. Recommendation: both, the template calling the
   function; the template can interpolate typed values in place.
7. Whether `xpath` should be an optional peer dependency (dynamic import when there is no
   native `document.evaluate`) or a hard dependency. Recommendation: optional peer; Node users
   install it once, add-in bundles never see it, and `pkg.xpathEngine` stays pluggable.
8. `CustomXmlNode.xml` / `appendChildNode(xml)` accept XML text, as Word does, but Office JS's
   `appendChildNode(name, namespaceUri, nodeType, nodeValue)` form also exists. Recommendation:
   both signatures, since the text form is what agents produce.

## 7. Implementation notes: phases B and D (2026-09-10)

`src/model/content/` (`Body`, `Paragraph`, `Range`, `Font`, `search`, `fragments`, `tree`),
exported from `.` and `./model`; `MainDocumentPart.body` / `getBody()`, the same on
`HeaderPart` and `FooterPart` (address prefixes `header:<relId>`, `footer:<relId>`);
`WordprocessingMLPackage.body`, `getBody()`, `outline()`, `paragraphAt()`. Seven tests in
`test/content.test.mjs`; `test/office-js-subset.ts` is the assignability check of section 3.4
(compiled by `npm run typecheck`). The README's Hello World is now
`pkg.body.insertParagraph('Hello World', 'End')`.

What is in: `Body` `paragraphs` (descending into tables and content controls), `text`,
`insertParagraph`, `insertText`, `insertBreak`, `insertElement` (validated against the Body
content union, PARENT linked), `insertXml`, `search`, `clear`, addresses (`elementAt`,
`paragraphAt`, `addressOf`, `outline`), the docx4j aliases; `Paragraph` `text` (get and set),
`style`, `styleBuiltIn`, `alignment`, the indents and spacing in points, `outlineLevel`,
`font`, `paraId`, `runs`, `insertText`, `insertParagraph`, `insertBreak`, `search`,
`getRange`, `delete`; `Range` `text`, `font` (runs split at the boundaries), `insertText`
(a replacement keeps the formatting of the first replaced character, as Word), `delete`,
`search`; `Font` bold, italic, underline, strike, sub/superscript, name, size, color,
highlightColor; search with `matchCase`, `matchWholeWord`, `matchWildcards`, matching
**across runs** (question 3 turned out cheap: matches are found on the paragraph's text and
mapped back to segments; only formatting needs the split). New paragraphs get a `w14:paraId`
when the document already uses them.

Departures and deferrals, all deliberate:

- `getOoxml()` is not offered; `getXml()` returns the part's or paragraph's XML. Wrapping a
  fragment in a `pkg:package` with the styles it needs is Phase C's `insertOoxml` work.
- `Font` reads report direct formatting of the first run in scope (Office JS reports effective
  formatting); CR-001 Phase B supplies the resolver. `Paragraph.font` is over the runs, not the
  paragraph mark. `alignment` is `'Unknown'` when `w:jc` is absent, for the same reason.
- `Body.tables` returns elements; `Table` views are Phase C. `insertXml` returns the inserted
  views (paragraphs, or `{ element, container }`), not a `Range`, since a fragment may hold
  several blocks.
- `styleBuiltIn` mapped ids to names by inserting spaces (`Heading1` to `Heading 1`), and `style`
  was the id. **Corrected 2026-09-16** to Office JS's meaning (section 11): `style` is the display
  name, `styleBuiltIn` the `Word.Style` value or `'Other'`, and the id is the `styleId` extension.
- `outlineLevel` is `w:outlineLvl + 1`, 10 when absent.
- Phase A landed in the objects package on 2026-09-10 (its CR-002, `builders/wml`), and the
  private copies here were replaced by imports the same day: `fragments.mts` and
  `runOptions.mts` are gone; `tree.mts` keeps only the paragraph text model (`segmentsOf`,
  `runsOf`, `runItemsOf`, `childrenOf`, `BLOCK_LEVEL_TYPES`, `runOf`, `paragraphOf`) and
  re-exports `walk`, `find`, `linkParents`, `textOf`, `isElement`, `typeNameOf`. `Font` goes
  through the builders' `applyRunOptions` / `readRunOptions`; `Body.insertXml` calls the
  builders' `wml` with `{ wrapper: 'body', preprocess }`. The `./model` subpath re-exports the
  builders so that one import serves the content API and the tree.
- A finding from that move: the model keeps the runs of `w:ins`, `w:del`, `w:moveFrom` and
  `w:moveTo` under docx4j's property name `customXmlOrSmartTagOrSdt`, not `content`. The text
  model here read `content` and silently dropped text inside tracked insertions; `runItemsOf`
  now reads by the model's names, and a test covers search and formatting inside a `w:ins`
  and the exclusion of `w:del`.

## 8. Implementation notes: phase C (2026-09-15)

`src/model/content/` gains `Table.mts` (`Table`, `TableRow`, `TableCell`, `cellOf`),
`InlinePicture.mts` (the view, the image header readers, `addImage`, `drawingFor`) and
`ContentControl.mts`, with `ooxml.mts` holding the `insertOoxml` machinery that `insertXml`
now shares; all are exported from `.` and `./model`. Nine tests in `test/content-c.test.mjs`
(`test/content.test.mjs` untouched), one new fixture, and `test/office-js-subset.ts` extended
with `Table`, `TableRow`, `TableCell`, `InlinePicture` and `ContentControl`.

What is in:

- **Tables.** `Body.tables` now returns `Table` views (it returned elements in phase B);
  `Body.insertTable(rowCount, columnCount, location, values?)` over the objects package's `tbl`
  builder, with the grid sized to the section's text width. `Table`: `rowCount`, `rows`,
  `values` (get and set), `style`, `styleBuiltIn`, `headerRowCount`, `getCell`, `addRows`,
  `deleteRows`, `delete`, `element`, plus `columnWidths()`, `text` and `parentTableCell`.
  `TableRow`: `rowIndex`, `cellCount`, `cells`, `values`, `isHeader`, `insertRows`, `delete`.
  `TableCell`: `body` (a `Body` over the `w:tc`), `paragraphs`, `tables`, `text`, `value`,
  `insertParagraph`, `insertText`, `rowIndex`, `cellIndex`, `parentRow`, `parentTable`,
  `width` / `columnWidth`. `Paragraph.parentTableCell` is section 3.1's.
- **Pictures.** `Body.insertInlinePictureFromBase64(base64, location, options?)` and the same on
  `Paragraph` (with `'Replace'`): an `ImagePart` under `/word/media/imageN.<ext>`, a
  relationship from the part the body belongs to (the main document part, or the header or
  footer), and a `w:drawing`/`wp:inline` built exactly as docx4j's
  `BinaryPartAbstractImage.createImageInline` writes it. `imageInfoOf` reads PNG (IHDR and
  pHYs), JPEG (SOFn and the JFIF APP0 densities), GIF and BMP headers, which is what docx4j
  gets from XML Graphics Commons' `ImageInfo`, with the same 96 dpi default; EMU is
  `px / dpi * 914400`, and an image wider than the text area is scaled down as docx4j's
  `CxCy.scale` does. `InlinePicture` has `width`, `height` (points), `altTextDescription`,
  `altTextTitle`, `imageFormat`, `getBase64()` (and Office JS's name `getBase64ImageSrc()`),
  `delete()`, `paragraph`, plus `relId`, `inline` and `imagePart()`. `Body.inlinePictures` and
  `Paragraph.inlinePictures` are Office JS collections that phase C needed for its own tests.
- **`insertOoxml`** on `Body`, `Paragraph` and `Range`. A flat OPC `pkg:package` is loaded with
  CR-001's container, its body content taken, and every part that content references through
  relationships copied into the target package under a free name with a fresh relationship id;
  the `r:embed` / `r:id` / `r:link` references in the content are rewritten. A bare `w:p` /
  `w:tbl` fragment is accepted as well, which is what `insertXml` takes: both go through
  `contentOf()`.
- **Content controls.** `ContentControl` over `w:sdt` in all four forms, with `form`, `type`,
  `tag`, `title`, `id`, `text`, `paragraphs`, `tables`, `contentControls`, `insertText`,
  `insertParagraph`, `search`, `getRange`, `delete(keepContent)`, `getXml` and `element`;
  `Body.contentControls` and `Paragraph.contentControls` in document order, nested ones
  included. `insertContentControl`, the typed kinds and `w:dataBinding` stay phase E.

Departures and deferrals, all deliberate:

- `insertOoxml` returns the inserted views, as `insertXml` does, not the single `Range` of
  section 3.1: a package may bring several blocks (section 7 recorded the same for `insertXml`).
  At paragraph and range level, a fragment of exactly one `w:p` has its runs merged into the
  paragraph at `'Start'` / `'End'` (range: `'Before'` / `'After'` / `'Replace'`), which is what
  Word's paste does; anything else is inserted as blocks before or after the paragraph.
- `insertInlinePictureFromBase64` is **synchronous**, as Office JS's is: the `wp:inline` is
  built from the generated factories (`factory/org_docx4j_dml`, `.../dml_picture`,
  `.../dml_wordprocessingDrawing`, `el.drawing`, `el.pic`) rather than parsed from the XML
  template docx4j uses, so nothing has to await the Jsonix context.
- `InlinePicture.altTextTitle` reads and writes `wp:docPr/@title`. Until objects 0.1.3 the
  model had no `title` on `CTNonVisualDrawingProps` (ECMA-376 1st edition) and the attribute was
  not marshalled; closed 2026-09-16 (section 14).
- Styles and numbering are not merged by `insertOoxml` (open question 5, as recommended):
  incoming content keeps its `w:pStyle` and `w:numPr` values, which resolve against the target's
  styles, or dangle. Style merging is what docx4j's MergeDocx does and stays out of scope.
- A copied part is carried as bytes (`ImagePart` for `image/*`, else `BinaryPart`), and its own
  relationships are copied recursively **keeping their ids**, so nothing inside a copied part
  (a chart's `c:chart`, say) has to be rewritten; only the inserted content's references change.
  Part names follow docx4j's `getNewPartName`: `/word/media/imageN.<ext>`, N free for that
  extension, so a package can hold `image1.gif` and `image1.png`, as docx4j's does.
- References are rewritten by attribute name (`embed`, `link`, `id`, `href`, the diagram and
  chart ones) **and** only when the incoming package really has a relationship of that id, so
  numeric ids (`w:bookmarkStart/@w:id`, `wp:docPr/@id`) are never touched. The walker is local
  (`visitReferences`) rather than the objects package's `walk` because it must also enter the
  DOM nodes an `xs:any` property may hold (an SVG twin's `r:embed` in a `a:extLst`).
- `Table.rows` and `TableRow.cells` descend into row- and cell-level content controls (an
  OpenDoPE repeat wraps its `w:tr` in a `w:sdt`), so `rows` is what Word shows. `addRows` puts
  new rows in the table's own content, never inside a repeat.
- `headerRowCount` counts **leading** rows carrying `w:tblHeader`, as Word's own property does;
  a `w:tblHeader` on a later row is left alone by the getter and cleared by the setter.
- `Table.style`, `styleBuiltIn` and `styleId` follow `Paragraph`'s (corrected 2026-09-16, section
  11); `insertTable` sets no style, so a new table is borderless until `styleBuiltIn = 'TableGrid'`
  (the default styles part has that style).
- `ContentControl.getRange()` is exact for a run-level control (the offsets its runs cover) and
  is the first paragraph's range for the block, row and cell forms, because a `Range` here is
  within one paragraph (section 7). `type` is `RichText` when `w:sdtPr` names no kind, as Word
  reports it; `w:text` is `PlainText`. `form` falls back to what the content holds, since
  objects CR-002 notes that a run-level `w:sdt` parsed at body level comes back as `SdtBlock`.
  `appearance`, `color`, `cannotDelete`, `placeholderText` and the kind-specific properties are
  phase E, along with `insertContentControl`. A row- or cell-level control refuses
  `insertParagraph` at `'Start'` / `'End'` and `insertText(..., 'Replace')`, naming what it
  holds, rather than putting a `w:p` next to a `w:tr`.
- Three extensions to phase B's code, all additive: `Body.sub(container, prefix)` (a view over a
  nested container: how a cell and a block content control get a body), `Body.paragraphFor(p)`
  (the view of a `w:p` reached through PARENT), and `Body.elementAt` accepting a prefix that
  itself holds slashes, so a cell's body is addressed `body/4/0/1` and its paragraphs
  `body/4/0/1/0`, which is what `outline()` already emits.
- A bug fixed on the way: `childrenOf` answered `sdtContent` for `SdtBlock` only, so paragraphs
  inside a row- or cell-level control (an OpenDoPE repeat) were invisible to `Body.paragraphs`,
  and `addressOf` gave up inside one. `childrenOf` now answers for every `w:sdt` kind, `pathOf`
  skips all four `SdtContent*` levels, and `rowsOf` / `cellsOf` / `SDT_TYPES` / `Located` are
  the shared traversal (`tree.mts`).
- `ooxml.mts` imports `packages/WordprocessingMLPackage.mjs` **dynamically** inside the
  function: `parts/wml` builds `Body` views, so a static import back would be the runtime cycle
  CLAUDE.md forbids. `insertOoxml` is asynchronous anyway.

Gaps found in the objects package (candidates for its own CRs, worked around here):

1. **No `wp:docPr/@title`** (closed in objects 0.1.3, section 14). `CTNonVisualDrawingProps`
   had `descr` but no `title`, so `altTextTitle` could not round trip.
2. **No drawing or inline-picture constructor in `builders/wml`** (closed in objects 0.1.4,
   section 15: `inlinePicture(relId, options)`). `drawingFor()` here was the
   counterpart of docx4j's `createImageInline` and needs only the tree (the relationship id is
   a string), so it belongs in the objects package next to `p`, `r` and `tbl`. It is written
   over the generated factories, so moving it is a copy.
3. **No row or cell builder** (closed in objects 0.1.4, section 15: `tr` and `tc`). `tbl(rows)`
   builds a whole table; `addRows` and `insertRows`
   need one row of the table's widths, so `rowElement()` here called `el.tr` / `el.tc` directly.
   `tr(cells, opts)` and `tc(blocks, opts)` in `builders/wml` would cover it.
4. `walk` does not enter DOM nodes held by `xs:any` properties (closed in objects 0.1.4,
   section 15: `walkAll`). That is the documented
   behaviour, not a defect, but anything that rewrites references has to know; a `walkAll` (or
   an option) in the objects package would save the copy here.

## 9. Implementation notes: phase G (2026-09-15)

Comments, as section 3.8 specifies them. `src/model/content/Comment.mts` (the `Comment` view),
`src/model/content/comments.mts` (the seam: the author identity, the marker traversal, the id and
date helpers, and the two registries below) and `src/parts/wml/comments.mts` (the part half:
loading and creating the four parts, the w16cex DOM, the two styles). `Body.getComments()`,
`Paragraph.getComments()` / `insertComment()`, `Range.getComments()` / `insertComment()`,
`WordprocessingMLPackage.author`. Eight tests in `test/comments.test.mjs`; `test/office-js-subset.ts`
gains a `Comment` interface and the three `getComments` / two `insertComment` members, and
`test/README.md` a manual Word check (6) and the `comments-two.docx` fixture.

What is in: reading a document Word wrote (author, initials, email from `w:people`, date,
content, `resolved` from `w15:done`, threads from `w15:paraIdParent`, the commented range as a
`Range` per paragraph); `insertComment` on a range or a paragraph, which splits the runs at the
boundaries, writes `w:commentRangeStart` / `End` and the reference run in the `CommentReference`
style, the comment with its paragraphs in `CommentText` style opened by a `w:annotationRef` run
and a fresh `w14:paraId`, and the `w15:commentsEx`, `w16cid:commentsIds`, `w:people` and (when
present) `w16cex` entries; creating any of the four parts that is absent, with its relationship
and content type through `addTargetPart`, and adding the two styles to the styles part;
`reply` (a new comment whose markers nest inside its parent's, linked by `paraIdParent`);
`resolved` as a read/write property; `delete()`, which removes the comment, its replies, every
marker, the reference runs and every side-part entry; `content` get and set; `getComments()` on a
body, a paragraph or a range, in document order with replies nested under their parent.

Departures and deferrals, all deliberate:

- **Asynchronous where Office JS is synchronous.** `getComments`, `insertComment`, `reply` and
  `delete` return promises, because they unmarshal (or create) the comment parts; section 3's
  rule is "only what marshals or unmarshals is asynchronous", and this is that. `resolved`,
  `content`, `replies` and `getRange()` are plain, since by then the parts are loaded.
- **One `Comment` type for comments and replies.** Office JS has `CommentReply` with a subset of
  `Comment`'s members; here a reply is a `Comment` (it is one in the file), `reply()` returns a
  `Comment` and `replies` is `Comment[]`. `parent` is an extension.
- **`id` is the OOXML `w:id`, a number**, not Office JS's opaque string, because that is what
  docx4j users and the markers work with. `paraId` (the w15/w16cid key), `initials`,
  `paragraphs`, `commentBody` (a `Body` over the comment's own content) and `element` (the
  `w:comment`; the model holds comments as values, not `{ name, value }` pairs) are extensions.
- **`creationDate` is `Date | undefined`** (Office JS always has one); a comment without `w:date`
  is legal and `comments-two.docx`-style documents exist.
- `getRange()` returns a `Range` per paragraph, as section 3.8 allows. A comment whose markers
  sit at block level (a range spanning whole tables) gets no range: only markers inside a
  paragraph are mapped to offsets. An empty range gets a reference run only, which is all Word
  requires (docx4j's `CommentsSample` says the same).
- **Reads unmarshal three parts, writes five.** `getComments()` unmarshals `w:comments`,
  `w15:commentsEx` and `w:people`, which is what the views read; `w16cid` and `w16cex` are only
  touched when a comment is inserted, replied to or deleted. A document whose comments are not
  read keeps all five byte for byte (a test). The styles part is only unmarshalled when its XML
  does not already name both comment styles, so a document that has them keeps it byte for byte.
- **No `w16cex` part is created**, only kept in step when the document has one. It was edited
  as a DOM until objects 0.1.3 typed it; now `CommentsExtensiblePart` (section 14).
- Comment **reactions** (`cr:`, Word 2020) and the `w16cid` "comments ids" of *replies to
  replies* beyond what Word writes are out of scope, as are comments on headers and footers
  (the API accepts them, but Word does not write them).
- The author identity is `WordprocessingMLPackage.author: { name, initials?, email? }`, default
  `{ name: 'docx4j' }`, read by the parts layer without importing the package (a structural
  read, to keep the import graph acyclic). Phase F's tracked changes use the same setting;
  `trackedChangeAuthor` in 3.7 should become `author` when it lands.
- **Two registries keep the import graph one-directional** (CR-001's "no runtime import cycles"):
  `setCommentPartsAccess` (the parts layer registers how to load and create the comment parts) and
  `setCommentApi` (`Comment.mts` registers the two entry points `Body`, `Paragraph` and `Range`
  call). `WordprocessingMLPackage` imports `src/parts/wml/comments.mts` for its registration, as
  `packages/index.mts` imports the package classes for theirs.

Objects-package gaps found (candidates for a CR there; none blocked this phase):

All three closed in objects 0.1.3 (2026-09-16, section 14):

1. **No `org_docx4j_w16cex` module.** `w16cex:commentsExtensible` was not generated, so the part
   loaded as a `DefaultXmlPart` and phase G edited its DOM. Now typed; `CommentsExtensiblePart`.
2. **`mc:Ignorable` was missing from `w:comments`**, `CTCommentsEx` and `CTPeople`, so a
   re-marshalled comments part lost Word's `mc:Ignorable="w14 w15 ..."`. Now kept (a test).
3. **`w15:CTPerson.contact` was declared required** though Word omits it. Now optional.

## 10. Implementation notes: phase I (2026-09-15)

`src/office-js/` behind a new subpath `@docx4j/core-ts/office-js` (`exports` `./office-js`, with
`types` and `import` as the others; `test/nodenext/consumer.mts` imports it): `errors.mts`
(`NotSupportedError`, `ItemNotFoundError`, `ValueNotLoadedError`), `enums.mts`, `proxy.mts`
(`ClientResult`, the collections, `Wrapper`, `unwrap`, `nullObject`), `document.mts` (`Document`,
`DocumentProperties`), `run.mts` (`RequestContext`, `run`), `extras.mts` (members the shim adds
over a view), `fragment.mts`, `supported.mts` with the generated `supported.generated.mts`, and
`toApiScript.mts`, which is also exported from `./model` (and so from `.`). The shim itself is
**not** in `src/index.mts`: an add-in bundle that wants only the content API never pulls the
proxies in. Nine tests in `test/office-js.test.mjs`; the README has a "Use in Node" section.

What is in:

- **`Word.run(pkg, fn, options?)`**, the package first since there is no host document. It
  unmarshals the main document part (and, unless `unmarshalSideParts: false`, the core and
  extended properties and the settings parts) before the callback, calls it with a proxied
  `RequestContext`, and syncs once more at the end. `context.document.body` is the package's
  `Body`; `getSelection()` returns `options.selection`; `properties` is `DocPropsCorePart` and
  `DocPropsExtendedPart` in Office JS's vocabulary (`title`, `author` = `dc:creator`, `subject`,
  `keywords`, `comments` = `dc:description`, `category`, `lastAuthor`, `revisionNumber`,
  `creationDate`, `lastSaveTime`, `lastPrintDate`, `manager`, `company`, `applicationName`,
  `template`, `security`), creating the part on write; `changeTrackingMode` reads and writes
  `w:trackRevisions`, creating the settings part if absent (the tracking of edits is phase F);
  `getComments()` calls `Body.getComments()` if it exists at run time (phase G) and otherwise
  reports none.
- **`load` / `track` / `untrack`** are no-ops returning the object, on every proxy and on the
  collections and null objects, so an add-in's `load` lines stay. **`context.sync()`** resolves
  what was asynchronous: `getOoxml`, `getXml`, `insertXml` (and phase C's `insertOoxml`) return
  `ClientResult`-shaped objects whose `value` it fills. `Word.run` syncs once at the end, so a
  callback that forgets to is still correct.
- **Collections**: an array of wrapped views with `items`, `getFirst`, `getFirstOrNullObject`,
  `getLast`, `getLastOrNullObject`, `getCount` and the no-ops, so `body.paragraphs.items[0]`,
  `body.paragraphs[0]` and `[...body.paragraphs]` all work. A null object answers
  `isNullObject: true` and throws `ItemNotFoundError` from every other member.
- **Enums** as frozen objects: `InsertLocation`, `Alignment`, `UnderlineType`,
  `ChangeTrackingMode`, `BreakType`, `ContentControlType`, `Style`, `ErrorCodes`, and
  `SearchOptions` (the interface, plus `newObject()`).
- **Proxies**: every object handed out is a `Proxy` over the view. A member the view has passes
  through and its result is wrapped in turn (view, array of views, `ClientResult`); a member it
  does not have throws `NotSupportedError` (`Word.Paragraph.getNextOrNullObject is not supported
  by @docx4j/core-ts`, `code: 'NotImplemented'`), on reads, writes and calls. `unwrap(proxy)`
  gives the view back, and `part` and `package_` are handed out unwrapped so that docx4j code
  keeps working on them.
- **`Word.supported`**: a `Set` of `Class.member` strings. `scripts/generate-supported.mjs` reads
  `test/office-js-subset.ts` and writes `src/office-js/supported.generated.mts`; it is
  data-driven (every `interface` in the file, whatever it is called), so phases C and G extend
  the subset and re-run `npm run generate` with no change here. The generated file is committed,
  so the build does not depend on the script; `npm test` regenerates it (`pretest`) and
  `--check` fails when it is stale. `supported.mts` adds the shim's own surface (the context,
  the document, the collection members and the `load` no-ops) to the generated half.
- **`toApiScript(target, options?)`** over a `Body`, a `Paragraph`, a `Range`, a `w:p` / `w:tbl`
  element or an array of them: `insertParagraph` with the text of a leading unformatted run,
  then `style`, `alignment`, the indents and spacing in points and `outlineLevel`, then one
  `insertText` per run with the `Range.font` assignments that differ from the previous run, and
  `insertBreak` for `w:br`. It falls back to ``await body.insertXml(`...`, 'End')`` with the
  marshalled fragment, preceded by a comment saying why, for anything else — and the decision is
  taken before a line is emitted, so a paragraph comes out whole either way.

Departures and deferrals, all deliberate:

- **`Word.run` takes the package**, and `Word` is a plain object, not a namespace with the whole
  Office JS surface. `Word.run(fn)` without a package cannot mean anything here.
- **Calls act at once; `sync()` only resolves promises.** Office JS queues them. This is section
  3.1's decision, and the observable difference is that a mutation is visible before the sync.
- **`ClientResult` is awaitable** (`then` delegates to the underlying call), which Office JS's
  is not. Without it `await body.insertXml(...)` — how the content API and `toApiScript`'s output
  read — would not wait, and the insertions would land out of order. The Office JS idiom
  (`const r = body.getOoxml(); await context.sync(); r.value`) works unchanged, and `value`
  before the sync throws `ValueNotLoadedError`.
- **`getOoxml()` is the shim's, not the views'** (section 7 left it out). On a `Body` it is
  `OpcPackage.saveFlatOpc()`, which is exactly Word's whole-document `pkg:package`; on a
  `Paragraph` or a `Range` it is the bare `w:p` fragment, because wrapping a fragment in a
  package with the styles it needs is phase C's `insertOoxml` work. A `Range` also gains
  `getXml()`, which marshals a copy of its paragraph trimmed to the span (`fragment.mts`).
- **`document.properties` and `changeTrackingMode` read synchronously**, so `Word.run`
  unmarshals `docProps/core.xml`, `docProps/app.xml` and `word/settings.xml` when they are
  present, which marks them for re-marshalling even if the callback never reads them. Pass
  `{ unmarshalSideParts: false }` to keep those parts byte for byte; reads then report nothing
  and writes throw with that message. Everything else obeys CR-001's rule.
- **`changeTrackingMode` is `'Off'` or `'TrackAll'`**: `w:trackRevisions` has no "mine only".
  Writing `'TrackMineOnly'` sets `w:trackRevisions`, as Word does for a document it later filters
  by author.
- **`Word.Style` uses Office JS's values** (`'Heading1'`), which `styleBuiltIn`'s setter accepts;
  at the time its getter still returned the spaced form (`'Heading 1'`, section 7). Corrected the
  next day, section 11; `toApiScript` now emits `styleBuiltIn` for a built-in style and `style`
  with the display name otherwise.
- **`toApiScript` fidelity.** A `w:tab` is emitted as `insertText('\t')`, which writes a tab
  character rather than a `w:tab` (the text model reads both as `\t`); `w:proofErr`,
  `w:lastRenderedPageBreak`, `w:bookmarkStart` and `w:bookmarkEnd` are dropped; a `w:br` becomes
  a run of its own. Everything else that a verb cannot express falls back to `insertXml`: any
  `w:pPr` beyond `pStyle`, `jc` (left/center/right/both), `ind` (left/right/firstLine/hanging),
  `spacing` and `outlineLvl` — the paragraph mark's `w:rPr` included; any `w:rPr` outside the
  `Font` vocabulary, `w:rStyle` included, since `Font` has no `style` member; `w:spacing` with a
  `w:lineRule` other than `exact`, because `lineSpacing` writes `exact`; and a run that drops a
  direct size after one that set it, because `Font.size = 0` writes `<w:sz w:val="0"/>` rather
  than removing it. A `Range` emits its runs only, without its paragraph's properties.
  `insertTable` is emitted only when `Body.prototype.insertTable` exists (phase C) and the table
  is a plain grid of single-paragraph text cells; until then a table is `insertXml`.
- **The emitted script is content-API code, not shim code**: run it against a `Body` (or the
  shim's proxied body, since `ClientResult` is awaitable) inside an async function, which is what
  the test does with `new Function('body', 'return (async () => {' + script + '})()')`.

For the objects package's next CR (`toSource` in `builders/wml`): what this one wants from it is
the complement, not the same thing. `toApiScript` needs a way to say "this element is beyond the
verbs, emit it as source": `toSource(element)` giving the `el.p({ ... })` / `p([r('x')])` calls
would be a better fallback than a marshalled XML string for a reader, and the two generators
should share the shape of their output (statements, a variable per element, stable names). The
pieces it should expose: `toSource(element, { variable, factory: 'el' | 'sugar' })` returning
statements; a predicate `isSugarExpressible(element)` so a caller can choose between `p('x')` and
`el.p({...})`; and the run-formatting inverse of `applyRunOptions`, which `readRunOptions`
already is, so that `r('x', { bold: true })` can be emitted instead of an `rPr` literal. With
that, `toApiScript`'s fallback becomes `body.insertElement(<source>, 'End')` and the XML string
is only for what neither can type.

## 11. Correction: `style` and `styleBuiltIn` (2026-09-16)

Section 3.1 had `style` as the style id and `styleBuiltIn` as the id with spaces inserted, and
phases B, C and I implemented that. Office JS means the opposite: `style` is the style's display
name (`'Heading 1'`, localised in Word) and `styleBuiltIn` a `Word.Style` value (`'Heading1'`,
`'TableGrid'`), reading `'Other'` for a style that is not built in and refusing to be set to it.
The section 3.4 promise is that a function written against the subset runs against both, so
the values have to agree, not only the types; there was no reason for the departure beyond the
tree holding the id. Fixed in `Paragraph`, `Range`, `Table` and `toApiScript`:

- `style` reads the display name: the styles part's `w:name` when that part is unmarshalled
  (nothing is unmarshalled for it), with Word's stored lower-case built-in names (`heading 1`,
  `toc 1`, `annotation text`) mapped to the display names (`Heading 1`, `TOC 1`, `Comment Text`)
  through docx4j's `KnownStyles.xml` list; otherwise derived from the id by inserting spaces,
  which is exact for the built-ins because Word derives the id from the English name. Setting
  accepts a display name, a stored name or, as an extension, an id; resolved against the styles
  part when unmarshalled, else by removing spaces.
- `styleBuiltIn` reads the `Word.Style` value for the id (`Toc1` for Word's `TOC1`) or `'Other'`;
  setting takes a `Word.Style` value (leniently the spaced name too) and writes Word's id spelling;
  `'Other'` throws `RangeError`. `BUILT_IN_STYLES` in `src/model/content/styles.mts` is the list,
  and a test keeps `Word.Style` in `office-js/enums.mts` equal to it.
- `styleId` is the docx4j-named extension for the id (`w:pStyle`, `w:tblStyle`) on `Paragraph`,
  `Range` and `Table`, for code that thinks in ids; `outline()` keeps reporting ids.
- `toApiScript` emits `styleBuiltIn = 'Heading1'` for a built-in and `style = '<name>'` otherwise.

## 12. Implementation notes: phase E (2026-09-16)

Custom XML, XML mapping and the typed content controls, as sections 3.5 and 3.6 specify them.
A new directory `src/model/customxml/`, exported from `.` and `./model`: `xpath.mts` (the
`XPathEngine` interface, the default engine, the prefix-mapping strings and the canonical XPath of
a node), `CustomXmlPart.mts` (`CustomXmlPart`, `CustomXmlNode`,
`CustomXmlPrefixMappingCollection`), `CustomXmlPartCollection.mts` (`pkg.customXmlParts`, `add`,
`applyBindings`, `updateFromContentControls`), `XmlMapping.mts`, `kinds.mts` (the kind-specific
views) and `insert.mts` (the `w:sdt` builder and the id allocator behind `insertContentControl`).
`src/model/content/ContentControl.mts` gains the phase E members; `Body`, `Paragraph` and `Range`
gain `insertContentControl`; `WordprocessingMLPackage` gains `customXmlParts` and `xpathEngine`;
`DefaultXmlPart` gains `parseDocument` / `parsedDocument` / `markModified`. Eleven tests in
`test/customxml.test.mjs`, one new fixture (`invoice2013.docx`), `test/office-js-subset.ts`
extended with `CustomXmlPart`, `CustomXmlNode`, `XmlMapping`, the kind views, the collection and
the new `ContentControl` members, and `test/README.md` a manual Word check (9).

What is in:

- **The XPath engine** (section 3.6). `XPathEngine` is `select`, `selectValue`, `ready()` and
  `isReady`; `DefaultXPathEngine` uses `document.evaluate` when the runtime has it and otherwise
  `import('xpath')`. `xpath` is an **optional peer dependency** (open question 7, as recommended)
  and a devDependency here; the import specifier is a variable, so a bundler does not follow it and
  a consumer without the package still type-checks.
- **`pkg.customXmlParts`**: `load()`, `items`, `getByNamespace`, `getItem` (case-insensitive, with
  or without braces), `add`, `applyBindings()`, `updateFromContentControls()`. `CustomXmlPart` has
  `id`, `namespaceUri`, `builtIn`, `documentElement`, `namespaceManager`, `schemaCollection`,
  `getXml`, `setXml`, `selectNodes`, `selectSingleNode`, the four element and attribute methods,
  `delete()` and `part`; `CustomXmlNode` has the whole node model of section 3.5, including the
  canonical `xpath` and both forms of `appendChildNode` (open question 8, as recommended).
- **`XmlMapping`** on every content control, over `w:dataBinding`: `isMapped`, `xpath`,
  `prefixMappings`, `storeItemID`, `customXmlPart`, `customXmlNode`, `setMapping`,
  `setMappingByNode`, `delete`. `add()` follows docx4j's `AbstractMigrator.addPropertiesPart`:
  `/customXml/itemN.xml` with `/customXml/itemPropsN.xml`, a brace-wrapped upper-case random UUID
  as the `ds:itemID`, `ds:schemaRefs` from the document element's namespace, and the data part's
  relationship **from the main document part** (Word silently drops a custom XML part the main part
  does not relate to).
- **The typed kinds**: `checkboxContentControl` (w14 `CTSdtCheckbox`), `datePickerContentControl`,
  `dropDownListContentControl` and `comboBoxContentControl` (`listItems`, `addListItem`,
  `deleteAllListItems`), `pictureContentControl`, `repeatingSectionContentControl` (w15
  `CTSdtRepeatedSection`, with `insertItemAfter`), `groupContentControl`, plus `placeholderText`,
  `appearance`, `color`, `cannotDelete`, `cannotEdit`, `removeWhenEdited` and the `findProperty` /
  `putProperty` / `removeProperty` extensions they are built on.
- **`insertContentControl(kind?)`** on `Body` (wraps the body's content), `Paragraph` (wraps the
  paragraph) and `Range` (wraps the runs, splitting at the boundaries as `Range.font` does), with a
  `w:id` free in the document and a `w:sdtPr` typed for the kind.

The **asynchrony decision** (section 3's rule, "only what marshals or unmarshals is asynchronous"):
loading the `xpath` package is asynchronous, so the engine is **warmed once** — `await
pkg.customXmlParts.load()`, which parses every custom XML part's DOM and awaits
`pkg.xpathEngine.ready()` — and everything after it is synchronous, as Office JS is. Before that,
`items` (and so `getItem`, `getByNamespace`) throws `Custom XML part ... is not parsed yet: await
pkg.customXmlParts.load() once`, and an engine call throws `The XPath engine is not ready yet`. In
a browser or a Word add-in `document.evaluate` is native, so the engine is ready from the start and
only the DOM parse needs the await. `applyBindings` and `updateFromContentControls` are
asynchronous because they unmarshal the main document part (and the headers and footers).

Departures and deferrals, all deliberate:

- **A bound control's `insertText` writes through to the custom XML node.** The 2026-09-16 Word
  acceptance run (check 7) found that replacing a bound control's text does not show, because Word
  refreshes a bound control from the custom XML part when it opens the document. So
  `ContentControl.insertText` now makes the same write as `updateFromContentControls()` when the
  mapping resolves; when the XPath engine is not warm (nobody called `load()`) the write is skipped
  silently and the explicit `updateFromContentControls()` is the fix. Nothing else in the content
  API writes through: an edit made on the control's paragraph, or through `Range`, is a document
  edit, and `updateFromContentControls()` is how it reaches the data.
- **Containers are never bound.** docx4j binds anything that is not explicitly rich text, which for
  a `w15:repeatingSection` would replace the repeat with one run. `applyBindings` and
  `updateFromContentControls` skip `RepeatingSection`, `RepeatingSectionItem`, `Group` and
  `BuildingBlockGallery`, and any control that holds a table or another control: their children
  carry the values. Word's own repeat semantics (a section per node of a node set) are OpenDoPE's
  `OpenDoPEHandler` work and are not in this phase.
- **Pictures and explicit rich text are not bound** (either way). docx4j replaces the `a:blip`
  embed from base64 image data, and writes a whole flat OPC package into a rich-text binding; both
  are worth their own change request, and both are counted as `skipped` in the result rather than
  failing.
- **Dates are formatted, checkboxes and lists round trip by value.** `applyBindings` formats the
  bound value with `w:dateFormat` in the `w:lid` locale (`formatDate`, the .NET vocabulary docx4j
  maps to `SimpleDateFormat`) and sets `w:fullDate`; a checkbox gets `w14:checked` and Word's ☒/☐
  glyph run in MS Gothic, taking the glyphs from `w14:checkedState` / `uncheckedState` when the
  control has them (docx4j ignores those, an explicit TODO there); a list shows the `displayText`
  of the entry whose `value` matches and records `w:lastValue`. The reverse writes `true`/`false`
  for a checkbox, the entry's `value` for a list and the stored `w:fullDate` form for a date —
  docx4j writes the rendered glyph and skips dates entirely, which loses the data's meaning.
- **`applyBindings` keeps docx4j's shapes otherwise**: the value is trimmed; the run properties
  come from `w:sdtPr/w:rPr`, not from the existing runs; an empty value gives docx4j's placeholder
  run (`w:rStyle` `PlaceholderText`, "Click here to enter text.") and sets `w:showingPlcHdr`;
  `xml:space="preserve"` only for a leading or trailing space; a multiline `w:text` control turns
  newlines into `w:br` runs and a single-line one drops them. It differs in two places: the content
  goes into the **first paragraph** of a block, row or cell control, keeping its `w:pPr` and the
  rest of the structure, rather than reducing a `w:tr` or `w:tbl` to one cell as docx4j does; and
  `w:placeholder` is **kept** (docx4j always removes it), since it names the glossary part Word
  shows the placeholder from.
- **Bindings are read from `w:dataBinding` only.** docx4j prefers the OpenDoPE XPaths part
  (`od:xpath=x1` in `w:tag`, resolved through `/customXml/itemN.xml`); here the control's own
  `w:dataBinding` is the mapping, which is what Word writes and what Office JS's `XmlMapping`
  reports. The invoice fixture's three bindings resolve either way. OpenDoPE's conditions, repeats
  and `od:Handler` are a separate piece of work.
- **The three well-known docProps store item ids** (`{6C3C8BC8-...}` core, `{6668398D-...}` app,
  `{55AF091B-...}` cover page) are not special-cased as docx4j does: a binding to them resolves
  only if the package really has that custom XML part.
- **`getXml()` is synchronous** and returns the string, where Office JS returns a `ClientResult`: a
  custom XML part is a DOM here, so nothing is marshalled. The subset file declares it that way.
- **The XPath-addressed mutators take `namespaceMappings` last** (`insertElement(xpath, xml,
  namespaceMappings?, index?)`, `insertAttribute(xpath, name, value, namespaceMappings?)`), as
  section 3.5 specifies, where Office JS's desktop-only forms put it second. They are therefore
  left out of `test/office-js-subset.ts` (the section 3.4 promise), with a comment saying why;
  everything else in the custom XML model is in it and asserted assignable.
- **`CustomXmlPart.delete()` unlinks the mappings it can reach**: the controls of the parts that
  are already unmarshalled. A document whose main part was never read keeps its `w:dataBinding`
  elements, which then name a part that is gone — Word treats those as unbound.
- **`placeholderText`** reads the control's own text while it is showing its placeholder and writes
  by replacing the content and setting `w:showingPlcHdr`; it refuses when the control holds
  content, so that a value is never lost. Word keeps the placeholder text in a glossary document
  part (`w:placeholder/w:docPart`); creating glossary parts is not in this phase.
- **Reading a custom XML part no longer costs its byte-for-byte round trip.** `DefaultXmlPart`
  separates "parsed for reading" from "adopted as the part's content": `parseDocument()` parses,
  `parsedDocument` is the synchronous accessor the views read, and `markModified()` (which every
  mutation of a node view calls) is what makes the part re-marshal. `getDocument()` keeps its old
  meaning — parse and adopt — so phase G's w16cex editing is unchanged. A test loads the invoice,
  reads every custom XML part and asserts the four parts come back byte-identical.
- **`insertContentControl` at range level refuses a span that crosses run holders** (a hyperlink, a
  tracked insertion), naming the reason, rather than producing a `w:sdt` that spans a boundary Word
  would reject; and `RepeatingSection` is refused at run level. A `Range` whose start equals its end
  gets an empty control at that position.
- **Nothing was added to the `Word` shim.** `context.document.customXmlParts` would mean editing
  `src/office-js/document.mts`, which phase F is changing in parallel; the collection is reachable
  as `unwrap(context.document).package_.customXmlParts` in the meantime, and wiring it is a small
  follow-up.
- **No fixture in the docx4j checkout has a drop-down, combo box or group control**, so those kinds
  are exercised on documents this package builds (`insertContentControl` then the kind view);
  `invoice2013.docx` covers the checkbox, the date, the picture and the two `w15` repeating
  sections on a document Word wrote.

Objects-package gaps found (candidates for a CR there; none blocked this phase):

1. **No builder for `w:sdt`** (closed in objects 0.1.4, section 15: `sdt`, `sdtPr`, `nextSdtId`).
   `builders/wml` had `p`, `r`, `t` and `tbl` but nothing for a content
   control, so `insert.mts` wrote the four `w:sdt` forms and their `w:sdtPr` over the generated
   factories here. `sdt(content, { kind, tag, title, id })` (and the `sdtPr` half) needs only the
   tree, so it belongs there next to `tbl`, as phase C said of `drawingFor`.
2. **`SdtPr`'s properties are a choice list** (closed in objects 0.1.4, section 15: `sdtProperty`
   and `sdtKindOf`). `rPrOrAliasOrLock` is a union of
   `TypedNamedValue<...>` typed element by element (it matches docx4j's own choice list, the objects
   session notes), so reading `w:tag` or `w14:checkbox` means a search by local name and a cast
   (`findProperty` here). The helpers of the objects package's CR-003 (`sdtProperty`, `sdtKindOf`)
   replaced them; not a compiler change.
3. **`w14:checkbox`'s `checked` was `CTOnOff` with `val?: string`** (closed in objects 0.1.3,
   section 14: docx4j CR-018 retyped it to `xsd:boolean`, so `val` is a boolean and `w14:val="1"`
   in a document unmarshals to `true`).
4. **No `xs:anyAttribute` escape on `DatastoreItem`** (closed in objects 0.1.3: `otherAttributes`).
   The properties part is still read from its DOM here, since nothing needs it typed.

## 13. Implementation notes: phase F (2026-09-16)

Change tracking and `replaceText` (section 3.7). New files
`src/model/content/tracking.mts` (the `ChangeTracker`, the markup writers and the
`w:rPr` / `w:rPrChange` conversions) and `src/model/content/TrackedChange.mts` (the views,
accept and reject); the mutation paths of `Paragraph`, `Range`, `Font`, `Body` and
`tree.mts` grew a tracked branch; `Table.mts`'s row mutations and `Comment.mts`'s marker
placement were adapted; `WordprocessingMLPackage` gained `trackedChangeDate`,
`changeTrackingMode`, `changeTracker` and `getTrackedChanges()`, and the shim's
`document.changeTrackingMode` now delegates to the package. Twenty tests in
`test/tracking.test.mjs`, one new fixture (`test/fixtures/tracked-changes.docx`), and
`test/office-js-subset.ts` now also asserts `TrackedChange`, `getTrackedChanges()` on
`Body`, `Paragraph`, `Range` and `Document`, and `changeTrackingMode` on `Document`
(`npm run generate` has rewritten `supported.generated.mts` accordingly).

What is in:

- **`pkg.changeTrackingMode`** (`Off` | `TrackAll` | `TrackMineOnly`) over `w:trackRevisions`
  in the settings part, which is created when the document has none. The revision author is
  phase G's `pkg.author` (the `Author` of section 9; only its `name` reaches `w:author`), so a
  comment and a revision made in one session carry one identity; `pkg.trackedChangeDate` is an
  optional fixed date. `pkg.changeTracker` is the tracker every mutation asks its `Body` for,
  and `pkg.getTrackedChanges()` is Office JS's `document.getTrackedChanges()` over the body.
  The shim's `document.changeTrackingMode` (phase I) is now a delegation to the package, with
  nothing of its own but the "settings part not unmarshalled" guard `Word.run` needs.
- **Every mutation writes revision markup when the mode is on**, in the paragraph-level
  primitives, so `Range`, `Body` and later phases inherit it: `Paragraph.splice` (behind
  `insertText`, the `text` setter, `Range.insertText` and `Range.delete`) deletes as a
  `w:del` and inserts as a `w:ins`; `Body.insertElement` (behind both `insertParagraph`s and
  `insertXml`) marks an inserted paragraph's mark and wraps its runs, and marks every row and
  paragraph of an inserted table; `Paragraph.delete()` moves the content into a `w:del` and
  marks the mark deleted; `Font` and the paragraph property setters record `w:rPrChange` and
  `w:pPrChange`; `Table.addRows`, `TableRow.insertRows`, `Table.deleteRows`, `TableRow.delete`
  and `Table.delete` go through `ChangeTracker.markRowInserted` / `markRowDeleted`, which write
  `w:trPr/w:ins` and `w:trPr/w:del` (a deleted row stays in the tree until the change is
  accepted) **and, since the acceptance run of 2026-09-19, the row's content too** — see the
  row-revisions note at the end of this section. `Table` and `ContentControl` needed no tracked
  branch of their own beyond the rows: their text edits already run through `Paragraph.splice`
  and `Body.insertText`.
- **Word's rules**, not just the markup: a run already inside a `w:ins` by the same author is
  extended rather than nested in another one; deleting text that author had inserted takes it
  back instead of nesting a `w:del`; a replacement writes the `w:del` first and the `w:ins`
  after it; `w:t` becomes `w:delText` and `w:instrText` becomes `w:delInstrText`; runs are
  split at the boundaries of the span, as `Range.font` already did, so that a partly deleted
  run is not wholly deleted.
- **`TrackedChange`** over `w:ins`, `w:del`, `w:moveFrom`, `w:moveTo`, `w:rPrChange`,
  `w:pPrChange`, the two paragraph-mark forms and the two row forms: `type`, `author`,
  `date`, `text`, `accept()`, `reject()`, `getRange()`, and `id`, `element` and `target` as
  extensions. `getTrackedChanges()` on `Body`, `Paragraph` and `Range`; `acceptAll()` and
  `rejectAll()` on `Body`. Accepting follows docx4j's
  `org.docx4j.convert.out.common.preprocess.AcceptTrackedChanges`: a `w:ins` is unwrapped, a
  `w:del` removed, a deleted paragraph mark joins its paragraph with the next (the joined
  paragraph keeps the first one's content and `w14:paraId` and takes the second one's
  properties), a deleted row removed.
- **`{ view: 'original' }`** on `Body.getText()` and `Paragraph.getText()`: the tree with
  `w:del` and `w:moveFrom` counted and `w:ins` and `w:moveTo` skipped. `segmentsOf` and
  `runsOf` in `tree.mts` take the option, and a segment now reports the `w:ins` / `w:del` it
  is inside (`TextSegment.revision`), which is what the tracked primitives work from.
- **`replaceText(find, replace, options?)`** on `Body`, `Paragraph` and `Range`: `search`
  then `insertText(replace, 'Replace')` from the last match to the first, returning the
  count, tracked or not.

Departures and deferrals, all deliberate:

- **Revision ids and comment ids are separate spaces, deliberately.** Revision ids come from a
  per-package counter one above the highest `w:id` on a `CTMarkup` in the parts already
  unmarshalled — the one annotation id space of ECMA-376 17.13.5.4, which revisions, bookmarks,
  comment range marks and permissions share. `w:comment/@w:id` is **excluded** from the scan
  (`MARKUP_TYPES` in `tracking.mts` leaves `org_docx4j_wml.Comments.Comment` out), because it
  is a different space: phase G keeps its own allocator (`nextCommentId` in `comments.mts`,
  which counts `w:comment` ids and the markers in the body) and the two never have to agree.
  They do share `pkg.author`.
- **A comment is not a revision.** `Range.insertComment` writes its markers and its reference
  run straight into the run arrays, never through `splice`, so it is not tracked as a text
  insertion — which is what Word does too. `placeAround` in `Comment.mts` now also hoists the
  markers out of a `w:ins` or `w:del` the anchor run sits in (`markerSite`), so that the
  reference run is never itself inside an insertion and accepting or rejecting the revision
  leaves the comment where it is. The cost is that a comment on part of an insertion widens to
  the whole of it, which is what the document shows once the insertion is accepted anyway.
- **`comments.mts`'s marker walk had to learn the revisions.** `RUN_HOLDERS` in `tree.mts` no
  longer contains `RunIns` and `RunTrackChange` (they are the new `REVISION_HOLDERS`, which
  `segmentsOf` treats by view), so the walk that finds comment markers and their offsets now
  descends into a `w:ins` or `w:moveTo` explicitly — the accepted view, which is the view the
  offsets are in.
- **Deleting a content control is not tracked.** `ContentControl.delete()` removes the `w:sdt`
  outright even while tracking is on: its content is not text, and Word records such a change
  as the deletion of what the control held rather than of the control. Its `insertText`,
  `insertParagraph` and `search` all run through the paragraph primitives and are tracked.
- **`changeTrackingMode` reads and writes synchronously**, so on a loaded package whose
  settings part has not been unmarshalled the getter reports `Off` until
  `await pkg.getChangeTrackingMode()`, and the setter defers the `w:trackRevisions` write to
  `saveTo` (it is applied there and then when the part is already unmarshalled, or absent).
  This keeps the CR-001 rule that the settings part is unmarshalled only when the mode is
  read or written. The async pair `getChangeTrackingMode()` / `setChangeTrackingMode()`
  mirrors `body` / `getBody()`.
- **`TrackMineOnly` is stored as `TrackAll`.** `w:trackRevisions` is a flag, so a file cannot
  tell the two apart; a document that has it on reads back as `TrackAll`. The distinction
  only matters once a second author edits the same package, which is Office JS's concern and
  not this package's.
- **An inserted paragraph carries its own mark.** Word, splitting a paragraph, marks the
  *first* paragraph's mark as inserted and leaves the new paragraph the original mark;
  `insertParagraph` marks the mark of the paragraph it added. The accepted and the rejected
  document are the same either way, and the edit stays on the element that was added. The
  one place it shows is rejecting an insertion at the end of a container, where there is no
  paragraph after it to join with: `joinWithNext(paragraph, true)` then removes the paragraph
  into the one before it instead.
- **`Range.getText({ view: 'original' })` throws**, because a range's offsets are
  accepted-view offsets; read the original view on the `Paragraph` or the `Body`. For the
  same reason `search` has no `view` option: a `Range` it returned could not be edited.
  A `Body`'s original view is text-level, not structural — a paragraph inserted whole still
  contributes its (empty) line.
- **`Body.text` still shows a paragraph whose mark is deleted** as its own line, since the
  join only happens on accept. Word shows it the same way while the change is pending.
- **A deleted row stays in the tree.** `Table.deleteRows`, `TableRow.delete` and
  `Table.delete` mark rows rather than removing them while tracking is on, so `rowCount` and
  `values` still report them until the change is accepted, exactly as Word shows them. That is
  the one place where a tracked call's return value differs from the untracked one.
- **`style` is set through `styleId`, and both record `w:pPrChange`.** The section 11
  correction made `style` the display name over `styleId`; the tracking hook sits in the one
  private `pPr()` accessor every paragraph-property setter goes through, so `style`,
  `styleBuiltIn`, `styleId`, `alignment`, the indents, the spacing and `outlineLevel` all
  record the original once, on the first write.
- **`Body.clear()` marks everything deleted** while tracking is on (every row's `w:trPr` takes
  a `w:del` and every paragraph is deleted), which is also what `insertText(text, 'Replace')`
  does before adding its paragraph. Accepting that leaves the empty paragraph whose mark could
  not join with a following one (the one before a table, say) — docx4j's
  `AcceptTrackedChanges` leaves the same, since a mark only ever joins with the paragraph
  after it.
- **Editing deleted text is refused** by `ChangeTracker.assertEditable`, which throws naming
  the author. In practice it is unreachable through the views, because the accepted text
  model never surfaces a `w:del`; it guards the primitives for callers that build their own
  segments. Deleting a paragraph whose mark is already deleted throws too.
- **`acceptAll()` and `rejectAll()` are one pass in reverse document order**, not docx4j's
  recursive rewrite: a paragraph join then never disturbs a change still to do. Unlike
  docx4j's `AcceptTrackedChanges`, which is a conversion preprocessor and leaves formatting
  revisions alone (the current properties are what the document shows), accepting here also
  drops `w:rPrChange` and `w:pPrChange`, which is what accepting means for a document that
  is saved again.

### Row revisions: the whole row, not just its `w:trPr` (acceptance run of 2026-09-19)

Word showed a tracked row **deletion** with nothing struck out (`check10.docx`, finding 2 of
`test/README.md`'s acceptance record). Word writes a deleted row as `w:trPr/w:del` **and** the
cells' content deleted with it: every run in a `w:del` with its `w:t` turned into `w:delText`,
and every paragraph mark carrying `w:pPr/w:rPr/w:del`. Only the `w:trPr` was written, so the row
was marked but its text still read as present, and Word had nothing to strike through. An
inserted row already did the symmetric thing (`markRowsInserted` marks each paragraph and wraps
its runs in `w:ins`), which is what made the asymmetry visible.

- **`markRowsDeleted` in `Table.mts`** is now the mirror of `markRowsInserted`: `markRowDeleted`
  for the `w:trPr`, then `Paragraph.markDeletedInPlace(tracker)` on every paragraph of every
  cell, and itself again on the rows of a nested table. `markDeletedInPlace` is the same pair of
  primitives `Paragraph.delete()` uses — `deleteText(tracker, 0, length)` and
  `markParagraphDeleted` — without its "take back my own insertion" step, which would remove the
  paragraph element and leave a `w:tc` with no block-level child.
- **One change per row, as Office JS reports it.** `Body.getTrackedChanges` now collects the
  revision markup inside a row that is itself a revision into that row's `TrackedChange`
  (`target.inner`) instead of listing it separately: a deleted row is one `Deleted` row change,
  an inserted row one `Added` row change, whatever their cells carry. Decided this way because
  Office JS's `TrackedChange` for a row is the row, because listing a dozen inner changes for one
  user action reads as a dozen changes in any UI over this API, and because it keeps the
  insertion and deletion sides symmetric (the inserted-row case was already listing its inner
  `w:ins`). A row carrying no revision of its own still reports its cells' changes as before.
- **Accept and reject apply the inner markup with the row**: accepting a deleted row removes it
  and everything in it; rejecting one drops the `w:trPr/w:del` and rejects the inner changes, so
  the text comes back as `w:t` and the marks lose their `w:del`; accepting an inserted row drops
  the `w:trPr/w:ins` and accepts the inner insertions (otherwise the row would keep a `w:ins`
  around every run); rejecting one removes the row. Both directions are applied last-first, as
  `acceptAll` is.
- **`TrackedChange.text` of a deleted row** reads the original view (`rowText(tr, 'original')`,
  a line per paragraph): the builders' `textOf` is the accepted view, in which a deleted row's
  content, now that it is `w:delText`, is empty.

Objects-package gaps found (candidates for its own CR, worked around here):

- `w:rPrChange/w:rPr` is `CTRPrChange.RPr`, an untyped `egrPrBase` element list, while `w:rPr`
  is `RPr` with named properties, and nothing converts between them. `rPrElements` and
  `rPrFromElements` in `tracking.mts` do it by name through `el/org_docx4j_wml`, in the
  schema's EG_RPrBase order. The w14 effects (`w14:glow`, `w14:shadow`, `w14:reflection`,
  `w14:textOutline`, `w14:textFill`, `w14:scene3D`, `w14:props3D`, `w14:ligatures`,
  `w14:numForm`, `w14:numSpacing`, `w14:stylisticSets`, `w14:cntxtAlts`) are dropped from a
  recorded original by `rPrElements` here. Correction (2026-09-16): their scoped wrappers do
  exist, in `factory/org_docx4j_wml` rather than `el`; the objects package's CR-003
  (`rPrToElements` / `rPrFromElements`) covers them and has replaced the copy here: closed in
  objects 0.1.4, section 15, the w14 effects included.
- `runItemsOf` had to learn `accOrBarOrBox`, the name docx4j gives `w:moveFrom` and
  `w:moveTo`'s run list (`RunTrackChange`), alongside `customXmlOrSmartTagOrSdt`. Closed in
  objects 0.1.3: `builders/wml` exports a structural `runItemsOf` with the three names and the
  run-level `sdtContent` case, and `tree.mts` re-exports it (section 14).
- `deepCopy` of a `w:pPr` into `w:pPrChange` needs the copy's `TYPE_NAME` changed to
  `org_docx4j_wml.PPrBase`, or the marshaller writes `xsi:type="w:CT_PPr"` on it (valid but
  not what Word writes). A `deepCopyAs(value, typeName)` would say this plainly. Closed in
  objects 0.1.4, section 15: `deepCopyAsSync(pPr, 'org_docx4j_wml.PPrBase')`.

## 14. Upgrade to objects 0.1.3 (2026-09-16)

`@docx4j/generated-objects-ts` 0.1.3 closes five of the gaps sections 8, 9, 12 and 13 recorded,
and this package now depends on `^0.1.3`:

- `w14:CTOnOff/@val` is a boolean (docx4j CR-018): `kinds.mts`'s `isOn` and the two writes, and
  `insert.mts`'s checkbox init, take booleans; `w14:val="1"` in a document unmarshals to `true`.
- `org_docx4j_w16cex` exists: `CommentsExtensiblePart` (`/word/commentsExtensible.xml`, typed
  `CTCommentsExtensible`, registered by content type, a `DocumentPart` shortcut) replaces the DOM
  code in `parts/wml/comments.mts`; `dateUtc` is an `XmlCalendar`, written with `dateToCalendar`.
  The marshalled attribute order is the model's (`dateUtc` before `durableId`), which Word does not
  mind; the comments test asserts the typed form. Still no part is created when absent.
- `mc:Ignorable` on `w:comments`, `w15:commentsEx` and `w:people` survives a re-marshal (a test;
  `test/README.md` item 8 no longer warns about it).
- `wp:docPr/@title` is typed, so `altTextTitle` round-trips; the cast is gone.
- `runItemsOf` is imported from `builders/wml` (structural: the holder's list under `content`,
  `customXmlOrSmartTagOrSdt`, `accOrBarOrBox` or `sdtContent`), and the local copy is gone.
- `w15:CTPerson.contact` is optional and `DatastoreItem` has `otherAttributes`; nothing here
  changed for either.

Still open, now the objects package's CR-003 phase A in this order: `sdt`/`sdtPr`/`nextSdtId`/
`sdtProperty`/`sdtKindOf`; `tr`/`tc` and `inlinePicture`; `rPrToElements`/`rPrFromElements` and
`deepCopyAs`; `walkAll`. Phase B (`toSource`, `isSugarExpressible`) is unscheduled.

## 15. Upgrade to objects 0.1.4 (2026-09-16)

`@docx4j/generated-objects-ts` 0.1.4 is the objects package's CR-003 phase A: the builders that
were written here first. This package now depends on `^0.1.4` and the seven copies are gone —
`insert.mts` is 27 lines instead of 95, and nothing in `src/model/` builds a `w:sdt`, a `w:tr`,
a `w:tc` or a `wp:inline` by hand any more:

1. **`kindElement`, `sdtPrFor`, `sdtBlockFor`, `sdtRunFor`, `nextControlId`, `checkKind`**
   (`customxml/insert.mts`) are `sdt(content, { kind, id, form })`, `sdtPr(options)` and
   `nextSdtId(root)`. `Body`, `Paragraph` and `Range` pass `form` explicitly ('block', 'block',
   'run'), so the builder's inference from the content never decides here. What is left in
   `insert.mts` is what needs a part or a view: `sdtKindFor` (Office JS's `Unknown` has no kind
   element, as `RichText` has none) and `controlIdScope`, which gives `nextSdtId` the **part's
   whole contents** rather than the body container, so that a control inserted in a cell, a
   header or another control still gets an id free in the part.
2. **`findProperty` and `TYPE_BY_ELEMENT`** (`ContentControl.mts`) are `sdtProperty(sdtPr,
   localPart, namespaceURI = W)` and `sdtKindOf(sdtPr)`; `XmlMapping.mts` and `bindings.mts`
   call the builder too. `ContentControlType` is now `SdtKind | 'Unknown'`, the same union
   spelled once.
3. **`rowElement`'s element building** (`Table.mts`) is `tr(cells, { widths })` over `tc`; the
   two-line adapter that turns `(widths, values)` into the builder's arguments stays, since
   `addRows` and `insertRows` call it per row.
4. **`drawingFor`** (`InlinePicture.mts`) is `inlinePicture(relId, { cx, cy, id, name, descr,
   title })`, and the dml/pic factory imports with it.
5. **`rPrElements` / `rPrFromElements`** (`tracking.mts`) are the builders' `rPrToElements` /
   `rPrFromElements`; `TrackedChange.mts` imports the inverse from `builders/wml` and
   `model/content/index.mts` re-exports both under the builders' names.
6. **The `deepCopy` plus `TYPE_NAME` override in `recordPPrChange`** is
   `deepCopyAsSync(pPr, 'org_docx4j_wml.PPrBase')`, which copies, types the copy as the base and
   drops what `CT_PPrBase` does not declare, so the three `delete`s are gone as well.
7. **`visitReferences`'s traversal** (`ooxml.mts`) is `walkAll(root, visitor, domVisitor)`. Only
   the walking moved: the `REL_ATTRIBUTES` rule for typed objects, the relationships-namespace
   rule for DOM attributes and the "only when the incoming package has that relationship id"
   guard are unchanged.

Observable changes:

- **A repeating section in the run form throws the builder's message**: `Range.insertContentControl('RepeatingSection')`
  now fails with `Error: A repeating section is a block-level control; wrap paragraphs or a table,
  or pass form` instead of a `Docx4JException` saying "insert it on a body or a paragraph". No
  test asserted the old text. `Range.insertContentControl` builds the empty control **before** it
  splits any run, so the refusal still leaves the paragraph untouched.
- **`ContentControl.findProperty` is gone** from the public surface (`sdtProperty(control.sdt.sdtPr, name)`
  replaces it, and is re-exported from `.`, `./model`). Its namespace argument was optional and
  matched any namespace; `sdtProperty` defaults to wml. That is a real difference in one place:
  Word writes **`w15:dataBinding`** on a repeating section and on a rich-text control bound to a
  container (three of the twenty bindings in `invoice2013.docx`), so `XmlMapping.dataBinding`
  asks for both namespaces. The customxml test caught it.
- **A recorded `w:rPrChange` now keeps the w14 run effects** (`w14:glow`, `w14:textFill`, the
  other ten): the builders go through the scoped `createCTRPrChangeRPr*` wrappers, so the
  section 13 note about dropped effects is history.
- **`w:pPrChange/w:pPr` is asserted, not just described**: a new test in `tracking.test.mjs`
  records a change on a paragraph whose `w:pPr` has a style and the inserted mark's `w:rPr`, and
  asserts the original carries the `w:pStyle`, no `w:rPr`, no `w:sectPr` and no `xsi:type`.
- **`InlinePictureOptions` gained `altTextTitle`**, which `addImage` passes to the builder's
  `title` (`wp:docPr/@title`); it is not written when absent, as before. `descr` is still written
  as `""` when no description is given, so existing output is byte for byte what it was.
- Exports: `drawingFor` and `rPrElements` are no longer exported from `./model`; `nextControlId`,
  `sdtPrFor`, `sdtBlockFor`, `sdtRunFor` and `checkKind` are no longer exported from the custom
  XML index. In their place `./model` re-exports the builders `sdt`, `sdtPr`, `nextSdtId`,
  `sdtProperty`, `sdtKindOf`, `tr`, `tc`, `inlinePicture`, `rPrToElements`, `rPrFromElements` and
  `walkAll`, and `deepCopyAs` / `deepCopyAsSync` / `getContextSync` come through the facade
  re-export in `src/index.mts`.

Nothing was kept as a copy. CR-003 phase B (`toSource`, `isSugarExpressible`, for
`toApiScript`'s XML fallback) is still unscheduled there.

## 16. Effective formatting (CR-001 Phase B step 2, 2026-09-19)

`PropertyResolver` landed (CR-001 section 15.1), so the reads section 7 said report direct
formatting now report the formatting that actually applies - the document defaults, the style
chain and the numbering level resolved - which is what Office JS reports and what section 7
promised. Writes are unchanged: they are always direct formatting on the element itself.

Which reads changed meaning:

- **`Font`**: `bold`, `italic`, `underline`, `strikeThrough`, `doubleStrikeThrough`,
  `subscript`, `superscript`, `name`, `size`, `color`, `highlightColor`, on
  `paragraph.font` and `range.font`. `size` is now never 0 for a loaded document (the document
  defaults always state one, and docx4j's 10 pt stands in where they do not); `name` is still
  `''` for a run whose face comes from the theme, until CR-001 Phase B step 4 resolves theme
  references.
- **`Paragraph`**: `alignment`, `leftIndent`, `rightIndent`, `firstLineIndent`, `spaceBefore`,
  `spaceAfter`, `lineSpacing`, `outlineLevel`. A paragraph in a heading style now reports that
  style's `outlineLevel` and spacing rather than 10 and 0.
- **`Alignment` loses `'Unknown'`**: an absent `w:jc` resolves to `'Left'`, as Word lays it out.
  The setter still accepts `'Unknown'` (it removes the element) and the direct read still
  returns it; the union for both is `AlignmentOrUnknown`.

How to get the direct values, unchanged from what these reads used to answer:

```ts
paragraph.getFont({ direct: true }).size;          // 0 when the run states no w:sz
range.getFont({ direct: true }).bold;
paragraph.formatting({ direct: true });            // the eight paragraph properties, direct
paragraph.formatting();                            // ... and effective, in one call
paragraph.effectivePPr;                            // the resolved w:pPr itself
paragraph.effectiveParagraphMarkRPr;               // the mark's resolved w:rPr
```

`toApiScript` still reads the tree directly (a script reproduces markup, not appearance), and
says so at `paragraphProperties`.

A read needs the package's resolver, which is built by `await pkg.getPropertyResolver()` and by
`getBody()`, `paragraphAt()` and `outline()`, which await it; `createPackage()` builds one
before it returns. A `Body` with no package, or one whose package has not built a resolver,
throws `PropertyResolverNotCreatedException` naming `getPropertyResolver()` rather than
silently falling back to direct formatting. This changes the values existing callers see, so it
is a minor-version note.

## 17. Implementation notes: phase H (2026-09-19)

Lists (section 3.9). `src/model/content/List.mts` (`List`, `ListItem` and the functions
`Paragraph` and `Body` call), the members `isListItem`, `list`, `listItem`,
`listOrNullObject`, `listItemOrNullObject`, `startNewList`, `attachToList`, `detachFromList`
and `restartList` on `Paragraph`, `lists` and `listLabels()` on `Body`,
`NumberingDefinitionsPart.makeLive()`, the three enums in the shim, the subset interfaces and
`test/lists.test.mjs` (16 tests) are in. **The labels are docx4j's**: for every paragraph of
every story of all 45 parity goldens - 618 paragraphs, 183 of them numbered -
`paragraph.listItem?.listString` is the `numString` docx4j recorded and `listItem.level` its
`ilvl`, with zero differences.

**What is in.**

- `List` is a view over a `w:numId` and the body it was found through; it resolves the `w:num`
  and the `w:abstractNum` through the numbering part's definitions on every call, so nothing
  goes stale. `id`, `levelExistences`, `levelTypes` (`Bullet` / `Number` / `Picture`, from
  `w:numFmt` and `w:lvlPicBulletId`), `levelExists(level)`, `getLevelString`, `getLevelFont`,
  `getLevelParagraphs`, `paragraphs`, `insertParagraph`, `setLevelNumbering`, `setLevelBullet`,
  `setLevelIndents`, `setLevelAlignment`, `setLevelStartingNumber`, `restart()`, `separate()`,
  and `element` / `abstractElement` / `definition` / `level(n)` as the extensions on to the tree.
- `ListItem` is a view over a paragraph: `level` (read and write), `listString`, `siblingIndex`,
  `getAncestor(parentOnly?)`, `getDescendants(directChildrenOnly?)`, `list`, `paragraph`.
- `isListItem` is `Emulator.numRefFor` (CR-001 section 15.2), so a "List Number" style counts,
  a `w:numId` of 0 does not, and neither does a dangling `w:numId` - all three tested.
- Reads never unmarshal `word/numbering.xml`: they go through the definitions the
  `PropertyResolver` built from a private read. A document whose lists are only read still
  saves the numbering part byte for byte (tested on `numbering-lvlrestart.docx`).

**Departures, and why.**

1. **`startNewList()` is asynchronous** where Office JS's is synchronous. It copies its
   `w:abstractNum` from docx4j's default definitions, which are XML that has to be unmarshalled
   (and a document with no numbering part gets one, whose default contents are unmarshalled
   too). The subset declares `startNewList(): Promise<List>`, which is the convention section
   3.4 already uses for the calls that marshal; `await p.startNewList()` compiles against a Word
   add-in's objects as well. Every other list verb is synchronous, as Office JS has them.
2. **The first write promotes the definitions' tree**, rather than unmarshalling the part a
   second time: `NumberingDefinitionsPart.makeLive()` makes the tree the definitions were built
   over the part's live contents (`setContents`). Unmarshalling afresh would give a second tree
   and leave every `List` view pointing at the first. The part is then re-marshalled on save,
   which is the price of a write and only of a write.
3. **`separate()` is the copy-on-shared rule, and it is public.** A `setLevel*` call copies the
   `w:abstractNum` when another `w:num` names it, gives the copy a fresh `w:abstractNumId` and
   `w:nsid`, drops its `w:styleLink` (which maps a numbering style to one definition), and
   repoints this `w:num` at it. Offering it as a verb lets a caller do the split once before a
   run of writes, and lets `getLevelFont` say what it does.
4. **`getLevelFont(level)` separates first**, so reading it costs the numbering part its
   byte-for-byte round trip. `Font` is a read *and* write view over one `w:rPr`, and the
   supplier it takes cannot tell the two apart; Office JS pairs `getLevelFont` with
   `resetLevelFont`, so it is a write accessor. A read that costs nothing is
   `list.level(n).labelRPr` (or `.rPr`) on the definition.
5. **The restart verb is `List.restart()`**, with `Paragraph.restartList()` as the sugar that
   also attaches the paragraph. `startNewList({ restartFrom })` was the alternative; a restart
   is not a new list (it is the same `w:abstractNum`, which is what makes Word continue the
   numbering), so naming it `startNewList` would have said the wrong thing. It writes a
   `w:lvlOverride` with a `w:startOverride` for **level 0** only: `incrementCounter` resets the
   deeper levels whenever level 0 is used, so one override is enough.
6. **`listString` costs a walk of the story per read.** Word's numbering is a running count, so
   a label cannot be read off the paragraph: the label is `Emulator.getNumber` counted in a
   fresh `NumberingState` walked over the paragraph's story in document order - the rule of
   CR-001 section 15.2, with headers and footers sharing one state and a `w:txbxContent`
   starting its own. `Body.listLabels()` does that walk once and returns a
   `Map<P, ListLabel>` (`listString`, `level`, `numId`, `siblingIndex`, `isBullet`) for every
   numbered paragraph of the story, which is what a caller that wants them all should use.
   The walk is `test/parity.test.mjs`'s `walkNumbering`, which is the Java harness's own.
7. **A header or footer that has not been unmarshalled contributes nothing to the count.**
   The story a footer counts in is shared with every header (and every other footer), and a
   synchronous read cannot unmarshal them. So a footer's label is the number it would have if it
   were the only part read, unless the headers have been read first (`await part.getBody()` on
   each). The golden cross-check reads every story first, which is why it agrees with docx4j
   there; nothing else is affected, since the main story is one part.
8. **`attachToList`, `detachFromList` and `listItem.level` do not refresh the resolver.** They
   write `w:pPr`, not the numbering part, so there is nothing to rebuild; only the writes that
   touch `word/numbering.xml` (`setLevel*`, `separate`, `restart`, `startNewList`) call
   `refreshPropertyResolver()`. All of them go through `Paragraph.pPr()` / `Paragraph.numPr()`,
   which is where a tracked write records `w:pPrChange` - `pPr()` stopped being private for
   that, and a tracked `attachToList` is tested.
9. **`detachFromList` writes `w:numId` 0 when the style still numbers the paragraph.** Removing
   the paragraph's own `w:numPr` would leave a "List Number" paragraph numbered; ECMA-376
   17.9.18's `w:numId` 0 is how Word switches an inherited list off, and is what
   `Emulator.resolve` reads.
10. **`getLevelString` returns the string** where Office JS returns a `ClientResult`: nothing is
    marshalled to answer it (the same reasoning as `CustomXmlPart.getXml`, section 12).
    `levelExistences` is Office JS's array; `levelExists(level)` is the extension beside it.
11. **The null objects are local.** `listOrNullObject` / `listItemOrNullObject` build the same
    proxy `src/office-js/proxy.mts` does for `getFirstOrNullObject`, with the same
    `code: 'ItemNotFound'`, rather than importing it: the content API must not pull the shim's
    proxies into a bundle that only wants the views.
12. **`setLevelBullet`'s six presets** are the characters and faces Word's bullet library uses:
    Symbol F0B7 (solid), Courier New `o` (hollow), and Wingdings F0A7, F076, F0D8 and F0FC
    (square, diamonds, arrow, checkmark). `Custom` takes `charCode` and `fontName`; the font is
    written to the level's `w:rPr/w:rFonts` (ascii, hAnsi and `w:hint="default"`), as Word does.
13. **`setLevelNumbering`'s `formatString`** is Office JS's array of strings and numbers, where
    a number is a 0-based level whose counter goes there: `['(', 0, ')']` is `w:lvlText`
    `"(%1)"`. Without one the level numbers itself with a trailing full stop (`"%<n>."`).
14. **A `w:lvlOverride/w:lvl` for the level being written is removed**, since it would mask the
    abstract level the write goes to; a `w:startOverride` beside it stays (that is the restart).
    `setLevelStartingNumber` removes the level's `w:startOverride` too, for the same reason.

**The shim.** `Word.ListLevelType`, `Word.ListNumbering` and `Word.ListBullet` are in
`src/office-js/enums.mts` and on the `Word` object. Nothing else was needed: `List` and
`ListItem` are classes, so the wrapper proxies them by name like every other view, `body.lists`
and the four paragraph members pass through, and `COLLECTION_ITEM_CLASS` gained
`getLevelParagraphs` and `getDescendants` so that an empty one still names `Paragraph`. The
null objects are plain objects with a `Proxy`, which `isView` leaves alone, so
`isNullObject` reads through the shim as it does directly. `supported.generated.mts` is
regenerated (230 members).

**`toApiScript`.** A direct `w:numPr` is now `p1.attachToList(3, 0);` (and `detachFromList()`
for a `w:numId` of 0) instead of the `insertXml` fallback - the emitter's property pairs gained
a call form. It names the `w:numId` exactly as the existing style assignment names a style id:
both assume the target document defines it. `startNewList()` is **not** emitted: it is
asynchronous, and choosing between the decimal and the bullet set would need the numbering
part, which `toApiScript` does not have (it takes elements as readily as views).

**Objects-package gaps: none.** Everything is built with `factory/org_docx4j_wml`
(`createNumberingNum`, `createNumberingNumAbstractNumId`, `createNumberingNumLvlOverride`,
`createNumberingNumLvlOverrideStartOverride`, `createLvl`, `createNumFmt`, `createLvlLvlText`,
`createLvlStart`, `createJc`, `createPPrBaseInd`, `createCTLongHexNumber`,
`createPPrBaseNumPrNumId`, `createPPrBaseNumPrIlvl`) and `deepCopy` / `unmarshalNode` from the
facade; `linkParents` comes from `builders/wml`. One small thing for the objects package's own
CR when it next opens: `deepCopy` of a `w:abstractNum` is the only place here that needs a
fragment copy of a non-global element, and it works.

**Tests.** `test/lists.test.mjs`: `startNewList` (the part and its relationship created, both
default sets present, 1. 2. 3., and the same after a round trip), the bullet set, `attachToList`
and `level` (1. a. b. 2.), `detachFromList` and the two null objects, `restart` (1. 2. 1. 2.,
the shared `w:abstractNum` and the `w:startOverride`), `setLevelNumbering` /
`setLevelStartingNumber` / `setLevelIndents` / `setLevelAlignment`, `setLevelBullet` including
`Custom`, the copy-on-shared-abstract rule (two lists on one abstract, one edited, the other
unchanged), `siblingIndex`, `getAncestor` / `getDescendants`, `List.insertParagraph`,
`isListItem` through a "List Number" style with the `w:numId` 0 and dangling-id cases, a tracked
`attachToList`, the byte-for-byte round trip of a read-only document, and the golden
cross-check. Two existing tests were updated: `office-js.test.mjs` used `paragraph.listItem` as
its example of an unsupported member (it is supported now) and expected the `insertXml`
fallback for `w:numPr` (a `w:framePr` paragraph is the fallback example instead).
