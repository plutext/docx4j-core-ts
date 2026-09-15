# CR-002: A content API in the shape of Office JS, over the docx4j tree

**Status:** Phases B and D implemented 2026-09-10 (section 7); phase A implemented 2026-09-10 as objects CR-002; phases C, G and I implemented 2026-09-15 (sections 8, 9 and 10); phases E, F and H proposed.
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
  `Emulator`; until then `listString` is `undefined` and the editor's list rendering waits
  for Phase B, as CR-003 section 3.4 says.
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
- `InlinePicture.altTextTitle` reads and writes `wp:docPr/@title`, which **is not marshalled**:
  docx4j's schema (ECMA-376 1st edition) has no `title` attribute on `CTNonVisualDrawingProps`,
  so the model has no property for it. Word's own title survives an untouched part, being read
  from the source bytes; a title set through this API does not. See the gaps below.
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

1. **No `wp:docPr/@title`.** `org_docx4j_dml.CTNonVisualDrawingProps` has `descr` but no
   `title`; docx4j's generated class has none either (its `xsd/` is the 1st edition), so
   Office JS's `altTextTitle` cannot round trip. The fix is in the schemas or in an
   `anyAttributes` escape on that type, not here.
2. **No drawing or inline-picture constructor in `builders/wml`.** `drawingFor()` here is the
   counterpart of docx4j's `createImageInline` and needs only the tree (the relationship id is
   a string), so it belongs in the objects package next to `p`, `r` and `tbl`. It is written
   over the generated factories, so moving it is a copy.
3. **No row or cell builder.** `tbl(rows)` builds a whole table; `addRows` and `insertRows`
   need one row of the table's widths, so `rowElement()` here calls `el.tr` / `el.tc` directly.
   `tr(cells, opts)` and `tc(blocks, opts)` in `builders/wml` would cover it.
4. `walk` does not enter DOM nodes held by `xs:any` properties. That is the documented
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
- **No `w16cex` part is created**, only kept in step when the document has one: the objects
  package has no module for that namespace (below), so it is edited as a DOM.
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

1. **No `org_docx4j_w16cex` module.** `w16cex:commentsExtensible` (Word 2018: a durable id and a
   UTC date per comment, and the anchor of comment reactions) is not generated, so the part loads
   as a `DefaultXmlPart` and phase G edits its DOM. docx4j's Java has the same gap, so this is a
   schema to add there first.
2. **`mc:Ignorable` is missing from `w:comments`.** `Styles` and `w16cid:CTCommentsIds` carry an
   `ignorable` property, `Comments`, `CTCommentsEx` and `CTPeople` do not, so a re-marshalled
   comments part loses Word's `mc:Ignorable="w14 w15 ..."` while keeping the `w14:paraId`
   attributes it covers. Modern Word understands w14 natively and opens the result; the manual
   check in `test/README.md` is where a regression would show.
3. **`w15:CTPerson.contact` is declared required** (`contact: string`) but Word omits it in every
   `w:people` part seen. The generated factory's init is partial, so the property is simply left
   out and the returned object's type claims a `contact` that is not there. Either the schema or
   the binding should make it optional.

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
