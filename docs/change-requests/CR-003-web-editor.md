# CR-003: A web docx editor over the engine: architecture, and what it needs from here

**Status:** Proposed 2026-09-11, decisions recorded the same day (section 10); design only, the editor is its own repository
**Depends on:** CR-001 Phase A (done), CR-002 phases B and D (done); for rendering fidelity
CR-001 Phase B; for tables and content controls CR-002 phases C and E
**Scope:** a browser editor that approximates WYSIWYG for paragraphs of text and tables, edits
sections and footnotes (3.1, 3.2), shows images (deletable, not editable) and headers and
footers (read-only), round-trips everything else untouched, and whose main purpose beyond text is content control insertion (including
nesting) and deletion for OpenDoPE templates. Complements Word; does not replace it. The
underlying design supports co-editing from day one; the UI surfaces it later.

## 1. Feasibility

High, and the engine is most of the hard part already. Three things make a docx editor hard:
round-tripping what you do not understand, mapping the rich model to something an editor can
work on, and the browser's editing surface. The first is CR-001's design rule (untouched parts
byte for byte, unknown content kept as typed or DOM objects, MCE resolved only where needed)
plus the object model's completeness. The second is a projection problem with a known answer,
described below. The third is solved by ProseMirror, which has had a decade of production use
on `contenteditable` + `beforeinput` and comes with schema validation, a transaction model,
history and a Yjs binding.

The nearest existing system is SuperDoc (Harbour Enterprises, AGPL/commercial), a docx editor
that began on ProseMirror, Yjs and JSZip and whose V2 has moved to its own OOXML-backed
document model and renderer. Appendix C describes it and compares the two; in short, it
validates the problem and the market, and the differences are the typed object model and
packaging engine underneath, the editing surface, the OpenDoPE focus and the licence.

## 2. Architecture: three layers, one source of truth per layer

```
 browser
 ┌──────────────────────────────────────────────────────────────────────┐
 │ view      ProseMirror EditorView (contenteditable + beforeinput today; │
 │           EditContext when it is cross-browser)                        │
 │ edit      ProseMirror EditorState over a Yjs Y.XmlFragment             │
 │           (y-prosemirror binding): the EDITABLE PROJECTION            │
 │ document  @docx4j/core-ts WordprocessingMLPackage: the DOCX itself,    │
 │           untouched parts byte-for-byte, typed tree for the rest       │
 └──────────────────────────────────────────────────────────────────────┘
```

- **Document layer** (this package, in the browser). The dropped file is loaded with
  `OpcPackage.load(bytes)`; the runtime uses the browser's DOM; `fflate` unzips. Only the
  main document part, styles, numbering, fonts, theme, settings and the headers and footers
  are unmarshalled; images are `ImagePart` bytes turned into blob URLs; everything else
  stays as bytes and is copied back on save. This layer is the source of truth for the
  *file*.
- **Edit layer.** A ProseMirror document whose schema covers exactly the editable subset
  (section 3), held in a Yjs `Y.XmlFragment` so that every edit is a CRDT operation from the
  start. This layer is the source of truth for the *edit session*: undo (`y-undo-manager`),
  co-editing (any Yjs provider: WebSocket, WebRTC, or none), cursor positions.
- **View layer.** ProseMirror's view renders the edit document to DOM; CSS generated from the
  document layer (section 5) makes it look like the page. Read-only regions (headers,
  footers, images, unknown blocks) are node views that render a snapshot and take no input.

Positions, selection, IME and accessibility are ProseMirror's problem, not this design's.

## 3. The projection: what becomes editable, and how nothing is lost

The rule: an editable node in the edit layer carries the **id of the element it came from**
and only the properties the editor can change. Everything else about that element stays in
the document layer and is reused on export. Effective formatting (styles resolved, Phase B)
is computed for display only and never written into the projection: node attributes hold
direct formatting and references, as SuperDoc's import rule also insists (Appendix C).

| Document layer | Edit layer (ProseMirror schema) | Editable |
|---|---|---|
| `w:p` | `paragraph { sourceId, style, align, ... }` with inline content | text, runs, style, alignment, indents (Phase B values for display) |
| `w:r` + `w:rPr` | `text` with marks `bold`, `italic`, `underline`, `strike`, `sub`, `sup`, `font { name, size, color, highlight }`, `runStyle`; a `runSource` mark carrying `sourceId` | the run mapping (`applyRunOptions` / `readRunOptions`) |
| `w:tbl` / `w:tr` / `w:tc` | `table`, `table_row`, `table_cell { sourceId, colspan, rowspan }` (prosemirror-tables) | cell text, rows and columns added or deleted; `tblPr`/`tcPr` carried opaque |
| `w:sdt` (block) | `sdtBlock { sourceId, sdtPr }` with block content | insert, delete, nest; `sdtPr` edited in a side panel |
| `w:sdt` (run) | `sdtInline { sourceId, sdtPr }`: an inline node with inline content | same; ProseMirror allows inline nodes with content, and nesting is by nodes, not marks, so nested run-level controls are ordinary |
| `w:drawing` / `w:pict` inline | `image { sourceId, src: blobUrl, width, height }` atom | delete only (the relationship and part go with it on export when unreferenced) |
| `w:hyperlink` | `link` mark `{ sourceId, href }` | text inside; the target in a later phase |
| `w:footnoteReference` in a run + `w:footnote` in the footnotes part | `footnoteRef { sourceId, footnoteId }` inline atom; the footnote's paragraphs as a `footnote` block document of the same schema in a side fragment (3.2) | text of the footnote, insert, delete; endnotes the same |
| `w:commentRangeStart` / `End` + `w:commentReference` + the comments parts | `comment { id }` mark on the commented text (one per paragraph the range touches); comment content as a small document per comment in a side fragment (3.3) | view, insert, reply, resolve, delete |
| `w:fldSimple`, `w:ins`, `w:del`, bookmarks, math, `w:altChunk`, unknown blocks | `opaque { sourceId, kind, snapshotHtml }` atom, block or inline | none; deletable as a whole |
| headers, footers | separate read-only ProseMirror views, same schema, `editable: false` | none |
| `w:sectPr` (in a paragraph's `pPr`, or the body's last) | `paragraph.attrs.sectPr` (opaque JSON of the original, plus the editable fields) and `doc.attrs.finalSectPr` | page size, orientation, margins, section type; insert or remove a section break (3.1) |

Import walks the body with `Body.paragraphs` / `childrenOf` and the run segments
(`segmentsOf`, `runItemsOf`), assigning each element an id kept in a side map
`id -> Element` (a `WeakMap` in reverse for the tree). Text inside `w:ins` is shown as
ordinary text (the insertion wrapper is opaque around it in a later phase, when tracked
changes are surfaced); `w:del` is not shown.

**Export** is the inverse and is where fidelity lives. For each edit-layer node:

- if its `sourceId` resolves and the node is **unchanged** since import (ProseMirror nodes
  are immutable, so this is an identity comparison against the imported node, cached per id),
  the original element is emitted as is;
- otherwise a new element is built from the original: `deepCopy` of the original `w:p` with
  its `pPr`, then the content array rebuilt from the node's inline content, each run from a
  text span with its marks through `applyRunOptions` on a copy of the original run's `rPr`
  (`runSource`), so a run that lost no formatting keeps its `w:lang`, `w:rsid` and whatever
  else the editor never showed;
- a node without a `sourceId` (typed by the user) is built from scratch through the builders
  (`p`, `r`, `t`, `tbl`, `el.sdt`);
- an opaque node emits its original element; a deleted image drops the element, and the
  package's `removePart` handles the relationship when nothing references the part any more.

The rebuilt elements replace the body's content array (`Body.content`), the part is marked
for re-marshalling by being unmarshalled, and `pkg.save()` writes the docx: untouched parts
byte for byte, the main part re-marshalled with the facade's prefix table. Headers and
footers were never unmarshalled for editing, only for display, so they round-trip untouched.

### 3.1 Sections (`w:sectPr`): in scope

A section's properties sit on the last paragraph of the section (`w:p/w:pPr/w:sectPr`) and,
for the final section, on the body after the last block. That maps cleanly: a `sectPr` attr
on the paragraph that ends a section, and a document attr for the final one. Editing is a
properties panel over the fields the editor understands (`w:pgSz` size and orientation,
`w:pgMar`, `w:type`, columns count), written onto a `deepCopy` of the original `sectPr` so
the header and footer references, line numbering, page numbering and anything else ride
through. **Insert a section break** is "set `sectPr` on this paragraph" (a copy of the
section it splits, as Word does); **remove** is "clear the attr", and the paragraphs join the
following section. The body's final `sectPr` is never removed, since Word requires it.

Display: a section-break decoration under the paragraph (Word's dotted line and label), and
the editor's page width and padding follow each section's `pgSz` and `pgMar` through a CSS
class per section. Header and footer references stay read-only in E1 (they show which header
part a section uses; changing them is E2, with the read-only header views following).

Effort: a day on the schema and export, two on the panel. In E1.

This is the same shape as docx4j's own docx4all (Swing, 2008) and as SuperDoc, with one
difference that matters: the projection is over a **typed** tree, so "carry the rest opaque"
is `deepCopy` of a typed object, not string surgery on XML.

### 3.2 Footnotes and endnotes: in scope, staged

A footnote is two things: a `w:footnoteReference w:id="n"` item in a run of the body, and a
`w:footnote w:id="n"` in the footnotes part (`FootnotesPart`, typed here as
`CTFootnotes`), whose content is ordinary paragraphs, the first starting with a
`w:footnoteRef` run. Ids -1 and 0 are the separator and continuation-separator notes and
are never shown or touched. Endnotes are the same with `endnote` in the names.

Projection: the reference is an inline atom `footnoteRef { sourceId, footnoteId }`, numbered
by a decoration (Word numbers by order of reference, restart rules in `sectPr`), and each
footnote's paragraphs are a small document of the same schema held in a side `Y.Map` of
fragments keyed by footnote id, so a footnote's text is edited with the same commands as the
body and co-edits like it. Display as Word Online does: the reference superscript in the
flow, the footnote text in a panel below the page (or a popover), editable in place.

Operations:

- **Read** (E1): references shown and numbered; the footnotes part is unmarshalled when the
  first footnote is displayed, not at load, which is what keeps a footnote-heavy document's
  `onReady` low (Appendix C).
- **Edit** the text of an existing footnote (E2): the same paragraph and run machinery over
  the footnote's fragment; export rebuilds the changed `w:footnote` from a `deepCopy` of the
  original, unchanged ones re-emitted.
- **Delete** (E2): deleting the reference deletes the footnote; deleting the last text of a
  footnote leaves an empty footnote, as Word does.
- **Insert** (E2): a new id (one above the highest), a reference run at the cursor copying
  the surrounding run's `rPr` plus `w:rStyle FootnoteReference`, and a footnote with one
  paragraph in `FootnoteText` style holding a `FootnoteReference`-styled `w:footnoteRef` run
  and the cursor. Both style ids come from the styles part when present and are added from
  docx4j's defaults when not.

Content controls inside footnotes work by the same schema but are not a target. Effort:
three days for read and edit, two for insert and delete. E1 read, E2 the rest.

### 3.3 Comments: view, insert, reply, delete

In scope. Comments are the other half of review (4.2): a reviewer annotates, an agent
explains what it changed and why, and both must survive into Word. The engine half is
CR-002 phase G (`getComments()`, `Comment` with `content`, `resolved`, `reply`, `delete`,
`Range.insertComment`, the four parts kept in step); the editor half:

- **Projection.** The commented text carries a `comment { id }` mark per paragraph it
  touches, so a range spanning paragraphs is several marks with one id; export writes one
  `w:commentRangeStart` before the first and one `w:commentRangeEnd` plus the reference run
  after the last. The comment's own paragraphs are a small document in a side `Y.Map` of
  fragments keyed by id, as footnotes are (3.2), so a comment's text is edited with the
  body's commands and co-edits like it.
- **Display.** A margin column with a bubble per thread aligned to its anchor, the author's
  colour on the anchored text, replies nested, resolved threads collapsed; a click selects
  the range. Word Online's shape, which users know.
- **Operations.** Select text and comment (or comment at the cursor, on the word); reply;
  resolve and reopen; delete a comment or a thread; edit one's own comment text. Author
  from the editor's user settings, or the agent's name for E4.
- **Staging.** View in E1 (read-only bubbles, the cheapest useful thing); insert, reply,
  resolve and delete in E4 with the review panel, since that is where the workflow needs
  them; a document without comment parts gets them created on first insert.

### 3.4 Formatting UI: runs, paragraphs, styles, lists, page breaks

Not everything Word offers; the working set, each control a call on the content API so that
the toolbar, the keyboard and an agent (4.2) do the same thing:

| Control | Content API | Engine status |
|---|---|---|
| Style picker: paragraph and character styles by their `w:name`, the quick styles first (`isQFormat`), a preview in each entry rendered from Phase B's effective properties | `paragraph.style`, `range.font` with `style` (`w:rStyle`) | in place (CR-002 B) |
| Font family (the document's fonts from the font table and theme, then a common list), size, bold, italic, underline; colour and highlight because they cost nothing | `range.font.name`, `size`, `bold`, `italic`, `underline`, `color`, `highlightColor`; runs split at the selection's edges | in place |
| Alignment; indent increase and decrease (left indent by half an inch, as Word's buttons); first-line and hanging indent, space before and after, line spacing in a paragraph dialog | `paragraph.alignment`, `leftIndent`, `firstLineIndent`, `spaceBefore`, `spaceAfter`, `lineSpacing` | in place |
| Bullets and numbering on and off, list level in and out, restart numbering, continue previous list | `paragraph.startNewList()`, `attachToList(listId, level)`, `detachFromList()`, `listItem.level`, `list.levelTypes` (Office JS `Word.List` and `Word.ListItem`) | **CR-002 phase H**, below; labels rendered by Phase B's `Emulator` |
| Page break at the cursor (and section break, 3.1) | `paragraph.insertBreak('Page', 'After')`, `range.insertBreak`; `sectPr` attr for a section break | in place |
| Clear formatting | `font` reset, `paragraph.style = 'Normal'` | in place |

Keyboard: the Word set (Ctrl+B, I, U; Ctrl+Shift+comma and period for size; Ctrl+M and
Ctrl+Shift+M for indent; Tab and Shift+Tab at the start of a list item for level; Ctrl+Enter
for a page break; Ctrl+Alt+1 to 3 for headings). The toolbar reflects the selection's
direct formatting today and its effective formatting once Phase B lands (a heading shows as
bold in the toolbar because its style says so), which is the same provisional note as
`Font` reads in CR-002 section 7.

Lists are the one gap. The content API has no numbering verbs and the model's numbering part
is not surfaced beyond `NumberingDefinitionsPart`; CR-002 phase H adds the Office JS list
surface: a `List` view over a `w:num` (and its `w:abstractNum`), `Paragraph.startNewList()`
creating a new definition from docx4j's default bullet or decimal definitions (the same
resource `WordprocessingMLPackage.createPackage` will use for numbering, docx4j's
`NumberingDefinitionsPart.unmarshalDefaultNumbering`), `attachToList`, `detachFromList`,
`listItem.level` writing `w:numPr` `ilvl`, and restart through a new `w:num` with a
`w:lvlOverride`. Rendering the labels ("1.", "a)", "•") is CR-001 Phase B's `Emulator`, so
numbering UI is E2, after Phase B; everything else in this section is E1.

### 3.5 Developer mode: reveal codes, a console, macros, and pasting Office JS

The editor's second audience is developers working with this API, and the editor is the
best place to learn it: the document is on screen, the selection is a range, and the tree
behind it is one click away. Four features, all cheap because the engine already has the
pieces.

**Reveal codes.** A panel that follows the selection and shows, side by side:

1. The Open XML of the selected element(s): the run(s) of a text selection, the paragraph,
   the table, the content control, whatever the caret is in, with the ancestor chain as
   breadcrumbs (`body > sdt (od:repeat) > tbl > tr > tc > p > r`). Marshalled from the tree
   (`marshalString`), so it is what the file will contain, prefixes and all.
2. The TypeScript that would build it with the factories: `el.p({ pPr: { pStyle: { val:
   'Heading1' } }, content: [el.r({ rPr: { b: {} }, content: [el.t({ value: 'Hello' })] })] })`.
   A generator over the tree (objects package, `builders/wml` `toSource(element)`: `el.<name>`
   for the module's own elements, the scoped `create...` factories for foreign ones, literal
   values, `TYPE_NAME` and `PARENT` omitted).
3. The content-API calls that would produce the same thing where the API has verbs for it:
   `const p = body.insertParagraph('Hello', 'End'); p.style = 'Heading1'; p.font.bold = true;`
   with a fallback to `body.insertXml(...)` for what the verbs do not cover (this package,
   `toApiScript(element)`). This is the "how would I write this" answer, and it is what an
   agent reads to learn the API from an example.

Edits made in the XML pane are applied through `insertXml` / `setXml` (with the MCE
preprocessor and the schema check giving the error on a bad edit), so reveal codes is also
a way to make a change the toolbar has no button for.

**A console.** A Monaco editor pane (the package's `.d.mts` declarations loaded as extra
libraries, so completion and type errors work against the real API) in which the user
writes TypeScript against `pkg`, `body`, `el`, `wml` and the selection, and runs it. The
TypeScript is stripped to JavaScript in the browser (Sucrase, about 100 KB, or esbuild-wasm)
and executed as an ES module from a blob URL with those bindings passed in, inside one Yjs
transaction so that the whole script is one undo step and one tracked-change batch, and
with the change-tracking mode honoured. Errors map back to the source line. The script runs
in the user's own page with the user's own document, the same trust as the browser's
devtools; a hosted deployment that lets others' scripts run is an E3 policy question, not a
mechanism question.

**Macros.** A saved console script with a name is a macro: stored per user (browser
storage) or, when the user asks, in the document itself as a custom XML part
(`docx4j-editor/macros`), where Word ignores it and the editor offers it on the toolbar the
next time the document opens. Together with the console this is a macro language for docx
with a typed API, source-level completion and an undo step per run, which is what VBA gave
Word and what add-ins took away. It is also the natural target for an agent: "write me a
macro that ..." produces a script the user can read, run and keep.

**Pasting Office JS.** Add-in code such as

```ts
await Word.run(async (context) => {
  const paragraphs = context.document.body.paragraphs;
  paragraphs.load('items/text,items/style');
  await context.sync();
  for (const p of paragraphs.items) if (p.style === 'Heading 1') p.insertText(' (reviewed)', 'End');
  await context.sync();
});
```

runs unchanged in the console, because the console provides a `Word` global: `Word.run`
hands the script a `context` whose `document.body` is this package's `Body`, `load()` is a
no-op (properties are already readable), `context.sync()` resolves the pending reads such
as `getOoxml().value`, collections are arrays with `items`, and the enums
(`Word.InsertLocation`, `Word.Alignment`, `Word.ChangeTrackingMode`, ...) are the string
values the API already accepts. No package names are rewritten: the point of CR-002
section 3.4 is that the API *is* a subset of `Word.*`, so the same source is valid on both
sides. What the subset does not have is reported, not guessed at: before the run, a static
pass over the script's member accesses against the subset's declarations highlights the
lines that use members with no equivalent here (`body.insertFileFromBase64`, `getReviewedText`,
`context.document.properties`); at run time, the shim's objects are proxies that throw a
`NotSupported` error naming the member and the line for anything the static pass missed.
The engine half is CR-002 phase I (the shim is a library, useful in Node for running an
add-in's code against a package in a test), the editor half is the console integration and
the highlighting.

One caveat to state plainly: code that works here is not guaranteed to work in Word,
because Office JS requires `load()` and `sync()` before a property is read and this
package does not. The console can lint for that (a property read on a proxy with no
preceding `load` of that property) and offer to insert the `load` calls, so that a script
developed here is portable back to the add-in.

**Asking for content.** The console's prompt box takes natural language as well as code:
"insert a paragraph about XYZ after this one", "make this a two-column table of the items
above", "write the definitions clause for a services agreement", "explain what this XML
does". The request goes to a model with the context the console already has: the
selection's XML and text, the outline, the styles in use with their names, and the API
reference (the same declarations Monaco holds). The model is asked to answer in one of
three fenced forms, and the console knows what to do with each:

- **Content**, as a WML fragment (`wml` block) or plain text: rendered in a preview pane
  through the same projection, over a scratch package carrying the document's styles so it
  looks as it will in place; buttons **Insert at selection**, **Replace selection**, **Copy**
  (as text, as XML, as factory source). Insert goes through `insertXml` in one Yjs
  transaction, tracked when the mode is on, so it is one undo step and one reviewable
  change.
- **A script** (`ts` block): shown in the console editor, not run; the user reads it,
  edits it if they like, and presses Run. Generated code never executes on its own.
- **An explanation** (prose): shown in the pane; the "explain this" companion to reveal
  codes, and the way a developer asks "how do I do X with this API" with the real
  declarations in the model's context.

Nothing happens to the document until the user presses a button, which is the difference
from E4's agent, where the model calls tools on the live document as a peer. The two share
the provider configuration (an API key of the user's, or an organisation's proxy, or the
E4 companion process relaying to whatever the organisation runs), the context builders and
the fenced-form contract, so the ask mode is E4's first half and ships before it. It is
also the cheapest way to get LLM value into the editor: no tools, no tracking of an
autonomous session, just a preview and an insert the user controls.

Staging: reveal codes (XML and factory source) and the console in E1, since they cost
days and they are how the team itself will debug the projection; the API-script generator,
macros, the Office JS shim and the ask mode with a bring-your-own key in E2; the
organisation-proxy and companion routes with E4.

### 3.6 Perspectives

The features above serve three different people, and a toolbar that shows all of them at
once serves none well. The UI is organised into **perspectives** in Eclipse's sense: one
document, one model, one selection; a perspective is a named arrangement of panels,
toolbar, keyboard map and defaults, switched with one click or a shortcut, remembered per
user and suggested from the document (a docx with OpenDoPE parts opens in Template; a docx
with pending tracked changes or comments offers Review).

| Perspective | For | Panels and toolbar | Defaults |
|---|---|---|---|
| **Edit** | writing and formatting | the formatting toolbar (3.4); comments in the margin (3.3); footnotes below (3.2); the ask box (3.5) as the one model feature; page boundaries from markers (Appendix D) | tracking off; content controls drawn lightly (Word's bounding boxes on hover); developer panels hidden |
| **Review** | seeing what changed and deciding | the review panel: tracked changes by author and time with accept and reject, comments and replies, the agent's log when E4 has run; "show as accepted" and "original" toggles; a compare view against the last saved version (Yjs snapshot) | tracking on for the user's own edits; the agent's edits highlighted by author colour; formatting toolbar reduced |
| **Template** | OpenDoPE authors | the content control tree (nesting, kinds, tags) with the properties panel (4); the custom XML parts as a tree with the XPath box and bind-by-drag (Word's XML Mapping pane, done properly); conditions, repeats and questions editors; sample data and the binding preview (`applyBindings`); row-level control commands on tables (4.1) | controls drawn with tags and coloured gutters; the ask box biased to template help ("make this row repeat over the line items") |
| **Developer** | people working with the API, and the team | reveal codes with the XML, factory-source and API-script panes (3.5); the console with Monaco and the `Word` shim; macros; the outline as an address table (3.4 addresses); the part list of the package with each part's XML; the projection diagnostics (what is opaque and why) | tracking off; the console bound to `pkg`, `body`, `el`, `wml`, the selection and the current perspective's panel state |

What the perspectives share is everything that matters: the projection, the commands (a
command is one content-API call, registered once, bound to toolbar buttons, keys, the
console and the MCP tools alike), the undo history, the change-tracking mode (a document
setting, not a perspective setting, so switching perspectives never silently changes
whether an edit is tracked), the selection, and the agent's peer session (E4), whose tool
surface does not depend on the perspective, though the Template perspective's ask box
supplies template-oriented context and the Developer perspective's console shows the
agent's calls as they happen.

Two rules keep it honest. A perspective **hides** panels and buttons, it never removes
capabilities: any command remains available from the command palette (Ctrl+Shift+P) in
every perspective, so that a Template author can bold a word without leaving Template.
And a perspective is a **preset, not a mode**: the user can open any panel in any
perspective, and the arrangement they end up with is what "their" version of that
perspective becomes, with a reset to the default.

Staging: Edit and Developer in E1 (Developer is how the team debugs the projection);
Template in E2 with the content controls and custom XML work; Review in E4 with change
tracking and comments editing.

### 3.7 Scripts and directionality: LTR first, RTL and CJK in the foundations

The first releases target left-to-right Latin text. The foundations must not assume it,
because the places where an LTR assumption hides (the text model, search, run splitting,
font selection, CSS direction) are the ones that cost the most to revisit. What the
foundations commit to:

- **Rendering is the browser's.** Bidirectional layout, CJK line breaking, IME composition
  and vertical metrics come from the browser through `contenteditable`; the editor sets
  `direction: rtl` from `w:bidi` on the paragraph and `unicode-bidi` from `w:rtl` on runs,
  and nothing else. This is the one place where not owning the layout engine is a plain
  advantage (SuperDoc's V2 right-to-left caret and tab-order issues, Appendix C, are the
  cost of owning it).
- **Fonts by script.** CR-001 Phase B's `RunFontSelector` chooses the font per character
  range from `w:rFonts` (`ascii`, `hAnsi`, `cs`, `eastAsia`, the `hint`) and the Unicode
  script of the text; the generated CSS carries a `font-family` per script span rather
  than one per run. The `Font` view's `name` is the ASCII font today; `nameBidi` and
  `nameEastAsia` (WordApiDesktop's names) follow with Phase B. The run mapping already
  writes the complex-script twins (`w:bCs`, `w:iCs`, `w:szCs`) alongside `w:b`, `w:i`,
  `w:sz`, so bold and size apply to Arabic and Hebrew text from day one; `sizeBidi` is a
  later refinement.
- **Offsets are UTF-16 code units** in the text model, the addresses and the API, as in
  Office JS and in Java, never bytes and never code points; and a run is never split
  inside a surrogate pair or a grapheme cluster (`Intl.Segmenter` with grapheme
  granularity at the split point), so combining marks, emoji sequences and Indic clusters
  survive formatting a span.
- **Search is Unicode.** `matchCase` folds through locale-aware case mapping;
  `matchWholeWord` uses `Intl.Segmenter` word boundaries rather than the ASCII `\b`, which
  is wrong for CJK (no spaces) and for Arabic (joined letters); wildcards operate on code
  points. Diacritic-insensitive and kashida-insensitive matching (Word's options) are later.
- **Numbering** formats for CJK and Arabic (`chineseCounting`, `arabicAlpha`, ...) are
  Phase B's `Emulator`, where docx4j already has them.
- **The UI** is mirrored by CSS logical properties from the start (`margin-inline-start`,
  not `margin-left`), so a right-to-left interface is a stylesheet direction switch, and
  the toolbar's indent buttons mean "start" and "end" (`w:ind/@w:start`, `@w:end`, the
  transitional `left`/`right` read as aliases).
- **Out of scope** until asked for: vertical text (`w:textDirection`), East Asian layout
  properties (`w:eastAsianLayout`), and Word's RTL table layouts (`w:bidiVisual`), which
  round-trip untouched as everything else.

Tests: the fixtures gain a Hebrew or Arabic document and a Japanese one from the docx4j
corpus; the round-trip and projection tests run over them; a search test with a CJK phrase
and a whole-word match on Arabic; a split-at-grapheme test with an emoji sequence.

## 4. Content controls and OpenDoPE

The editor's reason to exist beyond text. Operations:

- **Insert** a content control around the selection: block-level when the selection spans
  whole blocks, run-level otherwise; inside another control when the selection is inside
  one (nesting is natural in the schema). Kinds: rich text, plain text, repeating section
  (w15), and OpenDoPE's `od:condition` / `od:repeat` / `od:xpath` tags.
- **Delete** a control keeping or dropping its content (Office JS `delete(keepContent)`).
- **Properties panel**: tag, alias, id, lock, appearance, and for OpenDoPE the XPath, the
  condition or repeat it refers to, with the XPaths, conditions and questions parts edited
  alongside (custom XML parts are DOM here, XPath through CR-002 E's engine).
- **Binding preview**: `applyBindings` (CR-002 E) shows the sample data in place.

### 4.1 Content controls around table rows and cells

OpenDoPE's staple is a repeat on a table row: `w:sdt` around one or more `w:tr` (the model's
`CTSdtRow` with `CTSdtContentRow`), and sometimes around a cell (`CTSdtCell`). A control
around a whole table is the ordinary block case (`sdtBlock` containing `table`) and needs
nothing.

Rows are the problem, because `prosemirror-tables` (the table editing plugin the editor
relies on for cell selection, column resizing and row and column commands) assumes the
schema `table > table_row > table_cell` with nothing in between; a wrapper node between
`table` and `table_row` breaks its `TableMap`. Two ways out:

1. Extend `prosemirror-tables` to accept a `sdtRow` group node (content `(table_row |
   sdtRow)+`, `sdtRow` content `table_row+`) and flatten it in `TableMap`. Moderate work in a
   fork, to be maintained.
2. Keep the schema flat and put the control on the rows as an **attribute**:
   `table_row.attrs.sdt = [{ groupId, sourceId, sdtPr, first, last }]`, a stack so that a row
   inside nested controls (a repeat inside a condition) lists them outermost first. Rows with
   the same `groupId` are one control. Cells likewise: `table_cell.attrs.sdt`.

Option 2 is the recommendation. It leaves `prosemirror-tables` untouched, nesting is
expressible, and every operation is an attribute transaction:

- **Wrap**: select rows (a `CellSelection` spanning them, or the rows the cursor is in),
  "Wrap rows in content control" pushes a new group onto each selected row's stack. Rows
  must be contiguous and within one table; the command is disabled otherwise.
- **Unwrap**: remove the group from the rows that carry it; the rows stay. "Delete control
  with content" deletes the rows.
- **Edit**: the same properties panel as block controls, bound to the group's `sdtPr`
  (`Y.Map` per group, so co-editors converge).
- **Row commands** from `prosemirror-tables` keep working; an added row inside a group
  inherits the stack of the row it was added next to (a plugin's `appendTransaction`
  maintains the invariant that a group's rows stay contiguous, and splits or merges groups
  when a row is moved out).

Display: a coloured left gutter per group with the tag or alias, nested groups as nested
gutters, matching how block controls are shown.

Export rebuilds the structure from the attrs: walking a table's rows, a run of consecutive
rows sharing the outermost `groupId` becomes one `CTSdtRow` around them (recursively for
the inner groups), built from the original `sdtPr` when the group has a `sourceId` or from
the builders when it was created in the editor. The same for cells. Import is the inverse
walk, assigning `groupId`s. Because a row-level control that existed in the file keeps its
`sourceId`, an unchanged one is emitted as the original element, as everything else.

The engine side of this is CR-002 phases C and E; the editor needs nothing beyond them.

### 4.2 An agent driving the editor, and change tracking

In scope, as a phase of its own (E4), because it is the workflow the product is for: an
OpenDoPE author asks an LLM to draft, restructure or template a document and reviews the
result. Two design points, one architectural and one about what the user sees.

**The agent is a peer, not a plug-in.** Once E3 gives the editor a Yjs provider, anything
that can speak Yjs is a collaborator, with presence and attribution. The agent connects
that way: a small local companion process (`docx4j-editor-mcp`, Node) serves MCP over stdio
or streamable HTTP to Claude Code, Claude Desktop or any MCP client, holds the same package
as the browser through the same provider, and applies each tool call as one Yjs transaction.
The user watches the document change in place, with the agent's cursor and colour, and can
type at the same time. Nothing new is invented for it: the MCP tools are the content API
(CR-002) in Office JS shapes plus the addresses and `outline()` of section 3.3, which is
also why those exist. A no-editor variant (the same tools over a file on disk, `docx4j-mcp`
style) is the same server without a provider.

The tool surface, first cut (each maps to one content-API call, addresses as in 3.3):
`outline`, `get_text(address?)`, `search(text, options)`, `replace_text(find, replace,
options)`, `insert_paragraph(address, location, text, style?)`, `set_text(address, text)`,
`delete(address)`, `insert_xml(address, location, wml)`, `format(address | range, font)`,
`set_paragraph(address, { style, alignment, ... })`, `insert_table`, `insert_content_control(range
| rows, kind, sdtPr)`, `remove_content_control(address, keepContent)`,
`set_content_control(address, sdtPr)`, `bind(address, xpath, part)`, `get_tracked_changes`,
`accept`/`reject(changeId | all)`, `save`. Every mutating tool returns what changed
(addresses and a short diff of text), which is what the model needs to continue, and what
the reviewer sees in the log panel.

**Everything the agent does is a tracked change.** The editor keeps a change-tracking mode
in the shape of Office JS (`document.changeTrackingMode`: `Off`, `TrackAll`, `TrackMineOnly`)
and the agent's session runs with tracking on and the agent's name as author. Insertions are
written as `w:ins` around the new runs, deletions as `w:del` with `w:delText`, inserted and
deleted paragraph marks on `w:rPr` of the mark, format changes as `w:rPrChange` /
`w:pPrChange` with the old properties, table row insertions and deletions on `w:trPr`. That
is Word's own representation, so the reviewer can accept or reject in the editor **or in
Word**, and a document handed on carries the review state. The editor renders tracked
changes as Word does (author colour, underline for insertions, strikethrough for deletions,
a change bar in the margin), with a review panel listing changes by author and time and
accept / reject per change, per author or all, and a "show as accepted" toggle. This
upgrades E1's "tracked changes are opaque" (decision 2): E1 still displays existing
changes read-only; E4 makes them first-class, since it creates them.

The engine half is CR-002 phase F (recorded there): `changeTrackingMode` on the package,
`TrackedChange` views with `accept()` / `reject()` and `getTrackedChanges()` on `Body`,
`Paragraph` and `Range`, every content-API mutation honouring the mode, and `replaceText`
as a verb. The editor half is the rendering, the panel and the mode switch; the agent half
is the companion process.

Why tracked changes rather than Yjs snapshots for "what did the agent do": snapshots give
a diff of the edit session and are kept for undo and history, but they do not leave the
editor; a `w:ins` does. Both exist; the review surface is built on the tracked changes.

## 5. Approximating WYSIWYG

ProseMirror renders semantic DOM; the look comes from CSS generated by the document layer:

- **Styles**: `PropertyResolver.getEffectivePPr` / `getEffectiveRPr` (CR-001 Phase B) for
  every style in use, emitted once as CSS classes (`.s-Heading1 { font: ...; margin: ... }`);
  direct formatting as marks maps to inline CSS through `readRunOptions`.
- **Fonts**: `RunFontSelector` + `IdentityPlusMapper` (Phase B) give `font-family` stacks
  with the metric-compatible substitutes; the theme's major and minor fonts resolve through
  the theme part.
- **Numbering**: the `Emulator` (Phase B) gives each list paragraph its label and indent;
  rendered as a decoration widget before the paragraph so it is not editable text.
- **Page**: `sectPr` page size and margins set the editor's width and padding. No layout-based
  pagination; page boundaries are shown from the `w:lastRenderedPageBreak` markers Word left
  in the file, refreshed on demand (Appendix D), and headers and footers are shown at them.
- **Tables**: `tblPr` borders, widths and shading to CSS; `prosemirror-tables` for the
  editing behaviour.
- **Images**: `wp:inline` extent in EMUs to pixels; anchored drawings are opaque.

Fidelity is "close enough to edit confidently", not print fidelity. Phase B is therefore
the gating dependency for the editor looking right, which is why it is next.

## 6. Co-editing and the representation question

**Representation.** A piece table is the right structure for a plain-text buffer (VS Code's
piece tree): cheap edits, cheap undo, one flat string. A docx body is a tree of typed nodes,
and its edits are tree operations, so the structure that fits is a **persistent tree with
structural sharing**, which is what ProseMirror's document is (immutable nodes; an edit
copies the path from the root, everything else is shared), and, for co-editing, a **sequence
CRDT per text node** with a tree of them, which is what Yjs's `Y.XmlFragment` is (YATA over
items, with an internal representation that behaves like a rope of runs). Yjs's item store
is, in effect, a piece table whose pieces are CRDT items with identity; there is no need to
build another. Recommendation: Yjs is the edit-session representation from the first
commit, with `y-prosemirror` binding the view to it, even with a single user and no
provider. That gives undo, snapshots and a binary update log for free, and makes co-editing
a transport question (`y-websocket`, `y-webrtc`, or an in-house relay) rather than a
rewrite.

**Persistence beyond the docx.** The Yjs state (a few KB of updates) can be stored in a
custom XML part of the docx itself, so that reopening the file in the editor restores the
session, history and remote peers' positions, while Word ignores the part.

**What does not co-edit.** Opaque regions and headers and footers are single-writer by
construction; two users editing the same `sdtPr` in the panel is last-writer-wins on a
`Y.Map`, which is acceptable for template authoring.

## 7. The editing surface, and the W3C Editing WG

Today: `contenteditable` with `beforeinput` (Input Events Level 2, the WG's shipped
deliverable), which ProseMirror drives and normalises across browsers.

Later: the WG's **EditContext API** decouples text input (IME composition, emoji pickers,
accessibility) from the DOM rendering, so an editor can own its rendering completely. It is
in Chrome since 121; Firefox has a positive position but no implementation; Safari's
position is unknown. It is not a dependency for this editor; it becomes interesting if
pagination or a canvas renderer is ever wanted. ProseMirror will adopt it when it is
cross-browser, and this design inherits that.

## 8. Effort and phases (editor repository)

| Phase | Content | Effort |
|---|---|---|
| E1 | The Edit and Developer perspectives (3.6); reveal codes and the console (3.5); drop a docx, project paragraphs and runs to ProseMirror in Yjs, render with generated CSS (Phase B for effective values), export, download; the formatting toolbar (styles, font, size, bold, italic, underline, alignment, indents, page breaks; 3.4); sections editable; footnote references shown and their text readable; comments shown read-only in the margin; images shown; headers and footers read-only; the 400-page load benchmark, reported not gating | 3 weeks |
| E2 | The Template perspective (3.6); macros, the API-script generator, the Office JS shim and the ask-the-model mode (preview, insert, copy) in the console (3.5, CR-002 I); tables (needs CR-002 C), content controls with the properties panel, nesting and row-level controls (needs CR-002 E), OpenDoPE tags and the custom XML parts; footnotes and endnotes edited, inserted and deleted; bullets and numbering (needs CR-002 H and Phase B) | 4 weeks |
| E3 | Co-editing surfaced: a relay, presence | 2 weeks |
| E4 | The Review perspective (3.6); an agent as a peer: the `docx4j-editor-mcp` companion serving the content API as MCP tools over the same provider; change-tracking mode with Word's revision markup (CR-002 phase F), tracked-change rendering, the review panel with accept and reject; comments inserted, replied to, resolved and deleted (CR-002 phase G); find and replace in the UI | 4 weeks |

## 9. What the engine needs to provide (this repository)

- CR-001 Phase B as specified: `PropertyResolver`, `Emulator`, `RunFontSelector` with
  `IdentityPlusMapper`. (Next.)
- CR-002 C: `Table` views and `insertInlinePictureFromBase64`; CR-002 E: content controls
  and custom XML.
- Small additions, to be folded into those CRs: a stable element identity that survives a
  `deepCopy` (a `WeakMap` from element to id is enough, done by the editor, but a
  `Body.idOf(element)` helper avoids two of them); `ImagePart` to `Blob` and `wp:inline`
  extent readers; `removePart` when the last reference goes (`RelationshipsPart.removePart`
  exists; the check that nothing else references the part is a one-liner on
  `sourceRelationships`).
- Nothing from the runtime or the objects package beyond what is landed.

## 10. Decisions (2026-09-11)

The open questions of the first draft, decided as recommended:

1. Two packages in one repository, `docx4j-editor`: `@docx4j/editor-model` (the projection,
   headless, testable in Node with jsdom; what an agent or a test uses) and the app.
2. Tracked changes are opaque in E1 (insertions shown as ordinary text, deletions hidden),
   with a banner; they become first-class in E4 (4.2), which creates them.
3. E1 does not ship without Phase B: a document that looks wrong undermines trust in the
   round trip, which is the product.
4. Sections are in scope (3.1). Row- and cell-level content controls use the attribute
   representation (4.1).
5. Footnotes and endnotes are in scope, read in E1 and edited, inserted and deleted in E2
   (3.2), because the template audience is largely legal and footnotes are frequent there;
   the footnotes part is unmarshalled on first display, never at load.
6. Comments are in scope: viewed in E1, inserted, replied to, resolved and deleted in E4
   (3.3), as the other half of review. Decided 2026-09-12.
7. The formatting UI is the working set of 3.4: styles, font, size, bold, italic,
   underline, alignment, indents, page breaks in E1; bullets and numbering in E2 over
   CR-002 phase H and Phase B. Decided 2026-09-12.
8. Developer mode is in scope (3.5): reveal codes with XML, factory source and API script;
   a TypeScript console with completion from the package's declarations, one undo step
   per run; saved scripts as macros; pasted Office JS run through a `Word` shim with
   unsupported members highlighted; an ask mode in which a model proposes content or a
   script that the user previews and applies, nothing executing or inserting on its own.
   Decided 2026-09-12.
9. The UI is organised as four perspectives, Edit, Review, Template and Developer (3.6),
   presets over one model and one command set; a perspective hides, never removes, and the
   tracking mode is a document setting independent of it. Decided 2026-09-12.
10. The 400-page, 2-second load figure is an aspiration, not a requirement: the Developer
   and Template perspectives are the product's purpose and are served either way; the
   benchmark is measured and reported and never gates a release. Decided 2026-09-12.
11. Left-to-right Latin text first; the foundations (3.7) carry RTL and CJK: browser
   bidi and IME, fonts by script through Phase B, UTF-16 offsets with grapheme-safe
   splitting, Unicode-aware search, logical CSS properties. Decided 2026-09-12.
12. An edited document gets the editor's name and version in `docProps/app.xml`, unless
   turned off in the options pane; nothing else Word-visible is written without being
   asked (section 11, item 2). Decided 2026-09-12.
13. Repository `docx4j-editor` with the model, UI and MCP packages; ProseMirror directly;
   React for the application with the editing surface as a framework-free class and a web
   component wrapper (section 11, item 10). Decided 2026-09-12.
14. An LLM driving the editor through MCP is in scope as phase E4, with the agent joining
   as a Yjs peer and every edit it makes written as a Word tracked change under its own
   author name, so the user reviews in the editor or in Word (4.2). Decided 2026-09-12.

## 11. Foundations to settle before E1

Each of these would be baked in by E1's first weeks and is expensive to change afterwards.
A recommendation for each; decisions go into section 10 as they are made.

1. **Files: open, save, and the file Word may also have open.** Open is drag-and-drop or
   a file picker; save is a download by default, and save-in-place through the File System
   Access API where the browser has it (Chromium), with autosave to browser storage (the
   Yjs update log, which is what survives a crash) and a visible dirty state. The editor
   never writes the original file without asking, and before any save it runs the
   **round-trip self-check**: export to bytes, reload, re-project, compare with the live
   projection; a mismatch blocks the save with the reveal-codes diff and never overwrites.
   A file Word has open is a last-writer-wins situation the editor cannot detect; it warns
   when a file's modification time changed since it was opened. Recommendation: as above;
   the self-check from E1's first day, since it is the fidelity guarantee made
   operational.
2. **What the editor writes into a document that Word will see.** Nothing, by default,
   beyond the user's edits, with one exception: **when the document has been edited, the
   editor writes its name and version to `docProps/app.xml` (`Application`, `AppVersion`)
   as Word does, unless the user turns that off in the options pane** (decision, Jason,
   2026-09-12). No `w14:paraId`s unless Re-paginate or the user asks, no macros or Yjs
   state unless the user chooses "keep session in document", and the core properties'
   `lastModifiedBy` and `modified` are left as they are (a later option, off by default,
   could set them). A document opened and saved without edits round-trips byte for byte,
   and E1's tests assert it; an edited one differs only in the edit and in `app.xml`.
   docx4j's Java engine has the same switch (`docx4j.App.write`, off by default there);
   here the default is on, because a document edited in the editor should say so.
   **`AppVersion` must be of the form `XX.YYYY`, two digits, a dot, four digits** (Word
   2010 x64 treats a document as corrupt on any other value, a `-SNAPSHOT` suffix
   included; Jason, 2026-09-12). The editor therefore never writes its semver string:
   `major.minor.patch` maps to `MM.mmpp` (`0.3.1` becomes `00.0301`; `1.12.4` becomes
   `01.1204`), pre-release and build suffixes are dropped, minor and patch above 99 are
   clamped, and the writer validates the result against `^\d{2}\.\d{4}$` and writes
   nothing rather than an invalid value. Word's own values are of this form
   (`16.0000`). The same rule applies to any tool of ours that writes `app.xml`, the
   engine's `createPackage` included when it gains an `Application` default.
3. **Undo.** One `Y.UndoManager` over the body fragment and the side maps, scoped to the
   user's own origin, so a colleague's or the agent's concurrent edits are not undone by
   the user's Ctrl+Z; a console run and an agent tool call are one undo step each. In a
   tracked session, undo of an agent's tracked insertion removes it; reject does too but
   leaves the audit trail, which is the reviewer's tool. Recommendation: as above.
4. **Deployment shape.** A static site is the baseline: no server, everything in the
   browser, the document never leaves the machine, which is also the privacy story. The
   companion process (E4's MCP, Re-paginate through `docx4j-mcp`, a Yjs relay for E3) is
   optional and local by default; an organisation can host the relay. Recommendation:
   static-first, and every feature must say which of the two it needs.
5. **Security posture for the console, macros and the model.** The console runs the user's
   code in the user's page: the same trust as devtools, acceptable for a static site. A
   macro stored in a document is code from wherever the document came from: it never runs
   automatically, it is shown before its first run, and a hosted deployment can disable
   document macros. The model provider receives document content: the ask box says so,
   shows what is being sent (the context builder's output is inspectable in Developer),
   and the key or proxy is the user's or the organisation's, never the editor's.
   Recommendation: as above, written into the UI, not a policy page.
6. **Paste.** Plain text and the editor's own clipboard (ProseMirror's) in E1. Paste from
   Word arrives as HTML with `mso-` styles (browsers do not expose Word's OOXML clipboard
   flavour): E2 converts paragraphs, headings by outline level, bold, italic, underline,
   lists and simple tables, and drops the rest, stating so; docx4j's ImportXHTML is the
   reference for what to keep. Paste of a WML fragment (from reveal codes, from another
   instance) as XML is the developer route and is E1. Recommendation: as above.
7. **Fonts in the browser.** The document's fonts are usually not installed. Phase B's
   `IdentityPlusMapper` gives the metric-compatible substitutes as `font-family` stacks
   (Carlito for Calibri, Liberation for Arial and Times). Fonts embedded in the document
   (`ObfuscatedFontPart`) can be deobfuscated (the XOR with the GUID that docx4j
   implements) and loaded with the `FontFace` API for the session, which is a real
   fidelity win for branded templates. Recommendation: substitutes in E1, embedded fonts
   in E2; the editor ships no fonts of its own.
8. **Unsupported content is reported, not hidden.** A document with opaque blocks (tracked
   moves, math, fields, altChunks, text boxes) shows a banner in Edit and Template
   ("12 elements shown read-only") linking to the Developer perspective's diagnostics
   that list them by kind and address. Recommendation: from E1; SuperDoc's V2 issue
   tracker is what happens when a reader drops things silently.
9. **Accessibility.** `contenteditable` gives screen readers the text; the editor adds
   roles and names to its panels, keyboard access to everything (the perspectives' rule
   that every command is in the palette), visible focus, and author colours that also
   differ by pattern for tracked changes and comments. Recommendation: from E1 as a
   review item per phase, not a later phase.
10. **Repository, packages, framework.** Arrangements considered:

    *Repositories.* (a) One repository `docx4j-editor` holding `@docx4j/editor-model`
    (the projection and commands, headless), `@docx4j/editor` (the UI) and
    `docx4j-editor-mcp` (E4's companion), depending on the published engine. (b) A
    TypeScript monorepo of objects, engine and editor: atomic cross-package changes and one
    CI, which this design session would have welcomed (four cross-repository CRs in two
    days), but the objects package is generated output with its own regeneration flow tied
    to the compiler, and the objects-versus-engine split deliberately mirrors docx4j's Java
    modules; folding them together trades a stated design for convenience. (c) The editor
    inside the engine's repository as a subpath: wrong, since ProseMirror, Yjs and Monaco
    have no place in a package that Office add-ins bundle, and the release cadences differ.
    (d) One repository per package: too fragmented for three packages that change
    together. Recommendation: (a), with the editor packages versioned in lockstep with each
    other and depending on engine minors; sibling checkouts as now until the engine is on
    npm.

    *Editor toolkit.* (a) ProseMirror directly (Appendix A). (b) Tiptap, a ProseMirror
    wrapper with React and Vue bindings, an extension system and ready-made menus; its
    collaboration extension is `y-prosemirror` underneath. It speeds up generic rich-text
    UIs, but its schema conventions are HTML-shaped and would be fought at every docx
    node (content controls as inline nodes with content, row-level attributes, the
    `sourceId` discipline), and some of its extensions are paid. (c) Lexical or Slate:
    Appendix A's reasons. Recommendation: (a); Tiptap's conveniences are for a different
    document model.

    *UI framework for the panels and perspectives.* (a) React with TypeScript and Vite:
    the largest ecosystem, Monaco and ProseMirror integrate through thin wrappers, and it
    is what most contributors know. (b) Vue: what SuperDoc's shell uses; comparable,
    smaller pool. (c) Svelte: the least framework in the bundle and pleasant for panels,
    smaller pool again. (d) Lit web components, or no framework: the editor embeds anywhere
    (an Office add-in task pane, the docx4j site, another application) without dragging a
    framework, which fits "complements Word" and the rule that state lives in Yjs and
    ProseMirror; the cost is building tree views, dialogs and a command palette by hand
    or from a component library, which is where React saves the most time.
    Recommendation: **React for the application, and the editing surface as a
    framework-free class (`DocxEditorView`: ProseMirror view, commands, Yjs binding,
    reveal-codes data) with a thin web-component wrapper `<docx4j-editor>`**, so that the
    surface embeds without React while the four perspectives' panels are built quickly in
    it. The rule stands that no application state lives in React; a panel is a view over
    the model, and the framework is replaceable at the panel layer.

    *Versioning.* Semantic versioning; the editor packages in lockstep with each other,
    depending on the engine by minor range.
11. **Testing.** Three layers: the model package's projection round trips over every
    fixture (export equals import for untouched documents; deep-equal after edits), run in
    Node; component tests of the commands against the model; Playwright end-to-end runs of
    the perspectives' main flows, with the Word acceptance checklist extended by the
    editor's outputs. Recommendation: the first layer before any UI.
12. **Licences of what is pulled in.** ProseMirror, prosemirror-tables, Yjs and its
    providers, Monaco, Sucrase, fflate, xpath, Paged.js if used: all MIT; nothing AGPL or
    GPL, and no fonts. Recommendation: a licence check in CI from E1, since Apache-2.0
    throughout is a stated differentiator.

### 11.1 Coordinates (decided 2026-09-12)

| | |
|---|---|
| GitHub | `plutext/docx4j-editor-ts`, following `docx4j-core-ts` and `docx4j-generated-objects-ts`; the bare `docx4j-editor` stays free |
| npm | `@docx4j/editor-model` (headless projection and commands), `@docx4j/editor` (the React application and `DocxEditorView` with the `<docx4j-editor>` web component), `@docx4j/editor-mcp` (the E4 companion; bin `docx4j-editor-mcp`, so `npx @docx4j/editor-mcp` runs it) |
| Node | `>=20` for the editor repository (Vite, Playwright, Monaco tooling; `Intl.Segmenter` is present from 16); CI on 20 and 22; the engine stays `>=18` |
| Web component tag | `docx4j-editor` |
| `app.xml` | `Application` = `docx4j-editor`; `AppVersion` per item 2's `XX.YYYY` rule |
| Custom XML namespace for optional session state and macros | `http://docx4java.org/editor/2026` |

## Appendix A. ProseMirror, for readers who have not used it

ProseMirror (Marijn Haverbeke, MIT, 2015 to date) is not an editor but a toolkit for
building one. The pieces this design leans on:

- **Schema.** The document's grammar: node types with content expressions (`paragraph:
  inline*`, `table: table_row+`), attributes with defaults, mark types (bold, a link) that
  decorate inline text, and per-type `parseDOM` / `toDOM` rules. The library refuses any
  document or edit that violates the schema, which is the property the projection depends
  on: a paragraph cannot end up inside a run, a row cannot escape its table.
- **Document model.** An immutable tree of `Node`s (`Fragment`s of children, text nodes with
  a set of marks). Immutability is what makes "unchanged since import" an identity test and
  what makes undo and structural sharing cheap: an edit rebuilds the path from the root and
  shares everything else. Positions are integers counting tokens in the flattened document,
  so a selection is two numbers, and every change comes with a `Mapping` that carries
  positions across it.
- **State and transactions.** `EditorState` is immutable too; a `Transaction` is a list of
  `Step`s (replace, add mark, set attribute) plus selection and metadata, applied to produce
  the next state. Steps are invertible and mappable, which is what history and, in
  ProseMirror's own `collab` module, operational transformation use. This design uses the
  transaction model but routes the document through Yjs rather than `collab` (Appendix B).
- **Plugins.** State fields, key bindings, `appendTransaction` (react to a transaction with
  another, the invariant-keeping hook used in 4.1), decorations, node views. `history`,
  `keymap`, `inputrules`, `prosemirror-tables` are plugins.
- **Decorations.** Rendering-only additions the document does not contain: the list-number
  widgets of section 5, the section-break lines of 3.1, the control gutters of 4.1, remote
  users' cursors. Because they are not in the document they never reach the export.
- **Node views.** A node type can take over its own rendering and event handling. The
  read-only regions (images, opaque blocks, headers and footers as separate views with
  `editable: false`) are node views showing a snapshot.
- **The view.** `EditorView` renders the document to a `contenteditable` DOM, listens to
  `beforeinput`, composition and selection events, and turns them into transactions. It is
  the part that absorbs browser differences and IME behaviour, and it is where the WG's
  EditContext would slot in later.
- **Why ProseMirror rather than Lexical (Meta), Slate or Tiptap.** Tiptap is ProseMirror with
  a component layer, usable here without changing the design. Lexical and Slate are sound
  but younger and without an equivalent of `prosemirror-tables` or of `y-prosemirror`'s
  maturity; SuperDoc's choice of ProseMirror for a docx editor is independent confirmation.

## Appendix B. Yjs, for readers who have not used it

Yjs (Kevin Jahns, MIT, 2015 to date) is a CRDT library: shared data structures that several
peers can edit concurrently and that converge without a central server deciding an order.

- **Y.Doc and shared types.** A `Y.Doc` holds named shared types: `Y.Text` (a sequence of
  characters with formatting attributes), `Y.Array`, `Y.Map`, and `Y.XmlFragment` /
  `Y.XmlElement` / `Y.XmlText`, a tree whose leaves are `Y.Text`s. The edit layer of this
  design is one `Y.XmlFragment` for the body and one `Y.Map` per content control's
  properties.
- **The algorithm.** YATA: every inserted item carries a unique id (client, clock) and the id
  of its left neighbour at insertion time; concurrent inserts at the same place are ordered
  deterministically by id, deletions are tombstones. The item store is a doubly linked list
  of runs (consecutive characters typed by one client form one item), which is why it is,
  in effect, a piece table whose pieces have identity, and why it is compact.
- **Updates and providers.** Every change is encoded as a binary update; applying the same
  updates in any order gives the same document. A provider ships updates between peers:
  `y-websocket` (a small relay), `y-webrtc` (peer to peer), `y-indexeddb` (local
  persistence); with no provider the document is single-user and loses nothing. **Awareness**
  is the side channel for cursors and presence, not part of the document.
- **Undo.** `Y.UndoManager` undoes a client's own changes even when others have edited in
  between, scoped to the shared types it is given; this replaces ProseMirror's `history`.
- **Snapshots and versions.** A `Y.Snapshot` is a cheap point-in-time marker; the document
  at any snapshot can be restored, which gives version history and "what changed since"
  without storing copies. Garbage collection of tombstones can be switched off while
  snapshots are needed.
- **y-prosemirror.** `ySyncPlugin` mirrors the `Y.XmlFragment` and the ProseMirror document
  both ways (a local transaction becomes a Yjs update; a remote update becomes a transaction),
  `yCursorPlugin` renders remote selections from awareness, `yUndoPlugin` wires the undo
  manager. Positions across peers use Yjs **relative positions**, which survive concurrent
  edits.
- **One consequence for the export rule.** When a remote update arrives, `ySyncPlugin`
  rebuilds the affected ProseMirror nodes, so node identity is not a reliable "unchanged"
  test in a co-editing session. The export therefore compares an imported node's content by
  value (`Node.eq`, structural, cheap) when identity fails, and identity is only the fast
  path. `sourceId`s live in node attrs, which Yjs carries, so they survive remote rebuilds.
- **Why Yjs rather than ProseMirror's `collab`.** `collab` is operational transformation
  with a central authority that orders steps; it needs a server from day one and offline
  editing is awkward. Yjs needs no server to be correct, works offline, and its update log
  is the persistence format. Automerge is the other mature CRDT; Yjs has the ProseMirror
  binding and the larger editor ecosystem.

## Appendix C. SuperDoc, and how this product compares

Sources: SuperDoc's repository (renamed from `superdoc-dev/superdoc` to `superdoc/docx-editor`;
npm `superdoc`), its V2 docs at docs.superdoc.dev and archived V1 docs, the shipped
`@superdoc/docx-engine` package's README, licence, typings and third-party notices, its issue
tracker, and a maintainer's post of 2026 quoted by the user. Researched 2026-09-11; what is
stated by the maintainers is marked as such, the rest is inference from public types and
code. The engine bundle was not deobfuscated (its licence forbids it).

### What SuperDoc is

A browser docx editor with a server side, by Harbour Enterprises. V1 (`superdoc@1.0.0`,
2025-12-09) was built on ProseMirror, Yjs and JSZip: a converter mapped docx to a ProseMirror
document and back. V2 (prereleases from 2025-12-19; first stable `2.3.0` on 2026-07-28;
`2.13.0` current on 2026-09-08) replaced that with, in the README's words, "an OOXML-backed
document model" that "reads progressively, renders bounded windows, runs without a browser
DOM, and synchronizes document content and package state through one collaboration model".
Features: view, edit, suggest, comment, tracked changes, tables, images, footnotes, fields,
math, TOC, content controls (text, rich text, date, checkbox, combo, drop-down, repeating
section and item, group, with lock modes and appearance), Word-faithful pagination; Node and
Python SDKs, a CLI and an MCP server over the same Document API.

**Licence.** The repository is AGPLv3 with a commercial licence, but as of V2 the core is not
open source: the `superdoc` package depends on `@superdoc/docx-engine`, whose README says
"DOCX Engine is proprietary software. It is not open source." The AGPL repository holds the
UI shell, the layout engine, the DOM painter and the Document API types; the parser, the
editing model, the input manager and the Yjs binding are in the closed engine.

### V2's architecture, as far as it is public

- **Model.** ProseMirror is gone entirely (the migration guide says the types describing its
  nodes and transactions "have no v2 equivalent because the model they described is gone";
  the engine's third-party notices list no prosemirror package). The public types show a
  typed tree of discriminated-union nodes carrying OOXML identities (`w14:paraId`,
  `w14:textId`) and **raw OOXML properties**; the contributing guide states the rule: "the
  importer stores raw OOXML properties and the style engine resolves them at render time.
  Resolving styles during import bakes them into node attributes and loses the original
  document intent on export." Addressing is per block: a block id plus a UTF-16 range, never
  a document-wide position; the docs say "do not derive mutation locations from rendered DOM
  nodes or copied text offsets". Not a piece table, not CRDT-native. A SAX parser in the
  bundle is the basis for reading "reads progressively" as streaming XML parsing.
- **Rendering.** A bespoke layout engine (open, in the repository): its own line breaking,
  justification, tab stops, columns, floats, footnote page flow and keep constraints, with
  canvas `measureText` as the measurement primitive, painted to plain DOM. "Bounded windows"
  is virtualisation of paginated output: five pages plus one of overscan by default, page
  shells kept for a stable scroll extent, page content hydrated as it comes into view.
  Layout is incremental with dependency tracking, and runs partly in Web Workers. A
  non-paginated "web" flow mode exists without headers, footers or page numbers.
- **Input.** `contenteditable` with `beforeinput` and composition events, an input manager of
  their own, hit testing against the layout. Not EditContext (the string does not occur in
  the bundle), not canvas input.
- **Collaboration.** Yjs is retained as a peer dependency, but the application can no
  longer supply its own `Y.Doc`: the engine manages it in a worker, and the room holds
  identified content units plus package state, validated by exporting and diffing parts.
  V1 rooms are not compatible; migration is an offline server-side step.

### The stated drivers, and what actually cost the time

From the README, the changelog of 2026-03-22, the maintainer's post and the issue tracker:

1. **A package is not one editor tree.** "A DOCX is a package of related XML parts,
   relationships, and assets rather than one editor tree"; "direct XML editing loses
   formatting; conversion approaches break comments."
2. **Headless.** "Server use required a simulated browser DOM"; one engine now backs the
   browser, the SDKs, the CLI and the MCP server, with the same result on each.
3. **Agents.** "LLMs are good at deciding what should change ... But actually modifying a
   `.docx` file is a different problem": the engine turns it "from a generation problem into
   a tool-use problem", with "explicit, testable document operations" (the post: "A person
   would rarely sit down to make 5,000 or 10,000 edits ... Agents make that kind of bulk
   work easy").
4. **Load time on large, footnote-heavy documents.** The post: the same 400-page,
   footnote-heavy document reaches `onReady` in about 12 seconds in V1 and under 2 in V2. The
   corroborating primary source is issue #3869 (opened 2026-08-05, closed 2026-09-09 by the
   maintainer with a video): on V1, 247 pages took about 18 seconds, about 75 ms per page,
   because "the layout pipeline measures every block and paginates the entire document, up
   to several convergence passes, before a single page is painted", with measuring at about
   82% of the time; building the ProseMirror document was about 30 ms for 913 blocks, and
   JSZip was not a factor. On V2 2.3.0 first paint was about 1.1 seconds and length-independent,
   but a 372-footnote document still took about 23 seconds to settle pagination; by 2.13.0
   "current V2 completes layout before displaying the document to keep it stable."
5. **One collaboration model** over content and package state, where V1 "needed a separate
   synchronization layer for the rest of the package".

Not cited as drivers anywhere: ProseMirror's schema being unable to express OOXML,
`contenteditable` or IME trouble, or pagination as such (V1 already had the same layout
engine and painter behind a hidden ProseMirror instance, and already virtualised to a few
mounted pages). The slow part of V1 was measuring and paginating, which is a cost of
pagination, not of ProseMirror.

### What those drivers mean for this design

1. **Load time is measured, not promised.** Reaching `onReady` for a 400-page document here
   means unmarshalling `document.xml` (the runtime's unmarshaller over a very large DOM),
   footnotes, styles and numbering; projecting to ProseMirror; building the Yjs document;
   and ProseMirror's eager render of the whole body. Mitigations, in order of cost:
   unmarshal the main part off the main thread in a worker (the tree is
   structured-cloneable apart from `PARENT`, which `linkParents` restores); project and
   render progressively, the first screens first, with lightweight placeholders for blocks
   not yet in view; unmarshal footnotes, comments and headers only when viewed; and, if
   needed, a virtualised view over the same ProseMirror state. The engine already loads
   parts lazily and copies untouched ones byte for byte, so the cost is confined to what
   is shown. A 400-page footnote-heavy fixture goes into `test/fixtures/` and a load
   benchmark into the editor's CI, with SuperDoc's under-2-seconds as the aspiration.
   **Decision (Jason, 2026-09-12): an aspiration, not a requirement.** The product's main
   purposes, the Developer and Template perspectives, are served whether or not that number
   is reached: templates are short, and developers work on fixtures. The benchmark guides
   the mitigations and is reported; it never gates a release.
2. **Bulk agent edits do not go through the editor.** An agent programs the content API
   against the package directly (or joins as a peer, E4) and the editor re-projects the
   result as one Yjs transaction, with snapshots before and after giving the reviewer a
   diff and tracked changes giving Word one. One engine cost follows: `search` and
   `insertText` recompute a paragraph's segments per call, fine for a person and quadratic
   for thousands of edits on one paragraph; a per-paragraph index invalidated on write is
   the fix, when a benchmark shows it.
3. **One engine, browser and server, self-hosted** is what this engine is by construction.
4. **Nothing in the stated drivers argues against ProseMirror as the view.** Pagination is
   not in this product's scope and the input layer is not named as a problem; what remains
   is point 1, which the benchmark measures.

### Known limitations of V2 (open issues, September 2026)

Chinese IME input loses characters in tables and in headers and footers (#3936); typing
latency doubles in paragraphs with an inline break, tab or footnote reference (#3984), and
a bold edit on a 482-page fixture spends 385 to 550 ms in a full-document layout (#3842);
text boxes lose their content on import where V1 kept it (#3956); a cluster of
complex-script and right-to-left gaps (`w:bCs`, `w:szCs`, `w:kern`, caret and tab-stop
order); tracked moves rendered as delete plus insert (#3999); creating a content control
fails in V2 collaboration (#3974); an undo can leave a room unable to export (#3991). The
V1-to-V2 migration guide classifies none of its changes as mechanical.

### What transfers to this design, and what does not

- **Driver 1 does not apply; this design already has the stronger form of it.** The typed
  object model with untouched parts round-tripped byte for byte, and changed elements
  rebuilt from a `deepCopy` of the original, is "the no-conversion boundary" by construction.
  SuperDoc had to rewrite because V1 genuinely converted docx to ProseMirror and back and
  ProseMirror was the source of truth; here ProseMirror is a projection and the tree is the
  truth. Their V2 export-diff machinery (`semanticChangedParts`, `protectedChangedParts`) is a
  weaker approximation of what the identity-and-`deepCopy` rule gives.
- **Driver 2 is already met.** The engine runs in Node and the browser with no DOM
  dependency of its own; nothing in the editor is needed for a server side.
- **Driver 3 applies and argues for what CR-002 already is:** a typed operation layer over
  the tree that agents call, addressed by stable per-block ids (`w14:paraId`, the ordinal
  addresses) and block-local ranges, never by a document-wide position. The rule to adopt
  verbatim from their runtime types: nothing outside the view speaks ProseMirror; positions
  are opaque handles. That is also what makes the Office add-in surface natural.
- **Driver 4 largely does not apply, because this product does not paginate.** The cost
  centre that made V1 slow was measure-all plus paginate-all; without pagination and canvas
  measurement, and with the browser laying out text, it vanishes. What remains at 400 pages
  is DOM node count and ProseMirror's eager render, which section "What those drivers mean"
  point 1 addresses with the worker, progressive projection and, if needed, a virtualised
  view. The benchmark stays; the alarm is lower than the post's numbers suggest.
- **Driver 5 applies and is the one to think about before E3 ships.** If Yjs holds only the
  body, then footnotes, section attributes, content control properties, media additions and
  later comments are out of band. The room schema is the hardest thing to change afterwards
  (SuperDoc's V1 rooms could not be migrated in place). Decision for E3: one `Y.Doc` holding
  the body fragment, a map of footnote fragments, a map of control properties, and a map of
  package-level changes (added media, section edits), with the docx as the carrier and the
  Yjs state stored in a custom XML part, as section 6 proposes.

### Lessons adopted

1. **Keep `contenteditable` and ProseMirror's input handling.** SuperDoc kept
   `contenteditable` and `beforeinput` even after leaving ProseMirror, wrote its own input
   manager, and shipped 2.x with IME broken in tables and nine open right-to-left caret
   issues. That is the price of owning selection and composition. ProseMirror's decade of
   that work is a reason to stay, not a risk.
2. **Resolve styles at render time; never bake them into the projection.** Their import rule
   is the same as this design's: node attributes hold the editable subset and a reference to
   the source element; effective formatting (Phase B) is computed for display and never
   written back. Section 3's table is amended to say so explicitly.
3. **Content controls around rows are a model problem, not a view problem.** Their painter
   has a dedicated `sdt` module with cross-page label state, and their type union carries
   repeating sections and items. Section 4.1's attribute representation keeps the control
   in the tree and projects it; that is the right side of the line.
4. **Invest the tests in the projection, not the parse.** V2's hand-written reader dropped
   text-box content that the V1 converter kept (#3956); a generated, schema-complete model
   with untouched-part passthrough is immune to that class of bug, so the risk here is the
   projection and its inverse. E1's test suite is round trips through the projection on the
   fixtures, with the byte-identical and deep-equal checks of CR-001 as the oracle.
5. **The closed-source ending matters for anyone weighing "just use SuperDoc".** The parts
   one would learn from or depend on are no longer inspectable; the layout engine is, and is
   the one component worth reading if pagination ever enters scope.

### Where this product differs

| | SuperDoc V2 | This design |
|---|---|---|
| Model under the editor | Its own OOXML-backed typed tree with raw properties, in a proprietary engine | docx4j's generated object model (94 modules) and the CR-001 packaging engine, Apache-2.0; the editor is a projection over it |
| Fidelity mechanism | Writes edits back to the package; export-diff validation | Untouched parts byte for byte; unchanged elements re-emitted; changed ones rebuilt from a `deepCopy` of the original; parity tests against docx4j Java (Phase B) |
| Rendering and input | Own layout engine, canvas measurement, DOM painter, five-page virtualisation; `contenteditable` with an own input manager | ProseMirror view over `contenteditable`; browser text layout; no pagination; virtualisation only if the benchmark demands it |
| Pagination | Word-faithful, in-engine | Out of scope; Word paginates |
| Content controls | Broad kind support; creation broken in collaboration today | The centre of the product: insertion, nesting, deletion, row-level controls, OpenDoPE tags, bindings to custom XML with XPath, binding preview |
| Scope | Word in the browser: review, comments, tracked changes, pagination | Complements Word: paragraphs, tables, sections, footnotes, content controls; everything else preserved read-only |
| Agents and server | Node and Python SDKs, CLI, MCP server over the Document API | The content API in Office JS shapes over the same engine in Node and the browser; `docx4j-mcp` as the server half |
| Office add-in | Not a target | A target: the API is a `Word.*` subset and the flat OPC container is `getOoxml` / `insertOoxml` |
| Collaboration | Yjs inside the engine's worker; no bring-your-own `Y.Doc`; V1 rooms not migratable | Yjs `Y.Doc` owned by the application, provider-agnostic; room schema decided in E3 with the package state included from the start |
| Licence | AGPLv3 shell; proprietary engine | Apache-2.0 throughout |

### What to take from SuperDoc

Their Document API and MCP tool surface as a checklist of what agents ask for; their
per-block addressing and opaque-position rule; their render-time style resolution; the
layout engine as reading if pagination is ever wanted; and issue #3869 as the reference
measurement for the 400-page benchmark. Do not compete on breadth: SuperDoc aims to be Word
in the browser; this product aims to be the docx tool Word users and OpenDoPE authors reach
for alongside Word, with a fidelity guarantee backed by the engine and its parity tests, and
with content controls as the thing it does better than anything else.

## Appendix D. Pagination: out of scope, and what stands in for it

### D.1 Page boundaries as data: `w:lastRenderedPageBreak`

Word writes a `w:lastRenderedPageBreak` item into the run at every point where it last laid
out a page break, so a document saved by Word already carries its own pagination as data
(the model types it as `R.LastRenderedPageBreak`). The editor shows these as page-boundary
decorations: a gap and a rule between the lines, a page number, and the section's header
above and footer below the boundary (D.2). No layout is computed. After an edit, the markers
in and after the edited paragraph are stale; they stay where they are and are drawn muted
("approximate, since your edit") until a **Re-paginate** action refreshes them. Word itself
rewrites them on its next save, so nothing about the file depends on the editor keeping
them right, and the export leaves them untouched except where Re-paginate replaced them.

**Re-paginate exists: docx4j CR-012, phases 1 to 3 shipped 2026-09-11** (`org.docx4j.model.pagination`
in docx4j-core; the `docx4j-mcp` tool is its phase 4, pending). `Paginate.paginate(pkg,
settings)` lays the document out through export-fo and Apache FOP's area tree and rewrites
the markers; `Paginate.compute` returns the `PaginationMap` (page index and FOP's formatted
page number per paragraph, and the offsets where later pages begin in a paragraph that
spans pages) without touching the document. What the implementation settled, and what the
editor must know:

- **Granularity is the line.** Break offsets inside a paragraph come from run anchors in
  the area tree (`lineBreaks` is on by default), and the writer splits the run at the
  offset, copying `w:rPr`, with the marker opening the second run before the first
  character of the new page's line, which is where Word puts it. So after Re-paginate the
  runs of a paginated paragraph differ from before: the projection re-imports those
  paragraphs (their identity test fails, as it should), and `segmentsOf` here already
  treats the marker as a non-text item, so the paragraph's text and ranges are unchanged.
- **Paragraph ids are assigned.** `paginate` gives every paragraph without a `w14:paraId`
  one (eight hex digits, below `0x80000000`, paired with `w14:textId` as Word writes them),
  because the markers were written against those keys. For the editor that is a gift: after
  the first Re-paginate every paragraph has the stable address of section 3.3.
- **The layout is of the accepted view of tracked changes.** Deleted content takes no
  space and insertions are laid out as text, which is what the editor shows (section 10,
  decision 2), so the markers and the display agree.
- **Word's marker rule, measured on its files:** the marker goes in the first run that
  starts on a page, including after an explicit page or section break; an empty paragraph
  starting a page gets a `w:r` to hold its marker (Word writes none there); a document
  opening with a break-only paragraph follows FOP's layout rather than Word's. The editor
  draws boundaries from the markers wherever they are and needs no rule of its own.
- **Accuracy is FOP's.** Fonts differ in metrics from Word's, a field result inside a run
  may shift a boundary by the length difference, and hyphenation is reconciled by matching
  the new line's first word. Good enough for "page 12 of 40" and for headers and footers at
  plausible places; not for a filing's line numbering. The editor labels boundaries from
  Re-paginate as computed, those from Word as Word's.

Documents without markers (created by this editor or by docx4j, or by producers that omit
them) show no boundaries until Re-paginate runs, or a single page frame around the section
when the user asks for one. A browser-side variant by measurement over the rendered DOM
(the Paged.js approach) remains an option for a no-server deployment and is not planned.

Recommendation: markers in E1 (a day), Re-paginate through `docx4j-mcp` in E2 once its
phase 4 lands, the browser variant only if a no-server deployment needs it.

### D.2 Headers and footers without pagination

With page boundaries as data, the section's header is shown at the top of each page's
content and the footer at the bottom, choosing the first-page, even or odd variant from the
section's `w:titlePg` and `w:evenAndOddHeaders` settings and the page number as Word would.
Without markers, each section shows its default header once above its content and its
footer once below, as a frame. In both cases they are the separate read-only views of
section 2, rendered by the same pipeline as the body, with `PAGE` and `NUMPAGES` fields
displayed from the marker count (or as their cached field results when there are no markers).
Editing them stays out of scope; a "Open in Word" affordance sits on the frame.

### D.3 What real pagination would take

A layout engine that paginates like Word must do, at least: line breaking with the actual
font metrics (which means having the fonts, or metric-compatible substitutes, and their
kerning), justification and hyphenation, tab stops, paragraph spacing rules including
contextual spacing and the auto-spacing of styles, `keepNext`, `keepLines`, widow and orphan
control, page and column breaks, sections with different page sizes, orientations, columns
and header and footer sizes, tables with row splitting, repeated header rows and `cantSplit`,
floating and anchored drawings with text wrapping, text boxes, footnotes and endnotes that
take space on the page they are referenced from (and may split), line numbering, and the
fields that depend on the result (`PAGE`, `NUMPAGES`, `PAGEREF`, the TOC). Word's own
behaviour is the specification and is undocumented in places, so parity is measured, not
derived.

Calibration points: SuperDoc's layout engine (Appendix C) is this list implemented in
TypeScript with canvas text measurement, incremental relayout and virtualisation, and it is
the component their V2 issue tracker is mostly about (typing latency from full-document
relayout, IME in tables, right-to-left caret order). It is AGPL and readable, which makes it
a good study and an impossible dependency for an Apache-2.0 product. Apache FOP is the same
engine class in Java, mature, and docx4j has driven it since 2008 (export-fo), which is why
the server Re-paginate above exists already in substance. A browser engine of that class
from scratch is months of work by someone who has done it before, plus fonts: without the
document's fonts, no engine paginates like Word, and shipping fonts is a licensing question
of its own.

This product does not need it. Its users edit templates and text and go to Word for the
page; showing Word's last pagination, refreshing it on request, and never pretending to more
precision than that is the honest and cheap position, and it keeps the editor's typing
latency independent of document length, which is the property SuperDoc's V2 is still
working to recover. If print fidelity ever becomes the product, the route is the FOP area
tree first (it exists: docx4j CR-012) and a browser engine last.
