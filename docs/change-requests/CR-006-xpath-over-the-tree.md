# CR-006: XPath over the WordprocessingML tree: a read-only binder from marshalled nodes back to the objects

**Status:** Proposed 2026-09-27, at the editor's request; reviewed here the same day (section 4
rewritten against the runtime, section 8 added). Scheduled by Jason the same day with CR-002
section 30's release, and **implemented** 2026-09-27 (`7368f8c`, fixes from a review in CR-002
section 31), with section 5's docx4j oracle the same day: 46 XPath goldens from docx4j 17.2.1,
zero differences (section 9). Section 10's quadratic `pathOf`, found by the editor on 0.2.0, fixed
2026-09-28 for 0.2.1.
**Depends on:** CR-001 Phase A (`XmlPart.marshalToNode`); CR-002 phase E (`XPathEngine`,
`pkg.xpathEngine`, the default engine) and CR-005 phase A (`FontoXPathEngine`); a marshalling hook
below this package (section 4)
**Requested by:** `plutext/docx4j-ts-editor` ED-005 section 3.5 (E4.a: "Select by XPath" in the
Developer perspective, planned 2026-09-27). It does not gate the editor, which resolves hits to
blocks by position until a release carries this (section 6).
**Counterpart:** docx4j `JaxbXmlPart.getJAXBNodesViaXPath(xpath, refreshXmlFirst)` and
`MainDocumentPart.getJAXBNodesViaXPath`, over a JAXB `Binder<Node>`; `XmlUtils.getNodesViaXPath`.

## 1. Summary

CR-002 section 3.6 considered XPath over the WordprocessingML tree and left it out: "docx4j does
it by marshalling to a DOM and mapping nodes back through a JAXB `Binder`, which this runtime does
not have ... If demand appears, the route is a DOM view over the typed tree, its own CR." The
demand has appeared: the editor's Developer perspective selects by XPath
(`//w:sdt[w:sdtPr/w:tag/@w:val='x']`, `//w:p[w:pPr/w:pStyle/@w:val='Heading1']`,
`//w:r[w:rPr/w:rFonts]`), and its console wants the objects those select, so that a structural
replace is a line of script. The editor's agent's `find` tool is the likely next user.

Everything but one piece exists: `marshalToNode()` gives a part as a DOM, and `pkg.xpathEngine`
evaluates XPath 1.0 (or 3.1 through FontoXPath) over any DOM node. The missing piece is the way
back: a DOM node the XPath selected carries no link to the tree object it was marshalled from.
docx4j's `Binder` keeps that link in both directions and can write a DOM change back into the
objects; this CR proposes only the first direction, which is what selection needs and is the
cheap half.

## 2. The API

On `XmlPart`:

```ts
/** The objects an XPath selects in this part, through a DOM marshalled for the purpose. */
selectObjects(xpath: string, options?: { namespaces?: Record<string, string> }): Promise<Bound[]>;

interface Bound {
  /** The DOM node the XPath selected (in a snapshot: changing it changes nothing). */
  node: Node;
  /** The tree object of the element, or of the element owning the attribute or text selected. */
  object: unknown;
  /** Set when the node is an attribute: its qualified name on `object`. */
  attribute?: string;
  /** The node's path from the part's root, `/w:document/w:body/w:p[3]/w:r[2]`. */
  path: string;
}
```

On `Body` (and `Paragraph`, `Table`, `ContentControl` as contexts): `select(xpath, options?)`
returning the content-API view of each selected element where there is one (`Paragraph`, `Table`,
`TableRow`, `TableCell`, `ContentControl`; a `Range` over a `w:r` or a `w:t`) and the `Bound`
otherwise, with the context node the view's own element. A non-node result (a count, a string)
is `selectValue`'s business, as it is for custom XML parts today.

The namespaces default to the part's prefixes (`w`, `w14`, `r`, `wp`, `a`, `mc` and every other
the objects package declares), so the common expressions need no mapping. The expression goes to
`pkg.xpathEngine`, so its version is the consumer's choice (1.0 by default, 3.1 with FontoXPath).

Each call marshals afresh: the tree has no change log a cached DOM could be checked against, and
a stale answer is worse than a slow one. docx4j's `refreshXmlFirst` flag exists for the same
reason; here it is always true. The cost is one marshal of the part per call, which the editor
spends per query, not per keystroke; the implementation notes record it for the editor's
255-page document.

## 3. How

While `marshalNode` writes the DOM, every element it opens for a tree value is recorded against
that value, in a map from DOM element to object (a `WeakMap`, dropped with the DOM). A value that
the object model wraps as a name and a value (`JAXBElement`'s shape) maps its element to the
value. An attribute or a text node selected maps to its owning element's object, with the
attribute's name for an attribute. Nothing else changes in the marshal, so the bytes a part saves
are the same whether or not a select ever ran.

The alternative that needs no hook, a walk of the DOM and the tree in parallel matching elements
by name and order, breaks on `any` and mixed content and on the MCE preprocessor's work, and is
not proposed.

## 4. The hook below this package (read against the runtime, 2026-09-27)

The design holds, and the hook is smaller than the request supposed: **three lines in
`@docx4j/jsonix`, no signature changed.** Measured in `jsonix-factory.js` of 3.3.0.

The request reads as a callback at each site that opens an element, of which there are eight
(`writeStartElement` is called from `Marshalls.Element.marshalElement` and from the wrapper and
element-property paths at lines 3099, 3345, 3363, 3372, 3482 and so on), and most of those do not
have the tree value in hand. But one site downstream of all of them does:
**`Jsonix.Model.ClassInfo.marshal(value, context, output, scope)`** (line 2426) is called
immediately after the enclosing element is opened, whichever site opened it, and the open element is
`output.peek()` (line 1773). So:

```js
// Jsonix.Model.ClassInfo.marshal, first line
if (context.onElement) { context.onElement(output.peek(), value); }
```

records every class-typed value in the tree against its element, and nothing else needs touching.
Two riders:

- **`any` content is a *copy*.** DOM content goes through `output.writeNode(node)` (line 3511 and
  line 3761), which does `document.importNode(node, true)` (line 1761), so the element in the
  marshalled DOM is not the node in the tree and a `WeakMap` keyed on it would miss. `writeNode`
  returns the imported node, so those two call sites get the same one line, mapping the copy to the
  original: `var n = output.writeNode(value); if (context.onElement) context.onElement(n, value);`
  That is how a `Bound` over `sle:slicer` or a `w:drawing` inside a resolved Choice finds its object.
- **An element backed by a simple type has no object of its own** (its value is a string or a
  number held by the owning class), so it maps to the owner, as attributes and text nodes already
  do in section 2. Worth saying in the `Bound` doc comment rather than leaving the caller to find
  out.

**The hook is per-call, not global, and needs no new option plumbing in the facade.**
`marshalNode` already marshals through a *derived* context -
`Object.create(context, { namespacePrefixes: { value: ... } })` in `marshalToDocument` - precisely
so that the shared context is not mutated. An `onElement` goes on that same derived object, so two
concurrent marshals cannot see each other's callback, and `marshalNode(value, { onElement })` is a
one-line change in objects-ts over machinery that is already there.

**Nothing between jsonix and the DOM disturbs node identity.** `marshalNode` returns
`marshalToDocument(...).documentElement` and the only post-processing is
`fixRootNamespaceDeclarations`, which sets and removes attributes on the root; `XmlPart.marshalToNode`
then adds `declareIgnorablePrefixes`'s attributes. There is no serialize-and-reparse anywhere on
that path, which is what would have made the whole approach unworkable.

Where the request went: **jsonix**, as `plutext/jsonix` jsonix-CR-006, filed 2026-09-27 with
CR-002 section 27's unmarshalling hook as the other half - the same file, the same shape, one CR.
Its section 2 prototypes all four sites against 3.3.0 and reports what they answer, including the
`importNode` copy above, which is how that rider was found. objects-ts passes the option through
`marshalNode` when it lands; nothing here is scheduled either way.

## 5. Tests

The oracle is docx4j: a harness in `test/java` (as CR-005's goldens are made) runs
`getJAXBNodesViaXPath` over the main document parts of a fixture set with a list of expressions,
and records for each result the element's path and its local name. `selectObjects` over the same
fixtures and expressions must return the same paths in the same order, and each `object` must be
the tree object at that path (checked by walking the tree). Plus: `Body.select` returns views of
the right kinds; a select leaves `getXml()`'s bytes unchanged; attributes and text nodes map to
their owners.

## 6. What the editor does until then

ED-005 section 3.5: it evaluates the expression over `marshalToNode()` and maps a hit to a block
by its position under `w:body` (the ordinal address), which is exact for a block and cannot see
inside one past plain runs; the hit list says so. Its console's replace-by-XPath waits for this.

## 7. Not proposed

- **Writing back through the DOM** (`Binder.updateJAXB`): a change is made to the objects, through
  the content API or `setXml` on an element, never to the snapshot.
- **Other parts than XML parts**: the relationships and content types parts have their own API.

## 8. Review notes (2026-09-27)

Filed by the editor session from the installed 0.1.5 and CR-002 section 3.6, without reading the
marshalling code; this repository's review of it. The design stands - a read-only binder recorded
while marshalling, one marshal per call, nothing written back - and section 4 is rewritten with what
the runtime actually offers. Three things to settle before it is scheduled, none of them a blocker:

1. **The cost, measured rather than assumed.** Section 2 says "one marshal of the part per call,
   which the editor spends per query, not per keystroke", and records the 255-page document as the
   case to watch. That number should be measured before the API is offered, because it decides
   whether the DOM may be cached after all: if a marshal of a large `document.xml` is tens of
   milliseconds the note is right and no cache is needed; if it is hundreds, the editor's
   "Select by XPath" will want one, and then the staleness question section 2 dismisses has to be
   answered rather than avoided. `test/fixtures/` has nothing that size; the editor has the
   document.
2. **What `Bound.object` is for a wrapped value.** Section 3 says a value the model wraps as a name
   and a value "maps its element to the value", which is right for `{ name, value }` pairs, but the
   tree also holds `JAXBElement`-shaped entries inside `any` and `mixed` collections where the
   caller's `instanceof`-style test is against the *value*. The rule wants stating once, in the
   `Bound` doc comment: `object` is always what a walk of the typed tree would hand you at that
   path, never a wrapper.
3. **The oracle's reach.** docx4j's `getJAXBNodesViaXPath` runs over a `Binder`, so its answers are
   only defined for parts JAXB binds whole. A fixture whose part holds DOM content the model does
   not bind - `cr022-slicers-timelines.xlsx`'s drawings, the OMML of `omml.test.mjs` - will differ
   in kind, not degree, from docx4j's answer. Pick the fixture set for the harness accordingly, and
   record the DOM cases as this port's own tests rather than as parity goldens.

## 9. Implementation (2026-09-27)

Implemented as sections 2 to 4 describe, over `@docx4j/jsonix` 3.4.0's `onElement` (jsonix-CR-006)
through `@docx4j/generated-objects-ts` 0.3.0's `marshalNode(value, { onElement })` (objects CR-007).
`src/model/content/binder.mts` records the map and builds the `Bound`s; `XmlPart.selectObjects`,
`pkg.selectObjects` (the main document part by default) and `Body.select` are the API. Section 8's
three questions, answered:

1. **The cost** was measured by the editor on its 255-page main document part: 349 to 411 ms a
   marshal, against 975 ms for the first unmarshal, and a selection runs on Enter. No cache.
2. **`Bound.object`** is what a walk of the typed tree hands you at that path, never a wrapper, as
   the `Bound` doc comment says; for DOM content the model does not bind it is the node in the
   **tree**, not the snapshot's imported copy (section 4's first rider, which has its test on
   `cr022-slicers-timelines.xlsx`).
3. **The oracle's reach** stands as written: the fixtures whose main part holds what the two sides
   treat differently in kind (an `mc:AlternateContent`) are recorded without hits (section 9.1).

Departures from section 2, each deliberate:

- **`select` is on `Body` only**, not on `Paragraph`, `Table` or `ContentControl` as contexts, and
  the expression is evaluated over the whole part the body belongs to, not from the body's own
  element: `//w:p` from a header's body selects that header's paragraphs. A hit outside a sub-body
  (a cell's, a control's) comes back as a `Bound`, since the body has no view of it.
- **What `Body.select` maps to a view**, fixed after the review (CR-002 section 31, finding 10):
  the most specific view at any depth - `ContentControl` (inline ones included), `Table` (nested
  ones included), `TableRow`, `TableCell`, `Paragraph`; a `Range` over the paragraph for a `w:r` or
  anything in one, wherever the run sits; and the `Bound` for everything else, a `w:hyperlink`
  included. An attribute or a text node is always the `Bound`, so the caller keeps what matched.
- **`XmlPart.selectObjects` readies the engine itself**, on the snapshot it evaluates, rather than
  requiring a ready one (CR-002 section 31, finding 13). Whether an engine can evaluate a tree is a
  property of the document the tree is in (CR-002 section 21), and the snapshot is that document;
  readying on a marshal of its own cost `pkg.selectObjects` a second marshal of the part.

### 9.1 The oracle (section 5)

`test/java`'s `XPathHarness`, a second main class beside the parity harness, runs 47 expressions
over the main document part of every `.docx` fixture through docx4j's
`XmlUtils.getJAXBAssociationsForXPath` - what `getJAXBNodesViaXPath` is built on, without its throw
where a hit has no object - and records every hit's path, the kind of node, and the class of the
object docx4j's binder associates with it, spelled as a `TYPE_NAME`. `test/golden/xpath/` holds one
golden per fixture; `test/xpath-parity.test.mjs` runs the same expressions through
`pkg.selectObjects` and requires the same hits in the same order, each object of docx4j's class and
an object of the part's own tree (found by walking it, so not a copy). The expressions are the
structural ones (every paragraph, run, text, table, row, cell and control, the body's children),
what sits in a run (fields, tabs, breaks, a picture's `wp:inline`, `a:graphic` and `pic:pic`),
section 1's predicates, positional ones (`//w:p[last()]`, `(//w:p)[1]`), which are where an order
would show, and one attribute and one text selection.

**The result: zero differences**, 6,016 hits over 43 fixtures, on the first run. Checked not to
be vacuous by corrupting a golden three ways - two hits swapped, a run's class changed, a text's
class changed - each of which fails. Three things the harness had to decide, recorded in
`test/java/README.md`:

- **The binder is made from a fresh marshal.** docx4j's `getBinder()` binds the source XML when the
  part has not been unmarshalled, and a marshal of the tree when it has, so the harness unmarshals
  first. That is the same snapshot `selectObjects` evaluates, and what docx4j's `refreshXmlFirst`
  is for.
- **Section 8 item 3's fixtures.** Three fixtures - `loadAndSave`, `numbering-stories` and
  `numbering-stories-coretests`, each with a text box - marshal an `mc:AlternateContent` in their
  main part. docx4j keeps the element and both branches in its tree, so its marshal has the text
  box's paragraphs twice; this package resolves the choice before it unmarshals. They are recorded
  without hits, and the test lists them, so that a fixture joining them is a decision.
- **The departure of section 2, measured.** docx4j's binder associates nothing with an attribute or
  a text node (778 text nodes and 114 attributes here, all null); this package gives the owning
  element's object, and the test checks it is the object of the class docx4j gives that element.

The goldens are from the 17.2.1 release on Maven Central (`docx4j-17.2.1`, `0e8e7633e`) rather than
from `7fba7a150`, which the parity goldens record: the XPath and binder code has not moved in years,
a release needed no build, and `.github/workflows/parity.yml` now regenerates these with the parity
goldens every week from docx4j's head, diffing them the same way.

## 10. A defect the editor found: `pathOf` is quadratic in a parent's children (2026-09-27)

**Found by** `plutext/docx4j-ts-editor` ED-005 section 12.1 item 12 (E4.a's Select by XPath, on
0.2.0). `boundsOf` calls `pathOf` for every node, and `pathOf` counts the node's preceding
siblings of the same qualified name to write its position, walking from the parent's first child,
and does the same for every ancestor. A hit list of the 7,088 paragraphs of a 255-page document's
body walks up to 7,088 siblings for each: `//w:p[w:pPr/w:pStyle]` (6,972 hits) took 25 seconds in
Chromium, 21 of them in `pathOf` (a CPU profile of the select), where the marshal took 0.5 s and the
native evaluation 0.05 s; in Node over xmldom `boundsOf` took 161 ms for the same hits, which is
why the tests here never saw it.

The editor binds with its own copy until a release fixes it: the same `Bound`s and the same paths,
held to `pkg.selectObjects`' paths over every fixture by its `xpath.test.mts`, with the positions of
a parent's children counted once per call (a map per parent, filled on first need) and each
element's path kept for its descendants. That is the fix asked for here, in `boundsOf` or in a
`pathOf` that takes such a cache; a positional path is otherwise unchanged. With it the same select
takes 0.6 s in Chromium, the export and the marshal 0.53 s of that.

A second measurement, for whoever weighs the engines (not a defect of this package): FontoXPath's
evaluation of `//w:p[w:pPr/w:pStyle]` over that DOM takes 3.3 s in Node and about 25 s in Chromium;
the default engine takes 21.8 s in Node (the `xpath` package over xmldom) and 0.05 s in Chromium
(native `document.evaluate`). The editor's XPath tab therefore evaluates with the default engine
and offers XPath 3.1 as an option.

*(Fixed 2026-09-28 for 0.2.1, at Jason's request, as asked: `boundsOf` makes the paths of a hit list
through one cache, which counts a parent's children once, on first need, and keeps each element's
path for its descendants; `pathOf` on its own goes through a cache of its own. The paths are
unchanged - the XPath goldens hold every one of them. Measured in Node over xmldom, 30,000 siblings
took 4.2 s before and 14 ms after; the test selects that many and allows a second, which the old
code missed by three times and the new one meets seventy times over.)*
