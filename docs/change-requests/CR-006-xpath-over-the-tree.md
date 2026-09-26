# CR-006: XPath over the WordprocessingML tree: a read-only binder from marshalled nodes back to the objects

**Status:** Proposed 2026-09-27, at the editor's request
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

## 4. The hook below this package

The marshal happens in `@docx4j/jsonix`, reached through `@docx4j/generated-objects-ts`'s
`marshalNode`. A callback on the marshalling context, called with the value and the DOM element
when a complex value opens its element, passed through `marshalNode(value, { onElement })`, is
the whole of it. Where that request goes (jsonix, objects-ts, or both) is this repository's call;
CR-002 section 27 asks for the matching hook on the unmarshalling side, and the two could be one
change there.

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
