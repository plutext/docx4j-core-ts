# CR-005: An OpenDoPE processor in TypeScript, and an XPath 2 engine to evaluate it with

**Status:** Phase A implemented 2026-09-25 (section 7); phase B proposed; section 8 (escaped XHTML: the `html` module and the bind step, 2026-10-08) implemented, unreleased; Word check 37 pending
**Depends on:** CR-002 phase E (`ContentControl`, `XmlMapping`, `CustomXmlPartCollection`,
`DefaultXPathEngine`, `applyBindingsTo`); objects CR-003 phase A (`sdt`, `sdtProperty`,
`nextSdtId`, `walkAll`, `deepCopyAs`)
**Requested by:** `plutext/docx4j-ts-editor` ED-003 section 4 (E2.b, the Template perspective,
2026-09-25) and its section 9 items 10 and 13, accepted by Jason 2026-09-25. Neither phase gates
the editor's work: it evaluates verdicts itself and shows the instance document once phase B
lands.
**Counterpart:** docx4j `org.docx4j.model.datastorage` (`OpenDoPEHandler`, `OpenDoPEIntegrity`,
`OpenDoPEIntegrityAfterBinding`, `BindingHandler`, `RemovalHandler`, `XPathEnhancerParser`,
the `Docx4J.bind` facade), and the OpenDoPE Specification v3 WD 2026-09-18 (docx4j `docs/`),
which is the contract.

## 1. Summary

CR-002 phase E gave this package the binding step of the OpenDoPE processing model (the
specification's step 5: `applyBindingsTo` writes values into bound controls) and the parts it
needs. It did not give it the steps around that one: evaluating conditions and expanding repeats
with their contextualised expressions (step 3), repairing what those break (steps 4 and 6),
finishing, removing controls and parts (steps 7 to 9). Those are docx4j's `OpenDoPEHandler`
and companions, some 2,000 lines of Java, and they are what turns a template into an instance
document. Without them a TypeScript consumer can author a template and bind a value but cannot
produce a document from data; the editor's Template perspective shows verdicts ("this
condition is false", "this repeat gives three rows") where it would rather show the result,
and E4's companion, which serves the content API to an agent, has no `bind` tool to offer.

Both halves of the request are the engine's rather than the editor's for the same reason: they
are Apache-licensed, every Node consumer wants them, and the editor may not move code down
(its CLAUDE.md boundary rule). Two phases, because the first is a day and the second is weeks,
and the first is wanted by the editor's E2.b at once.

## 2. Phase A: a FontoXPath `XPathEngine`

The specification's boolean conversion modes (its section 7.2, table 7) are three: `java` (the
string equals `true`), `xpath1` (XPath 1.0 `boolean()`), `xpath2` (the effective boolean value,
`xs:boolean` cast accepting only `true`, `false`, `1` and `0`). `DefaultXPathEngine` evaluates
XPath 1.0 through the browser's `document.evaluate` or the `xpath` package; `xpath2` needs a
2.0 evaluator, and expressions in that mode may use 2.0 syntax (REQ-014's note). docx4j uses
Saxon.

[FontoXPath](https://github.com/FontoXML/fontoxpath) is an XPath 3.1 and XQuery 3.1 engine in
JavaScript, MIT, version 3.34.0, about 665 KB unpacked with two small dependencies, evaluating
over DOM nodes in browsers and in Node. XPath 3.1 is a superset of 2.0.

- `FontoXPathEngine implements XPathEngine` in `src/model/customxml/xpath-fonto.mts`, exported
  from a subpath (`@docx4j/core-ts/xpath-fonto`) so that the main entry does not import it;
  `fontoxpath` an **optional peer dependency**, as `xpath` is (CR-002 section 6 item 7's
  pattern), loaded lazily by `ready()`. `select` and `selectValue` over
  `evaluateXPathToNodes` and `evaluateXPath` with the namespace resolver built from the
  `namespaces` map; `selectValue` returns the XPath value in the four-type union the
  interface has.
- A `booleanValue(expression, context, namespaces, mode)` on the engine interface (added to
  `XPathEngine` with a default implementation in `DefaultXPathEngine` for `java` and `xpath1`
  and a throw for `xpath2`), so a caller asks for a boolean in a declared mode and the engine
  says whether it can. `FontoXPathEngine` implements all three (`xpath2` as
  `evaluateXPathToBoolean`, which is the effective boolean value).
- `pkg.xpathEngine = new FontoXPathEngine()` is the whole change for a consumer.
- Tests: the specification's table 7 examples in each mode; the CR-002 phase E XPath tests
  re-run over the Fonto engine; `invoice_Saxon_XPath2.docx` from docx4j's `sample-docs/
  databinding` evaluated and compared with what Saxon says in docx4j (a golden, as the parity
  harness makes).

Effort: one day.

## 3. Phase B: the processor

The specification's steps, in its order (REQ-015), over the main document part and every
header and footer related from it (REQ-016):

1. **Insert data** (step 1): replace the data part's content (`CustomXmlPart.setXml`).
2. **Preprocess** (step 3; specification sections 7 and 8): conditions evaluated in the declared
   mode through the engine (REQ-031 to REQ-034; markers for reverting optional, REQ-035 to
   REQ-037); repeats expanded (REQ-038 to REQ-044: the repeat base, the count, the tag rewrite
   to `od:rptd` and `od:RptOcc`, ids kept for the first instance and renewed for the rest,
   bookmarks renamed), expressions contextualised (REQ-045 to REQ-049, table 8, the
   `X_o_i` and `C_i` entries written to the XPaths and Conditions parts), position conditions
   (REQ-050, REQ-051), Word repeating sections (REQ-052 to REQ-055). docx4j's
   `XPathEnhancerParser` is the reference for contextualising; a port of its grammar rather
   than a regex.
3. **Integrity** (steps 4 and 6; specification 11.1): the repairs docx4j's `OpenDoPEIntegrity`
   and `OpenDoPEIntegrityAfterBinding` make.
4. **Bind** (step 5): `applyBindingsTo` as it is, plus the rich content the specification's
   section 9 defines where this package can (XHTML through `insertXml` of a converted fragment
   is out of scope until an XHTML importer exists here; Flat OPC through `insertOoxml`; pictures
   through `insertInlinePictureFromBase64`), each optional and reported when left unprocessed
   (REQ-076, REC-009).
5. **Finish, remove, delete parts** (steps 7 to 9; specification 11.2, 11.3): finishing takes a
   user function in place of XSLT; removal unwraps controls keeping content; the parts deleted.
6. **A facade**: `bind(pkg, data, { steps })` in the shape of `Docx4J.bind` and its flags.

Security (specification section 12): no external entities when parsing (REQ-073; the DOM
parsers here resolve none), cycles refused (REQ-029), missing conditions an error (REQ-030), a
version newer than 3.0 refused (REQ-010).

Tests: a `Docx4J.bind` harness in `test/java/` beside the parity one, with its goldens committed
and a workflow regenerating them against docx4j's head, so a difference arrives as a pull request
(section 6; Jason accepted this on 2026-09-25, and the editor copies the goldens rather than
making them). Agreed with the editor session, 2026-09-25:

- **Where they live: `test/golden/opendope/<template basename>.json`.** The editor's refresh
  script first assumed `test/fixtures/opendope/goldens/`; this repository's precedent is that
  goldens are `test/golden/` and fixtures are `test/fixtures/`, which CR-001 phase B settled, so
  the parity precedent wins and the editor has been told the path.
- **The shape**, which the editor's oracle test expects and this harness therefore writes:
  `{ "template", "data"?, "docx4j": <version>, "controls": { "<w:id>": { "value"?: string,
  "kept"?: boolean, "count"?: number, "unchanged"?: true } } }` - one file per template-and-data
  pair, recording what `Docx4J.bind` did to every content control after insert-XML, preprocess,
  integrity and bind, **with no removal step**, so a control is still there to be named. Any
  template and data pair in, a golden out, which is REQ-074 with docx4j as the judge.
- The pairs: `binding-simple`, `invoice` and `invoice2013` with their data files; `conventions`,
  `non-element-bind`, `hyperlink-binding-test`, `picture` and `invoice_Saxon_XPath2` against the
  data part they carry. Applying a data file means replacing the data part's XML before
  evaluating, which is step 1 of the pipeline, so the values line up either way.

Effort: two to three weeks, by the size of the Java.

## 4. Consumers

- The editor: phase A in E2.b's step B1 (its section 9 item 13); phase B replaces the
  Preview tab's verdicts with the instance document.
- E4's `docx4j-editor-mcp` companion: a `bind` tool.
- Node users of the content API: document generation from a template without Java, which is
  the package's stated purpose ("docx4j-core for TypeScript").

## 5. Open questions

1. Whether `booleanValue` belongs on `XPathEngine` or beside it (recommendation: on it, with the
   default engine implementing two of the three modes, so a caller learns what an engine can do
   from the engine).
2. Whether phase B lives in `src/model/datastorage/` as docx4j's package is named, or under
   `customxml/` with phase E (recommendation: `datastorage/`, since it is the processor and
   `customxml/` is the parts).
3. Reverting (specification 11.4) and components (section 10): both optional; recommendation:
   out of phase B, their own phase when a consumer asks.

## 6. Review by this package (2026-09-25)

The shape is right and phase A is ready to build as written. Four notes, one of which changes
phase B's estimate.

**Phase A's facts check out.** `fontoxpath` 3.34.0 is MIT, 664,772 bytes unpacked, with two
dependencies (`prsc`, `xspattern`). `XPathEngine` and the settable `pkg.xpathEngine` hook exist
as described, and `xpath` is already an optional peer, so the subpath plus optional-peer pattern
is the one CR-002 section 6 item 7 established rather than a new idea. A browser consumer who
never uses `xpath2` pays none of the 665 KB, which is the property that makes this acceptable
at all.

**Phase B is against about four times the Java the CR counts.** "Some 2,000 lines" understates
it: `OpenDoPEHandler` 1,796, `OpenDoPEIntegrity` 300, `OpenDoPEIntegrityAfterBinding` 217,
`BindingHandler` 491, `RemovalHandler` 379 - 3,183 - plus `XPathEnhancerParser` 2,458 and
`XPathEnhancerLexer` 2,292, so **7,933** in scope. Two to three weeks was estimated against the
smaller number and should be re-estimated before it is scheduled.

**`XPathEnhancerParser` is generated, so the thing to port is smaller than it looks and harder to
place.** It is ANTLR 3.3 output from `docx4j-core/src/main/antlr/XPathEnhancer.g`, 451 lines,
itself derived from a published XPath 1.0 grammar with `rewrite = true`. So the CR is right that
contextualising wants a grammar rather than a regex, but the artefact is the 451-line `.g`, not
the 4,750 generated lines, and it cannot be used as it stands: the TypeScript ANTLR runtimes
target ANTLR 4. Three ways, to be decided before estimating:

- convert the grammar to ANTLR 4 and generate with `antlr4ng` - closest to docx4j, and the
  generated code is not ours to maintain, but it adds a build step and a runtime dependency;
- hand-write a parser for the subset the rewrite needs (locations, predicates, the paths
  table 8 contextualises) - smallest dependency, and the one that can silently diverge;
- reuse **FontoXPath's** parser, which phase A already brings and which parses XPath to an AST.
  Tempting, and it would make the two phases share one XPath implementation - but it turns
  FontoXPath from an optional peer needed only for `xpath2` into a hard dependency of every
  consumer who binds a template, which is the opposite of phase A's virtue. Name the trade-off
  rather than drift into it.

**The goldens should be made here.** The CR offers the editor's Java oracle goldens as this
phase's tests. The discipline this package settled in CR-001 phase B is that the harness lives
in `test/java/`, the goldens are committed beside it, and `.github/workflows/parity.yml`
regenerates them weekly against docx4j's head so a difference arrives as a reviewable pull
request. A golden set produced by another repository's harness cannot be regenerated by that
workflow and will drift without anyone learning of it. So: the harness for `Docx4J.bind` over
`sample-docs/databinding` belongs in `test/java/` beside the existing one, and the editor
consumes it, not the other way round.

### Answers to section 5

1. **`booleanValue` optional on the interface, with a function beside it.** `XPathEngine` is
   public and a consumer may set `pkg.xpathEngine` to their own, so adding a *required* method
   breaks every external implementation at compile time. And two of the three modes need nothing
   an engine does not already have: `java` is `selectValue` as a string compared with `"true"`,
   `xpath1` is `selectValue` of `boolean(expr)`. So: `booleanValue?(expression, context,
   namespaces, mode)` **optional** on the interface, plus a module-level
   `booleanValue(engine, ...)` that answers `java` and `xpath1` over any engine, delegates to
   `engine.booleanValue` when it exists, and otherwise throws for `xpath2` naming
   `@docx4j/core-ts/xpath-fonto` in the message. Non-breaking, a third-party engine gets two
   modes for free, and the failure tells the caller what to install.
2. **`datastorage/`**, as proposed: `CLAUDE.md`'s naming rule is that names follow docx4j so its
   Java and documentation transfer, and `model/datastorage` is where this lives there.
3. **Out of phase B, as proposed, with one exception: write the reverting markers anyway.**
   Deferring the revert *operation* costs nothing, but if preprocess does not write the markers
   (REQ-035 to REQ-037), every document phase B produces is permanently unrevertable - the
   information is cheap while the template is in hand and impossible afterwards. Write the
   markers in phase B; implement reverting when a consumer asks.

## 7. Phase A implementation notes (2026-09-25)

Built as section 2 specifies, with the shape section 6 answered for: `booleanValue` is a
**function** over any engine plus an **optional** `booleanValue` member on `XPathEngine`, so no
existing implementation of that public interface breaks and a third-party engine gets the `java`
and `xpath1` modes for nothing. `FontoXPathEngine` is in `src/model/customxml/xpath-fonto.mts`,
exported from `@docx4j/core-ts/xpath-fonto`, with `fontoxpath` an optional peer beside `xpath`
and loaded lazily by `ready()`; the nodenext consumer check imports the new subpath. 496 tests.

**What table 7 turns out to mean, which is not what the CR assumed.** Section 2 proposed
`xpath2` as `evaluateXPathToBoolean`, "which is the effective boolean value". Measured, that is
wrong for the case the mode exists for - and, as the next paragraph records, it is not what docx4j
does either: FontoXPath's effective boolean value of `/invoice/wantspam`
where the element holds `false` is **true**, because the effective boolean value of a non-empty
node sequence is true - the opposite of table 7's "`"false"` is false", and a silent wrong answer
in a template processor, where a condition that should hide content would show it. Reading the
table's two sentences together gives the rule actually implemented: the effective boolean value,
**with a string cast to `xs:boolean`**. So a boolean result is itself (`/invoice/amt > 1000`); an
empty sequence and an empty string are false; any other value is taken by its string value and
cast, which accepts only `true`, `false`, `1` and `0` and raises an error otherwise -
`"yes"` is an error, as the table says, and so is `"True"`, `xs:boolean`'s lexical space being
case-sensitive where the `java` mode is not.

**No divergence from docx4j after all, and the implementation now shares its strategy.** The
paragraph that stood here supposed docx4j's `cast2` to be the plain effective boolean value, from
reading `XmlPart.xpathGetAsBoolean`. Jason asked the docx4j session to measure it instead
(Saxon-HE 9.9.0-2, docx4j-core 17.2.1-SNAPSHOT, 2026-09-25), and the supposition was wrong:
`XmlPart.cachedXPathGetBoolean` with `opendope.conditions.Xpathref.XPathBoolean=cast2` wraps the
**whole expression** in `xs:boolean(...)` before evaluating, so it is a cast of the atomized node,
not the effective boolean value of a node sequence. `cast1` is the effective boolean value;
`cast2` is table 7. So the specification, docx4j and this package agree, and FontoXPath's
`evaluateXPathToBoolean` - what section 2 proposed - would have implemented `cast1`.

**Counterpart:** `docx4j-core-tests org.docx4j.model.datastorage.ConditionBooleanSemanticsTest`
(docx4j `0d076bdde` on `VERSION_17_2_1`), nine tests pinning what was measured, of which the first
five are the eleven cases matched here: `cast2`'s lexical space, nothing-selected being false, the
four out-of-space texts and the two-node sequence being errors, and `typechecking` making no
difference on a bare path. The rest pin what this package does not attempt - the `cast1` and
Java-default contrasts, the legacy comparison shapes where `typechecking` does fire, and
`Xpathref.xpathEval` turning a cast error into an `InputIntegrityException`, which is the
abort-the-bind behaviour. When the two must be compared, that test is the oracle.

**Counterpart:** `docx4j-core-tests org.docx4j.model.datastorage.ConditionBooleanSemanticsTest`
(docx4j `0d076bdde` on `VERSION_17_2_1`), nine tests pinning what was measured, of which the first
five are the eleven cases matched here: `cast2`'s lexical space, nothing-selected being false, the
four out-of-space texts and the two-node sequence being errors, and `typechecking` making no
difference on a bare path. The rest pin what this package does not attempt - the `cast1` and
Java-default contrasts, the legacy comparison shapes where `typechecking` does fire, and
`Xpathref.xpathEval` turning a cast error into an `InputIntegrityException`, which is the
abort-the-bind behaviour. When the two must be compared, that test is the oracle.

This engine therefore does what docx4j does rather than casting by hand: `xs:boolean(${expression})`
evaluated by FontoXPath. All eleven of the measured cases match, and by construction rather than
by agreement, so the lexical space, the whitespace collapse, the multi-item rule and the empty
sequence cannot drift apart. Two of them had been got wrong by the hand-rolled cast and are fixed:
an **empty element** raises (`xs:boolean("")` is a cast error) where it had answered false, and a
**sequence of more than one node** raises where the first item had been taken. An expression
selecting **nothing** is false in both.

The one thing still unsettled is what a raise means. docx4j turns it into an
`InputIntegrityException` and the whole bind fails - measured: `wantspam` set to `yes` makes
`Docx4J.bind` throw, and now pinned by the last test of `ConditionBooleanSemanticsTest`. Here it is
a `Docx4JException` out of `booleanValue`, and what a processor does with it is phase B's decision;
the specification does not say whether a bad value is a failed document or a failed condition, and
the docx4j session agrees that is v3's line to write rather than either implementation's. Also measured, and unchanged: nothing in docx4j reads
`xpaths/@booleanConversion` (it exists only in `xsd/OpenDoPE/xpaths.xsd`), so the property applies
to every template and the v3 NOTE still holds at 17.2.1.

The `xpath2.typechecking` property is orthogonal, as read: it fires only in the catch, only when
Saxon's message contains `cannot compare xs:boolean to xs:string` **and** the expression contains
`=`, and rewrites only what follows the first `=`. It never fires for a bare path, so `cast2` and
`strict` answer a condition like `/invoice[1]/misc/wantspam` identically. It matters only for
legacy comparisons shaped `true() = 'true'`, which this package does not attempt to reproduce and
which phase B should not until a template needs it.

**A defect found in the `xpath` package, recorded per CR-001 section 19.** In Node the default
engine is the optional peer `xpath` 0.0.34, and it matches element names **case-insensitively**:
over a document holding `<t>` and `<T>` as distinct elements, `/invoice/t` answers both. A
browser's `document.evaluate` is correct, and so is FontoXPath, so the same template binds
differently in an add-in and in Node. `test/xpath-fonto.test.mjs` asserts the broken answer with
a `KNOWN` comment naming the package and version, so the fix announces itself; the fix is to
assert equality with FontoXPath instead. Worth reporting upstream, and worth knowing for CR-002
phase E users today.

**The module can be passed to the constructor** (2026-09-26, at the editor's request, ED-003
section 11.2 item 4). `ready()`'s dynamic `import()` of the bare specifier is resolvable in Node
and **not in a browser**: there is no import map, and the `@vite-ignore` that stops a bundler
carrying the optional peer against the consumer's will equally stops it resolving the specifier.
So a `./xpath-fonto` subpath that could not be used in a bundle was half of what phase A promised.
`new FontoXPathEngine(fontoxpath)` takes the module the application imported itself - its own
bundler resolving it statically - and is then ready at once, `ready()` resolving immediately;
`new FontoXPathEngine()` is unchanged for Node, still importing lazily and still naming the
package to install, now also naming the constructor form. Either the namespace object or the
default export is accepted, and anything else is refused **where it is given** rather than at the
first evaluation.

Two designs rejected, and why: a static `import` in the subpath module would satisfy bundlers but
turn a Node consumer's "npm install fontoxpath" into a bare `ERR_MODULE_NOT_FOUND` at import time
and make `ready()` a lie; an import map pushes this package's packaging problem onto every
application that consumes it. Taking the module keeps both audiences and leaves the loading
decision with the consumer, which is the principle `pkg.xpathEngine` being settable already
states.

**Tests** (`test/xpath-fonto.test.mjs`, 9): table 7's own examples across all three modes over
both engines; the default engine refusing `xpath2` with a message naming
`@docx4j/core-ts/xpath-fonto` and `fontoxpath` (REQ-032, evaluate in the declared mode or refuse);
the `java` default for a template that declares no mode; `select` and `selectValue` agreeing with
the default engine; an unready engine saying how to ready it; and docx4j's
`invoice_Saxon_XPath2.docx`, whose two conditions are the discriminating cases - `wantspam` holds
`false`, which is true in `xpath1` and false in `xpath2`, and `dateGt` is
`xs:date(/invoice/date) > xs:date('2018-12-31')`, XPath 2.0 syntax that the default engine cannot
evaluate in any mode, which is why the template is named for Saxon. **That fixture declares no
`xpaths/@booleanConversion`** (noticed by the editor session, ED-003 section 11.2 item 21), so
what it exercises about the engine is the 2.0 *syntax*, needed whatever the mode - a processor
reading this template would use its default, which is `java` both in the specification and in
docx4j's property - and not the `xpath2` cast, which is table 7's own test above. The test now
says so: it asserts the attribute's absence, and evaluates `dateGt` in all three modes, true
through this engine and a throw through the default one in every one of them. The last two cover the
constructor form: the module given either way answers as the imported one does, and a module that
is not `fontoxpath` is refused at the constructor. The nodenext consumer check exercises both
forms, so the declaration has to admit them.

## 8. Escaped XHTML: the HTML module, and the bind step (2026-10-08)

The specification's section 9 binds a node whose string value is XHTML markup, stored escaped as
the node's text, through a control whose tag carries `od:ContentType=application/xhtml+xml`. Item
4 of section 3 had put it out of scope "until an XHTML importer exists here". One does now, by the
editor's proposal (docx4j-ts-editor ED-005 proposal 44, accepted by Jason 2026-10-08) and his
decision as the copyright owner to move the editor's converter down under Apache-2.0 (its ED-001
decision 26): the converter is moved, never copied, and the editor imports it back from a release.

### 8.1 The `html` module (`@docx4j/core-ts/html`)

1. **`convert.mts`**: `convertHtml(html, lookup, { parser })` walks an HTML document to an
   intermediate form (`PasteBlock`: paragraphs with a style id, a list key and level, alignment,
   indents and space before, and inlines - runs with the six marks, a font, a size, a colour, a
   highlight, a character style or a link's address; tables with cells spanning rows and columns,
   widths in twips, a header row) and a report of what it kept and dropped, counted by kind.
   Written for Word's clipboard HTML and any XHTML; what it keeps and drops is the editor's
   ED-003 section 8.3 and its notes (`view/paste.ts`'s head). `clipboardMarkup` strips a
   `CF_HTML` header, `isTerminalHtml` tells a terminal's capture, `styleLookupOf(entries,
   tableStyleIds)` is the document's answer to a style name. **The parser is the caller's**: a
   browser's `DOMParser` by default; in Node, jsdom's or linkedom's (`linkedom`, ISC, is this
   package's test dependency). A bare fragment - the text between Word's markers, a data node's
   XHTML - is wrapped in `html` and `body` by a browser and not by a light DOM, so the walk's
   root is the first of the body, an `html` element and the document itself that holds an
   element (`rootOf`).
2. **`elements.mts`**: `blocksToXml(blocks, { numId, relId })` writes the form as a `w:p` /
   `w:tbl` fragment for `contentOf` or `Body.insertXml` (the prefixes undeclared, as those take
   it): run properties in the schema's order, a tab and a line break as runs of their own, the
   runs of one address under one `w:hyperlink` in the `Hyperlink` character style with no direct
   colour (as `Range.hyperlink` leaves a link), indents as `w:left`, `w:right`, `w:hanging` or
   `w:firstLine`, a table's grid from its first row's widths where every cell states one, else
   Word's text width shared equally, a cell spanning columns with its `gridSpan`, one spanning rows
   with `w:vMerge w:val="restart"` and a continuation cell in each row below (HTML lists no cell
   there), a header row's `w:tblHeader`. `needsOf(blocks)` lists what the document must give
   first - a definition per list key, a relationship per link address - and a list or link the
   resolvers do not answer is a plain paragraph or plain text.
3. **Tests** (`test/html.test.mjs`): the editor's corpus, fourteen captures from Word 15 and a
   terminal under `test/fixtures/html/` with their conversions pinned (`UPDATE_HTML=1`
   regenerates), unchanged by the move; every capture built, parsed by the engine and inserted
   into a new document, its paragraphs read back after a save; a list's numbering and a link's
   hyperlink given and withheld; vertical and horizontal merges and a header row. `tsconfig`'s
   `es2019` target has no iterable DOM collections: the walk uses `Array.from`.

### 8.2 The bind step (implemented 2026-10-08)

An XHTML-bound control carries **no `w:dataBinding`**: Word allows none on a block-level
control, and block content needs one, so docx4j's `bind.xslt` reads the XPath from the tag's
`od:xpath` entry (the XPaths part) and copies `w:sdtPr` as it is; Word, finding no binding, leaves
the content alone on open. The bind step handles a control whose tag says
`od:ContentType=application/xhtml+xml` and names an `od:xpath` entry: the entry's XPath over its
part gives the node, whose string value is the markup (unescaped by the XML parser already); the
converter and `blocksToXml` make the content with the document's styles as the lookup, the list
definitions made and the relationships added for what `needsOf` names; a block-level control is
given the blocks, a run-level one the first paragraph's inline content with the rest reported
(docx4j's `BindingTraverserXSLT.convertXHTML`, which refuses block content in a run-level control;
REQ-076, REC-009); the control's own run properties apply where the markup states none, as the
text binding's do. The reverse direction (`UpdateXmlFromDocumentSurface`) leaves such a control
alone, as docx4j's does. The parser comes through the bind's options; without one in a runtime
that has no `DOMParser`, the control is left and reported. What Word 2010 and 15 show for such a
file before and after the bind: check 37.

As built (`src/model/customxml/xhtml.mts`, `opendope.mts`; `test/html-binding.test.mjs`):

1. **`applyBindings({ html: { parser } })`** runs the XHTML pass after the text bindings, over the
   same controls. A control is XHTML-bound when its tag has `od:ContentType=application/xhtml+xml`
   and an `od:xpath`; its `xmlMapping.isMapped` is false, so the text pass never touches it.
   `opendope.mts` reads the tag's parameters (`tagParamsOf`) and the XPaths part's entries
   (`xpathsEntriesOf`, the first custom XML part in `http://opendope.org/xpaths`); phase B will
   read the rest of the parts. The result's new `notes` say what was left and why (no entry, no
   part, nothing selected, no parser) and what was dropped (the converter's counts; a run-level
   control's blocks after the first), one line each.
2. **The document first:** `styles.ensure` for the built-ins the markup names (a failure counted,
   not thrown), `numbering.newList` per list key (bullets each their own, numbered lists one
   definition then `restart`), a hyperlink relationship per address on the control's part with
   the `Hyperlink` style required at save, as `Range.hyperlink` does. The parsed fragment's
   elements are wrapped (`{ name, value }`), so the pass reads through the wrapper (`innerOf`).
3. **The content:** a block-level control takes every block; a run-level one (`SdtRun`, as
   `Range.insertContentControl` makes it; a paragraph's own `insertContentControl` wraps the
   paragraph as Office JS's does) takes the first paragraph's runs and hyperlinks. The control's
   `w:sdtPr/w:rPr` goes onto every run without properties of its own. `isShowingPlaceholder` is
   cleared. The package is the collection's `xhtmlHost`; a collection built without one leaves
   such controls and notes it.
4. **Held by:** the tag and the part read; a block-level control given a heading in its style,
   bold, a link with its relationship and a list with a definition, saved and reloaded with no
   `w:dataBinding` written; a run-level control given the first paragraph, the rest noted; a
   missing entry and a missing parser left and noted. The suite: 775.
