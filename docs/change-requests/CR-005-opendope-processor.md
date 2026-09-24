# CR-005: An OpenDoPE processor in TypeScript, and an XPath 2 engine to evaluate it with

**Status:** Phase A implemented 2026-09-25 (section 7); phase B proposed
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
wrong for the case the mode exists for: FontoXPath's effective boolean value of `/invoice/wantspam`
where the element holds `false` is **true**, because the effective boolean value of a non-empty
node sequence is true - the opposite of table 7's "`"false"` is false", and a silent wrong answer
in a template processor, where a condition that should hide content would show it. Reading the
table's two sentences together gives the rule actually implemented: the effective boolean value,
**with a string cast to `xs:boolean`**. So a boolean result is itself (`/invoice/amt > 1000`); an
empty sequence and an empty string are false; any other value is taken by its string value and
cast, which accepts only `true`, `false`, `1` and `0` and raises an error otherwise -
`"yes"` is an error, as the table says, and so is `"True"`, `xs:boolean`'s lexical space being
case-sensitive where the `java` mode is not.

**A divergence from docx4j, deliberate.** docx4j evaluates both XPath modes with
`XPathConstants.BOOLEAN` (`XmlPart.xpathGetAsBoolean`), which over Saxon is the plain effective
boolean value, so a selected element holding `false` is **true** there. The specification v3 is
what this CR names as the contract and it is newer than that code - and its own NOTE says docx4j
17.2.0 does not read the `booleanConversion` attribute at all, so docx4j is not yet a conformant
v3 implementation on this point. Recorded rather than resolved here; when phase B's oracle
harness runs, this is the first place the two will disagree, and the question of which moves is
docx4j's.

**A defect found in the `xpath` package, recorded per CR-001 section 19.** In Node the default
engine is the optional peer `xpath` 0.0.34, and it matches element names **case-insensitively**:
over a document holding `<t>` and `<T>` as distinct elements, `/invoice/t` answers both. A
browser's `document.evaluate` is correct, and so is FontoXPath, so the same template binds
differently in an add-in and in Node. `test/xpath-fonto.test.mjs` asserts the broken answer with
a `KNOWN` comment naming the package and version, so the fix announces itself; the fix is to
assert equality with FontoXPath instead. Worth reporting upstream, and worth knowing for CR-002
phase E users today.

**Tests** (`test/xpath-fonto.test.mjs`, 7): table 7's own examples across all three modes over
both engines; the default engine refusing `xpath2` with a message naming
`@docx4j/core-ts/xpath-fonto` and `fontoxpath` (REQ-032, evaluate in the declared mode or refuse);
the `java` default for a template that declares no mode; `select` and `selectValue` agreeing with
the default engine; an unready engine saying how to ready it; and docx4j's
`invoice_Saxon_XPath2.docx`, whose two conditions are the discriminating cases - `wantspam` holds
`false`, which is true in `xpath1` and false in `xpath2`, and `dateGt` is
`xs:date(/invoice/date) > xs:date('2018-12-31')`, XPath 2.0 syntax that the default engine cannot
evaluate in any mode, which is why the template is named for Saxon.
