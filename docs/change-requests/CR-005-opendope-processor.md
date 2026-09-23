# CR-005: An OpenDoPE processor in TypeScript, and an XPath 2 engine to evaluate it with

**Status:** Proposed 2026-09-25, two phases: A the FontoXPath `XPathEngine`, B the processor
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

Tests: the goldens the editor's Java oracle makes over docx4j's `sample-docs/databinding`
templates (ED-003 section 4.8: the instance document per template and data) become this
package's too, compared part by part after the removal step and text by text before it; the
parity harness pattern (a weekly workflow) keeps them current against docx4j.

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
