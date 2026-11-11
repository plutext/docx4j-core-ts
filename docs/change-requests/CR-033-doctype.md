# CR-033: A DOCTYPE, read by this package and never by the platform's parser

**Status:** Decided and implemented 2026-11-11. Unreleased. Jason's direction was: "DOCTYPE as per
/home/jharrop/git/docx4j-portfolio/docs/dtd_handling_policy.md : refuse where the open xml requires refusal; for other parts,
guard against malicious entity expansion and other attacks". He then decided the four points that left open (section 2).
**Counterparts:**
- The portfolio's `docs/dtd_handling_policy.md` (draft of 2026-11-10; its D1, D2, D3, D4, D6, D7 and D8 settled here as section 2
  has them, its D5 waiting for the WD).
- docx4j-python's CR-031, which refuses every DOCTYPE.
- docx4j, whose StAX reader ignores a DTD.
- CR-031 here, whose `repair` decides the forbidden parts' reading.
- CR-032 here, whose short read and save rule the relationships part's repair uses.

## 0. Summary

Before this CR the parser, xmldom in Node, read a DOCTYPE in any part without a word:
- **Entities:** it expanded none. `&bp;` read as `&bp;`, and a save after unmarshalling wrote `&amp;bp;`, so the text changed.
- **Attribute defaults:** it applied none.
- **Fetching:** it fetched nothing.
- **A browser:** its `DOMParser` (libxml2 or expat) may expand internal entities, so the reading depended on the runtime, and a
  part built to expand without end met the platform's limits, not this package's.

Now:
- **Where ECMA-376 Part 2 forbids a DOCTYPE** (section 1), the package is refused at the load, naming the part.
  - Under `repair: 'word'` or `'all'`, a part's relationships part and the core properties are read as Word 365 repairs them, and
    a signature part as Word opens it.
  - The content types and `/_rels/.rels` are refused at every value, as Word refuses them after its repair prompt.
- **Everywhere else** the part is read at every value of `repair`. `readDoctype` takes the DOCTYPE out of the text before any
  parser sees it, and does what XML 1.0 §5.1 tells a non-validating processor to do:
  - internal entities are expanded, and internal-subset attribute defaults applied, within limits;
  - nothing outside the package is ever fetched.
  - `xml.doctype` names the part, what the DOCTYPE held, and Word's reading of such a file.
- **A save after an edit** writes every part that held a DOCTYPE afresh, without it, so the edited file opens in Word. An untouched
  save passes each through.

## 1. What the standard says, and what Word does

From the policy's sections 2 and 3, which cite the sources.

**ECMA-376 Part 2 (5th edition), §6.2.5 b):** "DTD declarations shall not be used in the XML markup defined in this
document". The parts it defines are the Media Types stream, the Core Properties part, Digital Signature XML Signature parts and
Relationships parts. Parts 1, 3 and 4 forbid nothing. XML 1.0 §5.1 has a non-validating processor expand internal entities and
apply internal-subset attribute defaults, and not process declarations after a parameter entity reference it didn't read.

**Word 365**, bridge rounds `dtd-word-check` and `dtd2-word-check` (10 November 2026, `test/doctype/`):

| Where the DOCTYPE is | Word 365 |
|---|---|
| The main part (bare, an empty subset, `ELEMENT`, `ATTLIST`, an entity unused or used, an external DTD), a header, a footer | Refused, "Word experienced an error" |
| Styles, settings, numbering, footnotes, endnotes, comments, the core properties, the custom XML store | Repaired: the part read as absent |
| The glossary, an embedded package's main part | Repaired: the content lost |
| `word/_rels/document.xml.rels` | Repaired: every relationship lost |
| `[Content_Types].xml`, `_rels/.rels` | The repair prompt, then refused |
| A digital signature part | Opens silently |

## 2. Decisions

| | Question | Decided (Jason, 2026-11-11) |
|---|---|---|
| D1 | Where a DOCTYPE is a fault | **Only where Part 2 forbids one** ("refuse where the open xml requires refusal"): the content types, every relationships part, the core properties, every signature part. The policy's D1 and D6 |
| D2 | Under `repair: 'word'`, in those parts | **Word's repairs, as the policy's table lists them** (1). A part's relationships part loses every relationship (read as none kept, CR-032's machinery), the core properties read as absent, a signature part read. The content types and `/_rels/.rels` stay refused at every value |
| D3 | Internal entities and attribute defaults elsewhere | **Expanded and applied within the limits** (2). The policy's D3, D4 and D7: nesting 8 deep, 10,000 references and 1 MiB of replacement text a part. Past one, or at an entity that names itself, the part reads as XML not well formed |
| D4 | A save after an edit | **Every part holding a DOCTYPE written afresh without it** (4), so Word opens the result. An untouched save passes each through. The policy's D8 |
| D5 | The policy's `doctype: 'word'` text-model option | **Waits for the WD** (3), so the three ports agree. The text model reads such parts as the default reading does, and `xml.doctype` stays off its warnings until the WD names it |

The engine's own:
- **D6. Nothing is fetched, whatever the option:**
  - an external subset (`SYSTEM` or `PUBLIC`) is named and left;
  - an external general entity's references read as nothing;
  - a parameter entity's reference stops later declarations being processed (§5.1).

  Each is listed in the warning's `notRead`. Guarding against the attacks is this rule plus D3's limits.
- **D7. A reference to an entity no declaration names reads as its own text** (`&nope;`), as xmldom read it before. A browser's
  parser would have refused the part, so this keeps the two runtimes alike.
- **D8. A default declaring a namespace (`xmlns`, `xmlns:p`) isn't applied.** It would change what every name below means.
- **D9. A signature part is left as it came by an edited save.** Its bytes are what it signs.
- **D10. "An edit" is a part written afresh:** one unmarshalled or given content, a DOM part adopted, or a part the source didn't
  hold. The save has no change log to ask, so a part read by `getBody()` and not changed counts. A removed part alone doesn't.

## 3. As built

- **`src/xml/doctype.mts`:**
  - `readDoctype(text)` finds the prolog's DOCTYPE, reads its internal subset (`ENTITY`, `ATTLIST`, `ELEMENT`, `NOTATION`,
    comments, PIs, parameter entity references), and expands the references in the rest of the text. That covers text and
    attribute values, but not comments, PIs or CDATA.
    - An attribute's replacement text has `<` and its quote escaped.
    - The DOCTYPE's lines are kept, so a parse error's line is the source's.
  - `applyAttributeDefaults` sets each default, in its prefix's namespace, on the elements of that name that lack it.
  - `hasDoctype`; `DOCTYPE_LIMITS`.
- **`parseXml(text, onDoctype?)`** reads every DOCTYPE through it. The platform's parser never sees one, in Node or a browser.
- **`XmlPart.parse`:**
  - `forbidsDoctype` (a relationships part's) and `readForbiddenDoctype` decide the parts Part 2 defines.
  - Every other part records `xml.doctype` (`DoctypeWarning`: `holds`, `expanded`, `defaulted`, `notRead`, `word`, `repaired`) first
    among its warnings, and `doctypeRead`.
  - `DefaultXmlPart` (a custom XML store) records it on `doctype`.
- **`RelationshipsPart.readForbiddenDoctype`:**
  - It refuses at `repair: 'refuse'`, and refuses `/_rels/.rels` at every value.
  - Otherwise it reads the part with no relationships, `part.not_read` with none kept (CR-032), so the parts only it named are
    absent, their Overrides go, and the save is CR-032's. The warning is `xml.doctype` with `repaired`.
- **`ContentTypeManager.parse`** refuses at every value.
- **`Load.mts`, `forbiddenDoctypes`:** the core properties and each signature part, by their bytes, so an untouched part stays so.
  The core properties are removed under `'word'`. `pkg.doctypes` holds their warnings.
- **`OpcPackage`:** `pkg.repaired` is `'word'` where a DOCTYPE was read as Word repairs it. `pkg.warnings` gathers every
  `xml.doctype`.
- **`Save.mts`, `Saver.dropDoctypes`:** D4 and D10, before the losses are settled. A part read afresh for it is never copied for
  its skipped content (CR-025), so what a rewrite would lose goes to the gate. `lastSave.doctypes` lists every part written
  without its DOCTYPE.
- `DoctypeException` (`package.doctype`) extends `PartUnmarshalException`.

## 4. Tests

`test/doctype.test.mjs`, 11 tests, over `test/doctype/`:
- Word's two rounds, copied unedited from `bridge/portfolio/<case>-a/` with a SHA-1 manifest: each input, Word's save and text
  where Word opened it, and the bridge's `result.json`.
- **The forbidden parts:** each refused at the default, naming the part; the content types and `/_rels/.rels` at every value. Under
  `'word'` and `'all'`, the core properties absent and saved without, and the signature read.
- **dtd-18 under `'word'`:** every relationship lost, Word's text exactly, and a save without the parts only those relationships
  named, as Word's save is.
- **Every other case**, 17 at two values of `repair`: read, one `xml.doctype` naming the part, and its `word` Word's reading in
  `result.json`. The glossary and the embedded package aren't met by the reading.
- **The subset as XML 1.0 reads it:**
  - dtd-06's entity expanded, and saved expanded with no DOCTYPE;
  - dtd-04's default applied (the bold off);
  - dtd-11's entity in a style's `w:jc`;
  - dtd-07's external subset named, not read.
- **The safeguards:**
  - the parser never handed a DOCTYPE, and the lines kept;
  - entities in content and in attributes, markup in replacement text, a character reference at declaration and use, and an
    undeclared entity;
  - the billion laughs past the depth, a wide expansion past the references, a long one past the replacement text, and a
    recursive entity, each refused, and a part over a limit refused naming the part;
  - external entities read as nothing, a parameter entity stopping later declarations, and an external subset with an
    internal one;
  - defaults: literal, `#FIXED`, `#IMPLIED`, a stated value kept, `xmlns` not applied.
- **The save:**
  - an untouched save passing four parts through (a header, styles, a store, footnotes);
  - an edited save writing each without its DOCTYPE and reading back alike;
  - the main part's;
  - a signature part kept.

Nine mutants, each caught:
- the DOCTYPE left to the parser;
- no depth limit; no reference limit;
- an external entity read;
- declarations after a parameter entity processed;
- a relationships part repaired at `'refuse'`;
- an edited save keeping the DOCTYPEs; an untouched save rewriting them;
- the core properties read at `'refuse'`.

## 5. Not done, or not measured

- **D5's `doctype: 'word'`**, and `xml.doctype` in the text model's warnings and the WD's REQ-801 list, wait for the WD.
- **Whether Word opens this package's edited save** of each case. Word's own saves drop every DOCTYPE, so it is expected, but not
  run. A bridge round of the edited saves would settle it.
- **A DOCTYPE in an embedded package's parts**, and in the theme, font table and web settings parts, isn't measured in Word. They
  are read as any other part is, `word: 'not measured'`.
- **The 2nd edition's M1.18** ("shall treat the presence of DTD declarations as an error") is the 5th edition's rule here too, for
  Part 2's parts only.
