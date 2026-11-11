# Word fragments: ED-005 proposal 46's Word check, part 6(a)

Built by `node packages/editor-model/test/fixtures/build/fragments-for-word.mjs` from the repository root. The XML is
written by hand, with no engine, so that nothing of the editor's or the engine's own writing is in what Word is asked
about. Run in Word 365 through the portfolio's Word bridge (Jason, 2026-11-11). Word for the Mac has no bridge. Each file
is opened and saved as docx and as PDF. The saves go beside the inputs as `46a-word365.docx`, `46a-word365.pdf` and so on.
The inputs carry a settings part in compatibility mode 15, as a new document in Word 365 has. The first run's inputs had
none, so Word opened them in mode 12 (Word 2007); its saves are `-word365-mode12` (run 1 below).

## The fragment

Every file carries the same fragment: a Flat OPC `pkg:package`, as Word writes one, with the XML and `mso-application`
declarations. It holds:

- a **Heading 1** in red at 20 pt;
- a paragraph in **FragmentNote**, green italic, a style the document lacks;
- a **numbered list** of two items, upper-case Roman, under `numId` 1;
- a sentence with a **footnote**;
- an **inline picture**, a blue PNG of 160 by 80 pixels shown at 1.6 by 0.8 inches;
- a two-by-two **table**;
- a **control** tagged `inner`;
- a final **section** that is landscape A5.

## The document

The document around the fragment defines:

- its own **Heading 1**, in blue at 14 pt;
- its own **list** under `numId` 1, decimal;
- an A4 portrait section.

The text, in order:

1. a paragraph saying which file it is;
2. a Heading 1;
3. list items one and two;
4. the control;
5. list item three;
6. "After the control."

Each file also has a data part, with the fragment escaped as the text of `/data/fragment`, and an XPaths part naming that
node `x1`.

## The files

| File | The control's content |
|---|---|
| `46a-altchunk.docx` | `<w:altChunk>` pointing at `word/chunk1.xml` (the fragment, `application/xml`): docx4j's output (`BindingTraverserXSLT.convertFlatOPC`) |
| `46b-altchunk-matchsrc.docx` | the same, with `w:altChunkPr/w:matchSrc` (keep the chunk's own formatting) |
| `46c-template.docx` | (1) the `od:progid` control holding its placeholder; (2) a plain text control mapped to `/data/fragment` by `w:dataBinding`; (3) a rich text control mapped to it by `w15:dataBinding` |

## What to read from Word's saves

The editor reads the PDF for what Word shows, and the docx for what Word wrote.

- **46a and 46b: does Word resolve an altChunk inside a block-level control?** The chunk inside `w:sdtContent` is
  docx4j's form, though the schema puts `w:altChunk` among the body's elements.
- **Heading 1:** does the fragment's heading draw blue at 14 pt (the document's style wins) or red at 20 pt? In the
  saved `styles.xml`, is there one Heading 1 or two, the second renamed?
- **FragmentNote:** is the style copied into the document's `styles.xml`, with its definition?
- **The lists:**
  - Are the fragment's items I and II, or numbered 3 and 4 in the document's decimal list?
  - Is the document's item three 3, or 5?
  - In the saved `numbering.xml`, is there a new `w:num` and `w:abstractNum`?
- **The footnote:** kept, as a footnote in the document's notes part, with its number?
- **The picture:** kept as a new image part?
- **The table:** kept, with its borders.
- **The control tagged `inner`:** kept as a control, with its tag and its `w:id`?
- **The section:** does the fragment's landscape A5 section come in as a section break, or is it dropped?
- **The control tagged `od:progid`:** kept around the content, with its tag?
- **46b against 46a:** what `matchSrc` changes.
- **46c (1):** the placeholder unchanged. Word knows nothing of `od:progid`.
- **46c (2):** the plain text control: does Word fill it with the escaped Flat OPC as text, as Word 2010 did with a mapped
  node holding Flat OPC (core-ts check 35)?
- **46c (3):** the `w15` rich text control: **does Word 365 fill it from the Flat OPC on open**, with the fragment's
  content? This answers the proposal's decision (d).

## Run 1: Word 365 in compatibility mode 12 (2026-11-11, bridge 22479f7)

The inputs as committed in `e2edbdc` had no settings part, so Word opened them in mode 12 (Word 2007). The saves are
`*-word365-mode12.docx` and `.pdf`.

- **How each file opened:** silently, completely, with no repair and no refusal. On saving 46c, Word's Compatibility
  Checker said "Rich text content controls will no longer be mapped", and the bridge answered Continue. Its screenshot is
  `46c-word365-mode12-compatibility-checker.png`.
- **46a (the chunk):** Word resolved the altChunk inside the control. No `w:altChunk` and no `aFChunk` relationship are
  left. The `od:progid` control stays around the fragment's twelve paragraphs and table, with its tag and `w:id`; Word gave
  its `w:sdtPr` an `w:rPr`.
  - **Heading 1 is the document's** (blue, 14 pt): there's one Heading 1 in `styles.xml`, and the fragment's paragraph keeps
    `pStyle` Heading1.
  - **FragmentNote is copied** with its definition, as are FootnoteText, FootnoteReference and Word's linked
    FootnoteTextChar.
  - **The list is new:** `w:num` 2 over a new `w:abstractNum` (upper-case Roman), its items I and II. The document's list
    goes on at 3.
  - **Kept:** the footnote, numbered 1 in the document's new notes part; the picture, as `media/image1.png`; the table; and
    the control tagged `inner`, with its tag and `w:id` 201.
  - **The landscape section is dropped:** the document has one section, A4.
- **46b (`matchSrc`)** is as 46a, except that the fragment's styles became direct formatting:
  - the heading has no `pStyle`, and carries the fragment's Heading 1's paragraph and run properties (red, 20 pt);
  - the FragmentNote paragraph is green italic directly, though the style is still copied;
  - the footnote reference is superscript directly.
- **46c (the template):**
  - (1) the `od:progid` control keeps its placeholder;
  - (2) the plain text control keeps its placeholder, not the escaped XML that Word 2010 showed in core-ts check 35;
  - (3) **the `w15`-mapped rich text control is filled from the Flat OPC on open**, reconciled as 46a's chunk was (the
    document's Heading 1, FragmentNote copied, a new list). Its `w15:dataBinding` is dropped on the save, which mode 12
    cannot hold.

## Run 2: Word 365 in compatibility mode 15 (2026-11-11, bridge 22479f7)

The same files, rebuilt with the settings part (`fa5a1dc`), so that 46c's mapping is asked of Word 365 as it is. The saves
are `*-word365.docx` and `.pdf`, and record compatibility mode 15.

- **How each file opened:** silently, completely, with no dialog and no problem. The Compatibility Checker stayed quiet.
- **46a and 46b:** the same as run 1 in everything the README asks. The only differences are which `w:abstractNumId` each
  list got (the same two lists, numbered the other way round) and the revision ids.
- **46c:** the `w15`-mapped rich text control is filled from the Flat OPC on open, with the whole fragment: twelve
  paragraphs, the table, the footnote reference, the picture and the inner control. **The save keeps its
  `w15:dataBinding`.** The plain text control keeps its placeholder and its `w:dataBinding`; the `od:progid` control keeps
  its placeholder.

So the compatibility mode changes nothing in how Word reconciles a fragment. It decides only whether a `w15` mapping
survives the save.

