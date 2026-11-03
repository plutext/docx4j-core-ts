# Store-item and bound-date probes and Word 365's saves (CR-020)

docx4j CR-058's five probes, copied unedited from `docx4j-core-tests/src/test/resources/OpenDoPE/` at the 17.3.5 release
(`docx4j-17.3.5`, `64fae4cd0`), with Word 365's saves of them from the share (`Office 2016/bind-fun/`), and `manifest.json`'s
SHA-1 of each file. The probes are byte-identical to the copies Jason Harrop opened in Word 365 on 22 and 23 October 2026,
on a PC at UTC+11; the share names the two date probes `bind-typed-formats.docx` and `bind-typed-twelve.docx`.

- `bind-plain.docx`: one custom XML part `{...0001}`; B7's `w:storeItemID` `{...00FF}` names no part, and its XPath fits the
  one there. Before saving, something was typed into B1 (the share's `README.txt`, step 3), so Word's save shows B1's typing.
- `bind-plain-two-parts.docx`: two parts whose root the XPaths fit; B7's node is in both, B11's in part 2 only.
- `bind-plain-no-part.docx`: B7's XPath is in another namespace, B12's starts at an element no part has as its root.
- `typed-controls-dates.docx`: thirteen date controls DF1 to DF13, one `w:dateFormat` each.
- `typed-controls-twelve.docx`: six date controls TW1 to TW6, the twelve-hour clock at 12 and stored offsets.

Word's save is the oracle: what each control shows, and the `w:storeItemID` Word wrote. docx4j CR-058 section 8 records
the run. `store-items.test.mjs` holds this package's resolver and date formatter to the saves.
