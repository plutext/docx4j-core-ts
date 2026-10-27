# Integer attribute probes and Word 365's saves (CR-017)

docx4j CR-065's and CR-066's Word probes, copied unedited from the share (`Office 2016/__tmp/cr065/`, `__tmp/cr066/`) with
`manifest.json`'s SHA-1 of each file. docx4j-25 built them (docx4j CR-065 phase 0 and CR-066 phase 0, 2026-10-26); Jason
Harrop opened each in Word 365 the same day and saved it (`<name>-word365.docx`), or Word refused it and there is no save.
docx4j's `docs/developer/change-requests/CR-065-integer-attributes-read-as-word-reads-them.md` section 3 and
`CR-066-integer-attributes-within-words-range.md` sections 3.1 and 3.4 (at `109a8d47b`) record what the saves show.

- `cr065/`: rounds 2 to 4, a decimal or an odd form in an integer attribute. `cr065-decimals-probe.docx` is the all-cases file
  (refused by Word), `cr065-structure-control.docx` the same with plain integers (opened clean), `cases.txt` the 65 cases, and
  `families/` the cases by family: B to I opened clean; A was refused and split into A1 to A7, A2 again into A2a to A2f; A1, A2a
  to A2d and A6 opened. Word truncates a decimal toward zero, reads `""` and the no-digit forms as 0, and refuses the whole
  document for whitespace, a comma, hex, an exponent or letters. The READMEs are docx4j-25's.
- `cr066/`: round 5, integers at and beyond each attribute's storage width: `groups/cr066-{P,B,R,T,S,N}.docx` and Word's saves;
  190 cases in `cases.txt`. Every file opened clean.
- `cr066/round6/`: round 6, the attributes round 5 left unmeasured or inferred: `groups/cr066-r6-{A,F,R,T,P,S,N,E}.docx` and the
  settings singles `singles/cr066-r6-D{1,2}-<n>.docx`, with Word's saves; 490 cases. Every file opened clean.

Word's save is the oracle: it writes Word's reading of each value, or leaves out a value that reads as the default. The family
and single files Word was never given stay on the share. `integer-attributes.test.mjs` holds this package's warnings to the saves.
