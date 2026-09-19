// A Word add-in taskpane with two buttons.
//
//   Edit the selection   getOoxml() -> load the flat OPC package here -> one content API edit ->
//                        insertOoxml() it back. The engine runs inside the add-in; Word only
//                        hands over the markup and takes it back.
//
//   Run the same edit    the same function, over a sample package, through `@docx4j/core-ts`'s
//   without Word         `Word` shim, which is the API Office JS gives you with the `load()` and
//                        `sync()` lines removed. This is the path `node run-in-node.mjs` and the
//                        repository's test/examples.test.mjs exercise, so the add-in's logic is
//                        testable without Word at all.
//
// In a real add-in the imports are `from '@docx4j/core-ts'` and `from '@docx4j/core-ts/office-js'`.
import { WordprocessingMLPackage } from '@docx4j/core-ts';
import { Word as WordInProcess } from '@docx4j/core-ts/office-js';
import { edit, applyEdit } from './edit.mjs';

function show(message: string): void {
  const status = document.getElementById('status');
  if (status) status.textContent = message;
}

/** The add-in path: Word hands out the selection as a flat OPC package and takes one back. */
async function editSelection(): Promise<void> {
  await Word.run(async (context) => {
    const selection = context.document.getSelection();
    const ooxml = selection.getOoxml();
    await context.sync();

    const edited = await applyEdit(ooxml.value);

    selection.insertOoxml(edited, 'Replace');
    await context.sync();
  });
  show('Edited the selection.');
}

/** A small flat OPC package, standing in for `getOoxml()` when there is no Word. */
const SAMPLE_OOXML = `<?xml version="1.0" standalone="yes"?>
<?mso-application progid="Word.Document"?>
<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
  <pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml">
    <pkg:xmlData>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="/word/document.xml"/>
      </Relationships>
    </pkg:xmlData>
  </pkg:part>
  <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
    <pkg:xmlData>
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:body>
          <w:p><w:r><w:t>Quarterly report</w:t></w:r></w:p>
          <w:p><w:r><w:t>This draft is not final; the draft goes to legal.</w:t></w:r></w:p>
        </w:body>
      </w:document>
    </pkg:xmlData>
  </pkg:part>
</pkg:package>`;

/** The Node path, in the browser: the same `edit`, through the shim, over a package here. */
async function editWithoutWord(): Promise<void> {
  const pkg = await WordprocessingMLPackage.load(SAMPLE_OOXML);
  const highlighted = await WordInProcess.run(pkg, (context) => edit(context.document.body));
  show(`Edited ${highlighted} occurrences; first paragraph is now ${pkg.body.paragraphs[0]?.style}.`);
}

Office.onReady(() => {
  document.getElementById('edit-selection')?.addEventListener('click', () => void editSelection().catch((e: unknown) => show(String(e))));
  document.getElementById('edit-sample')?.addEventListener('click', () => void editWithoutWord().catch((e: unknown) => show(String(e))));
  show('Ready.');
});
