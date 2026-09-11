# Change requests

| CR | Title | Status |
|---|---|---|
| [CR-001](CR-001-engine.md) | The engine: container interface, Open Packaging layer, typed parts, resolution utilities | Phase A implemented 2026-09-10; Phases B and C proposed |
| [CR-002](CR-002-content-api.md) | A content API in the shape of Office JS (`body.insertParagraph`, `Paragraph`, `Range`, `Font`, `Table`, `insertOoxml`) over the docx4j tree, assignable to a `Word.*` subset so add-in code runs on both sides; docx4j's helper names as aliases; `wml` fragments in the objects package; addresses and `outline()` for agents; WordApiDesktop 1.3's `CustomXmlPart` / `CustomXmlNode` / `XmlMapping` and typed content controls over XPath 1.0 (native or the `xpath` package) | Phases B and D implemented 2026-09-10; A proposed as objects CR-002; C and E proposed |
| [CR-003](CR-003-web-editor.md) | A web docx editor over the engine (design only, the editor is its own repository): ProseMirror over a Yjs document as the editable projection of the typed tree, export by identity-checked re-emission with `deepCopy` of the untouched originals, content controls for OpenDoPE, CSS from Phase B; what the engine must provide | Proposed 2026-09-11 |
