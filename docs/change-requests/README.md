# Change requests

| CR | Title | Status |
|---|---|---|
| [CR-001](CR-001-engine.md) | The engine: container interface, Open Packaging layer, typed parts, resolution utilities | Phase A implemented 2026-09-10; Phase B implemented 2026-09-19 (parity with docx4j VERSION_17_1_1 7fba7a150 on 45 goldens); Phase C proposed |
| [CR-002](CR-002-content-api.md) | A content API in the shape of Office JS (`body.insertParagraph`, `Paragraph`, `Range`, `Font`, `Table`, `insertOoxml`) over the docx4j tree, assignable to a `Word.*` subset so add-in code runs on both sides; docx4j's helper names as aliases; `wml` fragments in the objects package; addresses and `outline()` for agents; WordApiDesktop 1.3's `CustomXmlPart` / `CustomXmlNode` / `XmlMapping` and typed content controls over XPath 1.0 (native or the `xpath` package) | Phases B and D implemented 2026-09-10; A implemented as objects CR-002; C, G and I implemented 2026-09-15; E and F implemented 2026-09-16; H proposed (unblocked 2026-09-19 by CR-001 Phase B) |
