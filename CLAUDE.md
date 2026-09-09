# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

`@docx4j/core-ts`: the TypeScript counterpart of docx4j's `docx4j-core` module. It provides the Open
Packaging layer (packages, parts, relationships, content types; zip and flat OPC containers), the
typed parts (MainDocumentPart, StyleDefinitionsPart, ...) and the resolution utilities
(PropertyResolver, list numbering, font selection) over `@docx4j/generated-objects-ts`, the
generated Office Open XML object model (the counterpart of `docx4j-generated-objects`), and the
`@docx4j/jsonix` runtime. Design and scope live in `docs/change-requests/` (CR-001 is the engine).

The dividing rule with the objects package: anything that needs only an object tree (helpers,
the `XmlUtils`-style facade, flat OPC typing) lives there; anything that needs parts or
relationships lives here. This package re-exports the objects facade so both read the same.

## Commands

```
# Until @docx4j/jsonix and @docx4j/generated-objects-ts are on npm, install them from sibling checkouts
# (build the objects package first: a directory dependency is symlinked, not built)
(cd ../docx4j-generated-objects-ts && npm install --no-save typescript@5.6.3 ../jsonix/nodejs/scripts && npm run build)
npm install --no-save typescript@5.6.3 ../jsonix/nodejs/scripts ../docx4j-generated-objects-ts

npm run build       # tsc -p tsconfig.build.json: src/ -> dist/
npm run typecheck   # tsc --strict over src/ and test/*.ts (lib includes dom: the runtime typings need Node/Document)
npm test            # build, then node test/smoke.mjs
```

## Rules

- ES modules only (`"type": "module"`, `.mts` sources built to `dist/*.mjs` with `.d.mts`). Unlike
  the objects package there are no UMD files here, so `"type": "module"` is fine.
- Public paths are the `exports` map only. Keep them stable; add subpaths deliberately.
- Names follow docx4j (`OpcPackage`, `WordprocessingMLPackage`, `MainDocumentPart`,
  `RelationshipsPart`, `PropertyResolver`, `Emulator`) so docx4j Java code and documentation
  transfer; document deliberate departures in the CR that introduces them.
- Never edit `../docx4j-generated-objects-ts/modules/`; it is generated. A gap in the objects
  facade is fixed there (its own release), then the dependency range is raised here.
- Work is proposed as numbered change requests in `docs/change-requests/`; record decisions and
  implementation notes inside the CR. Commit only when asked.
