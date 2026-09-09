# @docx4j/core-ts

docx4j-core for TypeScript: Open Packaging (docx, pptx and xlsx as zip, and the flat OPC packages
that Office JS `getOoxml()` returns), typed parts and relationships, and docx4j's style, numbering
and font resolution, over the Office Open XML object model of
[`@docx4j/generated-objects-ts`](https://github.com/plutext/docx4j-generated-objects-ts). Node,
browsers and Word add-ins.

Status: design stage. The engine is specified in
[CR-001](docs/change-requests/CR-001-engine.md); today the package only re-exports the object
model's facade.

```
npm install @docx4j/core-ts
```

The "docx4j-ts" line:

| Java | npm | Repository |
|---|---|---|
| `docx4j-generated-objects` | `@docx4j/generated-objects-ts` | [plutext/docx4j-generated-objects-ts](https://github.com/plutext/docx4j-generated-objects-ts) |
| `docx4j-core` | `@docx4j/core-ts` (this package) | [plutext/docx4j-core-ts](https://github.com/plutext/docx4j-core-ts) |

Runtime: [`@docx4j/jsonix`](https://github.com/plutext/jsonix).

## Development

Until the dependencies are on npm, install them from sibling checkouts (the objects package must be
built first, since a directory dependency is symlinked, not built):

```
(cd ../docx4j-generated-objects-ts && npm install --no-save typescript@5.6.3 ../jsonix/nodejs/scripts && npm run build)
npm install --no-save typescript@5.6.3 ../jsonix/nodejs/scripts ../docx4j-generated-objects-ts
npm run typecheck && npm test
```

## Licence

Apache-2.0, as docx4j. See NOTICE.
