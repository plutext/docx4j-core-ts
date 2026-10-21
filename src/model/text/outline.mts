// CR-012 D3: Body.outline() and pkg.outline() read the text model. The paragraph text is the model's (hidden
// text left out, a deleted mark joined, the view's), each text box's paragraphs follow their anchor at
// `textbox:` addresses, flagged, and the options are the model's; the shape is the outline's own.
import { setOutlineReader } from '../content/outline.mjs';
import type { Body, Outline, OutlineParagraph, OutlineTable } from '../content/Body.mjs';
import type { XmlPart } from '../../parts/XmlPart.mjs';
import { Extractor } from './Extractor.mjs';
import { PackageTrees, TextContext } from './TextContext.mjs';
import type { TextOptions, TextOptionsInit } from './TextOptions.mjs';
import type { TextBlock, TextStory } from './TextResult.mjs';
import { contextFor, optionsOf, isPackage } from './Texts.mjs';

/** The story a body's blocks are in, from its address prefix. */
function storyOf(prefix: string): TextStory {
  const kind = prefix.split(':')[0];
  switch (kind) {
    case 'header': return 'HEADER';
    case 'footer': return 'FOOTER';
    case 'footnote': return 'FOOTNOTE';
    case 'endnote': return 'ENDNOTE';
    case 'comment': return 'COMMENT';
    default: return 'MAIN';
  }
}

/** The outline of blocks: each paragraph with its address, paraId, style and text, each table with its size; a cell's blocks after its table. */
export function outlineOfBlocks(blocks: readonly TextBlock[]): Outline {
  const paragraphs: OutlineParagraph[] = [];
  const tables: OutlineTable[] = [];
  const visit = (list: readonly TextBlock[]): void => {
    for (const block of list) {
      if (block.kind === 'PARAGRAPH') {
        const entry: OutlineParagraph = { address: block.address, style: block.style ?? 'Normal', text: block.text };
        if (block.paraId) entry.paraId = block.paraId;
        if (block.textbox) entry.textbox = true;
        paragraphs.push(entry);
      } else if (block.kind === 'TABLE') {
        tables.push({ address: block.address, rows: block.rows.length, cols: block.rows[0]?.cells.length ?? 0 });
        for (const row of block.rows) for (const cell of row.cells) visit(cell.blocks);
      }
    }
  };
  visit(blocks);
  return { paragraphs, tables };
}

setOutlineReader(async (body: Body, options: unknown, context: unknown): Promise<Outline> => {
  const o = optionsOf(options as TextOptionsInit | TextOptions | undefined);
  const pkg = body.package_ as unknown;
  if (!isPackage(pkg)) {
    // content with no package: read as it stands, as Texts.of(content) does
    const result = new Extractor(PackageTrees.none(), o, 'planless', new TextContext(PackageTrees.none())).extractContent(body.container, storyOf(body.prefix), body.prefix, undefined);
    return outlineOfBlocks(result.blocks);
  }
  const ctx = await contextFor(pkg, context as TextContext | undefined, o);
  const part = body.part as unknown as XmlPart<unknown> | undefined;
  const result = new Extractor(ctx.trees, o, 'view', ctx).extractContent(body.container, storyOf(body.prefix), body.prefix, part as never);
  return outlineOfBlocks(result.blocks);
});
