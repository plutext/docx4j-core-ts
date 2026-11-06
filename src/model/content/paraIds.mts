// w14:paraId on what the content API adds. In a compatibility mode 15 file, one paragraph or row without an id makes Word 365
// renumber the whole part at the next open (check 79: 79d-table, 79f-one-missing, 79f-row-ids; check 80), which breaks a `w14:`
// address held across that open. Mode 12 files behaved otherwise (CR-002 section 66). Word's own saves hold an id on everything,
// so a new paragraph or row takes one wherever the story already uses them.
import { walk } from '@docx4j/generated-objects-ts/builders/wml';

const P = 'org_docx4j_wml.P';
const TR = 'org_docx4j_wml.Tr';

interface Identified {
  TYPE_NAME?: string;
  paraId?: string;
}

const holdsAnId = (value: object): value is Identified => {
  const type = (value as Identified).TYPE_NAME;
  return type === P || type === TR;
};

/** The top of the tree `within` stands in, up `PARENT`: the whole story, so a cell's own body sees the document's ids. */
function storyOf(within: object): object {
  let top = within as { PARENT?: object };
  for (let guard = 0; top.PARENT !== undefined && guard < 4096; guard++) top = top.PARENT as typeof top;
  return top;
}

/** The ids the story uses. A paragraph's own content isn't entered: what it holds matters only to a clash, one in two billion. */
function idsIn(story: object): Set<string> {
  const used = new Set<string>();
  walk(story, (value) => {
    if (!holdsAnId(value)) return undefined;
    if (value.paraId) used.add(value.paraId.toUpperCase());
    return (value as Identified).TYPE_NAME === P ? false : undefined;
  });
  return used;
}

function freeId(used: Set<string>): string {
  let id: string;
  do {
    id = Math.floor(Math.random() * 0x7fffffff).toString(16).toUpperCase().padStart(8, '0');
  } while (used.has(id));
  used.add(id);
  return id;
}

/**
 * Gives every paragraph and row in `added`, at any depth, a w14:paraId, where the story `within` stands in uses them (CR-002
 * section 66). A document that uses none stays so. `added` has to be in the tree already, or `within` be what it goes in.
 */
export function assignParaIds(within: object, added: Iterable<unknown>): void {
  const used = idsIn(storyOf(within));
  if (used.size === 0) return;
  for (const block of added) {
    walk(block, (value) => {
      if (holdsAnId(value) && !value.paraId) value.paraId = freeId(used);
      return undefined;
    });
  }
}
