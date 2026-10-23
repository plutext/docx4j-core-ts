// Which tracked changes a listing returns and a resolve-all resolves (CR-011 section 3.2): docx4j's
// TrackedChangeFilter (its CR-043, org.docx4j.model.revisions) by author, type, kind, date and part,
// as an object that travels as JSON, or a function over a single change.
import { PartName } from '../../opc/PartName.mjs';
import type { Part } from '../../parts/Part.mjs';
import type { TrackedChange, TrackedChangeKind, TrackedChangeType } from './TrackedChange.mjs';

/**
 * The conditions of docx4j's `TrackedChangeFilter` (CR-011 section 3.2). Every condition given must
 * hold, and one not given holds for all. They apply to the single changes a group is made of
 * (`TrackedChange.pieces`):
 * - a listing returns a change, whole, when any of its pieces matches;
 * - a resolve-all resolves only the pieces that match, so a `kinds` filter can take a deleted
 *   paragraph's text and leave its mark (docx4j CR-043's U19).
 *
 * The range conditions of docx4j's (`within`, `withinBookmark`) are not here (CR-011 D2), and a key
 * this does not know is refused rather than ignored: a misspelt condition would otherwise resolve
 * every change.
 */
export interface TrackedChangeFilter {
  /** Changes by these authors (`w:author`), compared exactly. */
  authors?: readonly string[];
  /** Changes of these types: Office JS's `Word.ChangeTrackingState` names or docx4j's `Type` names (CR-011 D3). */
  types?: readonly ('Added' | 'Deleted' | 'Formatted' | 'ADDED' | 'DELETED' | 'FORMATTED')[];
  /**
   * Changes made at or after this (`TrackedChange.date`). A string is read by `Date`, so an ISO date
   * travels as JSON. A change with no date matches no date range.
   */
  from?: Date | string;
  /** Changes made before this; as `from`. */
  to?: Date | string;
  /** Changes of these kinds (`TrackedChange.kind`: a single change's, as `RUN_INSERTION`). */
  kinds?: readonly TrackedChangeKind[];
  /** Changes in these parts (`TrackedChange.part`); a string is a part name, as `'/word/footnotes.xml'`. */
  parts?: readonly (Part | string)[];
}

/** A filter object, or a function over a single change (a piece) that returns whether it is taken. */
export type TrackedChangeSelection = TrackedChangeFilter | ((change: TrackedChange) => boolean);

const KEYS = new Set(['authors', 'types', 'from', 'to', 'kinds', 'parts']);

const TYPES: Readonly<Record<string, TrackedChangeType>> = {
  Added: 'Added', Deleted: 'Deleted', Formatted: 'Formatted',
  ADDED: 'Added', DELETED: 'Deleted', FORMATTED: 'Formatted',
};

const KINDS: ReadonlySet<string> = new Set<TrackedChangeKind>([
  'RUN_INSERTION', 'RUN_DELETION', 'MOVE', 'PARAGRAPH_MARK_INSERTION', 'PARAGRAPH_MARK_DELETION',
  'RUN_PROPERTIES', 'PARAGRAPH_MARK_PROPERTIES', 'PARAGRAPH_PROPERTIES', 'ROW_INSERTION', 'ROW_DELETION',
  'FIELD_INSERTION', 'CONTENT_CONTROL_INSERTION', 'CONTENT_CONTROL_DELETION', 'MATH_STRUCTURE_INSERTION', 'MATH_STRUCTURE_DELETION',
  'INSERTION', 'DELETION', 'FORMATTING', 'TABLE_PROPERTIES', 'SECTION_PROPERTIES', 'CELL_INSERTION', 'CELL_DELETION',
]);

/**
 * Whether a single change meets a selection, as a function: the function form as it is, the object form
 * checked once here (a `TypeError` for a key or value it does not know), so that nothing is resolved on a
 * filter that was misread.
 */
export function pieceMatcher(selection: TrackedChangeSelection): (piece: TrackedChange) => boolean {
  if (typeof selection === 'function') return selection;
  if (typeof selection !== 'object' || selection === null) throw new TypeError('A tracked-change filter is an object or a function');
  for (const key of Object.keys(selection)) {
    if (!KEYS.has(key)) throw new TypeError(`Unknown tracked-change filter condition '${key}' (known: ${[...KEYS].join(', ')})`);
  }
  const authors = listOf(selection.authors, 'authors', (a) => typeof a === 'string');
  const types = listOf(selection.types, 'types', (t) => typeof t === 'string' && t in TYPES)?.map((t) => TYPES[t]!);
  const kinds = listOf(selection.kinds, 'kinds', (k) => typeof k === 'string' && KINDS.has(k));
  const parts = listOf(selection.parts, 'parts', (p) => typeof p === 'string' || (typeof p === 'object' && p !== null && 'partName' in p));
  const names = parts?.map((p) => (typeof p === 'string' ? PartName.of(p) : undefined));
  const from = dateOf(selection.from, 'from');
  const to = dateOf(selection.to, 'to');
  return (piece) => {
    if (authors && !authors.includes(piece.author)) return false;
    if (types && !types.includes(piece.type)) return false;
    if (kinds && !kinds.includes(piece.kind)) return false;
    if (from !== undefined || to !== undefined) {
      const date = piece.date?.getTime();
      if (date === undefined) return false;
      if (from !== undefined && date < from) return false;
      if (to !== undefined && date >= to) return false;
    }
    if (parts) {
      const part = piece.part;
      if (!part || !parts.some((p, i) => (typeof p === 'string' ? names![i]!.equals(part.partName) : p === part))) return false;
    }
    return true;
  };
}

function listOf<T>(value: readonly T[] | undefined, key: string, valid: (item: unknown) => boolean): readonly T[] | undefined {
  if (value === undefined) return undefined;
  if (!Array.isArray(value)) throw new TypeError(`The tracked-change filter's '${key}' is an array`);
  for (const item of value) if (!valid(item)) throw new TypeError(`The tracked-change filter's '${key}' has ${JSON.stringify(String(item))}, which it does not know`);
  return value;
}

function dateOf(value: Date | string | undefined, key: string): number | undefined {
  if (value === undefined) return undefined;
  const time = (value instanceof Date ? value : typeof value === 'string' ? new Date(value) : undefined)?.getTime();
  if (time === undefined || Number.isNaN(time)) throw new TypeError(`The tracked-change filter's '${key}' is not a date: ${String(value)}`);
  return time;
}
