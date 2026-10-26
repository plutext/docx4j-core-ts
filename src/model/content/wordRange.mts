// CR-002 section 65: a measurement setter writes a value only where Word 365 stores it as written. Word keeps each
// integer attribute in a field of its binary format's width, 16 bits for these, and reads a value past it wrapped, with
// no error or repair (docx4j CR-065 round 2, CR-066 round 5, run by Jason 2026-10-26): a left indent of 400,000 pt,
// 8,000,000 twips, opens as 4,608. So the setters refuse such a value, as Office JS refuses an argument it won't take.
// The width is the attribute's, not its simple type's (CR-066 section 3.1). A file's value is read and kept as written.
import { InvalidArgument } from './generalException.mjs';

/** How Word 365 stores an attribute a setter writes, and the range the setter writes within, in the attribute's unit. */
export interface WordStorage {
  /** The attribute, as the part writes it. */
  attribute: string;
  /** The attribute's unit, and how many make a point. */
  unit: 'twips' | 'half-points' | 'fiftieths of a percent';
  min: number;
  max: number;
  /** What Word reads a value as, once it is stored: the wrap, as CR-066 section 3.1 measured it (or infers it). */
  reads(value: number): string;
  /** Measured by CR-066's round 5; otherwise the note says what the range is inferred from. */
  measured: boolean;
  /** Why a value Word would keep as written is refused all the same: the range narrowed below Word's storage. */
  why?: string;
}

const int16 = (v: number): number => { const u = ((v % 65536) + 65536) % 65536; return u >= 32768 ? u - 65536 : u; };
const uint16 = (v: number): number => ((v % 65536) + 65536) % 65536;
const signed = (attribute: string, measured: boolean): WordStorage =>
  ({ attribute, unit: 'twips', min: -32768, max: 32767, reads: (v) => String(int16(v)), measured });

/** The attributes the content API's setters write, by Word 365's storage (CR-066 section 3.1). */
export const WORD_STORAGE = {
  // int16, wraps (w:left measured by round 5, w:right by round 6)
  indLeft: signed('w:ind/@w:left', true),
  indRight: signed('w:ind/@w:right', true),
  // one signed 16-bit field, the binary format's dxaLeft1: a first-line indent as itself, a hanging indent as its negative,
  // so 0 to 32767 and 0 to 32768 (check 71: Word's firstLineIndent setter wrote 1638.4 pt as a hanging indent of 32768;
  // round 6, the file side: a first-line indent of 32768 reads as a hanging indent of 32768)
  indFirstLine: {
    attribute: 'w:ind/@w:firstLine', unit: 'twips', min: 0, max: 32767, measured: true,
    reads: (v) => { const stored = int16(v); return stored < 0 ? `a hanging indent of ${-stored}` : String(stored); },
  },
  indHanging: {
    attribute: 'w:ind/@w:hanging', unit: 'twips', min: 0, max: 32768, measured: true,
    reads: (v) => { const stored = int16(-v); return stored > 0 ? `a first-line indent of ${stored}` : String(-stored); },
  },
  // uint16, wraps, and 0xFFFF is written back as -1; a negative is refused, the schema's type being unsigned
  spacingBefore: { attribute: 'w:spacing/@w:before', unit: 'twips', min: 0, max: 65534, reads: (v) => String(uint16(v)), measured: true },
  // as w:before (round 6)
  spacingAfter: { attribute: 'w:spacing/@w:after', unit: 'twips', min: 0, max: 65534, reads: (v) => String(uint16(v)), measured: true },
  // an exact line is stored negated, in 16 bits: the stored value's sign gives the rule back and its size the line, so
  // only 1 to 32768 reads as the exact line written (exact 65535 reads as atLeast 1, 0 as atLeast 0)
  spacingLineExact: {
    attribute: 'w:spacing/@w:line', unit: 'twips', min: 1, max: 32768, measured: true,
    reads: (v) => { const stored = int16(-v); return stored < 0 ? `exact ${-stored}` : `atLeast ${stored}`; },
  },
  // int16, wraps, under dxa and under pct; a width is never negative
  tcWDxa: { ...signed('w:tcW/@w:w', true), min: 0, why: "a cell's width is never negative" },
  tcWPct: {
    attribute: 'w:tcW/@w:w', unit: 'fiftieths of a percent', min: 0, max: 32767, reads: (v) => String(int16(v)), measured: true,
    why: "a cell's width is never negative",
  },
  // uint16, wraps; Word's UI stops at 1638 pt, its storage at 65535 half-points
  sz: { attribute: 'w:sz/@w:val', unit: 'half-points', min: 0, max: 65535, reads: (v) => String(uint16(v)), measured: true },
} as const satisfies Record<string, WordStorage>;

const PER_POINT = { twips: 20, 'half-points': 2 } as const;

/**
 * `points` in the attribute's unit, rounded, where Word stores it as written; else `InvalidArgument`, naming the range and
 * what Word would read (CR-002 section 65). `where` is the setter (`Paragraph.leftIndent`). `units` is the value already
 * in the attribute's unit, where the setter computes it otherwise (a percentage width's fiftieths).
 */
export function withinWord(where: string, points: number, storage: WordStorage, units?: number): number {
  if (!Number.isFinite(points)) throw new InvalidArgument(`${where}: ${points} is not a measurement in points`);
  const unit = storage.unit;
  const value = units ?? Math.round(points * PER_POINT[unit as keyof typeof PER_POINT]);
  if (value >= storage.min && value <= storage.max) return value;
  const inPoints = unit in PER_POINT
    ? `, ${storage.min / PER_POINT[unit as keyof typeof PER_POINT]} to ${storage.max / PER_POINT[unit as keyof typeof PER_POINT]} pt`
    : '';
  const read = storage.reads(value);
  const reason = read === String(value) && storage.why !== undefined ? storage.why : `Word 365 would read it as ${read}`;
  throw new InvalidArgument(`${where}: ${points} pt is ${value} ${unit} in ${storage.attribute}, outside the range written there `
    + `(${storage.min} to ${storage.max} ${unit}${inPoints}): ${reason} (CR-002 section 65)`);
}
