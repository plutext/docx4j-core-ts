// CR-002 section 65: a measurement setter writes a value only where Word 365 stores it as written. Word keeps each
// integer attribute in a field of its binary format's width, 16 bits for these, and reads a value past it wrapped, with
// no error or repair (docx4j CR-065 round 2, CR-066 rounds 5 and 6, run by Jason 2026-10-26): a left indent of 400,000 pt,
// 8,000,000 twips, opens as 4,608. So the setters refuse such a value, as Office JS refuses an argument it won't take.
// The width is the attribute's, not its simple type's (CR-066 section 3.1). A file's value is read and kept as written.
// Since section 65.3 the storage comes from the objects package's table (`WORD_INTEGER_STORAGE`, its CR-016), through
// `keptRange`, one source for the setters' guards and the out-of-range warnings (CR-017); what each row adds is the
// setters' own policy, a range narrowed below Word's storage where a value Word keeps is still a caller's mistake.
import { wordIntegerReading } from '@docx4j/generated-objects-ts/helpers/wml';
import { keptRange } from '../wordIntegers.mjs';
import { InvalidArgument } from './generalException.mjs';

/** How Word 365 stores an attribute a setter writes, and the range the setter writes within, in the attribute's unit. */
export interface WordStorage {
  /** The attribute, as the part writes it. */
  attribute: string;
  /** The attribute's unit, and how many make a point. */
  unit: 'twips' | 'half-points' | 'fiftieths of a percent';
  min: number;
  max: number;
  /** What Word reads a value as, once it is stored: the wrap, as CR-066 measured it (the objects package's `wordIntegerReading`). */
  reads(value: number): string;
  /** Measured by CR-066's rounds 5 and 6: every row is, since round 6 (section 65.6). */
  measured: boolean;
  /** The range Word keeps as written, the objects table's (`keptRange`), before the setter's policy narrows it. */
  stored: readonly [number, number];
  /** Why a value Word would keep as written is refused all the same: the range narrowed below Word's storage. */
  why?: string;
  /** The objects table's entry the row reads: `TYPE_NAME`, property, and the rule beside the attribute where one applies. */
  objects: { typeName: string; property: string; rule?: 'auto' | 'exact' | 'atLeast' };
}

const WML = 'org_docx4j_wml.';
const WORD_SPACING_LIMIT = "Word's own setter stops at 22 inches, 1584 pt (check 71b); its storage runs to 65535, which its save writes back as -1";

/**
 * A row from the objects table: the range Word keeps as written there (`keptRange`), narrowed by the setter's policy
 * where `min`, `max` and `why` say so; `reads` through `wordIntegerReading`. A row whose entry the table lacks, or
 * no longer measures, fails here, at load, so a table change announces itself.
 */
function row(attribute: string, unit: WordStorage['unit'], typeName: string, property: string,
  policy: { rule?: 'auto' | 'exact' | 'atLeast'; min?: number; max?: number; why?: string } = {}): WordStorage {
  const range = keptRange(WML + typeName, property, policy.rule);
  if (range === undefined || range === null) throw new Error(`wordRange: the objects table has no measured range for ${typeName}.${property}`);
  const name = (p: string): string => (p === 'hanging' ? 'a hanging indent' : 'a first-line indent');
  return {
    attribute, unit, measured: true, stored: range,
    min: policy.min ?? range[0], max: policy.max ?? range[1],
    ...(policy.why !== undefined ? { why: policy.why } : {}),
    objects: { typeName: WML + typeName, property, ...(policy.rule !== undefined ? { rule: policy.rule } : {}) },
    reads(value) {
      const r = wordIntegerReading(WML + typeName, property, value, policy.rule);
      if (r === undefined || r.value === null) return 'not known';
      if (r.attribute !== undefined && r.attribute !== property) return `${name(r.attribute)} of ${r.value}`;
      if (r.rule !== undefined && policy.rule !== undefined) return `${r.rule} ${r.value}`;
      return String(r.value);
    },
  };
}

/**
 * The attributes the content API's setters write, by Word 365's storage (the objects table; CR-066 sections 3.1 and 3.4),
 * narrowed where Word's own setter stops short of its storage (check 71b, CR-002 section 65.9, D2): there this package
 * refuses where Word does, and no earlier.
 */
export const WORD_STORAGE = {
  // int16, wraps (w:left measured by round 5, w:right by round 6); Word's own indent setters wrap too, and this package
  // refuses instead (section 65.5, D1)
  indLeft: row('w:ind/@w:left', 'twips', 'PPrBase.Ind', 'left'),
  indRight: row('w:ind/@w:right', 'twips', 'PPrBase.Ind', 'right'),
  // one signed 16-bit field, the binary format's dxaLeft1: a first-line indent as itself, a hanging indent as its negative,
  // so 0 to 32767 and 0 to 32768 (check 71: Word's firstLineIndent setter wrote 1638.4 pt as a hanging indent of 32768;
  // round 6, the file side: a first-line indent of 32768 reads as a hanging indent of 32768)
  indFirstLine: row('w:ind/@w:firstLine', 'twips', 'PPrBase.Ind', 'firstLine'),
  indHanging: row('w:ind/@w:hanging', 'twips', 'PPrBase.Ind', 'hanging'),
  // a list level's text indent, the same int16 field, but Word's own setLevelIndents refuses at 1584 pt either side
  // (check 71b: 1583.95 and -1583.95 written, 1584 and -1584 InvalidArgument), so 31679 twips either way
  lvlIndLeft: row('w:ind/@w:left', 'twips', 'PPrBase.Ind', 'left',
    { min: -31679, max: 31679, why: "Word's own setLevelIndents stops short of 22 inches, 1584 pt, either way (check 71b)" }),
  // uint16, wraps, and Word's save writes 0xFFFF back as -1; Word's own setters stop at 22 inches, 1584 pt (check 71b: 1584
  // written as 31680, 1584.05 InvalidArgument; a negative refused, the schema's type being unsigned too)
  spacingBefore: row('w:spacing/@w:before', 'twips', 'PPrBase.Spacing', 'before', { max: 31680, why: WORD_SPACING_LIMIT }),
  spacingAfter: row('w:spacing/@w:after', 'twips', 'PPrBase.Spacing', 'after', { max: 31680, why: WORD_SPACING_LIMIT }),
  // Office JS's lineSpacing, in points, is the line as a multiple of 12 pt ("in the Word UI, this value is divided by 12"):
  // Word's own setter writes w:line = 20 × points under w:lineRule="auto" (check 71b: 0.75 pt as w:line="15"
  // w:lineRule="auto", 1584 as 31680), and takes 0.75 to 1584 pt (0.7 and 1584.05 InvalidArgument); int16, wraps
  spacingLineAuto: row('w:spacing/@w:line', 'twips', 'PPrBase.Spacing', 'line',
    { rule: 'auto', min: 15, max: 31680, why: "Word's own setter takes 0.75 to 1584 pt, 15 to 31680 twips under w:lineRule=\"auto\" (check 71b)" }),
  // int16, wraps, under dxa and under pct; a width is never negative
  tcWDxa: row('w:tcW/@w:w', 'twips', 'TblWidth', 'w', { min: 0, why: "a cell's width is never negative" }),
  tcWPct: row('w:tcW/@w:w', 'fiftieths of a percent', 'TblWidth', 'w', { min: 0, why: "a cell's width is never negative" }),
  // uint16, wraps; Word's own setter takes 1 to 1638 pt, in half-points after rounding (check 71b: 0.9 written as w:sz
  // w:val="2", 1638.05 as 3276, 0.5 and 1638.4 InvalidArgument), where its storage runs to 65535
  sz: row('w:sz/@w:val', 'half-points', 'HpsMeasure', 'val', { min: 2, max: 3276, why: "Word's own setter takes 1 to 1638 pt, 2 to 3276 half-points after rounding (check 71b)" }),
} as const satisfies Record<string, WordStorage>;

const PER_POINT = { twips: 20, 'half-points': 2 } as const;

/**
 * `points` in the attribute's unit, rounded, where Word stores it as written and the setter's policy allows; else
 * `InvalidArgument`, naming the range and what Word would read, or why the policy stops short of Word's storage (CR-002
 * section 65; the policy section 65.9's). `where` is the setter (`Paragraph.leftIndent`). `units` is the value already
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
  // within Word's storage the value is kept as written, so the refusal is the setter's policy; past it, what Word reads
  const kept = value >= storage.stored[0] && value <= storage.stored[1];
  const reason = kept && storage.why !== undefined ? storage.why : `Word 365 would read it as ${storage.reads(value)}`;
  throw new InvalidArgument(`${where}: ${points} pt is ${value} ${unit} in ${storage.attribute}, outside the range written there `
    + `(${storage.min} to ${storage.max} ${unit}${inPoints}): ${reason} (CR-002 section 65)`);
}
