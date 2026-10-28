// CR-017: what Word 365 keeps of an integer attribute, from the objects package's table (its CR-016, docx4j CR-066's
// rounds 5 and 6, run by Jason 2026-10-26). Word stores each attribute in a field of its binary format's width, 16 bits
// for most, and reads a value past it wrapped, sometimes then clamped, with no error or repair. The table gives the
// width, the sign and the rules per attribute; this module turns an entry into the range of values Word keeps as
// written, which the setters' guards (CR-002 section 65.3, `wordRange.mts`) and the out-of-range warnings (CR-017 P2
// and P3, `XmlPart.valueOutOfRange`) both read, so there is one source for both. A report names its entry (`typeName`,
// `property`, `storage` and the `rule` beside the value, objects 0.7.0), so `rangeOf` takes it straight.
import { WORD_INTEGER_STORAGE, wordIntegerReading, type WordIntegerReading, type WordIntegerStorage } from '@docx4j/generated-objects-ts/helpers/wml';

/** `[min, max]` in the attribute's own unit: every value in it Word reads as written, and none outside. */
export type KeptRange = readonly [number, number];

/** The certain range of an attribute the table has but no round measured: outside it every measured attribute of its type changes. */
export const UNMEASURED_RANGE: KeptRange = [-32768, 65535];

/** Whether a reading is of an attribute no round measured (`storage` says so), as against one Word reads to no number. */
export const isUnmeasured = (reading: WordIntegerReading): boolean => reading.storage.startsWith('unmeasured');

/**
 * The range Word 365 keeps as written in an attribute of the table (CR-017 E1's `range`; CR-002 section 65.3's bounds):
 * - a signed field `-2^(bits-1)` to `2^(bits-1) - 1`, an unsigned one `0` to `2^bits - 1`;
 * - `w:ind`'s `firstLine` and `hanging`, one signed field, `0` to `32767` and `0` to `32768`;
 * - a line or a row height whose sign is the rule: under `exact` `1` to `32768`, under `atLeast` `0` to `32767`, a line
 *   under `auto` the plain signed field; a row height under `hRule="auto"` is ignored, so null;
 * - a frame's height `0` to `32767` (15 bits);
 * - then the clamps (a tab stop, the page size), and a column's width or space, whose negative Word replaces (a
 *   computed width; -1, which it also keeps as written, its own sentinel).
 * The rule is the `lineRule` or `hRule` beside the attribute, where it has one. Undefined for an attribute not in the
 * table. 65535 in an attribute Word's save writes as -1 (`w:spacing`'s `before` and `after`, the unsigned page margins)
 * is in the range, and Word reads a -1 there as 65535 too, its own spelling.
 */
export function keptRange(typeName: string, property: string, rule?: string): KeptRange | null | undefined {
  const entry = WORD_INTEGER_STORAGE[typeName]?.[property];
  return entry === undefined ? undefined : rangeOf(entry, property, rule);
}

/** `keptRange` from the table's entry; `property` for the shared field, `rule` for a line or a height. */
export function rangeOf(entry: WordIntegerStorage, property: string, rule?: string): KeptRange | null {
  if (!entry.measured) return UNMEASURED_RANGE;
  const half = 2 ** (entry.bits - 1);
  let min: number;
  let max: number;
  if (entry.heightFlag) [min, max] = [0, half - 1];
  else if (entry.pair) [min, max] = property === entry.pair.negative ? [0, half] : [0, half - 1];
  else if (entry.signRule) {
    const read = rule ?? (entry.signRule === 'hRule' ? 'atLeast' : 'auto');
    if (entry.signRule === 'hRule' && read === 'auto') return null;
    if (read === 'exact') [min, max] = [1, half];
    else if (read === 'atLeast') [min, max] = [0, half - 1];
    else [min, max] = [-half, half - 1];
  } else if (entry.signed) [min, max] = [-half, half - 1];
  else [min, max] = [0, 2 ** entry.bits - 1];
  if (entry.negative !== undefined) min = Math.max(min, 0);
  if (entry.min !== undefined) min = Math.max(min, entry.min);
  if (entry.max !== undefined) max = Math.min(max, entry.max);
  return [min, max];
}

/** Whether a reading's storage string names a rule beside the value (a line's `lineRule`, a height's `hRule`). */
const dependsOnRule = (storage: string): boolean => /\brule\b|clearing exact/.test(storage);

/**
 * Word's reading, for a message: the value, in the shared field's other attribute where the sign moved it (`25536 in
 * w:hanging`), with the rule where the sign carries one (`atLeast 1`, `exact 32768`); `not known` for an unmeasured
 * attribute.
 */
export function describeReading(reading: WordIntegerReading, attribute: string): string {
  if (reading.value === null) return reading.storage.includes('computed') ? 'a width Word computes' : 'not known';
  const where = reading.attribute !== undefined && reading.attribute !== attribute ? ` in w:${reading.attribute}` : '';
  const rule = reading.rule !== undefined && dependsOnRule(reading.storage) ? `${reading.rule} ` : '';
  return `${rule}${reading.value}${where}`;
}
