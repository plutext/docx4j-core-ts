// The number formats a list label is written in (docx4j `NumberFormatter` and its
// `LabelFormatter` family, CR-014 phase 1).
//
// A registry of stateless formatters keyed by `w:numFmt`, and **fail-soft**: a format with no
// formatter, or a value a formatter cannot express (Roman above 3999, a circled digit above 20,
// a letter below 1), gives the **decimal** label and one warning per format, which is what Word
// does.  The formatters themselves are not fail-soft - they throw `RangeError`, so a caller that
// wants to know a value is out of range still can - exactly as docx4j's throw
// `NumberFormatException` and `NumberFormatter` catches it.
//
// `register(numFmt, formatter)` adds or replaces one, for the counting styles docx4j does not
// ship (Japanese, Korean, Vietnamese and most of the ideographic ones).
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import { log } from '../properties/log.mjs';

/**
 * Turns a list counter into the label text of one `w:numFmt`: 3 into "iii", "C", "3rd",
 * "Three", "Third", "03", "③".  docx4j `LabelFormatter.format(int)`.
 *
 * Throws `RangeError` where the format has no label for the value; {@link formatValue}
 * substitutes the decimal label.
 */
export type LabelFormatter = (value: number) => string;

// ---------------------------------------------------------------- the formatters

/** docx4j `NumberFormatRomanAbstract`: 1 to 3999, nothing outside. */
const ROMAN: ReadonlyArray<readonly [number, string]> = [
  [1000, 'm'], [900, 'cm'], [500, 'd'], [400, 'cd'], [100, 'c'], [90, 'xc'],
  [50, 'l'], [40, 'xl'], [10, 'x'], [9, 'ix'], [5, 'v'], [4, 'iv'], [1, 'i'],
];

function romanLower(value: number): string {
  if (value >= 4000 || value < 1) throw new RangeError('Numbers must be in range 1-3999');
  let n = value;
  let out = '';
  for (const [number, letters] of ROMAN) {
    while (n >= number) { n -= number; out += letters; }
  }
  return out;
}

/**
 * docx4j `NumberFormatAlphabet`: the n-th letter up to the alphabet's length, then that letter
 * repeated - after z comes aa, bb, cc, not ab (ECMA-376 17.18.59, and Word).  The Cyrillic,
 * Arabic and Thai sets are the ones Word's labels use, which leave out some letters of each
 * script (Russian has no ё, й, ъ, ы, ь; Thai omits ฃ and ฅ).
 */
function alphabet(letters: string): LabelFormatter {
  const glyphs = Array.from(letters);
  return (value: number): string => {
    if (value < 1) throw new RangeError(`No letter for ${value}`);
    const size = glyphs.length;
    return glyphs[(value - 1) % size]!.repeat(Math.floor((value - 1) / size) + 1);
  };
}

/** docx4j `NumberFormatDigits`: a decimal label in another script's digits. */
function digits(ten: string): LabelFormatter {
  const glyphs = Array.from(ten);
  if (glyphs.length !== 10) throw new Error('ten digits expected');
  return (value: number): string => String(value).replace(/[0-9]/g, (c) => glyphs[c.charCodeAt(0) - 48]!);
}

/** docx4j `NumberFormatDecimalZero`: 01, 02 ... 09, 10. */
function decimalZero(value: number): string {
  return value < 10 ? `0${value}` : String(value);
}

/** docx4j `NumberFormatOrdinal`: 1st, 2nd, 3rd, 4th, 11th, 21st, 101st, 111th. */
function ordinal(value: number): string {
  return `${value}${ordinalSuffix(value)}`;
}

function ordinalSuffix(n: number): string {
  const abs = Math.abs(n);
  const lastTwo = abs % 100;
  if (lastTwo >= 11 && lastTwo <= 13) return 'th';
  switch (abs % 10) {
    case 1: return 'st';
    case 2: return 'nd';
    case 3: return 'rd';
    default: return 'th';
  }
}

const ONES = ['Zero', 'One', 'Two', 'Three', 'Four', 'Five', 'Six', 'Seven', 'Eight', 'Nine',
  'Ten', 'Eleven', 'Twelve', 'Thirteen', 'Fourteen', 'Fifteen', 'Sixteen', 'Seventeen',
  'Eighteen', 'Nineteen'];
const TENS = ['', '', 'Twenty', 'Thirty', 'Forty', 'Fifty', 'Sixty', 'Seventy', 'Eighty', 'Ninety'];
const SCALES = ['', 'Thousand', 'Million', 'Billion'];
const ORDINAL_ONES = ['Zeroth', 'First', 'Second', 'Third', 'Fourth', 'Fifth', 'Sixth', 'Seventh',
  'Eighth', 'Ninth', 'Tenth', 'Eleventh', 'Twelfth', 'Thirteenth', 'Fourteenth', 'Fifteenth',
  'Sixteenth', 'Seventeenth', 'Eighteenth', 'Nineteenth'];

/** docx4j `NumberFormatCardinalText.words`: the cardinal words, every word capitalised. */
function cardinalWords(value: number): string {
  if (value < 20) return ONES[value]!;
  let n = value;
  const groups: number[] = [];
  while (n > 0) { groups.push(n % 1000); n = Math.floor(n / 1000); }
  let out = '';
  for (let i = groups.length - 1; i >= 0; i--) {
    if (groups[i] === 0) continue;
    if (out.length > 0) out += ' ';
    out += belowThousand(groups[i]!);
    if (i > 0) out += ` ${SCALES[i]}`;
  }
  return out;
}

function belowThousand(n: number): string {
  let out = '';
  let rest = n;
  if (rest >= 100) {
    out += `${ONES[Math.floor(rest / 100)]} Hundred`;
    rest %= 100;
    if (rest > 0) out += ' ';
  }
  if (rest >= 20) {
    out += TENS[Math.floor(rest / 10)];
    if (rest % 10 > 0) out += `-${ONES[rest % 10]}`;
  } else if (rest > 0) {
    out += ONES[rest];
  }
  return out;
}

/** docx4j `NumberFormatCardinalText`: One, Twenty-One, One Hundred One, One Thousand. */
function cardinalText(value: number): string {
  if (value < 0) throw new RangeError(`No words for ${value}`);
  return cardinalWords(value);
}

/** docx4j `NumberFormatOrdinalText`: the cardinal words with the last word made ordinal. */
function ordinalText(value: number): string {
  if (value < 0) throw new RangeError(`No words for ${value}`);
  const cardinal = cardinalWords(value);
  const cut = Math.max(cardinal.lastIndexOf(' '), cardinal.lastIndexOf('-'));
  const head = cut < 0 ? '' : cardinal.slice(0, cut + 1);
  return head + ordinalWord(cardinal.slice(cut + 1));
}

function ordinalWord(word: string): string {
  const index = ONES.indexOf(word);
  if (index >= 0) return ORDINAL_ONES[index]!;
  if (word.endsWith('y')) return `${word.slice(0, -1)}ieth`;   // Twenty -> Twentieth
  return `${word}th`;                                          // Hundred, Thousand, Million
}

/** docx4j `NumberFormatChicago`: * † ‡ §, then doubled, then tripled. */
const CHICAGO_SYMBOLS = ['*', '†', '‡', '§'];

function chicago(value: number): string {
  if (value < 1) throw new RangeError(`No symbol for ${value}`);
  const symbol = CHICAGO_SYMBOLS[(value - 1) % CHICAGO_SYMBOLS.length]!;
  return symbol.repeat(Math.floor((value - 1) / CHICAGO_SYMBOLS.length) + 1);
}

const CIRCLED = ['①', '②', '③', '④', '⑤', '⑥', '⑦', '⑧',
  '⑨', '⑩', '⑪', '⑫', '⑬', '⑭', '⑮', '⑯', '⑰',
  '⑱', '⑲', '⑳'];

/** docx4j `NumberFormatDecimalEnclosedCircle`: ① to ⑳, and nothing above. */
function decimalEnclosedCircle(value: number): string {
  if (value <= 0 || value > 20) throw new RangeError('Numbers must be in range 1-20');
  return CIRCLED[value - 1]!;
}

const HEBREW_UNITS = ['א', 'ב', 'ג', 'ד', 'ה', 'ו', 'ז', 'ח', 'ט'];
const HEBREW_TENS = ['י', 'כ', 'ל', 'מ', 'נ', 'ס', 'ע', 'פ', 'צ'];
const HEBREW_HUNDREDS = ['ק', 'ר', 'ש', 'ת'];

/** docx4j `NumberFormatHebrew1`: טו and טז for 15 and 16, ת repeated above 400, no gershayim. */
function hebrew1(value: number): string {
  if (value < 1) throw new RangeError(`No Hebrew numeral for ${value}`);
  let out = '';
  let n = value;
  while (n >= 400) { out += 'ת'; n -= 400; }
  if (n >= 100) { out += HEBREW_HUNDREDS[Math.floor(n / 100) - 1]; n %= 100; }
  if (n === 15) return `${out}טו`;
  if (n === 16) return `${out}טז`;
  if (n >= 10) { out += HEBREW_TENS[Math.floor(n / 10) - 1]; n %= 10; }
  if (n > 0) out += HEBREW_UNITS[n - 1];
  return out;
}

/** docx4j `NumberFormatChineseAbstract`, with the lower or the legal (upper) characters. */
function chinese(characters: readonly string[], units: readonly string[], bigUnits: readonly string[]): LabelFormatter {
  const fourDigit = (num: number): string => {
    let out = '';
    let n = num;
    let unitPosition = 0;
    let lastIsZero = false;
    while (n > 0) {
      const digit = n % 10;
      if (digit !== 0) {
        out = `${characters[digit]}${units[unitPosition]}${out}`;
        lastIsZero = false;
      } else if (!lastIsZero) {
        out = `${characters[0]}${out}`;
        lastIsZero = true;
      }
      n = Math.floor(n / 10);
      unitPosition++;
    }
    // remove the trailing 零
    if (out.endsWith(characters[0]!)) out = out.slice(0, -characters[0]!.length);
    return out;
  };
  return (value: number): string => {
    if (value === 0) return characters[0]!;
    let out = '';
    let n = value;
    let bigUnitPosition = 0;
    let lastIsZero = false;
    while (n > 0) {
      const part = n % 10000;
      if (part !== 0) {
        out = `${fourDigit(part)}${bigUnitPosition > 0 ? bigUnits[bigUnitPosition] : ''}${out}`;
        lastIsZero = false;
      } else if (!lastIsZero && out.length > 0) {
        out = `${characters[0]}${out}`;
        lastIsZero = true;
      }
      n = Math.floor(n / 10000);
      bigUnitPosition++;
    }
    return out;
  };
}

const chineseLower = chinese(
  ['零', '一', '二', '三', '四', '五', '六', '七', '八', '九'],
  ['', '十', '百', '千'], ['', '万', '亿']);
const chineseUpper = chinese(
  ['零', '壹', '贰', '叁', '肆', '伍', '陆', '柒', '捌', '玖'],
  ['', '拾', '佰', '仟'], ['', '万', '亿']);

const decimal: LabelFormatter = (value) => String(value);

// ---------------------------------------------------------------- the registry

/** `w:numFmt` -> its formatter. docx4j `NumberFormatter.REGISTRY`. */
const REGISTRY = new Map<wml.NumberFormat, LabelFormatter>([
  ['decimal', decimal],
  ['decimalHalfWidth', decimal],
  ['none', () => ''],
  // TODO (docx4j's own): revisit; the code elsewhere for bullets overlaps with this numFmt stuff
  ['bullet', () => '*'],
  ['upperRoman', (v) => romanLower(v).toUpperCase()],
  ['lowerRoman', romanLower],
  ['lowerLetter', alphabet('abcdefghijklmnopqrstuvwxyz')],
  ['upperLetter', alphabet('ABCDEFGHIJKLMNOPQRSTUVWXYZ')],
  ['decimalZero', decimalZero],
  ['ordinal', ordinal],
  ['cardinalText', cardinalText],
  ['ordinalText', ordinalText],
  ['hex', (v) => (v >>> 0).toString(16).toUpperCase()],
  ['chicago', chicago],
  ['numberInDash', (v) => `- ${v} -`],
  ['decimalFullWidth', digits('０１２３４５６７８９')],
  ['decimalFullWidth2', digits('０１２３４５６７８９')],
  ['thaiNumbers', digits('๐๑๒๓๔๕๖๗๘๙')],
  ['hindiNumbers', digits('०१२३४५६७८९')],
  ['russianLower', alphabet('абвгдежзиклмнопрстуфхцчшщэюя')],
  ['russianUpper', alphabet('АБВГДЕЖЗИКЛМНОПРСТУФХЦЧШЩЭЮЯ')],
  ['arabicAlpha', alphabet('أبتثجحخدذرزسشصضطظعغفقكلمنهوي')],
  ['thaiLetters', alphabet('กขคฆงจฉชซฌญฎฏฐฑฒณดตถทธนบปผฝพฟภมยรลวศษสหฬอฮ')],
  ['hebrew1', hebrew1],
  // these two are the same in Chinese, no need to be processed separately
  ['chineseCounting', chineseLower],
  ['chineseCountingThousand', chineseLower],
  // this one means use upper Chinese number characters
  ['chineseLegalSimplified', chineseUpper],
  // these two are the same, just to adapt to documents in Chinese
  ['decimalEnclosedCircle', decimalEnclosedCircle],
  ['decimalEnclosedCircleChinese', decimalEnclosedCircle],
]);

/** The formats warned about already: one warning per format, whatever the value or document. */
const WARNED = new Set<string>();

/** Adds or replaces the formatter for a `w:numFmt`; `undefined` removes it. docx4j `register`. */
export function register(numFmt: wml.NumberFormat, formatter: LabelFormatter | undefined): void {
  if (formatter === undefined) REGISTRY.delete(numFmt);
  else REGISTRY.set(numFmt, formatter);
}

/** The formatter registered for a `w:numFmt`, or undefined. docx4j `NumberFormatter.get`. */
export function formatterFor(numFmt: wml.NumberFormat | undefined): LabelFormatter | undefined {
  return numFmt === undefined ? undefined : REGISTRY.get(numFmt);
}

/** Every `w:numFmt` this package can format (a test's, and a caller's, inventory). */
export function registeredFormats(): wml.NumberFormat[] {
  return [...REGISTRY.keys()];
}

/** Forgets which formats have been warned about; a test that asserts on the warning needs it. */
export function resetWarnings(): void {
  WARNED.clear();
}

/**
 * A counter as its label. docx4j `NumberFormatter.getCurrentValueFormatted(numFmt, current, where)`.
 *
 * A `w:numFmt` of undefined is decimal; so is a format with no formatter, and a value the
 * formatter cannot express - with one warning per format naming `where` ("numId 3 ilvl 0") it
 * first happened.
 */
export function formatValue(numFmt: wml.NumberFormat | undefined, value: number, where?: string): string {
  if (numFmt === undefined) return String(value);
  const formatter = REGISTRY.get(numFmt);
  if (formatter === undefined) {
    warnOnce(numFmt, `no formatter for numFmt ${numFmt}; decimal labels used`, where);
    return String(value);
  }
  try {
    return formatter(value);
  } catch (e) {
    if (!(e instanceof RangeError)) throw e;
    warnOnce(numFmt, `numFmt ${numFmt} cannot express ${value} (${e.message}); decimal label used`, where);
    return String(value);
  }
}

function warnOnce(numFmt: string, message: string, where: string | undefined): void {
  if (WARNED.has(numFmt)) return;
  WARNED.add(numFmt);
  log.warn(`${message}${where === undefined ? '' : ` (first at ${where})`}; not reported again for this format`);
}
