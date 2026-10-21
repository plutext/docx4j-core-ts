// An equation's text as Office JS gives it (CR-012 phase 4; checks 48 case 17 and 51): the `m:t` text of its
// math runs in document order, each letter in the mathematical alphabet Word draws it in, a hyphen-minus as
// the minus sign U+2212, and each structure (a fraction, a script, a radical, an n-ary operator, a delimiter,
// a function...) written `,` then its arguments joined by `-` then `.`, as Word's text has it: a fraction a
// over b is `,𝑎-𝑏.`, sin x `,sin-𝑥.` (check 51 case 12). Word's text is the characters it draws and its own
// structure marks, not the ones the file stores. Measured: check 48's `x`, and check 51's alphabets, styles,
// scripts, Greek, digits and six structures; an argument a structure hides (a radical's degree under
// `m:degHide`) is left out, as check 51's radical had it.
import type { Element } from '@docx4j/generated-objects-ts/builders/wml';

/** The two views a paragraph's text is read in. */
type View = 'accepted' | 'original';

const MATH_RUN = 'org_docx4j_math.CTR';
const MATH_TEXT = 'org_docx4j_math.CTText';
const MATH_RUN_TRACK_CHANGE = 'org_docx4j_wml.CTMathRunTrackChange';
const CONTROL_PROPERTIES = 'org_docx4j_math.CTCtrlPr';

/** True for an equation at run-holder level: `m:oMath` or `m:oMathPara`. */
export function isMath(el: Element): boolean {
  const tn = (el.value as { TYPE_NAME?: string } | undefined)?.TYPE_NAME;
  return tn === 'org_docx4j_math.CTOMath' || tn === 'org_docx4j_math.CTOMathPara';
}

/**
 * The text of an equation (`m:oMath`, `m:oMathPara`) in a view: the `m:t` of every math run in document
 * order, a revision's runs as the view has them, each letter mapped as Word draws it (`mathLetter`).
 */
export function mathTextOf(math: unknown, view: View = 'accepted'): string {
  let out = '';
  const hides = (name: string | undefined): boolean =>
    view === 'accepted' ? name === 'del' || name === 'moveFrom' : name === 'ins' || name === 'moveTo';
  const runText = (run: { content?: unknown[] }): void => {
    const rPr = mathRunProperties(run);
    const add = (text: unknown): void => { out += styled(String(text ?? ''), rPr); };
    for (const o of run.content ?? []) {
      const el = o as { name?: { localPart?: string }; value?: { TYPE_NAME?: string; value?: unknown; rest?: unknown[] } };
      const v = el.value;
      if (v?.TYPE_NAME === MATH_RUN_TRACK_CHANGE) {
        if (hides(el.name?.localPart)) continue;
        for (const inner of v.rest ?? []) {
          const iv = (inner as { value?: { TYPE_NAME?: string; value?: unknown } }).value;
          if (iv?.TYPE_NAME === MATH_TEXT) add(iv.value);
        }
        continue;
      }
      if (v?.TYPE_NAME === MATH_TEXT) add(v.value);
      else if (v?.TYPE_NAME === 'org_docx4j_wml.Text' && el.name?.localPart === 't') add(v.value);
    }
  };
  const visit = (value: unknown, depth: number): void => {
    if (typeof value !== 'object' || value === null || depth > 64) return;
    if (Array.isArray(value)) { for (const v of value) visit(v, depth + 1); return; }
    const named = value as { name?: { localPart?: string }; value?: unknown };
    if (named.name !== undefined && 'value' in named) {
      if (hides(named.name.localPart) && isRevisionValue(named.value)) return;
      visit(named.value, depth + 1);
      return;
    }
    const tn = (value as { TYPE_NAME?: string }).TYPE_NAME;
    if (tn === MATH_RUN) { runText(value as { content?: unknown[] }); return; }
    if (isProperties(tn)) return;
    if (tn !== undefined && tn.startsWith('org_docx4j_math.') && !CONTAINERS.has(tn)) {
      // a structure: `,`, its arguments joined by `-`, `.` (check 51 case 12)
      out += ',';
      argumentsOf(value).forEach((argument, i) => {
        if (i > 0) out += '-';
        visit(argument, depth + 1);
      });
      out += '.';
      return;
    }
    for (const key of Object.keys(value)) {
      if (key === 'PARENT' || key === 'TYPE_NAME') continue;
      visit((value as Record<string, unknown>)[key], depth + 1);
    }
  };
  visit(math, 0);
  return out;
}

/** What holds math elements without being a structure: an equation, a display equation, an argument. */
const CONTAINERS = new Set(['org_docx4j_math.CTOMath', 'org_docx4j_math.CTOMathPara', 'org_docx4j_math.CTOMathArg']);

/** A math object's properties (`m:fPr`, `m:radPr`, `m:ctrlPr`, ...): no text. */
function isProperties(tn: string | undefined): boolean {
  return tn !== undefined && tn.startsWith('org_docx4j_math.') && (tn.endsWith('Pr') || tn === CONTROL_PROPERTIES);
}

/**
 * A structure's arguments (`m:e`, `m:num`, `m:den`, `m:sub`, `m:sup`, `m:deg`, `m:lim`, `m:fName`) in schema
 * order, a matrix's rows' arguments in turn, and a radical's degree left out where `m:degHide` hides it.
 */
function argumentsOf(structure: object): object[] {
  const out: object[] = [];
  const degHidden = on((structure as { radPr?: { degHide?: { val?: unknown } } }).radPr?.degHide);
  const take = (v: unknown): void => {
    if (typeof v !== 'object' || v === null) return;
    if (Array.isArray(v)) { for (const o of v) take(o); return; }
    const named = v as { name?: unknown; value?: unknown };
    const value = (named.name !== undefined && 'value' in named ? named.value : v) as { TYPE_NAME?: string };
    if (value?.TYPE_NAME === 'org_docx4j_math.CTOMathArg') { out.push(value); return; }
    if (isProperties(value?.TYPE_NAME)) return;
    if (value?.TYPE_NAME?.startsWith('org_docx4j_math.')) {
      for (const key of Object.keys(value)) if (key !== 'PARENT' && key !== 'TYPE_NAME') take((value as Record<string, unknown>)[key]);
    }
  };
  for (const key of Object.keys(structure)) {
    if (key === 'PARENT' || key === 'TYPE_NAME' || (key === 'deg' && degHidden)) continue;
    take((structure as Record<string, unknown>)[key]);
  }
  return out;
}

/** An `m:` on/off property: present and not `0`, `false` or `off`. */
function on(property: { val?: unknown } | undefined): boolean {
  return property !== undefined && property !== null && property.val !== false && property.val !== 'off' && property.val !== '0' && property.val !== 0;
}

/** A `w:ins`, `w:del` or `w:move*` around math runs. */
function isRevisionValue(value: unknown): boolean {
  const tn = (value as { TYPE_NAME?: string } | undefined)?.TYPE_NAME;
  return tn === 'org_docx4j_wml.RunIns' || tn === 'org_docx4j_wml.RunDel' || tn === 'org_docx4j_wml.RunTrackChange' || tn === MATH_RUN_TRACK_CHANGE;
}

interface MathRunProperties { sty?: string; scr?: string; nor: boolean }

function mathRunProperties(run: { content?: unknown[] }): MathRunProperties {
  for (const o of run.content ?? []) {
    const v = (o as { value?: { TYPE_NAME?: string; sty?: { val?: string }; scr?: { val?: string }; nor?: { val?: unknown } } }).value;
    if (v?.TYPE_NAME === 'org_docx4j_math.CTRPR') return { sty: v.sty?.val, scr: v.scr?.val, nor: on(v.nor) };
  }
  return { nor: false };
}

function styled(text: string, rPr: MathRunProperties): string {
  if (rPr.nor) return text;                                          // normal text: as the run's w:rPr draws it (check 51 case 11)
  let out = '';
  for (const c of text) out += c === '-' ? '\u2212' : mathLetter(c, rPr.sty, rPr.scr);   // a minus sign (check 51 case 6)
  return out;
}

// Unicode's Mathematical Alphanumeric Symbols: each alphabet's capital A and small a, its digit zero
// where it has digits, and its Greek block (capitals, nabla, smalls, then the variants).
const LATIN: Record<string, [number, number]> = {
  bold: [0x1D400, 0x1D41A], italic: [0x1D434, 0x1D44E], boldItalic: [0x1D468, 0x1D482],
  script: [0x1D49C, 0x1D4B6], boldScript: [0x1D4D0, 0x1D4EA], fraktur: [0x1D504, 0x1D51E],
  doubleStruck: [0x1D538, 0x1D552], boldFraktur: [0x1D56C, 0x1D586], sans: [0x1D5A0, 0x1D5BA],
  sansBold: [0x1D5D4, 0x1D5EE], sansItalic: [0x1D608, 0x1D622], sansBoldItalic: [0x1D63C, 0x1D656],
  monospace: [0x1D670, 0x1D68A],
};
const DIGITS: Record<string, number> = { bold: 0x1D7CE, doubleStruck: 0x1D7D8, sans: 0x1D7E2, sansBold: 0x1D7EC, monospace: 0x1D7F6 };
const GREEK: Record<string, number> = { bold: 0x1D6A8, italic: 0x1D6E2, boldItalic: 0x1D71C, sansBold: 0x1D756, sansBoldItalic: 0x1D790 };
/** The letters Unicode left where they were before the block: the BMP's letterlike symbols. */
const HOLES: Record<string, Record<string, number>> = {
  italic: { h: 0x210E },
  script: { B: 0x212C, E: 0x2130, F: 0x2131, H: 0x210B, I: 0x2110, L: 0x2112, M: 0x2133, R: 0x211B, e: 0x212F, g: 0x210A, o: 0x2134 },
  fraktur: { C: 0x212D, H: 0x210C, I: 0x2111, R: 0x211C, Z: 0x2128 },
  doubleStruck: { C: 0x2102, H: 0x210D, N: 0x2115, P: 0x2119, Q: 0x211A, R: 0x211D, Z: 0x2124 },
};
/** A Greek letter's place in an alphabet's Greek block of 58. */
function greekIndex(cp: number): number | undefined {
  if (cp >= 0x391 && cp <= 0x3A1) return cp - 0x391;               // Alpha to Rho
  if (cp === 0x3F4) return 17;                                        // capital theta symbol
  if (cp >= 0x3A3 && cp <= 0x3A9) return cp - 0x3A3 + 18;          // Sigma to Omega
  if (cp === 0x2207) return 25;                                       // nabla
  if (cp >= 0x3B1 && cp <= 0x3C9) return cp - 0x3B1 + 26;          // alpha to omega, final sigma among them
  const variants: Record<number, number> = { 0x2202: 51, 0x3F5: 52, 0x3D1: 53, 0x3F0: 54, 0x3D5: 55, 0x3F1: 56, 0x3D6: 57 };
  return variants[cp];
}

/**
 * A character as Word draws it in a math run with this `m:sty` (`p`, `b`, `i`, `bi`; none is Word's default,
 * italic) and `m:scr` (`roman`, the default, `script`, `fraktur`, `double-struck`, `sans-serif`,
 * `monospace`), as check 51 measured: letters, Greek capitals, small letters and nabla alike, in the
 * alphabet of the style and script; digits upright in the script's family (bold for bold or bold italic,
 * sans-serif under sans-serif, none of their own in script or fraktur); a character no alphabet has is
 * itself.
 */
export function mathLetter(c: string, sty: string | undefined, scr: string | undefined): string {
  const cp = c.codePointAt(0)!;
  const bold = sty === 'b' || sty === 'bi';
  const italic = sty === undefined || sty === 'i' || sty === 'bi';
  let alphabet: string | undefined;
  let digits: string | undefined;
  switch (scr ?? 'roman') {
    case 'roman': alphabet = bold ? (italic ? 'boldItalic' : 'bold') : italic ? 'italic' : undefined; digits = bold ? 'bold' : undefined; break;
    case 'script': alphabet = bold ? 'boldScript' : 'script'; break;
    case 'fraktur': alphabet = bold ? 'boldFraktur' : 'fraktur'; break;
    case 'double-struck': alphabet = 'doubleStruck'; digits = 'doubleStruck'; break;
    case 'sans-serif': alphabet = bold ? (italic ? 'sansBoldItalic' : 'sansBold') : italic ? 'sansItalic' : 'sans'; digits = bold ? 'sansBold' : 'sans'; break;
    case 'monospace': alphabet = 'monospace'; digits = 'monospace'; break;
    default: alphabet = undefined;
  }
  if (cp >= 0x30 && cp <= 0x39) {
    const zero = digits === undefined ? undefined : DIGITS[digits];
    return zero === undefined ? c : String.fromCodePoint(zero + cp - 0x30);
  }
  if (alphabet === undefined) return c;
  const hole = HOLES[alphabet]?.[c];
  if (hole !== undefined) return String.fromCodePoint(hole);
  const latin = LATIN[alphabet]!;
  if (cp >= 0x41 && cp <= 0x5A) return String.fromCodePoint(latin[0] + cp - 0x41);
  if (cp >= 0x61 && cp <= 0x7A) return String.fromCodePoint(latin[1] + cp - 0x61);
  const g = greekIndex(cp);
  if (g === undefined) return c;
  const base = GREEK[alphabet];
  return base === undefined ? c : String.fromCodePoint(base + g);
}
