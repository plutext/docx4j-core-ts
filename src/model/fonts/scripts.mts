// `Character.UnicodeScript.of(cp).name()` for JavaScript.
//
// docx4j's `RunFontSelector.spanScript` and `FontFallback.coverageGroupOf` key on the Java enum
// constant's name - "LATIN", "HAN", "GEORGIAN", "COMMON", "INHERITED", "UNKNOWN" - which is the
// Unicode script's long alias upper-cased with its hyphens turned into underscores.  There is
// no such API in JavaScript, but `\p{Script=...}` in a Unicode-mode regular expression is the
// same property, so the answer is found by testing the scripts the engine knows.
//
// Scripts are disjoint, so the order of the test does not matter; a code point no script claims
// is UNKNOWN, as `Character.UnicodeScript.of` answers for an unassigned one.  Answers are
// memoised per code point, since a document asks about the same few hundred over and over.

/** The Unicode script long aliases, as `\p{Script=...}` spells them.  A value this engine's
 *  Unicode version does not know is dropped when the pattern fails to compile, so a newer
 *  script simply answers UNKNOWN on an older Node, as it would on an older JDK. */
const SCRIPT_NAMES = [
  'Common', 'Latin', 'Greek', 'Cyrillic', 'Armenian', 'Hebrew', 'Arabic', 'Syriac', 'Thaana',
  'Devanagari', 'Bengali', 'Gurmukhi', 'Gujarati', 'Oriya', 'Tamil', 'Telugu', 'Kannada',
  'Malayalam', 'Sinhala', 'Thai', 'Lao', 'Tibetan', 'Myanmar', 'Georgian', 'Hangul', 'Ethiopic',
  'Cherokee', 'Canadian_Aboriginal', 'Ogham', 'Runic', 'Khmer', 'Mongolian', 'Hiragana',
  'Katakana', 'Bopomofo', 'Han', 'Yi', 'Old_Italic', 'Gothic', 'Deseret', 'Inherited', 'Tagalog',
  'Hanunoo', 'Buhid', 'Tagbanwa', 'Limbu', 'Tai_Le', 'Linear_B', 'Ugaritic', 'Shavian',
  'Osmanya', 'Cypriot', 'Braille', 'Buginese', 'Coptic', 'New_Tai_Lue', 'Glagolitic', 'Tifinagh',
  'Syloti_Nagri', 'Old_Persian', 'Kharoshthi', 'Balinese', 'Cuneiform', 'Phoenician',
  'Phags_Pa', 'Nko', 'Sundanese', 'Lepcha', 'Ol_Chiki', 'Vai', 'Saurashtra', 'Kayah_Li',
  'Rejang', 'Lycian', 'Carian', 'Lydian', 'Cham', 'Tai_Tham', 'Tai_Viet', 'Avestan',
  'Egyptian_Hieroglyphs', 'Samaritan', 'Lisu', 'Bamum', 'Javanese', 'Meetei_Mayek',
  'Imperial_Aramaic', 'Old_South_Arabian', 'Inscriptional_Parthian', 'Inscriptional_Pahlavi',
  'Old_Turkic', 'Brahmi', 'Kaithi', 'Meroitic_Hieroglyphs', 'Meroitic_Cursive', 'Sora_Sompeng',
  'Chakma', 'Sharada', 'Takri', 'Miao', 'Caucasian_Albanian', 'Bassa_Vah', 'Duployan',
  'Elbasan', 'Grantha', 'Pahawh_Hmong', 'Khojki', 'Linear_A', 'Mahajani', 'Manichaean',
  'Mende_Kikakui', 'Modi', 'Mro', 'Old_North_Arabian', 'Nabataean', 'Palmyrene', 'Pau_Cin_Hau',
  'Old_Permic', 'Psalter_Pahlavi', 'Siddham', 'Khudawadi', 'Tirhuta', 'Warang_Citi', 'Ahom',
  'Anatolian_Hieroglyphs', 'Hatran', 'Multani', 'Old_Hungarian', 'SignWriting', 'Adlam',
  'Bhaiksuki', 'Marchen', 'Newa', 'Osage', 'Tangut', 'Masaram_Gondi', 'Nushu', 'Soyombo',
  'Zanabazar_Square', 'Dogra', 'Gunjala_Gondi', 'Makasar', 'Medefaidrin', 'Hanifi_Rohingya',
  'Sogdian', 'Old_Sogdian', 'Elymaic', 'Nandinagari', 'Nyiakeng_Puachue_Hmong', 'Wancho',
  'Chorasmian', 'Dives_Akuru', 'Khitan_Small_Script', 'Yezidi', 'Cypro_Minoan', 'Old_Uyghur',
  'Tangsa', 'Toto', 'Vithkuqi', 'Kawi', 'Nag_Mundari',
] as const;

let TESTS: { name: string; re: RegExp }[] | undefined;

function tests(): { name: string; re: RegExp }[] {
  if (TESTS === undefined) {
    TESTS = [];
    for (const name of SCRIPT_NAMES) {
      try {
        TESTS.push({ name: name.toUpperCase(), re: new RegExp(`\\p{Script=${name}}`, 'u') });
      } catch {
        // a script this engine's Unicode version does not know
      }
    }
  }
  return TESTS;
}

const MEMO = new Map<number, string>();

/** The Unicode script of a code point, as `Character.UnicodeScript.of(cp).name()` spells it. */
export function scriptOf(cp: number): string {
  const memo = MEMO.get(cp);
  if (memo !== undefined) return memo;
  let answer = 'UNKNOWN';
  let c: string;
  try {
    c = String.fromCodePoint(cp);
  } catch {
    MEMO.set(cp, answer);
    return answer;
  }
  for (const { name, re } of tests()) {
    if (re.test(c)) { answer = name; break; }
  }
  MEMO.set(cp, answer);
  return answer;
}
