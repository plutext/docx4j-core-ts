// docx4j `org.docx4j.fonts.CJKToEnglish`.
//
// Certain fonts have CJK names in a docx; on Windows (English locale at least) the font names
// are in English.  Font discovery collects the English name where there is one, so that the
// mapper is populated with the name the machine would have.

const TO_ENGLISH = new Map<string, string>([
  ['ＭＳ ゴシック', 'MS Gothic'],   // <a:font script="Jpan" typeface="ＭＳ ゴシック"/>
  ['ＭＳ 明朝', 'MS Mincho'],
  ['맑은 고딕', 'Malgun Gothic'],   // <a:font script="Hang" typeface="맑은 고딕"/>
  ['宋体', 'SimSun'],               // <a:font script="Hans" typeface="宋体"/>
  ['新細明體', 'PMingLiU'],
]);

/** The English name of a CJK font name, or undefined where this is not one of the names we know. */
export function toEnglish(fontName: string): string | undefined {
  return TO_ENGLISH.get(fontName);
}
