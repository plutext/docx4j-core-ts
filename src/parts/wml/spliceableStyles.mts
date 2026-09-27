// The style definitions docx4j's own `styles.xml` carries but leaves **commented out**, so that
// `createPackage()` - which writes that file verbatim (`DEFAULT_STYLES_XML`, generated, faithful to
// docx4j) - produces a document without them. `pkg.styles.ensure(ids)` splices them in when a caller
// needs one: inserting a footnote into such a document otherwise shows the note in Normal
// (CR-002 section 22.2).
//
// Taken verbatim from the two comments in
// docx4j-core/src/main/resources/org/docx4j/openpackaging/parts/WordprocessingML/styles.xml
// (a UTF-16 file; docx4j VERSION_17_3_0), which is where Word's own definitions of these came from.
// 9 styles: BlockText, EndnoteText, EndnoteTextChar, Footer, FooterChar, FootnoteText, FootnoteTextChar, FootnoteReference, EndnoteReference.
export const SPLICEABLE_STYLES_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:style w:type="paragraph" w:styleId="BlockText">
    <w:name w:val="Block Text" />
    <w:basedOn w:val="Normal" />
    <w:uiPriority w:val="99" />
    <w:unhideWhenUsed />
    <w:rsid w:val="00841CD9" />
    <w:pPr>
      <w:pBdr>
        <w:top w:val="single" w:sz="2" w:space="10" w:color="4F81BD" w:themeColor="accent1" w:shadow="1" />
        <w:left w:val="single" w:sz="2" w:space="10" w:color="4F81BD" w:themeColor="accent1" w:shadow="1" />
        <w:bottom w:val="single" w:sz="2" w:space="10" w:color="4F81BD" w:themeColor="accent1" w:shadow="1" />
        <w:right w:val="single" w:sz="2" w:space="10" w:color="4F81BD" w:themeColor="accent1" w:shadow="1" />
      </w:pBdr>
      <w:ind w:left="1152" w:right="1152" />
    </w:pPr>
    <w:rPr>
      <w:rFonts w:eastAsiaTheme="minorEastAsia" />
      <w:i />
      <w:iCs />
      <w:color w:val="4F81BD" w:themeColor="accent1" />
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="EndnoteText">
    <w:name w:val="endnote text" />
    <w:basedOn w:val="Normal" />
    <w:link w:val="EndnoteTextChar" />
    <w:uiPriority w:val="99" />
    <w:unhideWhenUsed />
    <w:rsid w:val="00841CD9" />
    <w:rPr>
      <w:sz w:val="20" />
      <w:szCs w:val="20" />
    </w:rPr>
  </w:style>
  <w:style w:type="character" w:customStyle="1" w:styleId="EndnoteTextChar">
    <w:name w:val="Endnote Text Char" />
    <w:basedOn w:val="DefaultParagraphFont" />
    <w:link w:val="EndnoteText" />
    <w:uiPriority w:val="99" />
    <w:rsid w:val="00841CD9" />
    <w:rPr>
      <w:sz w:val="20" />
      <w:szCs w:val="20" />
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Footer">
    <w:name w:val="footer" />
    <w:basedOn w:val="Normal" />
    <w:link w:val="FooterChar" />
    <w:uiPriority w:val="99" />
    <w:unhideWhenUsed />
    <w:rsid w:val="00841CD9" />
    <w:pPr>
      <w:tabs>
        <w:tab w:val="center" w:pos="4680" />
        <w:tab w:val="right" w:pos="9360" />
      </w:tabs>
    </w:pPr>
  </w:style>
  <w:style w:type="character" w:customStyle="1" w:styleId="FooterChar">
    <w:name w:val="Footer Char" />
    <w:basedOn w:val="DefaultParagraphFont" />
    <w:link w:val="Footer" />
    <w:uiPriority w:val="99" />
    <w:rsid w:val="00841CD9" />
  </w:style>
  <w:style w:type="paragraph" w:styleId="FootnoteText">
    <w:name w:val="footnote text" />
    <w:basedOn w:val="Normal" />
    <w:link w:val="FootnoteTextChar" />
    <w:uiPriority w:val="99" />
    <w:unhideWhenUsed />
    <w:rsid w:val="00841CD9" />
    <w:rPr>
      <w:sz w:val="20" />
      <w:szCs w:val="20" />
    </w:rPr>
  </w:style>
  <w:style w:type="character" w:customStyle="1" w:styleId="FootnoteTextChar">
    <w:name w:val="Footnote Text Char" />
    <w:basedOn w:val="DefaultParagraphFont" />
    <w:link w:val="FootnoteText" />
    <w:uiPriority w:val="99" />
    <w:rsid w:val="00841CD9" />
    <w:rPr>
      <w:sz w:val="20" />
      <w:szCs w:val="20" />
    </w:rPr>
  </w:style>
<w:style w:type="character" w:styleId="FootnoteReference">
    <w:name w:val="footnote reference" />
    <w:basedOn w:val="DefaultParagraphFont" />
    <w:uiPriority w:val="99" />
    <w:unhideWhenUsed />
    <w:rsid w:val="00D1197D" />
    <w:rPr>
      <w:vertAlign w:val="superscript" />
    </w:rPr>
  </w:style>
  <w:style w:type="character" w:styleId="EndnoteReference">
    <w:name w:val="endnote reference" />
    <w:basedOn w:val="DefaultParagraphFont" />
    <w:uiPriority w:val="99" />
    <w:unhideWhenUsed />
    <w:rsid w:val="00D1197D" />
    <w:rPr>
      <w:vertAlign w:val="superscript" />
    </w:rPr>
  </w:style>
</w:styles>`;
