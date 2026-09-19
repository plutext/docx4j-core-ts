// The template markup `PresentationMLPackage.createPackage()` starts a new presentation from, as
// docx4j builds it (docx4j VERSION_17_1_1 a58cf10b8):
//
//   docx4j-core/src/main/java/org/docx4j/openpackaging/parts/PresentationML/JaxbPmlPart.java
//     COMMON_SLIDE_DATA, COLOR_MAPPING
//   .../PresentationML/SlideMasterPart.java      createSldMaster()
//   .../PresentationML/SlideLayoutPart.java      createSldLayout()   (cSld name "Title Slide")
//   .../PresentationML/SlidePart.java            createSld()
//
// docx4j unmarshals those fragments into a JAXB tree and marshals it back out; here they are the
// whole part, so a created slide, layout or master that nobody touches is written byte for byte
// as docx4j writes it.  Generated content, not hand markup: do not edit without the Java.
//
// The theme part is not here: `createPackage()` uses the same embedded Office themes as
// `WordprocessingMLPackage.createPackage()` (`src/model/fonts/themes.generated.mts`), so which
// theme a new presentation carries follows `pkg.fonts.defaultTheme`, as it does for a new
// document.  docx4j instead ships one fixed
// `org/docx4j/openpackaging/parts/PresentationML/theme.xml` (CR-001 section 16).

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const NS = ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
  + ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
  + ' xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';

/** docx4j `JaxbPmlPart.COMMON_SLIDE_DATA`: an empty shape tree. `name` is the `p:cSld/@name`. */
function commonSlideData(name?: string): string {
  return `<p:cSld${name === undefined ? '' : ` name="${name}"`}>`
    + '<p:spTree>'
    + '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
    + '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
    + '</p:spTree>'
    + '</p:cSld>';
}

/** docx4j `JaxbPmlPart.COLOR_MAPPING`: the identity colour map a master needs. */
const COLOR_MAPPING = '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1"'
  + ' accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6"'
  + ' hlink="hlink" folHlink="folHlink"/>';

/** `/ppt/slideMasters/slideMaster1.xml` as docx4j's `SlideMasterPart.createSldMaster()` writes it. */
export const DEFAULT_SLIDE_MASTER_XML = `${XML_DECL}<p:sldMaster${NS}>`
  + commonSlideData() + COLOR_MAPPING + '<p:sldLayoutIdLst/>'
  + '</p:sldMaster>';

/** `/ppt/slideLayouts/slideLayout1.xml` as docx4j's `SlideLayoutPart.createSldLayout()` writes it. */
export const DEFAULT_SLIDE_LAYOUT_XML = `${XML_DECL}<p:sldLayout${NS}>`
  + commonSlideData('Title Slide')
  + '</p:sldLayout>';

/** `/ppt/slides/slide1.xml` as docx4j's `SlidePart.createSld()` writes it. */
export const DEFAULT_SLIDE_XML = `${XML_DECL}<p:sld${NS}>`
  + commonSlideData()
  + '</p:sld>';
