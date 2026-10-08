// CR-002 phase E: custom XML parts, XML mapping, the XPath engine and the typed content controls
// (CR-002 sections 3.5 and 3.6). Exported from '.' and './model', as the content API is.
export { type XPathEngine, type XPathValue, type BooleanConversionMode, booleanValue, DefaultXPathEngine, parsePrefixMappings, formatPrefixMappings, canonicalXPathOf, xmlOf, type CanonicalXPath } from './xpath.mjs';
export { CustomXmlPart, CustomXmlNode, CustomXmlPrefixMappingCollection, type CustomXmlNodeType, type CustomXmlPrefixMapping, type CustomXmlPartOwner } from './CustomXmlPart.mjs';
export { CustomXmlPartCollection, type CustomXmlHost } from './CustomXmlPartCollection.mjs';
export { XmlMapping, type CustomXmlPartLookup } from './XmlMapping.mjs';
export {
  CheckboxContentControl, DatePickerContentControl, ListContentControl, ContentControlListItem,
  PictureContentControl, RepeatingSectionContentControl, GroupContentControl, checkboxRun,
  CHECKED_SYMBOL, UNCHECKED_SYMBOL, CHECKBOX_FONT,
} from './kinds.mjs';
export { applyBindingsTo, applyBindingTo, updateFromControls, updateFromControl, runsForValue, formatDate, holdsFlatOpc, PLACEHOLDER_STYLE, PLACEHOLDER_TEXT, type BindingResult } from './bindings.mjs';
export { controlIdScope, sdtKindFor, isRunLevel } from './insert.mjs';
export { tagParamsOf, xpathsEntriesOf, selectEntryNode, OPENDOPE_XPATHS_NS, XHTML_CONTENT_TYPE, PICTURE_HANDLER, type XPathsEntry } from './opendope.mjs';
export { applyXhtmlBindingTo, applyXhtmlBindingsTo, isXhtmlBound, type XhtmlHost, type XhtmlBindingOptions } from './xhtml.mjs';
export { applyPictureHandlerTo, applyPictureHandlersTo, isPictureHandlerBound, bindPicture, pictureWidthOf, type PictureWidth, type PictureOutcome } from './pictures.mjs';
