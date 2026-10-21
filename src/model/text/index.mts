// The text model (CR-012): docx4j's org.docx4j.model.text, held to the text extraction specification.
export { Texts } from './Texts.mjs';
export { TextContext, PackageTrees, type TextPackage } from './TextContext.mjs';
export {
  TextOptions, TextOptionError, TEXT_OPTION_NAMES,
  type TextOptionsInit, type TextModelView, type TextInclude, type TextFields, type TextBindings, type TextSymbols,
  type TextMath, type TextTables, type TextMarkupDetail, type TextLayout,
} from './TextOptions.mjs';
export {
  TextResult, TextBlock, TextRun, TextRow, TextCell, TextField, TextNumbering, TextRevision, TextWarning, TextComment,
  TextImage, ImageDescription, AltTextDescriber,
  type ImageDescriber, type TextBlockKind, type TextStory, type TextRunKind, type NoteType, type TextRunSource,
} from './TextResult.mjs';
export { PlainText, INDENT_SPACES, LISTED_FORMAT_NOTE } from './PlainText.mjs';
