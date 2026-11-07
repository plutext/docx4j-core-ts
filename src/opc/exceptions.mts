import type { MisplacedContentWarning } from './mce/preprocessor.mjs';

/** docx4j Docx4JException: anything the packaging layer cannot recover from. */
export class Docx4JException extends Error {
  readonly cause: unknown;
  constructor(message: string, options?: { cause?: unknown }) {
    super(message);
    this.name = 'Docx4JException';
    this.cause = options?.cause;
  }
}

/** docx4j InvalidFormatException: a package, part name or relationship violates Open Packaging. */
export class InvalidFormatException extends Docx4JException {
  constructor(message: string, options?: { cause?: unknown }) {
    super(message, options);
    this.name = 'InvalidFormatException';
  }
}

/**
 * A package whose container holds two entries of one name, OPC part names compared without case (WD REQ-206; CR-024),
 * refused when it is loaded, before anything is read. Word reads such a file only after a repair, and the readers that
 * don't refuse read different copies. `entries` gives each duplicated name, as its first entry spells it, with the
 * number of entries; `LoadOptions.duplicateEntries: 'first'` reads the first copy of each instead.
 */
export class DuplicateEntryException extends InvalidFormatException {
  /** The finding's code, shared with docx4j-python (its CR-032). */
  readonly code = 'package.duplicate_entry';

  constructor(readonly entries: Readonly<Record<string, number>>) {
    const list = Object.entries(entries);
    super(list.length === 1
      ? `Duplicate entry ${list[0]![0]} in this package: ${list[0]![1]} entries (REQ-206). Load with { duplicateEntries: 'first' } to read the first`
      : `Duplicate entries in this package: ${list.map(([name, n]) => `${name}, ${n} entries`).join('; ')} (REQ-206). `
        + "Load with { duplicateEntries: 'first' } to read the first of each");
    this.name = 'DuplicateEntryException';
  }
}

/**
 * Under `LoadOptions.duplicateEntries: 'first'`, a name the container holds more than one entry of (WD REQ-206; CR-024):
 * the first entry in the container's order is read, the rest ignored, and a save writes only the copy read, as Word's
 * repaired save does. On `pkg.warnings`, ahead of the parts' warnings. The code and the fields are docx4j-python's (its
 * CR-032): `entry`, the name as the first entry spells it (no leading '/'), and `count`, the entries of that name.
 */
export interface DuplicateEntryWarning {
  code: 'package.duplicate_entry';
  entry: string;
  count: number;
  message: string;
}

/**
 * A package whose package relationships hold more than one officeDocument relationship (WD REQ-206's statement, WD 0.100;
 * CR-024 section 9), refused when it is loaded, as a duplicated entry is and for the same reason: Word reads such a file
 * only through a repair, and which target it reads then alternates with each repaired open in a Word session. A
 * transitional and a strict relationship are two, and so are two naming one target. An External one names no part and
 * isn't counted. `targets` gives each as a part name (`/word/document.xml`), in the relationships' order.
 * `LoadOptions.duplicateEntries: 'first'` reads the last relationship's target instead, as a fresh Word does.
 */
export class MultipleMainPartsException extends InvalidFormatException {
  /** The finding's code, the WD's, shared with docx4j and docx4j-python. */
  readonly code = 'package.multiple_main_parts';

  constructor(readonly targets: readonly string[]) {
    super(`${targets.length} officeDocument relationships in this package: ${targets.join(', ')} (REQ-206). `
      + "Load with { duplicateEntries: 'first' } to read the last");
    this.name = 'MultipleMainPartsException';
  }
}

/**
 * Under `LoadOptions.duplicateEntries: 'first'`, a package with more than one officeDocument relationship (WD REQ-206's
 * statement; CR-024 section 9): the last relationship's target is read, as a fresh Word reads it, the other relationships
 * go at the load, and a save writes the one main part. On `pkg.warnings`, after any `DuplicateEntryWarning` and ahead of
 * the parts' warnings. The fields are docx4j-python's: `partName`, the part read (`/word/document.xml`), and `count`, the
 * relationships there were.
 */
export interface MultipleMainPartsWarning {
  code: 'package.multiple_main_parts';
  partName: string;
  count: number;
  message: string;
}

/**
 * docx4j CyclicStylesException: a style's `w:basedOn` chain comes back to a style it already
 * holds, or goes deeper than docx4j's limit of 32. Thrown only when
 * `throwOnCyclicStyles(true)` is set (docx4j's
 * `docx4j.openpackaging.exceptions.CyclicStylesException.throw` property, false by default);
 * otherwise the walk stops where the cycle was found and resolution uses as much of the
 * hierarchy as it has.
 */
export class CyclicStylesException extends Docx4JException {
  constructor(message: string) {
    super(message);
    this.name = 'CyclicStylesException';
  }
}

/**
 * Thrown by a synchronous read that needs the property resolver of a package which has not
 * built one yet. Building it unmarshals the styles part, which cannot be done synchronously.
 */
export class PropertyResolverNotCreatedException extends Docx4JException {
  constructor(what = 'Effective formatting') {
    super(`${what} needs the package's PropertyResolver; await getPropertyResolver() (or getBody()) first, or read direct formatting with { direct: true }`);
    this.name = 'PropertyResolverNotCreatedException';
  }
}

/** pptx4j Pptx4jException: a presentation's own structure (the slide list, a missing master). */
export class Pptx4jException extends Docx4JException {
  constructor(message: string, options?: { cause?: unknown }) {
    super(message, options);
    this.name = 'Pptx4jException';
  }
}

/** xlsx4j Xlsx4jException: a workbook's own structure (the sheet list). */
export class Xlsx4jException extends Docx4JException {
  constructor(message: string, options?: { cause?: unknown }) {
    super(message, options);
    this.name = 'Xlsx4jException';
  }
}

/** docx4j PartUnrecognisedException: no part class for a content type (the registry falls back before this is thrown). */
export class PartUnrecognisedException extends Docx4JException {
  constructor(message: string) {
    super(message);
    this.name = 'PartUnrecognisedException';
  }
}

/**
 * An XML part that could not be unmarshalled, naming the part (CR-002 section 27, item 9 of the
 * E4 release). The runtime's own message is specific about *what* is wrong and, since
 * `@docx4j/jsonix` 3.4.0, *where* - `w:sz/@w:val (line 12): Argument [NaN] must be an integer` -
 * but it cannot know which part it was reading, and a person editing a part's XML in an editor needs
 * that first. The location fields are lifted off the runtime's error so a caller does not have to
 * parse the message.
 */
export class PartUnmarshalException extends Docx4JException {
  constructor(
    message: string,
    /** The part being unmarshalled: `/word/styles.xml`. */
    readonly partName: string,
    options?: {
      cause?: unknown;
      /** The element or attribute the runtime was on, when it said. */
      location?: string | undefined;
      /** Its line in the source, when the parser supplied one (Node only: xmldom's locator). */
      line?: number | undefined;
      column?: number | undefined;
    },
  ) {
    super(message, { cause: options?.cause });
    this.name = 'PartUnmarshalException';
    this.location = options?.location;
    this.line = options?.line;
    this.column = options?.column;
  }

  readonly location: string | undefined;
  readonly line: number | undefined;
  readonly column: number | undefined;
}

/**
 * A save refused for a child its container never admits (CR-016 D2): the part being marshalled holds a run straight in
 * the body, a table in a paragraph, a string in a run, ..., which Word 365 refuses the file for or repairs by dropping
 * content (check 72). Nothing was written. `findings` are the part's, each with its path and docx4j's message naming
 * the fix; `pkg.checkContentOnSave = false` writes the file as held and reports them on `pkg.lastSave.misplaced` instead.
 */
export class MisplacedContentException extends Docx4JException {
  constructor(
    /** The part refused: `/word/document.xml`. */
    readonly partName: string,
    readonly findings: readonly MisplacedContentWarning[],
    options?: { cause?: unknown },
  ) {
    super(
      `Save refused: ${partName} holds ${findings.length === 1 ? 'a child its container never admits' : `${findings.length} children their containers never admit`}, `
      + 'which Word 365 refuses the file for or repairs by dropping content (set pkg.checkContentOnSave = false to write it as held):\n'
      + findings.map((f) => `  ${f.message}`).join('\n'),
      options,
    );
    this.name = 'MisplacedContentException';
  }
}

/** One part a save would lose content from (CR-025 section 2.3). */
export interface ContentLoss {
  partName: string;
  kind: 'skipped';
  /** The items its read dropped: `part.skipped`, one each. */
  count: number;
}

/**
 * A save refused because a part holding what its read dropped was changed, so a marshal would lose that (CR-025 section 2.3, the
 * repair policy's P7). Nothing is written. The code and the wording are docx4j-python's, with this package's way past it.
 */
export class SaveLossException extends Docx4JException {
  readonly code = 'save.content_lost';

  constructor(readonly losses: readonly ContentLoss[]) {
    super(`Saving would lose content: ${losses.map((l) => `${l.partName} holds ${l.count} item${l.count === 1 ? '' : 's'} the load skipped (part.skipped)`).join(', ')}. `
      + `Set pkg.checkLossOnSave = false to write ${losses.length === 1 ? 'the part without it' : 'the parts without them'}.`);
    this.name = 'SaveLossException';
  }
}
