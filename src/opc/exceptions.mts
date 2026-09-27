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
