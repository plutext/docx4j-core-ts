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

/** docx4j PartUnrecognisedException: no part class for a content type (the registry falls back before this is thrown). */
export class PartUnrecognisedException extends Docx4JException {
  constructor(message: string) {
    super(message);
    this.name = 'PartUnrecognisedException';
  }
}
