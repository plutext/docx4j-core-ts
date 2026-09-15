// The errors the Word shim throws (CR-002 section 3.10). Office JS raises OfficeExtension.Error
// with a `code` from Word.ErrorCodes; these carry the same codes so that add-in code which
// branches on `error.code` behaves the same here.

/**
 * A member of the Word API that this package does not implement. The message names the class and
 * the member, so an add-in's unsupported call fails at the call with something actionable:
 * `Word.Paragraph.getNextOrNullObject is not supported by @docx4j/core-ts`.
 */
export class NotSupportedError extends Error {
  /** Word.ErrorCodes.notImplemented. */
  readonly code = 'NotImplemented';

  constructor(
    /** The Word class, as Office JS names it ('Paragraph', 'Body', 'RequestContext'). */
    readonly className: string,
    /** The member that was read, written or called. */
    readonly member: string,
    detail?: string,
  ) {
    super(`Word.${className}.${member} is not supported by @docx4j/core-ts${detail ? `: ${detail}` : ''}`);
    this.name = 'NotSupportedError';
  }
}

/** Nothing matched: `getFirst()` on an empty collection, or a member of a null object. */
export class ItemNotFoundError extends Error {
  /** Word.ErrorCodes.itemNotFound. */
  readonly code = 'ItemNotFound';

  constructor(message: string) {
    super(message);
    this.name = 'ItemNotFoundError';
  }
}

/** A result read before `context.sync()` resolved it (Office JS: ValueNotLoaded). */
export class ValueNotLoadedError extends Error {
  readonly code = 'ValueNotLoaded';

  constructor(message = 'The value is not loaded yet; await context.sync() first') {
    super(message);
    this.name = 'ValueNotLoadedError';
  }
}
