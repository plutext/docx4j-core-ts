import { Docx4JException } from '../../opc/exceptions.mjs';

/**
 * What Office JS throws where the host refuses an operation it has no answer for: its `code`, so that
 * add-in code branching on `error.code` behaves the same here (check 19: a hyperlink set on an empty range;
 * check 52: a range starting or ending inside a table deleted or replaced).
 */
export class GeneralException extends Docx4JException {
  readonly code = 'GeneralException';
  constructor(message: string) {
    super(message);
    this.name = 'GeneralException';
  }
}

/**
 * What Office JS throws for an argument it won't take (`InvalidArgument`): its `code`, as `GeneralException` carries its
 * own (check 56: `insertTable` into a bound block content control).
 */
export class InvalidArgument extends Docx4JException {
  readonly code = 'InvalidArgument';
  constructor(message: string) {
    super(message);
    this.name = 'InvalidArgument';
  }
}
