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
