// The little of the Office JS API this taskpane uses, declared here rather than by depending on
// `@types/office-js`: the point of the example is that the edit is written against this package's
// types, and the Office globals are only the seam at the edges.
//
// In your own add-in, install `@types/office-js` and delete this file.

declare namespace Office {
  function onReady(callback?: (info: { host: unknown; platform: unknown }) => void): Promise<{ host: unknown; platform: unknown }>;
}

declare namespace Word {
  /** A loaded property (`ClientResult<T>`): `value` is there after the next `context.sync()`. */
  interface ClientResult<T> {
    value: T;
  }

  interface Range {
    getOoxml(): ClientResult<string>;
    insertOoxml(ooxml: string, insertLocation: 'Replace' | 'Before' | 'After'): Range;
    insertText(text: string, insertLocation: 'Replace' | 'Before' | 'After' | 'Start' | 'End'): Range;
  }

  interface Document {
    getSelection(): Range;
  }

  interface RequestContext {
    document: Document;
    sync(): Promise<void>;
  }

  function run<T>(batch: (context: RequestContext) => Promise<T>): Promise<T>;
}
