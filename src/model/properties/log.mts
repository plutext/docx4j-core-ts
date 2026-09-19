// The one place property resolution writes a diagnostic. docx4j logs through SLF4J; this
// package has no logging framework, so a missing style or a style cycle goes to `console.warn`
// behind a settable sink (an add-in that must stay silent, or a test that wants to assert what
// was logged, replaces it).
//
// Resolution logs a missing style *once per id* (a `w:pStyle` naming a deleted style is common
// in real documents, and docx4j used to log one ERROR per paragraph); the resolver keeps that
// set, not this module.

export interface Logger {
  warn(message: string): void;
}

const consoleLogger: Logger = {
  warn(message: string): void {
    // eslint-disable-next-line no-console
    console.warn(message);
  },
};

let sink: Logger = consoleLogger;

/** The logger property resolution writes to. */
export const log: Logger = {
  warn(message: string): void {
    sink.warn(message);
  },
};

/** Replaces the sink; pass nothing to go back to `console.warn`. */
export function setLogger(logger?: Logger): void {
  sink = logger ?? consoleLogger;
}
