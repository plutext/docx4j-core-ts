// Hand-written declarations for edit.mjs, so the taskpane typechecks without `allowJs`.
import type { Body } from '@docx4j/core-ts/model';

/** One edit on any `Word.Body`, Office JS's or this package's; returns how many it highlighted. */
export declare function edit(body: Body, term?: string): number;

/** A flat OPC package in (`getOoxml()`), the edited flat OPC package out (`insertOoxml()`). */
export declare function applyEdit(ooxml: string, term?: string): Promise<string>;
