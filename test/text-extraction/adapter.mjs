// CR-012 phases 2 and 3: the conformance suite's engine over this package's text model, a mirror of docx4j's
// TextsAdapter (GOLDENS.md, "adapter contract"): a scenario becomes options, and only the statements this
// package implements are advertised, so that every other row stays pending with its phase.
import { readFileSync } from 'node:fs';
import { WordprocessingMLPackage, Texts, PlainText, TextOptions, MarkdownExporter } from '../../dist/index.mjs';
import { HTML_READER_ABSENT } from './conformance.mjs';

/**
 * The statements CR-012 phases 2 and 3 implement (its sections 15 and 16). Not among them: REQ-203 (Strict,
 * CR-001 section 17.6); REQ-303, which matches `Requires` by prefix as REQ-303 permits (case 6 rebinds wps's
 * namespace to another prefix and expects the URI to decide); REQ-506, REQ-511, REQ-512, REQ-513 and
 * REQ-604, phase 5's; REQ-705, HTML, which this package doesn't write; REQ-1203, the whole-document goldens,
 * which hold phase 5's content; and the inspector's (REQ-901 to 907, REQ-1204).
 */
export const IMPLEMENTED = Object.freeze([
  'REQ-201', 'REQ-202', 'REQ-204',
  'REQ-301', 'REQ-302', 'REQ-304', 'REQ-305', 'REQ-306', 'REQ-307',
  'REQ-401', 'REQ-402', 'REQ-403', 'REQ-404', 'REQ-405', 'REQ-406',
  'REQ-501', 'REQ-502', 'REQ-503', 'REQ-504', 'REQ-505', 'REQ-507', 'REQ-508', 'REQ-509', 'REQ-510',
  'REQ-514', 'REQ-515', 'REQ-516', 'REQ-517', 'REQ-518', 'REQ-520',
  'REQ-601', 'REQ-602', 'REQ-603', 'REQ-605', 'REQ-606',
  'REQ-701', 'REQ-702', 'REQ-703', 'REQ-704', 'REQ-706', 'REQ-707', 'REQ-708', 'REQ-709',
  'REQ-801', 'REQ-802', 'REQ-1001', 'REQ-1202',
]);

/** A scenario's options: `default` (and `markdown`), or WD options as `name: value`, several `; ` apart. */
export function optionsOf(view, scenario) {
  const options = TextOptions.of({ view });
  if (scenario === undefined || scenario === 'default' || scenario === 'markdown' || scenario.startsWith('html reader')) return options;
  for (const option of scenario.split(';')) {
    const colon = option.indexOf(':');
    if (colon < 0) throw new Error(`Unknown scenario: ${scenario}`);
    options.set(option.slice(0, colon).trim(), option.slice(colon + 1).trim());
  }
  return options;
}

/** The engine: a package per fixture, loaded once; findings from the warnings and the metadata, as docx4j's adapter makes them. */
export function textsAdapter() {
  const packages = new Map();
  return {
    implementedStatements: () => [...IMPLEMENTED, HTML_READER_ABSENT],
    async extract(fixture, view, scenario, format) {
      if (format !== 'txt' && format !== 'md') throw new Error(`format ${format}`);
      let pkg = packages.get(fixture);
      if (!pkg) packages.set(fixture, (pkg = await WordprocessingMLPackage.load(readFileSync(fixture))));
      const result = await Texts.of(pkg, optionsOf(view, scenario));
      const findings = {};
      for (const w of result.warnings) {
        const key = `warning\t${w.kind}\t${w.selector ?? ''}`;
        findings[key] = String(Number(findings[key] ?? 0) + 1);
      }
      // the properties (core, app, custom) by their own names, as the findings table names them
      for (const [k, v] of Object.entries(result.metadata)) {
        if (v !== null && typeof v === 'object') for (const [pk, pv] of Object.entries(v)) findings[`metadata\t${pk}\t`] = String(pv);
        else findings[`metadata\t${k}\t`] = String(v);
      }
      // Markdown with the package for the quote style's indent, as docx4j's adapter writes it
      return { text: format === 'md' ? await MarkdownExporter.export(result, pkg) : PlainText.of(result), findings };
    },
  };
}

/** The agreement test's engine: the walk's plain text, and the resolved copy's read with no view logic. */
export function agreementEngine() {
  const options = (view, all) => {
    const o = TextOptions.of({ view });
    if (all) { o.include('headers', 'footers'); o.comments = true; o.hidden = true; o.fieldCodes = true; }
    return o;
  };
  return {
    walk: async (pkg, view, all) => PlainText.of(await Texts.of(pkg, options(view, all))),
    resolved: async (copy, view, all) => PlainText.of(await Texts.ofResolved(copy, options(view, all))),
  };
}
