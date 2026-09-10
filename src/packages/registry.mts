import type { OpcPackage } from './OpcPackage.mjs';

export type PackageConstructor = new () => OpcPackage;

const byMainContentType = new Map<string, PackageConstructor>();
let generic: PackageConstructor | undefined;

/** Called by each package module on evaluation; keeps `OpcPackage.load` free of subclass imports. */
export function registerPackageClass(mainContentTypes: readonly string[], ctor: PackageConstructor): void {
  for (const ct of mainContentTypes) byMainContentType.set(ct, ctor);
}

export function registerGenericPackageClass(ctor: PackageConstructor): void {
  generic = ctor;
}

/** The package class for a main part's content type; `OpcPackage` for anything else. */
export function createPackageForContentType(mainPartContentType: string | undefined): OpcPackage {
  const ctor = (mainPartContentType !== undefined ? byMainContentType.get(mainPartContentType) : undefined) ?? generic;
  if (!ctor) throw new Error('No package class registered');
  return new ctor();
}
