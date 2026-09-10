// Packages: OpcPackage and the three Office kinds. Importing this module registers all four.
export { OpcPackage, type PackageSource } from './OpcPackage.mjs';
export { WordprocessingMLPackage, type PageSizePaper, type CreatePackageOptions } from './WordprocessingMLPackage.mjs';
export { PresentationMLPackage } from './PresentationMLPackage.mjs';
export { SpreadsheetMLPackage } from './SpreadsheetMLPackage.mjs';
export { registerPackageClass, createPackageForContentType, type PackageConstructor } from './registry.mjs';
