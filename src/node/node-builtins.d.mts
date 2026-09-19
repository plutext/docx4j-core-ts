// The little of Node's API `src/opc/DirectoryPartStore.mts` uses, declared here rather than by
// depending on `@types/node`: this package has no Node types in its `tsconfig` (`"types": []`,
// `lib` es2019 + dom) because everything else in it is platform-neutral, and a types-only
// dependency for one module is not worth the install.
//
// These declarations are for this repository's own compilation only. They are not emitted (a
// `.d.mts` is never written to `dist/`), and the public declarations of `./node` name no Node
// type, so a consumer with `@types/node` — or without — sees no conflict.

declare module 'node:fs/promises' {
  interface Dirent {
    name: string;
  }
  interface Stats {
    size: number;
    isFile(): boolean;
    isDirectory(): boolean;
  }
  export function readFile(path: string): Promise<Uint8Array>;
  export function readdir(path: string, options: { withFileTypes: true }): Promise<Dirent[]>;
  export function stat(path: string): Promise<Stats>;
  export function mkdir(path: string, options?: { recursive?: boolean }): Promise<string | undefined>;
  export function writeFile(path: string, data: Uint8Array): Promise<void>;
}

declare module 'node:path' {
  export function join(...segments: string[]): string;
  export function dirname(path: string): string;
  export const sep: string;
}
