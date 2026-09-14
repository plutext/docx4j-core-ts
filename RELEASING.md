# Releasing

`@docx4j/core-ts` is published to npm from GitHub Actions (`.github/workflows/push-to-npm.yml`), like
`@docx4j/generated-objects-ts`, `@docx4j/jsonix` and `@docx4j/jsonix-schema-compiler`.

## Publishing target

Publishing uses npm **trusted publishing**: no npm token is stored anywhere. The settings of the npm
package `@docx4j/core-ts` (npmjs.com, scope `@docx4j`) name the publisher GitHub Actions and the GitHub
repository allowed to publish it: organization or user `plutext`, repository `docx4j-core-ts`, workflow
`push-to-npm.yml`, no environment, with **Allow npm publish** ticked. Without that permission the run gets
its OIDC token and signs provenance, then `npm publish` fails with `E403 ... OIDC permission denied for
this action` (fix the setting and re-run the failed job; nothing was published). The "Publishing access"
setting (2FA, tokens) does not affect trusted publishing; "Require two-factor authentication and disallow
bypass 2fa tokens" leaves CI, or a manual publish with a 2FA code, as the ways to publish. Renaming the
workflow file breaks publishing until the npmjs.com setting is changed to match.

A trusted publisher is configured on an existing package, so the first version (0.1.0) is published by
hand; see "First release" below.

npm versions cannot be reused once published (even after an unpublish): a problem found after
publishing ships as the next patch version.

The package ships `dist/` (built from `src/`), `README.md`, `LICENSE` and `NOTICE`. Its dependencies are
`@docx4j/generated-objects-ts` (which brings `@docx4j/jsonix`) and `fflate`.

## Dependencies

A release that needs a change in `@docx4j/generated-objects-ts` waits for that package's release: publish
it first (its own `RELEASING.md`), then raise the range here (`npm install @docx4j/generated-objects-ts@^X.Y.Z`),
typecheck, test, and commit `package.json`. `@docx4j/core-ts` is never published against
an unreleased objects package.

## First release (0.1.0)

`package.json` already carries 0.1.0, so there is no version change. From the repository root, on a
clean `main` whose CI (`test.yml`) has passed:

```bash
rm -rf node_modules dist && npm install
npm run typecheck && npm test
npm pack --dry-run                 # @docx4j/core-ts 0.1.0: dist/, README.md, LICENSE, NOTICE, package.json
git tag -a 0.1.0 -m "Version 0.1.0"
git push origin main 0.1.0
npm login                          # an npm account with publish rights on the @docx4j scope
npm publish --access public        # prepublishOnly runs typecheck and test again
```

Then, on npmjs.com, add the trusted publisher to the package's settings (as above, including **Allow npm
publish**). Do **not** create a GitHub release for 0.1.0: it would run `push-to-npm.yml`, which fails
because 0.1.0 is already published. Every later version follows "Steps".

## Steps

```bash
# 1. If the release needs a newer @docx4j/generated-objects-ts, release that first and raise the range
#    here (see "Dependencies").

# 2. Set the version, the next one after the latest on npm (npm view @docx4j/core-ts version);
#    no tag or commit yet
npm version 0.1.1 --no-git-tag-version

# 3. From a clean node_modules, check against the registry dependencies and inspect the package
rm -rf node_modules dist && npm install
npm run typecheck && npm test
npm pack --dry-run

# 4. Commit, tag and push (the tag is the bare version, as the workflow checks it; a leading v is tolerated)
git commit -am "Version 0.1.1"
git tag -a 0.1.1 -m "Version 0.1.1"
git push origin main 0.1.1
```

5. Publish: create a GitHub release for the tag at
   https://github.com/plutext/docx4j-core-ts/releases/new (choose the existing tag, title = the version,
   leave "Set as a pre-release" unticked, then **Publish release**; a saved draft does not publish).

Publishing the release runs `push-to-npm.yml`, which installs, fails unless the release tag
equals `package.json`'s version, runs typecheck and test, checks the tree is unchanged, and runs `npm pack`
and `npm publish` (with provenance, via OIDC).

If the workflow fails before the publish step, fix the problem, move the tag, and re-run it from the
Actions tab (or delete and recreate the release). If only the publish step fails (for example the E403
above), fix the npmjs.com setting and use **Re-run failed jobs**; the tag and release stay.
