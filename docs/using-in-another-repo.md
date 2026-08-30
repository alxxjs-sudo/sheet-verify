# Using it in another repository

The tool does not need the reports to live beside it. Point it at a folder and
it reads that folder in place — so a Playwright suite that generates reports
into its own repository can be verified there, with nothing copied anywhere.

```bash
cd ~/repos/edison
npx sheet-verify output_comparison
```

That finds the cases, writes `results/` beside each one, and writes the
summaries into `!summary/` folders — all inside the Edison repository.

## Installing it

Install it **per repository**, as a dev dependency, from a packed tarball. Not
globally: a global install is ambient, nothing records it, nothing reviews
updates to it, and it drifts between machines. A dev dependency is pinned in
`package-lock.json` with an integrity hash and reviewed like any other change.

**1. Pack it, here:**

```bash
cd ~/repos/sheet-verify
npm pack
# → sheet-verify-1.11.0.tgz
```

**2. Copy that file into the consuming repository** — `vendor/` is a good home
— and commit it.

**3. Add it and install:**

```jsonc
// package.json
"devDependencies": {
  "sheet-verify": "file:vendor/sheet-verify-1.11.0.tgz"
}
```

```bash
npm install
```

**4. Add a script, so nobody has to remember the folder name:**

```jsonc
"scripts": {
  "compare": "sheet-verify output_comparison",
  "compare:clean": "sheet-verify output_comparison --clean"
}
```

### Why a tarball rather than a git URL

A `github:` dependency needs credentials for a private repository at install
time, in every pipeline that runs `npm ci` — including pipelines that never run
the comparison. A tarball needs no network and no secret, so an install cannot
fail on it. The exact bytes are in the repository and an upgrade is a commit
somebody reviews.

Upgrading is the same three steps: `npm pack` here, replace the file there,
point the dependency at the new name. Run `npx sheet-verify --version` to see
what is actually installed.

## It must not interfere with the tests

Installing rather than copying is what guarantees this, and it is worth knowing
why rather than taking it on trust.

**The tool's own tests are not in the package.** `files` is `["dist",
"README.md", "docs"]` — there is no `tests/` folder inside the published
tarball, so Playwright cannot discover them however `testDir` is configured.

**Nothing pulls Playwright in unless you ask for it.** `@playwright/test` is
imported by exactly one file, and it is reached only through the
`sheet-verify/matcher` entry point. The CLI and the main API never load it. It
is declared as an *optional peer dependency*, so the matcher uses the
Playwright the consuming repository already has rather than installing a second
copy — two copies would mean `expect.extend` patching a different instance from
the one running the tests.

**It never runs as part of `playwright test`.** Nothing is added to
`playwright.config.ts` — no `globalSetup`, no `globalTeardown`, no reporter. The
comparison is a command you run deliberately.

> **Do not copy the source in.** Putting `src/` and `tests/` into the consuming
> repository is the one arrangement that does interfere: its 302 tests land
> inside your `testDir`, its fixtures write into `tests/.fixtures/` and
> `test-results/` alongside Playwright's own output, and its `tsconfig.json` and
> `prepare` script collide with yours. Install it instead.

## What to commit, and the trap in it

The golden outputs are the contract and belong in the repository. Everything
else under the tree is produced by a run and can be produced again.

```gitignore
# sheet-verify: goldens are the contract and are committed. Everything else
# under the tree is produced by a run.
#
# The summary folder is ESCAPED on purpose. A bare `!` is gitignore's negation
# prefix, so a line reading `!summary/` does not ignore that folder -- it reads
# as "un-ignore summary/", matches nothing, and the folder gets committed.
current/
results/
\!summary/
```

Check it rather than trusting it. `git add -An` lists what a commit would take
without taking it:

```bash
git add -An
#   add 'output_comparison/meta.json'
#   add 'output_comparison/pricing_report/meta.json'
#   add 'output_comparison/pricing_report/case_001/golden/rep_1786955263151.csv'
```

Configuration and goldens, nothing else. Without the escape, `!summary/s.md`
appears in that list — silently, with no error anywhere.

## Clearing stale results

```bash
npx sheet-verify output_comparison --clean --dry   # list what would go
npx sheet-verify output_comparison --clean         # do it
```

Only folders named `results` that sit **beside a golden** are removed, plus the
`!summary/` folders. The two inputs and any `meta.json` are never touched.

This matters more in a consuming repository than it does here: a case that
stops failing leaves its old `differences.xlsx` sitting there, and a case that
is renamed leaves a whole `results/` folder behind. Both read as current.

## Verifying inside the tests instead

The command above is the usual way — the comparison is a separate activity from
the UI run, and it produces the run summary and per-type summaries. But a spec
can also assert directly, which puts a comparison failure next to the Playwright
trace of the run that produced it:

```ts
import { expect } from 'sheet-verify/matcher';

await expect(downloadedPath).toMatchCase('output_comparison/pro-forma/case_001');
```

See [the API](api.md) for `toMatchSheetBaseline`, `toMatchWorkbookBaseline` and
`toMatchCase`. Both approaches can coexist; neither excludes the other.
