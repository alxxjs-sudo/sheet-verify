# Verifying downloaded templates

The report comparison answers *"did this release change the numbers"*. This
answers a different question about a different artefact: the templates the app
hands out to be edited and uploaded back.

A template that quietly drops a column, mis-scales a percentage, or stops
marking a field editable breaks a round trip nobody was watching — and none of
it shows up in a report comparison, because a template is not a report.

## Running it

```bash
npm run compare:templates              # verify ./comparison/templates
npm run compare:templates -- <folder>  # verify somewhere else
npm run bless:templates                # take the current downloads as the contract
npm run gaps:templates                 # what the cases have never exercised
npm run check:templates                # verify the checks themselves
npm run clean:templates                # clear the results folders
```

Exit code is 1 when anything fails, 0 otherwise.

`<folder>` is the templates root — the folder holding one folder per kind. It
does not have to sit in this repository:

```bash
npm run compare:templates -- ../edison-e2e/comparison/templates
```

Naming a single kind's folder works too, for when only that one is being worked
on:

```bash
npm run compare:templates -- ../edison-e2e/comparison/templates/program_selection_template
```

This is separate from `npm run compare` in every way that matters: separate
tree, separate command, separate results. Nothing it does touches the report
comparison, and it uses none of that machinery.

## The tree

One folder per template kind, one folder per case:

```
comparison/templates/
  program_selection_template/
    <case name>/
      current/template.xlsx       what the app produced
      golden/template.xlsx        what it produced when this was blessed
      data/payload_data.json      what the client sent
      data/table_data.json        what the screen showed
      results/report.md           written by the run
```

The same `golden` / `current` shape the report comparison uses, so a case reads
the same whichever tool is looking at it. A case with an older single
`template/` folder still runs, with nothing to compare against; the report names
which layout it found rather than treating a missing baseline as a clean one.

The kind's folder name must match a descriptor under
`template_comparison/templates/`. A folder with no descriptor is **reported
rather than skipped** — a case nobody is checking looks exactly like a case that
passed. A run that recognised nothing exits 1 for the same reason.

`data/` and `template/` are read and never written to.

## The two templates covered

**`program_selection_template`** — one `Treaties` sheet, header on row 3.
Editable columns are painted yellow, and A1 says so: *"Only fields highlighted
yellowish are editable in the UI"*.

**`overrides_template`** — two shapes under one kind, because overriding a
marketplace layer and overriding a MetaRisk treaty are the same feature with
different columns. The descriptor names both variants (`MarketPlace Layer`,
`MetaRisk Treaty`) and the workbook decides which it is. Header on row 4.
Editable columns are yellow here too, but reached a different way — see below.
A1 says *"Do not alter the existing "Type" row and sheet name, as it will make
the template unrecognizable by the application"*, which makes the sheet name
part of the contract: a renamed sheet fails the run.

## What gets checked

**The payload against the template.** For program selection, every column the
request determines, paired on the treaty id. For overrides the request body is
only `{ targetType, targetIds, timezoneOffset }` — a list of ids — so it proves
membership rather than values: exactly the rows asked for, no more and no fewer.
That is not a small check. A download returning a row nobody selected, or
dropping one that was, is invisible to a value comparison, because every row it
does return is correct.

**The table against the template.** What a person could actually see. The
overrides capture is self-describing — it carries a `headers` map from its own
field names to the exact column titles the template writes — so its column list
is taken from the capture rather than transcribed into the descriptor.

**Coverage.** How many of the sheet's columns anybody actually looked at.
Without this a case prints `31 columns clean` over a sheet holding 176, and the
145 nobody checked read exactly like 145 that passed. Columns no source can see
are declared in the descriptor with a reason; anything else left unchecked
**fails the case**, so a column added in a future release is noticed the first
time it appears rather than the first time it is wrong.

**Which columns say "you may edit this".** Checked in both directions: every
column that should say it does, and no column that should not does. The second
direction is the one worth having — a field that quietly became editable looks
exactly like a field that was always meant to be. The two templates say it
differently, and the difference matters:

- **Program selection** paints the six editable columns yellow (`FFFFFF99`) as
  an ordinary cell fill, and A1 says so in as many words.
- **Overrides** has no yellow cell anywhere. Every header cell instead carries a
  conditional format — `endsWith "*"` → paint `FFFFFF00` — so the yellow is
  *computed from the header text*. `Edison Treaty Premium *` is editable;
  `MetaRisk Treaty Premium` is not.

That makes the `*` and the colour one contract rather than two, so checking the
set of starred columns checks what a reader actually sees. The rules themselves
are checked separately, because dropping them would leave every `*` in place
while nothing was yellow any more: the sheet would still list its editable
columns and silently stop showing them.

**Columns the sheet computes.** `Edison Client Level GeoScope differs from
Property COE GeoScope` is a Yes/No the overrides sheet works out from two of its
own columns. No source carries it, so this is the only check it will ever get,
and getting it wrong tells a user a field was left alone when it was changed.

The same idea reaches eight columns of the ROLePlay block, which no capture can
speak for — see below.

**The rules the sheet carries about itself.** These templates ship data
validations — `Include` is `Yes` or `No`, `Treaty Participation` is a decimal
between 0 and 1, `Treaty Premium` is at least 0. They are a contract with
whoever edits the file, since Excel refuses anything else on the way back in, so
a wrong rule breaks the upload half of the round trip while the download looks
perfect.

Two things are checked, and neither needs a source of any kind — the template
states both sides of the argument itself:

1. **Every value the template wrote satisfies the rule it wrote beside it.** A
   sheet shipping a figure its own rule rejects is a sheet nobody can edit at
   that cell.
2. **A rule written per row refers to its own row.** A formula copied down
   without being advanced validates somebody else's cell, and passes for as long
   as the two happen to agree.

`custom` rules are reported as unevaluated rather than assumed good; this is not
an Excel formula engine, and claiming to have checked something that was skipped
is the failure this whole tool exists to avoid. Their row references are still
checked, which is what (2) is for.

**The current download against the golden.** Every other check asks *is this
right*, and answers only where a source can reach. This one asks *has this
changed*, and reaches everything — all 175 columns, including the 136 no capture
can speak for. Values are paired by the same business key, never by position,
and the header fills and data validations are compared too, so a contract the
descriptor never knew about still cannot drift quietly.

The same two files are also compared **as archives** — every XML part inside the
`.xlsx`, byte for byte. That needs no spreadsheet model, and a model only sees
what it was built to see: drawings, charts, merged cells, column widths, defined
names, sheet protection, themes and print setup all live in those parts and are
invisible to every other check here. The archive comparison cannot say what
changed in a way anyone enjoys reading, but it cannot miss anything either.

Across the ten captured cases the two files differ in exactly one part,
`docProps/core.xml`, and only in its timestamps — which are normalised away, so
a changed title or creator is still compared. If the two files were written by
different producers, that is reported as the single fact it is rather than as
every part differing for one reason.

The two questions are not interchangeable, and the report keeps them apart:

> *Unchanged is not the same as correct.* A figure that was wrong when the
> golden was blessed is still wrong, and this will not say so.

`npm run bless:templates` takes the current downloads as the new contract. It is
a separate command on purpose — blessing a drift nobody has read is how a wrong
figure becomes the thing everything else is measured against. Blessing also
cannot hide a real defect: it silences the golden difference and leaves the
payload and table checks failing exactly as they were.

**Optional blocks, all or nothing.** See below.

**That the capture and the template belong together.** Both captures state what
they were taken of; the template states the same in its `Type` row. A mismatched
pair is named as one line rather than reported as a page of differences that all
have a single cause. A table capture taken mid-list (`shown` < `total`) is
likewise named, rather than surfacing as "rows the template has that the screen
does not".

## The ROLePlay block

Downloading a program selection template opens a prompt — *"Would you like to
include ROLePlay data in the template for advanced filtration?"* — and the
answer decides whether the sheet is 31 columns or 176. The extra 145 are a
divider, `ROLePlay Data ->`, and the 144 modelling columns behind it.

**The answer is recorded in the request**, as `includeRoleplayData`, so the
request is the authority on whether the block belongs:

| `includeRoleplayData` | the template must |
| --- | --- |
| `true` | carry the divider and all 144 columns |
| `false` | carry none of them |

Letting the sheet decide for itself only proves it agrees with itself. A
download that was asked for ROLePlay data and came back with none of it is
perfectly self-consistent — no divider, no block — and is exactly the bug worth
catching.

Given that it is a choice, it is also checkable as a binary: the block arrives
**whole or not at all**. Half a block is a template that lost columns on the way
out, and it would otherwise pass, since every column that did arrive is
correct.

Those 144 columns are **declared unverifiable**, with the reason recorded in the
report:

> modelling output the server joins in; in neither capture

That is measured, not assumed. Of the 144, 68 hold a value that appears nowhere
in `payload_data.json` or `table_data.json`. They are listed by name rather than
as "everything after the divider", so a column added to the block in a future
release is reported as unchecked instead of being quietly excused by a rule.

**There is no capture that would fix this.** The download flow is a POST to
`report-creation-treaties-template`, which answers with a job ticket —
`{topic, downloadName, status, id}` — then a poll on `status`, then a `GET` of
`result`, which is the `.xlsx` itself. No JSON in the browser carries these
figures; the server computes them and writes them straight into the workbook.
The template cannot verify itself, so this is a genuine limit rather than a
missing capture.

### What can be checked without a source

Eight of the 144 are arithmetic on the columns beside them — standard
reinsurance identities, holding exactly on all three captured ROLePlay
templates (17 rows, worst relative gap `3.93e-16`, which is IEEE754 rounding
and nothing else). They are checked, and reported in two groups, because the
two are not the same claim:

| | rule |
| --- | --- |
| **anchored** | `Limit` = `Treaty Occurrence Limit` |
| | `Modeled Deposit Premium @ 100% Placed` = `Treaty Premium` |
| | `Modeled ROL` = `Treaty Premium` ÷ `Treaty Occurrence Limit` |
| **consistency only** | `LOL` = `Expected Loss @ 100% Placed` ÷ `Treaty Occurrence Limit` |
| | `Loss Ratio` = `Expected Loss @ 100% Placed` ÷ `Treaty Premium` |
| | `CV` = `Standard Deviation @ 100% Placed` ÷ `Expected Loss @ 100% Placed` |
| | `Modeled Expected Premium` = `Modeled Deposit Premium` + `Modeled Reinstatement Premium` |
| | `GCMP Deposit Premium @ 100% Placed` = `GCMP USD Rate` × `Treaty GCMP Premium` |

The **anchored** three trace every input back to a column verified against the
payload, so the result is verified too. The **consistency** five tie unverified
figures together without pinning any down: three equations in four unknowns says
nothing about whether `Expected Loss` is right, only that if it is wrong then
`LOL`, `Loss Ratio` and `CV` are wrong in exactly the matching way. Worth
having, and not verification — reporting it as such would be the overstatement
this tool exists to avoid.

That takes coverage on a ROLePlay template from 31 of 175 columns to 39. Each
rule is skipped when its columns are absent, and a skipped rule leaves its
column counted as **unchecked** rather than taking the rule's word for a check
that never ran.

The remaining 136 are covered by the golden comparison — for drift, which is
what a release actually introduces. So a ROLePlay template reads:

```
39/175 columns covered, 136 declared unverifiable
175 column(s) unchanged since the golden
```

Two different claims about the same sheet, and both are true.

`Edison Program Name` appears twice in the 176-column sheet, at E3 and again at
AI3, because the block repeats the identity columns it is keyed on. Lookup is by
name and the first wins, which leaves the second unreachable — declared in the
descriptor so it reads as known, and so a name that starts repeating for some
other reason is reported rather than absorbed.

## Numbers, and what a screen can prove about them

Excel stores 15 significant digits, so a figure rebuilt in a different order
lands a few of those away with nothing having changed. These templates hold
figures from `0.0002` to `1.3e11`, and no single absolute tolerance fits both
ends. Values from the payload agree when they are within `1e-12` of each other
in proportion.

A screen is different, and looser in a way that has to be handled honestly. The
overrides table shows `USD 4.86M` for a premium of `4,862,069`. Read as an exact
number that would report a correct template as wrong; ignored entirely it would
let `4,900,000` through. So an abbreviation is read as **the band of values it
actually stands for** — `4.86M` is every figure that rounds to it, a band 5,000
wide — and the template value has to fall inside. Nothing tighter is available,
and nothing looser is defensible.

Two smaller adjustments, both properties of the screen rather than defects:

| what | why |
| --- | --- |
| runs of whitespace collapse | a browser collapses them before it paints, so the screen could not show that the sheet holds `Layer 2 -  85M XS 35M` with two spaces |
| a currency may be missing | an override is shown as `18.00M` with no currency, so the currency is not compared against a guess |

Whitespace is collapsed only for sources read off a rendered page. A request
body carries the string exactly and is compared exactly.

## Rows are paired by key, never by position

The template does not write its rows in the order the request lists them. In the
first capture 36 of 68 lined up positionally and the rest had shuffled, so a
positional comparison would have reported every row after the first shuffle as
wrong.

Program selection pairs on `Treaty Id` for the payload, and on **program name
plus treaty name** for the table, because the screen never shows the id — and
not on company plus treaty name, which collides: one company can carry the same
treaty name in several programs, which it did on 12 of those 68 rows. Overrides
pair on `GCMP Layer ID` or `MetaRisk Treaty ID`, the ids the request asked for.

## The captured files are re-saves

Every `.xlsx` under the capture tree says
`Openpyxl 3.1.5` in `docProps/app.xml`. The capture step opens the download and
saves it again, so these are not the bytes the app produced.

Confirmed to survive that round trip: cell values, solid fills, conditional
formatting, data validations, defined names. Not present in any captured file,
so it is unknown whether the originals had them: images, charts, merged cells,
sheet protection, column widths, freeze panes.

This is survivable rather than fine, and the archive comparison is why. Golden
and current both go through the same pipeline and lose the same things, so drift
in everything that survives is still caught in full. What it cannot tell you is
that a feature was already missing when the golden was blessed.

Saving the `result` response bytes straight to disk, instead of re-saving them,
would remove the doubt entirely — and the comparison would simply keep working,
since it compares whatever parts it is given. Until then, if a check that should
fire does not, confirm the feature survives an openpyxl round trip before
looking for a bug in the tool.

## Standing findings

As of 2026-09-01, four of the ten captured cases fail, and they fail on the
template rather than on the tool. A red run here is expected until the
generator is fixed.

The program selection template writes its data validations correctly for the
first two treaty rows and wrongly for every row after them:

| column | rows 4-5 | rows 6+ |
| --- | --- | --- |
| `Treaty Aggregate Limit` | a rule referring to its own row | a copy of **row 5's** rule, referring to `Q5` |
| `Treaty Participating Limit` | the same rule as `Q`, itself the wrong column | `decimal between 0 and 1`, holding 1,000,000 to 6,000,000 |

So any template with three or more treaties ships broken validations from row 6
down, which is the half of the round trip Excel enforces on upload.
`us_cat_only_program` passes only because it has a single row.

Blessing does not silence this and is not meant to: it is a standing defect, not
drift, and the golden comparison has nothing to say about it.

## What the cases have never exercised

```bash
npm run gaps:templates -- <folder>
```

A green run says the cases that exist all pass. It says nothing about the ones
that do not exist, and *"we have cases"* reads exactly like *"we have coverage"*
until somebody checks. This reads every case of a kind together and reports what
none of them contains — the list of downloads still worth capturing.

Three questions, in the order they are worth answering:

1. **A column allows a value no case has ever held.** The sheet's own list rule
   says `Include` may be `Yes` or `No`; only `Yes` has ever appeared, so the
   other is a path through the app no capture has taken.
2. **A column is empty in every row of every case.** Nothing it does has been
   seen at all.
3. **A column holds one value everywhere.** Its comparison has never had to
   distinguish anything, so it would pass while broken.

Everything is measured from the files, so nothing goes stale: as cases are added
the gaps close on their own, and a column that starts varying stops being
reported without anyone editing a list. Findings are ranked, with editable
columns and payload flags first — a gap on a read-only column is usually just
what the test data happens to be, while an editable one never filled is a field
the round trip has never carried.

It never fails a build. It is a worklist, not a verdict.

## Checking the checks

```bash
npm run check:templates -- <folder>
```

A comparison that reports nothing is either a clean template or a broken check,
and the two read identically. This plants one defect at a time in a throwaway
copy of each real case — changes a value, drops a row, repaints a fill, strips a
marker, drops the rule that paints the yellow, half-writes a block, removes a
block that was asked for, breaks a data validation, drifts a column no source
can verify, merges two cells nothing here models, adds a column nobody checks —
and confirms the run notices.

Each fault is judged against **the case's own baseline**, not against zero: a
case that already carries a real finding can still be used to check that the
other checks work. Faults that do not apply to a case are skipped rather than counted as
passes, and the case itself is never written to.

Nothing in it knows a column name. Every fault is built from the descriptor and
from whatever the sheet turns out to hold, so a template added tomorrow is
covered without teaching it anything.

## Adding a template

One folder under `template_comparison/templates/`, named to match its folder in
the tree, exporting a descriptor:

```js
export default {
  sheet: 'Treaties',        // or `variants: [{ sheet, ... }, ...]`
  headerRow: 3,
  rowMarker: 'Treaty Name', // the column every real row fills
  sources: [
    {
      name: 'payload',
      file: 'payload_data.json',
      label: 'the download payload',
      columns: [/* names, or { name, compare } */],
      project: (data, columns) => new Map(/* key -> { [column]: value } */),
      key: (cell) => String(cell('Treaty Id')),
    },
  ],
  fills: { editable: { argb: 'FFFFFF99', columns: [/* ... */] } },
  markers: { suffix: ' *', columns: [/* ... */] },
  conditionalFills: { editable: { argb: 'FFFFFF00', type: 'endsWith', text: '*' } },
  derived: [{ column, from: [a, b], value: (x, y) => /* ... */ }],
  blocks: { 'name': { lead: 'divider column', columns: [/* ... */] } },
  unverifiable: { 'the reason': [/* column names */] },
  duplicateHeaders: [/* names the sheet writes twice on purpose */],
};
```

`project` turns one source into the rows the template should hold; `key` builds
the same identity from the template's own cells. Both `columns` and `project`
may be functions of the capture and of the sheet's own column names, for a
capture that describes itself. Nothing outside that folder changes, and neither
does sheet-verify.

Then run `npm run check:templates` against it, and confirm the faults that apply
are caught.
