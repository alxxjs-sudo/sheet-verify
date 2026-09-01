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
      template/template.xlsx      what the app produced
      data/payload_data.json      what the client sent
      data/table_data.json        what the screen showed
      results/report.md           written by the run
```

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
`MetaRisk Treaty`) and the workbook decides which it is. Header on row 4,
nothing painted yellow; editable columns are marked with a trailing `*` in the
header instead. A1 here says *"Do not alter the existing "Type" row and sheet
name, as it will make the template unrecognizable by the application"*, which
makes the sheet name part of the contract — a renamed sheet fails the run.

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

**The fills, and the markers.** The two ways these templates say "you may edit
this". Checked in both directions: every column that should carry it does, and
no column that should not does. The second direction is the one worth having —
a field that quietly became editable looks exactly like a field that was always
meant to be.

**Columns the sheet computes.** `Edison Client Level GeoScope differs from
Property COE GeoScope` is a Yes/No the overrides sheet works out from two of its
own columns. No source carries it, so this is the only check it will ever get,
and getting it wrong tells a user a field was left alone when it was changed.

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

Because it is a choice rather than a property of the data, it is checkable as a
binary: the block arrives **whole or not at all**. Half a block is a template
that lost columns on the way out, and it would otherwise pass, since every
column that did arrive is correct. The divider decides which case a sheet is in.

Those 144 columns are **declared unverifiable**, with the reason recorded in the
report:

> modelling output the server joins in; in neither capture

That is measured, not assumed. Of the 144, 68 hold a value that appears nowhere
in `payload_data.json` or `table_data.json` — the server joins them in when it
builds the workbook. Verifying them needs the download *response* captured,
which nothing does yet. They are listed by name rather than as "everything after
the divider", so a column added to the block in a future release is reported as
unchecked instead of being quietly excused by a rule.

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

## Checking the checks

```bash
npm run check:templates -- <folder>
```

A comparison that reports nothing is either a clean template or a broken check,
and the two read identically. This plants one defect at a time in a throwaway
copy of each real case — changes a value, drops a row, repaints a fill, strips a
marker, half-writes a block, adds a column nobody checks — and confirms the run
notices. Faults that do not apply to a case are skipped rather than counted as
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
