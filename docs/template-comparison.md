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
npm run clean:templates                # clear the results folders
```

Exit code is 1 when anything fails, 0 otherwise.

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
passed.

`data/` and `template/` are read and never written to.

## What gets checked

Three things per case.

**The payload against the template.** Every column the request determines,
paired on the treaty id, which both sides carry.

**The table against the template.** Four columns fewer, each left out for its
own reason:

| column | why the screen cannot verify it |
| --- | --- |
| `View Of Risk Id` | an internal id, never displayed |
| `Treaty Id` | never displayed |
| `Created At` | displayed, but in local time with the milliseconds dropped |
| `Treaty Participating Limit` | computed by the template, not shown |

All four are covered by the payload check, which reads them exactly. Nothing
goes unverified; it is simply verified from the one source that can see it.

**The fills.** A1 of these templates reads *"Only fields highlighted yellowish
are editable in the UI"*, which makes the colour a contract rather than
decoration. Checked in both directions: every column that should carry the fill
does, and no column that should not carries it. The second direction is the one
worth having — a field that quietly became editable looks exactly like a field
that was always meant to be.

## Why two sources and not one

The payload is what the client sent. The table is what a person could see. A bug
that corrupts the request is caught by the first; a bug that misrenders the
screen is caught by the second. One source alone would call either one correct,
having no second opinion.

## Rows are paired by key, never by position

The template does not write its rows in the order the request lists them. In the
first capture 36 of 68 lined up positionally and the rest had shuffled, so a
positional comparison would have reported every row after the first shuffle as
wrong.

The payload check pairs on `Treaty Id`. The table check pairs on **program name
plus treaty name**, because the screen never shows the id — and not on company
plus treaty name, which collides: one company can carry the same treaty name in
several programs, which it did on 12 of those 68 rows.

## Numbers are compared in proportion

Excel stores 15 significant digits, so a figure rebuilt in a different order
lands a few of those away with nothing having changed. These templates hold
figures from `0.0002` to `1.3e11`, and no single absolute tolerance fits both
ends: one loose enough for the large is meaningless for the small.

Values agree when they are within `1e-12` of each other in proportion — about
two orders of magnitude clear of anything recalculation produces, and far below
any difference a person made.

## Adding a template

One folder under `template_comparison/templates/`, named to match its folder in
the tree, exporting a descriptor:

```js
export default {
  sheet: 'Treaties',
  headerRow: 3,
  sources: [
    {
      name: 'payload',
      file: 'payload_data.json',
      label: 'the download payload',
      columns: [/* names, as the template writes them */],
      project: (data) => new Map(/* key -> { [column]: value } */),
      key: (cell) => String(cell('Treaty Id')),
    },
  ],
  fills: {
    editable: { argb: 'FFFFFF99', columns: [/* ... */] },
  },
};
```

`project` turns one source into the rows the template should hold. `key` builds
the same identity from the template's own cells. Nothing outside that folder
changes.
