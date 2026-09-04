# Changelog

Versions follow the policy in [the README](README.md#versioning): a major is
reserved for changes to configuration keys, CLI flags, exports and artefact
shapes. Detection changes ship as minors, each saying what to recheck.

## 1.13.0 — 2026-09-03

### A run says which report types nobody has configured

`--write-meta` writes the `meta.json` a report type would otherwise be typed out
by hand, and it existed for months while being mentioned in exactly one place:
`--help`. Trees went unconfigured because nobody knew the command was there.

Every run now ends by naming the type folders holding cases with no `meta.json`
of their own, and printing the command:

    2 report type folder(s) have no meta.json:
      output_comparison/marginal_analysis
      output_comparison/policy_ranking
      ...
      npm run write:meta -- output_comparison/marginal_analysis

It is a sibling of the existing "held no cases" note: a gap a green run would
otherwise hide. Not a failure, no effect on the exit code, and silent the moment
the file exists -- because a tree can be green and unconfigured at the same time,
and passing is not evidence that the configuration is right.

### An agent for the part that needs judgement

`propose.ts` has always said what it will not do:

> The one thing deliberately left out is per-sheet configuration -- header rows,
> end rows, keys. Those are what a report type genuinely needs a human for.

That gap has a cost, and one tree measured it: fifteen cases failing because no
one had named `Return Period` as a key, and four more because a header row held
the run date and so a column was literally named `2026-09-02`.

`.claude/agents/meta-config.md` is a Claude Code subagent for exactly that
layer. It runs `--write-meta` for the evidence, reads the worklist `report.md`
already prints, decides the keys and header rows, writes its reasoning into the
`//` comments, then re-runs the folder and reports before and after.

It inherits one rule verbatim. Keys are safe to get wrong -- a key column that
does not exist fails loudly. `metadata` is not, because it makes the tool go
quiet, so nothing the figures depend on goes in it and every entry is justified
with both observed values.

### A formula difference that printed the same string twice

Four cases failed on this row, and it says nothing:

    | Row | Column | Cell | Golden | Actual |
    | #4 | Non-client data | `B5` | `"Exposure Data as of "&TEXT(A1,…)` | `"Exposure Data as of "&TEXT(A1,…)` |

The finding is real. Formulas are compared as the tool resolves them, not as
Excel writes them -- a reference becomes `[column name]@row±n`, which is what
lets a formula survive its table moving down the sheet. Here `A1` is the
report's header cell and it holds the run date, so the column it names is
called `2026-09-02` in one file and `2026-09-03` in the other. The formula
points at a renamed column; its text never changed.

The markdown report printed the text. The plain-text report has always fallen
back to the resolved form when the two agree, and this one now does too, with a
line saying which it is showing and why.

### A bare run picks the tree that is there, and asks when there are two

One checkout holds two comparison trees now, one per source system, and that is
the right shape: each carries its own root `meta.json` — its own run identity,
its own tolerances — and merging them applies one system's rules to the other's
reports.

The cost of getting it wrong is not a near miss. The same nine cases, the same
goldens, compared under the two roots:

    edison_output_comparison/…/us-cat-and-risk-air-tsv12-market  No defects.
    catwb_output_comparison/…/us-cat-and-risk-air-tsv12-market   Differences found.

One cell apart — `Edison Build`, `20260901.4120.1` → `20260902.4135.1` — which
the Edison root sets aside as run identity and the CATWB root has never heard
of.

So a bare `npm run compare` now looks for `output_comparison`,
`edison_output_comparison` and `catwb_output_comparison`, runs the one it
finds, and where more than one is present names them and stops rather than
choosing. A silently-wrong tree is a green run that means nothing.

`compare:templates`, `gaps:templates` and `check:templates` default to
`edison_output_comparison/templates`, which is where that tree now lives.

### "Nothing checked them" said the one thing that was not true

The heading over layer 2's coverage gap read:

    ## Differing, and nothing checked them (17)

Which contradicts itself: if nothing had checked them, nobody would know they
differ. Layer 2 checked them -- that is how they got there. What did not reach
them is layer 1, because their table had no row key. Read literally, the old
heading said the tool had skipped 17 cells it had in fact examined, and on a
report with a five-hundred-cell gap list that is an alarming thing to say.

It now reads `## Differing, outside the keyed comparison`, and the paragraph
under it names both layers. The same wording follows through the run log, the
run summary and the per-type summaries, which all carried `nobody checked`.

No behaviour changes; the counts are the ones they always were.

### A clean verdict over a six-figure count no longer reads as a contradiction

A report opened with **Identical.** and then, four lines down:

    | total | within tolerance | above tolerance |
    | 50,274 | 0 | **50,274 (100%)** |

Both are true. The counts are layer 2's, and layer 2 compares by address: an
8,476-row CSV that both files held with the rows in a different order is
identical to layer 1, which pairs by key, and wall-to-wall different to layer 2,
which does not. Layer 2 never decides the verdict, so the case passes -- but
nobody reading that table would have believed it.

The table now says so where the two disagree, and only there.

### A file nobody compared no longer reads as a pass

A case held `results.csv`, `details.zip` and `unused.zip` on each side. The run
compared the CSV, reported **Identical**, and said nothing whatever about the
other two -- which held a `policy.csv` and an `unused_policy.csv` of real data.

The folder walker filtered to spreadsheets before it looked, so the archives
were never even seen. Anything in a `golden/` or `current/` folder that no
comparison read is now named, in the log and in `report.md`, and fails the case:
a `golden/` folder exists to hold one side's output, and a verdict has to mean
the whole case was checked.

### One case's settings no longer stop the tree

A generator wrote its own `metadata` -- an object describing the analysis --
into a `case.json`. This tool's `metadata` is a list of cells to read but not
judge, so the run died with `(given.metadata ?? []) is not iterable`, naming
neither the file nor the key. Thirty-six cases, none compared, over one file.

Recognised keys are now checked for shape, with a message naming the file, the
key, what was found and what was wanted. A key this tool does not recognise is
still left alone, exactly as before -- a `case.json` is a good place to describe
a case to whatever generates it. And a case whose settings cannot be read is now
one broken case, not a dead run.

### analysisType and entityType name a report type too

Six of twelve type folders in one tree said `analysisType` or `entityType`,
because a Conditional EP is an analysis and a Data Transmittal is an entity.
Every one of them was filed under *Unspecified report type* -- the heading that
exists to mean nobody set this, printed over folders where somebody plainly had.
Both are now read; `reportType` still wins where more than one is set.

## 1.12.0 — 2026-09-03

### report.md is the verdict, not the archive

One case produced 18,661 lines and 1.6 MB, 92% of it a single section. The rule
that got it there was deliberate and written down: nothing is truncated, because
"and 34 more" hides exactly the row somebody wanted.

It held right up until the file stopped being read. Nobody opens 1.6 MB of
markdown, so in practice everything was hidden rather than one row -- and the
complete record already sat beside it, in a form better suited to the job:
`differences.xlsx` holds one row per differing cell, sortable and filterable,
`compared.xlsx` holds every cell compared, `diff.json` holds the lot structured.

So the row-by-row listings are now capped at ten, each saying how many it left
out and which file holds them. The same case is 1,212 lines.

Never capped, whatever the setting: the counts, the per-column tallies -- the
part that says which figure moved -- and a new **Where the differences are**
table ranking every failing table by how much it differs. On the case above that
one table says 16,099 of 16,990 findings are in a single sheet, which the old
layout answered only by scrolling until the sections stopped.

The paragraph explaining what the two layers are is identical in every report
ever written, so it moves to `docs/reading-a-report.md` and the numbers stay.
That reclaims thirteen lines between the reader and the findings.

`--detail full` restores the old output exactly, for anyone who built on it.

## 1.11.1 — 2026-08-31

### A table that failed on errors now says so

The breakdown printed under a failing case named six measures -- values,
formulas, types, invariants, rows and columns -- and `errors` was not one of
them. But errors count towards `defects` like any difference, so a table fails
on them alone: a key column the config names and the sheet does not have, a
formula with no cached value.

Such a table appeared in the breakdown with a zero under every column and
nothing saying why. Reported from the field as "why do we display Geography if
it's 0 under Values", which is exactly the right question -- the row announced a
failure and then declined to explain it.

The quieter half was worse. Rows are ordered by the sum of those measures, so a
table failing only on errors scored zero, sorted last, and with five other
failures dropped off the end of the table entirely. The table nobody could
compare is the one most worth naming, and it was the first to go.

`errors` is now a measure. The column appears only when something is non-zero,
so a run without errors prints exactly what it printed before.

## 1.11.0 — 2026-08-31

### `relativeTolerance`, for reports that span magnitudes

A workbook holding both `0.0002` and `126,339,393,111.699` cannot be judged by
one absolute tolerance, and until now that was the only kind there was.

Excel stores 15 significant digits. What that buys in absolute terms depends
entirely on magnitude: at `1.3e11` the finest step it can represent is a
*thousandth*, so a total rebuilt in a different order lands whole thousandths
from the one stored last month with nothing having changed — and the default
tolerance of `0.001`, which is one such step, calls it a difference. At
`5,000,000` the same `0.001` is a hundred thousand times looser than the
format's own precision, and an error of a cent goes unreported.

Measured across this project's cases, the effect is stark. Among cells whose gap
is pure recalculation drift, the *absolute* gaps span fourteen orders of
magnitude while the *relative* gaps span two — all of them between `1e-16` and
`5.4e-14`, which is to say they are one phenomenon and only the size of the
number underneath makes them look different. Judged absolutely, five drifting
cells were absorbed in silence while another with *less* proportional drift was
reported as a difference: the same wobble, opposite verdicts, decided by nothing
but where the decimal point fell.

A cell may now also be forgiven for being close in proportion:

```
|golden − actual|  ≤  max( tolerance, relativeTolerance × max(|golden|, |actual|) )
```

`tolerance` remains the floor, because a proportion of almost-nothing is
meaningless and a value near zero needs an absolute answer. `1e-12` is a
reasonable starting point for recalculation drift: about 19× clear of the worst
drift measured, and ten orders of magnitude below any difference a person made.

It takes the same shapes as `tolerance` — one number, or a record keyed by
column with `*` as the fallback — merges the same way through `defaults` →
sheet → table, and applies in both layers, so a gap layer 1 forgives no longer
turns up in layer 2's headline count.

**Nothing changes unless you ask for it.** `relativeTolerance` defaults to `0`,
which reduces the rule to exactly the comparison that ran before it existed.

This is deliberately not the hidden float slack that used to live in
`equalValues` and was removed in 1.4.0, though the arithmetic is a cousin. That
one was on for everybody, was never written down, and could not be seen doing
its work — "nobody could say what it had swallowed" was the reason it went. This
one is off by default, is written in a `meta.json` where it can be read, and is
reported per cell: because a relative rule resolves to a different allowance on
every value, `differences.xlsx` and `report.md` now carry the allowance *that
cell* was measured by rather than the setting its column was given. Printing the
column's number would show a figure no cell was ever judged against.

### To recheck

Nothing, unless you set it. If you do, read the "Inside the tolerance you set"
section of a `report.md` afterwards and check the gaps it now absorbs are ones
you meant to give away — that section exists so a tolerance set too wide is
visible rather than silent.

## 1.10.2 — 2026-08-30

### A case.json may also describe the case to whatever generates it

The check that catches per-sheet settings written at the top level rejected
*every* key it did not recognise. That is a different and wider rule than the
mistake it was built for, and it refused a legitimate file: a `case.json` in a
Playwright repository carrying the automation's own description of the case --
the datasets it runs on, the units, the company name, the environment. One file
describing one case to both the generator and the verifier is the right shape,
and this tool has no business failing a run over fields that are not its.

It now objects only when a top-level value is *shaped like a sheet's settings*:
a plain object naming at least one of `keyColumns`, `headerRow`, `tolerance`,
`columns`, `tables`, `invariants` and the rest. So this is still refused, and
named on its own rather than alongside the automation's fields:

    { "dataSets": [...], "Occupancy": { "keyColumns": ["Portfolio"] } }
      -> has a sheet's settings at the top level, where they do nothing:
         "Occupancy"

while `dataSets`, `settings` and `configuration` beside it are left alone. The
protection is unchanged for the mistake it exists to catch; only the false
positive is gone.

## 1.10.1 — 2026-08-23

### A report type's summary moves into that report type's folder

1.10.0 put them all beside the run summary. Eleven report types is twenty-two
files in one folder, which is more to scan than it is worth — and the folder
that sorts to the top of the tree is the one place that should stay short.

Each type's pair now goes in a `!summary/` inside that type's own folder, which
is the folder already being worked in:

```
output_comparison/
  !summary/run-summary.md          the whole run, as before
  pro-forma/!summary/Pro Forma.md  this type alone
  pro-forma/case_001/
```

The file keeps the type in its name even though the folder says it too. That is
redundant in the tree and worth it the moment the file is detached and mailed:
three attachments called `summary.xlsx` are worse than one redundant name.

The folder chosen is the one whose `meta.json` named the type. A `reportType`
set in a `case.json` is ignored for placement — a summary of a report type does
not belong inside one of its cases. Two folders declaring the same type share
one summary, placed above both.

`npm run clean` removes every `!summary/` folder, not only the one at the root.
A full run clears the files 1.10.0 left at the root, so no manual tidying.

## 1.10.0 — 2026-08-23

### A summary for each report type, not only for the whole run

`!summary/` now holds a `.md` and `.xlsx` per report type beside the run-wide
pair, named after the type's `reportType`. The run summary answers "how did the
run go"; a type's own file answers the question asked next, and the one that
actually gets forwarded: how did *my* report type go, without the ten types the
reader does not work on. It carries the same case table with the totals spelled
out rather than left to be added up.

A type with no `meta.json`, or one that never set `reportType`, was grouped
under its folder path. It is now **Unspecified report type**, with the folder
path kept inside the name so two unnamed types stay two groups and two files.
That is a visible change to the headings in `run-summary.md` for any tree that
never named its types.

Each type's file is rewritten only by a run that covered that type, so a
narrowed run leaves the others describing their own last run. A run over the
whole tree also clears summaries for types that no longer exist — a renamed type
otherwise leaves a file reading exactly as current as the ones beside it.

### A failing case says which tables failed

The log gave a count and a path: `18 tables failing`, then `report.md`. Between
them sat the question everyone asked next, and a colleague reading only the
terminal concluded the tool had missed differences it had caught and written
down. A failing case now prints its worst five tables with the kind of
difference beside each — values, formulas, types, invariants, rows, columns.

Columns of nothing but zeroes are dropped, so value differences stay two columns
wide and only a report that changed shape widens it.

### Logging, elsewhere

- The report path in the log was absolute while every other path was relative;
  it wrapped the line and buried the block it was meant to close.
- A run says how long it took, which is the difference between waiting for a
  plain run and waiting for `--recalc`.
- A run that found differing cells nobody checked says so on its last line.
  That total lived in each case's report, which is not read on a green run.
- `--bare` counted files it rewrote and printed a line for each file it did not.
  Over forty cases that was eighty lines saying nothing happened. Counted now.
- `npm run summary` said how many cases it read and not how they stood.
- A rebuilt summary printed **0 differing cells nobody checked** for every case.
  `diff.json` holds layer 1 only, so a rebuild never had that number — and a
  zero it did not measure is the exact failure this tool exists to prevent. It
  now leaves the column empty and says why.
- A report type folder holding no cases at all was passed over in silence. With
  no cases it appears in no summary either, so a type nobody is checking looked
  exactly like a type where everything passed. Named now, without failing the
  run — a folder can be waiting for its first download.
- `npm run clean` printed a line per folder; grouped by report type, with the
  space each freed. `--dry` still lists every folder by name.
- `npm run links` groups broken links under the file that holds them, since a
  move between documents breaks several at once and reads as one edit.

## 1.9.1 — 2026-08-20

### Two pieces of documentation that had stopped being true

`headerName` still described the fallback it no longer has -- "failing that the
formula itself, which is stable if ugly" -- directly above the code that removed
it, and directly above the comment explaining why. A doc comment describing
behaviour that was taken out is worse than none: it is the version somebody
believes.

And the README pinned a test count, which is a number that goes stale on every
release and tells a reader nothing they need. `npm run check` reports it.

## 1.9.0 — 2026-08-20

### A totals row of formulas is not a row of column names

Reported from a real case as columns flagged added and removed in their
hundreds. Detection had named the columns after a row of formulas:

    headers: ["Total", "SUM(B15:B18)", "SUM(C15:C18)", "IFERROR(C14/B14,0)"]

`headerName` fell back to a formula's own text when the cell had no computed
result -- which is every formula in these reports. That was added so a heading
*built* by a formula still names its column, and the fallback was justified as
"stable if ugly". It is the opposite of stable: `SUM(B15:B18)` carries row
numbers, so inserting one row above renames every column in the table and both
files report the lot as removed and added.

Worse, it made a totals row look like twenty-five names against the real header
row's seven, so the header search preferred it -- taking the key column with it,
which became `Total`.

A formula that builds header text contains that text; one that computes a number
does not, and has no name to give. The fallback is now the leading string
literal or nothing.

That distinction had to be split from a second use of the same function.
Detection reads the whole sheet through it to decide which cells are *occupied*,
and there a formula with no result is emphatically not an empty cell -- a row of
them reading as blank splits the block it sits in. Occupancy and naming are now
two questions with two answers.

### Banded shading is not a header row

The same case, another sheet: rows 32, 34, 36 and 38 fully painted, rows 31, 33,
35 and 37 not. Zebra striping paints a row exactly the way these generators
paint a heading, so the search took the first banded row as the header. The key
column became `US - Northeast` -- a region name lifted from the data -- and the
rest of the columns were named after that row's figures.

Told apart by shape rather than by paint: a header row is text where its data is
numeric, so a candidate whose numeric columns are exactly the next row's is
data. It must hold a number itself for this to fire, or a heading of plain words
over rows of plain words would be refused.

### A numeric heading that drifted is still the same column

    removed: 788321.400221      added: 788321.4002209998     gap 1.16e-10
    removed: 33792307.66114401  added: 33792307.661144       gap 7.45e-9

Columns headed by computed figures recalculate. Compared as text those are two
different columns; a tolerance would have forgiven the same drift instantly had
the number been a value rather than a name. Numeric headings are now canonical
to twelve significant figures -- Excel keeps about fifteen, drift shows past
twelve -- so drift collapses and genuinely different headings stay apart.

### A block whose top row names nothing keeps its rows

Fixing the above meant fewer rows qualified as headers, and a block whose header
names fewer than two columns was being *dropped* -- 26 tables across the tree,
silently. The rows are still there and still comparable; they just have no names
of their own. They get positional ones now.

Measured over 38 real cases: **1,313 tables compared, up from 1,290**, with
**column churn nearly halved, 690 to 368**. Nothing lost -- tables not compared
and integrity errors both unchanged.

The 368 that remain are one case whose two files were produced differently: the
golden carries none of its 3,721 formula results and the current carries 3,271
of them, so a formula heading renders as a value on one side and as nothing on
the other. That is what `npm run bare` is for, and the run already warns about
the pair.

## 1.8.0 — 2026-08-20

### `clean` said it removed folders it had not removed

Reported from a real tree, with the output to prove it:

    removed  output_comparison/pro-forma/case_001/results
    [Error: EPERM: operation not permitted, rmdir '.../case_001/results']

It printed the line *before* attempting the delete, so a folder Windows refused
to release was announced as gone. The throw then took the whole script with it,
leaving every later folder untouched and unmentioned. A clean that says it
cleared the tree and cleared part of it is worse than one that fails outright,
because the run afterwards looks trustworthy -- which is how stale results came
to be mistaken for a cache.

Three changes:

- **Delete first, report second.** Only what actually went is called removed.
- **Retry.** The failure is usually transient: Windows marks a file
  delete-pending until the last handle closes, so the contents go and the
  `rmdir` behind them fails -- which is the exact shape of the error above.
  Eight attempts, 150ms apart.
- **Carry on, then fail loudly.** One locked folder costs that folder, not the
  rest. What could not be removed is listed with its error code, the likely
  causes are named, and the exit code is 1.

### `report.md` says which files it read

Not a cache, though it looked like one: swap the pair in a case folder, fail to
clear the old results, and the report still names the same two paths, so a stale
report is indistinguishable from a fresh one. Nothing in this tool reads its own
results -- a run only writes to that folder -- but there was no way to tell that
from the outside. The header now carries the stamp:

    | golden | ...\golden.xlsx — 22 sheet(s) — 181,914 bytes, modified 2026-08-17 11:20:01 |

Swap a pair and the size and time move. If they do not, the report is describing
the old files.

### `npm run bare`

`--bare` had no script, so the one command for a report accidentally saved in
Excel was the one that had to be typed out in full.

## 1.7.1 — 2026-08-20

### "33 sheets failing" on a report with 22 sheets

Reported from a real run, and the number was not wrong -- the noun was. There
are two units in play and the one-line summary mixed them.

`sheetSchema` counts sheets: a sheet added, removed or moved is a sheet.
`w.sheets` holds one entry **per table**, so a workbook of 22 sheets with
several tables on some of them has 49 of them. Counting those as sheets
overstates the damage and, worse, cannot be reconciled against the file by
anyone reading it -- which is exactly what happened.

The inconsistency was sitting inside the same function. `review` counts that
same collection and has always called them tables; `failed` and `skipped` called
them sheets. So a summary could read "33 sheets failing, 3 tables to review"
with both numbers drawn from the same list.

Now:

    1 sheet removed, 2 tables failing, 1 sheet moved, 1 table to review

Sheets where they are sheets, tables where they are tables. The same correction
reaches the run summary's columns, which were `Sheets compared` / `Sheets
failing` and are now `Tables compared` / `Tables failing`, and `report.md`'s
**Sheets to review** heading, which listed added *sheets*, uncompared *tables*
and moved *sheets* -- two of the three not sheets. It is **What to review** now.

`report.md` was right throughout: it has always said "tables compared" and
"tables not compared". Only the one-liner was wrong, which is the line that
reaches the run log and the run summary, and therefore the one most people read.

## 1.7.0 — 2026-08-20

### `npm run summary`: rebuild the overview without re-comparing

The run summary was only ever a side effect of a comparison, and that left it
drifting out of date through ordinary use. A run narrowed to one report type
replaced the tree-wide summary with its own three cases, and a single-case run
wrote nothing at all -- so the file at the root went on claiming thirty-four
cases while describing a run that was nothing of the sort.

Three changes, and the first two matter more than the new flag.

**A narrowed run says it was narrowed.** Without it, "3 case(s)" reads as the
size of the tree rather than the size of the run.

**A single-case run writes the summary too.** The case's own report says more,
but skipping it left the *previous* run's summary sitting at the root, current
to look at and wrong.

**`--summary` rebuilds from what is already on disk**, comparing nothing. It
walks the whole tree however narrow the last run was, reads each case's
`diff.json`, and takes about three seconds where re-running the tree is a minute
and `--recalc` is half an hour. A case that has never been run is listed as
**never run** rather than omitted -- a case missing from an overview is
indistinguishable from a case that passed.

Two things bit while building it. `readJson` validates a *config* and rejects
unknown keys, so reading `diff.json` through it threw on every case and the
rebuild reported all thirty-four as never run; it reads plainly now. And the
scope note landed without a blank line before it, which glued it to the line
above in every markdown renderer.

## 1.6.1 — 2026-08-20

### The summary folder sorts to the top

`results/run-summary.*` landed alphabetically between `pro-forma` and
`riskplay_report` -- in the middle of the report types, for the one artefact
that should be read first. It is now `!summary/`.

`_summary` was the obvious choice and is wrong. Windows orders a leading
underscore *after* letters, so it sorted to the bottom of the tree, below every
report type -- worse than where it started. Checked rather than assumed, against
a folder of candidates:

    !summary  0-summary  00-summary  @summary  comparison_report  …  _summary  ~summary

`!` comes ahead of both letters and digits, which is why it is the Windows
convention for pinning a folder to the top.

A run given `--results <name>` puts its summary in `!summary/<name>/`, so a
plain run and a `--recalc` run keep both rather than one overwriting the other.

## 1.6.0 — 2026-08-20

### A run summary: how the whole run went, by report type

Every artefact so far describes one case. That is the right shape for fixing
something and the wrong shape for the question asked first -- *how did the run
go?* -- whose answer lived only in terminal scrollback, gone as soon as anyone
scrolled and impossible to send to somebody who was not watching.

A run covering more than one case now writes `results/run-summary.md` and
`results/run-summary.xlsx` at the root of what was run.

Grouped by report type, because that is the unit people work in: a release
breaks a *kind* of report, and eight failures on one type with the rest clean is
a different morning from one failure on each of eight types. Failures sort
first inside each type, so nobody scrolls past twelve ticks to find them. A type
that never named itself gets its folder path.

The workbook opens on an **Overview** sheet -- every type, its counts, and a
totals row -- then a sheet per type carrying more than the markdown does: sheets
compared, sheets failing, tables not compared, unchecked differing cells,
whether the run was recalculated, and the path to each case's own report.

**Unchecked differing gets its own column and its own callout**, because it is
the one number that can be non-zero on a case that passed. That is what a table
with no row key looks like from outside, and it is the only way a green run can
be hiding something.

`clean` removes the summary with the results. One left from the last run reads
as current exactly as loudly as a stale `results/` folder does.

Built and immediately caught doing something worse than the problem it solved:
written at the *root* of the tree, `run-summary.xlsx` is a spreadsheet with no
golden beside it, so the next run read the root as a broken case and refused to
start -- *"no golden file. Rename one of [run-summary.xlsx]"*. It now goes in
the results folder, which the case walker already skips by name, and there is a
test pinning that a second run still works.

### Every script says what it does

`package.json` grew from nine scripts to twenty-four across two releases with no
explanation of any of them. The README now carries a table: one line each, and a
note that arguments need npm's `--` separator.

## 1.5.1 — 2026-08-20

### Nothing identifying ships with the package

`docs/` is listed in `package.json` `files`, so everything in it goes to anyone
who installs this. `docs/case-labels.md` was in there, and it records real
report names from the source system. It has moved to `notes/`, which does not
ship, and carries a header saying why.

Product and system names -- a portal, a model vendor, three report-type names --
are out of the docs, the source comments and the fixtures, replaced with neutral
equivalents that make the same point. A wildcard covering `Summary Report Name`,
`Quarterly Report Name` and `Regional Report Name` illustrates itself exactly as
well as the real three did. Two fixtures also carried a real person's name,
lifted from a report and used as filler; the tests assert on the label and never
the value, so there was no reason to keep it.

`notes/case-labels.md` is now the only file in the repository holding source
names, deliberately: stripping them would destroy what it is for, which is
matching a label back to a re-download. It says so at the top.

### The limitations list caught up with the tool

Reviewed every entry against what actually shipped this week.

**Cached values** was the big one and is now answered twice over -- `--recalc`
for the exact numbers, and the "Will recalculate" sheet for which cells move
when Excel is not available. The entry also carries the measurement that makes
the point: not *some* generators write formulas without results, but every one
of 379,959 formula cells on a real tree.

**Two limits added that nothing had written down.** Computed references --
`OFFSET`, `INDIRECT` -- are opaque to the impact graph, and 62,406 of those
formulas are `OFFSET`, so this is load-bearing rather than theoretical. And
`--recalc` needs Windows, Excel, and Excel closed, each of which it refuses on
rather than quietly comparing the files as they arrived.

**One entry was simply out of date.** "ExcelJS is dormant. No release since
October 2023" -- checked against the registry: 4.4.0 is indeed October 2023, but
there was a `4.4.1-prerelease.0` in December 2024. Reworded, and it makes the
point better: twenty months have passed since even the prerelease.

The rest were re-checked and still hold: `endRow` is still Excel-only, the
unknown-key check is still top level only, table numbers are still positional.

## 1.5.0 — 2026-08-20

### Scripts for the things that were being typed out

`package.json` carried nine scripts and the CLI had grown well past them:
`--recalc`, `--print-spec`, `--write-meta`, `--write-expect` and `--bless` were
all being typed as `node dist/cli.js ...` from memory. They now have names, and
so do `clean:dry` and a new `links`.

- **`npm run check`** -- typecheck, doc links and the full suite. One gate to
  run before pushing, rather than three commands to remember in order.
- **`npm run links`** -- a real script at `scripts/check-links.mjs` instead of
  the throwaway one that got rewritten by hand three times in a day. It walks
  every `](...)` across the README, the changelog and `docs/`, resolves relative
  paths and checks that each anchor exists in the file it points at. Links and
  headings inside fenced blocks are skipped: a fence showing sample `report.md`
  output is not document structure, and counting its headings as anchors would
  let a genuinely broken link pass.
- **Everything that runs the CLI rebuilds first**, through `pre` hooks. `dist/`
  is generated and not committed, so a stale one answers with yesterday's code
  and says nothing about it -- the same silent-staleness hazard `clean` exists
  for, one level up.

### A person's name is out of the test fixtures

Two fixtures carried a real name lifted from a report -- a creator field, used
as filler. The repository is private, so nothing was exposed, but there was
never a reason to carry it: the tests assert on the *label*, never the value.
Replaced with a placeholder.

## 1.4.1 — 2026-08-20

### `--recalc` could not find the file it was given

Shipped broken. The workbook path was passed as `powershell -Command <script>
-args <path>`, and `-Command` does not populate `$args` -- only `-File` does --
so the script opened an empty path and every run failed with Excel's own
"Sorry, we couldn't find ." The flag had been tested only down its refusal
paths, which all still worked, so the failure looked like the safeguard doing
its job.

The path now travels in an environment variable. That also removes a quoting
problem that was waiting to happen: a Windows path threaded through a shell into
a script that is itself a command-line argument has a dozen ways to be subtly
wrong, and folder names here contain spaces, brackets and ampersands.

`$wb.Save()` replaces `$wb.SaveAs($dst, 51)` while I was in there -- the copy is
already at the destination and already xlsx, so re-specifying the format was one
more thing to get wrong.

Verified end to end on real reports: a golden with **94 formulas and 0 stored
results** comes back with 41, the originals are untouched, and on one quarterly
case the run went from **2 value differences to 5** -- the extra three being
exactly what the "Will recalculate" sheet had predicted. That sheet correctly
disappears from a recalculated run, having nothing left to predict.

## 1.4.0 — 2026-08-20

### `--recalc`: let Excel work the formulas out first

The reports carry formulas and no stored results, so a formula compares as its
text alone -- identical text on both sides matches, and the number a reader sees
never reaches the comparison. `--recalc` copies both files, has Excel open and
save them, and compares the copies. Every formula then carries a result and
compares as an ordinary value.

Nothing here evaluates a formula, and nothing here should: a spreadsheet engine
is an enormous thing to own and be wrong about, and Excel already has one.

- **The inputs are never touched.** `golden/` and `current/` are the record of
  what the generator produced. The copies live in `recalculated/` beside the
  results and are cleared with them.
- **`report.md` says so**, above the numbers whose meaning it changes.
- **It refuses rather than degrades** -- wrong platform, no Excel, or Excel
  already open, and the run fails. Comparing the files as they arrived would
  find far less than the flag promised, and every silence would read as a pass.
- **It stops when Excel is open.** `New-Object -ComObject Excel.Application`
  can attach to a session someone is working in, and the script ends by quitting
  what it got hold of with alerts suppressed. That would close their workbooks
  and discard unsaved work. The cost of refusing is a message; the cost of
  attaching is somebody's morning.

This also closes a hole in the impact graph that no amount of static analysis
can. References are found by reading a formula's text, so `OFFSET($A$1,
MATCH(...), 0)` is recorded as reading `$A$1` while it actually reads wherever
the arithmetic lands -- and **62,406 of the 379,959 formulas in the tree are
OFFSET**. Changes reaching a cell that way were never reported and could not be.
Excel does not have to reason about it. The limit is now written down in
`impact.ts`, which had claimed only to over-approximate.

### `--results <name>`

Names the folder a case writes into, so a plain run and a `--recalc` run can sit
side by side instead of overwriting each other:

    npx sheet-verify                                    # -> results/
    npx sheet-verify --recalc --results results-recalc  # -> results-recalc/

Deliberately not one command producing both. Two folders mean two verdicts and
two exit codes, and when they disagree the tool has stopped answering the
question it exists to answer. `clean` takes the same flag, and the folder walker
skips both names, so an earlier run is never read as a case of its own.


### `differences.xlsx` says which cells will recalculate, and what drives them

A colleague opened the two reports, found a sheet of totals plainly different,
looked for them in `differences.xlsx`, found nothing, and concluded the tool had
missed them. It had not. These reports carry formulas with no stored results --
across a tree of 33 real cases, **379,959 formula cells and not one with a
cached value** -- so a formula whose inputs moved has nothing to compare. What
moved was an input on another sheet, and that was reported.

`report.md` has said so under "Will recalculate differently" since the sweep
existed. Nobody reads it, because the artefact that gets opened is the
spreadsheet. So it goes there too, as a second worksheet:

```
Sheet      Cell  Column          How                      Driven by  Golden      Actual
RDS Loss   K15   Ground-Up (#3)  reads it directly        E15        1536804.44  11536804.44
RDS Loss   E28   Ground-Up       reads it directly        E15        1536804.44  11536804.44
RDS Loss   K28   Ground-Up (#3)  through another formula  E28
```

Each row names both ends -- the cell that will move, and the change that moves
it, with that cell's two values spelled out. Chains are followed, so the third
row above points at the second and the second points at the difference.

It does not say what the recalculated number will *be*. Nothing here evaluates
formulas, so there is no such number; the sheet answers which cells and driven
by what. Opening both files in Excel and saving them before a run bakes the
results in, and then they compare as ordinary values.

Two fixes fell out of building it. `AffectedCell.indirect` was documented and
hardcoded `false`, so nothing had ever distinguished a cell that reads a
difference from one that reads a formula that reads a difference; it is now set
from the walk that already computed it, and 51 of 99 cells on one real case are
downstream rather than direct. And a run whose only finding was recalculation
wrote no `differences.xlsx` at all -- the file is now written when either kind
of finding exists, since "these twelve totals will move" is exactly what the
file was being opened for.

## 1.3.0 — 2026-08-19

### A block with no header row is read as data, not renamed by one

A key-value block -- a label in column A, its value in column B -- has no header
row. Detection had to pick one anyway, and that cost twice over: the row picked
stopped being data, and the value column took its name from a *value*.

Where that value is the report's own name or id, it differs between any two runs
by construction. The column then pairs with nothing in the other file, and every
row of the block arrives as one column removed and another added -- thirty
findings for a fifteen-row block, with the real change buried among them. The
rows above the chosen header were not compared at all: on a quarterly report,
`Report ID` and the report name sat outside every table, so an edit to either
was invisible to layer 1.

It went unseen here because a golden and its actual are usually the same report
with something planted in it, so the name matched. Two genuinely different runs
are what expose it.

Detection now tells the two shapes apart by asking what the painting marks out.
A header row is painted to stand out from its data -- bold white on navy across
the table, nothing beneath it. A key-value block is painted down its label
column on *every* row, so the paint describes a column, not a row, and no row
stands out. Where every row is painted alike, the block is read as data from its
first row down, with columns named `Column A`, `Column B` -- names no edit to
the report can change.

Blocks nobody painted are excluded from the rule deliberately. Every row of an
unstyled table is painted identically too -- not at all -- and concluding "no
header" there would throw away the column names of every plain grid.

`headerRow` accepts `0` to say this by hand, for a block detection reads wrong.
It is not a new concept: `headerRow` has always named the row *above* the data,
so `0` is simply the row above row 1 -- which is what a block at the top of a
sheet needs, and what no value could express before.

Measured over the same tree of 34 real cases: **1,092 tables before and after,
14,747 columns before and after, no case changing verdict** -- and **103 more
rows compared**, on 58 tables. Eleven of those rows carried a planted edit that
had never been compared.

Table ranges are more accurate as a side effect. The range said what rectangle a
table covered by starting at `headerRow`, which claimed a row that named nothing
when the header row was blank. It now starts at the header row only when that
row named the columns, correcting **163 ranges** across the tree.

## 1.2.0 — 2026-08-19

Detection finds blocks it used to walk past — tables printed side by side, and
tables under a bold section title. A run can assert what it expects to compare,
and a report type's configuration can be generated from the pairs. Table
numbering shifts where new blocks appear, so anything pinned by number wants
rechecking with `--print-spec`.

### A section title no longer hides the table underneath it

Detection picks a block's header row by looking for the row that names the most
columns, with formatting breaking the tie — these generators paint a header row
deliberately, and nothing else on the sheet is painted that way.

A bold section title sitting alone in column A beat that test. It was the only
painted row in its block, so it won the search; then the block was discarded for
naming a single column, and every row under the title went unseen by layer 1
with nothing said about it. On one report that lost an entire exchange-rate
block — four label/value pairs, one of which had changed.

A row naming fewer than two columns is now passed over during the search, which
is the same rule already applied to the row that wins it. Across a tree of 33
real cases this brought **55 previously invisible blocks into detection** and
44 more tables into layer 1, with no change to any table that was already found
and no case changing verdict.

Table numbers below a newly found block shift by one — `Table 2` means "the
second block detection finds", so anything pinned by number on an affected sheet
needs rechecking with `--print-spec`. None of the sheets configured in the tree
this was measured on were affected.

### The README is a starting point again, not the manual

It had grown to 1,181 lines, most of it reference that only matters once you
are already running. Split into a document per question, with the README down
to 321 lines: what it is, install, quick start, the five things worth doing
once, the one behaviour to know before trusting a result, how it works, and a
table saying where everything else lives.

| | |
| --- | --- |
| `docs/cli.md` | the command: flags, pair naming, case selection, labels, generating a config |
| `docs/configuration.md` | keys, sheets, tables, tolerance, metadata, the options reference |
| `docs/detection-tuning.md` | working out which correction a report needs |
| `docs/reading-a-report.md` | the case folder, each artefact, reading the verdict |
| `docs/api.md` | matchers, the direct API, invariants, extending it |

No prose was dropped -- every line of the original lands in exactly one of
these, checked mechanically -- and every internal link was repointed and
verified to resolve. `docs/` now ships with the package, so those links work
for a consumer too.

### Adding a table to a single-table sheet no longer replaces it

A sheet holding one table carries its header row and key at the sheet level,
with no `tables` block -- that is what makes reports read "Ledger" rather than
"Ledger · Table 1". Declaring a table on such a sheet *replaced* it: one entry
written to check a title block also stopped the sheet's real table being
compared, and nothing said so. The sheet's own table is now kept, filed under
the sheet's name, which is the name reports already give it:

```
| Ledger         | `A1:C4` | 3 columns | 3 rows | key |
| Ledger · Bands | `A5:C7` | 3 columns | 2 rows | key |
```

Only fires when a layer above declares `tables` and the one below has none, so
a sheet-level `keyColumns` written on its own -- the common correction, and the
one in every example -- behaves exactly as before.

### Overriding and adding tables: docs/detection-tuning.md

Two new sections, both written from behaviour that was checked rather than
assumed: which fields survive a partial override (all of them -- merging is per
table and per field), what a table name that matches nothing does (it adds a
phantom at `A1:B1` and fails the run, rather than doing nothing), and the three
things that decide where an added table lands -- `headerRow`, its bound, and
`columns` when it has a neighbour.

### `--write-meta`: a report type's meta.json, from the pairs

Starting a report type meant typing out a config by hand and finding out which
parts were needed from the failures. Now:

    sheet-verify --write-meta output_comparison/natural_cat_srq

It reads the pairs under that folder and writes a `meta.json` holding what it
can show evidence for, and comments for the rest:

- **`reportType`**, from the folder name, with a note that the spelling is
  yours to fix.
- **`metadata`** — run-identity labels found in the files, matched against a
  deliberately short vocabulary and kept only where the value was **observed to
  take more than one value** across the files scanned. A label found in a wide
  header row is passed over: "Program ID" as the 28th column heading of a table
  is not run identity, and listing it would drop that column out of the
  comparison. A report type that spells its own name into the label gets the
  wildcard form, `*Report Name`, which is what covers Summary / Quarterly /
  Regional in one entry. Labels found but unvarying are listed as candidates,
  not written.
- **`defaults`** — `requireCachedValues: false` when the files really do carry
  formulas with no stored result, `fillKeyDown: true` when a sheet really does
  write a group heading once and leave it blank beneath. Both quote the
  evidence.
- **the unkeyed tables**, largest first, as a note — the work that is left,
  visible without being pre-decided.

What it deliberately does **not** write is per-sheet header rows, end rows and
keys. Those are re-detected from the files on every run, which is what lets a
report change shape without breaking the config; frozen into a file they go
stale, and a generated entry is indistinguishable from one somebody meant.

It refuses to touch a `meta.json` that already has settings in it.

On a summary report it produced `Report ID`, `*Report Name`, `Creation Date`
and `Elapsed Processing Time` — the hand-written list plus one — and the tree
it wrote gave the same verdicts as the hand-written config.

### `expect`: a table that stops being compared now fails

Detection is remade from the files on every run. That is the point, and it
means coverage can shrink with nothing to show for it but a smaller number in
a summary line nobody was watching. On one report here an entire block went
unread by layer 1 for the life of the tool.

    { "expect": { "Report Info": ["A2:B17", "H2:J22", "A19:B24"], "Cover": 1 } }

Ranges, or a count. Checked after the comparison and reported as an integrity
error, which fails the run and names the sheet:

```
- Report Info: expected 5 table(s), compared 4 — not compared: A99:B120
- Comments: expected 3 table(s) compared by name and key, found 1
```

It is an assertion, never an instruction — it changes nothing about what is
compared, so an entry that is wrong stops the run and says so rather than
quietly comparing the wrong thing, which is what a pinned `headerRow` does when
a report shifts.

`sheet-verify --write-expect` records it from a run, into each case's
`case.json`, beside whatever is already there. Deliberately its own step rather
than part of `--bless`: blessing accepts a change to the *output*, and
accepting a change to what is *checked* is a separate decision.

### Every report says which rectangle each table covered

The question "what are we actually verifying" had no answer in the output. A
table's spec does not hold one either: with no `endRow` it runs to the bottom
of the sheet, and with no `columns` it spans the whole width. So the range is
now read off the model — what was *read*, not what was asked for — and reported
in three places.

`report.md` gains a **What was verified** section listing every table layer 1
read, with its range in each file, its width, its height and whether rows were
matched by key or by position. Each changed table also carries its own range
under the heading, because a finding at `B25` reads differently depending on
whether the table starts at row 2 or row 24:

```markdown
### Currency Info · Table 2

_`A18:B23` — 2 columns × 5 rows, rows matched by position_
```

`compared.xlsx` gains a banner on row 1 of every table's worksheet saying the
same thing — bold, on a filled band the width of the table, over a rule. The
column headers move to row 2 and the frozen pane and autofilter move with them.
The band is light where the header under it is dark, so the two read as two
things rather than one heavy block. `diff.json` gains `range: { base, next }`
on every compared sheet outcome.

The two files are named separately whenever they disagree. A table that grew,
moved or lost its bottom rows between runs explains a wall of differences that
would otherwise read as edits.

### Tables printed side by side are found, and can be bounded by column

A sheet was split into tables by blank rows alone. That cannot separate a
definitions table in `H:J` from the key-value block in `A:B` beside it: both
start on row 1, and the taller one keeps every row of the block non-blank. Read
as one table the two header rows fuse, the shorter table's rows read as blank,
and a key named in one of them can be found in the other. One report had four
tables on such a sheet and two were detected.

Detection now cuts a sheet by rows and by columns in turn, until neither moves —
cutting by rows can expose a column gap that spanned the old region, and the
other way round. A cut needs **two or more** blank columns: one is a spacer, and
cutting on one fragments 186 blocks across a tree of 33 real cases, most of them
wrongly. Two cuts 67, and those are genuinely two tables.

New `columns` option on a sheet or table, written as a range of letters:

```json
{ "tables": { "Definitions": { "columns": "H:J", "keyColumns": ["Return Year"] } } }
```

Detection sets it when a sheet holds tables side by side, and only then — a
stacked sheet gets no bound, since freezing a width that grows would help
nobody. It is what stops a header row reaching across into a neighbour: without
it the reader takes the whole width, finds the neighbour's columns hold data,
and folds them into this table.

Measured over the same tree: **37 fewer tables matched by row position**, and
118 cells moved from layer 1 to layer 2 — caption bands like `Label A / Label B
/ % Change` that were being read as one-row tables and are decoration, not data.
Both layers still cover every cell; no case changed verdict. Table numbers shift
on any sheet where this applies, so anything pinned by number wants rechecking
with `--print-spec`.

### Tuning detection: docs/detection-tuning.md

How to read a run for the tables detection got wrong, recipes for the two shapes
it reliably struggles with — a key-value block with no header row, and a data
table under title blocks — and which parts of the configuration are positional
rather than named, so it is clear which entries survive a report changing shape
and which do not.

## 1.1.0 — 2026-08-18

Comparison gained a tolerance — `0.001` by default, applied in both layers —
with the cells it forgives counted and listed rather than quietly dropped.

### Tolerance defaults to 0.001

Applied when no config names one, to both layers, and overridden by `tolerance`
anywhere in the tree — `0` for exactness. It absorbs recalculation noise, which
lands about `1.19e-7` from the stored figure on a report in the hundreds of
millions, and stays well under a cent so a change made on purpose still shows.

`0.01` was the first proposal and was measurably too wide: it swallowed a real
`34.45 → 34.46`, and the suite's own "a difference that matters is still caught,
however small" failed against it.

Four existing tests now ask for `tolerance: 0` explicitly. They were written
when exactness was the default, and each is about something else — how a wall of
differences is presented, whether the ledger and the comparer agree — so they
say what they need rather than relying on it.

### What differs gets its own table

```
**Cells that differ**

| total | within tolerance (±0.001) | above tolerance |
| ---: | ---: | ---: |
| 894 | 870 | **24 (2.7%)** |
```

894 reads as a disaster when 870 of those moved by less than a thousandth. It is
the figure the whole report exists to deliver, so it no longer sits as one line
among the file paths, and the tolerance that drew the line is named on the
column that used it. Where a run applies several — tolerance resolves per column
— the header gives the range rather than claiming a single number.

### `differences.xlsx` leaves out within-tolerance rows

It is read as a list of things to fix, and a cell the tolerance already forgave
is not one. `--ledger all` still carries them, and `report.md` lists every one
under "Inside the tolerance you set", so the rule stays auditable.

### Tolerance reaches layer 2

A tolerance used to quiet only the keyed comparison. The address sweep knew
nothing about it, so `cells differing` — the number at the top of every report,
and the one most people read first — went on counting gaps of `1.19e-7` that
the config had already declared immaterial, and the same cells filled the
"nothing checked them" list.

The sweep now reads the per-column tolerances off the tables layer 1 compared,
translating them from column names to the addresses it works in, and takes a
blanket `*` from `defaults` for the cells no compared table covers. Both layers
therefore agree about which gaps matter.

Cells inside tolerance are counted and listed on their own — `within tolerance`
beside the differing count, and a folded section showing each cell with the size
of its gap. Hiding them would make a tolerance set too wide invisible, which is
a worse failure than the noise it was set to remove.

Layer 2 compares the text a cell displays and never sees its type, so text that
reads as a number is treated as one. Documented rather than worked around: the
answer is to set a tolerance to the size of the rounding it absorbs.

## 1.0.0 — 2026-08-18

First release, covering everything through `ce721d0`. The changelog starts
here, so the entries below are the last day of that work and everything
earlier is in the git history.

The comparison tree moved to `output_comparison/`, cases can be named and
selected, and both layers stopped judging things that differ by construction.

### Report metadata is read, listed, and not judged

`metadata` in any `meta.json` names the cells that identify a run rather than
describe it — a report name, an id, a creation timestamp. They are read, listed
in `report.md` under **Not verified, on purpose** with both values, and left out
of the verdict.

The test is narrow: **does the value differ between two runs by construction?**
A minted id does. A creator name does not, when the same account generates every
report — a change there means the wrong account ran it, which is a finding. The
same goes for anything the figures depend on: view of risk, currency, model
version, the as-at date behind a "Data as of …" caption.

An entry is a label (`"Report ID"`, matching the cell and the value beside it,
including a fused `="Report ID: " & id`), or a cell reference (`"Cover!A3"`).
`*` covers a run of text, which is how one pattern reaches `Summary Report Name`, `Regional Report Name` and `Quarterly Report Name`. Either form takes a
sheet qualifier.

Exclusion reaches all three consumers — the keyed comparison, the address
sweep, and `differences.xlsx`. Each had to be told separately; a metadata cell
was briefly reported by one while another called it fine.

### `ignoreSheets` is gone from the shipped config

Cover, Contact, Table of Contents, Executive Summary, Glossary, Disclaimer,
Comments and Errata were excluded as boilerplate — 2,582 cells nobody checked.
Put back in scope they compare clean everywhere except three run-stamped cells
on Cover, now named as metadata. Images and drawings are invisible to both
layers on every sheet, so a tab full of logos costs nothing to verify. The
option still exists; nothing ships using it.

### Every difference is reported, including cascades

Numeric comparison briefly carried slack of scale × 1e-12, on the reasoning that
Excel keeps 15 significant digits so a smaller gap is rounding rather than
change. True, and the wrong trade: a cell whose stored values differ is a cell
that differs. Judging which differences matter is the reader's job, and
`tolerance` is how they say so, per column, deliberately.

Cascades are likewise reported in full. The root-cause mark stays — it says
where a problem starts — but a report that shows only causes understates the
reach of a change.

### `report.md` restructured

- a two-layer assurance block, replacing the old "cells unchecked" count
- differences grouped by column when a table has more than twelve of them
- one block per sheet, with its own count, in the recalculation and
  unchecked-cells sections
- long sections folded behind `Show`
- values printed at full stored precision, so the two sides of a difference can
  never print as the same string

### Cases can say what they are

`reportType` in a report type's `meta.json` and `label` in a case's `case.json`
now head the case in the run log and title its `report.md`:

```
✗ Validation Report · case_003 · an extra Cat Model Version row shifts the perils block down
    validation_report/case_003
    1 sheet failing
```

`reportType` inherits. **`label` does not** — one written a folder above would
describe every case beneath it identically. A label that only repeats the folder
name is treated as no label, so nothing reads `case_003 · case_003`.

### A pair can be named by folder

```
case_001/
  golden/case_1%1786955263151.xlsx
  current/case_1%1786957329031.xlsx
  results/
```

Which is what a downloader produces when it keeps the source system's name —
that name carries the download's timestamp, and renaming it to `golden.xlsx`
throws it away. Same words as the file prefixes (`golden|baseline|expected|
before`, `current|actual|new|after|report`). Exactly one spreadsheet per folder;
two stops the run rather than one being guessed at.

Re-blessing such a pair writes the new golden under the name the new report
arrived with and removes the file it replaced.

### `cases` selects what runs

```json
{ "cases": ["comparison_report/**", "!comparison_report/case_002"] }
```

In any `meta.json`, relative to that file's own folder, so the root can select
by report type and a type can select its own cases. What is left out is counted
with the results — a case set aside is not a case that passed.

### A folder meant to be a case is never skipped in silence

A malformed case folder used to be reported only when the run found no cases at
all. Beside working cases it vanished, and the total simply read one lower.
Those folders are now always reported, counted, and fail the run:

```
1 folder(s) meant to be cases could not be run:
  validation_report/case_002: a golden/ folder is here with no current/ folder beside it

5 cases, 5 failing, 1 could not be run
```

Folders that were never cases — groupings, stray directories — stay quiet unless
nothing runs at all.

### Fixes

- **Merged cells.** ExcelJS reports the master's value for every slave in a
  merge, so a merged banner was counted once per column it spanned. Slaves now
  read as blank, in the reader as well as the sweep.
- **Trailing blank rows.** Excel's used range outlives its contents: a sheet
  edited down to thirteen rows still reports thirty. Under positional matching
  every phantom row took an ordinal, so a disclaimer identical in both files
  arrived as seventeen removed rows. The trailing run is now trimmed; a blank
  row *between* populated ones is kept, since dropping it would shift everything
  under it out of step.
- **One equality rule.** `differences.xlsx` had its own, stricter than the
  comparison's, and reported cells the report said were fine. Both now call
  `equalValues`.
- **Full precision.** Report values were rounded to twelve significant digits,
  which printed the two sides of a smaller difference as the same string.

### Housekeeping

- the tree is `output_comparison/`, gitignored like the other data folders
- the per-case output folder is `results/`, renamed from `result/`
- `npm run clean` removes a `results/` beside a golden file **or** a `golden/`
  folder
- `npm run compare` defaults to `output_comparison`

### Known gaps

- **Row numbers in config are literal.** `validation_report`'s `Perils` table is
  pinned to `headerRow: 23`, which is right for two of its three cases; the
  third carries an extra `Cat Model Version` row that shifts the block to 24, so
  the key column is not found there and its differences fall through to the
  address sweep. Anchoring a header by content rather than row number would fix
  the class of problem.
- **A misspelled setting inside `sheets` is accepted and never read.** The
  unknown-key check covers the top level of a config file only.
- **`endRow` is ignored for CSV.**
- **Root cause is decided from same-row references only.** A full dependency
  graph exists in `impact.ts` and is not wired into that decision.
