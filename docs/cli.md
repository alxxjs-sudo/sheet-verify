# The command

Everything the CLI does, in the order you meet it. The [README](../README.md)
covers getting started; this is the reference.

```bash
npx sheet-verify                      # every case in ./output_comparison
npx sheet-verify path/to/case_007     # just one case
npx sheet-verify --bless              # accept the differences as the new golden
npx sheet-verify --print-spec         # show what was detected, as JSON
npx sheet-verify --write-meta <type>  # write a starting meta.json for a report type
npx sheet-verify --write-expect       # record what each case verified, as a guard
npx sheet-verify --ledger all         # record matching cells too
npx sheet-verify --help
```

**A case is any folder holding a golden file and the report to compare**, at any
depth. Everything above a case is just grouping, so file them by kind:

```
output_comparison/
  meta.json                          applies to every case below
  global_standard_cat_report/
    meta.json                        applies to this report type
    case_001/  case_002/  …
  quarterly_report/
    meta.json
    case_001/
      case.json                      only this case
  analyses/
    marginal/
      case_001/  case_002/  …
```

**Configuration is inherited down the tree.** Every `meta.json` from the root
down applies in order, and the case's own `case.json` wins. So the settings a
whole report type shares are written once, and only a case that genuinely
differs — an extra sheet or two — needs a file of its own.

Cases are named by their path, since `case_001` will exist under every type:

```
✗ Global Standard Cat Report · case_002 · three columns inserted into Geocoding
    global_standard_cat_report/case_002
    1 sheet failing
```

Targeting a subfolder runs only what is under it, but still applies the
configuration above it — so `sheet-verify quarterly_report` gives the same verdicts for
those cases as a full run.

## Two ways to name a pair

**By file name.** Matched by prefix, so `golden.xlsx` / `actual.xlsx` is the
convention but not the only option:

| role | any of |
| --- | --- |
| the trusted output | `golden` · `baseline` · `expected` · `before` |
| the output under test | `actual` · `new` · `current` · `after` · `report` |

The prefix must end at a non-word character, so `golden-2026.xlsx` and
`golden 2026.xlsx` are recognised and `golden_2026.xlsx` is not — an underscore
is a word character.

**By folder.** Put the role in the folder and each file keeps whatever name it
arrived with, which is what a downloader produces when it preserves the source
system's name — that name carries the download's timestamp, and renaming it to
`golden.xlsx` throws that away:

```
case_001/
  golden/case_1%1786955263151.xlsx      one spreadsheet, any name
  current/case_1%1786957329031.xlsx     the folder says which side it is
  results/
```

Folder names are the same words as the file prefixes. **Exactly one spreadsheet
per folder** — two stops the run rather than one being picked, because with
names like those a guess is a coin toss:

```
comparison_report/case_001: current/ holds 2 spreadsheets [rep_178699.csv, rep_178657.csv] — it must hold exactly one
```

A half-built case — `golden/` written and `current/` not yet — is reported and
fails the run, rather than quietly dropping out of the count:

```
1 folder(s) meant to be cases could not be run:
  validation_report/case_002: a golden/ folder is here with no current/ folder beside it

5 cases, 5 failing, 1 could not be run
```

`.xlsx`, `.xlsm` and `.csv` all work, and both files must be the same kind.

## Choosing which cases run

A run can be narrowed from the command line by targeting a folder, or written
down in any `meta.json` so a full run does what the tree says:

```json
{ "cases": ["comparison_report/**", "!comparison_report/case_002"] }
```

Paths are relative to the file that carries them, so the root can select by
report type while a type selects its own cases. `*` stays within one path
segment and `**` crosses them; naming a folder takes everything inside it; a
leading `!` excludes. Every file carrying a list narrows further — a case has to
be selected by all of them.

What is left out is counted with the results, never dropped in silence:

```
2 cases, 2 failing — 4 not selected by "cases"
```

## Naming what a case is

Two keys change nothing about the comparison and a great deal about reading it.
`reportType` goes in a report type's `meta.json`; `label` goes in a case's
`case.json` and is one sentence saying what that case is for:

```json
{ "label": "an extra Cat Model Version row shifts the whole perils block down" }
```

```
✗ Validation Report · case_003 · an extra Cat Model Version row shifts the whole perils block down
    validation_report/case_003
    1 sheet failing
```

The label also titles that case's `report.md`, with the folder name kept
underneath so the file is still findable from what it says:

```markdown
# an extra Cat Model Version row shifts the whole perils block down

_Validation Report · case_003_
```

`reportType` is inherited like any other setting. **`label` deliberately is
not** — one written a folder above would head every case beneath it with the
same sentence and distinguish none of them. A label that merely repeats the
folder name is treated as no label at all, so nothing reads `case_003 ·
case_003` before you have written a real one.

## Correcting anything else

`--print-spec` shows exactly what was worked out from your files:

```bash
npx sheet-verify --print-spec > output_comparison/comparison_report/case_001/case.json
```

Edit that file and it is layered over the detection on the next run. Everything
in it is optional — you only write the parts you want to change, and the rest of
the detection is kept. The most common entry by far silences a timestamp that is
rewritten on every run:

```json
{
  "defaults": { "ignoreRows": ["Generated At"], "ignoreColumns": ["Run Id"] }
}
```

**Put it at the right level.** The same file format works as `meta.json` at any
folder above a case, so a setting that applies to a whole report type belongs
there rather than copied into every case:

| goes in | applies to |
| --- | --- |
| `output_comparison/meta.json` | every case — timestamps, run ids |
| `reports/<type>/meta.json` | that report type — its keys, its decoration sheets |
| `<case>/case.json` | one case — an extra sheet it happens to have |

Lists accumulate down the layers rather than replacing each other, so a type
excluding its glossary and a case excluding one more sheet ends up excluding
both. `--print-spec` lists every layer that was applied, so it is always
answerable which rules a result came from.

## Starting a report type without typing a config

```bash
npx sheet-verify --write-meta output_comparison/natural_cat_srq
```

Reads the pairs under that folder and writes the `meta.json` you would
otherwise hand-write: the report type's name, the run-identity cells it can
show evidence for, the defaults the files themselves justify, and the unkeyed
tables listed as a note. Everything in it is either evidence or a comment.

It does **not** write per-sheet header rows, end rows or keys. Those are
re-detected from every file on every run, which is what lets a report change
shape without breaking your config -- frozen into a file they go stale, and a
generated entry is indistinguishable from one you meant. Add those yourself,
only where a run says detection got it wrong.

It refuses to touch a `meta.json` that already has settings in it.

## Guarding what gets verified

Detection is remade on every run, so a table can stop being compared and the
only sign is a smaller number in a summary line. `expect` turns that into a
failure with a name:

```json
{ "expect": { "Report Info": ["A2:B17", "H2:J22", "A19:B24"], "Cover": 1 } }
```

Ranges, or a count of tables. Checked after the comparison; a miss is an
integrity error, so the run fails and says which sheet and which range:

```
- Report Info: expected 5 table(s), compared 4 - not compared: A99:B120
```

`--write-expect` records it from a run into each case's `case.json`, beside
whatever is already there. It is deliberately separate from `--bless`:
blessing accepts a change to the output; accepting a change to what is
*checked* is a different decision.

It is an assertion, never an instruction -- it changes nothing about what is
compared. An entry that is wrong stops the run rather than quietly comparing
the wrong thing, which is what a pinned `headerRow` does when a report shifts.

**[docs/detection-tuning.md](detection-tuning.md)** goes further: how to
read a run for the tables detection got wrong, recipes for the two shapes it
reliably struggles with — a key-value block with no header row, and a data table
under title blocks — and the two things about this configuration that are
positional rather than named, so you know which of your entries will need
revisiting and which will not.

## Recalculating before comparing

These generators write formulas and no results — Excel works them out when you
open the file. Across a tree of 33 real cases that is **379,959 formula cells
and not one carrying a stored value**, so by default every one of them compares
as formula text alone: identical text on both sides matches, and a total that
moved by billions reads as identical.

```bash
npx sheet-verify --recalc
```

Both files are copied, opened and saved by Excel, and the copies are compared.
Every formula then carries a result and compares as an ordinary value.

- **The inputs are never touched.** `golden/` and `current/` stay exactly as the
  generator wrote them — they are the record of what it produced, and a run must
  not rewrite its own evidence. The copies go in `recalculated/` beside the
  results and are cleared with them.
- **`report.md` says so**, at the top, above every number it changes the meaning
  of.
- **It refuses rather than degrades.** No Windows, no Excel, or Excel already
  open, and the run fails with a message. Silently comparing the files as they
  arrived would find far less than the flag promised, and every one of those
  silences would read as a pass.
- **Excel must be closed.** Automation attaches to a running session and ends by
  quitting it with alerts suppressed, which would discard anything unsaved. The
  run checks, and stops.

It is slower than a plain run by however long Excel takes to open and save each
file, which on a large report is seconds rather than milliseconds.

### Keeping both runs

`--results` names the folder a case writes into, so the two views can sit side
by side rather than overwrite each other:

```bash
npx sheet-verify                                      # -> results/
npx sheet-verify --recalc --results results-recalc    # -> results-recalc/
```

There is deliberately no way to produce both from one command. Two folders mean
two verdicts and two exit codes, and when they disagree the tool has stopped
answering the question it exists to answer. Choosing at the command line keeps
one run, one answer.

`npm run clean -- --results=results-recalc` clears the named folder; without the
flag it clears `results/`.

### Which one is the real answer

The recalculated run, when you can have it. It **contains** the plain one —
formula text is still compared, so nothing the plain run finds is lost — and it
adds the values, which is most of the report. The plain run is what you fall
back to when Excel is not available, and its "Will recalculate" sheet is the
substitute: which cells would move, and driven by what, without the numbers.

