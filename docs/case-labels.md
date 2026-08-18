# Case labels

The `label` each case carried in `output_comparison/` before the report pairs
were cleared, with the notes that recorded how a few of them were built.

Kept here because that tree is gitignored: nothing inside it survives a
clean-out or a move to another machine, and these sentences were written by
reading each case’s own diff rather than guessed at.

**Reapply only where a case still holds the same pair.** A label names what one
specific comparison was for. Pasted onto a re-download that happens to reuse the
case number, it would describe the wrong report confidently, which is worse than
having no label at all.

To use one, put it in that case’s `case.json`:

```json
{ "label": "the Regions sheet is no longer produced" }
```

## comparison_report

- **case_001** — geography rebuilt: 851 cells restated across the breakdown
- **case_002** — one line-of-business figure moves, nothing else
- **case_003** — the same report downloaded twice, expected to match

## facility_report

- **case_001** — an unchanged facility report, downloaded twice
  - Named in the source system as "FacilityReport_0721".
- **case_002** — a single summary figure differs on an otherwise clean report
  - Named in the source system as "FacilityReport_NonCat_0626".
- **case_003** — one figure moves on the RiskPlay sheet
  - Named in the source system as "FacilityReport_NonCat_0629".

## global_standard_cat_report

- **case_001** — geography rebuilt: 848 cells restated
- **case_002** — three unnamed columns inserted into Geocoding, formulas rewritten
- **case_003** — a stray "waaaa" column added to Modeling Parameters

## natural_cat_srq

- **case_001** — two genuine downloads, one AAL figure apart
- **case_002** — the PML Exhibit sheet stops being produced
  - Named in the source system as "case_002 - PML Exhibit removed from the report". A test case, not a downloaded pair. Copied from case_001, then the 'PML Exhibit SCS TS' worksheet was removed from actual.xlsx at the package level so nothing else about the file changed. The golden still has it, so the run should fail with a removed sheet.
- **case_003** — planted: assumptions edited and a data-quality total scaled
  - Named in the source system as "case_003 - edits planted in the downloaded pair". Came from test_data/natural_cat_srq/case_001. The golden is the report as downloaded; the actual is a copy with known differences planted in it.
- **case_004** — planted: assumptions edited and the HU RMS region sheet removed
  - Named in the source system as "case_004 - from s.xlsx, edits planted". Came from test_data/natural_cat_srq/case_002, whose report file was named s.xlsx. The golden is that report; the actual is a copy with known differences planted in it.
- **case_005** — planted: assumptions edited, with a total scaled by 1.01
  - Named in the source system as "case_005 - from case_003.xlsx, edits planted". Came from test_data/natural_cat_srq/case_003. The golden is that report; the actual is a copy with known differences planted in it.

## pro-forma

- **case_001** — planted: a New Analysis sheet appears and a currency figure moves
  - Named in the source system as "CurrencyTest_0916_4.6.1 (Copy) (Copy) (Copy) (Copy)_0810".
- **case_002** — planted: currency figures edited and a summary formula scaled
  - Named in the source system as "Non-cat report_0807".
- **case_003** — planted: the loss-by-account VaR sheet is dropped
  - Named in the source system as "ProformaReport_886Contracts (Copy)_5.0.2 (Copy)_Latest_0807".
- **case_004** — planted: currency edits with a layer total scaled by 1.01
  - Named in the source system as "TestToConfirmLayerTypeInReport".

## riskplay_report

- **case_001** — planted: a New Analysis sheet on top of report-info edits
  - Named in the source system as "RiskplayReportWithOverrides0616".
- **case_002** — planted: report-info edits and a portfolio ratio scaled
  - Named in the source system as "RiskPlayReport_0629".
- **case_003** — planted: the Portfolio Composition sheet is dropped
  - Named in the source system as "RiskPlayReport_0709".

## roleplay_report

- **case_001** — planted: report-info figures edited, nothing structural
  - Named in the source system as "RoleplayReportToTestIndividualPremium".
- **case_002** — planted: a New Analysis sheet alongside report-info edits
  - Named in the source system as "RoleplayReport_0721_CustomVORs".
- **case_003** — planted: larger report-info figures on a 300-program report
  - Named in the source system as "RoleplayReport_300programs".

## validation_report

- **case_001** — planted: Project Uploader renamed, RMS Historical dropped, perils restated
  - Named in the source system as "REGRESSION_v5.0.1_Sample1b_RMS-validation-RMS_v18.1_ROLePlay".
- **case_002** — planted: Project Uploader renamed and both peril AALs restated
  - Named in the source system as "REGRESSION_v5.0_Edison Collection - Sample 1b (Simple, with NetPreCat FIX)-validation-RMS_v18.1_ROLePlay".
- **case_003** — an extra Cat Model Version row shifts the whole perils block down
  - Named in the source system as "Testing MPIUA to see Occ vs Agg-validation-RMS_v23_ROLePlay".

## workers_comp_comparison_report

- **case_001** — a full geography rebuild, with planted edits on top
- **case_002** — planted: RDS Loss dropped and portfolio totals restated
- **case_003** — planted: portfolio totals restated, "data as of" caption scaled

## workers_comp_report

- **case_001** — planted: a New Analysis sheet over a full geography rebuild
- **case_002** — planted: portfolio totals restated with the caption formula scaled
- **case_003** — planted: RDS Loss dropped and portfolio totals restated
