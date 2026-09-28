# Data Cleaning Procedure, Project 4

This document mirrors the eleven cleaning sheets (C1 to C11) in the workbook. Every figure below was produced by the same Python engine that wrote those sheets, and each sheet repeats the key counts as live Excel formulas so the reconciliation can be checked inside the workbook.

## Guiding principles

1. No row is deleted unless it is completely empty or an exact copy of an earlier row.
2. No value is changed silently: every changed cell is traceable through a flag or audit column (key_repaired_flag, outlier_fields, imputed_fields, and the cross field flags).
3. Where two fields contradict each other, one is declared authoritative, the decision is written down, and the other is kept for audit.
4. Dates, identifiers, survey answers and legacy model outputs are never imputed.

## Step summary

<table>
  <tr><th>Step</th><th>Sheet</th><th>Action</th><th>Rows in</th><th>Rows out</th><th>Cells changed</th></tr>
  <tr><td>S1</td><td>C2 Structural Integrity</td><td>Removed fully blank rows</td><td>15,394</td><td>15,372</td><td>0</td></tr>
  <tr><td>S2</td><td>C2 Structural Integrity</td><td>Removed exact duplicate records</td><td>15,372</td><td>15,250</td><td>0</td></tr>
  <tr><td>S3</td><td>C3 Key Collision Repair</td><td>Repaired duplicate encounter_id values</td><td>15,250</td><td>15,250</td><td>46</td></tr>
  <tr><td>S4</td><td>C4 Text Standardization</td><td>Collapsed categorical variants to one standard label</td><td>15,250</td><td>15,250</td><td>3,554</td></tr>
  <tr><td>S5</td><td>C5 Boolean Normalization</td><td>Converted 12 token encodings to TRUE/FALSE</td><td>15,250</td><td>15,250</td><td>152,500</td></tr>
  <tr><td>S6</td><td>C6 Date Standardization</td><td>Parsed four mixed date formats to true Excel dates</td><td>15,250</td><td>15,250</td><td>29,936</td></tr>
  <tr><td>S7</td><td>C7 Numeric Validation</td><td>Nullified impossible and extreme values</td><td>15,250</td><td>15,250</td><td>452</td></tr>
  <tr><td>S8</td><td>C8 Cross Field Consistency</td><td>Tested business rules between related fields</td><td>15,250</td><td>15,250</td><td>0</td></tr>
  <tr><td>S9</td><td>C9 Missing Value Treatment</td><td>Filled or explicitly retained every missing value</td><td>15,250</td><td>15,250</td><td>6,065</td></tr>
  <tr><td>S10</td><td>C10 Feature Engineering</td><td>Derived arrival calendar keys, season, acuity and SLA fields</td><td>15,250</td><td>15,250</td><td>137,250</td></tr>
  <tr><td>S11</td><td>Clean Data</td><td>Published analysis ready table</td><td>15,250</td><td>15,250</td><td>0</td></tr>
</table>

## C1 Data profile

The raw export holds 15,394 rows and 45 columns. Each column was measured for blanks, distinct values and distinct values after normalising case, whitespace, trailing periods and underscores. A gap between those two distinct counts proves the column carries formatting variants.

## C2 Structural integrity

22 rows were completely blank and were removed. 122 rows were exact copies of an earlier row and were removed, keeping the first occurrence. The original Excel row number of every removed row is listed on the sheet.

## C3 Key collision repair

After deduplication, 46 values of encounter_id were still shared by different records. The collision free anchor staffing_scenario_id carries the true record sequence, so 46 records were re keyed from it instead of being deleted. The live check on the sheet confirms that no key in Clean Data is duplicated.

Related identifiers that repeat for legitimate reasons are flagged but not changed:

* patient_id: 92 records flagged. Flagged in patient_id_conflict_flag. Repeat visits by one patient are normal in an ED, but these repeats carry different age and sex, so they are routed for registration review rather than overwritten.

## C4 Text standardisation

15 categorical columns were standardised. 174 non standard spellings covering 3,554 cells were mapped to one governed label. Each raw value was normalised (trim, drop trailing period, underscores to spaces, collapse spaces, ignore case) and mapped to the most frequent clean spelling. Identifier columns were trimmed and upper cased.

<table>
  <tr><th>Issue type</th><th>Variants</th><th>Cells</th></tr>
  <tr><td>Inconsistent letter case</td><td>80</td><td>1,981</td></tr>
  <tr><td>Leading or trailing whitespace</td><td>41</td><td>805</td></tr>
  <tr><td>Trailing period</td><td>41</td><td>631</td></tr>
  <tr><td>Underscore used as separator</td><td>12</td><td>137</td></tr>
</table>

## C5 Boolean normalisation

10 flag columns used twelve encodings of yes and no (true, TRUE, yes, Yes, Y, 1 and their negatives). They were converted to native TRUE and FALSE. No unrecognised tokens were found.

## C6 Date standardisation

Dates arrived as text in four formats. Unambiguous patterns were parsed directly; slash dates where both parts are 12 or below were resolved as month first (the convention shown in the data dictionary) and flagged in an _ambiguous_flag column.

<table>
  <tr><th>Column</th><th>Detected pattern</th><th>Records</th><th>Parsed</th></tr>
  <tr><td>admission_date</td><td>Slash ISO (YYYY/MM/DD)</td><td>3,096</td><td>3,096</td></tr>
  <tr><td>admission_date</td><td>Day Month abbrev (DD Mon YYYY)</td><td>2,943</td><td>2,943</td></tr>
  <tr><td>admission_date</td><td>ISO 8601 (YYYY MM DD)</td><td>2,939</td><td>2,939</td></tr>
  <tr><td>admission_date</td><td>Ambiguous NN/NN/YYYY resolved as MM/DD/YYYY</td><td>2,419</td><td>2,419</td></tr>
  <tr><td>admission_date</td><td>MM/DD/YYYY (day > 12, unambiguous)</td><td>1,844</td><td>1,844</td></tr>
  <tr><td>admission_date</td><td>DD/MM/YYYY (day > 12, unambiguous)</td><td>1,826</td><td>1,826</td></tr>
  <tr><td>admission_date</td><td>Missing</td><td>183</td><td>0</td></tr>
  <tr><td>discharge_date</td><td>Day Month abbrev (DD Mon YYYY)</td><td>2,991</td><td>2,991</td></tr>
  <tr><td>discharge_date</td><td>Slash ISO (YYYY/MM/DD)</td><td>2,937</td><td>2,937</td></tr>
  <tr><td>discharge_date</td><td>ISO 8601 (YYYY MM DD)</td><td>2,932</td><td>2,932</td></tr>
  <tr><td>discharge_date</td><td>Ambiguous NN/NN/YYYY resolved as MM/DD/YYYY</td><td>2,394</td><td>2,394</td></tr>
  <tr><td>discharge_date</td><td>MM/DD/YYYY (day > 12, unambiguous)</td><td>1,857</td><td>1,857</td></tr>
  <tr><td>discharge_date</td><td>DD/MM/YYYY (day > 12, unambiguous)</td><td>1,758</td><td>1,758</td></tr>
  <tr><td>discharge_date</td><td>Missing</td><td>381</td><td>0</td></tr>
</table>

## C7 Numeric validation

Each numeric column was tested against physical or definitional limits and, where no hard maximum exists, against the Tukey outer fence (Q3 plus 3 x IQR). Failing cells were blanked (never the whole row) and then treated in C9.

<table>
  <tr><th>Column</th><th>Unit</th><th>Minimum</th><th>Applied maximum</th><th>Below minimum</th><th>Above maximum</th><th>Rationale</th></tr>
  <tr><td>patient_age</td><td>years</td><td>0.00</td><td>110.00</td><td>0</td><td>0</td><td>Human age range</td></tr>
  <tr><td>length_of_stay_days</td><td>days</td><td>0.00</td><td>19.37</td><td>82</td><td>68</td><td>Negative stays impossible; sentinel 25.6 caught by fence</td></tr>
  <tr><td>time_to_triage_min</td><td>minutes</td><td>0.00</td><td>41.36</td><td>0</td><td>0</td><td>Cannot be negative</td></tr>
  <tr><td>time_to_provider_min</td><td>minutes</td><td>0.00</td><td>161.41</td><td>76</td><td>75</td><td>Negative waits impossible; sentinel 208 caught by fence</td></tr>
  <tr><td>time_to_disposition_min</td><td>minutes</td><td>0.00</td><td>465.72</td><td>0</td><td>0</td><td>Cannot be negative</td></tr>
  <tr><td>total_ed_los_min</td><td>minutes</td><td>0.00</td><td>623.05</td><td>86</td><td>65</td><td>Negative stays impossible; inflated stays (to 1,162 min) caught by fence</td></tr>
  <tr><td>physician_on_shift_count</td><td>providers</td><td>1.00</td><td>40.00</td><td>0</td><td>0</td><td>At least one provider on shift</td></tr>
  <tr><td>nurse_on_shift_count</td><td>nurses</td><td>1.00</td><td>80.00</td><td>0</td><td>0</td><td>At least one nurse on shift</td></tr>
  <tr><td>ed_bed_count</td><td>beds</td><td>1.00</td><td>200.00</td><td>0</td><td>0</td><td>Treatment spaces</td></tr>
  <tr><td>ed_occupancy_pct</td><td>%</td><td>0.00</td><td>150.00</td><td>0</td><td>0</td><td>Hallway overflow allowed to 150</td></tr>
  <tr><td>patient_satisfaction_score</td><td>score</td><td>1.00</td><td>10.00</td><td>0</td><td>0</td><td>Survey scale 1 to 10</td></tr>
  <tr><td>predicted_wait_time_min</td><td>minutes</td><td>0.00</td><td>158.56</td><td>0</td><td>0</td><td>Legacy model output; cannot be negative</td></tr>
  <tr><td>forecast_error_min</td><td>minutes</td><td>minus 300.00</td><td>300.00</td><td>0</td><td>0</td><td>Signed error of the legacy model</td></tr>
</table>

## C8 Cross field consistency

<table>
  <tr><th>Rule</th><th>Records failing</th><th>Fail rate</th><th>Decision</th></tr>
  <tr><td>Time components do not add up to ED stay</td><td>12,286</td><td>80.6 percent</td><td>Component times are authoritative (each is a measured interval and door to provider is the SLA metric). ed_los_components_min is derived; total_ed_los_min is descriptive only.</td></tr>
  <tr><td>Provider seen before triage (ESI 3 to 5)</td><td>1,808</td><td>11.9 percent</td><td>Clinically implausible for ESI 3 to 5 (ESI 1 and 2 may bypass triage). Flagged and retained; the SLA uses door to provider directly, which is unaffected.</td></tr>
  <tr><td>Left without being seen yet admitted</td><td>121</td><td>0.8 percent</td><td>Mutually exclusive outcomes. Flagged; neither flag is used in the staffing model.</td></tr>
  <tr><td>Day type contradicts date</td><td>5,735</td><td>37.6 percent</td><td>The date is authoritative. arrival_dow is derived from it and drives the day of week arrival index.</td></tr>
  <tr><td>Season contradicts arrival month</td><td>11,299</td><td>74.1 percent</td><td>The date is authoritative. season_derived replaces season in every seasonal index.</td></tr>
  <tr><td>Discharge date before arrival</td><td>7,301</td><td>47.9 percent</td><td>Same pattern as Project 1: discharge_date is unreliable and is not used; ED intervals come from the minute fields.</td></tr>
  <tr><td>ESI 1 waited over 10 minutes for a provider</td><td>383</td><td>2.5 percent</td><td>Resuscitation patients should be seen immediately. Retained and reported as a clinical safety finding; wait times in this extract do not vary by acuity (08).</td></tr>
</table>

## C9 Missing value treatment

<table>
  <tr><th>Column</th><th>Blanks before</th><th>Strategy</th><th>Filled</th><th>Blanks after</th></tr>
  <tr><td>patient_age</td><td>274</td><td>Median within facility_unit (overall median if group empty)</td><td>274</td><td>0</td></tr>
  <tr><td>length_of_stay_days</td><td>379</td><td>Median within facility_unit (overall median if group empty)</td><td>379</td><td>0</td></tr>
  <tr><td>time_to_triage_min</td><td>305</td><td>Median within triage_level (overall median if group empty)</td><td>305</td><td>0</td></tr>
  <tr><td>time_to_provider_min</td><td>456</td><td>Median within triage_level (overall median if group empty)</td><td>456</td><td>0</td></tr>
  <tr><td>time_to_disposition_min</td><td>458</td><td>Median within triage_level (overall median if group empty)</td><td>458</td><td>0</td></tr>
  <tr><td>total_ed_los_min</td><td>456</td><td>Median within triage_level (overall median if group empty)</td><td>456</td><td>0</td></tr>
  <tr><td>ed_occupancy_pct</td><td>305</td><td>Median within season (overall median if group empty)</td><td>305</td><td>0</td></tr>
  <tr><td>patient_satisfaction_score</td><td>1,220</td><td>Left blank: survey non response; imputing would invent patient opinions</td><td>0</td><td>1,220</td></tr>
  <tr><td>predicted_wait_time_min</td><td>305</td><td>Left blank: legacy model output; imputing a forecast would fabricate accuracy</td><td>0</td><td>305</td></tr>
  <tr><td>forecast_error_min</td><td>305</td><td>Left blank: legacy model error; imputing would fabricate accuracy</td><td>0</td><td>305</td></tr>
  <tr><td>patient_gender</td><td>152</td><td>Explicit category 'Unknown'</td><td>152</td><td>0</td></tr>
  <tr><td>race_ethnicity</td><td>458</td><td>Explicit category 'Unknown'</td><td>458</td><td>0</td></tr>
  <tr><td>primary_diagnosis_code</td><td>305</td><td>Explicit category 'Not Recorded'</td><td>305</td><td>0</td></tr>
  <tr><td>procedure_code</td><td>915</td><td>Explicit category 'Not Recorded'</td><td>915</td><td>0</td></tr>
  <tr><td>discharge_disposition</td><td>305</td><td>Explicit category 'Not Recorded'</td><td>305</td><td>0</td></tr>
  <tr><td>attending_physician</td><td>229</td><td>Explicit category 'Not Recorded'</td><td>229</td><td>0</td></tr>
  <tr><td>chief_complaint</td><td>305</td><td>Explicit category 'Not Recorded'</td><td>305</td><td>0</td></tr>
  <tr><td>code_status</td><td>458</td><td>Explicit category 'Unknown'</td><td>458</td><td>0</td></tr>
  <tr><td>ems_agency</td><td>305</td><td>Explicit category 'Not Recorded'</td><td>305</td><td>0</td></tr>
  <tr><td>zip_code</td><td>305</td><td>Left blank: location cannot be inferred</td><td>0</td><td>305</td></tr>
  <tr><td>admission_date</td><td>183</td><td>Left blank: arrival dates are never imputed; record excluded from dated views only</td><td>0</td><td>183</td></tr>
  <tr><td>discharge_date</td><td>381</td><td>Left blank: superseded by length of stay</td><td>0</td><td>381</td></tr>
</table>

## C10 Feature engineering

Derived fields are documented on the C10 sheet with their definition, worksheet equivalent and a live populated count.

## C11 Audit log and reconciliation

The final sheet lists every step with rows in, rows out and cells changed, and reconciles live: rows in Raw Data, less blank rows, less exact duplicates, must equal rows in Clean Data. The check returns PASS in the delivered workbook.
