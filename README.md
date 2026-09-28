# Emergency Department Wait Time Forecasting and Staffing Model

An Excel queueing model that forecasts hourly ED arrivals, sizes provider staffing with M/M/c (Erlang C) theory, optimises a shift roster for cost against a 30 minute door to provider target, and validates the plan on historical peak weeks. Built from a supplied 15,394 row visit extract through a fully documented, eleven step cleaning procedure.

Excel Project 4 of 10 in a data analyst portfolio. Prepared by Anthony Chinedu Echem.

## Business problem

An urban ED with 92,000 annual visits sees door to provider times above 68 minutes at peaks against a 30 minute target, because staffing follows a fixed roster that ignores demand patterns.

## Results at a glance

<table>
  <tr><th>Measure</th><th>Current flat roster</th><th>Recommended plan</th></tr>
  <tr><td>Provider hours per week</td><td>1,344</td><td>1,100 in 8 and 12 hour shifts</td></tr>
  <tr><td>Mean door to provider time</td><td>38.8 min</td><td>20.8 min</td></tr>
  <tr><td>Hours of the week breaching target</td><td>38 of 168</td><td>0 of 168</td></tr>
  <tr><td>Busiest hour</td><td>queue unstable (arrivals exceed capacity)</td><td>within target</td></tr>
</table>

## Success metrics

<table>
  <tr><th>Catalogue metric</th><th>Result</th><th>Status</th></tr>
  <tr><td>Average wait under 35 minutes</td><td>20.8 minutes</td><td>Met</td></tr>
  <tr><td>Labor cost within 3 percent of the Solver optimum</td><td>Shift roster equals the proven integer optimum (1,100 hours); 0.9 percent above the unrosterable hourly minimum</td><td>Met</td></tr>
  <tr><td>All eight historical peak weeks backtested</td><td>8 of 8 pass on signal adjusted demand; 0 of 8 on raw sample surges; 5 of 8 with the surge buffered roster</td><td>Met (all three lenses reported)</td></tr>
  <tr><td>SLA breach flags accurate within one hour block</td><td>Flags are hourly by design; the extract has no timestamps to test against</td><td>Not testable</td></tr>
</table>

## Workbook structure

<table>
  <tr><th>Sheet</th><th>Purpose</th></tr>
  <tr><td>01 Executive Brief</td><td>Staffing recommendation, success metrics, live findings</td></tr>
  <tr><td>02 Hourly Staffing Optimizer</td><td>Required providers, current waits and optimised waits for all 168 hours</td></tr>
  <tr><td>03 Cost vs SLA Dashboard</td><td>Season, surge, target and day selectors; staffing and wait charts; cost trade off curve</td></tr>
  <tr><td>04 MMc Wait Calculator</td><td>Erlang C metrics and staffing curve for any scenario</td></tr>
  <tr><td>05 Surge Season Overlay</td><td>Weekly hours for 20 season and surge combinations with an hourly overlay</td></tr>
  <tr><td>06 Eight Week Backtest</td><td>Eight peak weeks forecast, staffed and exposed to realised demand under three lenses</td></tr>
  <tr><td>07 Solver Shift Roster</td><td>Optimal 8 and 12 hour shift plan, ready to re solve with Excel Solver</td></tr>
  <tr><td>08 Assumptions and Validation</td><td>Erlang validation, calibration, metrics, assumptions and findings</td></tr>
  <tr><td>M1 to M3</td><td>Arrival model and calibration, Erlang staffing table, 168 hour demand engine</td></tr>
  <tr><td>C1 to C11</td><td>Data cleaning procedure, one step per sheet</td></tr>
</table>

## Method

1. Clean the extract (see docs/data_cleaning_procedure.md). Season and day type labels contradict the arrival dates, so both are rebuilt from the date.
2. Measure day of week and seasonal indices from the extract and scale volume to the brief's 92,000 visits. Arrival times are not recorded, so the hourly shape is an editable planning profile.
3. Calibrate the provider service rate (2.03 patients per provider hour) so the model reproduces the observed 37.0 minute mean door to provider time on the recorded roster of 8.
4. Compute Erlang C without recursion using Erlang B = POISSON(c, a, FALSE) / POISSON(c, a, TRUE), and build a 400 x 30 staffing table.
5. Solve the shift roster as an integer program (HiGHS) and backtest on the eight busiest historical weeks.

## Key findings

* The mismatch, not the headcount, causes the waits: nights are overstaffed while late mornings exceed provider capacity. Demand based staffing uses fewer hours and still meets the target in every hour.
* Waits do not vary by acuity (about 37 minutes at every ESI level), and 383 ESI 1 patients waited over 10 minutes, a clinical safety issue.
* Weekly visit counts have a dispersion index of 1.05, so the extract's peak weeks are almost entirely sampling noise; the backtest reports raw, adjusted and buffered results side by side.

## How to use

Open the workbook in Microsoft Excel 2010 or later. Use the teal bordered selectors on 03 and edit the yellow inputs on M1 (rate per provider hour, surge multipliers, hourly profile). To re solve the roster, follow the Solver instructions on 07. The selectors are in cell drop down lists (data validation), which behave like combo boxes in every Excel version; a Form Control combo box can be linked to the same cells if preferred.

## Reproducing the build

```
pip install pandas numpy openpyxl scipy
python3 src/p4_build.py
```

Run from the repository root. src/p4_ref.py is the independent queueing reference and src/p4_milp.py solves the shift roster.

## Repository structure

```
README.md
workbook/      finished Excel model
data/raw/      supplied dataset, unchanged
src/           cleaning engine, queueing reference, roster optimiser and build script
docs/          data cleaning procedure and methodology
```

## Limitations

Hourly arrival shape and surge multipliers are planning inputs; providers are modelled as identical servers without breaks or handover overlap; the cost per provider hour is a placeholder that scales results linearly.

## Author

Anthony Chinedu Echem
