# Emergency Department Wait Time Forecasting and Staffing Model

An Excel queueing model that forecasts hourly emergency department (ED) arrivals, sizes provider staffing with M/M/c (Erlang C) queueing theory, optimises a shift roster for labour cost against a 30 minute door to provider target, and validates the plan on historical peak weeks. The model is built from a supplied 15,394 row visit extract through a fully documented, eleven step cleaning procedure, and it runs in native Excel, with Excel Solver used only to re solve the shift roster.

## Project brief

The emergency department is the front door of the hospital and one of the few services where demand cannot be scheduled. Patients arrive when illness or injury strikes, with volumes that rise and fall through the day, across the week and with the seasons. The time from arrival to first assessment by a doctor or advanced practitioner, known as door to provider time, is one of the most closely watched measures of ED performance because it is where clinical risk, patient experience and operational pressure meet. Long waits increase the chance that a deteriorating patient is missed, drive up the number of people who leave without being seen, and contribute to crowding that spreads into every other part of the hospital.

The urban ED in this study handles about 92,000 visits a year against a door to provider target of 30 minutes, yet at peak times patients are waiting more than 68 minutes. The root cause is not simply a shortage of clinicians. Provider rosters are built on flat historical averages, so the same number of providers is on duty whatever the hour, even though arrivals follow a strong daily and weekly pattern. The result is a department that is overstaffed when it is quiet and overwhelmed when it is busy. Leadership faces a familiar dilemma: adding providers across the board is expensive and hard to recruit for, but without a clear view of when demand actually peaks, there is no evidence base for redistributing the hours already funded.

This project was commissioned to replace the flat roster with a staffing plan grounded in queueing theory and optimised for cost. The brief called for an hour by hour arrival model, an M/M/c queueing approximation to estimate the wait at every possible staffing level, and a labour cost optimiser that finds the cheapest roster able to meet the 30 minute target in every hour of the week. It also required holiday and flu season surge multipliers, an interactive dashboard showing the trade off between cost and service level, and a backtest of the recommended staffing against the eight busiest historical weeks. The ambition was to move the conversation from how many more providers the ED needs to where and when the existing provider hours should be placed.

## Objectives

* Model arrival rates for all 168 hours of the week, with day of week, seasonal and surge adjustments.
* Calibrate a provider service rate so the model reproduces observed waits on the current roster.
* Use Erlang C to calculate the minimum number of providers needed in each hour to meet the door to provider target.
* Optimise a practical roster of 8 and 12 hour shifts that meets the target at minimum labour cost.
* Show the cost against service level trade off interactively, including surge and seasonal scenarios.
* Backtest the recommended plan against the eight busiest historical weeks.

## What makes this model different

Most ED staffing analyses stop at a pivot table of average arrivals. This model treats each hour as a queue with multiple servers, so it captures the non linear way waits explode as a department approaches capacity. Erlang C is calculated without recursion using native worksheet functions, a 400 by 30 staffing table covers every realistic combination of workload and headcount, and the shift roster is solved as an integer program so it is provably the cheapest feasible plan rather than a sensible looking guess. Every queueing result is checked against an independent Python reference.

## Data

<img width="1173" height="356" alt="Screenshot 2026-09-28 at 23 40 36" src="https://github.com/user-attachments/assets/900b55b2-c36c-4db4-8077-b9d3cb13850e" />


The supplied extract contains 15,394 visit rows with arrival dates, acuity (Emergency Severity Index, ESI), door to provider times, season and day type labels, and recorded staffing. Cleaning followed eleven documented steps, one per worksheet (C1 to C11).

Several data issues shaped the analysis and are worth stating plainly:

* Season and day type labels contradict the arrival dates, so both were rebuilt directly from the date.
* Arrival times of day are not recorded. The hourly arrival shape is therefore an editable planning profile, while day of week and seasonal patterns are measured from the data.
* The extract is a sample, so visit volumes were scaled to the brief's 92,000 annual visits.
* Weekly visit counts have a dispersion index of 1.05, meaning weekly variation in the sample is almost entirely random. This affects how peak weeks should be interpreted (see finding 5).

## Results at a glance

<table>
  <tr><th>Measure</th><th>Current flat roster</th><th>Recommended plan</th></tr>
  <tr><td>Provider hours per week</td><td>1,344 (8 providers every hour)</td><td>1,100 in 8 and 12 hour shifts</td></tr>
  <tr><td>Change in provider hours</td><td></td><td>244 fewer per week (18.2 percent)</td></tr>
  <tr><td>Mean door to provider time</td><td>38.8 min</td><td>20.8 min</td></tr>
  <tr><td>Hours of the week breaching the 30 minute target</td><td>38 of 168</td><td>0 of 168</td></tr>
  <tr><td>Busiest hour</td><td>Queue unstable (arrivals exceed capacity)</td><td>Within target</td></tr>
  <tr><td>Calibrated service rate</td><td colspan="2">2.03 patients per provider hour</td></tr>
</table>

## Success metrics

<table>
  <tr><th>Target</th><th>Result</th><th>Status</th></tr>
  <tr><td>Average wait under 35 minutes</td><td>20.8 minutes</td><td>Met</td></tr>
  <tr><td>Labour cost within 3 percent of the Solver optimum</td><td>Shift roster equals the proven integer optimum of 1,100 hours; 0.9 percent above the hourly minimum, which cannot be rostered in practical shifts</td><td>Met</td></tr>
  <tr><td>All eight historical peak weeks backtested</td><td>8 of 8 pass on signal adjusted demand; 0 of 8 on raw sample surges; 5 of 8 with the surge buffered roster</td><td>Met (all three lenses reported)</td></tr>
  <tr><td>SLA breach flags accurate within one hour block</td><td>Flags are hourly by design; the extract has no timestamps to test against</td><td>Not testable</td></tr>
</table>

## Findings in detail

### 1. The problem is when providers work, not how many there are

The current roster places 8 providers on duty in every hour of the week, a total of 1,344 provider hours. The recommended plan uses 1,100 hours, 244 fewer each week, yet it cuts the mean door to provider time from 38.8 to 20.8 minutes and eliminates target breaches entirely, from 38 hours of the week to none. This is the central insight of the project. The department does not need more provider hours to meet its target; it needs the hours it already funds to follow demand. Nights are currently overstaffed while late mornings are overwhelmed, and moving hours from the former to the latter improves both cost and service at the same time.

<img width="1172" height="427" alt="Screenshot 2026-09-28 at 23 42 21" src="https://github.com/user-attachments/assets/829be18d-1c6d-489c-96c5-d9da69a84716" />

### 2. At peak, the current roster cannot keep up at all

In the busiest hour of the week, patients arrive faster than the flat roster can see them. In queueing terms the system is unstable: the queue does not settle at a long wait, it grows for as long as the peak lasts and only clears once arrivals fall away. This explains the waits above 68 minutes described in the brief. No amount of working harder within that hour can fix an unstable queue; only additional capacity at that time can. The recommended plan brings every hour, including the busiest, within the 30 minute target.

<img width="1216" height="481" alt="Screenshot 2026-09-28 at 23 43 48" src="https://github.com/user-attachments/assets/226eaf39-9897-40a9-8ae0-2127f1497031" />

### 3. The recommended roster is provably the cheapest practical plan

The shift roster was solved as an integer program, so 1,100 hours is the proven minimum for a roster built from 8 and 12 hour shifts that meets the target in every hour. It sits only 0.9 percent above the theoretical hourly minimum, which cannot be achieved in practice because providers cannot be rostered one hour at a time. For finance and workforce planning, this means the saving is not the result of an optimistic assumption; it is the lowest labour cost at which the service standard can be delivered with realistic shift patterns.

<img width="718" height="336" alt="Screenshot 2026-09-28 at 23 44 40" src="https://github.com/user-attachments/assets/db1fe9c1-b48a-48ce-9f87-35d880efe5fa" />

### 4. Triage is not prioritising the sickest patients

Door to provider times are almost identical across acuity levels, at about 37 minutes for every ESI level. In a well functioning ED, the most critical patients are seen almost immediately and lower acuity patients absorb the wait. In this extract, 383 ESI 1 patients, the most critically ill category, waited more than 10 minutes. This is a clinical safety issue that staffing alone will not resolve, and it should be reviewed by clinical leadership alongside any roster change.

<img width="714" height="325" alt="Screenshot 2026-09-28 at 23 45 22" src="https://github.com/user-attachments/assets/478ffbd9-f753-43af-8339-ba378d9b1f72" />

### 5. Most of the extract's peak weeks are statistical noise

Weekly visit counts have a dispersion index of 1.05, very close to the value of 1 expected from purely random arrivals. The apparent surges in the sample are therefore mostly noise rather than genuine demand events, and scaling them up to the full 92,000 visits exaggerates them. For this reason the backtest reports three lenses side by side. On signal adjusted demand, the recommended plan passes all 8 peak weeks. On raw sample surges it passes none, because those surges are inflated by sampling noise. With a surge buffered roster it passes 5 of 8. The practical conclusion is that the base plan is sound, and that real surge periods such as flu season should be handled with a planned surge buffer rather than by permanently staffing for the worst week of the year.

<img width="643" height="337" alt="Screenshot 2026-09-28 at 23 46 05" src="https://github.com/user-attachments/assets/13fec144-38e8-4873-ab87-3dfa05156194" />

### 6. Seasonal and surge planning can be priced in advance

The surge season overlay shows weekly provider hours for 20 combinations of season and surge level, with an hourly view of where extra hours are needed. This allows workforce planners to agree surge rosters and budgets before the winter pressures arrive, rather than relying on last minute agency cover at premium rates.

<img width="1039" height="389" alt="Screenshot 2026-09-28 at 23 46 58" src="https://github.com/user-attachments/assets/c9823e8a-7173-4647-918f-f81c87fbe8d0" />

<img width="714" height="482" alt="Screenshot 2026-09-28 at 23 47 51" src="https://github.com/user-attachments/assets/be11522b-5503-41ea-9dc9-e216aafe98e0" />

<img width="732" height="365" alt="Screenshot 2026-09-28 at 23 48 39" src="https://github.com/user-attachments/assets/f37716b7-636b-4365-8058-d62961a22956" />

### 7. The queueing engine is validated and transparent

The provider service rate of 2.03 patients per provider hour was calibrated so that the model reproduces the observed mean door to provider time of 37.0 minutes on the recorded roster of 8 providers. Erlang calculations are checked against an independent Python reference, and every input that drives the result is visible and editable, so clinical and finance stakeholders can test assumptions for themselves.

<img width="1198" height="435" alt="Screenshot 2026-09-28 at 23 49 45" src="https://github.com/user-attachments/assets/f7ef7a77-5e51-4aae-942b-0d653155aa46" />

## Recommendations

1. Replace the flat roster with the demand based plan of 1,100 provider hours in 8 and 12 hour shifts, moving hours from overnight to late morning and daytime peaks.
2. Use part of the 244 released weekly hours to fund a planned surge buffer for flu season and holiday periods.
3. Refer the acuity finding to clinical leadership and review the triage process so that ESI 1 and ESI 2 patients are seen immediately.
4. Start recording arrival times and provider start times so the hourly arrival profile can be measured rather than assumed, and so breach flags can be validated.
5. Review the roster quarterly using the model, recalibrating the service rate and seasonal indices as new data arrives.
6. Replace the placeholder cost per provider hour with actual pay rates to produce a budget ready financial case.

## Method

1. **Clean the extract.** Eleven documented steps (see docs/data_cleaning_procedure.md). Season and day type labels contradict the arrival dates, so both are rebuilt from the date.
2. **Model demand.** Day of week and seasonal indices are measured from the extract and volume is scaled to the brief's 92,000 visits. Because arrival times are not recorded, the hourly shape is an editable planning profile.
3. **Calibrate service.** The provider service rate (2.03 patients per provider hour) is calibrated so the model reproduces the observed 37.0 minute mean door to provider time on the recorded roster of 8.
4. **Compute Erlang C.** Erlang C is computed without recursion using Erlang B = POISSON(c, a, FALSE) / POISSON(c, a, TRUE), and a 400 by 30 staffing table gives the minimum providers for any workload.
5. **Optimise the roster.** The shift roster is solved as an integer program (HiGHS) and can be re solved in Excel Solver.
6. **Backtest.** The plan is tested on the eight busiest historical weeks under raw, signal adjusted and surge buffered lenses.

## Workbook structure

<table>
  <tr><th>Sheet</th><th>Purpose</th></tr>
  <tr><td>01 Executive Brief</td><td>Staffing recommendation, success metrics and live findings</td></tr>
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

## How to use

Open the workbook in Microsoft Excel 2010 or later; calculation is automatic.

* Use the teal bordered selectors on sheet 03 to choose season, surge level, wait target and day.
* Edit the yellow inputs on sheet M1, including the rate per provider hour, surge multipliers and the hourly arrival profile.
* To re solve the roster after changing inputs, follow the Solver instructions on sheet 07.

The selectors are in cell drop down lists (data validation), which behave like combo boxes in every Excel version. A Form Control combo box can be linked to the same cells if preferred.

## Reproducing the build

The workbook is generated by Python, so every sheet can be rebuilt from the raw data.

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

## Assumptions and limitations

* **Hourly profile.** The hourly arrival shape and surge multipliers are planning inputs, because arrival times are not recorded in the extract.
* **Identical providers.** Providers are modelled as identical servers without breaks or handover overlap, so real rosters may need a small allowance on top of the modelled hours.
* **Acuity mix.** The queueing model treats all patients alike; it does not model priority queues by acuity.
* **Cost.** The cost per provider hour is a placeholder, so financial results scale linearly once actual pay rates are entered.
* **Breach flag validation.** Breach flags are hourly by design, but the extract has no timestamps against which their accuracy can be tested.
* **Sample scaling.** Volumes are scaled from a sample to 92,000 visits, which inflates the apparent size of sample surges.
