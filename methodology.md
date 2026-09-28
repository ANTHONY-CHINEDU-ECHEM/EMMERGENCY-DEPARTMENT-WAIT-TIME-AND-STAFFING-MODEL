# Methodology, Project 4

## Arrival model
Daily volume = 92,000 / 365 x day of week index x season index x surge multiplier; hourly arrivals = daily volume x hourly share. Day and season indices are measured from the extract; the hourly profile is a planning input because arrival times are not recorded.

## Queueing model
M/M/c with offered load a = lambda / mu. Erlang B = POISSON(c, a, FALSE) / POISSON(c, a, TRUE); Erlang C = c B / (c minus a (1 minus B)); Wq = C / (c mu minus lambda). Door to provider = measured triage time + Wq. Required providers are the smallest c whose wait meets the target, found from a 400 x 30 table by counting failing staffing levels.

## Calibration
The service rate is the single free parameter, set so the arrival weighted door to provider time at the recorded roster of 8 equals the observed 37.0 minutes.

## Roster optimisation
An integer program chooses the number of 8 and 12 hour shifts starting at each of the 168 hours of the week to cover every hourly requirement at minimum paid hours (scipy HiGHS). The Excel sheet is Solver ready with the same model.

## Backtest
The eight busiest complete weeks are forecast from calendar indices, staffed with the model rule, and exposed to realised demand in three lenses: raw uplift, uplift shrunk by measured reliability (empirical Bayes: reliability = 1 minus mean / variance of weekly counts), and raw uplift against a surge buffered roster.

## Validation
Twelve Erlang cases against a recursive Python implementation (src/p4_ref.py), 168 hourly staffing matches, calibration error, and cleaning reconciliation.
