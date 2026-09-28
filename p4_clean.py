"""Project 4: ED Wait Time Forecasting and Staffing Model. Cleaning pipeline."""
import glob
import numpy as np
import pandas as pd
from lib.cleaning import Cleaner

SRC = sorted(glob.glob("data/raw/*.xlsx"))[0]
BOOL = ["left_without_being_seen_flag", "boarding_flag", "ambulance_diversion_flag", "lab_order_flag", "imaging_order_flag",
        "consult_requested_flag", "admit_flag", "return_within_72h_flag", "flu_season_flag", "mass_casualty_flag"]
SEASON_OF_MONTH = {12: "Winter", 1: "Winter", 2: "Winter", 3: "Spring", 4: "Spring", 5: "Spring",
                   6: "Summer", 7: "Summer", 8: "Summer", 9: "Fall", 10: "Fall", 11: "Fall"}
N = lambda **k: dict(**k)

CFG = dict(
    key="encounter_id", anchor="staffing_scenario_id", key_prefix="ENC-", key_width=7,
    flag_identity=[("patient_id", "Flagged in patient_id_conflict_flag. Repeat visits by one patient are normal in an ED, but these repeats carry different age and sex, so they are routed for registration review rather than overwritten.")],
    categorical=["patient_gender", "race_ethnicity", "insurance_type", "facility_unit", "admission_source", "discharge_disposition",
                 "triage_level", "arrival_mode", "chief_complaint", "day_type", "season", "code_status", "ems_agency",
                 "primary_diagnosis_code", "attending_physician"],
    canonical_override={"triage_level": {f"esi {i}": f"ESI-{i}" for i in range(1, 6)}},
    id_text=["patient_id", "encounter_id", "staffing_scenario_id"], zip=["zip_code"], code_text=["procedure_code"],
    bool=BOOL, dates=["admission_date", "discharge_date"],
    numeric={
        "patient_age": N(lo=0, hi=110, fence=False, unit="years", integer=True, rationale="Human age range"),
        "length_of_stay_days": N(lo=0, hi=90, unit="days", rationale="Negative stays impossible; sentinel 25.6 caught by fence"),
        "time_to_triage_min": N(lo=0, unit="minutes", rationale="Cannot be negative"),
        "time_to_provider_min": N(lo=0, unit="minutes", rationale="Negative waits impossible; sentinel 208 caught by fence"),
        "time_to_disposition_min": N(lo=0, unit="minutes", rationale="Cannot be negative"),
        "total_ed_los_min": N(lo=0, unit="minutes", rationale="Negative stays impossible; inflated stays (to 1,162 min) caught by fence"),
        "physician_on_shift_count": N(lo=1, hi=40, fence=False, unit="providers", integer=True, rationale="At least one provider on shift"),
        "nurse_on_shift_count": N(lo=1, hi=80, fence=False, unit="nurses", integer=True, rationale="At least one nurse on shift"),
        "ed_bed_count": N(lo=1, hi=200, fence=False, unit="beds", integer=True, rationale="Treatment spaces"),
        "ed_occupancy_pct": N(lo=0, hi=150, fence=False, unit="%", rationale="Hallway overflow allowed to 150"),
        "patient_satisfaction_score": N(lo=1, hi=10, fence=False, unit="score", rationale="Survey scale 1 to 10"),
        "predicted_wait_time_min": N(lo=0, unit="minutes", rationale="Legacy model output; cannot be negative"),
        "forecast_error_min": N(lo=-300, hi=300, fence=False, unit="minutes", rationale="Signed error of the legacy model"),
    },
    impute={
        "patient_age": ("median", "facility_unit"), "length_of_stay_days": ("median", "facility_unit"),
        "time_to_triage_min": ("median", "triage_level"), "time_to_provider_min": ("median", "triage_level"),
        "time_to_disposition_min": ("median", "triage_level"), "total_ed_los_min": ("median", "triage_level"),
        "ed_occupancy_pct": ("median", "season"),
        "patient_satisfaction_score": ("leave", "survey non response; imputing would invent patient opinions"),
        "predicted_wait_time_min": ("leave", "legacy model output; imputing a forecast would fabricate accuracy"),
        "forecast_error_min": ("leave", "legacy model error; imputing would fabricate accuracy"),
        "patient_gender": ("constant", "Unknown"), "race_ethnicity": ("constant", "Unknown"),
        "primary_diagnosis_code": ("constant", "Not Recorded"), "procedure_code": ("constant", "Not Recorded"),
        "discharge_disposition": ("constant", "Not Recorded"), "attending_physician": ("constant", "Not Recorded"),
        "chief_complaint": ("constant", "Not Recorded"), "code_status": ("constant", "Unknown"), "ems_agency": ("constant", "Not Recorded"),
        "zip_code": ("leave", "location cannot be inferred"),
        "admission_date": ("leave", "arrival dates are never imputed; record excluded from dated views only"),
        "discharge_date": ("leave", "superseded by length of stay"),
    },
)

FEATURES = [
    ("arrival_date", "ED arrival date (the cleaned admission_date field)", "=admission_date"),
    ("arrival_dow", "Day of week of arrival, from the date (replaces day_type text, C8)", "=TEXT(arrival_date,\"ddd\")"),
    ("arrival_week_start", "Monday of the arrival week", "=arrival_date-WEEKDAY(arrival_date,3)"),
    ("arrival_month", "First day of the arrival month", "=DATE(YEAR(d),MONTH(d),1)"),
    ("season_derived", "Meteorological season from the arrival month (Dec to Feb Winter, Mar to May Spring, Jun to Aug Summer, Sep to Nov Fall)", "=CHOOSE(MONTH(d),\"Winter\",\"Winter\",\"Spring\",...)"),
    ("esi_level", "Triage acuity as a number (1 most urgent, 5 least)", "=VALUE(RIGHT(triage_level,1))"),
    ("door_to_provider_min", "Door to provider time, the SLA metric (= time_to_provider_min after cleaning)", "=time_to_provider_min"),
    ("sla_met_flag", "Door to provider within the 30 minute SLA", "=door_to_provider_min<=30"),
    ("ed_los_components_min", "Triage plus provider plus disposition time, the component based ED length of stay", "=time_to_triage_min+time_to_provider_min+time_to_disposition_min"),
]


def run():
    raw = pd.read_excel(SRC, sheet_name="Dataset", dtype=object)
    dd = pd.read_excel(SRC, sheet_name="Data Dictionary", header=None, dtype=object)
    cl = Cleaner(raw, CFG)
    cl.profile(); cl.structural(); cl.key_collisions(); cl.text(); cl.booleans(); cl.dates(); cl.numerics()
    wkday = lambda d: np.where(d.admission_date.dt.weekday >= 5, "Weekend", "Weekday")
    checks = [
        dict(Check="Time components do not add up to ED stay", Description="|triage + provider + disposition minus total_ed_los_min| greater than 15 minutes",
             mask=lambda d: (d.time_to_triage_min + d.time_to_provider_min + d.time_to_disposition_min - d.total_ed_los_min).abs() > 15,
             Decision="Component times are authoritative (each is a measured interval and door to provider is the SLA metric). ed_los_components_min is derived; total_ed_los_min is descriptive only."),
        dict(Check="Provider seen before triage (ESI 3 to 5)", Description="time_to_provider_min below time_to_triage_min for lower acuity patients",
             mask=lambda d: (d.time_to_provider_min < d.time_to_triage_min) & d.triage_level.isin(["ESI-3", "ESI-4", "ESI-5"]), flag="sequence_conflict_flag",
             Decision="Clinically implausible for ESI 3 to 5 (ESI 1 and 2 may bypass triage). Flagged and retained; the SLA uses door to provider directly, which is unaffected."),
        dict(Check="Left without being seen yet admitted", Description="left_without_being_seen_flag TRUE and admit_flag TRUE",
             mask=lambda d: d.left_without_being_seen_flag & d.admit_flag, flag="lwbs_conflict_flag",
             Decision="Mutually exclusive outcomes. Flagged; neither flag is used in the staffing model."),
        dict(Check="Day type contradicts date", Description="day_type Weekday or Weekend disagrees with the weekday of admission_date (Holiday excluded)",
             mask=lambda d: d.admission_date.notna() & (d.day_type != "Holiday") & (d.day_type != wkday(d)),
             Decision="The date is authoritative. arrival_dow is derived from it and drives the day of week arrival index."),
        dict(Check="Season contradicts arrival month", Description="season differs from the meteorological season of the arrival month",
             mask=lambda d: d.admission_date.notna() & (d.season != d.admission_date.dt.month.map(SEASON_OF_MONTH)),
             Decision="The date is authoritative. season_derived replaces season in every seasonal index."),
        dict(Check="Discharge date before arrival", Description="discharge_date earlier than admission_date",
             mask=lambda d: d.discharge_date < d.admission_date,
             Decision="Same pattern as Project 1: discharge_date is unreliable and is not used; ED intervals come from the minute fields."),
        dict(Check="ESI 1 waited over 10 minutes for a provider", Description="triage_level ESI-1 and time_to_provider_min above 10",
             mask=lambda d: (d.triage_level == "ESI-1") & (d.time_to_provider_min > 10),
             Decision="Resuscitation patients should be seen immediately. Retained and reported as a clinical safety finding; wait times in this extract do not vary by acuity (08)."),
    ]
    cl.cross_field(checks)
    cl.missing()
    df = cl.df
    df["arrival_date"] = df.admission_date
    df["arrival_dow"] = df.arrival_date.dt.strftime("%a")
    df["arrival_week_start"] = df.arrival_date - pd.to_timedelta(df.arrival_date.dt.weekday, unit="D")
    df["arrival_month"] = df.arrival_date.dt.to_period("M").dt.to_timestamp()
    df["season_derived"] = df.arrival_date.dt.month.map(SEASON_OF_MONTH)
    df["esi_level"] = df.triage_level.str[-1].astype(int)
    df["door_to_provider_min"] = df.time_to_provider_min
    df["sla_met_flag"] = df.door_to_provider_min <= 30
    df["ed_los_components_min"] = df.time_to_triage_min + df.time_to_provider_min + df.time_to_disposition_min
    cl._log("S10", "C10 Feature Engineering", "Derived arrival calendar keys, season, acuity and SLA fields", len(df), len(df), len(FEATURES) * len(df), f"{len(FEATURES)} features added")
    cl._log("S11", "Clean Data", "Published analysis ready table", len(df), len(df), 0, "Clean Data sheet, with named ranges for model formulas")
    cl.df = df
    return raw, dd, cl


if __name__ == "__main__":
    raw, dd, cl = run()
    d = cl.df
    print(d.shape)
    print(cl.cross[["Check", "RecordsFailing", "FailRate"]])
    print(cl.numeric_rules[["Column", "AppliedMax", "BelowMin", "AboveMax"]].to_string())
    print("door to provider mean/median", d.door_to_provider_min.mean(), d.door_to_provider_min.median(), "triage mean", d.time_to_triage_min.mean())
    print("sla met", d.sla_met_flag.mean(), "physicians mean", d.physician_on_shift_count.mean())
    print(d.groupby("triage_level").door_to_provider_min.mean())
    last = d[d.arrival_date >= d.arrival_date.max() - pd.Timedelta(days=364)]
    print("last 12m visits", len(last), d.arrival_date.min(), d.arrival_date.max())
    days = pd.date_range(d.arrival_date.min(), d.arrival_date.max())
    cnt = d.groupby("arrival_date").size().reindex(days, fill_value=0)
    print("DOW index", (cnt.groupby(cnt.index.dayofweek).mean() / cnt.mean()).round(3).to_dict())
    print("season index", (cnt.groupby(cnt.index.month.map(SEASON_OF_MONTH)).mean() / cnt.mean()).round(3).to_dict())
    fl = d.groupby("arrival_date").flu_season_flag.mean().reindex(days)
    print("flu share by season", d.groupby("season_derived").flu_season_flag.mean().round(3).to_dict())
    print("facility", d.facility_unit.value_counts().to_dict())
