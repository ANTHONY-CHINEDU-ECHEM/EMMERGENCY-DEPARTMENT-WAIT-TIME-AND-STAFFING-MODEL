"""Shift roster optimisation (integer program) used to prefill the Solver ready sheet."""
import math, numpy as np
from scipy.optimize import milp, LinearConstraint, Bounds
from p4_ref import PROFILE, req_c

def required_week(daily, dow_idx, mu, tq, floor, season=1.0, surge=1.0):
    req, lam = [], []
    for d in range(7):
        for h in range(24):
            l = daily * dow_idx[d] * season * surge * PROFILE[h]
            lam.append(l)
            lr = math.ceil(round(l * 10, 9)) / 10
            req.append(max(floor, req_c(lr, mu, tq)))
    return lam, req

def solve_roster(req, lengths=(8, 12)):
    H = 168; K = len(lengths)
    A = np.zeros((H, H * K))
    for k, Lk in enumerate(lengths):
        for s in range(H):
            for t in range(Lk):
                A[(s + t) % H, k * H + s] = 1
    cost = np.concatenate([np.full(H, Lk) for Lk in lengths])
    res = milp(cost, constraints=LinearConstraint(A, lb=np.array(req), ub=np.inf), integrality=np.ones(H * K), bounds=Bounds(0, np.inf),
               options={"time_limit": 120})
    x = np.round(res.x).astype(int)
    return x, res, A

if __name__ == "__main__":
    from p4_clean import run
    import pandas as pd
    raw, dd, cl = run(); d = cl.df
    days = pd.date_range(d.arrival_date.min(), d.arrival_date.max())
    tot = d.arrival_date.notna().sum() / len(days)
    idx = []
    for k, nm in enumerate(["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]):
        n = (days.dayofweek == k).sum()
        idx.append((d.arrival_dow == nm).sum() / n / tot)
    mu = 2.0332; tq = 30 - d.time_to_triage_min.mean()
    lam, req = required_week(92000 / 365, idx, mu, tq, 2)
    x, res, A = solve_roster(req)
    cov = A @ x
    print(res.status, res.message, "hours", (x[:168] * 8).sum() + (x[168:] * 12).sum(), "lower bound", sum(req), "gap", ((x[:168] * 8).sum() + (x[168:] * 12).sum()) / sum(req) - 1, "shifts", x.sum(), "min slack", (cov - np.array(req)).min())
