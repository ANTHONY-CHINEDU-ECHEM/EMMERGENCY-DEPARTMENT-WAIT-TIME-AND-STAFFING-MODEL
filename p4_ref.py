"""Independent Python reference for the P4 queueing model (Erlang C via recursion, not the POISSON identity used in Excel)."""
import math
PROFILE = [2.4, 2.0, 1.7, 1.5, 1.4, 1.5, 2.0, 3.0, 4.3, 5.3, 5.9, 6.2, 6.1, 5.9, 5.7, 5.5, 5.4, 5.4, 5.5, 5.4, 5.0, 4.4, 3.7, 3.0]
S = sum(PROFILE); PROFILE = [p / S for p in PROFILE]

def erlang_c(lam, mu, c):
    a = lam / mu
    if c <= a: return 1.0
    b = 1.0
    for k in range(1, c + 1):
        b = a * b / (k + a * b)
    return c * b / (c - a * (1 - b))

def wq_min(lam, mu, c, cap=None):
    a = lam / mu
    if c <= a: return cap if cap is not None else float("inf")
    w = erlang_c(lam, mu, c) / (c * mu - lam) * 60
    return min(w, cap) if cap is not None else w

def req_c(lam, mu, target_wq, cmax=30, floor=1):
    for c in range(max(1, floor), cmax + 1):
        if wq_min(lam, mu, c) <= target_wq: return c
    return cmax
