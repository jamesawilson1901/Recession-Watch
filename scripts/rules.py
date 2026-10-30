"""Deterministic rule calculations for Australia Recession Watch.

Pure functions only: no I/O, no interpolation, no generated values.
Rule definitions follow the Stage 1 research pack (v0.1.0-stage1), section 6.
A series is a list of (period, Decimal | None) tuples in ascending period order.
None means missing; any rule whose window touches a missing value returns None (unknown).
"""
from decimal import Decimal

RULE_VERSIONS = {
    "R-GDP-2Q": "0.1.0",
    "R-GDPPC-2Q": "0.1.0",
    "R-GDPPC-RUN": "0.1.0",
    "R-SAHM-AU-050": "0.1.0",
    "R-EMP-3M": "0.1.0",
    "SIGNAL-INDEX": "0.1.0-experimental",
}

SAHM_THRESHOLD = Decimal("0.50")
SAHM_ALT_THRESHOLD = Decimal("0.75")  # reported RBA adaptation; primary source not obtained

# Experimental composite weights. Published in the UI; not validated.
INDEX_WEIGHTS = {
    "gdp_2q": Decimal("0.40"),
    "sahm": Decimal("0.30"),
    "gdppc_2q": Decimal("0.15"),
    "emp_3m": Decimal("0.15"),
}


def pct_change(new, old):
    if new is None or old is None or old == 0:
        return None
    return Decimal(100) * (new / old - 1)


def growth_series(series):
    """Period-on-period % change, computed from unrounded levels."""
    out = []
    for i in range(1, len(series)):
        out.append((series[i][0], pct_change(series[i][1], series[i - 1][1])))
    return out


def consecutive_declines(series):
    """Count of consecutive negative changes ending at the latest period.

    Returns None if the latest change is unknown or a gap breaks the run before a
    non-negative change is found.
    """
    g = growth_series(series)
    if not g:
        return None
    run = 0
    for _, v in reversed(g):
        if v is None:
            return None
        if v < 0:
            run += 1
        else:
            return run
    return run  # every available change negative; run limited by history


def two_quarter_rule(series):
    """R-GDP-2Q / R-GDPPC-2Q: latest two changes both negative. Needs 3 levels."""
    if len(series) < 3:
        return None
    a, b, c = (v for _, v in series[-3:])
    if a is None or b is None or c is None:
        return None
    return (b < a) and (c < b)


def annual_change(series, lag=4):
    if len(series) <= lag:
        return None
    return pct_change(series[-1][1], series[-1 - lag][1])


def _mean(vals):
    return sum(vals) / Decimal(len(vals))


def sahm_gap(series):
    """R-SAHM-AU-050 gap in percentage points. Needs 15 consecutive monthly values."""
    if len(series) < 15:
        return None
    window = [v for _, v in series[-15:]]
    if any(v is None for v in window):
        return None
    u3 = [_mean(window[i - 2:i + 1]) for i in range(2, 15)]  # 13 values, latest last
    return u3[-1] - min(u3[:-1])


def sahm_gap_history(series, n=24):
    out = []
    for end in range(max(15, len(series) - n + 1), len(series) + 1):
        out.append((series[end - 1][0], sahm_gap(series[:end])))
    return out


def emp_3m_change(series):
    """R-EMP-3M: mean of latest 3 months vs previous non-overlapping 3 months, in %."""
    if len(series) < 6:
        return None
    w = [v for _, v in series[-6:]]
    if any(v is None for v in w):
        return None
    return pct_change(_mean(w[3:]), _mean(w[:3]))


def emp_3m_history(series, n=24):
    out = []
    for end in range(max(6, len(series) - n + 1), len(series) + 1):
        out.append((series[end - 1][0], emp_3m_change(series[:end])))
    return out


def pp_change(series, lag=1):
    if len(series) <= lag:
        return None
    a, b = series[-1][1], series[-1 - lag][1]
    if a is None or b is None:
        return None
    return a - b


def proximity_two_quarter(declines):
    """0, 50 or 100: share of the two required consecutive declines observed."""
    if declines is None:
        return None
    return Decimal(min(declines, 2) * 50)


def proximity_sahm(gap, threshold=SAHM_THRESHOLD):
    if gap is None:
        return None
    return max(Decimal(0), min(Decimal(100), Decimal(100) * gap / threshold))


def proximity_emp(change):
    if change is None:
        return None
    return Decimal(100) if change < 0 else Decimal(0)


def signal_index(proximities):
    """Experimental weighted stress score 0-100. Not a recession probability.

    Missing inputs are not reweighted: any unknown input makes the index unknown.
    """
    if any(proximities.get(k) is None for k in INDEX_WEIGHTS):
        return None
    return sum(INDEX_WEIGHTS[k] * proximities[k] for k in INDEX_WEIGHTS)
