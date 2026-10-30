"""Fetch official ABS/RBA releases, extract verified series, evaluate rules, write data/.

Run: python scripts/update_data.py [--offline DIR]
  --offline DIR  read workbooks from DIR/<SOURCE_ID>.xlsx instead of the network (testing)

Behaviour on failure: the previous values in data/latest.json are kept and the source is
marked source_unreachable or parse_failed. Nothing is estimated or substituted.
"""
import argparse
import datetime as dt
import hashlib
import io
import json
import os
import re
import sys
import urllib.request
from decimal import Decimal
from zoneinfo import ZoneInfo

import openpyxl

sys.path.insert(0, os.path.dirname(__file__))
import rules  # noqa: E402

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA = os.environ.get("ARW_DATA_DIR") or os.path.join(ROOT, "data")
BRISBANE = ZoneInfo("Australia/Brisbane")
SYDNEY = ZoneInfo("Australia/Sydney")
UA = "AustraliaRecessionWatch/0.1 (+https://github.com/jamesawilson1901/Recession-Watch)"
PARSER_VERSION = "0.1.0"

SOURCES = {
    "ABS_NA_KEY": {
        "publisher": "ABS",
        "title": "Australian National Accounts: National Income, Expenditure and Product",
        "landing": "https://www.abs.gov.au/statistics/economy/national-accounts/australian-national-accounts-national-income-expenditure-and-product/latest-release",
        "file_pattern": r"5206001_Key_Aggregates\.xlsx",
        "keep_raw": True,
        "series": {
            "gdp_real_level": "A2304402X",
            "gdp_real_qoq_published": "A2304370T",
            "gdp_pc_real_level": "A2304404C",
            "gdp_pc_real_qoq_published": "A2304372W",
            "household_saving_ratio": "A2323382F",
            "terms_of_trade_index": "A2304200A",
            "terms_of_trade_qoq": "A2304400V",
        },
    },
    "ABS_LF_MAIN": {
        "publisher": "ABS",
        "title": "Labour Force, Australia (Table 1)",
        "landing": "https://www.abs.gov.au/statistics/labour/employment-and-unemployment/labour-force-australia/latest-release",
        "file_pattern": r"62020001\.xlsx",
        "keep_raw": True,
        "series": {
            "unemployment_rate": "A84423050A",
            "employment_level": "A84423043C",
            "employment_population_ratio": "A84423054K",
        },
    },
    "RBA_F1": {
        "publisher": "RBA",
        "title": "RBA Statistical Table F1.1 Interest Rates and Yields - Money Market",
        "url": "https://www.rba.gov.au/statistics/tables/xls/f01hist.xlsx",
        "keep_raw": False,  # changes daily; extracted series only
        "series": {"cash_rate_target": "FIRMMCRTD"},
    },
}


# ---------------------------------------------------------------- acquisition

def http_get(url, timeout=60):
    req = urllib.request.Request(url, headers={"User-Agent": UA})
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return r.read(), r.geturl()


def discover_abs(landing_html, file_pattern, base="https://www.abs.gov.au"):
    """Find the release workbook link and release metadata on an ABS latest-release page."""
    m = re.search(r'href="([^"]*/' + file_pattern + r')"', landing_html)
    if not m:
        raise ValueError(f"download link matching {file_pattern} not found on landing page")
    href = m.group(1)
    url = href if href.startswith("http") else base + href
    meta = {"download_url": url}
    issued = re.search(r'name="dcterms\.issued"\s+content="[A-Za-z]{3}, (\d{2})/(\d{2})/(\d{4}) - (\d{2}):(\d{2})"', landing_html)
    if issued:
        d, mo, y, h, mi = map(int, issued.groups())
        # ABS states release times in Canberra local time (AEST/AEDT).
        meta["release_at"] = dt.datetime(y, mo, d, h, mi, tzinfo=SYDNEY).isoformat()
    temporal = re.search(r'name="dcterms\.temporal"\s+content="([^"]+)"', landing_html)
    if temporal:
        meta["reference_period"] = temporal.group(1)
    nxt = re.search(r"Next release[^0-9]{0,200}?(\d{1,2})/(\d{1,2})/(\d{4})", landing_html, re.S)
    if nxt:
        d, mo, y = map(int, nxt.groups())
        meta["next_release_date"] = dt.date(y, mo, d).isoformat()  # date only; time not verified
    return meta


# ---------------------------------------------------------------- parsing

def _cell_str(v):
    return "" if v is None else str(v).strip()


def parse_timeseries_workbook(content, wanted_ids):
    """Parse an ABS/RBA time-series workbook. Locates the 'Series ID' row by label.

    Returns {series_id: {"meta": {...}, "obs": [(date, Decimal|None), ...]}}.
    Raises ValueError if any wanted ID is missing (whole source is quarantined).
    """
    wb = openpyxl.load_workbook(io.BytesIO(content), read_only=True, data_only=True)
    found = {}
    for ws in wb.worksheets:
        if not ws.title.lower().startswith("data"):
            continue
        rows = list(ws.iter_rows(values_only=True))
        id_row = next((i for i, r in enumerate(rows) if r and _cell_str(r[0]).lower() == "series id"), None)
        if id_row is None:
            continue
        header = rows[id_row]
        meta_rows = {(_cell_str(r[0]) if r else ""): r for r in rows[:id_row]}
        for col, sid in enumerate(header):
            sid = _cell_str(sid)
            if sid not in wanted_ids or sid in found:
                continue
            meta = {"description": _cell_str(rows[0][col]) if rows[0] and col < len(rows[0]) else ""}
            for label in ("Unit", "Units", "Series Type", "Data Type", "Frequency", "Type", "Title", "Description"):
                r = meta_rows.get(label)
                if r is not None and col < len(r):
                    meta[label.lower().replace(" ", "_")] = _cell_str(r[col])
            obs, seen = [], set()
            for r in rows[id_row + 1:]:
                if not r or not isinstance(r[0], (dt.datetime, dt.date)):
                    continue
                d = r[0].date() if isinstance(r[0], dt.datetime) else r[0]
                if d in seen:
                    raise ValueError(f"{sid}: duplicate date {d}")
                seen.add(d)
                v = r[col] if col < len(r) else None
                if v is None or (isinstance(v, str) and not v.strip()):
                    val = None
                elif isinstance(v, (int, float)):
                    val = Decimal(repr(v))
                else:
                    try:
                        val = Decimal(str(v).strip())
                    except Exception:
                        val = None
                obs.append((d, val))
            obs.sort(key=lambda x: x[0])
            # trim trailing empties (future rows), keep interior gaps as None
            while obs and obs[-1][1] is None:
                obs.pop()
            found[sid] = {"meta": meta, "obs": obs}
    missing = set(wanted_ids) - set(found)
    if missing:
        raise ValueError(f"series not found in workbook: {sorted(missing)}")
    return found


def quarter_label(d):
    return f"{d.year}-Q{(d.month - 1) // 3 + 1}"


def month_label(d):
    return f"{d.year}-{d.month:02d}"


def cash_rate_events(obs):
    """Collapse a daily cash-rate-target series into change events (effective dates)."""
    events, prev = [], None
    for d, v in obs:
        if v is None:
            continue
        if prev is None or v != prev:
            events.append((d, v))
            prev = v
    return events


# ---------------------------------------------------------------- storage helpers

def load_json(path, default):
    try:
        with open(path) as f:
            return json.load(f)
    except FileNotFoundError:
        return default


def write_json(path, obj):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    tmp = path + ".tmp"
    with open(tmp, "w") as f:
        json.dump(obj, f, indent=1, ensure_ascii=False)
        f.write("\n")
    os.replace(tmp, path)  # atomic


def f(x, nd=None):
    if x is None:
        return None
    return float(round(x, nd)) if nd is not None else float(x)


def ser_out(obs, label_fn):
    return [[label_fn(d), f(v)] for d, v in obs]


def diff_revisions(old_series, new_series, label):
    """Compare historical values present in both vintages; return changed periods."""
    old = {p: v for p, v in old_series}
    changes = []
    for p, v in new_series:
        if p in old and old[p] is not None and v is not None and abs(old[p] - v) > 1e-9:
            changes.append({"series": label, "period": p, "old": old[p], "new": v})
    return changes


# ---------------------------------------------------------------- main

def acquire(source_id, cfg, offline_dir):
    """Return (bytes, asset_meta)."""
    if offline_dir:
        with open(os.path.join(offline_dir, f"{source_id}.xlsx"), "rb") as fh:
            return fh.read(), {"download_url": f"offline:{source_id}.xlsx"}
    if "landing" in cfg:
        html, _ = http_get(cfg["landing"])
        meta = discover_abs(html.decode("utf-8", "replace"), cfg["file_pattern"])
        content, final = http_get(meta["download_url"])
        meta["final_url"] = final
        return content, meta
    content, final = http_get(cfg["url"])
    return content, {"download_url": cfg["url"], "final_url": final}


def build_indicators(series):
    """series: {indicator_id: [(date, Decimal|None)]}. Returns indicator + rule payload."""
    out, prox = {}, {}

    def q(ind):
        return series.get(ind) or []

    gdp = q("gdp_real_level")
    if gdp:
        dec = rules.consecutive_declines(gdp)
        g = rules.growth_series(gdp)
        prox["gdp_2q"] = rules.proximity_two_quarter(dec)
        out["gdp_real"] = {
            "latest_period": quarter_label(gdp[-1][0]),
            "level": f(gdp[-1][1]),
            "qoq_calc": f(g[-1][1], 4) if g and g[-1][1] is not None else None,
            "qoq_published": f(q("gdp_real_qoq_published")[-1][1]) if q("gdp_real_qoq_published") else None,
            "annual": f(rules.annual_change(gdp), 4),
            "consecutive_declines": dec,
            "rule": {"id": "R-GDP-2Q", "version": rules.RULE_VERSIONS["R-GDP-2Q"],
                     "met": rules.two_quarter_rule(gdp), "proximity": f(prox["gdp_2q"])},
            "history": [[quarter_label(d), f(v, 4)] for d, v in g[-24:]],
        }

    pc = q("gdp_pc_real_level")
    if pc:
        dec = rules.consecutive_declines(pc)
        g = rules.growth_series(pc)
        prox["gdppc_2q"] = rules.proximity_two_quarter(dec)
        out["gdp_pc_real"] = {
            "latest_period": quarter_label(pc[-1][0]),
            "level": f(pc[-1][1]),
            "qoq_calc": f(g[-1][1], 4) if g and g[-1][1] is not None else None,
            "qoq_published": f(q("gdp_pc_real_qoq_published")[-1][1]) if q("gdp_pc_real_qoq_published") else None,
            "annual": f(rules.annual_change(pc), 4),
            "consecutive_declines": dec,
            "rule": {"id": "R-GDPPC-2Q", "version": rules.RULE_VERSIONS["R-GDPPC-2Q"],
                     "met": rules.two_quarter_rule(pc), "proximity": f(prox["gdppc_2q"])},
            "history": [[quarter_label(d), f(v, 4)] for d, v in g[-24:]],
        }

    sav = q("household_saving_ratio")
    if sav:
        out["saving_ratio"] = {
            "latest_period": quarter_label(sav[-1][0]),
            "value": f(sav[-1][1]),
            "pp_change": f(rules.pp_change(sav), 4),
            "pp_change_annual": f(rules.pp_change(sav, 4), 4),
            "history": [[quarter_label(d), f(v)] for d, v in sav[-24:]],
        }

    tot = q("terms_of_trade_index")
    if tot:
        g = rules.growth_series(tot)
        out["terms_of_trade"] = {
            "latest_period": quarter_label(tot[-1][0]),
            "value": f(tot[-1][1]),
            "qoq_calc": f(g[-1][1], 4) if g and g[-1][1] is not None else None,
            "annual": f(rules.annual_change(tot), 4),
            "history": [[quarter_label(d), f(v)] for d, v in tot[-24:]],
        }

    u = q("unemployment_rate")
    if u:
        gap = rules.sahm_gap(u)
        prox["sahm"] = rules.proximity_sahm(gap)
        out["unemployment"] = {
            "latest_period": month_label(u[-1][0]),
            "value": f(u[-1][1]),
            "pp_change": f(rules.pp_change(u), 4),
            "pp_change_annual": f(rules.pp_change(u, 12), 4),
            "rule": {"id": "R-SAHM-AU-050", "version": rules.RULE_VERSIONS["R-SAHM-AU-050"],
                     "gap_pp": f(gap, 4), "threshold_pp": 0.50, "alt_threshold_pp": 0.75,
                     "met": None if gap is None else gap >= rules.SAHM_THRESHOLD,
                     "proximity": f(prox["sahm"]),
                     "proximity_alt": f(rules.proximity_sahm(gap, rules.SAHM_ALT_THRESHOLD)),
                     "status": "experimental"},
            "history": [[month_label(d), f(v)] for d, v in u[-36:]],
            "gap_history": [[month_label(d), f(v, 4)] for d, v in rules.sahm_gap_history(u)],
        }

    e = q("employment_level")
    if e:
        ch = rules.emp_3m_change(e)
        prox["emp_3m"] = rules.proximity_emp(ch)
        out["employment"] = {
            "latest_period": month_label(e[-1][0]),
            "value": f(e[-1][1]),
            "mom_change_pct": f(rules.pct_change(e[-1][1], e[-2][1]) if len(e) > 1 else None, 4),
            "rule": {"id": "R-EMP-3M", "version": rules.RULE_VERSIONS["R-EMP-3M"],
                     "change_3m_pct": f(ch, 4), "met": None if ch is None else ch < 0,
                     "proximity": f(prox["emp_3m"]), "status": "experimental"},
            "history": [[month_label(d), f(v, 4)] for d, v in rules.emp_3m_history(e)],
        }

    ep = q("employment_population_ratio")
    if ep:
        out["employment_population"] = {
            "latest_period": month_label(ep[-1][0]),
            "value": f(ep[-1][1]),
            "pp_change": f(rules.pp_change(ep), 4),
            "pp_change_annual": f(rules.pp_change(ep, 12), 4),
            "history": [[month_label(d), f(v)] for d, v in ep[-36:]],
        }

    cr = q("cash_rate_target")
    if cr:
        ev = cash_rate_events(cr)
        out["cash_rate"] = {
            "value": f(ev[-1][1]),
            "effective_date": ev[-1][0].isoformat(),
            "previous_value": f(ev[-2][1]) if len(ev) > 1 else None,
            "previous_effective_date": ev[-2][0].isoformat() if len(ev) > 1 else None,
            "last_observation_date": cr[-1][0].isoformat(),
            "announcement_note": "Effective dates from RBA F1.1. Announcement timestamps are not captured.",
            "history": [[d.isoformat(), f(v)] for d, v in ev[-16:]],
        }

    idx = rules.signal_index(prox)
    signal = {
        "id": "SIGNAL-INDEX", "version": rules.RULE_VERSIONS["SIGNAL-INDEX"],
        "value": f(idx, 1),
        "weights": {k: float(v) for k, v in rules.INDEX_WEIGHTS.items()},
        "inputs": {k: f(prox.get(k)) for k in rules.INDEX_WEIGHTS},
        "note": "Experimental stress score 0-100. Not a recession probability. Not validated against Australian history.",
    }

    gdppc = out.get("gdp_pc_real", {})
    emp_rule = out.get("employment", {}).get("rule", {})
    sahm_rule = out.get("unemployment", {}).get("rule", {})
    phases = [
        {"id": 1, "name": "Household squeeze",
         "check": "Real GDP per person fell in the latest quarter or over the past year",
         "observed": None if not gdppc else bool((gdppc.get("consecutive_declines") or 0) >= 1 or (gdppc.get("annual") or 0) < 0)},
        {"id": 2, "name": "Business & jobs breakpoint",
         "check": "Employment 3-month average falling, or unemployment Sahm gap at or above 0.50 pp (insolvency data pending)",
         "observed": None if (emp_rule.get("met") is None and sahm_rule.get("met") is None) else bool(emp_rule.get("met") or sahm_rule.get("met"))},
        {"id": 3, "name": "Technical recession",
         "check": "Real GDP fell in two consecutive quarters",
         "observed": out.get("gdp_real", {}).get("rule", {}).get("met")},
    ]
    current = max([p["id"] for p in phases if p["observed"]], default=0)
    return out, signal, {"phases": phases, "current_phase": current,
                         "note": "Scenario mapping of the original infographic onto observed data. Not a forecast or countdown."}


def main(argv=None):
    ap = argparse.ArgumentParser()
    ap.add_argument("--offline")
    args = ap.parse_args(argv)

    now = dt.datetime.now(dt.timezone.utc)
    sources_state = load_json(os.path.join(DATA, "sources.json"), {})
    store = load_json(os.path.join(DATA, "series.json"), {})  # indicator_id -> [[iso_date, value]]
    revisions = load_json(os.path.join(DATA, "revisions.json"), [])
    errors = []
    data_changed = False

    for sid, cfg in SOURCES.items():
        st = sources_state.setdefault(sid, {})
        st.update({"publisher": cfg["publisher"], "title": cfg["title"], "last_check_at": now.isoformat()})
        try:
            content, meta = acquire(sid, cfg, args.offline)
        except Exception as ex:  # network / discovery failure: keep last values
            st.update({"state": "source_unreachable", "last_error": f"{type(ex).__name__}: {ex}"[:300]})
            errors.append(f"{sid}: {st['last_error']}")
            continue
        digest = hashlib.sha256(content).hexdigest()
        if digest == st.get("content_hash") and all(k in store for k in cfg["series"]):
            st.update({"state": "current", "last_error": None})
            continue
        try:
            parsed = parse_timeseries_workbook(content, set(cfg["series"].values()))
        except Exception as ex:  # schema change: quarantine, no partial publish
            st.update({"state": "parse_failed", "last_error": f"{type(ex).__name__}: {ex}"[:300]})
            errors.append(f"{sid}: {st['last_error']}")
            continue

        new_vals = {}
        for ind, series_id in cfg["series"].items():
            obs = parsed[series_id]["obs"]
            new_vals[ind] = [[d.isoformat(), None if v is None else str(v)] for d, v in obs]
            old = [(p, float(v)) for p, v in store.get(ind, []) if v is not None]
            new = [(p, float(v)) for p, v in new_vals[ind] if v is not None]
            for ch in diff_revisions(old, new, ind):
                ch.update({"source": sid, "detected_at": now.isoformat()})
                revisions.append(ch)
        store.update(new_vals)
        data_changed = True

        if cfg["keep_raw"]:
            raw = os.path.join(DATA, "raw", sid, f"{digest[:16]}.xlsx")
            if not os.path.exists(raw):
                os.makedirs(os.path.dirname(raw), exist_ok=True)
                with open(raw, "wb") as fh:
                    fh.write(content)
        st.update({
            "state": "current", "last_error": None, "content_hash": digest,
            "last_success_at": now.isoformat(), "parser_version": PARSER_VERSION,
            "series_meta": {k: parsed[v]["meta"] for k, v in cfg["series"].items()},
            **{k: v for k, v in meta.items()},
        })

    revisions = revisions[-200:]
    as_dec = {k: [(dt.date.fromisoformat(p), None if v is None else Decimal(v)) for p, v in obs]
              for k, obs in store.items()}
    indicators, signal, scenario = build_indicators(as_dec)

    latest = {
        "schema": "arw-latest/1",
        "generated_at": now.isoformat(),
        "generated_at_brisbane": now.astimezone(BRISBANE).strftime("%d %b %Y %H:%M AEST"),
        "signal_index": signal,
        "scenario": scenario,
        "indicators": indicators,
        "sources": {k: {kk: vv for kk, vv in v.items() if kk != "series_meta"} for k, v in sources_state.items()},
        "revisions": revisions[-30:],
        "errors": errors,
    }
    write_json(os.path.join(DATA, "sources.json"), sources_state)
    write_json(os.path.join(DATA, "revisions.json"), revisions)
    if data_changed:
        write_json(os.path.join(DATA, "series.json"), store)
    write_json(os.path.join(DATA, "latest.json"), latest)

    for e in errors:
        print("ERROR", e, file=sys.stderr)
    print(f"data_changed={str(data_changed).lower()} errors={len(errors)}")
    gh_out = os.environ.get("GITHUB_OUTPUT")
    if gh_out:
        with open(gh_out, "a") as fh:
            fh.write(f"data_changed={str(data_changed).lower()}\nerrors={len(errors)}\n")
    return 0


if __name__ == "__main__":
    sys.exit(main())
