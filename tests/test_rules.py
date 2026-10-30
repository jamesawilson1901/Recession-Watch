"""Acceptance tests from research pack section 10. Synthetic fixtures, not real data.

Run: python -m unittest discover -s tests
"""
import datetime as dt
import io
import os
import sys
import unittest
from decimal import Decimal as D

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "scripts"))
import openpyxl  # noqa: E402
import rules  # noqa: E402
import update_data  # noqa: E402


def levels(*growths, start=D(100)):
    """Build a level series from successive % changes."""
    out, v = [("p0", start)], start
    for i, g in enumerate(growths, 1):
        v = v * (1 + D(g) / 100)
        out.append((f"p{i}", v))
    return out


class TestGdpRules(unittest.TestCase):
    def test_two_declines_met(self):
        self.assertTrue(rules.two_quarter_rule(levels("0.5", "-0.1", "-0.1")))

    def test_decline_then_flat_not_met(self):
        self.assertFalse(rules.two_quarter_rule(levels("0.5", "-0.1", "0.0")))

    def test_only_two_levels_unknown(self):
        self.assertIsNone(rules.two_quarter_rule(levels("-0.1")))

    def test_missing_level_unknown(self):
        s = levels("0.5", "-0.1", "-0.1")
        s[2] = (s[2][0], None)
        self.assertIsNone(rules.two_quarter_rule(s))

    def test_aggregate_grows_per_capita_falls(self):
        agg = levels("0.3", "0.2", "0.1")
        pc = levels("0.1", "-0.1", "-0.1")
        self.assertFalse(rules.two_quarter_rule(agg))
        self.assertTrue(rules.two_quarter_rule(pc))

    def test_small_level_rise_is_not_decline(self):
        # published rounded growth 0.0% but levels rise (pack snapshot: 24962 -> 24971)
        s = [("q1", D(24962)), ("q2", D(24971))]
        self.assertEqual(rules.consecutive_declines(s), 0)

    def test_consecutive_decline_count_and_proximity(self):
        s = levels("0.4", "-0.2")
        self.assertEqual(rules.consecutive_declines(s), 1)
        self.assertEqual(rules.proximity_two_quarter(1), D(50))
        self.assertEqual(rules.proximity_two_quarter(3), D(100))

    def test_pp_vs_proportional(self):
        s = [("a", D(4)), ("b", D(5))]
        self.assertEqual(rules.pp_change(s), D(1))
        self.assertEqual(rules.pct_change(D(5), D(4)), D(25))


class TestSahm(unittest.TestCase):
    def test_gap_exactly_threshold_met(self):
        # 12 months at 4.0, then three at 4.5 -> u3 latest 4.5, min prior u3 4.0, gap 0.50
        s = [(i, D("4.0")) for i in range(12)] + [(12, D("4.5")), (13, D("4.5")), (14, D("4.5"))]
        gap = rules.sahm_gap(s)
        self.assertEqual(gap, D("0.5"))
        self.assertTrue(gap >= rules.SAHM_THRESHOLD)
        self.assertEqual(rules.proximity_sahm(gap), D(100))

    def test_fewer_than_15_unknown(self):
        self.assertIsNone(rules.sahm_gap([(i, D(4)) for i in range(14)]))

    def test_missing_inside_window_unknown(self):
        s = [(i, D(4)) for i in range(15)]
        s[7] = (7, None)
        self.assertIsNone(rules.sahm_gap(s))

    def test_negative_gap_proximity_zero(self):
        s = [(i, D(5) - D(i) / 10) for i in range(15)]
        self.assertLess(rules.sahm_gap(s), 0)
        self.assertEqual(rules.proximity_sahm(rules.sahm_gap(s)), D(0))


class TestEmployment(unittest.TestCase):
    def test_non_overlapping_means(self):
        s = [(i, D(v)) for i, v in enumerate([100, 100, 100, 99, 99, 99])]
        self.assertEqual(rules.emp_3m_change(s), D(-1))

    def test_needs_six(self):
        self.assertIsNone(rules.emp_3m_change([(i, D(1)) for i in range(5)]))


class TestSignalIndex(unittest.TestCase):
    def test_missing_input_not_reweighted(self):
        self.assertIsNone(rules.signal_index({"gdp_2q": D(0), "sahm": D(0), "gdppc_2q": D(0)}))

    def test_weighted(self):
        v = rules.signal_index({"gdp_2q": D(50), "sahm": D(100), "gdppc_2q": D(0), "emp_3m": D(0)})
        self.assertEqual(v, D("50.00"))


def make_abs_workbook(series):
    """series: {sid: [(date, value)]} -> bytes in ABS time-series layout."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Data1"
    ids = list(series)
    ws.append([None] + [f"desc {s}" for s in ids])
    for label in ["Unit", "Series Type", "Data Type", "Frequency", "Collection Month",
                  "Series Start", "Series End", "No. Obs"]:
        ws.append([label] + ["x"] * len(ids))
    ws.append(["Series ID"] + ids)
    dates = sorted({d for s in series.values() for d, _ in s})
    for d in dates:
        ws.append([dt.datetime(d.year, d.month, d.day)] + [dict(series[s]).get(d) for s in ids])
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


class TestParser(unittest.TestCase):
    def test_locates_series_by_label(self):
        content = make_abs_workbook({
            "A1": [(dt.date(2026, 3, 1), 1.5), (dt.date(2026, 6, 1), 2.5)],
            "B2": [(dt.date(2026, 3, 1), None), (dt.date(2026, 6, 1), 7)],
        })
        out = update_data.parse_timeseries_workbook(content, {"A1", "B2"})
        self.assertEqual(out["A1"]["obs"][-1], (dt.date(2026, 6, 1), D("2.5")))
        self.assertIsNone(out["B2"]["obs"][0][1])  # null stays null, not zero

    def test_missing_identifier_quarantines(self):
        content = make_abs_workbook({"A1": [(dt.date(2026, 3, 1), 1)]})
        with self.assertRaises(ValueError):
            update_data.parse_timeseries_workbook(content, {"A1", "ZZZ"})

    def test_quarter_label(self):
        self.assertEqual(update_data.quarter_label(dt.date(2026, 6, 1)), "2026-Q2")

    def test_cash_rate_events(self):
        obs = [(dt.date(2026, 8, 11), D("4.10")), (dt.date(2026, 8, 12), D("4.35")),
               (dt.date(2026, 8, 13), D("4.35")), (dt.date(2026, 9, 30), D("4.60"))]
        ev = update_data.cash_rate_events(obs)
        self.assertEqual([e[1] for e in ev], [D("4.10"), D("4.35"), D("4.60")])

    def test_discover_abs(self):
        html = ('<meta name="dcterms.issued" content="Thu, 15/10/2026 - 11:30" />'
                '<a href="/statistics/labour/x/sep-2026/62020001.xlsx">t</a>'
                '<div>Released</div><div>15/10/2026</div><div>Next Release</div><div>19/11/2026</div>')
        m = update_data.discover_abs(html, r"62020001\.xlsx")
        self.assertTrue(m["download_url"].endswith("sep-2026/62020001.xlsx"))
        self.assertEqual(m["next_release_date"], "2026-11-19")
        # 11:30 AEDT Sydney == 10:30 Brisbane
        when = dt.datetime.fromisoformat(m["release_at"]).astimezone(update_data.BRISBANE)
        self.assertEqual((when.hour, when.minute), (10, 30))


if __name__ == "__main__":
    unittest.main()
