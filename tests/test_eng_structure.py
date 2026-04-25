"""
Test ENG model import: verify indicator counts, hierarchy, and period structure.

Checks:
1. Every sheet has a minimum number of indicators (not just top-level)
2. Sheets with hierarchical Excel data produce hierarchical records (parent_id set)
3. Period structure per sheet (visible_record_ids, period levels)

Run:
    python -m pytest tests/test_eng_structure.py -v --tb=short

Requires: ENG model already imported into pebble.db at working dir.
"""
from __future__ import annotations
import json, os, sqlite3, re
from pathlib import Path
import openpyxl
import pytest

DB_PATH = Path(os.environ.get("PEBBLE_DB", "pebble.db"))
EXCEL_PATH = Path("XLS-MODELS/ANNEX 1 Simply Ecosystem FinModel 2025-2029_ENG Final.xlsx")


def _get_eng_model_id(db: sqlite3.Connection) -> str | None:
    # Try by name first
    for mid, name in db.execute("SELECT id, name FROM models").fetchall():
        if name and ("ENG" in name or "ANNEX" in name or "Simply" in name):
            return mid
    # Fall back to the model that has sheets with Funnel/Goals excel_code
    for (mid,) in db.execute("SELECT DISTINCT m.id FROM models m JOIN sheets s ON s.model_id = m.id WHERE s.excel_code IN ('Funnel', 'Goals', 'Funnel QH')").fetchall():
        return mid
    # Last resort: most recent model
    row = db.execute("SELECT id FROM models ORDER BY rowid DESC LIMIT 1").fetchone()
    return row[0] if row else None


def _get_db():
    if not DB_PATH.exists():
        pytest.skip(f"DB not found: {DB_PATH}")
    return sqlite3.connect(str(DB_PATH))


def _count_excel_data_rows(ws, max_row: int = 500) -> dict:
    """Count text rows per column to determine label column and expected row count."""
    max_row = min(ws.max_row or 1, max_row)
    col_counts = {}
    for col in (1, 2, 3):
        text_rows = 0
        for r in range(3, max_row + 1):
            v = ws.cell(r, col).value
            if v is not None and str(v).strip() and not isinstance(v, (int, float)):
                text_rows += 1
        col_counts[col] = text_rows
    label_col = max(col_counts, key=col_counts.get)
    return {"label_col": label_col, "text_rows": col_counts[label_col], "col_counts": col_counts}


# ── Expected minimums per Excel sheet name ──
# (excel_sheet_name, min_indicators, min_hierarchical, min_depth)
# min_indicators: minimum total records
# min_hierarchical: minimum records with parent_id set
# min_depth: minimum hierarchy depth
EXPECTED_INDICATORS = {
    "Goals": (80, 50, 2),
    "Funnel": (80, 50, 1),
    "Funnel QH": (20, 8, 1),
    "packages (indiv)": (100, 50, 1),
    "Micro_30 (indiv)": (80, 60, 1),
    "Installment card_150 (indiv)": (100, 60, 1),
    "Installment card_1000 (indiv)": (100, 60, 1),
    "Debit card (indiv)": (70, 40, 1),
    "CashCredit_dossym (indiv)": (80, 30, 1),
    "BNPL_commiss (corp)": (50, 30, 1),
    "Deposits (indiv)": (120, 40, 1),
    "Merchants(corp)": (200, 60, 1),
    "OPEX_CAPEX": (60, 50, 1),
    "BS (Ecosystem)": (50, 40, 1),
    "PL (Ecosystem)": (90, 50, 1),
    "Деб.карты": (120, 50, 1),
    "CashLoans": (40, 3, 1),
    "Криейтеры (ЮЛ)": (40, 2, 1),
}


class TestENGIndicatorStructure:
    """Verify indicator counts and hierarchy for each sheet."""

    @pytest.fixture(autouse=True)
    def setup(self):
        self.db = _get_db()
        self.model_id = _get_eng_model_id(self.db)
        if not self.model_id:
            pytest.skip("ENG model not found in DB")
        yield
        self.db.close()

    def _get_sheet_indicators(self, excel_code: str):
        """Get indicator stats for a sheet by its excel_code."""
        row = self.db.execute(
            "SELECT id FROM sheets WHERE model_id = ? AND excel_code = ?",
            (self.model_id, excel_code),
        ).fetchone()
        if not row:
            return None
        sheet_id = row[0]

        # Find non-period analytics
        sas = self.db.execute("""
            SELECT sa.analytic_id FROM sheet_analytics sa
            JOIN analytics a ON a.id = sa.analytic_id
            WHERE sa.sheet_id = ? AND a.is_periods = 0
        """, (sheet_id,)).fetchall()

        total = 0
        with_parent = 0
        max_depth = 0
        for (aid,) in sas:
            recs = self.db.execute(
                "SELECT id, parent_id FROM analytic_records WHERE analytic_id = ?",
                (aid,),
            ).fetchall()
            parents = {r[0]: r[1] for r in recs}
            total += len(recs)
            with_parent += sum(1 for r in recs if r[1])
            for rid in parents:
                depth = 0
                cur = rid
                while parents.get(cur):
                    cur = parents[cur]
                    depth += 1
                    if depth > 20:
                        break
                max_depth = max(max_depth, depth)

        return {"total": total, "with_parent": with_parent, "max_depth": max_depth}

    @pytest.mark.parametrize("excel_code,expected", list(EXPECTED_INDICATORS.items()))
    def test_indicator_count(self, excel_code, expected):
        min_total, min_hier, min_depth = expected
        stats = self._get_sheet_indicators(excel_code)
        assert stats is not None, f"Sheet with excel_code={excel_code} not found"
        assert stats["total"] >= min_total, (
            f"{excel_code}: expected >= {min_total} indicators, got {stats['total']}"
        )

    @pytest.mark.parametrize("excel_code,expected", list(EXPECTED_INDICATORS.items()))
    def test_indicator_hierarchy(self, excel_code, expected):
        min_total, min_hier, min_depth = expected
        stats = self._get_sheet_indicators(excel_code)
        assert stats is not None, f"Sheet with excel_code={excel_code} not found"
        assert stats["with_parent"] >= min_hier, (
            f"{excel_code}: expected >= {min_hier} hierarchical records, got {stats['with_parent']}"
        )
        assert stats["max_depth"] >= min_depth, (
            f"{excel_code}: expected depth >= {min_depth}, got {stats['max_depth']}"
        )


class TestENGPeriodStructure:
    """Verify period structure per sheet."""

    @pytest.fixture(autouse=True)
    def setup(self):
        self.db = _get_db()
        self.model_id = _get_eng_model_id(self.db)
        if not self.model_id:
            pytest.skip("ENG model not found in DB")
        yield
        self.db.close()

    def _get_sheet_periods(self, excel_code: str) -> dict | None:
        """Get period info for a sheet."""
        row = self.db.execute(
            "SELECT id FROM sheets WHERE model_id = ? AND excel_code = ?",
            (self.model_id, excel_code),
        ).fetchone()
        if not row:
            return None
        sheet_id = row[0]

        sa_row = self.db.execute("""
            SELECT sa.analytic_id, sa.min_period_level, sa.visible_record_ids
            FROM sheet_analytics sa
            JOIN analytics a ON a.id = sa.analytic_id
            WHERE sa.sheet_id = ? AND a.is_periods = 1
        """, (sheet_id,)).fetchone()
        if not sa_row:
            return None

        aid, min_level, vis_rids_json = sa_row
        recs = self.db.execute(
            "SELECT id, data_json FROM analytic_records WHERE analytic_id = ?",
            (aid,),
        ).fetchall()

        # Parse period keys and levels
        levels = set()
        total = 0
        for rid, dj in recs:
            d = json.loads(dj)
            pk = d.get("period_key", "")
            if pk:
                total += 1
                if re.match(r'^\d{4}-Y$', pk):
                    levels.add("Y")
                elif re.match(r'^\d{4}-H\d$', pk):
                    levels.add("H")
                elif re.match(r'^\d{4}-Q\d$', pk):
                    levels.add("Q")
                elif re.match(r'^\d{4}-\d{2}$', pk) or re.match(r'^\d{4}-M\d{2}$', pk):
                    levels.add("M")

        visible_count = total
        if vis_rids_json:
            try:
                vis = json.loads(vis_rids_json)
                visible_count = len(vis)
            except:
                pass

        return {
            "total_periods": total,
            "visible_count": visible_count,
            "min_level": min_level,
            "levels": levels,
            "vis_rids_json": vis_rids_json,
        }

    def test_funnel_qh_has_qhy_not_monthly(self):
        """Funnel QH should show Q/H/Y periods, NOT monthly."""
        info = self._get_sheet_periods("Funnel QH")
        assert info is not None, "Funnel QH sheet not found"
        # If visible_record_ids is set, verify it excludes monthly
        # If min_period_level is set, verify it's Q or higher
        if info["min_level"]:
            assert info["min_level"] in ("Q", "H", "Y"), (
                f"Funnel QH min_period_level should be Q/H/Y, got {info['min_level']}"
            )
        # The visible count should be significantly less than total
        # (if all 114 periods are visible, monthly is included = wrong)
        if info["visible_count"] == info["total_periods"]:
            # All periods visible — check that means Q/H/Y only in the filter
            assert info["min_level"] in ("Q", "H", "Y"), (
                f"Funnel QH shows all {info['total_periods']} periods but min_level={info['min_level']} — should filter to Q/H/Y"
            )

    def test_all_sheets_have_periods(self):
        """Every sheet should have a period analytic bound."""
        sheets = self.db.execute(
            "SELECT id, name, excel_code FROM sheets WHERE model_id = ?",
            (self.model_id,),
        ).fetchall()
        for sid, sname, ecode in sheets:
            sa = self.db.execute("""
                SELECT COUNT(*) FROM sheet_analytics sa
                JOIN analytics a ON a.id = sa.analytic_id
                WHERE sa.sheet_id = ? AND a.is_periods = 1
            """, (sid,)).fetchone()
            assert sa[0] > 0, f"Sheet {sname} (xl={ecode}) has no period analytic"

    def test_periods_include_years(self):
        """Period analytic should include yearly periods."""
        sheets = self.db.execute(
            "SELECT excel_code FROM sheets WHERE model_id = ?",
            (self.model_id,),
        ).fetchall()
        for (ecode,) in sheets:
            info = self._get_sheet_periods(ecode)
            if info:
                assert "Y" in info["levels"], f"Sheet xl={ecode} has no yearly periods"
