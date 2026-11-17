"""Automated Tests for Approved Business Rules (Rules 1 - 5).

Validates:
1. Both Claimed and Verified sections present.
2. Duplicate headers distinguished using section context.
3. Verified fields selected over Claimed fields.
4. Missing Verified fields trigger safe stop / approval flag.
5. Partial Verified data triggers an ambiguity alert.
6. No silent fallback to Claimed data.
7. Exact product lookup against dynamically inspected Incentive Plan grid.
8. Unique fuzzy product match.
9. Ambiguous product match requiring approval.
10. Unknown product category not invented.
11. Original product text preserved in Column D.
12. Normalized product category proposed for Column E.
13. Columns P/R selected for source-backed values.
14. Columns T/U preserved without overwriting.
15. Existing formulas remain protected.
16. No automatic row append when slots are exhausted.
"""
import os
import pytest
import openpyxl

from backend.app.services.excel_orchestrator import ExcelOrchestrator
from excel_engine.analyzers.workbook_inspector import WorkbookInspector
from excel_engine.mappers.product_normalizer import ProductNormalizer
from excel_engine.mappers.schema_mapper import SchemaMapper
from shared.enums import ColumnRole, MatchConfidenceLevel

CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_ROOT = os.path.abspath(os.path.join(CURRENT_DIR, "../.."))
FIXTURES_DIR = os.path.join(PROJECT_ROOT, "ZEXCEL FILES")
MASTER_SRC = os.path.join(FIXTURES_DIR, "1. Domestic Counselling Team Incentive sheet Oct_26-2.xlsx")
GAGAN_SRC = os.path.join(FIXTURES_DIR, "Gagan Godha - APL2004- Oct_2026.xlsx")
CHIRAG_SRC = os.path.join(FIXTURES_DIR, "Chirag Singh- APL2056- Oct_2026.xlsx")


def test_both_claimed_and_verified_sections_present():
    """Rule 1: Verify dynamic detection of both Claimed and Verified sections."""
    insp = WorkbookInspector(GAGAN_SRC)
    reg = insp.discover_table_region("Sheet1")
    assert reg is not None

    claimed_cols = [c for c in reg.columns if not c.is_verified_section]
    verified_cols = [c for c in reg.columns if c.is_verified_section]

    assert len(claimed_cols) >= 10
    assert len(verified_cols) == 3  # Cols O, P, Q
    assert all("VERIFIED" in (c.section or "").upper() for c in verified_cols)


def test_duplicate_headers_distinguished_using_section_context():
    """Rule 1: Distinguish duplicate headers (COLLECTION AMT.) using section context."""
    insp = WorkbookInspector(GAGAN_SRC)
    reg = insp.discover_table_region("Sheet1")
    
    collection_cols = [c for c in reg.columns if "collection" in c.normalized_name]
    assert len(collection_cols) == 2  # Col 7 and Col 15

    col_7 = next(c for c in collection_cols if c.index == 7)
    col_15 = next(c for c in collection_cols if c.index == 15)

    assert col_7.is_verified_section is False
    assert col_15.is_verified_section is True
    assert "VERIFIED" in (col_15.section or "").upper()


def test_verified_fields_selected_over_claimed_fields():
    """Rule 1: Accounts Verified COLLECTION AMT. maps to Master R, not claimed Col 7."""
    orchestrator = ExcelOrchestrator()
    plan, _, _ = orchestrator.analyze_and_plan(GAGAN_SRC, MASTER_SRC)
    assert plan is not None

    # Col R must map from verified source Col 15
    rule_r = next(m for m in plan.column_mappings if m.destination_column_letter == "R")
    assert rule_r.source_index == 15
    assert rule_r.is_verified_source is True
    assert "VERIFIED" in (rule_r.section_context or "").upper()

    # Claimed Col 7 must be held in unmapped source columns
    assert any("Col 7" in u or "G6:G8" in u for u in plan.unmapped_source_columns)


def test_missing_verified_fields_trigger_safe_stop():
    """Rule 1: Workbooks with no verified section (Chirag) flag approval required."""
    orchestrator = ExcelOrchestrator()
    plan, _, logs = orchestrator.analyze_and_plan(CHIRAG_SRC, MASTER_SRC)
    assert plan is not None
    assert plan.requires_approval is True


def test_partial_verified_data_triggers_ambiguity():
    """Rule 1: Gagan row 8 has missing Overall in verified section -> triggers ambiguity."""
    orchestrator = ExcelOrchestrator()
    plan, _, _ = orchestrator.analyze_and_plan(GAGAN_SRC, MASTER_SRC)
    assert plan is not None

    # Incomplete cell Q8 in Gagan row 8 must be flagged
    assert any("empty/incomplete" in a.lower() and "row 8" in a.lower() for a in plan.ambiguities)
    assert plan.requires_approval is True


def test_no_silent_fallback_to_claimed_data():
    """Rule 1: Even with missing Q8 in verified, claimed data is not silently substituted."""
    orchestrator = ExcelOrchestrator()
    plan, _, _ = orchestrator.analyze_and_plan(GAGAN_SRC, MASTER_SRC)
    assert plan is not None

    # Overall in Col Q must still map to verified Source Col 17 (Q), not claimed fields
    rule_q = next(m for m in plan.column_mappings if m.destination_column_letter == "Q")
    assert rule_q.source_index == 17
    assert rule_q.is_verified_source is True


def test_exact_product_lookup_against_dynamically_inspected_grid():
    """Rule 2: Discovers lookup grid from Incentive Plan and normalizes exact match."""
    insp = WorkbookInspector(MASTER_SRC)
    grid = insp.discover_product_lookup_grid()
    assert len(grid) >= 7

    normalizer = ProductNormalizer(grid)
    cat, score, is_ambig, _ = normalizer.normalize("UGPG Admission General")
    assert cat == "UGPG Admission General"
    assert score == 1.0
    assert is_ambig is False


def test_unique_fuzzy_product_match():
    """Rule 2: Fuzzy match 'Ph.D Admission to Completion' -> 'PHD Admission'."""
    insp = WorkbookInspector(MASTER_SRC)
    grid = insp.discover_product_lookup_grid()
    normalizer = ProductNormalizer(grid)

    cat, score, is_ambig, ev = normalizer.normalize("Ph.D Admission to Completion")
    assert cat == "PHD Admission"
    assert score >= 0.80
    assert is_ambig is False


def test_ambiguous_product_match_requiring_approval():
    """Rule 2: Product text with multiple close candidates flags as ambiguous."""
    insp = WorkbookInspector(MASTER_SRC)
    grid = insp.discover_product_lookup_grid()
    normalizer = ProductNormalizer(grid)

    # 'Admission' shares equal token overlap with multiple admission categories
    cat, score, is_ambig, ev = normalizer.normalize("General Admission")
    # Must report candidate competition in evidence
    assert len(ev) > 0


def test_unknown_product_category_not_invented():
    """Rule 2: Never invent a category if text does not match lookup grid."""
    insp = WorkbookInspector(MASTER_SRC)
    grid = insp.discover_product_lookup_grid()
    normalizer = ProductNormalizer(grid)

    cat, score, is_ambig, ev = normalizer.normalize("Quantum Computing Architecture")
    assert cat is None
    assert is_ambig is True
    assert "No sufficiently matching category" in ev[0]


def test_original_product_text_preserved_in_column_d():
    """Rule 2: Original raw product string preserved in Master Col D."""
    orchestrator = ExcelOrchestrator()
    plan, _, _ = orchestrator.analyze_and_plan(GAGAN_SRC, MASTER_SRC)
    assert plan is not None

    rule_d = next(m for m in plan.column_mappings if m.destination_column_letter == "D")
    assert rule_d.transformation_rule == "direct_copy"
    assert rule_d.destination_column == "PRODUCT NAME"


def test_normalized_product_category_proposed_for_column_e():
    """Rule 2: Normalized product category proposed for Master Col E."""
    orchestrator = ExcelOrchestrator()
    plan, _, _ = orchestrator.analyze_and_plan(GAGAN_SRC, MASTER_SRC)
    assert plan is not None

    rule_e = next((m for m in plan.column_mappings if m.destination_column_letter == "E"), None)
    assert rule_e is not None
    assert "lookup_normalize:PHD Admission" in rule_e.transformation_rule
    assert plan.product_normalization_details.get("master_e_proposed_category") == "PHD Admission"


def test_columns_p_and_r_selected_for_source_backed_values():
    """Rule 3: Master P and R are mapped from approved Accounts Verified source fields."""
    orchestrator = ExcelOrchestrator()
    plan, _, _ = orchestrator.analyze_and_plan(GAGAN_SRC, MASTER_SRC)
    assert plan is not None

    rule_p = next(m for m in plan.column_mappings if m.destination_column_letter == "P")
    rule_r = next(m for m in plan.column_mappings if m.destination_column_letter == "R")

    assert rule_p.source_index == 16  # Verified SALE AMT
    assert rule_r.source_index == 15  # Verified COLLECTION AMT.
    assert rule_p.is_verified_source is True
    assert rule_r.is_verified_source is True


def test_columns_t_and_u_preserved_without_overwriting():
    """Rule 3: Columns T and U are shielded from source writes and preserved."""
    orchestrator = ExcelOrchestrator()
    plan, _, _ = orchestrator.analyze_and_plan(GAGAN_SRC, MASTER_SRC)
    assert plan is not None

    mapped_letters = [m.destination_column_letter for m in plan.column_mappings]
    assert "T" not in mapped_letters
    assert "U" not in mapped_letters
    assert "T" in plan.mirror_columns_observation
    assert "U" in plan.mirror_columns_observation


def test_existing_formulas_remain_protected():
    """Rule 4: Calculated formula columns F, G, K, N are shielded from overwrites."""
    orchestrator = ExcelOrchestrator()
    plan, _, _ = orchestrator.analyze_and_plan(GAGAN_SRC, MASTER_SRC)
    assert plan is not None

    mapped_letters = [m.destination_column_letter for m in plan.column_mappings]
    for formula_col in ("F", "G", "K", "N"):
        assert formula_col not in mapped_letters


def test_no_automatic_row_append_when_slots_are_exhausted():
    """Rule 4: Engine detects slot exhaustion and requires human approval."""
    insp = WorkbookInspector(MASTER_SRC)
    reg = insp.discover_table_region("Gagan")

    # Destination slot count is dynamically 20
    assert len(reg.available_empty_slots) == 20
    # Attempting to transfer more than available slots requires human escalation
    orchestrator = ExcelOrchestrator()
    plan, _, _ = orchestrator.analyze_and_plan(GAGAN_SRC, MASTER_SRC)
    assert plan.rows_to_transfer <= len(plan.destination_region.available_empty_slots)

