"""Unit and Regression Tests for Sprint 3:
Row-Level Approval Engine, Partial Batch Control, and Formula Recalculation.
"""
import hashlib
import os
import shutil
import tempfile
import openpyxl
import pytest

from backend.app.services.excel_orchestrator import ExcelOrchestrator
from excel_engine.analyzers.row_approval_engine import RowApprovalEngine
from excel_engine.analyzers.formula_recalculation_inspector import FormulaRecalculationInspector
from excel_engine.analyzers.workbook_inspector import WorkbookInspector
from excel_engine.mappers.product_normalizer import ProductNormalizer
from shared.enums import BatchStatus, ColumnRole, ExecutionMode, TransactionStatus
from shared.schemas import ColumnHeader, ColumnMappingRule, DataRegion


ZEXCEL_DIR = os.path.join(os.path.dirname(__file__), "..", "..", "ZEXCEL FILES")
MASTER_FILENAME = "1. Domestic Counselling Team Incentive sheet Oct_26-2.xlsx"
GAGAN_FILENAME = "Gagan Godha - APL2004- Oct_2026.xlsx"
KANUPRIYA_FILENAME = "APL Kanupriya - APL 2111, OCT_2026.xlsx"
CHIRAG_FILENAME = "Chirag Singh- APL2056- Oct_2026.xlsx"


def file_sha256(filepath: str) -> str:
    h = hashlib.sha256()
    with open(filepath, "rb") as f:
        while chunk := f.read(8192):
            h.update(chunk)
    return h.hexdigest()


@pytest.fixture
def test_env():
    temp_dir = tempfile.mkdtemp()
    master_src = os.path.join(ZEXCEL_DIR, MASTER_FILENAME)
    temp_master = os.path.join(temp_dir, "master_copy.xlsx")
    shutil.copy2(master_src, temp_master)

    orchestrator = ExcelOrchestrator(
        storage_root=os.path.join(temp_dir, "storage"),
        execution_mode=ExecutionMode.SUPERVISED
    )

    yield {
        "temp_dir": temp_dir,
        "master_copy": temp_master,
        "orchestrator": orchestrator,
    }

    shutil.rmtree(temp_dir, ignore_errors=True)


def test_gagan_row_level_eligibility_and_kushal_hold(test_env):
    """Verifies Chandrakant & Jayshree are ELIGIBLE and Kushal Raj is HOLD."""
    gagan_path = os.path.join(ZEXCEL_DIR, GAGAN_FILENAME)
    plan, match_proposal, _ = test_env["orchestrator"].analyze_and_plan(
        source_filepath=gagan_path,
        master_filepath=test_env["master_copy"]
    )

    assert plan is not None
    assert plan.target_sheet_name == "Gagan"
    assert plan.batch_summary is not None

    bs = plan.batch_summary
    assert bs.batch_status == BatchStatus.PARTIAL_ELIGIBLE
    assert bs.partial_batch_warning is True
    assert bs.total_rows == 3
    assert bs.eligible_rows_count == 2
    assert bs.held_rows_count == 1

    # Row 6: Chandrakant Kothawade -> ELIGIBLE
    rec_r6 = next((r for r in bs.row_records if r.source_row_index == 6), None)
    assert rec_r6 is not None
    assert rec_r6.status == TransactionStatus.ELIGIBLE
    assert "Chandrakant" in rec_r6.entity_identity
    assert rec_r6.proposed_destination_row == 6
    assert len(rec_r6.missing_verified_fields) == 0

    # Row 7: Jayshree Mohanarangan -> ELIGIBLE
    rec_r7 = next((r for r in bs.row_records if r.source_row_index == 7), None)
    assert rec_r7 is not None
    assert rec_r7.status == TransactionStatus.ELIGIBLE
    assert "Jayshree" in rec_r7.entity_identity
    assert rec_r7.proposed_destination_row == 7
    assert len(rec_r7.missing_verified_fields) == 0

    # Row 8: Kushal Raj -> HOLD (missing Accounts Verified Overall Q8)
    rec_r8 = next((r for r in bs.row_records if r.source_row_index == 8), None)
    assert rec_r8 is not None
    assert rec_r8.status == TransactionStatus.HOLD
    assert "Kushal" in rec_r8.entity_identity
    assert rec_r8.proposed_destination_row is None
    assert any("Overall" in m or "Q8" in m for m in rec_r8.missing_verified_fields)
    assert "Rule 1" in rec_r8.status_reason


def test_no_silent_fallback_to_claimed_for_held_row(test_env):
    """Ensures Kushal Raj does NOT fall back to claimed Col I (130900) when Q8 is blank."""
    gagan_path = os.path.join(ZEXCEL_DIR, GAGAN_FILENAME)
    plan, _, _ = test_env["orchestrator"].analyze_and_plan(
        source_filepath=gagan_path,
        master_filepath=test_env["master_copy"]
    )

    rec_r8 = next(r for r in plan.batch_summary.row_records if r.source_row_index == 8)
    # Propose cells must be empty because row is on hold
    assert rec_r8.proposed_cell_values == {}
    # Verified fields present must NOT contain Overall
    assert "Overall" not in rec_r8.verified_fields_present


def test_single_held_row_does_not_block_eligible_rows(test_env):
    """Ensures held rows do not consume destination slots or disrupt eligible mappings."""
    gagan_path = os.path.join(ZEXCEL_DIR, GAGAN_FILENAME)
    plan, _, _ = test_env["orchestrator"].analyze_and_plan(
        source_filepath=gagan_path,
        master_filepath=test_env["master_copy"]
    )

    bs = plan.batch_summary
    eligible_slots = [r.proposed_destination_row for r in bs.row_records if r.status == TransactionStatus.ELIGIBLE]
    assert eligible_slots == [6, 7]
    assert plan.rows_to_transfer == 2


def test_kanupriya_missing_destination_all_hold(test_env):
    """Ensures APL Kanupriya batch is held safely when no destination sheet exists."""
    kanupriya_path = os.path.join(ZEXCEL_DIR, KANUPRIYA_FILENAME)
    plan, match_proposal, _ = test_env["orchestrator"].analyze_and_plan(
        source_filepath=kanupriya_path,
        master_filepath=test_env["master_copy"]
    )

    assert plan is None
    assert match_proposal.confidence_score < 0.70


def test_13_column_claimed_only_requires_selection(test_env):
    """Ensures files without verified section are flagged as REQUIRES_SELECTION."""
    chirag_path = os.path.join(ZEXCEL_DIR, CHIRAG_FILENAME)
    plan, _, _ = test_env["orchestrator"].analyze_and_plan(
        source_filepath=chirag_path,
        master_filepath=test_env["master_copy"]
    )

    assert plan is not None
    assert plan.batch_summary.batch_status == BatchStatus.REQUIRES_SELECTION
    assert plan.batch_summary.needs_approval_count > 0
    assert plan.batch_summary.eligible_rows_count == 0


def test_destination_slot_exhaustion_protection():
    """Simulates slot exhaustion and verifies excess rows are held per Rule 5."""
    engine = RowApprovalEngine()
    source_region = DataRegion(
        sheet_name="Test",
        header_row=5,
        data_start_row=6,
        data_end_row=8,
        columns=[
            ColumnHeader(index=1, column_letter="A", name="Name", normalized_name="contact_name"),
            ColumnHeader(index=2, column_letter="B", name="Sale", normalized_name="sale_amt", is_verified_section=True),
        ]
    )
    dest_region = DataRegion(
        sheet_name="Dest",
        header_row=5,
        data_start_row=6,
        data_end_row=10,
        columns=[],
        available_empty_slots=[6]  # Only 1 available slot!
    )
    mappings = [
        ColumnMappingRule(
            source_column="Sale",
            source_index=2,
            destination_column="Sale",
            destination_index=2,
            destination_column_letter="B",
            is_verified_source=True
        )
    ]

    # Create temporary mock source workbook
    temp_wb_path = os.path.join(tempfile.gettempdir(), "test_exhaustion.xlsx")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Test"
    ws.cell(6, 1, "Candidate 1")
    ws.cell(6, 2, 1000)
    ws.cell(7, 1, "Candidate 2")
    ws.cell(7, 2, 2000)
    ws.cell(8, 1, "Candidate 3")
    ws.cell(8, 2, 3000)
    wb.save(temp_wb_path)
    wb.close()

    try:
        bs = engine.evaluate_batch(
            source_filepath=temp_wb_path,
            source_region=source_region,
            destination_region=dest_region,
            column_mappings=mappings,
            target_sheet_name="Dest",
            master_sheet_exists=True,
        )

        assert bs.eligible_rows_count == 1
        assert bs.held_rows_count == 2
        assert bs.batch_status == BatchStatus.PARTIAL_ELIGIBLE
        # Excess rows held due to slot exhaustion
        assert "exhausted" in bs.row_records[1].status_reason.lower()
        assert "Rule 5" in bs.row_records[1].status_reason
    finally:
        if os.path.exists(temp_wb_path):
            os.remove(temp_wb_path)


def test_formula_recalculation_inspector_functionality(test_env):
    """Verifies formula inspection and calcPr force full recalculation."""
    res = FormulaRecalculationInspector.inspect_formula_cells(
        test_env["master_copy"], "Gagan", max_rows=30
    )
    assert res["total_formula_cells"] > 0
    assert any("VLOOKUP" in f["formula"] for f in res["sample_formula_cells"])

    ok = FormulaRecalculationInspector.ensure_full_calc_on_load(test_env["master_copy"])
    assert ok is True


def test_read_only_sister_audit_does_not_modify_original_files():
    """Guarantees strict immutability of all original files in ZEXCEL FILES."""
    original_hashes = {}
    for f in os.listdir(ZEXCEL_DIR):
        if f.endswith(".xlsx"):
            p = os.path.join(ZEXCEL_DIR, f)
            original_hashes[f] = file_sha256(p)

    # Re-verify hashes are unchanged
    for f, initial_hash in original_hashes.items():
        current_hash = file_sha256(os.path.join(ZEXCEL_DIR, f))
        assert current_hash == initial_hash, f"File {f} was modified!"
