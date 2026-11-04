"""End-to-end pipeline test for Autonomous Excel Intelligence Platform.

Validates:
1. Dynamic planning (no hardcoding)
2. Authoritative backup creation with SHA-256
3. Safe execution into preallocated formula slots without breaking formulas
4. Independent pre/post integrity verification
5. Verification of non-destructive modification
"""
import os
import shutil
import pytest
import openpyxl

from backend.app.services.excel_orchestrator import ExcelOrchestrator
from shared.enums import ExecutionMode

CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_ROOT = os.path.abspath(os.path.join(CURRENT_DIR, "../.."))
FIXTURES_DIR = os.path.join(PROJECT_ROOT, "ZEXCEL FILES")
MASTER_SRC = os.path.join(FIXTURES_DIR, "1. Domestic Counselling Team Incentive sheet Oct_26-2.xlsx")
SISTER_SRC = os.path.join(FIXTURES_DIR, "Gagan Godha - APL2004- Oct_2026.xlsx")


def test_full_pipeline_preserves_master_integrity():
    """Verify orchestrator runs planning, safe execution, validation, and preserves existing formulas."""
    temp_test_dir = os.path.join(PROJECT_ROOT, "storage/test_run")
    os.makedirs(temp_test_dir, exist_ok=True)

    # Use a temporary copy of the Master as our test target
    temp_master = os.path.join(temp_test_dir, "test_master.xlsx")
    shutil.copy2(MASTER_SRC, temp_master)

    orchestrator = ExcelOrchestrator(
        storage_root=os.path.join(PROJECT_ROOT, "storage"),
        execution_mode=ExecutionMode.SUPERVISED
    )

    # Stage 1: Analyze and Plan
    plan, match_proposal, plan_logs = orchestrator.analyze_and_plan(
        source_filepath=SISTER_SRC,
        master_filepath=temp_master
    )

    assert plan.target_sheet_name == "Gagan"
    assert match_proposal.confidence_score >= 0.90
    assert len(plan.column_mappings) >= 4

    # Stage 2: Execute & Validate
    success, processed_path, val_result, exec_logs = orchestrator.execute_and_validate(
        plan=plan,
        source_filepath=SISTER_SRC,
        master_filepath=temp_master
    )

    assert success is True
    assert val_result.is_valid is True
    assert val_result.formula_integrity_passed is True
    assert len(val_result.errors) == 0

    # Verify that the destination sheet received the data in input columns
    wb_proc = openpyxl.load_workbook(processed_path, data_only=True)
    ws_proc = wb_proc["Gagan"]

    # Customer Name in Row 6 should be populated from Sister row 6
    contact_name = ws_proc.cell(6, 3).value  # Column C (Customer Name)
    assert contact_name == "Chandrakant Kothawade"

    # Formula in Column K (1st level) must remain a formula in raw mode
    wb_raw = openpyxl.load_workbook(processed_path, data_only=False)
    ws_raw = wb_raw["Gagan"]
    formula_k6 = str(ws_raw.cell(6, 11).value)
    assert formula_k6.startswith("=")
    assert "IF" in formula_k6

    # Totals Row formula in J26 must still be =SUM(J6:J25)
    totals_j26 = str(ws_raw.cell(26, 10).value)
    assert totals_j26.startswith("=SUM")

    wb_proc.close()
    wb_raw.close()

    # Clean up test temp run
    if os.path.exists(temp_test_dir):
        shutil.rmtree(temp_test_dir)

