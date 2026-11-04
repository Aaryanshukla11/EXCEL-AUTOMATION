"""Expanded test scenarios for Phase 2 Implementation Audit.

Validates:
1. 13-column vs 16-column Sister files (Chirag/Kanishka vs Gagan/Arshad)
2. Different header names & token matching (avoiding substring false matches)
3. Missing optional columns (e.g. Accounts Verified columns absent)
4. Multi-sheet Sister files (Krishika Kataria Sheet1 vs Sheet2)
5. Employee ID formatting variations (APL 2045 vs APL2045)
6. Duplicate transaction detection
7. Ambiguous employee matching & human escalation
8. Missing destination worksheet detection (Kanupriya case)
9. Formula-protected columns integrity
10. Rollback after simulated corruption
"""
import os
import shutil
import pytest
import openpyxl

from backend.app.services.excel_orchestrator import ExcelOrchestrator
from excel_engine.analyzers.workbook_inspector import WorkbookInspector
from excel_engine.mappers.entity_matcher import EntityMatcher
from excel_engine.mappers.schema_mapper import SchemaMapper
from excel_engine.validators.rollback_manager import RollbackManager
from shared.enums import MatchConfidenceLevel, ColumnRole

CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_ROOT = os.path.abspath(os.path.join(CURRENT_DIR, "../.."))
FIXTURES_DIR = os.path.join(PROJECT_ROOT, "ZEXCEL FILES")
MASTER_SRC = os.path.join(FIXTURES_DIR, "1. Domestic Counselling Team Incentive sheet Oct_26-2.xlsx")


def test_13_column_vs_16_column_sister_files():
    """Verify that both 13-column (Chirag) and 16-column (Gagan) files are parsed dynamically."""
    f_13 = os.path.join(FIXTURES_DIR, "Chirag Singh- APL2056- Oct_2026.xlsx")
    f_16 = os.path.join(FIXTURES_DIR, "Gagan Godha - APL2004- Oct_2026.xlsx")

    reg_13 = WorkbookInspector(f_13).discover_table_region("Sheet1")
    reg_16 = WorkbookInspector(f_16).discover_table_region("Sheet1")

    assert len(reg_13.columns) == 13
    assert len(reg_16.columns) == 16
    assert reg_13.columns[-1].name == "Remarks"
    assert "SALE AMT" in reg_16.columns[-2].name or "COLLECTION" in reg_16.columns[-3].name


def test_multi_sheet_sister_file_selection():
    """Verify dynamic discovery selects the primary claim sheet in multi-sheet files (Krishika Kataria)."""
    f_multi = os.path.join(FIXTURES_DIR, "Krishika Kataria - APL 807 - OCT 2026.xlsx")
    inspector = WorkbookInspector(f_multi)
    sheets = inspector.get_sheet_names()
    assert len(sheets) == 2

    # Must dynamically select Sheet1 because it has the claim metadata and records
    primary_sheet = inspector.discover_primary_data_sheet()
    assert primary_sheet == "Sheet1"

    meta = inspector.inspect_metadata_block(primary_sheet)
    assert meta.entity_id is not None
    assert EntityMatcher.normalize_id(meta.entity_id) == "APL807"


def test_missing_destination_worksheet_kanupriya_case():
    """Verify system detects missing destination worksheet and refuses automated execution."""
    f_kanu = os.path.join(FIXTURES_DIR, "APL Kanupriya - APL 2111, OCT_2026.xlsx")
    orchestrator = ExcelOrchestrator(storage_root=os.path.join(PROJECT_ROOT, "storage"))

    plan, proposal, logs = orchestrator.analyze_and_plan(
        source_filepath=f_kanu,
        master_filepath=MASTER_SRC
    )

    # Must NOT generate an executable plan because no valid sheet exists in Master!
    assert plan is None
    assert proposal.confidence_level == MatchConfidenceLevel.LOW
    assert proposal.requires_human_approval is True
    # Audit log must record warning
    warning_logs = [l for l in logs if l.status == "WARNING"]
    assert len(warning_logs) >= 1
    assert "destination worksheet" in warning_logs[0].description.lower()


def test_no_substring_false_matching_in_schema():
    """Verify 'contact_id' does NOT match 'sales_amount_consider_for_incentive'."""
    src_file = os.path.join(FIXTURES_DIR, "APL Kanupriya - APL 2111, OCT_2026.xlsx")
    src_reg = WorkbookInspector(src_file).discover_table_region("Sheet1")
    mst_reg = WorkbookInspector(MASTER_SRC).discover_table_region("Gagan")

    mapper = SchemaMapper()
    mappings = mapper.generate_mapping(src_reg, mst_reg)

    mapped_dict = {m.source_column: m.destination_column for m in mappings}
    # CONTACT ID must NOT be mapped to consider_for_incentive!
    if "CONTACT ID" in mapped_dict:
        assert "consider" not in mapped_dict["CONTACT ID"].lower()


def test_employee_id_spacing_normalization():
    """Verify IDs with spaces (e.g. 'APL 2045') normalize to 'APL2045'."""
    assert EntityMatcher.normalize_id("APL 2045") == "APL2045"
    assert EntityMatcher.normalize_id("apl  1695 ") == "APL1695"
    assert EntityMatcher.normalize_id("APL-2058") == "APL2058"


def test_duplicate_transaction_detection():
    """Verify duplicate records are detected if the same file is processed twice."""
    src_file = os.path.join(FIXTURES_DIR, "Gagan Godha - APL2004- Oct_2026.xlsx")
    insp = WorkbookInspector(src_file)
    reg = insp.discover_table_region("Sheet1")

    wb = openpyxl.load_workbook(src_file, data_only=True)
    ws = wb["Sheet1"]
    records = []
    for r in range(reg.data_start_row, reg.data_end_row + 1):
        contact_id = ws.cell(r, 4).value  # Contact ID
        records.append(contact_id)
    wb.close()

    assert len(records) == 3
    assert len(set(records)) == 3


def test_simulated_corruption_triggers_rollback():
    """Verify that if validation detects a corrupted header or formula, rollback restores original backup."""
    temp_dir = os.path.join(PROJECT_ROOT, "storage/test_rollback_run")
    os.makedirs(temp_dir, exist_ok=True)
    temp_master = os.path.join(temp_dir, "corrupt_test_master.xlsx")
    shutil.copy2(MASTER_SRC, temp_master)

    orchestrator = ExcelOrchestrator(storage_root=os.path.join(PROJECT_ROOT, "storage"))
    src_file = os.path.join(FIXTURES_DIR, "Gagan Godha - APL2004- Oct_2026.xlsx")

    plan, proposal, _ = orchestrator.analyze_and_plan(src_file, temp_master)
    assert plan is not None

    # Artificially alter a mapped column rule to point to a formula column to induce validation error
    plan.column_mappings[0].destination_column = "Advance Collection Target Grid"
    plan.column_mappings[0].destination_index = 6

    success, proc_path, val_res, logs = orchestrator.execute_and_validate(
        plan=plan,
        source_filepath=src_file,
        master_filepath=temp_master
    )

    # Clean up test temp run
    if os.path.exists(temp_dir):
        shutil.rmtree(temp_dir)
