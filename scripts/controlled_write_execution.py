"""Controlled Write Execution and Independent Verification Harness.

Executes controlled write testing on disposable copies:
1. Creates isolated disposable copy of Master workbook.
2. Uses Gagan Sister file as source.
3. Generates mapping plan & row-level approval preview.
4. Bypasses HELD row (Kushal Raj) and transfers only ELIGIBLE rows (Chandrakant & Jayshree).
5. OpenXML deterministic execution preserving tables, formulas, validations, and formatting.
6. Independent reopen verification: checks written cells, verifies held row is blank.
7. Generates cell-by-cell diff between original and written output.
8. Tests rollback mechanism on a separate disposable copy.
9. Verifies SHA-256 of original files in ZEXCEL FILES remain strictly untouched.
"""
import hashlib
import os
import shutil
import sys
import tempfile
import openpyxl
import xml.etree.ElementTree as ET
import zipfile

# Ensure project root in sys.path
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_ROOT = os.path.abspath(os.path.join(SCRIPT_DIR, ".."))
if PROJECT_ROOT not in sys.path:
    sys.path.insert(0, PROJECT_ROOT)

from backend.app.services.excel_orchestrator import ExcelOrchestrator
from excel_engine.analyzers.formula_recalculation_inspector import FormulaRecalculationInspector
from excel_engine.analyzers.workbook_inspector import WorkbookInspector
from excel_engine.validators.rollback_manager import RollbackManager
from shared.enums import ExecutionMode, TransactionStatus


def get_sha256(filepath: str) -> str:
    h = hashlib.sha256()
    with open(filepath, "rb") as f:
        while chunk := f.read(8192):
            h.update(chunk)
    return h.hexdigest()


def run_controlled_write_verification():
    zexcel_dir = os.path.join(PROJECT_ROOT, "ZEXCEL FILES")
    master_original = os.path.join(zexcel_dir, "1. Domestic Counselling Team Incentive sheet Oct_26-2.xlsx")
    gagan_original = os.path.join(zexcel_dir, "Gagan Godha - APL2004- Oct_2026.xlsx")

    # Record baseline hashes of original files
    orig_master_hash = get_sha256(master_original)
    orig_gagan_hash = get_sha256(gagan_original)

    print("=" * 100)
    print(" AUTONOMOUS EXCEL INTELLIGENCE PLATFORM — CONTROLLED WRITE EXECUTION & VERIFICATION")
    print("=" * 100)
    print(f"Original Master: {master_original} (SHA256: {orig_master_hash[:16]}...)")
    print(f"Original Sister: {gagan_original} (SHA256: {orig_gagan_hash[:16]}...)")

    # Step 1: Create an isolated disposable copy of Master workbook
    temp_dir = tempfile.mkdtemp(prefix="controlled_write_test_")
    disposable_master = os.path.join(temp_dir, "disposable_master.xlsx")
    shutil.copy2(master_original, disposable_master)
    print(f"\n[STEP 1] Created isolated disposable Master copy: {disposable_master}")

    # Step 2: Initialize Orchestrator in SUPERVISED mode
    orchestrator = ExcelOrchestrator(
        storage_root=os.path.join(temp_dir, "storage"),
        execution_mode=ExecutionMode.SUPERVISED
    )

    # Step 3: Generate Mapping and Row-Level Approval Preview
    print("\n[STEP 2 & 3] Generating mapping and row-level approval preview...")
    plan, match_proposal, plan_logs = orchestrator.analyze_and_plan(
        source_filepath=gagan_original,
        master_filepath=disposable_master
    )

    print(f"  Target Worksheet:      '{plan.target_sheet_name}' (Confidence: {match_proposal.confidence_score*100:.1f}%)")
    print(f"  Available Empty Slots: {len(plan.destination_region.available_empty_slots)} rows")
    print(f"  Batch Decision:        {plan.batch_summary.batch_status.value.upper()}")
    print(f"  Partial Batch Warning: {plan.batch_summary.partial_batch_warning}")

    for r_rec in plan.batch_summary.row_records:
        tag = f"[{r_rec.status.value.upper()}]"
        slot = f"-> Slot Row {r_rec.proposed_destination_row}" if r_rec.proposed_destination_row else "-> NO WRITE (HELD)"
        print(f"    {tag:<16} Row {r_rec.source_row_index} ({r_rec.entity_identity}) {slot}")
        if r_rec.status_reason:
            print(f"        Reason: {r_rec.status_reason}")

    # Step 4 & 5: Transfer only ELIGIBLE rows; Kushal Raj (Row 8) is HELD
    print("\n[STEP 4 & 5] Executing controlled write for ELIGIBLE rows only...")
    success, processed_path, val_result, exec_logs = orchestrator.execute_and_validate(
        plan=plan,
        source_filepath=gagan_original,
        master_filepath=disposable_master
    )

    for log in exec_logs:
        icon = "[OK]" if log.status == "SUCCESS" else "[WARN]"
        print(f"  {icon} Step {log.step_number}: {log.description}")

    # Set full calc on load for formula recalculation
    FormulaRecalculationInspector.ensure_full_calc_on_load(processed_path)

    # Step 6 & 7: Independent Reopen and Cell Value Inspection
    print("\n[STEP 6 & 7] Independent Reopen of Written Workbook...")
    wb_read = openpyxl.load_workbook(processed_path, data_only=True)
    ws_gagan = wb_read["Gagan"]

    # Read Row 6: Chandrakant Kothawade (ELIGIBLE)
    r6_vals = {
        "Name (C6)": ws_gagan["C6"].value,
        "Raw Product (D6)": ws_gagan["D6"].value,
        "Norm Product (E6)": ws_gagan["E6"].value,
        "Inv No (H6)": ws_gagan["H6"].value,
        "Sale Amt (P6)": ws_gagan["P6"].value,
        "Overall (Q6)": ws_gagan["Q6"].value,
        "Collection (R6)": ws_gagan["R6"].value,
        "Salary Shield (S6)": ws_gagan["S6"].value,
        "Mirror Col T (T6)": ws_gagan["T6"].value,
        "Mirror Col U (U6)": ws_gagan["U6"].value,
    }
    print("\n  >> Written Row 6 (Chandrakant Kothawade - ELIGIBLE):")
    for k, v in r6_vals.items():
        print(f"       {k:<20}: {v}")

    # Read Row 7: Jayshree Mohanarangan (ELIGIBLE)
    r7_vals = {
        "Name (C7)": ws_gagan["C7"].value,
        "Raw Product (D7)": ws_gagan["D7"].value,
        "Norm Product (E7)": ws_gagan["E7"].value,
        "Inv No (H7)": ws_gagan["H7"].value,
        "Sale Amt (P7)": ws_gagan["P7"].value,
        "Overall (Q7)": ws_gagan["Q7"].value,
        "Collection (R7)": ws_gagan["R7"].value,
        "Salary Shield (S7)": ws_gagan["S7"].value,
        "Mirror Col T (T7)": ws_gagan["T7"].value,
        "Mirror Col U (U7)": ws_gagan["U7"].value,
    }
    print("\n  >> Written Row 7 (Jayshree Mohanarangan - ELIGIBLE):")
    for k, v in r7_vals.items():
        print(f"       {k:<20}: {v}")

    # Read Row 8: Kushal Raj (HELD - MUST BE UNWRITTEN / EMPTY)
    r8_vals = {
        "Name (C8)": ws_gagan["C8"].value,
        "Raw Product (D8)": ws_gagan["D8"].value,
        "Sale Amt (P8)": ws_gagan["P8"].value,
        "Overall (Q8)": ws_gagan["Q8"].value,
        "Collection (R8)": ws_gagan["R8"].value,
    }
    print("\n  >> Inspected Row 8 (Kushal Raj - HELD):")
    for k, v in r8_vals.items():
        print(f"       {k:<20}: {v} (Expected: None / Unwritten)")

    wb_read.close()

    # Verify Kushal Raj was NOT written
    assert ws_gagan["C8"].value is None or ws_gagan["C8"].value == "", "Violation: Kushal Raj was written!"
    assert ws_gagan["P8"].value is None, "Violation: Row 8 Sale Amt was written!"
    print("\n  [VERIFIED] Held record (Kushal Raj) was NOT written to Row 8.")

    # Step 8: Detailed Cell Diff vs Original Master
    print("\n[STEP 8] Detailed Cell-by-Cell Diff Against Original Master:")
    wb_orig = openpyxl.load_workbook(master_original, data_only=True)
    ws_orig = wb_orig["Gagan"]
    wb_mod = openpyxl.load_workbook(processed_path, data_only=True)
    ws_mod = wb_mod["Gagan"]

    diffs = []
    # Check all cells in sheet Gagan up to row 35, col 26
    for r in range(1, 35):
        for c in range(1, 27):
            c_orig = ws_orig.cell(r, c).value
            c_mod = ws_mod.cell(r, c).value
            col_letter = openpyxl.utils.get_column_letter(c)
            ref = f"{col_letter}{r}"
            if c_orig != c_mod:
                diffs.append((ref, c_orig, c_mod))

    wb_orig.close()
    wb_mod.close()

    print(f"  Total Modified Cells in 'Gagan': {len(diffs)}")
    print(f"  {'Cell Ref':<10} {'Original Value':<30} {'Written Value':<35}")
    print("  " + "-" * 75)
    for ref, old_v, new_v in diffs:
        print(f"  {ref:<10} {str(old_v)[:28]:<30} {str(new_v)[:33]:<35}")

    # Check non-target sheets are 100% byte-exact
    insp_orig = WorkbookInspector(master_original)
    insp_mod = WorkbookInspector(processed_path)
    all_sheets = insp_orig.get_sheet_names()
    print(f"\n  Checking non-target sheets integrity across all {len(all_sheets)} sheets...")
    with zipfile.ZipFile(master_original, 'r') as z_orig, zipfile.ZipFile(processed_path, 'r') as z_mod:
        gagan_xml = insp_orig.get_sheet_target_path("Gagan")
        non_target_corrupted = 0
        for s in all_sheets:
            if s == "Gagan":
                continue
            s_xml = insp_orig.get_sheet_target_path(s)
            b_orig = z_orig.read(s_xml)
            b_mod = z_mod.read(s_xml)
            if b_orig != b_mod:
                non_target_corrupted += 1
                print(f"    [ERROR] Non-target sheet '{s}' changed!")

        assert non_target_corrupted == 0, "Non-target sheets were modified!"
        print(f"  [VERIFIED] All {len(all_sheets) - 1} non-target worksheets remain 100% byte-exact identical!")

    # Step 9: Test Rollback Mechanism on Separate Disposable Copy
    print("\n[STEP 9] Testing Rollback Mechanism on Separate Disposable Copy...")
    rollback_master = os.path.join(temp_dir, "rollback_test_master.xlsx")
    shutil.copy2(master_original, rollback_master)
    initial_hash = get_sha256(rollback_master)

    backup_path = os.path.join(temp_dir, "rollback_backup.xlsx")
    shutil.copy2(rollback_master, backup_path)
    assert os.path.exists(backup_path)

    # Simulate deliberate corruption
    with open(rollback_master, "wb") as f:
        f.write(b"CORRUPTED BY SIMULATION")
    assert get_sha256(rollback_master) != initial_hash

    # Trigger Rollback
    rb_mgr = RollbackManager()
    rb_success, rb_msg = rb_mgr.rollback(backup_path, rollback_master, initial_hash)
    assert rb_success is True
    assert get_sha256(rollback_master) == initial_hash
    print(f"  [VERIFIED] Rollback restored corrupted workbook to initial SHA-256 ({initial_hash[:16]}...) successfully.")

    # Step 10: Final Verification of Original Files Immutability
    final_master_hash = get_sha256(master_original)
    final_gagan_hash = get_sha256(gagan_original)
    assert final_master_hash == orig_master_hash, "CRITICAL: Original Master was modified!"
    assert final_gagan_hash == orig_gagan_hash, "CRITICAL: Original Gagan was modified!"
    print(f"\n[STEP 10] Verified original files in ZEXCEL FILES are 100% unchanged:")
    print(f"  Master SHA-256 unchanged: {final_master_hash[:16]}...")
    print(f"  Gagan  SHA-256 unchanged: {final_gagan_hash[:16]}...")

    # Cleanup temp directory
    shutil.rmtree(temp_dir, ignore_errors=True)
    print("\n" + "=" * 100)
    print(" CONTROLLED WRITE EXECUTION & INDEPENDENT VERIFICATION COMPLETED SUCCESSFULLY")
    print("=" * 100)
    return True


if __name__ == "__main__":
    run_controlled_write_verification()
