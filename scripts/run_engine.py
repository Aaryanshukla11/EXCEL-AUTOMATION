"""Interactive CLI Runner for Autonomous Excel Intelligence Platform.

Allows testing and running the generic Excel intelligence pipeline from the command line:
1. Creates safe temporary copies without touching original files.
2. Discovers entities, sheets, and pre-allocated slots dynamically.
3. Generates a comprehensive mapping preview.
4. Executes deterministic slot writing.
5. Runs independent integrity validation.
6. Displays granular audit logs.
"""
import argparse
import os
import shutil
import sys

# Ensure project root is in sys.path
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_ROOT = os.path.abspath(os.path.join(SCRIPT_DIR, ".."))
if PROJECT_ROOT not in sys.path:
    sys.path.insert(0, PROJECT_ROOT)

from backend.app.services.excel_orchestrator import ExcelOrchestrator
from shared.enums import ExecutionMode, MatchConfidenceLevel


def format_header(title: str, width: int = 80) -> str:
    return f"\n{'=' * width}\n {title}\n{'=' * width}"


def main():
    parser = argparse.ArgumentParser(description="Run Autonomous Excel Intelligence Engine.")
    parser.add_argument(
        "--sister",
        type=str,
        default=os.path.join(PROJECT_ROOT, "ZEXCEL FILES", "Gagan Godha - APL2004- Oct_2026.xlsx"),
        help="Path to Sister Excel file"
    )
    parser.add_argument(
        "--master",
        type=str,
        default=os.path.join(PROJECT_ROOT, "ZEXCEL FILES", "1. Domestic Counselling Team Incentive sheet Oct_26-2.xlsx"),
        help="Path to Master Excel workbook"
    )
    parser.add_argument(
        "--auto-approve",
        action="store_true",
        help="Automatically approve proposed mappings without interactive prompt"
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Run analysis, matching, schema mapping and row-level approval in read-only preview mode (stops before write)"
    )
    args = parser.parse_args()

    if not os.path.exists(args.sister):
        print(f"Error: Sister file not found at {args.sister}")
        sys.exit(1)
    if not os.path.exists(args.master):
        print(f"Error: Master file not found at {args.master}")
        sys.exit(1)

    print(format_header("AUTONOMOUS EXCEL INTELLIGENCE PLATFORM — CLI RUNNER"))
    print(f"Source Sister File: {args.sister}")
    print(f"Master Workbook:    {args.master}")

    # Create safe temporary working copy of Master (NEVER touch originals in ZEXCEL FILES)
    temp_dir = os.path.join(PROJECT_ROOT, "storage", "temporary")
    os.makedirs(temp_dir, exist_ok=True)
    temp_master = os.path.join(temp_dir, "interactive_master_working.xlsx")
    shutil.copy2(args.master, temp_master)
    print(f"Created isolated test working copy: {temp_master}")

    orchestrator = ExcelOrchestrator(
        storage_root=os.path.join(PROJECT_ROOT, "storage"),
        execution_mode=ExecutionMode.AUTONOMOUS if args.auto_approve else ExecutionMode.SUPERVISED
    )

    # ----------------------------------------------------
    # STAGE 1: ANALYZE AND PLAN
    # ----------------------------------------------------
    print(format_header("STAGE 1: DYNAMIC WORKBOOK DISCOVERY & MAPPING PREVIEW"))
    plan, match_proposal, plan_logs = orchestrator.analyze_and_plan(
        source_filepath=args.sister,
        master_filepath=temp_master
    )

    for log in plan_logs:
        icon = "[OK]" if log.status == "SUCCESS" else "[WARN]"
        print(f"{icon} Step {log.step_number}: {log.description}")

    if plan is None:
        print("\n[STOPPED] Execution halted: No matching worksheet exists for this entity in the Master workbook.")
        print(f"Best candidate score: {match_proposal.confidence_score * 100:.1f}% ({match_proposal.confidence_level.value})")
        print("Human-in-the-loop options required: Select an existing sheet or create a new one.")
        sys.exit(0)

    print(f"\n--- MATCHING EVIDENCE ---")
    print(f"  Target Worksheet:   '{plan.target_sheet_name}'")
    print(f"  Confidence Score:    {match_proposal.confidence_score * 100:.1f}% ({match_proposal.confidence_level.value.upper()})")
    for ev in match_proposal.matching_evidence:
        print(f"  - {ev}")

    print(f"\n--- PROPOSED COLUMN MAPPINGS (PREVIEW) ---")
    print(f"{'Source Column':<30} -> {'Destination Column':<35} {'Col':<5} {'Role':<15} {'Conf'}")
    print("-" * 90)
    for m in plan.column_mappings:
        role_str = "FORMULA (SHIELD)" if m.is_formula_column else "Input Value"
        ambig_str = f" [!] {m.ambiguity_reason}" if m.is_ambiguous else ""
        print(f"{m.source_column:<30} -> {m.destination_column:<35} {m.destination_column_letter:<5} {role_str:<15} {m.confidence * 100:.0f}%{ambig_str}")

    if plan.unmapped_source_columns:
        print(f"\n--- UNMAPPED SOURCE COLUMNS (HELD FROM TRANSFER) ---")
        for u in plan.unmapped_source_columns:
            print(f"  - {u}")

    if plan.ambiguities:
        print(f"\n--- DETECTED AMBIGUITIES / COMPETING FIELDS ---")
        for amb in plan.ambiguities:
            print(f"  - [WARNING] {amb}")

    print(f"\n--- DESTINATION SLOT AVAILABILITY ---")
    print(f"  Available Pre-allocated Empty Slots: {len(plan.destination_region.available_empty_slots)} rows")
    print(f"  Empty Slot Row Indices:             {plan.destination_region.available_empty_slots[:10]}")
    print(f"  Records to Transfer:                {plan.rows_to_transfer} rows")

    if plan.batch_summary:
        bs = plan.batch_summary
        print(f"\n--- ROW-LEVEL APPROVAL & ELIGIBILITY BREAKDOWN ---")
        print(f"  Batch Decision:          {bs.batch_status.value.upper()}")
        print(f"  Total Evaluated:         {bs.total_rows} rows")
        print(f"  Eligible Records:        {bs.eligible_rows_count}")
        print(f"  Held Records:            {bs.held_rows_count}")
        print(f"  Needs Approval:          {bs.needs_approval_count}")
        if bs.partial_batch_warning:
            print(f"  [!] PARTIAL BATCH WARNING: Contains held rows requiring human override.")

        print(f"\n  {'Status':<16} {'Row':<8} {'Entity / Candidate':<28} {'Slot Assignment'}")
        print("  " + "-" * 75)
        for r_rec in bs.row_records:
            icon = f"[{r_rec.status.value.upper()}]"
            slot_str = f"-> Slot Row {r_rec.proposed_destination_row}" if r_rec.proposed_destination_row else "-> NO WRITE (HELD)"
            print(f"  {icon:<16} R{r_rec.source_row_index:<7} {r_rec.entity_identity[:26]:<28} {slot_str}")
            if r_rec.status_reason:
                print(f"      Reason: {r_rec.status_reason}")
            if r_rec.normalized_product:
                print(f"      Product Normalization: '{r_rec.normalized_product}'")

    if args.dry_run:
        print(format_header("DRY RUN COMPLETE — READ-ONLY PREVIEW"))
        print("Analysis completed successfully. No cells were modified.")
        sys.exit(0)

    # ----------------------------------------------------
    # APPROVAL GATE
    # ----------------------------------------------------
    if not args.auto_approve:
        print(format_header("HUMAN-IN-THE-LOOP APPROVAL CENTER"))
        print("Do you approve the execution of this mapping plan? (y/n): ", end="", flush=True)
        resp = input().strip().lower()
        if resp not in ["y", "yes"]:
            print("Execution rejected by user. Master workbook left completely untouched.")
            sys.exit(0)

    # ----------------------------------------------------
    # STAGE 2: EXECUTE AND VALIDATE
    # ----------------------------------------------------
    print(format_header("STAGE 2: DETERMINISTIC EXECUTION & VALIDATION"))
    success, processed_path, val_result, exec_logs = orchestrator.execute_and_validate(
        plan=plan,
        source_filepath=args.sister,
        master_filepath=temp_master
    )

    for log in exec_logs:
        icon = "[OK]" if log.status == "SUCCESS" else "[WARN]"
        print(f"{icon} Step {log.step_number}: {log.description}")

    print(format_header("STAGE 3: INDEPENDENT QUALITY & INTEGRITY AUDIT"))
    print(f"  Overall Validation Status:   {'PASSED' if val_result.is_valid else 'FAILED'}")
    print(f"  Total Checks Executed:       {val_result.total_checks}")
    print(f"  Passed Checks:               {val_result.passed_checks}")
    print(f"  Failed Checks:               {val_result.failed_checks}")
    print(f"  Formula Integrity Intact:    {val_result.formula_integrity_passed}")

    if val_result.errors:
        print(f"  Errors Detected: {val_result.errors}")

    if success:
        print(format_header("EXECUTION SUCCESSFUL"))
        print(f"Updated Master Workbook saved to: {processed_path}")
        print(f"Original files in ZEXCEL FILES were NOT modified.")
    else:
        print(format_header("EXECUTION FAILED & ROLLED BACK"))
        print("Working copy was rolled back to pristine pre-execution backup.")


if __name__ == "__main__":
    main()

