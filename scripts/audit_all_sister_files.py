"""Comprehensive Read-Only Sister-File Auditor.

Dynamically discovers, inspects, matches, maps, and audits all Sister files
in the ZEXCEL FILES directory against the Master workbook using ExcelOrchestrator.
Operates strictly in READ-ONLY mode.
Evaluates:
- Worksheet identification & multi-sheet selection
- Column count & Accounts Verified banner detection (13 vs 16-col)
- Master target worksheet resolution & missing destination detection (Kanupriya)
- Pre-allocated slot discovery & exhaustion protection
- Row-level eligibility (Chandrakant/Jayshree vs Kushal Raj)
- Ambiguities, competing fields, and product normalizations
"""
import os
import sys
from typing import Any, Dict, List

# Ensure project root in sys.path
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_ROOT = os.path.abspath(os.path.join(SCRIPT_DIR, ".."))
if PROJECT_ROOT not in sys.path:
    sys.path.insert(0, PROJECT_ROOT)

from backend.app.services.excel_orchestrator import ExcelOrchestrator
from excel_engine.analyzers.workbook_inspector import WorkbookInspector
from shared.enums import ExecutionMode, MatchConfidenceLevel, TransactionStatus


def run_full_sister_audit(
    zexcel_dir: str = os.path.join(PROJECT_ROOT, "ZEXCEL FILES"),
    master_filename: str = "1. Domestic Counselling Team Incentive sheet Oct_26-2.xlsx"
) -> Dict[str, Any]:
    master_path = os.path.join(zexcel_dir, master_filename)
    if not os.path.exists(master_path):
        raise FileNotFoundError(f"Master workbook not found at {master_path}")

    sister_files = sorted([
        f for f in os.listdir(zexcel_dir)
        if f.endswith('.xlsx') and f != master_filename and not f.startswith('~$')
    ])

    print(f"=" * 100)
    print(f" AUTONOMOUS EXCEL INTELLIGENCE PLATFORM — READ-ONLY SISTER FILES AUDIT")
    print(f"=" * 100)
    print(f"Master Workbook: {master_filename}")
    print(f"Discovered Sister Files: {len(sister_files)}")
    print("-" * 100)

    # Initialize Orchestrator in SUPERVISED mode (strict read-only, no write)
    orchestrator = ExcelOrchestrator(
        storage_root=os.path.join(PROJECT_ROOT, "storage"),
        execution_mode=ExecutionMode.SUPERVISED
    )

    audit_summary: List[Dict[str, Any]] = []

    for idx, sister_fn in enumerate(sister_files, 1):
        sister_path = os.path.join(zexcel_dir, sister_fn)
        print(f"\n[{idx}/{len(sister_files)}] AUDITING: {sister_fn}")

        # Quick inspection of workbook sheets
        insp = WorkbookInspector(sister_path)
        sheets = insp.get_sheet_names()
        primary_sheet = insp.discover_primary_data_sheet()
        print(f"  Sheets in File:        {sheets}")
        print(f"  Primary Sheet:         '{primary_sheet}'")

        # Run Stage 1 analyze and plan (strictly read-only)
        plan, match_proposal, plan_logs = orchestrator.analyze_and_plan(
            source_filepath=sister_path,
            master_filepath=master_path
        )

        if plan is None:
            # Destination sheet missing in Master
            cand_name = match_proposal.candidate_sheet_name if match_proposal else "NONE"
            conf = match_proposal.confidence_score * 100 if match_proposal else 0.0
            conf_lvl = match_proposal.confidence_level.value.upper() if match_proposal else "UNKNOWN"
            print(f"  [STOPPED / ALL HOLD] Missing Destination Worksheet in Master Workbook!")
            print(f"  Best Candidate Guess:  '{cand_name}' (Score: {conf:.1f}%, Level: {conf_lvl})")
            print(f"  Batch Decision:        ALL_HOLD (Cannot map rows without destination worksheet)")
            audit_summary.append({
                "filename": sister_fn,
                "primary_sheet": primary_sheet,
                "target_sheet": None,
                "target_status": "MISSING_IN_MASTER",
                "batch_decision": "ALL_HOLD",
                "total_rows": 0,
                "eligible_count": 0,
                "held_count": 0,
                "requires_approval": True,
            })
            continue

        bs = plan.batch_summary
        target_sheet = plan.target_sheet_name
        avail_slots = len(plan.destination_region.available_empty_slots)
        total_cols = len(plan.source_region.columns)
        has_verified = any(c.is_verified_section for c in plan.source_region.columns)

        print(f"  Target Worksheet:      '{target_sheet}' (Confidence: {match_proposal.confidence_score*100:.1f}%)")
        print(f"  Columns:               {total_cols} cols (Verified Section: {'YES' if has_verified else 'NO - Claimed Only'})")
        print(f"  Available Slots:       {avail_slots} rows")
        print(f"  Batch Decision:        {bs.batch_status.value.upper() if bs else 'N/A'}")

        if bs:
            print(f"  Total Evaluated:       {bs.total_rows} rows")
            print(f"  Eligible Rows:         {bs.eligible_rows_count}")
            print(f"  Held Rows:             {bs.held_rows_count}")
            print(f"  Needs Approval:        {bs.needs_approval_count}")
            if bs.partial_batch_warning:
                print(f"  [!] PARTIAL BATCH WARNING: Contains held rows requiring human override.")

            for r_rec in bs.row_records:
                status_tag = f"[{r_rec.status.value.upper()}]"
                dest_str = f"-> Slot Row {r_rec.proposed_destination_row}" if r_rec.proposed_destination_row else "-> NO SLOT ASSIGNED"
                print(f"    {status_tag:<16} Row {r_rec.source_row_index} ({r_rec.entity_identity}) {dest_str}")
                if r_rec.status_reason:
                    print(f"        Reason: {r_rec.status_reason}")
                if r_rec.normalized_product:
                    print(f"        Product: '{r_rec.normalized_product}'")

        if plan.ambiguities:
            print(f"  Ambiguities ({len(plan.ambiguities)}):")
            for amb in plan.ambiguities[:3]:
                print(f"    - {amb}")

        audit_summary.append({
            "filename": sister_fn,
            "primary_sheet": primary_sheet,
            "target_sheet": target_sheet,
            "target_status": "MATCHED",
            "columns": total_cols,
            "has_verified": has_verified,
            "available_slots": avail_slots,
            "batch_decision": bs.batch_status.value if bs else "UNKNOWN",
            "total_rows": bs.total_rows if bs else 0,
            "eligible_count": bs.eligible_rows_count if bs else 0,
            "held_count": bs.held_rows_count if bs else 0,
            "needs_approval_count": bs.needs_approval_count if bs else 0,
            "partial_batch_warning": bs.partial_batch_warning if bs else False,
            "requires_approval": plan.requires_approval,
        })

    print(f"\n" + "=" * 100)
    print(f" SISTER FILES AUDIT SUMMARY TABLE")
    print(f"=" * 100)
    print(f"{'Filename':<35} {'Target Sheet':<22} {'Cols':<6} {'Slots':<6} {'Elig':<5} {'Held':<5} {'Decision'}")
    print("-" * 100)
    for row in audit_summary:
        tgt = row['target_sheet'] or "MISSING"
        print(f"{row['filename'][:34]:<35} {tgt[:21]:<22} {row.get('columns', 0):<6} {row.get('available_slots', 0):<6} {row.get('eligible_count', 0):<5} {row.get('held_count', 0):<5} {row['batch_decision']}")

    return {"summary": audit_summary}


if __name__ == "__main__":
    run_full_sister_audit()
