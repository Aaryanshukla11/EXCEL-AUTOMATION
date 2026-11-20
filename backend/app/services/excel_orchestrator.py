"""Generic Excel Orchestrator Service.

Coordinates the full autonomous pipeline across inspection, matching, mapping,
safe execution, validation, and rollback without hardcoded logic.
"""
import os
import shutil
import uuid
from typing import Any, Dict, List, Optional, Tuple
import openpyxl

from excel_engine.analyzers.workbook_inspector import WorkbookInspector
from excel_engine.analyzers.row_approval_engine import RowApprovalEngine
from excel_engine.mappers.entity_matcher import EntityMatcher
from excel_engine.mappers.schema_mapper import SchemaMapper
from excel_engine.mappers.product_normalizer import ProductNormalizer
from excel_engine.validators.integrity_validator import IntegrityValidator
from excel_engine.validators.rollback_manager import RollbackManager
from excel_engine.writers.deterministic_writer import DeterministicWriter
from shared.enums import AuditAction, ExecutionMode, MatchConfidenceLevel, TaskStatus, ColumnRole
from shared.schemas import (
    EntityMatchProposal,
    ExecutionStepLog,
    MappingPlan,
    ValidationResult,
)


class ExcelOrchestrator:
    """Core pipeline coordinator for autonomous Excel intelligence."""

    def __init__(
        self,
        storage_root: str = "storage",
        execution_mode: ExecutionMode = ExecutionMode.SUPERVISED
    ):
        self.storage_root = storage_root
        self.execution_mode = execution_mode
        self.backup_dir = os.path.join(storage_root, "backups")
        self.temp_dir = os.path.join(storage_root, "temporary")
        self.processed_dir = os.path.join(storage_root, "processed")

        for d in [self.backup_dir, self.temp_dir, self.processed_dir]:
            os.makedirs(d, exist_ok=True)

        self.matcher = EntityMatcher()
        self.schema_mapper = SchemaMapper()
        self.writer = DeterministicWriter(backup_dir=self.backup_dir)
        self.validator = IntegrityValidator()
        self.rollback_mgr = RollbackManager()

    def analyze_and_plan(
        self,
        source_filepath: str,
        master_filepath: str,
        summary_sheet_name: Optional[str] = None
    ) -> Tuple[Optional[MappingPlan], EntityMatchProposal, List[ExecutionStepLog]]:
        """Stage 1: Discover structure, match entity, and formulate mapping plan."""
        audit_logs: List[ExecutionStepLog] = []
        task_id = str(uuid.uuid4())[:8]

        # Step 1: Inspect Source File - Dynamically discover primary claim sheet
        src_inspector = WorkbookInspector(source_filepath)
        primary_src_sheet = src_inspector.discover_primary_data_sheet()

        src_meta = src_inspector.inspect_metadata_block(primary_src_sheet)
        src_region = src_inspector.discover_table_region(primary_src_sheet)

        audit_logs.append(ExecutionStepLog(
            step_number=1,
            action=AuditAction.STRUCTURE_INSPECT,
            status="SUCCESS",
            description=f"Inspected source file. Primary sheet: '{primary_src_sheet}', Identified entity: ID={src_meta.entity_id}, Name={src_meta.entity_name}",
            details={"metadata": src_meta.model_dump(), "source_sheet": primary_src_sheet}
        ))

        # Step 2: Inspect Master Workbook - Dynamically discover summary sheet if not given
        mst_inspector = WorkbookInspector(master_filepath)
        mst_sheets = mst_inspector.get_sheet_names()

        discovered_summary_sheet = summary_sheet_name or mst_inspector.discover_summary_sheet()

        linkages = []
        if discovered_summary_sheet and discovered_summary_sheet in mst_sheets:
            linkages = mst_inspector.extract_formula_linkages_from_summary(discovered_summary_sheet)

        # Match entity to destination sheet
        match_proposals = self.matcher.match(
            source_entity=src_meta,
            candidate_sheets=mst_sheets,
            summary_linkages=linkages
        )

        best_proposal = match_proposals[0] if match_proposals else None

        # Check for Missing Destination Sheet (e.g. Kanupriya case where no sheet exists)
        if not best_proposal or best_proposal.confidence_level == MatchConfidenceLevel.LOW:
            audit_logs.append(ExecutionStepLog(
                step_number=2,
                action=AuditAction.ENTITY_MATCH,
                status="WARNING",
                description=f"No high-confidence destination worksheet found for {src_meta.entity_name} (ID: {src_meta.entity_id}). Requires user selection or sheet creation.",
                details={"best_guess": best_proposal.candidate_sheet_name if best_proposal else None, "score": best_proposal.confidence_score if best_proposal else 0.0}
            ))
            return None, best_proposal, audit_logs

        audit_logs.append(ExecutionStepLog(
            step_number=2,
            action=AuditAction.ENTITY_MATCH,
            status="SUCCESS",
            description=f"Matched source to Master sheet '{best_proposal.candidate_sheet_name}' (Confidence: {best_proposal.confidence_score*100:.1f}%)",
            details={
                "sheet": best_proposal.candidate_sheet_name,
                "confidence": best_proposal.confidence_score,
                "evidence": best_proposal.matching_evidence
            }
        ))

        # Step 3: Inspect Destination Sheet Structure & Discovered Reference Grids
        dest_region = mst_inspector.discover_table_region(best_proposal.candidate_sheet_name)
        if not dest_region:
            raise ValueError(f"Could not discover tabular data region in destination sheet '{best_proposal.candidate_sheet_name}'")

        # Discover authoritative product lookup grid from Incentive Plan sheet (Rule 2)
        product_grid = mst_inspector.discover_product_lookup_grid()
        self.schema_mapper.product_grid = product_grid
        self.schema_mapper.product_normalizer = ProductNormalizer(product_grid) if product_grid else None

        # Sample raw product name from source for normalization resolution
        sample_product = None
        src_prod_col = next((c for c in src_region.columns if "product" in c.normalized_name), None)
        if src_prod_col:
            wb_temp = openpyxl.load_workbook(source_filepath, data_only=True)
            ws_temp = wb_temp[src_region.sheet_name]
            for r in range(src_region.data_start_row, min(src_region.data_start_row + 3, src_region.data_end_row + 1)):
                val = ws_temp.cell(r, src_prod_col.index).value
                if val and str(val).strip():
                    sample_product = str(val).strip()
                    break
            wb_temp.close()

        # Step 4: Generate Semantic Schema Mapping adhering to Rules 1-4
        mappings = self.schema_mapper.generate_mapping(src_region, dest_region, sample_source_product=sample_product)
        unmapped = getattr(self.schema_mapper, "last_unmapped_columns", [])
        ambiguities = list(getattr(self.schema_mapper, "last_ambiguities", []))
        prod_norm_details = getattr(self.schema_mapper, "last_product_normalization_details", {})

        # RULE 1 Validation: Check for incomplete / missing data in Accounts Verified section
        verified_mappings = [m for m in mappings if m.is_verified_source]
        if verified_mappings:
            wb_chk = openpyxl.load_workbook(source_filepath, data_only=True)
            ws_chk = wb_chk[src_region.sheet_name]
            for r in range(src_region.data_start_row, src_region.data_end_row + 1):
                for vm in verified_mappings:
                    cell_val = ws_chk.cell(r, vm.source_index).value
                    if cell_val is None or str(cell_val).strip() == "":
                        cell_ref = f"{openpyxl.utils.get_column_letter(vm.source_index)}{r}"
                        ambiguities.append(
                            f"Row {r}: Accounts Verified field '{vm.source_column}' (cell {cell_ref}) is empty/incomplete. Human approval required (never silently falling back to claimed data)."
                        )
            wb_chk.close()

        # RULE 3 Observation: Inspect columns T & U in destination table
        mirror_obs: Dict[str, str] = {}
        for letter in ("T", "U"):
            col_obj = next((c for c in dest_region.columns if c.column_letter == letter), None)
            if col_obj:
                role_desc = "Formulas" if col_obj.role == ColumnRole.CALCULATED_FORMULA else "Unwritten input slots (preserved per Rule 3)"
                mirror_obs[letter] = f"Column {letter} ('{col_obj.name}'): {role_desc}. Retained unchanged."
            else:
                mirror_obs[letter] = f"Column {letter}: Not present in destination table."

        audit_logs.append(ExecutionStepLog(
            step_number=3,
            action=AuditAction.SCHEMA_MAP,
            status="SUCCESS" if not ambiguities else "WARNING",
            description=f"Generated {len(mappings)} semantic column mappings" + (f" with {len(ambiguities)} ambiguities" if ambiguities else ""),
            details={
                "mapped_columns": [m.model_dump() for m in mappings],
                "unmapped_source_columns": unmapped,
                "ambiguities": ambiguities,
                "product_normalization": prod_norm_details,
                "mirror_columns_observation": mirror_obs
            }
        ))

        # Target start row: use first preallocated empty slot if available, else append at data_end_row + 1
        target_start_row = (
            dest_region.available_empty_slots[0]
            if dest_region.available_empty_slots
            else (dest_region.data_end_row + 1)
        )

        # Phase 1: Row-Level Approval Engine
        row_approval_engine = RowApprovalEngine(product_normalizer=self.schema_mapper.product_normalizer)
        batch_summary = row_approval_engine.evaluate_batch(
            source_filepath=source_filepath,
            source_region=src_region,
            destination_region=dest_region,
            column_mappings=mappings,
            target_sheet_name=best_proposal.candidate_sheet_name,
            master_sheet_exists=True,
        )

        has_unresolved_issues = (
            bool(ambiguities)
            or best_proposal.requires_human_approval
            or prod_norm_details.get("is_ambiguous", False)
            or batch_summary.partial_batch_warning
            or batch_summary.held_rows_count > 0
            or batch_summary.requires_explicit_approval
        )

        plan = MappingPlan(
            plan_id=task_id,
            source_filename=os.path.basename(source_filepath),
            master_filename=os.path.basename(master_filepath),
            target_sheet_name=best_proposal.candidate_sheet_name,
            source_region=src_region,
            destination_region=dest_region,
            column_mappings=mappings,
            rows_to_transfer=batch_summary.eligible_rows_count if batch_summary.eligible_rows_count > 0 else (src_region.data_end_row - src_region.data_start_row + 1),
            target_start_row=target_start_row,
            requires_approval=has_unresolved_issues or (self.execution_mode != ExecutionMode.AUTONOMOUS),
            status=TaskStatus.AWAITING_APPROVAL if has_unresolved_issues else TaskStatus.PENDING,
            unmapped_source_columns=unmapped,
            ambiguities=ambiguities,
            product_normalization_details=prod_norm_details,
            mirror_columns_observation=mirror_obs,
            batch_summary=batch_summary,
        )

        return plan, best_proposal, audit_logs

    def execute_and_validate(
        self,
        plan: MappingPlan,
        source_filepath: str,
        master_filepath: str,
    ) -> Tuple[bool, str, ValidationResult, List[ExecutionStepLog]]:
        """Stage 2: Execute approved plan on working copy, validate, and promote."""
        audit_logs: List[ExecutionStepLog] = []
        step_idx = 4

        # 1. Authoritative Backup
        backup_path, backup_hash = self.writer.create_backup(master_filepath, plan.plan_id)
        audit_logs.append(ExecutionStepLog(
            step_number=step_idx,
            action=AuditAction.BACKUP_CREATED,
            status="SUCCESS",
            description=f"Created SHA-256 verified backup of Master workbook",
            details={"backup_path": backup_path, "sha256": backup_hash}
        ))
        step_idx += 1

        # 2. Stage Working Copy
        working_copy_path = os.path.join(self.temp_dir, f"{plan.plan_id}_working.xlsx")
        shutil.copy2(backup_path, working_copy_path)

        # 3. Resolve target sheet XML path
        mst_inspector = WorkbookInspector(working_copy_path)
        target_xml_path = mst_inspector.get_sheet_target_path(plan.target_sheet_name)

        # 4. Deterministic Execution via OpenXML Patching
        exec_ok, exec_logs, stats = self.writer.execute_plan(
            source_filepath=source_filepath,
            master_working_copy_path=working_copy_path,
            target_sheet_name=plan.target_sheet_name,
            target_sheet_xml_path=target_xml_path,
            source_region=plan.source_region,
            destination_region=plan.destination_region,
            mappings=plan.column_mappings
        )
        audit_logs.extend(exec_logs)
        step_idx += len(exec_logs)

        # 5. Independent Validation
        val_result = self.validator.validate(
            backup_master_path=backup_path,
            processed_master_path=working_copy_path,
            target_sheet_name=plan.target_sheet_name,
            target_sheet_xml_path=target_xml_path,
            expected_destination_region=plan.destination_region
        )

        if not val_result.is_valid:
            # Trigger Rollback
            self.rollback_mgr.rollback(backup_path, working_copy_path, backup_hash)
            audit_logs.append(ExecutionStepLog(
                step_number=step_idx,
                action=AuditAction.ROLLBACK_TRIGGERED,
                status="WARNING",
                description="Validation failed! Triggered automatic rollback to original backup",
                details={"errors": val_result.errors}
            ))
            return False, working_copy_path, val_result, audit_logs

        # 6. Validation Passed: Promote to Processed Storage
        final_processed_path = os.path.join(self.processed_dir, f"updated_{plan.master_filename}")
        shutil.copy2(working_copy_path, final_processed_path)

        audit_logs.append(ExecutionStepLog(
            step_number=step_idx,
            action=AuditAction.TASK_COMPLETE,
            status="SUCCESS",
            description=f"Execution & Validation PASSED. Updated workbook ready for download.",
            details={"output_path": final_processed_path, "stats": stats}
        ))

        return True, final_processed_path, val_result, audit_logs
