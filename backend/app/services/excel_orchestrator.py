"""Generic Excel Orchestrator Service.

Coordinates the full autonomous pipeline across inspection, matching, mapping,
safe execution, validation, and rollback without hardcoded logic.
"""
import os
import shutil
import uuid
from typing import Any, Dict, List, Optional, Tuple

from excel_engine.analyzers.workbook_inspector import WorkbookInspector
from excel_engine.mappers.entity_matcher import EntityMatcher
from excel_engine.mappers.schema_mapper import SchemaMapper
from excel_engine.validators.integrity_validator import IntegrityValidator
from excel_engine.validators.rollback_manager import RollbackManager
from excel_engine.writers.deterministic_writer import DeterministicWriter
from shared.enums import AuditAction, ExecutionMode, MatchConfidenceLevel, TaskStatus
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
        summary_sheet_name: Optional[str] = "Incentive II Summary Sheet"
    ) -> Tuple[MappingPlan, EntityMatchProposal, List[ExecutionStepLog]]:
        """Stage 1: Discover structure, match entity, and formulate mapping plan."""
        audit_logs: List[ExecutionStepLog] = []
        task_id = str(uuid.uuid4())[:8]

        # Step 1: Inspect Source File
        src_inspector = WorkbookInspector(source_filepath)
        src_sheets = src_inspector.get_sheet_names()
        primary_src_sheet = src_sheets[0]

        src_meta = src_inspector.inspect_metadata_block(primary_src_sheet)
        src_region = src_inspector.discover_table_region(primary_src_sheet)

        audit_logs.append(ExecutionStepLog(
            step_number=1,
            action=AuditAction.STRUCTURE_INSPECT,
            status="SUCCESS",
            description=f"Inspected source file. Identified entity: ID={src_meta.entity_id}, Name={src_meta.entity_name}",
            details={"metadata": src_meta.model_dump(), "source_sheet": primary_src_sheet}
        ))

        # Step 2: Inspect Master Workbook
        mst_inspector = WorkbookInspector(master_filepath)
        mst_sheets = mst_inspector.get_sheet_names()

        # Extract Summary linkages if summary sheet is available
        linkages = []
        if summary_sheet_name and summary_sheet_name in mst_sheets:
            linkages = mst_inspector.extract_formula_linkages_from_summary(summary_sheet_name)

        # Match entity to destination sheet
        match_proposals = self.matcher.match(
            source_entity=src_meta,
            candidate_sheets=mst_sheets,
            summary_linkages=linkages
        )

        best_proposal = match_proposals[0]
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

        # Step 3: Inspect Destination Sheet Structure
        dest_region = mst_inspector.discover_table_region(best_proposal.candidate_sheet_name)
        if not dest_region:
            raise ValueError(f"Could not discover tabular data region in destination sheet '{best_proposal.candidate_sheet_name}'")

        # Step 4: Generate Semantic Schema Mapping
        mappings = self.schema_mapper.generate_mapping(src_region, dest_region)
        audit_logs.append(ExecutionStepLog(
            step_number=3,
            action=AuditAction.SCHEMA_MAP,
            status="SUCCESS",
            description=f"Generated {len(mappings)} semantic column mappings",
            details={"mapped_columns": [m.model_dump() for m in mappings]}
        ))

        plan = MappingPlan(
            plan_id=task_id,
            source_filename=os.path.basename(source_filepath),
            master_filename=os.path.basename(master_filepath),
            target_sheet_name=best_proposal.candidate_sheet_name,
            source_region=src_region,
            destination_region=dest_region,
            column_mappings=mappings,
            rows_to_transfer=src_region.data_end_row - src_region.data_start_row + 1,
            target_start_row=dest_region.data_start_row,
            requires_approval=best_proposal.requires_human_approval or (self.execution_mode != ExecutionMode.AUTONOMOUS),
            status=TaskStatus.AWAITING_APPROVAL if best_proposal.requires_human_approval else TaskStatus.PENDING
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
