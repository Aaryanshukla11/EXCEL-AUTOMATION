"""Shared Pydantic schemas for Autonomous Excel Intelligence Platform."""
from datetime import datetime, timezone
from typing import Any, Dict, List, Optional
from pydantic import BaseModel, Field

from shared.enums import (
    AuditAction,
    BatchStatus,
    ColumnRole,
    ExecutionMode,
    MatchConfidenceLevel,
    TaskStatus,
    TransactionStatus,
)


class CellCoordinate(BaseModel):
    sheet: str
    row: int
    column: int
    column_letter: str

    @property
    def ref(self) -> str:
        return f"{self.column_letter}{self.row}"


class ColumnHeader(BaseModel):
    index: int
    column_letter: str
    name: str
    normalized_name: str
    role: ColumnRole = ColumnRole.INPUT_VALUE
    sample_formula: Optional[str] = None
    has_formula_slots: bool = False
    section: Optional[str] = None
    is_verified_section: bool = False


class DataRegion(BaseModel):
    sheet_name: str
    header_row: int
    data_start_row: int
    data_end_row: int
    totals_row: Optional[int] = None
    columns: List[ColumnHeader]
    preallocated_slots_count: int = 0
    available_empty_slots: List[int] = Field(default_factory=list)


class EntityMetadata(BaseModel):
    entity_id: Optional[str] = None
    entity_name: Optional[str] = None
    supervisor_name: Optional[str] = None
    additional_attributes: Dict[str, Any] = Field(default_factory=dict)
    raw_source_cells: Dict[str, str] = Field(default_factory=dict)


class EntityMatchProposal(BaseModel):
    source_entity: EntityMetadata
    candidate_sheet_name: str
    confidence_score: float
    confidence_level: MatchConfidenceLevel
    score_breakdown: Dict[str, float]
    matching_evidence: List[str]
    requires_human_approval: bool


class ColumnMappingRule(BaseModel):
    source_column: str
    source_index: int
    destination_column: str
    destination_index: int
    destination_column_letter: str
    transformation_rule: Optional[str] = None
    is_formula_column: bool = False
    confidence: float = 1.0
    is_ambiguous: bool = False
    ambiguity_reason: Optional[str] = None
    section_context: Optional[str] = None
    source_cell_references: Optional[str] = None
    is_verified_source: bool = False


class ExecutionStepLog(BaseModel):
    step_number: int
    action: AuditAction
    status: str  # "SUCCESS", "WARNING", "FAILED"
    description: str
    details: Dict[str, Any] = Field(default_factory=dict)
    timestamp: datetime = Field(default_factory=lambda: datetime.now(timezone.utc))


class RowTransactionRecord(BaseModel):
    source_row_index: int
    entity_identity: str
    status: TransactionStatus
    status_reason: Optional[str] = None
    verified_fields_present: Dict[str, Any] = Field(default_factory=dict)
    missing_verified_fields: List[str] = Field(default_factory=list)
    proposed_destination_row: Optional[int] = None
    proposed_cell_values: Dict[str, Any] = Field(default_factory=dict)
    normalized_product: Optional[str] = None
    claimed_fields_observed: Dict[str, Any] = Field(default_factory=dict)


class BatchApprovalSummary(BaseModel):
    batch_status: BatchStatus
    total_rows: int
    eligible_rows_count: int
    held_rows_count: int
    needs_approval_count: int = 0
    rejected_count: int = 0
    partial_batch_warning: bool = False
    requires_explicit_approval: bool = False
    row_records: List[RowTransactionRecord] = Field(default_factory=list)


class MappingPlan(BaseModel):
    plan_id: str
    source_filename: str
    master_filename: str
    target_sheet_name: str
    source_region: DataRegion
    destination_region: DataRegion
    column_mappings: List[ColumnMappingRule]
    rows_to_transfer: int
    target_start_row: int
    requires_approval: bool
    status: TaskStatus = TaskStatus.PENDING
    unmapped_source_columns: List[str] = Field(default_factory=list)
    ambiguities: List[str] = Field(default_factory=list)
    product_normalization_details: Dict[str, Any] = Field(default_factory=dict)
    mirror_columns_observation: Dict[str, str] = Field(default_factory=dict)
    batch_summary: Optional[BatchApprovalSummary] = None



class ValidationResult(BaseModel):
    is_valid: bool
    total_checks: int
    passed_checks: int
    failed_checks: int
    warnings: List[str] = Field(default_factory=list)
    errors: List[str] = Field(default_factory=list)
    cell_diff_count: int = 0
    formula_integrity_passed: bool = True
