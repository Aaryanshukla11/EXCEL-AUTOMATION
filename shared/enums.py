"""Shared enums for Autonomous Excel Intelligence Platform."""
from enum import Enum


class ExecutionMode(str, Enum):
    LEARNING = "learning"
    SUPERVISED = "supervised"
    AUTONOMOUS = "autonomous"


class TaskStatus(str, Enum):
    PENDING = "pending"
    ANALYZING = "analyzing"
    MAPPING = "mapping"
    AWAITING_APPROVAL = "awaiting_approval"
    EXECUTING = "executing"
    COMPLETED = "completed"
    FAILED = "failed"
    ROLLED_BACK = "rolled_back"


class MatchConfidenceLevel(str, Enum):
    HIGH = "high"        # >= 0.90
    MEDIUM = "medium"    # >= 0.70 and < 0.90
    LOW = "low"          # < 0.70
    UNKNOWN = "unknown"


class ColumnRole(str, Enum):
    INPUT_VALUE = "input_value"
    CALCULATED_FORMULA = "calculated_formula"
    INDEX_SEQUENCE = "index_sequence"
    METADATA_HEADER = "metadata_header"
    TOTAL_AGGREGATE = "total_aggregate"


class AuditAction(str, Enum):
    FILE_UPLOAD = "file_upload"
    STRUCTURE_INSPECT = "structure_inspect"
    ENTITY_MATCH = "entity_match"
    SCHEMA_MAP = "schema_map"
    USER_APPROVAL = "user_approval"
    USER_CORRECTION = "user_correction"
    BACKUP_CREATED = "backup_created"
    CELL_WRITE = "cell_write"
    FORMULA_VALIDATION = "formula_validation"
    INTEGRITY_CHECK = "integrity_check"
    TASK_COMPLETE = "task_complete"
    ROLLBACK_TRIGGERED = "rollback_triggered"
    ERROR_DETECTED = "error_detected"

