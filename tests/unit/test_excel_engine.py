"""Unit & Integration tests for Autonomous Excel Intelligence Platform.

Validates that the system works across different reference workbooks without
hardcoding names, IDs, or column positions.
"""
import os
import pytest
from excel_engine.analyzers.workbook_inspector import WorkbookInspector
from excel_engine.mappers.entity_matcher import EntityMatcher
from excel_engine.mappers.schema_mapper import SchemaMapper
from excel_engine.writers.deterministic_writer import DeterministicWriter
from excel_engine.validators.integrity_validator import IntegrityValidator
from excel_engine.validators.rollback_manager import RollbackManager
from backend.app.services.excel_orchestrator import ExcelOrchestrator
from shared.enums import MatchConfidenceLevel, ColumnRole

CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_ROOT = os.path.abspath(os.path.join(CURRENT_DIR, "../.."))
FIXTURES_DIR = os.path.join(PROJECT_ROOT, "ZEXCEL FILES")
MASTER_FILE = os.path.join(FIXTURES_DIR, "1. Domestic Counselling Team Incentive sheet Oct_26-2.xlsx")


def test_workbook_inspector_metadata_discovery():
    """Verify inspector dynamically extracts metadata across multiple different Sister files."""
    test_files = [
        ("Gagan Godha - APL2004- Oct_2026.xlsx", "APL2004", "Gagan Godha"),
        ("Arshad - APL 1695 - Oct 2026.xlsx", "APL 1695", "Mohd Arshad"),
        ("Chirag Singh- APL2056- Oct_2026.xlsx", "APL2056", "Chirag SIngh"),
        ("Juhi - APL 2046 - Oct 2026.xlsx", "APL 2046", "Juhi"),
    ]

    for fname, exp_id, exp_name in test_files:
        fpath = os.path.join(FIXTURES_DIR, fname)
        inspector = WorkbookInspector(fpath)
        sheets = inspector.get_sheet_names()
        assert len(sheets) >= 1
        meta = inspector.inspect_metadata_block(sheets[0])
        assert meta.entity_id is not None
        assert EntityMatcher.normalize_id(meta.entity_id) == EntityMatcher.normalize_id(exp_id)
        assert meta.entity_name is not None
        assert EntityMatcher.normalize_name(meta.entity_name) == EntityMatcher.normalize_name(exp_name)


def test_entity_matcher_multi_signal():
    """Verify entity matcher accurately finds target sheet for multiple employees."""
    mst_inspector = WorkbookInspector(MASTER_FILE)
    mst_sheets = mst_inspector.get_sheet_names()
    summary_links = mst_inspector.extract_formula_linkages_from_summary("Incentive II Summary Sheet")

    matcher = EntityMatcher()

    test_cases = [
        ("Gagan Godha - APL2004- Oct_2026.xlsx", "Gagan"),
        ("Arshad - APL 1695 - Oct 2026.xlsx", "Mohmmad Arshad RSD"),
        ("Chirag Singh- APL2056- Oct_2026.xlsx", "Chirag"),
        ("Juhi - APL 2046 - Oct 2026.xlsx", "Juhi"),
    ]

    for fname, exp_sheet in test_cases:
        fpath = os.path.join(FIXTURES_DIR, fname)
        src_inspector = WorkbookInspector(fpath)
        meta = src_inspector.inspect_metadata_block(src_inspector.get_sheet_names()[0])

        proposals = matcher.match(meta, mst_sheets, summary_linkages=summary_links)
        assert len(proposals) > 0
        top = proposals[0]
        assert top.candidate_sheet_name == exp_sheet
        assert top.confidence_score >= 0.90
        assert top.confidence_level == MatchConfidenceLevel.HIGH


def test_schema_mapper_protects_formulas():
    """Verify schema mapper correctly aligns columns and protects calculated formulas."""
    src_file = os.path.join(FIXTURES_DIR, "Gagan Godha - APL2004- Oct_2026.xlsx")
    src_inspector = WorkbookInspector(src_file)
    src_region = src_inspector.discover_table_region("Sheet1")
    assert src_region is not None

    mst_inspector = WorkbookInspector(MASTER_FILE)
    dest_region = mst_inspector.discover_table_region("Gagan")
    assert dest_region is not None

    mapper = SchemaMapper()
    mappings = mapper.generate_mapping(src_region, dest_region)
    assert len(mappings) >= 4

    # Ensure no calculated formula columns are mapped as write targets
    dest_formula_indices = {c.index for c in dest_region.columns if c.role == ColumnRole.CALCULATED_FORMULA}
    for m in mappings:
        assert m.destination_index not in dest_formula_indices, f"Mapping violated formula protection on column {m.destination_column}"


def test_rollback_manager():
    """Verify rollback manager restores original byte copy accurately."""
    backup_path = os.path.join(PROJECT_ROOT, "storage/backups", "test_backup.tmp")
    target_path = os.path.join(PROJECT_ROOT, "storage/temporary", "test_target.tmp")
    
    os.makedirs(os.path.join(PROJECT_ROOT, "storage/backups"), exist_ok=True)
    os.makedirs(os.path.join(PROJECT_ROOT, "storage/temporary"), exist_ok=True)

    with open(backup_path, "wb") as f:
        f.write(b"ORIGINAL_BACKUP_CONTENT")

    with open(target_path, "wb") as f:
        f.write(b"CORRUPTED_CONTENT")

    rm = RollbackManager()
    exp_hash = rm.compute_sha256(backup_path)

    ok, msg = rm.rollback(backup_path, target_path, exp_hash)
    assert ok is True
    assert rm.compute_sha256(target_path) == exp_hash

    # Clean up test temp files
    if os.path.exists(backup_path):
        os.remove(backup_path)
    if os.path.exists(target_path):
        os.remove(target_path)
