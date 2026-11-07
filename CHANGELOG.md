# CHANGELOG

All notable changes to the Autonomous Excel Intelligence Platform will be documented in this file.
The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [0.2.0] - 2026-11-18
### Phase 2 Implementation Audit & Expanded Engine Verification
- Audited all 7 core modules (`workbook_inspector.py`, `entity_matcher.py`, `schema_mapper.py`, `deterministic_writer.py`, `integrity_validator.py`, `rollback_manager.py`, `excel_orchestrator.py`).
- Eliminated hardcoded summary sheet default (`Incentive II Summary Sheet`), replacing it with dynamic summary/roster discovery.
- Eliminated substring matching defect in `schema_mapper.py` where `'id'` in `'consider'` erroneously mapped `CONTACT ID` to `sales_amount_consider_for_incentive`.
- Fixed metadata regex boundary matching for punctuation/spaces (`EMP. ID` and `EMP. NAME`).
- Implemented `discover_primary_data_sheet` to score and distinguish primary claim sheets from secondary auxiliary logs in multi-sheet files (verified on `Krishika Kataria`).
- Handled missing destination worksheet scenario (`APL Kanupriya - APL 2111`) by refusing automatic write and properly escalating to human decision.
- Verified byte-level integrity across all 1,589 parts of the Master workbook (all 237 pivot parts 100% byte-for-byte identical, 145/146 sheets completely untouched).
- Added 7 expanded scenario test cases bringing total automated test coverage to 12 passing tests.

## [0.1.0] - 2026-11-18
### Initial Scaffold & Architecture
- Initialized fixed project structure adhering to strict software engineering standards.
- Deep binary structural inspection of real-world test fixtures.
- Created `ARCHITECTURE.md` and initial baseline test fixtures.
