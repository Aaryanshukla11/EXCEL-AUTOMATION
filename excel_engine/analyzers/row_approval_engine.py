"""Row-Level Approval and Granular Eligibility Engine.

Evaluates each transaction row independently against verified data rules,
destination slot constraints, and product normalization.
Enforces zero silent fallback from verified to claimed values,
partial batch warning signals, and granular approval control.
"""
from typing import Any, Dict, List, Optional, Tuple
import openpyxl

from shared.enums import (
    BatchStatus,
    ColumnRole,
    TransactionStatus,
)
from shared.schemas import (
    BatchApprovalSummary,
    ColumnMappingRule,
    DataRegion,
    RowTransactionRecord,
)
from excel_engine.mappers.product_normalizer import ProductNormalizer


class RowApprovalEngine:
    """Evaluates transaction row eligibility and manages partial batch approvals."""

    def __init__(self, product_normalizer: Optional[ProductNormalizer] = None):
        self.normalizer = product_normalizer

    def evaluate_batch(
        self,
        source_filepath: str,
        source_region: DataRegion,
        destination_region: Optional[DataRegion],
        column_mappings: List[ColumnMappingRule],
        target_sheet_name: Optional[str] = None,
        master_sheet_exists: bool = True,
    ) -> BatchApprovalSummary:
        """Evaluates every transaction row in the source data region.
        
        Args:
            source_filepath: Path to source Sister workbook.
            source_region: DataRegion detected in source sheet.
            destination_region: DataRegion detected in destination Master sheet.
            column_mappings: Resolved column mapping rules.
            target_sheet_name: Target worksheet name in Master.
            master_sheet_exists: True if target sheet was found in Master.
            
        Returns:
            BatchApprovalSummary with granular row-by-row decisions.
        """
        row_records: List[RowTransactionRecord] = []

        # If destination worksheet does not exist in Master (e.g. Kanupriya case)
        if not master_sheet_exists or destination_region is None or not target_sheet_name:
            # All rows in source must be held because there is no destination
            wb = openpyxl.load_workbook(source_filepath, data_only=True)
            ws = wb[source_region.sheet_name]
            for r in range(source_region.data_start_row, source_region.data_end_row + 1):
                identity = self._extract_row_identity(ws, r, source_region)
                if not identity:
                    continue
                row_records.append(RowTransactionRecord(
                    source_row_index=r,
                    entity_identity=identity,
                    status=TransactionStatus.HOLD,
                    status_reason=f"Destination worksheet '{target_sheet_name or 'UNKNOWN'}' does not exist in Master workbook. Entire batch held.",
                    missing_verified_fields=["DESTINATION_SHEET_MISSING"],
                    proposed_destination_row=None,
                    proposed_cell_values={},
                ))
            wb.close()

            return BatchApprovalSummary(
                batch_status=BatchStatus.ALL_HOLD,
                total_rows=len(row_records),
                eligible_rows_count=0,
                held_rows_count=len(row_records),
                needs_approval_count=0,
                rejected_count=0,
                partial_batch_warning=False,
                requires_explicit_approval=True,
                row_records=row_records,
            )

        # Check if source contains verified section
        verified_cols = [
            c for c in source_region.columns if c.is_verified_section
        ]
        has_verified_section = len(verified_cols) > 0

        # Required verified financial fields when verified section exists
        verified_mappings = [
            m for m in column_mappings if m.is_verified_source
        ]

        # Available destination empty slots
        available_slots = list(destination_region.available_empty_slots) if destination_region else []
        slot_pointer = 0

        wb = openpyxl.load_workbook(source_filepath, data_only=True)
        ws = wb[source_region.sheet_name]

        for r in range(source_region.data_start_row, source_region.data_end_row + 1):
            identity = self._extract_row_identity(ws, r, source_region)
            # If the row has no identity and no financial data, it's an empty template row
            if not identity and self._is_row_blank(ws, r, source_region):
                continue

            row_id_text = identity or f"Row {r}"

            # Case A: 13-column or unverified file (claimed only)
            if not has_verified_section or not verified_mappings:
                claimed_vals = self._extract_claimed_values(ws, r, source_region)
                row_records.append(RowTransactionRecord(
                    source_row_index=r,
                    entity_identity=row_id_text,
                    status=TransactionStatus.NEEDS_APPROVAL,
                    status_reason="Source file contains claimed figures only. Verification section absent. Requires managerial override.",
                    verified_fields_present={},
                    missing_verified_fields=["ACCOUNTS_VERIFIED_SECTION"],
                    proposed_destination_row=None,
                    proposed_cell_values={},
                    claimed_fields_observed=claimed_vals,
                ))
                continue

            # Case B: Verified section exists (e.g. 15 or 16-column file)
            verified_present: Dict[str, Any] = {}
            missing_verified: List[str] = []
            claimed_vals = self._extract_claimed_values(ws, r, source_region)

            for vm in verified_mappings:
                cell_val = ws.cell(r, vm.source_index).value
                cell_ref = f"{openpyxl.utils.get_column_letter(vm.source_index)}{r}"
                if cell_val is None or str(cell_val).strip() == "":
                    missing_verified.append(f"{vm.source_column} ({cell_ref})")
                else:
                    verified_present[vm.source_column] = cell_val

            # Also check any remarks in the verified section for context
            remark_text = self._extract_verified_remark(ws, r, source_region)

            # Evaluate eligibility
            if missing_verified:
                reason = f"Accounts Verified field(s) {', '.join(missing_verified)} missing or blank."
                if remark_text:
                    reason += f" Auditor Remark: '{remark_text}'."
                reason += " Held per Rule 1 (never silently falling back to claimed data)."

                row_records.append(RowTransactionRecord(
                    source_row_index=r,
                    entity_identity=row_id_text,
                    status=TransactionStatus.HOLD,
                    status_reason=reason,
                    verified_fields_present=verified_present,
                    missing_verified_fields=missing_verified,
                    proposed_destination_row=None,
                    proposed_cell_values={},
                    claimed_fields_observed=claimed_vals,
                ))
            else:
                # All required verified fields are present -> check slot availability
                if slot_pointer < len(available_slots):
                    dest_row = available_slots[slot_pointer]
                    slot_pointer += 1

                    # Build proposed cell values
                    proposed_cells, norm_prod = self._build_proposed_cells(
                        ws=ws,
                        source_row=r,
                        dest_row=dest_row,
                        column_mappings=column_mappings,
                        destination_region=destination_region,
                    )

                    row_records.append(RowTransactionRecord(
                        source_row_index=r,
                        entity_identity=row_id_text,
                        status=TransactionStatus.ELIGIBLE,
                        status_reason=None,
                        verified_fields_present=verified_present,
                        missing_verified_fields=[],
                        proposed_destination_row=dest_row,
                        proposed_cell_values=proposed_cells,
                        normalized_product=norm_prod,
                        claimed_fields_observed=claimed_vals,
                    ))
                else:
                    # Slot exhausted! Cannot append per Rule 5
                    row_records.append(RowTransactionRecord(
                        source_row_index=r,
                        entity_identity=row_id_text,
                        status=TransactionStatus.HOLD,
                        status_reason="Pre-allocated destination empty slots exhausted. Automatic row append prohibited per Rule 5.",
                        verified_fields_present=verified_present,
                        missing_verified_fields=["DESTINATION_SLOT_EXHAUSTED"],
                        proposed_destination_row=None,
                        proposed_cell_values={},
                        claimed_fields_observed=claimed_vals,
                    ))

        wb.close()

        # Batch Aggregation
        total = len(row_records)
        eligible = sum(1 for rec in row_records if rec.status == TransactionStatus.ELIGIBLE)
        held = sum(1 for rec in row_records if rec.status == TransactionStatus.HOLD)
        needs_appr = sum(1 for rec in row_records if rec.status == TransactionStatus.NEEDS_APPROVAL)
        rejected = sum(1 for rec in row_records if rec.status == TransactionStatus.REJECTED)

        if held > 0 and eligible > 0:
            batch_status = BatchStatus.PARTIAL_ELIGIBLE
            partial_warning = True
            req_approval = True
        elif held > 0 and eligible == 0:
            batch_status = BatchStatus.ALL_HOLD
            partial_warning = False
            req_approval = True
        elif needs_appr > 0:
            batch_status = BatchStatus.REQUIRES_SELECTION
            partial_warning = False
            req_approval = True
        elif eligible > 0 and held == 0:
            batch_status = BatchStatus.ALL_ELIGIBLE
            partial_warning = False
            req_approval = False
        else:
            batch_status = BatchStatus.ALL_HOLD
            partial_warning = False
            req_approval = True

        return BatchApprovalSummary(
            batch_status=batch_status,
            total_rows=total,
            eligible_rows_count=eligible,
            held_rows_count=held,
            needs_approval_count=needs_appr,
            rejected_count=rejected,
            partial_batch_warning=partial_warning,
            requires_explicit_approval=req_approval,
            row_records=row_records,
        )

    def _extract_row_identity(self, ws, row_idx: int, region: DataRegion) -> Optional[str]:
        """Dynamically finds entity identity (e.g. Contact Name or S.No) in row."""
        # 1. Look for Name column
        for col in region.columns:
            if "name" in col.normalized_name and not col.is_verified_section:
                val = ws.cell(row_idx, col.index).value
                if val and str(val).strip() and str(val).strip().upper() not in ("NA", "NONE", "NO"):
                    return str(val).strip()

        # 2. Look for ID or Serial column
        for col in region.columns:
            if any(k in col.normalized_name for k in ("id", "s_no", "sno", "sr_no")):
                val = ws.cell(row_idx, col.index).value
                if val is not None and str(val).strip():
                    return f"{col.name}: {str(val).strip()}"

        return None

    def _is_row_blank(self, ws, row_idx: int, region: DataRegion) -> bool:
        """Checks if row has no transaction data (ignoring s_no)."""
        for col in region.columns:
            if col.role == ColumnRole.INDEX_SEQUENCE:
                continue
            val = ws.cell(row_idx, col.index).value
            if val is not None and str(val).strip() not in ("", "None", "NA", "No"):
                return False
        return True

    def _extract_claimed_values(self, ws, row_idx: int, region: DataRegion) -> Dict[str, Any]:
        """Extracts claimed section figures for observation."""
        claimed: Dict[str, Any] = {}
        for col in region.columns:
            if not col.is_verified_section and col.role == ColumnRole.INPUT_VALUE:
                val = ws.cell(row_idx, col.index).value
                if val is not None and str(val).strip():
                    claimed[col.name] = val
        return claimed

    def _extract_verified_remark(self, ws, row_idx: int, region: DataRegion) -> Optional[str]:
        """Finds remarks within verified section."""
        for col in region.columns:
            if col.is_verified_section and "remark" in col.normalized_name:
                val = ws.cell(row_idx, col.index).value
                if val and str(val).strip():
                    return str(val).strip()
        return None

    def _build_proposed_cells(
        self,
        ws,
        source_row: int,
        dest_row: int,
        column_mappings: List[ColumnMappingRule],
        destination_region: DataRegion,
    ) -> Tuple[Dict[str, Any], Optional[str]]:
        """Constructs destination cell coordinates and values for an eligible row."""
        cells: Dict[str, Any] = {}
        norm_product: Optional[str] = None

        for rule in column_mappings:
            dest_col_letter = rule.destination_column_letter
            target_cell_ref = f"{dest_col_letter}{dest_row}"

            # Rule 3: Never write to mirror columns T or U
            if dest_col_letter in ("T", "U"):
                continue

            # Check if destination is a calculated formula
            dest_col_obj = next((c for c in destination_region.columns if c.column_letter == dest_col_letter), None)
            if dest_col_obj and dest_col_obj.role == ColumnRole.CALCULATED_FORMULA:
                continue

            raw_val = ws.cell(source_row, rule.source_index).value

            # Product Normalization logic (Rules 2A & 2B)
            if rule.transformation_rule == "preserve_raw_text":
                cells[target_cell_ref] = raw_val
            elif rule.transformation_rule == "lookup_normalize":
                if self.normalizer and raw_val:
                    res = self.normalizer.normalize(str(raw_val))
                    cells[target_cell_ref] = res.normalized_category
                    norm_product = res.normalized_category
                else:
                    cells[target_cell_ref] = raw_val
                    norm_product = str(raw_val) if raw_val else None
            else:
                cells[target_cell_ref] = raw_val

        return cells, norm_product
