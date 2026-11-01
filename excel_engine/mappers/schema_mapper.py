"""Generic Semantic Schema & Column Mapper.

Compares arbitrary source columns with existing destination table columns.
Discovers semantic alignments without assuming fixed column letters, column order,
or fixed table dimensions.
Shields formula columns from destructive overwrites.
"""
import re
from typing import Dict, List, Optional, Set
from rapidfuzz import fuzz

from shared.enums import ColumnRole
from shared.schemas import ColumnHeader, ColumnMappingRule, DataRegion


class SchemaMapper:
    """Generates dynamic, semantic source-to-destination column mappings."""

    # Generic semantic synonym clusters across business domains
    SYNONYM_GROUPS = [
        # Customer / Contact
        {"contact_name", "customer_name", "candidate", "client_name", "student_name", "name", "candidate_name"},
        # Date
        {"payment_date", "date", "invoice_date", "trans_date", "txn_date", "collection_date"},
        # Identifier
        {"contact_id", "customer_id", "client_id", "student_id", "emp_id", "id"},
        # Work Order
        {"work_order_id", "wo_id", "work_order_status", "work_order", "wo_status"},
        # Invoice
        {"inv_no", "pi_no", "invoice_number", "invoice_no", "pi_number", "bill_no"},
        # Collection Amount
        {"collection_amt", "overall_collection", "collection", "total_collection", "achieved_advance_collection", "collec_amt"},
        # Sales Amount
        {"sale_amt", "sale_amt_invoice_amount", "total_sales", "actual_sale", "sales_amount", "booked_sale"},
        # Product Name
        {"product_name", "product_as_per_counsllor", "product", "service_name", "course_task"},
        # University / Institution
        {"university_name", "phd_university_name", "institution", "college", "university"},
        # Remarks
        {"remarks", "payroll_remark", "audit_remarks", "comments", "notes"},
    ]

    @classmethod
    def are_synonyms(cls, name_a: str, name_b: str) -> bool:
        """Check if two normalized column names belong to the same synonym cluster."""
        norm_a = re.sub(r'[^a-z0-9]', '_', name_a.lower()).strip('_')
        norm_b = re.sub(r'[^a-z0-9]', '_', name_b.lower()).strip('_')
        if norm_a == norm_b:
            return True
        for group in cls.SYNONYM_GROUPS:
            if any(syn in norm_a for syn in group) and any(syn in norm_b for syn in group):
                return True
        return False

    def generate_mapping(
        self,
        source_region: DataRegion,
        destination_region: DataRegion
    ) -> List[ColumnMappingRule]:
        """Align source columns with existing destination columns semantically."""
        mappings: List[ColumnMappingRule] = []
        mapped_dest_indices: Set[int] = set()

        # Only allow mapping into destination columns that are NOT calculated formulas
        dest_input_cols = [c for c in destination_region.columns if c.role != ColumnRole.CALCULATED_FORMULA]

        for src_col in source_region.columns:
            # Skip pure index columns (e.g. S.No) as destination will generate or keep its own index
            if src_col.role == ColumnRole.INDEX_SEQUENCE:
                continue

            best_match: Optional[ColumnHeader] = None
            best_score = 0.0

            for dest_col in dest_input_cols:
                if dest_col.index in mapped_dest_indices:
                    continue

                # 1. Check exact normalized match
                if src_col.normalized_name == dest_col.normalized_name:
                    best_match = dest_col
                    best_score = 1.0
                    break

                # 2. Check synonym cluster
                if self.are_synonyms(src_col.name, dest_col.name):
                    score = 0.95
                    if score > best_score:
                        best_score = score
                        best_match = dest_col
                        continue

                # 3. Fuzzy string matching
                ratio = fuzz.token_sort_ratio(src_col.normalized_name, dest_col.normalized_name) / 100.0
                if ratio > 0.75 and ratio > best_score:
                    best_score = ratio
                    best_match = dest_col

            if best_match and best_score >= 0.70:
                mapped_dest_indices.add(best_match.index)
                mappings.append(ColumnMappingRule(
                    source_column=src_col.name,
                    source_index=src_col.index,
                    destination_column=best_match.name,
                    destination_index=best_match.index,
                    destination_column_letter=best_match.column_letter,
                    transformation_rule="direct_copy",
                    is_formula_column=best_match.role == ColumnRole.CALCULATED_FORMULA,
                    confidence=round(best_score, 3)
                ))

        return mappings
