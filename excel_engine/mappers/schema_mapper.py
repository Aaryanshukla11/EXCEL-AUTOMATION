"""Generic Semantic Schema & Column Mapper.

Compares arbitrary source columns with existing destination table columns.
Discovers semantic alignments without assuming fixed column letters, column order,
or fixed table dimensions.
Shields formula columns from destructive overwrites.
Uses full token matching to eliminate substring false positives.
"""
import re
from typing import Dict, List, Optional, Set
from rapidfuzz import fuzz

from shared.enums import ColumnRole
from shared.schemas import ColumnHeader, ColumnMappingRule, DataRegion


class SchemaMapper:
    """Generates dynamic, semantic source-to-destination column mappings."""

    # Explicit semantic roles for tokens to prevent cross-category false matches
    # (e.g., preventing 'contact_id' from matching 'sales_amount_consider_for_incentive')
    SEMANTIC_CATEGORIES = {
        "identifier": {
            "contact_id", "customer_id", "client_id", "student_id", "id", "code", "emp_id", "roll_no"
        },
        "entity_name": {
            "contact_name", "customer_name", "candidate", "candidate_name", "client_name", "student_name", "name"
        },
        "date": {
            "payment_date", "date", "invoice_date", "trans_date", "txn_date", "collection_date"
        },
        "invoice_ref": {
            "inv_no", "pi_no", "invoice_number", "invoice_no", "pi_number", "bill_no"
        },
        "work_order": {
            "work_order_id", "wo_id", "work_order", "order_id"
        },
        "collection_amount": {
            "collection_amt", "overall_collection", "collection", "total_collection", "achieved_advance_collection"
        },
        "sales_amount": {
            "sale_amt", "sale_amt_invoice_amount", "total_sales", "actual_sale", "sales_amount", "actual_price"
        },
        "product": {
            "product_name", "product_as_per_counsllor", "product", "service_name", "course_task"
        },
        "institution": {
            "university_name", "phd_university_name", "institution", "college", "university"
        },
        "remarks": {
            "remarks", "payroll_remark", "audit_remarks", "comments", "notes"
        }
    }

    @classmethod
    def get_token_category(cls, col_name: str) -> Optional[str]:
        """Determine semantic category using full token intersection to prevent substring bugs."""
        norm = re.sub(r'[^a-z0-9]', '_', col_name.lower()).strip('_')
        tokens = set(filter(None, norm.split('_')))

        for category, keywords in cls.SEMANTIC_CATEGORIES.items():
            # Exact match of normalized name
            if norm in keywords:
                return category
            # Multi-word match
            for kw in keywords:
                kw_tokens = set(kw.split('_'))
                if kw_tokens.issubset(tokens):
                    return category

        return None

    @classmethod
    def are_synonyms(cls, name_a: str, name_b: str) -> bool:
        """Check if two column names share the same semantic category."""
        cat_a = cls.get_token_category(name_a)
        cat_b = cls.get_token_category(name_b)
        if cat_a and cat_b and cat_a == cat_b:
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
            # Skip pure index columns (e.g. S.No) as destination generates or preserves its own index
            if src_col.role == ColumnRole.INDEX_SEQUENCE:
                continue

            best_match: Optional[ColumnHeader] = None
            best_score = 0.0

            src_cat = self.get_token_category(src_col.name)

            for dest_col in dest_input_cols:
                if dest_col.index in mapped_dest_indices:
                    continue

                dest_cat = self.get_token_category(dest_col.name)

                # 1. Exact normalized match
                if src_col.normalized_name == dest_col.normalized_name:
                    best_match = dest_col
                    best_score = 1.0
                    break

                # 2. Strict Semantic category match (prevents ID matching Amount or Date)
                if src_cat and dest_cat and src_cat == dest_cat:
                    score = 0.95
                    if score > best_score:
                        best_score = score
                        best_match = dest_col
                        continue

                # 3. Fuzzy string matching only if categories don't conflict
                if not src_cat or not dest_cat:
                    ratio = fuzz.token_sort_ratio(src_col.normalized_name, dest_col.normalized_name) / 100.0
                    if ratio > 0.80 and ratio > best_score:
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
