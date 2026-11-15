"""Generic Semantic Schema & Column Mapper.

Compares arbitrary source columns with existing destination table columns.
Discovers semantic alignments without assuming fixed column letters, column order,
or fixed table dimensions.
Shields formula columns and internal payroll columns from destructive overwrites.
Uses multi-pass matching: exact normalized matches take absolute precedence over
broad fuzzy or category matches.
Enforces mapping uniqueness and flags competing source columns.
"""
import re
from typing import Dict, List, Optional, Set, Tuple
from rapidfuzz import fuzz

from shared.enums import ColumnRole
from shared.schemas import ColumnHeader, ColumnMappingRule, DataRegion


class SchemaMapper:
    """Generates dynamic, semantic source-to-destination column mappings."""

    def __init__(self):
        self.last_unmapped_columns: List[str] = []
        self.last_ambiguities: List[str] = []

    # Explicit semantic roles for tokens to prevent cross-category false matches
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
            "collection_amt", "overall_collection", "collection", "advance_collection"
        },
        "sales_amount": {
            "sale_amt", "sale_amt_invoice_amount", "total_sales", "actual_sale", "sales_amount"
        },
        "list_price": {
            "actual_price", "list_price", "mrp", "standard_price"
        },
        "product": {
            "product_name", "product_as_per_counsllor", "product", "service_name", "course_task"
        },
        "institution": {
            "university_name", "phd_university_name", "institution", "college", "university"
        },
        "payroll_metric": {
            "salary", "salary_on_total_collection", "pay_on", "ctc", "payroll", "payable", "commission_rate"
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

        # Prioritize payroll metric protection
        if any(p in tokens for p in ("salary", "payroll", "payable", "ctc")):
            return "payroll_metric"

        # Prioritize list price protection
        if "actual_price" in norm or "list_price" in norm or "mrp" in tokens:
            return "list_price"

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
        """Align source columns with existing destination columns semantically.
        
        Enforces:
        1. Exact Header Matches (Score 1.0) take absolute priority over broad category/fuzzy matches.
        2. Weak/broad matches are never allowed to pre-empt later exact matches.
        3. Mapping uniqueness: One source column maps to at most one destination; one destination maps to at most one source.
        4. Competing source columns are detected and flagged.
        5. Calculated formula columns and internal payroll columns (Column S) are strictly shielded.
        6. Unmapped source columns are recorded explicitly rather than forced into unrelated destinations.
        """
        mappings: List[ColumnMappingRule] = []
        mapped_dest_indices: Set[int] = set()
        mapped_src_indices: Set[int] = set()
        ambiguities: List[str] = []

        # 1. Filter eligible destination columns:
        # Strictly shield calculated formula columns, totals aggregates, and payroll calculation metrics (e.g. Column S)
        dest_input_cols = [
            c for c in destination_region.columns
            if c.role not in (ColumnRole.CALCULATED_FORMULA, ColumnRole.TOTAL_AGGREGATE)
            and self.get_token_category(c.name) != "payroll_metric"
        ]

        # PASS 1: Exact Normalized Header Matches (Score 1.0)
        for src_col in source_region.columns:
            if src_col.index in mapped_src_indices:
                continue
            for dest_col in dest_input_cols:
                if dest_col.index in mapped_dest_indices:
                    continue
                if src_col.normalized_name == dest_col.normalized_name:
                    mapped_src_indices.add(src_col.index)
                    mapped_dest_indices.add(dest_col.index)
                    mappings.append(ColumnMappingRule(
                        source_column=src_col.name,
                        source_index=src_col.index,
                        destination_column=dest_col.name,
                        destination_index=dest_col.index,
                        destination_column_letter=dest_col.column_letter,
                        transformation_rule="direct_copy",
                        is_formula_column=dest_col.role == ColumnRole.CALCULATED_FORMULA,
                        confidence=1.0
                    ))
                    break

        # PASS 2: Strong Category & Token Identity Matches (Score 0.95 - 0.99)
        candidate_matches: List[Tuple[float, ColumnHeader, ColumnHeader, str]] = []

        for src_col in source_region.columns:
            if src_col.index in mapped_src_indices:
                continue
            src_cat = self.get_token_category(src_col.name)
            # Unmapped categories that shouldn't be forced into tables without matching definitions
            if not src_cat or src_cat in ("list_price", "identifier", "work_order"):
                continue

            for dest_col in dest_input_cols:
                if dest_col.index in mapped_dest_indices:
                    continue
                dest_cat = self.get_token_category(dest_col.name)
                if not dest_cat or src_cat != dest_cat:
                    continue

                # Category match base score
                score = 0.95
                src_tokens = set(filter(None, src_col.normalized_name.split('_')))
                dest_tokens = set(filter(None, dest_col.normalized_name.split('_')))
                if src_tokens == dest_tokens:
                    score = 0.99
                elif src_tokens.issubset(dest_tokens) or dest_tokens.issubset(src_tokens):
                    score = 0.96

                candidate_matches.append((score, src_col, dest_col, "category_token_match"))

        # Sort candidate matches by score descending
        candidate_matches.sort(key=lambda x: x[0], reverse=True)

        # Detect competing source columns for the same destination column
        dest_competition: Dict[int, List[ColumnHeader]] = {}
        for score, src_col, dest_col, _ in candidate_matches:
            dest_competition.setdefault(dest_col.index, []).append(src_col)

        for dest_idx, competitors in dest_competition.items():
            if len(competitors) > 1:
                comp_desc = ", ".join([f"Col {c.index} '{c.name}'" for c in competitors])
                dest_name = next(c.name for c in dest_input_cols if c.index == dest_idx)
                ambiguities.append(
                    f"Multiple source columns ({comp_desc}) compete for destination column '{dest_name}'."
                )

        # Assign best matches uniquely
        mapped_src_norm_names: Set[str] = set()
        for m in mappings:
            mapped_src_norm_names.add(re.sub(r'[^a-z0-9]', '_', m.source_column.lower()).strip('_'))

        for score, src_col, dest_col, match_type in candidate_matches:
            if src_col.index in mapped_src_indices or dest_col.index in mapped_dest_indices:
                continue

            # If another source column with identical normalized name was already mapped in this category,
            # do not force this duplicate source column into an unrelated secondary destination column!
            if src_col.normalized_name in mapped_src_norm_names:
                ambiguities.append(
                    f"Duplicate source field '{src_col.name}' (Col {src_col.index}) held from transfer: competing with already mapped '{src_col.name}' for category '{src_cat}'."
                )
                continue

            mapped_src_indices.add(src_col.index)
            mapped_dest_indices.add(dest_col.index)
            mapped_src_norm_names.add(src_col.normalized_name)

            is_competing = len(dest_competition.get(dest_col.index, [])) > 1
            mappings.append(ColumnMappingRule(
                source_column=src_col.name,
                source_index=src_col.index,
                destination_column=dest_col.name,
                destination_index=dest_col.index,
                destination_column_letter=dest_col.column_letter,
                transformation_rule="direct_copy",
                is_formula_column=dest_col.role == ColumnRole.CALCULATED_FORMULA,
                confidence=round(score, 3),
                is_ambiguous=is_competing,
                ambiguity_reason=f"Competing with other source fields for '{dest_col.name}'" if is_competing else None
            ))

        # PASS 3: High-Confidence Fuzzy Matches (> 0.85)
        for src_col in source_region.columns:
            if src_col.index in mapped_src_indices or src_col.role == ColumnRole.INDEX_SEQUENCE:
                continue
            src_cat = self.get_token_category(src_col.name)
            if src_cat in ("list_price", "identifier", "work_order"):
                continue

            best_dest = None
            best_score = 0.0

            for dest_col in dest_input_cols:
                if dest_col.index in mapped_dest_indices:
                    continue
                dest_cat = self.get_token_category(dest_col.name)
                if src_cat and dest_cat and src_cat != dest_cat:
                    continue

                ratio = fuzz.token_sort_ratio(src_col.normalized_name, dest_col.normalized_name) / 100.0
                if ratio > 0.85 and ratio > best_score:
                    best_score = ratio
                    best_dest = dest_col

            if best_dest:
                mapped_src_indices.add(src_col.index)
                mapped_dest_indices.add(best_dest.index)
                mappings.append(ColumnMappingRule(
                    source_column=src_col.name,
                    source_index=src_col.index,
                    destination_column=best_dest.name,
                    destination_index=best_dest.index,
                    destination_column_letter=best_dest.column_letter,
                    transformation_rule="direct_copy",
                    is_formula_column=best_dest.role == ColumnRole.CALCULATED_FORMULA,
                    confidence=round(best_score, 3)
                ))

        # Identify unmapped source columns
        unmapped_sources = [
            c.name for c in source_region.columns
            if c.index not in mapped_src_indices and c.role != ColumnRole.INDEX_SEQUENCE
        ]

        self.last_unmapped_columns = unmapped_sources
        self.last_ambiguities = ambiguities

        # Sort mappings by destination column index
        mappings.sort(key=lambda m: m.destination_index)
        return mappings
