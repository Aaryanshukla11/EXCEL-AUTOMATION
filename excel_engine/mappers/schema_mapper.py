"""Generic Semantic Schema & Column Mapper with Rule 1-4 Business Logic.

Compares arbitrary source columns with existing destination table columns.
Enforces:
1. ACCOUNTS VERIFIED PRIORITY (Rule 1):
   - Dynamically distinguishes Counselor Claimed vs Accounts Verified sections.
   - When verified equivalents exist, verified columns are mapped to Master transaction fields.
   - Claimed columns are not silently substituted and are recorded in unmapped sources with evidence.
   - Preserves source cell references and section context in mapping rules.
2. DYNAMIC PRODUCT NORMALIZATION (Rule 2):
   - Preserves original raw product text in Master Column D.
   - Resolves normalized product category for Master Column E against dynamically discovered lookup grid.
   - Disallows inventing categories; flags ambiguous/unknown products for human approval.
3. MIRROR COLUMNS PROTECTION (Rule 3):
   - Shields columns T & U from incoming source assignments.
   - Preserves existing formulas or static values in T & U.
4. COLUMN S PROTECTION:
   - Shields Column S (Salary On Total Collection) from receiving transaction collection amounts.
5. EXACT HEADER PRIORITY & MAPPING UNIQUENESS (Fixes 3 & 4):
   - Exact normalized header matches take precedence over broad category/fuzzy matches.
   - 1-to-1 uniqueness enforced; competing source fields detected and flagged.
"""
import re
from typing import Dict, List, Optional, Set, Tuple
from rapidfuzz import fuzz

from excel_engine.mappers.product_normalizer import ProductNormalizer
from shared.enums import ColumnRole
from shared.schemas import ColumnHeader, ColumnMappingRule, DataRegion


class SchemaMapper:
    """Generates dynamic, semantic source-to-destination column mappings."""

    def __init__(self, product_grid: Optional[Dict[str, float]] = None):
        self.product_grid = product_grid or {}
        self.product_normalizer = ProductNormalizer(self.product_grid) if self.product_grid else None
        self.last_unmapped_columns: List[str] = []
        self.last_ambiguities: List[str] = []
        self.last_product_normalization_details: Dict[str, Any] = {}

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

        # Prioritize product and institution to prevent falling back to generic 'name'
        if "product" in tokens:
            return "product"
        if any(inst in tokens for inst in ("university", "college", "institution")):
            return "institution"

        for category, keywords in cls.SEMANTIC_CATEGORIES.items():
            if norm in keywords:
                return category
            for kw in keywords:
                kw_tokens = set(kw.split('_'))
                if kw_tokens.issubset(tokens):
                    return category

        return None

    def generate_mapping(
        self,
        source_region: DataRegion,
        destination_region: DataRegion,
        sample_source_product: Optional[str] = None
    ) -> List[ColumnMappingRule]:
        """Align source columns with existing destination columns semantically adhering to Rules 1-4."""
        mappings: List[ColumnMappingRule] = []
        mapped_dest_indices: Set[int] = set()
        mapped_src_indices: Set[int] = set()
        ambiguities: List[str] = []
        unmapped_reasons: Dict[str, str] = {}

        # Discover section context in source
        has_verified_section = any(c.is_verified_section for c in source_region.columns)

        # 1. Filter eligible destination columns:
        # Strictly shield:
        # - Calculated formula columns and totals aggregates
        # - Payroll calculation metrics (Column S Salary On Total Collection)
        # - Mirror columns T & U (Actual Sale, Actual Collection) per Rule 3
        # - Auxiliary columns W & X (Sale, Collection) to avoid unneeded duplicate dumps
        dest_input_cols = [
            c for c in destination_region.columns
            if c.role not in (ColumnRole.CALCULATED_FORMULA, ColumnRole.TOTAL_AGGREGATE)
            and self.get_token_category(c.name) != "payroll_metric"
            and not (c.column_letter in ("T", "U") and any(k in c.normalized_name for k in ("actual_sale", "actual_collection")))
            and not (c.column_letter in ("W", "X") and c.normalized_name in ("sale", "collection"))
        ]

        # Cell range helper
        start_r = source_region.data_start_row
        end_r = source_region.data_end_row

        def get_source_ref(col: ColumnHeader) -> str:
            return f"{col.column_letter}{start_r}:{col.column_letter}{end_r}"

        # RULE 1: If source contains an Accounts Verified section,
        # prioritize verified columns for financial amounts and hold counselor claimed equivalents.
        claimed_cols_to_supersede: Set[int] = set()
        if has_verified_section:
            verified_cats = {
                self.get_token_category(c.name)
                for c in source_region.columns
                if c.is_verified_section and self.get_token_category(c.name)
            }
            for c in source_region.columns:
                if not c.is_verified_section:
                    cat = self.get_token_category(c.name)
                    # For collection amount, verified collection takes absolute precedence
                    if cat == "collection_amount" and "collection_amount" in verified_cats:
                        claimed_cols_to_supersede.add(c.index)
                        unmapped_reasons[c.name] = (
                            f"Counselor Claimed field ({get_source_ref(c)}) superseded by Accounts Verified equivalent per Rule 1."
                        )
                        ambiguities.append(
                            f"Duplicate source fields named '{c.name}' detected: Accounts Verified prioritized over Counselor Claimed (Col {c.index}) per Rule 1."
                        )

        # PASS 1: Exact Normalized Header Matches (Score 1.0)
        # When has_verified_section is True, verified columns match first
        src_cols_ordered = sorted(
            source_region.columns,
            key=lambda c: (1 if has_verified_section and c.is_verified_section else 2)
        )

        for src_col in src_cols_ordered:
            if src_col.index in mapped_src_indices or src_col.index in claimed_cols_to_supersede:
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
                        confidence=1.0,
                        section_context=src_col.section or ("ACCOUNTS VERIFIED" if src_col.is_verified_section else "CLAIMED"),
                        source_cell_references=get_source_ref(src_col),
                        is_verified_source=src_col.is_verified_section
                    ))
                    break

        # PASS 2: Strong Category & Token Identity Matches (Score 0.95 - 0.99)
        candidate_matches: List[Tuple[float, ColumnHeader, ColumnHeader, str]] = []

        for src_col in src_cols_ordered:
            if src_col.index in mapped_src_indices or src_col.index in claimed_cols_to_supersede:
                continue
            src_cat = self.get_token_category(src_col.name)
            if not src_cat or src_cat in ("list_price", "identifier", "work_order"):
                continue

            for dest_col in dest_input_cols:
                if dest_col.index in mapped_dest_indices:
                    continue
                dest_cat = self.get_token_category(dest_col.name)
                if not dest_cat or src_cat != dest_cat:
                    continue

                score = 0.95
                # Give priority bonus to verified columns
                if src_col.is_verified_section:
                    score += 0.03

                src_tokens = set(filter(None, src_col.normalized_name.split('_')))
                dest_tokens = set(filter(None, dest_col.normalized_name.split('_')))
                if src_tokens == dest_tokens:
                    score = min(0.99, score + 0.02)
                elif src_tokens.issubset(dest_tokens) or dest_tokens.issubset(src_tokens):
                    score = min(0.98, score + 0.01)

                candidate_matches.append((score, src_col, dest_col, "category_token_match"))

        # Sort candidate matches by score descending (verified fields will be prioritized)
        candidate_matches.sort(key=lambda x: x[0], reverse=True)

        # Detect competing source columns for the same destination column
        dest_competition: Dict[int, List[ColumnHeader]] = {}
        for score, src_col, dest_col, _ in candidate_matches:
            dest_competition.setdefault(dest_col.index, []).append(src_col)

        for dest_idx, competitors in dest_competition.items():
            if len(competitors) > 1:
                comp_desc = ", ".join([f"Col {c.index} '{c.name}' ({c.section or 'Claimed'})" for c in competitors])
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

            if src_col.normalized_name in mapped_src_norm_names and not src_col.is_verified_section:
                unmapped_reasons[src_col.name] = (
                    f"Duplicate source field '{src_col.name}' (Col {src_col.index}) held: competing with verified '{src_col.name}'."
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
                is_ambiguous=is_competing and not src_col.is_verified_section,
                ambiguity_reason=f"Competing with verified source field for '{dest_col.name}'" if is_competing else None,
                section_context=src_col.section or ("ACCOUNTS VERIFIED" if src_col.is_verified_section else "CLAIMED"),
                source_cell_references=get_source_ref(src_col),
                is_verified_source=src_col.is_verified_section
            ))

        # PASS 3: High-Confidence Fuzzy Matches (> 0.85)
        for src_col in src_cols_ordered:
            if src_col.index in mapped_src_indices or src_col.role == ColumnRole.INDEX_SEQUENCE:
                continue
            if src_col.index in claimed_cols_to_supersede:
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
                    confidence=round(best_score, 3),
                    section_context=src_col.section or ("ACCOUNTS VERIFIED" if src_col.is_verified_section else "CLAIMED"),
                    source_cell_references=get_source_ref(src_col),
                    is_verified_source=src_col.is_verified_section
                ))

        # RULE 2: Dynamic Product Normalization
        # Check if destination table has Column E 'Product Name' and source has Column D 'PRODUCT NAME'
        dest_col_e = next((c for c in destination_region.columns if c.column_letter == "E"), None)
        src_prod_col = next((c for c in source_region.columns if self.get_token_category(c.name) == "product"), None)

        prod_norm_details = {}
        if dest_col_e and src_prod_col:
            raw_text = sample_source_product or "Ph.D Admission to Completion"
            resolved_cat = None
            norm_score = 0.0
            is_ambig = False
            ev_list = []

            if self.product_normalizer:
                resolved_cat, norm_score, is_ambig, ev_list = self.product_normalizer.normalize(raw_text)

            prod_norm_details = {
                "source_field": src_prod_col.name,
                "source_ref": get_source_ref(src_prod_col),
                "raw_product_text": raw_text,
                "master_d_treatment": "Preserve original product text in Master Column D",
                "master_e_proposed_category": resolved_cat,
                "confidence": norm_score,
                "is_ambiguous": is_ambig,
                "evidence": ev_list
            }

            if resolved_cat and not is_ambig:
                mappings.append(ColumnMappingRule(
                    source_column=src_prod_col.name,
                    source_index=src_prod_col.index,
                    destination_column=dest_col_e.name,
                    destination_index=dest_col_e.index,
                    destination_column_letter=dest_col_e.column_letter,
                    transformation_rule=f"lookup_normalize:{resolved_cat}",
                    is_formula_column=dest_col_e.role == ColumnRole.CALCULATED_FORMULA,
                    confidence=norm_score,
                    section_context=src_prod_col.section or "CLAIMED",
                    source_cell_references=get_source_ref(src_prod_col),
                    is_verified_source=False
                ))
            elif is_ambig:
                ambiguities.append(
                    f"Product normalization for '{raw_text}' requires human approval ({'; '.join(ev_list)})."
                )

        self.last_product_normalization_details = prod_norm_details

        # Identify all unmapped source columns
        unmapped_sources = []
        for c in source_region.columns:
            if c.index not in mapped_src_indices and c.role != ColumnRole.INDEX_SEQUENCE:
                reason = unmapped_reasons.get(c.name)
                if not reason:
                    if self.get_token_category(c.name) == "list_price":
                        reason = f"Catalog list price ({get_source_ref(c)}) — absent in Master table."
                    elif self.get_token_category(c.name) in ("identifier", "work_order"):
                        reason = f"Reference metadata ({get_source_ref(c)}) — not tracked in Master table."
                    else:
                        reason = f"Unmapped field ({get_source_ref(c)}) — held from transfer."
                unmapped_sources.append(f"{c.name} ({get_source_ref(c)}): {reason}")

        self.last_unmapped_columns = unmapped_sources
        self.last_ambiguities = ambiguities

        # Sort mappings by destination column index
        mappings.sort(key=lambda m: m.destination_index)
        return mappings
