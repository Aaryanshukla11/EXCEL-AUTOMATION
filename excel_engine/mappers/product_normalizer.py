"""Dynamic Product Normalizer for Incentive Calculations.

Normalizes raw source product strings (e.g. 'Ph.D Admission to Completion', 'Writing')
into authoritative product categories discovered directly from the workbook's lookup grid
(e.g. 'Incentive Plan'!R:S).
Zero hardcoding: all valid categories are loaded dynamically from the target workbook.
"""
import re
from typing import Dict, List, Optional, Tuple
from rapidfuzz import fuzz


class ProductNormalizer:
    """Resolves arbitrary product text to dynamically discovered authoritative categories."""

    def __init__(self, authoritative_grid: Dict[str, float]):
        self.grid = authoritative_grid
        self.categories = list(authoritative_grid.keys())

    @staticmethod
    def clean_token_text(text: str) -> str:
        """Normalize punctuation and common acronym formats."""
        t = re.sub(r'ph[\.\s_-]*d', 'phd', text, flags=re.IGNORECASE)
        t = re.sub(r'[^a-zA-Z0-9\s]', ' ', t.lower())
        return ' '.join(t.split())

    def normalize(self, raw_product: str) -> Tuple[Optional[str], float, bool, List[str]]:
        """Resolves raw product name to a discovered category.
        
        Returns:
            Tuple of (resolved_category, confidence_score, is_ambiguous, evidence_list)
        """
        if not raw_product or not self.categories:
            return None, 0.0, True, ["Lookup grid is empty or raw product text is missing"]

        clean_raw = self.clean_token_text(raw_product)
        raw_tokens = set(clean_raw.split())
        evidence: List[str] = []

        # 1. Exact match (case-insensitive)
        for cat in self.categories:
            if clean_raw == self.clean_token_text(cat):
                evidence.append(f"Exact normalized match to authoritative category '{cat}'")
                return cat, 1.0, False, evidence

        # 2. Score all candidate categories
        scored_candidates: List[Tuple[float, str]] = []
        for cat in self.categories:
            clean_cat = self.clean_token_text(cat)
            cat_tokens = set(clean_cat.split())

            # Token overlap score
            token_overlap = len(raw_tokens.intersection(cat_tokens)) / max(len(cat_tokens), 1)
            fuzz_ratio = fuzz.token_sort_ratio(clean_raw, clean_cat) / 100.0

            # Composite score prioritizing token overlap
            composite = (token_overlap * 0.6) + (fuzz_ratio * 0.4)
            scored_candidates.append((composite, cat))

        scored_candidates.sort(key=lambda x: x[0], reverse=True)
        best_score, best_cat = scored_candidates[0]
        second_score = scored_candidates[1][0] if len(scored_candidates) > 1 else 0.0

        # Disallow match if score is too low (< 0.55) -> unknown category, do not invent!
        if best_score < 0.55:
            evidence.append(f"No sufficiently matching category found in Incentive Plan for '{raw_product}' (Best candidate: '{best_cat}' at {best_score:.2f})")
            return None, round(best_score, 3), True, evidence

        # Check for ambiguity: if top two candidates are within 0.12 score margin
        is_ambiguous = (best_score - second_score) < 0.12 and second_score > 0.50
        if is_ambiguous:
            evidence.append(
                f"Ambiguous candidates in Incentive Plan: '{best_cat}' ({best_score:.2f}) vs '{scored_candidates[1][1]}' ({second_score:.2f})"
            )
            return best_cat, round(best_score, 3), True, evidence

        evidence.append(
            f"Fuzzy token match from '{raw_product}' to authoritative category '{best_cat}' (Score: {best_score:.2f})"
        )
        return best_cat, round(best_score, 3), False, evidence

