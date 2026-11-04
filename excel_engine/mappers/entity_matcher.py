"""Generic Multi-Signal Entity Matcher.

Discovers the corresponding Master destination worksheet for any unfamiliar Sister file
using weighted composite signals:
1. Strong Unique Identifier match (Employee ID, Account ID, Code)
2. Normalized Name similarity (Token set ratio, Levenshtein distance)
3. Cross-sheet formula dependency linkages in master summary sheets
4. Contextual metadata alignment (Supervisor, Department, Target)

No hardcoded entities, names, or sheet names.
"""
import re
from typing import Dict, List, Optional
from rapidfuzz import fuzz

from shared.enums import MatchConfidenceLevel
from shared.schemas import EntityMatchProposal, EntityMetadata


class EntityMatcher:
    """Calculates matching confidence and proposes destination worksheet."""

    def __init__(
        self,
        weight_id: float = 0.50,
        weight_name: float = 0.30,
        weight_context: float = 0.20,
        auto_threshold: float = 0.90,
        review_threshold: float = 0.70
    ):
        self.weight_id = weight_id
        self.weight_name = weight_name
        self.weight_context = weight_context
        self.auto_threshold = auto_threshold
        self.review_threshold = review_threshold

    @staticmethod
    def normalize_id(identifier: Optional[str]) -> str:
        """Strip spaces, punctuation, and normalize to uppercase."""
        if not identifier:
            return ""
        return re.sub(r'[^A-Za-z0-9]', '', str(identifier)).upper()

    @staticmethod
    def normalize_name(name: Optional[str]) -> str:
        """Strip honorifics, extra spaces, and lowercase."""
        if not name:
            return ""
        s = str(name).lower()
        s = re.sub(r'\b(mr|ms|mrs|dr|sir|maam|mam)\b', '', s)
        s = re.sub(r'[^a-z0-9\s]', '', s)
        return " ".join(s.split())

    def match(
        self,
        source_entity: EntityMetadata,
        candidate_sheets: List[str],
        summary_linkages: Optional[List[Dict[str, str]]] = None,
        sheet_headers_map: Optional[Dict[str, EntityMetadata]] = None,
    ) -> List[EntityMatchProposal]:
        """Rank candidate sheets by multi-signal match confidence."""
        proposals = []
        norm_source_id = self.normalize_id(source_entity.entity_id)
        norm_source_name = self.normalize_name(source_entity.entity_name)

        summary_links = summary_linkages or []
        headers_map = sheet_headers_map or {}

        for sheet_name in candidate_sheets:
            score_breakdown = {}
            evidence = []
            id_score = 0.0
            name_score = 0.0
            context_score = 0.0

            # 1. Check for authoritative formula link in Summary Roster
            formula_linked = False
            for link in summary_links:
                if link.get("linked_sheet_name", "").strip().lower() == sheet_name.strip().lower():
                    link_id = self.normalize_id(link.get("entity_id"))
                    if norm_source_id and link_id == norm_source_id:
                        formula_linked = True
                        id_score = 1.0
                        evidence.append(f"Master summary sheet formula directly links ID '{link['entity_id']}' to sheet '{sheet_name}'")
                        break

            # 2. Check Sheet Header Metadata if formula link not already found
            if not formula_linked and sheet_name in headers_map:
                dest_meta = headers_map[sheet_name]
                dest_id = self.normalize_id(dest_meta.entity_id)
                if norm_source_id and dest_id:
                    if norm_source_id == dest_id:
                        id_score = 1.0
                        evidence.append(f"Destination sheet header contains exact ID match: '{dest_meta.entity_id}'")
                    else:
                        id_score = 0.0

            # 3. Name Similarity against Sheet Name & Target Entity Name
            norm_sheet_name = self.normalize_name(sheet_name)
            # Evaluate token set ratio between full source name and sheet name
            ratio_sheet = fuzz.token_set_ratio(norm_source_name, norm_sheet_name) / 100.0

            # Also check if destination sheet header has employee name
            dest_name_val = headers_map.get(sheet_name, EntityMetadata()).entity_name
            ratio_dest_header = 0.0
            if dest_name_val:
                ratio_dest_header = fuzz.token_set_ratio(norm_source_name, self.normalize_name(dest_name_val)) / 100.0

            name_score = max(ratio_sheet, ratio_dest_header)
            if name_score >= 0.80:
                evidence.append(f"Strong name token alignment between source '{source_entity.entity_name}' and target '{sheet_name}' (score: {name_score:.2f})")
            elif name_score >= 0.50:
                evidence.append(f"Partial name similarity between source and sheet name '{sheet_name}' (score: {name_score:.2f})")

            # 4. Contextual Signal (Supervisor / Manager / Designation)
            if source_entity.supervisor_name and sheet_name in headers_map:
                dest_supervisor = headers_map[sheet_name].supervisor_name
                if dest_supervisor:
                    sup_ratio = fuzz.token_set_ratio(
                        self.normalize_name(source_entity.supervisor_name),
                        self.normalize_name(dest_supervisor)
                    ) / 100.0
                    context_score = sup_ratio
                    if sup_ratio >= 0.75:
                        evidence.append(f"Supervisor alignment: '{source_entity.supervisor_name}' matches '{dest_supervisor}'")

            # Aggregate weighted confidence score
            # If ID is confirmed exact (1.0), boost confidence
            if id_score == 1.0:
                total_score = (self.weight_id * id_score) + (self.weight_name * name_score) + (self.weight_context * max(context_score, 0.5))
                # Boost if both exact ID and strong name match
                if name_score >= 0.70 or formula_linked:
                    total_score = max(total_score, 0.98)
            else:
                total_score = (self.weight_id * id_score) + (self.weight_name * name_score) + (self.weight_context * context_score)

            total_score = round(min(total_score, 1.0), 4)

            score_breakdown["id_score"] = round(id_score, 3)
            score_breakdown["name_score"] = round(name_score, 3)
            score_breakdown["context_score"] = round(context_score, 3)

            if total_score >= self.auto_threshold:
                level = MatchConfidenceLevel.HIGH
                needs_approval = False
            elif total_score >= self.review_threshold:
                level = MatchConfidenceLevel.MEDIUM
                needs_approval = True
            else:
                level = MatchConfidenceLevel.LOW
                needs_approval = True

            proposals.append(EntityMatchProposal(
                source_entity=source_entity,
                candidate_sheet_name=sheet_name,
                confidence_score=total_score,
                confidence_level=level,
                score_breakdown=score_breakdown,
                matching_evidence=evidence,
                requires_human_approval=needs_approval,
            ))

        proposals.sort(key=lambda x: x.confidence_score, reverse=True)
        return proposals

