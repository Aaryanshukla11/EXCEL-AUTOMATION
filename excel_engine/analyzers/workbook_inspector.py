"""Generic Workbook Inspector.

Analyzes any Excel workbook structure dynamically without hardcoded assumptions.
Discovers worksheets, metadata regions, header rows, data tables, preallocated formula slots,
and cross-sheet dependency relationships.
"""
import re
import zipfile
import xml.etree.ElementTree as ET
from datetime import datetime
from typing import Any, Dict, List, Optional, Tuple
import openpyxl
from openpyxl.utils import get_column_letter

from shared.enums import ColumnRole
from shared.schemas import ColumnHeader, DataRegion, EntityMetadata

NS = {
    'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
    'pkg_r': 'http://schemas.openxmlformats.org/package/2006/relationships'
}


class WorkbookInspector:
    """Dynamic inspector for unfamiliar Excel workbooks."""

    def __init__(self, filepath: str):
        self.filepath = filepath

    def get_sheet_names(self) -> List[str]:
        """Return list of sheet names in the workbook using fast zip inspection."""
        with zipfile.ZipFile(self.filepath, 'r') as z:
            wb_xml = z.read('xl/workbook.xml')
            root = ET.fromstring(wb_xml)
            return [elem.attrib.get('name') for elem in root.iter(f"{{{NS['main']}}}sheet")]

    def get_sheet_target_path(self, sheet_name: str) -> str:
        """Resolve sheet name to internal OpenXML path (e.g., 'xl/worksheets/sheet21.xml')."""
        with zipfile.ZipFile(self.filepath, 'r') as z:
            wb_rels_root = ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
            rels = {elem.attrib['Id']: elem.attrib['Target'] for elem in wb_rels_root.findall('pkg_r:Relationship', NS)}
            
            wb_root = ET.fromstring(z.read('xl/workbook.xml'))
            for s in wb_root.findall(f".//{{{NS['main']}}}sheet"):
                if s.attrib.get('name') == sheet_name:
                    r_id = s.attrib.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
                    target = rels.get(r_id, "")
                    if target.startswith("/"):
                        return target.lstrip("/")
                    elif not target.startswith("xl/"):
                        return f"xl/{target}"
                    return target
        raise ValueError(f"Sheet '{sheet_name}' not found in workbook {self.filepath}")

    def discover_primary_data_sheet(self) -> str:
        """Dynamically identifies the primary transaction/data sheet in a multi-sheet workbook.
        
        Evaluates candidate sheets across the workbook and scores metadata presence
        and record density. Solves multi-sheet files (e.g. Krishika Kataria) without hardcoding.
        """
        sheet_names = self.get_sheet_names()
        if len(sheet_names) == 1:
            return sheet_names[0]

        scored_candidates = []
        for s_name in sheet_names:
            score = 0
            # Inspect metadata block
            meta = self.inspect_metadata_block(s_name)
            # Valid entity ID must be an alphanumeric identifier code, not a date or phrase
            if meta.entity_id and re.match(r'^[A-Za-z0-9\s_-]{2,15}$', str(meta.entity_id)):
                score += 60
            if meta.entity_name and len(str(meta.entity_name)) > 2:
                score += 40

            # Discover table region
            region = self.discover_table_region(s_name)
            if region:
                row_count = region.data_end_row - region.data_start_row + 1
                col_count = len(region.columns)
                score += min(row_count * 2, 20) + min(col_count * 2, 20)

            scored_candidates.append((s_name, score))

        scored_candidates.sort(key=lambda x: x[1], reverse=True)
        return scored_candidates[0][0]

    def discover_summary_sheet(self) -> Optional[str]:
        """Dynamically discovers summary/roster sheets in a Master workbook.
        
        Scans sheet names for summary/roster keywords and checks for cross-sheet formula links.
        Eliminates hardcoded summary sheet names.
        """
        sheet_names = self.get_sheet_names()
        candidate_keywords = ["summary", "roster", "master", "overview", "payout", "team"]

        for s_name in sheet_names:
            norm = s_name.lower()
            if any(kw in norm for kw in candidate_keywords):
                try:
                    links = self.extract_formula_linkages_from_summary(s_name)
                    if len(links) >= 3:
                        return s_name
                except Exception:
                    pass

        return None

    def inspect_metadata_block(self, sheet_name: str, max_search_rows: int = 6) -> EntityMetadata:
        """Dynamically detect key-value metadata pairs in top rows of any sheet.
        
        Distinguishes key-value pairs (1-3 cells in row) from table headers (5+ cells in row).
        """
        wb = openpyxl.load_workbook(self.filepath, read_only=True, data_only=True)
        ws = wb[sheet_name]
        
        extracted_attributes = {}
        raw_cells = {}
        entity_id = None
        entity_name = None
        supervisor_name = None

        id_patterns = [r'\bemp[\.\s_-]*id\b', r'\bemployee[\.\s_-]*id\b', r'\bstaff[\.\s_-]*id\b', r'\bclient[\.\s_-]*id\b', r'\bvendor[\.\s_-]*id\b', r'\bcode\b']
        name_patterns = [r'\bemp[\.\s_-]*name\b', r'\bemployee[\.\s_-]*name\b', r'\bcounsellor[\.\s_-]*name\b', r'\bclient[\.\s_-]*name\b', r'\bperson[\.\s_-]*name\b']
        manager_patterns = [r'\bam[\.\s_-]*name\b', r'\btl[\.\s_-]*name\b', r'\bmanager\b', r'\bteam[\.\s_-]*leader\b', r'\bsupervisor\b']

        for r in range(1, max_search_rows + 1):
            row_cells = []
            for c in range(1, 15):
                val = ws.cell(r, c).value
                if val is not None and not isinstance(val, datetime):
                    row_cells.append((c, str(val).strip()))

            # If a row has >= 5 cells, it is likely a tabular data header, not a key-value metadata row!
            if len(row_cells) >= 5:
                continue

            for i, (col_idx, text) in enumerate(row_cells):
                norm_text = text.lower()
                next_val = row_cells[i + 1][1] if i + 1 < len(row_cells) else None

                if not next_val or isinstance(next_val, datetime):
                    continue

                # 1. Entity ID
                if any(re.search(pat, norm_text) for pat in id_patterns):
                    if not entity_id and re.match(r'^[A-Za-z0-9\s_-]+$', next_val):
                        entity_id = next_val
                        raw_cells["entity_id"] = f"{get_column_letter(row_cells[i+1][0])}{r}"
                # 2. Entity Name
                elif any(re.search(pat, norm_text) for pat in name_patterns):
                    if not entity_name:
                        entity_name = next_val
                        raw_cells["entity_name"] = f"{get_column_letter(row_cells[i+1][0])}{r}"
                # 3. Supervisor Name
                elif any(re.search(pat, norm_text) for pat in manager_patterns):
                    if not supervisor_name:
                        supervisor_name = next_val
                        raw_cells[text] = f"{get_column_letter(row_cells[i+1][0])}{r}"
                elif next_val:
                    extracted_attributes[text] = next_val

        wb.close()
        return EntityMetadata(
            entity_id=entity_id,
            entity_name=entity_name,
            supervisor_name=supervisor_name,
            additional_attributes=extracted_attributes,
            raw_source_cells=raw_cells,
        )

    TEMPLATE_PLACEHOLDER_VALUES = {
        "no", "na", "n/a", "-", "--", "none", "null", "false", "0", "0.0",
        "n.a", "n.a.", "#n/a", "#div/0!", "#value!", "#ref!", "#name?", "#num!", "#null!", "not applicable"
    }

    @staticmethod
    def is_primary_identity_column(col_name: str) -> bool:
        """Determines if a column is a primary transactional identity column."""
        norm = re.sub(r'[^a-z0-9]', '_', col_name.lower()).strip('_')
        primary_indicators = [
            "name", "customer", "contact", "candidate", "client", "person", "student", "employee",
            "date", "inv_no", "invoice", "wo_id", "order_id", "pi_no", "bill_no"
        ]
        # Ignore secondary lookup or status columns
        if any(sec in norm for sec in ["status", "grid", "target", "matrix", "remark", "as_per"]):
            return False
        return any(ind in norm for ind in primary_indicators)

    def discover_table_region(self, sheet_name: str, max_scan_rows: int = 35) -> Optional[DataRegion]:
        """Dynamically locate the primary data table, headers, data slots, and summary row.
        
        Scans both values and formula expressions directly from sheet XML to be fast and
        immune to openpyxl whole-workbook corruption.
        """
        sheet_path = self.get_sheet_target_path(sheet_name)
        with zipfile.ZipFile(self.filepath, 'r') as z:
            sheet_bytes = z.read(sheet_path)
            shared_strings = []
            if 'xl/sharedStrings.xml' in z.namelist():
                ss_root = ET.fromstring(z.read('xl/sharedStrings.xml'))
                for si in ss_root.findall(f"{{{NS['main']}}}si"):
                    shared_strings.append(''.join(t.text for t in si.findall(f".//{{{NS['main']}}}t") if t.text))

        root = ET.fromstring(sheet_bytes)
        sheet_data = root.find(f"{{{NS['main']}}}sheetData")

        rows: Dict[int, Dict[int, Tuple[Optional[str], Optional[str]]]] = {}
        for row_elem in sheet_data.findall(f"{{{NS['main']}}}row"):
            r_idx = int(row_elem.attrib.get('r', 0))
            cells = {}
            for c_elem in row_elem.findall(f"{{{NS['main']}}}c"):
                ref = c_elem.attrib.get('r', '')
                col_match = re.match(r'([A-Z]+)(\d+)', ref)
                if not col_match:
                    continue
                col_letters = col_match.group(1)
                c_idx = 0
                for char in col_letters:
                    c_idx = c_idx * 26 + (ord(char) - ord('A') + 1)

                t = c_elem.attrib.get('t')
                v_el = c_elem.find(f"{{{NS['main']}}}v")
                f_el = c_elem.find(f"{{{NS['main']}}}f")
                is_el = c_elem.find(f"{{{NS['main']}}}is")

                val = None
                if is_el is not None:
                    t_el = is_el.find(f"{{{NS['main']}}}t")
                    val = t_el.text if t_el is not None else None
                elif v_el is not None:
                    val = v_el.text
                    if t == 's' and val is not None and int(val) < len(shared_strings):
                        val = shared_strings[int(val)]

                formula = f_el.text if f_el is not None else None
                cells[c_idx] = (val, formula)
            rows[r_idx] = cells

        # 1. Discover Header Row by evaluating text density
        candidate_rows = []
        for r in range(1, max_scan_rows + 1):
            if r not in rows:
                continue
            row_dict = rows[r]
            text_cells = 0
            distinct_names = []
            for c_idx, (val, formula) in row_dict.items():
                if val and isinstance(val, str) and not formula:
                    text_cells += 1
                    distinct_names.append(val.strip())

            if text_cells >= 4:
                keywords = ["name", "date", "id", "amount", "no", "product", "price", "status", "sale", "collection", "total", "code", "type", "qty", "rate", "value", "desc", "item", "num", "ref"]
                score = sum(1 for w in distinct_names if any(k in w.lower() for k in keywords))
                candidate_rows.append((r, score, text_cells))

        if not candidate_rows:
            return None

        candidate_rows.sort(key=lambda x: (x[1], x[2]), reverse=True)
        header_row = candidate_rows[0][0]

        # 2. Extract Columns and Detect Surrounding Section Headers (e.g. ACCOUNTS VERIFIED)
        section_markers: Dict[int, str] = {}
        for r_sec in range(max(1, header_row - 2), header_row):
            if r_sec in rows:
                for c_idx, (val, _) in rows[r_sec].items():
                    if val and isinstance(val, str):
                        val_clean = val.strip()
                        val_lower = val_clean.lower()
                        if any(k in val_lower for k in ("verified", "accounts verified", "audit", "claimed", "counsellor")):
                            section_markers[c_idx] = val_clean

        sorted_sec_col_indices = sorted(section_markers.keys())

        columns = []
        header_cells = rows.get(header_row, {})
        for c_idx in sorted(header_cells.keys()):
            val, _ = header_cells[c_idx]
            if val is not None and str(val).strip():
                name = str(val).strip()
                col_letter = get_column_letter(c_idx)
                norm_name = re.sub(r'[^a-zA-Z0-9]', '_', name.lower()).strip('_')

                # Determine section membership
                current_section = None
                applicable_markers = [s_idx for s_idx in sorted_sec_col_indices if s_idx <= c_idx]
                if applicable_markers:
                    current_section = section_markers[applicable_markers[-1]]
                else:
                    current_section = "CLAIMED" if sorted_sec_col_indices else None

                is_verified = bool(current_section and any(k in current_section.lower() for k in ("verified", "accounts verified", "audit")))

                sample_formula = None
                has_formula = False
                for test_r in range(header_row + 1, min(header_row + 6, max(rows.keys(), default=header_row) + 1)):
                    if test_r in rows and c_idx in rows[test_r]:
                        _, f = rows[test_r][c_idx]
                        if f:
                            has_formula = True
                            sample_formula = f
                            break

                role = ColumnRole.CALCULATED_FORMULA if has_formula else ColumnRole.INPUT_VALUE
                if norm_name in ("s_no", "sno", "sr_no", "serial_no", "sl_no") or norm_name.startswith(("s_no_", "sr_no_")):
                    role = ColumnRole.INDEX_SEQUENCE

                columns.append(ColumnHeader(
                    index=c_idx,
                    column_letter=col_letter,
                    name=name,
                    normalized_name=norm_name,
                    role=role,
                    sample_formula=sample_formula,
                    has_formula_slots=has_formula,
                    section=current_section,
                    is_verified_section=is_verified
                ))

        # 3. Discover Data Boundaries, Preallocated Slots, and Totals Anchor
        empty_slots = []
        totals_row = None
        data_start_row = header_row + 1
        current_data_end = header_row

        max_row_in_sheet = max(rows.keys(), default=header_row)
        for r in range(data_start_row, min(max_row_in_sheet + 1, header_row + 100)):
            if r not in rows:
                empty_slots.append(r)
                continue

            r_cells = rows[r]
            row_vals_str = [str(val or "") for (val, formula) in r_cells.values()]
            is_totals = any("total" in v.lower() or "sum" in v.lower() or "average" in v.lower() for v in row_vals_str)
            has_sum_formula = any(str(formula or "").startswith("SUM") or str(formula or "").startswith("SUBTOTAL") for (val, formula) in r_cells.values())

            if is_totals or has_sum_formula:
                totals_row = r
                break

            # Distinguish genuine user transaction records from default template slots
            has_primary_identity = False
            non_placeholder_count = 0

            for col in columns:
                # Formula columns, totals, and index sequences are not user-entered transaction records
                if col.role in (ColumnRole.CALCULATED_FORMULA, ColumnRole.TOTAL_AGGREGATE, ColumnRole.INDEX_SEQUENCE):
                    continue

                if col.index in r_cells:
                    v, f = r_cells[col.index]
                    # Formulas are dormant calculation slots, not user-entered data records
                    if f is not None:
                        continue
                    if v is not None:
                        v_str = str(v).strip()
                        if not v_str:
                            continue
                        is_placeholder = v_str.lower() in self.TEMPLATE_PLACEHOLDER_VALUES
                        if not is_placeholder:
                            non_placeholder_count += 1
                            if self.is_primary_identity_column(col.name):
                                has_primary_identity = True

            # An occupied transaction row has either a valid primary identity (name, date, invoice)
            # or multiple distinct non-placeholder inputs.
            is_occupied_record = has_primary_identity or (non_placeholder_count >= 2)

            if is_occupied_record:
                current_data_end = r
            else:
                empty_slots.append(r)

        data_end_row = totals_row - 1 if totals_row else max(current_data_end, header_row + 1)

        return DataRegion(
            sheet_name=sheet_name,
            header_row=header_row,
            data_start_row=data_start_row,
            data_end_row=data_end_row,
            totals_row=totals_row,
            columns=columns,
            preallocated_slots_count=len(empty_slots),
            available_empty_slots=empty_slots,
        )

    def extract_formula_linkages_from_summary(self, summary_sheet_name: str) -> List[Dict[str, Any]]:
        """Inspects summary/roster sheet for cross-sheet formula links without whole-file openpyxl load."""
        sheet_path = self.get_sheet_target_path(summary_sheet_name)
        with zipfile.ZipFile(self.filepath, 'r') as z:
            sheet_bytes = z.read(sheet_path)
            shared_strings = []
            if 'xl/sharedStrings.xml' in z.namelist():
                ss_root = ET.fromstring(z.read('xl/sharedStrings.xml'))
                for si in ss_root.findall(f"{{{NS['main']}}}si"):
                    shared_strings.append(''.join(t.text for t in si.findall(f".//{{{NS['main']}}}t") if t.text))

        root = ET.fromstring(sheet_bytes)
        sheet_data = root.find(f"{{{NS['main']}}}sheetData")
        sheet_link_pattern = re.compile(r"['\"]?([A-Za-z0-9_\s\-\.]+)['\"]?!([A-Z]+[0-9]+)")

        linkages = []
        for row_elem in sheet_data.findall(f"{{{NS['main']}}}row"):
            r = int(row_elem.attrib.get('r', 0))
            row_id = None
            row_name = None
            linked_sheets = set()

            for c_elem in row_elem.findall(f"{{{NS['main']}}}c"):
                t = c_elem.attrib.get('t')
                v_el = c_elem.find(f"{{{NS['main']}}}v")
                f_el = c_elem.find(f"{{{NS['main']}}}f")
                is_el = c_elem.find(f"{{{NS['main']}}}is")

                val = ""
                if is_el is not None:
                    t_el = is_el.find(f"{{{NS['main']}}}t")
                    val = t_el.text if t_el is not None else ""
                elif v_el is not None:
                    val = v_el.text or ""
                    if t == 's' and val and int(val) < len(shared_strings):
                        val = shared_strings[int(val)]

                formula = (f_el.text if f_el is not None else None) or ""

                if re.match(r"^[A-Za-z]{2,5}\s*\d{3,6}$", val.strip()):
                    row_id = val.strip()
                elif len(val) > 2 and not formula and not row_name:
                    if any(char.isalpha() for char in val):
                        row_name = val.strip()

                if "!" in formula:
                    matches = sheet_link_pattern.findall(formula)
                    for target_sheet, _ in matches:
                        if target_sheet != summary_sheet_name:
                            linked_sheets.add(target_sheet)

            if row_id and linked_sheets:
                for target_sheet in linked_sheets:
                    linkages.append({
                        "summary_row": r,
                        "entity_id": row_id,
                        "entity_name": row_name,
                        "linked_sheet_name": target_sheet
                    })

        return linkages

    def discover_product_lookup_grid(self, lookup_sheet_name: str = "Incentive Plan") -> Dict[str, float]:
        """Dynamically inspects the product lookup grid from the workbook (e.g. Incentive Plan sheet).
        
        Discovers the lookup table columns (e.g. Product Name -> Advance Collection Target)
        without hardcoding specific product names, column letters, or row positions.
        """
        sheet_names = self.get_sheet_names()
        target_sheet = None
        for s in sheet_names:
            if s.lower() == lookup_sheet_name.lower():
                target_sheet = s
                break
        if not target_sheet:
            for s in sheet_names:
                if "plan" in s.lower() or "slab" in s.lower():
                    target_sheet = s
                    break

        if not target_sheet:
            return {}

        sheet_path = self.get_sheet_target_path(target_sheet)
        with zipfile.ZipFile(self.filepath, 'r') as z:
            sheet_bytes = z.read(sheet_path)
            shared_strings = []
            if 'xl/sharedStrings.xml' in z.namelist():
                ss_root = ET.fromstring(z.read('xl/sharedStrings.xml'))
                for si in ss_root.findall(f"{{{NS['main']}}}si"):
                    shared_strings.append(''.join(t.text for t in si.findall(f".//{{{NS['main']}}}t") if t.text))

        root = ET.fromstring(sheet_bytes)
        sheet_data = root.find(f"{{{NS['main']}}}sheetData")

        rows: Dict[int, Dict[int, Tuple[Optional[str], Optional[str]]]] = {}
        for row_elem in sheet_data.findall(f"{{{NS['main']}}}row"):
            r_idx = int(row_elem.attrib.get('r', 0))
            cells = {}
            for c_elem in row_elem.findall(f"{{{NS['main']}}}c"):
                ref = c_elem.attrib.get('r', '')
                col_match = re.match(r'([A-Z]+)(\d+)', ref)
                if not col_match:
                    continue
                col_letters = col_match.group(1)
                c_idx = 0
                for char in col_letters:
                    c_idx = c_idx * 26 + (ord(char) - ord('A') + 1)

                t = c_elem.attrib.get('t')
                v_el = c_elem.find(f"{{{NS['main']}}}v")
                f_el = c_elem.find(f"{{{NS['main']}}}f")
                is_el = c_elem.find(f"{{{NS['main']}}}is")

                val = None
                if is_el is not None:
                    t_el = is_el.find(f"{{{NS['main']}}}t")
                    val = t_el.text if t_el is not None else None
                elif v_el is not None:
                    val = v_el.text
                    if t == 's' and val is not None and int(val) < len(shared_strings):
                        val = shared_strings[int(val)]

                formula = f_el.text if f_el is not None else None
                cells[c_idx] = (val, formula)
            rows[r_idx] = cells

        product_grid: Dict[str, float] = {}
        candidates = []

        for r in range(1, 15):
            if r not in rows:
                continue
            for c_idx, (val, _) in rows[r].items():
                if val and isinstance(val, str) and "product" in val.lower():
                    next_c = c_idx + 1
                    if next_c in rows[r]:
                        next_val = str(rows[r][next_c][0] or "").lower()
                        score = 0
                        if "advance" in next_val:
                            score += 10
                        if any(k in next_val for k in ("target", "traget", "rate", "percent", "grid")):
                            score += 5
                        if score > 0:
                            candidates.append((score, r + 1, c_idx, next_c))

        candidates.sort(key=lambda x: x[0], reverse=True)

        if candidates:
            _, grid_start_row, prod_col_idx, rate_col_idx = candidates[0]
            for r in range(grid_start_row, grid_start_row + 25):
                if r in rows and prod_col_idx in rows[r]:
                    p_name, _ = rows[r][prod_col_idx]
                    rate_val, _ = rows[r].get(rate_col_idx, (None, None))
                    if p_name and isinstance(p_name, str) and p_name.strip():
                        try:
                            rate_float = float(rate_val) if rate_val is not None else 0.0
                        except ValueError:
                            rate_float = 0.0
                        product_grid[p_name.strip()] = rate_float

        return product_grid
