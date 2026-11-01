"""Generic Workbook Inspector.

Analyzes any Excel workbook structure dynamically without hardcoded assumptions.
Discovers worksheets, metadata regions, header rows, data tables, preallocated formula slots,
and cross-sheet dependency relationships.
"""
import re
import zipfile
import xml.etree.ElementTree as ET
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

    def inspect_metadata_block(self, sheet_name: str, max_search_rows: int = 6) -> EntityMetadata:
        """Dynamically detect key-value metadata pairs in the top rows of a sheet.
        
        Matches patterns like 'EMP. NAME: Gagan', 'Employee ID: APL2004', 'ID: 1234',
        'TL: John', 'Manager: Jane', etc.
        """
        wb = openpyxl.load_workbook(self.filepath, read_only=True, data_only=True)
        ws = wb[sheet_name]
        
        extracted_attributes = {}
        raw_cells = {}
        entity_id = None
        entity_name = None
        supervisor_name = None

        id_patterns = [r'emp.*id', r'employee.*id', r'staff.*id', r'id', r'code']
        name_patterns = [r'emp.*name', r'employee.*name', r'counsellor.*name', r'name']
        manager_patterns = [r'am.*name', r'tl.*name', r'manager', r'team.*leader', r'supervisor']

        for r in range(1, max_search_rows + 1):
            row_cells = []
            for c in range(1, 15):
                val = ws.cell(r, c).value
                if val is not None:
                    row_cells.append((c, str(val).strip()))

            for i, (col_idx, text) in enumerate(row_cells):
                norm_text = text.lower()
                next_val = row_cells[i + 1][1] if i + 1 < len(row_cells) else None

                # 1. Entity ID
                if any(re.search(pat, norm_text) for pat in id_patterns):
                    if next_val and not entity_id:
                        entity_id = next_val
                        raw_cells["entity_id"] = f"{get_column_letter(row_cells[i+1][0])}{r}"
                # 2. Entity Name
                elif any(re.search(pat, norm_text) for pat in name_patterns):
                    if next_val and not entity_name:
                        entity_name = next_val
                        raw_cells["entity_name"] = f"{get_column_letter(row_cells[i+1][0])}{r}"
                # 3. Supervisor Name
                elif any(re.search(pat, norm_text) for pat in manager_patterns):
                    if next_val and not supervisor_name:
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

    def discover_table_region(self, sheet_name: str, max_scan_rows: int = 35) -> Optional[DataRegion]:
        """Dynamically locate the primary data table, headers, data slots, and summary row.
        
        Scans both values and formula expressions directly from sheet XML to be fast and
        immune to openpyxl whole-workbook corruption.
        """
        sheet_path = self.get_sheet_target_path(sheet_name)
        with zipfile.ZipFile(self.filepath, 'r') as z:
            sheet_bytes = z.read(sheet_path)
            # Check for shared strings
            shared_strings = []
            if 'xl/sharedStrings.xml' in z.namelist():
                ss_root = ET.fromstring(z.read('xl/sharedStrings.xml'))
                for si in ss_root.findall(f"{{{NS['main']}}}si"):
                    shared_strings.append(''.join(t.text for t in si.findall(f".//{{{NS['main']}}}t") if t.text))

        root = ET.fromstring(sheet_bytes)
        sheet_data = root.find(f"{{{NS['main']}}}sheetData")

        # Parse rows and cells
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
                # Convert letter to 1-based index
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
                    if t == 's' and val is not None:
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
                keywords = ["name", "date", "id", "amount", "no", "product", "price", "status", "sale", "collection", "total"]
                score = sum(1 for w in distinct_names if any(k in w.lower() for k in keywords))
                candidate_rows.append((r, score, text_cells))

        if not candidate_rows:
            return None

        candidate_rows.sort(key=lambda x: (x[1], x[2]), reverse=True)
        header_row = candidate_rows[0][0]

        # 2. Extract Columns
        columns = []
        header_cells = rows.get(header_row, {})
        for c_idx in sorted(header_cells.keys()):
            val, _ = header_cells[c_idx]
            if val is not None and str(val).strip():
                name = str(val).strip()
                col_letter = get_column_letter(c_idx)
                norm_name = re.sub(r'[^a-zA-Z0-9]', '_', name.lower()).strip('_')

                # Check downstream rows (header_row+1 to header_row+5) for formulas
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
                if "s_no" in norm_name or "sno" in norm_name:
                    role = ColumnRole.INDEX_SEQUENCE

                columns.append(ColumnHeader(
                    index=c_idx,
                    column_letter=col_letter,
                    name=name,
                    normalized_name=norm_name,
                    role=role,
                    sample_formula=sample_formula,
                    has_formula_slots=has_formula,
                ))

        # 3. Discover Data Boundaries, Preallocated Slots, and Totals Anchor
        empty_slots = []
        totals_row = None
        data_start_row = header_row + 1
        current_data_end = header_row

        max_row_in_sheet = max(rows.keys(), default=header_row)
        for r in range(data_start_row, min(max_row_in_sheet + 1, header_row + 60)):
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

            input_cols_empty = True
            has_any_content = False
            for col in columns:
                if col.index in r_cells:
                    v, f = r_cells[col.index]
                    if v is not None or f is not None:
                        has_any_content = True
                        if col.role == ColumnRole.INPUT_VALUE and v is not None and str(v).strip():
                            input_cols_empty = False

            if has_any_content and not input_cols_empty:
                current_data_end = r
            elif input_cols_empty:
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
                    if t == 's' and val:
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
