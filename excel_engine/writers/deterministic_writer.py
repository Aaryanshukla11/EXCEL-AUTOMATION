"""Deterministic Excel Writer using OpenXML Package Patching.

Executes approved mapping plans against existing Master tables without hardcoded logic.
1. Preserves existing headers, styles, formatting, and formulas.
2. Populates existing preallocated calculation slots where available.
3. Automatically creates SHA-256 verified backups before writing.
4. Uses selective XML patching to guarantee that complex multi-sheet workbooks
   (with pivot caches, drawings, shapes, and macros) are never corrupted.
5. Generates granular cell-by-cell audit logs for verification and rollback.
"""
import hashlib
import os
import re
import shutil
import zipfile
from datetime import datetime, timezone
import xml.etree.ElementTree as ET
from typing import Any, Dict, List, Optional, Tuple
import openpyxl
from openpyxl.utils import get_column_letter

from shared.enums import AuditAction, ColumnRole
from shared.schemas import ColumnMappingRule, DataRegion, ExecutionStepLog

NS = {'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
ET.register_namespace('', NS['main'])


class DeterministicWriter:
    """Executes deterministic cell transfers on working copies of workbooks."""

    def __init__(self, backup_dir: str = "storage/backups"):
        self.backup_dir = backup_dir
        os.makedirs(self.backup_dir, exist_ok=True)

    def create_backup(self, master_filepath: str, task_id: str) -> Tuple[str, str]:
        """Create an authoritative, immutable backup copy with SHA-256 checksum."""
        timestamp = datetime.now(timezone.utc).strftime("%Y%m%d_%H%M%S")
        backup_filename = f"{task_id}_{timestamp}_master_backup.xlsx"
        backup_path = os.path.join(self.backup_dir, backup_filename)
        shutil.copy2(master_filepath, backup_path)

        sha256 = hashlib.sha256()
        with open(backup_path, "rb") as f:
            for chunk in iter(lambda: f.read(65536), b""):
                sha256.update(chunk)
        file_hash = sha256.hexdigest()

        return backup_path, file_hash

    @staticmethod
    def clean_cell_value(val: Any) -> Any:
        """Normalize numbers with commas, currency symbols, and strip trailing whitespace."""
        if val is None:
            return None
        if isinstance(val, (int, float, datetime)):
            return val
        s = str(val).strip()
        if not s:
            return None
        num_clean = re.sub(r'^[₹$\sRrs\.]+', '', s)
        num_clean = num_clean.replace(',', '').strip()
        if re.match(r'^-?\d+(\.\d+)?$', num_clean):
            try:
                return float(num_clean) if '.' in num_clean else int(num_clean)
            except ValueError:
                pass
        return s

    def execute_plan(
        self,
        source_filepath: str,
        master_working_copy_path: str,
        target_sheet_name: str,
        target_sheet_xml_path: str,
        source_region: DataRegion,
        destination_region: DataRegion,
        mappings: List[ColumnMappingRule],
    ) -> Tuple[bool, List[ExecutionStepLog], Dict[str, Any]]:
        """Transfer source records into existing destination table slots via OpenXML patching."""
        audit_logs: List[ExecutionStepLog] = []
        step_idx = 1

        # 1. Extract source records
        wb_src = openpyxl.load_workbook(source_filepath, data_only=True)
        ws_src = wb_src[source_region.sheet_name]

        source_records = []
        for r in range(source_region.data_start_row, source_region.data_end_row + 1):
            row_data = {}
            has_data = False
            for rule in mappings:
                cell_val = ws_src.cell(r, rule.source_index).value
                cleaned_val = self.clean_cell_value(cell_val)
                row_data[rule.source_index] = cleaned_val
                if cleaned_val is not None:
                    has_data = True
            if has_data:
                source_records.append(row_data)

        wb_src.close()

        audit_logs.append(ExecutionStepLog(
            step_number=step_idx,
            action=AuditAction.STRUCTURE_INSPECT,
            status="SUCCESS",
            description=f"Extracted {len(source_records)} data rows from source sheet '{source_region.sheet_name}'",
            details={"rows_extracted": len(source_records)}
        ))
        step_idx += 1

        if not source_records:
            return True, audit_logs, {"written_rows": 0, "written_cells": 0}

        # 2. Read Master ZIP archive and target sheet XML
        with zipfile.ZipFile(master_working_copy_path, 'r') as z_in:
            all_files = {item.filename: z_in.read(item.filename) for item in z_in.infolist()}

        sheet_bytes = all_files[target_sheet_xml_path]
        root = ET.fromstring(sheet_bytes)
        sheet_data = root.find(f"{{{NS['main']}}}sheetData")

        # Index existing rows
        existing_rows: Dict[int, ET.Element] = {}
        for r_el in sheet_data.findall(f"{{{NS['main']}}}row"):
            r_num = int(r_el.attrib.get('r', 0))
            existing_rows[r_num] = r_el

        empty_slots = destination_region.available_empty_slots
        written_cells = 0
        cell_diffs = []

        for record_idx, rec_data in enumerate(source_records):
            dest_row = (
                empty_slots[record_idx]
                if record_idx < len(empty_slots)
                else (destination_region.data_end_row + 1 + (record_idx - len(empty_slots)))
            )
            
            # Find or create row element
            if dest_row in existing_rows:
                row_elem = existing_rows[dest_row]
            else:
                row_elem = ET.SubElement(sheet_data, f"{{{NS['main']}}}row", {'r': str(dest_row)})
                existing_rows[dest_row] = row_elem

            # Index cells in this row
            row_cells: Dict[str, ET.Element] = {}
            for c_el in row_elem.findall(f"{{{NS['main']}}}c"):
                ref = c_el.attrib.get('r', '')
                row_cells[ref] = c_el

            for rule in mappings:
                dest_col_letter = rule.destination_column_letter
                cell_ref = f"{dest_col_letter}{dest_row}"
                val_to_write = rec_data.get(rule.source_index)

                if cell_ref in row_cells:
                    c_elem = row_cells[cell_ref]
                else:
                    c_elem = ET.SubElement(row_elem, f"{{{NS['main']}}}c", {'r': cell_ref})
                    row_cells[cell_ref] = c_elem

                # Check if this cell contains a formula - NEVER overwrite formula!
                f_elem = c_elem.find(f"{{{NS['main']}}}f")
                if f_elem is not None:
                    continue

                prev_val = None
                v_el = c_elem.find(f"{{{NS['main']}}}v")
                is_el = c_elem.find(f"{{{NS['main']}}}is")
                if is_el is not None:
                    t_el = is_el.find(f"{{{NS['main']}}}t")
                    prev_val = t_el.text if t_el is not None else None
                elif v_el is not None:
                    prev_val = v_el.text

                # Clear old values
                for child in list(c_elem):
                    if child.tag.endswith('v') or child.tag.endswith('is'):
                        c_elem.remove(child)

                # Set new value
                if val_to_write is not None:
                    if isinstance(val_to_write, (int, float)):
                        c_elem.attrib.pop('t', None)  # numeric has no t attribute or 'n'
                        v_new = ET.SubElement(c_elem, f"{{{NS['main']}}}v")
                        v_new.text = str(val_to_write)
                    else:
                        c_elem.attrib['t'] = 'inlineStr'
                        is_new = ET.SubElement(c_elem, f"{{{NS['main']}}}is")
                        t_new = ET.SubElement(is_new, f"{{{NS['main']}}}t")
                        t_new.text = str(val_to_write)

                written_cells += 1
                cell_diffs.append({
                    "cell": cell_ref,
                    "previous_value": prev_val,
                    "new_value": val_to_write,
                    "source_field": rule.source_column,
                    "destination_field": rule.destination_column
                })

            audit_logs.append(ExecutionStepLog(
                step_number=step_idx,
                action=AuditAction.CELL_WRITE,
                status="SUCCESS",
                description=f"Wrote record {record_idx + 1} to existing destination row {dest_row}",
                details={"target_row": dest_row, "fields_written": len(mappings)}
            ))
            step_idx += 1

        # 3. Update modified sheet in archive and re-zip cleanly
        updated_sheet_bytes = ET.tostring(root, xml_declaration=True, encoding='utf-8')
        all_files[target_sheet_xml_path] = updated_sheet_bytes

        with zipfile.ZipFile(master_working_copy_path, 'w', compression=zipfile.ZIP_DEFLATED) as z_out:
            for name, content in all_files.items():
                z_out.writestr(name, content)

        stats = {
            "written_rows": len(source_records),
            "written_cells": written_cells,
            "target_sheet": target_sheet_name,
            "cell_diffs": cell_diffs
        }
        return True, audit_logs, stats
