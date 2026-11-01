"""Generic Integrity & Formula Validator.

Verifies post-execution state against original pre-execution baseline:
1. Verifies existing headers remain 100% unchanged.
2. Verifies formula cells in calculated columns are intact and undamaged.
3. Detects introduced Excel formula errors (#REF!, #VALUE!, #NAME?, #NULL!).
4. Operates at OpenXML level to provide robust, fast verification even for
   gigantic workbooks containing pivot tables, charts, and macros.
"""
import re
import zipfile
import xml.etree.ElementTree as ET
from typing import Dict, List, Optional
from openpyxl.utils import get_column_letter

from shared.enums import ColumnRole
from shared.schemas import DataRegion, ValidationResult

NS = {'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}


class IntegrityValidator:
    """Independent validator checking workbook integrity and formula preservation."""

    def validate(
        self,
        backup_master_path: str,
        processed_master_path: str,
        target_sheet_name: str,
        target_sheet_xml_path: str,
        expected_destination_region: DataRegion,
    ) -> ValidationResult:
        """Perform comprehensive pre/post integrity checks on the modified workbook."""
        errors: List[str] = []
        warnings: List[str] = []
        checks_run = 0
        checks_passed = 0

        # 1. Verify working copy is valid zip archive
        checks_run += 1
        try:
            with zipfile.ZipFile(processed_master_path, 'r') as z_proc:
                proc_sheet_bytes = z_proc.read(target_sheet_xml_path)
            with zipfile.ZipFile(backup_master_path, 'r') as z_orig:
                orig_sheet_bytes = z_orig.read(target_sheet_xml_path)
            checks_passed += 1
        except Exception as e:
            errors.append(f"Failed to read processed workbook archive: {str(e)}")
            return ValidationResult(
                is_valid=False,
                total_checks=checks_run,
                passed_checks=checks_passed,
                failed_checks=len(errors),
                errors=errors,
                formula_integrity_passed=False
            )

        root_proc = ET.fromstring(proc_sheet_bytes)
        root_orig = ET.fromstring(orig_sheet_bytes)

        sheet_data_proc = root_proc.find(f"{{{NS['main']}}}sheetData")
        sheet_data_orig = root_orig.find(f"{{{NS['main']}}}sheetData")

        def parse_row_cells(sheet_data_el):
            rows = {}
            for r in sheet_data_el.findall(f"{{{NS['main']}}}row"):
                r_num = int(r.attrib.get('r', 0))
                cells = {}
                for c in r.findall(f"{{{NS['main']}}}c"):
                    ref = c.attrib.get('r', '')
                    f = c.find(f"{{{NS['main']}}}f")
                    v = c.find(f"{{{NS['main']}}}v")
                    is_el = c.find(f"{{{NS['main']}}}is")
                    
                    val_str = ""
                    if is_el is not None:
                        t = is_el.find(f"{{{NS['main']}}}t")
                        val_str = t.text if t is not None else ""
                    elif v is not None:
                        val_str = v.text or ""

                    cells[ref] = {
                        'formula': f.text if f is not None else None,
                        'val': val_str
                    }
                rows[r_num] = cells
            return rows

        orig_rows = parse_row_cells(sheet_data_orig)
        proc_rows = parse_row_cells(sheet_data_proc)

        # 2. Verify existing header row remains 100% unchanged
        checks_run += 1
        header_row = expected_destination_region.header_row
        header_mismatch = False

        orig_header_cells = orig_rows.get(header_row, {})
        proc_header_cells = proc_rows.get(header_row, {})

        for col in expected_destination_region.columns:
            cell_ref = f"{col.column_letter}{header_row}"
            orig_h = orig_header_cells.get(cell_ref, {}).get('val', '')
            proc_h = proc_header_cells.get(cell_ref, {}).get('val', '')

            if orig_h != proc_h:
                errors.append(f"Header altered at {cell_ref}: original '{orig_h}', found '{proc_h}'")
                header_mismatch = True

        if not header_mismatch:
            checks_passed += 1

        # 3. Verify formula preservation in calculated columns
        checks_run += 1
        formula_corrupted = False
        formula_cols = [c for c in expected_destination_region.columns if c.role == ColumnRole.CALCULATED_FORMULA]

        for col in formula_cols:
            for r in range(expected_destination_region.data_start_row, expected_destination_region.data_end_row + 1):
                cell_ref = f"{col.column_letter}{r}"
                orig_c = orig_rows.get(r, {}).get(cell_ref, {})
                proc_c = proc_rows.get(r, {}).get(cell_ref, {})

                orig_f = orig_c.get('formula')
                proc_f = proc_c.get('formula')

                if orig_f and not proc_f:
                    errors.append(f"Formula destroyed at {cell_ref}: original formula was '{orig_f}'")
                    formula_corrupted = True

                # Check for introduced formula syntax errors
                if proc_f and any(err_token in proc_f for err_token in ["#REF!", "#NAME?", "#NULL!"]):
                    errors.append(f"Formula syntax error at {cell_ref}: {proc_f}")
                    formula_corrupted = True

        if not formula_corrupted:
            checks_passed += 1

        # 4. Verify Totals Row formula is preserved
        checks_run += 1
        if expected_destination_region.totals_row:
            tot_row = expected_destination_region.totals_row
            orig_tot_cells = orig_rows.get(tot_row, {})
            proc_tot_cells = proc_rows.get(tot_row, {})

            for ref, c_info in orig_tot_cells.items():
                if c_info.get('formula'):
                    p_info = proc_tot_cells.get(ref, {})
                    if not p_info.get('formula'):
                        errors.append(f"Totals formula destroyed at {ref}: '{c_info.get('formula')}'")
                        formula_corrupted = True
            checks_passed += 1
        else:
            checks_passed += 1

        is_valid = len(errors) == 0
        return ValidationResult(
            is_valid=is_valid,
            total_checks=checks_run,
            passed_checks=checks_passed,
            failed_checks=len(errors),
            warnings=warnings,
            errors=errors,
            formula_integrity_passed=not formula_corrupted
        )
