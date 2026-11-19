"""Formula Recalculation Inspector and Engine.

Investigates and ensures formula recalculation across Master workbooks:
1. Explains OpenXML cached values (<v>) vs formula definitions (<f>).
2. Sets calcPr forceFullCalc/fullCalcOnLoad in xl/workbook.xml to guarantee
   instant full recalculation when opened in any Excel instance.
3. Provides headless COM Excel recalculation when Microsoft Office is present on Windows.
"""
import os
import shutil
import xml.etree.ElementTree as ET
import zipfile
from typing import Any, Dict, List, Optional, Tuple


NS = {
    'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
}


class FormulaRecalculationInspector:
    """Inspects formula dependencies, cached values, and coordinates recalculation."""

    @staticmethod
    def inspect_formula_cells(filepath: str, sheet_name: str, max_rows: int = 40) -> Dict[str, Any]:
        """Inspects formula cells and their cached <v> elements in a given sheet."""
        with zipfile.ZipFile(filepath, 'r') as z:
            # Map sheet name to target xml
            wb_root = ET.fromstring(z.read('xl/workbook.xml'))
            sheets_el = wb_root.find(f"{{{NS['main']}}}sheets")
            sheet_id = None
            r_id = None
            for s in sheets_el.findall(f"{{{NS['main']}}}sheet"):
                if s.attrib.get('name') == sheet_name:
                    r_id = s.attrib.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
                    break

            if not r_id:
                return {"error": f"Sheet {sheet_name} not found"}

            rels_root = ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
            target_path = None
            for rel in rels_root.findall('{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'):
                if rel.attrib.get('Id') == r_id:
                    target = rel.attrib.get('Target', '')
                    target_path = f"xl/{target}" if not target.startswith('xl/') else target
                    break

            if not target_path or target_path not in z.namelist():
                return {"error": f"Sheet target path not found for {sheet_name}"}

            sheet_bytes = z.read(target_path)

        root = ET.fromstring(sheet_bytes)
        sheet_data = root.find(f"{{{NS['main']}}}sheetData")

        formula_cells: List[Dict[str, Any]] = []
        for row in sheet_data.findall(f"{{{NS['main']}}}row"):
            r_idx = int(row.attrib.get('r', 0))
            if r_idx > max_rows:
                continue
            for c in row.findall(f"{{{NS['main']}}}c"):
                ref = c.attrib.get('r', '')
                f_el = c.find(f"{{{NS['main']}}}f")
                v_el = c.find(f"{{{NS['main']}}}v")
                if f_el is not None:
                    formula_cells.append({
                        "cell_ref": ref,
                        "formula": f_el.text,
                        "cached_value": v_el.text if v_el is not None else None,
                        "has_cached_value": v_el is not None,
                    })

        return {
            "sheet_name": sheet_name,
            "total_formula_cells": len(formula_cells),
            "sample_formula_cells": formula_cells[:15],
        }

    @staticmethod
    def ensure_full_calc_on_load(filepath: str) -> bool:
        """Sets fullCalcOnLoad and forceFullCalc attributes in xl/workbook.xml.
        
        This guarantees that Microsoft Excel will force a full dependency recalculation
        of all formulas upon opening the file, updating all cached <v> values.
        """
        temp_zip = f"{filepath}.temp_calc"
        with zipfile.ZipFile(filepath, 'r') as zin, zipfile.ZipFile(temp_zip, 'w', compression=zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                content = zin.read(item.filename)
                if item.filename == 'xl/workbook.xml':
                    root = ET.fromstring(content)
                    calc_pr = root.find(f"{{{NS['main']}}}calcPr")
                    if calc_pr is None:
                        calc_pr = ET.Element(f"{{{NS['main']}}}calcPr")
                        root.append(calc_pr)
                    
                    calc_pr.attrib['fullCalcOnLoad'] = '1'
                    calc_pr.attrib['forceFullCalc'] = '1'
                    calc_pr.attrib['calcCompleted'] = '0'
                    content = ET.tostring(root, encoding='utf-8', xml_declaration=True)
                zout.writestr(item, content)

        shutil.move(temp_zip, filepath)
        return True

    @staticmethod
    def recalculate_via_com(filepath: str) -> Tuple[bool, str]:
        """Performs headless recalculation using Microsoft Excel COM automation if available.
        
        Args:
            filepath: Absolute path to the Excel workbook to recalculate.
            
        Returns:
            Tuple of (success: bool, message: str)
        """
        try:
            import win32com.client
            import pythoncom
        except ImportError:
            return False, "win32com.client not available"

        abs_path = os.path.abspath(filepath)
        pythoncom.CoInitialize()
        excel = None
        try:
            excel = win32com.client.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            excel.ScreenUpdating = False

            wb = excel.Workbooks.Open(
                abs_path,
                UpdateLinks=0,
                ReadOnly=False,
                CorruptLoad=1,
            )
            excel.CalculateFull()
            wb.Save()
            wb.Close(SaveChanges=True)
            return True, "Headless recalculation via Microsoft Excel COM completed successfully."
        except Exception as e:
            return False, f"COM recalculation error: {str(e)}"
        finally:
            if excel:
                try:
                    excel.Quit()
                except Exception:
                    pass
            pythoncom.CoUninitialize()
