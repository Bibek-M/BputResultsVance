"""
BPUT Excel Exporter Engine
Generates a beautifully designed, executive-standard Excel workbook (.xlsx)
containing complete academic results, semester-by-semester grade sheets,
and KPI summaries using openpyxl.
"""

import os
import re
from typing import Dict, Any, List, Optional
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter


class BPUTExcelExporter:
    """Exports scraped BPUT student results to an elegantly styled Excel workbook."""

    # Color Palette Tokens
    NAVY_PRIMARY = "192F46"       # BPUT Portal Primary Header
    NAVY_DARK = "111E2E"          # Deep Midnight Navy
    BLUE_ACCENT = "1E40AF"        # Royal Accent
    BLUE_LIGHT = "EFF6FF"         # Soft Blue Highlight
    ICE_BG = "F8FAFC"             # Crisp light gray/blue background
    ZEBRA_LIGHT = "F1F5F9"        # Alternating row fill
    WHITE = "FFFFFF"
    BORDER_LIGHT = "CBD5E1"       # Subtle gray border
    BORDER_DARK = "94A3B8"

    # Grade Color Map: (Fill Color, Text Color)
    GRADE_STYLES = {
        "O": ("D1E7DD", "0F5132"),   # Outstanding: Soft Emerald
        "E": ("D4EDDA", "155724"),   # Excellent: Soft Forest Green
        "A": ("CFE2FF", "084298"),   # Very Good: Soft Ocean Blue
        "B": ("D1ECF1", "0C5460"),   # Good: Soft Cyan Blue
        "C": ("E2E3E5", "383D41"),   # Fair: Cool Slate Gray
        "D": ("FFF3CD", "664D03"),   # Below Average / Pass: Soft Amber
        "F": ("F8D7DA", "842029"),   # Fail: Soft Crimson
        "M": ("F8D7DA", "842029"),   # Malpractice: Soft Crimson
        "S": ("F8D7DA", "842029"),   # Absent: Soft Crimson
        "P": ("FFE5D0", "8A3B00"),   # Pending: Soft Apricot
    }

    def __init__(self, data: Dict[str, Any]):
        self.data = data
        self.student_info = data.get("student_info", {})
        self.semesters = data.get("semesters", [])
        self.summary = data.get("summary", {})
        self.wb = Workbook()

    def _get_border(self, top=False, bottom=False, left=False, right=False, color=BORDER_LIGHT, style="thin") -> Border:
        line = Side(border_style=style, color=color)
        return Border(
            top=line if top else None,
            bottom=line if bottom else None,
            left=line if left else None,
            right=line if right else None
        )

    def _box_border(self, color=BORDER_LIGHT, style="thin") -> Border:
        line = Side(border_style=style, color=color)
        return Border(top=line, bottom=line, left=line, right=line)

    def _auto_fit_columns(self, ws, max_len_cap: int = 60):
        """Auto-adjusts column widths to fit content with elegant padding."""
        for col in ws.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                # Skip merged cells or header rows if too long
                if cell.row in [1, 2, 3] and ws.title == "Academic Overview":
                    continue
                val = str(cell.value or "")
                if val:
                    lines = val.split("\n")
                    max_len = max(max_len, max(len(l) for l in lines))
            col_width = min(max(max_len + 3, 11), max_len_cap)
            ws.column_dimensions[col_letter].width = col_width

    def _enable_sheet_view(self, ws):
        """Ensures gridlines are visible and aesthetics are clean."""
        ws.views.sheetView[0].showGridLines = True

    # -------------------------------------------------------------
    # SHEET 1: ACADEMIC OVERVIEW & KPI DASHBOARD
    # -------------------------------------------------------------
    def _create_overview_sheet(self):
        ws = self.wb.active
        ws.title = "Academic Overview"
        self._enable_sheet_view(ws)

        # 1. Main University Header Banner
        ws.merge_cells("A1:H1")
        cell_a1 = ws["A1"]
        cell_a1.value = "BIJU PATNAIK UNIVERSITY OF TECHNOLOGY, ODISHA"
        cell_a1.font = Font(name="Segoe UI", size=15, bold=True, color=self.WHITE)
        cell_a1.fill = PatternFill("solid", fgColor=self.NAVY_PRIMARY)
        cell_a1.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 36

        ws.merge_cells("A2:H2")
        cell_a2 = ws["A2"]
        cell_a2.value = "OFFICIAL CONSOLIDATED STUDENT PERFORMANCE RECORD"
        cell_a2.font = Font(name="Segoe UI", size=10, bold=True, color="94A3B8")
        cell_a2.fill = PatternFill("solid", fgColor=self.NAVY_DARK)
        cell_a2.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[2].height = 20

        # Empty spacer row
        ws.row_dimensions[3].height = 10

        # 2. Student Identity Card Box
        ws.merge_cells("A4:H4")
        sec_hdr = ws["A4"]
        sec_hdr.value = "  STUDENT PROFILE & ACADEMIC REGISTRATION"
        sec_hdr.font = Font(name="Segoe UI", size=11, bold=True, color=self.WHITE)
        sec_hdr.fill = PatternFill("solid", fgColor="2B4C7E")
        sec_hdr.alignment = Alignment(horizontal="left", vertical="center")
        ws.row_dimensions[4].height = 24

        profile_fields = [
            ("Student Name:", self.student_info.get("studentName", "N/A"), "Registration No:", self.student_info.get("rollNo", "N/A")),
            ("Course / Degree:", self.student_info.get("courseName", "N/A"), "Admission Batch:", self.student_info.get("batch", "N/A")),
            ("Branch / Program:", self.student_info.get("branchName", "N/A"), "College Code:", self.student_info.get("collegeCode", "N/A")),
            ("College / Institute:", self.student_info.get("collegeName", "N/A"), "Academic Standing:", self.summary.get("academic_standing", "N/A"))
        ]

        row_idx = 5
        for f1_label, f1_val, f2_label, f2_val in profile_fields:
            ws.row_dimensions[row_idx].height = 22
            # Left Field
            ws.cell(row=row_idx, column=1, value=f1_label).font = Font(name="Segoe UI", size=9.5, bold=True, color="475569")
            ws.cell(row=row_idx, column=1).fill = PatternFill("solid", fgColor=self.ICE_BG)
            ws.cell(row=row_idx, column=1).alignment = Alignment(horizontal="right", vertical="center")

            ws.merge_cells(start_row=row_idx, start_column=2, end_row=row_idx, end_column=4)
            val1_cell = ws.cell(row=row_idx, column=2, value=str(f1_val))
            val1_cell.font = Font(name="Segoe UI", size=9.5, bold=True, color="0F172A")
            val1_cell.fill = PatternFill("solid", fgColor=self.WHITE)
            val1_cell.alignment = Alignment(horizontal="left", vertical="center")

            # Right Field
            ws.cell(row=row_idx, column=5, value=f2_label).font = Font(name="Segoe UI", size=9.5, bold=True, color="475569")
            ws.cell(row=row_idx, column=5).fill = PatternFill("solid", fgColor=self.ICE_BG)
            ws.cell(row=row_idx, column=5).alignment = Alignment(horizontal="right", vertical="center")

            ws.merge_cells(start_row=row_idx, start_column=6, end_row=row_idx, end_column=8)
            val2_cell = ws.cell(row=row_idx, column=6, value=str(f2_val))
            
            # Highlight Academic Standing
            if f2_label == "Academic Standing:":
                standing = str(f2_val).upper()
                if "CLEARED" in standing:
                    val2_cell.font = Font(name="Segoe UI", size=9.5, bold=True, color="0F5132")
                    val2_cell.fill = PatternFill("solid", fgColor="D1E7DD")
                else:
                    val2_cell.font = Font(name="Segoe UI", size=9.5, bold=True, color="842029")
                    val2_cell.fill = PatternFill("solid", fgColor="F8D7DA")
            else:
                val2_cell.font = Font(name="Segoe UI", size=9.5, bold=True, color="0F172A")
                val2_cell.fill = PatternFill("solid", fgColor=self.WHITE)
                
            val2_cell.alignment = Alignment(horizontal="left", vertical="center")

            for col in range(1, 9):
                ws.cell(row=row_idx, column=col).border = self._box_border(color="E2E8F0")

            row_idx += 1

        # Spacer row
        ws.row_dimensions[row_idx].height = 14
        row_idx += 1

        # 3. KPI Metric Summary Cards (4 Cards across columns A-H)
        kpi_row = row_idx
        ws.row_dimensions[kpi_row].height = 20
        ws.row_dimensions[kpi_row + 1].height = 30

        kpis = [
            ("CUMULATIVE CGPA", f"{self.summary.get('cgpa', 0.0):.2f}", "1E40AF", "DBEAFE", 1, 2),
            ("TOTAL CREDITS", str(self.summary.get("total_credits", 0)), "065F46", "D1FAE5", 3, 4),
            ("SEMESTERS RECORDED", str(self.summary.get("total_semesters_found", 0)), "3730A3", "E0E7FF", 5, 6),
            ("ACTIVE BACKLOGS", str(self.summary.get("active_backlog_count", 0)), "991B1B" if self.summary.get("active_backlog_count", 0) > 0 else "15803D", "FEE2E2" if self.summary.get("active_backlog_count", 0) > 0 else "DCFCE7", 7, 8)
        ]

        for title, val, text_color, bg_color, start_c, end_c in kpis:
            ws.merge_cells(start_row=kpi_row, start_column=start_c, end_row=kpi_row, end_column=end_c)
            t_cell = ws.cell(row=kpi_row, column=start_c, value=title)
            t_cell.font = Font(name="Segoe UI", size=8.5, bold=True, color="64748B")
            t_cell.fill = PatternFill("solid", fgColor="F1F5F9")
            t_cell.alignment = Alignment(horizontal="center", vertical="center")

            ws.merge_cells(start_row=kpi_row + 1, start_column=start_c, end_row=kpi_row + 1, end_column=end_c)
            v_cell = ws.cell(row=kpi_row + 1, column=start_c, value=val)
            v_cell.font = Font(name="Segoe UI", size=15, bold=True, color=text_color)
            v_cell.fill = PatternFill("solid", fgColor=bg_color)
            v_cell.alignment = Alignment(horizontal="center", vertical="center")

            for r in [kpi_row, kpi_row + 1]:
                for c in range(start_c, end_c + 1):
                    ws.cell(row=r, column=c).border = self._box_border(color="CBD5E1")

        row_idx += 3

        # 4. Semester Performance Breakdown Table
        ws.merge_cells(start_row=row_idx, start_column=1, end_row=row_idx, end_column=8)
        tbl_hdr = ws.cell(row=row_idx, column=1, value="  SEMESTER-WISE PERFORMANCE BREAKDOWN")
        tbl_hdr.font = Font(name="Segoe UI", size=11, bold=True, color=self.WHITE)
        tbl_hdr.fill = PatternFill("solid", fgColor=self.NAVY_PRIMARY)
        tbl_hdr.alignment = Alignment(horizontal="left", vertical="center")
        ws.row_dimensions[row_idx].height = 24
        row_idx += 1

        headers = ["Semester", "Exam Session", "Exam Type", "Subjects", "Credits", "Grade Points", "SGPA", "Status"]
        col_widths = [14, 22, 22, 12, 12, 14, 12, 16]

        ws.row_dimensions[row_idx].height = 22
        for col_i, (h, w) in enumerate(zip(headers, col_widths), start=1):
            cell = ws.cell(row=row_idx, column=col_i, value=h)
            cell.font = Font(name="Segoe UI", size=9.5, bold=True, color=self.WHITE)
            cell.fill = PatternFill("solid", fgColor="334155")
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = self._box_border(color="64748B")

        row_idx += 1
        start_table_row = row_idx

        for i, sem in enumerate(self.semesters):
            ws.row_dimensions[row_idx].height = 20
            bg = self.ZEBRA_LIGHT if i % 2 == 1 else self.WHITE

            sem_name = str(sem.get("semester", "")).strip()
            if not sem_name.endswith("Semester"):
                sem_name += " Sem" if sem_name else f"Sem {sem.get('semId')}"

            # Status pill styling
            status = sem.get("status", "PASS")
            is_pass = "PASS" in status

            row_data = [
                (sem_name, "center"),
                (sem.get("session", ""), "center"),
                (sem.get("exam_type", "Regular"), "center"),
                (len(sem.get("grades", [])), "center"),
                (sem.get("credits", 0), "center"),
                (round(sem.get("points", 0.0), 2), "right"),
                (f"{sem.get('sgpa', 0.0):.2f}" if sem.get("sgpa") else "N/A", "center"),
                (status, "center")
            ]

            for col_i, (val, align) in enumerate(row_data, start=1):
                cell = ws.cell(row=row_idx, column=col_i, value=val)
                cell.font = Font(name="Segoe UI", size=9.5, bold=(col_i in [1, 7]))
                cell.fill = PatternFill("solid", fgColor=bg)
                cell.alignment = Alignment(horizontal=align, vertical="center")
                cell.border = self._box_border(color="E2E8F0")

                # Format SGPA bold
                if col_i == 7 and val != "N/A":
                    cell.font = Font(name="Segoe UI", size=9.5, bold=True, color="1E3A8A")

                # Format Status
                if col_i == 8:
                    if is_pass:
                        cell.font = Font(name="Segoe UI", size=9.0, bold=True, color="166534")
                        cell.fill = PatternFill("solid", fgColor="DCFCE7")
                    else:
                        cell.font = Font(name="Segoe UI", size=9.0, bold=True, color="991B1B")
                        cell.fill = PatternFill("solid", fgColor="FEE2E2")

            row_idx += 1

        # Table Totals Row
        ws.row_dimensions[row_idx].height = 22
        ws.merge_cells(start_row=row_idx, start_column=1, end_row=row_idx, end_column=3)
        tot_lbl = ws.cell(row=row_idx, column=1, value="CUMULATIVE TOTALS & OVERALL CGPA")
        tot_lbl.font = Font(name="Segoe UI", size=9.5, bold=True, color=self.WHITE)
        tot_lbl.fill = PatternFill("solid", fgColor=self.NAVY_PRIMARY)
        tot_lbl.alignment = Alignment(horizontal="center", vertical="center")

        total_subjects = sum(len(s.get("grades", [])) for s in self.semesters)
        total_credits = self.summary.get("total_credits", 0)
        total_points = self.summary.get("total_grade_points", 0.0)
        cgpa_val = f"{self.summary.get('cgpa', 0.0):.2f}"

        totals_data = [
            (4, total_subjects),
            (5, total_credits),
            (6, total_points),
            (7, cgpa_val),
            (8, self.summary.get("academic_standing", "N/A"))
        ]

        for col_i, val in totals_data:
            cell = ws.cell(row=row_idx, column=col_i, value=val)
            cell.font = Font(name="Segoe UI", size=9.5, bold=True, color=self.WHITE)
            cell.fill = PatternFill("solid", fgColor=self.NAVY_PRIMARY)
            cell.alignment = Alignment(horizontal="center", vertical="center")

        for col in range(1, 9):
            ws.cell(row=row_idx, column=col).border = self._box_border(color="475569")

        self._auto_fit_columns(ws)

    # -------------------------------------------------------------
    # SHEET 2: CONSOLIDATED GRADE SHEET (DETAILED BY SEMESTER)
    # -------------------------------------------------------------
    def _create_detailed_grades_sheet(self):
        ws = self.wb.create_sheet(title="Consolidated Grade Sheet")
        self._enable_sheet_view(ws)

        # Main Header
        ws.merge_cells("A1:I1")
        h1 = ws["A1"]
        h1.value = f"CONSOLIDATED GRADE SHEET — {self.student_info.get('studentName', '')} ({self.student_info.get('rollNo', '')})"
        h1.font = Font(name="Segoe UI", size=13, bold=True, color=self.WHITE)
        h1.fill = PatternFill("solid", fgColor=self.NAVY_PRIMARY)
        h1.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 32

        row_idx = 3

        for sem in self.semesters:
            sem_title = f"SEMESTER {sem.get('semId', '')} ({sem.get('semester', '')}) — {sem.get('session', '')}"
            if sem.get("exam_type") == "Backlog / Improvement":
                sem_title += " [BACKLOG / IMPROVEMENT]"
            
            sgpa_disp = f"SGPA: {sem.get('sgpa', 0.0):.2f}" if sem.get("sgpa") else "SGPA: N/A"
            credits_disp = f"Credits: {sem.get('credits', 0)}"

            # Semester Banner
            ws.merge_cells(start_row=row_idx, start_column=1, end_row=row_idx, end_column=6)
            b_cell = ws.cell(row=row_idx, column=1, value=f"  ▶ {sem_title}")
            b_cell.font = Font(name="Segoe UI", size=10.5, bold=True, color=self.WHITE)
            b_cell.fill = PatternFill("solid", fgColor="1E3A8A")
            b_cell.alignment = Alignment(horizontal="left", vertical="center")

            ws.merge_cells(start_row=row_idx, start_column=7, end_row=row_idx, end_column=9)
            kpi_cell = ws.cell(row=row_idx, column=7, value=f"{credits_disp}   |   {sgpa_disp}")
            kpi_cell.font = Font(name="Segoe UI", size=10, bold=True, color=self.WHITE)
            kpi_cell.fill = PatternFill("solid", fgColor="1E3A8A")
            kpi_cell.alignment = Alignment(horizontal="right", vertical="center")
            ws.row_dimensions[row_idx].height = 24

            for col in range(1, 10):
                ws.cell(row=row_idx, column=col).border = self._box_border(color="1E3A8A")

            row_idx += 1

            # Table Header
            sub_headers = ["Sl.", "Subject Code", "Subject Name", "Type", "Credits", "Grade", "Points", "Credit Points", "Status"]
            ws.row_dimensions[row_idx].height = 20

            for col_i, h in enumerate(sub_headers, start=1):
                cell = ws.cell(row=row_idx, column=col_i, value=h)
                cell.font = Font(name="Segoe UI", size=9, bold=True, color=self.WHITE)
                cell.fill = PatternFill("solid", fgColor="475569")
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.border = self._box_border(color="64748B")

            row_idx += 1

            grades = sem.get("grades", [])
            for sl_no, g in enumerate(grades, start=1):
                ws.row_dimensions[row_idx].height = 19
                bg = self.ZEBRA_LIGHT if sl_no % 2 == 1 else self.WHITE

                sub_tp = str(g.get("subjectTP", "")).strip().upper()
                type_desc = "Theory" if sub_tp == "T" else ("Practical" if sub_tp == "P" else sub_tp)
                grade = str(g.get("grade", "")).strip().upper()
                recheck = int(g.get("recheck", 0) or 0)
                grade_disp = f"{grade}*" if recheck == 1 else grade

                # Status
                if grade in ["F", "M", "S"]:
                    status = "FAILED"
                elif grade == "P":
                    status = "PENDING"
                else:
                    status = "CLEARED"

                cells_data = [
                    (sl_no, "center"),
                    (g.get("subjectCODE", ""), "center"),
                    (g.get("subjectName", ""), "left"),
                    (type_desc, "center"),
                    (g.get("subjectCredits", 0), "center"),
                    (grade_disp, "center"),
                    (g.get("points", 0), "center"),
                    (g.get("creditPoints", 0), "center"),
                    (status, "center")
                ]

                for col_i, (val, align) in enumerate(cells_data, start=1):
                    cell = ws.cell(row=row_idx, column=col_i, value=val)
                    cell.font = Font(name="Segoe UI", size=9, bold=(col_i in [2, 6]))
                    cell.fill = PatternFill("solid", fgColor=bg)
                    cell.alignment = Alignment(horizontal=align, vertical="center")
                    cell.border = self._box_border(color="E2E8F0")

                    # Color-coded Grade Cell
                    if col_i == 6:
                        fill_c, text_c = self.GRADE_STYLES.get(grade, ("FFFFFF", "000000"))
                        cell.fill = PatternFill("solid", fgColor=fill_c)
                        cell.font = Font(name="Segoe UI", size=9.5, bold=True, color=text_c)

                    # Status Cell Styling
                    if col_i == 9:
                        if status == "CLEARED":
                            cell.font = Font(name="Segoe UI", size=8.5, bold=True, color="166534")
                        else:
                            cell.font = Font(name="Segoe UI", size=8.5, bold=True, color="991B1B")

                row_idx += 1

            # Semester Subtotal Footer
            ws.row_dimensions[row_idx].height = 20
            ws.merge_cells(start_row=row_idx, start_column=1, end_row=row_idx, end_column=4)
            ftr_lbl = ws.cell(row=row_idx, column=1, value=f"SEMESTER {sem.get('semId')} TOTALS")
            ftr_lbl.font = Font(name="Segoe UI", size=9, bold=True, color="1E293B")
            ftr_lbl.fill = PatternFill("solid", fgColor="E2E8F0")
            ftr_lbl.alignment = Alignment(horizontal="right", vertical="center")

            # Credits
            c_cell = ws.cell(row=row_idx, column=5, value=sem.get("credits", 0))
            c_cell.font = Font(name="Segoe UI", size=9, bold=True, color="1E293B")
            c_cell.fill = PatternFill("solid", fgColor="E2E8F0")
            c_cell.alignment = Alignment(horizontal="center", vertical="center")

            # SGPA label
            ws.cell(row=row_idx, column=6, value="SGPA:").alignment = Alignment(horizontal="right", vertical="center")
            ws.cell(row=row_idx, column=6).font = Font(name="Segoe UI", size=9, bold=True, color="1E293B")
            ws.cell(row=row_idx, column=6).fill = PatternFill("solid", fgColor="E2E8F0")

            # SGPA value
            sgpa_cell = ws.cell(row=row_idx, column=7, value=f"{sem.get('sgpa', 0.0):.2f}" if sem.get("sgpa") else "N/A")
            sgpa_cell.font = Font(name="Segoe UI", size=9.5, bold=True, color="1E40AF")
            sgpa_cell.fill = PatternFill("solid", fgColor="E2E8F0")
            sgpa_cell.alignment = Alignment(horizontal="center", vertical="center")

            # Points
            pts_cell = ws.cell(row=row_idx, column=8, value=round(sem.get("points", 0.0), 2))
            pts_cell.font = Font(name="Segoe UI", size=9, bold=True, color="1E293B")
            pts_cell.fill = PatternFill("solid", fgColor="E2E8F0")
            pts_cell.alignment = Alignment(horizontal="center", vertical="center")

            # Empty status
            ws.cell(row=row_idx, column=9, value=sem.get("status", "")).alignment = Alignment(horizontal="center", vertical="center")
            ws.cell(row=row_idx, column=9).font = Font(name="Segoe UI", size=8.5, bold=True, color="1E293B")
            ws.cell(row=row_idx, column=9).fill = PatternFill("solid", fgColor="E2E8F0")

            for col in range(1, 10):
                ws.cell(row=row_idx, column=col).border = self._box_border(color="CBD5E1")

            # Spacing between semesters
            row_idx += 2

        self._auto_fit_columns(ws)

    # -------------------------------------------------------------
    # SHEET 3: ALL SUBJECTS MASTER LIST (FLAT TABLE FOR FILTERING)
    # -------------------------------------------------------------
    def _create_master_subjects_sheet(self):
        ws = self.wb.create_sheet(title="All Subjects Master List")
        self._enable_sheet_view(ws)

        headers = [
            "Semester", "Sem ID", "Exam Session", "Exam Type",
            "Subject Code", "Subject Name", "Type", "Credits",
            "Grade", "Grade Points", "Credit Points", "Recheck", "Status"
        ]

        ws.row_dimensions[1].height = 24
        for col_i, h in enumerate(headers, start=1):
            cell = ws.cell(row=1, column=col_i, value=h)
            cell.font = Font(name="Segoe UI", size=9.5, bold=True, color=self.WHITE)
            cell.fill = PatternFill("solid", fgColor=self.NAVY_PRIMARY)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = self._box_border(color="475569")

        row_idx = 2
        for sem in self.semesters:
            sem_name = sem.get("semester", "")
            sem_id = sem.get("semId", "")
            session = sem.get("session", "")
            exam_type = sem.get("exam_type", "")

            for g in sem.get("grades", []):
                ws.row_dimensions[row_idx].height = 19
                bg = self.ZEBRA_LIGHT if row_idx % 2 == 1 else self.WHITE

                sub_tp = str(g.get("subjectTP", "")).strip().upper()
                type_desc = "Theory" if sub_tp == "T" else ("Practical" if sub_tp == "P" else sub_tp)
                grade = str(g.get("grade", "")).strip().upper()
                recheck = "Yes" if int(g.get("recheck", 0) or 0) == 1 else "No"
                status = "FAILED" if grade in ["F", "M", "S"] else ("PENDING" if grade == "P" else "CLEARED")

                row_vals = [
                    (sem_name, "center"),
                    (sem_id, "center"),
                    (session, "center"),
                    (exam_type, "center"),
                    (g.get("subjectCODE", ""), "center"),
                    (g.get("subjectName", ""), "left"),
                    (type_desc, "center"),
                    (g.get("subjectCredits", 0), "center"),
                    (grade, "center"),
                    (g.get("points", 0), "center"),
                    (g.get("creditPoints", 0), "center"),
                    (recheck, "center"),
                    (status, "center")
                ]

                for col_i, (val, align) in enumerate(row_vals, start=1):
                    cell = ws.cell(row=row_idx, column=col_i, value=val)
                    cell.font = Font(name="Segoe UI", size=9, bold=(col_i in [5, 9]))
                    cell.fill = PatternFill("solid", fgColor=bg)
                    cell.alignment = Alignment(horizontal=align, vertical="center")
                    cell.border = self._box_border(color="E2E8F0")

                    if col_i == 9:
                        fill_c, text_c = self.GRADE_STYLES.get(grade, ("FFFFFF", "000000"))
                        cell.fill = PatternFill("solid", fgColor=fill_c)
                        cell.font = Font(name="Segoe UI", size=9.5, bold=True, color=text_c)

                row_idx += 1

        # Enable Auto-filters on all columns
        ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{row_idx - 1}"
        self._auto_fit_columns(ws)

    def export(self, output_path: Optional[str] = None) -> str:
        """
        Renders the complete workbook and writes it to disk.
        Returns the absolute file path of the generated Excel file.
        """
        self._create_overview_sheet()
        self._create_detailed_grades_sheet()
        self._create_master_subjects_sheet()

        if not output_path:
            roll_no = self.student_info.get("rollNo", "Student")
            name = self.student_info.get("studentName", "Result")
            # Sanitize filename for Windows
            safe_name = re.sub(r'[\/\\:\*\?"<>\|]', '_', f"{roll_no}_{name}")
            output_path = f"BPUT_Result_{safe_name}.xlsx"

        output_path = os.path.abspath(output_path)
        self.wb.save(output_path)
        return output_path
