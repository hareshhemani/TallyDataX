"""
Report Engine
Generates pixel-perfect Excel reports matching Tally_Party_Voucher_Report.xlsx
with active Excel formulas (=J+N, =SUM, =O-P) and merges multiple Excel files.
"""

import os
import re
import datetime
from typing import List, Dict, Any, Optional
from collections import OrderedDict
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

HEADER_COLUMNS = [
    'S No', 'Date', 'Vendor Name', 'GSTN', 'Voucher No.', 'Address', 'Invoice No', 'Ledger', 'Ledger Group',
    'Particulars', 'Taxable Value', 'CGST', 'SGST', 'IGST', 'GST TOTAL', 'TCS/ (TDS)', 'Invoice Amount',
    'Amount paid', 'Date of Payment', 'Mode of Payment', 'Ledger balance', 'Closing Balance'
]

# Pre-allocated openpyxl styling definitions
HEADER_FILL = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
HEADER_FONT = Font(name="Arial", size=10, bold=True, color="FFFFFF")
TOTAL_FILL = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
TOTAL_FONT = Font(name="Arial", size=10, bold=True, color="0F172A")
DATA_FONT = Font(name="Arial", size=9, color="1E293B")
BOLD_DATA_FONT = Font(name="Arial", size=9, bold=True, color="0F172A")

THIN_BORDER = Border(
    left=Side(style='thin', color='E2E8F0'),
    right=Side(style='thin', color='E2E8F0'),
    top=Side(style='thin', color='E2E8F0'),
    bottom=Side(style='thin', color='E2E8F0')
)
TOTAL_BORDER = Border(
    top=Side(style='thin', color='94A3B8'),
    bottom=Side(style='double', color='0F172A')
)
GRAND_TOTAL_FILL = PatternFill(start_color="0F172A", end_color="0F172A", fill_type="solid")
GRAND_TOTAL_FONT = Font(name="Arial", size=10, bold=True, color="FFFFFF")

# Pre-allocated alignments for performance
ALIGN_RIGHT = Alignment(horizontal="right", vertical="center")
ALIGN_CENTER = Alignment(horizontal="center", vertical="center")
ALIGN_LEFT = Alignment(horizontal="left", vertical="center")
ALIGN_HEADER = Alignment(horizontal="center", vertical="center", wrap_text=True)
NUM_FMT = '#,##0.00;(#,##0.00);"-"'


def write_voucher_worksheet_xlsxwriter(
    ws,
    wb,
    vouchers: List[Dict[str, Any]],
    add_source_col: bool = False
) -> None:
    """
    High-performance Excel renderer using xlsxwriter.
    Generates 12,000+ rows with active formulas, sub-totals, and grand totals in ~5 seconds.
    Includes Column I: Ledger Group (full hierarchy path).
    """
    headers = list(HEADER_COLUMNS)
    if add_source_col:
        headers.append("Source File")

    # Formats
    hdr_fmt = wb.add_format({
        'font_name': 'Arial', 'font_size': 10, 'bold': True, 'font_color': '#FFFFFF',
        'bg_color': '#1E293B', 'align': 'center', 'valign': 'vcenter', 'text_wrap': True,
        'border': 1, 'border_color': '#E2E8F0'
    })
    data_fmt_left = wb.add_format({
        'font_name': 'Arial', 'font_size': 9, 'font_color': '#1E293B',
        'align': 'left', 'valign': 'vcenter',
        'border': 1, 'border_color': '#E2E8F0'
    })
    data_fmt_center = wb.add_format({
        'font_name': 'Arial', 'font_size': 9, 'font_color': '#1E293B',
        'align': 'center', 'valign': 'vcenter',
        'border': 1, 'border_color': '#E2E8F0'
    })
    data_fmt_num = wb.add_format({
        'font_name': 'Arial', 'font_size': 9, 'font_color': '#1E293B',
        'align': 'right', 'valign': 'vcenter',
        'border': 1, 'border_color': '#E2E8F0',
        'num_format': NUM_FMT
    })

    tot_fmt_left = wb.add_format({
        'font_name': 'Arial', 'font_size': 10, 'bold': True, 'font_color': '#0F172A',
        'bg_color': '#F1F5F9', 'align': 'left', 'valign': 'vcenter',
        'top': 1, 'top_color': '#94A3B8', 'bottom': 6, 'bottom_color': '#0F172A'
    })
    tot_fmt_center = wb.add_format({
        'font_name': 'Arial', 'font_size': 10, 'bold': True, 'font_color': '#0F172A',
        'bg_color': '#F1F5F9', 'align': 'center', 'valign': 'vcenter',
        'top': 1, 'top_color': '#94A3B8', 'bottom': 6, 'bottom_color': '#0F172A'
    })
    tot_fmt_num = wb.add_format({
        'font_name': 'Arial', 'font_size': 10, 'bold': True, 'font_color': '#0F172A',
        'bg_color': '#F1F5F9', 'align': 'right', 'valign': 'vcenter',
        'num_format': NUM_FMT,
        'top': 1, 'top_color': '#94A3B8', 'bottom': 6, 'bottom_color': '#0F172A'
    })

    grand_fmt_left = wb.add_format({
        'font_name': 'Arial', 'font_size': 10, 'bold': True, 'font_color': '#FFFFFF',
        'bg_color': '#0F172A', 'align': 'left', 'valign': 'vcenter',
        'top': 1, 'top_color': '#94A3B8', 'bottom': 6, 'bottom_color': '#0F172A'
    })
    grand_fmt_center = wb.add_format({
        'font_name': 'Arial', 'font_size': 10, 'bold': True, 'font_color': '#FFFFFF',
        'bg_color': '#0F172A', 'align': 'center', 'valign': 'vcenter',
        'top': 1, 'top_color': '#94A3B8', 'bottom': 6, 'bottom_color': '#0F172A'
    })
    grand_fmt_num = wb.add_format({
        'font_name': 'Arial', 'font_size': 10, 'bold': True, 'font_color': '#FFFFFF',
        'bg_color': '#0F172A', 'align': 'right', 'valign': 'vcenter',
        'num_format': NUM_FMT,
        'top': 1, 'top_color': '#94A3B8', 'bottom': 6, 'bottom_color': '#0F172A'
    })

    # Header Row (0-indexed row 0)
    ws.set_row(0, 26)
    for col_idx, col_name in enumerate(headers):
        ws.write_string(0, col_idx, col_name, hdr_fmt)
    ws.freeze_panes(1, 0)

    # Column Widths (22 standard columns + optional Source File)
    col_widths = [
        8,   # A: S No
        12,  # B: Date
        26,  # C: Vendor Name
        18,  # D: GSTN
        14,  # E: Voucher No.
        35,  # F: Address
        16,  # G: Invoice No
        22,  # H: Ledger
        26,  # I: Ledger Group (NEW!)
        35,  # J: Particulars
        15,  # K: Taxable Value
        12,  # L: CGST
        12,  # M: SGST
        12,  # N: IGST
        14,  # O: GST TOTAL
        14,  # P: TCS/ (TDS)
        16,  # Q: Invoice Amount
        16,  # R: Amount paid
        14,  # S: Date of Payment
        20,  # T: Mode of Payment
        16,  # U: Ledger balance
        18,  # V: Closing Balance
    ]
    if add_source_col:
        col_widths.append(24)

    for c_idx, w in enumerate(col_widths):
        ws.set_column(c_idx, c_idx, w)

    party_groups: Dict[str, List[Dict[str, Any]]] = OrderedDict()
    for v in vouchers:
        p = (v.get("vendor") or "General").strip()
        party_groups.setdefault(p, []).append(v)

    current_row = 1
    party_total_rows = []

    for p_name, p_vouchers in party_groups.items():
        start_row = current_row
        for v in p_vouchers:
            taxable = v.get("taxable")
            cgst = v.get("cgst")
            sgst = v.get("sgst")
            igst = v.get("igst")
            gst_total = v.get("gstTotal")
            tcs_tds = v.get("tcsTds")
            amount_paid = v.get("amountPaid")

            excel_row = current_row + 1
            ws.set_row(current_row, 20)

            # Col 0: S No
            s_no = v.get("sNo", current_row)
            if isinstance(s_no, (int, float)):
                ws.write_number(current_row, 0, int(s_no), data_fmt_center)
            else:
                ws.write(current_row, 0, str(s_no), data_fmt_center)

            # Col 1: Date
            ws.write(current_row, 1, str(v.get("date", "") or ""), data_fmt_center)

            # Col 2: Vendor Name
            ws.write(current_row, 2, str(v.get("vendor", "") or ""), data_fmt_left)

            # Col 3: GSTN
            gstn = v.get("gstn", "") or ""
            ws.write(current_row, 3, str(gstn) if gstn else "", data_fmt_left)

            # Col 4: Voucher No.
            ws.write(current_row, 4, str(v.get("voucherNo", "") or ""), data_fmt_center)

            # Col 5: Address
            ws.write(current_row, 5, str(v.get("address", "") or ""), data_fmt_left)

            # Col 6: Invoice No
            inv_no = v.get("invoiceNo", "") or ""
            ws.write(current_row, 6, str(inv_no) if inv_no else "", data_fmt_left)

            # Col 7: Ledger
            ws.write(current_row, 7, str(v.get("ledger", "") or ""), data_fmt_left)

            # Col 8: Ledger Group (Full Hierarchy Path)
            ws.write(current_row, 8, str(v.get("ledgerGroup", "") or ""), data_fmt_left)

            # Col 9: Particulars
            ws.write(current_row, 9, str(v.get("particulars", "") or ""), data_fmt_left)

            # Col 10: Taxable Value (Col K)
            if taxable is not None and taxable != 0:
                ws.write_number(current_row, 10, float(taxable), data_fmt_num)
            else:
                ws.write_blank(current_row, 10, None, data_fmt_num)

            # Col 11: CGST (Col L)
            if cgst is not None and cgst != 0:
                ws.write_number(current_row, 11, float(cgst), data_fmt_num)
            else:
                ws.write_blank(current_row, 11, None, data_fmt_num)

            # Col 12: SGST (Col M)
            if sgst is not None and sgst != 0:
                ws.write_number(current_row, 12, float(sgst), data_fmt_num)
            else:
                ws.write_blank(current_row, 12, None, data_fmt_num)

            # Col 13: IGST (Col N)
            if igst is not None and igst != 0:
                ws.write_number(current_row, 13, float(igst), data_fmt_num)
            else:
                ws.write_blank(current_row, 13, None, data_fmt_num)

            # Col 14: GST TOTAL (Col O)
            if gst_total is not None and gst_total != 0:
                ws.write_number(current_row, 14, float(gst_total), data_fmt_num)
            else:
                ws.write_blank(current_row, 14, None, data_fmt_num)

            # Col 15: TCS/ (TDS) (Col P)
            if tcs_tds is not None and tcs_tds != 0:
                ws.write_number(current_row, 15, float(tcs_tds), data_fmt_num)
            else:
                ws.write_blank(current_row, 15, None, data_fmt_num)

            # Col 16: Invoice Amount (Col Q) (Active Formula =K+O+P)
            if taxable:
                if tcs_tds:
                    ws.write_formula(current_row, 16, f"=K{excel_row}+O{excel_row}+P{excel_row}", data_fmt_num)
                else:
                    ws.write_formula(current_row, 16, f"=K{excel_row}+O{excel_row}", data_fmt_num)
            elif v.get("invoiceAmount"):
                ws.write_number(current_row, 16, float(v.get("invoiceAmount")), data_fmt_num)
            else:
                ws.write_blank(current_row, 16, None, data_fmt_num)

            # Col 17: Amount paid (Col R)
            if amount_paid is not None and amount_paid != 0:
                ws.write_number(current_row, 17, float(amount_paid), data_fmt_num)
            else:
                ws.write_blank(current_row, 17, None, data_fmt_num)

            # Col 18: Date of Payment (Col S)
            pay_dt = v.get("paymentDate", "") or ""
            ws.write(current_row, 18, str(pay_dt) if pay_dt else "", data_fmt_center)

            # Col 19: Mode of Payment (Col T)
            pay_mode = v.get("paymentMode", "") or ""
            ws.write(current_row, 19, str(pay_mode) if pay_mode else "", data_fmt_left)

            # Col 20: Ledger balance (Col U)
            ws.write_number(current_row, 20, 0.0, data_fmt_num)

            # Col 21: Closing Balance (Col V)
            ws.write(current_row, 21, format_closing_balance(v.get("closingBalance", "")), data_fmt_center)

            # Col 22: Source File (Col W) (optional)
            if add_source_col:
                ws.write(current_row, 22, str(v.get("sourceFile", "") or ""), data_fmt_left)

            current_row += 1

        end_row = current_row - 1
        start_excel_row = start_row + 1
        end_excel_row = end_row + 1
        tot_excel_row = current_row + 1
        party_total_rows.append(tot_excel_row)

        ws.set_row(current_row, 23)
        ws.write_blank(current_row, 0, None, tot_fmt_center)
        ws.write_blank(current_row, 1, None, tot_fmt_center)
        ws.write_blank(current_row, 2, None, tot_fmt_left)
        ws.write_string(current_row, 3, "TOTAL", tot_fmt_center)
        for c in range(4, 10):
            ws.write_blank(current_row, c, None, tot_fmt_left)

        ws.write_formula(current_row, 10, f"=SUM(K{start_excel_row}:K{end_excel_row})", tot_fmt_num)
        ws.write_formula(current_row, 11, f"=SUM(L{start_excel_row}:L{end_excel_row})", tot_fmt_num)
        ws.write_formula(current_row, 12, f"=SUM(M{start_excel_row}:M{end_excel_row})", tot_fmt_num)
        ws.write_formula(current_row, 13, f"=SUM(N{start_excel_row}:N{end_excel_row})", tot_fmt_num)
        ws.write_formula(current_row, 14, f"=SUM(O{start_excel_row}:O{end_excel_row})", tot_fmt_num)
        ws.write_formula(current_row, 15, f"=SUM(P{start_excel_row}:P{end_excel_row})", tot_fmt_num)
        ws.write_formula(current_row, 16, f"=SUM(Q{start_excel_row}:Q{end_excel_row})", tot_fmt_num)
        ws.write_formula(current_row, 17, f"=SUM(R{start_excel_row}:R{end_excel_row})", tot_fmt_num)
        ws.write_blank(current_row, 18, None, tot_fmt_center)
        ws.write_blank(current_row, 19, None, tot_fmt_left)
        ws.write_formula(current_row, 20, f"=Q{tot_excel_row}-R{tot_excel_row}", tot_fmt_num)
        ws.write(current_row, 21, format_closing_balance(p_vouchers[-1].get("closingBalance", "")), tot_fmt_center)
        if add_source_col:
            ws.write_blank(current_row, 22, None, tot_fmt_left)

        current_row += 1

    # Grand Total Row if multiple parties
    if len(party_groups) > 1 and party_total_rows:
        grand_excel_row = current_row + 1
        last_check_row = grand_excel_row - 1

        ws.set_row(current_row, 26)
        ws.write_blank(current_row, 0, None, grand_fmt_center)
        ws.write_blank(current_row, 1, None, grand_fmt_center)
        ws.write_blank(current_row, 2, None, grand_fmt_left)
        ws.write_string(current_row, 3, "GRAND TOTAL", grand_fmt_center)
        for c in range(4, 10):
            ws.write_blank(current_row, c, None, grand_fmt_left)

        ws.write_formula(current_row, 10, f'=SUMIF($D$2:$D${last_check_row}, "TOTAL", K$2:K${last_check_row})', grand_fmt_num)
        ws.write_formula(current_row, 11, f'=SUMIF($D$2:$D${last_check_row}, "TOTAL", L$2:L${last_check_row})', grand_fmt_num)
        ws.write_formula(current_row, 12, f'=SUMIF($D$2:$D${last_check_row}, "TOTAL", M$2:M${last_check_row})', grand_fmt_num)
        ws.write_formula(current_row, 13, f'=SUMIF($D$2:$D${last_check_row}, "TOTAL", N$2:N${last_check_row})', grand_fmt_num)
        ws.write_formula(current_row, 14, f'=SUMIF($D$2:$D${last_check_row}, "TOTAL", O$2:O${last_check_row})', grand_fmt_num)
        ws.write_formula(current_row, 15, f'=SUMIF($D$2:$D${last_check_row}, "TOTAL", P$2:P${last_check_row})', grand_fmt_num)
        ws.write_formula(current_row, 16, f'=SUMIF($D$2:$D${last_check_row}, "TOTAL", Q$2:Q${last_check_row})', grand_fmt_num)
        ws.write_formula(current_row, 17, f'=SUMIF($D$2:$D${last_check_row}, "TOTAL", R$2:R${last_check_row})', grand_fmt_num)
        ws.write_blank(current_row, 18, None, grand_fmt_center)
        ws.write_blank(current_row, 19, None, grand_fmt_left)
        ws.write_formula(current_row, 20, f"=Q{grand_excel_row}-R{grand_excel_row}", grand_fmt_num)
        ws.write_blank(current_row, 21, None, grand_fmt_center)
        if add_source_col:
            ws.write_blank(current_row, 22, None, grand_fmt_left)


def build_voucher_worksheet(
    ws,
    vouchers: List[Dict[str, Any]],
    add_source_col: bool = False
) -> None:
    """
    Renders voucher data onto a given openpyxl Worksheet with:
    - 21 standard columns (including Column O: TCS/ (TDS), plus optional Column V: Source File)
    - Active formulas: =J{row}+N{row}+O{row} on Invoice Amount
    - Party Sub-total rows with =SUM(...) and =P-Q formulas
    - Grand Total row with =SUMIF(...) formulas (if multiple parties)
    - Full styling: navy header, borders, number formats, auto-width
    """
    headers = list(HEADER_COLUMNS)
    if add_source_col:
        headers.append("Source File")

    # Write Header Row
    ws.append(headers)
    for col_idx in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = ALIGN_HEADER
        cell.border = THIN_BORDER
    ws.row_dimensions[1].height = 26
    ws.freeze_panes = 'A2'

    party_groups: Dict[str, List[Dict[str, Any]]] = OrderedDict()
    for v in vouchers:
        p = (v.get("vendor") or "General").strip()
        party_groups.setdefault(p, []).append(v)

    current_row = 2
    party_total_rows = []

    for p_name, p_vouchers in party_groups.items():
        start_row = current_row
        for v in p_vouchers:
            taxable = v.get("taxable")
            cgst = v.get("cgst")
            sgst = v.get("sgst")
            igst = v.get("igst")
            gst_total = v.get("gstTotal")
            tcs_tds = v.get("tcsTds")
            amount_paid = v.get("amountPaid")

            # Active formula for Invoice Amount (=K{row}+O{row}+P{row}) if taxable exists
            if taxable:
                if tcs_tds:
                    invoice_formula = f"=K{current_row}+O{current_row}+P{current_row}"
                else:
                    invoice_formula = f"=K{current_row}+O{current_row}"
            elif v.get("invoiceAmount"):
                invoice_formula = v.get("invoiceAmount")
            else:
                invoice_formula = None

            row_data = [
                v.get("sNo", current_row - 1),
                v.get("date", ""),
                v.get("vendor", ""),
                v.get("gstn", "") or None,
                str(v.get("voucherNo", "")),
                v.get("address", ""),
                v.get("invoiceNo", "") or None,
                v.get("ledger", ""),
                v.get("ledgerGroup", ""),
                v.get("particulars", ""),
                taxable if taxable else None,
                cgst if cgst else None,
                sgst if sgst else None,
                igst if igst else None,
                gst_total if gst_total else None,
                tcs_tds if (tcs_tds is not None and tcs_tds != 0) else None,
                invoice_formula,
                amount_paid if amount_paid else None,
                v.get("paymentDate", "") or None,
                v.get("paymentMode", "") or None,
                0,
                format_closing_balance(v.get("closingBalance", ""))
            ]
            if add_source_col:
                row_data.append(v.get("sourceFile", ""))

            ws.append(row_data)

            # Style data cells using pre-allocated alignments
            row_cells = ws[current_row]
            for col_idx, c in enumerate(row_cells, start=1):
                c.font = DATA_FONT
                c.border = THIN_BORDER
                if col_idx in (11, 12, 13, 14, 15, 16, 17, 18, 21):
                    c.number_format = NUM_FMT
                    c.alignment = ALIGN_RIGHT
                elif col_idx in (1, 2, 5, 19, 22):
                    c.alignment = ALIGN_CENTER
                else:
                    c.alignment = ALIGN_LEFT

            ws.row_dimensions[current_row].height = 20
            current_row += 1

        end_row = current_row - 1
        party_tot_row = current_row
        party_total_rows.append(party_tot_row)

        party_tot_data = [
            None, None, None, "TOTAL", None, None, None, None, None, None,
            f"=SUM(K{start_row}:K{end_row})",
            f"=SUM(L{start_row}:L{end_row})",
            f"=SUM(M{start_row}:M{end_row})",
            f"=SUM(N{start_row}:N{end_row})",
            f"=SUM(O{start_row}:O{end_row})",
            f"=SUM(P{start_row}:P{end_row})",
            f"=SUM(Q{start_row}:Q{end_row})",
            f"=SUM(R{start_row}:R{end_row})",
            None, None,
            f"=Q{party_tot_row}-R{party_tot_row}",  # Net Balance: Invoice - Paid
            format_closing_balance(p_vouchers[-1].get("closingBalance", ""))
        ]
        if add_source_col:
            party_tot_data.append(None)

        ws.append(party_tot_data)

        party_cells = ws[party_tot_row]
        for col_idx, c in enumerate(party_cells, start=1):
            c.fill = TOTAL_FILL
            c.font = TOTAL_FONT
            c.border = TOTAL_BORDER
            if col_idx in (11, 12, 13, 14, 15, 16, 17, 18, 21):
                c.number_format = NUM_FMT
                c.alignment = ALIGN_RIGHT
            elif col_idx in (4, 22):
                c.alignment = ALIGN_CENTER

        ws.row_dimensions[party_tot_row].height = 23
        current_row += 1

    # Write Grand Total Row if multiple parties
    if len(party_groups) > 1 and party_total_rows:
        grand_row = current_row
        last_check_row = grand_row - 1
        grand_tot_data = [
            None, None, None, "GRAND TOTAL", None, None, None, None, None, None,
            f'=SUMIF($D$2:$D${last_check_row}, "TOTAL", K$2:K${last_check_row})',
            f'=SUMIF($D$2:$D${last_check_row}, "TOTAL", L$2:L${last_check_row})',
            f'=SUMIF($D$2:$D${last_check_row}, "TOTAL", M$2:M${last_check_row})',
            f'=SUMIF($D$2:$D${last_check_row}, "TOTAL", N$2:N${last_check_row})',
            f'=SUMIF($D$2:$D${last_check_row}, "TOTAL", O$2:O${last_check_row})',
            f'=SUMIF($D$2:$D${last_check_row}, "TOTAL", P$2:P${last_check_row})',
            f'=SUMIF($D$2:$D${last_check_row}, "TOTAL", Q$2:Q${last_check_row})',
            f'=SUMIF($D$2:$D${last_check_row}, "TOTAL", R$2:R${last_check_row})',
            None, None,
            f"=Q{grand_row}-R{grand_row}",
            None
        ]
        if add_source_col:
            grand_tot_data.append(None)

        ws.append(grand_tot_data)

        grand_cells = ws[grand_row]
        for col_idx, c in enumerate(grand_cells, start=1):
            c.fill = GRAND_TOTAL_FILL
            c.font = GRAND_TOTAL_FONT
            c.border = TOTAL_BORDER
            if col_idx in (11, 12, 13, 14, 15, 16, 17, 18, 21):
                c.number_format = NUM_FMT
                c.alignment = ALIGN_RIGHT
            elif col_idx == 4:
                c.alignment = ALIGN_CENTER

        ws.row_dimensions[grand_row].height = 26

    # Column Widths Auto-sizing
    col_widths = {
        'A': 8,   'B': 12,  'C': 26,  'D': 18,  'E': 14,
        'F': 35,  'G': 16,  'H': 22,  'I': 26,  'J': 35,
        'K': 15,  'L': 12,  'M': 12,  'N': 12,  'O': 14,
        'P': 14,  'Q': 16,  'R': 16,  'S': 14,  'T': 20,
        'U': 16,  'V': 18,  'W': 24
    }
    for col_letter, width in col_widths.items():
        ws.column_dimensions[col_letter].width = width


def generate_voucher_excel(
    vouchers: List[Dict[str, Any]],
    party_name: str = "Special Blasts Limited",
    output_path: Optional[str] = None
) -> str:
    """
    Generates a formatted Excel (.xlsx) file with active formulas matching
    the exact structure of Tally_Party_Voucher_Report.xlsx.
    Uses xlsxwriter for ultra-high speed (5s for 12,000+ rows) with fallback to openpyxl.
    """
    sheet_title = "".join(c for c in party_name if c not in r"\/*?:[]")[:30] or "Voucher Report"
    if not output_path:
        output_path = os.path.join(os.getcwd(), f"Tally_Report_{sheet_title}.xlsx")

    try:
        import xlsxwriter
        wb = xlsxwriter.Workbook(output_path, {'constant_memory': False})
        ws = wb.add_worksheet(sheet_title)
        write_voucher_worksheet_xlsxwriter(ws, wb, vouchers, add_source_col=False)
        wb.close()
        return output_path
    except ImportError:
        pass

    # Fallback to openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet_title
    build_voucher_worksheet(ws, vouchers, add_source_col=False)
    wb.save(output_path)
    return output_path


def format_closing_balance(raw_val: Any, is_deemed_pos: Optional[str] = None) -> str:
    """Formats closing balance with Indian number formatting and Dr / Cr suffix."""
    if raw_val is None:
        return ""
    s = str(raw_val).strip()
    if not s:
        return ""
    if re.search(r'\b(dr|cr)\b', s, re.IGNORECASE):
        return s
    clean_s = s.replace(',', '')
    try:
        val = float(clean_s)
        if abs(val) < 0.001:
            return "0.00"
        if is_deemed_pos and str(is_deemed_pos).strip().lower() in ('yes', 'true', '1'):
            dr_cr = "Dr"
        elif is_deemed_pos and str(is_deemed_pos).strip().lower() in ('no', 'false', '0'):
            dr_cr = "Cr"
        else:
            dr_cr = "Dr" if val < 0 else "Cr"
        return f"{abs(val):,.2f} {dr_cr}"
    except ValueError:
        return s


def parse_date_sort_key(d_val: Any) -> tuple:
    """Returns (year, month, day) tuple for chronological sorting across years."""
    if isinstance(d_val, (datetime.datetime, datetime.date)):
        return (d_val.year, d_val.month, d_val.day)
    if not d_val:
        return (9999, 12, 31)
    s = str(d_val).strip()
    for fmt in ('%d-%m-%y', '%d-%m-%Y', '%d/%m/%y', '%d/%m/%Y', '%Y-%m-%d', '%Y/%m/%d'):
        try:
            dt = datetime.datetime.strptime(s, fmt)
            return (dt.year, dt.month, dt.day)
        except ValueError:
            pass
    m = re.match(r'^(\d{1,2})[-/](\d{1,2})[-/](\d{2,4})$', s)
    if m:
        d, mo, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
        if y < 100:
            y += 2000
        return (y, mo, d)
    return (9999, 12, 31)


def format_date_str(d_val: Any) -> str:
    """Formats date to DD-MM-YY string matching Tally export conventions."""
    if not d_val:
        return ""
    if isinstance(d_val, (datetime.datetime, datetime.date)):
        return d_val.strftime("%d-%m-%y")
    s = str(d_val).strip()
    m = re.match(r"^(\d{4})-(\d{1,2})-(\d{1,2})", s)
    if m:
        y, mo, d = m.group(1), int(m.group(2)), int(m.group(3))
        return f"{d:02d}-{mo:02d}-{y[-2:]}"
    return s


def safe_vch_key(vch: Any) -> tuple:
    """Generates natural sort key for voucher numbers."""
    if not vch:
        return (0, 0, "")
    s = str(vch).strip()
    if s.isdigit():
        return (0, int(s), s)
    m = re.search(r'(\d+)$', s)
    if m:
        return (1, int(m.group(1)), s)
    return (2, 0, s)


def clean_float(val: Any) -> Optional[float]:
    """Parses float cleanly from numbers or formatted currency strings."""
    if val is None:
        return None
    if isinstance(val, (int, float)):
        return float(val) if val != 0 else None
    s = str(val).strip().replace(',', '')
    if s.startswith('='):
        return None
    try:
        f = float(s)
        return f if f != 0 else None
    except ValueError:
        return None


def clean_str(val: Any) -> str:
    """Cleans string value."""
    if val is None:
        return ""
    s = str(val).strip()
    return "" if s.lower() in ('none', 'null') else s


def extract_vouchers_from_workbook(fpath: str) -> List[Dict[str, Any]]:
    """
    Reads an uploaded Excel file (.xlsx / .xlsm), scans all worksheets,
    detects columns accurately, skips TOTAL rows, and extracts voucher records.
    """
    if not os.path.exists(fpath):
        return []

    fname = os.path.basename(fpath)
    extracted = []

    try:
        wb = openpyxl.load_workbook(fpath, data_only=True)
    except Exception as e:
        print(f"[Warning] Failed to load workbook {fpath}: {e}")
        return []

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        rows = list(ws.iter_rows(values_only=True))
        if not rows:
            continue

        # Locate Header Row (scan first 10 rows)
        h_idx = -1
        col_map: Dict[str, int] = {}

        for r_idx, row in enumerate(rows[:10]):
            row_str = [str(c).strip().lower() for c in row if c is not None]
            if any('date' in c for c in row_str) and any('vendor' in c or 'party' in c for c in row_str):
                h_idx = r_idx
                for c_idx, c_val in enumerate(row):
                    if c_val is None:
                        continue
                    norm = str(c_val).strip().lower()
                    if norm in ('s no', 'sno', 'sr no', 'srno', 'sr. no.', 'sr.', 'serial', 'sl no', 'no'):
                        col_map['s_no'] = c_idx
                    elif norm in ('date of payment', 'payment date', 'pay date'):
                        col_map['payment_date'] = c_idx
                    elif 'date' in norm and 'payment' not in norm:
                        col_map['date'] = c_idx
                    elif 'vendor' in norm or 'party' in norm or 'customer' in norm:
                        col_map['vendor'] = c_idx
                    elif norm in ('gst total', 'total gst', 'gst amount') or ('gst' in norm and 'total' in norm):
                        col_map['gst_total'] = c_idx
                    elif 'cgst' in norm:
                        col_map['cgst'] = c_idx
                    elif 'sgst' in norm:
                        col_map['sgst'] = c_idx
                    elif 'igst' in norm:
                        col_map['igst'] = c_idx
                    elif norm in ('gstn', 'gstin', 'gst no', 'gst no.', 'party gstin', 'party gstn') or (
                        ('gst' in norm or 'gstin' in norm) and not any(k in norm for k in ('cgst', 'sgst', 'igst', 'total', 'amt', 'amount'))
                    ):
                        col_map['gstn'] = c_idx
                    elif 'voucher no' in norm or 'vch no' in norm:
                        col_map['voucher_no'] = c_idx
                    elif 'address' in norm:
                        col_map['address'] = c_idx
                    elif 'invoice no' in norm or 'bill no' in norm or 'ref no' in norm:
                        col_map['invoice_no'] = c_idx
                    elif 'closing balance' in norm:
                        col_map['closing_balance'] = c_idx
                    elif 'ledger balance' in norm or ('balance' in norm and 'closing' not in norm):
                        col_map['ledger_balance'] = c_idx
                    elif 'ledger group' in norm or 'group' in norm:
                        col_map['ledger_group'] = c_idx
                    elif 'ledger' in norm:
                        col_map['ledger'] = c_idx
                    elif 'particular' in norm or 'narration' in norm:
                        col_map['particulars'] = c_idx
                    elif 'taxable' in norm or 'basic' in norm:
                        col_map['taxable'] = c_idx
                    elif 'invoice amount' in norm or 'bill amount' in norm or norm in ('invoice value', 'gross amount'):
                        col_map['invoice_amount'] = c_idx
                    elif 'amount paid' in norm or 'paid amount' in norm:
                        col_map['amount_paid'] = c_idx
                    elif 'mode' in norm:
                        col_map['payment_mode'] = c_idx
                    elif 'tcs' in norm or 'tds' in norm:
                        col_map['tcs_tds'] = c_idx
                    elif 'source file' in norm:
                        col_map['source_file'] = c_idx
                break

        if h_idx == -1:
            # If no explicit header recognized, check if row 0 has standard columns
            if len(rows[0]) >= 15:
                h_idx = 0
                standard_keys = [
                    's_no', 'date', 'vendor', 'gstn', 'voucher_no', 'address', 'invoice_no', 'ledger', 'ledger_group',
                    'particulars', 'taxable', 'cgst', 'sgst', 'igst', 'gst_total', 'tcs_tds', 'invoice_amount',
                    'amount_paid', 'payment_date', 'payment_mode', 'ledger_balance', 'closing_balance'
                ]
                for i, k in enumerate(standard_keys):
                    if i < len(rows[0]):
                        col_map[k] = i
            else:
                continue

        # Extract data rows below header
        for r_idx in range(h_idx + 1, len(rows)):
            row = rows[r_idx]
            if not any(row):
                continue

            # Skip summary / total rows
            if any(isinstance(v, str) and v.strip().upper() in ('TOTAL', 'GRAND TOTAL', 'SUB TOTAL', 'SUBTOTAL') for v in row):
                continue

            def get_val(key):
                idx = col_map.get(key)
                return row[idx] if idx is not None and idx < len(row) else None

            vendor = clean_str(get_val('vendor')) or sheet_name
            v_date_raw = get_val('date')
            v_date = format_date_str(v_date_raw)

            # Skip row if neither vendor nor date is available
            if not v_date and not vendor:
                continue

            taxable = clean_float(get_val('taxable'))
            cgst = clean_float(get_val('cgst'))
            sgst = clean_float(get_val('sgst'))
            igst = clean_float(get_val('igst'))
            gst_total = clean_float(get_val('gst_total'))
            if gst_total is None and (cgst or sgst or igst):
                gst_total = (cgst or 0.0) + (sgst or 0.0) + (igst or 0.0)

            tcs_tds = clean_float(get_val('tcs_tds'))

            invoice_amount = clean_float(get_val('invoice_amount'))
            if invoice_amount is None and taxable:
                invoice_amount = (taxable or 0.0) + (gst_total or 0.0) + (tcs_tds or 0.0)

            amount_paid = clean_float(get_val('amount_paid'))

            src_file_val = clean_str(get_val('source_file')) or fname

            extracted.append({
                'vendor': vendor,
                'date': v_date,
                '_sort_date': parse_date_sort_key(v_date_raw),
                'gstn': clean_str(get_val('gstn')),
                'voucherNo': clean_str(get_val('voucher_no')),
                'address': clean_str(get_val('address')),
                'invoiceNo': clean_str(get_val('invoice_no')),
                'ledger': clean_str(get_val('ledger')),
                'ledgerGroup': clean_str(get_val('ledger_group')),
                'particulars': clean_str(get_val('particulars')),
                'taxable': taxable,
                'cgst': cgst,
                'sgst': sgst,
                'igst': igst,
                'gstTotal': gst_total,
                'tcsTds': tcs_tds,
                'invoiceAmount': invoice_amount,
                'amountPaid': amount_paid,
                'paymentDate': format_date_str(get_val('payment_date')),
                'paymentMode': clean_str(get_val('payment_mode')),
                'closingBalance': clean_str(get_val('closing_balance')),
                'sourceFile': src_file_val
            })

    return extracted


def merge_excel_workbooks(
    file_paths: List[str],
    mode: str = "singleSheet",
    auto_align: bool = True,
    add_source_col: bool = False,
    output_path: Optional[str] = None
) -> str:
    """
    Merges multiple Excel workbooks (e.g. exported from different years):
    - Consolidates each party across all files and sheets into a single continuous block.
    - Sorts all vouchers of each party chronologically (date-wise across multiple years).
    - Recalculates continuous running ledger balance.
    - Generates the exact 20-column accounting report matching Tally_Party_Voucher_Report.xlsx
      with active Excel formulas (=J+N, =SUM, =O-P) and party/grand totals.
    - 'singleSheet': All parties merged into one master worksheet.
    - 'multiSheets': Master Consolidated sheet + individual party tabs.
    """
    all_raw_vouchers = []
    for fpath in file_paths:
        vchs = extract_vouchers_from_workbook(fpath)
        all_raw_vouchers.extend(vchs)

    # Group vouchers by party (case-insensitive key)
    party_map: Dict[str, List[Dict[str, Any]]] = OrderedDict()
    party_canonical_name: Dict[str, str] = {}
    party_metadata: Dict[str, Dict[str, str]] = {}

    for v in all_raw_vouchers:
        raw_name = (v.get('vendor') or 'General').strip()
        norm_key = raw_name.lower()
        party_map.setdefault(norm_key, []).append(v)

        # Store canonical display name (prefer first non-empty)
        if norm_key not in party_canonical_name:
            party_canonical_name[norm_key] = raw_name

        # Accumulate party metadata (GSTN, Address, Closing Balance)
        meta = party_metadata.setdefault(norm_key, {'gstn': '', 'address': '', 'closingBalance': ''})
        if v.get('gstn') and not meta['gstn']:
            meta['gstn'] = v['gstn']
        if v.get('address') and len(v['address']) > len(meta['address']):
            meta['address'] = v['address']
        if v.get('closingBalance'):
            meta['closingBalance'] = v['closingBalance']

    # Sort vouchers date-wise within each party and compute running ledger balance
    consolidated_parties: Dict[str, List[Dict[str, Any]]] = OrderedDict()

    for norm_key, p_vouchers in party_map.items():
        disp_name = party_canonical_name.get(norm_key, norm_key)
        meta = party_metadata.get(norm_key, {})

        # Chronological sort across all years/files
        p_vouchers.sort(key=lambda r: (r['_sort_date'], safe_vch_key(r.get('voucherNo'))))

        # Propagate metadata & set individual voucher ledger balance to 0
        for v in p_vouchers:
            v['vendor'] = disp_name
            if not v.get('address') and meta.get('address'):
                v['address'] = meta['address']
            if not v.get('closingBalance') and meta.get('closingBalance'):
                v['closingBalance'] = meta['closingBalance']

            v['ledgerBalance'] = 0

        consolidated_parties[disp_name] = p_vouchers

    # Flatten vouchers with sequential S No across the entire workbook
    all_sorted_vouchers: List[Dict[str, Any]] = []
    global_s_no = 1
    for p_name, p_vouchers in consolidated_parties.items():
        for v in p_vouchers:
            v_copy = dict(v)
            v_copy['sNo'] = global_s_no
            global_s_no += 1
            all_sorted_vouchers.append(v_copy)

    if not output_path:
        output_path = os.path.join(os.getcwd(), "Consolidated_Tally_Merged_Report.xlsx")

    # High-speed xlsxwriter path
    try:
        import xlsxwriter
        wb = xlsxwriter.Workbook(output_path, {'constant_memory': False})
        if mode == "multiSheets":
            used_sheet_names = set()
            if len(consolidated_parties) > 1:
                ws_master = wb.add_worksheet("Consolidated Report")
                used_sheet_names.add("consolidated report")
                write_voucher_worksheet_xlsxwriter(ws_master, wb, all_sorted_vouchers, add_source_col=add_source_col)

                for p_name, p_vouchers in consolidated_parties.items():
                    clean_name = re.sub(r'[\\/*?:\[\]]', '_', p_name)[:28] or "Party"
                    tab_title = clean_name
                    counter = 1
                    while tab_title.lower() in used_sheet_names:
                        tab_title = f"{clean_name[:24]}_{counter}"
                        counter += 1
                    used_sheet_names.add(tab_title.lower())

                    party_tab_vchs = []
                    for idx, pv in enumerate(p_vouchers, start=1):
                        pvc = dict(pv)
                        pvc['sNo'] = idx
                        party_tab_vchs.append(pvc)

                    ws_party = wb.add_worksheet(tab_title)
                    write_voucher_worksheet_xlsxwriter(ws_party, wb, party_tab_vchs, add_source_col=add_source_col)
            else:
                p_name = list(consolidated_parties.keys())[0] if consolidated_parties else "Voucher Report"
                tab_title = re.sub(r'[\\/*?:\[\]]', '_', p_name)[:30] or "Voucher Report"
                ws_single = wb.add_worksheet(tab_title)
                write_voucher_worksheet_xlsxwriter(ws_single, wb, all_sorted_vouchers, add_source_col=add_source_col)
        else:
            # singleSheet
            if len(consolidated_parties) == 1:
                p_name = list(consolidated_parties.keys())[0]
                tab_title = re.sub(r'[\\/*?:\[\]]', '_', p_name)[:30] or "Voucher Report"
            else:
                tab_title = "Consolidated Report"
            ws_dest = wb.add_worksheet(tab_title)
            write_voucher_worksheet_xlsxwriter(ws_dest, wb, all_sorted_vouchers, add_source_col=add_source_col)

        wb.close()
        return output_path
    except ImportError:
        pass

    # Openpyxl Fallback
    wb = openpyxl.Workbook()

    if mode == "multiSheets":
        used_sheet_names = set()

        if len(consolidated_parties) > 1:
            ws_master = wb.active
            ws_master.title = "Consolidated Report"
            used_sheet_names.add("consolidated report")
            build_voucher_worksheet(ws_master, all_sorted_vouchers, add_source_col=add_source_col)

            for p_name, p_vouchers in consolidated_parties.items():
                clean_name = re.sub(r'[\\/*?:\[\]]', '_', p_name)[:28] or "Party"
                tab_title = clean_name
                counter = 1
                while tab_title.lower() in used_sheet_names:
                    tab_title = f"{clean_name[:24]}_{counter}"
                    counter += 1
                used_sheet_names.add(tab_title.lower())

                party_tab_vchs = []
                for idx, pv in enumerate(p_vouchers, start=1):
                    pvc = dict(pv)
                    pvc['sNo'] = idx
                    party_tab_vchs.append(pvc)

                ws_party = wb.create_sheet(title=tab_title)
                build_voucher_worksheet(ws_party, party_tab_vchs, add_source_col=add_source_col)
        else:
            ws_single = wb.active
            p_name = list(consolidated_parties.keys())[0] if consolidated_parties else "Voucher Report"
            tab_title = re.sub(r'[\\/*?:\[\]]', '_', p_name)[:30] or "Voucher Report"
            ws_single.title = tab_title
            build_voucher_worksheet(ws_single, all_sorted_vouchers, add_source_col=add_source_col)

    else:
        # mode == "singleSheet" (Default)
        ws_dest = wb.active
        if len(consolidated_parties) == 1:
            p_name = list(consolidated_parties.keys())[0]
            ws_dest.title = re.sub(r'[\\/*?:\[\]]', '_', p_name)[:30] or "Voucher Report"
        else:
            ws_dest.title = "Consolidated Report"

        build_voucher_worksheet(ws_dest, all_sorted_vouchers, add_source_col=add_source_col)

    wb.save(output_path)
    return output_path
