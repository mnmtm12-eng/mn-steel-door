# -*- coding: utf-8 -*-
"""توليد وثائق Excel احترافية (xlsx) لكل نوع وثيقة تصدير."""
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from io import BytesIO

NAVY = "1E2B48"; STEEL = "27608B"; LIGHT = "F5F7FA"; ROW = "FAFBFC"
WHITE = "FFFFFF"; GREY = "6B7177"
F_TITLE = Font(name="Arial", size=18, bold=True, color=NAVY)
F_SUB = Font(name="Arial", size=8, color=GREY)
F_DOC = Font(name="Arial", size=16, bold=True, color=NAVY)
F_H = Font(name="Arial", size=9, bold=True, color=WHITE)
F_LBL = Font(name="Arial", size=8, bold=True, color=STEEL)
F_K = Font(name="Arial", size=9, color=GREY)
F_V = Font(name="Arial", size=10, bold=True, color=NAVY)
F_TXT = Font(name="Arial", size=10, color="1A1D21")
F_BOLD = Font(name="Arial", size=10, bold=True, color=NAVY)
F_GRAND = Font(name="Arial", size=12, bold=True, color=WHITE)
FILL_NAVY = PatternFill("solid", fgColor=NAVY)
FILL_LIGHT = PatternFill("solid", fgColor=LIGHT)
FILL_ROW = PatternFill("solid", fgColor=ROW)
CTR = Alignment(horizontal="center", vertical="center", wrap_text=True)
LFT = Alignment(horizontal="left", vertical="center", wrap_text=True)
RGT = Alignment(horizontal="right", vertical="center")
thin = Side(style="thin", color="E1E4E8")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)


def _f(n):
    try: return f"{float(n):,.2f}"
    except Exception: return n

def _hdr(ws, ctx, ncols=5):
    s, co, cu, doc = ctx["s"], ctx["company"], ctx["cust"], ctx["doc"]
    last = get_column_letter(ncols)
    ws.merge_cells(f"A1:{get_column_letter(max(2,ncols-1))}1")
    ws["A1"] = "MN STEEL DOOR"; ws["A1"].font = F_TITLE
    ws.merge_cells(f"A2:{get_column_letter(max(2,ncols-1))}2")
    ws["A2"] = "Steel Doors · Export"; ws["A2"].font = F_SUB
    ws[f"{last}1"] = doc["label"]; ws[f"{last}1"].font = F_DOC; ws[f"{last}1"].alignment = RGT
    ws[f"{last}2"] = f"Ref {s['ref_no']}"; ws[f"{last}2"].font = F_K; ws[f"{last}2"].alignment = RGT
    r = 4
    pairs = [("REFERENCE", s["ref_no"]), ("DATE", s["created_at"]), ("HS CODE", ctx["hs"]), ("ORIGIN", ctx["origin"])]
    for i, (k, v) in enumerate(pairs):
        c = get_column_letter(1 + i)
        ws[f"{c}{r}"] = k; ws[f"{c}{r}"].font = F_LBL
        ws[f"{c}{r+1}"] = v; ws[f"{c}{r+1}"].font = F_V
    r += 3
    # exporter
    ws[f"A{r}"] = "EXPORTER"; ws[f"A{r}"].font = F_LBL
    ws[f"A{r+1}"] = co["name"]; ws[f"A{r+1}"].font = F_BOLD
    ex = [x for x in [co["address"], (("Tel: " + co["phone"]) if co["phone"] else ""),
          (("Email: " + co["email"]) if co["email"] else ""),
          (("Tax No (VKN): " + co["tax_no"]) if (("tax_no" in co.keys()) and co["tax_no"]) else "")] if x]
    for i, line in enumerate(ex):
        ws[f"A{r+2+i}"] = line; ws[f"A{r+2+i}"].font = F_TXT
    # consignee
    midc = get_column_letter(max(3, ncols - 1))
    ws[f"{midc}{r}"] = "CONSIGNEE"; ws[f"{midc}{r}"].font = F_LBL
    ws[f"{midc}{r+1}"] = cu["name"]; ws[f"{midc}{r+1}"].font = F_BOLD
    cn = [x for x in [cu["contact"], cu["address"], (("Tel: " + cu["phone"]) if cu["phone"] else ""), (("Tax No: " + cu["tax_no"]) if cu["tax_no"] else "")] if x]
    for i, line in enumerate(cn):
        ws[f"{midc}{r+2+i}"] = line; ws[f"{midc}{r+2+i}"].font = F_TXT
    r += 2 + max(len(ex), len(cn)) + 1
    # shipment strip
    ship = [("Incoterm", s["incoterm"]), ("Loading", s["port_loading"] or "—"),
            ("Discharge", s["port_discharge"] or "—"), ("ACID No", s["acid_no"] or "—"),
            ("Country", cu["country"] or "—")]
    for i, (k, v) in enumerate(ship[:ncols]):
        c = get_column_letter(1 + i)
        ws[f"{c}{r}"] = k.upper(); ws[f"{c}{r}"].font = F_LBL; ws[f"{c}{r}"].fill = FILL_LIGHT
        ws[f"{c}{r+1}"] = v; ws[f"{c}{r+1}"].font = F_V; ws[f"{c}{r+1}"].fill = FILL_LIGHT
    return r + 3

def _table_header(ws, r, headers):
    for i, h in enumerate(headers):
        c = ws.cell(row=r, column=i + 1, value=h)
        c.font = F_H; c.fill = FILL_NAVY; c.alignment = CTR; c.border = BORDER
    return r + 1

def _widths(ws, widths):
    for i, w in enumerate(widths):
        ws.column_dimensions[get_column_letter(i + 1)].width = w


def build(doctype, ctx):
    s, co, cu = ctx["s"], ctx["company"], ctx["cust"]
    items, containers = ctx["items"], ctx["containers"]
    cur = s["currency"] or "USD"
    wb = Workbook(); ws = wb.active; ws.title = ctx["doc"]["label"][:31]
    ws.sheet_view.showGridLines = False

    if doctype in ("commercial", "proforma"):
        _widths(ws, [6, 40, 10, 18, 18]); r = _hdr(ws, ctx, 5)
        r = _table_header(ws, r, ["#", "Description", "Qty", f"Unit Price ({cur})", f"Amount ({cur})"])
        for i, it in enumerate(items):
            lt = it.get("line_total", it["qty"] * it["unit_price"])
            vals = [i + 1, f"{it['name']}  ({it['width']}×{it['height']}×{it['beam']} cm)", int(it["qty"]), _f(it["unit_price"]), _f(lt)]
            for j, v in enumerate(vals):
                c = ws.cell(row=r, column=j + 1, value=v); c.border = BORDER
                c.alignment = LFT if j == 1 else CTR; c.font = F_TXT
                if i % 2: c.fill = FILL_ROW
            r += 1
        # total
        ws.cell(row=r, column=4, value="TOTAL").font = F_BOLD
        gc = ws.cell(row=r, column=5, value=f"{cur} {_f(s['total'])}"); gc.font = F_GRAND; gc.fill = FILL_NAVY; gc.alignment = CTR
        r += 2
        if s["payment_terms"]:
            ws.cell(row=r, column=1, value="Payment Terms:").font = F_LBL
            ws.cell(row=r, column=2, value=s["payment_terms"]).font = F_TXT; r += 1
        if co["iban"]:
            for k, v in [("Beneficiary", co["name"]), ("Bank", co["bank_name"]), ("Branch", f"{co['branch']} ({co['branch_code']})"), ("IBAN", co["iban"]), ("SWIFT", co["swift"])]:
                ws.cell(row=r, column=1, value=k).font = F_K; ws.cell(row=r, column=2, value=v).font = F_TXT; r += 1
        r += 1
        ws.cell(row=r, column=1, value="THIS INVOICE IS AUTHENTIC AND APPROVED BY THE EXPORTER").font = F_LBL

    elif doctype == "packing":
        _widths(ws, [6, 34, 18, 22, 20]); r = _hdr(ws, ctx, 5)
        r = _table_header(ws, r, ["#", "Description of Goods", "Quantity", "", ""])
        for i, it in enumerate(items):
            for j, v in enumerate([i + 1, f"{it['name']}  ({it['width']}×{it['height']}×{it['beam']} cm)", f"{int(it['qty'])} pcs", "", ""]):
                c = ws.cell(row=r, column=j + 1, value=v); c.border = BORDER; c.font = F_TXT
                c.alignment = LFT if j == 1 else CTR
                if i % 2: c.fill = FILL_ROW
            r += 1
        ws.cell(row=r, column=2, value="TOTAL").font = F_BOLD
        ws.cell(row=r, column=3, value=f"{s['qty_total']} pcs").font = F_BOLD; ws.cell(row=r, column=3).alignment = CTR
        r += 2
        ws.cell(row=r, column=1, value="CONTAINER / PACKING DETAILS").font = F_LBL; r += 1
        r = _table_header(ws, r, ["#", "Container No.", "Seal No.", "Content", "Gross Weight (KG)"])
        for i, ct in enumerate(containers):
            for j, v in enumerate([i + 1, ct.get("no", "—") or "—", ct.get("seal", "—") or "—", ct.get("content", "—") or "—", _f(ct.get("weight", 0))]):
                c = ws.cell(row=r, column=j + 1, value=v); c.border = BORDER; c.font = F_TXT
                c.alignment = LFT if j == 3 else CTR
                if i % 2: c.fill = FILL_ROW
            r += 1
        ws.cell(row=r, column=4, value="TOTAL GROSS WEIGHT").font = F_BOLD
        gc = ws.cell(row=r, column=5, value=f"{_f(s['gross_weight'] or 0)} KG"); gc.font = F_GRAND; gc.fill = FILL_NAVY; gc.alignment = CTR

    elif doctype == "origin":
        _widths(ws, [6, 40, 14, 16, 18]); r = _hdr(ws, ctx, 5)
        ws.cell(row=r, column=1, value="CERTIFICATE OF ORIGIN / شهادة منشأ").font = F_DOC; r += 2
        r = _table_header(ws, r, ["#", "Description of Goods", "Quantity", "HS Code", "Country of Origin"])
        for i, it in enumerate(items):
            for j, v in enumerate([i + 1, f"{it['name']}  ({it['width']}×{it['height']}×{it['beam']} cm)", f"{int(it['qty'])} pcs", ctx["hs"], ctx["origin"]]):
                c = ws.cell(row=r, column=j + 1, value=v); c.border = BORDER; c.font = F_TXT
                c.alignment = LFT if j == 1 else CTR
                if i % 2: c.fill = FILL_ROW
            r += 1
        ws.cell(row=r, column=2, value="TOTAL").font = F_BOLD
        ws.cell(row=r, column=3, value=f"{s['qty_total']} pcs").font = F_BOLD; ws.cell(row=r, column=3).alignment = CTR
        r += 2
        ws.cell(row=r, column=1, value=f"We hereby certify that the goods described above are of {ctx['origin']} origin.").font = F_TXT

    elif doctype == "shipping":
        _widths(ws, [6, 30, 16, 18, 22]); r = _hdr(ws, ctx, 5)
        ws.cell(row=r, column=1, value="To: Shipping Line / Freight Forwarder").font = F_BOLD; r += 1
        ws.cell(row=r, column=1, value=f"Please arrange shipment for reference {s['ref_no']} per the details below.").font = F_TXT; r += 2
        r = _table_header(ws, r, ["Commodity", "HS Code", "Quantity", "Containers", "Gross Weight"])
        for j, v in enumerate(["Steel Doors", ctx["hs"], f"{s['qty_total']} pcs", s["container_count"], f"{_f(s['gross_weight'] or 0)} KG"]):
            c = ws.cell(row=r, column=j + 1, value=v); c.border = BORDER; c.font = F_TXT; c.alignment = CTR
        r += 2
        if containers:
            r = _table_header(ws, r, ["#", "Container No.", "Seal No.", "Content", "Weight (KG)"])
            for i, ct in enumerate(containers):
                for j, v in enumerate([i + 1, ct.get("no", "—") or "—", ct.get("seal", "—") or "—", ct.get("content", "—") or "—", _f(ct.get("weight", 0))]):
                    c = ws.cell(row=r, column=j + 1, value=v); c.border = BORDER; c.font = F_TXT
                    c.alignment = LFT if j == 3 else CTR
                r += 1
        r += 1
        if s["acid_no"]:
            ws.cell(row=r, column=1, value="ACID No:").font = F_LBL
            ws.cell(row=r, column=2, value=s["acid_no"]).font = F_TXT

    return wb


def to_bytes(wb):
    bio = BytesIO(); wb.save(bio); return bio.getvalue()
