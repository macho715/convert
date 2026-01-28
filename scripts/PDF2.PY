from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib import colors
from reportlab.lib.units import mm

W, H = A4

def euro_fmt(x: float) -> str:
    s = f"{x:,.2f}"
    return s.replace(",", "X").replace(".", ",").replace("X", ".")

def draw_logo(c):
    # Match look: logo centered, modest size
    cx = W/2
    y_top = H - 15*mm
    r = 12
    c.setLineWidth(1.2)
    c.circle(cx - 85, y_top - 15, r, stroke=1, fill=0)
    c.setFont("Helvetica-Bold", 18)
    c.drawString(cx - 91, y_top - 26, "T")
    c.setFont("Helvetica-Bold", 26)
    c.drawString(cx - 65, y_top - 30, "TRENCH")
    c.setFont("Helvetica", 11)
    c.drawString(cx - 63, y_top - 42, "Sense the Power")

def box(c, x, y, w, h, fill=None, lw=1):
    c.setLineWidth(lw)
    if fill is not None:
        c.setFillColor(fill)
        c.rect(x, y, w, h, stroke=1, fill=1)
        c.setFillColor(colors.black)
    else:
        c.rect(x, y, w, h, stroke=1, fill=0)

def make_invoice_v5(out_path):
    c = canvas.Canvas(out_path, pagesize=A4)
    margin_l = 22*mm
    margin_r = 22*mm
    content_w = W - margin_l - margin_r
    
    draw_logo(c)
    
    # Title box - right aligned, below logo
    title_w = 150*mm
    title_h = 16*mm
    title_x = W - margin_r - title_w
    title_y = H - 55*mm
    box(c, title_x, title_y, title_w, title_h, fill=colors.lightgrey, lw=1)
    c.setFont("Helvetica-Bold", 14)
    c.drawString(title_x+10, title_y+title_h/2+1, "Shipping Invoice  No. 131801")
    c.setFont("Helvetica", 11)
    c.drawString(title_x+118, title_y+title_h/2+1, "of 15.12.2025")
    c.setFont("Helvetica", 11)
    c.drawString(title_x+10, title_y-16, "OdV 21522")
    
    # Customer block
    x0 = margin_l
    y_customer = H - 82*mm
    c.setFont("Helvetica-Bold", 11)
    c.drawString(x0, y_customer, "Customer")
    c.setFont("Helvetica", 11)
    cust_lines = [
        "Samsung C& T Corporation Abu Dhabi",
        "Office M04, M-Flr Saeed Mohamed Hamda",
        "Al Najda St",
        "42654 Abu Dhabi",
        "UTD.ARAB.EMIR."
    ]
    for i, line in enumerate(cust_lines):
        c.drawString(x0, y_customer-14*(i+1), line)
    
    # TEL
    c.setFont("Helvetica", 11)
    c.drawString(W/2 + 20, y_customer-14*4, "TEL: +974-28015100")
    
    # Shipping address
    y_ship = H - 124*mm
    c.setFont("Helvetica-Bold", 11)
    c.drawString(x0, y_ship, "Shipping Address")
    c.setFont("Helvetica", 11)
    for i, line in enumerate(cust_lines):
        c.drawString(x0, y_ship-14*(i+1), line)
    
    # Order box
    ob_x = margin_l
    ob_y = H - 182*mm
    ob_h = 32*mm
    box(c, ob_x, ob_y, content_w, ob_h, fill=colors.lightgrey, lw=1)
    c.setFont("Courier", 12)
    order_lines = [
        "Your Order: OdV 21522",
        "Shipment by: AIR",
        "Packing list n°: 5019590",
        "Delivery: FCA our factory",
        "Payment: 60 gg data fattura"
    ]
    for i, line in enumerate(order_lines):
        c.drawString(ob_x+10, ob_y+ob_h-16 - i*14, line)
    
    # Table header
    th_h = 9*mm
    th_y = ob_y - 14*mm
    box(c, ob_x, th_y, content_w, th_h, fill=colors.lightgrey, lw=1)
    c.setFont("Helvetica-Bold", 9.5)
    col_item = ob_x + 6
    col_desc = ob_x + 32
    col_qty = ob_x + content_w*0.63
    col_unit_r = ob_x + content_w*0.82
    col_disc_r = ob_x + content_w*0.90
    col_amt_r = ob_x + content_w*0.985
    c.drawString(col_item, th_y+th_h-7, "ITEM")
    c.drawString(col_desc, th_y+th_h-7, "DESCRIPTION OF GOODS")
    c.drawString(col_qty, th_y+th_h-7, "Q.TY")
    c.drawRightString(col_unit_r, th_y+th_h-7, "UNIT PRICE")
    c.drawRightString(col_disc_r, th_y+th_h-7, "DISCOUNT %")
    c.drawRightString(col_amt_r, th_y+th_h-7, "AMOUNT")
    
    # Main box
    mb_y = 32*mm
    mb_h = th_y - mb_y
    box(c, ob_x, mb_y, content_w, mb_h, fill=None, lw=1)
    
    # Item line
    row_y = th_y - 16
    c.setFont("Helvetica", 10.5)
    c.drawString(col_item, row_y, ".001 RC Divider")
    c.drawString(col_qty, row_y, "2,00")
    unit_price = 19000.00
    total_amount = 38000.00
    c.drawRightString(col_unit_r, row_y, euro_fmt(unit_price))
    # Discount empty
    c.drawRightString(col_amt_r, row_y, euro_fmt(total_amount))
    
    # Detail lines
    c.setFont("Courier", 10)
    detail_lines = [
        "sn. 30186699-R and sn 30186700-R",
        "Customs tariff 85043121",
        "At closure of temporary import n° 25ITQRX05DAU8710R0  del 23/10/2025",
        "",
        "Export Control Codes :  ECCN=N - AL=N",
        "",
        "Re-imported after repair / Value addition due to repair.",
        "Customs duty/taxes basis: repair value (value added) only.",
        "",
        f"Total costs for custom duty only : EUR  {euro_fmt(total_amount)}",
        "",
        "N ° 2 Case dimensions cm154x107x256h",
        "Net weight 740 Kg     Gross weight 1.200 Kg",
    ]
    y = row_y - 18
    for line in detail_lines:
        c.drawString(ob_x+25, y, line)
        y -= 12
    
    # Compliance paragraph starts after some gap but above VAT strip
    vat_y = 20*mm
    vat_h = 16*mm
    para_top = max(y - 18, vat_y + vat_h + 110)  # ensure not overlapping details
    c.setFont("Helvetica", 11)
    para_lines = [
        'Goods labeled with "AL not equal to N" are subject to European or Italian export regulations',
        "and",
        'relevant authorizations when being exported out of the EU. Goods labeled with "ECCN not',
        "equal to",
        'N" are subject to US re-export regulations and relevant authorizations. Even without a label, or',
        'with label "AL:N" or "ECCN:N", authorization may be required due to the final end-use and the',
        "destination for which the goods are to be used.",
        "Notice: Compliance with legal and internal regulations is an integral part of all business pro-",
        "cesses",
        'at Trench Italia. Possible Infringements can be reported to our HelpDesk "Tell us” at',
        "http://www.trenchgroup.com/Trench_Compliance.html",
    ]
    yy = para_top
    for line in para_lines:
        c.drawString(ob_x+10, yy, line)
        yy -= 14
    
    # VAT strip
    box(c, ob_x, vat_y, content_w, vat_h, fill=colors.lightgrey, lw=1)
    c.setFont("Courier", 9.5)
    c.drawString(ob_x+10, vat_y+vat_h-12, "VAT Base")
    c.drawString(ob_x+85, vat_y+vat_h-12, "VAT %")
    c.drawString(ob_x+135, vat_y+vat_h-12, "VAT Amount")
    c.drawString(ob_x+content_w-95, vat_y+vat_h-12, "Total Amount")
    c.setFont("Courier", 10)
    c.drawString(ob_x+10, vat_y+6, euro_fmt(total_amount))
    c.drawString(ob_x+85, vat_y+6, "0,00")
    c.drawString(ob_x+135, vat_y+6, "0,00")
    c.drawString(ob_x+content_w-145, vat_y+6, "EUR")
    c.drawRightString(ob_x+content_w-10, vat_y+6, euro_fmt(total_amount))
    
    # Footer & page no (above footer)
    c.setFont("Helvetica", 9)
    c.drawRightString(W - margin_r, 14*mm, "Page No. 1")
    c.setFont("Helvetica", 7.5)
    c.drawString(margin_l, 14*mm, "Mod:ZMG_INVOICEMAN")
    
    c.setFont("Helvetica-Bold", 12)
    c.drawCentredString(W/2, 10*mm, "Trench Italia S.r.l. Strada Curagnata, 37 | 17014 Cairo Montenotte (SV) - Italia -")
    c.setFont("Helvetica", 8)
    c.drawCentredString(W/2, 6.5*mm, "Tel. (39) 01951611 Fax (39) 0195161401 | PEC: info@pec.trenchitalia.it")
    
    c.showPage()
    c.save()

import os
from pathlib import Path

# Windows 환경에 맞게 출력 경로 설정
output_dir = Path(__file__).parent / "output"
output_dir.mkdir(exist_ok=True)
out_pdf_v5 = str(output_dir / "Trench_Corrected_Invoice_131801_UNIFIED_v5.pdf")

if __name__ == "__main__":
    print(f"PDF 생성 중: {out_pdf_v5}")
    make_invoice_v5(out_pdf_v5)
    print(f"✅ PDF 생성 완료: {out_pdf_v5}")
