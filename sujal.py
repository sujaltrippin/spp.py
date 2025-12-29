import csv
import os
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Table, TableStyle, Spacer
)
from reportlab.lib.styles import ParagraphStyle
# ---------------- PDF CREATOR ----------------
def create_invoice_pdf(booking_id, vendor_name, property_name, amount, output_folder):
    os.makedirs(output_folder, exist_ok=True)
    filename = os.path.join(output_folder, f"{booking_id}.pdf")
    doc = SimpleDocTemplate(
        filename,
        pagesize=A4,
        leftMargin=40,
        rightMargin=40,
        topMargin=40,
        bottomMargin=40
    )
    # ---------- PASTEL PEACH THEME ----------
    PEACH_BG = colors.HexColor("#FFF1E6")
    PEACH_DARK = colors.HexColor("#E07A5F")
    PEACH_LIGHT = colors.HexColor("#FDE8D7")
    BORDER = colors.HexColor("#E6A57E")
    TEXT = colors.HexColor("#333333")
    CONTENT_WIDTH = 420
    # ---------- STYLES ----------
    title = ParagraphStyle(
        "Title",
        fontName="Helvetica-Bold",
        fontSize=22,
        textColor=PEACH_DARK,
        alignment=1
    )
    vendor_style = ParagraphStyle(
        "Vendor",
        fontName="Helvetica-Bold",
        fontSize=13,
        alignment=1,
        textColor=TEXT
    )
    property_style = ParagraphStyle(
        "Property",
        fontName="Helvetica",
        fontSize=10,
        alignment=1,
        textColor=colors.grey
    )
    normal = ParagraphStyle(
        "Normal",
        fontName="Helvetica",
        fontSize=10,
        textColor=TEXT
    )
    footer_style = ParagraphStyle(
        "Footer",
        fontName="Helvetica",
        fontSize=9,
        alignment=1,
        textColor=colors.grey
    )
    elements = []
    # ---------------- MAIN CONTENT ----------------
    content = []
    # ---------------- HEADER ----------------
    content.append(Paragraph("STAYVISTA", title))
    content.append(Spacer(1, 6))
    content.append(Paragraph("INVOICE", title))
    content.append(Spacer(1, 20))
    # ---------------- PAYMENT DETAILS ----------------
    content.append(Paragraph(vendor_name, vendor_style))
    content.append(Spacer(1, 4))
    content.append(Paragraph(property_name, property_style))
    content.append(Spacer(1, 4))
    content.append(Paragraph(f"Booking ID: {booking_id}", property_style))
    content.append(Spacer(1, 20))
    # ---------------- ITEM TABLE ----------------
    amt = f"Rs. {amount}"
    items = [
        ["Description", "Qty", "Rate", "Amount"],
        [f"Cook Arranged – Booking {booking_id}", "1", amt, amt]
    ]
    item_table = Table(items, colWidths=[220, 50, 75, 75])
    item_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), PEACH_DARK),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("BACKGROUND", (0, 1), (-1, -1), PEACH_LIGHT),
        ("GRID", (0, 0), (-1, -1), 0.5, BORDER),
        ("ALIGN", (1, 1), (1, -1), "CENTER"),
        ("ALIGN", (2, 1), (-1, -1), "RIGHT"),
        ("TOPPADDING", (0, 0), (-1, -1), 10),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 10),
    ]))
    content.append(item_table)
    content.append(Spacer(1, 22))
    # ---------------- TOTALS ----------------
    totals = [
        ["Subtotal", amt],
        ["Tax", "Rs. 0"],
        ["Total Amount", amt],
        ["Amount Paid", "Rs. 0"]
    ]
    totals_table = Table(
        totals,
        colWidths=[CONTENT_WIDTH * 0.6, CONTENT_WIDTH * 0.4]
    )
    totals_table.setStyle(TableStyle([
        ("BOX", (0, 0), (-1, -1), 1, PEACH_DARK),
        ("INNERGRID", (0, 0), (-1, -1), 0.25, BORDER),
        ("BACKGROUND", (0, 0), (-1, -1), colors.whitesmoke),
        ("BACKGROUND", (0, 2), (-1, 3), PEACH_LIGHT),
        ("FONTNAME", (0, 2), (-1, 3), "Helvetica-Bold"),
        ("ALIGN", (1, 0), (-1, -1), "RIGHT"),
        ("LEFTPADDING", (0, 0), (-1, -1), 12),
        ("RIGHTPADDING", (0, 0), (-1, -1), 12),
        ("TOPPADDING", (0, 0), (-1, -1), 8),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
    ]))
    content.append(totals_table)
    content.append(Spacer(1, 18))
    # ---------------- FOOTER ----------------
    content.append(Paragraph(
        "This is a system-generated invoice. No signature is required.",
        footer_style
    ))
    # ---------------- PEACH BACKGROUND WRAPPER ----------------
    wrapper = Table([[content]], colWidths=[CONTENT_WIDTH])
    wrapper.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), PEACH_BG),
        ("BOX", (0, 0), (-1, -1), 1, BORDER),
        ("LEFTPADDING", (0, 0), (-1, -1), 20),
        ("RIGHTPADDING", (0, 0), (-1, -1), 20),
        ("TOPPADDING", (0, 0), (-1, -1), 20),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 20),
        ("ALIGN", (0, 0), (-1, -1), "CENTER")
    ]))
    elements.append(wrapper)
    doc.build(elements)
    print(f":white_tick: Invoice generated: {filename}")
# ---------------- MAIN ----------------
def main():
    with open("bills.csv", newline="", encoding="utf-8") as file:
        reader = csv.reader(file)
        next(reader)
        for row in reader:
            create_invoice_pdf(
                booking_id=row[0],
                vendor_name=row[1],
                property_name=row[2],
                amount=row[3],
                output_folder="stayvista_invoices_pdf"
            )
if __name__ == "__main__":
    main()