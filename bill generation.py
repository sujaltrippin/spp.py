import csv
import os
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Table, TableStyle, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from dotenv import load_dotenv
from oauth2client.service_account import ServiceAccountCredentials
from googleapiclient.discovery import build
from datetime import datetime
import pytz
import gspread
from googleapiclient.http import MediaFileUpload
from google_auth_oauthlib.flow import InstalledAppFlow
from google.oauth2.credentials import Credentials

# Load environment variables
load_dotenv()
ist = pytz.timezone('Asia/Kolkata')
now_ist = datetime.now(ist)


with open("credentials.json", "w") as f:
    f.write(os.getenv("GOOGLE_SHEET_CONNECTOR"))
    
# Google Sheets Auth
scope = ["https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/spreadsheets"]


def get_oauth_credentials():
    creds = None

    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", scope)

    if not creds or not creds.valid:
        flow = InstalledAppFlow.from_client_secrets_file(
            "oauth.json",
            scope
        )
        creds = flow.run_local_server(port=0)

        with open("token.json", "w") as token:
            token.write(creds.to_json())

    return creds

creds = get_oauth_credentials()
gs_client = gspread.authorize(creds)    
drive_service = build('drive', 'v3', credentials=creds)
sheets_service = build("sheets", "v4", credentials=creds)


# Try to register DejaVu font, but fall back to default if not found
try:
    pdfmetrics.registerFont(TTFont('DejaVu', 'DejaVuLGCSansCondensed.ttf'))
    pdfmetrics.registerFont(TTFont('DejaVu-Bold', 'DejaVuLGCSansCondensed-Bold.ttf'))
    DEFAULT_FONT = 'DejaVu'
except:
    print("Note: DejaVu font not found. Using default font.")
    DEFAULT_FONT = 'Helvetica'

def upload_to_drive(file_path, drive_folder_id):
    file_metadata = {
        'name': os.path.basename(file_path),
        'parents': [drive_folder_id]
    }

    media = MediaFileUpload(
        file_path,
        mimetype='application/pdf',
        resumable=False
    )

    uploaded_file = drive_service.files().create(
        body=file_metadata,
        media_body=media,
        fields='id, name'
    ).execute()

    return uploaded_file


def create_invoice_pdf(booking_id, vendor_name, property_name, amount, output_folder):
    """
    Create a StayVista invoice PDF for a single booking
    """
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    filename = os.path.join(output_folder, f"{booking_id}.pdf")
    doc = SimpleDocTemplate(filename, pagesize=A4)
    styles = getSampleStyleSheet()

    # Create custom styles with the registered font
    normal_style = ParagraphStyle(
        'CustomNormal',
        parent=styles['Normal'],
        fontName=DEFAULT_FONT,
        fontSize=10,
        leading=12
    )

    title_style = ParagraphStyle(
        'CustomTitle',
        parent=styles['Title'],
        fontName=DEFAULT_FONT,
        fontSize=18,
        spaceAfter=20,
        alignment=1  # Center alignment
    )

    heading2_style = ParagraphStyle(
        'CustomHeading2',
        parent=styles['Heading2'],
        fontName=DEFAULT_FONT,
        fontSize=12,
        spaceAfter=6
    )

    vendor_style = ParagraphStyle(
        'VendorStyle',
        parent=styles['Normal'],
        fontName=DEFAULT_FONT,
        fontSize=14,
        spaceAfter=12,
        spaceBefore=6
    )

    elements = []

    # Title
    title = Paragraph("<b>Vista Invoice</b>", title_style)
    elements.append(title)
    elements.append(Spacer(1, 12))

    # Bill To section
    elements.append(Paragraph("<b>Payment to</b>", heading2_style))
    elements.append(Paragraph(f"<b>{vendor_name}</b>", vendor_style))
    elements.append(Paragraph(f"Property: {property_name}", normal_style))
    elements.append(Spacer(1, 12))

    # Table for invoice details
    data = [
        ["StayVista", "", "", ""],
        ["ITEMS", "QTY.", "RATE", "AMOUNT"],
        [f"Booking id {booking_id} - Cook Arranged", "1", f"Rs. {amount}", f"Rs. {amount}"]
    ]

    table = Table(data, colWidths=[250, 50, 100, 100])
    table.setStyle(TableStyle([
        ("SPAN", (0, 0), (-1, 0)),  # Span first row across all columns
        ("BACKGROUND", (0, 1), (-1, 1), colors.lightgrey),
        ("ALIGN", (0, 0), (-1, 0), "CENTER"),  # Center first row
        ("ALIGN", (1, 1), (-1, -1), "CENTER"),  # Center quantity, rate, amount
        ("ALIGN", (0, 1), (0, 1), "CENTER"),  # Center "ITEMS" header
        ("FONTNAME", (0, 0), (-1, -1), DEFAULT_FONT),
        ("FONTSIZE", (0, 0), (-1, 0), 12),
        ("BOTTOMPADDING", (0, 0), (-1, 0), 10),
        ("BOX", (0, 0), (-1, -1), 1, colors.black),
        ("GRID", (0, 1), (-1, -1), 0.5, colors.grey),
        ("FONTNAME", (0, 1), (-1, 1), DEFAULT_FONT + "-Bold" if DEFAULT_FONT == 'DejaVu' else 'Helvetica-Bold'),
    ]))
    elements.append(table)
    elements.append(Spacer(1, 20))

    # Terms and conditions
    elements.append(Paragraph("<b>StayVista</b>", heading2_style))
    elements.append(Paragraph("*.Food & beverage Expenses.", normal_style))
    elements.append(Spacer(1, 12))

    # Amount details
    elements.append(Paragraph(f"<b>TAXABLE AMOUNT Rs. {amount}</b>", normal_style))
    elements.append(Paragraph(f"<b>TOTAL AMOUNT Rs. {amount}</b>", normal_style))
    elements.append(Paragraph("<b>Received Amount Rs. 0</b>", normal_style))

    # Build PDF
    doc.build(elements)
    print(f"Created invoice PDF: {filename}")
    # ---- Upload to Drive ----
    DRIVE_FOLDER_ID = "1iamK2iKtDv6L91-5010fPv191rAX4qtE"
    uploaded = upload_to_drive(filename, DRIVE_FOLDER_ID)

    print(f"Uploaded {uploaded['name']} to Drive")


def main():
    output_folder = "/tmp/stayvista_invoices_pdf"

    # Open worksheet
    worksheet = gs_client.open("test data exp").worksheet("a")

    # Fetch all rows
    rows = worksheet.get_all_values()

    if not rows or len(rows) < 2:
        print("Error: Sheet is empty or has no data rows")
        return

    headers = rows[0]
    data_rows = rows[1:]

    print(f"Processing bills from Google Sheet...\n")
    print(f"Headers: {headers}\n")

    for row_num, row in enumerate(data_rows, start=2):
        # Pad row to avoid index errors
        row += [""] * (4 - len(row))

        booking_id    = row[0].strip()
        vendor_name   = row[1].strip()
        property_name = row[2].strip()
        amount        = row[3].strip()

        # Validation
        if not booking_id or not vendor_name or not amount:
            print(f"Skipping row {row_num}: Missing required data")
            continue

        print(f"Processing booking {booking_id}...")
        create_invoice_pdf(
            booking_id,
            vendor_name,
            property_name,
            amount,
            output_folder
        )


if __name__ == "__main__":
    main()