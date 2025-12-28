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
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import TimeoutException
import time
from google.oauth2 import service_account

# Load environment variables
load_dotenv()
ist = pytz.timezone('Asia/Kolkata')
now_ist = datetime.now(ist)

with open("credentials.json", "w") as f:
    f.write(os.getenv("GOOGLE_SHEET_CONNECTOR"))

# Google Sheets Auth
scope = ["https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/spreadsheets"]


creds = service_account.Credentials.from_service_account_file(
    "credentials.json",
    scopes=scope
)

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
        fields='id, name',
        supportsAllDrives=True
    ).execute()

    return uploaded_file


def create_invoice_pdf(srno, booking_id, vendor_name, property_name, amount, output_folder):
    """
    Create a StayVista invoice PDF for a single booking
    """
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    filename = os.path.join(output_folder, f"{srno}_{booking_id}.pdf")
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
    DRIVE_FOLDER_ID = "1St6hd_7veFTcaK7dAJC29yfmcDNQo4wf"
    uploaded = upload_to_drive(filename, DRIVE_FOLDER_ID)

    print(f"Uploaded {uploaded['name']} to Drive")

# ------------------ Setup Driver (HEADLESS) ------------------
def setup_driver():
    chrome_options = Options()
    chrome_options.add_argument("--headless=new")
    chrome_options.add_argument("--window-size=1920,1080")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--disable-notifications")
    chrome_options.add_argument("--ignore-certificate-errors")
    chrome_options.add_argument("--disable-blink-features=AutomationControlled")
    chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
    chrome_options.add_experimental_option("useAutomationExtension", False)
    chrome_options.add_experimental_option("prefs", {
        "download.prompt_for_download": False,
        "safebrowsing.enabled": True
    })
    driver = webdriver.Chrome(options=chrome_options)
    driver.execute_cdp_cmd(
        "Page.addScriptToEvaluateOnNewDocument",
        {
            "source": """
                Object.defineProperty(navigator, 'webdriver', {
                    get: () => undefined
                })
            """
        }
    )
    return driver

# ------------------ Login ------------------
def login_to_stayvista(driver, username, password):
    print("Logging in...")
    driver.get("https://admin.vistarooms.com/dashboard")
    try:
        WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.NAME, "email"))
        ).send_keys(username)
        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "loginViaPasswordBtn"))
        ).click()
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.NAME, "password"))
        ).send_keys(password)
        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "loginViaPasswordBtn"))
        ).click()
        WebDriverWait(driver, 20).until(EC.url_contains("dashboard"))
        print("✅ Login successful")
        return True
    except Exception as e:
        print("❌ Login failed:", e)
        driver.save_screenshot("login_error.png")
        return False
    
# ------------------ Navigate ------------------
def navigate_to_expenses_add_page(driver):
    try:
        driver.get("https://admin.vistarooms.com/expenses/log")
        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.ID, "select2-expensetype-container"))
        )
        return True
    except Exception as e:
        print(":x: Navigation failed:", e)
        return False
# ------------------ Handle Duplicate Popup ------------------
def handle_duplicate_popup(driver, timeout=6):
    try:
        yes_btn = WebDriverWait(driver, timeout).until(
            EC.element_to_be_clickable((By.ID, "btnYes"))
        )
        driver.execute_script("arguments[0].click();", yes_btn)
        print(":warning: Duplicate popup detected — clicked YES")
        time.sleep(1)
        return True
    except TimeoutException:
        return False
    
def wait_for_redirect(driver, old_url, timeout=10):
    try:
        WebDriverWait(driver, timeout).until(
            lambda d: d.current_url != old_url
        )
        return True
    except TimeoutException:
        return False
        
# ------------------ Select Vendor ------------------
def select_vendor(driver, vendor_name):
    driver.find_element(By.ID, "select2-vendor_name-container").click()
    search = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.CLASS_NAME, "select2-search__field"))
    )
    search.send_keys(vendor_name)
    time.sleep(0.5)
    search.send_keys(Keys.RETURN)
    
# ------------------ Tax ------------------
def set_tax_percentage(driver):
    Select(
        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.NAME, "tax_percentage[]"))
        )
    ).select_by_visible_text("0")
    
# ------------------ Upload Bill ------------------
def upload_bill(driver,srno ,booking_id, bills_folder):
    if not bills_folder:
        return
    path = os.path.join(bills_folder, f"{srno}_{booking_id}.pdf")
    path = os.path.abspath(path)
    if not os.path.exists(path):
        print(f":x: Bill missing: {path}")
        return
    driver.find_element(By.ID, "bill").send_keys(path)
    
def move_row_to_log(gs_client, srno):
    ss = gs_client.open("test data exp")
    source_ws = ss.worksheet("a")
    log_ws = ss.worksheet("log")

    rows = source_ws.get_all_values()

    for idx, row in enumerate(rows[1:], start=2):  # row index in sheet
        if row and row[0].strip() == str(srno):
            # Append to log sheet
            log_ws.append_row(row, value_input_option="USER_ENTERED")

            # Remove from source sheet
            source_ws.delete_rows(idx)

            print(f"Moved SRNO {srno} from 'a' → 'log'")
            return True

    print(f"SRNO {srno} not found in sheet 'a'")
    return False

    
# ------------------ Log Expense ------------------
def log_expense(driver,srno, booking_id, vendor, property_name, amount, sub_desc, bills_folder):
    # Expense Type
    driver.find_element(By.ID, "select2-expensetype-container").click()
    search = driver.find_element(By.CLASS_NAME, "select2-search__field")
    search.send_keys("f&b")
    search.send_keys(Keys.RETURN)
    # Expense Head
    driver.find_element(By.ID, "select2-expenshead-container").click()
    search = driver.find_element(By.CLASS_NAME, "select2-search__field")
    search.send_keys("Cook Arranged")
    time.sleep(0.5)
    search.send_keys(Keys.RETURN)
    # Category
    driver.find_element(By.ID, "expense_head_categoriespart").send_keys(sub_desc)
    # Vendor
    select_vendor(driver, vendor)
    # Property
    driver.find_element(By.ID, "select2-expense_villa_list-container").click()
    search = driver.find_element(By.CLASS_NAME, "select2-search__field")
    search.send_keys(property_name)
    time.sleep(0.5)
    search.send_keys(Keys.RETURN)
    # Cost bearer
    # Select(driver.find_element(By.NAME, "cost_bearer")).select_by_visible_text("VISTA")
    # Cost bearer (VISTA → fallback to SV Managed)
    select = Select(driver.find_element(By.NAME, "cost_bearer"))

    vista_option = None
    sv_option = None

    for opt in select.options:
        text = opt.text.strip()
        if text == "VISTA":
            vista_option = opt
        elif text == "SV Managed":
            sv_option = opt

    if vista_option and vista_option.is_enabled():
        select.select_by_visible_text("VISTA")
    elif sv_option and sv_option.is_enabled():
        select.select_by_visible_text("SV Managed")
    else:
        raise Exception("No valid cost bearer available (VISTA / SV Managed)")
        # sheet logs ERROR

    driver.find_element(By.ID, "invoice_number").send_keys("1")
    driver.find_element(By.ID, "bill_date").send_keys(now_ist.strftime("%d-%m-%Y"))
    # Booking ID
    driver.find_element(By.ID, "select2-bookingid_expenses-container").click()
    search = driver.find_element(By.CLASS_NAME, "select2-search__field")
    search.send_keys(booking_id)
    time.sleep(0.5)
    search.send_keys(Keys.RETURN)
    driver.find_element(By.NAME, "quantity[]").send_keys("1")
    driver.find_element(By.NAME, "rate_per_unit[]").send_keys(amount)
    set_tax_percentage(driver)
    upload_bill(driver,srno ,booking_id, bills_folder)
    submit = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.NAME, "submitButton"))
    )
    old_url = driver.current_url
    driver.execute_script("arguments[0].click();", submit)
    if wait_for_redirect(driver, old_url, timeout=8):
        return True  # ✅ success 
    if handle_duplicate_popup(driver):
        # After clicking YES, redirect MUST happen
        if wait_for_redirect(driver, old_url, timeout=8):
            return True  
    # :white_tick: Handle duplicate popup
    # WebDriverWait(driver, 15).until(EC.url_contains("expenses"))
    return False

def upload_expenses(driver, bills_data, bills_folder, gs_client):
    for row in bills_data:
        if not navigate_to_expenses_add_page(driver):
            print(f"Could not open expense page for {row['booking_id']}")
            continue

        success = log_expense(
            driver,
            row["sr_no"],
            row["booking_id"],
            row["vendor"],
            row["property_name"],
            row["amount"],
            row["sub"],
            bills_folder
        )

        if success:
            move_row_to_log(gs_client, row["sr_no"])
            print(f"✅ Expense logged for {row['booking_id']}")
        else:
            print(f"⚠️ Expense FAILED for {row['booking_id']} (SRNO {row['sr_no']})")

        time.sleep(2)


def generate_pdfs_from_gsheet(output_folder):
    worksheet = gs_client.open("test data exp").worksheet("a") #Change INput sheet name here
    rows = worksheet.get_all_values()

    headers = rows[0]
    data_rows = rows[1:]

    bill_rows = []

    for row in data_rows:
        row += [""] * (4 - len(row))

        srno          = row[0].strip()
        booking_id    = row[1].strip()
        vendor_name   = row[2].strip()
        property_name = row[3].strip()
        amount        = row[4].strip()

        if not booking_id or not vendor_name or not amount:
            continue

        create_invoice_pdf(
            srno,
            booking_id,
            vendor_name,
            property_name,
            amount,
            output_folder
        )

        bill_rows.append({
            "sr_no": srno,
            "booking_id": booking_id,
            "vendor": vendor_name,
            "property_name": property_name,
            "amount": amount,
            "sub": f"Expense for booking {booking_id}"
        })

    return bill_rows
    

def main():
    username = "sujal.uttekar@stayvista.com"
    password = "Sujal@2025"
    bills_folder = "/tmp/stayvista_invoices_pdf"

    os.makedirs(bills_folder, exist_ok=True)

    # 1️ Generate PDFs + data from Google Sheet
    bills_data = generate_pdfs_from_gsheet(bills_folder)

    if not bills_data:
        print("No valid bills found")
        return

    # 2 Start Selenium
    driver = setup_driver()

    try:
        if not login_to_stayvista(driver, username, password):
            return

        # 3 Upload expenses
        upload_expenses(driver, bills_data, bills_folder, gs_client)

    finally:
        driver.quit()
        print("Browser closed")

if __name__ == "__main__":
    main()