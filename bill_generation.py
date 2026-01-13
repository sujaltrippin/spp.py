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
    file_name = os.path.basename(file_path)

    # -------- 1. Find existing files with same name in folder --------
    query = (
        f"name = '{file_name}' "
        f"and '{drive_folder_id}' in parents "
        f"and mimeType = 'application/pdf' "
        f"and trashed = false"
    )

    existing = drive_service.files().list(
        q=query,
        spaces="drive",
        fields="files(id, name)",
        supportsAllDrives=True,
        includeItemsFromAllDrives=True
    ).execute().get("files", [])

    # -------- 2. Delete existing files --------
    for file in existing:
        drive_service.files().delete(
            fileId=file["id"],
            supportsAllDrives=True
        ).execute()
        print(f"Deleted existing file: {file['name']}")

    # -------- 3. Upload new file --------
    file_metadata = {
        "name": file_name,
        "parents": [drive_folder_id]
    }

    media = MediaFileUpload(
        file_path,
        mimetype="application/pdf",
        resumable=False
    )

    uploaded_file = drive_service.files().create(
        body=file_metadata,
        media_body=media,
        fields="id, name",
        supportsAllDrives=True
    ).execute()

    return uploaded_file

def create_invoice_pdf(unqid, booking_id, vendor_name, property_name, amount, output_folder):
    """
    Create a StayVista invoice PDF for a single booking
    """
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    filename = os.path.join(output_folder, f"{unqid}.pdf")
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
    # styles = getSampleStyleSheet()
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
    print('Setting up drivers...')
    nilesh_token = os.getenv("X_AUTH_TOKEN")
    print(nilesh_token)
    driver.execute_cdp_cmd("Network.enable", {})
    driver.execute_cdp_cmd(
        "Network.setExtraHTTPHeaders",
        {
            "headers": {
                "x-am-automation-key": nilesh_token
            }
        }
    )
    
    return driver

# ------------------ Login ------------------
def login_to_stayvista(driver, username, password, max_retries=5):
    for attempt in range(1, max_retries + 1):
        print(f"Login attempt {attempt}/{max_retries}")

        try:
            driver.get("https://admin.vistarooms.com/dashboard")

            WebDriverWait(driver, 15).until(
                EC.presence_of_element_located((By.NAME, "email"))
            ).clear()
            driver.find_element(By.NAME, "email").send_keys(username)

            WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.ID, "loginViaPasswordBtn"))
            ).click()

            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.NAME, "password"))
            ).clear()
            driver.find_element(By.NAME, "password").send_keys(password)

            WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.ID, "loginViaPasswordBtn"))
            ).click()

            WebDriverWait(driver, 20).until(
                EC.url_contains("dashboard")
            )

            print("✅ Login successful")
            return True

        except Exception as e:
            BASE_DIR = os.getcwd()
            screenshot_name = f"login_error_attempt_{attempt}.png"
            screenshot_path = os.path.join(BASE_DIR, screenshot_name)

            driver.save_screenshot(screenshot_path)

            print(f"Saved screenshot: {screenshot_path}")
            print(f"❌ Login failed on attempt {attempt}: {e}")

            if attempt < max_retries:
                time.sleep(3)
            else:
                print("Max login attempts reached")


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
def upload_bill(driver,unqid ,booking_id, bills_folder):
    if not bills_folder:
        print('no bill folder found...')
        return
    path = os.path.join(bills_folder, f"{unqid}.pdf")
    path = os.path.abspath(path)
    path = os.path.normpath(path)
    print(path)
    if not os.path.exists(path):
        print(f":x: Bill missing: {path}")
        return
    driver.find_element(By.ID, "bill").send_keys(path)
    

def move_row_to_log(gs_client, unqid):
    ss = gs_client.open("vista logs")
    source_ws = ss.worksheet("to be logged")
    log_ws = ss.worksheet("admin logs")

    rows = source_ws.get_all_values(
        value_render_option="UNFORMATTED_VALUE"
    )

    target = str(unqid).strip()

    for idx, row in enumerate(rows[1:], start=2):
        if not row:
            continue

        cell_value = str(row[0]).strip()

        if cell_value == target:
            # ---- prepend current date ----
            today = now_ist.strftime("%d-%b-%Y")
            new_row = [today] + row

            # ---- append to log ----
            log_ws.append_row(
                new_row,
                value_input_option="USER_ENTERED"
            )

            # ---- delete from source ----
            source_ws.delete_rows(idx)

            print(f"Moved SRNO {unqid} from 'to be logged' → 'admin logs'")
            return True

    print(f"SRNO {unqid} not found in sheet 'to be logged'")
    return False


def log(step):
    print(f"➡️ {step}", flush=True)


import time

def select2_search(driver, container_id, value, timeout=25):
    wait = WebDriverWait(driver, timeout)

    log(f"Select2 open: {container_id}")
    container = wait.until(EC.element_to_be_clickable((By.ID, container_id)))
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", container)
    container.click()

    time.sleep(0.6)  # ⏸ allow dropdown animation/render

    search = wait.until(EC.visibility_of_element_located(
        (By.CSS_SELECTOR, ".select2-container--open .select2-search__field")
    ))

    search.clear()
    time.sleep(0.3)

    search.send_keys(value)
    time.sleep(0.8)  # ⏸ allow results to load

    search.send_keys(Keys.RETURN)
    time.sleep(0.6)  # ⏸ allow selection to apply

    
# ------------------ Log Expense ------------------
def log_expense(driver,unqid, booking_id, head, comment, vendor, property_name, amount, cost_bearer, bills_folder):
    wait = WebDriverWait(driver, 30)

    try:
        # Expense Type
        log("Expense Type")
        select2_search(driver, "select2-expensetype-container", "F&B")

        # Expense Head
        log("Expense Head")
        select2_search(driver, "select2-expenshead-container", head)

        # Category / Comment
        log("Category / Comment")
        comment_el = wait.until(
            EC.visibility_of_element_located((By.ID, "expense_head_categoriespart"))
        )
        comment_el.clear()
        comment_el.send_keys(comment)

        # Vendor
        log("Vendor")
        select_vendor(driver, vendor)

        # Property
        log("Property")
        select2_search(driver, "select2-expense_villa_list-container", property_name)

        # Cost Bearer
        log("Cost Bearer")
        cost_select = Select(
            wait.until(EC.element_to_be_clickable((By.NAME, "cost_bearer")))
        )

        options = {opt.text.strip(): opt for opt in cost_select.options}

        if cost_bearer in options and options[cost_bearer].is_enabled():
            options[cost_bearer].click()
        elif "SV Managed" in options and options["SV Managed"].is_enabled():
            options["SV Managed"].click()
        else:
            raise Exception("No valid cost bearer available")

        # Invoice number
        log("Invoice Number")
        wait.until(EC.visibility_of_element_located(
            (By.ID, "invoice_number"))
        ).send_keys("1")

        # Bill date (timezone safe)
        log("Bill Date")
        d = now_ist
        driver.execute_script("""
            const el = document.getElementById('bill_date');
            const fixedDate = new Date(Date.UTC(arguments[0], arguments[1]-1, arguments[2]));
            el.valueAsDate = fixedDate;
            el.dispatchEvent(new Event('input', { bubbles: true }));
            el.dispatchEvent(new Event('change', { bubbles: true }));
        """, d.year, d.month, d.day)

        # Booking ID
        log("Booking ID")
        select2_search(driver, "select2-bookingid_expenses-container", booking_id)

        # Quantity & Rate
        log("Quantity & Rate")
        wait.until(EC.visibility_of_element_located(
            (By.NAME, "quantity[]"))
        ).send_keys("1")

        wait.until(EC.visibility_of_element_located(
            (By.NAME, "rate_per_unit[]"))
        ).send_keys(str(amount))

        # Tax
        log("Tax")
        set_tax_percentage(driver)

        # Upload Bill
        log("Upload Bill")
        upload_bill(driver, unqid, booking_id, bills_folder)

        # Submit
        log("Submit Expense")
        submit = wait.until(EC.presence_of_element_located((By.NAME, "submitButton")))
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", submit)
        wait.until(EC.element_to_be_clickable((By.NAME, "submitButton")))
        driver.execute_script("arguments[0].click();", submit)
        
        driver.execute_script("window.scrollTo(0, 0);")
        time.sleep(0.5) 

        old_url = driver.current_url

        # ✅ SUCCESS CONDITIONS (any one is enough)
        try:
            WebDriverWait(driver, 12).until(
                EC.any_of(
                    lambda d: d.current_url != old_url,
                    EC.presence_of_element_located((By.CLASS_NAME, "toast-success")),
                    EC.presence_of_element_located((By.CLASS_NAME, "alert-success")),
                    EC.invisibility_of_element_located((By.NAME, "submitButton"))
                )
            )
            log("✅ Expense submitted successfully")
            return True
        except TimeoutException:
            pass

        # Duplicate popup handling
        log("Check duplicate popup")
        if handle_duplicate_popup(driver):
            try:
                WebDriverWait(driver, 10).until(
                    EC.any_of(
                        lambda d: d.current_url != old_url,
                        EC.presence_of_element_located((By.CLASS_NAME, "toast-success")),
                        EC.presence_of_element_located((By.CLASS_NAME, "alert-success"))
                    )
                )
                log("✅ Expense submitted after duplicate confirm")
                return True
            except TimeoutException:
                pass


    except Exception as e:
        log(f"❌ Exception in log_expense: {e}")
        driver.save_screenshot(f"log_expense_error_{unqid}.png")
        raise

def upload_expenses(driver, bills_data, bills_folder, gs_client):
    for row in bills_data:
        if not navigate_to_expenses_add_page(driver):
            print(f"Could not open expense page for {row['booking_id']}")
            continue

        success = log_expense(
            driver,
            row["unqid"],
            row["booking_id"],
            row["head"],
            row["comment"],
            row["vendor"],
            row["property_name"],
            row["amount"],
            row["cost_bearer"],
            bills_folder
        )

        if success:
            move_row_to_log(gs_client, row["unqid"])
            print(f"✅ Expense logged for {row['booking_id']}")
        else:
            print(f"⚠️ Expense FAILED for {row['booking_id']} (unqid {row['unqid']})")

        time.sleep(2)


def generate_pdfs_from_gsheet(output_folder):
    worksheet = gs_client.open("vista logs").worksheet("to be logged") #Change INput sheet name here
    rows = worksheet.get("A:I",value_render_option="UNFORMATTED_VALUE")

    print(rows)
    headers = rows[0]
    data_rows = rows[1:]

    bill_rows = []

    for row in data_rows:
        row += [""] * (4 - len(row))

        unqid          = row[0].strip()
        booking_id     = row[1]
        head           = row[2].strip()
        comment        = row[3].strip()
        cost_bearer    = row[4].strip()
        amount         = row[5]
        tax            = row[6]
        vendor_name    = row[7].strip()
        property_name  = row[8].strip()

        if not booking_id or not vendor_name or not amount:
            continue

        create_invoice_pdf(
            unqid,
            booking_id,
            vendor_name,
            property_name,
            amount,
            output_folder
        )

        bill_rows.append({
            "unqid": unqid,
            "booking_id": booking_id,
            "head": head,
            "comment": comment,
            "cost_bearer": cost_bearer,
            "vendor": vendor_name,
            "property_name": property_name,
            "amount": amount,
        })

    return bill_rows
    

def update_status(gs_client, text, bg_color):
    """
    Updates status text & background color in to be logged!M1
    """
    try:
        ss = gs_client.open("vista logs")
        ws = ss.worksheet("to be logged")

        # ✅ Correct way to update single cell
        ws.update_acell("M1", text)
        ws.update_acell("M2", now_ist.strftime("%d-%b-%Y %I:%M %p"))

        sheets_service.spreadsheets().batchUpdate(
            spreadsheetId=ss.id,
            body={
                "requests": [
                    {
                        "repeatCell": {
                            "range": {
                                "sheetId": ws.id,
                                "startRowIndex": 0,
                                "endRowIndex": 1,
                                "startColumnIndex": 12,  # Column M
                                "endColumnIndex": 13
                            },
                            "cell": {
                                "userEnteredFormat": {
                                    "backgroundColor": bg_color,
                                    "textFormat": {
                                        "bold": True
                                    }
                                }
                            },
                            "fields": "userEnteredFormat(backgroundColor,textFormat)"
                        }
                    }
                ]
            }
        ).execute()

    except Exception as e:
        print("⚠️ Failed to update status cell:", e)


def main():
    try:
        update_status(
            gs_client,
            "Logging expense to admin module",
            {"red": 0.8, "green": 1, "blue": 0.8}
        )

        username = os.getenv("EMAIL")
        password = os.getenv("PASSWORD")
        bills_folder = "/tmp/stayvista_invoices_pdf"

        bills_data = generate_pdfs_from_gsheet(bills_folder)

        if not bills_data:
            raise Exception("No valid bills found")

        driver = setup_driver()

        if not login_to_stayvista(driver, username, password):
            raise Exception("Login failed")

        upload_expenses(driver, bills_data, bills_folder, gs_client)

        # ✅ SUCCESS
        update_status(
            gs_client,
            "All expenses logged",
            {"red": 0.8, "green": 1, "blue": 0.8}
        )

    except Exception as e:
        # ❌ FAILURE
        update_status(
            gs_client,
            "Failed to Log",
            {"red": 1, "green": 0.8, "blue": 0.8}
        )
        print("❌ Script failed:", e)
        raise

    finally:
        try:
            driver.quit()
        except:
            pass


if __name__ == "__main__":
    main()