import csv
import os
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import TimeoutException


# ------------------ Setup Driver (HEADLESS) ------------------
def setup_driver():
    chrome_options = Options()
    # chrome_options.add_argument("--headless=new")
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
        print(":white_tick: Login successful")
        return True
    except Exception as e:
        print(":x: Login failed:", e)
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
def upload_bill(driver, booking_id, bills_folder):
    if not bills_folder:
        return
    path = os.path.join(bills_folder, f"{booking_id}.pdf")
    if not os.path.exists(path):
        print(f":x: Bill missing: {path}")
        return
    driver.find_element(By.ID, "bill").send_keys(path)
    
# ------------------ Log Expense ------------------
def log_expense(driver, booking_id, vendor, property_name, amount, sub_desc, bills_folder):
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
    Select(driver.find_element(By.NAME, "cost_bearer")).select_by_visible_text("VISTA")
    driver.find_element(By.ID, "invoice_number").send_keys("1")
    driver.find_element(By.ID, "bill_date").send_keys(time.strftime("%d-%m-%Y"))
    # Booking ID
    driver.find_element(By.ID, "select2-bookingid_expenses-container").click()
    search = driver.find_element(By.CLASS_NAME, "select2-search__field")
    search.send_keys(booking_id)
    time.sleep(0.5)
    search.send_keys(Keys.RETURN)
    driver.find_element(By.NAME, "quantity[]").send_keys("1")
    driver.find_element(By.NAME, "rate_per_unit[]").send_keys(amount)
    set_tax_percentage(driver)
    upload_bill(driver, booking_id, bills_folder)
    submit = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.NAME, "submitButton"))
    )
    driver.execute_script("arguments[0].click();", submit)
    # :white_tick: Handle duplicate popup
    handle_duplicate_popup(driver)
    WebDriverWait(driver, 15).until(EC.url_contains("expenses"))
    
# ------------------ Main ------------------
def main():
    username = "sujal.uttekar@stayvista.com"
    password = "Sujal@2025"
    csv_file = "bills.csv"
    bills_folder = r"C:\Users\sujal\Untitled Folder 9\stayvista_invoices_pdf"
    driver = setup_driver()
    try:
        if not login_to_stayvista(driver, username, password):
            return
        with open(csv_file, newline="", encoding="utf-8") as file:
            reader = csv.DictReader(file)
            for row in reader:
                booking_id = row["booking_id"].strip()
                vendor = row["vendor_name"].strip()
                property_name = row["property_name"].strip()
                amount = row["amount"].strip()
                sub_desc = row.get("sub", f"Expense for booking {booking_id}")
                if not navigate_to_expenses_add_page(driver):
                    continue
                log_expense(
                    driver,
                    booking_id,
                    vendor,
                    property_name,
                    amount,
                    sub_desc,
                    bills_folder
                )
                print(f":white_tick: Expense logged for {booking_id}")
                time.sleep(2)
    finally:
        driver.quit()
        print("Browser closed")
# ------------------ Run ------------------
if __name__ == "__main__":
    main()
    
    
# def main():
#     output_folder = "/tmp/stayvista_invoices_pdf"
#     username = "sujal.uttekar@stayvista.com"
#     password = "Sujal@2025"

#     # Open worksheet 
#     worksheet = gs_client.open("test data exp").worksheet("a")

#     # Fetch all rows
#     rows = worksheet.get_all_values()

#     if not rows or len(rows) < 2:
#         print("Error: Sheet is empty or has no data rows")
#         return

#     headers = rows[0]
#     data_rows = rows[1:]

#     print(f"Processing bills from Google Sheet...\n")
#     print(f"Headers: {headers}\n")

#     for row_num, row in enumerate(data_rows, start=2):
#         # Pad row to avoid index errors
#         row += [""] * (4 - len(row))

#         booking_id    = row[0].strip()
#         vendor_name   = row[1].strip()
#         property_name = row[2].strip()
#         amount        = row[3].strip()

#         # Validation
#         if not booking_id or not vendor_name or not amount:
#             print(f"Skipping row {row_num}: Missing required data")
#             continue

#         print(f"Processing booking {booking_id}...")
#         create_invoice_pdf(
#             booking_id,
#             vendor_name,
#             property_name,
#             amount,
#             output_folder
#         )
