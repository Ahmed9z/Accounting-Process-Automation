"""
Logistaas Invoice Reader v5
- Reads invoice numbers from Excel
- Searches each invoice on Logistaas
- Extracts product lines and maps them to Odoo product names
- Output columns: Invoice No | Product | Quantity | Unit Price | Tax
- Ready to import directly into Odoo
"""

import asyncio
import re
import os
import sys
import io
import difflib
import pandas as pd
import pdfplumber
from playwright.async_api import async_playwright


# ============================================================
# Find Chromium - works both as .py and as .exe
# ============================================================
def setup_chromium_path():
    # When running as exe, look next to the exe file
    if getattr(sys, 'frozen', False):
        base = os.path.dirname(sys.executable)
    else:
        base = os.path.dirname(os.path.abspath(__file__))

    chromium = os.path.join(base, "ms-playwright", "chromium-1217", "chrome-win64", "chrome.exe")
    if os.path.exists(chromium):
        # Tell Playwright where browsers are
        os.environ["PLAYWRIGHT_BROWSERS_PATH"] = os.path.join(base, "ms-playwright")
        print(f"[INFO] Using Chromium from: {chromium}")
        return chromium
    print(f"[WARN] Chromium not found at: {chromium}")
    return None


# ============================================================
# Port detection from MB/L
# ============================================================
def get_port(mbl: str) -> str:
    """Returns 'sokhna export', 'sokhna import', 'alexandria export', 'alexandria import', or 'unknown'"""
    mbl_clean = re.sub(r'[^A-Z0-9]', '', mbl.strip().upper())
    if mbl_clean.startswith('SOK'):
        return "sokhna export"
    elif mbl_clean.startswith('ALY'):
        return "alexandria export"
    elif 'SOK' in mbl_clean:
        return "sokhna import"
    elif 'ALY' in mbl_clean:
        return "alexandria import"
    return "unknown"


# ============================================================
# Map Logistaas description -> Odoo product name
# ============================================================
def map_to_odoo_product(description: str, port: str) -> str:
    desc = description.lower().strip()
    is_alex = "alexandria" in port
    is_sokhna = "sokhna" in port
    is_export = "export" in port
    is_alex_import = port == "alexandria import"

    if "demurrage" in desc:
        return "Port Demurrage Charge"

    elif "storage" in desc or "bonded" in desc:
        return "Import Storage Alexandria Port" if is_alex_import else "Storage"

    # THC must come before DTHC check
    elif re.search(r'^t+hc', desc) and "dthc" not in desc:
        return "THC - (EGYPT) Alexandria Port" if is_alex else "THC"

    elif "dthc" in desc:
        return "DTHC Alex" if is_alex else "DTHC"

    elif "admin fee" in desc or "administration fee" in desc:
        return "Admin Fee Alex" if is_alex else "Admin Fee"

    elif "delivery order" in desc:
        return "Delivery Order Alex" if is_alex else "Delivery Order Fees"

    elif "nafeza" in desc:
        return "Nafeza Fees"

    elif "agency" in desc:
        return "Agency DOC fee Alex" if is_alex else "Agency Fees"

    elif "container movement & security" in desc or "movement & security" in desc:
        return "Container movement Alex"

    elif "container movement" in desc or "movement in yard" in desc:
        return "Container movement in yard"

    elif "container inspection" in desc:
        return "Container Inspection Charges"

    elif "cleaning" in desc:
        return "Cleaning Alex" if is_alex else "Miscellaneous Charge"

    elif "decree" in desc:
        return "DECREE 800" if is_alex else "Miscellaneous Charge"

    elif "equipment recovery" in desc:
        return "Equipment Recovery Handling Alex" if is_alex else "Miscellaneous Charge"

    elif "special service" in desc:
        return "SPECIAL SERVICE REQUEST Alexandria Port" if is_alex else "Miscellaneous Charge"

    elif "documentation fee" in desc or "doc fee" in desc or "repositioning" in desc:
        return "DOC Fee & Repositioning Alex" if is_alex else "Miscellaneous Charge"

    elif "customs duty" in desc or "customs inspection" in desc or "customs" in desc:
        return "CUSTOMS INSPECTION Alexandria Port"

    elif "equipment release" in desc:
        return "Equipment Release Charge"

    elif "ocean freight" in desc:
        return "Ocean Freight Charges"

    elif "pti" in desc or "pre-trip" in desc or "pre trip" in desc:
        return "PTI (Pre-Trip Inspection)"

    elif "reefer monitoring" in desc:
        return "Reefer Monitoring Alexandria Port" if is_alex else "Reefer Monitoring"

    elif "survey" in desc:
        return "Survey Revenue"

    elif "transportation" in desc:
        return "Transportation - Unifeeder"

    elif "uploading after cutoff" in desc or "after cutoff" in desc:
        # Differentiate based on bracket content
        bracket_match = re.search(r'\((.*?)\)', description)
        bracket = bracket_match.group(1).lower().strip() if bracket_match else ""
        if "after" in bracket and "berth" in bracket:
            return "After Vessel Berth"
        else:
            return "Before Vessel Berth"

    elif "shutout" in desc or "shut out" in desc or "shut-out" in desc:
        return "Shut Out"

    elif "late payment" in desc:
        return "Late Payment Penalty-Unifeeder"

    elif "telex release" in desc:
        return "TELEX RELEASE Alexandria Port" if is_alex else "Telex Release Fees"

    elif "detention" in desc:
        return "Detention Collection Commission"

    elif "seal" in desc:
        return "SEAL FEE Alexandria Port" if is_alex else "Seal Fees"

    elif "bl fee" in desc or "b/l fee" in desc:
        return "BL Fees"

    else:
        print(f"  [INFO] Unknown description mapped to Miscellaneous Charge: '{description}'")
        return "Miscellaneous Charge"


# ============================================================
# Smart collapse: only collapse if text is genuinely doubled
# ============================================================
def maybe_collapse(text: str) -> str:
    if not text or len(text) < 4:
        return text
    # Count consecutive duplicate character pairs
    pairs = sum(
        1 for i in range(len(text) - 1)
        if text[i].lower() == text[i + 1].lower()
    )
    ratio = pairs / max(len(text) - 1, 1)
    # Only collapse if more than 40% of chars are duplicated pairs
    if ratio > 0.4:
        result = [text[0]]
        for char in text[1:]:
            if char.lower() != result[-1].lower():
                result.append(char)
        return ''.join(result)
    return text


# ============================================================
# Clean price: "80.0 USD" -> "80", "385.0 USD" -> "385"
# ============================================================
def clean_price(value: str) -> str:
    cleaned = re.sub(r'[^0-9.]', '', str(value).strip())
    if not cleaned:
        return "0"
    try:
        f = float(cleaned)
        return str(int(f)) if f == int(f) else str(round(f, 2))
    except:
        return cleaned


# ============================================================
# Clean quantity
# ============================================================
def clean_qty(value: str) -> str:
    cleaned = re.sub(r'[^0-9.]', '', str(value).strip())
    if not cleaned:
        return "1"
    try:
        f = float(cleaned)
        return str(int(f)) if f == int(f) else str(round(f, 2))
    except:
        return cleaned


# ============================================================
# Normalize tax value
# ============================================================
def normalize_tax(value: str) -> str:
    v = str(value).strip()
    if re.search(r'14', v):
        return "VAT 14%"
    # Empty or no tax = Zero Rated 0%
    return "Zero Rated 0%"


# ============================================================
# Pick Excel file
# ============================================================
def pick_excel_file(title: str) -> str:
    try:
        from tkinter import Tk, filedialog
        root = Tk()
        root.withdraw()
        root.attributes('-topmost', True)
        filepath = filedialog.askopenfilename(
            title=title,
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        root.destroy()
        return filepath
    except:
        # tkinter not available - ask for manual path
        print(f"\n[INFO] {title}")
        print("Please type or paste the full path to your Excel file:")
        print("Example: C:\\Users\\akhai\\Downloads\\invoices.xlsx")
        path = input("Path: ").strip().strip('"')
        return path


# ============================================================
# Read invoice numbers from Excel
# ============================================================
def read_invoice_numbers(filepath: str) -> list:
    print(f"\n[INFO] Reading: {os.path.basename(filepath)}")
    df = pd.read_excel(filepath)
    print(f"[INFO] Columns: {list(df.columns)}")
    print(df.head())

    print("\nWhich column has invoice numbers?")
    for i, col in enumerate(df.columns):
        print(f"  {i} - {col}")

    col_index = int(input("Enter column number: ").strip())
    col_name = df.columns[col_index]

    invoices = df[col_name].dropna().astype(str).str.strip().tolist()
    invoices = [x for x in invoices if x and x.lower() != 'nan']

    print(f"\n[INFO] Found {len(invoices)} invoices (in original Excel order):")
    for inv in invoices:
        print(f"  - {inv}")
    return invoices


# ============================================================
# Search invoice on Logistaas and get PDF URL
# ============================================================
async def search_and_get_pdf_url(page, serial_no: str, first_invoice: bool = False) -> str:
    if first_invoice:
        print(f"  [STEP 1] Going to invoices page...")
        await page.goto("https://internal-company-website/invoices")
        await page.wait_for_load_state("networkidle")
        await page.wait_for_timeout(800)

        print(f"  [STEP 2] Opening search panel...")
        try:
            await page.click('svg[viewBox="0 0 512 512"]')
            await page.wait_for_timeout(700)
        except:
            try:
                await page.evaluate("""
                    () => {
                        const svgs = document.querySelectorAll('svg');
                        for (const svg of svgs) {
                            if (svg.getAttribute('viewBox') === '0 0 512 512') {
                                svg.parentElement.click();
                                return;
                            }
                        }
                    }
                """)
                await page.wait_for_timeout(700)
            except Exception as e:
                print(f"  [WARN] Could not click search icon: {e}")
                input("  >> Please open the search panel manually, then press Enter...")

        print(f"  [STEP 4] Setting Fully Paid filter (first invoice only)...")
        try:
            await page.click('input.select2-search__field[role="searchbox"]')
            await page.wait_for_timeout(500)
            await page.click('.select2-results__option:has-text("Fully Paid")')
            await page.wait_for_timeout(300)
            print("  [OK] Fully Paid selected")
        except Exception as e:
            print(f"  [WARN] Filter: {e}")
    else:
        # Reopen the search panel for each subsequent invoice
        print(f"  [STEP 2] Reopening search panel...")
        try:
            # Try fastest approach first - click via JavaScript directly
            clicked = await page.evaluate("""
                () => {
                    const svgs = document.querySelectorAll('svg');
                    for (const svg of svgs) {
                        if (svg.getAttribute('viewBox') === '0 0 512 512') {
                            svg.parentElement.click();
                            return true;
                        }
                    }
                    return false;
                }
            """)
            await page.wait_for_timeout(300)
            if not clicked:
                await page.click('svg[viewBox="0 0 512 512"]')
                await page.wait_for_timeout(300)
        except Exception as e:
            print(f"  [WARN] Could not reopen search: {e}")
            input("  >> Please click the magnifying glass manually, then press Enter...")

    print(f"  [STEP 3] Entering serial number: {serial_no}")
    await page.fill('input[name="filterrific[with_serial_no]"]', serial_no)
    await page.wait_for_timeout(300)

    print(f"  [STEP 5] Clicking Search...")
    await page.click('input[name="commit"]')
    await page.wait_for_load_state("networkidle")
    await page.wait_for_timeout(1200)

    print(f"  [STEP 6] Getting PDF URL...")
    attachment_link = await page.query_selector('a:has(i.fa-file-pdf)')
    if not attachment_link:
        print("  [WARN] No PDF found - invoice may not exist or not Fully Paid")
        return ""

    href = await attachment_link.get_attribute('href')
    if href.startswith('/'):
        href = f"https://internal-company-website{href}"
    print(f"  [OK] PDF URL: {href}")
    return href


# ============================================================
# Download PDF and extract data
# ============================================================
async def extract_pdf_data(page, pdf_url: str) -> dict:
    result = {"port": "unknown", "items": []}

    print(f"  [STEP 7] Downloading PDF...")
    pdf_data = None
    for attempt in range(1, 4):  # try up to 3 times
        try:
            response = await page.context.request.get(pdf_url)
            pdf_data = await response.body()
            print(f"  [OK] Downloaded {len(pdf_data)} bytes")
            break
        except Exception as download_err:
            if attempt < 3:
                print(f"  [WARN] Download attempt {attempt} failed: {download_err} - retrying in 3s...")
                await page.wait_for_timeout(3000)
            else:
                print(f"  [ERROR] Download failed after 3 attempts: {download_err}")
                return result
    if not pdf_data:
        return result
    try:

        with pdfplumber.open(io.BytesIO(pdf_data)) as pdf:
            full_text = ""
            all_tables = []
            for pdf_page in pdf.pages:
                full_text += pdf_page.extract_text() or ""
                all_tables.extend(pdf_page.extract_tables())

        print(f"  [OK] Parsed - {len(full_text)} chars, {len(all_tables)} tables")

        # Collapse full text for MB/L search
        full_text_clean = maybe_collapse(full_text)

        # Detect port from MB/L - regex handles both MB/L and MBB//LL (doubled PDF chars)
        mbl_match = re.search(
            r'M+B+[/\\]+L+\s*[:\-]?\s*([A-Z0-9][A-Z0-9/]+)',
            full_text_clean, re.IGNORECASE
        )
        if mbl_match:
            raw_mbl = mbl_match.group(1).strip()
            result["port"] = get_port(raw_mbl)
            print(f"  [OK] MB/L: {raw_mbl} -> Port: {result['port']}")
        else:
            print(f"  [WARN] MB/L not found - port unknown")

        # Extract table rows
        skip_keywords = ["description", "total", "subtotal", "taxes", "tax", "rate per"]
        seen = set()

        for table in all_tables:
            for row in table:
                if not row or not row[0]:
                    continue

                cells = [maybe_collapse(str(c).strip()) if c else "" for c in row]
                first = cells[0].lower().strip()

                if any(k in first for k in skip_keywords):
                    continue
                if len(cells[0].strip()) < 3:
                    continue

                # Check bracket content for "late payment" BEFORE stripping
                # Uses fuzzy matching to handle typos/extra spaces
                bracket_match = re.search(r'\((.*?)\)', cells[0])
                bracket_content = bracket_match.group(1).lower().strip() if bracket_match else ""

                # Fuzzy check: does bracket content resemble "late payment"?
                late_payment_detected = False
                if bracket_content:
                    # Direct keyword check first
                    if "late" in bracket_content and "pay" in bracket_content:
                        late_payment_detected = True
                    else:
                        # Fuzzy match against "late payment"
                        score = difflib.SequenceMatcher(
                            None, bracket_content, "late payment"
                        ).ratio()
                        if score >= 0.7:
                            late_payment_detected = True

                # Strip bracketed descriptions e.g. "Storage (20ft Dry)" -> "Storage"
                clean_desc = re.sub(r'\s*\(.*?\)', '', cells[0]).strip()
                if not clean_desc:
                    clean_desc = cells[0]

                # Override mapping for late payment
                if late_payment_detected:
                    clean_desc = "Late Payment"

                # Deduplicate based on cleaned description + price combination
                # Same description with different price = different line (e.g. storage periods)
                price_for_key = clean_price(cells[2]) if len(cells) > 2 else "0"
                key = f"{clean_desc.strip()[:30]}_{price_for_key}"
                if key in seen:
                    continue
                seen.add(key)

                if len(cells) >= 4:
                    price_val = clean_price(cells[2]) if len(cells) > 2 else "0"
                    qty_val   = clean_qty(cells[3])   if len(cells) > 3 else "1"

                    # Skip lines with zero price (e.g. free storage days)
                    try:
                        if float(price_val) == 0:
                            print(f"  [SKIP] Zero price line skipped: {clean_desc[:55]}")
                            continue
                    except:
                        pass

                    result["items"].append({
                        "description": clean_desc,
                        "unit_price":  price_val,
                        "quantity":    qty_val,
                        "tax":         normalize_tax(cells[4] if len(cells) > 4 else ""),
                    })
                    print(f"  [ITEM] {clean_desc[:55]} | price:{price_val} | qty:{qty_val}")

    except Exception as e:
        print(f"  [ERROR] {e}")

    return result


# ============================================================
# Main
# ============================================================
async def main():
    print("=" * 55)
    print("   Logistaas Invoice Reader v5")
    print("=" * 55)

    chromium_path = setup_chromium_path()

    excel_path = pick_excel_file("Select Excel file with invoice numbers")
    if not excel_path:
        print("[ERROR] No file selected.")
        input("Press Enter to exit...")
        return

    invoice_numbers = read_invoice_numbers(excel_path)
    if not invoice_numbers:
        print("[ERROR] No invoice numbers found.")
        input("Press Enter to exit...")
        return

    print(f"\n[INFO] {len(invoice_numbers)} invoices to process")
    print("[INFO] You can minimize the browser after login - do not click inside it")
    input("Press Enter to open the browser...")

    all_results = []

    async with async_playwright() as p:
        browser = await p.chromium.launch(
            headless=False,
            slow_mo=50,
            executable_path=chromium_path
        )
        context = await browser.new_context()
        page = await context.new_page()

        await page.goto("https://internal-company-website/invoices")
        await page.wait_for_load_state("networkidle")

        if "login" in page.url or "signin" in page.url:
            print("\n[INFO] Please login first")
            input("Login in the browser, then press Enter...")
            await page.wait_for_load_state("networkidle")

        for idx, serial_no in enumerate(invoice_numbers, 1):
            print(f"\n{'='*55}")
            print(f"[{idx}/{len(invoice_numbers)}] {serial_no}")
            print(f"{'='*55}")

            try:
                pdf_url = await search_and_get_pdf_url(page, serial_no, first_invoice=(idx == 1))

                if not pdf_url:
                    all_results.append({
                        "Reference":                serial_no,
                        "Invoice lines / Product":  "NOT FOUND",
                        "Invoice lines / Unit Price": "",
                        "Invoice lines / Quantity": "",
                        "Invoice lines / Taxes":    "",
                    })
                    continue

                data = await extract_pdf_data(page, pdf_url)
                port = data["port"]
                print(f"  Port: {port} | Items: {len(data['items'])}")

                if data["items"]:
                    for item in data["items"]:
                        odoo_name = map_to_odoo_product(item["description"], port)
                        if not odoo_name:
                            print(f"  [INFO] '{item['description']}' has no mapping for port '{port}' -> Miscellaneous Charge")
                            odoo_name = "Miscellaneous Charge"
                        all_results.append({
                            "Reference":                serial_no,
                            "Invoice lines / Product":  odoo_name,
                            "Invoice lines / Unit Price": item["unit_price"],
                            "Invoice lines / Quantity": item["quantity"],
                            "Invoice lines / Taxes":    item["tax"],
                        })
                else:
                    all_results.append({
                        "Reference":                serial_no,
                        "Invoice lines / Product":  "No items found",
                        "Invoice lines / Unit Price": "",
                        "Invoice lines / Quantity": "",
                        "Invoice lines / Taxes":    "",
                    })

            except Exception as e:
                print(f"  [ERROR] {serial_no}: {e}")
                all_results.append({
                    "Reference":                serial_no,
                    "Invoice lines / Product":  f"ERROR: {e}",
                    "Invoice lines / Unit Price": "",
                    "Invoice lines / Quantity": "",
                    "Invoice lines / Taxes":    "",
                })

        await browser.close()

    # Save logistaas_results.xlsx
    if not all_results:
        print("\n[WARN] No data extracted.")
        input("\nPress Enter to exit...")
        return

    results_path = os.path.join(os.path.dirname(excel_path), "logistaas_results.xlsx")
    try:
        pd.DataFrame(all_results).to_excel(results_path, index=False)
        print(f"\n[DONE] Results saved to: {results_path}")
    except PermissionError:
        # File is open in Excel - save with timestamp instead
        from datetime import datetime
        ts = datetime.now().strftime("%H%M%S")
        results_path = os.path.join(os.path.dirname(excel_path), f"logistaas_results_{ts}.xlsx")
        pd.DataFrame(all_results).to_excel(results_path, index=False)
        print(f"\n[DONE] File was open in Excel - saved as: {results_path}")
    print(f"[DONE] Total rows: {len(all_results)}")
    input("\nPress Enter to exit...")


if __name__ == "__main__":
    try:
        asyncio.run(main())
    except Exception as e:
        print(f"\n[CRASH] {e}")
        import traceback
        traceback.print_exc()
        input("\nPress Enter to exit...")
