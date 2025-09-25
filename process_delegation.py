import sys
import re
import time
import fitz  # PyMuPDF
import xlwings as xw

# --- Constants ---
WORKBOOK_PATH = r"\\reiltys\iomgroot\DeptShare_DHSS_Nobles\Management\Director of Finance, Performance & Delivery\16. Manx Care\Financials\Reporting\Delegations\Delegation PDFs\Delegated Officers & Budget Holders - Full List .xlsm"
TARGET_SHEET_NAME = "FULL LIST"

FINANCIAL_HEADINGS = [
    "Issuing orders/ approving payment commitments",
    "Authorising Invoices - in relation to orders/commitments made within the Budget Area (where the Delegated Officer has not authorised the order)",
    "Authorising Invoices/Payments – in relation to orders/commitments authorised external to the Budget Area",
    "Purchase Card Single Transaction Limit",
    "Purchase Card Monthly Limit",
    "Approval of the write - Off of assets, inventory or stock",
    "Authorisation of Imprest / Petty Cash Payments",
    "Approval of Off – Island Travel Requests",
    "Approval of overtime & enhanced payments",
    "Approval of mileage, travel & subsistence claims",
    "Approval of Credit Facilities to 3rd Parties",
    "Approval to write-off individual debts"
]


def extract_page_text(pdf_path, page_number):
    """Extract text from a specific page number (zero-based) in a PDF using PyMuPDF."""
    print(f"Extracting text from page {page_number + 1}...")
    with fitz.open(pdf_path) as doc:
        if page_number < len(doc):
            text = doc[page_number].get_text("text")
            print(f"Page {page_number + 1} text extracted, length: {len(text)} characters")
            return text
        else:
            raise IndexError(f"PDF has only {len(doc)} pages. Requested page {page_number + 1} not found.")


def extract_cost_centre_codes(page_text: str) -> dict:
    """
    Extract all digit sequences 8 or more digits long from page text.
    Joins results with ' // ' separator.
    """
    print("---- Raw Cost Centre Codes Page Text Start ----")
    print(page_text[:500])  # print first 500 chars max to avoid flooding console
    print("---- Raw Cost Centre Codes Page Text End ----")

    clean_text = page_text.replace('\u00A0', ' ').replace('\u200B', '')
    clean_text = re.sub(r'[^\x00-\x7F]', '', clean_text)  # remove non-ASCII
    clean_text = clean_text.replace('|', ' ')  # remove pipes

    codes = re.findall(r'\b\d{8,}\b', clean_text)
    unique_codes = sorted(set(codes))

    print(f"Extracted {len(unique_codes)} unique cost centre codes: {unique_codes}")

    # Separate 10-digit and 8-digit codes (10 digit takes priority)
    ten_digit_codes = [c for c in unique_codes if len(c) == 10]
    eight_digit_codes = [c for c in unique_codes if len(c) == 8]

    return {
        'ten_digit': ' // '.join(ten_digit_codes),
        'eight_digit': ' // '.join(eight_digit_codes)
    }


def extract_limits(page_text):
    clean_text = re.sub(r"\s+", " ", page_text).strip()
    boundaries = []
    for heading in FINANCIAL_HEADINGS:
        pattern = re.escape(heading)
        match = re.search(pattern, clean_text, flags=re.IGNORECASE)
        if match:
            boundaries.append((heading, match.start()))
        else:
            print(f"⚠️ Warning: Heading not found: {heading}")
            boundaries.append((heading, None))

    extracted = []
    for i, (heading, start_pos) in enumerate(boundaries):
        if start_pos is None:
            extracted.append("None")
            continue

        next_start = None
        for j in range(i + 1, len(boundaries)):
            if boundaries[j][1] is not None:
                next_start = boundaries[j][1]
                break

        chunk = clean_text[start_pos:next_start].strip() if next_start else clean_text[start_pos:].strip()
        chunk = re.sub(re.escape(heading), "", chunk, flags=re.IGNORECASE).strip()
        chunk = re.split(r"FINANCIAL DIRECTION [A-Z]:|OTHER FINANCIAL DELEGATIONS", chunk, flags=re.IGNORECASE)[
            0].strip()
        extracted.append(chunk if chunk else "None")

    print(f"Extracted financial limits: {extracted}")
    return extracted


def parse_page4(text):
    def extract_section(ref_label):
        pattern = rf"{ref_label}.*?Name:\s*(.*?)\s*Job Title:\s*(.*?)\s*Date:\s*([\d/-]+)"
        match = re.search(pattern, text, re.DOTALL | re.IGNORECASE)
        if match:
            return match.group(1).strip(), match.group(2).strip(), match.group(3).strip()
        else:
            return "", "", ""

    delegated_name, delegated_title, _ = extract_section("DELEGATED OFFICER")
    manager_name, manager_title, _ = extract_section("LINE MANAGER")

    budget_match = re.search(
        r"AUTHORISING BUDGET HOLDER NAME:\s*Name:\s*(.*?)\s*Job Title:\s*(.*?)\s*Date:\s*([\d/-]+)",
        text, re.DOTALL | re.IGNORECASE
    )
    if budget_match:
        budget_name = budget_match.group(1).strip()
        budget_title = budget_match.group(2).strip()
        signed_date = budget_match.group(3).strip()
    else:
        budget_name = budget_title = signed_date = ""

    print(f"Delegated Officer: {delegated_name}, {delegated_title}")
    print(f"Line Manager: {manager_name}, {manager_title}")
    print(f"Budget Holder: {budget_name}, {budget_title}, Signed Date: {signed_date}")

    return {
        "delegated_name": delegated_name,
        "delegated_title": delegated_title,
        "manager_name": manager_name,
        "manager_title": manager_title,
        "budget_name": budget_name,
        "budget_title": budget_title,
        "signed_date": signed_date,
    }


print("About to start find_first_empty_row")


def find_first_empty_row(sheet):
    row = 5
    while True:
        val = sheet.range(f"A{row}").value
        if val is None:
            return row
        row += 1


print("Finished find_first_empty_row")
print("About to start get_next_row_and_count")


def get_next_row_and_count(sheet):
    # Get last used row in column A (or any column that always has data)
    last_used_row = sheet.range('A' + str(sheet.cells.last_cell.row)).end('up').row
    prev_count = 0
    for row in range(last_used_row, 4, -1):  # from last used down to 5
        val = sheet.range(f"A{row}").value
        if val is not None:
            try:
                prev_count = int(val)
            except (ValueError, TypeError):
                prev_count = 0
            break
    next_row = last_used_row + 1
    new_count = prev_count + 1
    return next_row, new_count


print("Finished get_next_row_and_count")
print("About to start update_excel")


import difflib
import datetime
import ctypes

MB_OKCANCEL = 1
MB_TOPMOST = 0x00040000

def update_excel(data, limits, codes):
    print(f"Opening workbook: {WORKBOOK_PATH}")
    print(f"Target sheet: {TARGET_SHEET_NAME}")
    print("Codes extracted:", codes)

    app = xw.App(visible=True)
    app.display_alerts = False
    app.screen_updating = False

    # Close default new workbook if open
    if app.books.count > 0:
        for book in app.books:
            if book.name == "Book1":
                book.close()

    wb = None
    try:
        import os

        wb_name = os.path.basename(WORKBOOK_PATH)

        # Attach to already open workbook or open it
        try:
            wb = app.books[wb_name]
            print(f"Attached to already open workbook: {wb_name}")
        except Exception:
            print(f"Workbook not open, opening: {WORKBOOK_PATH}")
            wb = app.books.open(WORKBOOK_PATH)

        ws = wb.sheets[TARGET_SHEET_NAME]

        new_row, new_count = get_next_row_and_count(ws)
        next_row = find_first_empty_row(ws)

        names = data['delegated_name'].strip().split()
        first_name = " ".join(names[:-1]) if len(names) >= 2 else (names[0] if names else "")
        last_name = names[-1] if len(names) >= 2 else ""

        full_name = f"{last_name} {first_name}".strip()
        lower_full_name = full_name.lower()

        matched_row = None
        matched_type_value = None

        used_rows = ws.range(f"D5:D{next_row}").value
        if not isinstance(used_rows, list):
            used_rows = [used_rows]

        for idx, name in enumerate(used_rows):
            row_number = idx + 5
            if isinstance(name, str) and name.strip().lower() == lower_full_name:
                # MessageBox with topmost flag
                answer = ctypes.windll.user32.MessageBoxW(
                    0,
                    f"Match found: {name}\nUse this match?",
                    "Exact Match",
                    MB_OKCANCEL | MB_TOPMOST
                )
                if answer == 1:  # Yes clicked
                    status = str(ws.range(f"G{row_number}").value).strip().lower()
                    if status == "active":
                        ws.range(f"G{row_number}").value = "Inactive"
                        today_str = datetime.datetime.now().strftime("%d/%m/%Y")
                        ws.range(f"H{row_number}").value = f"Set inactive on {today_str} due to new delegation."
                        matched_row = row_number
                        matched_type_value = ws.range(f"E{row_number}").value
                        break
                    else:
                        continue

        # If no exact match accepted, try fuzzy match
        if matched_row is None:
            possible_names = [str(name) for name in used_rows if isinstance(name, str)]
            close_matches = difflib.get_close_matches(full_name, possible_names, n=1, cutoff=0.75)
            if close_matches:
                guessed = close_matches[0]
                for idx, name in enumerate(used_rows):
                    row_number = idx + 5
                    if name == guessed:
                        answer = ctypes.windll.user32.MessageBoxW(
                            0,
                            f"Possible match: {guessed}\nUse this match?",
                            "Fuzzy Match",
                            MB_OKCANCEL | MB_TOPMOST
                        )
                        if answer == 1:
                            status = str(ws.range(f"G{row_number}").value).strip().lower()
                            if status == "active":
                                ws.range(f"G{row_number}").value = "Inactive"
                                today_str = datetime.datetime.now().strftime("%d/%m/%Y")
                                ws.range(f"H{row_number}").value = f"Set inactive on {today_str} due to new delegation."
                                matched_row = row_number
                                matched_type_value = ws.range(f"E{row_number}").value
                                break

        # Write new data
        names = full_name.split()
        first_name = " ".join(names[:-1]) if len(names) >= 2 else (names[0] if names else "")
        last_name = names[-1] if len(names) >= 2 else ""

        ws.range(f"A{new_row}").value = new_count
        ws.range(f"B{next_row}").value = first_name
        ws.range(f"C{next_row}").value = last_name
        ws.range(f"D{next_row}").value = full_name

        ws.range(f"F{next_row}").value = "Delegated Officer"
        if matched_type_value:
            ws.range(f"E{next_row}").value = matched_type_value
        ws.range(f"G{next_row}").value = "Active"

        ws.range(f"I{next_row}").value = data['delegated_title']
        ws.range(f"T{next_row}").value = codes['ten_digit']
        ws.range(f"U{next_row}").value = codes['eight_digit']

        # Process 10-digit codes
        ten_codes = codes['ten_digit'].split(' // ')
        ten_codes = sorted([code for code in ten_codes if len(code) == 10])
        if ten_codes:
            smallest = ten_codes[0]
            largest = ten_codes[-1]

            ws.range(f"J{next_row}").value = smallest[:4]
            ws.range(f"L{next_row}").value = smallest[:6]
            ws.range(f"N{next_row}").value = smallest
            ws.range(f"R{next_row}").value = smallest
            ws.range(f"P{next_row}").value = largest
            ws.range(f"S{next_row}").value = largest

        # Financial Limits columns
        columns = ["AE", "AH", "AK", "AN", "AQ", "AT", "AW", "AZ", "BC", "BF", "BI", "BL"]
        for i, value in enumerate(limits[:12]):
            ws.range(f"{columns[i]}{next_row}").value = value

        # Line Manager & Budget Holder
        ws.range(f"BS{next_row}").value = data['manager_name']
        ws.range(f"BT{next_row}").value = data['manager_title']
        ws.range(f"BU{next_row}").value = data['budget_name']
        ws.range(f"BV{next_row}").value = data['budget_title']

        ws.range(f"BW{next_row}").value = data['signed_date']
        wb.save()
        wb.close()
        wb.app.quit()
    finally:
            print()

def main(pdf_path):
    print(f"Processing PDF: {pdf_path}")
    try:
        # Extract text from relevant pages
        page2_text = extract_page_text(pdf_path, 1)  # page 2 (0-based index 1)
        page3_text = extract_page_text(pdf_path, 2)  # page 3 (0-based index 2)
        page4_text = extract_page_text(pdf_path, 3)  # page 4 (0-based index 3)

        # Extract cost centre codes from page 2
        codes = extract_cost_centre_codes(page2_text)

        # Extract financial limits from page 3
        limits = extract_limits(page3_text)

        # Extract names, titles, dates from page 4
        data = parse_page4(page4_text)

        # Update Excel workbook
        update_excel(data, limits, codes)

        print("Processing complete.")
    except Exception as exc:
        print(f"An error occurred: {exc}")


if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python script.py <path_to_pdf>")
        sys.exit(1)
    pdf_file_path = sys.argv[1]
    main(pdf_file_path)
    print("✅ Script completed. This window will close in 5 seconds...")
    time.sleep(5)
    sys.exit()

