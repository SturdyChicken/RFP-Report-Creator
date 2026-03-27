import csv  # This library lets us read and reqrite to excel csv files
import os  # Used to check for files (for errors)
import sys  # Used to exit the program (for errors)
from datetime import date, datetime  # ... gives the current date and time ...
from dateutil.relativedelta import (
    relativedelta,
)  # Allows us to compare dates inside the library dataset
from dateutil import parser
from openpyxl import Workbook  # Allows us to write directly to an excel spreadsheet
from openpyxl.styles import (
    PatternFill,
    Font,
    Alignment,
)  # Allows us to more finly adjust the formatting of the excel spreadsheet
import requests  # Allows us to scrape web pages
from bs4 import (
    BeautifulSoup,
)  # Allows us to organise the scraped web data into a more easy to work with format
import re  # Let's us use more robust string searching (case sensitivity, etc)
from transformers import pipeline
import logging  # Lets us disable errors
from transformers import (
    logging as hf_logging,
)  # these last two are just for the next 3 lines of code and make the terminal look cleaner lol
import asyncio  # Allows us to open up a webpage automatically, so we can read addition loaded javascript data without backend api stuff
from playwright.async_api import (
    async_playwright,
)  # This allows async to open, close, and run in backgroun more automatically and smoothly

# specialized code to silence the "tied weights" warnings
hf_logging.set_verbosity_error()
logging.getLogger("transformers").setLevel(logging.ERROR)

# Define a global variable to hold the AI model in memory
_shared_classifier = None


async def async_princeton(url, filename="princeton_grants.html"):
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)

        # Create a context with a real user agent (helps avoid being blocked)
        context = await browser.new_context(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:140.0) Gecko/20100101 Firefox/140.0"
        )
        page = await context.new_page()

        print(f"Navigating to {url}...")

        # 'networkidle' ensures we wait for the background JSON data to arrive
        await page.goto(url, wait_until="networkidle")

        print("Waiting for data to populate...")
        try:
            await page.wait_for_selector(
                "#search-result-count", state="visible", timeout=30000
            )
            print("   Data detected!")
            import asyncio

            await asyncio.sleep(2)

        except Exception as e:
            print(
                f"   Warning: Timed out waiting for data. Saving whatever we have... Error: {e}"
            )

        content = await page.content()
        with open(filename, "w", encoding="utf-8") as f:
            f.write(content)

        print(f"HTML saved to '{filename}' in program folder. Closing browser.")
        await browser.close()
        return filename


def clean_and_deduplicate(rfp_library):
    cleaned_library = []
    seen_fingerprints = set()

    for entry in rfp_library:
        raw_deadline = str(entry.get("Deadline", "")).strip()

        try:
            dt_object = datetime.strptime(raw_deadline, "%m/%d/%Y")
            entry["Deadline"] = dt_object.strftime("%m/%d/%Y")
        except ValueError:
            entry["Deadline"] = raw_deadline

        # We use (Name, URL, Deadline) as the unique identity of each rant to check against duplicates
        fingerprint = (
            entry.get("Grant Name", "").strip().lower(),
            entry.get("Deadline", "").strip(),
        )

        if fingerprint in seen_fingerprints:
            continue

        seen_fingerprints.add(fingerprint)
        cleaned_library.append(entry)

    return cleaned_library


def get_keywords(text, top_k=3):

    global _shared_classifier

    possible_keywords = [
        "Climate & Environment",
        "Data Science",
        "Food & Agriculture",
        "Life Sciences/Health",
        "AI/ Machine Learning",
        "Neuroscience",
        "Physical Sciences",
        "Science/Technology",
        "Public Health",
        "Business & Economics",
        "Law",
        "Economics",
        "Journalism",
        "Justice",
        "Policy",
        "Social Sciences",
        "Arts and Humanities",
        "Museums & Libraries",
        "BIPOC",
        "Education",
        "Sustainability",
        "Public Safety",
        "Conservation",
    ]
    # Load the model weights (Only happens the first time this function runs)
    if _shared_classifier is None:
        # Silence the library's built-in progress bars and warnings
        hf_logging.set_verbosity_error()
        try:
            _shared_classifier = pipeline(
                "zero-shot-classification", model="valhalla/distilbart-mnli-12-1"
            )
        except Exception as e:
            print(f"\nModel load failed: {e}")
            return "Manual Review Needed"
    try:
        if not text or not isinstance(text, str):
            return "Manual Review Needed"
        result = _shared_classifier(text, possible_keywords, multi_label=True)

        sorted_keywords = sorted(
            zip(result["labels"], result["scores"]), key=lambda x: x[1], reverse=True
        )
        top_list = [label for label, score in sorted_keywords[:top_k]]
        return ", ".join(top_list)
    except Exception as e:
        print(f"\n[!] Analysis Failed on text: '{text[:30]}...' | Error: {e}")
        return "Manual Review Needed"


def UCB_csv_reader(file_name):
    ## ---------- FUNCTION OBJECTIVES ------------
    # This file will execute the following:
    # Read the contents of the provided CSV file
    # Create a dictionary to store the data
    # The following data will be exported in the library
    # 0. Granting Organization
    # 1. Grant Name
    # 2. Deadline
    # 3. Funding Amount
    # 4. Keywords
    # 5. URL

    if not os.path.exists(file_name):
        print(f"ERROR: The file {file_name} was not found")
        print(
            "Please ensure the file has that exact name, and is in the same folder as this script"
        )
        sys.exit()
    print(f"File found: '{file_name}'")
    print("Processing data now...")

    rfp_library = []

    try:
        # "with open" basically opens the file AND gaurentees it will be closed later
        # 'r' is for 'read mode', 'utf-8-sig' is standard for excel csv
        with open(file_name, mode="r", encoding="utf-8-sig") as UCB_csv:

            reader = csv.reader(UCB_csv)

            next(reader, None)

            for row in reader:
                if len(row) <= 8:
                    continue

                original_dealine = row[2]
                if not original_dealine.strip():
                    final_dealine = "rolling"
                else:
                    final_dealine = original_dealine

                data_entry = {
                    "Granting Organization": row[0],  # Column A
                    "Grant Name": row[1],  # Column B
                    "Deadline": final_dealine,  # C
                    "Funding Amount": row[3],  # Column D
                    "Keywords": row[6],  # Column G
                    "URL": row[8],  # Column I
                }

                rfp_library.append(data_entry)

            print(
                f"Successfully extracted {len(rfp_library)} RFPs from UCB csv provided"
            )

    except Exception as error:
        print(f"An unexpected error occurred while reading the file: {error}")

    return rfp_library


def parse_date(date_str):
    ## ---------- FUNCTION OBJECTIVES ------------
    # Attempts to turn a string like '1/15/2026' into a date object
    # Returns "NONE" if false
    if not date_str or date_str.lower() == "rolling":
        return None

    formats = ["%m/%d/%Y", "%Y-%m-%d", "%m-%d-%Y"]

    for fmt in formats:
        try:
            return datetime.strptime(date_str, fmt).date()
        except ValueError:
            continue
    return None


def report_maker(rfp_library, output_folder):
    ## ---------- FUNCTION OBJECTIVES ------------
    # This file will create a formatted excel rfp report as a csv file with the following specifications
    # ouput into the given folder with the file name informed by current date and time
    # formatted in a readable and professional manor using openpyxl library
    # The following columns:
    # 1. Deadline
    # 2. Funding Organization Name
    # 3. <=HYPERLINK(I#,J#)>
    # 4. Funding Amount
    # 5. Key Words (Description / Abstract?)
    # 6. Natural Science?
    # 7. Social Science?
    # 8. Humanities?
    # 9. Direct URL
    # 10. Grant Name

    ## -------------------- SET UP FILE NAME AND PATH --------------------

    def get_sort_key(entry):
        date_str = str(entry.get("Deadline", ""))

        try:
            date_obj = datetime.strptime(date_str, "%m/%d/%Y")
            return (0, date_obj, entry["Granting Organization"])

        except (ValueError, TypeError):
            return (1, datetime.max, entry["Granting Organization"])

    rfp_library.sort(key=get_sort_key)
    today = date.today()
    start_date = (today + relativedelta(months=1)).replace(day=1)
    end_date = start_date + relativedelta(months=6)
    file_name = f"{start_date.strftime('%B_%Y')}_RFP_-_The_College.xlsx"
    
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
    full_path = os.path.join(output_folder, file_name)

    ## -------------------- "BUCKETS" FOR FORMATTING" --------------------

    fill_data_row = PatternFill(
        start_color="F2F2F2", end_color="F2F2F2", fill_type="solid"
    )  # CURRENT COLOR: LIGHT GREY
    fill_month_header = PatternFill(
        start_color="D9E1F2", end_color="D9E1F2", fill_type="solid"
    )  # CURRENT COLOR: LIGHT BLUE
    fill_main_header = PatternFill(
        start_color="8EA9DB", end_color="8EA9DB", fill_type="solid"
    )  # CURRENT COLOR: DAKRER BLUE
    fill_white = PatternFill(
        start_color="FFFFFF", end_color="FFFFFF", fill_type="solid"
    )  # CURRENT COLOR: PURE WHITE
    fill_ns = PatternFill(
        start_color="FCE4D6", end_color="FCE4D6", fill_type="solid"
    )  # CURRENT COLOR: LIGHT ORANGE
    fill_ss = PatternFill(
        start_color="DDEBF7", end_color="DDEBF7", fill_type="solid"
    )  # CURRENT COLOR: LIGHT BLUE
    fill_h = PatternFill(
        start_color="E2EFDA", end_color="E2EFDA", fill_type="solid"
    )  # CURRENT COLOR: LIGHT GREEN
    font_bold = Font(bold=True)
    font_link = Font(color="0000FF", underline="single")  # Blue & Underlined
    align_center = Alignment(horizontal="center", vertical="center")

    align_general = Alignment(vertical="center", wrap_text=False)

    # Now we need to sort the data into "buckets", one for each month + one for rolling
    buckets = {}
    current_month_cursor = start_date
    while current_month_cursor < end_date:
        buckets[current_month_cursor] = []
        current_month_cursor += relativedelta(months=1)
    buckets["Rolling"] = []

    for rfp in rfp_library:
        original_deadline = rfp["Deadline"]

        if original_deadline == "rolling":
            buckets["Rolling"].append(rfp)
        else:
            parsed_date = parse_date(original_deadline)
            if parsed_date and start_date <= parsed_date < end_date:
                bucket_key = parsed_date.replace(day=1)
                if bucket_key in buckets:
                    buckets[bucket_key].append(rfp)
                else:
                    pass

    ## -------------------- SET UP WORKBOOK AND HEADERS --------------------

    wb = Workbook()
    ws = wb.active
    ws.title = "RFP Report"
    ws.sheet_view.showGridLines = False

    # Approx conversion: Width = (Pixels - 5) / 7
    ws.column_dimensions["A"].width = 11  # 75px
    ws.column_dimensions["B"].width = 60  # 420px
    ws.column_dimensions["C"].width = 82  # 575px
    ws.column_dimensions["D"].width = 28  # 200px
    ws.column_dimensions["E"].width = 42  # 300px
    ws.column_dimensions["F"].width = 17  # 110px
    ws.column_dimensions["G"].width = 15  # 100px
    ws.column_dimensions["H"].width = 12  # 90px
    ws.column_dimensions["I"].width = 11  # 75px
    ws.column_dimensions["J"].width = 14  # 100px

    headers = [
        "Deadline",  # A
        "Granting Organization",  # B
        "Link",  # C
        "Funding Amount",  # D
        "Keywords",  # E
        "Discipline: NS",  # F
        "Discipline: SS",  # G
        "Discipline: H",  # H
        "URL",  # I
        "Grant Name",  # J
    ]

    ws.append(headers)
    ws.row_dimensions[1].height = 25  # Set Header Row Height

    for cell in ws[1]:
        # If it is Column K (11), make it white/invisible. Otherwise, make it Blue.
        if cell.column == 11:
            cell.fill = fill_white
        else:
            cell.fill = fill_main_header
            cell.font = font_bold
            cell.alignment = align_center

    ## -------------------- WRITE BUCKETS TO EXCEL --------------------

    excel_row_index = 2

    def write_rfp_list(rfp_list, block_title):
        nonlocal excel_row_index

        ws.merge_cells(
            start_row=excel_row_index,
            start_column=1,
            end_row=excel_row_index,
            end_column=10,
        )

        title_cell = ws.cell(row=excel_row_index, column=1, value=block_title)
        title_cell.fill = fill_month_header
        title_cell.font = font_bold
        title_cell.alignment = align_center
        ws.row_dimensions[excel_row_index].height = 25
        
        excel_row_index += 1

        for rfp in rfp_list:
            hyperlink_formula = f"=HYPERLINK(I{excel_row_index}, J{excel_row_index})"

            val_ns, fill_ns_cell = "", fill_data_row
            val_ss, fill_ss_cell = "", fill_data_row
            val_h, fill_h_cell = "", fill_data_row

            if rfp.get("Disc_NS", "").strip().upper() != "NONE":
                val_ns = rfp.get("Disc_NS", "Natural Sciences")
                fill_ns_cell = fill_ns  # Orange

            if rfp.get("Disc_SS", "").strip().upper() != "NONE":
                val_ss = rfp.get("Disc_SS", "Social Sciences")
                fill_ss_cell = fill_ss  # Blue
                
            if rfp.get("Disc_H", "").strip().upper() != "NONE":
                val_h = rfp.get("Disc_H", "Humanities")
                fill_h_cell = fill_h  # Green

            row_data = [
                rfp["Deadline"],
                rfp["Granting Organization"],
                hyperlink_formula,
                rfp["Funding Amount"],
                rfp["Keywords"],
                val_ns,
                val_ss,
                val_h,
                rfp["URL"],
                rfp["Grant Name"],
                " ",
            ]

            for col_num, value in enumerate(row_data, start=1):
                # This line is to make sure we increment our counter correctly to keep track of where we are in the worksheet
                cell = ws.cell(row=excel_row_index, column=col_num, value=value)
                cell.alignment = align_general

                if col_num == 11:
                    cell.fill = fill_white
                else:
                    cell.fill = fill_data_row

                if col_num == 6:  # NS
                    cell.fill = fill_ns_cell
                elif col_num == 7:  # SS
                    cell.fill = fill_ss_cell
                elif col_num == 8:  # H
                    cell.fill = fill_h_cell

                if col_num == 3:
                    cell.font = font_link

            ws.row_dimensions[excel_row_index].height = 25
            excel_row_index += 1

        excel_row_index += 2

    sorted_keys = sorted([k for k in buckets.keys() if isinstance(k, date)])

    for key in sorted_keys:
        rfps_in_month = buckets[key]
        if rfps_in_month:
            month_title = key.strftime("%B %Y")
            write_rfp_list(rfps_in_month, month_title)

    if buckets["Rolling"]:
        write_rfp_list(buckets["Rolling"], "Rolling Deadlines")

    ## -------------------- COMPLETE - PRINT INFORMATION --------------------
    try:
        wb.save(full_path)
        print(f"Success! Excel Report generated at: {full_path}")
        print(
            f"Filter applied: Only dates between {start_date} and {end_date} (plus 'rolling')."
        )
    except PermissionError:
        print(f"Error: Could not save to '{full_path}'. Check if the file is open")
    except Exception as error:
        print(f"An unexpected error occurred: {error}")


def add_discipline_data_from_keywords(rfp_library):
    ## Scans the keywords field of the library, and appends with discipline data

    # if it includes one of these, it is definitly [discipline].
    ns_keywords = {
        "Climate & Environment",
        "Data Science",
        "Food & Agriculture",
        "Life Sciences/Health",
        "AI/ Machine Learning",
        "Neuroscience",
        "Physical Sciences",
        "Science/Technology",
        "Public Health",
    }

    ss_keywords = {
        "Business & Economics",
        "Law",
        "Economics",
        "Journalism",
        "Justice",
        "Policy",
        "Social Sciences",
    }

    h_keywords = {"Arts and Humanities", "Museums & Libraries"}

    for rfp in rfp_library:
        raw_keywords = rfp.get("Keywords", "")

        current_tags = [tag.strip() for tag in raw_keywords.split(",")]

        is_ns = False
        is_ss = False
        is_h = False
        for tag in current_tags:
            if tag in ns_keywords:
                is_ns = True
            if tag in ss_keywords:
                is_ss = True
            if tag in h_keywords:
                is_h = True

        if not (is_ns or is_ss or is_h):
            is_ns = True
            is_ss = True
            is_h = True

        rfp["Disc_NS"] = "Natural Sciences" if is_ns else "NONE"
        rfp["Disc_SS"] = "Social Sciences" if is_ss else "NONE"
        rfp["Disc_H"] = "Humanities" if is_h else "NONE"

    print("Disciplines added based on provided keywords.")
    return rfp_library


def web_scraper_stanford(url):
    ## ---------- FUNCTION OBJECTIVES ------------
    # This function will scrape the website provided by "url" and return the following information:
    # Granting Organization
    # Grant Name
    # Funding Amount (May be N/A or Varies; if this is the case always return "Varies")
    # Deadline
    # Link (the hyperlink attached the text "guidelines")
    # (TBD): Keyword analysis from abstract

    ## -------------------- PHASE 1: SCRAPE HTML TEXT AND DEFINE BLOCKS --------------------

    print(f"fetching data from {url}")

    rfp_library = []

    try:
        # this will let the website know we are a person from a real computer, not a bot
        # this "user-agent" is accosiated with Blake Allen and his computer in ARM 172L.
        # if you are having problems with this code, try changing the user-agent to your own computer's
        # you can find this by googling "what is my user agent", and google will return it for you directly
        user_agent_blake = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:140.0) Gecko/20100101 Firefox/140.0"
        }

        response = requests.get(url, headers=user_agent_blake, timeout=15)

        if response.status_code != 200:
            print(f"Error: Failed to load page (Status Code: {response.status_code})")
            return []

    except Exception as e:
        print(f"Connection Error: {e}")
        return []

    current_offers_only = response.text.split("Deadline passed - for reference only")[0]

    soup = BeautifulSoup(current_offers_only, "html.parser")

    main_content = soup.find("main") or soup.body

    if not main_content:
        print("Error: Could not find main content area")
        return []

    org_names = [h for h in main_content.find_all("h3") if h.get_text(strip=True)]

    grant_blocks = []

    for header in org_names:
        block_content = []

        current_element = header.find_next_sibling()

        while current_element:
            if current_element.name == "h3":
                break

            block_content.append(current_element)
            current_element = current_element.find_next_sibling()
        grant_blocks.append(
            {
                "organization_name": header.get_text(strip=True),
                "elements": block_content,
            }
        )

    ## -------------------- PHASE 2: FIND OTHER INFO IN EACH BLOCK --------------------

    total_grants = len(grant_blocks)

    for i, block in enumerate(grant_blocks, 1):
        print(f"Completed {i}/{total_grants} AI Keyword Analysis...", end="\r")

        entry = {
            "Granting Organization": block["organization_name"],
            "Grant Name": "General Research Grant",
            "Funding Amount": "Varies",
            "Deadline": "Rolling",
            "URL": "",
            "Keywords": "",
        }

        # ------ GRANT NAME ------
        # Find the FIRST <p> tag that contains a <strong> tag
        for element in block["elements"]:
            if element.name == "p":
                bold_tag = element.find("strong")
                if bold_tag:
                    entry["Grant Name"] = bold_tag.get_text(strip=True)
                    break
        # ------ /GRANT NAME ------

        # ------ DEADLINE ------
        for element in block["elements"]:
            text = element.get_text(separator=" ", strip=True)

            if "Deadline:" in text:
                match = re.search(
                    r"Deadline:\s*([A-Za-z]+\s+\d{1,2},\s+\d{4})", text, re.IGNORECASE
                )

                if match:
                    raw_date_str = match.group(1).strip()

                    # Try to convert "January 7, 2026" -> "01/07/2026"
                    try:
                        dt_object = datetime.strptime(raw_date_str, "%B %d, %Y")
                        entry["Deadline"] = dt_object.strftime("%m/%d/%Y")
                    except ValueError:
                        entry["Deadline"] = "Rolling"

                break
        # ------ /DEADLINE ------

        # ------ FUNDING AMOUNT ------
        for element in block["elements"]:
            text = element.get_text(separator=" ", strip=True)

            if "Funding amount:" in text:
                parts = re.split(r"Funding amount:", text, flags=re.IGNORECASE)

                if len(parts) > 1:
                    full_amount_text = parts[1].strip()
                    
                    if full_amount_text.lower() in ["n/a", "varies"]:
                        entry["Funding Amount"] = "Varies"
                    else:
                        entry["Funding Amount"] = full_amount_text

                break
        # ------ /FUNDING AMOUNT ------

        # ------ URL ------
        # in case we don't find anything, we want the program to keep running
        backup_link = ""

        for element in block["elements"]:
            if not hasattr(element, "find_all"):
                continue

            links = element.find_all("a", href=True)

            for link in links:
                link_text = link.get_text(strip=True).lower()
                href = link["href"]

                if not backup_link:
                    backup_link = href

                if "guidelines" in link_text:
                    entry["URL"] = href
                    break

            if entry["URL"]:
                break
        if not entry["URL"]:
            entry["URL"] = backup_link
        # ------ /URL ------

        # ------ KEYWORDS ------
        description_parts = []

        description_parts.append(entry["Grant Name"])

        for element in block["elements"]:
            text = element.get_text(separator=" ", strip=True)

            if "Funding amount:" in text:
                continue
            if "Deadline:" in text:
                continue
            if "guidelines" in text.lower() and len(text) < 100:
                continue

            description_parts.append(text)

        clean_text = " ".join(description_parts)

        entry["Keywords"] = get_keywords(clean_text)
        # ------ /KEYWORDS ------

        rfp_library.append(entry)

    # TEMP Print results to check
    # for item in rfp_library:
    # print(f"Org: {item['Granting Organization']}")
    # print(f"Grant: {item['Grant Name']}")
    # print(f"Funding Amount: {item['Funding Amount']}")
    # print(f"Deadline: {item['Deadline']}")
    # print(f"URL: {item['URL']}")
    # print(f"Keywords: {item['Keywords']}")
    # print("-" * 20)

    ## ------------ FINISHED ------------

    ## Now that we have iterated through them all, we can export the final list to be used in the report generator
    print(f"Extracted {len(rfp_library)} grants from {url}!")
    return rfp_library


def analyze_local_html(filename):
    timezones = {
        "ET": 0,
        "EST": 0,
        "EDT": 0,
        "CT": 0,
        "CST": 0,
        "CDT": 0,
        "PT": 0,
        "PST": 0,
        "PDT": 0,
    }
    print(f"Opening '{filename}' for analysis...")

    ## -------------------- PHASE 1: SCRAPE HTML TEXT AND DEFINE BLOCKS --------------------

    with open(filename, "r", encoding="utf-8") as f:
        html_content = f.read()

    soup = BeautifulSoup(html_content, "html.parser")
    rfp_library = []
    
    # This effectively finds the <div class="cards-title-row">...</div> block
    start_marker = soup.find("h2", id="search-result-count").parent
    end_marker = soup.find("div", id="paginationContainer")

    main_content = []
    for sibling in start_marker.find_next_siblings():
        if sibling == end_marker:
            break
        main_content.append(sibling)

    # Finds all h3 headers with the specific class "card-title", wihch on this site is all grant names
    all_h3_headers = soup.find_all("h3", class_="card-title")

    orgs_and_names = []

    for header in all_h3_headers:
        full_text = header.get_text(strip=True)

        parts = full_text.split(":", 1)

        if len(parts) == 2:
            org_name = parts[0].strip()
            grant_name = parts[1].strip()
        else:
            org_name = full_text
            grant_name = full_text

        orgs_and_names.append({"Organization": org_name, "Grant": grant_name})

    rfp_library = []
    i = 0

    for header in all_h3_headers:
        i = i + 1
        print(f"Completed {i}/{len(all_h3_headers)} AI Keyword Analysis...", end="\r")

        # ------ GRANT ORG + NAME ------
        full_title = header.get_text(strip=True)
        parts = full_title.split(":", 1)
        org_name = parts[0].strip() if len(parts) > 1 else full_title
        grant_name = parts[1].strip() if len(parts) > 1 else full_title
        # ------ /GRANT ORG + NAME ------

        card_body = header.parent

        # ----- KEYWORDS ------
        description_tag = card_body.find("p", class_="card-text")
        description_text = (
            description_tag.get_text(strip=True)
            if description_tag
            else "no desciption found"
        )
        # Call an AI model to get keywords
        keywords = get_keywords(description_text)
        # ----- /KEYWORDS ------

        # ------ FUNDING & DEADLINE ------
        funding_amt = "N/A"
        deadline = "rolling"

        all_keys = card_body.find_all("div", class_="category-key")

        for key_element in all_keys:
            key_text = key_element.get_text(strip=True)
            
            if "Funding Amt" in key_text:
                val_element = key_element.find_next_sibling(
                    "div", class_="category-value"
                )
                if val_element:
                    funding_amt = val_element.get_text(strip=True)

            elif "Due Date" in key_text:
                val_element = key_element.find_next_sibling(
                    "div", class_="category-value"
                )
                if val_element:
                    deadline = val_element.get_text(strip=True)

        if deadline and deadline not in ["rolling"]:
            try:
                dt = parser.parse(deadline, fuzzy=True, tzinfos=timezones)

                deadline = dt.strftime("%m/%d/%Y")
            except:
                pass
        # ------ /FUNDING & DEADLINE ------

        # ------ URL ------
        url = "N/A"
        card_footer = card_body.find_next_sibling("div", class_="card-footer")

        if card_footer:
            link_tag = card_footer.find(
                "a", string=lambda text: text and "Learn More" in text
            )
            if link_tag:
                raw_url = link_tag.get("href")
                if raw_url.startswith("#"):
                    url = "https://princeton.infoready4.com/" + raw_url
                else:
                    url = raw_url
        # ------ /URL ------

        rfp_library.append(
            {
                "Granting Organization": org_name,
                "Grant Name": grant_name,
                "Funding Amount": funding_amt,
                "Deadline": deadline,
                "URL": url,
                "Keywords": keywords,
            }
        )
    print(f"Found {len(rfp_library)} sources from Princeton!")
    return rfp_library


def web_scraper_princeton(url):
    html_text_file = asyncio.run(async_princeton(url))
    rfp_library = analyze_local_html(html_text_file)
    return rfp_library
