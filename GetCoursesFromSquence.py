import requests
from bs4 import BeautifulSoup
import re
from collections import Counter
import pandas as pd

def scrape_program_page(url, year_range="2025-2026"):
    """
    Scrape a program sequence page for course codes.
    Returns a dict of {caption_text: DataFrame} where each DataFrame has Category, Count, %.
    """
    response = requests.get(url)
    soup = BeautifulSoup(response.text, "html.parser")

    tables = {}

    for caption in soup.find_all("caption"):
        caption_text = caption.get_text(strip=True)
        if "Course sequence" in caption_text and year_range in caption_text:
            target_table = caption.find_parent("table")
            codes = []

            for li in target_table.find_all("li"):
                strong = li.find("strong")
                if strong:
                    text = strong.get_text(strip=True)
                    code = text.replace(" ", "")
                    codes.append(code)

            categories = Counter()
            for code in codes:
                match = re.match(r"([A-Z]+)", code)
                if match:
                    prefix = match.group(1)
                    categories[prefix] += 1

            total_courses = len(codes)
            rows = []
            for category, count in categories.items():
                pct = round((count / total_courses) * 100, 2)
                rows.append({
                    "Category": category,
                    "Count": count,
                    "%": f"{pct}%"
                })

            df = pd.DataFrame(rows)
            tables[caption_text] = df

    return tables

def main():
    year_range = "2025-2026"

    # Ask user for one or more URLs
    print("Enter program URLs (comma separated if multiple):")
    urls_input = input().strip()
    program_urls = [u.strip() for u in urls_input.split(",") if u.strip()]

    with pd.ExcelWriter("program_reports.xlsx", engine="openpyxl") as writer:
        for url in program_urls:
            print(f"\nScraping from {url}")
            tables = scrape_program_page(url, year_range)

            if tables:
                # Use domain or last part of URL as sheet name
                program_name = url.split("/")[-1] or "Program"
                sheet_name = program_name[:31]

                start_row = 1
                pd.DataFrame().to_excel(writer, sheet_name=sheet_name, index=False)
                ws = writer.sheets[sheet_name]

                ws.cell(row=start_row, column=1, value=f"{program_name} – {year_range}")
                start_row += 2

                for caption, df in tables.items():
                    ws.cell(row=start_row, column=1, value=caption)
                    df.to_excel(writer, sheet_name=sheet_name, index=False, startrow=start_row)
                    start_row += (1 + 1 + len(df)) + 2

                print(f"Added summaries for {program_name} to Excel.")
            else:
                print("No data found.")

    print("\nSaved all program reports to program_reports.xlsx")

if __name__ == "__main__":
    main()