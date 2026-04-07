import openpyxl
import argparse
import logging
import re
import os

from req_loader import load_requirements, add_filter_args, parse_filters

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)


def extract_reqs_from_lld_file(file_path):
    """
    Reads the LLD Excel file and extracts mapped REQs.
    Looks for a column named 'REQ'. Handles multiple reqs per cell.
    """
    logging.info(f"Loading LLD file: {file_path}")
    wb = openpyxl.load_workbook(file_path, data_only=True)
    sheet = wb.active

    lld_reqs = set()
    req_col_idx = None

    for col_idx, cell in enumerate(sheet[1], start=1):
        if cell.value and str(cell.value).strip().upper() == 'REQ':
            req_col_idx = col_idx
            break

    if not req_col_idx:
        logging.error("Could not find a 'REQ' column in the LLD file.")
        return lld_reqs

    logging.info(f"Found REQ column at index {req_col_idx}. Extracting requirements...")

    for row in sheet.iter_rows(min_row=2, min_col=req_col_idx, max_col=req_col_idx):
        val = row[0].value
        if val:
            split_reqs = re.split(r'[\s,]+', str(val).strip())
            for req in split_reqs:
                if req:
                    lld_reqs.add(req)

    logging.info(f"Extracted {len(lld_reqs)} unique requirements mapped in LLD file.")
    return lld_reqs


def generate_html_report(req_ids, lld_reqs, output_file="traceability_report.html"):
    """
    Generates a modern, interactive HTML report showing the traceability gaps and coverages.
    """
    logging.info("Analyzing traceability coverages and gaps...")

    covered_in_lld = req_ids.intersection(lld_reqs)
    missing_in_lld = req_ids - lld_reqs
    missing_in_reqs = lld_reqs - req_ids

    html_content = f"""
    <!DOCTYPE html>
    <html lang="en">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Traceability Matrix Report</title>
        <style>
            body {{ font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background-color: #f4f7f6; margin: 0; padding: 20px; color: #333; }}
            .container {{ max-width: 900px; margin: auto; background: #fff; padding: 30px; border-radius: 8px; box-shadow: 0 4px 8px rgba(0,0,0,0.1); }}
            h1 {{ text-align: center; color: #2c3e50; }}
            .summary {{ display: flex; justify-content: space-around; margin: 20px 0; }}
            .card {{ background: #ecf0f1; padding: 20px; border-radius: 8px; text-align: center; width: 30%; box-shadow: 0 2px 4px rgba(0,0,0,0.05); cursor: pointer; transition: 0.3s; }}
            .card:hover {{ background: #bdc3c7; }}
            .card h3 {{ margin: 0; font-size: 24px; color: #2980b9; }}
            .card p {{ margin: 5px 0 0; color: #7f8c8d; font-weight: bold; }}

            /* Specific Colors */
            .card.success h3 {{ color: #27ae60; }}
            .card.warning h3 {{ color: #e67e22; }}
            .card.danger h3 {{ color: #e74c3c; }}

            .section {{ display: none; margin-top: 30px; }}
            .section.active {{ display: block; }}
            h2 {{ color: #34495e; border-bottom: 2px solid #ecf0f1; padding-bottom: 5px; }}
            ul {{ list-style-type: none; padding: 0; max-height: 400px; overflow-y: auto; background: #fafafa; border: 1px solid #ddd; border-radius: 4px; }}
            li {{ padding: 10px; border-bottom: 1px solid #eee; }}
            li:last-child {{ border-bottom: none; }}
        </style>
        <script>
            function toggleSection(sectionId) {{
                document.querySelectorAll('.section').forEach(sec => sec.classList.remove('active'));
                document.getElementById(sectionId).classList.add('active');
            }}
        </script>
    </head>
    <body>
        <div class="container">
            <h1>Bidirectional Traceability Report</h1>

            <div class="summary">
                <div class="card success" onclick="toggleSection('covered')">
                    <h3>{len(covered_in_lld)}</h3>
                    <p>Covered Requirements</p>
                </div>
                <div class="card danger" onclick="toggleSection('missing-lld')">
                    <h3>{len(missing_in_lld)}</h3>
                    <p>Missing in LLD</p>
                </div>
                <div class="card warning" onclick="toggleSection('missing-reqs')">
                    <h3>{len(missing_in_reqs)}</h3>
                    <p>Missing in Reqs File</p>
                </div>
            </div>

            <div id="covered" class="section active">
                <h2>Covered Requirements (Exist in Both)</h2>
                <ul>
                    {''.join(f'<li>{req}</li>' for req in sorted(covered_in_lld))}
                </ul>
            </div>

            <div id="missing-lld" class="section">
                <h2>Missing in LLD (Defined in Reqs, not implemented)</h2>
                <ul>
                    {''.join(f'<li>{req}</li>' for req in sorted(missing_in_lld))}
                </ul>
            </div>

            <div id="missing-reqs" class="section">
                <h2>Unmapped in LLD (Exist in LLD, not in Reqs file)</h2>
                <ul>
                    {''.join(f'<li>{req}</li>' for req in sorted(missing_in_reqs))}
                </ul>
            </div>
        </div>
    </body>
    </html>
    """

    with open(output_file, 'w', encoding='utf-8') as f:
        f.write(html_content)

    logging.info(f"Report successfully generated: {os.path.abspath(output_file)}")


def main():
    parser = argparse.ArgumentParser(description="Generate 2-way Traceability Matrix from Excel files.")
    parser.add_argument("req_file", help="Path to the Requirements Excel file")
    parser.add_argument("lld_file", help="Path to the LLD Excel file")
    parser.add_argument("-o", "--output", default="traceability_report.html", help="Output HTML report file name")
    add_filter_args(parser)

    args = parser.parse_args()

    if not os.path.exists(args.req_file):
        logging.error(f"Requirements file not found: {args.req_file}")
        return

    if not os.path.exists(args.lld_file):
        logging.error(f"LLD file not found: {args.lld_file}")
        return

    filters = parse_filters(args.filter)
    req_ids = load_requirements(args.req_file, filters=filters, filter_logic=args.filter_logic)
    lld_reqs = extract_reqs_from_lld_file(args.lld_file)

    if not req_ids and not lld_reqs:
        logging.error("No requirements found in both files. Exiting.")
        return

    generate_html_report(req_ids, lld_reqs, args.output)


if __name__ == "__main__":
    main()
