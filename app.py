import os
import re
import io
import gc
import time
import logging
import traceback
from flask import Flask, render_template, request, send_file, redirect, url_for, flash
import pdfplumber
import pandas as pd
from werkzeug.utils import secure_filename

# ------------------ Flask Setup ------------------
app = Flask(__name__)
app.secret_key = "bsp_secret_key"
app.config["UPLOAD_FOLDER"] = "uploads"
app.config["OUTPUT_FOLDER"] = "outputs"
app.config["ALLOWED_EXTENSIONS"] = {"pdf"}
app.config["MAX_CONTENT_LENGTH"] = 16 * 1024 * 1024  # 16 MB per uploaded PDF

os.makedirs(app.config["UPLOAD_FOLDER"], exist_ok=True)
os.makedirs(app.config["OUTPUT_FOLDER"], exist_ok=True)

EXCEL_FILE = "BSP_Extracted_Data.xlsx"

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(name)s: %(message)s")
logger = logging.getLogger(__name__)


# ------------------ Helper Functions ------------------
def allowed_file(filename):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in app.config["ALLOWED_EXTENSIONS"]


def normalize_amount(value):
    """Convert string like '20,565.45' to float 20565.45"""
    if value is None:
        return None

    text = str(value).strip()
    if not text:
        return None

    text = text.replace(",", "").replace(" ", "")
    if text.startswith("(") and text.endswith(")"):
        text = "-" + text[1:-1]

    try:
        return float(text)
    except (TypeError, ValueError):
        return None


def extract_data_from_page(page_text):
    """Extract IATA numbers and corresponding Grand Totals from a single page of text."""
    if not page_text:
        return []

    text = page_text.replace("\r\n", "\n")
    iata_pattern = re.compile(r"\b\d{2}-\d+\s+\d{3,4}\s+\d\b")
    grand_pattern = re.compile(
        r"grand total[^\n\r]*?([-+]?(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d+)?)",
        re.IGNORECASE,
    )

    results = []
    for match in iata_pattern.finditer(text):
        iata = match.group(0)
        snippet_start = max(0, match.start())
        snippet = text[snippet_start: snippet_start + 1500]
        grand_match = grand_pattern.search(snippet)
        if grand_match:
            amount = normalize_amount(grand_match.group(1))
            if amount is not None:
                results.append((iata, amount))

    if not results:
        global_match = grand_pattern.search(text)
        if global_match:
            amount = normalize_amount(global_match.group(1))
            if amount is not None:
                results.append(("GLOBAL TOTAL", amount))

    return results


def extract_from_pdf(pdf_bytes):
    """Extract text from PDF page by page and return extracted results without retaining full document text."""
    start_time = time.perf_counter()
    results = []

    try:
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            page_count = len(pdf.pages)
            logger.info("Starting PDF extraction: %s pages", page_count)

            for page_index, page in enumerate(pdf.pages, start=1):
                page_text = page.extract_text(x_tolerance=2, y_tolerance=2)
                if not page_text:
                    logger.debug("Page %s produced no extractable text", page_index)
                    continue

                page_results = extract_data_from_page(page_text)
                if page_results:
                    results.extend(page_results)

                del page_text
                del page_results
                gc.collect()

            logger.info(
                "Completed PDF extraction in %.2f seconds. Extracted %s rows.",
                time.perf_counter() - start_time,
                len(results),
            )
            return results
    except Exception:
        logger.exception("Failed during PDF extraction")
        raise


# ------------------ Flask Routes ------------------
@app.route("/")
def index():
    return render_template("index.html")


@app.route("/process", methods=["POST"])
def process():
    if "pdfs" not in request.files:
        flash("No files uploaded.", "danger")
        return redirect(url_for("index"))

    files = request.files.getlist("pdfs")
    if not files or files[0].filename == "":
        flash("Please select at least one PDF file.", "warning")
        return redirect(url_for("index"))

    output_path = os.path.join(app.config["OUTPUT_FOLDER"], EXCEL_FILE)
    processed = 0
    total_rows = 0

    try:
        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            for file in files:
                if not allowed_file(file.filename):
                    logger.warning("Skipping non-PDF file: %s", file.filename)
                    continue

                filename = secure_filename(file.filename)
                pdf_bytes = file.read()
                if not pdf_bytes:
                    logger.warning("Empty PDF received: %s", filename)
                    continue

                pdf_start_time = time.perf_counter()
                try:
                    extracted_data = extract_from_pdf(pdf_bytes)
                    df = pd.DataFrame(extracted_data, columns=["IATA", "GRAND_TOTAL"])
                    if df.empty:
                        df = pd.DataFrame(columns=["IATA", "GRAND_TOTAL"])

                    sheet_name = os.path.splitext(filename)[0][:31]
                    df.to_excel(writer, sheet_name=sheet_name, index=False)

                    rows_extracted = len(extracted_data)
                    total_rows += rows_extracted
                    processed += 1

                    logger.info(
                        "Processed %s in %.2f seconds. Rows extracted: %s",
                        filename,
                        time.perf_counter() - pdf_start_time,
                        rows_extracted,
                    )

                    del extracted_data
                    del df
                    gc.collect()
                except Exception:
                    logger.exception("Error processing file: %s", filename)
                    flash(f"Error processing {filename}: {traceback.format_exc()}", "danger")

        if processed > 0:
            flash(f"Processed {processed} file(s) successfully with {total_rows} extracted rows.", "success")
            return redirect(url_for("download_file"))

        flash("No valid PDFs processed.", "warning")
        return redirect(url_for("index"))
    except Exception:
        logger.exception("Failed while creating Excel output")
        flash("An internal error occurred while generating the Excel file.", "danger")
        return redirect(url_for("index"))


@app.route("/download")
def download_file():
    output_path = os.path.join(app.config["OUTPUT_FOLDER"], EXCEL_FILE)
    if not os.path.exists(output_path):
        flash("No Excel file found. Please process PDFs first.", "warning")
        return redirect(url_for("index"))
    return send_file(output_path, as_attachment=True, download_name=EXCEL_FILE)


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)), debug=False)
