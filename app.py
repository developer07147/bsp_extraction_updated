import os
import re
import io
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

os.makedirs(app.config["UPLOAD_FOLDER"], exist_ok=True)
os.makedirs(app.config["OUTPUT_FOLDER"], exist_ok=True)

EXCEL_FILE = "BSP_Extracted_Data.xlsx"


# ------------------ Helper Functions ------------------
def allowed_file(filename):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in app.config["ALLOWED_EXTENSIONS"]


def normalize_amount(value):
    """Convert string like '20,565.45' to float 20565.45"""
    if not value:
        return None
    value = value.replace(",", "").replace(" ", "")
    try:
        return float(value)
    except:
        return None


def extract_data_from_text(text):
    """Extract IATA numbers and corresponding Grand Totals from text"""
    iata_pattern = re.compile(r"\b\d{2}-\d+\s+\d{3,4}\s+\d\b")
    grand_pattern = re.compile(r"grand total[^\n\r]*?([-+]?\d{1,3}(?:[,.\d{3}]*\d)?(?:\.\d+)?)", re.IGNORECASE)

    results = []
    for match in iata_pattern.finditer(text):
        iata = match.group()
        snippet = text[match.start(): match.start() + 1500]
        grand = grand_pattern.search(snippet)
        if grand:
            amount = normalize_amount(grand.group(1))
            results.append((iata, amount))

    # fallback: if none found, look for a global grand total
    if not results:
        global_match = grand_pattern.search(text)
        if global_match:
            results.append(("GLOBAL TOTAL", normalize_amount(global_match.group(1))))
    return results


def extract_from_pdf(pdf_bytes):
    """Extract text from PDF and parse data"""
    text = ""
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            page_text = page.extract_text(x_tolerance=2, y_tolerance=2)
            if page_text:
                text += page_text + "\n"
                print("TEXT LENGTH:",len(text))
                print(text[:5000])
    return extract_data_from_text(text)


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
    writer = pd.ExcelWriter(output_path, engine="openpyxl")

    processed = 0
    for file in files:
        if allowed_file(file.filename):
            filename = secure_filename(file.filename)
            pdf_bytes = file.read()
            try:
                extracted_data = extract_from_pdf(pdf_bytes)
                df = pd.DataFrame(extracted_data, columns=["IATA", "GRAND_TOTAL"])
                if df.empty:
                    df = pd.DataFrame(columns=["IATA", "GRAND_TOTAL"])
                sheet_name = os.path.splitext(filename)[0][:31]
                df.to_excel(writer, sheet_name=sheet_name, index=False)
                processed += 1
            except Exception as e:
                flash(f"Error processing {filename}: {e}", "danger")

    writer.close()

    if processed > 0:
        flash(f"Processed {processed} file(s) successfully!", "success")
        return redirect(url_for("download_file"))
    else:
        flash("No valid PDFs processed.", "warning")
        return redirect(url_for("index"))


@app.route("/download")
def download_file():
    output_path = os.path.join(app.config["OUTPUT_FOLDER"], EXCEL_FILE)
    if not os.path.exists(output_path):
        flash("No Excel file found. Please process PDFs first.", "warning")
        return redirect(url_for("index"))
    return send_file(output_path, as_attachment=True, download_name=EXCEL_FILE)


if __name__ == "__main__":
    app.run(debug=True)
