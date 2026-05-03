"""
routes.py — handles 1 or 2 bill uploads per request.
"""
import os
from flask import Blueprint, render_template, request, send_file
from modules.extractor    import extract_bill_data
from modules.calculator   import calculate_solar
from modules.excel_writer import build_excel, build_combined_excel

bp = Blueprint("main", __name__)
OUTPUT_DIR = "outputs"


def is_valid_bill(bill):
    units  = float(bill.get("units_consumed",    0) or 0)
    amount = float(bill.get("total_bill_amount", 0) or 0)
    name   = str(bill.get("consumer_name",   "") or "").strip()
    number = str(bill.get("consumer_number", "") or "").strip()
    if units == 0 and amount == 0 and not name and not number:
        return False, "⚠️ This does not look like an electricity bill."
    if units == 0:
        return False, "⚠️ Could not find 'Units Consumed'. Please upload a clear bill image."
    return True, None


@bp.route("/", methods=["GET"])
def index():
    return render_template("index.html")


@bp.route("/process", methods=["POST"])
def process():
    file1 = request.files.get("bill")
    file2 = request.files.get("bill2")

    if not file1 or not file1.filename:
        return render_template("index.html", error="Please select at least one bill file.")

    allowed = {"pdf","jpg","jpeg","png"}
    def bad(f): return f.filename.rsplit(".",1)[-1].lower() not in allowed

    if bad(file1):
        return render_template("index.html", error="Only PDF, JPG, PNG supported.")
    if file2 and file2.filename and bad(file2):
        return render_template("index.html", error="Only PDF, JPG, PNG supported for Bill 2.")

    try:
        bill1 = extract_bill_data(file1.read(), file1.content_type or "application/pdf")
        ok, msg = is_valid_bill(bill1)
        if not ok: return render_template("index.html", error=msg)
        solar1 = calculate_solar(bill1)

        bill2 = solar2 = None
        if file2 and file2.filename:
            bill2 = extract_bill_data(file2.read(), file2.content_type or "application/pdf")
            ok2, msg2 = is_valid_bill(bill2)
            if not ok2: return render_template("index.html", error=f"Bill 2: {msg2}")
            solar2 = calculate_solar(bill2)

        _, ind_file  = build_excel(bill1, solar1, OUTPUT_DIR, bill2, solar2)
        _, mst_file  = build_combined_excel(bill1, solar1, OUTPUT_DIR, bill2, solar2)

        return render_template("results.html",
            bill=bill1, solar=solar1, bill2=bill2,
            ind_filename=ind_file, combined_filename=mst_file)

    except Exception as e:
        return render_template("index.html", error=f"Error: {str(e)}")


@bp.route("/download/<filename>")
def download(filename):
    path = os.path.join(OUTPUT_DIR, os.path.basename(filename))
    if not os.path.exists(path):
        return render_template("index.html", error="File not found. Please process again.")
    return send_file(path, as_attachment=True)