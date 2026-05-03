"""
extractor.py
------------
Uses Gemini API to read electricity bill (PDF or image)
and extract key fields + monthly history as a Python dict.
"""

import base64
import json
import re
import google.generativeai as genai

from modules.config import GEMINI_API_KEY

if not GEMINI_API_KEY or GEMINI_API_KEY == "paste-your-gemini-key-here":
    raise ValueError(
        "\n\n❌ Gemini API key not set!\n"
        "Open  modules/config.py  and paste your Gemini key.\n"
        "Get it free from: https://aistudio.google.com/apikey\n"
    )

genai.configure(api_key=GEMINI_API_KEY)
model = genai.GenerativeModel("gemini-2.5-flash")

EXTRACTION_PROMPT = """You are analyzing an Indian MSEDCL Maharashtra electricity bill.
Extract ALL fields below. Return ONLY a valid JSON object — no explanation, no markdown.

{
  "consumer_name": "",
  "consumer_number": "",
  "address": "",
  "meter_number": "",
  "billing_month": "",
  "units_consumed": 0,
  "sanctioned_load_kw": 0,
  "contract_demand_kva": 0,
  "tariff_category": "",
  "previous_reading": 0,
  "current_reading": 0,
  "total_bill_amount": 0,
  "energy_charges": 0,
  "fixed_charges": 130,
  "due_date": "",
  "bill_date": "",
  "voltage_supply": "",
  "connection_type": "Single Phase",
  "monthly_history": [
    {"month": "February 2025", "units": 99,  "bill_amount": 0, "unit_cost": 0},
    {"month": "March 2025",    "units": 151, "bill_amount": 0, "unit_cost": 0}
  ]
}

Rules:
- units_consumed: kWh used this billing period (most recent month)
- sanctioned_load_kw: connected load in kW (look for "मंजूर भार" or "Sanctioned Load")
- connection_type: "Single Phase" or "Three Phase"
- monthly_history: extract ALL months shown in the bill's consumption history graph/table
  (oldest first, newest last). For each month extract:
    - month: month name and year e.g. "January 2026"
    - units: units consumed that month (number)
    - bill_amount: bill amount for that month if shown (number, else 0)
    - unit_cost: cost per unit for that month if shown (number, else 0)
  Extract as many months as visible — typically 12-13 months.
- fixed_charges: look for fixed/demand charges amount
- Use 0 for missing numbers, "" for missing strings
- Return ONLY the JSON, nothing else"""


def extract_bill_data(file_bytes: bytes, file_type: str) -> dict:
    """Send bill to Gemini, get structured data including monthly history."""
    if "pdf" in file_type:
        mime = "application/pdf"
    elif "png" in file_type:
        mime = "image/png"
    else:
        mime = "image/jpeg"

    image_part = {
        "inline_data": {
            "mime_type": mime,
            "data": base64.b64encode(file_bytes).decode("utf-8")
        }
    }

    response = model.generate_content([image_part, EXTRACTION_PROMPT])
    raw = response.text.strip()
    raw = re.sub(r"^```json\s*", "", raw)
    raw = re.sub(r"\s*```$", "", raw)

    return json.loads(raw)