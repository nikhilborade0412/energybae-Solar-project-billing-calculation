"""
calculator.py
-------------
Uses Groq API (fast + free) to calculate solar system sizing
based on extracted bill data.

Groq receives the bill data as text and returns solar recommendations.
"""

import json
import re
from groq import Groq

from modules.config import GROQ_API_KEY

# ── Validate key ──────────────────────────────────────────────────────────────
if not GROQ_API_KEY or GROQ_API_KEY == "paste-your-groq-key-here":
    raise ValueError(
        "\n\n❌ Groq API key not set!\n"
        "Open  modules/config.py  and paste your Groq key.\n"
        "Get it free from: https://console.groq.com\n"
    )

client = Groq(api_key=GROQ_API_KEY)

# ── Prompt ────────────────────────────────────────────────────────────────────
SOLAR_PROMPT = """You are a solar energy expert for Maharashtra, India.
Given the electricity bill data below, calculate the solar system recommendation.
Return ONLY a valid JSON object — no explanation, no markdown.

Bill data:
{bill_data}

Return this exact JSON with calculated values:
{{
  "daily_units": 0.0,
  "required_kw": 0.0,
  "recommended_kw": 0.0,
  "num_panels": 0,
  "panel_watt": 400,
  "battery_kw": 0.0,
  "estimated_cost": 0,
  "annual_savings": 0,
  "payback_years": 0.0,
  "co2_saved_per_year": 0.0,
  "peak_sun_hours": 5,
  "system_efficiency": 0.8,
  "avg_tariff": 7.5
}}

Calculation rules (follow exactly):
- daily_units = units_consumed / 30
- required_kw = daily_units / (5 peak sun hours × 0.8 efficiency)
- recommended_kw = round required_kw UP to nearest 0.5 kW
- num_panels = ceil(recommended_kw × 1000 / 400)
- battery_kw = round(daily_units / 2, 1)
- estimated_cost = recommended_kw × 50000  (Rs per kW installed)
- annual_savings = units_consumed × 12 × 7.5  (avg Rs per unit)
- payback_years = estimated_cost / annual_savings  (round to 2 decimals)
- co2_saved_per_year = units_consumed × 12 × 0.82  (kg CO2 per kWh)
- Return ONLY the JSON, nothing else"""


# ── Main function ─────────────────────────────────────────────────────────────
def calculate_solar(bill_data: dict) -> dict:
    """
    Send extracted bill data to Groq and get solar sizing recommendation.
    """
    prompt = SOLAR_PROMPT.format(bill_data=json.dumps(bill_data, indent=2))

    response = client.chat.completions.create(
        model="llama-3.1-8b-instant",
        messages=[{"role": "user", "content": prompt}],
        max_tokens=500,
        temperature=0,                # deterministic — we want exact numbers
    )

    raw = response.choices[0].message.content.strip()

    # Clean markdown fences if present
    raw = re.sub(r"^```json\s*", "", raw)
    raw = re.sub(r"\s*```$", "", raw)

    return json.loads(raw)
