# ⚡ Energybae Solar Load Calculator

Automatically reads an MSEDCL electricity bill (PDF or image) using AI,
calculates the recommended solar system size, and generates a filled Excel report.

---

## Which API does what

| Task | API | Cost |
|---|---|---|
| Read bill — PDF / Image | **Gemini** (gemini-2.5-flash) | Free |
| Solar calculation logic | **Groq** (llama-3.1-8b-instant) | Free |

---

## Project Structure

```
energybae_v4/
│
├── app.py                   ← Entry point — run this to start the server
├── routes.py                ← URL routes (/, /process, /download) + bill validation
├── requirements.txt         ← Python dependencies
├── .gitignore               ← Keeps API keys and outputs off GitHub
│
├── modules/
│   ├── config.py            ← ⭐ PASTE YOUR API KEYS HERE (only file to edit)
│   ├── extractor.py         ← Gemini reads bill image/PDF → returns Python dict
│   ├── calculator.py        ← Groq calculates solar sizing → returns Python dict
│   └── excel_writer.py      ← Builds Excel report matching Energybae template
│
├── templates/
│   ├── index.html           ← Upload form with loading overlay animation
│   └── results.html         ← Results page (Jinja2 fills values from Python)
│
├── static/
│   └── css/style.css        ← All styles
│
├── outputs/                 ← Generated Excel files saved here (auto-created)
└── uploads/                 ← Temp folder (auto-created)
```

---

## Excel Output Format

The generated Excel matches the Energybae template:
- Consumer Name, Number, Fixed Charges, Sanctioned Load, Connection Type
- Monthly units history table (up to 13 months extracted from bill)
- Average units, kW calculation, Solar Panels count
- Solar capacity (green) and Number of Panels (red) highlighted rows
- Total solar capacity and panels summary at bottom

---

## How to Run

### Step 1 — Get FREE API keys

**Gemini** (for reading bill images/PDFs):
1. Go to https://aistudio.google.com/apikey
2. Sign in with Google → Click "Create API Key"
3. Copy the key (starts with AIzaSy...)

**Groq** (for solar calculations):
1. Go to https://console.groq.com
2. Sign up → API Keys → Create API Key
3. Copy the key (starts with gsk_...)

### Step 2 — Paste keys in config.py

Open `modules/config.py` and fill in:
```python
GEMINI_API_KEY = "AIzaSy..."
GROQ_API_KEY   = "gsk_..."
```
Save the file (Ctrl+S).

### Step 3 — Install Python 3.9+
https://www.python.org/downloads/

### Step 4 — Create virtual environment (recommended)
```
python -m venv myenv
myenv\Scripts\Activate      # Windows
source myenv/bin/activate   # Mac/Linux
```

### Step 5 — Install dependencies
```
pip install -r requirements.txt
```

### Step 6 — Run
```
python app.py
```

### Step 7 — Open browser
```
http://localhost:5000
```

---

## Usage

1. Open http://localhost:5000
2. Click the upload area → select MSEDCL electricity bill (PDF or image, max 50MB)
3. Click **Extract Data & Calculate Solar**
4. Loading screen shows progress (Gemini takes 10–20 seconds to read the bill)
5. Results page shows extracted data + solar recommendation
6. Click **Download Excel Report** to get the filled Excel file
7. Click **← Process another bill** to go back and upload again

---

## Validations & Error Handling

- Only PDF, JPG, PNG files accepted
- Max file size: 50 MB
- If uploaded file is not an electricity bill → warning shown, no Excel generated
- If units consumed not found in bill → clear error message shown
- Friendly error pages for all server errors (no blank browser error pages)
