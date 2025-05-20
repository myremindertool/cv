import streamlit as st
import pdfplumber
import docx
import gspread
import re
from io import BytesIO
from google.oauth2.service_account import Credentials
from datetime import datetime
from dateutil.relativedelta import relativedelta
import pandas as pd
import json
import openai  # ✅ Use this form of import

# ✅ Initialize OpenAI using the builder method (Streamlit-safe)
client = openai.OpenAI(api_key=st.secrets["openai"]["api_key"])

# Column order for Google Sheets & Excel
FIELDS_ORDER = [
    "Name", "Nationality", "Qualification", "Experience",
    "Current Salary", "Expected Salary", "Position", "Source",
    "Status", "Remark"
]

# Connect to Google Sheet
def connect_to_gsheet():
    scope = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/drive"
    ]
    creds = Credentials.from_service_account_info(st.secrets["gcp_json"], scopes=scope)
    client_gsheet = gspread.authorize(creds)
    return client_gsheet.open_by_key("1DX9RAwrLZv6uJ_aCKLMOi8DmSQXVuL6rfOgwsVKOHXw").sheet1

# Extract text from PDF
def extract_text_from_pdf(uploaded_file):
    with pdfplumber.open(uploaded_file) as pdf:
        return "\n".join(page.extract_text() or "" for page in pdf.pages)

# Extract text from DOCX
def extract_text_from_docx(uploaded_file):
    doc = docx.Document(uploaded_file)
    return "\n".join(p.text for p in doc.paragraphs)

# Extract fields from CV using GPT-4
def extract_fields_with_ai(text):
    prompt = (
        "Extract the following fields from this CV:\n"
        "1. Name\n2. Nationality\n3. Qualification\n\n"
        f"CV Text:\n{text}\n\n"
        "Respond in JSON format with keys: Name, Nationality, Qualification."
    )
    response = client.chat.completions.create(
        model="gpt-4",
        messages=[{"role": "user", "content": prompt}],
        temperature=0.2
    )
    try:
        content = response.choices[0].message.content
        return json.loads(content)
    except:
        return {key: "" for key in ["Name", "Nationality", "Qualification"]}

# Normalize date like "Jan 2021"
def normalize_date(date_str):
    try:
        return datetime.strptime(date_str, "%b %Y")
    except:
        return None

# Extract work periods
def extract_experience_blocks(text):
    pattern = re.compile(r"(?i)(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*[\s,]+\d{4}\s*[-–]+\s*(Present|\d{4})")
    matches = pattern.findall(text)
    results = []
    for match in matches:
        start = normalize_date(f"{match[0]} {match[1] if match[1] != 'Present' else datetime.today().year}")
        end = datetime.today() if match[1] == 'Present' else normalize_date(match[1])
        if start and end:
            results.append((start, end))
    return results

# Merge overlapping periods
def merge_periods(periods):
    periods.sort()
    merged = []
    for start, end in periods:
        if not merged:
            merged.append((start, end))
        else:
            last_start, last_end = merged[-1]
            if start <= last_end + relativedelta(months=1):
                merged[-1] = (last_start, max(last_end, end))
            else:
                merged.append((start, end))
    return merged

# Calculate total experience
def calculate_total_experience(periods):
    total_months = sum((relativedelta(e, s).years * 12 + relativedelta(e, s).months) for s, e in periods)
    return f"{total_months // 12} years {total_months % 12} months"

# Generate Excel download
def generate_excel_download(data):
    buffer = BytesIO()
    df = pd.DataFrame([data])
    df = df.reindex(columns=FIELDS_ORDER, fill_value="")
    df.to_excel(buffer, index=False)
    buffer.seek(0)
    return buffer

# Streamlit App UI
st.title("📄 Smart CV Extractor for Schools (AI-Powered)")

uploaded_file = st.file_uploader("Upload CV (.pdf or .docx)", type=["pdf", "docx"])

if uploaded_file:
    file_type = uploaded_file.name.split(".")[-1].lower()
    text = extract_text_from_pdf(uploaded_file) if file_type == "pdf" else extract_text_from_docx(uploaded_file)

    if not text.strip():
        st.error("❌ Could not extract text. Please check the file format or content.")
    else:
        fields = extract_fields_with_ai(text)
        periods = extract_experience_blocks(text)
        merged = merge_periods(periods)
        experience = calculate_total_experience(merged)

        st.subheader("📋 Extracted Information")
        st.write("Experience:", experience)
        for key in ["Name", "Nationality", "Qualification"]:
            st.write(f"{key}:", fields.get(key, ""))

        row = [
            fields.get("Name", ""),
            fields.get("Nationality", ""),
            fields.get("Qualification", ""),
            experience,
            "", "", "", "", "", ""
        ]

        if st.button("✅ Append to Google Sheet"):
            sheet = connect_to_gsheet()
            sheet.append_row(row)
            st.success("Data added to Google Sheet successfully.")

        download_data = dict(zip(FIELDS_ORDER, row))
        excel_file = generate_excel_download(download_data)
        st.download_button(
            label="⬇️ Download Excel",
            data=excel_file,
            file_name="extracted_candidate.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
