import streamlit as st
import pandas as pd
import requests
import json
import io
from typing import Dict, List, Any
from google import genai
import pdfplumber
from bs4 import BeautifulSoup
import time

# ------------------ Page Config ------------------
st.set_page_config(page_title="Itinerary Extractor", layout="wide")
st.title("📄 AI Itinerary Extractor (Gemini 2.0 Flash)")
st.markdown("Upload an Excel file with columns: **Country, Client name, Website, Source URL/Path**")

# ------------------ Helper Functions ------------------
def extract_text_from_url(url: str) -> str:
    """Download and extract text from a URL (HTML or PDF) with browser headers."""
    try:
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8",
            "Accept-Language": "en-US,en;q=0.5",
            "Accept-Encoding": "gzip, deflate, br",
            "Connection": "keep-alive",
            "Upgrade-Insecure-Requests": "1",
        }
        response = requests.get(url, timeout=30, headers=headers)
        response.raise_for_status()
        content_type = response.headers.get("content-type", "").lower()
        
        if "application/pdf" in content_type or url.lower().endswith(".pdf"):
            # PDF
            with pdfplumber.open(io.BytesIO(response.content)) as pdf:
                text = "\n".join(page.extract_text() or "" for page in pdf.pages)
            return text
        else:
            # HTML
            soup = BeautifulSoup(response.text, "html.parser")
            for script in soup(["script", "style"]):
                script.decompose()
            return soup.get_text(separator="\n", strip=True)
    except Exception as e:
        return f"ERROR: {str(e)}"

def call_gemini(api_key: str, document_text: str) -> Dict[str, Any]:
    """Send document text to Gemini 2.0 Flash and return parsed JSON."""
    client = genai.Client(api_key=api_key)
    
    prompt = f"""
You are an expert travel itinerary parser. Extract the following information from the document below.
Return ONLY valid JSON (no extra text) in the following structure:

{{
  "itinerary": [
    {{"day": integer, "city": "string", "activity": "string"}}
  ],
  "hotels": [
    {{"city": "string", "hotel_name": "string", "note": "string (optional)"}}
  ],
  "restaurants": [
    {{"city": "string", "restaurant_name": "string", "meal_type": "string", "note": "string"}}
  ],
  "attractions": [
    {{"city": "string", "attraction_name": "string", "description": "string"}}
  ]
}}

If any category is not found, return an empty list for that key.

DOCUMENT:
{document_text[:70000]}
"""
    try:
        response = client.models.generate_content(
            model="gemini-2.0-flash",
            contents=prompt
        )
        raw = response.text.strip()
        # Remove markdown code fences if present
        if raw.startswith("```json"):
            raw = raw[7:]
        if raw.startswith("```"):
            raw = raw[3:]
        if raw.endswith("```"):
            raw = raw[:-3]
        raw = raw.strip()
        data = json.loads(raw)
        # Ensure all keys exist
        for key in ["itinerary", "hotels", "restaurants", "attractions"]:
            if key not in data:
                data[key] = []
        return data
    except Exception as e:
        st.error(f"Gemini parsing error: {str(e)}")
        return {"itinerary": [], "hotels": [], "restaurants": [], "attractions": []}

def process_row(row: pd.Series, api_key: str, progress_bar, status_text) -> Dict[str, List]:
    """Process one row: download, call Gemini, return structured data."""
    country = row["Country"]
    client = row["Client name"]
    website = row["Website"]
    source = row["Source URL/Path"]
    
    status_text.text(f"Processing: {client} - {source}")
    text = extract_text_from_url(source)
    if text.startswith("ERROR"):
        st.warning(f"Failed to fetch {source}: {text}")
        return {"itinerary": [], "hotels": [], "restaurants": [], "attractions": []}
    
    extracted = call_gemini(api_key, text)
    # Add metadata to each item
    for item in extracted["itinerary"]:
        item["Country"] = country
        item["Client name"] = client
        item["Website"] = website
    for item in extracted["hotels"]:
        item["Country"] = country
        item["Client name"] = client
        item["Website"] = website
    for item in extracted["restaurants"]:
        item["Country"] = country
        item["Client name"] = client
        item["Website"] = website
    for item in extracted["attractions"]:
        item["Country"] = country
        item["Client name"] = client
        item["Website"] = website
    return extracted

# ------------------ Main App ------------------
api_key = st.text_input("🔑 Enter your Google Gemini API Key", type="password", 
                        help="Get one free at https://aistudio.google.com/apikey")
uploaded_file = st.file_uploader("📂 Upload Excel file", type=["xlsx", "xls"])

if api_key and uploaded_file:
    df = pd.read_excel(uploaded_file)
    required_cols = ["Country", "Client name", "Website", "Source URL/Path"]
    if not all(col in df.columns for col in required_cols):
        st.error(f"Excel must contain columns: {', '.join(required_cols)}")
        st.stop()
    
    if st.button("🚀 Start Extraction"):
        all_itinerary = []
        all_hotels = []
        all_restaurants = []
        all_attractions = []
        
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        for idx, row in df.iterrows():
            extracted = process_row(row, api_key, progress_bar, status_text)
            all_itinerary.extend(extracted["itinerary"])
            all_hotels.extend(extracted["hotels"])
            all_restaurants.extend(extracted["restaurants"])
            all_attractions.extend(extracted["attractions"])
            progress_bar.progress((idx + 1) / len(df))
            time.sleep(0.5)  # avoid hitting rate limits
        
        status_text.text("✅ Extraction complete! Building Excel...")
        
        # Create Excel with four sheets
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            # Only create sheet if there is data, otherwise empty placeholder
            pd.DataFrame(all_itinerary).to_excel(writer, sheet_name="Itinerary", index=False)
            pd.DataFrame(all_restaurants).to_excel(writer, sheet_name="Restaurant", index=False)
            pd.DataFrame(all_hotels).to_excel(writer, sheet_name="Hotel", index=False)
            pd.DataFrame(all_attractions).to_excel(writer, sheet_name="Attraction", index=False)
        
        output.seek(0)
        st.download_button(
            label="📥 Download Excel file",
            data=output,
            file_name="extracted_itineraries.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        st.success("Done!")
