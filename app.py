import streamlit as st
import pandas as pd
import requests
import json
import io
import time
import random
from typing import Dict, List, Any
from google import genai
from google.genai import errors
import pdfplumber
from bs4 import BeautifulSoup

# ------------------ Page Config ------------------
st.set_page_config(page_title="Itinerary Extractor", layout="wide")
st.title("📄 AI Itinerary Extractor (Gemini 2.0 Flash)")
st.markdown("Upload an Excel file with columns: **Country, Client name, Website, Source URL/Path**")

@st.cache_data
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
            with pdfplumber.open(io.BytesIO(response.content)) as pdf:
                text = "\n".join(page.extract_text() or "" for page in pdf.pages)
            return text
        else:
            soup = BeautifulSoup(response.text, "html.parser")
            for script in soup(["script", "style"]):
                script.decompose()
            return soup.get_text(separator="\n", strip=True)
    except Exception as e:
        return f"ERROR: {str(e)}"

def call_gemini_with_retry(api_key: str, document_text: str, max_retries: int = 3) -> Dict[str, Any]:
    """Send document text to Gemini with retry logic for 429 and 503 errors."""
    client = genai.Client(api_key=api_key)
    
    # More concise prompt to save tokens
    prompt = f"""
Analyze the travel itinerary document and extract structured data. Return ONLY valid JSON in this format:

{{
  "itinerary": [{{"day": integer, "city": "string", "activity": "string"}}],
  "hotels": [{{"city": "string", "hotel_name": "string", "note": "string"}}],
  "restaurants": [{{"city": "string", "restaurant_name": "string", "meal_type": "string", "note": "string"}}],
  "attractions": [{{"city": "string", "attraction_name": "string", "description": "string"}}]
}}

If a category has no data, return an empty list [].

DOCUMENT (truncated for length):
{document_text[:20000]}
"""
    for attempt in range(max_retries):
        try:
            response = client.models.generate_content(
                model="gemini-2.0-flash",
                contents=prompt
            )
            raw = response.text.strip()
            # Clean markdown code fences
            if raw.startswith("```json"):
                raw = raw[7:]
            if raw.startswith("```"):
                raw = raw[3:]
            if raw.endswith("```"):
                raw = raw[:-3]
            raw = raw.strip()
            data = json.loads(raw)
            for key in ["itinerary", "hotels", "restaurants", "attractions"]:
                if key not in data:
                    data[key] = []
            return data
        except errors.ClientError as e:
            if "429" in str(e) or "503" in str(e):
                if attempt < max_retries - 1:
                    # Exponential backoff: 2^attempt seconds + jitter
                    wait_time = (2 ** attempt) + random.uniform(0, 1)
                    st.warning(f"⚠️ Rate limit hit. Retrying in {wait_time:.1f} seconds... (Attempt {attempt + 1}/{max_retries})")
                    time.sleep(wait_time)
                    continue
                else:
                    st.error(f"Gemini API error after {max_retries} retries: {str(e)}")
                    return {"itinerary": [], "hotels": [], "restaurants": [], "attractions": []}
            else:
                st.error(f"Gemini API error: {str(e)}")
                return {"itinerary": [], "hotels": [], "restaurants": [], "attractions": []}
        except Exception as e:
            st.error(f"Unexpected error: {str(e)}")
            return {"itinerary": [], "hotels": [], "restaurants": [], "attractions": []}
    
    return {"itinerary": [], "hotels": [], "restaurants": [], "attractions": []}

def process_row(row: pd.Series, api_key: str, skip_on_quota: bool, progress_bar, status_text) -> Dict[str, List]:
    """Process one row with quota skip option."""
    country = row["Country"]
    client = row["Client name"]
    website = row["Website"]
    source = row["Source URL/Path"]
    
    status_text.text(f"Processing: {client} - {source}")
    text = extract_text_from_url(source)
    if text.startswith("ERROR"):
        st.warning(f"Failed to fetch {source}: {text}")
        return {"itinerary": [], "hotels": [], "restaurants": [], "attractions": []}
    
    extracted = call_gemini_with_retry(api_key, text)
    
    # Check if we hit quota and user wants to skip
    if skip_on_quota and not any(extracted.values()):
        st.info(f"⚠️ Skipping {client} due to quota limit. Check 'quota_skipped.txt'")
        with open("quota_skipped.txt", "a") as f:
            f.write(f"{client},{source}\n")
        return {"itinerary": [], "hotels": [], "restaurants": [], "attractions": []}
    
    # Add metadata
    for item in extracted["itinerary"]:
        item.update({"Country": country, "Client name": client, "Website": website})
    for item in extracted["hotels"]:
        item.update({"Country": country, "Client name": client, "Website": website})
    for item in extracted["restaurants"]:
        item.update({"Country": country, "Client name": client, "Website": website})
    for item in extracted["attractions"]:
        item.update({"Country": country, "Client name": client, "Website": website})
    
    return extracted

# ------------------ Main App ------------------
api_key = st.text_input("🔑 Enter your Google Gemini API Key", type="password", 
                        help="Get one free at https://aistudio.google.com/apikey")
uploaded_file = st.file_uploader("📂 Upload Excel file", type=["xlsx", "xls"])
skip_on_quota = st.checkbox("Skip URLs that fail due to quota limits", value=True)

if api_key and uploaded_file:
    df = pd.read_excel(uploaded_file)
    required_cols = ["Country", "Client name", "Website", "Source URL/Path"]
    if not all(col in df.columns for col in required_cols):
        st.error(f"Excel must contain columns: {', '.join(required_cols)}")
        st.stop()
    
    if st.button("🚀 Start Extraction"):
        all_itinerary, all_hotels, all_restaurants, all_attractions = [], [], [], []
        
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        for idx, row in df.iterrows():
            extracted = process_row(row, api_key, skip_on_quota, progress_bar, status_text)
            all_itinerary.extend(extracted["itinerary"])
            all_hotels.extend(extracted["hotels"])
            all_restaurants.extend(extracted["restaurants"])
            all_attractions.extend(extracted["attractions"])
            progress_bar.progress((idx + 1) / len(df))
            time.sleep(1)  # 1 second delay between rows
        
        status_text.text("✅ Extraction complete! Building Excel...")
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
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
        if skip_on_quota:
            st.info("ℹ️ URLs skipped due to quota are listed in 'quota_skipped.txt'. You can retry them later.")
        st.success("Processing complete!")
