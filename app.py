import os
import io
import json
from datetime import datetime

import streamlit as st
import pandas as pd
from docx import Document
from PyPDF2 import PdfReader
from openai import OpenAI


# -----------------------------
# Page setup
# -----------------------------
st.set_page_config(
    page_title="Allocator Diligence Engine",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Allocator Diligence Engine")
st.caption("AI-assisted manager diligence intake, red-flag review, follow-up question generation, and IC memo drafting.")


# -----------------------------
# Helper functions
# -----------------------------
def read_docx(uploaded_file) -> str:
    doc = Document(uploaded_file)
    paragraphs = [p.text for p in doc.paragraphs if p.text.strip()]
    return "\n".join(paragraphs)


def read_pdf(uploaded_file) -> str:
    reader = PdfReader(uploaded_file)
    pages = []
    for page in reader.pages:
        text = page.extract_text() or ""
        pages.append(text)
    return "\n\n".join(pages)


def read_xlsx(uploaded_file) -> str:
    xls = pd.ExcelFile(uploaded_file)
    output = []
    for sheet in xls.sheet_names:
        df = pd.read_excel(uploaded_file, sheet_name=sheet)
        output.append(f"\n--- Sheet: {sheet} ---\n")
        output.append(df.to_string(index=False))
    return "\n".join(output)


def read_txt(uploaded_file) -> str:
    return uploaded_file.read().decode("utf-8", errors="ignore")


def extract_file_text(uploaded_files) -> str:
    combined = []
    for file in uploaded_files:
        name = file.name.lower()
        try:
            if name.endswith(".docx"):
                text = read_docx(file)
            elif name.endswith(".pdf"):
                text = read_pdf(file)
            elif name.endswith(".xlsx"):
                text = read_xlsx(file)
            elif name.endswith(".txt"):
                text = read_txt(file)
            else:
                text = f"[Unsupported file type: {file.name}]"

            combined.append(f"\n\n================ FILE: {file.name} ================\n{text}")
        except Exception as e:
            combined.append(f"\n\n================ FILE: {file.name} ================\n[Error reading file: {e}]")
    return "\n".join(combined)


def build_prompt(manager_info, extracted_docs, raw_notes):
    return f"""
You are an experienced institutional allocator preparing diligence materials for an investment committee.

Your task is to review the supplied manager materials, notes, and quantitative tear sheet data. Do not simply summarize. Apply a skeptical allocator lens.

Important rules:
- Separate facts from interpretation.
- Identify missing information.
- Flag unsupported manager claims.
- Look for hidden beta, crowding, liquidity mismatch, capacity issues, incentive misalignment, weak short alpha, aggressive underwriting, or other allocator-relevant risks.
- Do not make up facts that are not supported by the materials.
- Use a concise, investment-committee-ready tone.
- This is not an investment recommendation system. The output is a draft for human allocator review.

Manager / allocator input:
{json.dumps(manager_info, indent=2)}

Raw allocator notes:
{raw_notes}

Uploaded diligence materials:
{extracted_docs}

Produce the following output:

# 1. Structured Intake Summary
- Manager / fund
- Strategy
- AUM
- Proposed role in portfolio
- Liquidity / terms
- Performance profile
- Exposure / risk profile
- Key facts extracted from materials

# 2. Investment Committee Memo Draft
Use this structure:
## Executive Summary
## Strategy & Portfolio Fit
## Investment Rationale
## Team Assessment
## Process & Edge
## Performance Assessment
## Risk Management
## Business / Operational Risk
## Key Risks / Red Flags
## Follow-Up Questions
## Preliminary Recommendation

# 3. Red Flag Review
Group into:
- Hard red flags
- Soft concerns
- Missing information
- Claims requiring verification

# 4. Follow-Up Questions
Group by:
- Team
- Strategy / process
- Portfolio construction
- Risk management
- Performance / regime behavior
- Liquidity / terms
- Operations / business risk

# 5. Preliminary Recommendation
Choose one:
- Proceed
- Watchlist / Continue Diligence
- Decline

Explain the rationale clearly.
"""


def call_openai(prompt, model, api_key):
    client = OpenAI(api_key=api_key)
    response = client.responses.create(
        model=model,
        input=prompt,
    )
    return response.output_text


# -----------------------------
# Sidebar
# -----------------------------
with st.sidebar:
    st.header("Settings")

    # Use Streamlit secret if deployed
    api_key = st.secrets.get("OPENAI_API_KEY", "")

    # Fallback for local testing
    if not api_key:
        api_key = st.text_input(
            "OpenAI API Key",
            type="password",
            help="For local testing only."
        )

    model = st.text_input(
        "Model",
        value="gpt-4.1-mini",
        help="You can change this to another OpenAI model available to your account."
    )

    st.markdown("---")
    st.markdown("### Demo guidance")
    st.write(
        "Use fictional/public materials only. "
        "Do not upload confidential manager documents or internal work materials."
    )
# -----------------------------
# Main input area
# -----------------------------
st.subheader("1. Manager Overview")

col1, col2 = st.columns(2)

with col1:
    manager_name = st.text_input("Manager Name", value="North River Capital Partners")
    fund_name = st.text_input("Fund Name", value="North River Market Neutral Fund")
    strategy = st.selectbox(
        "Strategy Type",
        [
            "Market Neutral",
            "L/S Equity",
            "Credit",
            "Private Credit",
            "Macro",
            "Event Driven",
            "Relative Value",
            "Private Equity",
            "Real Assets",
            "Other"
        ],
        index=0
    )
    aum = st.text_input("AUM", value="$2.5bn")

with col2:
    proposed_allocation = st.text_input("Proposed Allocation", value="$50mm")
    portfolio_role = st.text_area(
        "Proposed Portfolio Role",
        value="Potential low-beta equity diversifier within the absolute return portfolio.",
        height=90
    )
    initial_concerns = st.text_area(
        "Initial Allocator Concerns",
        value="Rapid AUM growth, potential crowding in quality-growth/AI names, and unclear short alpha contribution.",
        height=90
    )

st.subheader("2. Notes and Diligence Materials")

raw_notes = st.text_area(
    "Paste raw call notes or transcript excerpt",
    height=180,
    placeholder="Paste rough notes, call transcript, or meeting notes here..."
)

uploaded_files = st.file_uploader(
    "Upload manager materials",
    type=["pdf", "docx", "xlsx", "txt"],
    accept_multiple_files=True,
    help="Upload pitchbooks, DDQs, tear sheets, notes, or transcripts."
)

generate = st.button("Generate Memo Package", type="primary")

# -----------------------------
# Generate output
# -----------------------------
if generate:
    if not api_key:
        st.error("Please enter your OpenAI API key in the sidebar before generating.")
        st.stop()

    manager_info = {
        "manager_name": manager_name,
        "fund_name": fund_name,
        "strategy": strategy,
        "aum": aum,
        "proposed_allocation": proposed_allocation,
        "portfolio_role": portfolio_role,
        "initial_concerns": initial_concerns,
        "generated_at": datetime.now().isoformat(timespec="seconds"),
    }

    with st.spinner("Reading uploaded files..."):
        extracted_docs = extract_file_text(uploaded_files) if uploaded_files else ""

    with st.spinner("Generating allocator diligence package..."):
        prompt = build_prompt(manager_info, extracted_docs, raw_notes)
        try:
            output = call_openai(prompt, model, api_key)
        except Exception as e:
            st.error(f"OpenAI call failed: {e}")
            st.stop()

    st.success("Memo package generated. Review carefully before using or sharing.")

    tab1, tab2, tab3 = st.tabs(["AI Output", "Extracted Source Text", "Prompt Sent to Model"])

    with tab1:
        st.markdown(output)
        st.download_button(
            "Download Output as Markdown",
            data=output,
            file_name=f"{manager_name.replace(' ', '_')}_memo_package.md",
            mime="text/markdown"
        )

    with tab2:
        st.text_area("Extracted text from uploaded files", extracted_docs, height=500)

    with tab3:
        st.text_area("Prompt", prompt, height=500)
else:
    st.info("Enter manager details, upload fictional diligence materials, then click Generate Memo Package.")
