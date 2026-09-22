import io
import json
import streamlit as st
from pypdf import PdfReader
from openai import OpenAI
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor

st.set_page_config(page_title="TrialDeck - Clinical Study to PPTX", layout="centered")

st.title("📊 TrialDeck: Study to Journal Club Slides")
st.caption("Upload a clinical trial publication PDF to generate an editable, structured PowerPoint deck.")

# API Key handling (Sidebar input or Streamlit Secrets)
api_key = st.sidebar.text_input("OpenAI API Key", type="password")
if not api_key and "OPENAI_API_KEY" in st.secrets:
    api_key = st.secrets["OPENAI_API_KEY"]

uploaded_file = st.file_uploader("Upload Trial PDF (Primary Publication)", type=["pdf"])

def extract_pdf_text(file) -> str:
    reader = PdfReader(file)
    text = ""
    # Extract first 4 pages (sufficient for title, abstract, design, primary results, and safety)
    for page in reader.pages[:4]:
        text += page.extract_text() or ""
    return text[:14000]

def extract_clinical_json(text: str, client: OpenAI) -> dict:
    prompt = f"""
    You are an expert clinical biostatistician and evidence synthesis specialist.
    Extract the pivotal data from this clinical study text into a strict JSON object.
    
    Guardrail: Do NOT calculate or guess numbers. If an element is absent, set value to 'Not reported'.

    JSON Schema:
    {{
      "trial_acronym": "Acronym or Primary Drug Name",
      "full_title": "Full study title",
      "authors_journal": "Lead author, Journal Name, Year",
      "study_design": "Allocation, blinding, control type, group structure",
      "study_duration": "Total on-treatment time (weeks/months) + follow-up",
      "sample_size": "Total N (ITT population)",
      "population_criteria": "Inclusion criteria / disease baseline (1-2 sentences)",
      "arms": "Active arm(s) with doses vs. Comparator/Placebo",
      "primary_endpoint": {{
        "metric_name": "Primary outcome name and target timeframe",
        "results": "Intervention result vs. Control result",
        "statistics": "Treatment difference/effect, 95% CI, p-value"
      }},
      "safety_discontinuations": "Overall and AE/GI-specific discontinuation rates",
      "clinical_takeaway": [
        "Primary clinical headline & statistical significance",
        "Key trial strength",
        "Key trial limitation"
      ]
    }}

    Source Text:
    {text}
    """

    response = client.chat.completions.create(
        model="gpt-4o-mini",
        messages=[{"role": "user", "content": prompt}],
        response_format={"type": "json_object"}
    )
    return json.loads(response.choices[0].message.content)

def build_pptx(data: dict) -> io.BytesIO:
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    navy = RGBColor(26, 54, 93)
    charcoal = RGBColor(45, 55, 72)
    accent_blue = RGBColor(49, 130, 206)

    def add_header(slide, title_text, category="EVIDENCE SYNTHESIS & APPRAISAL"):
        cat_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(11.5), Inches(0.4))
        p0 = cat_box.text_frame.paragraphs[0]
        p0.text = category.upper()
        p0.font.size = Pt(11)
        p0.font.bold = True
        p0.font.color.rgb = accent_blue

        tb = slide.shapes.add_textbox(Inches(0.8), Inches(0.7), Inches(11.5), Inches(0.8))
        p = tb.text_frame.paragraphs[0]
        p.text = title_text
        p.font.size = Pt(24)
        p.font.bold = True
        p.font.color.rgb = navy

    # --- Slide 1: Title & Citation ---
    s1 = prs.slides.add_slide(blank_layout)
    tb1 = s1.shapes.add_textbox(Inches(1.0), Inches(2.0), Inches(11.3), Inches(3.5))
    tf1 = tb1.text_frame
    tf1.word_wrap = True
    
    p = tf1.paragraphs[0]
    p.text = data.get("trial_acronym", "Clinical Trial Evaluation")
    p.font.size = Pt(36)
    p.font.bold = True
    p.font.color.rgb = navy

    p2 = tf1.add_paragraph()
    p2.text = data.get("full_title", "")
    p2.font.size = Pt(16)
    p2.font.color.rgb = charcoal

    p3 = tf1.add_paragraph()
    p3.text = f"\nCitation: {data.get('authors_journal', 'Published Clinical Evidence')}"
    p3.font.size = Pt(13)
    p3.font.italic = True

    # --- Slide 2: Design & Methodology ---
    s2 = prs.slides.add_slide(blank_layout)
    add_header(s2, "Study Design & Protocol Parameters")
    tb2 = s2.shapes.add_textbox(Inches(0.8), Inches(1.8), Inches(11.5), Inches(5.0))
    tf2 = tb2.text_frame
    tf2.word_wrap = True

    items_s2 = [
        f"Design: {data.get('study_design', 'Not reported')}",
        f"Duration: {data.get('study_duration', 'Not reported')}",
        f"Sample Size: {data.get('sample_size', 'Not reported')}",
        f"Target Population: {data.get('population_criteria', 'Not reported')}",
        f"Treatment Arms: {data.get('arms', 'Not reported')}"
    ]
    for idx, item in enumerate(items_s2):
        bp = tf2.paragraphs[0] if idx == 0 else tf2.add_paragraph()
        bp.text = f"•  {item}"
        bp.font.size = Pt(15)
        bp.font.color.rgb = charcoal

    # --- Slide 3: Primary Efficacy Endpoints ---
    s3 = prs.slides.add_slide(blank_layout)
    add_header(s3, "Primary Efficacy Outcomes (ITT Analysis)")
    tb3 = s3.shapes.add_textbox(Inches(0.8), Inches(1.8), Inches(11.5), Inches(5.0))
    tf3 = tb3.text_frame
    tf3.word_wrap = True
    pe = data.get("primary_endpoint", {})

    p = tf3.paragraphs[0]
    p.text = f"Primary Metric: {pe.get('metric_name', 'Primary Outcome')}"
    p.font.size = Pt(18)
    p.font.bold = True
    p.font.color.rgb = navy

    p2 = tf3.add_paragraph()
    p2.text = f"\nObserved Values:\n{pe.get('results', 'Not reported')}\n"
    p2.font.size = Pt(16)
    p2.font.color.rgb = charcoal

    p3 = tf3.add_paragraph()
    p3.text = f"Treatment Difference & Statistical Significance:\n{pe.get('statistics', 'Not reported')}"
    p3.font.size = Pt(17)
    p3.font.bold = True
    p3.font.color.rgb = accent_blue

    # --- Slide 4: Safety & Discontinuations ---
    s4 = prs.slides.add_slide(blank_layout)
    add_header(s4, "Safety, Tolerability & Discontinuations")
    tb4 = s4.shapes.add_textbox(Inches(0.8), Inches(1.8), Inches(11.5), Inches(5.0))
    tf4 = tb4.text_frame
    tf4.word_wrap = True

    p = tf4.paragraphs[0]
    p.text = "Discontinuation & Adverse Event Profile:"
    p.font.size = Pt(18)
    p.font.bold = True
    p.font.color.rgb = navy

    p2 = tf4.add_paragraph()
    p2.text = f"\n{data.get('safety_discontinuations', 'Not reported')}"
    p2.font.size = Pt(16)
    p2.font.color.rgb = charcoal

    # --- Slide 5: Critical Appraisal & Conclusions ---
    s5 = prs.slides.add_slide(blank_layout)
    add_header(s5, "Clinical Appraisal & Key Takeaways")
    tb5 = s5.shapes.add_textbox(Inches(0.8), Inches(1.8), Inches(11.5), Inches(5.0))
    tf5 = tb5.text_frame
    tf5.word_wrap = True
    for idx, item in enumerate(data.get("clinical_takeaway", [])):
        bp = tf5.paragraphs[0] if idx == 0 else tf5.add_paragraph()
        bp.text = f"•  {item}"
        bp.font.size = Pt(15)
        bp.font.color.rgb = charcoal

    pptx_io = io.BytesIO()
    prs.save(pptx_io)
    pptx_io.seek(0)
    return pptx_io

if uploaded_file and st.button("Generate PowerPoint Deck"):
    if not api_key:
        st.error("Please provide an OpenAI API key.")
    else:
        with st.spinner("Extracting parameters and compiling presentation..."):
            client = OpenAI(api_key=api_key)
            raw_text = extract_pdf_text(uploaded_file)
            trial_json = extract_clinical_json(raw_text, client)
            deck = build_pptx(trial_json)

            st.success("Slide deck generated successfully!")
            
            with st.expander("View Structured Extraction Preview"):
                st.json(trial_json)

            st.download_button(
                label="📥 Download Presentation (.pptx)",
                data=deck,
                file_name=f"{trial_json.get('trial_acronym', 'Trial')}_Journal_Club.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )
