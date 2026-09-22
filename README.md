# TrialDeck: Clinical Study Evidence Synthesizer & Slide Engine

TrialDeck is an automated clinical informatics utility that transforms peer-reviewed clinical trial PDFs into structured, presentation-ready 6-slide PowerPoint decks (`.pptx`) for Journal Clubs, Pharmacy & Therapeutics (P&T) evaluations, and grand rounds.

---

## 🎯 Clinical Problem
Clinicians, residents, and formulary specialists spend hours manually transcribing study parameters, baseline demographics, statistical outcome deltas, and safety profiles from primary literature PDFs into presentation slides. General AI summarizers frequently miss vital statistical endpoints ($p$-values, confidence intervals, hazard ratios) or omit protocol architecture.

TrialDeck addresses this bottleneck by enforcing a deterministic clinical evidence schema tailored to standard comparative effectiveness reviews.

---

## ✨ Features
* **PDF Ingestion & Text Parsing:** Ingests primary trial publications and parses pivotal sections (Abstract, Methods, Primary Outcomes, Adverse Events).
* **Structured Clinical Extraction:** Enforces a rigid JSON extraction schema capturing:
  * Study design architecture (randomization, masking, active/placebo control, parallel/crossover).
  * On-treatment duration and safety follow-up windows.
  * Intention-to-Treat (ITT) population baselines.
  * Primary efficacy endpoints with exact effect sizes, 95% CIs, and $p$-values.
  * Safety profiles, adverse event rates, and discontinuation differentials.
* **Deterministic Guardrails:** Strict prompt constraints prevent hallucinations; metrics not explicitly reported in the text are populated with `"Not reported"`.
* **Automated `.pptx` Compilation:** Generates a downloadable, cleanly styled 16:9 widescreen slide deck using `python-pptx`.

---

## 🛠️ Architecture & Tech Stack
* **Frontend / UI:** [Streamlit](https://streamlit.io/)
* **Document Parsing:** `pypdf`
* **Extraction Engine:** OpenAI API (`gpt-4o-mini`) via structured JSON schema enforcement
* **Slide Generation Engine:** `python-pptx`
* **Deployment:** Streamlit Community Cloud

---

## 🚀 Getting Started Locally

### Prerequisites
* Python 3.10+
* OpenAI API Key

### Installation
1. Clone this repository:
   ```bash
   git clone [https://github.com/YOUR_USERNAME/trialdeck-app.git](https://github.com/YOUR_USERNAME/trialdeck-app.git)
   cd trialdeck-app
