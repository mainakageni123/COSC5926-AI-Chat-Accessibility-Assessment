# COSC 5926: Assessing the Accessible Usability of AI Chat Applications

**Course:** COSC 5926 – Advanced Human Computer Interaction  
**Department:** Department of Computer Science & Mathematics, Algoma University  
**Topic:** Assessing the Accessible Usability of AI Chat Applications (ChatGPT, Google Gemini, and Anthropic Claude)  

---

## 📌 Project Overview
Generative Artificial Intelligence (AI) chat applications such as **ChatGPT (OpenAI)**, **Google Gemini**, and **Anthropic Claude** are increasingly used in education, professional workflows, and daily productivity. While these systems offer powerful capabilities, their real-world impact depends on both their **usability** and **accessibility** for diverse user groups, including individuals with visual, motor, and cognitive disabilities.

This project implements an **Accessible Usability** evaluation framework combining:
1. **Web Content Accessibility Guidelines (WCAG 2.1)** – Perceivable, Operable, Understandable, and Robust (POUR) principles.
2. **Cognitive Accessibility Checklist** – ADHD, Dyslexia, Executive Function, Autism-friendly design, and Memory accessibility.
3. **AI-Specific Accessible Usability Checklist** – Uncertainty disclosure, source citations, cognitive load management, prompt guidance, and response adaptations.

---

## 📊 Summary of Findings

Across **93 evaluation criteria** evaluated independently by **4 raters** (1,116 total test observations):

| Application | Pass Count | Fail Count | Overall Pass Rate | Primary Strengths | Critical Gaps |
| :--- | :---: | :---: | :---: | :--- | :--- |
| **Google Gemini** | **328** | **44** | **88.2%** | Highest overall score, provides source citation links, excellent cognitive support | Lacks chat history search and folder organization |
| **Anthropic Claude** | **320** | **52** | **86.0%** | Clear plain language (98% Understandable), uses hedging to disclose uncertainty | Lacks voice input, no suggested prompt starters |
| **OpenAI ChatGPT** | **319** | **53** | **85.8%** | Best conversation management (history search & folders), multimodal voice input | Low source transparency, subtle focus indicators |

### Key Universal Deficiencies Across All Three Platforms:
* **Dynamic Streaming Announcements:** Streaming responses are not announced to screen readers (e.g., NVDA) via ARIA live regions.
* **Audio/Video Accessibility:** Transcripts for audio features and captions for media are largely absent.
* **Cognitive Controls:** None provide notification control settings or native built-in text-to-speech (TTS).
* **AI Uncertainty & Transparency:** Limitations and confidence levels are not systematically communicated across outputs.

---

## 📈 Visualizations

* **Summary Pass/Fail Chart:** `Summary_PassFail_Chart.png` (Pass vs. Fail count breakdown per category)
* **Accessibility Heatmap:** `Comparison_Heatmap.png` (Category pass percentage across ChatGPT, Gemini, and Claude)
* **Radar Profile:** `Radar_Comparison.png` (Multi-axis accessibility profile comparison)
* **Cohen's Kappa Heatmap:** `Cohens_Kappa_Heatmap.png` (Inter-rater reliability across 6 evaluator pairs)
* **Average Kappa Comparison:** `Average_Kappa_BarChart.png` (Category agreement relative to Landis & Koch benchmarks)

---

## 📁 Repository Structure

```
├── Assessment_Assignment_1_COSC5926_Completed.xlsx  # Complete evaluation data for ChatGPT, Gemini, Claude (4 evaluators)
├── Assessment_Summary_Completed.xlsx                 # Category-level pass/fail summary matrix across all evaluators
├── Cohens_Kappa_Results.xlsx                         # Statistical inter-rater reliability matrix for all member pairs
├── COSC5926_Assignment1_Report.docx                  # Full academic assessment report (200+ paragraphs, 7 tables, embedded charts)
├── complete_assessment.py                            # Main Python analysis script (data processing, Kappa calculations, plots)
├── generate_report.py                                # Script generating the structured academic report (.docx)
├── Summary_PassFail_Chart.png                        # Pass/Fail distribution chart
├── Comparison_Heatmap.png                            # Cross-application category heatmap
├── Radar_Comparison.png                              # Radar comparative chart
├── Cohens_Kappa_Heatmap.png                          # Inter-rater reliability heatmap
├── Average_Kappa_BarChart.png                        # Average Kappa bar chart
└── README.md                                         # Project documentation
```

---

## 🛠️ Reproduction & Source Code Execution

### Requirements
* Python 3.10+
* Required packages:
  ```bash
  pip install openpyxl pandas numpy matplotlib python-docx
  ```

### Run the Evaluation Pipeline
```bash
python complete_assessment.py
```
This script will:
1. Populate and format `Assessment_Assignment_1_COSC5926_Completed.xlsx`
2. Generate `Assessment_Summary_Completed.xlsx`
3. Compute Cohen's Kappa ($P_o$, $P_e$, and $\kappa$) across all categories and evaluator pairs into `Cohens_Kappa_Results.xlsx`
4. Render all five visualization figures (`.png`)

### Generate Report Document
```bash
python generate_report.py
```
Outputs `COSC5926_Assignment1_Report.docx` with all sections, tables, and figures embedded.
