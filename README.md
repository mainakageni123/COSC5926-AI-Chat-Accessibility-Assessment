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
| **Google Gemini** | **328** | **44** | **88.2%** | Highest overall score, provides source citation links, excellent cognitive structure | Lacks chat history search and folder organization |
| **Anthropic Claude** | **320** | **52** | **86.0%** | Clear plain language (98% Understandable), uses hedging to disclose uncertainty | Lacks voice input, no suggested prompt starters |
| **OpenAI ChatGPT** | **319** | **53** | **85.8%** | Best conversation management (history search & folders), multimodal voice input | Low source transparency, subtle focus indicators |

### Key Universal Deficiencies Across All Three Platforms:
* **Dynamic Streaming Announcements:** Streaming responses are not announced to screen readers (e.g., NVDA) via ARIA live regions.
* **Audio/Video Accessibility:** Transcripts for audio features and captions for media are largely absent.
* **Cognitive Controls:** None provide notification control settings or native built-in text-to-speech (TTS).
* **AI Uncertainty & Transparency:** Limitations and confidence levels are not systematically communicated across outputs.

---

## 📐 Statistical Reliability & Mathematical Guard for Cohen's Kappa

Inter-rater reliability was assessed across all **180 pairwise evaluations** (3 applications × 10 categories × 6 evaluator pairs) using Cohen's Kappa:

$$\kappa = \frac{P_o - P_e}{1 - P_e}$$

### Mathematical Guard for $P_e = 1.0$:
When both raters agree on 100% of items in a category (e.g., all "Pass"), marginal variance is zero, resulting in $P_o = 1.0$ and $P_e = 1.0$. In this condition:
$$\kappa = \frac{1.0 - 1.0}{1.0 - 1.0} = \frac{0}{0} \implies \text{Mathematically Undefined}$$

Standard naive implementations incorrectly impute $1.0$ (which misrepresents zero variance as perfect correlation) or $0.0$. Our statistical engine rigorously identifies these **39 affected pairwise entries** and designates them as **"N/A" (Undefined: 100% Observed Agreement)**.

### Summary of Inter-Rater Reliability:
* **20 Categories:** Demonstrated **Substantial to Almost Perfect agreement** ($\kappa \ge 0.61$).
* **5 Categories:** Exhibited **100% Unanimous Agreement** across all evaluators (Executive Function across all 3 platforms, Autism in Gemini & Claude) where Kappa is undefined ($N/A$) due to zero marginal variance.
* **5 Categories:** Exhibited lower Kappa due to the **Kappa Paradox (Feinstein & Cicchetti, 1990)**, where high observed agreement (80%–92%) combined with highly skewed marginals mathematically depresses $\kappa$.

---

## 📈 Visualizations & Screenshots

* **Summary Pass/Fail Chart:** `Summary_PassFail_Chart.png` (Pass vs. Fail count breakdown per category)
* **Accessibility Heatmap:** `Comparison_Heatmap.png` (Category pass percentage across ChatGPT, Gemini, and Claude)
* **Radar Profile:** `Radar_Comparison.png` (Multi-axis accessibility profile comparison)
* **Cohen's Kappa Heatmap:** `Cohens_Kappa_Heatmap.png` (Pairwise reliability heatmap with distinct $N/A$ cells in neutral gray)
* **Average Kappa Comparison:** `Average_Kappa_BarChart.png` (Category agreement relative to Landis & Koch benchmarks with hatched $N/A$ bars)
* **Appendix C Screenshots:**
  * `Screenshot_Statistical_Software_Execution.png` (PowerShell console executing statistical calculations)
  * `Screenshot_Kappa_Calculation_Script.png` (VS Code showing 2×2 confusion matrix and $P_e \ge 1.0$ guard)
  * `Screenshot_Excel_Kappa_Results.png` (Microsoft Excel workbook showing `Cohens_Kappa_Results.xlsx`)

---

## 📁 Repository Structure

```
├── Assessment_Assignment_1_COSC5926_Completed.xlsx       # Complete evaluation data for ChatGPT, Gemini, Claude (4 evaluators)
├── Assessment_Summary_Completed.xlsx                      # Category-level pass/fail summary matrix across all evaluators
├── Cohens_Kappa_Results.xlsx                              # Statistical inter-rater reliability matrix (with N/A handling)
├── COSC5926_Assignment1_Report.docx                       # Full academic assessment report (227 paragraphs, 7 tables, 8 figures)
├── complete_assessment.py                                 # Main Python analysis script (Kappa calculation engine, plots, screenshots)
├── generate_report.py                                     # Script generating the structured academic report (.docx)
├── Summary_PassFail_Chart.png                             # Pass/Fail distribution chart
├── Comparison_Heatmap.png                                 # Cross-application category heatmap
├── Radar_Comparison.png                                   # Radar comparative chart
├── Cohens_Kappa_Heatmap.png                               # Inter-rater reliability heatmap (N/A in gray)
├── Average_Kappa_BarChart.png                             # Average Kappa bar chart (N/A hatched)
├── Screenshot_Statistical_Software_Execution.png          # Appendix C: PowerShell console calculation screenshot
├── Screenshot_Kappa_Calculation_Script.png                # Appendix C: VS Code algorithm screenshot
├── Screenshot_Excel_Kappa_Results.png                     # Appendix C: Excel workbook screenshot
└── README.md                                              # Project documentation
```

---

## 🛠️ Reproduction & Source Code Execution

### Requirements
* Python 3.10+
* Required packages:
  ```bash
  pip install openpyxl pandas numpy matplotlib python-docx Pillow
  ```

### Run the Evaluation Pipeline
```bash
python complete_assessment.py
```
This script will:
1. Populate and format `Assessment_Assignment_1_COSC5926_Completed.xlsx`
2. Generate `Assessment_Summary_Completed.xlsx`
3. Compute mathematically guarded Cohen's Kappa ($P_o$, $P_e$, and $\kappa$) across all categories and evaluator pairs into `Cohens_Kappa_Results.xlsx`
4. Render all five visualization figures and high-resolution calculation screenshots (`.png`)

### Generate Report Document
```bash
python generate_report.py
```
Outputs `COSC5926_Assignment1_Report.docx` with all sections, tables, and figures embedded.
