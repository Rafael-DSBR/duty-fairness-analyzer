# ⚖️ Duty Fairness Analyzer (Roster Audit Tool)

> **"Data Intelligence applied to Workforce Management."**

<div align="center">
  <img src="https://img.shields.io/badge/Python-3.10+-3776AB?style=for-the-badge&logo=python&logoColor=white" />
  <img src="https://img.shields.io/badge/Data-Extraction-FFD43B?style=for-the-badge&logo=pandas&logoColor=blue" />
  <img src="https://img.shields.io/badge/Format-PDF_to_Excel-217346?style=for-the-badge&logo=microsoft-excel&logoColor=white" />
</div>

---

### 📋 The Problem (Diagnosis)
In military and security organizations, duty rosters (scales) are often published in unstructured PDF formats (Bulletins). 
Analyzing the **fairness** of these rosters—ensuring no soldier is overworked or underutilized—is manually impossible when dealing with hundreds of pages and variable naming conventions (e.g., "3rd Sgt Smith" vs "Sgt Smith").

**The Pain Point:** High-ranking officers (Colonels/Commanders) spend hours manually cross-referencing PDFs to detect burnout risks or favoritism.

### 🛠️ The Solution (The Tool)
**Duty Fairness Analyzer** is a Python-based ETL (Extract, Transform, Load) tool that parses military PDF bulletins, normalizes rank/name variations, and generates a statistical Excel report.

**Key Capabilities:**
* ✅ **PDF Scraping:** Extracts duty data from unstructured text using `pdfplumber`.
* ✅ **Name Normalization:** Uses Regex to strip ranks (Gen, Cel, Maj, Sgt) and standardize names for accurate aggregation.
* ✅ **Fairness Algorithm:** Calculates the exact load (% of total shifts) for each individual.
* ✅ **Visual Reporting:** Exports a formatted Excel dashboard with heatmaps and summaries.

---

### 💻 Technical Implementation

The software employs a cleaning pipeline to handle data inconsistencies common in legacy government documents.

**Workflow:**
1.  **Ingestion:** Scans a target folder for all `.pdf` roster files.
2.  **Regex Parsing:** Identifies patterns like *"FOR THE DAY [DATE]"* and *"Duty Officer: [NAME]"*.
3.  **Normalization:** Removes 40+ military rank variations to create unique IDs for personnel.
4.  **Analytics:** Computes total services vs. total days to derive the "Workload Share".

### How to Run

1.  **Install dependencies:**
    ```bash
    pip install -r requirements.txt
    ```

2.  **Execute the GUI:**
    ```bash
    python main.py
    ```

3.  **Operation:**
    * Select the folder containing the PDF Bulletins.
    * Select the destination folder for the Report.
    * Click "Exportar Dados".

---

**Author:** Rafael Cavalheiro
*QA Automation Engineer & Tier 3 Support Lead*
