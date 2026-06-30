# 💰 AI-Powered Engineering Parts Cost Estimation System


An intelligent automation platform that extracts machine and component names from engineering construction drawings using OCR, performs real-time online price discovery through internet search APIs, and automatically generates structured Excel-based cost estimation reports.

---

# 📖 About the Project

Engineering and construction drawings often contain hundreds of machine names, component labels, equipment references, and part identifiers. Manually identifying these components and gathering market pricing information is a labor-intensive process.

This project automates the complete workflow by:

* Detecting machine names from engineering drawings
* Extracting all textual information using OCR
* Identifying equipment and component references
* Performing live internet searches for pricing information
* Aggregating part names and estimated market prices
* Generating structured Excel reports automatically

The system significantly reduces the time required for cost estimation, procurement analysis, and engineering documentation workflows.

---

# 🧩 Key Features

✅ OCR-Based Text Detection

✅ Machine & Component Identification

✅ Engineering Drawing Text Extraction

✅ Automated Part Name Recognition

✅ Live Internet Price Search

✅ DDGS Search Integration

✅ Automated Cost Estimation

✅ Excel Report Generation

---

# 🛠️ Tech Stack

## OCR & Text Extraction

* EasyOCR
* Tesseract OCR (Optional)

## Computer Vision

* OpenCV
* NumPy

## Internet Search & Data Retrieval

* DDGS (DuckDuckGo Search)
* Search APIs

## Data Processing

* Pandas
* Regex
* Python Collections

## Excel Automation

* OpenPyXL
* Pandas Excel Writer

## Development Tools

* Python
* Jupyter Notebook
* VS Code

---

# 🧠 Learning Outcomes

Through this project, I gained practical experience in:

* OCR Pipeline Development
* Engineering Drawing Analysis
* Text Mining
* Data Extraction Automation
* API Integration
* Live Data Retrieval
* Cost Estimation Workflows
* Excel Automation
* Data Cleaning & Validation
* End-to-End Automation Development

---

# 🧰 Installation & Setup

## Clone Repository

```bash
git clone https://github.com/noobnetizen/Engineering-Parts-Cost-Estimation-System.git

cd Engineering-Parts-Cost-Estimation-System
```

## Create Virtual Environment

```bash
python -m venv venv

venv\Scripts\activate
```

## Install Dependencies

```bash
pip install -r requirements.txt
```

## Run Application

```bash
python main.py
```

---

# 📁 Project Structure

```text
Engineering-Parts-Cost-Estimation-System
│
├── input_drawings
│   ├── drawing_1.pdf
│   ├── drawing_2.pdf
│   └── ...
│
├── extracted_text
│   ├── parts_list.csv
│   └── extracted_labels.csv
│
├── price_search_results
│   ├── part_prices.xlsx
│   └── search_results.csv
│
├── src
│   ├── pdf_processor.py
│   ├── ocr_engine.py
│   ├── text_extractor.py
│   ├── component_identifier.py
│   ├── price_search.py
│   ├── excel_exporter.py
│   └── utils.py
│
├── assets
│   ├── workflow.png
│   ├── ocr_detection.png
│   ├── search_results.png
│   ├── excel_output.png
│   └── dashboard.png
│
├── requirements.txt
├── README.md
├── LICENSE
└── .gitignore
```

---

# 🔄 Workflow

```text
Construction Drawing PDF
            │
            ▼
OCR Text Detection
            │
            ▼
Machine & Part Identification
            │
            ▼
Component Name Extraction
            │
            ▼
Live Internet Search
            │
            ▼
Price Discovery
            │
            ▼
Data Validation
            │
            ▼
Excel Report Generation
```

---

# 📊 Output

The system generates structured cost estimation reports containing:

| Part Name        | Estimated Price |
| ---------------- | --------------- |
| Hydraulic Pump   | ₹15,000         |
| Bearing Assembly | ₹1,250          |
| Gear Motor       | ₹8,750          |

Outputs include:

* Extracted Part Lists
* Price Lookup Results
* Procurement Reports
* Excel-Based Cost Estimation Sheets

---

# 📈 Business Applications

The solution can be applied in:

* Construction Project Estimation
* Procurement Planning
* Asset Management
* Engineering Documentation
* Equipment Cost Analysis
* Maintenance Planning
* Inventory Assessment

---

# 🙌 Contributions

Contributions, issues, and feature requests are welcome! Feel free to fork this repo and submit a pull request.

---

# 🧑‍💻 Developed by

R.Krishnan

Machine Learning & Data Analytics Enthusiast

🔗 GitHub: https://github.com/noobnetizen

🔗 LinkedIn: https://www.linkedin.com/in/krishnan-r-8697771b6/

---

⭐ If you found this project useful, consider giving it a star on GitHub.
