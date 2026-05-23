# AI DDR Report Generator

AI-powered property inspection and thermal analysis system that automatically generates Detailed Diagnostic Reports (DDR) from inspection PDFs and thermal scan reports.

---

## 🚀 Features

- Upload Inspection + Thermal PDFs
- AI-powered document analysis
- Thermal anomaly interpretation
- Moisture/leakage detection
- Root cause analysis
- Severity assessment
- Automated DDR report generation
- DOCX/PDF export support
- Modern responsive UI
- Streamlit + Netlify deployment support

---

## 🧠 Workflow

```text
Inspection PDF + Thermal PDF
            ↓
PDF Page Extraction
            ↓
AI Analysis (Claude / Gemini)
            ↓
Issue Correlation
            ↓
Severity Detection
            ↓
DDR Report Generation
            ↓
Download Final Report
```

---

## 🛠️ Tech Stack

### Frontend
- HTML
- CSS
- JavaScript

### Backend
- Python
- Streamlit
- Netlify Functions

### AI & Processing
- Claude AI API
- Gemini API
- PDF2Image
- Pillow
- python-docx

---

## 📂 Project Structure

```text
AI-DDR-Report-Generator/
│
├── app.py
├── ddr_generator.py
├── requirements.txt
├── index.html
├── architecture.html
├── netlify.toml
├── README.md
│
├── netlify/
│   └── functions/
│       └── gemini.js
│
├── sample_inputs/
│   ├── Sample_Report.pdf
│   └── Thermal_Images.pdf
│
├── sample_outputs/
│   ├── DDR_Report_Sample_Output.docx
│   └── DDR_Report_Sample_Output.pdf
│
└── assets/
```

---

## ⚙️ Installation

### Clone Repository

```bash
git clone https://github.com/your-username/AI-DDR-Report-Generator.git
cd AI-DDR-Report-Generator
```

### Install Dependencies

```bash
pip install -r requirements.txt
```

---

## 📦 System Dependency

Install Poppler for PDF → image conversion.

### Ubuntu/Debian

```bash
sudo apt-get install poppler-utils
```

### macOS

```bash
brew install poppler
```

### Windows

Download:
https://github.com/oschwartz10612/poppler-windows

---

## 🔑 API Setup

### Windows

```bash
set ANTHROPIC_API_KEY=your_api_key
```

### macOS/Linux

```bash
export ANTHROPIC_API_KEY="your_api_key"
```

---

## ▶️ Run Project

### Streamlit Version

```bash
streamlit run app.py
```

---

## 🌐 Netlify Deployment

Folder structure:

```text
netlify/functions/gemini.js
```

Deploy using Netlify CLI:

```bash
netlify deploy
```

---

## 📄 Input Files

- Property Inspection Reports
- Thermal Imaging PDFs
- Site Observation Reports

---

## 📊 Output Generated

The system generates:

1. Property Issue Summary
2. Area-wise Observations
3. Root Cause Analysis
4. Severity Assessment
5. Recommended Actions
6. Additional Notes
7. Missing Information

---

## 🎯 Use Cases

- Property Inspection Automation
- Thermal Leak Detection
- Structural Audit Assistance
- Building Diagnostics
- Civil Engineering Documentation
- Facility Maintenance Reporting

---

## 🔮 Future Improvements

- OCR optimization
- Multi-language support
- Dashboard analytics
- AI confidence scoring
- Cloud storage integration
- Multi-property batch processing

---

## 👨‍💻 Author

Unnati Lunawat