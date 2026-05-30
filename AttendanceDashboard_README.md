# 📅 Attendance Analysis System

A web-based tool built with Streamlit that automates the processing of employee attendance data exported from biometric systems.

Built to replace manual HR calculations with an automated, auditable, and fully configurable process — no code changes required to adjust attendance policy.

---

## 🌐 Live Demo

Upload your own attendance Excel file to explore the full workflow.

---

## 📌 Background

Built to automate monthly attendance reporting across multiple directorates. The tool processes raw biometric export data and produces a structured report ready for HR and payroll use.

Dashboard interface and core logic developed with AI-assisted coding (Claude by Anthropic), based on operational defined and specified by the author.

---

## ⚙️ What It Does

Upload a monthly attendance Excel file exported from any standard biometric system, configure your shift rules, and export a formatted report — all in a few clicks.

### Core Calculations

| Metric | Description |
|--------|-------------|
| **Tardiness** | Configurable grace period, rounding rules, and base time |
| **Early Departure** | Detected against expected shift end time |
| **Net Overtime** | Calculated with optional tardiness deduction |
| **Absences** | Classified as justified or unjustified |

### Exceptions Handling
A separate Excel file allows marking specific employee-days as justified — absence, tardiness, or early departure — so they are excluded from penalties in the final report.

---

## 📊 Output Report

A clean Excel file with four sheets:

1. **Employee Summary** — monthly totals per employee
2. **Daily Detail** — day-by-day breakdown per employee
3. **Monthly / Department Statistics** — aggregated by department
4. **Configuration** — the shift rules used to generate the report (for auditability)

---

## 🚀 How to Run

```bash
# 1. Clone the repository
git clone https://github.com/Nedalsa/AttendanceDashboard.git
cd AttendanceDashboard

# 2. Install dependencies
pip install -r requirements.txt

# 3. Launch the app
streamlit run app.py
```

Then open your browser at `http://localhost:8501`

---

## 📁 Repository Structure

```
AttendanceDashboard/
│
├── app.py            # Streamlit UI and main application flow
├── processor.py      # Core attendance calculation logic
├── exporter.py       # Excel report generation
├── requirements.txt  # Python dependencies
└── README.md         # This file
```

---

## 🎯 Goals

- Replace manual HR calculations with an automated, auditable process
- Make attendance policy fully configurable without touching code
- Support any organization using a standard biometric attendance system

---

## 👤 Author

**Nidal Al-Saqqa**
[linkedin.com/in/nidal-al-saqqa](https://linkedin.com/in/nidal-al-saqqa)
