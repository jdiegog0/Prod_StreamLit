# 📊 Productivity Analysis Dashboard

An interactive Streamlit application for tracking performance through two distinct productivity models: **Effort** (Efficiency-focused) and **Points** (Output-focused).

## 🚀 Key Features
* **Automated Mapping:** Intelligent column detection for uploaded Excel files.
* **Dual Analysis Models:**
    * **Less is Best (Effort):** Analyzes workload and time expenditure.
    * **More is Best (Points):** Analyzes delivery and story points.
* **Moving Baseline:** Calculates productivity by comparing current windows against historical performance.

## 🛠️ Installation
```bash
pip install streamlit pandas plotly openpyxl
```
### or

```bash
pip install -r requirements.txt
```

## 📈 Calculation Logic
The engine calculates productivity by establishing an **Effort per Unit (EpU)** ratio from a baseline period to determine expected values.

### 1. Baseline Performance (EpU_BL)
**Formula:** EpU_BL = (Sum of Baseline Value) / (Sum of Baseline Units)

### 2. Expected Value (V_exp)
**Formula:** V_exp = EpU_BL * (Sum of Current Units)

### 3. Productivity Index (P)
**Formula:** P = ((Sum of Current Value - V_exp) / V_exp) * σ

* **Note:** σ = 1 for "More is Best" and σ = -1 for "Less is Best".

## 📋 Data Requirements
The system expects an Excel file (Sheet: RawData or first sheet) with the following detectable columns:
* **Effort Model:** Assigned To, Group, EndDate, Effort.
* **Points Model:** Points, Developer, Status, Period.

## 💻 Installation and Usage
To run the application, ensure you have the following libraries installed: **streamlit**, **pandas**, **plotly** and **openpyxl**. Use the command `streamlit run` followed by your script name to launch the dashboard.

```bash
streamlit run main.py
```