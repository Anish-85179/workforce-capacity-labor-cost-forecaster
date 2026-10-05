# Executive Workforce Capacity & Labor Cost Forecaster

An enterprise-grade Excel operational modeling engine designed to analyze employee shift logs, compute regular vs. overtime payroll liabilities, and monitor workforce capacity utilization across departments.

---

## 📌 Project Overview

Unplanned overtime, scheduling inefficiencies, and manual payroll tracking lead to severe budget leakage for mid-market organizations. This model automates time-card data aggregation, applies multi-condition overtime pricing (1.5x base rate), and dynamically projects total payroll liabilities through an executive dashboard.

* **Target Revenue / Fee Scope:** ₹65,000 – ₹95,000 ($800 – $1,150 USD)
* **Core Excel Architecture:** `XLOOKUP`, `SUMIF`, `MIN`/`MAX` shift boundary calculations, and range-aligned lookup engines.

---

## 📁 Workbook Architecture

The workbook is structured into four dedicated, interconnected sheets:

1. **`Employee_Master`**: Central employee repository containing role classifications, department mappings, standard shift lengths, and base hourly pay rates.
2. **`Shift_Log_Engine`**: Operational timecard log computing standard hours worked vs. overtime variances and tracking departmental absence impacts.
3. **`Labor_Cost_Forecaster`**: Multi-tier payroll aggregation engine calculating regular pay, 1.5x overtime premium costs, and total liability per employee.
4. **`Workforce_Analytics_Dashboard`**: C-suite visual reporting suite featuring high-level KPI cards and departmental cost breakdown tables.

---



## 🛠️ Key Calculations & Excel Logic

### 1. Dynamic Shift Hour Classification (`Shift_Log_Engine`)
To accurately separate standard hours from overtime without manual inputs:
* **Regular Hours:**
  ```excel
  =IF(Absent_Flag="Yes", 0, MIN(Scheduled_Hours, Actual_Hours))
