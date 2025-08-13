# Excel vs Python: Data Analysis Comparison

This project compares **Excel** and **Python** workflows for analyzing the same dataset, with a focus on functionality, ease of use, and output quality.  
It contains both an **Excel dashboard** and equivalent **Python-based analysis** so you can see how each tool handles the same tasks.

## Project Structure

```

- `data/` — Raw and processed datasets  
- `excel_dashboard/` — Excel dashboards and sheets  
- `images/` — Plots and charts used in README and docs  
- `notebooks/` — Jupyter notebooks for exploration & analysis  
- `python_analysis/` — Standalone Python scripts  
- `.gitignore` — Files & folders to be ignored by Git  
- `LICENSE` — License for reuse  
- `README.md` — Project overview (this file)  
- `requirements.txt` — Python dependencies


````

---

## How to Reproduce the Code

### 1. Clone this repository
```bash
git clone https://github.com/sakshikharkwal/Excel-vs-Python-Analysis.git
cd Excel-vs-Python-Analysis
````

### 2. Create and activate a virtual environment

```bash
python -m venv .venv
```

**Activate:**

* **Windows (PowerShell)**

  ```bash
  .venv\Scripts\Activate.ps1
  ```
* **Windows (Git Bash)**

  ```bash
  source .venv/Scripts/activate
  ```
* **macOS/Linux**

  ```bash
  source .venv/bin/activate
  ```

### 3. Install dependencies

```bash
pip install -r requirements.txt
```

### 4. Explore the project

* **Excel Dashboard** → Open files in `excel_dashboard/` using Excel
* **Python Scripts** → Run `.py` files inside `python_analysis/`
* **Notebooks** → Launch:

  ```bash
  jupyter notebook
  ```

---

## Observations & Findings

### Excel-specific challenges:

1. **YEAR() bug with text-formatted dates:**
   If your date is stored as text (e.g., `08/11/2016`), Excel’s `YEAR()` fails because it tries to interpret `"2016"` as a serial date number (2016 days after 1 Jan 1900 → 08 Jul 1905).
   `TEXT([@[Order Date]],"yyyy")` works because it just extracts the year from the string.
   **Temporary fix used:** Convert column type to **Number** instead of **Date**.
   **Limitation:** In slicers, there is no way to change the data type.

2. **Excel Online limitations:**

   * Cannot add slicers to regular tables (only available in Pivot Tables).
   * Reduced functionality in the free Excel Online version compared to desktop Excel.

---

### Python advantages:

1. Jupyter Notebook (and Python in general) has **no feature restrictions** in the free version.
2. More **automation and reproducibility** — the same analysis can run on any dataset without manual rework.
3. Libraries like **pandas**, **matplotlib**, and **seaborn** provide more control over transformations and visuals.

```
