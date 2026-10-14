# SendGrid Master Deliverability & Metrics Reporter

> An automated Python CLI tool for auditing email delivery performance, analyzing campaign metrics, and exporting professional, presentation-ready Excel reports with embedded charts and executive dashboards.

---

## 🚀 Key Features

* **Interactive CLI Date Selector:** Instantly choose from quick presets (**Last 7 days**, **Last 30 days**, **This month**) or specify a custom date range with built-in validation.
* **Dual-Engine Data Architecture:**
  * **Global Stats API (`/v3/stats`):** Guarantees macro-level historical coverage (delivered counts, bounces, blocks, deferred metrics) even beyond standard API retention windows.
  * **Messages API & Event Logs (`/v3/messages`):** Surfaces granular sender, recipient, status, and failure remarks when querying within SendGrid's active retention window.
* **Executive Excel Dashboards (`openpyxl`):**
  * Styled KPI summary blocks and color-coded status highlights.
  * Embedded native **Line Charts** (Daily request volume) and **Bar Charts** (Status distribution).
  * Auto-fitted columns, gridline suppression, and professional corporate color grading (`#17365D` Navy theme).
* **Resilient Error Handling:** Automatically handles locked Excel files (generating timestamped fallbacks if a file is open in Microsoft Excel) and API retention boundaries.

---

## 📂 Project Structure

```text
sendgrid-deliverability-reporter/
├── sendgrid_master_report.py   # Main Python automation script
├── reports/                    # Generated Excel reports output directory
└── README.md                   # Project documentation
```

---

## 📋 Prerequisites

* **Python:** Version `3.8` or higher installed on your system.
* **SendGrid Account:** An active SendGrid account with an API Key possessing read permissions for **Stats** and **Mail Send / Messages**.

---

## 🛠️ Step-by-Step Installation & Setup

### Step 1: Clone or Create the Project Directory
Open your terminal and create your project folder:
```bash
mkdir sendgrid-deliverability-reporter
cd sendgrid-deliverability-reporter
```

### Step 2: Create a Python Virtual Environment (Recommended)
```bash
python -m venv venv

# Activate the virtual environment:
# On Linux/macOS:
source venv/bin/activate

# On Windows (PowerShell):
venv\Scripts\Activate.ps1
```

### Step 3: Install Required Dependencies
Install the required packages (`requests`, `pandas`, and `openpyxl`):
```bash
pip install requests pandas openpyxl
```

### Step 4: Configure Your SendGrid API Key
The script requires your SendGrid API key to be stored securely as an environment variable.

* **Linux / macOS (Bash/Zsh):**
  ```bash
  export SENDGRID_API_KEY="your_sendgrid_api_key_here"
  ```
* **Windows (PowerShell):**
  ```powershell
  $env:SENDGRID_API_KEY="your_sendgrid_api_key_here"
  ```
* **Windows (Command Prompt):**
  ```cmd
  set SENDGRID_API_KEY=your_sendgrid_api_key_here
  ```

---

## 💻 Usage Instructions

Run the master script from your terminal:

```bash
python sendgrid_master_report.py
```

### Interactive Menu Walkthrough
Upon execution, the script will present an interactive date selection menu:

```text
=== SendGrid Master Deliverability Reporter ===

--- Select Reporting Period ---
1. Last 7 days (Recommended for individual message logs)
2. Last 30 days
3. This month
4. Custom date range (YYYY-MM-DD)

Choose an option (1-4) [1]: 
```

* Hit **Enter** to select the default option (`1`), or input a number (`1`–`4`).
* If option `4` is chosen, you will be prompted to input custom start and end dates in `YYYY-MM-DD` format.

Once completed, the script compiles the data, builds the workbook, and saves it inside the `reports/` folder (e.g., `reports/sendgrid_master_report_2026-10-01_2026-10-14.xlsx`).

---

## 📊 Understanding Data Retention & Architecture

* **Macro Metrics & Charts:** Powered by the SendGrid **Global Stats API (`/v3/stats`)**, allowing you to pull historical performance across weeks or months.
* **Granular Message Details:** Individual recipient addresses, sender details, and failure remarks are pulled via the **Messages API (`/v3/messages`)**. 
  > **Note on Retention:** SendGrid standard plans retain individual message activity logs for **3 to 7 days**. Queries outside this window will correctly populate macro totals and stats charts while indicating retention limits on individual log tabs.

---

## 🔧 Troubleshooting & Common Issues

### 1. `PermissionError: [Errno 13] Permission denied`
* **Cause:** The target Excel file is currently open in Microsoft Excel, locking the file from being overwritten.
* **Solution:** Close the open Excel workbook and re-run the script. (The script also includes an automatic fallback that appends a timestamp if a lock is detected).

### 2. "No message-level records available" in Email Details / Failed Messages
* **Cause:** You queried a historical date range older than SendGrid's 3–7 day message retention window.
* **Solution:** Use a recent date range (Option 1) if you need granular recipient logs, or rely on the `Daily Stats` and `Status Summary` tabs for long-term historical audits.

---

## 📄 License

This project is open-source and available under the [MIT License](LICENSE).
