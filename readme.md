# ✉️ Policy Email Automation Suite

An automated, web-based desktop application built with **Streamlit** and **Microsoft Outlook COM Automation (`pywin32`)**. This suite allows operations and support teams to process bulk recipient data from Excel, automatically match corresponding PDF policy documents from a local directory, and save formatted email drafts directly into Microsoft Outlook.

---

## 🌟 Key Features

* **Intelligent File Matching:** Automatically searches a designated local directory for PDF files containing matching **Policy Numbers**.
* **Flexible Email Content:** Supports customized per-row email content provided directly within the Excel file (Column G), with fallback template support.
* **Thread-Safe Outlook Integration:** Uses `pythoncom` initialization to safely communicate with local desktop Microsoft Outlook and save emails directly into your Drafts folder.
* **Live Execution Tracking:** Stage 2 execution includes real-time progress indicators, status metrics, and a detailed audit log dataframe.
* **Native Desktop Folder Selection:** Integrated Tkinter file dialog for intuitive directory selection directly within the Streamlit UI.

---

## 📊 Excel File Requirements

The application maps columns sequentially by their index position (Columns A through G):

| Column | Position | Description | Required / Optional |
| :--- | :--- | :--- | :--- |
| **Col A** | Index 0 | User / Recipient Name | Required |
| **Col B** | Index 1 | Email Address | Required |
| **Col C** | Index 2 | **Policy Number** *(PDF Matching Key)* | Required |
| **Col D** | Index 3 | Policy Period | Required |
| **Col E** | Index 4 | To Email Address *(Defaults to Col B if empty)* | Required |
| **Col F** | Index 5 | CC Email Address | Optional |
| **Col G** | Index 6 | Custom Email Body Content | Optional |

---

## 🚀 Getting Started

### Prerequisites

* **Operating System:** Windows (required for Microsoft Outlook desktop integration)
* **Python:** Python 3.8 or higher
* **Microsoft Outlook:** Local desktop client installed and logged in

### Installation

1. **Clone or Download the Repository:**
   ```powershell
   git clone <repository-url>
   cd <repository-folder>
   ```

2. **Set Up a Virtual Environment:**
   ```powershell
   python -m venv .venv
   .venv\Scripts\Activate.ps1
   ```

3. **Install Dependencies:**
   ```powershell
   pip install streamlit pandas pywin32 openpyxl
   ```

---

## 🏃 Usage

1. **Launch the Application:**
   Run the Streamlit server from your terminal:
   ```powershell
   streamlit run app.py
   ```

2. **Stage 1: Configuration & Inputs:**
   * Set your desired **Email Subject Line**.
   * Optionally set a default **Fallback Email Body** (used if Column G is empty or absent).
   * Upload your recipient Excel file (`.xlsx` or `.xls`).
   * Click **Browse Folder** to select the directory containing your policy PDFs.
   * Click **🚀 Start Processing & Generate Drafts**.

3. **Stage 2: Execution & Output:**
   * Monitor the live progress bar as each row is processed.
   * View generated drafts in your local **Microsoft Outlook > Drafts** folder.
   * Inspect summary totals (Processed, Drafts Created, Skipped, Failed) and review the detailed execution table.

---

## ⚠️ Troubleshooting & Common Issues

### COM / Outlook Error `(-2147352567, 'Exception occurred.', ...)`
* **Cause:** Privilege mismatch between your terminal/python process and Microsoft Outlook (e.g., one running as Administrator and the other as a standard user).
* **Resolution:**
  1. Close Microsoft Outlook completely.
  2. Close your current PowerShell terminal.
  3. Open a fresh terminal **without** Administrator privileges.
  4. Launch Streamlit and open Outlook as a standard user.

---

## 🛠️ Built With

* [Streamlit](https://streamlit.io/) - Web UI framework
* [Pandas](https://pandas.pydata.org/) - Excel parsing and data handling
* [PyWin32](https://github.com/mhammond/pywin32) - Windows COM Client for Outlook automation