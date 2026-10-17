import os
import re
import pandas as pd
import streamlit as st
import tkinter as tk
from tkinter import filedialog
import pythoncom

# Try importing pywin32 for Outlook automation on Windows
try:
    import win32com.client as win32
    OUTLOOK_AVAILABLE = True
except ImportError:
    OUTLOOK_AVAILABLE = False


def select_folder():
    """Opens a native desktop folder picker dialog."""
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    folder_path = filedialog.askdirectory(master=root)
    root.destroy()
    return folder_path


def create_outlook_draft(to_email, cc_email, subject, body, attachment_path=None):
    """Creates a draft email in local Microsoft Outlook with thread-safe COM handling."""
    try:
        pythoncom.CoInitialize()
        
        outlook = win32.Dispatch('Outlook.Application')
        mail = outlook.CreateItem(0)  # 0 = olMailItem
        
        mail.To = str(to_email).strip() if pd.notna(to_email) else ""
        if pd.notna(cc_email) and str(cc_email).strip():
            mail.CC = str(cc_email).strip()
            
        mail.Subject = str(subject)
        mail.Body = str(body)
        
        if attachment_path and os.path.exists(attachment_path):
            abs_attachment_path = os.path.abspath(attachment_path)
            mail.Attachments.Add(abs_attachment_path)
            
        mail.Save()  # Saves as draft in Outlook
        return True, "Draft saved successfully"
    except Exception as e:
        return False, str(e)
    finally:
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


# --- Streamlit UI Configuration ---
st.set_page_config(
    page_title="Policy Email Automation Suite",
    page_icon="✉️",
    layout="wide"
)

# Custom Styling
st.markdown("""
    <style>
        .main-header {
            font-size: 1.8rem;
            font-weight: 700;
            color: #1E293B;
            margin-bottom: 0.2rem;
        }
        .sub-header {
            font-size: 0.95rem;
            color: #64748B;
            margin-bottom: 1.5rem;
        }
        .stButton>button {
            width: 100%;
            background-color: #2563EB;
            color: white;
            font-weight: 600;
            border-radius: 8px;
            padding: 0.6rem 1rem;
        }
    </style>
""", unsafe_allow_html=True)


# ==========================================
# 1. LEFT PANEL: SETTINGS & MAPPING CONFIG
# ==========================================
with st.sidebar:
    st.subheader("⚙️ Configuration & Guidelines")
    st.markdown("---")
    
    st.markdown("### 📊 Excel Column Mapping")
    st.info(
        "The application maps columns by their exact sequence position (Columns A to F):\n\n"
        "• **Col A:** User Name\n"
        "• **Col B:** Email Address\n"
        "• **Col C:** Policy Number *(Matching Key)*\n"
        "• **Col D:** Policy Period\n"
        "• **Col E:** To Email Address\n"
        "• **Col F:** CC Email Address"
    )
    
    st.markdown("### 📁 PDF Matching Rule")
    st.warning(
        "The system scans the selected folder and checks if the **Policy Number** (Col C) is contained anywhere inside the PDF filename."
    )
    
    st.markdown("---")
    st.caption("Environment Status:")
    if OUTLOOK_AVAILABLE:
        st.success("🟢 Outlook Integration Active")
    else:
        st.warning("🟡 Simulation Mode (Outlook not detected)")


# ==========================================
# 2. MAIN PANEL: STAGE 1 & STAGE 2
# ==========================================
st.markdown('<div class="main-header">✉️ Automated Policy Email Draft Generator</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">Streamlined draft creation with intelligent file matching and tracking.</div>', unsafe_allow_html=True)

# Default Template Options
PRESET_TEMPLATES = {
    "Policy Document Dispatch": """Dear {User Name},

Please find attached your policy document for Policy Number: {Policy Number} (Period: {Policy Period}).

If you have any questions or require further details, feel free to reach out.

Best regards,
Operations Team""",

    "Policy Renewal Notice": """Dear {User Name},

We are pleased to inform you that your policy ({Policy Number}) document covering {Policy Period} is attached to this email.

Thank you for your continued trust.

Sincerely,
Customer Support Team"""
}

# --- STAGE 1: INPUT FORM & SETUP ---
st.markdown("### 🛠️ Stage 1: Configure Template & Data Source")

with st.form("draft_form"):
    col_t1, col_t2 = st.columns(2)
    
    with col_t1:
        template_choice = st.selectbox(
            "Choose a pre-defined template:",
            options=list(PRESET_TEMPLATES.keys())
        )
        
    with col_t2:
        email_subject = st.text_input(
            "Email Subject Line:",
            value="Policy Document - {Policy Number}"
        )
    
    email_template = st.text_area(
        "Email Body Template (Available Placeholders: {User Name}, {Email Address}, {Policy Number}, {Policy Period}, {To}, {CC}):",
        value=PRESET_TEMPLATES[template_choice],
        height=140
    )
    
    st.divider()
    
    col_u1, col_u2 = st.columns(2)
    
    with col_u1:
        excel_file = st.file_uploader(
            "Upload Recipient Excel Sheet (.xlsx, .xls)",
            type=["xlsx", "xls"]
        )
        
    with col_u2:
        st.write("**Select PDF Storage Folder:**")
        if "folder_path" not in st.session_state:
            st.session_state.folder_path = ""
            
        sub_c1, sub_c2 = st.columns([3, 1])
        with sub_c1:
            st.text_input(
                "Path",
                value=st.session_state.folder_path,
                placeholder="No folder selected",
                disabled=True,
                label_visibility="collapsed"
            )
        with sub_c2:
            if st.form_submit_button("Browse Folder"):
                selected = select_folder()
                if selected:
                    st.session_state.folder_path = selected
                    st.rerun()

    st.divider()
    submit_btn = st.form_submit_button("🚀 Start Processing & Generate Drafts", use_container_width=True)


# --- STAGE 2: PROCESSING STATUS & TOTALS ---
if submit_btn:
    st.markdown("---")
    st.markdown("### 📈 Stage 2: Processing Status & Results")
    
    if not excel_file:
        st.error("⚠️ Please upload an Excel sheet to proceed.")
        st.stop()
        
    if not st.session_state.folder_path or not os.path.exists(st.session_state.folder_path):
        st.error("⚠️ Please select a valid folder containing the PDF files.")
        st.stop()

    try:
        df = pd.read_excel(excel_file)
    except Exception as e:
        st.error(f"Failed to read Excel file: {e}")
        st.stop()

    if len(df.columns) < 6:
        st.error(f"⚠️ Your Excel sheet has {len(df.columns)} columns, but at least 6 columns are required.")
        st.stop()

    c_username = df.columns[0]
    c_email = df.columns[1]
    c_policy_no = df.columns[2]
    c_policy_period = df.columns[3]
    c_to = df.columns[4]
    c_cc = df.columns[5]

    pdf_files = [f for f in os.listdir(st.session_state.folder_path) if f.lower().endswith(".pdf")]
    
    if not pdf_files:
        st.warning("⚠️ No PDF files found in the selected folder.")

    success_count = 0
    skipped_count = 0
    failed_count = 0
    logs = []

    progress_bar = st.progress(0)
    total_rows = len(df)

    for idx, row in df.iterrows():
        policy_no = str(row[c_policy_no]).strip() if pd.notna(row[c_policy_no]) else ""
        user_name = str(row[c_username]).strip() if pd.notna(row[c_username]) else ""
        email_addr = str(row[c_email]).strip() if pd.notna(row[c_email]) else ""
        policy_period = str(row[c_policy_period]).strip() if pd.notna(row[c_policy_period]) else ""
        to_email = row[c_to] if pd.notna(row[c_to]) else email_addr
        cc_email = row[c_cc] if pd.notna(row[c_cc]) else ""

        if not policy_no:
            logs.append({"Row": idx + 1, "Policy No": "N/A", "Status": "Skipped", "Details": "Policy Number missing"})
            skipped_count += 1
            continue

        matched_pdf = None
        for pdf in pdf_files:
            if re.search(re.escape(policy_no), pdf, re.IGNORECASE):
                matched_pdf = os.path.join(st.session_state.folder_path, pdf)
                break

        if not matched_pdf:
            logs.append({"Row": idx + 1, "Policy No": policy_no, "Status": "Skipped", "Details": "Matching PDF not found in folder"})
            skipped_count += 1
            continue

        context = {
            "User Name": user_name,
            "Email Address": email_addr,
            "Policy Number": policy_no,
            "Policy Period": policy_period,
            "To": to_email,
            "CC": cc_email
        }

        formatted_subject = email_subject.format(**context)
        formatted_body = email_template.format(**context)

        if OUTLOOK_AVAILABLE:
            ok, msg = create_outlook_draft(to_email, cc_email, formatted_subject, formatted_body, matched_pdf)
            if ok:
                success_count += 1
                logs.append({"Row": idx + 1, "Policy No": policy_no, "Status": "Success", "Details": f"Draft created ({os.path.basename(matched_pdf)})"})
            else:
                failed_count += 1
                logs.append({"Row": idx + 1, "Policy No": policy_no, "Status": "Failed", "Details": f"Outlook error: {msg}"})
        else:
            success_count += 1
            logs.append({"Row": idx + 1, "Policy No": policy_no, "Status": "Simulated", "Details": "Simulated draft (Pywin32 unavailable)"})

        progress_bar.progress((idx + 1) / total_rows)

    # --- Summary Metrics Display ---
    st.markdown("### 📊 Summary Totals")
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Total Rows Processed", total_rows)
    m2.metric("Drafts Created", success_count, delta="Success")
    m3.metric("Skipped Rows", skipped_count, delta_color="off")
    m4.metric("Failed", failed_count, delta_color="inverse")

    st.markdown("---")
    st.subheader("Detailed Execution Log")
    st.dataframe(pd.DataFrame(logs), use_container_width=True)