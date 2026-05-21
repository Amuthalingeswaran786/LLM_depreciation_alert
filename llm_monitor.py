import os
import smtplib
import pandas as pd
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

def process_and_alert():
    file_path = "Scheduler_Testing.xlsx"
    
    if not os.path.exists(file_path):
        print(f"Error: {file_path} not found.")
        return

    # 1. Ingest spreadsheet data
    df = pd.read_excel(file_path, engine='openpyxl')
    
    # Clean up column names in the file (remove accidental spaces)
    df.columns = df.columns.str.strip()
    
    # DYNAMIC COLUMN MAPPER: Finds columns regardless of spaces, hyphens, or casing
    column_mapping = {}  # FIX 1: was being overwritten as a string instead of dict entries
    for col in df.columns:
        clean_name = col.lower().replace(" ", "").replace("-", "").replace("_", "")
        if clean_name in ["model", "llmmodel", "llm"]:
            column_mapping["Model"] = col
        elif clean_name in ["submodule", "sub_module"]:
            column_mapping["SubModule"] = col  # FIX 1: was `column_mapping = col`
        elif clean_name in ["depreciationdate", "depreciation_date", "expirydate", "expirationdate"]:
            column_mapping["DepreciationDate"] = col  # FIX 1: was `column_mapping = col`
        elif clean_name in ["recipientemail", "recipient_email", "email", "recipients"]:
            column_mapping["RecipientEmail"] = col  # FIX 1: was `column_mapping = col`

    # Verify that we matched all 4 required fields
    required_keys = ["Model", "SubModule", "DepreciationDate", "RecipientEmail"]  # FIX 2: was never defined
    missing_keys = [key for key in required_keys if key not in column_mapping]
    
    if missing_keys:
        print(f"Error: Could not identify columns for: {missing_keys}")
        print(f"Columns present in your Excel file are: {list(df.columns)}")
        print("Please ensure your Excel file contains columns representing: Model, SubModule, DepreciationDate, and RecipientEmail.")
        return

    # Rename the mapped columns to standard names for consistent coding
    rename_dict = {column_mapping[key]: key for key in required_keys}
    df = df.rename(columns=rename_dict)

    # 2. Clean and format the data columns safely
    df["Model"] = df["Model"].fillna("").astype(str).str.strip()
    df["SubModule"] = df["SubModule"].fillna("").astype(str).str.strip()        # FIX 3: was bare `df = df.fillna(...)`
    df["RecipientEmail"] = df["RecipientEmail"].fillna("").astype(str).str.strip()  # FIX 3: same issue
    
    # Safe date conversion: Parse values and convert invalid formats to NaT (Not a Time)
    df["DepreciationDate"] = pd.to_datetime(df["DepreciationDate"], errors='coerce')  # FIX 4: was applied to whole df
    
    # Filter out rows with unparseable dates or empty email addresses
    # valid_mask = df["DepreciationDate"].notna() & (df["RecipientEmail"] != "")  # FIX 5: broken syntax
    # df = df[valid_mask]
    
    # Filter out rows with unparseable dates or empty email addresses
    valid_mask = df.notna() & (df!= "")
    df = df[valid_mask].copy()  # <--- ADD.copy() HERE

    if df.empty:
        print("No valid rows containing both a DepreciationDate and RecipientEmail were found.")
        return

    # 3. Calculate remaining days mathematically
    today = pd.Timestamp.now().normalize()
    df["DaysRemaining"] = (df["DepreciationDate"] - today).dt.days  # FIX 6: was bare `df = (...).dt.days`
    
    # 4. FILTER: Select rows expiring in 90 days or less (including expired ones < 0)
    critical_records = df[df["DaysRemaining"] <= 90].copy()  # FIX 7: was `df <= 90]` (broken bracket)
    
    if critical_records.empty:
        print("Excellent! No LLM deprecations are scheduled within the next 90 days.")
        return

    # Sort records: most urgent first
    critical_records = critical_records.sort_values(by="DaysRemaining", ascending=True)

    # SMTP configuration loaded from GitHub Environment Secrets
    smtp_server = os.getenv("SMTP_SERVER", "smtp.gmail.com")
    smtp_port = int(os.getenv("SMTP_PORT", "465"))
    sender_email = os.getenv("SENDER_EMAIL")
    sender_password = os.getenv("SENDER_PASSWORD")  # App-specific password

    if not sender_email or not sender_password:
        print("SMTP credentials are not configured in your GitHub Environment Secrets.")
        return

    # Group alerts by email recipient to send one consolidated summary email
    grouped = critical_records.groupby("RecipientEmail")
    
    for email_recipient, group in grouped:
        email_group = group.copy()
        
        # Create a user-friendly status countdown column
        def get_status_label(days):
            if days < 0:
                return f"⚠️ EXPIRED ({abs(days)} days ago)"
            elif days == 0:
                return "🚨 Expiring TODAY"
            else:
                return f"{days} days left"
                
        email_group["Status"] = email_group["DaysRemaining"].apply(get_status_label)  # FIX 10: was bare `email_group = email_group.apply(...)`
        
        # Format dates visually for the email table (e.g., "Aug 20, 2026")
        email_group["DepreciationDate"] = email_group["DepreciationDate"].dt.strftime('%b %d, %Y')  # FIX 11: was bare `email_group = email_group.dt.strftime(...)`
        
        # Define the exact columns to print inside the HTML email table
        display_cols = ["Model", "SubModule", "DepreciationDate", "Status"]  # FIX 9: was never defined
        
        # Convert columns into styled HTML
        table_html = email_group[display_cols].to_html(index=False, classes='styled-table')
        
        # Build styled HTML email
        msg = MIMEMultipart('alternative')
        msg["Subject"] = "🚨 URGENT: Daily LLM Deprecation Warning (<90 Days Left)"  # FIX 8: was `msg = ...`
        msg['From'] = sender_email
        msg["To"] = email_recipient  # FIX 8: was `msg = email_recipient`
        
        body_content = f"""
        <html>
        <head>
            <style>
                body {{ font-family: 'Segoe UI', Arial, sans-serif; background-color: #f4f6f9; color: #333; padding: 20px; }}
              .container {{ max-width: 650px; background: #ffffff; border-radius: 8px; border-top: 5px solid #d9534f; box-shadow: 0 4px 10px rgba(0,0,0,0.05); margin: auto; padding: 30px; }}
                h2 {{ color: #d9534f; font-size: 20px; margin-top: 0; }}
                p {{ font-size: 14px; line-height: 1.6; color: #555555; }}
              .styled-table {{ width: 100%; border-collapse: collapse; margin: 20px 0; font-size: 14px; }}
              .styled-table th,.styled-table td {{ padding: 12px; text-align: left; border-bottom: 1px solid #e1e1e1; }}
              .styled-table th {{ background-color: #f8f9fa; font-weight: 600; color: #2c3e50; }}
              .styled-table tr:hover {{ background-color: #f1f1f1; }}
              .action-banner {{ background-color: #fdf7f7; border-left: 4px solid #d9534f; padding: 15px; border-radius: 4px; font-weight: bold; margin-top: 20px; font-size: 13px; color: #b94a48; }}
            </style>
        </head>
        <body>
            <div class="container">
                <h2>LLM Service Deprecation Digest</h2>
                <p>This is your daily automated reminder that the following Large Language Models in active sub-modules have <strong>90 days or less</strong> before deprecation:</p>
                {table_html}
                <div class="action-banner">
                    Action Required: Please transition these systems to a modern replacement model and update the tracking sheet to stop receiving this daily alert.
                </div>
            </div>
        </body>
        </html>
        """
        msg.attach(MIMEText(body_content, 'html'))
        
        # Connect securely and dispatch the email
        try:
            with smtplib.SMTP_SSL(smtp_server, smtp_port) as server:
                server.login(sender_email, sender_password)
                server.sendmail(sender_email, email_recipient, msg.as_string())
            print(f"Daily digest successfully sent to: {email_recipient}")
        except Exception as e:
            print(f"Failed to send email to {email_recipient}: {str(e)}")

if __name__ == "__main__":
    process_and_alert()
