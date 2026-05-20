import os
import smtplib
import pandas as pd
from datetime import datetime
from dateutil.relativedelta import relativedelta
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

def process_and_alert():
    file_path = "Scheduler_Testing.xlsx"
    
    if not os.path.exists(file_path):
        print(f"Error: {file_path} not found.")
        return

    # Ingest spreadsheet data
    df = pd.read_excel(file_path, engine='openpyxl')
    
    # Ensure specific column formatting is clean
    df['DepreciationDate'] = pd.to_datetime(df['DepreciationDate']).dt.date
    df['Notified'] = df['Notified'].fillna('').astype(str).str.strip()
    df['RecipientEmail'] = df['RecipientEmail'].fillna('').astype(str).str.strip()

    # Precise calendar matching (exactly 3 months from today)
    today = datetime.now().date()
    target_date = today + relativedelta(months=+3)
    
    # Filter for target dates where alerts haven't been sent yet
    critical_records = df[
        (df['DepreciationDate'] == target_date) & 
        (df['Notified'] != 'Yes') &
        (df['RecipientEmail'] != '')
    ]
    
    if critical_records.empty:
        print("No matching LLM deprecations scheduled for exactly 3 months from today.")
        return

    # SMTP configuration loaded from GitHub Environment Secrets
    smtp_server = os.getenv("SMTP_SERVER", "smtp.gmail.com")
    smtp_port = int(os.getenv("SMTP_PORT", "465"))
    sender_email = os.getenv("SENDER_EMAIL")
    sender_password = os.getenv("SENDER_PASSWORD") # App-specific password

    if not sender_email or not sender_password:
        print("SMTP credentials are not configured in environment variables.")
        return

    # Group alerts by email recipient to avoid spamming individuals
    grouped = critical_records.groupby('RecipientEmail')
    
    for email_recipient, group in grouped:
        # Convert matching records for this user into a clean HTML table
        display_cols = ['Model', 'SubModule', 'DepreciationDate']
        table_html = group[display_cols].to_html(index=False, classes='styled-table')
        
        # Build styled HTML email headers correctly
        msg = MIMEMultipart('alternative')
        msg['Subject'] = "Action Required: LLM Deprecation Warning (3 Months Notice)"
        msg['From'] = sender_email
        msg['To'] = email_recipient
        
        body_content = f"""
        <html>
        <head>
            <style>
                body {{ font-family: 'Segoe UI', Arial, sans-serif; background-color: #f4f6f9; color: #333; padding: 20px; }}
               .container {{ max-width: 600px; background: #ffffff; border-radius: 8px; border-top: 4px solid #d9534f; box-shadow: 0 4px 10px rgba(0,0,0,0.05); margin: auto; padding: 30px; }}
                h2 {{ color: #d9534f; font-size: 20px; margin-top: 0; }}
                p {{ font-size: 14px; line-height: 1.6; color: #555555; }}
               .styled-table {{ width: 100%; border-collapse: collapse; margin: 20px 0; font-size: 14px; }}
               .styled-table th,.styled-table td {{ padding: 12px; text-align: left; border-bottom: 1px solid #e1e1e1; }}
               .styled-table th {{ background-color: #f8f9fa; font-weight: 600; color: #2c3e50; }}
               .action-banner {{ background-color: #fdf7f7; border-left: 4px solid #d9534f; padding: 15px; border-radius: 4px; font-weight: bold; margin-top: 20px; font-size: 13px; color: #b94a48; }}
            </style>
        </head>
        <body>
            <div class="container">
                <h2>Urgent: Large Language Model Deprecation Warning</h2>
                <p>The automated monitoring system has detected that the following LLMs integrated into your sub-modules are scheduled for retirement in exactly <strong>3 months</strong> ({target_date.strftime('%B %d, %Y')}):</p>
                {table_html}
                <div class="action-banner">
                    Action Required: Please plan transition paths, perform regression tests, and update configurations to avoid endpoint failures.
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
            print(f"Alert successfully sent to: {email_recipient}")
            
            # Update original DataFrame rows state using indices from the processed group
            df.loc[group.index, 'Notified'] = 'Yes'
        except Exception as e:
            print(f"Failed to send email to {email_recipient}: {str(e)}")

    # Write updated state back to Excel
    df.to_excel(file_path, index=False, engine='openpyxl')
    print("Execution states saved back to tracking spreadsheet.")

if __name__ == "__main__":
    process_and_alert()