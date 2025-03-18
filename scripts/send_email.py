import smtplib
import os
import mimetypes
from email.message import EmailMessage
from dotenv import load_dotenv
import yagmail


# Load environment variables from .env file
load_dotenv()

SENDER_EMAIL = os.getenv("EMAIL_SENDER")
SENDER_PASSWORD = os.getenv("EMAIL_PASSWORD")

# 📌 Email Configuration
SMTP_SERVER = "smtp.gmail.com"  # Update for your provider
SMTP_PORT = 587
SENDER_EMAIL = os.getenv("EMAIL_SENDER")  # Store in environment variables
SENDER_PASSWORD = os.getenv("EMAIL_PASSWORD")  # Store in environment variables
RECIPIENT_EMAIL = "meos.r7@gmail.com"  # Change to desired recipient

# 📌 Attachments (Reports)
REPORT_FILE = r"C:\Users\Emilios Richards\VSCode\Manuav Mock File Automation\reports\Merged_Automation_Report_November.xlsx"
ERROR_LOG_FILE = r"C:\Users\Emilios Richards\VSCode\Manuav Mock File Automation\reports\Error_Summary_Report.txt"

# 📌 Function to Send Email
def send_email():
    """Sends an automated email with the merged report and error log."""

    msg = EmailMessage()
    msg["Subject"] = "📊 Manuav Automation Report - Data Processing Completed"
    msg["From"] = SENDER_EMAIL
    msg["To"] = RECIPIENT_EMAIL
    msg.set_content("""
    Hello,

    The automated data processing for Manuav has been completed. 

    📌 Attached:
    - Merged Data Report (Excel)
    - Error Summary Report (TXT)

    🔍 Key Findings:
    - Missing data entries were found and flagged in red.
    - Time mismatches detected and highlighted.
    - Possible training days (Over 4 hours) flagged in blue.
    - Non-numeric values detected in call logs (if any).

    Please review the attached reports and take necessary actions.

    Best regards,  
    Automation Script
    """)

    # 📌 Attach the Excel Report
    for file_path in [REPORT_FILE, ERROR_LOG_FILE]:
        if os.path.exists(file_path):
            with open(file_path, "rb") as file:
                file_type, _ = mimetypes.guess_type(file_path)
                file_type = file_type or "application/octet-stream"
                file_name = os.path.basename(file_path)
                msg.add_attachment(file.read(), maintype=file_type.split("/")[0], 
                                   subtype=file_type.split("/")[1], filename=file_name)

    # 📌 Connect & Send Email
    try:
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()  # Secure connection
            server.login(SENDER_EMAIL, SENDER_PASSWORD)
            server.send_message(msg)

        print("✅ Email sent successfully to:", RECIPIENT_EMAIL)

    except Exception as e:
        print(f"❌ Failed to send email: {e}")

# 📌 Run the email function if script is executed
if __name__ == "__main__":
    send_email()
