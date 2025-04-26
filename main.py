# main.py
import os
import time
import re
import fitz  # PyMuPDF
import smtplib
from email.message import EmailMessage
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

# --- Configuration ---
WATCH_FOLDER = "./confirmations"
SMTP_SERVER = "smtp.office365.com"
SMTP_PORT = 587
SMTP_USER = "athdesk@caltech.edu"
SMTP_PASSWORD = os.getenv("SMTP_PASSWORD")  # Set in .env file or environment

EMAIL_SUBJECT = "Your Reservation Confirmation – The Athenaeum"
EMAIL_BODY = """Dear Guest,\n\nThank you for your reservation at The Athenaeum.\nAttached is your confirmation letter for your upcoming visit.\n\nIf you have any questions or need to make changes, please contact us.\n\nBest regards,\nAthenaeum Front Desk\nathdesk@caltech.edu"""

EMAIL_REGEX = r"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}"


# --- Functions ---
def extract_email_from_pdf(pdf_path):
    try:
        doc = fitz.open(pdf_path)
        for page in doc:
            text = page.get_text()
            match = re.search(EMAIL_REGEX, text)
            if match:
                return match.group(0)
        return None
    except Exception as e:
        print(f"Error reading PDF: {e}")
        return None

def send_email_with_attachment(to_email, pdf_path):
    msg = EmailMessage()
    msg["Subject"] = EMAIL_SUBJECT
    msg["From"] = SMTP_USER
    msg["To"] = to_email
    msg.set_content(EMAIL_BODY)

    with open(pdf_path, "rb") as f:
        file_data = f.read()
        file_name = os.path.basename(pdf_path)
        msg.add_attachment(file_data, maintype="application", subtype="pdf", filename=file_name)

    with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
        server.starttls()
        server.login(SMTP_USER, SMTP_PASSWORD)
        server.send_message(msg)
        print(f"Email sent to {to_email}")

def log_failure(pdf_path):
    with open("sent_log.csv", "a") as log_file:
        log_file.write(f"{os.path.basename(pdf_path)},No email found,{time.strftime('%Y-%m-%d %H:%M:%S')}\n")


# --- Watchdog Handler ---
class PDFHandler(FileSystemEventHandler):
    def on_created(self, event):
        if event.is_directory or not event.src_path.endswith(".pdf"):
            return

        time.sleep(1)  # Wait for file to finish writing
        pdf_path = event.src_path
        print(f"Detected PDF: {pdf_path}")

        email = extract_email_from_pdf(pdf_path)
        if email:
            send_email_with_attachment(email, pdf_path)
        else:
            print(f"No email found in {pdf_path}. Leaving for manual processing.")
            log_failure(pdf_path)


# --- Main Execution ---
if __name__ == "__main__":
    os.makedirs(WATCH_FOLDER, exist_ok=True)
    observer = Observer()
    event_handler = PDFHandler()
    observer.schedule(event_handler, path=WATCH_FOLDER, recursive=False)
    print(f"Watching folder: {WATCH_FOLDER}")
    observer.start()
    try:
        while True:
            time.sleep(10)
    except KeyboardInterrupt:
        observer.stop()
    observer.join()