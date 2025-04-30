# main.py

# --- Import necessary libraries ---
import os
import time
import re
import fitz  # PyMuPDF for reading PDF files
import smtplib
from email.message import EmailMessage
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

# --- Configuration Section ---
WATCH_FOLDER = "./confirmations"  # Folder to watch for new PDF files
SMTP_SERVER = "smtp.office365.com"  # SMTP server address
SMTP_PORT = 587  # SMTP port for TLS
SMTP_USER = "email@domain.com"  # Email address used to send confirmations
SMTP_PASSWORD = os.getenv("SMTP_PASSWORD")  # Email account password loaded from environment variable

# Check if SMTP password is set to avoid runtime errors
if not SMTP_PASSWORD:
    raise EnvironmentError("SMTP_PASSWORD environment variable not set.")

# Email contents configuration
EMAIL_SUBJECT = "Your Reservation Confirmation – The Athenaeum"
EMAIL_BODY = """Dear Guest,\n\nThank you for your reservation at The Athenaeum.\nAttached is your confirmation letter for your upcoming visit.\n\nIf you have any questions or need to make changes, please contact us.\n\nBest regards,\nAthenaeum Front Desk\nathdesk@caltech.edu"""

# Regular expression to detect email addresses in text
EMAIL_REGEX = r"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}"

# Log file name
LOG_FILE = "sent_log.csv"

# --- Function to extract email address from PDF ---
def extract_email_from_pdf(pdf_path):
    try:
        doc = fitz.open(pdf_path)  # Open the PDF file
        for page in doc:
            text = page.get_text()  # Extract text from each page
            match = re.search(EMAIL_REGEX, text)  # Search for email pattern
            if match:
                return match.group(0)  # Return the first email found
        return None  # No email found in the PDF
    except Exception as e:
        print(f"Error reading PDF: {e}")
        return None

# --- Function to send email with PDF attached ---
def send_email_with_attachment(to_email, pdf_path):
    msg = EmailMessage()
    msg["Subject"] = EMAIL_SUBJECT
    msg["From"] = SMTP_USER
    msg["To"] = to_email
    msg.set_content(EMAIL_BODY)

    # Attach the PDF file to the email
    with open(pdf_path, "rb") as f:
        file_data = f.read()
        file_name = os.path.basename(pdf_path)
        msg.add_attachment(file_data, maintype="application", subtype="pdf", filename=file_name)

    # Connect to the SMTP server and send the email
    with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
        server.starttls()  # Secure the connection
        server.login(SMTP_USER, SMTP_PASSWORD)
        server.send_message(msg)
        print(f"Email sent to {to_email}")

    # Log the successful email send
    log_success(pdf_path, to_email)

# --- Function to log failure to find an email ---
def log_failure(pdf_path):
    write_log_entry(os.path.basename(pdf_path), "No email found", "N/A")

# --- Function to log successful email send ---
def log_success(pdf_path, to_email):
    write_log_entry(os.path.basename(pdf_path), "Email sent", to_email)

# --- Function to write an entry to the log file ---
def write_log_entry(file_name, status, email):
    header_needed = not os.path.exists(LOG_FILE)  # Check if header needs to be written
    with open(LOG_FILE, "a") as log_file:
        if header_needed:
            log_file.write("Filename,Status,Email,Timestamp\n")  # Write header if log file is new
        log_file.write(f"{file_name},{status},{email},{time.strftime('%Y-%m-%d %H:%M:%S')}\n")  # Write log entry

# --- Watchdog Event Handler for new PDFs ---
class PDFHandler(FileSystemEventHandler):
    def on_created(self, event):
        # Only act on new PDF files
        if event.is_directory or not event.src_path.endswith(".pdf"):
            return

        time.sleep(2)  # Wait for file to finish writing completely
        pdf_path = event.src_path
        print(f"Detected PDF: {pdf_path}")

        # Extract email and either send or log failure
        email = extract_email_from_pdf(pdf_path)
        if email:
            send_email_with_attachment(email, pdf_path)
        else:
            print(f"No email found in {pdf_path}. Leaving for manual processing.")
            log_failure(pdf_path)

# --- Main program execution ---
if __name__ == "__main__":
    os.makedirs(WATCH_FOLDER, exist_ok=True)  # Ensure watch folder exists
    observer = Observer()
    event_handler = PDFHandler()
    observer.schedule(event_handler, path=WATCH_FOLDER, recursive=False)  # Monitor folder for new PDFs
    print(f"Watching folder: {WATCH_FOLDER}")

    observer.start()  # Start observing the folder
    try:
        while True:
            time.sleep(10)  # Keep script running
    except KeyboardInterrupt:
        observer.stop()  # Gracefully stop on Ctrl+C
    observer.join()  # Wait for observer to finish
