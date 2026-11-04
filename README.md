# CUD Bot Moodle Coursework Automation

An automation pipeline that watches a student's inbox for new coursework, fetches it from the university's **Moodle** portal, and turns each brief into a short planning notes file. It combines email processing, LLM-based classification, and browser automation with **Playwright**.

---

## Technology

| Area | Technology |
| --- | --- |
| Language | Python 3 |
| Browser automation | Playwright (Chromium) |
| Email | IMAP via `imaplib` (Gmail) |
| AI | Groq API (Llama 3.3 70B) |
| Documents | python-docx, pdfplumber / PyPDF2, python-pptx, openpyxl |

---

## How It Works

### 1. Email detection
Connects to Gmail over IMAP and reads new messages. Instead of matching subject keywords, an LLM reads the full email body and decides whether it is coursework, what type it is (assignment, quiz, lab, midterm, project, report, presentation, and more), which course it belongs to, and the task's name. Processed emails are logged so nothing is handled twice.

### 2. Portal navigation
Logs into the Moodle portal with Playwright, dismisses any pop-ups without blocking, and opens the course navigation drawer. It then follows a direct path from the dashboard to the course and the task. When the page layout is ambiguous, the LLM is shown the visible links and picks the right one.

### 3. Download
Downloads the task file using the browser's suggested filename, so the original extension (DOCX, PDF, PPTX, XLSX) is kept. If there is no attachment, the task description is saved as text instead.

### 4. Task notes
Reads the downloaded file and writes a `Notes_<task>.docx` brief with a summary, the deadline, the deliverables, a checklist of every requirement, a suggested plan of work, and questions to ask the professor. It describes what has to be done and does not attempt the work itself.

---

## Concepts Demonstrated

End-to-end workflow automation, LLM-based classification and decision-making, resilient browser automation against a real web portal, multi-format document parsing, idempotent processing, and keeping credentials out of source code.

---

## Running Locally

```bash
pip install -r requirements.txt
playwright install chromium

cp cud_credentials.example ~/.cud_credentials   # then fill in your own values
python coursework_bot.py
```

Credentials are read from `~/.cud_credentials`, outside the repository, and are never stored in the code. Gmail requires an [app password](https://myaccount.google.com/apppasswords).

---

## Author

**Reyan Fethi**
