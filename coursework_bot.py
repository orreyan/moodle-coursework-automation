#!/usr/bin/env python3
"""
Moodle Coursework Automation (CUD Bot)
Checks email -> logs into the Moodle portal -> finds the coursework -> downloads it
-> writes a task-notes brief (requirements, deliverables, deadline, plan of work)
"""

import re
import json
import imaplib
import email
from groq import Groq
from pathlib import Path
from datetime import datetime
from playwright.sync_api import sync_playwright
from docx import Document

OUTPUT_DIR = Path.home() / "Desktop"  # where task files and notes are saved
CREDENTIALS = Path.home() / ".cud_credentials"
PORTAL_URL = "https://my.cud.ac.ae/login/index.php"
PROCESSED_LOG = Path.home() / ".cud_processed_emails"


# ─────────────────────────────────────────────
#  CREDENTIALS
# ─────────────────────────────────────────────

def load_credentials():
    creds = {}
    with open(CREDENTIALS) as f:
        for line in f:
            line = line.strip()
            if "=" in line and not line.startswith("#"):
                key, _, val = line.partition("=")
                creds[key.strip()] = val.strip()
    return creds


# ─────────────────────────────────────────────
#  JSON HELPER
# ─────────────────────────────────────────────

def extract_json(text):
    """Extract the first valid JSON object from any AI response."""
    try:
        return json.loads(text.strip())
    except Exception:
        pass
    brace_count = 0
    start = -1
    for i, ch in enumerate(text):
        if ch == "{":
            if brace_count == 0:
                start = i
            brace_count += 1
        elif ch == "}":
            brace_count -= 1
            if brace_count == 0 and start != -1:
                try:
                    return json.loads(text[start:i + 1])
                except Exception:
                    start = -1
    return {"error": "no_json"}


# ─────────────────────────────────────────────
#  EMAIL
# ─────────────────────────────────────────────

def analyse_email(subject, body, client):
    """
    Ask AI whether this professor email requires the student to find and complete
    any academic work on the portal. Covers every possible task type.
    """
    prompt = f"""You are helping a university student manage their coursework.

Read this email from a professor and decide:
1. Is the professor directing the student to go find and complete some academic work
   on the course portal?
2. What is the work? (assignment, lab, lab report, midterm, quiz, final exam,
   homework, worksheet, essay, report, project, presentation, reading, exercise,
   problem set, case study, reflection, or anything similar)

Say YES if the professor clearly expects the student to open the portal and do
something — submit work, take a quiz, download a file, read material, etc.
Say NO for general announcements, grade notifications, schedule changes, or
emails where nothing needs to be done on the portal.

EMAIL SUBJECT: {subject}

EMAIL BODY:
{body[:3000]}

Reply ONLY with valid JSON, no markdown, no extra text:
{{
  "requires_portal_action": true or false,
  "task_type": "lab / quiz / midterm / essay / report / homework / etc — or null",
  "task_name": "exact name or title of the task as mentioned in the email, or null",
  "course_hint": "course code or name if mentioned, or null",
  "confidence": "high / medium / low",
  "reasoning": "one sentence"
}}"""

    try:
        resp = client.chat.completions.create(
            model="llama-3.3-70b-versatile",
            messages=[{"role": "user", "content": prompt}],
            max_tokens=300,
            temperature=0.1,
        )
        result = extract_json(resp.choices[0].message.content)
        print(f"  [Email AI] requires_portal_action={result.get('requires_portal_action')} "
              f"({result.get('confidence')}) | {result.get('task_type')} | "
              f"{result.get('reasoning', '')}")
        return result
    except Exception as e:
        print(f"  [Email AI error] {e}")
        # Broad keyword fallback — never silently drop legitimate emails
        keywords = [
            "assignment", "lab", "homework", "quiz", "midterm", "exam",
            "worksheet", "essay", "report", "project", "submit", "complete",
            "deadline", "due", "upload", "presentation", "exercise", "problem set",
        ]
        found = any(kw in (subject + " " + body).lower() for kw in keywords)
        return {
            "requires_portal_action": found,
            "task_type": None, "task_name": None, "course_hint": None,
            "confidence": "low", "reasoning": "fallback keyword check",
        }


def check_email(gmail_address, app_password, groq_client):
    print("Connecting to Gmail...")
    mail = imaplib.IMAP4_SSL("imap.gmail.com")
    mail.login(gmail_address, app_password)
    mail.select("inbox")

    _, data = mail.search(None, "UNSEEN")
    uids = data[0].split()
    if not uids:
        print("No unread emails.")
        mail.logout()
        return []

    processed = set()
    if PROCESSED_LOG.exists():
        processed = set(PROCESSED_LOG.read_text().splitlines())

    results = []
    for uid in uids:
        uid_str = uid.decode()
        if uid_str in processed:
            continue

        _, msg_data = mail.fetch(uid, "(RFC822)")
        msg = email.message_from_bytes(msg_data[0][1])
        subject = msg.get("Subject", "No Subject")
        sender  = msg.get("From", "Unknown")

        # Extract full body — plain text preferred, HTML as fallback
        body = ""
        if msg.is_multipart():
            for part in msg.walk():
                if part.get_content_type() == "text/plain":
                    body = part.get_payload(decode=True).decode(errors="ignore")
                    break
            if not body:
                for part in msg.walk():
                    if part.get_content_type() == "text/html":
                        body = part.get_payload(decode=True).decode(errors="ignore")
                        break
        else:
            body = msg.get_payload(decode=True).decode(errors="ignore")

        print(f"\nChecking: '{subject}' from {sender}")
        analysis = analyse_email(subject, body, groq_client)

        if not analysis.get("requires_portal_action"):
            print("  -> No portal action needed, skipping.")
            continue

        results.append({
            "uid":        uid_str,
            "subject":    subject,
            "sender":     sender,
            "body":       body,
            "task_type":  analysis.get("task_type"),
            "task_name":  analysis.get("task_name"),
            "course_hint":analysis.get("course_hint"),
        })
        print(f"  -> Portal action required: {analysis.get('task_type')} — "
              f"'{analysis.get('task_name')}'")

    mail.logout()
    return results


def mark_processed(uid):
    with open(PROCESSED_LOG, "a") as f:
        f.write(uid + "\n")


# ─────────────────────────────────────────────
#  PORTAL HELPERS  (no stray clicking)
# ─────────────────────────────────────────────

def dismiss_popups(page):
    """
    Non-blocking. Tries common dismiss buttons once, moves on immediately
    whether or not a popup was found. Never waits for one to appear.
    """
    for selector in [
        "button:has-text('Cancel')", "button:has-text('Close')",
        "button:has-text('Later')",  "button:has-text('Dismiss')",
        "button:has-text('Skip')",   "button:has-text('No thanks')",
        "[aria-label='Close']", ".modal-close",
    ]:
        try:
            btn = page.query_selector(selector)
            if btn and btn.is_visible():
                btn.click()
                page.wait_for_timeout(300)
                print(f"  Dismissed popup ({selector})")
                return
        except Exception:
            pass


def open_nav_drawer(page):
    """
    Opens the course navigation drawer ONLY — nothing else.
    Does not touch notifications, messages, profile, or any other UI element.
    """
    try:
        drawer = page.query_selector(
            "#nav-drawer-toggle, "
            "button[aria-label='Open drawer'], "
            "button[aria-label='Side panel']"
        )
        if drawer and drawer.is_visible():
            if drawer.get_attribute("aria-expanded") == "false":
                drawer.click()
                page.wait_for_timeout(600)
                print("  Opened nav drawer.")
    except Exception:
        pass


def find_and_click(page, target_text):
    """Find a link or button by text and click it. Returns True on success."""
    if not target_text or target_text == "NOT_FOUND":
        return False
    print(f"  Clicking: '{target_text[:80]}'")

    el = None
    for selector in [
        f'a:has-text("{target_text}")',
        f'button:has-text("{target_text}")',
        f'span:has-text("{target_text}")',
    ]:
        try:
            el = page.query_selector(selector)
            if el:
                break
        except Exception:
            pass

    # Partial-match fallback
    if not el:
        try:
            for candidate in page.query_selector_all("a, button"):
                text = candidate.inner_text().strip()
                if target_text.lower() in text.lower() or text.lower() in target_text.lower():
                    el = candidate
                    print(f"    (partial match: '{text[:60]}')")
                    break
        except Exception:
            pass

    if not el:
        print(f"  Not found: '{target_text[:60]}'")
        return False

    try:
        el.scroll_into_view_if_needed()
        page.wait_for_timeout(200)
        el.click(timeout=5000)
        return True
    except Exception:
        try:
            el.evaluate("el => el.click()")
            return True
        except Exception as e:
            print(f"  Click failed: {e}")
            return False


# ─────────────────────────────────────────────
#  AI PAGE DECISION
# ─────────────────────────────────────────────

def ai_pick_element(page, email_data, task_description, client, step):
    """Ask AI which element to click on the current page."""
    print(f"[AI Step {step}] {task_description}")

    visible_text = page.inner_text("body")[:4000]
    elements = []
    for el in page.query_selector_all("a, button")[:80]:
        text = el.inner_text().strip()
        if text and len(text) < 150:
            elements.append({
                "text":    text[:100],
                "visible": el.is_visible(),
            })

    prompt = f"""You are navigating a university Moodle portal for a student.

PROFESSOR EMAIL SUBJECT: {email_data['subject']}
PROFESSOR EMAIL BODY: {email_data['body'][:1200]}
TASK TYPE: {email_data.get('task_type') or 'unknown'}
TASK NAME: {email_data.get('task_name') or 'unknown'}
COURSE:    {email_data.get('course_hint') or 'unknown'}

YOUR JOB NOW: {task_description}

PAGE TEXT (truncated):
{visible_text[:2500]}

CLICKABLE ELEMENTS ON PAGE:
{json.dumps(elements, indent=2)[:2000]}

Choose the single best element to click to accomplish the job.
Return ONLY valid JSON:
{{"target_text": "exact text of the element", "confidence": "high/medium/low", "reasoning": "brief"}}

If nothing is suitable: {{"target_text": "NOT_FOUND", "confidence": "low", "reasoning": "why"}}"""

    try:
        resp = client.chat.completions.create(
            model="llama-3.3-70b-versatile",
            messages=[{"role": "user", "content": prompt}],
            max_tokens=300,
            temperature=0.1,
        )
        result = extract_json(resp.choices[0].message.content)
        print(f"  [AI] '{result.get('target_text','?')[:70]}' "
              f"({result.get('confidence','?')}) — {result.get('reasoning','')}")
        return result
    except Exception as e:
        print(f"  [AI error] {e}")
        return {"target_text": "NOT_FOUND", "confidence": "low", "reasoning": str(e)}


# ─────────────────────────────────────────────
#  PORTAL NAVIGATION  (login -> course -> task)
# ─────────────────────────────────────────────

def navigate_to_task(page, email_data, client):
    """
    Straight-line navigation:
      1. Open nav drawer (only safe expansion)
      2. AI picks the course link
      3. AI picks the task link inside the course
    Returns the task name string, or None on failure.
    """

    # Step 1: open drawer if needed (no other expansion)
    open_nav_drawer(page)
    page.wait_for_timeout(1000)

    # Step 2: find and click the course
    print("\n[Step 1] Finding course...")
    decision = ai_pick_element(
        page, email_data,
        "Click the link for the course mentioned in the email. "
        "Look for a course code or course name in the list of courses.",
        client, 1,
    )
    if decision["target_text"] == "NOT_FOUND":
        print("Could not find the course.")
        return None
    if not find_and_click(page, decision["target_text"]):
        return None
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(1500)
    dismiss_popups(page)

    # Step 3: find and click the specific task inside the course
    print("\n[Step 2] Finding task...")
    decision = ai_pick_element(
        page, email_data,
        f"Find and click the specific {email_data.get('task_type') or 'task'} "
        f"called '{email_data.get('task_name') or 'unknown'}' as mentioned in the email. "
        "It could be a lab, quiz, assignment, essay, midterm, report, homework, etc.",
        client, 2,
    )
    if decision["target_text"] == "NOT_FOUND":
        print("Could not find the task.")
        return None
    if not find_and_click(page, decision["target_text"]):
        return None
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(1500)
    dismiss_popups(page)

    return decision["target_text"]


# ─────────────────────────────────────────────
#  DOWNLOAD  (pdf, docx, doc, pptx, ppt, txt)
# ─────────────────────────────────────────────

def download_task_file(page, task_name):
    """
    Download whatever file the task page offers.
    Supports PDF, DOCX, DOC, PPTX, PPT, XLSX, and generic Moodle pluginfile links.
    Falls back to saving the page text if no file is found.
    """
    print("\n[Download] Looking for task file...")

    safe = re.sub(r"[^\w\s-]", "", task_name).strip()
    safe = re.sub(r"\s+", "_", safe) or "task"

    # Extension -> save suffix mapping
    file_selectors = [
        ('a[href*=".pdf"]',         ".pdf"),
        ('a[href*=".docx"]',        ".docx"),
        ('a[href*=".doc"]',         ".doc"),
        ('a[href*=".pptx"]',        ".pptx"),
        ('a[href*=".ppt"]',         ".ppt"),
        ('a[href*=".xlsx"]',        ".xlsx"),
        ('a[href*="pluginfile.php"]',".bin"),   # Moodle generic; real ext set after download
        ('a:has-text("Download")',  ".bin"),
    ]

    for selector, ext in file_selectors:
        try:
            links = page.query_selector_all(selector)
            for link in links:
                try:
                    save_path = OUTPUT_DIR / f"{safe}{ext}"
                    with page.expect_download(timeout=20000) as dl_info:
                        link.click()
                    download = dl_info.value
                    # Use the browser-suggested filename to get the real extension
                    suggested = download.suggested_filename
                    if suggested:
                        real_ext = Path(suggested).suffix or ext
                        save_path = OUTPUT_DIR / f"{safe}{real_ext}"
                    download.save_as(str(save_path))
                    print(f"  Downloaded: {save_path.name}")
                    return save_path
                except Exception as e:
                    print(f"  Download attempt failed: {e}")
        except Exception:
            continue

    # Fallback: save visible page text
    print("  No file found — saving page text.")
    save_path = OUTPUT_DIR / f"{safe}.txt"
    save_path.write_text(page.inner_text("body"))
    print(f"  Saved page text: {save_path.name}")
    return save_path


# ─────────────────────────────────────────────
#  PORTAL AUTOMATION  (top-level)
# ─────────────────────────────────────────────

def run_portal(email_data, username, password, groq_client):
    print("\n=== Portal Automation ===")
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        context = browser.new_context(accept_downloads=True)
        page    = context.new_page()

        # Login
        print("Logging in...")
        page.goto(PORTAL_URL)
        page.fill("#username", username)
        page.fill("#password", password)
        page.click("#loginbtn")
        page.wait_for_load_state("networkidle")

        if "login" in page.url:
            print("Login failed.")
            browser.close()
            return None
        print("Logged in.")
        dismiss_popups(page)

        # Go straight to dashboard — no other pages
        page.goto("https://my.cud.ac.ae/my/")
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(1500)
        dismiss_popups(page)

        # Navigate to the task
        task_name = navigate_to_task(page, email_data, groq_client)
        if not task_name:
            print("Navigation failed.")
            browser.close()
            return None

        # Download the file
        file_path = download_task_file(page, task_name)
        browser.close()
        return file_path, task_name


# ─────────────────────────────────────────────
#  READ DOCUMENT  (pdf, docx, pptx, txt)
# ─────────────────────────────────────────────

def read_document(path: Path) -> str:
    """Extract text from PDF, DOCX, PPTX, PPT, or plain text files."""
    suffix = path.suffix.lower()

    if suffix == ".txt":
        return path.read_text(errors="ignore")

    if suffix in (".docx", ".doc"):
        doc = Document(str(path))
        return "\n".join(p.text for p in doc.paragraphs)

    if suffix == ".pdf":
        try:
            import pdfplumber
            with pdfplumber.open(str(path)) as pdf:
                return "\n".join(
                    page.extract_text() or "" for page in pdf.pages
                )
        except ImportError:
            try:
                import PyPDF2
                with open(path, "rb") as f:
                    reader = PyPDF2.PdfReader(f)
                    return "\n".join(
                        page.extract_text() or "" for page in reader.pages
                    )
            except ImportError:
                return f"[PDF reading requires pdfplumber or PyPDF2 — pip install pdfplumber]"

    if suffix in (".pptx", ".ppt"):
        try:
            from pptx import Presentation
            prs = Presentation(str(path))
            lines = []
            for slide in prs.slides:
                for shape in slide.shapes:
                    if hasattr(shape, "text") and shape.text.strip():
                        lines.append(shape.text.strip())
            return "\n".join(lines)
        except ImportError:
            return "[PPTX reading requires python-pptx — pip install python-pptx]"

    if suffix in (".xlsx", ".xls"):
        try:
            import openpyxl
            wb = openpyxl.load_workbook(str(path), read_only=True, data_only=True)
            lines = []
            for ws in wb.worksheets:
                for row in ws.iter_rows(values_only=True):
                    row_text = "\t".join(str(c) for c in row if c is not None)
                    if row_text.strip():
                        lines.append(row_text)
            return "\n".join(lines)
        except ImportError:
            return "[XLSX reading requires openpyxl — pip install openpyxl]"

    # Generic fallback
    return path.read_text(errors="ignore")


# ─────────────────────────────────────────────
#  TASK NOTES
# ─────────────────────────────────────────────

def generate_task_notes(file_path, task_name, email_data, groq_client):
    """Turn the downloaded brief into a to-do notes file: what is being asked,
    what to hand in, when it is due, and a plan of work. It deliberately does
    not attempt the work itself."""
    print("\nWriting task notes...")

    content = read_document(file_path)
    if len(content) > 8000:
        content = content[:8000] + "\n...[truncated]..."

    prompt = f"""You are helping a university student plan their coursework.
Read the professor's email and the task brief, then write concise planning notes.
Do NOT answer or solve any of the questions; only describe what has to be done.

TASK TYPE: {email_data.get('task_type') or 'academic work'}
TASK NAME: {task_name}

PROFESSOR'S EMAIL:
{email_data['body'][:1500]}

TASK BRIEF:
{content}

Use exactly these headings:
SUMMARY - two or three sentences on what the task is about
DEADLINE - the due date and time if stated, otherwise "Not stated"
DELIVERABLES - what must be submitted, in what format, and where
REQUIREMENTS CHECKLIST - every question, part or marking requirement, as a numbered list
PLAN OF WORK - suggested steps in order, with rough time estimates
QUESTIONS FOR THE PROFESSOR - anything unclear in the brief"""

    resp = groq_client.chat.completions.create(
        model="llama-3.3-70b-versatile",
        messages=[{"role": "user", "content": prompt}],
        max_tokens=1500,
    )
    notes_text = resp.choices[0].message.content

    clean = re.sub(r"[^\w\s-]", "", task_name).strip()
    clean = re.sub(r"\s+", "_", clean) or "task"
    notes_path = OUTPUT_DIR / f"Notes_{clean}.docx"

    doc = Document()
    doc.add_heading(f"Task notes: {clean}", 0)
    doc.add_paragraph(f"Generated:  {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    doc.add_paragraph(f"Task file:  {file_path.name}")
    doc.add_paragraph(f"Task type:  {email_data.get('task_type') or 'unknown'}")
    doc.add_paragraph("=" * 50)
    for line in notes_text.split("\n"):
        doc.add_paragraph(line)
    doc.save(str(notes_path))
    print(f"Notes saved: {notes_path.name}")
    return notes_path


# ─────────────────────────────────────────────
#  MAIN
# ─────────────────────────────────────────────

def main():
    print("=" * 60)
    print("  CUD Bot — checks email, downloads coursework, writes task notes")
    print("=" * 60)

    try:
        creds = load_credentials()
    except FileNotFoundError:
        print("~/.cud_credentials not found. Required keys:")
        print("  CUD_USERNAME, CUD_PASSWORD, GMAIL_ADDRESS, GMAIL_APP_PASSWORD, GROQ_API_KEY")
        return

    username  = creds.get("CUD_USERNAME", "")
    password  = creds.get("CUD_PASSWORD", "")
    gmail_addr = creds.get("GMAIL_ADDRESS", "")
    gmail_pw  = creds.get("GMAIL_APP_PASSWORD", "")
    groq_key  = creds.get("GROQ_API_KEY", "")

    if not all([username, password, gmail_addr, gmail_pw, groq_key]):
        print("Missing credentials in ~/.cud_credentials")
        return

    groq_client = Groq(api_key=groq_key)
    emails = check_email(gmail_addr, gmail_pw, groq_client)

    if not emails:
        print("\nNothing to do.")
        return

    for email_data in emails:
        print(f"\n{'='*60}")
        print(f"Processing: {email_data['subject']}")
        print(f"{'='*60}")
        try:
            result = run_portal(email_data, username, password, groq_client)
            if not result:
                print("Failed.")
                continue

            file_path, task_name = result
            notes_path = generate_task_notes(file_path, task_name, email_data, groq_client)
            mark_processed(email_data["uid"])

            print(f"\n{'='*60}")
            print(f"Done!")
            print(f"  Task file : {file_path.name}")
            print(f"  Notes     : {notes_path.name}")
            print(f"{'='*60}")

        except Exception as e:
            import traceback
            print(f"Error: {e}")
            traceback.print_exc()

    print("\nAll done!")


if __name__ == "__main__":
    main()