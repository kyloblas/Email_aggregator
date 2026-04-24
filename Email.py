import os
import email
import pandas as pd
from bs4 import BeautifulSoup
from email import policy
from email.parser import BytesParser


# ==============================
# Load EML
# ==============================
def load_eml(filepath):
    with open(filepath, 'rb') as f:
        msg = BytesParser(policy=policy.default).parse(f)
    return msg


def get_html_body(msg):
    """Get HTML body whether the email is multipart or single-part."""
    if msg.is_multipart():
        for part in msg.walk():
            if part.get_content_type() == "text/html":
                return part.get_content()
    elif msg.get_content_type() == "text/html":
        return msg.get_content()
    return ""


def extract_metadata(msg):
    return {
        "date": msg.get("Date"),
        "sender": msg.get("From", "")
    }


def detect_source(sender):
    sender = sender.lower()
    if "linkedin" in sender:
        return "linkedin"
    elif "jooble" in sender:
        return "jooble"
    elif "indeed" in sender:
        return "indeed"
    elif "efinancialcareers" in sender:
        return "efinancialcareers"
    elif "adzuna" in sender:
        return "adzuna"
    else:
        return "other"


# ==============================
# LINKEDIN PARSER
# ==============================
def parse_linkedin(soup):
    """
    LinkedIn emails have two <a> tags per job with /comm/jobs/view/ in href:
      1. An EMPTY anchor wrapping the full card (contains all info in the parent <tr>)
      2. A text anchor containing just the job title (isolated, no company sibling)

    We target the EMPTY anchors and read title + company from the parent <tr>.
    Company appears as "Company · Location" format.
    """
    jobs = []
    seen = set()

    for link in soup.find_all("a", href=True):
        href = link["href"]
        if "/comm/jobs/view/" not in href:
            continue
        if link.get_text(strip=True):
            # This is the text-only title link — skip it
            continue

        # Deduplicate by URL
        base_href = href.split("?")[0]
        if base_href in seen:
            continue
        seen.add(base_href)

        row = link.find_parent("tr")
        if not row:
            continue

        texts = [t.strip() for t in row.get_text("\n", strip=True).split("\n") if t.strip()]
        if not texts:
            continue

        title = texts[0]
        company = ""

        # Company is typically "Company · Location"
        for t in texts[1:]:
            if "·" in t:
                company = t.split("·")[0].strip()
                break
            elif t.lower() not in ("high experience match", "medium experience match", "low experience match"):
                company = t
                break

        jobs.append({
            "job_title": title,
            "company": company,
            "job_link": href
        })

    return jobs


# ==============================
# JOOBLE PARSER
# ==============================
def parse_jooble(soup, subject=""):
    """
    Jooble emails have two layout types, but both share these properties:
      - Every job link (/away/) contains an <h2> with the CLEAN job title
      - Company lives in an ancestor <table> that starts with the title text

    Type A (bulk "recommended" emails):
      The <a> link text is dirty (title + company + location + emoji tags concatenated).
      Immediate parent <table>: Title | £salary | Company | • | Location | emojis

    Type B ("you've got a match" single-job emails):
      The <a> link text is clean (just the title).
      Must walk up ~10 table levels to find: Title | £salary | Company | Location | ago | emojis
      Fallback: subject line format is "🔥 You've got a match: COMPANY is looking for TITLE"

    Strategy: use h2.get_text() as the clean title, then walk up ALL ancestor tables
    to find the first one where texts[0] == title AND there's a valid company after it.
    If not found, try extracting company from the email subject line.
    """
    jobs = []
    seen = set()

    EMOJI_PREFIXES = ("✅", "💼", "🔥", "📍", "🏢", "🌍", "🔍", "⭐")
    LOCATION_KEYWORDS = ("London", "Greater", "United Kingdom", "England", "Scotland",
                         "Wales", "Remote", "Hybrid", "West London", "East London",
                         "North London", "South London", "City of London")
    SKIP_TOKENS = {"•", "·", "TOP 4 jobs", "based on your search preferences."}
    SKIP_PREFIXES = ("Go to ", "Job Description", "Location:", "Travel:", "Compensation:",
                     "Show more", "Hybrid -")

    def is_skip(text):
        if text in SKIP_TOKENS:
            return True
        if text.startswith("£") or text.startswith("$"):
            return True
        if any(text.startswith(e) for e in EMOJI_PREFIXES):
            return True
        if any(kw in text for kw in LOCATION_KEYWORDS):
            return True
        if any(text.startswith(p) for p in SKIP_PREFIXES):
            return True
        if "ago" in text.lower() and len(text) < 25:
            return True
        if len(text) <= 2:
            return True
        return False

    for link in soup.find_all("a", href=True):
        href = link["href"]
        if "/away/" not in href:
            continue

        base_href = href.split("?")[0]
        if base_href in seen:
            continue
        seen.add(base_href)

        # -----------------------------------------------------------------
        # Get title — three layout types:
        # Type A (bulk recommended): link wraps <h2> with title
        # Type B (you-got-a-match):  link text IS the clean title
        # Type C (bulk digest):      link is EMPTY, title in parent <td>
        # -----------------------------------------------------------------
        import re
        h2 = link.find("h2")
        link_text = link.get_text(strip=True)

        if h2:
            title = h2.get_text(strip=True)        # Type A
        elif link_text:
            title = link_text                       # Type B
        else:
            td = link.find_parent("td")             # Type C
            if not td:
                continue
            td_texts = [t.strip() for t in td.get_text("\n", strip=True).split("\n") if t.strip()]
            title = td_texts[0] if td_texts else ""

        if not title:
            continue

        company = ""

        # Type C — company directly in parent <td>
        # Structure: Title | £salary | Company | • | Location | emojis
        if not link_text and not h2:
            td = link.find_parent("td")
            if td:
                td_texts = [t.strip() for t in td.get_text("\n", strip=True).split("\n") if t.strip()]
                for t in td_texts[1:]:
                    if is_skip(t):
                        continue
                    company = t
                    break

        # Type A & B — walk up ancestor tables to find company
        if not company:
            for table in link.parents:
                if table.name != "table":
                    continue
                texts = [t.strip() for t in table.get_text("\n", strip=True).split("\n") if t.strip()]
                if not texts or texts[0] != title:
                    continue
                for t in texts[1:]:
                    if is_skip(t):
                        continue
                    company = t
                    break
                if company:
                    break

        # Fallback: subject "... match: COMPANY is looking for ..."
        if not company and subject:
            m = re.search(r'match[:\s]+(.+?)\s+is looking for', subject, re.IGNORECASE)
            if m:
                company = m.group(1).strip()

        jobs.append({
            "job_title": title,
            "company": company,
            "job_link": href
        })

    return jobs


# ==============================
# INDEED PARSER
# ==============================
def parse_indeed(soup):
    """
    Indeed emails have two <a> tags per job with /rc/clk/dl or /pagead/clk/dl in href:
      1. An EMPTY anchor wrapping the full card (parent <tr> has Title, Company, Location...)
      2. A text anchor containing just the job title (isolated)

    We target EMPTY anchors and read title + company from the parent <tr>.
    Row order: Title | Company | Location | Description | N days ago
    """
    jobs = []
    seen = set()

    def is_indeed_job_link(href):
        return "/rc/clk/dl" in href or "/pagead/clk/dl" in href

    for link in soup.find_all("a", href=True):
        href = link["href"]
        if not is_indeed_job_link(href):
            continue
        if link.get_text(strip=True):
            # Text-only title link — skip
            continue

        base_href = href.split("?")[0]
        if base_href in seen:
            continue
        seen.add(base_href)

        row = link.find_parent("tr")
        if not row:
            continue

        texts = [t.strip() for t in row.get_text("\n", strip=True).split("\n") if t.strip()]
        if not texts:
            continue

        title = texts[0]
        company = ""

        # Company is second text item; skip if it looks like metadata
        for t in texts[1:]:
            if (
                "£" in t
                or "ago" in t.lower()
                or "days" in t.lower()
                or "hours" in t.lower()
                or t.lower() in ("london", "remote", "hybrid")
                or len(t) <= 2
            ):
                continue
            company = t
            break

        jobs.append({
            "job_title": title,
            "company": company,
            "job_link": href
        })

    return jobs


# ==============================
# EFINANCIALCAREERS PARSER
# ==============================
def parse_efinancialcareers(soup):
    """
    eFinancialCareers emails contain job title links pointing to their pub/cc tracking URL.
    Walking up to the grandparent <table> gives: Title | Company | Location & contract type | Salary
    We skip 'Apply now' and 'View online' links.
    """
    jobs = []
    seen = set()

    SKIP_TEXTS = {
        "apply now", "view online", "view all jobs", "see all jobs",
        "show more", "manage my job alerts", "unsubscribe from this job alert",
        "sign in", "privacy policy", "terms"
    }

    for link in soup.find_all("a", href=True):
        href = link["href"]
        title = link.get_text(strip=True)

        if not title or title.lower() in SKIP_TEXTS:
            continue
        if "efinancialcareers" not in href and "pub/cc" not in href:
            continue

        # Deduplicate by title+href combo (all tracking URLs are identical)
        key = title.lower()
        if key in seen:
            continue
        seen.add(key)

        # Walk up to find a <table> that contains the company info
        container = None
        p = link
        for _ in range(10):
            p = p.parent
            if p is None:
                break
            if p.name == "table":
                texts = [t.strip() for t in p.get_text("\n", strip=True).split("\n") if t.strip()]
                # A job card table will have at least title + company
                if len(texts) >= 2 and texts[0] == title:
                    container = p
                    break

        if not container:
            continue

        texts = [t.strip() for t in container.get_text("\n", strip=True).split("\n") if t.strip()]

        company = ""
        if len(texts) > 1:
            raw = texts[1]
            # Format 1: "Company Name" (standalone)
            # Format 2: "Company Name | Location, Country"
            # Format 3: "Company Name • Location"
            if "|" in raw:
                company = raw.split("|")[0].strip()
            elif "•" in raw:
                company = raw.split("•")[0].strip()
            elif not raw.startswith("£") and "United Kingdom" not in raw and "London" not in raw:
                company = raw

        jobs.append({
            "job_title": title,
            "company": company,
            "job_link": href
        })

    return jobs


# ==============================
# ADZUNA PARSER
# ==============================
def parse_adzuna(soup):
    """
    Adzuna emails have job title links (<h2> or direct <a>) pointing to adzuna.co.uk/jobs/land/ad/
    The parent <tr> contains: Title | 'TOP MATCH' (optional) | 'Company - Location - £Salary' | 'more details »'
    Company is extracted from the combined "Company - Location - Salary" line.
    """
    jobs = []
    seen = set()

    for link in soup.find_all("a", href=True):
        href = link["href"]
        if "adzuna.co.uk/jobs/land/ad/" not in href:
            continue

        title = link.get_text(strip=True)
        if not title or title.lower() == "more details »":
            continue

        base_href = href.split("?")[0]
        if base_href in seen:
            continue
        seen.add(base_href)

        row = link.find_parent("tr")
        if not row:
            continue

        texts = [t.strip() for t in row.get_text("\n", strip=True).split("\n") if t.strip()]

        company = ""
        for t in texts:
            if t == title or t.lower() in ("top match", "more details »"):
                continue
            # Company line format: "CompanyName - Location - £Salary" or "CompanyName - Location - Salary"
            if " - " in t:
                company = t.split(" - ")[0].strip()
                break

        jobs.append({
            "job_title": title,
            "company": company,
            "job_link": href
        })

    return jobs


# ==============================
# MAIN EMAIL PARSER
# ==============================
def parse_email(filepath):
    msg = load_eml(filepath)
    meta = extract_metadata(msg)
    html = get_html_body(msg)

    if not html:
        print(f"  [WARNING] No HTML body found in: {os.path.basename(filepath)}")
        return []

    soup = BeautifulSoup(html, "lxml")
    source = detect_source(meta["sender"])

    if source == "linkedin":
        jobs = parse_linkedin(soup)
    elif source == "jooble":
        jobs = parse_jooble(soup, subject=msg.get("Subject", ""))
    elif source == "indeed":
        jobs = parse_indeed(soup)
    elif source == "efinancialcareers":
        jobs = parse_efinancialcareers(soup)
    elif source == "adzuna":
        jobs = parse_adzuna(soup)
    else:
        print(f"  [SKIPPED] Unknown source for sender: {meta['sender']}")
        jobs = []

    for job in jobs:
        job["date"] = meta["date"]
        job["email_sender"] = meta["sender"]
        job["source"] = source

    return jobs


# ==============================
# FOLDER PROCESSOR
# ==============================
def parse_folder(folder_path):
    all_jobs = []

    eml_files = [f for f in os.listdir(folder_path) if f.endswith(".eml")]
    print(f"Found {len(eml_files)} .eml files in {folder_path}\n")

    for filename in eml_files:
        filepath = os.path.join(folder_path, filename)
        print(f"Processing: {filename}")
        jobs = parse_email(filepath)
        print(f"  -> {len(jobs)} jobs extracted")
        all_jobs.extend(jobs)

    return all_jobs


# ==============================
# EXPORT TO EXCEL
# ==============================
def export_to_excel(data, output_file):
    df = pd.DataFrame(data)

    if df.empty:
        print("\nNo jobs found. Nothing exported.")
        return

    columns_order = [
        "date",
        "email_sender",
        "source",
        "company",
        "job_title",
        "job_link"
    ]

    df = df[columns_order]

    # Sort by date descending, then company
    df["date"] = pd.to_datetime(df["date"], errors="coerce", utc=True)
    df = df.sort_values(["date", "company"], ascending=[False, True])
    df["date"] = df["date"].dt.strftime("%Y-%m-%d")

    df.to_excel(output_file, index=False)
    print(f"\nExported {len(df)} jobs to {output_file}")


# ==============================
# RUN
# ==============================
if __name__ == "__main__":

    folder_path = r"G:\My Drive\Python\Email_python"
    output_file = "job_results.xlsx"

    jobs = parse_folder(folder_path)
    export_to_excel(jobs, output_file)