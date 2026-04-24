"""
gmail_connector.py
==================
Handles all Gmail API interaction:
  - Authentication (OAuth2 with token caching)
  - Fetching messages by label
  - Converting raw Gmail messages to MSG objects (same format as .eml)
  - Applying the Job_Processed label after successful processing

FIRST-TIME SETUP
----------------
1. Go to https://console.cloud.google.com
2. Create a project → Enable "Gmail API"
3. Create OAuth 2.0 credentials → Desktop App
4. Download credentials JSON → save as credentials.json next to this file
5. pip install google-auth google-auth-oauthlib google-auth-httplib2 google-api-python-client

GMAIL LABEL SETUP (do this once in Gmail)
------------------------------------------
Create these labels in Gmail:
  Job_Alerts_LinkedIn
  Job_Alerts_Jooble
  Job_Alerts_Indeed
  Job_Alerts_eFinancial
  Job_Alerts_Adzuna
  Job_Processed1

Then set up filters in Gmail to auto-label incoming job alert emails.
"""

import os
import base64
import pickle
from email import policy
from email.parser import BytesParser

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build


# ==============================
# CONFIG
# ==============================

SCOPES = ["https://www.googleapis.com/auth/gmail.modify"]

CREDENTIALS_FILE = "credentials.json"   # Downloaded from Google Cloud Console
TOKEN_FILE       = "token.pickle"        # Auto-created after first login

# Labels to fetch from (query unprocessed emails from any of these)
JOB_ALERT_LABELS = [
    "Job_Alerts_LinkedIn",
    "Job_Alerts_Jooble",
    "Job_Alerts_Indeed",
    "Job_Alerts_eFinancial",
    "Job_Alerts_Adzuna",
]

PROCESSED_LABEL = "Job_Processed1"


# ==============================
# AUTHENTICATION
# ==============================

def get_gmail_service():
    creds = None

    if os.path.exists(TOKEN_FILE):
        try:
            with open(TOKEN_FILE, "rb") as f:
                creds = pickle.load(f)
        except Exception:
            print("Token corrupted — deleting")
            os.remove(TOKEN_FILE)
            creds = None

    if not creds or not creds.valid:
        try:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                raise Exception("Need new login")
        except Exception:
            if os.path.exists(TOKEN_FILE):
                os.remove(TOKEN_FILE)

            flow = InstalledAppFlow.from_client_secrets_file(
                CREDENTIALS_FILE,
                SCOPES
            )
            creds = flow.run_local_server(port=0)

        with open(TOKEN_FILE, "wb") as f:
            pickle.dump(creds, f)

    service = build("gmail", "v1", credentials=creds)
    print("✓ Gmail API authenticated")
    return service


# ==============================
# LABEL HELPERS
# ==============================

def get_label_id(service, label_name):
    """
    Look up a Gmail label ID by name.
    Gmail API needs IDs, not names.
    Returns None if label doesn't exist.
    """
    result = service.users().labels().list(userId="me").execute()
    for label in result.get("labels", []):
        if label["name"] == label_name:
            return label["id"]
    return None


def get_all_label_ids(service):
    """
    Returns a dict of {label_name: label_id} for all job alert labels + processed label.
    Warns if any label is missing from Gmail.
    """
    result = service.users().labels().list(userId="me").execute()
    all_labels = {l["name"]: l["id"] for l in result.get("labels", [])}

    label_ids = {}
    for name in JOB_ALERT_LABELS + [PROCESSED_LABEL]:
        if name in all_labels:
            label_ids[name] = all_labels[name]
        else:
            print(f"  [WARNING] Gmail label not found: '{name}' — create it in Gmail first")

    return label_ids


# ==============================
# FETCH MESSAGES
# ==============================

def fetch_unprocessed_messages(service, label_ids, dry_run=False):
    """
    Fetch all messages that have a Job_Alerts_* label but NOT Job_Processed.

    Args:
        service:    Gmail API service object
        label_ids:  dict from get_all_label_ids()
        dry_run:    if True, fetch messages but do NOT apply Job_Processed label

    Returns:
        list of (message_id, msg_object, label_name) tuples
    """
    processed_id = label_ids.get(PROCESSED_LABEL)
    if not processed_id:
        raise ValueError(f"Label '{PROCESSED_LABEL}' not found in Gmail. Create it first.")

    all_messages = []

    for label_name in JOB_ALERT_LABELS:
        label_id = label_ids.get(label_name)
        if not label_id:
            continue

        print(f"\nFetching: {label_name}")

        # Query: has this label AND does NOT have Job_Processed
        messages = _list_messages(service, label_id, exclude_label_id=processed_id)
        print(f"  Found {len(messages)} unprocessed messages")

        for msg_meta in messages:
            msg_id = msg_meta["id"]
            msg_obj = _get_message_as_email(service, msg_id)
            if msg_obj:
                all_messages.append((msg_id, msg_obj, label_name))

    return all_messages


def _list_messages(service, label_id, exclude_label_id=None, max_results=500):
    """
    List all message IDs with the given label, excluding another label.
    Handles Gmail API pagination automatically.
    """
    messages = []
    query = f"-label:{PROCESSED_LABEL}" if exclude_label_id else ""
    page_token = None

    while True:
        result = service.users().messages().list(
            userId="me",
            labelIds=[label_id],
            q=query,
            maxResults=min(max_results, 500),
            pageToken=page_token
        ).execute()

        batch = result.get("messages", [])
        messages.extend(batch)

        page_token = result.get("nextPageToken")
        if not page_token or len(messages) >= max_results:
            break

    return messages


def _get_message_as_email(service, msg_id):
    """
    Fetch a Gmail message as raw MIME bytes and parse it into an email.message object.
    This gives us the same object our existing parsers already work with.
    """
    try:
        raw = service.users().messages().get(
            userId="me",
            id=msg_id,
            format="raw"
        ).execute()

        raw_bytes = base64.urlsafe_b64decode(raw["raw"])
        msg = BytesParser(policy=policy.default).parsebytes(raw_bytes)
        return msg

    except Exception as e:
        print(f"  [ERROR] Could not fetch message {msg_id}: {e}")
        return None


# ==============================
# MARK AS PROCESSED
# ==============================

def mark_as_processed(service, msg_id, label_ids, dry_run=False):
    """
    Apply the Job_Processed label to a Gmail message.
    In dry_run mode, prints what would happen but does nothing.
    """
    processed_id = label_ids.get(PROCESSED_LABEL)
    if not processed_id:
        print(f"  [WARNING] Cannot mark processed — label ID not found")
        return

    if dry_run:
        print(f"  [DRY RUN] Would apply '{PROCESSED_LABEL}' to message {msg_id}")
        return

    try:
        service.users().messages().modify(
            userId="me",
            id=msg_id,
            body={"addLabelIds": [processed_id]}
        ).execute()
    except Exception as e:
        print(f"  [ERROR] Could not label message {msg_id}: {e}")


def mark_batch_as_processed(service, msg_ids, label_ids, dry_run=False):
    """
    Apply Job_Processed label to a list of message IDs.
    Uses Gmail batch modify for efficiency (up to 1000 at a time).
    """
    processed_id = label_ids.get(PROCESSED_LABEL)
    if not processed_id:
        return

    if dry_run:
        print(f"\n[DRY RUN] Would mark {len(msg_ids)} messages as '{PROCESSED_LABEL}'")
        return

    # Gmail batchModify supports up to 1000 IDs at once
    chunk_size = 1000
    for i in range(0, len(msg_ids), chunk_size):
        chunk = msg_ids[i:i + chunk_size]
        try:
            service.users().messages().batchModify(
                userId="me",
                body={
                    "ids": chunk,
                    "addLabelIds": [processed_id]
                }
            ).execute()
            print(f"  ✓ Marked {len(chunk)} messages as '{PROCESSED_LABEL}'")
        except Exception as e:
            print(f"  [ERROR] Batch label failed: {e}")
