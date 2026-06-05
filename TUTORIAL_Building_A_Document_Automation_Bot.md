# Tutorial: Building a Document Automation Bot from Scratch

> **Goal:** By the end of this tutorial you will have built—and understood every line of—a production document-automation bot that:
> - Watches an Adobe Sign widget for new signed agreements
> - Downloads and parses the attached PDF and Excel files
> - Validates the customer data against a master Excel spreadsheet
> - Updates the master spreadsheet with confirmed opt-out decisions
> - Emails different teams depending on the outcome
> - Runs automatically on a schedule via Windows Task Scheduler
>
> **Who this is for:** Developers who know a little Python but have never built a real-world integration system. Every concept is explained from first principles.

---

## Table of Contents

1. [Why This Project Exists](#1-why-this-project-exists)
2. [How to Read This Tutorial](#2-how-to-read-this-tutorial)
3. [Setting Up Your Development Environment](#3-setting-up-your-development-environment)
4. [Python Fundamentals](#4-python-fundamentals)
5. [Project Structure and the Module System](#5-project-structure-and-the-module-system)
6. [Configuration: Environment Variables and .env Files](#6-configuration-environment-variables-and-env-files)
7. [Understanding REST APIs](#7-understanding-rest-apis)
8. [OAuth 2.0 Authentication](#8-oauth-20-authentication)
9. [Building the Adobe Sign API Client](#9-building-the-adobe-sign-api-client)
10. [Working with Excel Files](#10-working-with-excel-files)
11. [Reading PDF Files](#11-reading-pdf-files)
12. [Data Validation: The Attachment Validator](#12-data-validation-the-attachment-validator)
13. [Sending Email via Microsoft Graph API](#13-sending-email-via-microsoft-graph-api)
14. [The Response Processor Pipeline](#14-the-response-processor-pipeline)
15. [State Management](#15-state-management)
16. [The Widget Monitor: The Main Loop](#16-the-widget-monitor-the-main-loop)
17. [Logging](#17-logging)
18. [Scheduling with Windows Task Scheduler](#18-scheduling-with-windows-task-scheduler)
19. [Debugging Techniques](#19-debugging-techniques)
20. [Security Considerations](#20-security-considerations)
21. [Extending the Bot](#21-extending-the-bot)

---

## 1. Why This Project Exists

### The Business Problem

A telecommunications company is switching off its legacy PSTN (Public Switched Telephone Network) over the next few years. Hundreds of reseller partners need to decide, circuit by circuit, whether to opt their customers in or out of the new service.

The old way of doing this: email chains, manually updated spreadsheets, and hours of admin work per day.

### The Automated Solution

Instead of emails, partners are sent a **digital form** (an Adobe Sign widget). They fill in their customer details and sign electronically. The bot:

1. Detects the completed form within minutes of it being submitted
2. Downloads the signed PDF and any attached Excel files automatically
3. Checks whether the data matches what the company holds on record
4. Updates the company's master spreadsheet with the partner's decisions
5. Sends the right email to the right team (data alert, partial opt-out notification, or opt-out confirmation)

No human has to open an email, download a file, or update a spreadsheet. The whole workflow is automated.

### What You Will Learn

Building this bot touches almost every skill a working Python developer needs:

- Making API calls to third-party services
- Authenticating securely with OAuth 2.0
- Reading and writing Excel and PDF files
- Validating data against a reference dataset
- Sending HTML emails
- Running background tasks on a schedule
- Structuring a multi-file Python project
- Logging, error handling, and debugging

---

## 2. How to Read This Tutorial

Each section introduces a concept, shows the underlying theory, then walks through the actual code. You will build the solution incrementally—each section adds a new piece until the whole system works.

**Conventions used:**

```
# Comments in code blocks explain what the line does
```

```python
# Python code you should actually write
def example():
    return "Hello"
```

> **Note:** Callouts like this highlight important points or common mistakes.

**Do the exercises.** Theory without practice does not produce skill. Every section has a small exercise to check your understanding before you move on.

---

## 3. Setting Up Your Development Environment

### 3.1 Install Python

Go to [https://www.python.org/downloads/](https://www.python.org/downloads/) and download the latest Python 3.12 or 3.13 installer for Windows.

During installation, tick **"Add Python to PATH"**. This lets you run `python` from any terminal.

Verify it worked:

```powershell
python --version
# Should print: Python 3.13.x
```

### 3.2 Install Visual Studio Code

Download from [https://code.visualstudio.com/](https://code.visualstudio.com/). Install the **Python extension** (search for "Python" in the Extensions panel, install the one from Microsoft).

### 3.3 Create a Project Folder

```powershell
mkdir C:\Projects\doc_bot
cd C:\Projects\doc_bot
```

### 3.4 Create a Virtual Environment

A **virtual environment** is an isolated copy of Python for your project. It means the packages you install for this project don't interfere with any other Python project on your machine.

```powershell
python -m venv .venv
```

This creates a `.venv` folder. Every time you open a new terminal to work on this project, activate the virtual environment:

```powershell
.venv\Scripts\Activate.ps1
```

Your prompt will change to show `(.venv)` at the start.

> **Why this matters:** Without a venv, all projects share the same Python packages. One project needing `requests==2.28` and another needing `requests==2.31` would conflict. Virtual environments solve this.

### 3.5 Install Dependencies

Create a file called `requirements.txt` in your project folder:

```
requests==2.32.3
python-dotenv==1.0.1
pandas==2.2.2
openpyxl==3.1.5
xlwings==0.33.3
pdfminer.six==20231228
msal==1.31.0
```

Install everything at once:

```powershell
pip install -r requirements.txt
```

### 3.6 Exercise

1. Create the project folder and virtual environment as described.
2. Run `pip list` and confirm all the packages are installed.
3. Open the folder in VS Code (`code .`) and confirm VS Code is using the `.venv` interpreter (bottom-left corner of VS Code shows the Python version).

---

## 4. Python Fundamentals

This section is a fast-paced refresher of the Python you will need. If you are already comfortable with Python, skim this section and move on.

### 4.1 Variables and Data Types

```python
# A string — text
name = "Acme Resellers Ltd"

# An integer — whole number
count = 42

# A float — decimal number
confidence = 0.95

# A boolean — True or False
is_valid = True

# None — the absence of a value
result = None
```

### 4.2 Collections

```python
# A list — ordered, can change, allows duplicates
emails = ["alice@example.com", "bob@example.com"]
emails.append("carol@example.com")   # add to end
first = emails[0]                    # "alice@example.com"

# A dictionary — key → value pairs
agreement = {
    "id": "CBJCHBCAABAAHgM-E9Mf",
    "status": "SIGNED",
    "name": "Acme Opt-Out Form"
}
status = agreement["status"]         # "SIGNED"
agreement["processed"] = True        # add new key

# A set — unique values, unordered
seen_ids = {"id1", "id2", "id3"}
seen_ids.add("id4")
"id1" in seen_ids                    # True

# A tuple — like a list but immutable (cannot change)
point = (10, 20)
x, y = point                         # unpack
```

### 4.3 Control Flow

```python
# if / elif / else
if status == "SIGNED":
    print("Process this agreement")
elif status == "OUT_FOR_SIGNATURE":
    print("Still waiting")
else:
    print("Unknown status:", status)

# for loop over a list
for email in emails:
    print("Sending to:", email)

# for loop over a range of numbers
for i in range(5):          # 0, 1, 2, 3, 4
    print(i)

# while loop
attempts = 0
while attempts < 3:
    attempts += 1
    print("Attempt", attempts)

# List comprehension — compact way to build a list
upper_emails = [e.upper() for e in emails]
```

### 4.4 Functions

```python
# Define a function
def send_email(to_address, subject, body):
    """Send an email. Returns True if successful."""
    # ... implementation ...
    return True

# Call it
success = send_email("alice@example.com", "Hello", "Hi Alice!")

# Default parameter values
def send_email(to_address, subject, body, cc=None):
    pass

# *args — accepts any number of positional arguments
def log(*messages):
    for msg in messages:
        print(msg)

# **kwargs — accepts any number of keyword arguments
def create_record(**fields):
    return fields   # returns a dict

record = create_record(name="Acme", status="active")
```

### 4.5 Classes

```python
# A class is a blueprint for creating objects
class Agreement:
    def __init__(self, agreement_id, partner_name):
        # __init__ runs when you create a new Agreement
        self.agreement_id = agreement_id
        self.partner_name = partner_name
        self.processed = False

    def mark_processed(self):
        self.processed = True

    def __repr__(self):
        # Controls how the object prints
        return f"Agreement({self.agreement_id}, {self.partner_name})"

# Create instances
agr = Agreement("CBJCHBCAABAAHgM", "Acme Ltd")
agr.mark_processed()
print(agr.processed)   # True
```

### 4.6 Exception Handling

```python
# try / except — handle errors gracefully
try:
    response = requests.get("https://api.example.com/data")
    response.raise_for_status()   # raises an exception if status >= 400
    data = response.json()
except requests.exceptions.HTTPError as exc:
    print("HTTP error:", exc)
except requests.exceptions.ConnectionError:
    print("Could not connect to the API")
except Exception as exc:
    # Catch-all — always log the actual error
    print("Unexpected error:", exc)
finally:
    # This block always runs, even if an exception occurred
    print("Finished API call")
```

> **Common mistake:** Writing `except:` (bare except with no exception type) catches *everything*, including `KeyboardInterrupt` and `SystemExit`. Always name the exception or use `except Exception`.

### 4.7 f-strings

```python
name = "Acme"
count = 42
# f-strings embed expressions inside {}
message = f"Partner {name} has {count} circuits to migrate."
# Equivalent: "Partner Acme has 42 circuits to migrate."

# You can use expressions
total = f"Total rows: {count * 2}"

# Format numbers
price = 1234.5678
formatted = f"£{price:,.2f}"   # "£1,234.57"
```

### 4.8 Working with Files

```python
# Writing a text file
with open("output.txt", "w", encoding="utf-8") as f:
    f.write("Hello, World!\n")

# Reading a text file
with open("output.txt", "r", encoding="utf-8") as f:
    content = f.read()

# Reading line by line
with open("output.txt", "r", encoding="utf-8") as f:
    for line in f:
        print(line.strip())
```

> **The `with` statement:** Automatically closes the file even if an exception occurs. Always use it for file operations.

### 4.9 Imports

```python
# Import a standard library module
import os
import json
import datetime

# Import specific things from a module
from pathlib import Path
from datetime import datetime, timedelta

# Import a third-party package
import requests
import pandas as pd     # "as pd" creates an alias
```

### 4.10 Exercise

1. Write a function `count_unique_partners(records)` that takes a list of dictionaries (each with a `"partner_name"` key) and returns the number of unique partner names.
2. Test it with sample data.
3. Add exception handling: if a record is missing the `"partner_name"` key, skip it and log a warning.

---

## 5. Project Structure and the Module System

### 5.1 Why Structure Matters

A single `bot.py` file will quickly become unmanageable. Good structure means:
- Each file has a single, clear responsibility
- Files can be tested and reused independently
- Team members can work on different parts without stepping on each other

### 5.2 Our Project Layout

```
doc_bot/
├── .venv/                     # virtual environment (never edit this)
├── .env                       # secrets and settings (never commit to git)
├── requirements.txt
├── widget_monitor.py          # main entry point — runs the loop
├── response_processor.py      # orchestrates one agreement
├── attachment_validator.py    # validates attachments and updates master
├── adobe_sign_client.py       # Adobe Sign API wrapper
├── graph_mailbox_check.py     # Microsoft Graph API email
├── config.py                  # loads .env settings into constants
├── state/                     # folder for runtime state files
│   └── processed_agreements.json
├── logs/                      # folder for log files
│   └── bot_2026-06-04.log
└── input/                     # temporary working files
```

### 5.3 The Module System

Every `.py` file is a **module**. You import code from one module into another:

```python
# In response_processor.py:
from attachment_validator import validate_attachment, flush_master_writes
from adobe_sign_client import AdobeSignClient
from config import WIDGET_ID, UPDATE_MASTER_DATA
```

Python finds these modules because they are in the same folder. For larger projects you would use **packages** (folders with an `__init__.py` file), but for our bot a flat structure is fine.

### 5.4 Circular Imports

Avoid importing module A from module B if module B is already imported by module A. This creates a circular dependency and Python will throw an `ImportError`. The solution is to put shared constants and utilities into a dedicated `config.py` that nothing else imports from.

---

## 6. Configuration: Environment Variables and .env Files

### 6.1 Why Not Hardcode Secrets?

Never write passwords, API keys, or URLs directly in your code:

```python
# WRONG — do not do this
API_KEY = "sk-abc123supersecret"
```

Reasons:
- If you commit the file to git, the secret is public forever
- Different environments (dev, test, production) need different values
- Rotating a secret means editing code and redeploying

### 6.2 The .env File

Create a file called `.env` in your project root. It is never committed to git (add it to `.gitignore`):

```
# Adobe Sign
ADOBE_SIGN_CLIENT_ID=your_client_id_here
ADOBE_SIGN_CLIENT_SECRET=your_client_secret_here
ADOBE_SIGN_REFRESH_TOKEN=your_refresh_token_here
WIDGET_ID=CBJCHBCAABAAHgM-E9MfQ4zVKGAAPqPXbkWE7lMnPp7T

# Microsoft Graph
GRAPH_TENANT_ID=your_tenant_id
GRAPH_CLIENT_ID=your_client_id
GRAPH_CLIENT_SECRET=your_client_secret
EMAIL_SENDER=BRProdUKCOP03@example.com

# Master data
CIRCUIT_MASTER_PATH=C:\Users\you\OneDrive\customer_data_master\Master.xlsx
CIRCUIT_MASTER_SHEET_NAME=Sams MPF Sheet

# Feature flags
UPDATE_MASTER_DATA=true
DATA_ALERT_EMAIL=admin@example.com,manager@example.com
```

### 6.3 Loading .env in Python

```python
# config.py
import os
from pathlib import Path
from dotenv import load_dotenv

# Load .env file into environment variables
load_dotenv()

# Now read them — os.environ["KEY"] raises an error if missing
# os.getenv("KEY") returns None if missing
# os.getenv("KEY", "default") returns default if missing

ADOBE_SIGN_CLIENT_ID     = os.environ["ADOBE_SIGN_CLIENT_ID"]
ADOBE_SIGN_CLIENT_SECRET = os.environ["ADOBE_SIGN_CLIENT_SECRET"]
ADOBE_SIGN_REFRESH_TOKEN = os.environ["ADOBE_SIGN_REFRESH_TOKEN"]
WIDGET_ID                = os.environ["WIDGET_ID"]

GRAPH_TENANT_ID          = os.environ["GRAPH_TENANT_ID"]
GRAPH_CLIENT_ID          = os.environ["GRAPH_CLIENT_ID"]
GRAPH_CLIENT_SECRET      = os.environ["GRAPH_CLIENT_SECRET"]
EMAIL_SENDER             = os.environ["EMAIL_SENDER"]

CIRCUIT_MASTER_PATH      = Path(os.environ["CIRCUIT_MASTER_PATH"])
CIRCUIT_MASTER_SHEET_NAME = os.getenv("CIRCUIT_MASTER_SHEET_NAME", "Sheet1")

# Boolean settings need special handling — env vars are always strings
UPDATE_MASTER_DATA = os.getenv("UPDATE_MASTER_DATA", "false").lower() == "true"

# Comma-separated email lists become Python lists
DATA_ALERT_EMAIL = [
    e.strip()
    for e in os.getenv("DATA_ALERT_EMAIL", "").split(",")
    if e.strip()
]
```

> **Note:** `"false".lower() == "true"` evaluates to `False`. `"true".lower() == "true"` evaluates to `True`. This is how you parse boolean environment variables safely.

### 6.4 Exercise

1. Create a `.env` file with a fake `API_KEY` value.
2. Write a `config.py` that loads and prints it.
3. Add `API_KEY` to `.gitignore` — actually, add `.env` itself.

---

## 7. Understanding REST APIs

### 7.1 What is a REST API?

An API (Application Programming Interface) lets two programs talk to each other over the internet. REST APIs use standard HTTP — the same protocol your browser uses.

When you make an API call you send an **HTTP request** and receive an **HTTP response**.

### 7.2 HTTP Methods

| Method | Meaning | Example |
|--------|---------|---------|
| `GET` | Retrieve data | Get a list of agreements |
| `POST` | Create or submit data | Submit a new form |
| `PUT` | Replace a resource | Update a record |
| `PATCH` | Partially update a resource | Change one field |
| `DELETE` | Delete a resource | Remove a draft |

### 7.3 HTTP Status Codes

| Code range | Meaning |
|------------|---------|
| 200–299 | Success |
| 400–499 | Your request was wrong (client error) |
| 500–599 | The server had a problem (server error) |

Common codes:
- `200 OK` — success
- `201 Created` — new resource was created
- `400 Bad Request` — malformed request
- `401 Unauthorized` — not authenticated
- `403 Forbidden` — authenticated but not allowed
- `404 Not Found` — resource doesn't exist
- `429 Too Many Requests` — rate limited, slow down
- `500 Internal Server Error` — server bug

### 7.4 Making API Calls with requests

```python
import requests

# Simple GET request
response = requests.get("https://api.example.com/agreements")

# Check status
print(response.status_code)   # 200

# Parse the JSON body
data = response.json()

# POST with a JSON body
payload = {"name": "New Agreement", "status": "DRAFT"}
response = requests.post(
    "https://api.example.com/agreements",
    json=payload,           # automatically sets Content-Type: application/json
    headers={"Authorization": "Bearer my_token_here"}
)

# Always check for errors
response.raise_for_status()   # raises requests.HTTPError if status >= 400
```

### 7.5 Headers and Authentication

APIs require you to prove who you are. The most common way is the **Authorization header**:

```python
headers = {
    "Authorization": "Bearer eyJhbGci...",   # your access token
    "Content-Type": "application/json",
    "Accept": "application/json"
}
response = requests.get(url, headers=headers)
```

### 7.6 Query Parameters

Some API calls filter data using query parameters (the `?key=value` part of a URL):

```python
params = {
    "status": "SIGNED",
    "startDate": "2026-01-01",
    "pageSize": 100
}
response = requests.get("https://api.example.com/agreements", params=params)
# requests builds the URL: .../agreements?status=SIGNED&startDate=2026-01-01&pageSize=100
```

### 7.7 Exercise

Using the free API at `https://jsonplaceholder.typicode.com`:

1. Fetch all posts (`GET /posts`) and print the title of the first 5.
2. Fetch a single post by ID (`GET /posts/1`).
3. Create a new post (`POST /posts`) with a title and body of your choosing.
4. Handle the case where the server returns a `404`.

---

## 8. OAuth 2.0 Authentication

### 8.1 Why Not Just a Password?

Giving your bot your username and password to an API is dangerous. If the bot is compromised, the attacker has your credentials for everything. OAuth 2.0 solves this by issuing short-lived **access tokens**.

### 8.2 Key Concepts

- **Client** — your application (the bot)
- **Resource Server** — the API you want to call (Adobe Sign)
- **Authorization Server** — the server that issues tokens (Adobe's auth server)
- **Access Token** — a short-lived credential (usually expires in 1–4 hours)
- **Refresh Token** — a long-lived credential used to get new access tokens without re-authenticating

### 8.3 Refresh Token Flow (Adobe Sign)

Adobe Sign uses this flow because the bot acts on behalf of a user (a human who signed in once and approved the bot's access):

```
1. Human signs in to Adobe Sign and approves the bot's permissions
   → Adobe gives back a refresh token (valid for months/years)
   → Store this in .env

2. Each time the bot starts (or the access token expires):
   Bot sends: POST /oauth/v2/refresh
       grant_type=refresh_token
       refresh_token=<stored refresh token>
       client_id=<your app's client ID>
       client_secret=<your app's client secret>
   
   Adobe replies: { "access_token": "eyJ...", "expires_in": 3600 }

3. Bot uses access_token in every subsequent API call header:
   Authorization: Bearer eyJ...

4. When access_token expires, bot requests a new one using refresh_token again
```

```python
# adobe_sign_client.py — token management
import time
import requests

class AdobeSignClient:
    BASE_URL = "https://api.na1.adobesign.com/api/rest/v6"
    TOKEN_URL = "https://api.na1.adobesign.com/oauth/v2/refresh"

    def __init__(self, client_id, client_secret, refresh_token):
        self.client_id = client_id
        self.client_secret = client_secret
        self.refresh_token = refresh_token
        self._access_token = None
        self._token_expiry = 0   # Unix timestamp

    def _ensure_token(self):
        """Get a new access token if the current one has expired."""
        if time.time() < self._token_expiry - 60:
            return   # still valid (with 60s safety margin)

        response = requests.post(self.TOKEN_URL, data={
            "grant_type": "refresh_token",
            "refresh_token": self.refresh_token,
            "client_id": self.client_id,
            "client_secret": self.client_secret,
        })
        response.raise_for_status()
        token_data = response.json()
        self._access_token = token_data["access_token"]
        self._token_expiry = time.time() + token_data["expires_in"]

    def _headers(self):
        self._ensure_token()
        return {"Authorization": f"Bearer {self._access_token}"}
```

### 8.4 Client Credentials Flow (Microsoft Graph)

For the email-sending part the bot acts as itself — no user involved. This uses the **client credentials** flow:

```
1. Register the app in Azure Active Directory
   → Get: tenant_id, client_id, client_secret

2. Bot requests a token directly:
   POST https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token
       grant_type=client_credentials
       client_id=<your app client ID>
       client_secret=<your app client secret>
       scope=https://graph.microsoft.com/.default

3. Microsoft replies: { "access_token": "eyJ...", "expires_in": 3599 }

4. Bot uses token to call Graph API endpoints
```

The `msal` library handles all of this automatically:

```python
import msal

def get_graph_token(tenant_id, client_id, client_secret):
    app = msal.ConfidentialClientApplication(
        client_id=client_id,
        client_credential=client_secret,
        authority=f"https://login.microsoftonline.com/{tenant_id}"
    )
    result = app.acquire_token_for_client(
        scopes=["https://graph.microsoft.com/.default"]
    )
    if "access_token" not in result:
        raise RuntimeError(f"Token request failed: {result.get('error_description')}")
    return result["access_token"]
```

### 8.5 Exercise

1. Read the Adobe Sign OAuth documentation (search "Adobe Sign refresh token").
2. Explain in one paragraph why client credentials flow is appropriate for sending emails but refresh token flow is required for reading agreements.

---

## 9. Building the Adobe Sign API Client

### 9.1 Design Principles

A good API client class:
- Hides all authentication details from callers
- Raises meaningful exceptions when things go wrong
- Retries on transient failures (network blips, rate limits)
- Logs what it is doing for debugging

### 9.2 Custom Exception

```python
class AdobeSignError(Exception):
    """Raised when the Adobe Sign API returns an error."""
    def __init__(self, message, status_code=None, response_body=None):
        super().__init__(message)
        self.status_code = status_code
        self.response_body = response_body
```

Having a custom exception lets callers catch specifically Adobe Sign errors:

```python
try:
    agreements = client.get_widget_agreements(WIDGET_ID)
except AdobeSignError as exc:
    logger.error("Adobe Sign error: %s (HTTP %s)", exc, exc.status_code)
```

### 9.3 Retry Logic

Networks are unreliable. APIs get temporarily overloaded. A good bot retries failed requests:

```python
import time
import logging

logger = logging.getLogger(__name__)

def _request_with_retry(self, method, path, max_retries=3, **kwargs):
    """Make an HTTP request, retrying on 429 (rate limit) or 5xx (server error)."""
    url = f"{self.BASE_URL}{path}"
    
    for attempt in range(max_retries):
        try:
            response = requests.request(
                method, url,
                headers=self._headers(),
                timeout=30,
                **kwargs
            )
            
            if response.status_code == 429:
                # Rate limited — wait and retry
                wait = int(response.headers.get("Retry-After", 5))
                logger.warning("Rate limited. Waiting %ds before retry.", wait)
                time.sleep(wait)
                continue
            
            if response.status_code >= 500:
                # Server error — retry with exponential backoff
                if attempt < max_retries - 1:
                    wait = 2 ** attempt   # 1s, 2s, 4s...
                    logger.warning("Server error %s. Retrying in %ds.", 
                                   response.status_code, wait)
                    time.sleep(wait)
                    continue
            
            response.raise_for_status()
            return response
            
        except requests.exceptions.ConnectionError:
            if attempt < max_retries - 1:
                time.sleep(2 ** attempt)
                continue
            raise
    
    raise AdobeSignError(f"All {max_retries} attempts failed for {method} {path}")
```

### 9.4 Key API Methods

```python
def get_widget_agreements(self, widget_id, cursor=None):
    """
    Get a page of agreements submitted to a widget.
    Adobe Sign returns up to 100 at a time with a cursor for pagination.
    """
    params = {"widgetId": widget_id, "pageSize": 100}
    if cursor:
        params["cursor"] = cursor
    
    response = self._request_with_retry("GET", "/agreements", params=params)
    return response.json()   # {"userAgreementList": [...], "page": {"nextCursor": "..."}}


def get_agreement_info(self, agreement_id):
    """Get metadata for a single agreement (status, dates, participant emails)."""
    response = self._request_with_retry("GET", f"/agreements/{agreement_id}")
    return response.json()


def get_agreement_documents(self, agreement_id):
    """List the documents (PDF + attachments) for a signed agreement."""
    response = self._request_with_retry(
        "GET", f"/agreements/{agreement_id}/documents"
    )
    return response.json()   # {"documents": [{"id": "...", "name": "...", ...}]}


def download_document(self, agreement_id, document_id):
    """Download a document's binary content (PDF or Excel bytes)."""
    response = self._request_with_retry(
        "GET", f"/agreements/{agreement_id}/documents/{document_id}",
        stream=True
    )
    return response.content   # bytes
```

### 9.5 Exercise

1. Implement the full `AdobeSignClient` class with token management, retry logic, and the four methods above.
2. Write a small test script that creates the client and prints the name of every document on a test agreement.

---

## 10. Working with Excel Files

### 10.1 Why Excel is Tricky

Excel files are deceptively complex. They contain:
- Multiple sheets
- Formulas that recalculate on open
- External links to other workbooks
- Merged cells, hidden rows, formatted numbers
- Potentially thousands of rows

The bot needs to both **read** (quickly, at startup) and **write** (carefully, without corrupting the file).

### 10.2 Reading with pandas + calamine

pandas with the `calamine` engine is extremely fast for reading large files because calamine is written in Rust and does not load Excel's full object model:

```python
import pandas as pd

def load_master_data(path, sheet_name):
    """Load the master spreadsheet into a DataFrame."""
    df = pd.read_excel(
        path,
        sheet_name=sheet_name,
        engine="calamine",    # fast Rust-based reader
        dtype=str,            # read everything as text — avoids type conversion surprises
    )
    # Clean up column names: strip whitespace, lowercase
    df.columns = [str(c).strip().lower() for c in df.columns]
    # Fill NaN (empty cells) with empty string so comparisons work
    df = df.fillna("")
    return df
```

> **Why `dtype=str`?** Excel often stores account numbers as numbers (losing leading zeros) or as text. Forcing everything to string means `"00123"` stays `"00123"` rather than becoming `123`.

### 10.3 Caching the Master Data

Reading 20,000 rows takes a second or two. Do it once and cache it:

```python
import os

_master_cache: pd.DataFrame | None = None
_master_mtime: float = 0.0

def get_master_data():
    """Return cached master data, reloading if the file has changed."""
    global _master_cache, _master_mtime
    
    current_mtime = os.path.getmtime(CIRCUIT_MASTER_PATH)
    if _master_cache is not None and current_mtime == _master_mtime:
        return _master_cache
    
    _master_cache = load_master_data(CIRCUIT_MASTER_PATH, CIRCUIT_MASTER_SHEET_NAME)
    _master_mtime = current_mtime
    return _master_cache
```

### 10.4 Writing: The Queue-and-Flush Pattern

Writing to a large Excel file is expensive. Rather than saving after every change, we **queue** all changes and **flush** (write) them all at once.

```python
# The queue — accumulates changes during processing
_pending_updates: list[tuple[int, str, str]] = []
# Each tuple: (excel_row_number, column_name, new_value)

def queue_update(excel_row: int, column: str, value: str):
    """Queue a cell update to be written later."""
    _pending_updates.append((excel_row, column, value))

def flush_master_writes():
    """Apply all queued updates to the Excel file and save."""
    if not _pending_updates:
        return
    
    updates = list(_pending_updates)
    _pending_updates.clear()
    
    try:
        try:
            _flush_via_xlwings(updates)     # fast Windows COM-based writer
        except ImportError:
            _flush_via_openpyxl(updates)    # fallback pure-Python writer
    except Exception as exc:
        logger.error("flush_master_writes failed: %s", exc, exc_info=True)
```

### 10.5 Writing with openpyxl

openpyxl reads and writes the full Excel XML format. It is slower than xlwings but works on any OS:

```python
import openpyxl

def _flush_via_openpyxl(updates: list[tuple[int, str, str]]):
    wb = openpyxl.load_workbook(LOCAL_WORKING_COPY, keep_links=False)
    
    # Find the sheet — case-insensitive search
    sheet_name_lower = CIRCUIT_MASTER_SHEET_NAME.lower()
    ws = next(
        (wb[s] for s in wb.sheetnames if s.lower() == sheet_name_lower),
        None
    )
    if ws is None:
        raise ValueError(f"Sheet '{CIRCUIT_MASTER_SHEET_NAME}' not found")
    
    # Build a map from column name to column index
    header_row = [cell.value for cell in ws[1]]
    col_index = {
        str(h).strip().lower(): i + 1   # openpyxl uses 1-based column indices
        for i, h in enumerate(header_row)
        if h is not None
    }
    
    # Apply updates
    for excel_row, col_name, value in updates:
        col_idx = col_index.get(col_name.lower())
        if col_idx is None:
            logger.warning("Column not found: %s", col_name)
            continue
        ws.cell(row=excel_row, column=col_idx).value = value
    
    wb.save(LOCAL_WORKING_COPY)
    logger.info("openpyxl: saved %d cell updates", len(updates))
```

### 10.6 Writing with xlwings

xlwings uses Windows COM automation — the same mechanism as a macro. It is faster for large files and preserves formulas and formatting:

```python
import xlwings as xw

def _flush_via_xlwings(updates: list[tuple[int, str, str]]):
    app = xw.App(visible=False)
    try:
        wb = app.books.open(str(LOCAL_WORKING_COPY))
        ws = wb.sheets[CIRCUIT_MASTER_SHEET_NAME]
        
        # Get column positions from the header row
        header_range = ws.range("A1").expand("right")
        headers = [str(c.value).strip().lower() if c.value else "" 
                   for c in header_range]
        col_index = {name: i + 1 for i, name in enumerate(headers) if name}
        
        for excel_row, col_name, value in updates:
            col_idx = col_index.get(col_name.lower())
            if col_idx is None:
                continue
            ws.range(excel_row, col_idx).value = value
        
        wb.save()
        wb.close()
        logger.info("xlwings: saved %d cell updates", len(updates))
    finally:
        app.quit()
```

### 10.7 Exercise

1. Create a small test Excel file with columns: `id`, `name`, `status`.
2. Load it with pandas and print the rows where `status` is empty.
3. Use openpyxl to set the `status` column of row 2 to `"active"` and save.

---

## 11. Reading PDF Files

### 11.1 Why PDFs Are Hard

A PDF is not a document — it is a set of drawing instructions. Text is placed at specific X,Y coordinates on the page. There is no concept of "rows" or "columns" unless you infer them from position.

When Adobe Sign generates a completed form PDF, the field labels and their values are placed at fixed positions. We extract text boxes and use their X position to determine which column of a table they belong to.

### 11.2 Extracting Text Boxes with pdfminer

pdfminer.six parses PDF files into a hierarchy of layout objects. We want `LTTextBox` objects, which correspond to visually distinct blocks of text:

```python
from pdfminer.high_level import extract_pages
from pdfminer.layout import LAParams, LTTextBox, LTTextLine

def extract_text_boxes(pdf_path):
    """
    Extract all text boxes from a PDF with their position.
    Returns a list of (x0, y0, text) tuples.
    """
    params = LAParams(
        char_margin=0.1,    # small value: keeps columns separate
        line_margin=0.5,
        word_margin=0.1,
    )
    boxes = []
    for page_layout in extract_pages(pdf_path, laparams=params):
        for element in page_layout:
            if isinstance(element, LTTextBox):
                text = element.get_text().strip()
                if text:
                    boxes.append((element.x0, element.y0, text))
    return boxes
```

### 11.3 Assigning Columns by X Position

A form table has columns at specific X positions. By sorting text boxes and grouping them by X range, we can reconstruct the table:

```python
def assign_columns(boxes, column_boundaries):
    """
    Assign each text box to a column based on its X position.
    column_boundaries: list of (col_name, x_min, x_max) tuples
    """
    rows = {}   # y_position → {col_name: text}
    
    for x0, y0, text in boxes:
        col_name = None
        for name, x_min, x_max in column_boundaries:
            if x_min <= x0 < x_max:
                col_name = name
                break
        
        if col_name:
            # Round Y to nearest 10 to group text on the same row
            row_key = round(y0 / 10) * 10
            if row_key not in rows:
                rows[row_key] = {}
            rows[row_key][col_name] = text
    
    # Sort rows from top to bottom (higher Y = higher on page)
    sorted_rows = sorted(rows.items(), key=lambda item: -item[0])
    return [row_data for _, row_data in sorted_rows]
```

### 11.4 Handling Merged Text

pdfminer sometimes merges text from adjacent cells into one box, separated by newlines:

```
"n\nvulnerable customer"
```

The first line is the field value; the rest is from the neighbouring column. Guard against this:

```python
value = text_box_content.strip()
if "\n" in value:
    value = value.split("\n")[0].strip()   # take only the first line
```

### 11.5 Exercise

1. Create a simple PDF (or use any existing PDF).
2. Use pdfminer to extract all text boxes and print their X, Y coordinates and text.
3. Observe how changing `char_margin` from `0.1` to `2.0` affects which boxes are merged together.

---

## 12. Data Validation: The Attachment Validator

### 12.1 The Validation Problem

The partner submits a form claiming they have certain circuits. We need to verify that those circuits actually belong to that partner in our master data, and then record their decision.

### 12.2 The Result Object

Use a dataclass to represent the validation outcome. A dataclass is like a class but with automatic `__init__`, `__repr__`, and `__eq__` methods:

```python
from dataclasses import dataclass, field

@dataclass
class AttachmentValidationResult:
    is_valid: bool = False
    
    # Extracted from the form
    reseller_name: str = ""
    account_number: str = ""
    opt_out_full: bool = False
    
    # Counts from validation
    total_circuits: int = 0
    matched_circuits: int = 0
    partial_circuits: int = 0   # circuits where include=n (partial opt-out)
    
    # Details
    unmatched_rows: list[dict] = field(default_factory=list)
    errors: list[str] = field(default_factory=list)
```

> **`field(default_factory=list)`** is required for mutable defaults in dataclasses. You cannot write `unmatched_rows: list = []` because that would share the same list across all instances.

### 12.3 The Validation Loop

```python
def validate_attachment(pdf_path, excel_path, agreement_id):
    """Validate a partner's submission against the master data."""
    result = AttachmentValidationResult()
    
    # Step 1: Load the master data
    master_df = get_master_data()
    
    # Step 2: Read the partner's submitted Excel file
    try:
        submitted_df = pd.read_excel(excel_path, dtype=str).fillna("")
    except Exception as exc:
        result.errors.append(f"Could not read attachment: {exc}")
        return result
    
    # Step 3: Extract data from the signed PDF
    pdf_data = extract_pdf_fields(pdf_path)
    result.reseller_name = pdf_data.get("reseller_name", "")
    result.account_number = pdf_data.get("account_number", "")
    
    # Step 4: Validate each row in the submitted data
    for _, row in submitted_df.iterrows():
        result.total_circuits += 1
        circuit_id = row.get("circuit_id", "").strip()
        
        # Find matching row in master
        master_rows = master_df[master_df["circuit_id"] == circuit_id]
        
        if master_rows.empty:
            result.unmatched_rows.append({"circuit_id": circuit_id})
            continue
        
        result.matched_circuits += 1
        
        # Check the include/exclude decision
        include_val = row.get("include in migration y/n", "").strip().lower()
        if "\n" in include_val:
            include_val = include_val.split("\n")[0].strip()
        
        if include_val in {"n", "no"}:
            result.partial_circuits += 1
            # Queue master update
            excel_row = master_rows.index[0] + 2   # +1 for 0→1 index, +1 for header
            queue_update(excel_row, "include in migration y/n", "n")
            queue_update(excel_row, "reason", row.get("reason", ""))
    
    result.is_valid = len(result.errors) == 0
    return result
```

### 12.4 Exercise

1. Create a sample master DataFrame with 10 rows.
2. Create a sample submission DataFrame with 5 rows, some matching and some not.
3. Run the validation loop and check the counts are correct.

---

## 13. Sending Email via Microsoft Graph API

### 13.1 Why Graph API?

The Graph API lets you send email as a specific Office 365 mailbox without needing SMTP credentials. This is more secure (uses OAuth 2.0) and does not require firewall rules for SMTP ports.

### 13.2 Azure App Registration

Before you can use Graph API you need to register an application in Azure Active Directory:

1. Go to `portal.azure.com` → Azure Active Directory → App registrations → New registration
2. Note down the **Application (client) ID** and **Directory (tenant) ID**
3. Go to Certificates & secrets → New client secret → Note it down immediately (it only shows once)
4. Go to API permissions → Add a permission → Microsoft Graph → Application permissions → `Mail.Send`
5. Click "Grant admin consent"

### 13.3 Sending an Email

```python
import msal
import requests

def send_email(
    tenant_id, client_id, client_secret,
    sender, to_addresses, subject, html_body,
    attachments=None
):
    """Send an email via Microsoft Graph API."""
    
    # 1. Get access token
    app = msal.ConfidentialClientApplication(
        client_id=client_id,
        client_credential=client_secret,
        authority=f"https://login.microsoftonline.com/{tenant_id}"
    )
    token_result = app.acquire_token_for_client(
        scopes=["https://graph.microsoft.com/.default"]
    )
    if "access_token" not in token_result:
        raise RuntimeError(f"Auth failed: {token_result.get('error_description')}")
    
    token = token_result["access_token"]
    
    # 2. Build the message payload
    message = {
        "subject": subject,
        "body": {"contentType": "HTML", "content": html_body},
        "toRecipients": [
            {"emailAddress": {"address": addr}} for addr in to_addresses
        ]
    }
    
    # 3. Add attachments if any
    if attachments:
        import base64
        message["attachments"] = []
        for filename, file_bytes in attachments:
            message["attachments"].append({
                "@odata.type": "#microsoft.graph.fileAttachment",
                "name": filename,
                "contentBytes": base64.b64encode(file_bytes).decode()
            })
    
    # 4. Send via Graph API
    url = f"https://graph.microsoft.com/v1.0/users/{sender}/sendMail"
    response = requests.post(
        url,
        headers={
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json"
        },
        json={"message": message}
    )
    response.raise_for_status()   # Graph returns 202 Accepted on success
    return True
```

### 13.4 Building HTML Emails

Plain text emails work but HTML emails are easier to read. A minimal HTML email:

```python
def build_alert_email(partner_name, unmatched_count, unmatched_rows):
    rows_html = "".join(
        f"<tr><td>{row['circuit_id']}</td><td>{row.get('reason','')}</td></tr>"
        for row in unmatched_rows
    )
    return f"""
    <html><body>
    <h2>Data Alert: Unmatched Circuits</h2>
    <p>Partner <strong>{partner_name}</strong> submitted {unmatched_count} 
       circuit(s) that could not be found in the master data.</p>
    <table border="1" cellpadding="4">
        <tr><th>Circuit ID</th><th>Reason</th></tr>
        {rows_html}
    </table>
    </body></html>
    """
```

### 13.5 Exercise

1. Register a test app in Azure (or use an existing one if available).
2. Write a script that sends yourself a test HTML email with your name in the body.

---

## 14. The Response Processor Pipeline

### 14.1 The Pipeline Pattern

A **pipeline** is a sequence of steps where the output of each step feeds into the next. Our response processor:

1. Gets the agreement details from Adobe Sign
2. Downloads the PDF and Excel attachments
3. Validates the attachments
4. Updates the master sheet (if valid)
5. Sends the appropriate email
6. Records the agreement as processed

Each step is wrapped in its own `try/except` so a failure in one step does not silently skip the others.

### 14.2 The Output Dictionary

```python
def process_agreement(agreement_id: str, client: AdobeSignClient) -> dict:
    """Process one signed agreement end-to-end."""
    
    # The output dict accumulates results from every step
    output = {
        "agreement_id": agreement_id,
        "success": False,
        "errors": [],
        "reseller_name": None,
        "account_number": None,
        "opt_out_full": False,
        "partial_circuits": 0,
        "total_circuits": 0,
        "matched_circuits": 0,
    }
    
    # --- Step 1: Get agreement info ---
    try:
        info = client.get_agreement_info(agreement_id)
        output["reseller_name"] = info.get("name", "Unknown")
    except Exception as exc:
        output["errors"].append(f"Step 1 failed: {exc}")
        return output   # can't continue without basic info
    
    # --- Step 2: Download documents ---
    pdf_path = None
    xlsx_path = None
    try:
        docs = client.get_agreement_documents(agreement_id)
        for doc in docs["documents"]:
            content = client.download_document(agreement_id, doc["id"])
            if doc["name"].endswith(".pdf"):
                pdf_path = save_temp_file(content, f"{agreement_id}.pdf")
            elif doc["name"].endswith(".xlsx"):
                xlsx_path = save_temp_file(content, f"{agreement_id}.xlsx")
    except Exception as exc:
        output["errors"].append(f"Step 2 failed: {exc}")
        return output
    
    # --- Step 3: Validate ---
    try:
        validation = validate_attachment(pdf_path, xlsx_path, agreement_id)
        output["opt_out_full"] = validation.opt_out_full
        output["partial_circuits"] = validation.partial_circuits
        output["total_circuits"] = validation.total_circuits
        output["errors"].extend(validation.errors)
    except Exception as exc:
        output["errors"].append(f"Step 3 failed: {exc}")
    
    # --- Step 4: Send appropriate email ---
    try:
        if output["errors"]:
            send_data_alert_email(output)
        elif output["opt_out_full"]:
            send_opt_out_email(output)
        elif output["partial_circuits"] > 0:
            send_partial_opt_out_email(output)
    except Exception as exc:
        output["errors"].append(f"Step 4 (email) failed: {exc}")
    
    output["success"] = len(output["errors"]) == 0
    return output
```

### 14.3 Exercise

1. Trace through the pipeline with a hypothetical agreement that has 3 partial opt-out circuits and no errors.
2. Which email gets sent?
3. What would `output` look like after the pipeline completes?

---

## 15. State Management

### 15.1 The Problem

The bot runs every 15 minutes. Each run it fetches all agreements from the widget. Without state, it would reprocess every agreement on every run.

### 15.2 The JSON State File

A simple JSON file records which agreement IDs have been processed:

```python
import json
from pathlib import Path

STATE_FILE = Path("state/processed_agreements.json")

def load_state() -> dict:
    """Load the persistent state file."""
    if not STATE_FILE.exists():
        return {"processed_ids": [], "last_run": None}
    try:
        return json.loads(STATE_FILE.read_text(encoding="utf-8"))
    except Exception:
        return {"processed_ids": [], "last_run": None}

def save_state(state: dict):
    """Save state to disk atomically."""
    STATE_FILE.parent.mkdir(parents=True, exist_ok=True)
    # Write to a temp file then rename — prevents partial writes on crash
    tmp = STATE_FILE.with_suffix(".tmp")
    tmp.write_text(json.dumps(state, indent=2), encoding="utf-8")
    tmp.replace(STATE_FILE)   # atomic on most OS

def mark_processed(agreement_id: str):
    """Record an agreement as processed."""
    state = load_state()
    if agreement_id not in state["processed_ids"]:
        state["processed_ids"].append(agreement_id)
    save_state(state)

def is_processed(agreement_id: str) -> bool:
    """Check if an agreement has already been processed."""
    return agreement_id in load_state()["processed_ids"]
```

> **Atomic save:** Writing to a temp file and then renaming is important. If the bot crashes mid-write, the original file is intact. A direct write could leave a partial, corrupt JSON file.

### 15.3 Exercise

1. Write the state file functions as described.
2. Add a `get_unprocessed(agreement_ids)` function that returns only IDs not yet in state.

---

## 16. The Widget Monitor: The Main Loop

### 16.1 What It Does

The widget monitor is the entry point. It:

1. Creates the Adobe Sign client
2. Fetches all agreements from the widget (handling pagination)
3. Filters out already-processed agreements
4. Processes each new agreement
5. Flushes the master Excel writes
6. Sleeps or exits (if run by Task Scheduler)

### 16.2 Pagination

Adobe Sign returns agreements in pages of up to 100. To get all of them, follow the `nextCursor`:

```python
def get_all_widget_agreements(client, widget_id):
    """Fetch every agreement for a widget, handling pagination."""
    all_agreements = []
    cursor = None
    
    while True:
        page = client.get_widget_agreements(widget_id, cursor=cursor)
        all_agreements.extend(page.get("userAgreementList", []))
        
        next_cursor = page.get("page", {}).get("nextCursor")
        if not next_cursor:
            break   # no more pages
        cursor = next_cursor
    
    return all_agreements
```

### 16.3 The Main Loop

```python
def run_monitor():
    """Main entry point — process all new signed agreements."""
    logger.info("Widget monitor starting...")
    
    client = AdobeSignClient(
        client_id=ADOBE_SIGN_CLIENT_ID,
        client_secret=ADOBE_SIGN_CLIENT_SECRET,
        refresh_token=ADOBE_SIGN_REFRESH_TOKEN,
    )
    
    # Fetch all agreements
    agreements = get_all_widget_agreements(client, WIDGET_ID)
    logger.info("Found %d total agreements", len(agreements))
    
    # Filter to signed, unprocessed ones
    new_agreements = [
        a for a in agreements
        if a.get("status") == "SIGNED" and not is_processed(a["id"])
    ]
    logger.info("%d new agreements to process", len(new_agreements))
    
    for agreement in new_agreements:
        agreement_id = agreement["id"]
        logger.info("Processing agreement: %s", agreement_id)
        
        try:
            result = process_agreement(agreement_id, client)
            if result["success"]:
                mark_processed(agreement_id)
                logger.info("✅  Agreement %s processed successfully", agreement_id)
            else:
                logger.warning("⚠️  Agreement %s had errors: %s", 
                               agreement_id, result["errors"])
        except Exception as exc:
            logger.error("💥 Unexpected error processing %s: %s", 
                         agreement_id, exc, exc_info=True)
    
    # Write all queued Excel updates in one save
    if UPDATE_MASTER_DATA:
        try:
            flush_master_writes()
        except Exception as exc:
            logger.error("flush_master_writes failed: %s", exc, exc_info=True)
    
    logger.info("Widget monitor finished.")

if __name__ == "__main__":
    run_monitor()
```

---

## 17. Logging

### 17.1 Why Not print()?

`print()` is fine for tiny scripts. For a production bot you need:
- Log levels (DEBUG, INFO, WARNING, ERROR, CRITICAL)
- Timestamps
- The module that emitted the message
- Log files you can review after the fact
- The ability to turn verbose logging on/off without changing code

### 17.2 Configuring Logging

```python
import logging
import logging.handlers
from datetime import datetime
from pathlib import Path

def setup_logging():
    """Configure logging to both console and daily rotating file."""
    log_dir = Path("logs")
    log_dir.mkdir(exist_ok=True)
    
    # Create root logger
    root_logger = logging.getLogger()
    root_logger.setLevel(logging.DEBUG)   # capture everything
    
    # Format: timestamp | level | module | message
    formatter = logging.Formatter(
        "%(asctime)s | %(levelname)-8s | %(name)s | %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S"
    )
    
    # Console handler — INFO and above only
    console = logging.StreamHandler()
    console.setLevel(logging.INFO)
    console.setFormatter(formatter)
    root_logger.addHandler(console)
    
    # File handler — DEBUG and above, one file per day
    today = datetime.now().strftime("%Y-%m-%d")
    file_handler = logging.FileHandler(
        log_dir / f"bot_{today}.log",
        encoding="utf-8"
    )
    file_handler.setLevel(logging.DEBUG)
    file_handler.setFormatter(formatter)
    root_logger.addHandler(file_handler)
```

### 17.3 Per-Module Loggers

Each module should have its own logger named after the module:

```python
# At the top of each module
import logging
logger = logging.getLogger(__name__)
# __name__ is the module's name, e.g. "attachment_validator"

# Then use it
logger.debug("Loading master data from %s", path)
logger.info("Validation complete: %d/%d circuits matched", matched, total)
logger.warning("Column not found: %s", col_name)
logger.error("Failed to save Excel file: %s", exc, exc_info=True)
```

> **`exc_info=True`** tells the logger to include the full stack trace in the log file. Use this on `logger.error()` and `logger.exception()` calls.

### 17.4 Log Levels Guide

| Level | When to use |
|-------|-------------|
| `DEBUG` | Verbose detail for debugging (e.g., "Checking circuit XYZ123") |
| `INFO` | Normal operational events (e.g., "Processed 5 agreements") |
| `WARNING` | Something unexpected but recoverable (e.g., "Column not found, skipping") |
| `ERROR` | Something failed but the bot continues (e.g., "Email send failed") |
| `CRITICAL` | The bot cannot continue at all |

---

## 18. Scheduling with Windows Task Scheduler

### 18.1 The Batch File

Create `run_bot.bat` in your project folder:

```batch
@echo off
cd /d C:\Projects\doc_bot
call .venv\Scripts\activate.bat
python widget_monitor.py
```

### 18.2 Creating the Scheduled Task

1. Open **Task Scheduler** (search for it in the Start menu)
2. Click **Create Task** (not the basic version — use "Create Task" for full options)
3. **General tab:**
   - Name: `Document Automation Bot`
   - Run whether user is logged on or not: tick this
   - Run with highest privileges: tick this
4. **Triggers tab → New:**
   - Begin the task: On a schedule
   - Daily, starting at 06:00
   - Repeat task every 15 minutes for a duration of 18 hours
5. **Actions tab → New:**
   - Program/script: `C:\Projects\doc_bot\run_bot.bat`
   - Start in: `C:\Projects\doc_bot`
6. **Settings tab:**
   - If the task is already running, do not start a new instance

### 18.3 Exporting and Importing Tasks

Once created, export the task to XML for version control:

```powershell
Export-ScheduledTask -TaskName "Document Automation Bot" | Out-File bot_task.xml
```

To import on another machine:

```powershell
Register-ScheduledTask -Xml (Get-Content bot_task.xml | Out-String) -TaskName "Document Automation Bot"
```

---

## 19. Debugging Techniques

### 19.1 The VS Code Debugger

The VS Code Python debugger lets you pause your code and inspect every variable:

1. Click the gutter next to a line number to set a **breakpoint** (red dot appears)
2. Press **F5** or go to Run → Start Debugging
3. When execution hits your breakpoint, it pauses
4. Use the **Variables** panel to inspect variables
5. Use the **Debug Console** to evaluate expressions
6. Press **F10** (step over) or **F11** (step into) to move line by line

### 19.2 The process_single_agreement.py Script

When debugging a specific agreement, it is useful to have a script that processes just one agreement by ID, without the monitor loop:

```python
# process_single_agreement.py
import sys
from config import *
from adobe_sign_client import AdobeSignClient
from response_processor import process_agreement
from attachment_validator import flush_master_writes
from logging_setup import setup_logging

setup_logging()

if len(sys.argv) < 2:
    print("Usage: python process_single_agreement.py <agreement_id>")
    sys.exit(1)

agreement_id = sys.argv[1]
client = AdobeSignClient(ADOBE_SIGN_CLIENT_ID, ADOBE_SIGN_CLIENT_SECRET, ADOBE_SIGN_REFRESH_TOKEN)

result = process_agreement(agreement_id, client)
print("Result:", result)

if UPDATE_MASTER_DATA:
    flush_master_writes()
```

Run it:

```powershell
python process_single_agreement.py CBJCHBCAABAAHgM-E9Mf
```

### 19.3 Reading Logs

Search a log file for errors:

```powershell
Select-String -Path "logs\bot_2026-06-04.log" -Pattern "ERROR|WARNING"
```

Find all lines related to a specific agreement:

```powershell
Select-String -Path "logs\bot_2026-06-04.log" -Pattern "CBJCHBCAABAAHgM-E9Mf"
```

### 19.4 Common Issues and Solutions

| Symptom | Likely cause | Fix |
|---------|-------------|-----|
| `ModuleNotFoundError` | venv not activated or package not installed | Activate venv, run `pip install -r requirements.txt` |
| `KeyError: 'access_token'` | OAuth token request failed | Check client ID, secret, and refresh token in `.env` |
| `FileNotFoundError` for master Excel | Path in `.env` is wrong | Check `CIRCUIT_MASTER_PATH` |
| `ValueError: sheet not found` | Sheet name mismatch | Check `CIRCUIT_MASTER_SHEET_NAME` |
| `partial=0` when circuits were opted out | pdfminer newline merge | Add `\n` split guard on `include_val` |
| Emails sent to wrong people | Multiple email lists used | Check `_build_alert_recipients()` function |
| Master not updating | `UPDATE_MASTER_DATA=false` or flush silently failing | Check `.env` and add `exc_info=True` to error handler |

---

## 20. Security Considerations

### 20.1 Never Hardcode Secrets

Already covered in Section 6, but worth repeating: API keys, passwords, and tokens **never** belong in source code. Use `.env` files, environment variables, or a secrets manager.

### 20.2 Validate All External Input

Every piece of data arriving from the API, from PDF files, or from Excel files is **untrusted input**. Validate it before using it:

```python
# Bad — assumes the data is always present and correct
circuit_id = row["circuit_id"]

# Good — handles missing or unexpected values
circuit_id = str(row.get("circuit_id", "")).strip()
if not circuit_id:
    logger.warning("Skipping row with empty circuit_id")
    continue
```

### 20.3 Path Injection

If you construct file paths from user-supplied data, an attacker could traverse to sensitive files:

```python
# Bad — allows path traversal
filename = user_supplied_name
path = base_dir / filename   # could be "../../windows/system32/config/sam"

# Good — validate and restrict
import os
filename = os.path.basename(user_supplied_name)   # strips any path components
path = base_dir / filename
if not path.is_relative_to(base_dir):
    raise ValueError("Path traversal detected")
```

### 20.4 Backups Before Writes

Before modifying the master Excel file, make a backup:

```python
import shutil
from datetime import datetime

def backup_master():
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_path = LOCAL_WORKING_COPY.with_name(
        f"{LOCAL_WORKING_COPY.stem}_backup_{timestamp}.xlsx"
    )
    shutil.copy2(LOCAL_WORKING_COPY, backup_path)
    logger.info("Backup created: %s", backup_path)
```

### 20.5 Least Privilege

The Azure app registration should only have `Mail.Send` permission — not `Mail.ReadWrite`, not `Files.ReadWrite`. Give each component only the permissions it actually needs.

### 20.6 OWASP Top 10 Relevance

The OWASP Top 10 is a list of the most common web application vulnerabilities. Several apply to automation bots:

- **Injection** — don't use unsanitised data in file paths or shell commands
- **Broken Authentication** — use OAuth 2.0 properly; rotate secrets regularly
- **Sensitive Data Exposure** — log what you need, not access tokens or full customer records
- **Security Misconfiguration** — keep permissions minimal; audit `.env` contents
- **Logging & Monitoring** — log errors with enough context to detect attacks or misuse

---

## 21. Extending the Bot

Once the bot is working, here are natural next steps:

### 21.1 Web Dashboard

Build a simple web page (using Flask or FastAPI) that shows:
- Agreements processed today
- Error counts
- A list of unmatched circuits

```python
from flask import Flask, jsonify
import json

app = Flask(__name__)

@app.route("/status")
def status():
    state = json.loads(Path("state/processed_agreements.json").read_text())
    return jsonify({
        "processed_today": len(state["processed_ids"]),
        "last_run": state.get("last_run")
    })
```

### 21.2 Webhooks Instead of Polling

Instead of checking every 15 minutes, Adobe Sign can push a notification to your server the moment an agreement is signed. This requires a public HTTPS endpoint (use ngrok for testing):

```python
from flask import Flask, request

app = Flask(__name__)

@app.route("/webhook", methods=["POST"])
def webhook():
    data = request.json
    if data.get("event") == "AGREEMENT_WORKFLOW_COMPLETED":
        agreement_id = data["agreement"]["id"]
        process_agreement_async(agreement_id)
    return "", 200
```

### 21.3 Database Instead of JSON State

For large volumes, replace the JSON state file with SQLite:

```python
import sqlite3

def init_db():
    conn = sqlite3.connect("state/bot.db")
    conn.execute("""
        CREATE TABLE IF NOT EXISTS agreements (
            id TEXT PRIMARY KEY,
            processed_at TEXT,
            success INTEGER,
            errors TEXT
        )
    """)
    conn.commit()
    return conn
```

### 21.4 Automated Testing

Write tests using `pytest` to catch regressions:

```python
# test_attachment_validator.py
import pytest
import pandas as pd
from attachment_validator import validate_attachment

def test_partial_opt_out_counted_correctly():
    # Create mock data
    master_df = pd.DataFrame({
        "circuit_id": ["C001", "C002", "C003"],
        "grandparentname": ["Acme"] * 3,
        "include in migration y/n": ["y", "y", "y"]
    })
    submission_df = pd.DataFrame({
        "circuit_id": ["C001", "C002"],
        "include in migration y/n": ["n", "y"],
        "reason": ["Cost", ""]
    })
    result = run_validation(master_df, submission_df)
    assert result.partial_circuits == 1
    assert result.matched_circuits == 2
```

### 21.5 Docker

Package the bot in a Docker container so it runs the same way on any machine:

```dockerfile
FROM python:3.13-slim
WORKDIR /app
COPY requirements.txt .
RUN pip install -r requirements.txt
COPY . .
CMD ["python", "widget_monitor.py"]
```

---

## Summary

You have built a complete document automation bot from scratch. Here is what you now know how to do:

| Skill | Where it appears |
|-------|-----------------|
| Python fundamentals | Throughout |
| Virtual environments and pip | Section 3 |
| .env configuration | Section 6 |
| HTTP and REST APIs | Section 7 |
| OAuth 2.0 (refresh token + client credentials) | Section 8 |
| API client with retry logic | Section 9 |
| Reading large Excel files efficiently | Section 10 |
| Writing Excel with queue-and-flush | Section 10 |
| PDF text extraction | Section 11 |
| Data validation | Section 12 |
| Sending HTML emails via Graph API | Section 13 |
| Pipeline architecture | Section 14 |
| Atomic state management | Section 15 |
| Polling loop with pagination | Section 16 |
| Production logging | Section 17 |
| Windows Task Scheduler | Section 18 |
| Debugging with VS Code | Section 19 |
| Security best practices | Section 20 |

The bot you have built handles real-world complexity: unreliable APIs, malformed input, large Excel files, and multiple notification workflows. These same patterns — configuration from environment, retry logic, queue-and-flush, pipeline with per-step error isolation, atomic state — appear in virtually every production automation system.

---

*Document version: 1.0 — Generated 2026-06-04*
