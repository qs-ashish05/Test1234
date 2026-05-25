# Outlook Email Search Macros

A collection of VBA macros for Microsoft Excel that search Outlook for emails by subject line or email address, with results written back to dedicated Excel sheets. Built for analysts, MDM teams, and anyone who needs to audit large volumes of mail traffic without manually digging through Outlook.

All macros use **late binding**, so no Outlook reference setup is required. Just drop the code into a module and run.

---

## Table of Contents

- [Features](#features)
- [Prerequisites](#prerequisites)
- [Setup](#setup)
- [Macros](#macros)
  - [1. SearchEmailsBySubject](#1-searchemailsbysubject)
  - [2. SearchEmailsByEMailId](#2-searchemailsbyemailid)
  - [3. SearchAllEmailsByEMailId](#3-searchallemailsbyemailid)
  - [4. SearchInboxBySubject](#4-searchinboxbysubject)
- [Shared Helpers](#shared-helpers)
- [Performance Notes](#performance-notes)
- [Troubleshooting](#troubleshooting)
- [License](#license)

---

## Features

- Case-insensitive exact-match searches on email subject lines
- Exact-match searches on recipient and sender SMTP addresses
- Recursive folder scanning (parent folder + all subfolders)
- DASL-based date filtering via `Items.Restrict` for fast scans on large mailboxes
- Cleans Exchange Legacy DN strings (e.g. `/o=ExchangeLabs/...`) and extracts plain SMTP addresses using regex
- Handles distribution lists, Exchange users, and external SMTP senders
- Writes "No data found" rows (highlighted grey) for inputs that produced zero matches, so nothing silently disappears
- Per-row time-window control where applicable

---

## Prerequisites

- Microsoft Excel with VBA enabled (Office 2016 or later recommended)
- Microsoft Outlook installed and configured with at least one mail profile
- Macros must be enabled in the workbook (`.xlsm` format)

No external libraries or references required.

---

## Setup

1. Open your Excel workbook and press `Alt + F11` to open the VBA editor.
2. `Insert → Module` and paste the macro code into the new module.
3. Make sure your workbook has a sheet named **`Sheet1`** — this is where all inputs go.
4. Save the workbook as a macro-enabled file (`.xlsm`).
5. Run a macro via `Alt + F8` → select the macro name → `Run`.

The output sheets (`Sheet2`, `Sheet3`, `Sheet5`, `Sheet6`) are created automatically if they don't exist.

---

## Macros

### 1. SearchEmailsBySubject

Search the **Sent Items** folder (and subfolders) for emails whose subject line exactly matches the input.

**Use case:** You sent a series of named emails (e.g. approval requests, vendor confirmations) and need to audit who they went to.

**Input — Sheet1**

| Column | Description |
| --- | --- |
| A | Email subject line (one per row) |

**Output — Sheet2**

| Column | Description |
| --- | --- |
| A | Input Subject |
| B | Matched Subject |
| C | To |
| D | CC |
| E | BCC |
| F | Date & Time |
| G | Folder |

**Matching:** Case-insensitive exact match on subject. No date restriction.

---

### 2. SearchEmailsByEMailId

Search the **Sent Items** folder (and subfolders) for emails sent to specific recipients. Matches across To, CC, and BCC.

**Use case:** You need to pull every email you sent to a vendor, customer, or internal team alias.

**Input — Sheet1**

| Column | Description |
| --- | --- |
| A | Recipient email address (one per row) |

**Output — Sheet3**

| Column | Description |
| --- | --- |
| A | Input Email ID |
| B | Sent To |
| C | CC |
| D | Subject |
| E | Date & Time |
| F | Folder |

**Matching:** Case-insensitive exact match on recipient SMTP address. No date restriction.

---

### 3. SearchAllEmailsByEMailId

Search **both Inbox and Sent Items** for all emails related to specific addresses — both directions (sent to them, received from them). Restricted to the last N days for speed.

**Use case:** You need the full two-way email history with a contact or group alias, e.g. for an audit or a handover document.

**Input — Sheet1**

| Cell | Description |
| --- | --- |
| A1:Axx | Email addresses to search (one per row) |
| B1 | Number of days to look back (e.g. `30`) |

**Output — Sheet5**

| Column | Description |
| --- | --- |
| A | Input Email ID |
| B | Direction (`Sent` or `Received`) |
| C | From |
| D | To |
| E | CC |
| F | Subject |
| G | Date & Time |
| H | Folder |

**Matching:** Case-insensitive exact match.
- For Inbox scans → matches sender SMTP address (handles Exchange users and distribution lists).
- For Sent Items scans → matches any recipient SMTP address (To, CC, or BCC).

**Performance:** Uses `Items.Restrict` with a DASL date filter (`urn:schemas:httpmail:datereceived` / `datesent`), so only items inside the time window are loaded into memory. Output is auto-sorted by Input Email ID then by Date & Time descending.

---

### 4. SearchInboxBySubject

Search the **Inbox** folder (and subfolders) for emails whose subject line exactly matches the input, with a **per-subject** time window. Each subject can have its own lookback period.

**Use case:** Inbox is too large for an unrestricted scan, and different subjects have different relevance windows (e.g. monthly reminders → last 7 days; quarterly approvals → last 90 days).

**Input — Sheet1**

| Column | Description |
| --- | --- |
| A | Email subject line |
| B | Number of days to look back (per row) |

| | A | B |
| --- | --- | --- |
| 1 | `Vendor Code Creation Request` | 30 |
| 2 | `Monthly Rent Reminder` | 7 |
| 3 | `RRSR Approval Pending` | 90 |

**Output — Sheet6**

| Column | Description |
| --- | --- |
| A | Input Subject |
| B | Last N Days |
| C | From |
| D | To |
| E | CC |
| F | Matched Subject |
| G | Date & Time |
| H | Folder |