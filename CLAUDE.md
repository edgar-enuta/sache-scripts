# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Email-to-Excel automation tool that reads unread emails from an IMAP server, extracts structured data using regex patterns defined in `field_config.yaml`, and exports the results to timestamped Excel files. Processed emails are flagged to prevent reprocessing.

## Commands

```bash
# Install dependencies
pip install -r requirements.txt

# Run the tool
python main.py

# Run tests
python tests/tests.py
python tests/test_enhanced_features.py
```

## Architecture

Four modules with clear separation of concerns, orchestrated by `main.py`:

- **main.py** — Workflow orchestrator with transactional safety. On failure, deletes partial Excel files and unflags any flagged emails.
- **imap.py** — IMAP connection management and email retrieval. Searches for UNSEEN+UNFLAGGED emails, filters by subject pattern from config. Uses global `mail` connection object.
- **extract_fields.py** — Regex-based field extraction. Concatenates email subject+body, applies patterns from `field_config.yaml`, and adds system-generated columns (OrderNo, EmailDate).
- **excel.py** — Creates new timestamped Excel files (never appends). Column order: system columns first (from config order), then extras alphabetically.

**Data flow:** IMAP fetch → filter by subject pattern → extract fields via regex → export to Excel → flag emails as processed.

## Configuration

- **field_config.yaml** — Defines regex patterns and Excel column names. Fields without a `pattern` key are system-generated columns (e.g., `order_number`, `email_date`).
- **.env** — Required credentials: `EMAIL`, `EMAIL_PASSWORD`, `IMAP_SERVER`, `IMAP_PORT` (default 993), `EXCEL_PATH` (default `output/data.xlsx`).

## Key Design Decisions

- Timestamped filenames (`YYYY-MM-DD_HH-MM-SS_filename.xlsx`) prevent data loss from overwrites.
- Email flagging happens only after successful Excel creation (transactional guarantee).
- IMAP `\Flagged` flag is used to track processed emails across runs.
- Date parsing uses multiple fallback strategies (RFC 2822 variants → raw string → current time).
