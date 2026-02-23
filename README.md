# Exam Scheduler

A Google Apps Script web app for scheduling and managing exams, using Google Sheets as the database.

## Overview

Exam Scheduler lets staff schedule individual or bulk exam sessions, view them on a calendar, and check venue availability — all through a web interface backed by a Google Sheet.

## Features

- Schedule exams with venue, class, date, and time
- Calendar view with colour-coded venues
- Clash detection (prevents double-booking a venue at overlapping times)
- Venue availability checker (shows capacity and conflicts)
- Bulk upload via Excel/CSV (up to 50 exams per batch, with per-row validation)
- Export schedule to Excel

## Tech Stack

- **Backend:** Google Apps Script (server-side logic + Google Sheets as DB)
- **Frontend:** Vanilla HTML/CSS/JS (single `index.html`, no framework)
- **Library:** [SheetJS v0.20.0](https://sheetjs.com/) for Excel parsing/export (loaded via CDN)

## Setup & Deployment

1. Open [Google Apps Script](https://script.google.com) and create a new project.
2. Copy the contents of `code.gs` into the default `Code.gs` file.
3. Create a new HTML file named `index` and paste the contents of `index.html` into it.
4. Click **Deploy > New deployment**, choose **Web app**, set access to your organisation, and click **Deploy**.
5. On first load, `initializeApp()` runs automatically and creates the Google Sheet.

> The Sheet ID is stored in Script Properties and persists across deployments.

## Usage

The app has three tabs:

| Tab | Description |
|-----|-------------|
| **Schedule Exam** | Add a single exam or upload a bulk Excel/CSV file |
| **Calendar View** | Browse scheduled exams by month; click an event to delete it |
| **Check Exam Venue** | Enter a date and time range to see which venues are free |

## Data Schema

Each exam is stored as a row in the Google Sheet:

```
ID | Exam Name | Date | Start Time | End Time | Venue | Class | Laptops Needed
```

Dates are stored as `YYYY-MM-DD` strings; times as `HH:MM` (24-hour).

## Development Notes

- **No local dev server** — all code runs inside Google Apps Script.
- **Backend debugging:** Add `Logger.log()` calls and check **Executions** in the Apps Script dashboard.
- **Frontend debugging:** Use browser DevTools; the client logs extensively via `console.log`.
- **Manual testing:** Run `testGetExams()` or `initializeApp()` directly from the Apps Script editor.
