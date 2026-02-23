# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Google Apps Script web app for scheduling and managing exams, using Google Sheets as the database. Two-file codebase with no build step, package manager, or test framework.

## Development & Deployment

- **No local dev server** — code runs inside Google Apps Script. Copy `code.gs` and `index.html` into the Apps Script editor.
- **Deploy:** Apps Script editor > Deploy > Web app
- **Debug backend:** Use `Logger.log()` and check Executions in the Apps Script dashboard
- **Debug frontend:** Browser DevTools console (client-side JS logs extensively with `console.log`)
- **Test backend manually:** Run `testGetExams()` or `initializeApp()` from the Apps Script editor

## Architecture

```
code.gs       - Server-side Google Apps Script (all backend logic)
index.html    - Client-side SPA (HTML + CSS + JS, all inline in one file)
```

**Database:** Google Sheets with auto-creation. Sheet ID stored in Script Properties. Schema: `ID | Exam Name | Date | Start Time | End Time | Venue | Class | Laptops Needed`

**Client-server communication:** All calls use `google.script.run.withSuccessHandler().withFailureHandler().backendFunction(args)`. GAS cannot return complex objects, so backend uses `JSON.parse(JSON.stringify(data))` before returning.

**Frontend SPA pattern:** Three tabs (Schedule Exam, Calendar View, Check Exam Venue) toggled via `.active` class on both button and content div. No router. Calendar tab triggers data load if exams array is empty.

**External dependencies:** SheetJS v0.20.0 (xlsx) loaded via CDN at ~line 877 of `index.html`. No other frameworks or libraries.

## Key Backend Functions

- **Setup:** `doGet()`, `getSpreadsheet()`, `getExamsSheet()`, `initializeApp()`
- **CRUD:** `saveExam()`, `getExams()`, `deleteExam()`, `deleteAllExams()`
- **Clash/Availability:** `checkForClashes()`, `checkAvailability()`, `getVenueCapacities()`
- **Bulk:** `validateBulkExams()`, `saveBulkExams()`
- **Helpers:** `parseTimeToMinutes()`, `getOverlapType()`
- **Utilities:** `getSpreadsheetUrl()`, `testGetExams()`, `migrateToEndTimeSchema()`

`migrateToEndTimeSchema()` adds an End Time column to older spreadsheets (uses `DEFAULT_DURATION_MINUTES = 180`).

## Critical Conventions

**Response objects** — Backend functions return `{ success: true/false, message: "..." }`. Exceptions:
- `checkForClashes()` returns `{ hasClash: true/false, clashes: [...] }` (different convention)
- `checkAvailability()` returns `{ success: true, data: { venues: [{ name, capacity, available, hasEnoughSeats, conflictingExams }] } }`

**Clash detection** — Overlap formula: `(StartA < EndB) AND (EndA > StartB)`. Back-to-back is allowed (exact boundary match is not a clash). Clashes are per **venue + date** only.

**Validation rules:**
- End time must be after start time (enforced in form, bulk upload client-side, and server-side)
- Bulk upload max 50 exams per batch
- Required fields: name, date, time, endTime, venue, class
- Optional: laptopsNeeded (defaults to 0)

**Bulk upload three-phase pattern:**
1. Client-side parse (SheetJS) with flexible column name matching (normalized headers)
2. Server-side validate (`validateBulkExams`) — checks existing DB + intra-batch clashes
3. Server-side save (`saveBulkExams`) — batch insert via `setValues()`

Each phase is a separate `google.script.run` call.

**Data formats:** Dates as `YYYY-MM-DD` strings, times as `HH:MM` 24-hour strings, IDs as `Date.now()` timestamps.

## Venue Configuration (Keep in Sync)

Venues are defined in **three separate places** that must stay synchronized:
1. `VENUE_CAPACITIES` object in `code.gs` (~line 85) — capacity data
2. `<select id="venue">` options in `index.html` (~line 633) — form dropdown
3. `VENUE_COLORS` object in `index.html` (~line 1781) — calendar color mapping

Note: The dropdown includes venues (Design Studio B, Design Studio C) not in `VENUE_CAPACITIES`, and `VENUE_CAPACITIES` includes venues (ISH) not in the dropdown. This is a known inconsistency.

## Detailed Patterns

See [.claude/docs/architectural_patterns.md](.claude/docs/architectural_patterns.md) for:
- Response object examples for all endpoints
- Clash detection algorithm details
- Data serialization patterns for GAS
- Defensive date/time parsing strategies
- Column name mapping for bulk upload
