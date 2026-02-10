# Architectural Patterns

## 1. Client-Server Communication (`google.script.run`)

All frontend-to-backend calls use the GAS client API:

```javascript
google.script.run
  .withSuccessHandler((response) => { /* handle response */ })
  .withFailureHandler((error) => { /* handle error.message */ })
  .backendFunctionName(args);
```

- Success handler receives the return value of the server function
- Failure handler receives an Error object (only `.message` is reliably available)
- UI buttons are disabled during calls and re-enabled in both handlers

## 2. Response Object Convention

Backend functions return consistent response objects:

```javascript
// Success
{ success: true, message: "...", ...additionalData }

// Failure
{ success: false, message: "Error description" }

// Clash-specific (saveExam)
{ success: false, clash: true, message: "...", clashes: [...] }

// Clash check (checkForClashes)
{ hasClash: true/false, clashes: [...] }

// Bulk validation (validateBulkExams)
{ results: [...], summary: { totalRows, validRows, invalidRows, clashingRows } }
```

Note: `checkForClashes` uses `hasClash` instead of `success` — this is a different convention used only for internal clash checking.

## 3. Clash Detection (Time Overlap Algorithm)

Used in `checkForClashes`, `checkAvailability`, and `validateBulkExams`:

```javascript
// Two events overlap if: (StartA < EndB) AND (EndA > StartB)
// Back-to-back is allowed (exact end/start match is NOT a clash)
if (newStart < existingEnd && newEnd > existingStart) {
  // CLASH
}
```

Clashes are checked per **venue + date** combination only. Same time in different venues is allowed.

`validateBulkExams` checks clashes in two passes:
1. Each new exam against all existing exams in the database
2. Each new exam against earlier exams within the same batch

## 4. Tab-Based SPA Navigation

Frontend uses a simple show/hide pattern:

```javascript
// All tabs: display:none by default, .active sets display:block
// Tab buttons toggle .active class on both button and content div
// Data is loaded/refreshed when switching to "view" or "calendar" tabs
```

No router or URL-based navigation — purely DOM class toggling.

## 5. Data Serialization for GAS Compatibility

Google Apps Script cannot return complex objects (Date objects, custom classes) through `google.script.run`. The codebase uses:

```javascript
return JSON.parse(JSON.stringify(exams));
```

This pattern appears in `getExams()` and `checkAvailability()` to strip prototypes and convert Dates to strings, ensuring the data survives the GAS client-server boundary.

## 6. Defensive Date/Time Parsing

The codebase handles multiple input formats due to Google Sheets and Excel inconsistencies:

**Server-side (`getExams`):**
- Dates: checks `instanceof Date` → uses `Utilities.formatDate()`, otherwise casts to string
- Times: checks `instanceof Date` → extracts HH:mm, checks `typeof string`, falls back to `String()`

**Client-side (bulk upload):**
- `parseDate()`: handles YYYY-MM-DD, MM/DD/YYYY, DD/MM/YYYY, Excel serial numbers (30000-60000 range)
- `parseTime12To24()`: handles HH:MM (24h), H:MM AM/PM (12h), Excel decimal times (0-1 range)

## 7. Bulk Operation Pattern

The bulk upload follows a three-phase pattern:

```
Phase 1: Client-side parse (SheetJS)
  → Parse Excel/CSV → Map columns by normalized header names → Client-validate formats

Phase 2: Server-side validate (validateBulkExams)
  → Validate required fields → Check clashes vs existing DB → Check clashes within batch
  → Return per-row results with status

Phase 3: Server-side save (saveBulkExams)
  → Only valid exams sent → Batch insert via sheet.getRange().setValues()
  → Unique IDs: baseTimestamp + index
```

Each phase is a separate `google.script.run` call, giving the user a chance to review validation results before committing.

## 8. Column Name Mapping (Bulk Upload)

Headers are normalized by lowercasing and stripping all non-alphanumeric characters, then matched against known aliases:

```javascript
// "Teaching Group" → "teachinggroup" → matches alias for "class" field
// "Laptops/Candidate" → "laptopscandidate" → matches alias for "laptopsNeeded" field
```

This allows flexible spreadsheet formats while maintaining backward compatibility with old column names.
