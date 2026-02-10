# Exam Scheduler

Google Apps Script web app for scheduling and managing exams, using Google Sheets as the database.

## Tech Stack

- **Backend:** Google Apps Script (`code.gs`)
- **Frontend:** Single-file HTML/CSS/JS (`index.html`)
- **Database:** Google Sheets (auto-created, ID stored in Script Properties)
- **Excel Parsing:** SheetJS (xlsx) loaded via CDN for bulk upload

## Project Structure

```
code.gs       - Server-side Google Apps Script (all backend logic)
index.html    - Client-side SPA (HTML + CSS + JS, all inline)
```

There is no build step, package manager, or test framework. Deployment is done via the Apps Script editor (Deploy > Web app).

## Sheet Schema

The "Exams" sheet has columns: `ID | Exam Name | Date | Start Time | End Time | Venue | Class | Laptops Needed`

- IDs are timestamp-based (`Date.now()`)
- Dates stored as `YYYY-MM-DD` strings
- Times stored as `HH:MM` (24-hour) strings

## Backend Functions (`code.gs`)

| Function | Purpose |
|---|---|
| `doGet()` | Serves the HTML web app |
| `getSpreadsheet()` | Gets or auto-creates the backing spreadsheet |
| `getExamsSheet()` | Returns the "Exams" sheet |
| `getExams()` | Returns all exams as plain objects (with `JSON.parse(JSON.stringify())` for GAS serialization) |
| `saveExam(examData)` | Saves a single exam after clash checking |
| `deleteExam(examId)` | Deletes one exam by ID |
| `deleteAllExams()` | Deletes all exam rows (keeps header) |
| `checkForClashes(newExam)` | Checks venue/date/time overlaps against existing exams |
| `checkAvailability(date, startTime, endTime, seatsRequired)` | Returns all venues with availability status |
| `validateBulkExams(examDataArray)` | Validates a batch: field checks + clash detection (existing + intra-batch) |
| `saveBulkExams(examDataArray)` | Batch-inserts validated exams via `setValues()` |
| `getVenueCapacities()` | Returns the `VENUE_CAPACITIES` config object |
| `parseTimeToMinutes(timeString)` | Converts `HH:MM` to minutes since midnight |
| `getOverlapType(...)` | Classifies overlap as exact/starts-during/ends-during/encompasses |
| `migrateToEndTimeSchema()` | One-time migration to add End Time column |

## Frontend Features (`index.html`)

Four tabs (SPA pattern using `display: none`/`block`):

1. **Check Exam Venue** - Links to an external venue-checking tool
2. **Schedule Exam** - Form to create a single exam + bulk upload via drag-and-drop
3. **View Scheduled Exams** - Table listing all exams with delete actions
4. **Calendar View** - Month/week views with color-coded venue badges and exam popups

### Bulk Upload Flow
1. User drops `.xlsx`/`.csv` file
2. SheetJS parses client-side with flexible column name matching
3. Client validates formats (dates, times, required fields)
4. "Validate" button calls `validateBulkExams()` server-side for clash detection
5. "Schedule" button calls `saveBulkExams()` for valid exams only

### Key Frontend Functions
- `parseDate()` - Handles YYYY-MM-DD, MM/DD/YYYY, DD/MM/YYYY, Excel serial numbers
- `parseTime12To24()` - Converts 12-hour, 24-hour, and Excel decimal times to HH:MM
- `loadExams()` - Fetches exams via `google.script.run.getExams()` and updates all views
- Calendar rendering: `renderMonthView()`, `renderWeekView()`, `positionWeekExams()`

## Venue Configuration

Venues and capacities are defined in `VENUE_CAPACITIES` (code.gs, line ~85). Frontend venue `<select>` options and `VENUE_COLORS` map are maintained separately in `index.html`.

## Key Patterns

See [.claude/docs/architectural_patterns.md](.claude/docs/architectural_patterns.md) for detailed architectural patterns including:
- Client-server communication via `google.script.run`
- Response object conventions
- Clash detection algorithm
- Data serialization for GAS compatibility
