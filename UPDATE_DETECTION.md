# Deadline Update Detection

## How It Works

The scraper now automatically detects deadline changes in two ways:

### 1. Canvas Announcements Monitoring ✅

Every time you run the scraper, it:
- Fetches the latest 50 announcements from each course
- Uses AI to identify deadline-related announcements
- Extracts any deadline changes, extensions, or new deadlines
- Highlights these updates in your output

**What gets detected from announcements:**
- Extended deadlines
- Postponed exams
- Rescheduled submissions
- New assignments announced mid-semester
- Cancelled assessments

### 2. Change Detection Between Runs ✅

The scraper compares your current run with the previous run to detect:
- **Date changes**: Deadlines that moved to different dates
- **Added deadlines**: New deadlines that weren't there before
- **Removed deadlines**: Cancelled or deleted deadlines

## Example Output

When changes are detected, your `deadlines.md` will start with a summary:

```markdown
# Assignment Deadlines - Generated 2025-01-15 14:30

## ⚠️ RECENT CHANGES DETECTED

### 📅 Date Changes:
- **CS2103T Software Engineering** - Assignment 2: ~~2025-01-20~~ → **2025-01-25**
- **EE2026 Digital Design** - Lab Report 3: ~~2025-01-22~~ → **2025-01-28**

### ✨ New Deadlines:
- **CS2107 Intro to Info Security** - Bonus Assignment (2025-02-01)

### 🗑️ Removed/Cancelled:
- **CS2040C Data Structures** - Quiz 2

---

## January 2025

### Mon, 25 Jan - **CS2103T Software Engineering**
- **[ASSIGNMENT]** Assignment 2 (15%)
  - 🔔 **UPDATE:** EXTENDED from Jan 20 to Jan 25
  - _Source: Announcement: Assignment 2 Extension_

...
```

## Files Generated

### 1. `deadlines.md`
- Human-readable format
- **Highlights changes at the top**
- Shows update markers (🔔) for modified deadlines
- Includes source attribution for announcements

### 2. `deadlines.json`
- Machine-readable format
- Contains all deadline data
- Used for change detection in next run

### 3. `changes.json` (when changes detected)
- Detailed change log
- Lists added, removed, and modified deadlines
- Useful for automated notifications

## What About Email Updates?

### Current Status: Announcements Only

Right now, the scraper monitors **Canvas announcements only**. This covers most cases because professors typically:
1. Post announcements on Canvas
2. Canvas auto-emails you the announcement

So you're already getting email notifications from Canvas itself!

### Future: Email Monitoring

For cases where professors **only** send updates via NUS Outlook (not Canvas), I can add email monitoring in the future. This would:

1. Connect to your NUS Outlook via IMAP/Microsoft Graph API
2. Scan emails from professors for deadline keywords
3. Extract deadline changes using AI
4. Merge with Canvas data

**When you might need this:**
- Professor sends deadline updates only via email
- Administrative emails about exam schedules
- Department-wide deadline announcements

**To add later (let me know if needed):**
```bash
# Add to .env
OUTLOOK_EMAIL=your_email@u.nus.edu
OUTLOOK_PASSWORD=your_password
# OR use Microsoft Graph API (recommended)
MICROSOFT_CLIENT_ID=...
MICROSOFT_CLIENT_SECRET=...
```

## Running for Updates

### First Time
```bash
python scrape_deadlines.py
```
Creates initial baseline - no changes detected.

### Subsequent Runs
```bash
python scrape_deadlines.py
```
Compares with previous run and highlights changes.

### Recommended Schedule

**Weekly runs** to catch updates:
```bash
# Monday 9 AM
0 9 * * 1 cd /path/to/scraper && python scrape_deadlines.py
```

This ensures you catch:
- Weekend announcements
- Last-minute Monday updates
- Weekly deadline adjustments

## Change Tracking Logic

The scraper identifies deadlines by `course + title`. For example:

- "CS2103T Software Engineering - Assignment 1" is tracked separately from
- "CS2103T Software Engineering - Assignment 2"

### Date Changes Detected When:
- Same course + same title
- Different date values
- Example: "CS2103T - Assignment 1" moved from Jan 15 → Jan 20

### New Deadlines Detected When:
- Course + title combination doesn't exist in previous run
- Can be truly new OR renamed assignment

### Removed Deadlines Detected When:
- Course + title from previous run not found in current run
- Likely cancelled or renamed

## Limitations

### What It CANNOT Detect:

1. **Silent syllabus updates**
   - If professor updates intro PDF without announcing
   - Re-download and re-run scraper to catch these

2. **Verbal announcements**
   - In-class deadline changes not posted anywhere
   - Solution: Check manually or ask professor to post on Canvas

3. **External platform deadlines**
   - Assignments on other platforms (CodeCrunch, Coursemology, etc.)
   - These need separate scrapers

4. **Outlook-only emails** (currently)
   - Future feature - let me know if you need this!

## Tips for Best Results

1. **Run weekly** - Don't wait until deadlines are close
2. **Check `changes.json`** - Quick summary of what changed
3. **Look for 🔔 markers** - These indicate updates from announcements
4. **Keep old `deadlines.json`** - Needed for change detection
5. **Star important announcements** - AI focuses on deadline keywords

## Troubleshooting

**"No changes detected" but I saw an announcement:**
- Check if announcement contains deadline keywords
- Look at `scraper.log` to see which announcements were processed
- The announcement might not mention specific dates

**"False positive changes detected":**
- Happens if professor slightly renamed an assignment
- Review the change and update manually if needed

**"Missed an update":**
- Was it posted on Canvas or only email?
- Canvas announcements = ✅ detected
- Outlook-only emails = ❌ not detected yet

## Want Email Monitoring?

If you find that professors frequently send deadline updates ONLY via Outlook (not Canvas), let me know and I can add:

1. Microsoft Graph API integration for NUS email
2. Email parsing for deadline keywords
3. Automatic merging with Canvas data
4. Duplicate detection (same update on Canvas + email)

For now, the Canvas announcements monitoring should cover 90% of cases since most professors post on Canvas (which also sends you email notifications automatically).
