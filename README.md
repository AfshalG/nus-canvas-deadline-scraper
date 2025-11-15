# NUS Canvas Deadline Scraper

An AI-powered tool to automatically extract assignment deadlines and exam dates from your Canvas course introduction documents.

## What This Does

- Fetches all your enrolled courses from NUS Canvas
- Downloads introduction documents/slides from each course
- Uses AI to extract assignment deadlines and exam dates
- Outputs them in chronological order
- Can be run periodically to catch any updated dates

## Prerequisites

1. **Canvas API Token** (I'll show you how to get this)
2. **Python 3.8+**
3. **Anthropic API Key** (for Claude AI)

## Step-by-Step Setup Guide

### Step 1: Get Your Canvas API Token

1. Go to https://canvas.nus.edu.sg/
2. Click on your profile picture (top left)
3. Go to **Settings**
4. Scroll down to **Approved Integrations**
5. Click **+ New Access Token**
6. Give it a purpose (e.g., "Deadline Scraper")
7. Set expiration date (optional, but recommended to set it for end of semester)
8. Click **Generate Token**
9. **IMPORTANT**: Copy this token immediately - you won't see it again!

### Step 2: Get Anthropic API Key

1. Go to https://console.anthropic.com/
2. Sign up or log in
3. Go to API Keys section
4. Create a new API key
5. Copy the key

### Step 3: Install Dependencies

```bash
cd canvas-deadline-scraper
pip install -r requirements.txt
```

### Step 4: Configure

Create a `.env` file in this directory:

```bash
CANVAS_API_TOKEN=your_canvas_token_here
CANVAS_API_URL=https://canvas.nus.edu.sg/api/v1
ANTHROPIC_API_KEY=your_anthropic_key_here
```

### Step 5: Run the Scraper

```bash
python scrape_deadlines.py
```

This will:
1. Fetch all your enrolled courses
2. Download course files (PDFs, PowerPoints, etc.)
3. Use AI to extract deadlines
4. Save results to `deadlines.json` and `deadlines.md`

## Usage

### Basic Usage

```bash
python scrape_deadlines.py
```

### View Deadlines in Terminal

```bash
python scrape_deadlines.py --display
```

### Search for Specific Keywords

```bash
python scrape_deadlines.py --search "assignment,exam,quiz"
```

## Output Format

The script generates two files:

1. **deadlines.json** - Structured data for programmatic use
2. **deadlines.md** - Human-readable markdown format

Example output:

```markdown
# Assignment Deadlines - AY2024/25 Semester 1

## 2024-09-15 - CS2103T Software Engineering
- Assignment 1: UML Diagrams (20%)

## 2024-09-22 - CS3230 Design and Analysis of Algorithms
- Problem Set 1 (10%)

## 2024-10-01 - CS2101 Effective Communication
- OP1 Presentation (15%)
```

## Automating Updates

To automatically check for deadline changes weekly:

### On Mac/Linux:

Add to crontab (`crontab -e`):

```bash
0 9 * * 1 cd /path/to/canvas-deadline-scraper && python scrape_deadlines.py
```

This runs every Monday at 9 AM.

### On Windows:

Use Task Scheduler to run the script weekly.

## Features

- ✅ Automatic course enrollment detection
- ✅ Intelligent intro document detection (supports various naming conventions)
- ✅ Multi-format support (PDF, DOCX, PPTX, TXT)
- ✅ AI-powered deadline extraction
- ✅ **Canvas announcements monitoring** - catches deadline changes posted by professors
- ✅ **Automatic change detection** - compares with previous runs to highlight updates
- ✅ Chronological sorting
- ✅ Duplicate detection
- ✅ Email notifications (optional)
- ✅ Update highlighting with 🔔 markers

## Supported Intro Document Patterns

The scraper automatically identifies intro/syllabus documents using these patterns:

**Tested with real NUS course filenames:**
- `L01a Introduction and Module Information.pdf` ✓
- `EE2026_L0_intro.pdf` ✓
- `2025_CS2107_Topic_0.pdf` ✓
- `AY25-26 S1 ES2631 Course Schedule and Assessments_STUDENT.pdf` ✓
- `LAH 1201 Syllabus.pdf` ✓

**Detected patterns include:**
- Lecture numbers: `L0`, `L00`, `L01`, `L1`, `L01a`, `Lec0`, `Lecture 1`
- Topic numbers: `Topic 0`, `Topic_0`
- Keywords: `intro`, `introduction`, `syllabus`, `course outline`, `schedule`, `overview`, `module information`, `assessment`, `course schedule`
- Week numbers: `Week 0`, `Week 1`, `Week01`

The detection is smart enough to avoid false positives like "Lecture 5" or "Week 8"!

## Deadline Update Detection 🔔

The scraper automatically monitors for deadline changes:

### What Gets Detected:

1. **Canvas Announcements** (✅ Implemented)
   - Extended deadlines
   - Postponed exams
   - Rescheduled submissions
   - New assignments announced mid-semester
   - Cancelled assessments

2. **Changes Between Runs** (✅ Implemented)
   - Date changes (e.g., Jan 20 → Jan 25)
   - New deadlines added
   - Removed/cancelled deadlines

### Example Output:

When you re-run the scraper, changes are highlighted at the top of `deadlines.md`:

```markdown
## ⚠️ RECENT CHANGES DETECTED

### 📅 Date Changes:
- **CS2103T** - Assignment 2: ~~2025-01-20~~ → **2025-01-25**

### ✨ New Deadlines:
- **CS2107** - Bonus Assignment (2025-02-01)
```

Updates from announcements are marked with 🔔:
```markdown
- **[ASSIGNMENT]** Assignment 2 (15%)
  - 🔔 **UPDATE:** EXTENDED from Jan 20 to Jan 25
  - _Source: Announcement: Assignment 2 Extension_
```

**See [UPDATE_DETECTION.md](UPDATE_DETECTION.md) for complete details.**

### What About Email Updates?

Currently monitors **Canvas announcements only** (which Canvas also emails you).

**Email-only updates** (NUS Outlook, no Canvas announcement):
- Not yet implemented
- Can be added if needed - let me know!
- Would use Microsoft Graph API to scan Outlook emails

Most professors post on Canvas, so this covers 90%+ of cases.

## Troubleshooting

### "Invalid API token"
- Make sure you copied the full token
- Check that it hasn't expired
- Regenerate if needed

### "No files found"
- Course materials might not be uploaded yet
- Check `scraper.log` to see what files were detected
- If your professor uses unusual naming, the script will list all available files in the log

### "AI extraction failed"
- Some documents might be scanned images - use OCR preprocessing
- Check your Anthropic API credits

## Privacy & Security

- Your Canvas token is stored locally only
- No data is sent anywhere except Canvas API and Anthropic API
- Add `.env` to `.gitignore` to prevent token leakage

## Cost Estimate

- Canvas API: **Free**
- Anthropic API: ~$0.10-0.50 per semester (depends on document size)

## Advanced: Email Notifications

To get email alerts when deadlines change, set up:

```bash
SMTP_EMAIL=your_email@gmail.com
SMTP_PASSWORD=your_app_password
NOTIFICATION_EMAIL=your_email@gmail.com
```

## License

MIT

## Support

If you encounter issues:
1. Check the logs in `scraper.log`
2. Verify your API tokens
3. Ensure course materials are uploaded
