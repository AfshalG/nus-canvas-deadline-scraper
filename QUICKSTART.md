# Quick Start Guide - 5 Minutes to Your Deadlines!

## 1. Get Canvas API Token (2 minutes)

1. Go to https://canvas.nus.edu.sg/
2. Click your profile picture → **Settings**
3. Scroll to **Approved Integrations**
4. Click **+ New Access Token**
5. Purpose: "Deadline Scraper"
6. Click **Generate Token**
7. **Copy the token immediately!** (You won't see it again)

## 2. Get Anthropic API Key (2 minutes)

1. Go to https://console.anthropic.com/
2. Sign up (it's free to start)
3. Go to **API Keys**
4. Click **Create Key**
5. **Copy the key**

Free tier includes enough credits for several semesters of deadline scraping!

## 3. Setup (1 minute)

```bash
cd canvas-deadline-scraper

# Install dependencies
pip install -r requirements.txt

# Copy example config
cp .env.example .env

# Edit .env and paste your tokens
nano .env  # or use any text editor
```

In `.env`, replace:
- `your_canvas_api_token_here` with your Canvas token
- `your_anthropic_api_key_here` with your Anthropic key

## 4. Run It!

```bash
python scrape_deadlines.py
```

That's it! Your deadlines will be in:
- `deadlines.md` (easy to read)
- `deadlines.json` (for apps/scripts)

## 5. View Results

```bash
# View in terminal
python scrape_deadlines.py --display

# Or just open the markdown file
cat deadlines.md
```

## Troubleshooting

**"CANVAS_API_TOKEN not set"**
→ Make sure you saved the `.env` file

**"Failed to authenticate"**
→ Double-check your Canvas token is correct

**"No intro documents found"**
→ Your course might use different file names. Check the logs to see what files were found.

**"ANTHROPIC_API_KEY not set"**
→ Make sure you added your Anthropic API key to `.env`

## Tips

- Run this at the start of each semester
- Run weekly to catch any deadline updates
- The script remembers downloaded files, so re-runs are fast!

## What Gets Extracted?

The AI will find:
- Assignment deadlines
- Quiz dates
- Exam dates
- Project milestones
- Presentation dates
- Any other important dates with percentages/weights

**Plus automatic update detection:**
- Checks Canvas announcements for deadline changes
- Compares with previous runs to highlight updates
- Marks changes with 🔔 in your output
- Shows what was extended, added, or cancelled

## Example Output

```markdown
## September 2024

### Mon, 16 Sep - **CS2103T Software Engineering**
- **[ASSIGNMENT]** Individual Project (20%)
  - _Submit via GitHub_

### Fri, 20 Sep - **CS3230 Design & Analysis of Algorithms**
- **[QUIZ]** Quiz 1 (10%)

## October 2024

### Tue, 01 Oct - **CS2101 Effective Communication**
- **[PRESENTATION]** OP1 (15%)
  - _8-10 minutes_
```

## Next Steps

See `README.md` for:
- Automation (run weekly automatically)
- Email notifications
- Advanced configuration
