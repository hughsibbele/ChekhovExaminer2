# Chekhov Oral Examiner

An oral defense examination system for student essays, built as a Google Apps Script web app. Students submit essays through a web portal, defend them in a voice conversation with **ChekhovBot** (an ElevenLabs conversational AI agent), and are graded automatically by Gemini.

ChekhovBot speaks with the formal manner of a 19th century Russian household servant — respectful, warm, but academically rigorous.

## How It Works

1. **Student submits essay** through the web portal
2. **System selects randomized questions** (content + writing process) and builds a custom prompt
3. **Student has a voice defense** with ChekhovBot via the ElevenLabs widget
4. **Transcript is received** via webhook (or recovered via API if the call errored)
5. **Gemini grades the defense** against a rubric, producing a multiplier and structured comments

The grade multiplier adjusts the student's essay grade based on how well they defended it orally. A score of 3 on all rubric elements = 1.00 multiplier (no change). Stronger defenses can earn up to 1.05; weak ones drop to 0.90.

## Architecture

- **Google Apps Script** web app (V8 runtime) — no build step, no package manager
- **Google Sheets** as database, config store, and prompt store
- **ElevenLabs Conversational AI** for voice-based oral defense
- **Gemini API** for automated grading

## Files

| File | Purpose |
|------|---------|
| `code.gs` | All backend logic: web endpoints, submission processing, question selection, prompt building, webhook handling, ElevenLabs API integration, Gemini grading, defense recovery |
| `index.html` | Single-page frontend with 5 screens (welcome, submit, ready, defense, complete). Inline CSS and JS |
| `appsscript.json` | Apps Script manifest |
| `Prompts` | Local mirror of the Prompts sheet (tab-separated) |
| `CLAUDE.md` | Development context and conventions |

## Google Sheets Structure

The spreadsheet serves as both database and configuration:

| Tab | Purpose |
|-----|---------|
| **Database** | Student submissions, transcripts, grades |
| **Config** | Key-value configuration (API keys, thresholds) |
| **Prompts** | Named prompts for the agent and grading system |
| **Questions** | Question bank by category (content / process) |
| **Logs** | Debug log entries from `sheetLog()` |

### Key Config Values

| Key | Description |
|-----|-------------|
| `elevenlabs_agent_id` | ElevenLabs conversational agent ID |
| `elevenlabs_api_key` | ElevenLabs API key (for conversation details + recovery) |
| `gemini_api_key` | Gemini API key for grading |
| `gemini_model` | Gemini model to use (default: `gemini-3-flash-preview`) |
| `webhook_secret` | Secret for authenticating 11Labs webhook calls |
| `min_call_length` | Minimum call duration in seconds before auto-excluding (default: 60) |
| `max_paper_length` | Maximum essay character count (default: 15000) |
| `app_title` | Portal title displayed in the header |

## Grading Rubric

Four elements scored on different scales:

| Element | Scale | Description |
|---------|-------|-------------|
| Paper Knowledge | 1-3 | How well the student knows their own essay |
| Text Knowledge | 1-5 | Knowledge of Chekhov's texts vs. what's in the essay |
| Content Understanding | 1-5 | Depth of analysis vs. the essay |
| Writing Process | 1-3 | Metacognitive awareness of writing as a process |

**3 = meets expectations** (defense consistent with essay quality).

**Multiplier formula:** `1.00 + (average - 3) * 0.05`, clamped to [0.90, 1.05].

**Integrity flags** are raised when any element scores 1 or the average is 1.5 or below.

## Call Exclusion

Short calls (mic failures, immediate disconnects) are automatically excluded from grading:

- Calls shorter than `min_call_length` (default 60s) get status **"Excluded"** instead of "Defense Complete"
- Excluded submissions are skipped by "Grade All Pending"
- To manually exclude: change the status cell to `Excluded` in the spreadsheet
- To re-include: change it back to `Defense Complete`

## Defense Recovery

If a call ends with an ElevenLabs error (e.g., token exhaustion), the webhook may not fire. The recovery mechanism handles this:

1. Open the spreadsheet
2. Go to **Oral Defense** menu > **Recover Stuck Defenses**
3. The system queries the ElevenLabs API for any submissions stuck in "Submitted" or "Defense Started" status
4. Transcripts and call duration are retrieved even from failed conversations
5. Results are shown in a dialog

## Deployment

1. Create a Google Spreadsheet with the tabs listed above
2. Update `SPREADSHEET_ID` in `code.gs`
3. Add your API keys and agent ID to the Config tab
4. Push code via [clasp](https://github.com/google/clasp) or paste into the Apps Script editor
5. Deploy as a web app (Execute as: Me, Access: Anyone)
6. Configure the ElevenLabs webhook to point to your web app URL with `?secret=YOUR_SECRET`

## Spreadsheet Menu

When the spreadsheet opens, an **Oral Defense** menu is added:

- **Grade All Pending** — Grades all "Defense Complete" submissions via Gemini
- **Recover Stuck Defenses** — Retrieves transcripts from ElevenLabs API for stuck submissions
- **Refresh Status Counts** — Shows a summary of submission statuses
- **Format Database Sheet** — Applies column widths and row formatting
