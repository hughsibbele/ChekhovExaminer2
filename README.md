# Oral Examiner 4.0

An oral defense examination system for student essays, built as a Google Apps Script web app. Students submit essays through a web portal, defend them in a voice conversation with an AI examiner (ElevenLabs Conversational AI), and are graded automatically by Gemini.

The grade multiplier adjusts the student's essay grade based on how well they defended it orally. A score of 3 on all rubric elements = 1.00 multiplier (no change). Stronger defenses can earn up to 1.05; weak ones drop to 0.90.

## Getting Started (~5 minutes)

### 1. Copy the Template

Click the template link to make your own copy of the spreadsheet (includes all tabs, prompts, and questions as sample data).

### 2. Run the Setup Wizard

1. Open your copied spreadsheet
2. Go to **Oral Defense** menu > **Setup Wizard (start here)**
3. Paste your three API keys:
   - **ElevenLabs Agent ID** — from your ElevenLabs Conversational AI agent settings
   - **ElevenLabs API Key** — from elevenlabs.io > Profile > API Keys
   - **Gemini API Key** — from aistudio.google.com > API Keys
4. Click **Save & Complete Setup**

### 3. Deploy as a Web App

1. Go to **Extensions > Apps Script**
2. Click **Deploy > New deployment**
3. Select type: **Web app**
4. Set "Execute as": **Me** and "Who has access": **Anyone**
5. Click **Deploy** and copy the URL

Share that URL with your students. No webhook configuration needed — transcripts are fetched automatically.

## How It Works

1. **Student submits essay** through the web portal
2. **System selects randomized questions** (content + writing process) and builds a custom prompt
3. **Student has a voice defense** with the examiner via the ElevenLabs widget
4. **Student clicks "Finish Defense"** — the system fetches the transcript from the ElevenLabs API
5. **Gemini grades the defense** against a rubric, producing a multiplier and structured comments

Transcripts are also recovered automatically every 5 minutes by a background trigger, as a safety net.

## Architecture

- **Google Apps Script** web app (V8 runtime) — no build step, no package manager
- **Google Sheets** as database, config store, and prompt store
- **ElevenLabs Conversational AI** for voice-based oral defense
- **Gemini API** for automated grading

## Customization

The Prompts and Questions tabs contain sample data meant to be customized for your course:

- **Prompts tab**: Edit the agent personality, examination flow, first message, grading rubric
- **Questions tab**: Replace with your own content and process questions
- **Config tab**: Adjust thresholds (`min_call_length`, `max_paper_length`, question counts)
- **Avatar**: Set `avatar_url` in the Config tab to use a custom bot avatar

## Google Sheets Structure

| Tab | Purpose |
|-----|---------|
| **Database** | Student submissions, transcripts, grades |
| **Config** | Key-value configuration (thresholds, UI settings) |
| **Prompts** | Named prompts for the agent and grading system |
| **Questions** | Question bank by category (content / process) |
| **Logs** | Debug log entries from `sheetLog()` |

## Grading Rubric

Four elements scored on different scales:

| Element | Scale | Description |
|---------|-------|-------------|
| Paper Knowledge | 1-3 | How well the student knows their own essay |
| Text Knowledge | 1-5 | Knowledge of source texts vs. what's in the essay |
| Content Understanding | 1-5 | Depth of analysis vs. the essay |
| Writing Process | 1-3 | Metacognitive awareness of writing as a process |

**3 = meets expectations** (defense consistent with essay quality).

**Multiplier formula:** `1.00 + (average - 3) * 0.05`, clamped to [0.90, 1.05].

**Integrity flags** are raised when any element scores 1 or the average is 1.5 or below.

## Spreadsheet Menu

When the spreadsheet opens, an **Oral Defense** menu is added:

- **Grade All Pending** — Grades all "Defense Complete" submissions via Gemini
- **Recover Stuck Defenses** — Retrieves transcripts from ElevenLabs API for stuck submissions
- **Refresh Status Counts** — Shows a summary of submission statuses
- **Format Database Sheet** — Applies column widths and row formatting
- **Re-run Setup Wizard** — Update API keys or settings

## Call Exclusion

Short calls (mic failures, immediate disconnects) are automatically excluded from grading:

- Calls shorter than `min_call_length` (default 60s) get status **"Excluded"** instead of "Defense Complete"
- Excluded submissions are skipped by "Grade All Pending"
- To manually exclude: change the status cell to `Excluded` in the spreadsheet
- To re-include: change it back to `Defense Complete`

## Secrets Management

API keys are stored in **Script Properties** (`PropertiesService`), not in the spreadsheet. The Setup Wizard handles this automatically. Non-secret config stays in the Config sheet for easy editing.
