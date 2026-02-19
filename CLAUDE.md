# Chekhov Oral Examiner

An oral defense examination system for student essays, built as a Google Apps Script web app. Students submit essays, then defend them in a voice conversation with "ChekhovBot" (an 11Labs voice agent), and are graded by Gemini AI.

## Architecture

- **Google Apps Script** web app deployed as a single project
- **Google Sheets** as the database, config store, and prompt store (spreadsheet ID: `181ZO5_JPRYsbDgJZSyzww0dqKoUmKsF_UNGfO8LZJZk`)
- **ElevenLabs Conversational AI** widget for voice-based oral defense
- **Gemini API** for automated grading of defense transcripts

## Files

- `code.gs` — All backend logic: web endpoints (`doGet`/`doPost`), submission processing, question selection, prompt building, transcript webhook handling, Gemini grading
- `index.html` — Single-page frontend with 5 screens (welcome → submit → ready → defense → complete). Inline CSS and JS. Uses `google.script.run` to call backend functions
- `appsscript.json` — Apps Script manifest (V8 runtime, Sheets + external request scopes)
- `Prompts/` — Local reference directory (prompts are stored in the Google Sheet "Prompts" tab, not in code)

## Data Flow

1. Student submits essay via portal → `processSubmission()` generates UUID session_id, selects random questions, stores in Sheets
2. Frontend configures 11Labs widget with `override-prompt` (essay + questions baked in) and `dynamic-variables` (session_id)
3. Student has voice defense with ChekhovBot
4. 11Labs sends transcript webhook → `doPost()` → `handleTranscriptWebhook()` matches by session_id and stores transcript
5. `gradeDefense()` sends essay + transcript to Gemini with rubric from Prompts sheet → stores grade

## Google Sheets Structure

Tabs: **Database** (submissions), **Config** (key-value pairs), **Prompts** (prompt name → text), **Questions** (category + question), **Logs** (debug log)

Key config values: `elevenlabs_agent_id`, `elevenlabs_api_key`, `gemini_api_key`, `gemini_model`, `webhook_secret`, `max_paper_length`, `app_title`

## Development Notes

- This is NOT a Node.js project — it's Google Apps Script (V8 runtime). No package manager, no build step
- Code is pushed to Apps Script via `clasp` or copy-paste to the script editor
- All backend functions are global (Apps Script requirement) — no module system
- Frontend uses `google.script.run` for server calls (Apps Script's built-in RPC)
- The 11Labs widget is loaded from `unpkg.com/@elevenlabs/convai-widget-embed@beta`
- Secrets (API keys) live in the Config sheet, not in code — `getConfig()` reads them at runtime
- Prompts are fetched from the Prompts sheet via `getPrompt()` with hardcoded fallbacks in `buildDefensePrompt()` and `getFirstMessage()`

## Style Conventions

- Use JSDoc comments on functions
- Constants use UPPER_SNAKE_CASE; column indices are 1-based (matching Sheets)
- Status values are defined in the `STATUS` object
- Log to the Logs sheet via `sheetLog(source, message, data)` for debugging — visible in the spreadsheet
