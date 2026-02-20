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
- `Prompts` — Local reference file (tab-separated: prompt_name → prompt_text). Canonical copy lives in the Google Sheet "Prompts" tab; this file is the local mirror

## Data Flow

1. Student submits essay via portal → `processSubmission()` generates UUID session_id, selects random questions, stores in Sheets
2. Frontend configures 11Labs widget with `override-prompt` (essay + questions baked in) and `dynamic-variables` (session_id)
3. Student has voice defense with ChekhovBot
4. 11Labs sends transcript webhook → `doPost()` → `handleTranscriptWebhook()` matches by session_id, fetches `call_duration_secs` from ElevenLabs API, stores transcript + call length
5. `gradeDefense()` sends essay + transcript to Gemini with rubric from Prompts sheet → stores multiplier and structured comments

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
- **Recovery**: `recoverStuckDefenses()` (Oral Defense menu → "Recover Stuck Defenses") queries the ElevenLabs API to retrieve transcripts and call duration for submissions stuck in Submitted/Defense Started status — covers cases where the webhook never fired (e.g., 11Labs errors, token exhaustion)

## Grading System

- **Rubric**: 4 elements — Paper Knowledge (1-3) and Writing Process (1-3) are capped at 3; Text Knowledge (1-5) and Content Understanding (1-5) can go higher. 3 = meets expectations
- **Multiplier formula**: `1.00 + (average - 3) × 0.05`, clamped to [0.90, 1.05], rounded to 2 decimal places
- **Integrity flags**: Any element scoring 1 or average ≤ 1.5 triggers a flag. Comments are prefixed with "⚠ INTEGRITY FLAG ⚠"
- **Parser** (`parseGradingResponse`): extracts `Multiplier: X.XX` line from Gemini's structured output; falls back to computing from individual scores if that line is missing
- Prompts: `grading_system_prompt` (role/persona) and `grading_rubric` (rubric + scoring formula + output format)

## Style Conventions

- Use JSDoc comments on functions
- Constants use UPPER_SNAKE_CASE; column indices are 1-based (matching Sheets)
- Status values are defined in the `STATUS` object: Submitted → Defense Started → Defense Complete → Graded → Reviewed (also: Excluded)
- **Exclusion**: Calls shorter than `min_call_length` config (default 60s) are auto-set to "Excluded" status and skipped by grading. To manually exclude, change the status cell to "Excluded" in the spreadsheet; to re-include, change it back to "Defense Complete"
- Log to the Logs sheet via `sheetLog(source, message, data)` for debugging — visible in the spreadsheet
