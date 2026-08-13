# SRS Flashcards

A Google Apps Script flashcard application for spaced-repetition vocabulary study. It runs inside Google Sheets as a sidebar, a modeless dialog, or an Apps Script web app.

## Features

- Due, new, difficult, and non-due study pools
- Flip, typing, multiple-choice, true/false, cloze, word-tile, faded-hint, and audio study modes
- SM-2-style interval scheduling with bounded fuzzing
- Dashboard mastery stages, study history, heatmap, rewards, and studied-card tools
- Translation quick-add with Google Translate, DeepL, and Gemini results
- Cached audio and translation responses
- Card editing and archive workflows
- English, Norwegian, and Italian interface localization
- Pseudonymous spreadsheet user identifiers instead of email-based `UserKey` values

## Project Structure

- `Code.js`: Apps Script server functions, scheduling, persistence, translation, audio, archive, and web-app entry points
- `Sidebar.html`: HTML template shell that includes all client partials
- `SidebarMarkup.html`: dashboard, quiz, modal, and control markup
- `SidebarStyles.html`: theme tokens and shared CSS
- `SidebarState.html`: shared client state
- `SidebarLocalization.html`: language names and UI translations
- `SidebarUtils.html`: UI helpers, translation, editing, and navigation
- `SidebarPersistence.html`: startup, refresh, lifecycle, and server synchronization
- `SidebarDashboard.html`: dashboard indexing, stages, heatmap, rewards, and study-pool preparation
- `SidebarQuiz.html`: quiz modes, grading, audio, and answer submission
- `appsscript.json`: Apps Script runtime, timezone, web-app, and access settings
- `LOGO KK banner cropped.png`: local logo asset

## Requirements

- A Google Spreadsheet
- Google Apps Script access
- Node.js and `clasp` for local deployment
- Optional script properties for translation providers:
  - `GEMINI_API_KEY`
  - `DEEPL_API_KEY`

The spreadsheet timezone is configured as `Europe/Berlin` in `appsscript.json`. Client timezone and calendar-day values are also sent with requests.

## Spreadsheet Requirements

The first sheet is the deck. Its first row contains language headers. Every card row must contain a valid hidden or visible column named exactly:

```text
SRS Card ID
```

Each card ID must be non-empty and stable. The application intentionally does not create missing IDs or migrate legacy deck formats.

The app creates and maintains these sheets as needed:

### `SRS_Progress`

```text
UserKey | CardId | Interval | NextReview | FailCount | IsLeech | LastReviewed | EF | PrevInterval
```

### `SRS_History`

```text
UserKey | Date | NewCount | OldCount | TotalReviews
```

### `SRS_ReviewEvents`

Stores idempotent study events and their scheduled result. It is used to prevent duplicate persistence and to count distinct review days for progression caps.

### `Archived cards`

Receives archived card rows. The archive sheet must also contain the `SRS Card ID` column.

## User Privacy

New `UserKey` values are salted SHA-256 pseudonymous identifiers in the form:

```text
user_<hash>
```

The logged-in email may be displayed in the dashboard for the user, but it is not written to spreadsheet tracking rows. The user email is also used for the browser-local reward scope so reward notifications remain stable across the identity-key change.

## Local Setup

Clone or open the project directory, then authenticate `clasp`:

```bash
clasp login
```

For a bound Apps Script project, `.clasp.json` supplies the project connection. Push local files with:

```bash
clasp push
```

Useful checks:

```bash
node --check Code.js
git diff --check
```

The HTML partials are evaluated by Apps Script before delivery. VS Code may report CSS errors on template directives such as:

```html
<?!= include_("SidebarStyles"); ?>
```

Those are editor parsing false positives, not runtime Apps Script errors.

## Deployment

1. Open the target Google Spreadsheet.
2. Open **Extensions > Apps Script**.
3. Ensure the Apps Script project is connected to this source directory.
4. Run `clasp push`.
5. Reload the spreadsheet.
6. Use the **SRS Flashcards** menu to launch the sidebar.

For web-app use, deploy the Apps Script project as a web app using the settings represented in `appsscript.json`:

- Execute as: user accessing the web app
- Access: anyone

Google Apps Script may request authorization when the spreadsheet, external translation APIs, audio, or user session APIs are first used.

## Scheduling Model

Intervals are stored in days. The dashboard maps them to these stages:

| Stage | Interval |
| --- | --- |
| 1 | Under 1 day |
| 2 | 1-3 days |
| 3 | 4-7 days |
| 4 | 8-14 days |
| 5 | 15-30 days |
| 6 | 31-60 days |
| 7 | 61-90 days |
| 8 | 91-180 days |
| 9 | 181-359 days |
| 10 | 360 days and above |

New progression caps are based on distinct review dates, not raw quiz events. Repeating a card in several modes on the same day does not unlock several new stages. Mature-card EF growth is capped at `1.5` to reduce very large jumps.

## Translation

The quick-add flow requests all three providers together:

- Google Translate through `LanguageApp.translate`
- DeepL through its API
- Gemini through the configured Google Generative Language API key

DeepL and Gemini require valid script properties. Provider failures are shown as unavailable rather than preventing the other results from appearing. Only successful translation results are cached.

## Troubleshooting

### Sidebar does not load

- Confirm the spreadsheet and Apps Script project are accessible to the current Google account.
- Confirm the first sheet has headers and a valid `SRS Card ID` column.
- Reload after `clasp push`.
- Check Apps Script execution logs under **Executions**.

### Cards appear in the wrong stage

- Confirm the `Interval` values are numeric days.
- Confirm `LastReviewed` and `SRS_ReviewEvents` contain the current pseudonymous `UserKey`.
- Recalculate the dashboard after refreshing the sidebar.

### Translation provider is unavailable

- Google uses the built-in `LanguageApp` service and may require authorization or quota.
- DeepL requires `DEEPL_API_KEY`.
- Gemini requires `GEMINI_API_KEY`.
- Check Apps Script executions for provider HTTP or quota errors.

### Due count is temporarily stale

The app flushes pending study events before requesting fresh spreadsheet data. If a save fails, the dashboard remains guarded and reports a refresh error instead of starting a new quiz from stale state.

### Reward popup appears unexpectedly

Reward announcements use browser-local storage scoped to the user and deck. Clearing site data, private browsing, or changing browser profiles resets notification history, although earned badge counts remain in the dashboard data.

## Development Notes

- Preserve the shared global scope across HTML partials. Inline HTML handlers call functions defined in the partial scripts.
- Keep server scheduling authoritative. Client code should send study outcomes, not calculated intervals.
- Do not store email addresses in spreadsheet `UserKey` columns.
- Validate both server and client syntax after edits.
- Do not manually alter `SRS_Progress` IDs or card IDs unless the migration impact is understood.
