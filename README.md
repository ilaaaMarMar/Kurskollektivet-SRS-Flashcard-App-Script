# SRS Flashcards

A spaced-repetition flashcard app that lives inside your Google Sheet. Study vocabulary with due/new/difficult pools, a smart SM-2 scheduler, progress dashboards, rewards, and one-click translation quick-add. Runs as a sidebar, a pop-out dialog, or a web app from the same Apps Script project.

---

## Part 1 — For Users

### What you get

- **Smart scheduling** — cards return based on how well you know them, using a proven spaced-repetition (SM-2) algorithm with a gentle, language-learning-friendly progression curve.
- **Study pools** — **Due** (ready to review), **New** (never studied), and **Difficult** (cards you've struggled with).
- **Eight study modes** — flip, typing, multiple-choice, true/false, cloze, word tiles, faded hints, and audio (listen & type).
- **Dashboard** — mastery stages (1–10), "studied today" counts, a review heatmap (week/month/year), gamified reward badges, and a list of today's cards with quick edit/archive.
- **Translate & Quick Add** — look up a word and add it to your deck straight from the app using Google, DeepL, and Gemini.
- **Languages & themes** — English, Norwegian, and Italian UI; light and dark themes.

### Getting started

1. Open your spreadsheet, then open **Extensions → Apps Script**.
2. From the spreadsheet menu, choose **📚 SRS Flashcards → Launch Sidebar View** (or **Launch Pop-Out Window** for a larger, floating window).
3. Pick your two language columns from the **Language Combination** selector on the dashboard.
4. Tap **🚀 Study Due Cards** (or **✨ Study New Cards**) and answer each card.
5. After each segment you'll see a summary; the dashboard updates with your progress.

### Your spreadsheet

- The **first sheet is your deck**. Row 1 contains the column headers (e.g. `English`, `Norwegian`).
- Every card row must have a value in the **`SRS Card ID`** column — a hidden or visible column named exactly that. The app fills in any missing IDs automatically and keeps them stable.
- The app maintains four helper sheets automatically: `SRS_Progress`, `SRS_History`, `SRS_ReviewEvents`, and `Archived cards`. Don't rename or hand-edit them.

### Dashboard explained

| Item | Meaning |
| --- | --- |
| Study Due / Study New / Study Difficult | Start a session from that pool (number in parentheses = how many cards). |
| Studied Today | New + review cards recorded today. |
| Stage boxes (1–10) | Distribution of your cards by interval; click **View** to see that group. |
| Heatmap | Green-scale history of how many reviews you did per day. |
| Rewards | Badges unlocked by consistency (e.g. 4 days/week, 20 days/month). |

### Buttons & settings

- **🌐 Web App** — get a shareable link to study in any browser on any device.
- **🇬🇧 / 🇳🇴 / 🇮🇹** — switch the interface language.
- **☀️ / 🌙** — switch the theme (remembered per browser).
- **🔄 Change Deck** — paste a spreadsheet URL or ID to switch decks.
- **🗃️ Archive** — move mastered cards (360-day interval) out of the active deck.
- **✏️ Edit** — change a card's text directly from the quiz or the "studied today" list.

### Privacy

- Your email is shown to you in the header ("Logged in as…") but is **never written** into the spreadsheet. Tracking rows use a pseudonymous `UserKey` — a salted hash, e.g. `user_<hash>`.
- Reward notifications are remembered in your browser only.

### Troubleshooting

**Sidebar won't load** — Make sure you have edit access to the spreadsheet, that the first row has headers, and that an `SRS Card ID` column exists. Then refresh the spreadsheet.

**Due count looks wrong** — The app flushes recent study results before refreshing; if a save failed it shows an error rather than a stale count. Refresh the sidebar.

**Translation provider shows unavailable** — Google is built in but needs authorization/quota. DeepL and Gemini require API keys configured by the app owner.

**"Logged in as" not showing** — You're running as an anonymous/incognito user; study progress is still tracked under a temporary key.

---

## Part 2 — For Developers

### Project structure

```
Code.js                          Server: scheduling, persistence, translation, audio, archive, entry points
Sidebar.html                     Template shell that includes all client partials (shared scope)
SidebarStyles.html               All CSS + theme tokens (included in <head>)
SidebarMarkup.html               Static HTML: top-bar, dashboard, quiz, all modals
SidebarLocalizationAndUtils.html i18n (en/no/it) + shared UI helpers (dialogs, quick-add, edit, selectors)
SidebarStateAndStorage.html      Shared client state, localStorage persistence, server-call plumbing
SidebarDashboard.html            Pool indexing, stages, heatmap, rewards + reward queue, session sync queue
SidebarQuiz.html                 Quiz modes, grading, flip select-then-submit, archive
appsscript.json                  V8 runtime, Europe/Berlin timezone, web-app settings
```

`Sidebar.html` evaluates the partials via `include_()` so **all scripts share one global scope**. Top-level names are cross-file contracts — do not rename without checking every file (inline `onclick` handlers also reference them).

### Local setup & deployment

```bash
clasp login                 # authenticate
clasp push                  # push all files to the bound Apps Script project
clasp push -f               # force-upload (normalizes line endings)
```

```bash
node --check Code.js        # syntax check server code
```

The pop-out dialog and sidebar run the bound script's **HEAD** version, so client-only changes are live after `clasp push` (no version deploy needed). The **web app** serves the latest *deployment*; redeploy after changes:

```bash
clasp deploy -i <DEPLOYMENT_ID> -d "message"
```

Web-app settings (`appsscript.json`): execute as the accessing user, access anyone. Optional script property `SRS_WEBAPP_URL` overrides the auto-derived web-app URL shown to users.

### Scheduling model

Intervals are stored in **days**. Stage mapping used by the dashboard:

| Stage | Interval |
| --- | --- |
| 1 | < 1 day |
| 2 | 1–3 d |
| 3 | 4–7 d |
| 4 | 8–14 d |
| 5 | 15–30 d |
| 6 | 31–60 d |
| 7 | 61–90 d |
| 8 | 91–180 d |
| 9 | 181–359 d |
| 10 | ≥ 360 d |

Core scheduler: `scheduleAnswerResult_` (Code.js). SM-2 EF update, EF floor 1.3, mature growth factor capped at 1.5, bounded overdue bonus (≤ 50% of interval), ±5% fuzz on mature intervals, max 360 days. Progression is further capped by distinct review-day count via `getReviewCountIntervalCap_` (`[0.25,1,3,7,14,30,60,90,180,360]`), so re-studying a card multiple times on the same day does not unlock new stages. **The server is the sole scheduling authority** — the client sends outcomes (`isCorrect`, `confidence`, `isTypo`), never computed intervals.

### Sync architecture

- **`submitAnswer`** optimistically updates the local dashboard and enqueues a result with a **session-unique `clientId`** (`CLIENT_SESSION_ID + ':' + seq`, SidebarStateAndStorage) — never a plain counter, since server-side dedup keys on `clientId` and counters reset between sessions.
- Results are streamed to the server one-at-a-time via a **serialized queue** (`streamQueue`/`streamInFlight`, SidebarDashboard) so at most one `saveSegmentResults` is in flight; a final **flush** on dashboard return / page-hide syncs the whole batch.
- `saveSegmentResults` (Code.js) is **idempotent**: an existing `APPLIED` event with the same `clientId` is skipped, so retries can't double-schedule. Corrections (marked `corrected: true`) rewrite the event instead.
- History counts are recomputed from the event sheet as absolute totals; the `SRS_History` cache (300 s) serves reads.

### Sheet schemas

`SRS_Progress`: `UserKey | CardId | Interval | NextReview | FailCount | IsLeech | LastReviewed | EF | PrevInterval`

`SRS_History`: `UserKey | Date | NewCount | OldCount | TotalReviews`

`SRS_ReviewEvents`: `UserKey | ClientId | CardId | IsCorrect | Confidence | WasNew | IsTypo | InDifficultMode | LastReviewed | ResultJson | Status | AppliedAt | LangA | LangB`

Datetimes (`NextReview`, `AppliedAt`) are stored as UTC text in `yyyy-MM-dd HH:mm:ss`; date-only fields (`LastReviewed`, `SRS_History.Date`) use `yyyy-MM-dd`. The app parses these as UTC, so scheduling is timezone-stable.

`Archived cards`: same columns as the main deck (must include `SRS Card ID`).

### External providers

| Provider | Config | Notes |
| --- | --- | --- |
| Google Translate | built-in `LanguageApp` | authorization/quota limited |
| DeepL | script property `DEEPL_API_KEY` | free tier auto-detected via `:fx` suffix |
| Gemini | script property `GEMINI_API_KEY` | model fallback chain in `GEMINI_MODELS`; spellcheck + translation |
| Audio (TTS) | none | unofficial `translate.google.com/translate_tts` endpoint (fragile; see Suggestions) |

Rate limits (`enforceRateLimit_`): translate 20/min, spellcheck 30/min, audio 60/min.

### Development guidelines

- Preserve the shared global scope; keep top-level names stable across partials.
- Server remains the scheduling authority — never persist client-computed intervals.
- Never write emails into `UserKey` columns; use `getActiveUserId()`.
- Run `node --check Code.js` and the harness tests in `tests/` before pushing.
- `git diff --check` before committing.

### Known limitations & roadmap

- Per-card `saveSegmentResults` re-reads the full deck for validation — batching + a cached deck snapshot is the highest-value optimization.
- `SRS_ReviewEvents` grows unbounded; consider periodic pruning.
- Leech flag currently fires on a single failure and never decays — a double-failure threshold with decay is recommended.
- Audio relies on an unofficial Google endpoint.