# NBA Live Market Tracker

Internal Streamlit tool for sportsbook traders and QC analysts. Tracks live NBA
play-by-play, results the first-event / timeframe / second-half markets that can
be determined directly from play-by-play, and flags stat corrections — separating
market-impacting corrections from harmless feed churn.

Layout and terminology follow `NBA Tracker Layout.xlsx`.

```
nba_live_tracker/
├── app.py                  # the whole app (12 commented sections)
├── requirements.txt
├── .gitignore
├── .streamlit/config.toml  # dark compact theme
└── .tracker_state/         # created at runtime, gitignored
```

---

## 1. Running it

### Locally

```bash
python -m venv .venv
source .venv/Scripts/activate     # Git Bash on Windows
pip install -r requirements.txt
streamlit run app.py
```

### Flow

1. **Load Live Games** (sidebar) — pulls today's slate. Change **Slate date** for
   another day. Or use **Add game by ID** (below) to skip the scoreboard entirely.
2. **Select a game** — live games sort first, then upcoming, then finals.
3. Both rosters load automatically.
4. Pick **two key players per team** from the roster dropdowns.
5. **Track Game** — starts polling and enables the three tabs.

No play-by-play request is made until you press **Track Game**. Once tracking, the
app auto-refreshes every 4s (15s if the game has not tipped off, and not at all
once the game is final). **Refresh now** in the sidebar forces a fetch.

### What sits above the tabs

A **game header** — both teams, both scores, the period and clock, and the time of
the last successful poll — then a **one-line status**. The big house banner
(`warning_box`, as in the NFL and NHL tools) is still there, but only for a state
that wants attention: a live stat correction, a data delay, a feed warning or a
rate-limit cooldown. `STATUS: OK` does not want attention, so it drops to
`status_line`, a compact line with a green dot. `banner_state()` still decides the
message and the precedence; only the rendering differs.

### Screen conventions

- **Section headings** (`sect`) are 13px uppercase with a hairline rule and no accent
  bar. The bar was a coloured object on every heading of a screen already using colour
  to mean something; the rule alone separates them.
- **The gap above a heading is unconditional** (`margin-top: 20px`). It used to be
  overridden by `.sect:first-child { margin-top: 4px }`, intended for the first heading
  on a tab — but Streamlit wraps each `st.markdown` in its own container, so the div is
  *always* the only child and the override matched every heading. The visible gap was
  therefore whatever the preceding block happened to leave behind (6px after a key
  player card, 5px after a feed row, more after a table), which is why some sections
  sat tighter than others. Measured after the fix: 22px above all three headings of the
  Live tab, identical.
- **No explanatory captions.** Legends and methodology notes ("Yes means both teams
  scored inside the window", how corrections are fingerprinted, how dunks are
  classified) are in this README and in comments at the call site, not on screen. The
  grey `note` helper survives only for *state*: "Lineup not available yet", "Lineup
  source: boxscore", the lineup-unverified warning, "Activates when the third quarter
  begins", and the two empty-state lines. No section carries a caption.
- **Teams are named in full wherever a panel is headed by one** — "San Antonio Spurs",
  not "SA KP". Abbreviations appear only inside rows that mix teams.

### Add game by ID

The sidebar's **Add game by ID** expander takes a bare ESPN event id
(`401810433`) or a pasted ESPN game URL — the id is extracted from the URL, so
you can copy straight out of the address bar. Use it when:

- the game is not on the slate you loaded,
- the scoreboard has already dropped the game,
- or the scoreboard endpoint itself is failing. The summary endpoint is keyed
  only on the event id, so it still works when `/scoreboard` does not.

The added game is merged into the same list the **Game** dropdown reads and is
labelled `(added by ID)`, so from that point it behaves identically to a game
picked off the slate — rosters auto-load, key players, tracking, corrections, all
the same. Loading a slate afterwards does not drop it.

A bad id gives a plain warning in the sidebar, not a traceback. Two details worth
knowing:

- Identity/status for a hand-added game come from the summary endpoint's `header`
  block, which — unlike the scoreboard — carries **no `period` and no
  `displayClock`**. Nothing depends on them: every market derives the period and
  clock from the play-by-play itself.
- While tracking a hand-added game, status refreshes read the summary header
  rather than the scoreboard. Both are cached at 20s, so the request rate is
  unchanged.

**Edit Key Players** sits inside the Live tab. It writes to the same canonical
slot as the sidebar control, so changing key players mid-game does not reset the
correction log, first-basket results, or key-player alert history.

---

## 2. Deploying to Streamlit Cloud from GitHub

### 2.1 Create the repo and push

Run these from Git Bash, in order:

```bash
cd /c/Users/e.colwell/nba_live_tracker
git init
git add .
git commit -m "NBA Live Market Tracker: initial version"
git branch -M main
```

Create an empty GitHub repo named `NBA_Live_Market_Tracker` (no README, no
.gitignore — the repo must be empty), then:

```bash
cd /c/Users/e.colwell/nba_live_tracker
git remote add origin https://github.com/Ecolwell99/NBA_Live_Market_Tracker.git
git push -u origin main
```

Verify:

```bash
cd /c/Users/e.colwell/nba_live_tracker
git status
git log --oneline -1
```

### 2.2 Deploy

1. Go to <https://share.streamlit.io> and sign in with the same GitHub account.
2. **Create app** → **Deploy a public app from GitHub**.
3. Fill in:
   - **Repository:** `Ecolwell99/NBA_Live_Market_Tracker`
   - **Branch:** `main`
   - **Main file path:** `app.py`
4. **Advanced settings** → **Python version:** `3.11` (3.11 or 3.12 both work;
   pin one so a future default change cannot break the build).
5. **Deploy**. First build installs `requirements.txt` and takes 1–3 minutes.

No secrets or environment variables are needed — every endpoint used is public
and unauthenticated.

### 2.3 Pushing a change later

```bash
cd /c/Users/e.colwell/nba_live_tracker
git add -A
git commit -m "<what changed>"
git push
```

Streamlit Cloud redeploys on push. If it does not, use **Manage app → Reboot**.

### 2.4 Streamlit Cloud caveats

- **`.tracker_state/` does not survive a container restart.** On Cloud the
  filesystem is ephemeral, so the JSON sidecar only protects you against a hard
  browser refresh within the same container — which is what it is for. Session
  state is the real store; the sidecar is best-effort. If the app is rebooted
  mid-game, press **Track Game** again and the correction log restarts from a
  fresh baseline (the log is append-only *within a tracking session*, per spec).
- **The app sleeps after inactivity** on the free tier. Wake it before tip-off,
  not at tip-off.
- **One container, shared state.** Every viewer gets their own session state, so
  two traders watching the same game each keep their own correction log. That is
  fine for QC, but two people will not see each other's dismissals.

---

## 3. Data source: what was chosen and why

### Chosen: ESPN public JSON

```
https://site.api.espn.com/apis/site/v2/sports/basketball/nba/scoreboard
https://site.api.espn.com/apis/site/v2/sports/basketball/nba/summary?event={id}
https://site.api.espn.com/apis/site/v2/sports/basketball/nba/teams/{id}/roster
```

The official NBA feed was tried first and rejected on evidence, not preference:

| Endpoint | Result from this network |
|---|---|
| `cdn.nba.com/static/json/liveData/playbyplay/...` | **HTTP 403** (Akamai "Access Denied") — persists with browser `User-Agent`, `Referer` and `Origin` headers |
| `stats.nba.com/stats/playbyplayv3` | **timeout**, no response at all |
| `www.nba.com/robots.txt` | 200 — so it is the CDN JSON specifically that is blocked, not NBA.com |
| `site.api.espn.com/...` (ESPN) | **200**, and already the working provider in the existing `nfl_qc_tool` |

Every field the markets depend on was verified against a real ESPN NBA payload
before any code was written:

| Need | ESPN field | Verified |
|---|---|---|
| FG attempt, and 2 vs 3 | `shootingPlay` + `pointsAttempted` | ✅ free throws also carry `shootingPlay: true` with `pointsAttempted: 1`, so FG detection requires `pointsAttempted in (2, 3)` |
| Made vs missed | `scoringPlay` / `scoreValue` | ✅ misses report `scoreValue: 0` |
| Shot detail / dunk | `type.text` (e.g. `"Driving Dunk Shot"`) | ✅ |
| Shooter | `participants[0].athlete.id` | ✅ correct even on blocked shots, where the *description* leads with the blocker's name |
| Team | `team.id` | ✅ |
| Clock bucketing | `period.number`, `clock.displayValue` | ✅ |
| Score at the time | `awayScore` / `homeScore` on each play | ✅ |
| Starters | `boxscore.players[].statistics[].athletes[].starter` | ✅ exactly 5 per team, once the boxscore publishes |
| Roster | `athletes` (flat list for NBA), `jersey`, `position.abbreviation` | ✅ |
| Substitutions | `type.id == "584"`, `participants[0]` on / `participants[1]` off | ✅ 480/480 across 9 games; **no on-court field exists** on `summary`, the core-API competition object, `/situation` or the competitor `/roster`, so the five are derived — see 4.6 |

### Swapping providers later

All source-specific code is confined to **SECTION 3 (fetch)** and **SECTION 4
(normalise)**. Everything downstream — markets, correction engine, UI — operates
only on the neutral frozen `GameEvent` dataclass. To move to the NBA CDN feed from
a network that can reach it, rewrite those two sections and nothing else.
`parse_clock_seconds` already handles the NBA CDN's ISO-8601 clock shape
(`PT11M39.00S`) alongside ESPN's `M:SS` and sub-minute `SS.T`.

---

## 4. Known limitations

Read this section before resulting anything off the tool.

### 4.1 Live starters

- **Before tip-off, ESPN publishes no lineup.** The boxscore `starter` flag only
  appears once the boxscore itself does — around tip-off, not before.
- `resolve_starters()` therefore falls back in three steps, and the UI **always
  prints which one is in use** under each player table:
  1. `boxscore ... starter: true` — real starters. Trust this.
  2. **First five distinct players from that team to appear in play-by-play** — a
     good proxy a minute into the quarter, but a fast substitution or a player
     whose first action is a rebound can distort it.
  3. **Roster order** — labelled as such. This is **not a lineup**. Do not result
     a prematch player market off it.
- Consequence: the Prematch tab's per-player tables are only reliable once source
  (1) is live. Check the "Lineup source" line, every time.

### 4.2 Dunk classification

- Dunks come from a **text match on the feed's shot type** (`\bdunk`,
  case-insensitive, word-boundary anchored), because ESPN exposes no dunk flag.
- The word boundary matters: the mock-up's own `Ryan Dunn` would otherwise
  register as a dunk. Names ending in "Dunn"/"Dunning" are safe; a shot type
  ESPN chooses to word differently would not be caught.
- **ESPN's own classification is inconsistent.** Some finishes come through as
  `"Layup Shot"` when broadcast called a dunk, and alley-oops are sometimes
  `"Alley Oop Layup"`. First Dunk markets should be QC'd against video, not
  resulted off the feed alone.
- Only **made** dunks resolve First Dunk. A missed/blocked dunk attempt does not.
- If ESPN later reclassifies a layup as a dunk, that **is** caught — category 8
  below.

### 4.3 Stat corrections

This is where the tool is most useful and also where it is most opinionated.

**How detection works.** Every poll, the tracked events (FG / FT / turnover only —
rebounds, fouls and substitutions are excluded to keep the diff quiet) are
fingerprinted and diffed against the previous snapshot. Three shapes are logged:

- **REMOVED** — an event id in the previous snapshot is gone.
- **CHANGED** — same id, different fingerprint.
- **INSERTED** — an id we have not seen that is *not normal forward progress*.

**Insertion is the hard case, and it is guarded twice.** ESPN's
`sequenceNumber` values are sparse and non-contiguous (4, 7, 8, 9, 11, 13…),
which makes them look like stable per-play ids — but how a *retroactive*
insertion gets numbered could not be confirmed against live data. So an unseen
event counts as a correction if **either** its sequence sits at or below the
watermark, **or** its game clock is more than `RETROACTIVE_TOLERANCE_SECONDS`
(45s) behind the live edge. Belt and braces, deliberately.

**Substitutions are kept out of this arithmetic entirely**, not just out of the
fingerprint diff. They were quietly generating false insertions two ways, both
measured over 9 games / 3,746 plays:

- The first play of every new quarter is a substitution stamped with the **new**
  period at `12:00`. Computing the live edge over every event therefore jumped a
  quarter ahead before any shot in that quarter arrived, and a genuine
  buzzer-beater from the quarter just ended read as a backdated insertion.
- Substitutions also carry sequence numbers above the newest tracked play (329 vs
  233 in one game), which tripped the watermark test on the next real plays.

The watermark and the live edge now come from tracked events only. Replaying all
9 games as polls: **279 false insertions before, 199 after**. The remaining 199
are *not* substitution-related — see 4.7.

**Market-impacting categories** (the eight from the spec):

1. Made shot added or removed
2. Missed shot added or removed
3. Make → miss, or miss → make
4. Two-pointer → three-pointer, or vice versa
5. Shooter attribution changed
6. Team attribution changed
7. Shot → turnover, or turnover → shot
8. Dunk classification added or removed

**One deliberate addition — a ninth category.** Made free throw added, removed or
flipped is flagged as market-impacting, because it moves **Timeframe Both Teams
To Score** (any made FT is a score inside a window). It is not in the spec's list
of eight, so it is gated behind `FLAG_SCORING_FT_CHANGES = True` in SECTION 1 —
set it to `False` to get exactly the eight canonical categories.

**Changes deliberately treated as non-impacting** (logged, visible under *All
Corrections*, never banner-worthy): period adjusted, clock adjusted, score
adjusted, shot type detail reworded, description reworded.

**False-positive sources that are actively suppressed:**

- **Player name text.** Names come from an id→name map that fills in as the
  boxscore publishes, so the *text* changes with no correction behind it.
  `player_id` is the sole authority for shooter attribution (category 5); the
  name string is excluded from comparison entirely.
- **Whitespace / formatting.** All text is normalised (`"Bad Pass\nTurnover"` →
  `"Bad Pass Turnover"`) before fingerprinting.
- **Coordinates, shot distance, assist credit.** Excluded from the fingerprint —
  they churn constantly and move no market in this tool.
- **Truncated payloads.** If the feed returns fewer than
  `REMOVAL_SANITY_RATIO` (90%) of the previous play count, removal detection is
  **skipped for that cycle** and a feed warning shows in the banner, rather than
  logging every missing play as a correction.
- **The first poll never logs.** It establishes the baseline snapshot and returns
  nothing, so pressing Track Game mid-game does not dump the whole first half
  into the log as insertions.

**What the tool cannot see:**

- **Corrections ESPN never publishes.** If the official scorer fixes a stat and
  ESPN does not update its play-by-play, there is nothing to diff. This tool
  detects *feed* corrections, which is a subset of *stat* corrections.
- **Corrections that landed before you pressed Track Game.** Baseline is taken at
  that moment.
- **Silent id reuse.** If ESPN were to reassign an existing play id to a
  different event without changing the fingerprint's market fields, that would
  read as unchanged. Not observed, but not detectable either.
- **Corrections after the game goes final.** Polling stops when the game is
  final, so post-game stat corrections (the most common kind in the NBA) are not
  caught. Keep the tab open and press **Refresh now** if you need to check.

**The log is append-only.** A correction is never removed because the feed changed
again later, and an event corrected twice logs twice — dedupe is on the
*transition* (`event_id | original | updated | impacting`), not on the event. The
same correction will not re-log on every 4s refresh, and the banner will not
re-flash the same correction; it stays visible for `CORRECTION_BANNER_SECONDS`
(120s) then clears.

### 4.4 Timeframe Both Teams To Score

- Windows are **upper-bound inclusive**: `12:00` lands in `12:00–11:01`, `11:01`
  lands in `12:00–11:01`, `11:00` lands in `11:00–10:01`. `int((720 - clock) // 60)`.
- The final window `1:00–0:00` is inclusive of `0.0`, so a buzzer-beater counts.
- Any **made** score counts — field goal or free throw — which is what the market
  prices.
- A window shows `-` until it is fully complete, then `Yes` or `No`. Completion is
  taken from the lowest clock reached in that period, an explicit end-of-period
  play, or the game being final.
- **ESPN's sub-minute clock has no colon** (`"51.7"` not `"0:51.7"`). This is
  handled, but it is the single most likely place a provider change breaks
  bucketing — every event in the last minute of a quarter depends on it.
- Overtime uses 5-minute periods (5 windows), regulation 12 (12 windows).
- **The table drops the Quarter column** the selectbox above it already states, and
  is rendered narrow rather than full width by `render_timeframe`. The dropped
  column and the muted Yes / No colours are presentation only: `timeframe_table`
  still returns the same three keys and still decides every verdict.
- **The highlighted row** is the window in play, or the last completed one if the
  quarter is over — derived in the UI from `period_progress` and `window_index`,
  the same two helpers the table itself uses.

### 4.5 Other

- **Second-half markets** activate once a Q3 event appears in the feed, and read
  only from period ≥ 3 events.
- **Workbook vs prompt conflict — timeframe row order.** The mock-up sheet lists
  the windows reverse-chronologically (`1:00–0:00` at the top); the written spec
  lists them chronologically. Chronological is the default, per the written spec.
  Set `TIMEFRAME_ORDER = "reverse"` in SECTION 1 to match the sheet.
- **Number of recent attempts** is one constant: `RECENT_FG_COUNT = 3`.
- **Do not drop `.block-container { padding-top }` below ~3.5rem.** Streamlit's top
  bar overlaps the main block rather than sitting in flow, so at the sibling tools'
  `1rem` the first element on the page — the scoreline — renders underneath it and
  cannot be scrolled to, because the page is already at scroll 0. Hiding the bar is
  not the fix: it holds the sidebar toggle, which is the only way to reopen a
  collapsed sidebar.
- **Rate limits.** ESPN's public endpoints are unauthenticated and undocumented,
  with no published limit. Fetches are cached (`scoreboard` 20s, `summary` ~3s,
  `roster` 1h) and a `429` surfaces as a "DATA DELAY" banner rather than a
  traceback. Nothing is requested until tracking starts.
- **All markets in the spec are supported by the data structures and the
  correction engine**, but this first version surfaces the tables the workbook
  draws (first-event, recent attempts, key players, timeframe, second half). The
  Possession Result and Next-FG-Attempt markets resolve from the same
  `GameEvent` primitives in SECTION 5 and are not yet given their own tables.

### 4.6 Players on the floor

At the top of the Live tab: the five on the floor per team, plus the last two
players substituted off. Not in the spec — added because subs are entered by hand
in the trading system and this is the glance that tells you what to change.

**It is derived, not reported.** ESPN publishes no on-court field on any endpoint
reachable from here: the `summary` payload, the core-API competition object,
`/situation` and the per-competitor `/roster` were all checked, and the only
lineup information in any of them is the boxscore `starter` flag. The five are
therefore the starters with every substitution play applied in feed order
(`type.id` 584, `participants[0]` on, `participants[1]` off).

**How well that holds up**, measured over 9 games / 3,746 plays / 480
substitutions before it was built:

- every substitution had exactly two participants, and `[0]`/`[1]` matched the
  description's "X enters the game for Y" in **480/480** cases;
- both teams held exactly five players at all **7,492** team-checkpoints, with no
  substitution ever taking off a player the reconstruction did not have on;
- **2,335 of 2,338** single-actor plays were by a player it had on the floor.

**The three exceptions are the limitation.** All three are same-clock ordering
ties — the feed lists the substitution just ahead of one last play by the man
going off (Q4 30.1 of ATL@CLE: "Dean Wade enters the game for Donovan Mitchell",
then Mitchell's turnover at the same 30.1). For a moment the panel shows a player
as off who then records a play. It resolves as soon as play moves on. The panel
shows the entry clock on a freshly substituted player rather than trying to
reorder the feed.

Other things to know:

- **It inherits 4.1 entirely.** No boxscore starters means no trustworthy five,
  and the panel prints the lineup source whenever it is not the boxscore flag.
  Before tip-off it shows the starting five with no substitution clocks.
- **If a substitution takes off a player the panel did not have on**, the swap is
  still applied but the team is labelled *Lineup unverified* with a count. That
  never happened in the measured games; if it appears, check the boxscore.
- **Shirt numbers come from the boxscore, not from `teams/{id}/roster`.** The
  roster endpoint is the *current* roster, so it cannot number a player who has
  since left the team: for game 401859966 it knows only 11 of the 15 athletes who
  played for San Antonio (no Olynyk, Waters, Biyombo or Plumlee) and carries no
  jersey at all for 5 of the 19 it does list, while the boxscore has one for
  **30/30** athletes in the game. Chips were rendering as bare names because of
  it. The rosters are still the fallback, for the pre-tip window before a boxscore
  exists.
- **One panel per team, five rows, full names.** It is read at a glance from a
  distance while players are being subbed by hand in another system, so
  legibility beats compactness. Rows are in shirt-number order, so the same team
  reads the same way twice and only the arrow moves.
- **A green up arrow marks the player who just came on, and it never expires.** It
  comes from `latest_substitutions`, which returns the most recent substitution
  *break* — every substitution sharing the last one's period and clock, because a
  timeout change is five or six plays at the same clock reading and showing one of
  them would hide the rest. Nothing times out: the arrow stands until the feed
  publishes the next substitution, because the trader may be mid-entry elsewhere
  when it lands. `SUB_HIGHLIGHT_MAX = 6` caps one break. The arrow slot is emitted
  on every row, empty where there is nothing to mark, so shirt numbers stay in one
  column and the five never shuffle sideways when a substitution lands.
- **Recent substitutions are listed under each team's own five**, newest first, as
  `<clock> <period> ↑ <in> ↓ <out>` — pairs, not a list of who left, because the
  pair is what gets copied into the other system. `FLOOR_RECENT_SUBS = 3` sets the
  count, enough to cover a whole timeout change for one team. Every row reads at
  the same contrast; the order already says which is newest.
- **No orange in the substitution UI.** Two rounds of review removed it: first a
  full-width orange alert bar per team (more alarm than a routine event deserves,
  and detached from the team it belonged to), then the orange row fill and the
  brighter newest row. Direction is now carried by colour-coded arrows —
  `ARROW_IN`/`ARROW_OUT`, coloured with the greens and reds already used by the
  event feed and the timeframe table. Orange still
  means *this wants your attention*: the correction and key-player banners.
- Both live **inside the team column, below the five** — never above. Above them,
  every substitution would push the five down the screen, the one thing a panel
  meant to be glanced at cannot do.
- **Unverified against a live feed.** Everything above is measured on completed
  games. If a live `summary` response returns only a trailing window of plays
  rather than the whole history, the five would have to be accumulated across
  polls instead of re-derived. The truncation guard in 4.3 implies full history is
  normally returned, but that has not been confirmed mid-game.

### 4.7 Remaining false insertions (known, not fixed)

Keeping substitutions out of the sequence and live-edge arithmetic (4.3) removed
80 of 279 false insertions in the 9-game replay. The other 199 come from two
assumptions the feed violates, both still present:

- **A coach's challenge re-issues the overturned play with a far-later sequence
  number at the same clock.** In PHI@WSH the Embiid foul and turnover come back as
  `seq 327/328` among neighbours numbered 271–281. That poisons the watermark for
  the rest of the game — 227 of that game's 549 plays then sit below it, and the
  replay logs 117 false insertions in that one game.
- **`_is_retroactive` measures a new play against a live edge computed from a
  window that includes that play.** So any poll bringing in more than
  `RETROACTIVE_TOLERANCE_SECONDS` of game clock — a rate-limit skip, a timeout, a
  long stoppage — flags everything except the newest play. At a 12-play cadence
  that is 70–94 rows per game. Comparing against the *previous* poll's edge
  instead would fix it.

Until these are addressed, treat a burst of INSERTED rows sharing one clock
reading as suspect, and check *Market-Impacting Corrections Only*.

### 4.8 Next Field Goal market anchors

The market is *Next Field Goal after `<score>`*, and that score is a **checkpoint
only a made field goal moves**. Free throws and technicals move the scoreboard
without opening a new market; a missed field goal moves nothing. So the score shown
against an attempt is not the live board — it is the board as it stood after the
previous made field goal. `fg_market_events` derives it and stores four values per
attempt: `market_anchor_score`, `event_result`, `post_event_score`, and
`new_market_anchor_score` (made field goals only).

- A row shows **three things and nothing else**, under one header per panel —
  **After / Result / Time** — so the word "After" is written once instead of on every
  line:

  ```
  AFTER     RESULT           TIME
  14-5      Made 2        6:20 3Q
  12-5      Missed 3      7:04 3Q
  ```

  Header and rows share a 3-column grid and the same padding and 3px left border, so
  the columns line up. Makes and misses take the same shape.
  `new_market_anchor_score` is stored but **deliberately not rendered**: on a make it
  is the same number as the next row's anchor, and printing it on every make made the
  panel unreadable. **Scores are away-home**, as everywhere else in the tool.
- **The panels carry no caption, and one consequence is worth knowing.** The
  checkpoint the market is open on *right now* is the newest made field goal's
  post-score, so it is never itself an `After` value and appears nowhere on the
  Live tab — it will show up as the next attempt's `After` once that attempt happens.
  The caption used to state it; it was removed on the user's instruction along with the
  rest of the explanatory text, and `open_market_anchor()` was deleted with its only
  call site. Reinstate both if the open checkpoint is wanted on screen again.
- **The team abbreviation appears only where a panel mixes teams.** `render_feed`
  takes `include_team`, on for the middle *Made Field Goals* panel (`SA Made 2`) and
  off for the two team panels, where every row would carry the same abbreviation and
  it says nothing. One helper, `_result_text`, defines `Made 2` / `Missed 3`; the
  prefixed form is `event_result` and the bare form is the `result_no_team` property,
  so the two can never drift.
- **The checkpoint is read off the feed, never added up from field-goal points.**
  Measured on game 401859966 (498 plays, 164 field-goal attempts by the app's own
  classifier, 72 made, 48 free throws): **20 of the 72 new checkpoints would be
  wrong if computed from field-goal points alone**, by 1 to 4 points. The game's
  first basket is the clearest case — Towns makes 2 on a 0-0 checkpoint and the new
  checkpoint is 2-2, not 0-2, because Fox had already made two free throws.
- **There is no arithmetic to do, because ESPN's per-play score is post-play.**
  Every one of that game's 109 scoring plays had (its own score − the previous
  play's score) equal to its own `scoreValue`, 0 exceptions, and the last play's
  106-107 is the official final. So `ev.away_score` / `ev.home_score` *is*
  `post_event_score`.
- **What this fixed:** 25 of the game's 92 missed attempts were displaying a board
  ahead of the market anchor by 1–4 points, i.e. naming a market that did not
  exist. Made rows were showing the score after the basket, which is the market that
  *opened*, where the row is about the one that *settled*.
- **Anchors are derived over the whole game and only then filtered per panel.** A
  team's checkpoint is moved by its opponent's baskets too, so deriving from one
  team's attempts alone would name the wrong market.
- **Corrections are handled by re-derivation, not patching.** Like every other
  derived view here it is a pure function of the event list, recomputed each poll,
  so a correction that adds, removes or re-values a made field goal re-anchors
  everything after it on the next refresh. This relies on ESPN restating the
  affected plays' `awayScore` / `homeScore` when it restates the play — consistent
  with the whole play list being re-sent each poll, but **not yet observed on a
  live correction**.
- **Guard:** if a made field goal's board is *below* the current checkpoint, the
  checkpoint is not moved and the row is flagged `score?`. That means an absent
  `awayScore` / `homeScore`, which normalises to 0. It happened 0 times in the
  measured game; the alternative was walking the market backwards.
- Checked over the same game: 0 chain breaks (every row's anchor equals the
  previous made row's post score), 0 makes without a new checkpoint, 0 misses with
  one, and the final checkpoint equals both the last made field goal's board and
  the official final score.
- **The first-shot tables still show the plain board**, because they are not market
  names. The Key Player Tracker shows an anchor, but a per-player one that is not a
  market name either — see 4.9.

### 4.9 Key player cards, and the expand toggle

Each key player card shows the After / Result / Time table from 4.8: one row closed,
all of that player's made field goals when expanded.

- **The anchor here is per player, not the game-wide checkpoint — on purpose, and it
  is the one place in the tool where "After" does not name a market.** The panels
  answer *which open market did this attempt settle*, so their anchor moves on the
  opponent's baskets too. A card answers *where the game stood each time this player
  scored*, so `player_market_events` runs its own walk that only advances on his own
  makes. Two consequences, both intended and both requested:
  - **the oldest row is always `0-0`**, because there is no basket of his behind it;
  - **an anchor may never have been an open market name.** He scores, board 2-0; the
    opponent scores, board 2-3; he scores again. That row reads `2-0`, the board after
    his own last basket, while the market open at that moment was "after 2-3". If you
    need exact market names, the panels below the cards are where they are exact.
- `score_suspect` still guards the row, but it can now only fire on a feed that
  contradicts itself: the anchor is this player's own previous post-score and the
  board cannot fall.
- **Closed and open are the same table**, one row or all of them. The card used to
  carry a summary line with the live scoreboard (`format_event_line`, now deleted —
  it had no other caller). That would have contradicted row one of the list it
  expands into: the board after a basket is not the checkpoint the basket settled,
  and the two differ by however many free throws fell between them. Showing both was
  a way to make a trader distrust the panel.
- **The toggle cannot disturb the flash alert**, which was the requirement:
  - alert *creation* is idempotent — it fires only when `kp_seen[pid]` differs from
    the newest make's event id, and writes the id as it fires, so a rerun from a
    click finds nothing new;
  - alert *expiry* is wall-clock (`now - alert["ts"] <= KEY_ALERT_SECONDS`), not a
    tick count, so extra reruns neither shorten nor extend the 30s;
  - a click makes no request — `fetch_summary` is cached at `REFRESH_SECONDS - 1`, so
    the rerun sees identical events and `record_corrections` diffs against an
    already-current snapshot.
- **The list is built into the card's own HTML** with `feed_html` rather than a second
  `st.markdown`, so `.kp.hot` frames the baskets it is flashing about and not just
  the name.
- **Open state lives in `session_state.kp_open`, not in the widget.** An `st.expander`
  would have been the obvious control and the wrong one: Streamlit identifies it by
  label and position, so a label carrying live data (`Wembanyama — 7 made`) is a new
  element on every basket and snaps shut. That is the open bug on the NFL tool's
  Drives tab. Holding the state ourselves also means the 15s autorefresh can't close
  it.
- **The button sits on the card's own title line, top right, and is overlaid there by
  CSS.** A Streamlit widget cannot be nested inside a block of our own HTML, so it is
  emitted immediately *before* the card and then pulled back over it by
  `div[class*="st-key-kpall_"]`. Three numbers, all in that rule's comment:
  - `margin-bottom: -36px` cancels the button out of the flow — its own 20px height
    plus the one extra 1rem `stVerticalBlock` gap that inserting an element costs
    (Streamlit's flex gap does not collapse with margins). Without it, a player with
    the toggle would sit lower than a player without one. **This is the only value
    inferred from Streamlit's own layout, so it is the one to change if the button
    ever appears clear above or below the card rather than on it.**
  - `transform: translateY(9px)` drops the button onto the name line. A transform is a
    paint offset, so it costs nothing in layout. 9px = the card's 2px border + 8px
    padding, less half the 2px by which the button overhangs the 18px name line.
  - `margin-right: 14px` **on the button, not the wrapper** — Streamlit gives element
    containers `width: 100%`, so padding or margin on the wrapper only overflows it to
    the right and moves nothing. 14px is the card's border + padding again, so the
    button's right edge meets the card's inner edge.
  Measured in a headless-Chrome mock of the real wrapper chain (flex column,
  `gap: 1rem`): button centred on the name line to 0px, fully inside the card box,
  inset 14px, and the card sitting exactly where it does with no button
  (21px below the heading, 22px between cards — both unchanged). `.kp .nm` reserves
  66px on the right so a long name wraps rather than running under the button, and
  `pointer-events` is off on the full-width wrapper so it cannot swallow clicks on the
  name.
- **One rerun per click**, and now for free: the button is rendered before the card, so
  the click is read before the card is built. No `st.rerun()`, and the `st.empty()`
  placeholder the previous below-the-card layout needed is gone.
- **No `help=` on the button.** Streamlit positions the `help` tooltip on hover and it
  does not reliably tear down when the page reruns underneath the cursor, so on a tab
  that repolls every 15s it gets left behind on screen. Reported that way by the user;
  the label (`All 7` / `Hide`) says enough on its own.
- The button is hidden until a player has more than one make, since a one-row list is
  what the closed card already shows.
