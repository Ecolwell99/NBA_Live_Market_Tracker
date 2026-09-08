"""
NBA Live Market Tracker
=======================

Internal tool for sportsbook traders / QC analysts.

Purpose
-------
1. Track live NBA play-by-play.
2. Result the first-event, timeframe and second-half markets that can be
   determined directly from play-by-play.
3. Immediately flag stat corrections, and separate the market-impacting ones
   from harmless feed churn.

Layout follows "NBA Tracker Layout.xlsx" (tabs: Prematch / Live / Stat Corrections).

Data source
-----------
ESPN's public NBA endpoints (see SECTION 3). The official NBA CDN feed
(cdn.nba.com/static/json/liveData/...) is structurally richer, but it is served
behind Akamai and returns HTTP 403 from corporate / datacenter egress, and
stats.nba.com times out entirely from the same network. ESPN's `summary`
endpoint carries everything the markets in this tool need:

    shootingPlay + pointsAttempted -> field goal attempt and 2 vs 3
    scoringPlay                    -> made vs missed
    type.text                      -> shot detail, incl. dunk classification
    participants[0].athlete.id     -> shooter (correct even on blocked shots)
    team.id                        -> team attribution
    period.number / clock          -> timeframe bucketing
    awayScore / homeScore          -> score at the time of the event
    boxscore ... starter           -> real starters (5 per team)
    type.id 584 + participants     -> substitution: [0] comes on, [1] goes off,
                                      which is what the on-floor five is built from

All source-specific parsing is confined to SECTION 3 (fetch) and SECTION 4
(normalise). Everything downstream works only on the neutral `GameEvent`
dataclass, so swapping providers means rewriting those two sections only.
"""

from __future__ import annotations

import html
import json
import re
import time
from copy import deepcopy
from dataclasses import dataclass, asdict, field
from datetime import date, datetime
from pathlib import Path
from typing import Any, Callable, Iterable, Sequence

import requests
import streamlit as st

try:
    from streamlit_autorefresh import st_autorefresh
except ImportError:  # pragma: no cover - surfaced in the UI instead of crashing
    st_autorefresh = None


# ===========================================================================
# SECTION 1 - CONFIGURATION
# Everything a trader might want to retune lives here.
# ===========================================================================

APP_TITLE = "NBA Live Market Tracker"

# "3 most recent FG attempts per team" - change this one value to show more.
RECENT_FG_COUNT = 3

# Auto-refresh cadence while a game is tracked (seconds). Spec: 3-5s.
REFRESH_SECONDS = 4

# Slower cadence before tip-off: nothing can change yet, so do not hammer the feed.
PREGAME_REFRESH_SECONDS = 15

# How long a detected stat correction stays in the status banner (seconds).
CORRECTION_BANNER_SECONDS = 120

# How long a key-player made-basket flash alert stays on screen (seconds).
KEY_ALERT_SECONDS = 30

# Key players tracked per team.
KEY_PLAYERS_PER_TEAM = 2

# On-floor panel (top of the Live tab). How many recent substitutions to list under
# each team's five, and how many pairs of one simultaneous change can highlight at
# once (a timeout change is five or six substitutions at the same clock reading).
#
# There is deliberately no expiry on the highlight: it is driven by the most recent
# substitution in the feed and stands until the next one arrives. The trader may be
# mid-entry in the other system when a sub lands, and a highlight that had timed out
# by the time they looked up would be worse than none.
FLOOR_RECENT_SUBS = 3
SUB_HIGHLIGHT_MAX = 6

# Timeframe table ordering. "chronological" -> 12:00-11:01 first (matches the
# written spec). "reverse" -> 1:00-0:00 first (matches the mock-up sheet).
TIMEFRAME_ORDER = "chronological"

# Regulation / overtime period lengths in seconds.
REGULATION_PERIOD_SECONDS = 720
OVERTIME_PERIOD_SECONDS = 300

# Extra correction category beyond the eight canonical ones: a made free throw
# being added or removed changes "Timeframe Both Teams To Score". Set to False
# to restrict market-impacting flags to the eight categories in the spec.
FLAG_SCORING_FT_CHANGES = True

# Guard against false "event removed" corrections when the feed briefly returns
# a truncated payload: skip removal detection if the play count drops by more
# than this fraction in a single poll.
REMOVAL_SANITY_RATIO = 0.90

# A newly appeared event whose game clock sits this far BEHIND the live edge is
# a retroactive insertion (a correction), not the game moving forward. Kept
# generous because plays around the live edge can arrive slightly out of order.
RETROACTIVE_TOLERANCE_SECONDS = 45

HTTP_TIMEOUT = 12
DASH = "-"

# Substitution direction markers, coloured green / red by .arw in the CSS. Literal
# glyphs like the sidebar's play triangle rather than HTML entities, so they also
# read correctly in a plain `st.caption` or a log line if one ever needs them.
ARROW_IN = "↑"    # up, green: coming on
ARROW_OUT = "↓"   # down, red: going off

# Cooldown after an HTTP 429: skip this many refresh ticks before polling again,
# as the NHL / NFL / MLB tools all do. Without it the autorefresh keeps hammering
# a feed that has already told us to back off.
RATE_LIMIT_SKIP_TICKS = 2

# On-disk state so a hard browser refresh (new Streamlit session) does not
# forget an already-detected correction. Best-effort only.
STATE_DIR = Path(__file__).resolve().parent / ".tracker_state"


# ===========================================================================
# SECTION 2 - STYLING + TAB IDENTIFIERS
#
# The house style, ported from the NFL / NHL tools: a near-empty stylesheet,
# with everything else drawn from the active Streamlit theme
# (var(--text-color), var(--secondary-background-color)) instead of a private
# palette. Hardcoding a palette here is exactly what made this tool look
# unrelated to the others. Strong colour is reserved for the status banner and
# the Yes/No pills, both of which live in the shared components in SECTION 8.
# ===========================================================================

# Tab identifiers. Defined this early because DEFAULT_STATE (SECTION 7) seeds
# `active_tab` with one of them. They are opaque keys, never labels - see
# `render_tab_strip` for why that distinction matters.
TAB_PREMATCH = "prematch"
TAB_LIVE = "live"
TAB_CORRECTIONS = "corrections"
TABS = (TAB_PREMATCH, TAB_LIVE, TAB_CORRECTIONS)

CSS = """
<style>
/* Tighten default Streamlit padding. padding-top has to stay above the height of
   Streamlit's own top bar (2.875rem), which overlaps the main block rather than
   sitting in flow: at the sibling tools' 1rem the first element on the page — here
   the game header — renders underneath it and cannot be scrolled to, because the
   page is already at scroll 0. 3.5rem clears it with ~10px to spare. Do not hide the bar
   instead; it holds the sidebar toggle, which is the only way back when the sidebar
   is collapsed. */
.block-container { padding-top: 3.5rem; padding-bottom: 1rem; }
/* Remove red underline from metric delta */
[data-testid="stMetricDelta"] svg { display: none; }

/* --- section labels ------------------------------------------------------
   Full contrast (no opacity), uppercase, and a hairline rule underneath - that
   is the whole treatment. No accent bar: it was one more coloured thing to look
   past on a screen that is read at a glance.

   The top margin is unconditional. It used to be overridden by
   `.sect:first-child { margin-top: 4px }`, which was meant for the first heading
   on a tab but matched EVERY heading, because Streamlit wraps each st.markdown in
   its own container and the div is always the only child. The gap above a heading
   was therefore whatever the element above happened to leave, which is why some
   sections sat tight against the block above them and others did not. */
.sect {
  font-size: 13px; text-transform: uppercase; letter-spacing: .06em;
  font-weight: 800; color: var(--text-color);
  margin: 20px 0 8px 0; padding: 0 0 5px 0;
  border-bottom: 1px solid var(--secondary-background-color);
}
.subsect {
  font-size: 13px; font-weight: 800; color: var(--text-color);
  letter-spacing: .01em; margin: 10px 0 5px 0;
}
.note { font-size: 11px; color: var(--text-color); opacity: .55; margin: 3px 0 0 0; }

/* --- game header ---------------------------------------------------------
   The one block that answers "what am I looking at": both teams, both scores,
   period, clock, and when the feed last came back. Grid rather than flex so the
   scores stay put as the clock text changes width. */
.gh {
  display: grid; grid-template-columns: 1fr auto 1fr; align-items: center; gap: 14px;
  padding: 12px 18px; margin: 0 0 10px 0; border-radius: 12px;
  background: rgba(128,128,128,0.05);
  border: 1px solid var(--secondary-background-color);
  color: var(--text-color);
}
.gh .tm { display: flex; align-items: baseline; gap: 12px; min-width: 0; }
.gh .tm.h { justify-content: flex-end; }
.gh .nm { font-size: 16px; font-weight: 800; }
.gh .sc { font-size: 34px; font-weight: 900; line-height: 1; }
.gh .mid { text-align: center; }
.gh .ck { font-size: 20px; font-weight: 800; white-space: nowrap; }
.gh .pd { font-size: 13px; font-weight: 700; opacity: .7; margin-left: 6px; }
.gh .up { font-size: 11px; opacity: .55; margin-top: 3px; white-space: nowrap; }

/* --- compact status line (the normal, nothing-wrong state) ---------------
   The big house banner is kept for every state that needs attention; STATUS: OK
   does not, and at 22px it was the loudest thing on the page. */
.statusline {
  display: flex; align-items: center; gap: 8px; margin: 0 0 12px 0;
  font-size: 12px; font-weight: 700; color: var(--text-color); opacity: .85;
}
.statusline .dot {
  width: 8px; height: 8px; border-radius: 50%; background: #2ecc71; flex: 0 0 auto;
}

/* --- alerts (key-player flash) --- */
.alert {
  font-size: 14px; font-weight: 700; padding: 10px 14px; border-radius: 8px;
  margin-bottom: 6px;
  background-color: #3a1600; color: #ffd966; border: 2px solid #ff9900;
}

/* --- event feed rows (Next Field Goal markets) ----------------------------
   Each row is a market, not a play: the anchor score it was priced under, the
   result, the game time. Three things only, so the words "After" and the team
   abbreviation live in the header instead of repeating down every row. Header and
   rows share the same grid and the same padding + 3px left border so the three
   columns line up; the clock is right-aligned and dimmed because it is the least
   of the three. */
.feed { margin-bottom: 2px; }
.fhead, .frow {
  display: grid; grid-template-columns: 58px 1fr auto; gap: 8px;
  align-items: baseline; padding: 5px 10px; border-left: 3px solid transparent;
}
.fhead {
  font-size: 10px; font-weight: 800; letter-spacing: .06em; text-transform: uppercase;
  color: var(--text-color); opacity: .45; padding-bottom: 1px;
}
.frow {
  font-size: 13px; margin-bottom: 3px;
  background: rgba(128,128,128,0.06); color: var(--text-color);
}
.frow.made { border-left-color: #00cc44; }
.frow.miss { border-left-color: #cc2200; }
.frow .anch { font-weight: 700; }
.frow .res { font-weight: 700; }
.frow .mk { opacity: .45; }
.frow .mk.bad { color: #e08a80; opacity: 1; font-weight: 800; }
.fhead .sc, .frow .sc { text-align: right; white-space: nowrap; }
.frow .sc { opacity: .6; }
.frow.empty { display: block; font-size: 13px; opacity: .5; }

/* --- on-floor five (top of the Live tab) ---------------------------------
   One panel per team, five rows each, full names at 14px: this is read at a
   glance from a distance while subbing players by hand in another system, so
   legibility beats compactness. Direction is carried by a green up arrow rather
   than by an orange fill: the fill read as a warning for what is a routine event,
   and green / red for on / off is the same language as .frow.made / .frow.miss
   and the timeframe Yes / No. Greens and reds are the existing ones. */
.lu {
  border: 1px solid var(--secondary-background-color); border-radius: 10px;
  overflow: hidden; margin: 2px 0 4px 0;
}
.lurow {
  display: flex; align-items: center; gap: 10px;
  padding: 7px 12px; font-size: 14px; color: var(--text-color);
  border-left: 3px solid transparent;
}
.lurow + .lurow { border-top: 1px solid rgba(128,128,128,0.12); }
.lurow:nth-child(odd) { background: rgba(128,128,128,0.04); }
.lurow .num {
  min-width: 30px; text-align: right; font-size: 12px; font-weight: 800; opacity: .55;
}
.lurow .nm { font-weight: 700; }
.lurow .cl { margin-left: auto; font-size: 12px; font-weight: 700; opacity: .75; }
/* After the striping, so an entering player wins on equal specificity. */
.lurow.in { border-left-color: #6fd68d; }
.lurow.in .cl { color: #6fd68d; opacity: 1; }

/* Shared by both blocks so an up arrow means the same thing in each. */
.arw {
  min-width: 10px; text-align: center; font-weight: 900; line-height: 1;
}
.arw.in { color: #6fd68d; }
.arw.out { color: #e08a80; }

/* --- recent substitutions (under each team's own five) -------------------
   Under the team panel rather than in a full-width banner: the trader works one
   team at a time while subbing by hand, so the change belongs next to that
   team's list, and a bar across the page was far more alarm than a routine
   substitution deserves. Every row reads at the same contrast - the newest was
   brighter than the rest for one round and the difference just looked like a
   rendering fault, since the list is already ordered newest first and the arrow
   in the panel above marks the current change.
   Persistent on purpose - no expiry, same as the highlight. */
.subs { margin: 5px 0 10px 0; }
.subhd {
  font-size: 10px; font-weight: 800; letter-spacing: .08em; text-transform: uppercase;
  color: var(--text-color); opacity: .45; margin: 0 0 2px 2px;
}
.subrow {
  display: flex; align-items: center; flex-wrap: wrap; gap: 6px;
  padding: 2px; font-size: 12px; color: var(--text-color); opacity: .85;
}
.subrow .gt { min-width: 54px; font-weight: 700; opacity: .7; }
.subrow .io { font-weight: 700; }
.subrow .io.out { margin-left: 2px; }

/* --- timeframe table -----------------------------------------------------
   Its own narrow table rather than `html_table`: two columns of short values do
   not need the full page width, and the loud Yes / No pills read as alarms in a
   twelve-row grid where most rows are routine. */
.tfwrap { margin: 2px 0 0 0; }
.tf { border-collapse: collapse; width: auto; min-width: 300px; max-width: 380px; }
.tf th {
  font-size: 11px; text-transform: uppercase; letter-spacing: .06em; font-weight: 800;
  text-align: left; padding: 5px 16px 5px 12px; white-space: nowrap;
  color: var(--text-color); opacity: .7;
  border-bottom: 2px solid var(--secondary-background-color);
}
.tf td {
  font-size: 13px; padding: 5px 16px 5px 12px; white-space: nowrap;
  color: var(--text-color); border-left: 3px solid transparent;
}
.tf tr:nth-child(even) td { background: rgba(128,128,128,0.05); }
.tf td.v { font-weight: 800; }
.tf td.yes { color: #6fd68d; }
.tf td.no { color: #e08a80; }
.tf td.na { opacity: .35; font-weight: 700; }
/* Current / most recently completed window. After the striping, same reason. */
.tf tr.cur td { background: rgba(255,153,0,0.10); font-weight: 800; }
.tf tr.cur td:first-child { border-left-color: #ff9900; }

/* --- key player card --- */
.kp {
  background: rgba(128,128,128,0.06); border: 2px solid transparent;
  border-radius: 8px; padding: 8px 12px; margin-bottom: 6px;
}
.kp .nm { font-size: 13px; font-weight: 700; color: var(--text-color); }
.kp .ln { font-size: 13px; color: var(--text-color); opacity: .85; margin-top: 2px; }
.kp .ln.none { opacity: .45; }
.kp.hot { border-color: #ff9900; background: rgba(255,153,0,0.12); }

/* --- tab strip -----------------------------------------------------------
   Cosmetic only - makes the keyed radio in `render_tab_strip` read as a tab
   strip. Scoped to that widget via the .st-key-active_tab wrapper class
   Streamlit adds for any keyed widget, so it cannot leak onto the sidebar. If
   Streamlit changes these internal selectors nothing breaks: it degrades to a
   plain horizontal radio, which still navigates correctly. */
.st-key-active_tab div[role="radiogroup"] { gap: 4px; }
.st-key-active_tab div[role="radiogroup"] > label {
  padding: 6px 16px; border-radius: 8px 8px 0 0;
  font-weight: 700; font-size: 15px; border-bottom: 2px solid transparent;
}
.st-key-active_tab div[role="radiogroup"] > label:hover {
  background: var(--secondary-background-color);
}
.st-key-active_tab div[role="radiogroup"] > label:has(input:checked) {
  background: var(--secondary-background-color);
  border-bottom: 2px solid #ff4b4b;
}
/* Hide the radio dot so it reads as a tab, not a form control. The :has(input)
   guard matters: without it, a future Streamlit DOM change could match the div
   holding the tab TEXT and hide the labels entirely. */
.st-key-active_tab div[role="radiogroup"] > label > div:first-child:has(input) {
  display: none;
}

/* --- small secondary button ----------------------------------------------
   Scoped to the Edit Key Players toggle by its keyed-widget wrapper class, the
   same mechanism as the tab strip above. Streamlit has no small-button size, so
   this is the only way to get one; if the selector ever stops matching, the
   button reverts to full size and still works. */
.st-key-edit_kp button {
  padding: 2px 12px; min-height: 0; font-size: 12px; font-weight: 700;
}
.st-key-edit_kp button p { font-size: 12px; margin: 0; }
</style>
"""


# ===========================================================================
# SECTION 3 - DATA SOURCE (ESPN)
# The only place that knows about ESPN's URL shapes and JSON keys, alongside
# SECTION 4. Swap this out to change providers.
# ===========================================================================

ESPN_BASE = "https://site.api.espn.com/apis/site/v2/sports/basketball/nba"

# DO NOT put a browser User-Agent here. ESPN sits behind Akamai, which matches the
# LEADING token of the UA against known HTTP-client signatures and serves 403 to
# anything else - including a spoofed Chrome string. Measured 2026-09-06, stable
# over repeated rounds, identical on all four endpoints we call:
#
#   python-requests/2.31.0                        200      <- requests' own default
#   python-requests/2.31.0 nba-live-tracker/1.0   200      <- prefix, then our name
#   curl/8.7.1  libcurl/8.0  Python-urllib/3.11   200
#   Mozilla/5.0 ... Chrome/124.0.0.0 ...          403      <- the old value
#   nba-live-tracker/1.0 python-requests/2.31.0   403      <- prefix must be FIRST
#   nba-live-tracker/1.0   "requests"   ""        403      <- no version, no match
#
# So: keep a real `python-requests/<version>` prefix, then identify ourselves after
# it. The version is read from the installed library rather than hardcoded, so the
# string can never claim a version we are not actually running.
_HTTP_HEADERS = {
    "User-Agent": f"python-requests/{requests.__version__} nba-live-tracker/1.0",
    "Accept": "application/json, text/plain, */*",
}


class DataSourceError(Exception):
    """Any failure talking to the feed. Rendered as a message, never a traceback."""


class RateLimitedError(DataSourceError):
    pass


@dataclass(frozen=True)
class TeamInfo:
    team_id: str
    abbr: str
    location: str
    name: str
    display_name: str


@dataclass(frozen=True)
class GameInfo:
    game_id: str
    start_iso: str
    state: str          # "pre" | "in" | "post"
    status_detail: str  # e.g. "7:00 PM ET", "Q3 4:12", "Final"
    period: int
    display_clock: str
    away: TeamInfo
    home: TeamInfo
    away_score: int
    home_score: int

    @property
    def label(self) -> str:
        return f"{self.away.abbr} @ {self.home.abbr}"

    @property
    def is_live(self) -> bool:
        return self.state == "in"

    @property
    def is_final(self) -> bool:
        return self.state == "post"


@dataclass(frozen=True)
class RosterPlayer:
    player_id: str
    name: str
    jersey: str
    position: str

    @property
    def label(self) -> str:
        bits = [b for b in (f"#{self.jersey}" if self.jersey else "", self.position) if b]
        return f"{self.name}" + (f"  ({' '.join(bits)})" if bits else "")


def _http_get_json(url: str) -> dict:
    try:
        resp = requests.get(url, headers=_HTTP_HEADERS, timeout=HTTP_TIMEOUT)
    except requests.exceptions.Timeout as exc:
        raise DataSourceError("Feed timed out.") from exc
    except requests.exceptions.RequestException as exc:
        raise DataSourceError(f"Network error reaching the feed: {exc.__class__.__name__}") from exc

    if resp.status_code == 429:
        raise RateLimitedError("Rate limited by the feed (HTTP 429). Backing off.")
    if resp.status_code == 404:
        raise DataSourceError("Game not found on the feed (HTTP 404).")
    if resp.status_code == 403:
        # Called out separately because the generic message sent us hunting for a bad
        # game ID when the real cause was the request headers. A 403 here is Akamai
        # bot filtering and is never about the game.
        raise DataSourceError(
            "Feed refused the request (HTTP 403) - bot filtering, not a bad game ID. "
            "Check the User-Agent in _HTTP_HEADERS."
        )
    if resp.status_code >= 500:
        raise DataSourceError(f"Feed is unavailable (HTTP {resp.status_code}).")
    if resp.status_code != 200:
        raise DataSourceError(f"Feed returned HTTP {resp.status_code}.")

    try:
        return resp.json()
    except ValueError as exc:
        raise DataSourceError("Feed returned a malformed (non-JSON) response.") from exc


def _team_from_raw(raw: dict) -> TeamInfo:
    return TeamInfo(
        team_id=str(raw.get("id", "")),
        abbr=(raw.get("abbreviation") or raw.get("shortDisplayName") or "").upper(),
        location=raw.get("location", "") or "",
        name=raw.get("name", "") or "",
        display_name=raw.get("displayName") or raw.get("name") or "",
    )


def _game_from_dict(g: dict) -> GameInfo:
    """Rehydrate a GameInfo from the plain-dict form kept in session_state."""
    return GameInfo(
        game_id=g["game_id"],
        start_iso=g["start_iso"],
        state=g["state"],
        status_detail=g["status_detail"],
        period=g["period"],
        display_clock=g["display_clock"],
        away=TeamInfo(**g["away"]),
        home=TeamInfo(**g["home"]),
        away_score=g["away_score"],
        home_score=g["home_score"],
    )


@st.cache_data(ttl=20, show_spinner=False)
def fetch_scoreboard(day: str | None) -> list[dict]:
    """Games for `day` (YYYYMMDD). Pass None to let the feed decide 'today'.

    Returns a list of plain dicts so Streamlit's cache can hash the result.
    """
    url = f"{ESPN_BASE}/scoreboard" + (f"?dates={day}" if day else "")
    payload = _http_get_json(url)

    games: list[dict] = []
    for event in payload.get("events") or []:
        comps = event.get("competitions") or []
        if not comps:
            continue
        comp = comps[0]
        status = comp.get("status") or event.get("status") or {}
        stype = status.get("type") or {}

        away_raw = home_raw = None
        away_score = home_score = 0
        for c in comp.get("competitors") or []:
            side = c.get("homeAway")
            try:
                score = int(float(c.get("score") or 0))
            except (TypeError, ValueError):
                score = 0
            if side == "home":
                home_raw, home_score = c.get("team") or {}, score
            elif side == "away":
                away_raw, away_score = c.get("team") or {}, score
        if not away_raw or not home_raw:
            continue

        games.append(
            {
                "game_id": str(event.get("id", "")),
                "start_iso": comp.get("date") or event.get("date") or "",
                "state": stype.get("state", "pre"),
                "status_detail": stype.get("shortDetail") or stype.get("description") or "",
                "period": int(status.get("period") or 0),
                "display_clock": status.get("displayClock") or "",
                "away": asdict(_team_from_raw(away_raw)),
                "home": asdict(_team_from_raw(home_raw)),
                "away_score": away_score,
                "home_score": home_score,
            }
        )
    return games


def list_games(day: str | None) -> list[GameInfo]:
    """Scoreboard as GameInfo, sorted live -> upcoming -> final."""
    out = [_game_from_dict(g) for g in fetch_scoreboard(day)]
    rank = {"in": 0, "pre": 1, "post": 2}
    out.sort(key=lambda g: (rank.get(g.state, 3), g.start_iso))
    return out


# --- manual game ID -------------------------------------------------------
# Traders need to track a game the scoreboard will not hand us: a game on a
# different slate, one the scoreboard has already dropped, or any game at all
# when the scoreboard endpoint itself is failing. The summary endpoint is keyed
# only on the event id, so it works in every one of those cases.

_GAME_ID_RE = re.compile(r"(\d{6,})")


def extract_game_id(text: str) -> str:
    """Accept a bare ESPN event id or a pasted ESPN game URL.

    'https://www.espn.com/nba/game/_/gameId/401810433/magic-grizzlies' -> '401810433'
    """
    found = _GAME_ID_RE.search(str(text or ""))
    return found.group(1) if found else ""


@st.cache_data(ttl=20, show_spinner=False)
def fetch_game_header(game_id: str) -> dict:
    """One game's identity + status from the summary endpoint's `header` block.

    Same plain-dict shape as `fetch_scoreboard` entries so both paths feed the
    identical downstream code.

    NOTE: unlike the scoreboard, the summary header's `status` carries only
    `type` - there is no `period` and no `displayClock`. Both are returned empty.
    Nothing downstream depends on them: every market derives the period and clock
    from the play-by-play itself (see `TrackedGame.max_period`).
    """
    payload = _http_get_json(f"{ESPN_BASE}/summary?event={game_id}")
    header = payload.get("header") or {}
    comps = header.get("competitions") or []
    if not comps:
        raise DataSourceError(f"ID {game_id} is not an NBA game on this feed.")
    comp = comps[0]
    stype = ((comp.get("status") or {}).get("type")) or {}

    away_raw = home_raw = None
    away_score = home_score = 0
    for c in comp.get("competitors") or []:
        try:
            score = int(float(c.get("score") or 0))
        except (TypeError, ValueError):
            score = 0
        if c.get("homeAway") == "home":
            home_raw, home_score = c.get("team") or {}, score
        elif c.get("homeAway") == "away":
            away_raw, away_score = c.get("team") or {}, score
    if not away_raw or not home_raw:
        raise DataSourceError(f"ID {game_id} did not return two NBA teams.")

    return {
        "game_id": str(header.get("id") or game_id),
        "start_iso": comp.get("date") or "",
        "state": stype.get("state", "pre"),
        "status_detail": stype.get("shortDetail") or stype.get("description") or "",
        "period": 0,
        "display_clock": "",
        "away": asdict(_team_from_raw(away_raw)),
        "home": asdict(_team_from_raw(home_raw)),
        "away_score": away_score,
        "home_score": home_score,
    }


def game_from_id(game_id: str) -> GameInfo:
    return _game_from_dict(fetch_game_header(game_id))


@st.cache_data(ttl=max(REFRESH_SECONDS - 1, 2), show_spinner=False)
def fetch_summary(game_id: str) -> dict:
    """Play-by-play + boxscore for one game.

    The cache TTL is just under the refresh interval so incidental Streamlit
    reruns (widget clicks, tab switches) reuse the last payload instead of
    hitting the feed again.
    """
    return _http_get_json(f"{ESPN_BASE}/summary?event={game_id}")


@st.cache_data(ttl=3600, show_spinner=False)
def fetch_roster(team_id: str) -> list[dict]:
    """Full team roster. Used for the key-player dropdowns before tip-off,
    when the boxscore is not populated yet."""
    payload = _http_get_json(f"{ESPN_BASE}/teams/{team_id}/roster")

    players: list[dict] = []
    for group in payload.get("athletes") or []:
        # ESPN returns either a flat athlete list or position-grouped buckets.
        entries = group.get("items") if isinstance(group, dict) and "items" in group else [group]
        for a in entries or []:
            if not isinstance(a, dict):
                continue
            pid = str(a.get("id", ""))
            if not pid:
                continue
            pos = a.get("position") or {}
            players.append(
                {
                    "player_id": pid,
                    "name": a.get("fullName") or a.get("displayName") or "",
                    "jersey": str(a.get("jersey") or ""),
                    "position": (pos.get("abbreviation") or "") if isinstance(pos, dict) else "",
                }
            )
    players.sort(key=lambda p: p["name"])
    return players


def roster_players(team_id: str) -> list[RosterPlayer]:
    return [RosterPlayer(**p) for p in fetch_roster(team_id)]


# ===========================================================================
# SECTION 4 - NORMALISATION
# Raw feed plays -> neutral GameEvent list. This is the boundary: nothing
# downstream of here touches provider-specific keys.
# ===========================================================================

# Field-goal kinds. Free throws are tracked separately because they score
# points (Timeframe Both Teams To Score) but are not field goals.
KIND_FG = "fg"
KIND_FT = "ft"
KIND_TURNOVER = "turnover"
# Substitutions get their own kind rather than falling into KIND_OTHER: the
# on-floor panel needs to find them, and naming them makes it explicit that they
# are outside TRACKED_KINDS and so can never become a correction row.
KIND_SUB = "sub"
KIND_OTHER = "other"

# Kinds that participate in stat-correction comparison. Rebounds, fouls and
# substitutions (KIND_SUB) are excluded so routine feed churn cannot produce
# noise. Substitutions are also kept out of the sequence / live-edge arithmetic
# in SECTION 6 - see the comment in `detect_corrections` for why that matters.
TRACKED_KINDS = (KIND_FG, KIND_FT, KIND_TURNOVER)

# Word-boundary match so player names like "Ryan Dunn" never read as a dunk.
_DUNK_RE = re.compile(r"\bdunk", re.IGNORECASE)

# Shot distance is the last resort for telling a 2 from a 3, needed only where
# `pointsAttempted` is absent and the text does not say "three point" - an older
# payload renders a missed three as just "misses 27-foot step back jumpshot".
# The NBA arc is 23.75ft (22ft in the corners), so 23 is the lowest threshold that
# does not start swallowing long twos. Measured over 27 games / 54 official team
# box-score lines: 23ft and 24ft each left 1 mismatch, 22ft made it worse at 5.
_SHOT_FEET_RE = re.compile(r"(\d+)-foot")
_THREE_POINT_FEET = 23
_ISO_CLOCK_RE = re.compile(r"^PT(?:(\d+)M)?(?:([\d.]+)S)?$", re.IGNORECASE)
_WS_RE = re.compile(r"\s+")

# ESPN's substitution play. `type.id` 584 is the structured signal and is the one
# to trust: it was present on 480/480 substitutions measured across 9 games
# (2025-10-28, Christmas, both 2026 All-Star games, 2026-04-08, Finals game
# 401859966). The two text checks are a backstop for an older payload that might
# number play types differently, the same defensive shape as the shot classifier.
ESPN_SUB_TYPE_ID = "584"
_SUB_TEXT_RE = re.compile(r"^(?P<incoming>.+?) enters the game for (?P<outgoing>.+?)$")


@dataclass(frozen=True)
class GameEvent:
    """One play, provider-neutral."""

    event_id: str
    sequence: int
    period: int
    clock_display: str
    clock_seconds: float | None   # seconds remaining in the period
    team_id: str
    player_id: str
    player_name: str
    kind: str
    made: bool
    points: int                  # 2 or 3 for a field goal, 1 for a free throw
    score_value: int             # points actually scored by this play
    is_dunk: bool
    away_score: int
    home_score: int
    type_text: str
    description: str
    short_desc: str
    # Substitutions only. The player coming ON is already in `player_id` /
    # `player_name`, because the feed puts them in participants[0] exactly as it
    # puts a shooter there. These two carry the player going OFF, who would
    # otherwise be discarded.
    sub_out_id: str = ""
    sub_out_name: str = ""

    @property
    def is_fg_attempt(self) -> bool:
        return self.kind == KIND_FG

    @property
    def is_scoring(self) -> bool:
        return self.score_value > 0

    @property
    def sub_in_id(self) -> str:
        """For a substitution, the player coming on. Named rather than leaving
        callers to know that `player_id` means the incoming player here."""
        return self.player_id if self.kind == KIND_SUB else ""

    @property
    def sub_in_name(self) -> str:
        return self.player_name if self.kind == KIND_SUB else ""


def norm_text(value: Any) -> str:
    """Collapse whitespace / newlines so cosmetic feed reformatting is not
    mistaken for a stat correction. ESPN emits e.g. 'Bad Pass\\nTurnover'."""
    return _WS_RE.sub(" ", str(value or "")).strip()


def parse_clock_seconds(display: Any) -> float | None:
    """Seconds remaining in the period.

    Handles every clock shape we have observed or might have to support:
      "11:39"        -> 699.0   (ESPN, above one minute)
      "51.7"         -> 51.7    (ESPN, under one minute, tenths)
      "PT11M39.00S"  -> 699.0   (ISO-8601, official NBA CDN shape)
    """
    raw = str(display or "").strip()
    if not raw:
        return None

    iso = _ISO_CLOCK_RE.match(raw)
    if iso:
        mins = float(iso.group(1) or 0)
        secs = float(iso.group(2) or 0)
        return mins * 60 + secs

    try:
        if ":" in raw:
            parts = raw.split(":")
            if len(parts) == 2:
                return int(parts[0]) * 60 + float(parts[1])
            if len(parts) == 3:
                return int(parts[0]) * 3600 + int(parts[1]) * 60 + float(parts[2])
            return None
        return float(raw)
    except (TypeError, ValueError):
        return None


def period_label(period: int) -> str:
    """1 -> '1Q' ... 4 -> '4Q', 5 -> 'OT', 6 -> '2OT'."""
    if period <= 0:
        return DASH
    if period <= 4:
        return f"{period}Q"
    n = period - 4
    return "OT" if n == 1 else f"{n}OT"


def period_seconds(period: int) -> int:
    return REGULATION_PERIOD_SECONDS if period <= 4 else OVERTIME_PERIOD_SECONDS


def _shooter_name_from_text(text: str) -> str:
    """Fallback shooter name when the ID -> name map has no entry.

    Two description shapes matter:
      "Paolo Banchero makes driving layup (...)"        -> name before makes/misses
      "Cedric Coward blocks Wendell Carter Jr. 's ..."  -> shooter is the blocked man
    """
    clean = norm_text(text)
    blocked = re.match(r"^.+?\s+blocks\s+(.+?)\s*'s\b", clean)
    if blocked:
        return blocked.group(1).strip()
    plain = re.match(r"^(.+?)\s+(?:makes|misses)\b", clean)
    if plain:
        return plain.group(1).strip()
    return ""


def _is_substitution(type_id: str, type_text: str, short_desc: str) -> bool:
    """Is this play a substitution?

    Checked before `_classify_play` rather than inside it, because the reliable
    signal is `type.id`, which the classifier is not given. A substitution is not
    a shooting play and carries no turnover text, so it would otherwise fall
    through to KIND_OTHER and be invisible to the on-floor panel.
    """
    return (
        type_id == ESPN_SUB_TYPE_ID
        or type_text.lower() == "substitution"
        or short_desc.lower() == "substitution"
    )


def _sub_names_from_text(description: str) -> tuple[str, str]:
    """(coming on, going off) read from 'X enters the game for Y'.

    Name fallback only - the IDs always come from `participants`. Needed because a
    player can change without ever appearing in the boxscore name map (he never
    touched the ball), and a bare ESPN athlete id in the panel is useless to a
    trader. The shape matched 480/480 substitutions across the 9 games measured.
    """
    match = _SUB_TEXT_RE.match(norm_text(description))
    if not match:
        return "", ""
    return match.group("incoming").strip(), match.group("outgoing").strip()


def _classify_play(short_desc: str, type_text: str, description: str, shooting: bool,
                   pts_attempted: int, scoring: bool) -> tuple[str, int]:
    """(kind, points) for one play.

    Primary signals, in order of reliability:
      pointsAttempted  2/3 -> field goal, 1 -> free throw
      shortDescription "+2 Points" / "Missed 3PT" / "Turnover"
      type.text        contains "Turnover" / "Free Throw"

    `shootingPlay` is the last authority, and it applies to makes and misses
    alike: a shooting play that is not a free throw is a field-goal attempt
    whether or not it scored. Older payloads need this - see the branch comment.
    """
    sd = short_desc.lower()
    tt = type_text.lower()
    desc = description.lower()
    # "three point" comes from the full play text and is the only 3PT signal that
    # survives in older payloads: a missed three there reads shortDescription
    # "Jump Shot" with pointsAttempted 0, and text "misses 26-foot three point
    # jumper". Without it those attempts are counted as 2PT and the 3PT lines go
    # badly wrong (measured 12-16 against an official 12-35).
    is_three = "3pt" in sd or "+3 points" in sd or "three point" in desc
    if not is_three:
        feet = _SHOT_FEET_RE.search(desc)
        is_three = bool(feet and int(feet.group(1)) >= _THREE_POINT_FEET)
    is_ft = "free throw" in tt or sd in ("+1 point", "missed ft")

    if shooting:
        if pts_attempted in (2, 3):
            return KIND_FG, pts_attempted
        if pts_attempted == 1 or is_ft:
            return KIND_FT, 1
        # pointsAttempted absent or zero. Pre-2015-ish payloads put the shot TYPE
        # in shortDescription ("Jump Shot", "Hook Shot"), so no text branch below
        # fires either. This branch used to be `if shooting and scoring`, which
        # rescued made shots but silently dropped missed ones: that lost 33 of the
        # 2012 finals game's 164 field-goal attempts, and every dropped play was a
        # miss. Never gate this on `scoring`.
        return KIND_FG, (3 if is_three else 2)

    if is_three:
        return KIND_FG, 3
    if sd in ("missed fg", "+2 points"):
        return KIND_FG, 2
    if is_ft:
        return KIND_FT, 1
    if "turnover" in sd or "turnover" in tt:
        return KIND_TURNOVER, 0
    return KIND_OTHER, 0


def normalise_plays(payload: dict, name_map: dict[str, str]) -> list[GameEvent]:
    """Raw `summary.plays` -> ordered GameEvent list (oldest first)."""
    events: list[GameEvent] = []

    for idx, raw in enumerate(payload.get("plays") or []):
        if not isinstance(raw, dict):
            continue

        try:
            sequence = int(str(raw.get("sequenceNumber") or idx))
        except (TypeError, ValueError):
            sequence = idx

        period_raw = raw.get("period") or {}
        period = int(period_raw.get("number") or 0) if isinstance(period_raw, dict) else 0

        clock_raw = raw.get("clock") or {}
        clock_display = (
            clock_raw.get("displayValue") if isinstance(clock_raw, dict) else str(clock_raw or "")
        ) or ""

        type_raw = raw.get("type") or {}
        type_id = str(type_raw.get("id") or "") if isinstance(type_raw, dict) else ""
        type_text = norm_text(type_raw.get("text") if isinstance(type_raw, dict) else "")
        short_desc = norm_text(raw.get("shortDescription"))
        description = norm_text(raw.get("text"))

        shooting = bool(raw.get("shootingPlay"))
        scoring = bool(raw.get("scoringPlay"))
        try:
            pts_attempted = int(raw.get("pointsAttempted") or 0)
        except (TypeError, ValueError):
            pts_attempted = 0
        try:
            score_value = int(raw.get("scoreValue") or 0)
        except (TypeError, ValueError):
            score_value = 0

        if _is_substitution(type_id, type_text, short_desc):
            kind, points = KIND_SUB, 0
        else:
            kind, points = _classify_play(
                short_desc, type_text, description, shooting, pts_attempted, scoring
            )

        # participants[0] is the shooter, including on blocked shots where the
        # description leads with the blocker's name - and, on a substitution, the
        # player coming ON.
        participant_ids: list[str] = []
        for part in raw.get("participants") or []:
            athlete = (part or {}).get("athlete") or {}
            pid = str(athlete.get("id", ""))
            if pid:
                participant_ids.append(pid)
        player_id = participant_ids[0] if participant_ids else ""

        player_name = name_map.get(player_id, "") or _shooter_name_from_text(description)

        # participants[1] on a substitution is the player going OFF. Verified by
        # mapping both ids back to the boxscore names on 480/480 substitutions
        # across 9 games: [0] was always the one the description names as coming
        # on, [1] always the one it names as going off, and every substitution
        # carried exactly two participants. Anything else leaves these empty, and
        # `team_floor` skips a substitution it cannot read rather than guessing -
        # a half-applied swap would corrupt the five for the rest of the game.
        sub_out_id = (
            participant_ids[1] if kind == KIND_SUB and len(participant_ids) >= 2 else ""
        )
        sub_out_name = ""
        if kind == KIND_SUB:
            text_in, text_out = _sub_names_from_text(description)
            # `_shooter_name_from_text` cannot help here: it looks for makes /
            # misses / blocks, none of which a substitution description contains.
            player_name = player_name or text_in
            if sub_out_id:
                sub_out_name = name_map.get(sub_out_id, "") or text_out

        team_raw = raw.get("team") or {}
        team_id = str(team_raw.get("id", "")) if isinstance(team_raw, dict) else ""

        try:
            away_score = int(raw.get("awayScore") or 0)
            home_score = int(raw.get("homeScore") or 0)
        except (TypeError, ValueError):
            away_score = home_score = 0

        # A dunk is identified from the structured shot type first; the free
        # text is only a backstop, matched on a word boundary.
        is_dunk = kind == KIND_FG and bool(
            _DUNK_RE.search(type_text) or _DUNK_RE.search(description)
        )

        events.append(
            GameEvent(
                event_id=str(raw.get("id") or f"seq-{sequence}"),
                sequence=sequence,
                period=period,
                clock_display=clock_display,
                clock_seconds=parse_clock_seconds(clock_display),
                team_id=team_id,
                player_id=player_id,
                player_name=player_name,
                kind=kind,
                made=scoring if kind in (KIND_FG, KIND_FT) else False,
                points=points,
                score_value=score_value,
                is_dunk=is_dunk,
                away_score=away_score,
                home_score=home_score,
                type_text=type_text,
                description=description,
                short_desc=short_desc,
                sub_out_id=sub_out_id,
                sub_out_name=sub_out_name,
            )
        )

    # Deliberately NOT sorted by `sequence`. ESPN's `sequenceNumber` is not monotonic
    # with game time: in game 401812485 the plays at 2:51 / 2:31 / 1:56 of Q4 carry
    # sequence 778 / 782 / 787 while the plays from 1:38 down to 0.3 carry 723-772,
    # so sorting by it drags mid-quarter plays past the end of the game. Measured over
    # 6 games / ~2900 plays: the payload's own array order had 0 chronology
    # violations in every game, while sequence order had 1-5 in five of six and moved
    # 19-70 plays. The feed is already chronological, so preserve it.
    #
    # This is what made the "most recent field goal attempts" feed show a missed 3 at
    # 1:56 Q4 as Miami's last shot instead of the made 3 at 39.7. It also silently
    # affected every order-dependent market - first basket, first 3, latest made FG.
    #
    # `sequence` is still kept on the event: the correction engine uses it to spot a
    # retroactive insertion, which does not require it to be ordered.
    return events


def extract_boxscore_players(
    payload: dict,
) -> tuple[dict[str, str], dict[str, list[RosterPlayer]], dict[str, str]]:
    """(id -> name map, team_id -> starters, id -> jersey) from the boxscore.

    `starter` is a real flag in this feed and yields exactly five players per
    team once the boxscore is published (typically well before tip-off). It is
    empty for a game that has not been posted yet - see `resolve_starters`.

    The jersey map covers every athlete in the boxscore, not just the starters,
    because `teams/{id}/roster` is the CURRENT roster and cannot number a player
    who has since left: for game 401859966 the boxscore carries a jersey for
    30/30 athletes, while San Antonio's roster endpoint knows only 11 of the 15
    who played for them (Olynyk, Waters, Biyombo and Plumlee are gone) and lists
    no jersey at all for 5 of its own 19 athletes. On-floor chips were rendering
    without a number because of it.
    """
    names: dict[str, str] = {}
    starters: dict[str, list[RosterPlayer]] = {}
    jerseys: dict[str, str] = {}

    box = payload.get("boxscore") or {}
    for team_block in box.get("players") or []:
        team_id = str(((team_block or {}).get("team") or {}).get("id", ""))
        found: list[RosterPlayer] = []
        for stat_block in team_block.get("statistics") or []:
            for entry in (stat_block or {}).get("athletes") or []:
                athlete = (entry or {}).get("athlete") or {}
                pid = str(athlete.get("id", ""))
                if not pid:
                    continue
                nm = athlete.get("displayName") or athlete.get("shortName") or ""
                if nm:
                    names[pid] = nm
                num = str(athlete.get("jersey") or "")
                if num:
                    jerseys[pid] = num
                pos = athlete.get("position") or {}
                if entry.get("starter"):
                    found.append(
                        RosterPlayer(
                            player_id=pid,
                            name=nm,
                            jersey=num,
                            position=(pos.get("abbreviation") or "") if isinstance(pos, dict) else "",
                        )
                    )
        if team_id and found:
            starters[team_id] = found[:5]
    return names, starters, jerseys


def resolve_starters(
    team_id: str,
    box_starters: dict[str, list[RosterPlayer]],
    events: Sequence[GameEvent],
    roster: Sequence[RosterPlayer],
) -> tuple[list[RosterPlayer], str]:
    """Best available starting five, plus a label saying where it came from.

    Fallback chain, most to least trustworthy:
      1. boxscore `starter` flag - the real lineup, 5 per team.
      2. first five distinct players from that team to appear in the play-by-play
         - effectively the starters once the game is a few minutes old.
      3. roster order - NOT a real lineup. Surfaced in the UI as such so nobody
         results a market off it.
    """
    if box_starters.get(team_id):
        return box_starters[team_id][:5], "boxscore starters"

    seen: list[str] = []
    by_id = {p.player_id: p for p in roster}
    for ev in events:
        if ev.team_id != team_id or not ev.player_id:
            continue
        if ev.kind not in (KIND_FG, KIND_FT, KIND_TURNOVER):
            continue
        if ev.player_id not in seen:
            seen.append(ev.player_id)
        if len(seen) >= 5:
            break
    if len(seen) >= 5:
        return (
            [
                by_id.get(pid, RosterPlayer(pid, _name_of(pid, events), "", ""))
                for pid in seen[:5]
            ],
            "first five to appear in play-by-play",
        )

    return list(roster[:5]), "roster order (not a confirmed lineup)"


def _name_of(player_id: str, events: Sequence[GameEvent]) -> str:
    for ev in events:
        if ev.player_id == player_id and ev.player_name:
            return ev.player_name
    return player_id


# ===========================================================================
# SECTION 5 - MARKET COMPUTATION
# Pure functions over GameEvent. Every market listed in the workbook resolves
# from these primitives.
# ===========================================================================

def _first(events: Iterable[GameEvent], pred: Callable[[GameEvent], bool]) -> GameEvent | None:
    for ev in events:
        if pred(ev):
            return ev
    return None


def first_event_row(events: Sequence[GameEvent], team_id: str | None = None) -> dict[str, str]:
    """First FG Exact / First FG Scorer / First 3 Make / First Dunk.

    Backs: (Prematch|Second Half) [Team] First Field Goal Scorer (Exact),
    First 3pt Scorer, First Dunk Scorer, Field Goal Exact/Team/Type.
    """
    def mine(ev: GameEvent) -> bool:
        return team_id is None or ev.team_id == team_id

    fg = _first(events, lambda e: e.kind == KIND_FG and e.made and mine(e))
    three = _first(events, lambda e: e.kind == KIND_FG and e.made and e.points == 3 and mine(e))
    dunk = _first(events, lambda e: e.kind == KIND_FG and e.made and e.is_dunk and mine(e))

    return {
        "First FG Exact": f"{fg.player_name} Made {fg.points}" if fg else DASH,
        "First FG Scorer": (fg.player_name or DASH) if fg else DASH,
        "First 3 Make": (three.player_name or DASH) if three else DASH,
        "First Dunk": (dunk.player_name or DASH) if dunk else DASH,
    }


def player_first_shot_row(events: Sequence[GameEvent], player_id: str) -> dict[str, str]:
    """First FG Attempt / First FG Type / First 3 Attempt for one player.

    Backs: Prematch Player First Field Goal Attempt / Type / First 3pt Attempt.
    """
    fga = _first(events, lambda e: e.kind == KIND_FG and e.player_id == player_id)
    tpa = _first(events, lambda e: e.kind == KIND_FG and e.points == 3 and e.player_id == player_id)
    return {
        "First FG Attempt": f"{'Made' if fga.made else 'Missed'} {fga.points}" if fga else DASH,
        "First FG Type": str(fga.points) if fga else DASH,
        "First 3 Attempt": ("Made" if tpa.made else "Missed") if tpa else DASH,
    }


def fg_attempts(events: Sequence[GameEvent], team_id: str | None = None) -> list[GameEvent]:
    return [
        e for e in events
        if e.kind == KIND_FG and (team_id is None or e.team_id == team_id)
    ]


# --- "Next Field Goal after <score>" market anchors ------------------------
#
# The market is named after a checkpoint score, and ONLY A MADE FIELD GOAL moves
# that checkpoint. Free throws and technical points move the scoreboard without
# opening a new market; a missed field goal moves nothing at all. So the score to
# show against an attempt is not the live board - it is the board as it stood
# after the previous made field goal.
#
# The checkpoint is read straight off the feed rather than added up from field-goal
# points, because points arrive between made field goals. Measured on game
# 401859966 (498 plays, 156 attempts, 72 made): 20 of the 72 new checkpoints would
# be wrong if derived from field-goal points alone, by 1 to 4 points - e.g. the
# game's first basket sits on a 0-0 checkpoint and produces 2-2, not 0-2, because
# Fox had made two free throws first. This costs nothing to get right, because a
# play's own awayScore / homeScore is already the board immediately after it:
# every one of the 109 scoring plays in that game had (this play's score minus the
# previous play's score) equal to its own scoreValue, and the last play's 106-107
# is the official final. There is no arithmetic to do.
#
# Recomputed from the event list on every poll like every other derived view here,
# which is what makes it survive a stat correction: if a correction adds, removes
# or re-values a made field goal, the checkpoints from that point on re-derive on
# the next refresh instead of having to be patched.


@dataclass(frozen=True)
class FGMarketEvent:
    """A field-goal attempt and the market it belongs to.

    Named in market vocabulary, not feed vocabulary, because these values get read
    straight against a market name in the trader's other system.
    """

    event: GameEvent
    market_anchor_score: tuple[int, int]              # (away, home) the market is named after
    event_result: str                                 # "NYK Made 3"
    post_event_score: tuple[int, int]                 # actual board right after this event
    new_market_anchor_score: tuple[int, int] | None   # made field goals only
    score_suspect: bool = False                       # feed board moved backwards

    @property
    def made(self) -> bool:
        return self.event.made

    @property
    def result_no_team(self) -> str:
        """`event_result` without the team, for a panel that is already one team's:
        'SA Made 2' repeated down the San Antonio column says nothing."""
        return _result_text(self.event)


def _result_text(ev: GameEvent) -> str:
    """'Made 2' / 'Missed 3'. One definition, used with and without a team prefix."""
    return f"{'Made' if ev.made else 'Missed'} {ev.points}"


def fg_market_events(events: Sequence[GameEvent],
                     abbr_for: dict[str, str] | None = None) -> list[FGMarketEvent]:
    """Every field-goal attempt in feed order, each carrying its market anchor.

    Walks only the field goals: a play's own score already includes every free
    throw before it, so nothing else has to be visited to keep the board right.
    The opening market is 0-0.
    """
    abbrs = abbr_for or {}
    anchor = (0, 0)
    rows: list[FGMarketEvent] = []

    for ev in fg_attempts(events):
        post = (ev.away_score, ev.home_score)
        # A board that has gone backwards cannot be right - most likely an absent
        # awayScore / homeScore, which normalises to 0. Report the attempt against
        # the last checkpoint we trust rather than moving the market backwards.
        suspect = post[0] < anchor[0] or post[1] < anchor[1]
        new_anchor = post if (ev.made and not suspect) else None

        abbr = abbrs.get(ev.team_id, "")
        rows.append(
            FGMarketEvent(
                event=ev,
                market_anchor_score=anchor,
                event_result=f"{abbr} {_result_text(ev)}".strip(),
                post_event_score=post,
                new_market_anchor_score=new_anchor,
                score_suspect=suspect,
            )
        )
        if new_anchor is not None:
            anchor = new_anchor

    return rows


def recent_fg_market_events(events: Sequence[GameEvent], abbr_for: dict[str, str],
                            team_id: str | None, count: int,
                            made_only: bool = False) -> list[FGMarketEvent]:
    """Newest-first slice of `fg_market_events`, filtered for one panel.

    The anchors are derived over the whole game first and only then filtered: a
    team's checkpoint is moved by the opponent's baskets too, so deriving from one
    team's attempts in isolation would name the wrong market.
    """
    rows = fg_market_events(events, abbr_for)
    if team_id is not None:
        rows = [r for r in rows if r.event.team_id == team_id]
    if made_only:
        rows = [r for r in rows if r.made]
    return list(reversed(rows[-count:]))


def open_market_anchor(events: Sequence[GameEvent]) -> tuple[int, int]:
    """The checkpoint the currently open market is named after: the board after the
    last made field goal, or 0-0 before the first one.

    Taken from `fg_market_events` rather than by scanning backwards for a made
    field goal, so it cannot disagree with the panels about which checkpoint is
    open - including when one of them was rejected as suspect.
    """
    for row in reversed(fg_market_events(events)):
        if row.new_market_anchor_score is not None:
            return row.new_market_anchor_score
    return (0, 0)


def latest_made_fg(events: Sequence[GameEvent], player_id: str) -> GameEvent | None:
    for ev in reversed(events):
        if ev.kind == KIND_FG and ev.made and ev.player_id == player_id:
            return ev
    return None


# --- Timeframe Both Teams To Score ----------------------------------------

def timeframe_windows(period: int) -> list[tuple[str, int, int]]:
    """One-minute windows for a period as (label, high_seconds, low_seconds).

    A window covers game clock in (low, high]; the final window is inclusive of
    0.0 so a buzzer-beater lands in "1:00-0:00".
    """
    total = period_seconds(period)
    out: list[tuple[str, int, int]] = []
    for i in range(total // 60):
        high = total - i * 60
        low = high - 60
        low_label = "0:00" if low == 0 else _clock_label(low + 1)
        out.append((f"{_clock_label(high)}-{low_label}", high, low))
    return out


def _clock_label(seconds: int) -> str:
    return f"{seconds // 60}:{seconds % 60:02d}"


def window_index(clock_seconds: float, period: int) -> int | None:
    """Which one-minute window a clock reading belongs to.

    12:00 (720s) -> window 0 ("12:00-11:01")
    11:01 (661s) -> window 0
    11:00 (660s) -> window 1 ("11:00-10:01")   labels are upper-bound inclusive
    51.7s        -> window 11 ("1:00-0:00")
    """
    total = period_seconds(period)
    if clock_seconds is None or clock_seconds > total or clock_seconds < 0:
        return None
    return min(total // 60 - 1, int((total - clock_seconds) // 60))


def period_progress(events: Sequence[GameEvent], period: int) -> tuple[float | None, bool]:
    """(lowest clock reached in this period, period is finished).

    "Finished" is taken from an explicit end-of-period play, or from the feed
    having moved on to a later period.
    """
    lowest: float | None = None
    finished = False
    max_period = 0

    for ev in events:
        max_period = max(max_period, ev.period)
        if ev.period != period:
            continue
        if ev.clock_seconds is not None:
            lowest = ev.clock_seconds if lowest is None else min(lowest, ev.clock_seconds)
        sd = ev.short_desc.lower()
        tt = ev.type_text.lower()
        if sd.startswith("end of") or sd == "halftime" or tt in ("end period", "end game"):
            finished = True

    if max_period > period:
        finished = True
    return lowest, finished


def timeframe_table(events: Sequence[GameEvent], period: int, away_id: str, home_id: str,
                    game_is_final: bool) -> list[dict[str, str]]:
    """Yes / No / - per one-minute window for Timeframe Both Teams To Score.

    Yes  - both teams scored at least once inside the exact window.
    No   - the window is fully complete and both teams did not score.
    -    - the window has not happened yet, or is still in progress.

    Any scoring play counts (made field goal or made free throw), which is what
    the market prices.
    """
    windows = timeframe_windows(period)
    scored: list[set[str]] = [set() for _ in windows]

    for ev in events:
        if ev.period != period or not ev.is_scoring or not ev.team_id:
            continue
        idx = window_index(ev.clock_seconds, period)
        if idx is None:
            continue
        scored[idx].add(ev.team_id)

    lowest, finished = period_progress(events, period)
    finished = finished or game_is_final

    rows: list[dict[str, str]] = []
    for idx, (label, _high, low) in enumerate(windows):
        both = away_id in scored[idx] and home_id in scored[idx]
        # The window is complete once the clock has run down to its lower bound.
        complete = finished or (lowest is not None and lowest <= low)
        if both:
            verdict = "Yes"
        elif complete:
            verdict = "No"
        else:
            verdict = DASH
        rows.append({"Quarter": period_label(period), "Time": label, "Yes/No": verdict})

    if TIMEFRAME_ORDER == "reverse":
        rows.reverse()
    return rows


# --- Players on the floor --------------------------------------------------
#
# Derived, never reported. ESPN publishes no on-court field on any endpoint
# reachable from here - the `summary` payload, the core-API competition object,
# `/situation` and the per-competitor `/roster` were all checked, and the only
# lineup information in any of them is the boxscore `starter` flag. So the five
# are the starters with every substitution applied in feed order.
#
# Measured over 9 games / 3,746 plays / 480 substitutions before this was built:
#   - both teams held exactly five players at all 7,492 team-checkpoints;
#   - no substitution ever took off a player the reconstruction did not have on,
#     and none ever brought on a player it already had on;
#   - 2,335 of 2,338 single-actor plays were by a player it had on the floor.
# The 3 exceptions are same-clock ordering ties, where the feed lists the
# substitution just ahead of one last play by the man going off (Q4 30.1 of
# ATL@CLE: "Dean Wade enters the game for Donovan Mitchell", then Mitchell's
# turnover). They correct themselves as soon as play moves on, which is why the
# panel shows the entry clock rather than trying to reorder the feed.

@dataclass(frozen=True)
class FloorEntry:
    """A player, plus the game time of the substitution that put them here.

    `period` 0 means "on since the start of the game" - a starter who has not
    been substituted, for whom there is no substitution clock to show.
    """

    player_id: str
    name: str
    period: int
    clock_display: str


@dataclass(frozen=True)
class TeamFloor:
    on_floor: tuple[FloorEntry, ...]
    unverified: int                      # subs applied to a player we had on the bench


def team_floor(events: Sequence[GameEvent], team_id: str,
               starters: Sequence[RosterPlayer]) -> TeamFloor:
    """Who is on the floor for one team right now.

    Pure function of the event list, recomputed every poll like every other table
    in this tool - so it cannot drift out of step with the feed the way stored
    lineup state would, and a corrected or removed substitution simply stops
    counting on the next refresh.
    """
    if not starters:
        return TeamFloor((), 0)

    on: dict[str, FloorEntry] = {
        p.player_id: FloorEntry(p.player_id, p.name or p.player_id, 0, "")
        for p in starters[:5]
    }
    unverified = 0

    for ev in events:
        if ev.kind != KIND_SUB or ev.team_id != team_id:
            continue
        if not ev.sub_in_id or not ev.sub_out_id:
            continue  # unreadable substitution: leave the five alone

        if ev.sub_out_id in on:
            del on[ev.sub_out_id]
        else:
            # Never seen in the measured games. Counted and surfaced rather than
            # silently swallowed, because it means the five on screen are wrong.
            unverified += 1
        on[ev.sub_in_id] = FloorEntry(
            ev.sub_in_id, ev.sub_in_name or ev.sub_in_id, ev.period, ev.clock_display
        )

    return TeamFloor(tuple(on.values()), unverified)


def recent_team_subs(events: Sequence[GameEvent], team_id: str,
                     limit: int = FLOOR_RECENT_SUBS) -> list[GameEvent]:
    """One team's last `limit` substitutions, most recent first.

    Substitution events rather than the players who went off, so each line can be
    read as the pair it actually was ("X for Y"): the trader is copying that pair
    into another system, and a bare list of names who left does not say who took
    their place. Repeats are kept - a player subbed off twice in a quarter really
    did leave twice, and collapsing that would misdate the more recent one.
    """
    subs = [
        e for e in events
        if e.kind == KIND_SUB and e.team_id == team_id and e.sub_in_id and e.sub_out_id
    ]
    return list(reversed(subs[-limit:]))


def latest_substitutions(events: Sequence[GameEvent],
                         limit: int = SUB_HIGHLIGHT_MAX) -> list[GameEvent]:
    """The most recent substitution break, in feed order.

    A break rather than a single play: at a timeout the feed publishes five or six
    substitutions at the same clock reading, and showing only the last of them
    would hide the rest of the change. Every substitution sharing the final one's
    period and clock is returned, both teams together.
    """
    subs = [e for e in events if e.kind == KIND_SUB and e.sub_in_id and e.sub_out_id]
    if not subs:
        return []
    last = subs[-1]
    same = [
        e for e in subs
        if e.period == last.period and e.clock_display == last.clock_display
    ]
    return same[-limit:]


# ===========================================================================
# SECTION 6 - STAT CORRECTION ENGINE
#
# Each poll we rebuild the tracked-event map and diff it against the previous
# snapshot. Three change shapes matter:
#
#   INSERTED  an event id we have not seen whose sequence number sits at or
#             below the highest sequence we had already seen. A genuinely new
#             live play always arrives with a HIGHER sequence, so this test
#             separates "the game moved on" from "the feed rewrote history".
#   REMOVED   an event id that has disappeared. Suppressed when the payload
#             looks truncated, so a partial response cannot fake a correction.
#   CHANGED   an event id whose market-relevant fields differ.
#
# The log is append-only for the session. If the same event is corrected twice,
# the before/after pair differs the second time and both entries are kept.
# ===========================================================================

# Fields whose change makes a correction market-impacting.
MARKET_FIELDS = ("kind", "made", "points", "player_id", "team_id", "is_dunk")


def fingerprint(ev: GameEvent) -> dict[str, Any]:
    """Comparison snapshot for one event.

    Deliberately excludes coordinates, shot distance and assist lists: those
    churn without ever moving a market. Text fields are whitespace-normalised
    so reformatting alone never registers.
    """
    return {
        "kind": ev.kind,
        "made": ev.made,
        "points": ev.points,
        "player_id": ev.player_id,
        "player_name": ev.player_name,
        "team_id": ev.team_id,
        "is_dunk": ev.is_dunk,
        "period": ev.period,
        "clock_display": ev.clock_display,
        "type_text": ev.type_text,
        "description": ev.description,
        "away_score": ev.away_score,
        "home_score": ev.home_score,
        "sequence": ev.sequence,
        "label": event_label(ev),
    }


def event_label(ev: GameEvent) -> str:
    """Human-readable event, matching the sheet's 'Devin Booker Missed 2' style."""
    name = ev.player_name or "Unknown"
    if ev.kind == KIND_FG:
        base = f"{name} {'Made' if ev.made else 'Missed'} {ev.points}"
        return f"{base} (Dunk)" if ev.is_dunk else base
    if ev.kind == KIND_FT:
        return f"{name} {'Made' if ev.made else 'Missed'} FT"
    if ev.kind == KIND_TURNOVER:
        return f"{name} Turnover"
    return ev.type_text or ev.short_desc or "Event"


def _label_from_fp(fp: dict[str, Any]) -> str:
    return fp.get("label") or "Event"


def classify_change(before: dict[str, Any], after: dict[str, Any]) -> tuple[list[str], bool]:
    """(impact descriptions, is_market_impacting) for a modified event.

    Maps onto the eight categories from the spec:
      3 make<->miss  4 2pt<->3pt  5 shooter  6 team  7 shot<->turnover  8 dunk
    (1 and 2 are additions/removals, handled by classify_presence.)
    """
    impacts: list[str] = []

    if before["kind"] != after["kind"]:
        pair = {before["kind"], after["kind"]}
        if pair == {KIND_FG, KIND_TURNOVER}:
            impacts.append(
                "Shot changed to turnover" if after["kind"] == KIND_TURNOVER
                else "Turnover changed to shot"
            )
        elif FLAG_SCORING_FT_CHANGES and pair == {KIND_FG, KIND_FT}:
            impacts.append(
                "Field goal changed to free throw" if after["kind"] == KIND_FT
                else "Free throw changed to field goal"
            )
        else:
            impacts.append(f"Event type changed ({before['kind']} to {after['kind']})")

    if before["kind"] == KIND_FG and after["kind"] == KIND_FG:
        if before["made"] != after["made"]:
            impacts.append("Make changed to miss" if before["made"] else "Miss changed to make")
        if before["points"] != after["points"]:
            impacts.append(f"{before['points']}-pointer changed to {after['points']}-pointer")
        if before["is_dunk"] != after["is_dunk"]:
            impacts.append(
                "Dunk classification added" if after["is_dunk"] else "Dunk classification removed"
            )

    if before["kind"] == KIND_FT and after["kind"] == KIND_FT and before["made"] != after["made"]:
        if FLAG_SCORING_FT_CHANGES:
            impacts.append(
                "Free throw make changed to miss" if before["made"]
                else "Free throw miss changed to make"
            )

    if before["player_id"] != after["player_id"]:
        impacts.append("Shooter attribution changed")
    if before["team_id"] != after["team_id"]:
        impacts.append("Team attribution changed")

    if impacts:
        return impacts, True

    # NOTE: player_name is deliberately not compared. It is looked up from the
    # id -> name map, which fills in as the boxscore publishes, so its text can
    # change with no correction behind it. player_id above is the authority.

    # Nothing market-relevant moved. Log it anyway so "All Corrections" is a
    # complete audit trail, but do not raise a banner.
    minor: list[str] = []
    if before["period"] != after["period"]:
        minor.append("Period adjusted")
    if before["clock_display"] != after["clock_display"]:
        minor.append("Clock adjusted")
    if (before["away_score"], before["home_score"]) != (after["away_score"], after["home_score"]):
        minor.append("Score adjusted")
    if before["type_text"] != after["type_text"]:
        minor.append("Shot type detail changed")
    elif before["description"] != after["description"]:
        minor.append("Description changed")
    return minor, False


def classify_presence(fp: dict[str, Any], added: bool) -> tuple[list[str], bool]:
    """Categories 1, 2 and the add/remove side of 7 for an inserted or deleted event."""
    verb = "added" if added else "removed"
    kind = fp["kind"]

    if kind == KIND_FG:
        impacts = [f"{'Made' if fp['made'] else 'Missed'} shot {verb}"]
        if fp["is_dunk"]:
            impacts.append(f"Dunk classification {verb}")
        return impacts, True
    if kind == KIND_TURNOVER:
        return [f"Turnover {verb}"], True
    if kind == KIND_FT:
        # Not one of the eight canonical categories, but a made free throw
        # appearing or vanishing does move Timeframe Both Teams To Score.
        if fp["made"] and FLAG_SCORING_FT_CHANGES:
            return [f"Made free throw {verb} (scoring event)"], True
        return [f"Free throw {verb}"], False
    return [f"Event {verb}"], False


def _live_edge(events: Sequence[GameEvent]) -> tuple[int, float]:
    """(furthest period reached, game clock of the newest play in it).

    Used to tell a backdated insertion apart from the game simply advancing.
    """
    period = max((e.period for e in events if e.period > 0), default=0)
    clock = float(period_seconds(period or 1))
    for ev in reversed(events):
        if ev.period == period and ev.clock_seconds is not None:
            clock = ev.clock_seconds
            break
    return period, clock


def _is_retroactive(fp: dict[str, Any], live_period: int, live_clock: float) -> bool:
    period = int(fp.get("period") or 0)
    if period <= 0 or live_period <= 0:
        return False
    if period < live_period:
        return True
    if period > live_period:
        return False
    clock = parse_clock_seconds(fp.get("clock_display"))
    if clock is None:
        return False
    # Clock counts DOWN, so a larger value means earlier in the period.
    return clock > live_clock + RETROACTIVE_TOLERANCE_SECONDS


def detect_corrections(events: Sequence[GameEvent], prev_snapshot: dict[str, dict],
                       prev_max_seq: int, prev_total: int) -> tuple[list[dict], dict[str, dict], int, int, str | None]:
    """Diff the current play-by-play against the last snapshot.

    Returns (new correction rows, new snapshot, new max sequence, new total,
    warning message or None).
    """
    tracked = [e for e in events if e.kind in TRACKED_KINDS]
    current = {e.event_id: fingerprint(e) for e in tracked}

    # `max_seq` and the live edge below come from `tracked`, NOT from every event,
    # and substitutions are the reason. Measured over 9 games / 3,746 plays:
    #
    #   - the first play of every new quarter is a substitution stamped with the
    #     NEW period at 12:00. An all-events live edge therefore jumps a quarter
    #     ahead before any shot in that quarter arrives, and a real buzzer-beater
    #     from the quarter just ended - first seen in that same poll - reads as a
    #     backdated insertion. CLE@NY (Merrill, Q2 13.4s) and ATL@CLE (Mitchell,
    #     Q1 0.5s) both did exactly that.
    #   - substitutions also carry sequence numbers above the newest tracked play
    #     (329 vs 233 in MIN@ORL), which trips the `by_sequence` test below on the
    #     next few genuine plays.
    #
    # Replaying all 9 games as polls: 279 false insertions before, 199 after
    # (MIN@ORL 31 -> 0, CHA@MIA 54 -> 12). `total` deliberately stays on every
    # event: it only feeds the truncation ratio below, both sides of which shrink
    # together, and a persisted sidecar from an earlier session holds an
    # all-events count - switching it would make the first poll after a restore
    # look like a 50% truncation and stall the snapshot.
    total = len(events)
    max_seq = max((e.sequence for e in tracked), default=0)
    detected_at = datetime.now().astimezone()
    rows: list[dict] = []
    warning: str | None = None

    # First poll of a game: take a baseline, do not report the whole history.
    if not prev_snapshot:
        return [], current, max_seq, total, None

    def row(fp_for_time: dict, original: str, updated: str, impacts: list[str],
            market: bool, event_id: str) -> dict:
        return {
            "detected_at": detected_at.isoformat(timespec="seconds"),
            "detected_display": detected_at.strftime("%H:%M:%S"),
            "period": int(fp_for_time.get("period") or 0),
            "clock": fp_for_time.get("clock_display") or DASH,
            "original": original,
            "updated": updated,
            "market_impacting": market,
            "impact": "; ".join(impacts) if impacts else "No market-relevant change",
            "event_id": event_id,
            # Dedupe key: the exact transition. A second, different correction
            # to the same event produces a different key and is logged too.
            "key": f"{event_id}|{original}|{updated}|{'M' if market else 'm'}",
        }

    # --- INSERTED ----------------------------------------------------------
    # Two independent tests, because we cannot assume how the feed numbers a
    # retroactively added play:
    #   a) its sequence number lands at or below the highest we had already
    #      seen (the feed rewrote history in place), or
    #   b) it carries a brand new high sequence number but its game clock sits
    #      well behind the live edge (the feed appended a backdated play).
    # A genuinely new live play fails both tests.
    live_period, live_clock = _live_edge(tracked)
    for eid, fp in current.items():
        if eid in prev_snapshot:
            continue
        by_sequence = fp["sequence"] <= prev_max_seq
        if not (by_sequence or _is_retroactive(fp, live_period, live_clock)):
            continue  # normal forward progress, not a correction
        impacts, market = classify_presence(fp, added=True)
        rows.append(row(fp, DASH, _label_from_fp(fp), impacts, market, eid))

    # --- REMOVED: id gone. Guarded against truncated payloads. -------------
    missing = [eid for eid in prev_snapshot if eid not in current]
    if missing:
        if prev_total and total < prev_total * REMOVAL_SANITY_RATIO:
            warning = (
                f"Feed returned {total} plays vs {prev_total} previously - "
                "removal detection skipped this cycle (suspected partial payload)."
            )
        else:
            for eid in missing:
                fp = prev_snapshot[eid]
                impacts, market = classify_presence(fp, added=False)
                rows.append(row(fp, _label_from_fp(fp), DASH, impacts, market, eid))

    # --- CHANGED ----------------------------------------------------------
    for eid, fp in current.items():
        before = prev_snapshot.get(eid)
        if before is None:
            continue
        if all(before.get(f) == fp.get(f) for f in MARKET_FIELDS) and \
           before.get("clock_display") == fp.get("clock_display") and \
           before.get("period") == fp.get("period") and \
           before.get("type_text") == fp.get("type_text") and \
           before.get("description") == fp.get("description") and \
           before.get("away_score") == fp.get("away_score") and \
           before.get("home_score") == fp.get("home_score"):
            continue
        impacts, market = classify_change(before, fp)
        if not impacts:
            continue  # nothing worth logging
        rows.append(row(fp, _label_from_fp(before), _label_from_fp(fp), impacts, market, eid))

    rows.sort(key=lambda r: (r["period"], not r["market_impacting"]))

    # The returned snapshot already reflects removals, and every removal above
    # has been turned into a log row, so nothing is dropped silently. When
    # `warning` is set the caller keeps the OLD snapshot instead, so a truncated
    # payload never becomes the new baseline.
    return rows, dict(current), max(max_seq, prev_max_seq), total, warning


# ===========================================================================
# SECTION 7 - SESSION STATE + BEST-EFFORT PERSISTENCE
# Streamlit reruns must never drop tracked state. session_state covers reruns;
# the JSON sidecar covers a hard browser refresh (a brand-new session).
# ===========================================================================

DEFAULT_STATE: dict[str, Any] = {
    "games": [],
    "games_error": None,
    "games_loaded": False,
    "manual_ids": set(),      # game ids loaded by hand, not from the scoreboard
    "scoreboard_day": None,
    "selected_game_id": None,
    "selected_game_label": None,
    "tracking": False,
    "active_tab": TAB_PREMATCH,
    "show_key_editor": False,  # UI only, like active_tab: never cleared, never saved
    "rate_limit_skip_remaining": 0,
    "key_players": {"away": [], "home": []},
    "kp_active": set(),
    "kp_seen": {},
    "kp_alerts": [],
    "corrections": [],
    "correction_keys": set(),
    "snapshot": {},
    "snapshot_max_seq": 0,
    "snapshot_total": 0,
    "last_fetch_ok": None,
    "last_error": None,
    "error_streak": 0,
    "feed_warning": None,
    "banner_dismissed_key": None,
    "timeframe_period": None,
    "last_save_ts": 0.0,
}

# Written to / restored from the JSON sidecar.
PERSISTED_KEYS = (
    "corrections", "correction_keys", "snapshot", "snapshot_max_seq",
    "snapshot_total", "key_players", "kp_seen",
)

# Wiped by "Reset tracked state". Deliberately excludes key_players: resetting
# the log should not silently desync the key-player dropdowns.
CLEARED_KEYS = (
    "corrections", "correction_keys", "snapshot", "snapshot_max_seq",
    "snapshot_total", "kp_seen",
)

# The snapshot is large, so routine saves are throttled. A newly detected
# correction always forces an immediate write.
SAVE_THROTTLE_SECONDS = 10


def init_state() -> None:
    # deepcopy, not .copy(): key_players is a dict OF LISTS, and a shallow copy
    # would hand every session the same inner lists as the module-level default.
    for key, value in DEFAULT_STATE.items():
        if key not in st.session_state:
            st.session_state[key] = deepcopy(value)


def _state_path(game_id: str) -> Path:
    return STATE_DIR / f"{game_id}.json"


def save_state(game_id: str, force: bool = False) -> None:
    """Persist the append-only log and diff snapshot. Never raises."""
    if not game_id:
        return
    now = time.time()
    if not force and now - float(st.session_state.get("last_save_ts") or 0) < SAVE_THROTTLE_SECONDS:
        return
    st.session_state.last_save_ts = now
    try:
        STATE_DIR.mkdir(parents=True, exist_ok=True)
        blob = {}
        for key in PERSISTED_KEYS:
            value = st.session_state.get(key)
            blob[key] = sorted(value) if isinstance(value, set) else value
        _state_path(game_id).write_text(json.dumps(blob), encoding="utf-8")
    except Exception:
        # A read-only filesystem is fine: session_state still covers reruns.
        pass


def load_state(game_id: str) -> bool:
    """Restore a previous tracking session for this game. Never raises."""
    if not game_id:
        return False
    try:
        path = _state_path(game_id)
        if not path.exists():
            return False
        blob = json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return False

    for key in PERSISTED_KEYS:
        if key not in blob:
            continue
        value = blob[key]
        st.session_state[key] = set(value) if key == "correction_keys" else value
    return True


def reset_game_state() -> None:
    """Drop everything derived from one specific game, keeping the sidecar file.

    Mirrors `reset_game_state` in the NFL tool, and is called on every change of
    selected game. Without it the previous game's snapshot survives into the
    next one, and the first poll diffs game B against game A: none of A's event
    ids exist in B, so every play in B reads as a retroactive insertion and the
    correction log fills with fiction.
    """
    for key in CLEARED_KEYS:
        st.session_state[key] = deepcopy(DEFAULT_STATE[key])
    st.session_state.kp_alerts = []
    st.session_state.kp_active = set()
    st.session_state.banner_dismissed_key = None


def clear_state(game_id: str) -> None:
    """Explicit "Reset tracked state": also deletes the on-disk sidecar, so the
    log does not come back on the next Track Game."""
    reset_game_state()
    try:
        _state_path(game_id).unlink(missing_ok=True)
    except Exception:
        pass


def record_corrections(rows: Sequence[dict], game_id: str) -> list[dict]:
    """Append genuinely new corrections. Returns those actually added."""
    added: list[dict] = []
    keys: set[str] = st.session_state.correction_keys
    for r in rows:
        if r["key"] in keys:
            continue
        keys.add(r["key"])
        st.session_state.corrections.append(r)
        added.append(r)
    if added:
        save_state(game_id, force=True)
    return added


# ===========================================================================
# SECTION 8 - UI HELPERS
# ===========================================================================

def sect(text: str) -> None:
    st.markdown(f'<div class="sect">{html.escape(text)}</div>', unsafe_allow_html=True)


def subsect(text: str) -> None:
    st.markdown(f'<div class="subsect">{html.escape(text)}</div>', unsafe_allow_html=True)


def note(text: str) -> None:
    st.markdown(f'<div class="note">{html.escape(text)}</div>', unsafe_allow_html=True)


_TH_STYLE = (
    "padding:6px 14px; text-align:left; border-bottom:2px solid "
    "var(--secondary-background-color); font-size:12px; color:var(--text-color); "
    "font-weight:700; white-space:nowrap; text-transform:uppercase; "
    "letter-spacing:0.04em;"
)


def _pill(text: str, bg: str, fg: str) -> str:
    return (
        f'<span style="background-color:{bg}; color:{fg}; padding:2px 10px; '
        f'border-radius:12px; font-weight:700; font-size:12px; '
        f'white-space:nowrap;">{html.escape(text)}</span>'
    )


def _render_cell(value: str) -> str:
    """Yes / No as the house pills, so a resulted market reads the same here as
    it does in the NFL and NHL tools. Everything else is plain escaped text."""
    if value == "Yes":
        return _pill("Yes", "#00cc44", "#000000")
    if value == "No":
        return _pill("No", "#cc2200", "#ffffff")
    if value == DASH:
        return f'<span style="opacity:0.45;">{html.escape(value)}</span>'
    return html.escape(value)


def html_table(rows: Sequence[dict], wrap_columns: set[str] | None = None) -> None:
    """House HTML table, ported from `nfl_qc_tool/components/tables.py`.

    Takes a list of dicts and derives the headers from the first row - the same
    signature as the NFL / NHL renderer, so a table defined here is laid out and
    coloured identically to one defined there. Colours come from the Streamlit
    theme rather than a private palette, which is what makes it match. Static
    HTML means no internal scrollbars, ever.

    wrap_columns: columns allowed to wrap onto several lines. Cells default to
    nowrap so short values never break mid-value, but a long impact description
    would otherwise force horizontal scrolling.
    """
    if not rows:
        st.info("No data.")
        return
    wrap_columns = wrap_columns or set()
    headers = list(rows[0].keys())
    head = "".join(f'<th style="{_TH_STYLE}">{html.escape(str(h))}</th>' for h in headers)
    body = []
    for i, row in enumerate(rows):
        bg = "rgba(128,128,128,0.04)" if i % 2 == 0 else "rgba(128,128,128,0.10)"
        cells = []
        for h in headers:
            if h in wrap_columns:
                sizing = ("white-space:normal; overflow-wrap:break-word; "
                          "width:100%; min-width:200px; line-height:1.5;")
            else:
                sizing = "white-space:nowrap;"
            cells.append(
                f'<td style="padding:5px 14px; font-size:13px; {sizing} '
                f'vertical-align:top; color:var(--text-color); font-weight:500;">'
                f'{_render_cell(str(row.get(h, "")))}</td>'
            )
        body.append(f'<tr style="background-color:{bg};">{"".join(cells)}</tr>')
    st.markdown(
        '<div style="overflow-x:auto; width:100%;">'
        '<table style="width:100%; border-collapse:collapse;">'
        f'<thead><tr>{head}</tr></thead><tbody>{"".join(body)}</tbody></table></div>',
        unsafe_allow_html=True,
    )


def format_event_line(ev: GameEvent, away_abbr: str, home_abbr: str,
                      team_abbr: str, include_team: bool = True) -> str:
    """'PHX Made 3 - 2:34 1Q - PHX 42, NYK 38'."""
    result = f"{'Made' if ev.made else 'Missed'} {ev.points}"
    prefix = f"{team_abbr} " if include_team and team_abbr else ""
    clock = ev.clock_display or DASH
    score = f"{away_abbr} {ev.away_score}, {home_abbr} {ev.home_score}"
    return f"{prefix}{result} — {clock} {period_label(ev.period)} — {score}"


def _anchor_text(score: tuple[int, int]) -> str:
    """'2-0'. Away first, home second, matching the market name."""
    return f"{score[0]}-{score[1]}"


def render_feed(rows: Sequence[FGMarketEvent], empty_text: str,
                include_team: bool = False) -> None:
    """Field-goal attempts against the market each was priced under.

    Three columns under one header - After, Result, Time - rather than repeating
    "After" and the team on every line. `include_team` is for the combined panel
    only; inside a team's own panel the abbreviation is the same on every row.

    The score a make opens the next market on is held in `new_market_anchor_score`
    and deliberately NOT rendered: it is the same number as the next row's anchor.
    """
    if not rows:
        st.markdown(
            f'<div class="frow empty">{html.escape(empty_text)}</div>', unsafe_allow_html=True
        )
        return

    parts = [
        '<div class="fhead"><span>After</span><span>Result</span>'
        '<span class="sc">Time</span></div>'
    ]
    for row in rows:
        ev = row.event
        result = row.event_result if include_team else row.result_no_team
        # Inside the Result cell, not as a fourth column, so the grid stays aligned.
        flag = ' <span class="mk bad">score?</span>' if row.score_suspect else ""
        parts.append(
            f'<div class="frow {"made" if row.made else "miss"}">'
            f'<span class="anch">{_anchor_text(row.market_anchor_score)}</span>'
            f'<span class="res">{html.escape(result)}{flag}</span>'
            f'<span class="sc">{html.escape(ev.clock_display or DASH)} '
            f'{html.escape(period_label(ev.period))}</span>'
            "</div>"
        )
    st.markdown(f'<div class="feed">{"".join(parts)}</div>', unsafe_allow_html=True)


def active_correction_alert() -> dict | None:
    """Most recent market-impacting correction still inside the banner window."""
    now = datetime.now().astimezone()
    for row in reversed(st.session_state.corrections):
        if not row.get("market_impacting"):
            continue
        if row["key"] == st.session_state.banner_dismissed_key:
            return None
        try:
            age = (now - datetime.fromisoformat(row["detected_at"])).total_seconds()
        except (ValueError, TypeError):
            age = 0
        if age <= CORRECTION_BANNER_SECONDS:
            return row
        return None
    return None


_WARNING_STYLES = {
    "alert": "background-color:#3a1600; color:#ffd966; border:2px solid #ff9900",
    "ok":    "background-color:#132117; color:#66ff99; border:2px solid #2e6b45",
    "info":  "background-color:#0d1f3c; color:#66aaff; border:2px solid #2255aa",
}


def warning_box(message: str, warning_type: str = "ok") -> None:
    """The house status banner, ported from `nfl_qc_tool/components/warning_box.py`
    (itself a direct port of the NHL tool's). `message` is injected as HTML, as it
    is there, so callers escape any feed-derived text themselves."""
    style = _WARNING_STYLES.get(warning_type, _WARNING_STYLES["ok"])
    st.markdown(
        f'<div style="margin-top:10px; margin-bottom:18px; padding:16px; border-radius:10px;'
        f' font-size:22px; font-weight:700; {style}">{message}</div>',
        unsafe_allow_html=True,
    )


def status_line(message: str) -> None:
    """The compact form of the banner, for the state that needs no attention.

    `warning_box` is kept for every state that does. `message` is injected as
    HTML, exactly as there, so callers escape feed-derived text themselves.
    """
    st.markdown(
        f'<div class="statusline"><span class="dot"></span>{message}</div>',
        unsafe_allow_html=True,
    )


def game_header(game: GameInfo, clock_text: str, period_text: str, updated: str) -> None:
    """Teams, scores, period, clock and the last successful poll, in one block.

    Scores sit inboard of the names so the eye lands on the two numbers together.
    """
    period = (
        f'<span class="pd">{html.escape(period_text)}</span>' if period_text else ""
    )
    st.markdown(
        '<div class="gh">'
        f'<div class="tm"><span class="nm">{html.escape(game.away.display_name)}</span>'
        f'<span class="sc">{game.away_score}</span></div>'
        '<div class="mid">'
        f'<div class="ck">{html.escape(clock_text or DASH)}{period}</div>'
        f'<div class="up">Updated {html.escape(updated or DASH)}</div>'
        '</div>'
        f'<div class="tm h"><span class="sc">{game.home_score}</span>'
        f'<span class="nm">{html.escape(game.home.display_name)}</span></div>'
        '</div>',
        unsafe_allow_html=True,
    )


_TF_CLASS = {"Yes": "yes", "No": "no"}


def render_timeframe(rows: Sequence[dict], current_label: str = "") -> None:
    """Narrow Time / Result table for the timeframe market.

    Not `html_table`: two columns of short values do not need the page width, and
    the house Yes / No pills read as alarms down twelve rows where most rows are
    routine. The Quarter column the selectbox above already states is dropped
    here rather than in `timeframe_table`, which stays the market's own function.

    `current_label` is the window in play, or the last completed one - the row a
    trader is actually pricing.
    """
    if not rows:
        st.info("No data.")
        return
    body = []
    for row in rows:
        label = str(row.get("Time", ""))
        verdict = str(row.get("Yes/No", DASH))
        cls = _TF_CLASS.get(verdict, "na")
        cur = ' class="cur"' if label and label == current_label else ""
        body.append(
            f'<tr{cur}><td>{html.escape(label)}</td>'
            f'<td class="v {cls}">{html.escape(verdict)}</td></tr>'
        )
    st.markdown(
        '<div class="tfwrap"><table class="tf">'
        '<thead><tr><th>Time</th><th>Both Scored</th></tr></thead>'
        f'<tbody>{"".join(body)}</tbody></table></div>',
        unsafe_allow_html=True,
    )


def banner_state() -> tuple[str, str]:
    """(escaped message, house warning type) for the status banner.

    Precedence is unchanged - a live stat correction outranks a feed error, which
    outranks a feed warning - but every non-ok state maps onto the house "alert"
    key, so this banner is indistinguishable from the NFL tool's.
    """
    alert = active_correction_alert()
    if alert:
        detail = (
            f"{period_label(alert['period'])} {alert['clock']} | "
            f"{alert['original']} → {alert['updated']} | {alert['impact']} | "
            f"detected {alert['detected_display']}"
        )
        return f"⚠ STAT CORRECTION — {html.escape(detail)}", "alert"

    if st.session_state.last_error:
        text = f"DATA DELAY — {st.session_state.last_error}"
        if st.session_state.error_streak > 1:
            text += f" ({st.session_state.error_streak} consecutive)"
        return f"⚠ {html.escape(text)}", "alert"

    if st.session_state.feed_warning:
        return f"⚠ FEED WARNING — {html.escape(st.session_state.feed_warning)}", "alert"

    logged = len(st.session_state.corrections)
    impacting = sum(1 for r in st.session_state.corrections if r.get("market_impacting"))
    if logged:
        return (
            f"STATUS: OK — {logged} correction(s) logged, {impacting} market impacting",
            "ok",
        )
    return "STATUS: OK", "ok"


def render_tab_strip() -> str:
    """The house tab strip: a keyed radio, NOT st.tabs().

    st.tabs keeps the active tab in the frontend only, so the autorefresh rerun
    every few seconds rebuilds the strip and it defaults back to the first tab -
    sitting on Stat Corrections kicked you to Prematch within seconds. Worse, the
    Corrections label carries a live count, and a changed label remounts the whole
    strip, so the bounce fired exactly when a correction landed. The NFL and MLB
    tools both hit this and both moved to a keyed radio; this is the same fix.

    The options MUST stay static and the count MUST live only in `format_func`:
    were the count in the option values, the stored value "Stat Corrections (3)"
    would cease to exist the moment the count hit 4, and Streamlit raises on a
    keyed radio whose stored value is not in `options`.
    """
    # A session from an older build can hold a value that is no longer an option.
    if st.session_state.get("active_tab") not in TABS:
        st.session_state.active_tab = TAB_PREMATCH

    def label(key: str) -> str:
        if key == TAB_PREMATCH:
            return "Prematch"
        if key == TAB_LIVE:
            return "Live"
        total = len(st.session_state.corrections)
        return f"Stat Corrections {'🔴' if total else '✅'} ({total})"

    # No index= here: with a key set, session state is the source of truth for the
    # selection, which is the whole point.
    st.radio(
        "View", options=list(TABS), format_func=label,
        horizontal=True, key="active_tab", label_visibility="collapsed",
    )
    return st.session_state.active_tab


def stretch() -> dict:
    """Full-width widget kwargs.

    The other tools all pass `use_container_width=True`, which is what gives them
    full-width sidebar buttons. It was deprecated in favour of `width="stretch"`
    in Streamlit 1.49, so this picks whichever the installed version wants: same
    result on screen, no deprecation notice.
    """
    try:
        major, minor = (int(p) for p in st.__version__.split(".")[:2])
    except (ValueError, AttributeError):
        return {"use_container_width": True}
    return {"width": "stretch"} if (major, minor) >= (1, 49) else {"use_container_width": True}


def game_option_label(game: GameInfo) -> str:
    marker = {"in": "LIVE", "pre": "UPCOMING", "post": "FINAL"}.get(game.state, game.state.upper())
    when = ""
    if game.start_iso:
        try:
            dt = datetime.fromisoformat(game.start_iso.replace("Z", "+00:00")).astimezone()
            when = dt.strftime("%a %I:%M %p").replace(" 0", " ")
        except ValueError:
            when = ""
    score = ""
    if game.state in ("in", "post"):
        score = f"  {game.away.abbr} {game.away_score}-{game.home_score} {game.home.abbr}"
    detail = game.status_detail or when
    return f"[{marker}] {game.away.display_name} @ {game.home.display_name}  ·  {detail}{score}"


# ===========================================================================
# SECTION 9 - SETUP PANEL (sidebar)
# Flow: Load Live Games -> select game -> rosters auto-load -> pick 2 key
# players per team -> Track Game.
# ===========================================================================

def sync_key_player(widget_key: str, side: str, slot: int) -> None:
    st.session_state.key_players[side][slot] = st.session_state[widget_key]


def key_player_controls(container, side: str, team: TeamInfo,
                        roster: Sequence[RosterPlayer], prefix: str) -> None:
    """Two dropdowns for one team, writing into the canonical key_players slot.

    Rendered both in the sidebar (pre-track) and in the Live tab ("Edit Key
    Players"). Both write to the same slot, so editing mid-game never resets
    corrections, first-basket results or anything else already recorded.
    """
    if not roster:
        container.caption(f"{team.abbr}: roster unavailable")
        return

    ids = [p.player_id for p in roster]
    label_by_id = {p.player_id: p.label for p in roster}
    current = st.session_state.key_players.get(side) or []
    while len(current) < KEY_PLAYERS_PER_TEAM:
        current.append(ids[len(current)] if len(current) < len(ids) else ids[0])
    st.session_state.key_players[side] = current[:KEY_PLAYERS_PER_TEAM]

    container.markdown(
        f'<div class="subsect">{html.escape(team.display_name)} key players</div>',
        unsafe_allow_html=True,
    )
    for slot in range(KEY_PLAYERS_PER_TEAM):
        wkey = f"{prefix}_{side}_{slot}"
        value = st.session_state.key_players[side][slot]
        # Switching games swaps the roster, so a stored id can fall outside the
        # new option set. Repoint it before the widget is built, or Streamlit
        # raises on a session_state value that is not in `options`.
        if value not in ids:
            value = ids[min(slot, len(ids) - 1)]
            st.session_state.key_players[side][slot] = value
        # Mirror the canonical slot into the widget so the sidebar copy and the
        # Live-tab copy of this control can never drift apart.
        if st.session_state.get(wkey) != value:
            st.session_state[wkey] = value
        container.selectbox(
            f"KP {slot + 1}",
            options=ids,
            format_func=lambda pid: label_by_id.get(pid, pid),
            key=wkey,
            on_change=sync_key_player,
            args=(wkey, side, slot),
            label_visibility="collapsed",
        )


def render_manual_id_entry(sb) -> None:
    """Load a game by ESPN game ID, bypassing the scoreboard.

    Deliberately the same shape as the NHL / NFL / MLB tools: a plainly visible
    text input plus a button that selects the id directly. It was previously
    buried in a collapsed expander, and it made a feed request before the id
    counted as loaded - so a slow or unhelpful response left nothing selected.
    Now the id is resolved when it is needed (see `resolve_selected_game`), which
    puts any feed problem in the status banner alongside every other one.
    """
    manual_id = sb.text_input(
        "Or enter a Game ID manually",
        placeholder="e.g. 401810433",
        disabled=st.session_state.tracking,
    )
    if sb.button("Load Manual Game ID", disabled=st.session_state.tracking, **stretch()):
        gid = extract_game_id(manual_id)
        if not gid:
            sb.error("Enter a numeric ESPN game ID, or paste the full game URL.")
        else:
            if gid != st.session_state.selected_game_id:
                st.session_state.selected_game_id = gid
                st.session_state.selected_game_label = f"Manual ({gid})"
                reset_game_state()
            st.session_state.manual_ids = set(st.session_state.manual_ids) | {gid}
            sb.success(f"Game ID {gid} loaded.")


def resolve_selected_game() -> tuple[GameInfo | None, str | None]:
    """(GameInfo, error) for whatever id is currently selected.

    A game picked off the slate is already in `games`. A manually loaded id is
    not, so its identity comes from the summary header - one feed call the house
    pattern cannot avoid here, because unlike the NHL and MLB providers, ESPN
    only exposes team ids alongside the game, and every market in this tool is
    keyed on team id.
    """
    gid = st.session_state.selected_game_id
    if not gid:
        return None, None
    for blob in st.session_state.games:
        if blob["game_id"] == gid:
            return _game_from_dict(blob), None
    try:
        return game_from_id(gid), None
    except DataSourceError as exc:
        return None, f"Game ID {gid}: {exc}"


def render_setup_sidebar() -> tuple[GameInfo | None, str | None]:
    """The house sidebar: title, Load Live Games, game selector, manual id,
    Track Game. Returns (game, error) so `main` can banner a resolution failure.
    """
    sb = st.sidebar
    sb.markdown("## NBA Markets")

    day_choice = sb.date_input(
        "Slate date", value=date.today(), format="YYYY-MM-DD", key="slate_date"
    )
    if sb.button("Load Live Games", type="primary", **stretch()):
        day = None if day_choice == date.today() else day_choice.strftime("%Y%m%d")
        st.session_state.scoreboard_day = day
        try:
            st.session_state.games = [asdict(g) for g in list_games(day)]
            st.session_state.games_error = (
                None if st.session_state.games else "No games found for that date."
            )
        except DataSourceError as exc:
            st.session_state.games = []
            st.session_state.games_error = str(exc)
        st.session_state.games_loaded = True

    if st.session_state.games_error:
        sb.warning(st.session_state.games_error, icon="⚠️")

    games = [_game_from_dict(g) for g in st.session_state.games]
    ids = [g.game_id for g in games]
    by_id = {g.game_id: g for g in games}

    # index=None plus a placeholder, as in the NFL tool: a manually loaded id is
    # not on the slate, and the selector must show "no slate game picked" rather
    # than silently snapping the selection onto the first game of the day.
    selected = st.session_state.selected_game_id
    chosen_id = sb.selectbox(
        "Game",
        options=ids,
        index=ids.index(selected) if selected in ids else None,
        format_func=lambda gid: game_option_label(by_id[gid]),
        placeholder="Load games first",
        label_visibility="collapsed",
        disabled=st.session_state.tracking,
    )
    if chosen_id and chosen_id != st.session_state.selected_game_id:
        st.session_state.selected_game_id = chosen_id
        st.session_state.selected_game_label = game_option_label(by_id[chosen_id])
        st.session_state.manual_ids = set(st.session_state.manual_ids) - {chosen_id}
        reset_game_state()

    sb.divider()
    render_manual_id_entry(sb)
    sb.divider()

    game, game_error = resolve_selected_game()
    if game_error:
        sb.warning(game_error, icon="⚠️")
        return None, game_error
    if game is None:
        if not st.session_state.games_loaded:
            sb.caption("Click **Load Live Games** to begin, or load a game by ID.")
        return None, None

    # Rosters load automatically for both teams once a game is selected. A failure
    # is reported but never blocking: the play-by-play carries its own player
    # names, so every market still results without a roster.
    away_roster, home_roster, roster_error = load_rosters(game)
    if roster_error:
        sb.warning(roster_error, icon="⚠️")
    else:
        sb.caption(
            f"Rosters loaded · {game.away.abbr} {len(away_roster)} · "
            f"{game.home.abbr} {len(home_roster)}"
        )

    if not st.session_state.tracking:
        key_player_controls(sb, "away", game.away, away_roster, "sb")
        key_player_controls(sb, "home", game.home, home_roster, "sb")

        sb.divider()
        if sb.button("▶  Track Game", type="primary", **stretch()):
            st.session_state.tracking = True
            st.session_state.last_error = None
            st.session_state.error_streak = 0
            st.session_state.rate_limit_skip_remaining = 0
            # Restore an earlier session for this game if one exists, so a
            # browser refresh does not lose the correction log.
            if load_state(game.game_id):
                sb.caption("Restored previous tracking state for this game.")
            st.rerun()
    else:
        sb.caption(f"{game.away.display_name} @ {game.home.display_name}")
        if sb.button("Stop Tracking", **stretch()):
            st.session_state.tracking = False
            st.rerun()
        if sb.button("Refresh now", **stretch()):
            fetch_summary.clear()
            fetch_scoreboard.clear()
            st.rerun()
        with sb.expander("Reset tracked state"):
            st.caption(
                "Clears the correction log, snapshot and key-player alert history "
                "for this game. Cannot be undone."
            )
            if st.button("Confirm reset", type="secondary"):
                clear_state(game.game_id)
                st.rerun()

        ok = st.session_state.last_fetch_ok
        sb.markdown(
            f'<div style="font-size:12px; opacity:0.6;">Last good fetch: {html.escape(ok or "never")}'
            + (f" · errors: {st.session_state.error_streak}"
               if st.session_state.error_streak else "")
            + "</div>",
            unsafe_allow_html=True,
        )
        sb.markdown(
            f'<div style="font-size:12px; opacity:0.6;">Interval: {REFRESH_SECONDS}s</div>',
            unsafe_allow_html=True,
        )

    return game, None


def load_rosters(game: GameInfo) -> tuple[list[RosterPlayer], list[RosterPlayer], str | None]:
    """Both teams' rosters. A single failed team is reported, not fatal."""
    errors: list[str] = []
    away: list[RosterPlayer] = []
    home: list[RosterPlayer] = []
    for team, bucket in ((game.away, "away"), (game.home, "home")):
        try:
            players = roster_players(team.team_id)
        except DataSourceError as exc:
            errors.append(f"{team.abbr}: {exc}")
            players = []
        if bucket == "away":
            away = players
        else:
            home = players
    return away, home, "; ".join(errors) or None


# ===========================================================================
# SECTION 10 - TABS
# ===========================================================================

@dataclass
class TrackedGame:
    """Everything the tabs need, computed once per rerun."""

    game: GameInfo
    events: list[GameEvent] = field(default_factory=list)
    away_roster: list[RosterPlayer] = field(default_factory=list)
    home_roster: list[RosterPlayer] = field(default_factory=list)
    away_starters: list[RosterPlayer] = field(default_factory=list)
    home_starters: list[RosterPlayer] = field(default_factory=list)
    away_starter_source: str = ""
    home_starter_source: str = ""
    # id -> shirt number, boxscore first and the current rosters only as a fallback
    # (see `extract_boxscore_players`). Empty until the boxscore is published.
    jersey_by_id: dict[str, str] = field(default_factory=dict)
    max_period: int = 0

    @property
    def abbr_for(self) -> dict[str, str]:
        return {self.game.away.team_id: self.game.away.abbr,
                self.game.home.team_id: self.game.home.abbr}


def _first_event_rows(events: Sequence[GameEvent], game: GameInfo) -> list[dict]:
    """Game / away / home rows for the first-event markets.

    Shared by Prematch and Second Half so the two tables cannot drift apart -
    they previously carried different column headers for the same market.
    """
    return [
        {"Team": "Game First", **first_event_row(events)},
        {"Team": game.away.display_name, **first_event_row(events, game.away.team_id)},
        {"Team": game.home.display_name, **first_event_row(events, game.home.team_id)},
    ]


def render_prematch_tab(tg: TrackedGame) -> None:
    game = tg.game

    sect("Game / Team First Field Goal")
    # First Dunk resolves on the first MADE dunk, classified from the feed's shot type
    # ("Driving Dunk Shot" etc). Limitations are in the README, not on screen.
    html_table(_first_event_rows(tg.events, game))

    sect("Player First Shot — Starters")
    left, right = st.columns(2, gap="medium")
    for col, team, starters, source in (
        (left, game.away, tg.away_starters, tg.away_starter_source),
        (right, game.home, tg.home_starters, tg.home_starter_source),
    ):
        with col:
            subsect(f"{team.display_name} Players")
            if not starters:
                note("Lineup not available yet.")
                continue
            html_table([
                {"Player": p.name, **player_first_shot_row(tg.events, p.player_id)}
                for p in starters
            ])
            note(f"Lineup source: {source}")


def _jersey_num(jersey: str) -> int:
    """Jersey as an int for ordering; unnumbered players sort last.

    The five are shown in shirt-number order rather than in the order they came
    on, so that a glance at the same team twice reads the same way and only the
    arrow moves.
    """
    return int(jersey) if jersey.isdigit() else 999


def _lineup_row(entry: FloorEntry, jersey: str, entering: bool) -> str:
    """One player row. Full name, shirt number, and - for the player who just came
    on - a green up arrow and the clock they came on at.

    The arrow slot is emitted on every row, empty where there is nothing to mark,
    so the shirt numbers stay in one column and the five do not shuffle sideways
    when a substitution lands.
    """
    num = f"#{jersey}" if jersey else DASH
    arrow = f'<span class="arw in">{ARROW_IN}</span>' if entering else '<span class="arw"></span>'
    when = (
        f'<span class="cl">{html.escape(entry.clock_display)} '
        f'{html.escape(period_label(entry.period))}</span>'
        if entering and entry.period and entry.clock_display
        else ""
    )
    return (
        f'<div class="lurow{" in" if entering else ""}">{arrow}'
        f'<span class="num">{html.escape(num)}</span>'
        f'<span class="nm">{html.escape(entry.name)}</span>{when}</div>'
    )


def render_recent_subs(events: Sequence[GameEvent], team_id: str) -> None:
    """One team's recent substitutions, newest first, under its own five.

    Green up arrow for the player coming on, red down arrow for the player going
    off, so the direction is readable without parsing the sentence. Every row is
    at the same contrast: the order already says which is newest. No team label
    and no `SUB` tag either - it sits inside that team's column, directly under
    that team's players, which says both already.
    """
    subs = recent_team_subs(events, team_id)
    if not subs:
        return

    rows = "".join(
        '<div class="subrow">'
        f'<span class="gt">{html.escape(ev.clock_display or DASH)} '
        f'{html.escape(period_label(ev.period))}</span>'
        f'<span class="arw in">{ARROW_IN}</span>'
        f'<span class="io in">{html.escape(ev.sub_in_name or ev.sub_in_id)}</span>'
        f'<span class="arw out">{ARROW_OUT}</span>'
        f'<span class="io out">{html.escape(ev.sub_out_name or ev.sub_out_id)}</span>'
        "</div>"
        for ev in subs
    )
    st.markdown(
        f'<div class="subs"><div class="subhd">Recent subs</div>{rows}</div>',
        unsafe_allow_html=True,
    )


def render_floor_panel(tg: TrackedGame) -> None:
    """The five on the floor per team, plus that team's recent substitutions.

    Sits at the very top of the Live tab, above the key-player flash alerts, so
    its position never moves: an alert appearing would otherwise push it down the
    screen, which is the one thing a panel meant to be glanced at cannot do.
    See `team_floor` for how the five are derived and how well that was measured.

    The green up arrow on a row comes from `latest_substitutions` (both teams, the
    whole break) while the list underneath comes from `recent_team_subs` (this
    team, longer history), so an arrowed player is always the incoming name on the
    top line of the list below and the two never contradict each other.
    """
    game = tg.game
    jersey_by_id = tg.jersey_by_id
    entering = {ev.sub_in_id for ev in latest_substitutions(tg.events)}

    sect("On The Floor")
    left, right = st.columns(2, gap="medium")
    for col, team, starters, source in (
        (left, game.away, tg.away_starters, tg.away_starter_source),
        (right, game.home, tg.home_starters, tg.home_starter_source),
    ):
        with col:
            floor = team_floor(tg.events, team.team_id, starters)
            subsect(team.display_name)

            if not floor.on_floor:
                note("Lineup not available yet.")
                continue

            rows = "".join(
                _lineup_row(
                    entry,
                    jersey_by_id.get(entry.player_id, ""),
                    entry.player_id in entering,
                )
                for entry in sorted(
                    floor.on_floor,
                    key=lambda e: (_jersey_num(jersey_by_id.get(e.player_id, "")), e.name),
                )
            )
            st.markdown(f'<div class="lu">{rows}</div>', unsafe_allow_html=True)
            render_recent_subs(tg.events, team.team_id)

            # Silent while the five are trustworthy, which is the normal case.
            # It speaks up only when they are not.
            if floor.unverified or len(floor.on_floor) != 5:
                note(
                    f"Lineup unverified: {len(floor.on_floor)} players shown, "
                    f"{floor.unverified} substitution(s) for a player the feed had not "
                    f"shown on the floor. Check the boxscore before resulting anything."
                )
            elif "boxscore" not in source:
                note(f"Lineup source: {source}")


def render_live_tab(tg: TrackedGame) -> None:
    game = tg.game
    away, home = game.away, game.home

    # --- 0. players on the floor (first, so alerts never shift it) ---------
    render_floor_panel(tg)

    # --- key player flash alerts (deduped, time-limited) ------------------
    fresh = []
    now = time.time()
    for alert in st.session_state.kp_alerts:
        if now - alert["ts"] <= KEY_ALERT_SECONDS:
            fresh.append(alert)
    st.session_state.kp_alerts = fresh
    for alert in reversed(fresh[-4:]):
        st.markdown(
            f'<div class="alert">{html.escape(alert["text"])}</div>',
            unsafe_allow_html=True,
        )

    # --- 1. key player tracker -------------------------------------------
    # Heading and its action on one row. The editor is a plain container behind a
    # small secondary button rather than an expander, so the heading is not
    # competing with a full-width clickable bar for attention.
    h_col, b_col = st.columns([8, 2], vertical_alignment="bottom")
    with h_col:
        sect("Key Player Tracker")
    with b_col:
        if st.button("Edit Key Players", key="edit_kp", type="secondary"):
            st.session_state.show_key_editor = not st.session_state.show_key_editor

    if st.session_state.show_key_editor:
        with st.container(border=True):
            st.caption(
                "Changing a key player keeps the correction log, first-basket results and "
                "everything else already recorded. A newly selected player starts from their "
                "current state, so no stale alert fires."
            )
            e_away, e_home = st.columns(2, gap="medium")
            key_player_controls(e_away, "away", away, tg.away_roster, "live")
            key_player_controls(e_home, "home", home, tg.home_roster, "live")

    name_by_id = {p.player_id: p.name for p in list(tg.away_roster) + list(tg.home_roster)}
    k_away, k_home = st.columns(2, gap="medium")
    for col, side, team in ((k_away, "away", away), (k_home, "home", home)):
        with col:
            subsect(team.display_name)
            for pid in st.session_state.key_players.get(side) or []:
                display = name_by_id.get(pid) or _name_of(pid, tg.events) or pid
                latest = latest_made_fg(tg.events, pid)
                hot = any(a["player_id"] == pid for a in st.session_state.kp_alerts)
                if latest:
                    line = format_event_line(
                        latest, away.abbr, home.abbr, team.abbr, include_team=False
                    )
                    cls = "kp hot" if hot else "kp"
                    body = f'<div class="ln">{html.escape(line)}</div>'
                else:
                    cls = "kp"
                    body = '<div class="ln none">No made field goal yet</div>'
                st.markdown(
                    f'<div class="{cls}"><div class="nm">{html.escape(display)}</div>{body}</div>',
                    unsafe_allow_html=True,
                )

    # --- 2. three most recent FG attempts per team + made-shots feed -------
    sect(f"{RECENT_FG_COUNT} Most Recent Field Goal Attempts")
    # Live market state, not a description of the section - kept, unlike the
    # explanatory captions.
    note(
        f"Open market: Next Field Goal after "
        f"{_anchor_text(open_market_anchor(tg.events))} ({away.abbr}-{home.abbr})"
    )
    c_away, c_made, c_home = st.columns(3, gap="medium")
    with c_away:
        subsect(away.display_name)
        render_feed(
            recent_fg_market_events(tg.events, tg.abbr_for, away.team_id, RECENT_FG_COUNT),
            "No field goal attempts yet.",
        )
    with c_made:
        subsect("Made Field Goals")
        # The only panel that mixes teams, so the only one that needs the abbreviation.
        render_feed(
            recent_fg_market_events(
                tg.events, tg.abbr_for, None, RECENT_FG_COUNT, made_only=True
            ),
            "No made field goals yet.",
            include_team=True,
        )
    with c_home:
        subsect(home.display_name)
        render_feed(
            recent_fg_market_events(tg.events, tg.abbr_for, home.team_id, RECENT_FG_COUNT),
            "No field goal attempts yet.",
        )

    # --- 3. timeframe table ----------------------------------------------
    sect("Timeframe Both Teams To Score")
    periods = sorted({e.period for e in tg.events if e.period > 0} | {1, 2, 3, 4})
    # First render defaults to the quarter in play; after that the trader's
    # choice sticks across auto-refreshes (the widget key holds it).
    if st.session_state.timeframe_period not in periods:
        st.session_state.timeframe_period = min(max(tg.max_period, 1), max(periods))
    tf_col, _spacer = st.columns([1, 3])
    with tf_col:
        period = st.selectbox(
            "Quarter", options=periods, format_func=period_label, key="timeframe_period",
        )

    # The row being priced: the window the clock is in, or the last one if the
    # quarter is over. Read-only use of the same two helpers the table itself uses.
    windows = timeframe_windows(period)
    lowest, finished = period_progress(tg.events, period)
    finished = finished or game.is_final
    current_label = ""
    if finished:
        current_label = windows[-1][0]
    elif lowest is not None:
        idx = window_index(lowest, period)
        if idx is not None:
            current_label = windows[idx][0]

    render_timeframe(
        timeframe_table(tg.events, period, away.team_id, home.team_id, game.is_final),
        current_label,
    )
    # Legend, for the record rather than the screen: Yes = both teams scored inside
    # the exact window (field goals and free throws); No = window complete with both
    # teams not scoring; - = not reached or in progress. The highlighted row is the
    # window in play, or the last one completed.

    # --- 4. second half ---------------------------------------------------
    sect("Second Half")
    if tg.max_period < 3:
        note("Activates when the third quarter begins.")
    else:
        html_table(_first_event_rows([e for e in tg.events if e.period >= 3], game))


def render_corrections_tab() -> None:
    log = st.session_state.corrections
    impacting = [r for r in log if r.get("market_impacting")]

    top = st.columns([2, 1, 1])
    with top[0]:
        view = st.radio(
            "View", options=["All Corrections", "Market-Impacting Corrections Only"],
            horizontal=True, label_visibility="collapsed", key="corr_view",
        )
    with top[1]:
        st.caption(f"Logged: {len(log)}  ·  Market impacting: {len(impacting)}")
    with top[2]:
        if active_correction_alert() and st.button("Dismiss banner"):
            newest = active_correction_alert()
            if newest:
                st.session_state.banner_dismissed_key = newest["key"]
            st.rerun()

    rows = impacting if view.startswith("Market-Impacting") else log
    if not rows:
        note("No corrections detected in this tracking session.")
        return

    html_table(
        [
            {
                "Detected": r["detected_display"],
                "Period": period_label(r["period"]),
                "Clock": r["clock"],
                "Original Event": r["original"],
                "Updated Event": r["updated"],
                "Market Impacting": "Yes" if r["market_impacting"] else "No",
                "Impact Type": r["impact"],
            }
            for r in reversed(rows)  # newest first
        ],
        wrap_columns={"Impact Type"},
    )
    # Append-only for this tracking session; an event corrected more than once is logged
    # once per change. Keyed on the feed's stable event id plus a fingerprint of period,
    # clock, team, shooter, result, shot type and description. See README 4.7.


# ===========================================================================
# SECTION 11 - TRACKING PIPELINE
# ===========================================================================

def build_tracked_game(game: GameInfo) -> TrackedGame:
    """Fetch, normalise, diff for corrections, and fire key-player alerts."""
    tg = TrackedGame(game=game)

    away_roster, home_roster, _ = load_rosters(game)
    tg.away_roster, tg.home_roster = away_roster, home_roster

    try:
        payload = fetch_summary(game.game_id)
        st.session_state.last_fetch_ok = datetime.now().astimezone().strftime("%H:%M:%S")
        st.session_state.last_error = None
        st.session_state.error_streak = 0
    except RateLimitedError as exc:
        # Must precede DataSourceError - RateLimitedError subclasses it, and a 429
        # needs the cooldown rather than another attempt in four seconds.
        st.session_state.rate_limit_skip_remaining = RATE_LIMIT_SKIP_TICKS
        st.session_state.last_error = str(exc)
        st.session_state.error_streak += 1
        payload = None
    except DataSourceError as exc:
        # Keep showing the last known good state rather than blanking the tool.
        st.session_state.last_error = str(exc)
        st.session_state.error_streak += 1
        payload = None

    if payload is None:
        return tg

    box_names, box_starters, box_jerseys = extract_boxscore_players(payload)
    name_map = {p.player_id: p.name for p in away_roster + home_roster}
    name_map.update(box_names)  # boxscore names win: they match the pbp exactly

    tg.jersey_by_id = {p.player_id: p.jersey for p in away_roster + home_roster if p.jersey}
    tg.jersey_by_id.update(box_jerseys)  # and boxscore numbers win, same reason

    tg.events = normalise_plays(payload, name_map)
    tg.max_period = max((e.period for e in tg.events), default=0)

    tg.away_starters, tg.away_starter_source = resolve_starters(
        game.away.team_id, box_starters, tg.events, away_roster
    )
    tg.home_starters, tg.home_starter_source = resolve_starters(
        game.home.team_id, box_starters, tg.events, home_roster
    )

    # --- stat corrections -------------------------------------------------
    rows, snapshot, max_seq, total, warning = detect_corrections(
        tg.events,
        st.session_state.snapshot,
        st.session_state.snapshot_max_seq,
        st.session_state.snapshot_total,
    )
    st.session_state.feed_warning = warning
    # On a suspected truncated payload the OLD snapshot is kept: letting a short
    # payload become the baseline would make the next full one look like mass
    # additions.
    if not warning:
        st.session_state.snapshot = snapshot
        st.session_state.snapshot_max_seq = max_seq
        st.session_state.snapshot_total = total
    record_corrections(rows, game.game_id)

    # --- key player made-basket alerts -----------------------------------
    active = {
        pid
        for side in ("away", "home")
        for pid in (st.session_state.key_players.get(side) or [])
        if pid
    }
    previously_active: set[str] = set(st.session_state.kp_active or set())

    # A newly selected key player is baselined to their current state so we do
    # not flash an alert for a basket made before they were being tracked.
    for pid in active - previously_active:
        latest = latest_made_fg(tg.events, pid)
        st.session_state.kp_seen[pid] = latest.event_id if latest else ""
    st.session_state.kp_active = active

    name_by_id = {p.player_id: p.name for p in away_roster + home_roster}
    for pid in active:
        latest = latest_made_fg(tg.events, pid)
        if not latest:
            continue
        if st.session_state.kp_seen.get(pid) == latest.event_id:
            continue
        st.session_state.kp_seen[pid] = latest.event_id
        display = name_by_id.get(pid) or latest.player_name or pid
        shot = "3-pointer" if latest.points == 3 else "2-pointer"
        if latest.is_dunk:
            shot = "dunk"
        st.session_state.kp_alerts.append(
            {
                "key": latest.event_id,
                "player_id": pid,
                "ts": time.time(),
                "text": (
                    f"KEY PLAYER ALERT: {display} made a {shot} at "
                    f"{latest.clock_display} in {period_label(latest.period)}"
                ),
            }
        )
    save_state(game.game_id)  # throttled; a new correction forces its own write
    return tg


def refresh_live_status(game: GameInfo) -> GameInfo:
    """Keep status/score current for the tracked game.

    A game added by ID may not be on the selected slate at all, so it reads its
    status from the summary header instead of the scoreboard. Both calls are
    cached at 20s, so either way this is one request per 20s, not per rerun.
    """
    if game.game_id in st.session_state.manual_ids:
        try:
            return game_from_id(game.game_id)
        except DataSourceError:
            return game
    try:
        for g in list_games(st.session_state.scoreboard_day):
            if g.game_id == game.game_id:
                return g
    except DataSourceError:
        pass
    return game


# ===========================================================================
# SECTION 12 - MAIN
# ===========================================================================

def main() -> None:
    st.set_page_config(page_title=APP_TITLE, layout="wide", initial_sidebar_state="expanded")
    st.markdown(CSS, unsafe_allow_html=True)
    init_state()

    game, game_error = render_setup_sidebar()

    if not st.session_state.tracking or game is None:
        if game_error:
            warning_box(f"⚠ {html.escape(game_error)}", "alert")
        else:
            status_line("STATUS: OK — Load a game and click Track Game")
        note(
            "Sidebar: Load Live Games (or enter a Game ID) → select a game → set two "
            "key players per team → Track Game. No play-by-play requests are made "
            "until tracking starts."
        )
        return

    # Auto-refresh only while a tracked game can still change. Registered before
    # the cooldown check below, so a rate-limited page still wakes up to resume.
    if not game.is_final:
        interval = PREGAME_REFRESH_SECONDS if game.state == "pre" else REFRESH_SECONDS
        if st_autorefresh is not None:
            st_autorefresh(interval=interval * 1000, key="nba_tick")
        else:
            st.warning(
                "`streamlit-autorefresh` is not installed, so the page will not update on "
                "its own. Install it (see requirements.txt) or use **Refresh now**.",
                icon="⚠️",
            )

    # Rate-limit cooldown, as in the NHL / NFL / MLB tools: after a 429, sit out a
    # couple of ticks instead of polling straight back into it.
    if st.session_state.rate_limit_skip_remaining > 0:
        st.session_state.rate_limit_skip_remaining -= 1
        secs_left = st.session_state.rate_limit_skip_remaining * REFRESH_SECONDS
        warning_box(f"⚠ RATE LIMITED — resuming in ~{secs_left}s", "alert")
        return

    game = refresh_live_status(game)
    tg = build_tracked_game(game)

    # Same status derivation as before, split into clock and period for the header.
    clock_text = game.status_detail or ("Not started" if game.state == "pre" else "")
    period_text = ""
    if game.is_live and tg.events:
        last = tg.events[-1]
        clock_text, period_text = last.clock_display, period_label(last.period)
    game_header(game, clock_text, period_text, st.session_state.last_fetch_ok or "")

    # One status, above the tab strip, exactly where the other tools place it.
    # Anything that needs attention keeps the house banner; STATUS: OK does not,
    # so it drops to a single compact line.
    message, kind = banner_state()
    if kind == "ok":
        status_line(message)
    else:
        warning_box(message, kind)

    if game.state == "pre":
        note("Game has not started. Prematch tables will populate from the first play.")
    elif not tg.events and st.session_state.last_error is None:
        note("Play-by-play is not published for this game yet.")

    active = render_tab_strip()
    if active == TAB_PREMATCH:
        render_prematch_tab(tg)
    elif active == TAB_LIVE:
        render_live_tab(tg)
    else:
        render_corrections_tab()


if __name__ == "__main__":
    main()
