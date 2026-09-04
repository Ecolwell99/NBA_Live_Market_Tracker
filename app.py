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

All source-specific parsing is confined to SECTION 3 (fetch) and SECTION 4
(normalise). Everything downstream works only on the neutral `GameEvent`
dataclass, so swapping providers means rewriting those two sections only.
"""

from __future__ import annotations

import html
import json
import re
import time
from dataclasses import dataclass, asdict, field
from datetime import date, datetime
from pathlib import Path
from typing import Any, Callable, Iterable, Sequence

import pandas as pd
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

# On-disk state so a hard browser refresh (new Streamlit session) does not
# forget an already-detected correction. Best-effort only.
STATE_DIR = Path(__file__).resolve().parent / ".tracker_state"


# ===========================================================================
# SECTION 2 - STYLING
# Dark, compact, no decoration. Strong colour is reserved for the two alert
# types (stat correction, key-player basket).
# ===========================================================================

CSS = """
<style>
div.block-container { padding-top: 1.1rem; padding-bottom: 1.5rem; max-width: 1560px; }
#MainMenu, footer, header [data-testid="stStatusWidget"] { visibility: hidden; }

/* --- section label: small caps rule instead of a big header --- */
.sect {
  font-size: 10.5px; text-transform: uppercase; letter-spacing: .09em;
  color: #7f8ba0; font-weight: 700; margin: 16px 0 6px 0;
  border-bottom: 1px solid #232a38; padding-bottom: 3px;
}
.sect:first-child { margin-top: 4px; }
.subsect {
  font-size: 11.5px; font-weight: 700; color: #c3cad8; letter-spacing: .02em;
  margin: 10px 0 4px 0;
}
.note { font-size: 10.5px; color: #6c7688; margin: 3px 0 0 0; }

/* --- status banner --- */
.banner {
  font-size: 12px; font-weight: 600; letter-spacing: .02em;
  padding: 6px 10px; border-radius: 3px; margin-bottom: 10px;
  font-family: ui-monospace, "Cascadia Mono", Consolas, monospace;
}
.banner-ok   { background: #111722; border: 1px solid #26303f; color: #7f8ba0; }
.banner-warn { background: #2b2410; border: 1px solid #6f5a1c; color: #f0dda6; }
.banner-corr { background: #3a1216; border: 1px solid #a02a35; color: #ffd6da; }

/* --- alerts --- */
.alert {
  font-size: 12px; padding: 6px 10px; border-radius: 3px; margin-bottom: 4px;
  font-family: ui-monospace, "Cascadia Mono", Consolas, monospace;
}
.alert-kp   { background: #33280d; border: 1px solid #8a6b1c; color: #ffeec2; }
.alert-corr { background: #3a1216; border: 1px solid #a02a35; color: #ffd6da; }

/* --- compact tables --- */
table.nbat { border-collapse: collapse; width: 100%; font-size: 12.5px; margin-bottom: 2px; }
table.nbat th {
  text-align: left; font-weight: 600; color: #8b94a7; text-transform: uppercase;
  letter-spacing: .05em; font-size: 10px; padding: 5px 8px;
  border-bottom: 1px solid #2c3444; white-space: nowrap;
}
table.nbat td {
  padding: 5px 8px; border-bottom: 1px solid #1b212d; color: #dfe3ec;
  white-space: nowrap; vertical-align: top;
}
table.nbat tr:last-child td { border-bottom: none; }
table.nbat td.k { color: #98a2b6; font-weight: 600; }
table.nbat tr.grp td { background: #151a24; font-weight: 600; color: #e8ecf4; }
table.nbat td.dim { color: #5d6675; }
.yes { color: #6fae7f; font-weight: 600; }
.no  { color: #c07a80; font-weight: 600; }

/* --- event feed rows --- */
.feed { margin-bottom: 2px; }
.frow {
  font-size: 12px; padding: 4px 8px; margin-bottom: 3px; background: #12161f;
  border-left: 2px solid #2a3040; color: #dfe3ec;
  font-family: ui-monospace, "Cascadia Mono", Consolas, monospace;
}
.frow.made { border-left-color: #3f7d52; }
.frow.miss { border-left-color: #6b3a3a; }
.frow .sc { color: #79839a; }
.frow.empty { color: #5d6675; border-left-color: #232a38; }

/* --- key player card --- */
.kp { background: #12161f; border: 1px solid #232a38; border-radius: 3px;
      padding: 7px 9px; margin-bottom: 6px; }
.kp .nm { font-size: 12.5px; font-weight: 700; color: #e8ecf4; }
.kp .ln { font-size: 12px; color: #b9c1d1; margin-top: 2px;
          font-family: ui-monospace, "Cascadia Mono", Consolas, monospace; }
.kp .ln.none { color: #5d6675; }
.kp.hot { border-color: #8a6b1c; background: #1d1a10; }

/* correction before/after colouring, as in the mock-up sheet */
.was { color: #d9848c; }
.now { color: #7fb98d; }
.arr { color: #7f8ba0; }

div[data-testid="stVerticalBlock"] { gap: 0.35rem; }
</style>
"""


# ===========================================================================
# SECTION 3 - DATA SOURCE (ESPN)
# The only place that knows about ESPN's URL shapes and JSON keys, alongside
# SECTION 4. Swap this out to change providers.
# ===========================================================================

ESPN_BASE = "https://site.api.espn.com/apis/site/v2/sports/basketball/nba"

_HTTP_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
    ),
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
    out = [
        GameInfo(
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
        for g in fetch_scoreboard(day)
    ]
    rank = {"in": 0, "pre": 1, "post": 2}
    out.sort(key=lambda g: (rank.get(g.state, 3), g.start_iso))
    return out


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
KIND_OTHER = "other"

# Kinds that participate in stat-correction comparison. Rebounds, fouls and
# substitutions are excluded so routine feed churn cannot produce noise.
TRACKED_KINDS = (KIND_FG, KIND_FT, KIND_TURNOVER)

# Word-boundary match so player names like "Ryan Dunn" never read as a dunk.
_DUNK_RE = re.compile(r"\bdunk", re.IGNORECASE)
_ISO_CLOCK_RE = re.compile(r"^PT(?:(\d+)M)?(?:([\d.]+)S)?$", re.IGNORECASE)
_WS_RE = re.compile(r"\s+")


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

    @property
    def is_fg_attempt(self) -> bool:
        return self.kind == KIND_FG

    @property
    def is_scoring(self) -> bool:
        return self.score_value > 0


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


def _classify_play(short_desc: str, type_text: str, shooting: bool, pts_attempted: int,
                   scoring: bool) -> tuple[str, int]:
    """(kind, points) for one play.

    Primary signals, in order of reliability:
      pointsAttempted  2/3 -> field goal, 1 -> free throw
      shortDescription "+2 Points" / "Missed 3PT" / "Turnover"
      type.text        contains "Turnover" / "Free Throw"
    """
    sd = short_desc.lower()
    tt = type_text.lower()

    if shooting and pts_attempted in (2, 3):
        return KIND_FG, pts_attempted
    if shooting and pts_attempted == 1:
        return KIND_FT, 1

    # pointsAttempted missing or zero: fall back to the text signals.
    if "3pt" in sd or "+3 points" in sd:
        return KIND_FG, 3
    if sd in ("missed fg", "+2 points"):
        return KIND_FG, 2
    if "free throw" in tt or sd in ("+1 point", "missed ft"):
        return KIND_FT, 1
    if "turnover" in sd or "turnover" in tt:
        return KIND_TURNOVER, 0
    if shooting and scoring:
        return KIND_FG, 2
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

        kind, points = _classify_play(short_desc, type_text, shooting, pts_attempted, scoring)

        # participants[0] is the shooter, including on blocked shots where the
        # description leads with the blocker's name.
        player_id = ""
        for part in raw.get("participants") or []:
            athlete = (part or {}).get("athlete") or {}
            pid = str(athlete.get("id", ""))
            if pid:
                player_id = pid
                break

        player_name = name_map.get(player_id, "") or _shooter_name_from_text(description)

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
            )
        )

    events.sort(key=lambda e: e.sequence)
    return events


def extract_boxscore_players(payload: dict) -> tuple[dict[str, str], dict[str, list[RosterPlayer]]]:
    """(id -> name map, team_id -> starters) from the boxscore.

    `starter` is a real flag in this feed and yields exactly five players per
    team once the boxscore is published (typically well before tip-off). It is
    empty for a game that has not been posted yet - see `resolve_starters`.
    """
    names: dict[str, str] = {}
    starters: dict[str, list[RosterPlayer]] = {}

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
                pos = athlete.get("position") or {}
                if entry.get("starter"):
                    found.append(
                        RosterPlayer(
                            player_id=pid,
                            name=nm,
                            jersey=str(athlete.get("jersey") or ""),
                            position=(pos.get("abbreviation") or "") if isinstance(pos, dict) else "",
                        )
                    )
        if team_id and found:
            starters[team_id] = found[:5]
    return names, starters


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


def recent_fg_attempts(events: Sequence[GameEvent], team_id: str | None, count: int,
                       made_only: bool = False) -> list[GameEvent]:
    """Most recent field-goal attempts, newest first."""
    pool = fg_attempts(events, team_id)
    if made_only:
        pool = [e for e in pool if e.made]
    return list(reversed(pool[-count:]))


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

    total = len(events)
    max_seq = max((e.sequence for e in events), default=0)
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
    live_period, live_clock = _live_edge(events)
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
    "scoreboard_day": None,
    "selected_game_id": None,
    "tracking": False,
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
    for key, value in DEFAULT_STATE.items():
        if key not in st.session_state:
            st.session_state[key] = (
                value.copy() if isinstance(value, (dict, list, set)) else value
            )


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


def clear_state(game_id: str) -> None:
    for key in CLEARED_KEYS:
        default = DEFAULT_STATE[key]
        st.session_state[key] = default.copy() if isinstance(default, (dict, list, set)) else default
    st.session_state.kp_alerts = []
    st.session_state.kp_active = set()
    st.session_state.banner_dismissed_key = None
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


def html_table(headers: Sequence[str], rows: Sequence[Sequence[Any]],
               first_col_key: bool = True, group_rows: Sequence[int] = ()) -> None:
    """Compact static table. Static HTML means no internal scrollbars, ever."""
    head = "".join(f"<th>{html.escape(str(h))}</th>" for h in headers)
    body = []
    for i, row in enumerate(rows):
        cls = ' class="grp"' if i in group_rows else ""
        cells = []
        for j, cell in enumerate(row):
            text = str(cell)
            css = []
            if first_col_key and j == 0:
                css.append("k")
            if text == DASH:
                css.append("dim")
            attr = f' class="{" ".join(css)}"' if css else ""
            # Yes/No get their own muted colouring.
            if text == "Yes":
                inner = '<span class="yes">Yes</span>'
            elif text == "No":
                inner = '<span class="no">No</span>'
            else:
                inner = html.escape(text)
            cells.append(f"<td{attr}>{inner}</td>")
        body.append(f"<tr{cls}>{''.join(cells)}</tr>")
    st.markdown(
        f'<table class="nbat"><thead><tr>{head}</tr></thead>'
        f'<tbody>{"".join(body)}</tbody></table>',
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


def render_feed(events: Sequence[GameEvent], away_abbr: str, home_abbr: str,
                abbr_for: dict[str, str], include_team: bool, empty_text: str) -> None:
    if not events:
        st.markdown(
            f'<div class="frow empty">{html.escape(empty_text)}</div>', unsafe_allow_html=True
        )
        return
    parts = []
    for ev in events:
        line = format_event_line(
            ev, away_abbr, home_abbr, abbr_for.get(ev.team_id, ""), include_team
        )
        head, _, tail = line.rpartition(" — ")
        cls = "made" if ev.made else "miss"
        parts.append(
            f'<div class="frow {cls}">{html.escape(head)} — '
            f'<span class="sc">{html.escape(tail)}</span></div>'
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


def render_banner() -> None:
    """Compact status banner. Rendered identically at the top of every tab."""
    alert = active_correction_alert()
    if alert:
        text = (
            f"STAT CORRECTION DETECTED — {period_label(alert['period'])} {alert['clock']} "
            f"| {alert['original']} → {alert['updated']} | {alert['impact']} "
            f"| detected {alert['detected_display']}"
        )
        st.markdown(f'<div class="banner banner-corr">{html.escape(text)}</div>',
                    unsafe_allow_html=True)
        return

    if st.session_state.last_error:
        streak = st.session_state.error_streak
        text = f"STATUS: DATA DELAY — {st.session_state.last_error}"
        if streak > 1:
            text += f" ({streak} consecutive)"
        st.markdown(f'<div class="banner banner-warn">{html.escape(text)}</div>',
                    unsafe_allow_html=True)
        return

    if st.session_state.feed_warning:
        st.markdown(
            f'<div class="banner banner-warn">'
            f'{html.escape("STATUS: FEED WARNING - " + st.session_state.feed_warning)}</div>',
            unsafe_allow_html=True,
        )
        return

    logged = len(st.session_state.corrections)
    impacting = sum(1 for r in st.session_state.corrections if r.get("market_impacting"))
    suffix = ""
    if logged:
        suffix = f"  ·  {logged} correction(s) logged, {impacting} market impacting"
    st.markdown(
        f'<div class="banner banner-ok">{html.escape("STATUS: OK" + suffix)}</div>',
        unsafe_allow_html=True,
    )


def dataframe_kwargs() -> dict:
    """`use_container_width` was renamed to `width="stretch"` in Streamlit 1.49."""
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


def render_setup_sidebar() -> GameInfo | None:
    sb = st.sidebar
    sb.markdown('<div class="sect">Track Game Flow</div>', unsafe_allow_html=True)

    day_choice = sb.date_input(
        "Slate date", value=date.today(), format="YYYY-MM-DD", key="slate_date"
    )
    if sb.button("Load Live Games", type="primary"):
        day = None if day_choice == date.today() else day_choice.strftime("%Y%m%d")
        st.session_state.scoreboard_day = day
        try:
            games = list_games(day)
            st.session_state.games = [asdict(g) for g in games]
            st.session_state.games_error = None if games else "No games found for that date."
        except DataSourceError as exc:
            st.session_state.games = []
            st.session_state.games_error = str(exc)
        st.session_state.games_loaded = True

    if st.session_state.games_error:
        sb.warning(st.session_state.games_error, icon="⚠️")
    if not st.session_state.games:
        if not st.session_state.games_loaded:
            sb.caption("Click **Load Live Games** to begin.")
        return None

    games = [
        GameInfo(
            game_id=g["game_id"], start_iso=g["start_iso"], state=g["state"],
            status_detail=g["status_detail"], period=g["period"],
            display_clock=g["display_clock"], away=TeamInfo(**g["away"]),
            home=TeamInfo(**g["home"]), away_score=g["away_score"], home_score=g["home_score"],
        )
        for g in st.session_state.games
    ]
    ids = [g.game_id for g in games]
    by_id = {g.game_id: g for g in games}

    # No explicit widget key here: `options` changes whenever a different slate
    # is loaded, and an auto-keyed widget re-derives from `index` instead of
    # holding a game id that no longer exists in the list.
    selected = st.session_state.selected_game_id
    index = ids.index(selected) if selected in ids else 0
    chosen_id = sb.selectbox(
        "Game", options=ids, index=index,
        format_func=lambda gid: game_option_label(by_id[gid]),
        disabled=st.session_state.tracking,
    )
    if chosen_id != st.session_state.selected_game_id:
        st.session_state.selected_game_id = chosen_id
    game = by_id[chosen_id]

    # Rosters load automatically for both teams once a game is selected.
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

        sb.markdown('<div class="sect">Start</div>', unsafe_allow_html=True)
        if sb.button("Track Game", type="primary", disabled=not (away_roster and home_roster)):
            st.session_state.tracking = True
            st.session_state.last_error = None
            st.session_state.error_streak = 0
            # Restore an earlier session for this game if one exists, so a
            # browser refresh does not lose the correction log.
            if load_state(game.game_id):
                sb.caption("Restored previous tracking state for this game.")
            st.rerun()
    else:
        sb.markdown('<div class="sect">Tracking</div>', unsafe_allow_html=True)
        sb.caption(f"{game.away.display_name} @ {game.home.display_name}")
        if sb.button("Stop Tracking"):
            st.session_state.tracking = False
            st.rerun()
        if sb.button("Refresh now"):
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
        sb.caption(
            f"Last good fetch: {ok or 'never'}"
            + (f" · errors: {st.session_state.error_streak}" if st.session_state.error_streak else "")
        )
        sb.caption(f"Auto-refresh: {REFRESH_SECONDS}s")

    return game


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
    max_period: int = 0

    @property
    def abbr_for(self) -> dict[str, str]:
        return {self.game.away.team_id: self.game.away.abbr,
                self.game.home.team_id: self.game.home.abbr}


def render_prematch_tab(tg: TrackedGame) -> None:
    render_banner()
    game = tg.game

    sect("Game / Team First Field Goal")
    rows = [
        ["Game First", *first_event_row(tg.events).values()],
        [game.away.display_name, *first_event_row(tg.events, game.away.team_id).values()],
        [game.home.display_name, *first_event_row(tg.events, game.home.team_id).values()],
    ]
    html_table(
        ["Team", "First FG Exact", "First FG Scorer", "First 3 Make", "First Dunk"],
        rows, group_rows=(0,),
    )
    note(
        "First Dunk resolves on the first MADE dunk. Dunks are classified from the feed's "
        "shot type (e.g. 'Driving Dunk Shot'); see limitations in the README."
    )

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
            table_rows = []
            for p in starters:
                r = player_first_shot_row(tg.events, p.player_id)
                table_rows.append([
                    p.name, r["First FG Attempt"], r["First FG Type"], r["First 3 Attempt"],
                ])
            html_table(
                ["Player", "First FG Attempt", "First FG Type", "First 3 Attempt"], table_rows
            )
            note(f"Lineup source: {source}")


def render_live_tab(tg: TrackedGame) -> None:
    render_banner()
    game = tg.game
    away, home = game.away, game.home

    # --- key player flash alerts (deduped, time-limited) ------------------
    fresh = []
    now = time.time()
    for alert in st.session_state.kp_alerts:
        if now - alert["ts"] <= KEY_ALERT_SECONDS:
            fresh.append(alert)
    st.session_state.kp_alerts = fresh
    for alert in reversed(fresh[-4:]):
        st.markdown(
            f'<div class="alert alert-kp">{html.escape(alert["text"])}</div>',
            unsafe_allow_html=True,
        )

    # --- 1. three most recent FG attempts per team + made-shots feed -------
    sect(f"{RECENT_FG_COUNT} Most Recent Field Goal Attempts")
    c_away, c_made, c_home = st.columns(3, gap="medium")
    with c_away:
        subsect(away.display_name)
        render_feed(
            recent_fg_attempts(tg.events, away.team_id, RECENT_FG_COUNT),
            away.abbr, home.abbr, tg.abbr_for, True, "No field goal attempts yet.",
        )
    with c_made:
        subsect("Made Shots")
        render_feed(
            recent_fg_attempts(tg.events, None, RECENT_FG_COUNT, made_only=True),
            away.abbr, home.abbr, tg.abbr_for, True, "No made field goals yet.",
        )
    with c_home:
        subsect(home.display_name)
        render_feed(
            recent_fg_attempts(tg.events, home.team_id, RECENT_FG_COUNT),
            away.abbr, home.abbr, tg.abbr_for, True, "No field goal attempts yet.",
        )

    # --- 2. key player tracker -------------------------------------------
    sect("Key Player Tracker")
    with st.expander("Edit Key Players (injuries / lineup changes)"):
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
            subsect(f"{team.abbr} KP")
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
    rows = timeframe_table(
        tg.events, period, away.team_id, home.team_id, game.is_final
    )
    html_table(
        ["Quarter", "Time", "Yes/No"],
        [[r["Quarter"], r["Time"], r["Yes/No"]] for r in rows],
        first_col_key=False,
    )
    note(
        "Yes = both teams scored inside the exact window (field goals and free throws). "
        "No = window complete with both teams not scoring. - = not reached or in progress."
    )

    # --- 4. second half ---------------------------------------------------
    sect("Second Half")
    if tg.max_period < 3:
        note("Activates when the third quarter begins.")
    else:
        second_half = [e for e in tg.events if e.period >= 3]
        sh_rows = [
            ["Game First", *first_event_row(second_half).values()],
            [away.display_name, *first_event_row(second_half, away.team_id).values()],
            [home.display_name, *first_event_row(second_half, home.team_id).values()],
        ]
        html_table(
            ["Team", "First FG Exact", "First FG Scorer", "First 3", "First Dunk"],
            sh_rows, group_rows=(0,),
        )


def render_corrections_tab() -> None:
    render_banner()

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

    frame = pd.DataFrame(
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
        ]
    )
    st.dataframe(
        frame, hide_index=True,
        height=min(38 * (len(frame) + 1) + 4, 620),
        **dataframe_kwargs(),
    )
    note(
        "Append-only for this tracking session. An event corrected more than once is "
        "logged once per change. Comparison is keyed on the feed's stable event id, with "
        "a fingerprint of period, clock, team, shooter, result, shot type and description."
    )


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
    except DataSourceError as exc:
        # Keep showing the last known good state rather than blanking the tool.
        st.session_state.last_error = str(exc)
        st.session_state.error_streak += 1
        payload = None

    if payload is None:
        return tg

    box_names, box_starters = extract_boxscore_players(payload)
    name_map = {p.player_id: p.name for p in away_roster + home_roster}
    name_map.update(box_names)  # boxscore names win: they match the pbp exactly

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
    """Re-read the scoreboard (cached, 20s) so status/score stay current."""
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

    game = render_setup_sidebar()

    if not st.session_state.tracking or game is None:
        st.markdown(
            f'<div class="sect">{html.escape(APP_TITLE)}</div>', unsafe_allow_html=True
        )
        for tab in st.tabs(["Prematch", "Live", "Stat Corrections"]):
            with tab:
                render_banner()
                note(
                    "Use the sidebar: Load Live Games → select a game → set two key "
                    "players per team → Track Game. No feed requests are made for "
                    "play-by-play until tracking starts."
                )
        return

    game = refresh_live_status(game)

    # Auto-refresh only while a tracked game can still change.
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

    tg = build_tracked_game(game)

    header = (
        f"{game.away.display_name} {game.away_score}  @  "
        f"{game.home.display_name} {game.home_score}"
    )
    status = game.status_detail or ("Not started" if game.state == "pre" else "")
    if game.is_live and tg.events:
        last = tg.events[-1]
        status = f"{period_label(last.period)} {last.clock_display}"
    st.markdown(
        f'<div class="sect">{html.escape(header)}  ·  {html.escape(status)}</div>',
        unsafe_allow_html=True,
    )

    if game.state == "pre":
        note("Game has not started. Prematch tables will populate from the first play.")
    elif not tg.events and st.session_state.last_error is None:
        note("Play-by-play is not published for this game yet.")

    prematch_tab, live_tab, corrections_tab = st.tabs(["Prematch", "Live", "Stat Corrections"])
    with prematch_tab:
        render_prematch_tab(tg)
    with live_tab:
        render_live_tab(tg)
    with corrections_tab:
        render_corrections_tab()


if __name__ == "__main__":
    main()
