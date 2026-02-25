"""
SLS_ROBUSTER v1.4.c — Single-window GUI combining:
- Retitle current live streams (Y1/Y2 or both)
- Find active/testing livestream + copy/save embeds
- Archive livestream uploads into channel playlists

Notes for first run (plain English):
- Preview Mode: lets you click around without touching YouTube. Turn it on in the app header.
- Tokens: place per-channel token files in app data as `token-<channel_id>.pickle`.
  The repo already includes these. If you need to (re)auth, disable Preview Mode and run an action.
- Client secret filename and channel defaults are loaded from `app_config.json`.
"""

from __future__ import annotations

import os
import pickle
import shutil
import subprocess
import sys
import threading
import time
import webbrowser
import datetime as dt
import json
import logging
import urllib.parse
import urllib.request
import ssl
import certifi
from zoneinfo import ZoneInfo
from pathlib import Path
from typing import Dict, List, Optional, Callable, Tuple

__version__ = "1.4.c"
VERSION_CODE = 10403  # Increment for each release (e.g. 1.4.a=10401, 1.4.b=10402, 1.4.c=10403)

logging.basicConfig(
    level=os.environ.get("SLS_LOG_LEVEL", "INFO").upper(),
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
)
APP_LOG = logging.getLogger("sls_robuster")

# GUI deps
try:
    import tkinter as tk
    from tkinter import ttk, messagebox, filedialog
    from tkinter import font as tkfont
    import customtkinter as ctk
except Exception as e:
    raise SystemExit(
        "customtkinter (and tkinter) are required. Install with: pip install customtkinter\n"
        f"Error: {e}"
    )

try:
    import pyperclip
except Exception:
    pyperclip = None  # type: ignore
try:
    from dateutil import rrule
except Exception:
    rrule = None  # type: ignore

# ------------------------
# Shared Config
# ------------------------
if getattr(sys, "frozen", False):
    BASE_DIR = os.path.abspath(getattr(sys, "_MEIPASS", os.path.dirname(sys.executable)))
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))


def _default_app_data_dir() -> Path:
    home = Path.home()
    if sys.platform == "darwin":
        return home / "Library" / "Application Support" / "SLSRobuster"
    if sys.platform.startswith("win"):
        base = Path(os.environ.get("APPDATA", home / "AppData" / "Roaming"))
        return base / "SLSRobuster"
    return home / ".config" / "SLSRobuster"


APP_DATA_PATH = _default_app_data_dir()
try:
    APP_DATA_PATH.mkdir(parents=True, exist_ok=True)
except Exception as exc:
    APP_LOG.warning("Could not create the app data directory.", exc_info=exc)
APP_DATA_DIR = str(APP_DATA_PATH)

APP_CONFIG_TEMPLATE_FILE = os.path.join(BASE_DIR, "app_config.json")
APP_CONFIG_FILE = os.path.join(APP_DATA_DIR, "app_config.json")


def _load_app_config() -> Dict[str, object]:
    defaults: Dict[str, object] = {
        "client_secret_filename": "",
        "default_channels": [],
        "update_metadata_url": "",
    }
    try:
        if not os.path.exists(APP_CONFIG_FILE) and os.path.exists(APP_CONFIG_TEMPLATE_FILE):
            shutil.copy(APP_CONFIG_TEMPLATE_FILE, APP_CONFIG_FILE)
    except Exception as exc:
        APP_LOG.warning("Failed to seed app_config.json template", exc_info=exc)

    try:
        if os.path.exists(APP_CONFIG_FILE):
            with open(APP_CONFIG_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
            if isinstance(data, dict):
                defaults.update(data)
    except Exception as exc:
        APP_LOG.warning("Failed to load app_config.json; using defaults", exc_info=exc)
    return defaults


def _default_client_secret_filename() -> str:
    configured = str(APP_CONFIG.get("client_secret_filename") or "").strip()
    if configured:
        return configured
    env_override = str(os.environ.get("SLS_CLIENT_SECRET_FILENAME", "")).strip()
    if env_override:
        return env_override
    matches = sorted(Path(BASE_DIR).glob("client_secret*.json"))
    return matches[0].name if matches else "client_secret.json"


def _normalize_channel_dict(c: dict) -> Dict[str, str]:
    name = str(c.get("name") or "").strip()
    channel_id = str(c.get("channel_id") or "").strip()
    return {
        "key": str(c.get("key") or (name.replace(" ", "_") if name else "")),
        "label": str(c.get("label") or (name.split()[-1] if name else "CH")),
        "name": name,
        "channel_id": channel_id,
        "playlist_id": str(c.get("playlist_id") or "").strip(),
        "thumbnail_template_file": str(c.get("thumbnail_template_file") or "").strip(),
    }


def _default_channels_from_config() -> List[Dict[str, str]]:
    raw = APP_CONFIG.get("default_channels")
    if not isinstance(raw, list):
        return []
    out: List[Dict[str, str]] = []
    for item in raw:
        if not isinstance(item, dict):
            continue
        out.append(_normalize_channel_dict(item))
    return out


APP_CONFIG = _load_app_config()
CLIENT_SECRET_FILENAME = _default_client_secret_filename()
CLIENT_SECRETS_BUNDLED = os.path.join(BASE_DIR, CLIENT_SECRET_FILENAME)
CLIENT_SECRETS_FILE = os.path.join(APP_DATA_DIR, CLIENT_SECRET_FILENAME)
try:
    if not os.path.exists(CLIENT_SECRETS_FILE) and os.path.exists(CLIENT_SECRETS_BUNDLED):
        shutil.copy(CLIENT_SECRETS_BUNDLED, CLIENT_SECRETS_FILE)
except Exception as exc:
    APP_LOG.warning("Could not seed a bundled config file into the app-data directory.", exc_info=exc)

SCOPES = ["https://www.googleapis.com/auth/youtube.force-ssl"]

# Store token files in app data directory by default
TOKEN_DIR = APP_DATA_DIR
GENERAL_SETTINGS_FILE = os.path.join(APP_DATA_DIR, "general_settings.json")
DREAMCLASS_SETTINGS_FILE = os.path.join(APP_DATA_DIR, "dreamclass_settings.json")
DREAMCLASS_API_BASE = "https://api.dreamclass.io/dreamclassapi/v1"

# Optional Chrome shortcut settings; override via env vars SLS_CHROME_PROFILE / SLS_CHROME_URL
CHROME_PROFILE_DIRECTORY = os.environ.get("SLS_CHROME_PROFILE", "Profile 1")
CHROME_TARGET_URL = os.environ.get("SLS_CHROME_URL", "https://slsdashboard.com/wp-admin/options-general.php?page=sls-live",)
UPDATE_METADATA_URL = str(APP_CONFIG.get("update_metadata_url") or os.environ.get("SLS_UPDATE_METADATA_URL", "")).strip()

# Channels registry stored in app data; optionally seeded from bundled template
CHANNELS_TEMPLATE_FILE = os.path.join(BASE_DIR, "channels.json")
CHANNELS_FILE = os.path.join(APP_DATA_DIR, "channels.json")
try:
    if not os.path.exists(CHANNELS_FILE) and os.path.exists(CHANNELS_TEMPLATE_FILE):
        shutil.copy(CHANNELS_TEMPLATE_FILE, CHANNELS_FILE)
except Exception as exc:
    APP_LOG.warning("Could not seed a bundled config file into the app-data directory.", exc_info=exc)
DEFAULT_CHANNELS: List[Dict[str, str]] = [
    c for c in _default_channels_from_config() if c.get("name") and c.get("channel_id")
]
if not DEFAULT_CHANNELS:
    try:
        if os.path.exists(CHANNELS_TEMPLATE_FILE):
            with open(CHANNELS_TEMPLATE_FILE, "r", encoding="utf-8") as f:
                ch_data = json.load(f)
            if isinstance(ch_data, dict):
                ch_data = ch_data.get("channels", [])
            if isinstance(ch_data, list):
                DEFAULT_CHANNELS = [
                    _normalize_channel_dict(c)
                    for c in ch_data
                    if isinstance(c, dict) and c.get("name") and c.get("channel_id")
                ]
    except Exception as exc:
        APP_LOG.warning("Failed to read default channels from channels.json", exc_info=exc)

def load_channels_config() -> List[Dict[str, str]]:
    """Load channels from channels.json; write defaults if missing. Returns a list of channel dicts."""
    try:
        if os.path.exists(CHANNELS_FILE):
            with open(CHANNELS_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
            if isinstance(data, dict) and isinstance(data.get("channels"), list):
                chans = [c for c in data["channels"] if isinstance(c, dict)]
            elif isinstance(data, list):
                chans = [c for c in data if isinstance(c, dict)]
            else:
                chans = []
            # minimal validation and fill-ins
            out: List[Dict[str, str]] = []
            for c in chans:
                name = str(c.get("name") or "").strip()
                ch_id = str(c.get("channel_id") or "").strip()
                pl_id = str(c.get("playlist_id") or "").strip()
                monitor = bool(c.get("monitor", True))
                if not name or not ch_id:
                    continue
                norm = _normalize_channel_dict(c)
                norm.update({
                    "name": name,
                    "channel_id": ch_id,
                    "playlist_id": pl_id,
                    "monitor": monitor,
                    "show_in_monitor": bool(c.get("show_in_monitor", True)),
                })
                out.append(norm)
            if out:
                return out
        # If file missing or invalid, write defaults for convenience
        os.makedirs(APP_DATA_DIR, exist_ok=True)
        with open(CHANNELS_FILE, "w", encoding="utf-8") as f:
            defaults = []
            for c in DEFAULT_CHANNELS:
                d = dict(c)
                d.setdefault("monitor", True)
                d.setdefault("show_in_monitor", True)
                defaults.append(d)
            json.dump({"channels": defaults}, f, indent=2)
        return [dict(c, monitor=True, show_in_monitor=True) for c in DEFAULT_CHANNELS]
    except Exception as exc:
        APP_LOG.warning("Failed to load channels config; using defaults", exc_info=exc)
        return [dict(c, monitor=True, show_in_monitor=True) for c in DEFAULT_CHANNELS]


def save_channels_config(channels: List[Dict[str, str]]) -> None:
    try:
        os.makedirs(APP_DATA_DIR, exist_ok=True)
        with open(CHANNELS_FILE, "w", encoding="utf-8") as f:
            cleaned = []
            for c in channels:
                norm = _normalize_channel_dict(c)
                norm.update({
                    "monitor": bool(c.get("monitor", True)),
                    "show_in_monitor": bool(c.get("show_in_monitor", True)),
                })
                cleaned.append(norm)
            json.dump({"channels": cleaned}, f, indent=2)
    except Exception as exc:
        APP_LOG.warning("Failed to write channels config to disk.", exc_info=exc)

LIVE_PRIVACY_STATUSES = {"unlisted", "public"}
TEXT_COLOR = "#EBEBEB"
BACKGROUND_COLOR = "#3F3F3F"
LIVE_INDICATOR_COLOR = "#2C5234"
CHIP_DEFAULT_COLOR = "#546263"
CHIP_NOT_COLOR = "#262C2C"
CHIP_ERROR_COLOR = "#D43E27"
CHIP_MONITOR_OFF_COLOR = "#849697"
ACTIVE_TAB_COLOR = "#13272D"
SURFACE_COLOR = "#2B2B2B"        # textboxes, listboxes, dark panels
SURFACE_MID_COLOR = "#2A2A2A"    # preview wrap, column headers
SURFACE_CARD_COLOR = "#252525"   # schedule slot cards
SURFACE_DEEP_COLOR = "#1F1F1F"   # "Other Locations" inset cards
INPUT_COLOR = "#4A4A4A"          # entries and comboboxes
INPUT_HOVER_COLOR = "#5A5A5A"    # combobox button hover
SEPARATOR_COLOR = "#3A3A3A"      # column dividers
SUBTEXT_COLOR = "#B8B8B8"        # description/hint labels
SECONDARY_TEXT_COLOR = "#D9D9D9" # secondary status labels
SCHEDULE_COL_LEFT = "Family Center"
SCHEDULE_COL_RIGHT = "Chapel"
PAD_SM = 4
PAD_MD = 8
PAD_LG = 12
PULSE_ANIM_COLORS = ["#546263", "#5D6E6F", "#667879", "#5D6E6F"]
PULSE_ANIM_INTERVAL_MS = 400
MAX_UPLOAD_PAGES = 10
AUTO_ARCHIVE_INTERVAL_MS = 30 * 60 * 1000  # 30 minutes
QUOTA_COSTS = {
    "videos.update": 50,
    "playlistitems.insert": 50,
}


def _version_key(version: str) -> Tuple[int, ...]:
    """Best-effort comparator for simple versions like 1.4.c / 1.4.1."""
    parts: List[int] = []
    for token in str(version or "").strip().lower().replace("-", ".").split("."):
        token = token.strip()
        if not token:
            continue
        if token.isdigit():
            parts.extend((1, int(token)))
            continue
        if token.isalpha():
            alpha_vals = [max(0, ord(ch) - 96) for ch in token]
            parts.extend((0, *alpha_vals))
            continue
        num = "".join(ch for ch in token if ch.isdigit())
        alpha = "".join(ch for ch in token if ch.isalpha())
        if num:
            parts.extend((1, int(num)))
        if alpha:
            parts.extend((0, *[max(0, ord(ch) - 96) for ch in alpha]))
    return tuple(parts or [0])


def _fetch_update_metadata(url: str) -> Dict[str, object]:
    target = str(url or "").strip()
    if not target:
        raise ValueError("Update metadata URL is not configured.")
    req = urllib.request.Request(
        target,
        headers={
            "User-Agent": f"SLSRobuster/{__version__}",
            "Cache-Control": "no-cache",
        },
    )
    context = ssl.create_default_context(cafile=certifi.where())
    with urllib.request.urlopen(req, timeout=10, context=context) as resp:
        payload = resp.read().decode("utf-8", errors="replace")
    data = json.loads(payload)
    if not isinstance(data, dict):
        raise ValueError("Update metadata must be a JSON object.")
    return data


def _is_update_available(metadata: Dict[str, object]) -> bool:
    remote_code_raw = metadata.get("version_code")
    if isinstance(remote_code_raw, (int, float)):
        try:
            return int(remote_code_raw) > int(VERSION_CODE)
        except Exception:
            pass
    remote_version = str(metadata.get("latest_version") or "").strip()
    if not remote_version:
        raise ValueError("Update metadata is missing 'latest_version'.")
    return _version_key(remote_version) > _version_key(__version__)


# ------------------------
# Lightweight YouTube helpers
# ------------------------
def _lazy_google_build(credentials):
    from googleapiclient.discovery import build
    return build("youtube", "v3", credentials=credentials)


def _lazy_oauth_flow():
    from google_auth_oauthlib.flow import InstalledAppFlow
    return InstalledAppFlow


def token_file_for_channel_id(channel_id: str) -> str:
    os.makedirs(TOKEN_DIR, exist_ok=True)
    return os.path.join(TOKEN_DIR, f"token-{channel_id}.pickle")


def _chrome_launch_command(profile: str, url: str) -> List[str]:
    """Build the platform-specific command to launch Chrome with a profile and URL."""
    profile = (profile or "").strip()
    url = (url or "").strip()
    if not url:
        raise ValueError("Target URL is required to launch Chrome.")

    if sys.platform == "darwin":
        cmd = ["open", "-na", "Google Chrome", "--args"]
        if profile:
            cmd.append(f"--profile-directory={profile}")
        cmd.append(url)
        return cmd

    chrome_executable: Optional[str] = None
    if sys.platform.startswith("win"):
        candidates = [
            os.path.join(os.environ[k], "Google", "Chrome", "Application", "chrome.exe")
            for k in ("PROGRAMFILES", "PROGRAMFILES(X86)", "LOCALAPPDATA")
            if os.environ.get(k)
        ]
        chrome_executable = next((p for p in candidates if os.path.exists(p)), None)
        if not chrome_executable:
            chrome_executable = shutil.which("chrome") or shutil.which("chrome.exe")
    else:
        for name in ("google-chrome", "chrome", "chromium", "chromium-browser"):
            found = shutil.which(name)
            if found:
                chrome_executable = found
                break

    if not chrome_executable:
        raise FileNotFoundError("Could not locate a Google Chrome executable on this system.")

    cmd = [chrome_executable]
    if profile:
        cmd.append(f"--profile-directory={profile}")
    cmd.append(url)
    return cmd


def authenticate_youtube(token_file: str) :
    """OAuth using the given token file path; returns a YouTube API client.
    Lazily imports google libs to avoid import errors in preview-only use.
    """
    creds = None
    if os.path.exists(token_file):
        try:
            with open(token_file, "rb") as f:
                creds = pickle.load(f)
        except Exception as exc:
            APP_LOG.warning("Failed to load token file; re-auth will be attempted", exc_info=exc)
            creds = None

    if creds and getattr(creds, "expired", False) and getattr(creds, "refresh_token", None):
        try:
            from google.auth.transport.requests import Request  # type: ignore
            creds.refresh(Request())
            with open(token_file, "wb") as f:
                pickle.dump(creds, f)
        except Exception as exc:
            APP_LOG.warning("Token refresh failed; falling back to OAuth flow", exc_info=exc)
            creds = None

    if not creds or not getattr(creds, "valid", False):
        Flow = _lazy_oauth_flow()
        flow = Flow.from_client_secrets_file(CLIENT_SECRETS_FILE, SCOPES)
        creds = flow.run_local_server(port=0)
        with open(token_file, "wb") as f:
            pickle.dump(creds, f)
    return _lazy_google_build(creds)


def get_authenticated_channel_info(youtube, logger: Optional[Callable[[str], None]] = None) -> Optional[Dict[str, str]]:
    resp = _yt_execute(youtube.channels().list(part="id,snippet", mine=True), logger, what="channels.list(mine)")
    items = resp.get("items", [])
    if not items:
        return None
    it = items[0]
    return {"id": it.get("id"), "title": (it.get("snippet") or {}).get("title", "")}


_QUOTA_TOTAL = 0
_QUOTA_LISTENERS: List[Callable[[int], None]] = []


def add_quota_listener(cb: Callable[[int], None]) -> None:
    _QUOTA_LISTENERS.append(cb)


def _inc_quota(units: int):
    global _QUOTA_TOTAL
    try:
        _QUOTA_TOTAL += int(units)
        for cb in list(_QUOTA_LISTENERS):
            try:
                cb(_QUOTA_TOTAL)
            except Exception as exc:
                APP_LOG.warning("Could not update or notify quota usage counters.", exc_info=exc)
    except Exception as exc:
        APP_LOG.warning("Could not update or notify quota usage counters.", exc_info=exc)


def _yt_execute(req, logger: Optional[Callable[[str], None]] = None, what: str = "request", retries: int = 3):
    """Execute a YouTube API request with basic retries/backoff and optional logging."""
    try:
        from googleapiclient.errors import HttpError  # type: ignore
    except Exception:
        HttpError = Exception  # fallback

    delay = 0.5
    last_err = None
    for attempt in range(1, retries + 1):
        try:
            res = req.execute()
            # Estimate quota cost and report
            w = (what or "").lower()
            units = next((v for k, v in QUOTA_COSTS.items() if k in w), 1)
            _inc_quota(units)
            return res
        except HttpError as e:  # type: ignore
            last_err = e
            status = getattr(e, "status_code", None) or getattr(getattr(e, "resp", None), "status", None)
            msg = ""
            try:
                content = getattr(e, "content", b"")
                msg = content.decode("utf-8", errors="ignore") if isinstance(content, (bytes, bytearray)) else str(e)
            except Exception:
                msg = str(e)
            if logger:
                logger(f"API error ({what}) [attempt {attempt}/{retries}]: {status or ''} {msg[:200]}")
            reason = msg.lower()
            # Retry on 5xx or quota/rate issues
            if (status and int(status) >= 500) or ("ratelimit" in reason or "quota" in reason or "userratelimit" in reason):
                time.sleep(delay)
                delay *= 2
                continue
            raise
        except Exception as e:
            last_err = e
            if logger:
                logger(f"Request failed ({what}) [attempt {attempt}/{retries}]: {e}")
            time.sleep(delay)
            delay *= 2
    # Final attempt or raise last error
    if last_err:
        raise last_err


def _is_auth_error(exc: Exception) -> bool:
    """Return True if an exception looks like an auth/token problem."""
    status = getattr(exc, "status_code", None) or getattr(getattr(exc, "resp", None), "status", None)
    if status in (401, 403):
        return True
    msg = str(exc).lower()
    keywords = [
        "invalid_grant",
        "invalid credentials",
        "invalidcredential",
        "unauthorized",
        "login required",
        "permissiondenied",
        "insufficientpermission",
        "token",
    ]
    return any(k in msg for k in keywords)


def get_channel_title_public(youtube, channel_id: str, logger: Optional[Callable[[str], None]] = None) -> str:
    resp = _yt_execute(youtube.channels().list(id=channel_id, part="snippet"), logger, what="channels.list(snippet)")
    items = resp.get("items", [])
    if items:
        return items[0]["snippet"]["title"]
    return channel_id


def find_active_or_testing_live_video_id(
    youtube,
    channel_id: str,
    ignore_privacy: bool = False,
    logger: Optional[Callable[[str], None]] = None,
) -> Optional[str]:
    """Backward-compatible wrapper that returns only video ID for active/testing stream."""
    details = find_latest_active_or_testing_stream(
        youtube,
        channel_id,
        ignore_privacy=ignore_privacy,
        logger=logger,
    )
    return (details or {}).get("video_id")


def find_latest_active_or_testing_stream(
    youtube,
    channel_id: str,
    ignore_privacy: bool = False,
    logger: Optional[Callable[[str], None]] = None,
) -> Optional[Dict[str, Optional[str]]]:
    """Find the freshest active/testing broadcast and return stream details."""
    resp = _yt_execute(
        youtube.liveBroadcasts().list(part="id,snippet,status", mine=True, maxResults=50),
        logger,
        what="liveBroadcasts.list",
    )
    items = resp.get("items", []) or []
    if not items:
        return None

    def _ok(b):
        st = (b.get("status") or {})
        life = (st.get("lifeCycleStatus") or "").lower()
        allowed = {"testing", "live", "teststarting", "livestarting"}
        privacy = st.get("privacyStatus")
        if ignore_privacy:
            priv_ok = True
        elif LIVE_PRIVACY_STATUSES:
            priv_ok = privacy in LIVE_PRIVACY_STATUSES
        else:
            priv_ok = True
        return priv_ok and life in allowed

    cands = [b for b in items if _ok(b)]
    if not cands:
        return None
    cands.sort(
        key=lambda b: (
            b.get("snippet", {}).get("actualStartTime") or "",
            b.get("snippet", {}).get("scheduledStartTime") or "",
        ),
        reverse=True,
    )
    top = cands[0]
    sn = top.get("snippet") or {}
    return {
        "video_id": top.get("id"),
        "stream_title": sn.get("title") or "",
        "actual_start_time": sn.get("actualStartTime"),
        "scheduled_start_time": sn.get("scheduledStartTime"),
    }


# ------------------------
# Archive: add livestream uploads to archive playlist
# ------------------------
def get_mine_channel_context(youtube, wanted_channel_id: str, logger: Optional[Callable[[str], None]] = None) -> Dict[str, str]:
    token = None
    while True:
        resp = _yt_execute(
            youtube.channels().list(mine=True, part="id,snippet,contentDetails", maxResults=50, pageToken=token),
            logger,
            what="channels.list(mine,contentDetails)",
        )
        for ch in resp.get("items", []):
            if ch["id"] == wanted_channel_id:
                return {
                    "id": ch["id"],
                    "title": ch["snippet"]["title"],
                    "uploads_playlist": ch["contentDetails"]["relatedPlaylists"]["uploads"],
                }
        token = resp.get("nextPageToken")
        if not token:
            break
    return {}


def get_playlist_video_ids(youtube, playlist_id, logger: Optional[Callable[[str], None]] = None) -> set:
    ids = set()
    token = None
    while True:
        resp = _yt_execute(
            youtube.playlistItems().list(playlistId=playlist_id, part="contentDetails", maxResults=50, pageToken=token),
            logger,
            what="playlistItems.list(contentDetails)",
        )
        for it in resp.get("items", []):
            ids.add(it["contentDetails"]["videoId"])
        token = resp.get("nextPageToken")
        if not token:
            break
    return ids


def _chunk(lst: List[str], n: int) -> List[List[str]]:
    return [lst[i:i + n] for i in range(0, len(lst), n)]


def get_recent_upload_ids_not_in_archive(
    youtube,
    uploads_playlist_id: str,
    already_in_archive: set,
    logger: Optional[Callable[[str], None]] = None,
) -> List[str]:
    candidates = []
    token = None
    pages = 0
    while True:
        resp = _yt_execute(
            youtube.playlistItems().list(playlistId=uploads_playlist_id, part="contentDetails", maxResults=50, pageToken=token),
            logger,
            what="playlistItems.list(uploads)",
        )
        for it in resp.get("items", []):
            vid = it["contentDetails"]["videoId"]
            if vid not in already_in_archive:
                candidates.append(vid)
        token = resp.get("nextPageToken")
        pages += 1
        if not token or pages >= MAX_UPLOAD_PAGES:
            break
    return candidates


def filter_to_livestreams(youtube, video_ids: List[str], logger: Optional[Callable[[str], None]] = None) -> List[Dict[str, str]]:
    results = []
    for batch in _chunk(video_ids, 50):
        resp = _yt_execute(
            youtube.videos().list(id=",".join(batch), part="snippet,status,liveStreamingDetails"),
            logger,
            what="videos.list(snippet,status,liveStreamingDetails)",
        )
        for v in resp.get("items", []):
            if "liveStreamingDetails" not in v:
                continue
            privacy = v.get("status", {}).get("privacyStatus")
            if LIVE_PRIVACY_STATUSES and privacy not in LIVE_PRIVACY_STATUSES:
                continue
            results.append({"videoId": v["id"], "title": v["snippet"]["title"]})
    return results


def get_recent_completed_livestreams_not_archived(
    youtube,
    profile: Dict[str, str],
    limit: int = 5,
    logger: Optional[Callable[[str], None]] = None,
) -> List[Dict[str, Optional[str]]]:
    """Return up to `limit` completed livestreams not already in the channel archive playlist."""
    ch_id = str(profile.get("channel_id") or "").strip()
    if not ch_id:
        return []

    archive_playlist_id = str(profile.get("playlist_id") or "").strip()
    archived_ids: set = set()
    if archive_playlist_id:
        try:
            archived_ids = get_playlist_video_ids(youtube, archive_playlist_id, logger=logger)
        except Exception as exc:
            if logger:
                logger(f"Failed to load archive playlist items: {exc}")

    mine = get_mine_channel_context(youtube, ch_id, logger=logger)
    uploads_playlist_id = (mine or {}).get("uploads_playlist")
    if not uploads_playlist_id:
        return []

    candidates = get_recent_upload_ids_not_in_archive(
        youtube, uploads_playlist_id, archived_ids, logger=logger
    )

    if not candidates:
        return []

    found: List[Dict[str, Optional[str]]] = []
    for batch in _chunk(candidates, 50):
        resp = _yt_execute(
            youtube.videos().list(id=",".join(batch), part="snippet,status,liveStreamingDetails"),
            logger,
            what="videos.list(snippet,status,liveStreamingDetails)",
        )
        for v in resp.get("items", []):
            live = v.get("liveStreamingDetails") or {}
            if not live:
                continue
            if not live.get("actualEndTime"):
                continue
            privacy = (v.get("status") or {}).get("privacyStatus")
            if LIVE_PRIVACY_STATUSES and privacy not in LIVE_PRIVACY_STATUSES:
                continue
            found.append(
                {
                    "video_id": v.get("id"),
                    "title": ((v.get("snippet") or {}).get("title") or "").strip(),
                    "actual_start_time": live.get("actualStartTime"),
                    "scheduled_start_time": live.get("scheduledStartTime"),
                }
            )

    found.sort(
        key=lambda x: (
            x.get("actual_start_time") or "",
            x.get("scheduled_start_time") or "",
        ),
        reverse=True,
    )
    return found[: max(limit, 1)]


def add_videos_to_playlist(youtube, playlist_id: str, videos: List[Dict[str, str]], logger: Callable[[str], None]) -> int:
    added = 0
    for vid in videos:
        try:
            _yt_execute(
                youtube.playlistItems().insert(
                    part="snippet",
                    body={
                        "snippet": {
                            "playlistId": playlist_id,
                            "resourceId": {"kind": "youtube#video", "videoId": vid["videoId"]},
                        }
                    },
                ),
                logger,
                what="playlistItems.insert",
            )
            logger(f"Added: {vid['title']} ({vid['videoId']})")
            added += 1
        except Exception as e:
            logger(f"Skip {vid['videoId']}: {e}")
    return added


def archive_video_to_channel_playlist(
    profile: Dict[str, str],
    youtube,
    video_id: str,
    title: str,
    dry_run: bool,
    logger: Callable[[str], None],
) -> int:
    """Archive a single video to the channel's configured playlist."""
    playlist_id = (profile.get("playlist_id") or "").strip()
    vid_clean = (video_id or "").strip()
    if not playlist_id:
        logger(f"No playlist configured for {profile['name']}; skipping archive.")
        return 0
    if not vid_clean:
        logger("No video id to archive; skipping.")
        return 0
    if dry_run:
        logger(f"[Preview] Would add {vid_clean or 'current live video'} to playlist {playlist_id}.")
        return 0
    if youtube is None:
        logger("No YouTube client available; cannot archive.")
        return 0
    try:
        existing_ids = get_playlist_video_ids(youtube, playlist_id, logger=logger)
    except Exception as exc:
        logger(f"Failed to load playlist items for duplicate check: {exc}")
        existing_ids = None
    if isinstance(existing_ids, set) and vid_clean in existing_ids:
        logger(f"Stream already archived for {profile['name']} ({vid_clean}); skipping duplicate.")
        return 0
    added = add_videos_to_playlist(
        youtube,
        playlist_id,
        [{"videoId": vid_clean, "title": title}],
        logger,
    )
    if added:
        logger(f"Archived {profile['name']} stream to playlist {playlist_id}.")
    return added


def archive_process_profile(profile: Dict[str, str], logger: Callable[[str], None]) -> int:
    ch_id = profile["channel_id"]
    dest_playlist = profile["playlist_id"]
    yt = authenticate_youtube(token_file_for_channel_id(ch_id))

    ch = get_mine_channel_context(yt, ch_id, logger=logger)
    if not ch:
        logger(
            f"Channel {profile['name']} not accessible with this Google login. "
            "Authenticate with the Google/Brand Account that owns the channel."
        )
        return 0

    logger(f"Fetching existing archive playlist items for {ch['title']}…")
    existing = get_playlist_video_ids(yt, dest_playlist, logger=logger)
    logger(f"Scanning uploads for {ch['title']}…")
    candidates = get_recent_upload_ids_not_in_archive(yt, ch["uploads_playlist"], existing, logger=logger)
    if not candidates:
        logger("No candidate uploads found.")
        return 0
    logger("Filtering to livestreams (live or completed)…")
    streams = filter_to_livestreams(yt, candidates, logger=logger)
    if not streams:
        logger("No new livestreams to add.")
        return 0
    logger(f"Adding {len(streams)} livestream(s) to the archive playlist…")
    return add_videos_to_playlist(yt, dest_playlist, streams, logger)


# ------------------------
# Retitle: update current live video title
# ------------------------
def process_channel(
    profile: Dict[str, str],
    new_title: str,
    dry_run: bool,
    logger: Callable[[str], None],
    existing_video_id: Optional[str] = None,
    youtube_client=None,
    thumbnail_path: Optional[str] = None,
) -> int:
    """Retitle the current live/testing broadcast for the given channel.
    Returns 1 if a title was updated; 0 otherwise.
    """
    ch_id = profile["channel_id"]
    token_path = token_file_for_channel_id(ch_id)

    if dry_run:
        logger(f"[Preview] Would authenticate with token {os.path.basename(token_path)}")
        logger(f"[Preview] Would retitle live broadcast on {profile['name']} to: {new_title}")
        archive_video_to_channel_playlist(
            profile,
            youtube_client,
            existing_video_id or "",
            new_title,
            dry_run=True,
            logger=logger,
        )
        if thumbnail_path:
            logger(f"[Preview] Would apply thumbnail '{os.path.basename(thumbnail_path)}' to {existing_video_id or 'current live video'}.")
        return 0

    yt = youtube_client or authenticate_youtube(token_path)
    mine = get_authenticated_channel_info(yt)
    if not mine or mine["id"] != ch_id:
        logger(
            f"Token is bound to {mine['title']} [{mine['id']}] not target [{ch_id}]" if mine else "No channel bound to token"
        )
        return 0

    vid = existing_video_id or find_active_or_testing_live_video_id(yt, ch_id, ignore_privacy=False, logger=logger)
    if not vid:
        logger("No active/testing broadcast found to retitle.")
        return 0

    # Fetch current snippet to preserve categoryId etc.
    vresp = _yt_execute(yt.videos().list(id=vid, part="snippet"), logger, what="videos.list(snippet)")
    items = vresp.get("items", [])
    if not items:
        logger("Could not fetch video snippet; skipping.")
        return 0
    sn = items[0]["snippet"]
    before = sn.get("title", "")
    sn["title"] = new_title
    try:
        _yt_execute(yt.videos().update(part="snippet", body={"id": vid, "snippet": sn}), logger, what="videos.update(snippet)")
        logger(f"Updated: '{before}' -> '{new_title}' (video {vid})")
        archive_video_to_channel_playlist(
            profile,
            yt,
            vid,
            new_title,
            dry_run=False,
            logger=logger,
        )
        if thumbnail_path:
            apply_video_thumbnail(
                yt,
                vid,
                thumbnail_path,
                logger=logger,
            )
        return 1
    except Exception as e:
        logger(f"Retitle failed for {vid}: {e}")
        return 0


# ------------------------
# Preset helpers for Retitle tab
# ------------------------
PRESETS_TEMPLATE_FILE = os.path.join(BASE_DIR, "retitle_presets.json")
PRESETS_FILE = os.path.join(APP_DATA_DIR, "retitle_presets.json")
THUMBNAIL_TEMPLATE_FILE = os.path.join(BASE_DIR, "Thumbnail_Default.jpg")
THUMBNAIL_PREVIEW_FILE = os.path.join(APP_DATA_DIR, "thumbnail_preview.jpg")
THUMBNAIL_OUTPUT_FILE = os.path.join(APP_DATA_DIR, "thumbnail_apply.jpg")
THUMBNAIL_SAFE_LEFT_X = int(APP_CONFIG.get("thumbnail_safe_left_x", 240))
THUMBNAIL_SAFE_RIGHT_X = int(APP_CONFIG.get("thumbnail_safe_right_x", 1040))
THUMBNAIL_FONT_FILE = str(
    APP_CONFIG.get("thumbnail_font_file") or os.environ.get("SLS_THUMBNAIL_FONT", "")
).strip()
THUMBNAIL_TITLE_FONT_FILE = str(
    APP_CONFIG.get("thumbnail_title_font_file")
    or os.environ.get("SLS_THUMBNAIL_TITLE_FONT", "Druk-Bold-Trial.otf")
).strip()
THUMBNAIL_SUBTITLE_FONT_FILE = str(
    APP_CONFIG.get("thumbnail_subtitle_font_file")
    or os.environ.get("SLS_THUMBNAIL_SUBTITLE_FONT", "Druk-Medium-Trial.otf")
).strip()


def _normalize_thumbnail_template_path(path: str) -> str:
    raw = str(path or "").strip()
    if not raw:
        return ""
    try:
        abs_path = os.path.abspath(raw)
        base_abs = os.path.abspath(BASE_DIR)
        common = os.path.commonpath([abs_path, base_abs])
        if common == base_abs:
            return os.path.relpath(abs_path, base_abs)
    except Exception:
        return raw
    return raw


def _resolve_thumbnail_template_path(path: str) -> str:
    raw = str(path or "").strip()
    if not raw:
        return THUMBNAIL_TEMPLATE_FILE
    candidates: List[str] = []
    if os.path.isabs(raw):
        candidates.append(raw)
    else:
        candidates.append(os.path.join(BASE_DIR, raw))
        candidates.append(os.path.join(APP_DATA_DIR, raw))
        candidates.append(raw)
    resolved = next((p for p in candidates if os.path.exists(p)), None)
    if resolved:
        return resolved
    return THUMBNAIL_TEMPLATE_FILE


try:
    if not os.path.exists(PRESETS_FILE) and os.path.exists(PRESETS_TEMPLATE_FILE):
        shutil.copy(PRESETS_TEMPLATE_FILE, PRESETS_FILE)
except Exception as exc:
    APP_LOG.warning("Could not seed the bundled presets file into the app-data directory.", exc_info=exc)


def _today_str() -> str:
    return dt.datetime.now().strftime("%m.%d.%y")


def load_presets() -> Dict[str, List[str]]:
    defaults = {
        "class_titles": [],
        "instructors": [],
        "dates": [_today_str()],
        "last_used": {"class_title": "", "instructor": "", "date": _today_str()},
    }
    try:
        if os.path.exists(PRESETS_FILE):
            with open(PRESETS_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
            for k, v in defaults.items():
                if k not in data:
                    data[k] = v
            t = _today_str()
            dates = [d for d in data.get("dates", []) if isinstance(d, str)]
            if t not in dates:
                dates.insert(0, t)
            data["dates"] = dates
            return _sort_presets_lists(data)
    except Exception as exc:
        APP_LOG.warning("Could not load presets from disk; using defaults.", exc_info=exc)
    return _sort_presets_lists(defaults)


def _deduplicated(
    items: List[str],
    prepend: Optional[str] = None,
    sort: bool = True,
    limit: Optional[int] = 100,
) -> List[str]:
    seq = [prepend] + list(items) if prepend else list(items)
    if sort:
        seq = sorted(seq, key=lambda s: str(s).casefold())
    seen = set()
    out: List[str] = []
    for x in seq:
        if not isinstance(x, str):
            continue
        x = x.strip()
        if not x or x in seen:
            continue
        seen.add(x)
        out.append(x)
    return out[:limit] if limit is not None else out


def _sort_presets_lists(data: Dict[str, List[str]]) -> Dict[str, List[str]]:
    result = dict(data)
    for key in ("class_titles", "dates", "instructors"):
        raw = result.get(key, [])
        cleaned = [str(x).strip() for x in raw if isinstance(x, str) and x.strip()]
        result[key] = _deduplicated(cleaned, sort=(key != "dates"))
    return result


def save_presets(data: Dict) -> None:
    try:
        with open(PRESETS_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
    except Exception as exc:
        APP_LOG.warning("Failed to write presets to disk.", exc_info=exc)


def apply_video_thumbnail(
    youtube,
    video_id: str,
    image_path: str,
    logger: Callable[[str], None],
) -> bool:
    """Upload a thumbnail image for a video using YouTube thumbnails.set."""
    vid = (video_id or "").strip()
    img = (image_path or "").strip()
    if not vid:
        logger("No video id for thumbnail; skipping thumbnail upload.")
        return False
    if not img or not os.path.exists(img):
        logger(f"Thumbnail image not found: {img or '(empty)'}")
        return False
    try:
        from googleapiclient.http import MediaFileUpload  # type: ignore
        media = MediaFileUpload(img, mimetype="image/jpeg", resumable=False)
        _yt_execute(
            youtube.thumbnails().set(videoId=vid, media_body=media),
            logger,
            what="thumbnails.set",
        )
        logger(f"Applied thumbnail to video {vid}.")
        return True
    except Exception as exc:
        logger(f"Thumbnail upload failed for {vid}: {exc}")
        return False


# ------------------------
# GUI App
# ------------------------
class SLSRobusterApp(ctk.CTk):
    def __init__(self) -> None:
        super().__init__()
        self._configure_window()
        self._init_state()
        self._build_header()
        self._build_status_bar()
        self._build_tabs()
        self._build_debug_console()

        self._presets_class_list: Optional[ctk.CTkScrollableFrame] = None
        self._presets_instr_list: Optional[ctk.CTkScrollableFrame] = None
        self._presets_status_label: Optional[ctk.CTkLabel] = None
        self._schedule_gap_overrides: Dict[str, dt.datetime] = {}
        self._schedule_gap_overrides_date: Optional[str] = None

        self._apply_dreamclass_enabled()

    def _configure_window(self) -> None:
        window_name = Path(sys.argv[0]).name if getattr(sys, "frozen", False) else Path(__file__).name
        self.title(window_name)
        self.geometry("700x980")
        self.resizable(False, False)
        self.configure(fg_color=BACKGROUND_COLOR)
        ctk.set_appearance_mode("system")
        ctk.set_default_color_theme("blue")

    def _init_state(self) -> None:
        self.preview_mode = tk.BooleanVar(value=False)  # default OFF per request
        self._init_fonts()
        self._init_animation_state()
        self._init_channel_state()

        self._dreamclass_settings = self._load_dreamclass_settings()
        self._dreamclass_enabled = bool(self._dreamclass_settings.get("enabled", True))
        self._general_status_label: Optional[ctk.CTkLabel] = None
        self._dreamclass_status_label: Optional[ctk.CTkLabel] = None
        self._general_settings = self._load_general_settings()
        try:
            self.preview_mode.set(self._general_settings.get("start_in_preview", False))
        except Exception as exc:
            APP_LOG.warning("Could not apply saved preview-mode startup preference.", exc_info=exc)
        if self._general_settings.get("stay_on_top"):
            try:
                self.attributes("-topmost", True)
            except Exception as exc:
                APP_LOG.warning("Could not apply 'stay on top' window setting.", exc_info=exc)
        if self._general_settings.get("auto_pulse_after_launch") and not self._pulse_active:
            self.after(500, self._pulse_toggle)

    def _init_fonts(self) -> None:
        self.FONT_HEADING = ctk.CTkFont(size=18, weight="bold")
        self.FONT_SUBHEADING = ctk.CTkFont(size=16, weight="bold")
        self.FONT_LABEL = ctk.CTkFont(size=14, weight="bold")
        self.FONT_SMALL = ctk.CTkFont(size=13, weight="bold")
        self.FONT_CARD = ctk.CTkFont(size=12, weight="bold")
        self.FONT_SCHEDULE = ctk.CTkFont(size=24, weight="bold")
        self.FONT_ROW_LABEL = ctk.CTkFont(size=15, weight="bold")
        self.FONT_BODY = ctk.CTkFont(size=14)
        self.FONT_CARD_ITALIC = ctk.CTkFont(size=12, weight="bold", slant="italic")
        self.FONT_CARD_SMALL = ctk.CTkFont(size=11, weight="bold")

    def _init_animation_state(self) -> None:
        self._header_icon_label: Optional[ctk.CTkLabel] = None
        self._header_icon_spin_job = None
        self._header_icon_spinning = False
        self._header_icon_angle = 0
        self._header_icon_pil_light = None
        self._header_icon_pil_dark = None
        self._upnext_frame = None
        self._upnext_pointer_label = None
        self._upnext_view_btn = None
        self._upnext_status_label = None
        self._upnext_fetch_in_progress = False
        self._upnext_auto_job = None
        self._schedule_spinner_canvas = None
        self._schedule_spinner_job = None
        self._schedule_spinner_angle = 0
        self._pulse_spinner_canvas = None
        self._pulse_spinner_job = None
        self._pulse_spinner_angle = 0
        self._pulse_spinner_holder = None
        self._pulse_job_id: Optional[str] = None
        self._pulse_active = False
        self._pulse_anim_colors = list(PULSE_ANIM_COLORS)
        self._pulse_anim_index = 0
        self._pulse_anim_job: Optional[str] = None
        self._pulse_button_anim_frames = ["Checking", "Checking.", "Checking..", "Checking..."]
        self._pulse_button_anim_index = 0
        self._pulse_button_anim_job: Optional[str] = None

    def _init_channel_state(self) -> None:
        self.channels: List[Dict[str, str]] = load_channels_config()
        for ch in self.channels:
            ch["monitor"] = False
            ch.setdefault("show_in_monitor", True)
            ch.setdefault("thumbnail_template_file", "")

        self._yt_clients: Dict[str, Dict[int, object]] = {}
        self._last_live_ids: Dict[str, Optional[str]] = {}
        self._indicators: Dict[str, ctk.CTkFrame] = {}
        self._indicator_labels: Dict[str, ctk.CTkLabel] = {}
        self._indicator_states: Dict[str, str] = {}
        self._label_by_id: Dict[str, str] = {}
        self._monitor_vars: Dict[str, tk.BooleanVar] = {}
        self._inds_frame = None
        self._pulse_controls = None
        self._settings_popup: Optional[ctk.CTkToplevel] = None
        self._whats_new_popup: Optional[ctk.CTkToplevel] = None
        self.chan_tree: Optional[ttk.Treeview] = None
        self._chan_monitor_scroll: Optional[ctk.CTkScrollableFrame] = None
        self._chan_monitor_vars: Dict[str, tk.BooleanVar] = {}
        self._settings_tabview: Optional[ctk.CTkTabview] = None
        self._thumb_subtitle_var = tk.StringVar(value="")
        self._thumb_preview_image = None
        self._thumb_preview_size = (320, 180)
        self._thumb_preview_label = None
        self._thumb_subtitle_entry = None
        self._thumb_template_info_label = None
        self._recovery_subtitle_entry = None
        self._presets_frames: Dict[str, Optional[ctk.CTkScrollableFrame]] = {
            "class_titles": None,
            "instructors": None,
        }
        self._presets_rows: Dict[str, Dict[str, ctk.CTkFrame]] = {
            "class_titles": {},
            "instructors": {},
        }
        self._presets_selected: Dict[str, Optional[str]] = {
            "class_titles": None,
            "instructors": None,
        }
        self._general_update_btn = None
        self._general_update_status_label = None

    def _make_combobox(self, parent, values, variable, **kwargs) -> ctk.CTkComboBox:
        return ctk.CTkComboBox(
            parent,
            values=values,
            variable=variable,
            text_color=TEXT_COLOR,
            fg_color=INPUT_COLOR,
            border_color=INPUT_COLOR,
            button_color=INPUT_COLOR,
            button_hover_color=INPUT_HOVER_COLOR,
            dropdown_fg_color=INPUT_COLOR,
            dropdown_text_color=TEXT_COLOR,
            **kwargs,
        )

    def _make_entry(self, parent, textvariable, **kwargs) -> ctk.CTkEntry:
        return ctk.CTkEntry(
            parent,
            textvariable=textvariable,
            fg_color=INPUT_COLOR,
            border_color=INPUT_COLOR,
            text_color=TEXT_COLOR,
            **kwargs,
        )

    def _build_header(self) -> None:
        hdr = ctk.CTkFrame(self, fg_color=BACKGROUND_COLOR)
        hdr.pack(fill=tk.X, padx=PAD_MD + 2, pady=(10, 6))
        header_top = ctk.CTkFrame(hdr, fg_color=BACKGROUND_COLOR)
        header_top.pack(fill=tk.X)
        title_frame = ctk.CTkFrame(header_top, fg_color=BACKGROUND_COLOR)
        title_frame.pack(side=tk.LEFT)

        self._header_icon = self._load_header_icon()
        if self._header_icon:
            self._header_icon_label = ctk.CTkLabel(title_frame, image=self._header_icon, text="")
            self._header_icon_label.pack(side=tk.LEFT, padx=(0, 6))
            try:
                self._header_icon_label.bind("<Button-1>", lambda _e: self._spin_header_icon())
            except Exception as exc:
                APP_LOG.warning("Could not finish building the header UI section.", exc_info=exc)
        self._settings_icon = self._load_settings_icon()
        self.btn_channels_popup = ctk.CTkButton(
            header_top,
            text="" if self._settings_icon else "Settings",
            image=self._settings_icon,
            width=48,
            height=36,
            fg_color=SEPARATOR_COLOR,
            hover_color=INPUT_COLOR,
            text_color=TEXT_COLOR,
            command=self._open_settings_popup,
        )
        self.btn_channels_popup.pack(side=tk.LEFT, padx=(8, 0))
        self.btn_whats_new = ctk.CTkButton(
            header_top,
            text="Updates",
            width=48,
            height=36,
            fg_color=SEPARATOR_COLOR,
            hover_color=INPUT_COLOR,
            text_color=TEXT_COLOR,
            command=self._open_whats_new_popup,
        )
        self.btn_whats_new.pack(side=tk.LEFT, padx=(8, 0))

        preview_wrap = ctk.CTkFrame(header_top, fg_color=SURFACE_MID_COLOR, corner_radius=12)
        preview_wrap.pack(side=tk.RIGHT, padx=0, pady=0, ipadx=2, ipady=0)
        ctk.CTkLabel(preview_wrap, text="Preview Mode", text_color=TEXT_COLOR).pack(side=tk.LEFT, padx=(8, 2), pady=2)
        self.preview_checkbox = ctk.CTkCheckBox(
            preview_wrap,
            text="",
            variable=self.preview_mode,
            onvalue=True,
            offvalue=False,
            width=18,
            height=18,
            fg_color="#22C55E",
            hover_color="#16A34A",
            border_color="#6B7280",
            border_width=2,
            checkmark_color="#0B4212",
        )
        self.preview_checkbox.pack(side=tk.LEFT, padx=PAD_SM, pady=2)

    def _build_status_bar(self) -> None:
        stat = ctk.CTkFrame(self, fg_color=BACKGROUND_COLOR)
        stat.pack(fill=tk.X, padx=PAD_MD + 2, pady=(0, 6))
        stat.grid_columnconfigure(0, weight=1)

        inds = ctk.CTkFrame(stat, fg_color=BACKGROUND_COLOR)
        inds.pack(side=tk.LEFT)
        self._inds_frame = inds
        for ch in self.channels:
            ch_id = ch["channel_id"]
            short = ch.get("label") or ch.get("key") or ch.get("name") or ch_id[:4]
            self._label_by_id[ch_id] = str(short)
            if not ch.get("show_in_monitor", True):
                continue
            # Chip container with rounded corners and colored background
            chip = ctk.CTkFrame(self._inds_frame, corner_radius=10, fg_color=CHIP_DEFAULT_COLOR)
            chip.pack(side=tk.LEFT, padx=(0, 6))
            chip.grid_columnconfigure(0, weight=0)
            chip.grid_columnconfigure(1, weight=0)
            var = tk.BooleanVar(value=False)
            self._monitor_vars[ch_id] = var
            ckb = ctk.CTkCheckBox(
                chip,
                text="",
                width=0,                 # remove reserved text area width
                checkbox_width=18,
                checkbox_height=18,
                variable=var,
                command=lambda cid=ch_id: self._on_monitor_toggled(cid),
            )
            ckb.grid(row=0, column=0, padx=(6, 2), pady=PAD_SM, sticky="w")
            lbl = ctk.CTkLabel(chip, text=str(short), text_color=TEXT_COLOR)
            lbl.grid(row=0, column=1, padx=(0, 8), pady=PAD_SM, sticky="w")
            self._indicators[ch_id] = chip
            self._indicator_labels[ch_id] = lbl
            self._set_indicator(ch_id, "off")

        pulse_controls = ctk.CTkFrame(stat, fg_color=BACKGROUND_COLOR)
        pulse_controls.pack(side=tk.RIGHT)
        self._pulse_controls = pulse_controls

        pulse_spinner_holder = ctk.CTkFrame(pulse_controls, fg_color=BACKGROUND_COLOR, width=18, height=18)
        pulse_spinner_holder.pack_propagate(False)
        self._pulse_spinner_holder = pulse_spinner_holder

        self.btn_pulse = ctk.CTkButton(
            pulse_controls,
            text="Check if Live",
            command=self._pulse_toggle,
            width=120,
            text_color=TEXT_COLOR,
        )
        self.btn_pulse.pack(side=tk.LEFT)
        self._update_pulse_button_enabled()

    def _build_tabs(self) -> None:
        self.tabs = ctk.CTkTabview(self, fg_color=BACKGROUND_COLOR)
        tabs = self.tabs
        tabs.pack(fill=tk.BOTH, expand=True, padx=PAD_MD + 2, pady=(0, 10))
        tab_retitle = tabs.add("Title + Archive")
        self._tab_name_find = "Stream Info"
        tab_find = tabs.add(self._tab_name_find)
        self._tab_name_schedule = "Schedule"
        self._tab_schedule: Optional[ctk.CTkFrame] = None
        tab_schedule = None
        if self._dreamclass_enabled:
            tab_schedule = tabs.add(self._tab_name_schedule)
            self._tab_schedule = tab_schedule
        for tab in (tab_retitle, tab_find, tab_schedule):
            try:
                tab.configure(fg_color=BACKGROUND_COLOR)
            except Exception as exc:
                APP_LOG.warning("Could not finish building the tabs UI section.", exc_info=exc)

        # Enlarge tab buttons to feel like regular buttons
        try:
            tabs._segmented_button.configure(
                height=40,
                font=self.FONT_SUBHEADING,
                text_color=TEXT_COLOR,
                selected_color=ACTIVE_TAB_COLOR,
                selected_hover_color=ACTIVE_TAB_COLOR,
                unselected_color=INPUT_COLOR,
                unselected_hover_color=INPUT_HOVER_COLOR,
                fg_color=SURFACE_COLOR,
            )
        except Exception as exc:
            APP_LOG.warning("Could not finish building the tabs UI section.", exc_info=exc)
        try:
            seg = tabs._segmented_button
            tabs.grid_columnconfigure(0, weight=1)
            manager = seg.winfo_manager()
            if manager == "grid":
                seg.grid_configure(sticky="ew")
            elif manager == "pack":
                seg.pack_configure(fill=tk.X, expand=True)
            # CTkTabview.bind is intentionally not implemented in current CustomTkinter versions.
            # Keep a one-shot sync and trigger additional syncs from explicit tab add/remove paths.
            self._sync_main_tab_width()
        except Exception as exc:
            APP_LOG.warning("Could not finish building the tabs UI section.", exc_info=exc)

        self._build_retitle_tab(tab_retitle)
        self._build_find_tab(tab_find)
        if tab_schedule is not None:
            self._build_schedule_tab(tab_schedule)

    def _build_debug_console(self) -> None:
        dbg = ctk.CTkFrame(self, fg_color=BACKGROUND_COLOR)
        dbg.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(
            dbg,
            text="Activity Log",
            font=self.FONT_LABEL,
            text_color=TEXT_COLOR,
        ).grid(row=0, column=0, sticky="w")
        ctk.CTkButton(
            dbg,
            text="Clear",
            width=80,
            command=self._debug_clear,
            text_color=TEXT_COLOR,
        ).grid(row=0, column=1, sticky="e")
        self.debug_text = ctk.CTkTextbox(
            dbg,
            height=160,
            text_color=TEXT_COLOR,
            fg_color=SURFACE_COLOR,
        )
        self.debug_text.grid(row=1, column=0, columnspan=2, sticky="nsew", pady=(6, 0))
        self._activity_frame = dbg
        self._set_activity_log_visible(self._general_settings.get("show_activity_log", True))

    # ---------------- Pulse helpers -----------------
    def _set_indicator(self, channel_id: str, state: str):
        # state in {"live", "not", "error", "no-token", "preview", "off"}
        self._indicator_states[channel_id] = state
        chip = self._indicators.get(channel_id)
        lbl = self._indicator_labels.get(channel_id)
        label = self._label_by_id.get(channel_id, channel_id[:4])
        if chip is None or lbl is None:
            return
        colors = {
            "live": (LIVE_INDICATOR_COLOR, label),
            "not": (CHIP_MONITOR_OFF_COLOR, label),
            "error": (CHIP_ERROR_COLOR, label),
            "no-token": ("#D97706", label),
            "preview": ("#64748B", label),
            "off": (CHIP_NOT_COLOR, label),
        }
        fg, text = colors.get(state, (CHIP_DEFAULT_COLOR, label))
        try:
            lbl.configure(text=text)
            if not (state == "not" and self._pulse_active):
                chip.configure(fg_color=fg)
        except Exception as exc:
            APP_LOG.warning("Could not apply indicator state.", exc_info=exc)

    def _sync_main_tab_width(self) -> None:
        tabs = getattr(self, "tabs", None)  # None until tab view is built during startup
        if tabs is None:
            return
        seg = getattr(tabs, "_segmented_button", None)
        if seg is None:
            return
        buttons = []
        try:
            btn_list = getattr(seg, "_buttons", None)
            if isinstance(btn_list, (list, tuple)) and btn_list:
                buttons = list(btn_list)
            else:
                btn_dict = getattr(seg, "_button_dict", None)
                if isinstance(btn_dict, dict) and btn_dict:
                    buttons = [b for b in btn_dict.values() if b is not None]
        except Exception as exc:
            APP_LOG.warning("Failed to inspect segmented tab buttons.", exc_info=exc)
            return
        tab_count = len(buttons)
        if tab_count <= 0:
            return
        try:
            seg_width = seg.winfo_width()
            tabs_width = tabs.winfo_width()
            total = max(seg_width or 0, tabs_width or 0)
        except Exception as exc:
            APP_LOG.warning("Failed to measure tab width.", exc_info=exc)
            return
        if total <= 1:
            try:
                self.after(50, self._sync_main_tab_width)
            except Exception as exc:
                APP_LOG.warning("Failed to schedule tab width sync retry.", exc_info=exc)
            return
        btn_width = max(160, int((total - 8) / tab_count))
        for btn in buttons:
            try:
                btn.configure(width=btn_width)
            except Exception as exc:
                APP_LOG.warning("Failed to apply width to tab button.", exc_info=exc)

    def _load_header_icon(self):
        path = Path(BASE_DIR) / "Icon.png"
        if not path.exists():
            return None
        try:
            from PIL import Image  # type: ignore

            img = Image.open(path)
            img_light = img.copy()
            img_dark = img.copy()
            img.close()
            self._header_icon_pil_light = img_light.copy()
            self._header_icon_pil_dark = img_dark.copy()
            self._header_icon = ctk.CTkImage(light_image=img_light, dark_image=img_dark, size=(48, 48))
            return self._header_icon
        except Exception:
            try:
                return tk.PhotoImage(file=str(path))
            except Exception:
                return None

    def _spin_header_icon(self) -> None:
        if self._header_icon_spinning:
            return
        if not self._header_icon_label or self._header_icon_pil_light is None:
            return
        self._header_icon_spinning = True
        self._header_icon_angle = 0
        self._spin_header_icon_tick()

    def _spin_header_icon_tick(self) -> None:
        label = self._header_icon_label
        if label is None or not bool(label.winfo_exists()):
            self._header_icon_spinning = False
            return
        angle = self._header_icon_angle
        try:
            from PIL import Image  # type: ignore

            light = self._header_icon_pil_light
            dark = self._header_icon_pil_dark or self._header_icon_pil_light
            if light is None:
                raise RuntimeError("Header icon missing")
            light_rot = light.rotate(-angle, resample=Image.BICUBIC, expand=False)
            dark_rot = dark.rotate(-angle, resample=Image.BICUBIC, expand=False)
            self._header_icon = ctk.CTkImage(light_image=light_rot, dark_image=dark_rot, size=(48, 48))
            label.configure(image=self._header_icon)
        except Exception:
            self._header_icon_spinning = False
            return
        self._header_icon_angle = (angle + 15) % 360
        if self._header_icon_angle == 0:
            try:
                if self._header_icon_pil_light is not None:
                    base_dark = self._header_icon_pil_dark or self._header_icon_pil_light
                    self._header_icon = ctk.CTkImage(
                        light_image=self._header_icon_pil_light,
                        dark_image=base_dark,
                        size=(48, 48),
                    )
                    label.configure(image=self._header_icon)
            except Exception as exc:
                APP_LOG.warning("Could not update spinner animation state (header icon tick).", exc_info=exc)
            self._header_icon_spinning = False
            return
        self._header_icon_spin_job = self.after(20, self._spin_header_icon_tick)

    def _load_settings_icon(self):
        path = Path(BASE_DIR) / "settings_cog.png"
        if not path.exists():
            return None
        try:
            from PIL import Image  # type: ignore

            img = Image.open(path)
            img_light = img.copy()
            img_dark = img.copy()
            img.close()
            return ctk.CTkImage(light_image=img_light, dark_image=img_dark, size=(24, 24))
        except Exception:
            try:
                return tk.PhotoImage(file=str(path))
            except Exception:
                return None

    def _open_settings_popup(self):
        existing = getattr(self, "_settings_popup", None)  # None until the Settings popup is opened
        if existing is not None and bool(existing.winfo_exists()):
            try:
                existing.deiconify()
                existing.lift()
                existing.focus()
            except Exception as exc:
                APP_LOG.warning("Could not open the settings popup window.", exc_info=exc)
            return
        popup = ctk.CTkToplevel(self)
        popup.title("Settings")
        popup.transient(self)
        popup.geometry("820x760")
        try:
            popup.configure(fg_color=BACKGROUND_COLOR)
        except Exception as exc:
            APP_LOG.warning("Could not open the settings popup window.", exc_info=exc)
        container = ctk.CTkFrame(popup, fg_color=BACKGROUND_COLOR)
        container.pack(fill=tk.BOTH, expand=True, padx=PAD_LG, pady=PAD_LG)

        tabview = ctk.CTkTabview(container, fg_color=BACKGROUND_COLOR)
        tabview.pack(fill=tk.BOTH, expand=True)
        self._settings_tabview = tabview
        tab_general = tabview.add("General Settings")
        tab_dreamclass = tabview.add("DreamClass")
        tab_presets = tabview.add("Presets")
        tab_channels = tabview.add("Channels")
        for t in (tab_channels, tab_general, tab_presets, tab_dreamclass):
            try:
                t.configure(fg_color=BACKGROUND_COLOR)
            except Exception as exc:
                APP_LOG.warning("Could not open the settings popup window.", exc_info=exc)
        try:
            tabview.set("General Settings")
        except Exception as exc:
            APP_LOG.warning("Could not open the settings popup window.", exc_info=exc)

        self._build_general_settings_tab(tab_general)
        self._build_dreamclass_tab(tab_dreamclass)
        self._build_presets_tab(tab_presets)
        self._build_channels_tab(tab_channels)

        def on_close():
            self._close_settings_popup()

        popup.protocol("WM_DELETE_WINDOW", on_close)
        popup.focus()
        self._settings_popup = popup

    def _open_whats_new_popup(self):
        existing = getattr(self, "_whats_new_popup", None)  # None until the Updates popup is opened
        if existing is not None and bool(existing.winfo_exists()):
            try:
                existing.deiconify()
                existing.lift()
                existing.focus()
            except Exception as exc:
                APP_LOG.warning("Could not open the whats new popup window.", exc_info=exc)
            return
        popup = ctk.CTkToplevel(self)
        popup.title("What's New")
        popup.transient(self)
        popup.geometry("560x520")
        try:
            popup.configure(fg_color=BACKGROUND_COLOR)
        except Exception as exc:
            APP_LOG.warning("Could not open the whats new popup window.", exc_info=exc)
        container = ctk.CTkFrame(popup, fg_color=BACKGROUND_COLOR)
        container.pack(fill=tk.BOTH, expand=True, padx=PAD_LG, pady=PAD_LG)
        container.grid_columnconfigure(0, weight=1)
        container.grid_rowconfigure(1, weight=1)

        ctk.CTkLabel(
            container,
            text="What's New",
            font=self.FONT_SUBHEADING,
            text_color=TEXT_COLOR,
        ).grid(row=0, column=0, sticky="w")

        text = ctk.CTkTextbox(
            container,
            text_color=TEXT_COLOR,
            fg_color=SURFACE_COLOR,
            wrap="word",
        )
        text.grid(row=1, column=0, sticky="nsew", pady=(8, 0))
        text.insert("1.0", self._get_latest_changelog_section())
        text.configure(state="disabled")

        def on_close():
            self._whats_new_popup = None
            try:
                popup.destroy()
            except Exception as exc:
                APP_LOG.warning("Could not persist or clean up popup state during window close.", exc_info=exc)

        popup.protocol("WM_DELETE_WINDOW", on_close)
        popup.focus()
        self._whats_new_popup = popup

    def _get_latest_changelog_section(self) -> str:
        try:
            with open(Path(BASE_DIR) / "CHANGELOG.txt", "r", encoding="utf-8") as f:
                lines = f.read().splitlines()
        except Exception:
            return "Unable to load CHANGELOG.txt."
        header_idx = None
        for i, line in enumerate(lines):
            if line.startswith("SLS_ROBUSTER v"):
                header_idx = i
                break
        if header_idx is None:
            return "No version entries found."
        end_idx = len(lines)
        for j in range(header_idx + 1, len(lines)):
            if lines[j].startswith("SLS_ROBUSTER v"):
                end_idx = j
                break
        return "\n".join(lines[header_idx:end_idx]).strip() or "No details available."

    def _close_settings_popup(self):
        popup = getattr(self, "_settings_popup", None)  # None if Settings popup was never opened
        if popup is not None:
            try:
                if popup.winfo_exists():
                    popup.destroy()
            except Exception as exc:
                APP_LOG.warning("Could not close the settings popup window cleanly.", exc_info=exc)
        self._settings_popup = None
        self.chan_tree = None
        self._chan_monitor_scroll = None
        self._chan_monitor_vars = {}
        self._settings_tabview = None

    def _on_monitor_toggled(self, channel_id: str):
        enabled = bool(self._monitor_vars.get(channel_id).get()) if channel_id in self._monitor_vars else True
        self.debug(f"Monitor {'ON' if enabled else 'OFF'} for {self._label_by_id.get(channel_id, channel_id[:4])}")
        if not enabled:
            # Immediately reflect as Off
            self._set_indicator(channel_id, "off")
        else:
            # Immediately reflect likely state without forcing a network call
            if self.preview_mode.get():
                self._set_indicator(channel_id, "preview")
            else:
                token_path = token_file_for_channel_id(channel_id)
                if not os.path.exists(token_path):
                    self._set_indicator(channel_id, "no-token")
                else:
                    vid = self._last_live_ids.get(channel_id)
                    self._set_indicator(channel_id, "live" if vid else "not")
        # Persist monitor state to channels.json
        try:
            for ch in self.channels:
                if ch.get("channel_id") == channel_id:
                    ch["monitor"] = enabled
                    break
            save_channels_config(self.channels)
        except Exception as exc:
            APP_LOG.warning("Non-critical operation failed while running _on_monitor_toggled.", exc_info=exc)
        self._update_pulse_button_enabled()
        if not self._has_monitored_channels() and self._pulse_active:
            self.debug("Pulse: stopped (no channels selected)")
            self._stop_pulse()
        try:
            self._update_thumbnail_preview()
        except Exception as exc:
            APP_LOG.warning("Failed to refresh thumbnail preview after monitor toggle.", exc_info=exc)

    def _rebuild_header_indicators(self):
        try:
            for w in list(self._inds_frame.winfo_children()):
                w.destroy()
        except Exception:
            return
        # Preserve previous monitor settings if present
        prev_monitor = getattr(self, "_monitor_vars", {}) or {}
        prev_states = getattr(self, "_indicator_states", {}) or {}
        self._indicators = {}
        self._indicator_labels = {}
        self._indicator_states = {}
        self._label_by_id = {}
        self._monitor_vars = {}
        for ch in self.channels:
            ch_id = ch["channel_id"]
            short = ch.get("label") or ch.get("key") or ch.get("name") or ch_id[:4]
            self._label_by_id[ch_id] = str(short)
            if not ch.get("show_in_monitor", True):
                continue
            chip = ctk.CTkFrame(self._inds_frame, corner_radius=10, fg_color=CHIP_DEFAULT_COLOR)
            chip.pack(side=tk.LEFT, padx=(0, 6))
            chip.grid_columnconfigure(0, weight=0)
            chip.grid_columnconfigure(1, weight=0)
            if ch_id in prev_monitor:
                try:
                    initial = bool(prev_monitor[ch_id].get())
                except Exception:
                    initial = False
            else:
                initial = bool(ch.get("monitor", False))
            var = tk.BooleanVar(value=initial)
            self._monitor_vars[ch_id] = var
            ckb = ctk.CTkCheckBox(
                chip,
                text="",
                width=0,
                checkbox_width=18,
                checkbox_height=18,
                variable=var,
                command=lambda cid=ch_id: self._on_monitor_toggled(cid),
            )
            ckb.grid(row=0, column=0, padx=(6, 2), pady=PAD_SM, sticky="w")
            lbl = ctk.CTkLabel(chip, text=str(short), text_color=TEXT_COLOR)
            lbl.grid(row=0, column=1, padx=(0, 8), pady=PAD_SM, sticky="w")
            self._indicators[ch_id] = chip
            self._indicator_labels[ch_id] = lbl
            self._set_indicator(ch_id, prev_states.get(ch_id, "off"))
        self._update_pulse_button_enabled()

    def _update_pulse_button_enabled(self) -> None:
        btn = getattr(self, "btn_pulse", None)  # None until the header controls are built
        if btn is None:
            return
        enabled = self._has_monitored_channels()
        try:
            btn.configure(state=tk.NORMAL if enabled else tk.DISABLED)
        except Exception as exc:
            APP_LOG.warning("Could not update pulse button enabled.", exc_info=exc)

    def _pulse_toggle(self):
        if not self._has_monitored_channels():
            if self._pulse_active:
                self._stop_pulse()
            self.debug("Pulse: no channels selected")
            self._update_pulse_button_enabled()
            return
        self._pulse_active = not self._pulse_active
        self.debug(f"Pulse: {'started' if self._pulse_active else 'stopped'}")
        if self._pulse_active:
            # kick off immediately
            self._start_pulse_animation()
            self._show_pulse_spinner()
            self._start_pulse_button_animation()
            self._schedule_pulse(0)
        else:
            # stop future ticks
            if self._pulse_job_id is not None:
                try:
                    self.after_cancel(self._pulse_job_id)
                except Exception as exc:
                    APP_LOG.warning("Could not update live-monitor pulse state (toggle).", exc_info=exc)
                self._pulse_job_id = None
            self._stop_pulse_animation()
            self._hide_pulse_spinner()
            self._stop_pulse_button_animation()

    def _stop_pulse(self):
        if not self._pulse_active:
            return
        self._pulse_active = False
        if self._pulse_job_id is not None:
            try:
                self.after_cancel(self._pulse_job_id)
            except Exception as exc:
                APP_LOG.warning("Could not stop pulse cleanly.", exc_info=exc)
            self._pulse_job_id = None
        self._stop_pulse_animation()
        self._hide_pulse_spinner()
        self._stop_pulse_button_animation()
        self.debug("Pulse: auto-stopped (2 channels live)")
        # After auto-stop, switch to Find tab and run a scan (monitored channels only)
        try:
            if hasattr(self, "tabs") and hasattr(self, "_tab_name_find"):
                self.tabs.set(self._tab_name_find)
            self._find_scan()
        except Exception as exc:
            APP_LOG.warning("Could not stop pulse cleanly.", exc_info=exc)

    def _start_pulse_button_animation(self) -> None:
        if not self._pulse_button_anim_frames:
            return
        if self._pulse_button_anim_job is not None:
            try:
                self.after_cancel(self._pulse_button_anim_job)
            except Exception as exc:
                APP_LOG.warning("Could not start pulse button animation.", exc_info=exc)
            self._pulse_button_anim_job = None
        self._pulse_button_anim_index = 0
        self._pulse_button_animation_tick()

    def _stop_pulse_button_animation(self) -> None:
        if self._pulse_button_anim_job is not None:
            try:
                self.after_cancel(self._pulse_button_anim_job)
            except Exception as exc:
                APP_LOG.warning("Could not stop pulse button animation cleanly.", exc_info=exc)
            self._pulse_button_anim_job = None
        self._set_pulse_button_idle_text()

    def _set_pulse_button_idle_text(self) -> None:
        try:
            self.btn_pulse.configure(text="Check if Live")
        except Exception as exc:
            APP_LOG.warning("Could not apply pulse button idle text state.", exc_info=exc)

    def _show_pulse_spinner(self) -> None:
        holder = getattr(self, "_pulse_spinner_holder", None)  # None until pulse spinner UI is created
        if holder is None:
            return
        try:
            if holder.winfo_manager() == "":
                holder.pack(side=tk.LEFT, padx=(0, 8))
        except Exception as exc:
            APP_LOG.warning("Could not show the pulse spinner UI element.", exc_info=exc)
        if getattr(self, "_pulse_spinner_canvas", None) is not None:  # Spinner canvas exists only while pulse spinner is shown
            return
        canvas = tk.Canvas(
            holder,
            width=18,
            height=18,
            bg=BACKGROUND_COLOR,
            highlightthickness=0,
        )
        canvas.pack(fill=tk.BOTH, expand=True)
        self._pulse_spinner_canvas = canvas
        self._pulse_spinner_angle = 0
        self._spin_pulse_spinner()

    def _spin_pulse_spinner(self) -> None:
        canvas = getattr(self, "_pulse_spinner_canvas", None)  # None until pulse spinner UI is created
        if canvas is None:
            return
        if not self._pulse_active:
            self._hide_pulse_spinner()
            return
        canvas.delete("all")
        pad = 2
        size = 18
        start = self._pulse_spinner_angle
        canvas.create_arc(
            pad,
            pad,
            size - pad,
            size - pad,
            start=start,
            extent=300,
            style="arc",
            outline=SECONDARY_TEXT_COLOR,
            width=3,
        )
        self._pulse_spinner_angle = (start + 15) % 360
        self._pulse_spinner_job = self.after(60, self._spin_pulse_spinner)

    def _hide_pulse_spinner(self) -> None:
        job = getattr(self, "_pulse_spinner_job", None)  # None when no pulse spinner animation is running
        if job is not None:
            try:
                self.after_cancel(job)
            except Exception as exc:
                APP_LOG.warning("Could not hide or tear down the pulse spinner UI element.", exc_info=exc)
        self._pulse_spinner_job = None
        canvas = getattr(self, "_pulse_spinner_canvas", None)  # None if pulse spinner was never shown
        if canvas is not None:
            try:
                canvas.destroy()
            except Exception as exc:
                APP_LOG.warning("Could not hide or tear down the pulse spinner UI element.", exc_info=exc)
        self._pulse_spinner_canvas = None
        holder = getattr(self, "_pulse_spinner_holder", None)  # None if pulse spinner container was never shown
        if holder is not None:
            try:
                if holder.winfo_manager():
                    holder.pack_forget()
            except Exception as exc:
                APP_LOG.warning("Could not hide or tear down the pulse spinner UI element.", exc_info=exc)

    def _pulse_button_animation_tick(self) -> None:
        if not self._pulse_active:
            self._pulse_button_anim_job = None
            self._set_pulse_button_idle_text()
            return
        if not self._pulse_button_anim_frames:
            self._set_pulse_button_idle_text()
            return
        text = self._pulse_button_anim_frames[self._pulse_button_anim_index % len(self._pulse_button_anim_frames)]
        self._pulse_button_anim_index = (self._pulse_button_anim_index + 1) % len(self._pulse_button_anim_frames)
        try:
            self.btn_pulse.configure(text=text)
        except Exception as exc:
            APP_LOG.warning("Could not update live-monitor pulse state (button animation tick).", exc_info=exc)
        self._pulse_button_anim_job = self.after(600, self._pulse_button_animation_tick)

    def _schedule_pulse(self, delay_ms: int = 5000):
        if not self._pulse_active:
            return
        # ensure single scheduled job
        if self._pulse_job_id is not None:
            try:
                self.after_cancel(self._pulse_job_id)
            except Exception as exc:
                APP_LOG.warning("Could not process the schedule step (pulse).", exc_info=exc)
            self._pulse_job_id = None
        self._pulse_job_id = self.after(delay_ms, self._pulse_tick)

    def _start_pulse_animation(self) -> None:
        if self._pulse_anim_job is not None:
            try:
                self.after_cancel(self._pulse_anim_job)
            except Exception as exc:
                APP_LOG.warning("Could not start pulse animation.", exc_info=exc)
            self._pulse_anim_job = None
        self._pulse_anim_index = 0
        self._pulse_animation_tick()

    def _stop_pulse_animation(self) -> None:
        if self._pulse_anim_job is not None:
            try:
                self.after_cancel(self._pulse_anim_job)
            except Exception as exc:
                APP_LOG.warning("Could not stop pulse animation cleanly.", exc_info=exc)
            self._pulse_anim_job = None
        for cid, state in list(self._indicator_states.items()):
            if state == "not":
                self._set_indicator(cid, "not")

    def _pulse_animation_tick(self) -> None:
        if not self._pulse_active:
            self._pulse_anim_job = None
            return
        if not self._pulse_anim_colors:
            return
        color = self._pulse_anim_colors[self._pulse_anim_index % len(self._pulse_anim_colors)]
        self._pulse_anim_index = (self._pulse_anim_index + 1) % len(self._pulse_anim_colors)
        for ch in self.channels:
            cid = ch["channel_id"]
            var = self._monitor_vars.get(cid)
            if var is None or not bool(var.get()):
                continue
            if self._indicator_states.get(cid) != "not":
                continue
            chip = self._indicators.get(cid)
            if chip is None:
                continue
            try:
                chip.configure(fg_color=color)
            except Exception as exc:
                APP_LOG.warning("Could not update live-monitor pulse state (animation tick).", exc_info=exc)
        self._pulse_anim_job = self.after(PULSE_ANIM_INTERVAL_MS, self._pulse_animation_tick)

    def _pulse_tick(self):
        if not self._pulse_active:
            return
        # Run checks in background, then reschedule
        def worker():
            try:
                # Mark a tick in the Activity Log with timestamp
                try:
                    self.debug(f"Pulse: tick at {dt.datetime.now().strftime('%H:%M:%S')}")
                except Exception:
                    self.debug("Pulse: tick")
                # Determine monitored channels up front
                monitored_ids = [c["channel_id"] for c in self._get_monitored_channels()]
                monitored_set = set(monitored_ids)
                total_monitored = len(monitored_ids)
                live_count = 0
                for ch in self.channels:
                    ch_id = ch["channel_id"]
                    label = self._label_by_id.get(ch_id, ch_id[:4])
                    # Skip if not monitored
                    if ch_id not in monitored_set:
                        self.after(0, lambda cid=ch_id: self._set_indicator(cid, "off"))
                        continue
                    if self.preview_mode.get():
                        self.after(0, lambda cid=ch_id: self._set_indicator(cid, "preview"))
                        self.debug(f"Pulse: {label} preview")
                        continue
                    token_path = token_file_for_channel_id(ch_id)
                    if not os.path.exists(token_path):
                        self.after(0, lambda cid=ch_id: self._set_indicator(cid, "no-token"))
                        self.debug(f"Pulse: {label} no token found")
                        continue
                    try:
                        # Reuse client and check live
                        yt = self._get_youtube(ch_id)
                        details = find_latest_active_or_testing_stream(yt, ch_id, ignore_privacy=False, logger=self.debug)
                        vid = (details or {}).get("video_id")
                        self._last_live_ids[ch_id] = vid
                        self.after(0, lambda cid=ch_id, live=bool(vid): self._set_indicator(cid, "live" if live else "not"))
                        self.debug(f"Pulse: {label} -> {'LIVE' if vid else '–'}")
                        if vid:
                            live_count += 1
                            if total_monitored > 0 and live_count >= total_monitored:
                                # Auto-stop when all monitored channels are live
                                self.after(0, self._stop_pulse)
                                break
                    except Exception as e:
                        if not self._handle_auth_error(ch_id, e, label):
                            self.after(0, lambda cid=ch_id: self._set_indicator(cid, "error"))
                            self.debug(f"Pulse: {label} error during check")
            finally:
                # schedule next tick in 5s
                self._schedule_pulse(5000)

        threading.Thread(target=worker, daemon=True).start()

    # ---------------- Debug console helpers -----------------
    def debug(self, msg: str):
        def _append():
            try:
                self.debug_text.insert("end", str(msg) + "\n")
                self.debug_text.see("end")
            except Exception as exc:
                APP_LOG.warning("Could not append text to the activity log widget.", exc_info=exc)
        try:
            self.after(0, _append)
        except Exception:
            _append()

    # ---------------- Background job runner -----------------
    def _run_job(
        self,
        name: str,
        work: Callable[[], Optional[object]],
        disable: Optional[List[object]] = None,
        status_widget: Optional[object] = None,
        start_text: Optional[str] = None,
        done_text: Optional[str] = None,
        on_done: Optional[Callable[[Optional[object]], None]] = None,
        on_error: Optional[Callable[[Exception], None]] = None,
    ) -> None:
        try:
            if disable:
                for w in disable:
                    try:
                        w.configure(state="disabled")
                    except Exception as exc:
                        APP_LOG.warning("Could not update UI state while running or scheduling a background job.", exc_info=exc)
            if status_widget and start_text:
                try:
                    status_widget.configure(text=start_text)
                except Exception as exc:
                    APP_LOG.warning("Could not update UI state while running or scheduling a background job.", exc_info=exc)
        except Exception as exc:
            APP_LOG.warning("Could not update UI state while running or scheduling a background job.", exc_info=exc)

        def runner():
            try:
                result = work()
                if on_done:
                    self.after(0, lambda: on_done(result))
                if status_widget and done_text:
                    self.after(0, lambda: status_widget.configure(text=done_text))
            except Exception as e:
                self.debug(f"{name} error: {e}")
                if on_error:
                    try:
                        self.after(0, lambda err=e: on_error(err))
                    except Exception as exc:
                        APP_LOG.warning("Could not restore UI state after the background worker finished.", exc_info=exc)
            finally:
                if disable:
                    self.after(0, lambda: [
                        (getattr(w, 'configure')(state='normal')) for w in disable if hasattr(w, 'configure')
                    ])

        threading.Thread(target=runner, daemon=True).start()

    def _debug_clear(self):
        try:
            self.debug_text.delete("1.0", "end")
        except Exception as exc:
            APP_LOG.warning("Non-critical operation failed while running _debug_clear.", exc_info=exc)

    # ---------------- YouTube client + live helper -----------------
    def _get_youtube(self, channel_id: str):
        """Return a YouTube client cached per channel + thread (avoid cross-thread reuse)."""
        tid = threading.get_ident()
        per_channel = self._yt_clients.get(channel_id)
        if per_channel is None:
            per_channel = {}
            self._yt_clients[channel_id] = per_channel
        yt = per_channel.get(tid)
        if yt is not None:
            return yt
        yt = authenticate_youtube(token_file_for_channel_id(channel_id))
        per_channel[tid] = yt
        return yt

    def _clear_channel_client(self, channel_id: str):
        """Drop cached clients + last live id for a channel (forces fresh auth next call)."""
        self._yt_clients.pop(channel_id, None)
        self._last_live_ids.pop(channel_id, None)

    def _handle_auth_error(self, channel_id: str, exc: Exception, label: Optional[str] = None) -> bool:
        """If error is auth-related, clear caches and surface a helpful log."""
        if not _is_auth_error(exc):
            return False
        self._clear_channel_client(channel_id)
        tag = label or self._label_by_id.get(channel_id, channel_id[:4])
        self.debug(f"Auth issue for {tag}: {exc}. Re-auth in Settings > Tokens.")
        try:
            self._set_indicator(channel_id, "error")
        except Exception as exc:
            APP_LOG.warning("Non-critical operation failed while running _handle_auth_error.", exc_info=exc)
        return True

    def _check_live(self, channel_id: str) -> Dict[str, Optional[str]]:
        """Return {live: bool, video_id: str|None} for a channel using cached client."""
        try:
            yt = self._get_youtube(channel_id)
            details = find_latest_active_or_testing_stream(yt, channel_id, ignore_privacy=False, logger=self.debug)
            vid = (details or {}).get("video_id")
            self._last_live_ids[channel_id] = vid
            return {"live": bool(vid), "video_id": vid}
        except Exception as e:
            if not self._handle_auth_error(channel_id, e):
                self.debug(f"check_live error for {channel_id}: {e}")
            return {"live": False, "video_id": None}

    @staticmethod
    def _format_stream_start_time(actual_start: Optional[str], scheduled_start: Optional[str]) -> str:
        raw = actual_start or scheduled_start
        if not raw:
            return "Unknown"
        try:
            norm = raw.replace("Z", "+00:00") if raw.endswith("Z") else raw
            dt_obj = dt.datetime.fromisoformat(norm)
            if dt_obj.tzinfo is None:
                dt_obj = dt_obj.replace(tzinfo=dt.timezone.utc)
            local_dt = dt_obj.astimezone()
            text = local_dt.strftime("%H:%M:%S")
        except Exception:
            text = raw
        if actual_start:
            return text
        return f"{text} (Scheduled Start Time)"

    def _rebuild_retitle_actions(self):
        try:
            for w in list(self._retitle_actions.winfo_children()):
                w.destroy()
        except Exception:
            return
        self._build_retitle_actions_content()

    def _retitle_channel_choices(self) -> Tuple[List[str], Dict[str, Dict[str, str]]]:
        choices: List[str] = []
        mapping: Dict[str, Dict[str, str]] = {}
        for ch in self.channels:
            base = str(ch.get("name") or ch.get("label") or ch.get("key") or ch.get("channel_id") or "").strip()
            if not base:
                continue
            display = base
            if display in mapping:
                display = f"{base} ({str(ch.get('channel_id') or '')[:8]})"
            choices.append(display)
            mapping[display] = ch
        return choices, mapping

    def _build_retitle_actions_content(self) -> None:
        actions = getattr(self, "_retitle_actions", None)  # None until the Retitle tab actions area is built
        if actions is None:
            return
        actions.grid_columnconfigure(0, weight=1)
        big_font = self.FONT_HEADING

        btn_monitored = ctk.CTkButton(
            actions,
            text="Apply Change",
            command=self._retitle_apply_monitored,
            font=big_font,
            text_color=TEXT_COLOR,
            fg_color="#2B6CB0",
            hover_color="#245A93",
            height=52,
            corner_radius=10,
        )
        actions.grid_rowconfigure(0, weight=0)
        btn_monitored.grid(row=0, column=0, sticky="ew", pady=(0, 6))
        info = ctk.CTkLabel(
            actions,
            text="This applies the selected title to all monitored livestreams.",
            text_color=TEXT_COLOR,
            justify="center",
        )
        actions.grid_rowconfigure(1, weight=0)
        info.grid(row=1, column=0, sticky="ew", pady=(0, 8))

        self.btn_open_recovery_popup = ctk.CTkButton(
            actions,
            text="Recover Missed Streams",
            command=self._open_recovery_popup,
            text_color=TEXT_COLOR,
            fg_color=SEPARATOR_COLOR,
            hover_color=INPUT_COLOR,
            height=40,
        )
        actions.grid_rowconfigure(2, weight=0)
        self.btn_open_recovery_popup.grid(row=2, column=0, sticky="ew", pady=(0, 8))

        ctk.CTkLabel(
            actions,
            text="Opens a popup for selecting one channel and retitle+archive of one completed stream at a time.",
            text_color=TEXT_COLOR,
            justify="center",
        ).grid(row=3, column=0, sticky="ew")
        actions.grid_rowconfigure(4, weight=1)

    def _open_recovery_popup(self) -> None:
        popup = getattr(self, "_retitle_recovery_popup", None)  # None until the recovery popup is opened
        if popup is not None:
            try:
                if popup.winfo_exists():
                    popup.lift()
                    popup.focus_force()
                    return
            except Exception as exc:
                APP_LOG.warning("Could not open the recovery popup window.", exc_info=exc)

        popup = ctk.CTkToplevel(self)
        self._retitle_recovery_popup = popup
        popup.title("Recover Missed Streams")
        popup.geometry("980x860")
        popup.minsize(920, 820)
        try:
            popup.transient(self)
        except Exception as exc:
            APP_LOG.warning("Could not open the recovery popup window.", exc_info=exc)
        popup.grid_columnconfigure(0, weight=1)
        popup.grid_rowconfigure(4, weight=1)

        ctk.CTkLabel(
            popup,
            text="Recover Missed Stream (Single Channel)",
            text_color=TEXT_COLOR,
            font=self.FONT_SUBHEADING,
        ).grid(row=0, column=0, sticky="w", padx=PAD_LG + 2, pady=(12, 8))

        # Popup-local title section (copied from main retitle area).
        title_section = ctk.CTkFrame(popup, fg_color=BACKGROUND_COLOR)
        title_section.grid(row=1, column=0, sticky="ew", padx=PAD_LG + 2, pady=(0, 8))
        title_section.grid_columnconfigure((0, 1, 2), weight=1)
        presets = load_presets()
        self._recovery_class_var = tk.StringVar(value=(self.cb_class.get() or "").strip())
        self._recovery_instr_var = tk.StringVar(value=(self.cb_instr.get() or "").strip())

        ctk.CTkLabel(title_section, text="Class Title", text_color=TEXT_COLOR).grid(row=0, column=0, sticky="w", padx=PAD_SM)
        self.cb_recovery_class = self._make_combobox(
            title_section,
            values=presets.get("class_titles", []),
            variable=self._recovery_class_var,
            command=lambda *_: self._update_thumbnail_preview(),
        )
        self.cb_recovery_class.grid(row=1, column=0, sticky="ew", padx=PAD_SM, pady=(0, 8))

        ctk.CTkLabel(title_section, text="Date (auto)", text_color=TEXT_COLOR).grid(row=0, column=1, sticky="w", padx=PAD_SM)
        self.lbl_recovery_date = ctk.CTkLabel(title_section, text=_today_str(), anchor="w", text_color=TEXT_COLOR)
        self.lbl_recovery_date.grid(row=1, column=1, sticky="ew", padx=PAD_SM, pady=(0, 8))

        ctk.CTkLabel(title_section, text="Instructor (optional)", text_color=TEXT_COLOR).grid(row=0, column=2, sticky="w", padx=PAD_SM)
        self.cb_recovery_instr = self._make_combobox(
            title_section,
            values=presets.get("instructors", []),
            variable=self._recovery_instr_var,
        )
        self.cb_recovery_instr.grid(row=1, column=2, sticky="ew", padx=PAD_SM, pady=(0, 8))

        preview_chip = ctk.CTkFrame(title_section, fg_color="#2C2C2C", corner_radius=18)
        preview_chip.grid(row=2, column=0, columnspan=3, sticky="ew", padx=PAD_SM + 2, pady=(4, 0))
        preview_chip.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(
            preview_chip,
            text="Title Preview",
            text_color="#AEC4D5",
            font=self.FONT_SMALL,
            anchor="w",
        ).grid(row=0, column=0, sticky="w", padx=PAD_LG + 2, pady=(10, 0))
        self.lbl_recovery_preview = ctk.CTkLabel(
            preview_chip,
            text="",
            font=self.FONT_HEADING,
            text_color=TEXT_COLOR,
            anchor="w",
            wraplength=760,
            justify="left",
        )
        self.lbl_recovery_preview.grid(row=1, column=0, sticky="ew", padx=PAD_LG + 2, pady=(6, 12))

        def update_recovery_preview(*_):
            self.lbl_recovery_date.configure(text=_today_str())
            title = self._compose_recovery_title()
            if title:
                self.lbl_recovery_preview.configure(text=title)
            else:
                self.lbl_recovery_preview.configure(text="Add a class title and instructor to build the livestream title.")

        self._recovery_class_var.trace_add("write", lambda *_: update_recovery_preview())
        self._recovery_instr_var.trace_add("write", lambda *_: update_recovery_preview())
        update_recovery_preview()

        controls = ctk.CTkFrame(popup, fg_color="transparent")
        controls.grid(row=2, column=0, sticky="ew", padx=PAD_LG + 2)
        controls.grid_columnconfigure(0, weight=1)
        controls.grid_columnconfigure(1, weight=0)
        choices, mapping = self._retitle_channel_choices()
        self._retitle_channel_choice_map = mapping
        if choices and not self._retitle_recovery_channel_var.get():
            self._retitle_recovery_channel_var.set(choices[0])
        if self._retitle_recovery_channel_var.get() not in mapping and choices:
            self._retitle_recovery_channel_var.set(choices[0])
        self.cb_recovery_channel = self._make_combobox(
            controls,
            values=choices or [""],
            variable=self._retitle_recovery_channel_var,
        )
        self.cb_recovery_channel.grid(row=0, column=0, sticky="ew", padx=(0, 8))

        self.btn_recovery_scan = ctk.CTkButton(
            controls,
            text="Load Last 5 Streams",
            command=self._retitle_load_recent_completed_streams,
            text_color=TEXT_COLOR,
            width=180,
        )
        self.btn_recovery_scan.grid(row=0, column=1, sticky="e")

        subtitle_row = ctk.CTkFrame(popup, fg_color="transparent")
        subtitle_row.grid(row=3, column=0, sticky="ew", padx=PAD_LG + 2, pady=(0, PAD_SM))
        ctk.CTkLabel(subtitle_row, text="Subtitle", text_color=TEXT_COLOR, anchor="w").pack(anchor="w")
        self._recovery_subtitle_entry = self._make_entry(
            subtitle_row,
            textvariable=self._thumb_subtitle_var,
            placeholder_text="(Insert Subtitle Here)",
        )
        self._recovery_subtitle_entry.pack(fill=tk.X, pady=(PAD_SM, 0))

        results_host = ctk.CTkFrame(popup, fg_color="#1E1E1E", corner_radius=10)
        results_host.grid(row=4, column=0, sticky="nsew", padx=PAD_LG + 2, pady=(10, 10))
        results_host.grid_columnconfigure(0, weight=1)
        results_host.grid_rowconfigure(0, weight=1)
        self._retitle_recovery_results_container = results_host

        self.btn_recovery_apply = ctk.CTkButton(
            popup,
            text="Retitle + Archive Selected Stream",
            command=self._retitle_apply_selected_completed_stream,
            text_color=TEXT_COLOR,
            height=42,
        )
        self.btn_recovery_apply.grid(row=5, column=0, sticky="ew", padx=PAD_LG + 2, pady=(0, 10))

        ctk.CTkLabel(
            popup,
            text="Loads last 5 completed livestreams that are not already archived. Popup is sized to show all rows.",
            text_color=TEXT_COLOR,
            justify="left",
        ).grid(row=6, column=0, sticky="w", padx=PAD_LG + 2, pady=(0, 12))

        popup.protocol("WM_DELETE_WINDOW", self._close_recovery_popup)
        self._retitle_render_recent_stream_rows([])

    def _close_recovery_popup(self) -> None:
        popup = getattr(self, "_retitle_recovery_popup", None)  # None if the recovery popup is already closed
        if popup is not None:
            try:
                if popup.winfo_exists():
                    popup.destroy()
            except Exception as exc:
                APP_LOG.warning("Could not close the recovery popup window cleanly.", exc_info=exc)
        self._retitle_recovery_popup = None
        self._retitle_recovery_results_container = None
        self.btn_recovery_scan = None
        self.btn_recovery_apply = None
        self.cb_recovery_channel = None
        self.cb_recovery_class = None
        self.cb_recovery_instr = None
        self._recovery_subtitle_entry = None
        self.lbl_recovery_date = None
        self.lbl_recovery_preview = None

    def _compose_recovery_title(self) -> str:
        try:
            c = (self.cb_recovery_class.get() or "").strip()
        except Exception:
            c = (self.cb_class.get() or "").strip()
        d = _today_str()
        try:
            if getattr(self, "lbl_recovery_date", None) is not None:  # Recovery popup label exists only while popup is open
                self.lbl_recovery_date.configure(text=d)
        except Exception as exc:
            APP_LOG.warning("Non-critical operation failed while running _compose_recovery_title.", exc_info=exc)
        try:
            i = (self.cb_recovery_instr.get() or "").strip()
        except Exception:
            i = (self.cb_instr.get() or "").strip()
        return " | ".join([x for x in (c, d, i) if x])

    def _retitle_get_selected_channel_profile(self) -> Optional[Dict[str, str]]:
        key = (self._retitle_recovery_channel_var.get() or "").strip()
        return (self._retitle_channel_choice_map or {}).get(key)

    @staticmethod
    def _format_stream_start_time_with_date(actual_start: Optional[str], scheduled_start: Optional[str]) -> str:
        raw = actual_start or scheduled_start
        if not raw:
            return "Unknown"
        try:
            norm = raw.replace("Z", "+00:00") if raw.endswith("Z") else raw
            dt_obj = dt.datetime.fromisoformat(norm)
            if dt_obj.tzinfo is None:
                dt_obj = dt_obj.replace(tzinfo=dt.timezone.utc)
            local_dt = dt_obj.astimezone()
            text = local_dt.strftime("%Y-%m-%d %H:%M:%S")
        except Exception:
            text = raw
        if actual_start:
            return text
        return f"{text} (Scheduled Start Time)"

    def _retitle_render_recent_stream_rows(self, streams: List[Dict[str, Optional[str]]]) -> None:
        parent = getattr(self, "_retitle_recovery_results_container", None)  # None until recovery results UI is built
        if parent is None:
            return
        for child in list(parent.winfo_children()):
            child.destroy()
        self._retitle_recovery_stream_map = {}
        self._retitle_recovery_selected_video_id.set("")
        if not streams:
            empty_wrap = ctk.CTkFrame(parent, fg_color="transparent")
            empty_wrap.pack(fill=tk.BOTH, expand=True, padx=PAD_MD, pady=PAD_MD)
            ctk.CTkLabel(
                empty_wrap,
                text="No unarchived completed livestreams found.",
                text_color=SECONDARY_TEXT_COLOR,
            ).pack(expand=True)
            return
        for stream in streams:
            vid = str(stream.get("video_id") or "").strip()
            if not vid:
                continue
            self._retitle_recovery_stream_map[vid] = stream
            row = ctk.CTkFrame(parent, fg_color=SEPARATOR_COLOR, corner_radius=8)
            row.pack(fill=tk.X, expand=True, padx=PAD_SM, pady=PAD_SM)
            top = ctk.CTkFrame(row, fg_color="transparent")
            top.pack(fill=tk.X, padx=PAD_MD, pady=(6, 2))
            rb = ctk.CTkRadioButton(
                top,
                text="",
                variable=self._retitle_recovery_selected_video_id,
                value=vid,
                width=14,
            )
            rb.pack(side=tk.LEFT, padx=(0, 6))
            ctk.CTkLabel(
                top,
                text=stream.get("title") or "(Untitled stream)",
                text_color=TEXT_COLOR,
                anchor="w",
                justify="left",
                wraplength=680,
            ).pack(side=tk.LEFT, fill=tk.X, expand=True)
            start_text = self._format_stream_start_time_with_date(
                stream.get("actual_start_time"),
                stream.get("scheduled_start_time"),
            )
            ctk.CTkLabel(
                row,
                text=f"Start: {start_text}  |  Video ID: {vid}",
                text_color="#BFC8D2",
                anchor="w",
                justify="left",
            ).pack(fill=tk.X, padx=34, pady=(0, 6))

    def _retitle_load_recent_completed_streams(self) -> None:
        profile = self._retitle_get_selected_channel_profile()
        if not profile:
            self.lbl_retitle_status.configure(text="Select a channel first")
            return
        self.lbl_retitle_status.configure(text="Loading recent completed streams…")

        def work():
            if self.preview_mode.get():
                demo = []
                now = dt.datetime.now().replace(microsecond=0)
                for idx in range(5):
                    demo_dt = now - dt.timedelta(days=idx, hours=idx)
                    demo.append(
                        {
                            "video_id": f"DEMO_RECOVER_{idx + 1}",
                            "title": f"Preview Completed Stream {idx + 1}",
                            "actual_start_time": demo_dt.isoformat(),
                            "scheduled_start_time": None,
                        }
                    )
                return demo, None
            ch_id = str(profile.get("channel_id") or "")
            yt = self._get_youtube(ch_id)
            mine = get_authenticated_channel_info(yt)
            if not mine or mine.get("id") != ch_id:
                return [], "token mismatch"
            streams = get_recent_completed_livestreams_not_archived(yt, profile, limit=5, logger=self.debug)
            return streams, None

        def on_done(result):
            streams, err = result
            if err:
                self._retitle_render_recent_stream_rows([])
                self.lbl_retitle_status.configure(text=f"Load failed: {err}")
                return
            self._retitle_render_recent_stream_rows(streams)
            if streams:
                self.lbl_retitle_status.configure(text=f"Loaded {len(streams)} stream(s)")
            else:
                self.lbl_retitle_status.configure(text="No unarchived completed livestreams found")

        def on_error(exc: Exception):
            if profile and self._handle_auth_error(profile.get("channel_id", ""), exc, profile.get("name")):
                self.lbl_retitle_status.configure(text="Auth error - re-auth token")
            else:
                self.lbl_retitle_status.configure(text=f"Load failed: {exc}")
            self._retitle_render_recent_stream_rows([])

        self._run_job(
            name="retitle-recent-scan",
            work=work,
            disable=[self.btn_recovery_scan, self.btn_recovery_apply],
            status_widget=self.lbl_retitle_status,
            start_text="Loading recent completed streams…",
            done_text=None,
            on_done=on_done,
            on_error=on_error,
        )

    def _retitle_apply_selected_completed_stream(self) -> None:
        profile = self._retitle_get_selected_channel_profile()
        if not profile:
            messagebox.showinfo("Select channel", "Choose a channel before applying a stream update.")
            return
        video_id = (self._retitle_recovery_selected_video_id.get() or "").strip()
        if not video_id:
            messagebox.showinfo("Select stream", "Pick one completed stream from the list first.")
            return
        title = self._compose_recovery_title()
        self._retitle_start(
            [profile],
            video_id_overrides={str(profile.get('channel_id') or ''): video_id},
            save_to_presets=False,
            title_override=title,
            thumbnail_title_override=(self.cb_recovery_class.get() or "").strip(),
            thumbnail_subtitle_override=(self._thumb_subtitle_var.get() or "").strip(),
        )

    def _retitle_apply_monitored(self) -> None:
        profiles = self._get_monitored_channels()
        if not profiles:
            messagebox.showinfo("No monitored channels", "Enable monitoring for at least one channel before applying the title.")
            try:
                self.lbl_retitle_status.configure(text="No monitored channels selected")
            except Exception as exc:
                APP_LOG.warning("Could not complete a Retitle workflow UI step (apply monitored).", exc_info=exc)
            return
        self._retitle_start(profiles)

    def _apply_channels_update(self, new_channels: List[Dict[str, str]]):
        new_ids = {c["channel_id"] for c in new_channels}
        # Replace channels
        self.channels = new_channels
        for ch in self.channels:
            ch.setdefault("show_in_monitor", True)
            ch.setdefault("thumbnail_template_file", "")
            # Reset monitor state to OFF when applying changes
            ch["monitor"] = False
        # Drop clients/cache for removed channels
        for cid in list(self._yt_clients.keys()):
            if cid not in new_ids:
                del self._yt_clients[cid]
        for cid in list(self._last_live_ids.keys()):
            if cid not in new_ids:
                del self._last_live_ids[cid]
        # Rebuild visuals
        self._rebuild_header_indicators()
        self._rebuild_retitle_actions()
        try:
            self._chan_refresh_monitor_checks()
            self._chan_check_tokens()
        except Exception as exc:
            APP_LOG.warning("Failed refreshing settings channel panels after apply", exc_info=exc)
        if not self._has_monitored_channels() and self._pulse_active:
            self._stop_pulse()

    def _copy_and_log(self, text: str, status_widget, what: str = "text"):
        """Copy text to clipboard, update status, and log to Activity Log."""
        if not text:
            return
        copied = False
        last_err: Optional[Exception] = None
        if pyperclip is not None:
            try:
                pyperclip.copy(text)
                copied = True
            except Exception as e:
                last_err = e
        if not copied:
            try:
                self.clipboard_clear()
                self.clipboard_append(text)
                copied = True
            except Exception as e:  # pragma: no cover - depends on OS clipboard
                last_err = e
        if copied:
            try:
                status_widget.configure(text=f"Copied {what}")
            except Exception as exc:
                APP_LOG.warning("Could not copy text to the clipboard or update copy status UI.", exc_info=exc)
            self.debug(f"Copied {what}: {text}")
        else:
            if status_widget:
                try:
                    status_widget.configure(text=f"Copy failed for {what}")
                except Exception as exc:
                    APP_LOG.warning("Could not copy text to the clipboard or update copy status UI.", exc_info=exc)
            self.debug(f"Copy failed for {what}: {last_err}")

    def _flash_button(self, button: ctk.CTkButton, flashes: int = 2, interval_ms: int = 240) -> None:
        if button is None:
            return
        try:
            base_fg = button.cget("fg_color")
            base_text = button.cget("text_color")
            base_hover = button.cget("hover_color")
        except Exception:
            return

        def apply_colors(fg, text, hover):
            kwargs = {}
            if fg is not None:
                kwargs["fg_color"] = fg
            if text is not None:
                kwargs["text_color"] = text
            if hover is not None:
                kwargs["hover_color"] = hover
            try:
                button.configure(**kwargs)
            except Exception as exc:
                APP_LOG.warning("Could not apply hover/selection colors to recovery list rows.", exc_info=exc)

        def normalize_color(val):
            if isinstance(val, (list, tuple)) and val:
                return val[0]
            return val

        def blend_with_white(hex_color: str, ratio: float) -> str:
            try:
                if not isinstance(hex_color, str) or not hex_color.startswith("#") or len(hex_color) != 7:
                    raise ValueError("Invalid color")
                r = int(hex_color[1:3], 16)
                g = int(hex_color[3:5], 16)
                b = int(hex_color[5:7], 16)
                r = int(r + (255 - r) * ratio)
                g = int(g + (255 - g) * ratio)
                b = int(b + (255 - b) * ratio)
                return f"#{r:02X}{g:02X}{b:02X}"
            except Exception:
                return "#636363"

        base_fg_norm = normalize_color(base_fg)
        flash_fg = blend_with_white(str(base_fg_norm), 0.25)
        flash_text = base_text
        steps = max(flashes, 1) * 2

        def tick(i: int):
            if not bool(button.winfo_exists()):
                return
            if i % 2 == 0:
                apply_colors(flash_fg, flash_text, flash_fg)
            else:
                apply_colors(base_fg, base_text, base_hover)
            if i + 1 < steps:
                self.after(interval_ms, lambda: tick(i + 1))

        tick(0)

    # ---------------- Retitle Tab -----------------
    def _build_retitle_tab(self, parent):
        try:
            parent.configure(fg_color=BACKGROUND_COLOR)
        except Exception as exc:
            APP_LOG.warning("Could not finish building the retitle tab UI section.", exc_info=exc)
        parent.grid_columnconfigure((0, 1, 2), weight=1)
        # Make the action area stretch to fill remaining vertical space
        parent.grid_rowconfigure(6, weight=1)

        presets = load_presets()
        self._save_class_to_presets_var = tk.BooleanVar(value=True)
        self._save_instr_to_presets_var = tk.BooleanVar(value=True)

        ctk.CTkLabel(parent, text="Class Title", text_color=TEXT_COLOR).grid(row=0, column=0, sticky="w", padx=PAD_SM)
        self._class_var = tk.StringVar()
        self.cb_class = self._make_combobox(
            parent,
            values=presets.get("class_titles", []),
            variable=self._class_var,
            command=lambda *_: self._update_thumbnail_preview(),
        )
        self.cb_class.grid(row=1, column=0, sticky="ew", padx=PAD_SM, pady=(0, 8))
        ctk.CTkCheckBox(
            parent,
            text="Save to dropdown",
            variable=self._save_class_to_presets_var,
            onvalue=True,
            offvalue=False,
            text_color=TEXT_COLOR,
        ).grid(row=2, column=0, sticky="w", padx=PAD_SM, pady=(0, 2))

        ctk.CTkLabel(parent, text="Date (auto)", text_color=TEXT_COLOR).grid(row=0, column=1, sticky="w", padx=PAD_SM)
        self.lbl_date = ctk.CTkLabel(parent, text=_today_str(), anchor="w", text_color=TEXT_COLOR)
        self.lbl_date.grid(row=1, column=1, sticky="ew", padx=PAD_SM, pady=(0, 8))

        ctk.CTkLabel(parent, text="Instructor (optional)", text_color=TEXT_COLOR).grid(row=0, column=2, sticky="w", padx=PAD_SM)
        self._instr_var = tk.StringVar()
        self.cb_instr = self._make_combobox(
            parent,
            values=presets.get("instructors", []),
            variable=self._instr_var,
            command=lambda *_: update_preview(),
        )
        self.cb_instr.grid(row=1, column=2, sticky="ew", padx=PAD_SM, pady=(0, 8))
        ctk.CTkCheckBox(
            parent,
            text="Save to dropdown",
            variable=self._save_instr_to_presets_var,
            onvalue=True,
            offvalue=False,
            text_color=TEXT_COLOR,
        ).grid(row=2, column=2, sticky="w", padx=PAD_SM, pady=(0, 2))

        # Preview chip
        preview_chip = ctk.CTkFrame(parent, fg_color="#2C2C2C", corner_radius=18)
        preview_chip.grid(row=3, column=0, columnspan=3, sticky="ew", padx=PAD_SM + 2, pady=(6, 0))
        preview_chip.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(
            preview_chip,
            text="Title Preview",
            text_color="#AEC4D5",
            font=self.FONT_SMALL,
            anchor="w",
        ).grid(row=0, column=0, sticky="w", padx=PAD_LG + 2, pady=(10, 0))
        self.lbl_preview = ctk.CTkLabel(
            preview_chip,
            text="",
            font=self.FONT_HEADING,
            text_color=TEXT_COLOR,
            anchor="w",
            wraplength=560,
            justify="left",
        )
        self.lbl_preview.grid(row=1, column=0, sticky="ew", padx=PAD_LG + 2, pady=(6, 12))

        def update_preview(*_):
            self.lbl_date.configure(text=_today_str())
            title = self._compose_title()
            if title:
                self.lbl_preview.configure(text=title)
            else:
                self.lbl_preview.configure(text="Add a class title and instructor to build the livestream title.")
            self._update_thumbnail_preview()

        self._class_var.trace_add("write", lambda *_: update_preview())
        self._instr_var.trace_add("write", lambda *_: update_preview())
        update_preview()

        # Thumbnail preview block (quick template builder)
        ctk.CTkLabel(parent, text="Thumbnail Preview", text_color=TEXT_COLOR).grid(
            row=4, column=0, columnspan=3, sticky="w", padx=PAD_LG + 2, pady=(PAD_SM, 0)
        )
        thumb_wrap = ctk.CTkFrame(parent, fg_color=SURFACE_MID_COLOR, corner_radius=18)
        thumb_wrap.grid(row=5, column=0, columnspan=3, sticky="nsew", padx=PAD_SM + 2, pady=(PAD_SM, 0))
        thumb_wrap.grid_columnconfigure(0, weight=3)
        thumb_wrap.grid_columnconfigure(1, weight=1)

        self._thumb_preview_label = ctk.CTkLabel(thumb_wrap, text="Loading preview…")
        self._thumb_preview_label.grid(row=0, column=0, sticky="w", padx=PAD_LG + 2, pady=PAD_LG)
        subtitle_col = ctk.CTkFrame(thumb_wrap, fg_color="transparent")
        subtitle_col.grid(row=0, column=1, sticky="n", padx=PAD_LG, pady=PAD_LG)
        ctk.CTkLabel(subtitle_col, text="Subtitle", text_color=TEXT_COLOR, anchor="w").pack(anchor="w")
        self._thumb_subtitle_entry = self._make_entry(
            subtitle_col,
            textvariable=self._thumb_subtitle_var,
            width=260,
            placeholder_text="(Insert Subtitle Here)",
        )
        self._thumb_subtitle_entry.pack(anchor="w", pady=(PAD_SM, 0))
        self._thumb_template_info_label = ctk.CTkLabel(
            subtitle_col,
            text="",
            text_color=SUBTEXT_COLOR,
            anchor="w",
            justify="left",
            wraplength=260,
        )
        self._thumb_template_info_label.pack(anchor="w", pady=(PAD_SM, 0))
        self._thumb_subtitle_var.trace_add("write", lambda *_: self._update_thumbnail_preview())
        self.after(0, self._update_thumbnail_preview)

        # Actions area: stacked, full-width, stretched vertically
        actions = ctk.CTkFrame(parent, fg_color=BACKGROUND_COLOR)
        actions.grid(row=6, column=0, columnspan=3, sticky="nsew", padx=PAD_MD, pady=(8, 8))
        self._retitle_recovery_channel_var = tk.StringVar(value="")
        self._retitle_recovery_selected_video_id = tk.StringVar(value="")
        self._retitle_recovery_stream_map: Dict[str, Dict[str, Optional[str]]] = {}
        self._retitle_channel_choice_map: Dict[str, Dict[str, str]] = {}
        self._retitle_recovery_popup = None
        self._retitle_recovery_results_container = None
        self.cb_recovery_channel = None
        self.btn_recovery_scan = None
        self.btn_recovery_apply = None

        self._retitle_actions = actions
        self._build_retitle_actions_content()

        # Status + log
        self.lbl_retitle_status = ctk.CTkLabel(parent, text="", text_color=TEXT_COLOR)
        self.lbl_retitle_status.grid(row=7, column=0, columnspan=3, sticky="w", pady=(6, 0))

    def _compose_title(self) -> str:
        c = (self.cb_class.get() or "").strip()
        d = _today_str()
        self.lbl_date.configure(text=d)
        i = (self.cb_instr.get() or "").strip()
        return " | ".join([x for x in (c, d, i) if x])

    @staticmethod
    def _thumbnail_wrapped_lines(
        draw,
        text: str,
        font,
        max_width: int,
        max_lines: int,
    ) -> List[str]:
        lines, _complete = SLSRobusterApp._thumbnail_wrapped_lines_with_fit(
            draw=draw,
            text=text,
            font=font,
            max_width=max_width,
            max_lines=max_lines,
        )
        return lines

    @staticmethod
    def _thumbnail_wrapped_lines_with_fit(
        draw,
        text: str,
        font,
        max_width: int,
        max_lines: int,
    ) -> Tuple[List[str], bool]:
        words = [w for w in str(text or "").split() if w]
        if not words:
            return [], True
        lines: List[str] = []
        idx = 0
        while idx < len(words) and len(lines) < max_lines:
            line = words[idx]
            idx += 1
            while idx < len(words):
                candidate = f"{line} {words[idx]}"
                bbox = draw.textbbox((0, 0), candidate, font=font)
                if (bbox[2] - bbox[0]) <= max_width:
                    line = candidate
                    idx += 1
                    continue
                break
            lines.append(line)
        complete = idx >= len(words)
        return lines, complete

    def _render_thumbnail_to_path(
        self,
        class_title: str,
        subtitle: str,
        out_path: str,
        template_file: Optional[str] = None,
    ) -> bool:
        title_text = (class_title or "").strip().upper()
        if not title_text:
            return False
        try:
            from PIL import Image, ImageDraw, ImageFont  # type: ignore
        except Exception as exc:
            APP_LOG.warning("Failed to import PIL for thumbnail rendering.", exc_info=exc)
            return False

        resolved_template = _resolve_thumbnail_template_path(template_file or "")
        if not os.path.exists(resolved_template):
            APP_LOG.warning("Failed to render thumbnail: template not found at %s", resolved_template)
            return False

        try:
            img = Image.open(resolved_template).convert("RGB")
            resample_lanczos = getattr(getattr(Image, "Resampling", Image), "LANCZOS", Image.BICUBIC)
            img = img.resize((1280, 720), resample_lanczos)
            draw = ImageDraw.Draw(img)
        except Exception as exc:
            APP_LOG.warning("Failed to load thumbnail template image.", exc_info=exc)
            return False

        def resolve_font_path(font_value: str, fallback_name: str) -> Optional[str]:
            candidate = (font_value or "").strip()
            if candidate and not os.path.isabs(candidate):
                candidate = os.path.join(BASE_DIR, candidate)
            fallback_candidates = [
                candidate,
                os.path.join(BASE_DIR, fallback_name),
                os.path.join(APP_DATA_DIR, fallback_name),
            ]
            if THUMBNAIL_FONT_FILE:
                shared = THUMBNAIL_FONT_FILE
                if not os.path.isabs(shared):
                    shared = os.path.join(BASE_DIR, shared)
                fallback_candidates.append(shared)
            fallback_candidates.extend(
                [
                    "/Library/Fonts/Impact.ttf",
                    "/Library/Fonts/Impact.otf",
                    "/Library/Fonts/Arial Bold.ttf",
                    "/System/Library/Fonts/Supplemental/Impact.ttf",
                    "/System/Library/Fonts/Supplemental/Impact.otf",
                    "/System/Library/Fonts/Supplemental/Arial Bold.ttf",
                ]
            )
            return next((p for p in fallback_candidates if p and os.path.exists(p)), None)

        title_font_path = resolve_font_path(THUMBNAIL_TITLE_FONT_FILE, "Druk-Bold-Trial.otf")
        subtitle_font_path = resolve_font_path(THUMBNAIL_SUBTITLE_FONT_FILE, "Druk-Medium-Trial.otf")

        def load_title_font(size: int):
            if title_font_path:
                try:
                    return ImageFont.truetype(title_font_path, size=size)
                except Exception:
                    pass
            return ImageFont.load_default()

        def load_subtitle_font(size: int):
            if subtitle_font_path:
                try:
                    return ImageFont.truetype(subtitle_font_path, size=size)
                except Exception:
                    pass
            return ImageFont.load_default()

        safe_left = max(0, min(1279, THUMBNAIL_SAFE_LEFT_X))
        safe_right = max(safe_left + 1, min(1280, THUMBNAIL_SAFE_RIGHT_X))
        title_box = (safe_left, 115, safe_right, 295)
        subtitle_box = (safe_left, 295, safe_right, 560)
        fill_color = (235, 235, 235)

        # Fit title using the requested rule:
        # 22 chars per line, max 2 lines, and shrink font until it visually fits.
        def split_title_lines_22(text: str) -> Tuple[List[str], bool]:
            tokens = [t for t in str(text or "").split() if t]
            if not tokens:
                return [], True
            lines: List[str] = []
            current = ""
            idx = 0
            while idx < len(tokens) and len(lines) < 2:
                tok = tokens[idx]
                if len(tok) > 22:
                    if current:
                        lines.append(current)
                        current = ""
                        if len(lines) >= 2:
                            break
                    lines.append(tok[:22])
                    tokens[idx] = tok[22:]
                    if len(lines) >= 2:
                        break
                    continue
                if not current:
                    current = tok
                    idx += 1
                    continue
                cand = f"{current} {tok}"
                if len(cand) <= 22:
                    current = cand
                    idx += 1
                else:
                    lines.append(current)
                    current = ""
            if current and len(lines) < 2:
                lines.append(current)
            return lines[:2], idx >= len(tokens)

        title_font = load_title_font(140)
        title_lines, title_complete = split_title_lines_22(title_text)
        if not title_lines:
            title_lines = [title_text[:22], title_text[22:44].strip()]
            title_lines = [ln for ln in title_lines if ln]
            title_complete = len(title_text) <= 44
        title_line_gap = 6
        title_w = title_box[2] - title_box[0]
        title_h = title_box[3] - title_box[1]
        for size in range(140, 28, -2):
            f = load_title_font(size)
            lines = title_lines
            if not lines or not title_complete:
                continue
            line_h = draw.textbbox((0, 0), "Hg", font=f)[3]
            total_h = len(lines) * line_h + max((len(lines) - 1), 0) * title_line_gap
            if total_h > title_h:
                continue
            widest = 0
            for line in lines:
                lb = draw.textbbox((0, 0), line, font=f)
                widest = max(widest, lb[2] - lb[0])
            if widest <= title_w:
                title_font = f
                title_lines = lines
                break

        if not title_complete:
            APP_LOG.warning("Failed to fit full class title within 2 lines of 22 characters: %s", title_text)

        title_line_h = draw.textbbox((0, 0), "Hg", font=title_font)[3]
        title_total_h = len(title_lines) * title_line_h + max((len(title_lines) - 1), 0) * title_line_gap
        title_start_y = title_box[1] + ((title_h - title_total_h) // 2)
        for idx, line in enumerate(title_lines):
            lb = draw.textbbox((0, 0), line, font=title_font)
            lw = lb[2] - lb[0]
            lx = title_box[0] + ((title_w - lw) // 2)
            ly = title_start_y + idx * (title_line_h + title_line_gap)
            draw.text((lx, ly), line, font=title_font, fill=fill_color)

        subtitle_text = (subtitle or "").strip().upper()
        if subtitle_text:
            sub_font = load_subtitle_font(96)
            lines: List[str] = []
            for size in range(96, 26, -2):
                f = load_subtitle_font(size)
                cand = self._thumbnail_wrapped_lines(
                    draw,
                    subtitle_text,
                    f,
                    max_width=(subtitle_box[2] - subtitle_box[0]),
                    max_lines=3,
                )
                if not cand:
                    continue
                line_h = draw.textbbox((0, 0), "Hg", font=f)[3]
                total_h = len(cand) * line_h + max((len(cand) - 1), 0) * 8
                if total_h <= (subtitle_box[3] - subtitle_box[1]):
                    sub_font = f
                    lines = cand
                    break
            if not lines:
                lines = self._thumbnail_wrapped_lines(
                    draw,
                    subtitle_text,
                    sub_font,
                    max_width=(subtitle_box[2] - subtitle_box[0]),
                    max_lines=3,
                )
            line_h = draw.textbbox((0, 0), "Hg", font=sub_font)[3]
            total_h = len(lines) * line_h + max((len(lines) - 1), 0) * 8
            sy = subtitle_box[1] + ((subtitle_box[3] - subtitle_box[1] - total_h) // 2)
            for idx, line in enumerate(lines):
                lb = draw.textbbox((0, 0), line, font=sub_font)
                lw = lb[2] - lb[0]
                lx = subtitle_box[0] + ((subtitle_box[2] - subtitle_box[0] - lw) // 2)
                ly = sy + idx * (line_h + 8)
                draw.text((lx, ly), line, font=sub_font, fill=fill_color)

        try:
            os.makedirs(APP_DATA_DIR, exist_ok=True)
            img.save(out_path, "JPEG", quality=92, optimize=True)
            return True
        except Exception as exc:
            APP_LOG.warning("Failed to save rendered thumbnail image.", exc_info=exc)
            return False

    def _update_thumbnail_preview(self) -> None:
        lbl = getattr(self, "_thumb_preview_label", None)  # None until the Retitle thumbnail preview UI is built
        if lbl is None:
            return
        try:
            if not bool(lbl.winfo_exists()):
                return
        except Exception:
            return
        class_title = (self.cb_class.get() or "").strip()
        subtitle = (self._thumb_subtitle_var.get() or "").strip() if self._thumb_subtitle_var is not None else ""
        monitored_profiles = self._get_monitored_channels()
        preview_template = THUMBNAIL_TEMPLATE_FILE
        template_lines: List[str] = []
        if monitored_profiles:
            preview_template = self._channel_template_file(monitored_profiles[0])
            for p in monitored_profiles:
                label = str(p.get("label") or p.get("name") or p.get("channel_id") or "Channel").strip()
                template_name = os.path.basename(self._channel_template_file(p))
                template_lines.append(f"{label}: {template_name}")
            if len(monitored_profiles) > 1:
                template_lines.insert(0, f"Preview uses: {template_lines[0]}")
                template_lines.insert(1, "Will apply by monitored channel:")
        else:
            template_lines.append(f"No monitored channels selected.")
            template_lines.append(f"Preview uses default: {os.path.basename(THUMBNAIL_TEMPLATE_FILE)}")

        info_lbl = getattr(self, "_thumb_template_info_label", None)  # None until the Retitle thumbnail preview UI is built
        if info_lbl is not None:
            try:
                info_text = "\n".join(template_lines[:6])
                if len(template_lines) > 6:
                    info_text += f"\n... +{len(template_lines) - 6} more"
                info_lbl.configure(text=info_text)
            except Exception:
                pass

        ok = self._render_thumbnail_to_path(
            class_title,
            subtitle,
            THUMBNAIL_PREVIEW_FILE,
            template_file=preview_template,
        )
        if not ok:
            try:
                lbl.configure(text="Thumbnail preview unavailable", image=None)
            except tk.TclError:
                # Recreate the label if Tk has already invalidated its image handle.
                parent = lbl.master
                try:
                    lbl.destroy()
                except Exception:
                    pass
                self._thumb_preview_label = ctk.CTkLabel(parent, text="Thumbnail preview unavailable")
                self._thumb_preview_label.grid(row=0, column=0, sticky="w", padx=PAD_LG + 2, pady=PAD_LG)
            return
        try:
            from PIL import Image  # type: ignore
            preview = Image.open(THUMBNAIL_PREVIEW_FILE).convert("RGB")
            new_img = ctk.CTkImage(light_image=preview, dark_image=preview, size=self._thumb_preview_size)
            # Important: configure first, then swap the strong reference.
            # Swapping first can garbage-collect the currently displayed Tk image too early.
            lbl.configure(image=new_img, text="")
            self._thumb_preview_image = new_img
        except Exception as exc:
            APP_LOG.warning("Failed to load thumbnail preview image for UI.", exc_info=exc)
            try:
                lbl.configure(text="Thumbnail preview unavailable", image=None)
            except Exception:
                pass

    def _channel_template_file(self, profile: Dict[str, str]) -> str:
        raw = str(profile.get("thumbnail_template_file") or "").strip()
        return _resolve_thumbnail_template_path(raw)

    def _retitle_start(
        self,
        profiles: List[Dict[str, str]],
        video_id_overrides: Optional[Dict[str, str]] = None,
        save_to_presets: bool = True,
        title_override: Optional[str] = None,
        thumbnail_title_override: Optional[str] = None,
        thumbnail_subtitle_override: Optional[str] = None,
    ):
        c = self.cb_class.get().strip()
        i = self.cb_instr.get().strip()
        title = title_override if isinstance(title_override, str) else self._compose_title()
        thumbnail_title = (thumbnail_title_override if isinstance(thumbnail_title_override, str) else c).strip()
        thumbnail_subtitle = (
            thumbnail_subtitle_override
            if isinstance(thumbnail_subtitle_override, str)
            else (self._thumb_subtitle_var.get() or "")
        ).strip()

        # Persist presets only for normal monitored live retitles.
        if save_to_presets:
            try:
                save_class = bool(self._save_class_to_presets_var.get())
            except Exception:
                save_class = True
            try:
                save_instr = bool(self._save_instr_to_presets_var.get())
            except Exception:
                save_instr = True
            data = load_presets()
            d = _today_str()
            if save_class:
                data["class_titles"] = _deduplicated(data.get("class_titles", []), prepend=c, sort=False)
            if save_instr:
                data["instructors"] = _deduplicated(data.get("instructors", []), prepend=i, sort=False)
            data["dates"] = _deduplicated(data.get("dates", []), prepend=d, sort=False)
            data["last_used"] = {"class_title": c, "date": d, "instructor": i}
            data = _sort_presets_lists(data)
            save_presets(data)
            self._apply_preset_dropdown_values(data)

        self.lbl_retitle_status.configure(text="Working…")
        self.debug(f"Retitle: starting for {', '.join([p['name'] for p in profiles])} -> '{title}' (preview={self.preview_mode.get()})")

        def logger(msg: str):
            self.lbl_retitle_status.configure(text=msg)
            self.debug(msg)

        def work():
            total = 0
            errs: List[str] = []
            for p in profiles:
                try:
                    ch_id = p["channel_id"]
                    thumbnail_path = None
                    if thumbnail_title:
                        safe_cid = "".join(ch for ch in str(ch_id) if ch.isalnum() or ch in ("-", "_")) or "channel"
                        per_channel_thumb = os.path.join(APP_DATA_DIR, f"thumbnail_apply_{safe_cid}.jpg")
                        if self._render_thumbnail_to_path(
                            thumbnail_title,
                            thumbnail_subtitle,
                            per_channel_thumb,
                            template_file=self._channel_template_file(p),
                        ):
                            thumbnail_path = per_channel_thumb
                    existing_vid = (video_id_overrides or {}).get(ch_id) or self._last_live_ids.get(ch_id)
                    yt_client = None if self.preview_mode.get() else self._get_youtube(ch_id)
                    total += process_channel(
                        p,
                        title,
                        dry_run=self.preview_mode.get(),
                        logger=logger,
                        existing_video_id=existing_vid,
                        youtube_client=yt_client,
                        thumbnail_path=thumbnail_path,
                    )
                except Exception as e:
                    errs.append(f"{p['name']}: {e}")
            return total, errs

        def on_done(result):
            total, errs = result
            msg = f"Done. Titles changed: {total}"
            if errs:
                msg += f"  Errors: {' | '.join(errs)}"
            self.lbl_retitle_status.configure(text=msg)
            self.debug(msg)

        self._run_job(
            name="retitle",
            work=work,
            disable=None,
            status_widget=self.lbl_retitle_status,
            start_text="Working…",
            done_text=None,
            on_done=on_done,
        )

    def _build_upnext_header(self, parent):
        upnext = ctk.CTkFrame(parent, fg_color=BACKGROUND_COLOR)
        upnext.pack(fill=tk.X, pady=(8, 0))
        self._upnext_frame = upnext
        upnext.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(
            upnext,
            text="Next Up",
            font=self.FONT_LABEL,
            text_color=TEXT_COLOR,
        ).grid(row=0, column=0, sticky="w")
        self._upnext_pointer_label = ctk.CTkLabel(upnext, text="Loading…", text_color=SECONDARY_TEXT_COLOR)
        self._upnext_pointer_label.grid(row=0, column=1, sticky="w", padx=(8, 0))
        self._upnext_view_btn = ctk.CTkButton(
            upnext,
            text="View Schedule",
            command=lambda: self.tabs.set(self._tab_name_schedule),
            text_color=TEXT_COLOR,
            width=130,
        )
        self._upnext_view_btn.grid(row=0, column=2, sticky="e")
        self._upnext_fetch_in_progress = False
        self._upnext_status_label: Optional[ctk.CTkLabel] = None
        self._schedule_items: List[Dict[str, object]] = []
        self._schedule_period_label: str = ""
        self._schedule_window: Tuple[Optional[dt.datetime], Optional[dt.datetime]] = (None, None)

        if not self._dreamclass_enabled:
            self._upnext_pointer_label.configure(text="DreamClass disabled.")
            try:
                self._upnext_view_btn.grid_remove()
            except Exception as exc:
                APP_LOG.warning("Could not finish building the upnext header UI section.", exc_info=exc)
            return
        if not self._dreamclass_is_configured():
            self._upnext_pointer_label.configure(text="DreamClass not configured.")

    def _schedule_upnext_refresh(self):
        # Auto-refresh removed; manual refresh only via button.
        self._upnext_auto_job = None

    def _dreamclass_is_configured(self) -> bool:
        settings = self._dreamclass_settings or {}
        return all(settings.get(k) for k in ("api_key", "school_code", "tenant"))

    def _dreamclass_headers(self) -> Dict[str, str]:
        settings = self._dreamclass_settings
        return {
            "dreamclass-secret-key": str(settings.get("api_key", "")),
            "schoolCode": str(settings.get("school_code", "")),
            "tenant": str(settings.get("tenant", "")),
        }

    def _dreamclass_instructor(self, item: Dict[str, object]) -> str:
        candidates = [
            item.get("instructor"),
            item.get("instructorName"),
            item.get("teacher"),
            item.get("teacherName"),
        ]
        for key in ("instructor", "teacher", "lecturer"):
            val = item.get(key)
            if isinstance(val, dict):
                candidates.append(val.get("name"))
        for k, v in item.items():
            if isinstance(k, str) and ("instructor" in k.lower() or "teacher" in k.lower()):
                if isinstance(v, dict):
                    candidates.append(v.get("name"))
                else:
                    candidates.append(v)
        for cand in candidates:
            if isinstance(cand, str) and cand.strip():
                return cand.strip()
        return ""

    def _remove_schedule_tab(self) -> None:
        tab = getattr(self, "_tab_schedule", None)  # None when the Schedule tab is disabled/not created
        if tab is not None:
            try:
                self.tabs.delete(self._tab_name_schedule)
            except Exception:
                try:
                    tab.destroy()
                except Exception as exc:
                    APP_LOG.warning("Could not remove schedule tab UI state.", exc_info=exc)
        self._tab_schedule = None
        self._schedule_list = None
        self.after(0, self._sync_main_tab_width)

    def _ensure_schedule_tab(self) -> None:
        if not self._dreamclass_enabled:
            self._remove_schedule_tab()
            return
        if self._tab_schedule is None:
            tab_schedule = self.tabs.add(self._tab_name_schedule)
            self._tab_schedule = tab_schedule
            try:
                tab_schedule.configure(fg_color=BACKGROUND_COLOR)
            except Exception as exc:
                APP_LOG.warning("Could not ensure schedule tab is available.", exc_info=exc)
            self._build_schedule_tab(tab_schedule)
        self.after(0, self._sync_main_tab_width)

    def _clear_schedule_ui(self) -> None:
        self._schedule_items = []
        self._schedule_window = (None, None)
        self._schedule_period_label = ""
        if getattr(self, "_schedule_status_label", None) is not None:  # Schedule status label exists only when Schedule tab is built
            try:
                self._schedule_status_label.configure(text="DreamClass disabled.")
            except Exception as exc:
                APP_LOG.warning("Could not clear schedule ui.", exc_info=exc)
        frame = getattr(self, "_schedule_list", None)  # None until the Schedule tab UI is built
        if frame is not None:
            for child in list(frame.winfo_children()):
                try:
                    child.destroy()
                except Exception as exc:
                    APP_LOG.warning("Could not clear schedule ui.", exc_info=exc)
            ctk.CTkLabel(
                frame,
                text="DreamClass schedule is disabled.",
                text_color=SECONDARY_TEXT_COLOR,
            ).pack(anchor="w", padx=PAD_SM + 2, pady=PAD_SM + 2)

    def _apply_dreamclass_enabled(self) -> None:
        enabled = bool(self._dreamclass_enabled)
        btn = getattr(self, "_upnext_view_btn", None)  # None when the Up Next header UI has not been built
        if btn is not None:
            try:
                if enabled:
                    btn.grid()
                else:
                    btn.grid_remove()
            except Exception as exc:
                APP_LOG.warning("Non-critical operation failed while running _apply_dreamclass_enabled.", exc_info=exc)
        frame = getattr(self, "_upnext_frame", None)  # None when the Up Next header UI has not been built
        if frame is not None:
            try:
                if enabled:
                    if not frame.winfo_manager():
                        frame.pack(fill=tk.X, pady=(8, 0))
                else:
                    frame.pack_forget()
            except Exception as exc:
                APP_LOG.warning("Non-critical operation failed while running _apply_dreamclass_enabled.", exc_info=exc)
        if enabled:
            self._ensure_schedule_tab()
            # Manual refresh only via the Refresh button; no auto-cycle.
            self._refresh_upnext()
        else:
            self._clear_schedule_ui()
            self._remove_schedule_tab()
            lbl = getattr(self, "_upnext_pointer_label", None)  # None when Up Next header UI is not built
            if lbl is not None:
                try:
                    lbl.configure(text="DreamClass disabled.")
                except Exception as exc:
                    APP_LOG.warning("Non-critical operation failed while running _apply_dreamclass_enabled.", exc_info=exc)

    def _dreamclass_request(self, path: str, params: Optional[Dict[str, str]] = None):
        url = f"{DREAMCLASS_API_BASE}{path}"
        if params:
            url = f"{url}?{urllib.parse.urlencode(params)}"
        req = urllib.request.Request(url, headers=self._dreamclass_headers(), method="GET")
        context = ssl.create_default_context(cafile=certifi.where())
        try:
            with urllib.request.urlopen(req, timeout=15, context=context) as resp:
                raw = resp.read().decode("utf-8")
        except ssl.SSLCertVerificationError as e:
            raise RuntimeError(
                "DreamClass SSL verification failed. On macOS, run the Python 'Install Certificates.command' "
                "script, then retry."
            ) from e
        return json.loads(raw)

    def _dreamclass_parse_datetime(self, value: Optional[str]) -> Optional[dt.datetime]:
        if not value:
            return None
        try:
            parsed = dt.datetime.fromisoformat(value)
        except Exception:
            return None
        return self._dreamclass_to_cst(parsed)

    def _dreamclass_to_cst(self, value: dt.datetime) -> dt.datetime:
        chicago = ZoneInfo("America/Chicago")
        if value.tzinfo is None:
            value = value.replace(tzinfo=dt.timezone.utc)
        return value.astimezone(chicago).replace(tzinfo=None)

    def _dreamclass_current_period(self, now: dt.datetime) -> Optional[Dict[str, object]]:
        periods = self._dreamclass_request("/curriculum/schoolperiods/list")
        if not isinstance(periods, list):
            return None
        today = now.date()
        for period in periods or []:
            start_raw = period.get("startDate")
            end_raw = period.get("endDate")
            if not start_raw or not end_raw:
                continue
            try:
                start_date = dt.date.fromisoformat(str(start_raw))
                end_date = dt.date.fromisoformat(str(end_raw))
            except Exception:
                continue
            if start_date <= today <= end_date:
                return period
        return None

    def _dreamclass_fetch_lectures(
        self,
        period_id: str,
        window_start_utc: dt.datetime,
        window_end_utc: dt.datetime,
    ) -> Tuple[List[Dict[str, object]], List[Dict[str, object]]]:
        date_from = window_start_utc.strftime("%Y-%m-%dT%H:%M:%S")
        date_to = window_end_utc.strftime("%Y-%m-%dT%H:%M:%S")
        singles = self._dreamclass_request(
            "/calendar/lecture/single",
            {"periodId": period_id, "dateFrom": date_from, "dateTo": date_to},
        )
        recurring = self._dreamclass_request("/calendar/lecture/recurring", {"periodId": period_id})
        return singles or [], recurring or []

    def _dreamclass_fetch_curriculum(
        self, period_id: str
    ) -> Tuple[Dict[int, str], Dict[int, str], Dict[int, Tuple[int, int]]]:
        classes = self._dreamclass_request(f"/curriculum/schoolperiod/{period_id}/classes")
        classcourses = self._dreamclass_request(f"/curriculum/schoolperiod/{period_id}/classcourses")
        courses = self._dreamclass_request("/curriculum/courses/list")

        class_by_id: Dict[int, str] = {}
        course_by_id: Dict[int, str] = {}
        classcourse_by_id: Dict[int, Tuple[int, int]] = {}

        for item in classes or []:
            try:
                class_id = int(item.get("id"))
                name = str(item.get("name") or "").strip()
                if name:
                    class_by_id[class_id] = name
            except Exception:
                continue

        for item in courses or []:
            try:
                course_id = int(item.get("id"))
                name = str(item.get("name") or "").strip()
                if name:
                    course_by_id[course_id] = name
            except Exception:
                continue

        for item in classcourses or []:
            try:
                cc_id = int(item.get("id"))
                class_id = int(item.get("classId"))
                course_id = int(item.get("courseId"))
                classcourse_by_id[cc_id] = (class_id, course_id)
            except Exception:
                continue

        return class_by_id, course_by_id, classcourse_by_id

    def _dreamclass_class_info(
        self,
        item: Dict[str, object],
        class_by_id: Dict[int, str],
        course_by_id: Dict[int, str],
        classcourse_by_id: Dict[int, Tuple[int, int]],
    ) -> Tuple[str, str]:
        class_name = ""
        course_name = ""
        cc = item.get("classCourse") or {}
        class_id = cc.get("classId")
        course_id = cc.get("courseId")
        cc_id = cc.get("id")

        try:
            if class_id is None and course_id is None and cc_id is not None:
                pair = classcourse_by_id.get(int(cc_id))
                if pair:
                    class_id, course_id = pair
            if class_id is not None:
                class_name = class_by_id.get(int(class_id), "")
            if course_id is not None:
                course_name = course_by_id.get(int(course_id), "")
        except Exception as exc:
            APP_LOG.warning("Could not process DreamClass data ( class info).", exc_info=exc)

        return class_name, course_name

    def _dreamclass_duration_minutes(self, value: Optional[object]) -> int:
        try:
            minutes = int(value) if value is not None else 60
            return max(minutes, 1)
        except Exception:
            return 60

    def _dreamclass_collect_occurrences(
        self,
        singles: List[Dict[str, object]],
        recurring: List[Dict[str, object]],
        window_start: dt.datetime,
        window_end: dt.datetime,
        class_by_id: Dict[int, str],
        course_by_id: Dict[int, str],
        classcourse_by_id: Dict[int, Tuple[int, int]],
    ) -> Tuple[List[Dict[str, object]], bool]:
        occurrences: List[Dict[str, object]] = []
        skipped_recurring = False

        for item in singles:
            start = self._dreamclass_parse_datetime(item.get("startDateTime"))
            if not start:
                try:
                    self.debug(f"DreamClass single skipped: missing/invalid start ({item.get('title')})")
                except Exception as exc:
                    APP_LOG.warning("Could not process DreamClass data ( collect occurrences).", exc_info=exc)
                continue
            try:
                raw = item.get("startDateTime")
                self.debug(f"DreamClass single raw start={raw} parsed={start}")
            except Exception:
                raw = None
            duration = self._dreamclass_duration_minutes(item.get("duration"))
            end = start + dt.timedelta(minutes=duration)
            classroom = item.get("classroom") or {}
            location = str(classroom.get("name") or "").strip()
            title = str(item.get("title") or "Untitled Class").strip()
            try:
                raw_start = item.get("startDateTime")
                raw_end = item.get("endDateTime")
                self.debug(
                    "DreamClass single item: "
                    f"title='{title}' loc='{location}' "
                    f"raw_start={raw_start} raw_end={raw_end} "
                    f"duration={duration} parsed_start={start} parsed_end={end}"
                )
            except Exception as exc:
                APP_LOG.warning("Could not process DreamClass data ( collect occurrences).", exc_info=exc)
            class_name, course_name = self._dreamclass_class_info(
                item, class_by_id, course_by_id, classcourse_by_id
            )
            instructor = self._dreamclass_instructor(item)
            occurrences.append(
                {
                    "title": title,
                    "start": start,
                    "end": end,
                    "start_raw": raw,
                    "location": location,
                    "duration": duration,
                    "class_name": class_name,
                    "course_name": course_name,
                    "instructor": instructor,
                }
            )

        for item in recurring:
            rule = item.get("rrule")
            if not rule:
                try:
                    self.debug(f"DreamClass recurring skipped: missing rrule ({item.get('title')})")
                except Exception as exc:
                    APP_LOG.warning("Could not process DreamClass data ( collect occurrences).", exc_info=exc)
                continue
            if rrule is None:
                skipped_recurring = True
                continue
            try:
                schedule = rrule.rrulestr(str(rule), ignoretz=False)
                ref = window_start
                dtstart = getattr(schedule, "_dtstart", None)
                if isinstance(dtstart, dt.datetime) and dtstart.tzinfo is not None:
                    ref = window_start.replace(tzinfo=ZoneInfo("America/Chicago")).astimezone(dtstart.tzinfo)
                next_start = schedule.after(ref, inc=True)
                if not next_start:
                    try:
                        self.debug(f"DreamClass recurring skipped: no next start ({item.get('title')})")
                    except Exception as exc:
                        APP_LOG.warning("Could not process DreamClass data ( collect occurrences).", exc_info=exc)
                    continue
                next_start = self._dreamclass_to_cst(next_start)
                try:
                    raw = item.get("startDateTime")
                    self.debug(f"DreamClass recurring raw start={raw} computed={next_start}")
                except Exception:
                    raw = None
                duration = self._dreamclass_duration_minutes(item.get("duration"))
                end = next_start + dt.timedelta(minutes=duration)
                classroom = item.get("classroom") or {}
                location = str(classroom.get("name") or "").strip()
                title = str(item.get("title") or "Untitled Class").strip()
                try:
                    raw_start = item.get("startDateTime")
                    raw_end = item.get("endDateTime")
                    self.debug(
                        "DreamClass recurring item: "
                        f"title='{title}' loc='{location}' "
                        f"raw_start={raw_start} raw_end={raw_end} "
                        f"duration={duration} parsed_start={next_start} parsed_end={end}"
                    )
                except Exception as exc:
                    APP_LOG.warning("Could not process DreamClass data ( collect occurrences).", exc_info=exc)
                class_name, course_name = self._dreamclass_class_info(
                    item, class_by_id, course_by_id, classcourse_by_id
                )
                instructor = self._dreamclass_instructor(item)
                occurrences.append(
                    {
                        "title": title,
                        "start": next_start,
                        "end": end,
                        "start_raw": raw,
                        "location": location,
                        "duration": duration,
                        "class_name": class_name,
                        "course_name": course_name,
                        "instructor": instructor,
                    }
                )
            except Exception:
                skipped_recurring = True
                continue

        return occurrences, skipped_recurring

    def _dreamclass_overlap(self, a: Dict[str, object], b: Dict[str, object]) -> bool:
        return bool(a["start"] < b["end"] and b["start"] < a["end"])

    def _dreamclass_pick_slots(
        self, occurrences: List[Dict[str, object]]
    ) -> Dict[str, Optional[Dict[str, object]]]:
        chapel = [o for o in occurrences if o.get("location") == "Chapel"]
        classified = []
        for occ in occurrences:
            location = occ.get("location")
            if location == "Chapel":
                occ = dict(occ)
                occ["class_type"] = "y2"
                classified.append(occ)
                continue
            if location == "Family Center":
                overlap = any(self._dreamclass_overlap(occ, other) for other in chapel)
                occ = dict(occ)
                occ["class_type"] = "y1" if overlap else "combined"
                classified.append(occ)

        def earliest(kind: str) -> Optional[Dict[str, object]]:
            items = [o for o in classified if o.get("class_type") == kind]
            return min(items, key=lambda o: o["start"]) if items else None

        y1 = earliest("y1")
        y2 = earliest("y2")
        combined = earliest("combined")

        y1_slot = y1
        if combined and (y1 is None or combined["start"] < y1["start"]):
            y1_slot = combined
        return {"y1": y1_slot, "y2": y2}

    def _schedule_adjust_key(self, occ: Dict[str, object]) -> str:
        title = str(occ.get("title") or "").strip()
        course = str(occ.get("course_name") or "").strip()
        location = str(occ.get("location") or "").strip()
        return f"{title}|{course}|{location}"

    def _schedule_compute_gaps(
        self,
        occs: List[Dict[str, object]],
        day_start: dt.datetime,
        day_end: dt.datetime,
    ) -> List[Tuple[dt.datetime, dt.datetime]]:
        gaps: List[Tuple[dt.datetime, dt.datetime]] = []
        cursor = day_start
        for occ in sorted(occs, key=lambda o: o["start"]):
            start = occ["start"]
            end = occ["end"]
            if start > cursor:
                gaps.append((cursor, start))
            if end > cursor:
                cursor = end
        if cursor < day_end:
            gaps.append((cursor, day_end))
        return gaps

    def _schedule_prompt_gap_choice(
        self,
        occ: Dict[str, object],
        gaps: List[Tuple[dt.datetime, dt.datetime]],
    ) -> Optional[Tuple[dt.datetime, dt.datetime]]:
        if not gaps:
            return None
        result: Dict[str, Optional[Tuple[dt.datetime, dt.datetime]]] = {"gap": None}
        popup = ctk.CTkToplevel(self)
        popup.title("Adjust Class Time")
        popup.transient(self)
        popup.geometry("460x420")
        try:
            popup.configure(fg_color=BACKGROUND_COLOR)
        except Exception as exc:
            APP_LOG.warning("Could not process the schedule step (prompt gap choice).", exc_info=exc)
        popup.grid_columnconfigure(0, weight=1)
        popup.grid_rowconfigure(1, weight=1)

        title = occ.get("title") or "Untitled Class"
        location = occ.get("location") or "Unknown location"
        ctk.CTkLabel(
            popup,
            text="Select the correct time slot for this class:",
            text_color=TEXT_COLOR,
            font=self.FONT_SMALL,
            anchor="w",
            justify="left",
        ).grid(row=0, column=0, sticky="w", padx=PAD_LG, pady=(12, 4))
        ctk.CTkLabel(
            popup,
            text=f"{title} · {location}",
            text_color=SUBTEXT_COLOR,
            anchor="w",
        ).grid(row=0, column=0, sticky="w", padx=PAD_LG, pady=(36, 0))

        list_frame = ctk.CTkScrollableFrame(popup, fg_color=BACKGROUND_COLOR)
        list_frame.grid(row=1, column=0, sticky="nsew", padx=PAD_LG, pady=(12, 8))
        list_frame.grid_columnconfigure(0, weight=1)

        def choose(gap: Tuple[dt.datetime, dt.datetime]):
            result["gap"] = gap
            try:
                popup.destroy()
            except Exception as exc:
                APP_LOG.warning("Could not apply the selected schedule gap override.", exc_info=exc)

        for gap_start, gap_end in gaps:
            label = f"{gap_start.strftime('%I:%M %p')} - {gap_end.strftime('%I:%M %p')}"
            ctk.CTkButton(
                list_frame,
                text=label,
                command=lambda g=(gap_start, gap_end): choose(g),
                text_color=TEXT_COLOR,
                fg_color=SEPARATOR_COLOR,
                hover_color=INPUT_COLOR,
            ).grid(row=list_frame.grid_size()[1], column=0, sticky="ew", pady=(0, 6))

        ctk.CTkButton(
            popup,
            text="Cancel",
            command=lambda: choose(None),
            text_color=TEXT_COLOR,
            fg_color="#2F2F2F",
            hover_color=SEPARATOR_COLOR,
            width=120,
        ).grid(row=2, column=0, sticky="e", padx=PAD_LG, pady=(0, 12))

        popup.grab_set()
        self.wait_window(popup)
        return result.get("gap")

    def _schedule_apply_gap_overrides(
        self,
        occurrences: List[Dict[str, object]],
        window: Tuple[Optional[dt.datetime], Optional[dt.datetime]],
    ) -> None:
        if not occurrences:
            return
        window_start, window_end = window
        date_ref = window_start.date() if window_start else dt.date.today()
        date_key = date_ref.isoformat()
        if self._schedule_gap_overrides_date != date_key:
            self._schedule_gap_overrides_date = date_key
            self._schedule_gap_overrides = {}
        baseline = dt.datetime.combine(date_ref, dt.time(8, 30))
        day_end = window_end or (baseline + dt.timedelta(days=1))

        location_map: Dict[str, List[Dict[str, object]]] = {}
        for occ in occurrences:
            loc = str(occ.get("location") or "").strip()
            if occ.get("start") >= baseline:
                location_map.setdefault(loc, []).append(occ)

        for occ in sorted(occurrences, key=lambda o: o["start"]):
            if occ.get("start") >= baseline:
                continue
            loc = str(occ.get("location") or "").strip()
            if not loc:
                continue
            key = self._schedule_adjust_key(occ)
            override_start = self._schedule_gap_overrides.get(key)
            if override_start is not None:
                duration = dt.timedelta(minutes=int(occ.get("duration") or 60))
                occ["start"] = override_start
                occ["end"] = override_start + duration
                continue
            gaps = self._schedule_compute_gaps(location_map.get(loc, []), baseline, day_end)
            if not gaps:
                continue
            chosen = gaps[0] if len(gaps) == 1 else self._schedule_prompt_gap_choice(occ, gaps)
            if not chosen:
                continue
            chosen_start, _chosen_end = chosen
            duration = dt.timedelta(minutes=int(occ.get("duration") or 60))
            occ["start"] = chosen_start
            occ["end"] = chosen_start + duration
            self._schedule_gap_overrides[key] = chosen_start
            location_map.setdefault(loc, []).append(occ)

    def _schedule_apply_slot_corrections(self, occurrences: List[Dict[str, object]]) -> None:
        """Apply targeted slot corrections for known DreamClass swap patterns."""
        if not occurrences:
            return
        slots: Dict[dt.datetime, List[Dict[str, object]]] = {}
        for occ in occurrences:
            start = occ.get("start")
            if isinstance(start, dt.datetime):
                slots.setdefault(start, []).append(occ)
        for start, occs in list(slots.items()):
            if len(occs) != 1:
                continue
            primary = occs[0]
            loc = str(primary.get("location") or "").strip()
            if loc != "Family Center":
                continue
            duration = int(primary.get("duration") or 0)
            if duration <= 0:
                continue
            # DreamClass sometimes returns a 60-min Family Center slot that should follow paired 45-min classes.
            next_start = start + dt.timedelta(minutes=duration)
            next_occs = slots.get(next_start, [])
            if len(next_occs) != 2:
                continue
            family_next = [o for o in next_occs if str(o.get("location") or "").strip() == "Family Center"]
            chapel_next = [o for o in next_occs if str(o.get("location") or "").strip() == "Chapel"]
            if len(family_next) != 1 or len(chapel_next) != 1:
                continue
            if int(family_next[0].get("duration") or 0) != 45:
                continue
            if int(chapel_next[0].get("duration") or 0) != 45:
                continue
            for occ in (family_next[0], chapel_next[0]):
                occ["start"] = start
                occ["end"] = start + dt.timedelta(minutes=int(occ.get("duration") or 45))
            primary["start"] = next_start
            primary["end"] = next_start + dt.timedelta(minutes=duration)
            try:
                self.debug(
                    "Schedule correction applied: swapped Family Center 60-min class at "
                    f"{start.strftime('%I:%M %p')} with paired 45-min classes at {next_start.strftime('%I:%M %p')}"
                )
            except Exception as exc:
                APP_LOG.warning("Could not process the schedule step (apply slot corrections).", exc_info=exc)

    def _format_upnext_body(self, occ: Dict[str, object]) -> str:
        start = occ["start"].strftime("%a %b %d, %I:%M %p")
        end = occ["end"].strftime("%I:%M %p")
        title = occ.get("title") or "Untitled Class"
        location = occ.get("location") or "No location"
        class_name = str(occ.get("class_name") or "").strip()
        course_name = str(occ.get("course_name") or "").strip()
        lines = []
        if class_name:
            lines.append(self._format_class_line(class_name, course_name))
        lines.append(str(title))
        lines.append(f"{start} - {end} CST")
        lines.append(str(location))
        return "\n".join(lines)

    def _format_class_line(self, class_name: str, course_name: str) -> str:
        base = class_name
        if course_name:
            base = f"{base} ({course_name})"
        return base

    def _render_schedule_card(
        self,
        parent,
        occ: Dict[str, object],
        inset: bool = False,
        use_grid: bool = False,
        grid_row: int = 0,
        centered: bool = False,
        wrap_len: Optional[int] = None,
    ) -> None:
        card = ctk.CTkFrame(parent, fg_color="#2C2C2C", corner_radius=12)
        if use_grid:
            card.grid(row=grid_row, column=0, sticky="ew", padx=6 if inset else 4, pady=PAD_SM)
        else:
            card.pack(fill=tk.X, expand=True, padx=6 if inset else 4, pady=PAD_SM)
        card.grid_columnconfigure(0, weight=1)
        title = occ.get("title") or "Untitled Class"
        class_name = occ.get("class_name") or ""
        course_name = occ.get("course_name") or ""
        top_text = self._format_class_line(class_name, course_name) if class_name else title
        effective_wrap = int(wrap_len) if isinstance(wrap_len, int) and wrap_len > 0 else (280 if not inset else 360)
        anchor = "center" if centered else "w"
        justify = "center" if centered else "left"
        sticky = "ew" if centered else "w"
        ctk.CTkLabel(
            card,
            text=top_text,
            font=self.FONT_SMALL,
            text_color=TEXT_COLOR,
            anchor=anchor,
            justify=justify,
            wraplength=effective_wrap,
        ).grid(row=0, column=0, sticky=sticky, padx=PAD_MD + 2, pady=(8, 0))
        if class_name:
            ctk.CTkLabel(
                card,
                text=title,
                font=self.FONT_CARD_ITALIC,
                text_color=TEXT_COLOR,
                anchor=anchor,
                justify=justify,
                wraplength=effective_wrap,
            ).grid(row=1, column=0, sticky=sticky, padx=PAD_MD + 2, pady=(2, 8))

    def _refresh_upnext(self, auto: bool = False):
        if not self._dreamclass_enabled:
            self._hide_schedule_spinner()
            lbl = getattr(self, "_upnext_status_label", None)  # None when Up Next/Schedule status UI is not built
            if lbl is not None:
                lbl.configure(text="DreamClass disabled.")
            return
        if self._upnext_fetch_in_progress:
            return
        if not self._dreamclass_is_configured():
            self._hide_schedule_spinner()
            lbl = getattr(self, "_upnext_status_label", None)  # None when Up Next/Schedule status UI is not built
            if lbl is not None:
                lbl.configure(text="DreamClass not configured.")
            self._update_upnext_ui({"y1": None, "y2": None}, note=None)
            return
        if auto and self._pulse_active:
            # Avoid spamming API while live monitoring is active.
            return
        self._upnext_fetch_in_progress = True
        lbl = getattr(self, "_upnext_status_label", None)  # None when Up Next/Schedule status UI is not built
        if lbl is not None:
            lbl.configure(text="Loading schedule…")
        self._show_schedule_spinner()

        def work():
            tz_chi = ZoneInfo("America/Chicago")
            now_cst_aware = dt.datetime.now(tz_chi)
            window_start_local = dt.datetime.combine(now_cst_aware.date(), dt.time(0, 0, tzinfo=tz_chi))
            window_end_local = window_start_local + dt.timedelta(days=1)
            window_start_utc = window_start_local.astimezone(dt.timezone.utc)
            window_end_utc = window_end_local.astimezone(dt.timezone.utc)
            window_start_naive = window_start_local.replace(tzinfo=None)
            window_end_naive = window_end_local.replace(tzinfo=None)

            period = self._dreamclass_current_period(now_cst_aware.replace(tzinfo=None))
            if not period:
                return {
                    "slots": {"y1": None, "y2": None},
                    "period": None,
                    "skipped": False,
                    "now": now_cst_aware.replace(tzinfo=None),
                    "occurrences": [],
                    "window": (window_start_naive, window_end_naive),
                }
            period_id = str(period.get("id"))
            singles, recurring = self._dreamclass_fetch_lectures(period_id, window_start_utc, window_end_utc)
            class_by_id, course_by_id, classcourse_by_id = self._dreamclass_fetch_curriculum(period_id)
            occurrences, skipped = self._dreamclass_collect_occurrences(
                singles,
                recurring,
                window_start_naive,
                window_end_naive,
                class_by_id,
                course_by_id,
                classcourse_by_id,
            )
            slots = self._dreamclass_pick_slots(occurrences)
            return {
                "slots": slots,
                "period": period,
                "skipped": skipped,
                "now": now_cst_aware.replace(tzinfo=None),
                "occurrences": occurrences,
                "window": (window_start_naive, window_end_naive),
            }

        def on_done(result):
            self._upnext_fetch_in_progress = False
            if not isinstance(result, dict):
                self._hide_schedule_spinner()
                self._upnext_status_label.configure(text="DreamClass refresh failed.")
                return
            slots = result.get("slots", {})
            period = result.get("period")
            skipped = bool(result.get("skipped"))
            now = result.get("now") or dt.datetime.now()
            occurrences = result.get("occurrences") or []
            window = result.get("window") or (now, now + dt.timedelta(hours=48))
            note = None
            label = "No active period"
            if period:
                period_label = period.get("shortName") or period.get("name") or "Active Period"
                label = f"{period_label}"
            if note:
                label = f"{label} · {note}"
            lbl = getattr(self, "_upnext_status_label", None)  # None when Up Next/Schedule status UI is not built
            if lbl is not None:
                lbl.configure(text=label)
            self.debug(f"Up Next refreshed at {now.strftime('%I:%M %p')} CST")
            self._schedule_apply_slot_corrections(occurrences)
            self._schedule_apply_gap_overrides(occurrences, window)
            self._schedule_items = sorted(occurrences, key=lambda o: o["start"])
            self._schedule_period_label = label
            self._schedule_window = window
            self._update_upnext_ui(slots, note=note)
            self._render_schedule_items()

        def on_error(err: Exception):
            self._upnext_fetch_in_progress = False
            self._hide_schedule_spinner()
            lbl = getattr(self, "_upnext_status_label", None)  # None when Up Next/Schedule status UI is not built
            if lbl is not None:
                lbl.configure(text=f"DreamClass error: {err}")

        self._run_job(name="dreamclass", work=work, on_done=on_done, on_error=on_error)

    def _update_upnext_ui(self, slots: Dict[str, Optional[Dict[str, object]]], note: Optional[str]):
        pointer_parts: List[str] = []
        y1 = slots.get("y1")
        y2 = slots.get("y2")
        if y1:
            pointer_parts.append(f"Y1: {y1.get('title') or 'Class'} at {y1['start'].strftime('%I:%M %p')}")
        if y2:
            pointer_parts.append(f"Y2: {y2.get('title') or 'Class'} at {y2['start'].strftime('%I:%M %p')}")
        if not pointer_parts:
            pointer = "No upcoming classes"
        else:
            pointer = " | ".join(pointer_parts)
        if note:
            pointer = f"{pointer} · {note}"
        lbl = getattr(self, "_upnext_pointer_label", None)  # None when Up Next header UI is not built
        if lbl is not None:
            try:
                lbl.configure(text=pointer)
            except Exception as exc:
                APP_LOG.warning("Could not update upnext ui.", exc_info=exc)


    # ---------------- Find Live Tab -----------------
    def _build_find_tab(self, parent):
        try:
            parent.configure(fg_color=BACKGROUND_COLOR)
        except Exception as exc:
            APP_LOG.warning("Could not finish building the find tab UI section.", exc_info=exc)
        parent.grid_columnconfigure(0, weight=1)
        parent.grid_rowconfigure(0, weight=0)
        parent.grid_rowconfigure(2, weight=0)
        parent.grid_rowconfigure(3, weight=1)
        parent.grid_rowconfigure(4, weight=0)

        # Scan button at top
        self.btn_scan = ctk.CTkButton(
            parent,
            text="Scan Channels",
            command=self._find_scan,
            height=42,
            text_color=TEXT_COLOR,
        )
        self.btn_scan.grid(row=0, column=0, sticky="ew", padx=PAD_MD, pady=(4, 0))
        status_frame = ctk.CTkFrame(parent, fg_color=SURFACE_MID_COLOR, corner_radius=8)
        status_frame.grid(row=1, column=0, sticky="ew", padx=PAD_MD, pady=(PAD_SM, 0))
        self.lbl_find_status = ctk.CTkLabel(status_frame, text="Idle", text_color=SUBTEXT_COLOR)
        self.lbl_find_status.pack(anchor="w", padx=10, pady=6)

        # Results chips
        table_frm = ctk.CTkFrame(parent, fg_color=BACKGROUND_COLOR)
        table_frm.grid(row=3, column=0, sticky="nsew", padx=PAD_MD, pady=PAD_MD)
        self.find_results_container = ctk.CTkScrollableFrame(table_frm, fg_color=BACKGROUND_COLOR)
        self.find_results_container.pack(fill=tk.BOTH, expand=True)
        self.find_results_container.grid_columnconfigure(0, weight=1)

        # Single wide Copy ID button spanning horizontally; tall by +100px
        try:
            base_h = int(self.btn_scan.cget("height"))
        except Exception:
            base_h = 42
        big_font = self.FONT_SUBHEADING
        self.btn_copy_id = ctk.CTkButton(
            parent,
            text="Copy Video ID",
            command=self._find_copy_video_id,
            height=base_h,
            font=big_font,
            text_color=TEXT_COLOR,
        )
        self.btn_copy_id.grid(row=4, column=0, sticky="ew", padx=PAD_MD, pady=(0, 8))

        self._find_results: List[Dict[str, Optional[str]]] = []
        self._find_last_result: Optional[Dict[str, Optional[str]]] = None

    def _find_scan(self):
        # Clear old rows
        for child in list(self.find_results_container.winfo_children()):
            child.destroy()
        self._find_results.clear()
        self._find_last_result = None
        self.lbl_find_status.configure(text="Scanning…")

        def add_row(res: Dict[str, Optional[str]]):
            row = ctk.CTkFrame(self.find_results_container, fg_color=SEPARATOR_COLOR, corner_radius=12)
            row.grid_columnconfigure(0, weight=1)
            row.pack(fill=tk.X, expand=True, pady=PAD_SM + 2, padx=PAD_SM)

            channel_label = ctk.CTkLabel(
                row,
                text=res.get("channel_title") or res.get("channel_id") or "",
                text_color="#A7C7FF",
                font=self.FONT_ROW_LABEL,
                anchor="w",
            )
            channel_label.grid(row=0, column=0, sticky="w", padx=PAD_LG, pady=(8, 4))

            chips = ctk.CTkFrame(row, fg_color=SEPARATOR_COLOR)
            chips.grid(row=1, column=0, sticky="ew", padx=PAD_LG, pady=(0, 12))
            chips.grid_columnconfigure(0, weight=1)
            chip_row = 0

            def copy_value(value: str, what: str, widget: Optional[ctk.CTkButton] = None):
                if not value:
                    return
                self._find_last_result = res
                self._copy_and_log(value, self.lbl_find_status, what=what)
                if widget is not None:
                    self._flash_button(widget, flashes=2)

            vid = res.get("video_id") or ""
            if vid:
                btn_vid = self._find_create_chip(
                    chips,
                    label="Video ID",
                    value=vid,
                    width=300,
                    command=lambda: None,
                )
                btn_vid.grid(row=chip_row, column=0, sticky="ew", pady=(0, 6))
                btn_vid.configure(command=lambda v=vid, b=btn_vid: copy_value(v, "video ID", b))
                chip_row += 1
                if self._find_last_result is None:
                    self._find_last_result = res

            url = res.get("video_url") or ""
            if url:
                btn_url = self._find_create_chip(
                    chips,
                    label="Livestream URL",
                    value=url,
                    width=540,
                    command=lambda: None,
                )
                btn_url.grid(row=chip_row, column=0, sticky="ew", pady=(0, 6))
                btn_url.configure(command=lambda u=url, b=btn_url: copy_value(u, "video URL", b))
                chip_row += 1

            stream_title = res.get("stream_title") or ""
            if stream_title:
                lbl_title = self._find_create_chip(
                    chips,
                    label="Stream Title",
                    value=stream_title,
                    width=700,
                    command=None,
                    passive=True,
                )
                lbl_title.grid(row=chip_row, column=0, sticky="ew", pady=(0, 6))
                chip_row += 1

            stream_start_time = res.get("stream_start_time") or ""
            if stream_start_time:
                lbl_start = self._find_create_chip(
                    chips,
                    label="Start Time",
                    value=stream_start_time,
                    width=420,
                    command=None,
                    passive=True,
                )
                lbl_start.grid(row=chip_row, column=0, sticky="ew", pady=(0, 6))
                chip_row += 1

            status_msg = res.get("status") or ""
            if status_msg:
                lbl_status = self._find_create_chip(
                    chips,
                    label="Status",
                    value=status_msg,
                    width=320,
                    command=None,
                    passive=True,
                )
                lbl_status.grid(row=chip_row, column=0, sticky="ew")

            self._find_results.append(res)
            self.debug(f"Find: {res['channel_title']} -> {res.get('video_id', '')}")

        def add_msg(channel_title: str, channel_id: str, msg: str):
            res = {
                "channel_title": channel_title,
                "channel_id": channel_id,
                "video_id": "",
                "video_url": "",
                "status": msg,
            }
            add_row(res)
            self.debug(f"Find: {channel_title} ({channel_id}) -> {msg}")

        def work():
            try:
                self.debug("Find: scanning channels…")
                # Only scan channels that are currently monitored (checkbox ON)
                monitored = self._get_monitored_channels()
                if not monitored:
                    return "No monitored channels"
                for ch in monitored:
                    cid = ch["channel_id"]
                    ch_name = ch.get("name", cid)
                    if self.preview_mode.get():
                        # Preview: fake result
                        demo_vid = "DEMO12345"
                        res = {
                            "channel_title": ch_name,
                            "channel_id": cid,
                            "video_id": demo_vid,
                            "video_url": f"https://www.youtube.com/watch?v={demo_vid}",
                            "stream_title": "Preview Stream",
                            "stream_start_time": dt.datetime.now().strftime("%H:%M:%S"),
                        }
                        self.after(0, lambda r=res: add_row(r))
                        continue
                    try:
                        yt = self._get_youtube(cid)
                        mine = get_authenticated_channel_info(yt)
                        if not mine or mine["id"] != cid:
                            self.after(0, lambda c=cid, n=ch_name: add_msg(n, c, "token mismatch"))
                            continue
                        details = find_latest_active_or_testing_stream(yt, cid, ignore_privacy=False, logger=self.debug)
                        vid = (details or {}).get("video_id")
                        self._last_live_ids[cid] = vid
                        if vid:
                            title = get_channel_title_public(yt, cid, logger=self.debug)
                            res = {
                                "channel_title": title,
                                "channel_id": cid,
                                "video_id": vid,
                                "video_url": f"https://www.youtube.com/watch?v={vid}",
                                "stream_title": (details or {}).get("stream_title") or "",
                                "stream_start_time": self._format_stream_start_time(
                                    (details or {}).get("actual_start_time"),
                                    (details or {}).get("scheduled_start_time"),
                                ),
                            }
                            self.after(0, lambda r=res: add_row(r))
                        else:
                            self.after(0, lambda c=cid, n=ch_name: add_msg(n, c, "no live"))
                    except Exception as e:
                        if self._handle_auth_error(cid, e, ch_name):
                            self.after(0, lambda c=cid, n=ch_name: add_msg(n, c, "auth error - re-auth token"))
                        else:
                            self.after(0, lambda c=cid, n=ch_name, m=str(e): add_msg(n, c, f"error: {m}"))
            finally:
                self.after(0, lambda: (self.lbl_find_status.configure(text="Done"), self.debug("Find: done")))
            return "OK"

        def on_done(result):
            if result == "No monitored channels":
                self.lbl_find_status.configure(text="No monitored channels")
        self._run_job(
            name="find-scan",
            work=work,
            disable=[self.btn_scan],
            status_widget=self.lbl_find_status,
            start_text="Scanning…",
            done_text=None,
            on_done=on_done,
        )

    def _find_create_chip(
        self,
        parent,
        *,
        label: str,
        value: str,
        width: int,
        command: Optional[Callable[[], None]],
        passive: bool = False,
    ):
        text = f"{label}: {value}" if value else label
        font = self.FONT_BODY
        base_color = SURFACE_COLOR
        hover_color = INPUT_COLOR if not passive else SURFACE_COLOR
        chip = ctk.CTkButton(
            parent,
            text=text,
            command=command or (lambda: None),
            fg_color=base_color,
            hover_color=hover_color,
            text_color=TEXT_COLOR,
            text_color_disabled=TEXT_COLOR,
            font=font,
            corner_radius=8,
        )
        chip.configure(anchor="w")
        if passive or not command:
            chip.configure(state="disabled")
        chip.configure(width=width)
        try:
            chip._text_label.configure(wraplength=width - 24, justify="left")
        except Exception:
            try:
                chip.configure(wraplength=width - 24)
            except Exception as exc:
                APP_LOG.warning("Could not update the Find tab UI (create chip).", exc_info=exc)
        return chip

    def _find_copy_video_id(self):
        res = getattr(self, "_find_last_result", None)  # None until a Find tab scan has produced a result
        if not res or not res.get("video_id"):
            return
        self._copy_and_log(res["video_id"], self.lbl_find_status, what="video ID")

    # ---------------- Archive Tab -----------------
    # ---------------- Channels Tab -----------------
    def _build_channels_tab(self, parent):
        try:
            parent.configure(fg_color=BACKGROUND_COLOR)
        except Exception as exc:
            APP_LOG.warning("Could not finish building the channels tab UI section.", exc_info=exc)
        parent.grid_columnconfigure(0, weight=1)
        parent.grid_rowconfigure(0, weight=1)
        parent.grid_rowconfigure(3, weight=1)

        # Table of channels
        table_frm = ctk.CTkFrame(parent, fg_color=BACKGROUND_COLOR)
        table_frm.grid(row=0, column=0, sticky="nsew", padx=PAD_MD + 2, pady=(10, 6))
        cols = ("label", "name", "channel_id", "playlist_id", "thumbnail_template", "show_monitor")
        self.chan_tree = ttk.Treeview(table_frm, columns=cols, show="headings", height=6)
        column_configs = {
            "label": 80,
            "name": 200,
            "channel_id": 260,
            "playlist_id": 260,
            "thumbnail_template": 260,
            "show_monitor": 140,
        }
        for c in cols:
            self.chan_tree.heading(c, text=c.replace("_", " ").title())
            width = column_configs.get(c, 160)
            anchor = "center" if c == "show_monitor" else "w"
            self.chan_tree.column(c, width=width, anchor=anchor)
        self.chan_tree.pack(fill=tk.BOTH, expand=True)
        btn_remove = ctk.CTkButton(
            table_frm,
            text="Remove Selected Row",
            command=self._chan_delete_selected,
            text_color=TEXT_COLOR,
            fg_color="#D43E27",
            hover_color="#B7321F",
        )
        btn_remove.pack(side=tk.RIGHT, padx=0, pady=(6, 0))
        try:
            style = ttk.Style()
            style.configure(
                "Robuster.Treeview",
                background=INPUT_COLOR,
                fieldbackground=INPUT_COLOR,
                foreground=TEXT_COLOR,
                rowheight=26,
                bordercolor=INPUT_COLOR,
            )
            style.configure(
                "Robuster.Treeview.Heading",
                foreground=TEXT_COLOR,
                background=SURFACE_COLOR,
                relief="flat",
            )
            style.map(
                "Robuster.Treeview",
                background=[("selected", INPUT_HOVER_COLOR)],
                foreground=[("selected", TEXT_COLOR)],
            )
            self.chan_tree.configure(style="Robuster.Treeview")
        except Exception as exc:
            APP_LOG.warning("Could not finish building the channels tab UI section.", exc_info=exc)

        # Form
        form = ctk.CTkFrame(parent, fg_color=BACKGROUND_COLOR)
        form.grid(row=1, column=0, sticky="ew", padx=PAD_MD + 2)
        for i in range(8):
            form.grid_columnconfigure(i, weight=1)

        self.var_label = tk.StringVar()
        self.var_name = tk.StringVar()
        self.var_cid = tk.StringVar()
        self.var_pid = tk.StringVar()
        self.var_thumbnail_template = tk.StringVar()

        entry_kwargs = {
            "fg_color": INPUT_COLOR,
            "border_color": INPUT_COLOR,
            "text_color": TEXT_COLOR,
        }

        ctk.CTkLabel(form, text="Label", text_color=TEXT_COLOR).grid(row=0, column=0, sticky="w")
        ctk.CTkEntry(form, textvariable=self.var_label, **entry_kwargs).grid(row=1, column=0, sticky="ew", padx=(0, 6))
        ctk.CTkLabel(form, text="Name", text_color=TEXT_COLOR).grid(row=0, column=1, sticky="w")
        ctk.CTkEntry(form, textvariable=self.var_name, **entry_kwargs).grid(row=1, column=1, sticky="ew", padx=(0, 6))
        ctk.CTkLabel(form, text="Channel ID", text_color=TEXT_COLOR).grid(row=0, column=2, sticky="w")
        ctk.CTkEntry(form, textvariable=self.var_cid, **entry_kwargs).grid(row=1, column=2, sticky="ew", padx=(0, 6))
        ctk.CTkLabel(form, text="Playlist ID", text_color=TEXT_COLOR).grid(row=0, column=3, sticky="w")
        ctk.CTkEntry(form, textvariable=self.var_pid, **entry_kwargs).grid(row=1, column=3, sticky="ew", padx=(0, 6))
        ctk.CTkLabel(form, text="Thumbnail Template", text_color=TEXT_COLOR).grid(row=0, column=4, sticky="w")
        ctk.CTkEntry(form, textvariable=self.var_thumbnail_template, **entry_kwargs).grid(row=1, column=4, sticky="ew", padx=(0, 6))
        ctk.CTkButton(
            form,
            text="Browse",
            command=self._chan_browse_thumbnail_template,
            text_color=TEXT_COLOR,
            width=90,
        ).grid(row=1, column=5, sticky="ew", padx=(0, 6))
        ctk.CTkButton(
            form,
            text="Add Channel",
            command=self._chan_add_channel,
            text_color=TEXT_COLOR,
        ).grid(row=2, column=0, columnspan=6, sticky="ew", pady=(8, 0))

        token_box = ctk.CTkFrame(parent, fg_color=BACKGROUND_COLOR)
        token_box.grid(row=2, column=0, sticky="nsew", padx=PAD_MD + 2, pady=(8, 0))
        token_box.grid_columnconfigure(0, weight=1)
        token_box.grid_rowconfigure(1, weight=1)
        ctk.CTkLabel(token_box, text="Tokens", text_color=TEXT_COLOR).grid(row=0, column=0, sticky="w")
        tokens_btns = ctk.CTkFrame(token_box, fg_color=BACKGROUND_COLOR)
        tokens_btns.grid(row=0, column=0, sticky="e")
        ctk.CTkButton(
            tokens_btns,
            text="Check Tokens",
            command=self._chan_check_tokens,
            text_color=TEXT_COLOR,
            width=120,
            height=34,
        ).pack(side=tk.RIGHT, padx=(6, 0))
        self._chan_token_results = ctk.CTkScrollableFrame(token_box, fg_color=BACKGROUND_COLOR, height=120)
        self._chan_token_results.grid(row=1, column=0, sticky="nsew", pady=(6, 0))
        self._chan_token_results.grid_columnconfigure(0, weight=1)

        monitor_box = ctk.CTkFrame(parent, fg_color=BACKGROUND_COLOR)
        monitor_box.grid(row=3, column=0, sticky="nsew", padx=PAD_MD + 2, pady=(10, 0))
        monitor_box.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(monitor_box, text="Monitor Chips", text_color=TEXT_COLOR).grid(row=0, column=0, sticky="w")
        self._chan_monitor_scroll = ctk.CTkScrollableFrame(monitor_box, fg_color=BACKGROUND_COLOR, height=160)
        self._chan_monitor_scroll.grid(row=1, column=0, sticky="nsew", pady=(6, 6))
        self._chan_monitor_scroll.grid_columnconfigure(0, weight=1)

        # Buttons row
        btns = ctk.CTkFrame(parent, fg_color=BACKGROUND_COLOR)
        btns.grid(row=4, column=0, sticky="ew", padx=PAD_MD + 2, pady=(6, 10))
        btns.grid_columnconfigure(0, weight=1)
        btns.grid_columnconfigure(1, weight=0)
        btns.grid_columnconfigure(2, weight=0)

        ctk.CTkButton(btns, text="Close", command=self._close_settings_popup, text_color=TEXT_COLOR).grid(row=0, column=1, sticky="e", padx=PAD_SM)
        ctk.CTkButton(btns, text="Save + Apply Changes", command=self._chan_save_apply, text_color=TEXT_COLOR).grid(row=0, column=2, sticky="e", padx=PAD_SM)

        # Bind selection -> populate form
        def on_tree_select(_e=None):
            sel = self.chan_tree.selection()
            if not sel:
                return
            vals = self.chan_tree.item(sel[0], "values")
            if not vals or len(vals) < 6:
                return
            self.var_label.set(vals[0])
            self.var_name.set(vals[1])
            self.var_cid.set(vals[2])
            self.var_pid.set(vals[3])
            self.var_thumbnail_template.set(vals[4])
        self.chan_tree.bind("<<TreeviewSelect>>", on_tree_select)

        self._chan_refresh_table()
        self._chan_check_tokens()

    def _apply_preset_dropdown_values(self, data: Dict[str, List[str]]):
        classes = [str(x).strip() for x in data.get("class_titles", []) if isinstance(x, str)]
        instructors = [str(x).strip() for x in data.get("instructors", []) if isinstance(x, str)]
        self.cb_class.configure(values=classes)
        self.cb_instr.configure(values=instructors)

    def _chan_check_tokens(self):
        frame = getattr(self, "_chan_token_results", None)  # None until the Channels settings tab is opened
        if frame is None:
            return
        for child in list(frame.winfo_children()):
            try:
                child.destroy()
            except Exception as exc:
                APP_LOG.warning("Could not update Channels settings UI (check tokens).", exc_info=exc)
        for ch in self.channels:
            cid = ch.get("channel_id")
            if not cid:
                continue
            label = ch.get("label") or ch.get("name") or cid[:8]
            row = ctk.CTkFrame(frame, fg_color=SEPARATOR_COLOR, corner_radius=10)
            row.grid_columnconfigure(0, weight=1)
            row.pack(fill=tk.X, expand=True, pady=PAD_SM, padx=2)
            ctk.CTkLabel(row, text=label, text_color=TEXT_COLOR, anchor="w").grid(row=0, column=0, sticky="w", padx=PAD_MD + 2, pady=PAD_SM + 2)
            status_lbl = ctk.CTkLabel(row, text="", text_color=SECONDARY_TEXT_COLOR, anchor="w")
            status_lbl.grid(row=1, column=0, sticky="w", padx=PAD_MD + 2, pady=(0, 6))

            token_path = token_file_for_channel_id(cid)
            has_token = os.path.exists(token_path)
            if has_token:
                status_lbl.configure(text="Token found")
            else:
                status_lbl.configure(text="No token")

            def do_auth(channel_id: str, lbl: ctk.CTkLabel, name: str):
                def work():
                    try:
                        yt = authenticate_youtube(token_file_for_channel_id(channel_id))
                        info = get_authenticated_channel_info(yt)
                        return info
                    except Exception as e:
                        return e

                def on_done(res):
                    if isinstance(res, Exception):
                        try:
                            lbl.configure(text=f"Auth failed: {res}")
                        except Exception as exc:
                            APP_LOG.warning("Could not update UI state after a background task finished.", exc_info=exc)
                        self.debug(f"Auth error for {name}: {res}")
                    else:
                        try:
                            lbl.configure(text="Token saved")
                        except Exception as exc:
                            APP_LOG.warning("Could not update UI state after a background task finished.", exc_info=exc)
                        self.debug(f"Auth success for {name}")
                        self._clear_channel_client(channel_id)

                self._run_job(
                    name=f"auth-{channel_id}",
                    work=work,
                    disable=None,
                    status_widget=None,
                    start_text=None,
                    done_text=None,
                    on_done=on_done,
                )

            btn_auth = ctk.CTkButton(
                row,
                text="Sign in" if not has_token else "Re-auth",
                text_color=TEXT_COLOR,
                command=lambda cid=cid, lbl=status_lbl, name=label: do_auth(cid, lbl, name),
                width=120,
                height=34,
            )
            btn_auth.grid(row=0, column=1, sticky="e", padx=PAD_SM + 2, pady=PAD_SM + 2)

            def do_delete(channel_id: str, lbl: ctk.CTkLabel):
                try:
                    os.remove(token_file_for_channel_id(channel_id))
                    lbl.configure(text="Token deleted")
                    self.debug(f"Token deleted for {channel_id}")
                    self._clear_channel_client(channel_id)
                except FileNotFoundError:
                    lbl.configure(text="No token to delete")
                except Exception as e:
                    lbl.configure(text=f"Delete failed: {e}")

            if has_token:
                btn_del = ctk.CTkButton(
                    row,
                    text="Delete Token",
                    text_color=TEXT_COLOR,
                    fg_color="#D43E27",
                    hover_color="#B7321F",
                    command=lambda cid=cid, lbl=status_lbl: do_delete(cid, lbl),
                    width=120,
                    height=34,
                )
                btn_del.grid(row=1, column=1, sticky="e", padx=PAD_SM + 2, pady=(0, 6))

    # ---------------- Settings Tab -----------------
    def _build_general_settings_tab(self, parent):
        try:
            parent.configure(fg_color=BACKGROUND_COLOR)
        except Exception as exc:
            APP_LOG.warning("Could not finish building the general settings tab UI section.", exc_info=exc)
        parent.grid_columnconfigure(0, weight=1)
        parent.grid_rowconfigure(5, weight=1)

        desc = ctk.CTkLabel(
            parent,
            text="General settings for Robuster. These apply immediately when you hit Apply.",
            text_color=TEXT_COLOR,
            wraplength=720,
            justify="left",
        )
        desc.grid(row=0, column=0, sticky="w", padx=PAD_MD + 2, pady=(10, 8))

        box = ctk.CTkFrame(parent, fg_color=BACKGROUND_COLOR)
        box.grid(row=1, column=0, sticky="nsew", padx=PAD_MD + 2, pady=(0, 10))
        box.grid_columnconfigure(0, weight=1)
        self._general_settings_frame = box

        self._general_start_preview_var = tk.BooleanVar(value=self._general_settings.get("start_in_preview", False))
        self._general_auto_pulse_var = tk.BooleanVar(value=self._general_settings.get("auto_pulse_after_launch", False))
        self._general_topmost_var = tk.BooleanVar(value=self._general_settings.get("stay_on_top", False))
        self._general_show_log_var = tk.BooleanVar(value=self._general_settings.get("show_activity_log", True))

        ctk.CTkCheckBox(
            box,
            text="Start in Preview Mode",
            variable=self._general_start_preview_var,
            text_color=TEXT_COLOR,
        ).grid(row=0, column=0, sticky="w", pady=(4, 4))
        ctk.CTkLabel(
            box,
            text="Runs actions in preview/dry-run by default until you turn it off.",
            text_color=SUBTEXT_COLOR,
        ).grid(row=1, column=0, sticky="w", pady=(0, 6))

        ctk.CTkCheckBox(
            box,
            text="Auto-start Live Monitor pulse after launch",
            variable=self._general_auto_pulse_var,
            text_color=TEXT_COLOR,
        ).grid(row=2, column=0, sticky="w", pady=(4, 4))
        ctk.CTkLabel(
            box,
            text="Automatically begin the live check pulse when the app opens.",
            text_color=SUBTEXT_COLOR,
        ).grid(row=3, column=0, sticky="w", pady=(0, 6))

        ctk.CTkCheckBox(
            box,
            text="Keep Robuster window on top",
            variable=self._general_topmost_var,
            text_color=TEXT_COLOR,
        ).grid(row=4, column=0, sticky="w", pady=(4, 4))
        ctk.CTkLabel(
            box,
            text="Prevents other windows from covering the app.",
            text_color=SUBTEXT_COLOR,
        ).grid(row=5, column=0, sticky="w", pady=(0, 6))
        ctk.CTkCheckBox(
            box,
            text="Show Activity Log",
            variable=self._general_show_log_var,
            text_color=TEXT_COLOR,
        ).grid(row=6, column=0, sticky="w", pady=(4, 4))
        ctk.CTkLabel(
            box,
            text="Hide or show the Activity Log panel below the tabs.",
            text_color=SUBTEXT_COLOR,
        ).grid(row=7, column=0, sticky="w", pady=(0, 6))

        actions_row = ctk.CTkFrame(parent, fg_color=BACKGROUND_COLOR)
        actions_row.grid(row=2, column=0, sticky="ew", padx=PAD_MD + 2, pady=(0, 10))
        actions_row.grid_columnconfigure(0, weight=1)
        self._general_update_btn = ctk.CTkButton(
            actions_row,
            text="Check for Updates",
            command=self._check_for_updates_manual,
            text_color=TEXT_COLOR,
            width=150,
        )
        self._general_update_btn.grid(row=0, column=0, sticky="w", padx=(0, PAD_SM))
        ctk.CTkButton(
            actions_row,
            text="Apply General Settings",
            command=self._apply_general_settings,
            text_color=TEXT_COLOR,
        ).grid(row=0, column=1, sticky="e")
        self._general_status_label = ctk.CTkLabel(parent, text="", text_color=SUBTEXT_COLOR)
        self._general_status_label.grid(row=3, column=0, sticky="w", padx=PAD_MD + 2, pady=(0, 6))
        self._general_update_status_label = ctk.CTkLabel(parent, text="", text_color=SUBTEXT_COLOR)
        self._general_update_status_label.grid(row=4, column=0, sticky="w", padx=PAD_MD + 2, pady=(0, 6))
        spacer = ctk.CTkFrame(parent, fg_color="transparent")
        spacer.grid(row=5, column=0, sticky="nsew")

    def _build_dreamclass_tab(self, parent):
        try:
            parent.configure(fg_color=BACKGROUND_COLOR)
        except Exception as exc:
            APP_LOG.warning("Could not finish building the dreamclass tab UI section.", exc_info=exc)
        parent.grid_columnconfigure(0, weight=1)
        parent.grid_rowconfigure(4, weight=1)
        desc = ctk.CTkLabel(
            parent,
            text="DreamClass schedule + Up Next configuration. Toggle the feature on/off and provide the API credentials.",
            text_color=TEXT_COLOR,
            wraplength=720,
            justify="left",
        )
        desc.grid(row=0, column=0, sticky="w", padx=PAD_MD + 2, pady=(10, 8))

        box = ctk.CTkFrame(parent, fg_color=BACKGROUND_COLOR)
        box.grid(row=1, column=0, sticky="nsew", padx=PAD_MD + 2, pady=(0, 10))
        box.grid_columnconfigure(1, weight=1)

        self._dreamclass_enabled_var = tk.BooleanVar(value=self._dreamclass_enabled)
        ctk.CTkCheckBox(
            box,
            text="Enable DreamClass schedule + Up Next",
            variable=self._dreamclass_enabled_var,
            text_color=TEXT_COLOR,
        ).grid(row=0, column=0, columnspan=2, sticky="w", pady=(4, 6))
        ctk.CTkLabel(
            box,
            text="When off, the Schedule tab and View Schedule button are hidden and DreamClass checks stop running.",
            text_color=SUBTEXT_COLOR,
            wraplength=700,
            justify="left",
        ).grid(row=1, column=0, columnspan=2, sticky="w", pady=(0, 10))

        ctk.CTkLabel(box, text="API Key", text_color=TEXT_COLOR).grid(row=2, column=0, sticky="w")
        self._dreamclass_api_key_var = tk.StringVar(value=self._dreamclass_settings.get("api_key", ""))
        self._make_entry(
            box,
            textvariable=self._dreamclass_api_key_var,
            show="*",
        ).grid(row=2, column=1, sticky="ew", padx=(8, 0))
        ctk.CTkLabel(box, text="School Code", text_color=TEXT_COLOR).grid(row=3, column=0, sticky="w", pady=(6, 0))
        self._dreamclass_school_code_var = tk.StringVar(value=self._dreamclass_settings.get("school_code", ""))
        self._make_entry(
            box,
            textvariable=self._dreamclass_school_code_var,
        ).grid(row=3, column=1, sticky="ew", padx=(8, 0), pady=(6, 0))
        ctk.CTkLabel(box, text="Tenant", text_color=TEXT_COLOR).grid(row=4, column=0, sticky="w", pady=(6, 0))
        self._dreamclass_tenant_var = tk.StringVar(value=self._dreamclass_settings.get("tenant", ""))
        self._make_entry(
            box,
            textvariable=self._dreamclass_tenant_var,
        ).grid(row=4, column=1, sticky="ew", padx=(8, 0), pady=(6, 0))

        save_row = ctk.CTkFrame(parent, fg_color=BACKGROUND_COLOR)
        save_row.grid(row=2, column=0, sticky="ew", padx=PAD_MD + 2, pady=(0, 10))
        save_row.grid_columnconfigure(0, weight=1)
        ctk.CTkButton(
            save_row,
            text="Save DreamClass Settings",
            command=self._save_dreamclass_settings,
            text_color=TEXT_COLOR,
        ).grid(row=0, column=1, sticky="e")
        self._dreamclass_status_label = ctk.CTkLabel(parent, text="", text_color=SUBTEXT_COLOR)
        self._dreamclass_status_label.grid(row=3, column=0, sticky="w", padx=PAD_MD + 2, pady=(0, 10))
        spacer = ctk.CTkFrame(parent, fg_color="transparent")
        spacer.grid(row=4, column=0, sticky="nsew")

    def _build_schedule_tab(self, parent):
        try:
            parent.configure(fg_color=BACKGROUND_COLOR)
        except Exception as exc:
            APP_LOG.warning("Could not finish building the schedule tab UI section.", exc_info=exc)
        parent.grid_columnconfigure(0, weight=1)

        container = ctk.CTkFrame(parent, fg_color=BACKGROUND_COLOR)
        container.grid(row=0, column=0, sticky="nsew", padx=PAD_MD, pady=(8, 8))
        container.grid_columnconfigure(0, weight=1)
        container.grid_rowconfigure(1, weight=1)

        # Persistent column header row (outside scroll area)
        col_header = ctk.CTkFrame(container, fg_color=SURFACE_MID_COLOR, corner_radius=10)
        col_header.grid(row=0, column=0, sticky="ew", pady=(0, 6))
        try:
            col_header.grid_columnconfigure(0, weight=0)
            col_header.grid_columnconfigure(1, weight=1, minsize=220, uniform="schedcols")
            col_header.grid_columnconfigure(2, weight=0)
            col_header.grid_columnconfigure(3, weight=1, minsize=220, uniform="schedcols")
        except Exception as exc:
            APP_LOG.warning("Could not finish building the schedule tab UI section.", exc_info=exc)
        ctk.CTkLabel(col_header, text="", text_color=TEXT_COLOR).grid(row=0, column=0, sticky="w")
        ctk.CTkLabel(
            col_header,
            text=SCHEDULE_COL_LEFT,
            text_color=TEXT_COLOR,
            font=self.FONT_SCHEDULE,
        ).grid(row=0, column=1, sticky="nsew", padx=PAD_SM)
        ctk.CTkFrame(col_header, fg_color=SEPARATOR_COLOR, width=2, height=28).grid(
            row=0, column=2, sticky="n", padx=2, pady=PAD_SM
        )
        ctk.CTkLabel(
            col_header,
            text=SCHEDULE_COL_RIGHT,
            text_color=TEXT_COLOR,
            font=self.FONT_SCHEDULE,
        ).grid(row=0, column=3, sticky="nsew", padx=PAD_SM)
        self._schedule_list = ctk.CTkScrollableFrame(container, fg_color=BACKGROUND_COLOR)
        self._schedule_list.grid(row=1, column=0, sticky="nsew")
        self._schedule_list.grid_columnconfigure(0, weight=1)
        parent.grid_rowconfigure(0, weight=1)

        # Schedule controls bar (only on Schedule tab)
        sched_bar = ctk.CTkFrame(parent, fg_color=BACKGROUND_COLOR)
        sched_bar.grid(row=1, column=0, sticky="ew", padx=PAD_MD, pady=(0, 6))
        sched_bar.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(
            sched_bar,
            text="Schedule",
            font=self.FONT_LABEL,
            text_color=TEXT_COLOR,
        ).grid(row=0, column=0, sticky="w")
        self._schedule_status_label = ctk.CTkLabel(sched_bar, text="Loading…", text_color=SUBTEXT_COLOR)
        self._schedule_status_label.grid(row=0, column=1, sticky="w", padx=(8, 0))
        btns = ctk.CTkFrame(sched_bar, fg_color=BACKGROUND_COLOR)
        btns.grid(row=0, column=2, sticky="e")
        ctk.CTkButton(
            btns,
            text="Refresh",
            command=self._refresh_upnext,
            text_color=TEXT_COLOR,
            width=110,
        ).pack(side=tk.LEFT, padx=(0, 6))
        ctk.CTkButton(
            btns,
            text="Dump to Log",
            command=self._schedule_dump_to_log,
            text_color=TEXT_COLOR,
            width=110,
        ).pack(side=tk.LEFT)

    def _show_schedule_spinner(self) -> None:
        frame = getattr(self, "_schedule_list", None)  # None until the Schedule tab UI is built
        if frame is None:
            return
        for child in list(frame.winfo_children()):
            try:
                child.destroy()
            except Exception as exc:
                APP_LOG.warning("Could not show the schedule spinner UI element.", exc_info=exc)
        spinner_holder = ctk.CTkFrame(frame, fg_color=BACKGROUND_COLOR)
        spinner_holder.pack(fill=tk.BOTH, expand=True)
        spinner_holder.grid_columnconfigure(0, weight=1)
        spinner_holder.grid_rowconfigure(0, weight=1)
        canvas = tk.Canvas(
            spinner_holder,
            width=48,
            height=48,
            bg=BACKGROUND_COLOR,
            highlightthickness=0,
        )
        canvas.grid(row=0, column=0)
        self._schedule_spinner_canvas = canvas
        self._schedule_spinner_angle = 0
        self._spin_schedule_spinner()

    def _spin_schedule_spinner(self) -> None:
        canvas = getattr(self, "_schedule_spinner_canvas", None)  # None until the Schedule spinner is shown
        if canvas is None:
            return
        canvas.delete("all")
        pad = 6
        size = 48
        start = self._schedule_spinner_angle
        extent = 300
        canvas.create_arc(
            pad,
            pad,
            size - pad,
            size - pad,
            start=start,
            extent=extent,
            style="arc",
            outline=SECONDARY_TEXT_COLOR,
            width=4,
        )
        self._schedule_spinner_angle = (start + 15) % 360
        self._schedule_spinner_job = self.after(60, self._spin_schedule_spinner)

    def _hide_schedule_spinner(self) -> None:
        job = getattr(self, "_schedule_spinner_job", None)  # None when no Schedule spinner animation is running
        if job is not None:
            try:
                self.after_cancel(job)
            except Exception as exc:
                APP_LOG.warning("Could not hide or tear down the schedule spinner UI element.", exc_info=exc)
        self._schedule_spinner_job = None
        canvas = getattr(self, "_schedule_spinner_canvas", None)  # None if the Schedule spinner was never shown
        if canvas is not None:
            try:
                parent = canvas.master
                canvas.destroy()
                if parent is not None:
                    parent.destroy()
            except Exception as exc:
                APP_LOG.warning("Could not hide or tear down the schedule spinner UI element.", exc_info=exc)
        self._schedule_spinner_canvas = None

    def _render_schedule_items(self):
        self._hide_schedule_spinner()
        frame = getattr(self, "_schedule_list", None)  # None until the Schedule tab UI is built
        if frame is None:
            return
        for child in list(frame.winfo_children()):
            try:
                child.destroy()
            except Exception as exc:
                APP_LOG.warning("Non-critical operation failed while running _render_schedule_items.", exc_info=exc)
        items = getattr(self, "_schedule_items", [])
        window = getattr(self, "_schedule_window", (None, None))
        status_parts = []
        if items:
            status_parts.append(f"{len(items)} item(s)")
        start, end = window
        if start and end:
            status_parts.append(f"{start.strftime('%b %d %I:%M %p')} → {end.strftime('%b %d %I:%M %p')} CST")
        if self._schedule_status_label is not None:
            self._schedule_status_label.configure(text=" · ".join(status_parts) if status_parts else "No items")

        if not items:
            ctk.CTkLabel(
                frame,
                text="No classes found in the current window.",
                text_color=SECONDARY_TEXT_COLOR,
            ).pack(anchor="w", padx=PAD_SM + 2, pady=PAD_SM + 2)
            return
        try:
            frame.grid_columnconfigure((0, 1, 2, 3), weight=1)
        except Exception as exc:
            APP_LOG.warning("Non-critical operation failed while running _render_schedule_items.", exc_info=exc)

        def loc_bucket(loc: str) -> str:
            loc_lower = (loc or "").lower()
            if "family" in loc_lower:
                return "family"
            if "chapel" in loc_lower:
                return "chapel"
            return "other"

        # Group occurrences by time slot (start/end) to preserve separation while laying out columns
        groups: Dict[Tuple[dt.datetime, dt.datetime], Dict[str, List[Dict[str, object]]]] = {}
        ordered_keys: List[Tuple[dt.datetime, dt.datetime]] = []
        for occ in items:
            key = (occ["start"], occ["end"])
            if key not in groups:
                groups[key] = {"family": [], "chapel": [], "other": []}
                ordered_keys.append(key)
            bucket = loc_bucket(str(occ.get("location") or ""))
            groups[key][bucket].append(occ)

        ordered_keys.sort(key=lambda k: (k[0], k[1]))

        for key in ordered_keys:
            slot = groups[key]
            start_dt, end_dt = key
            slot_label = f"{start_dt.strftime('%a %b %d, %I:%M %p')} - {end_dt.strftime('%I:%M %p')} CST"

            slot_frame = ctk.CTkFrame(frame, fg_color=SURFACE_CARD_COLOR, corner_radius=10)
            slot_frame.pack(fill=tk.X, expand=True, padx=PAD_SM, pady=PAD_SM + 2)
            try:
                slot_frame.grid_columnconfigure(0, weight=0)
                slot_frame.grid_columnconfigure(1, weight=1, minsize=220, uniform="schedcols")
                slot_frame.grid_columnconfigure(2, weight=0)
                slot_frame.grid_columnconfigure(3, weight=1, minsize=220, uniform="schedcols")
                slot_frame.grid_rowconfigure(1, weight=1)
            except Exception as exc:
                APP_LOG.warning("Could not configure the schedule slot layout row.", exc_info=exc)

            ctk.CTkLabel(
                slot_frame,
                text=slot_label,
                text_color=TEXT_COLOR,
                font=self.FONT_CARD,
                anchor="w",
            ).grid(row=0, column=0, columnspan=4, sticky="w", padx=PAD_MD + 2, pady=(6, 4))

            def render_column(col_idx: int, occs: List[Dict[str, object]]):
                col_frame = ctk.CTkFrame(slot_frame, fg_color=SURFACE_CARD_COLOR)
                col_frame.grid(row=1, column=col_idx, sticky="nsew", padx=PAD_SM + 2, pady=(0, 8))
                col_frame.grid_columnconfigure(0, weight=1)
                if not occs:
                    # Add a spacer to preserve layout/height even when no class in this column.
                    spacer = ctk.CTkFrame(col_frame, fg_color=SURFACE_CARD_COLOR, height=20)
                    spacer.pack(fill=tk.X, expand=True, pady=PAD_SM)
                    try:
                        spacer.pack_propagate(False)
                    except Exception as exc:
                        APP_LOG.warning("Could not preserve schedule column spacer sizing.", exc_info=exc)
                    return
                for occ_item in occs:
                    self._render_schedule_card(col_frame, occ_item)

            family_only_combined = bool(slot["family"]) and not bool(slot["chapel"])
            if family_only_combined:
                ctk.CTkLabel(
                    slot_frame,
                    text="Combined",
                    text_color=SECONDARY_TEXT_COLOR,
                    font=self.FONT_CARD_SMALL,
                    fg_color=SURFACE_DEEP_COLOR,
                    corner_radius=6,
                    padx=PAD_MD,
                    pady=2,
                ).grid(row=0, column=3, sticky="e", padx=(0, PAD_MD), pady=(6, 4))
                combined_frame = ctk.CTkFrame(slot_frame, fg_color=SURFACE_CARD_COLOR)
                combined_frame.grid(row=1, column=1, columnspan=3, sticky="nsew", padx=PAD_SM + 2, pady=(0, 8))
                combined_frame.grid_columnconfigure(0, weight=1)
                for occ_item in slot["family"]:
                    self._render_schedule_card(combined_frame, occ_item, centered=True, wrap_len=760)
            else:
                render_column(1, slot["family"])
                # Visual separator between columns (shorter height to reduce visual noise)
                ctk.CTkFrame(slot_frame, fg_color=SEPARATOR_COLOR, width=2, height=40).grid(
                    row=1, column=2, sticky="ns", pady=(4, 8)
                )
                render_column(3, slot["chapel"])

            # If there are locations that don't match either bucket, render them under the time label spanning both columns
            if slot["other"]:
                other_frame = ctk.CTkFrame(slot_frame, fg_color=SURFACE_DEEP_COLOR, corner_radius=8)
                other_frame.grid(row=2, column=0, columnspan=4, sticky="ew", padx=PAD_MD, pady=(4, 8))
                other_frame.grid_columnconfigure(0, weight=1)
                ctk.CTkLabel(
                    other_frame,
                    text="Other Locations",
                    text_color=SECONDARY_TEXT_COLOR,
                    font=self.FONT_CARD_SMALL,
                    anchor="w",
                ).grid(row=0, column=0, sticky="w", padx=PAD_MD, pady=(6, 2))
                row_idx = 1
                for occ_item in slot["other"]:
                    self._render_schedule_card(other_frame, occ_item, inset=True, use_grid=True, grid_row=row_idx)
                    row_idx += 1

    def _schedule_dump_to_log(self):
        items = getattr(self, "_schedule_items", [])
        if not items:
            self.debug("Schedule dump: no items")
            return
        self.debug(f"Schedule dump ({len(items)} items):")
        for occ in items:
            title = occ.get("title") or "Untitled Class"
            class_name = occ.get("class_name") or ""
            course_name = occ.get("course_name") or ""
            loc = occ.get("location") or "No location"
            ctype = occ.get("class_type") or ""
            when = f"{occ['start'].strftime('%a %b %d %I:%M %p')} - {occ['end'].strftime('%I:%M %p')} CST"
            raw = occ.get("start_raw") or ""
            self.debug(f"- {title} | {class_name} ({course_name}) | {loc} | {when} | raw={raw} | {ctype}")


    def _build_presets_tab(self, parent):
        try:
            parent.configure(fg_color=BACKGROUND_COLOR)
        except Exception as exc:
            APP_LOG.warning("Could not finish building the presets tab UI section.", exc_info=exc)
        parent.grid_columnconfigure((0, 1), weight=1)
        parent.grid_rowconfigure(3, weight=1)
        info = ctk.CTkLabel(
            parent,
            text="Manage dropdown presets for Class Title and Instructor. Select an entry and click Delete to remove typos.",
            text_color=TEXT_COLOR,
            wraplength=720,
            justify="left",
        )
        info.grid(row=0, column=0, columnspan=2, sticky="w", padx=PAD_MD + 2, pady=(10, 6))

        def build_panel(col: int, title: str, key: str):
            frame = ctk.CTkFrame(parent, fg_color=BACKGROUND_COLOR)
            frame.grid(row=1, column=col, sticky="nsew", padx=PAD_MD + 2, pady=(0, 10))
            frame.grid_rowconfigure(1, weight=1)
            ctk.CTkLabel(frame, text=title, text_color=TEXT_COLOR).grid(row=0, column=0, sticky="w", pady=(0, 6))
            rows = ctk.CTkScrollableFrame(frame, fg_color=BACKGROUND_COLOR, height=280)
            rows.grid(row=1, column=0, sticky="nsew")
            rows.grid_columnconfigure(0, weight=1)
            self._presets_frames[key] = rows
            btn = ctk.CTkButton(
                frame,
                text="Delete Selected",
                command=lambda k=key: self._presets_delete_selected(k),
                text_color=TEXT_COLOR,
                fg_color="#D43E27",
                hover_color="#B7321F",
            )
            btn.grid(row=2, column=0, sticky="ew", pady=(6, 0))
            return rows

        self._presets_class_list = build_panel(0, "Class Title Presets", "class_titles")
        self._presets_instr_list = build_panel(1, "Instructor Presets", "instructors")
        self._presets_status_label = ctk.CTkLabel(parent, text="", text_color=TEXT_COLOR)
        self._presets_status_label.grid(row=2, column=0, columnspan=2, sticky="w", padx=PAD_MD + 2, pady=(0, 6))
        spacer = ctk.CTkFrame(parent, fg_color="transparent")
        spacer.grid(row=3, column=0, columnspan=2, sticky="nsew")
        self._presets_refresh_lists()

    def _presets_select_row(self, key: str, value: str) -> None:
        self._presets_selected[key] = value
        rows = self._presets_rows.get(key, {})
        for row_val, row in rows.items():
            color = ACTIVE_TAB_COLOR if row_val == value else SURFACE_COLOR
            row.configure(fg_color=color)

    def _presets_refresh_lists(self):
        data = load_presets()
        self._apply_preset_dropdown_values(data)
        class_vals = data.get("class_titles", [])
        instr_vals = data.get("instructors", [])
        values_by_key = {
            "class_titles": class_vals,
            "instructors": instr_vals,
        }
        for key, values in values_by_key.items():
            frame = self._presets_frames.get(key)
            if frame is None:
                continue
            for child in list(frame.winfo_children()):
                child.destroy()
            self._presets_rows[key] = {}
            self._presets_selected[key] = None
            for val in values:
                row = ctk.CTkFrame(frame, fg_color=SURFACE_COLOR, corner_radius=6)
                row.pack(fill=tk.X, expand=True, padx=PAD_SM, pady=(0, PAD_SM))
                lbl = ctk.CTkLabel(row, text=val, text_color=TEXT_COLOR, anchor="w")
                lbl.pack(fill=tk.X, padx=PAD_MD, pady=6)
                row.bind("<Button-1>", lambda _e, k=key, v=val: self._presets_select_row(k, v))
                lbl.bind("<Button-1>", lambda _e, k=key, v=val: self._presets_select_row(k, v))
                self._presets_rows[key][val] = row
        if self._presets_status_label is not None:
            self._presets_status_label.configure(text="")

    def _presets_delete_selected(self, key: str):
        label = "Class Title" if key == "class_titles" else "Instructor"
        value = str(self._presets_selected.get(key) or "").strip()
        if not value:
            messagebox.showinfo("Delete Preset", f"Select a {label} preset to delete.")
            return
        if not messagebox.askyesno("Delete Preset", f"Remove '{value}' from {label} presets?"):
            return
        data = load_presets()
        values = [str(x).strip() for x in data.get(key, []) if isinstance(x, str)]
        if value not in values:
            messagebox.showinfo("Delete Preset", f"'{value}' is already removed.")
            self._presets_refresh_lists()
            return
        data[key] = [x for x in values if x != value]
        save_presets(data)
        self._presets_refresh_lists()
        if key == "class_titles" and self._class_var.get().strip() == value:
            self._class_var.set("")
        if key == "instructors" and self._instr_var.get().strip() == value:
            self._instr_var.set("")
        if self._presets_status_label is not None:
            self._presets_status_label.configure(text=f"Deleted '{value}' from {label} presets.")

    def _apply_settings_with_status(
        self,
        apply_fn: Callable[[], None],
        status_label: Optional[ctk.CTkLabel],
        success_message: str,
    ) -> None:
        try:
            apply_fn()
            if status_label is not None:
                status_label.configure(text=success_message)
        except Exception as e:
            if status_label is not None:
                status_label.configure(text=f"Error: {e}")

    def _check_for_updates_manual(self) -> None:
        status_lbl = getattr(self, "_general_update_status_label", None)  # None until General tab is built
        btn = getattr(self, "_general_update_btn", None)  # None until General tab is built
        if not UPDATE_METADATA_URL:
            msg = "Update check is not configured yet (missing update_metadata_url in app_config.json)."
            if status_lbl is not None:
                status_lbl.configure(text=msg)
            messagebox.showinfo("Check for Updates", msg)
            return

        def work():
            return _fetch_update_metadata(UPDATE_METADATA_URL)

        def on_done(result):
            if not isinstance(result, dict):
                if status_lbl is not None:
                    status_lbl.configure(text="Update check failed: invalid metadata response.")
                messagebox.showerror("Check for Updates", "Received an invalid update response.")
                return

            latest_version = str(result.get("latest_version") or "").strip() or "(unknown)"
            notes = str(result.get("notes") or "").strip()
            download_url = str(result.get("download_url_macos") or "").strip()
            changelog_url = str(result.get("changelog_url") or "").strip()
            release_date = str(result.get("release_date") or "").strip()

            try:
                update_available = _is_update_available(result)
            except Exception as exc:
                if status_lbl is not None:
                    status_lbl.configure(text=f"Update check failed: {exc}")
                messagebox.showerror("Check for Updates", f"Could not compare versions.\n\n{exc}")
                return

            if not update_available:
                if status_lbl is not None:
                    status_lbl.configure(text=f"Up to date (current v{__version__})")
                messagebox.showinfo("Check for Updates", f"You're up to date.\n\nCurrent version: v{__version__}")
                return

            if status_lbl is not None:
                status_lbl.configure(text=f"Update available: v{latest_version}")

            lines = [
                f"Current version: v{__version__}",
                f"Latest version: v{latest_version}",
            ]
            if release_date:
                lines.append(f"Release date: {release_date}")
            if notes:
                lines.append("")
                lines.append(notes)
            if not download_url:
                lines.append("")
                lines.append("No macOS download link is configured in version.json yet.")

            open_download = False
            if download_url:
                open_download = messagebox.askyesno(
                    "Update Available",
                    "\n".join(lines) + "\n\nOpen the download page now?",
                )
                if open_download:
                    webbrowser.open(download_url)
                    return
            else:
                messagebox.showinfo("Update Available", "\n".join(lines))

            if changelog_url and messagebox.askyesno(
                "View Changes",
                "Open the online changelog?",
            ):
                webbrowser.open(changelog_url)

        def on_error(err: Exception):
            if status_lbl is not None:
                status_lbl.configure(text=f"Update check failed: {err}")
            messagebox.showerror(
                "Check for Updates",
                "Could not check for updates.\n\n"
                f"{err}\n\n"
                "This usually means the update URL is missing, the internet is unavailable, or the hosted version.json is invalid.",
            )

        self._run_job(
            name="check-updates",
            work=work,
            disable=[btn] if btn is not None else None,
            status_widget=status_lbl,
            start_text="Checking for updates…",
            done_text=None,
            on_done=on_done,
            on_error=on_error,
        )

    def _apply_general_settings(self):
        def apply():
            self._general_settings["start_in_preview"] = bool(self._general_start_preview_var.get())
            self._general_settings["auto_pulse_after_launch"] = bool(self._general_auto_pulse_var.get())
            self._general_settings["stay_on_top"] = bool(self._general_topmost_var.get())
            self._general_settings["show_activity_log"] = bool(self._general_show_log_var.get())

            try:
                self.preview_mode.set(self._general_settings["start_in_preview"])
            except Exception as exc:
                APP_LOG.warning("Could not apply settings changes to the app state.", exc_info=exc)

            if self._general_settings["auto_pulse_after_launch"]:
                if not self._pulse_active:
                    self._pulse_toggle()
            else:
                if self._pulse_active:
                    self._stop_pulse()

            try:
                self.attributes("-topmost", bool(self._general_settings["stay_on_top"]))
            except Exception as exc:
                APP_LOG.warning("Could not apply settings changes to the app state.", exc_info=exc)
            self._set_activity_log_visible(self._general_settings["show_activity_log"])
            self._save_general_settings()

        self._apply_settings_with_status(
            apply, getattr(self, "_general_status_label", None), "General settings applied."  # Label is None until General tab is built
        )

        self._apply_settings_with_status(apply, getattr(self, "_general_status_label", None), "General settings applied.")  # Label is None until General tab is built

    def _load_general_settings(self) -> Dict[str, bool]:
        defaults = {
            "start_in_preview": bool(self.preview_mode.get()),
            "auto_pulse_after_launch": False,
            "stay_on_top": False,
            "show_activity_log": True,
        }
        try:
            if os.path.exists(GENERAL_SETTINGS_FILE):
                with open(GENERAL_SETTINGS_FILE, "r", encoding="utf-8") as f:
                    data = json.load(f)
                if isinstance(data, dict):
                    defaults.update(
                        {
                            "start_in_preview": bool(data.get("start_in_preview", defaults["start_in_preview"])),
                            "auto_pulse_after_launch": bool(
                                data.get("auto_pulse_after_launch", defaults["auto_pulse_after_launch"])
                            ),
                            "stay_on_top": bool(data.get("stay_on_top", defaults["stay_on_top"])),
                            "show_activity_log": bool(data.get("show_activity_log", defaults["show_activity_log"])),
                        }
                    )
        except Exception as exc:
            APP_LOG.warning("Non-critical operation failed while running _load_general_settings.", exc_info=exc)
        return defaults

    def _save_general_settings(self) -> None:
        try:
            os.makedirs(APP_DATA_DIR, exist_ok=True)
            with open(GENERAL_SETTINGS_FILE, "w", encoding="utf-8") as f:
                json.dump(self._general_settings, f, indent=2)
        except Exception as exc:
            APP_LOG.warning("Non-critical operation failed while running _save_general_settings.", exc_info=exc)

    def _set_activity_log_visible(self, show: bool) -> None:
        frame = getattr(self, "_activity_frame", None)  # None until the activity log UI is built
        if frame is None:
            return
        try:
            frame.pack_forget()
        except Exception as exc:
            APP_LOG.warning("Could not apply activity log visible state.", exc_info=exc)
        if show:
            try:
                frame.pack(fill=tk.BOTH, expand=False, padx=PAD_MD + 2, pady=(0, 10))
            except Exception as exc:
                APP_LOG.warning("Could not apply activity log visible state.", exc_info=exc)

    def _load_dreamclass_settings(self) -> Dict[str, str]:
        defaults = {"api_key": "", "school_code": "", "tenant": "", "enabled": True}
        try:
            if os.path.exists(DREAMCLASS_SETTINGS_FILE):
                with open(DREAMCLASS_SETTINGS_FILE, "r", encoding="utf-8") as f:
                    data = json.load(f)
                if isinstance(data, dict):
                    defaults.update(
                        {
                            "api_key": str(data.get("api_key", "")).strip(),
                            "school_code": str(data.get("school_code", "")).strip(),
                            "tenant": str(data.get("tenant", "")).strip(),
                            "enabled": bool(data.get("enabled", defaults["enabled"])),
                        }
                    )
        except Exception as exc:
            APP_LOG.warning("Non-critical operation failed while running _load_dreamclass_settings.", exc_info=exc)
        return defaults

    def _save_dreamclass_settings(self) -> None:
        def apply():
            data = {
                "api_key": str(self._dreamclass_api_key_var.get()).strip(),
                "school_code": str(self._dreamclass_school_code_var.get()).strip(),
                "tenant": str(self._dreamclass_tenant_var.get()).strip(),
                "enabled": bool(self._dreamclass_enabled_var.get()),
            }
            os.makedirs(APP_DATA_DIR, exist_ok=True)
            with open(DREAMCLASS_SETTINGS_FILE, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=2)
            self._dreamclass_settings = data
            self._dreamclass_enabled = bool(data.get("enabled", True))
            self._apply_dreamclass_enabled()

        self._apply_settings_with_status(
            apply, getattr(self, "_dreamclass_status_label", None), "DreamClass settings saved."  # Label is None until DreamClass tab is built
        )

    def _chan_refresh_table(self):
        # Repopulate tree from self.channels
        tree = getattr(self, "chan_tree", None)  # None until the Channels settings tab is opened
        if tree is None:
            return
        try:
            for iid in tree.get_children():
                tree.delete(iid)
        except Exception:
            return
        for ch in self.channels:
            tree.insert(
                "",
                tk.END,
                values=(
                    ch.get("label", ""),
                    ch.get("name", ""),
                    ch.get("channel_id", ""),
                    ch.get("playlist_id", ""),
                    ch.get("thumbnail_template_file", ""),
                    "Yes" if ch.get("show_in_monitor", True) else "No",
                ),
            )
        self._chan_resize_columns()
        self._chan_refresh_monitor_checks()

    def _chan_refresh_monitor_checks(self):
        frame = getattr(self, "_chan_monitor_scroll", None)  # None until the Channels settings tab is opened
        if frame is None:
            return
        for child in list(frame.winfo_children()):
            try:
                child.destroy()
            except Exception as exc:
                APP_LOG.warning("Could not update Channels settings UI (refresh monitor checks).", exc_info=exc)
        self._chan_monitor_vars = {}
        max_cols = 3
        for col in range(max_cols):
            try:
                frame.grid_columnconfigure(col, weight=1)
            except Exception as exc:
                APP_LOG.warning("Could not update Channels settings UI (refresh monitor checks).", exc_info=exc)
        for idx, ch in enumerate(self.channels):
            cid = ch.get("channel_id")
            if not cid:
                continue
            label = ch.get("label") or ch.get("name") or cid[:6]
            var = tk.BooleanVar(value=bool(ch.get("show_in_monitor", True)))
            self._chan_monitor_vars[cid] = var
            cb = ctk.CTkCheckBox(
                frame,
                text=str(label),
                variable=var,
                text_color=TEXT_COLOR,
                command=lambda c=cid, v=var: self._chan_on_monitor_checkbox_toggle(c, v),
            )
            row = idx // max_cols
            col = idx % max_cols
            cb.grid(row=row, column=col, sticky="w", padx=PAD_SM + 2, pady=PAD_SM)

    def _chan_on_monitor_checkbox_toggle(self, channel_id: str, var: tk.BooleanVar):
        show = bool(var.get())
        for ch in self.channels:
            if ch.get("channel_id") == channel_id:
                ch["show_in_monitor"] = show
                if not show:
                    ch["monitor"] = False
                break
        self._chan_update_tree_show_value(channel_id, show)

    def _chan_update_tree_show_value(self, channel_id: str, show: bool):
        tree = getattr(self, "chan_tree", None)  # None until the Channels settings tab is opened
        if tree is None:
            return
        for iid in tree.get_children():
            vals = tree.item(iid, "values")
            if vals and len(vals) >= 6 and vals[2] == channel_id:
                new_vals = list(vals)
                new_vals[5] = "Yes" if show else "No"
                tree.item(iid, values=new_vals)
                break

    def _chan_reset_form(self):
        self.var_label.set("")
        self.var_name.set("")
        self.var_cid.set("")
        self.var_pid.set("")
        self.var_thumbnail_template.set("")

    def _chan_browse_thumbnail_template(self):
        chosen = filedialog.askopenfilename(
            title="Select Thumbnail Template",
            filetypes=[("Image Files", "*.jpg *.jpeg *.png *.webp"), ("All Files", "*.*")],
        )
        if not chosen:
            return
        self.var_thumbnail_template.set(_normalize_thumbnail_template_path(chosen))

    def _chan_add_channel(self):
        label = (self.var_label.get() or "").strip()
        name = (self.var_name.get() or "").strip()
        cid = (self.var_cid.get() or "").strip()
        pid = (self.var_pid.get() or "").strip()
        thumb_tpl = _normalize_thumbnail_template_path((self.var_thumbnail_template.get() or "").strip())
        if not name or not cid:
            messagebox.showerror("Missing", "Name and Channel ID are required")
            return
        updated = False
        for ch in self.channels:
            if ch.get("channel_id") == cid:
                ch.update(
                    {
                        "label": label or ch.get("label", ""),
                        "name": name,
                        "playlist_id": pid,
                        "thumbnail_template_file": thumb_tpl,
                    }
                )
                updated = True
                break
        if not updated:
            key = (label or name or cid).replace(" ", "_")
            self.channels.append(
                {
                    "key": key,
                    "label": label or key,
                    "name": name,
                    "channel_id": cid,
                    "playlist_id": pid,
                    "thumbnail_template_file": thumb_tpl,
                    "monitor": False,
                    "show_in_monitor": True,
                }
            )
        self._chan_refresh_table()
        self._rebuild_header_indicators()
        self.debug(f"Channels: {'updated' if updated else 'added'} {name} [{cid}]")
        self._chan_reset_form()

    def _chan_delete_selected(self):
        sel = self.chan_tree.selection()
        if not sel:
            return
        to_delete = []
        for iid in sel:
            vals = self.chan_tree.item(iid, "values")
            if vals and len(vals) >= 3:
                to_delete.append(vals[2])  # channel_id
        if not to_delete:
            return
        self.channels = [c for c in self.channels if c.get("channel_id") not in to_delete]
        self._chan_refresh_table()
        self._rebuild_header_indicators()
        self.debug(f"Channels: deleted {len(to_delete)} channel(s)")

    def _chan_save_apply(self):
        save_channels_config(self.channels)
        new_list = load_channels_config()
        self._apply_channels_update(new_list)
        self._chan_refresh_table()
        try:
            self._update_thumbnail_preview()
        except Exception as exc:
            APP_LOG.warning("Failed to refresh thumbnail preview after saving channels.", exc_info=exc)
        self.debug("Channels: saved and applied")

    def _iter_monitored_channels(self):
        for ch in self.channels:
            cid = ch.get("channel_id")
            if not cid:
                continue
            if not ch.get("show_in_monitor", True):
                continue
            var = self._monitor_vars.get(cid)
            if var is not None:
                if bool(var.get()):
                    yield ch
            elif bool(ch.get("monitor", False)):
                yield ch

    def _get_monitored_channels(self) -> List[Dict[str, str]]:
        return list(self._iter_monitored_channels())

    def _has_monitored_channels(self) -> bool:
        return next(self._iter_monitored_channels(), None) is not None

    def _chan_resize_columns(self):
        tree = getattr(self, "chan_tree", None)  # None until the Channels settings tab is opened
        if tree is None:
            return
        try:
            font = tkfont.nametofont(tree.cget("font"))
        except Exception:
            font = tkfont.nametofont("TkDefaultFont")
        padding = 24
        for col in ("label", "name", "channel_id", "playlist_id", "thumbnail_template", "show_monitor"):
            header = col.replace("_", " ").title()
            max_width = font.measure(header)
            for iid in tree.get_children():
                text = tree.set(iid, col)
                max_width = max(max_width, font.measure(str(text)))
            width = max_width + padding
            tree.column(col, width=int(width), stretch=False)

def main():
    app = SLSRobusterApp()
    app.mainloop()


if __name__ == "__main__":
    main()
