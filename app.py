import os
import sqlite3
import time
import json
import re
import html
from datetime import datetime
from functools import wraps
from concurrent.futures import ThreadPoolExecutor, as_completed
from flask import Flask, render_template, request, jsonify, redirect, url_for, session, send_file
from authlib.integrations.flask_client import OAuth
import io
import requests
import openpyxl
from dotenv import load_dotenv

import db
from scim import scim_bp

load_dotenv(override=True)
db.init_db()

app = Flask(__name__)
app.secret_key = os.getenv("FLASK_SECRET_KEY", os.urandom(24))

# --- Okta / OIDC config ---
OKTA_DOMAIN        = os.getenv("OKTA_DOMAIN")       # e.g. https://yourcompany.okta.com
OKTA_CLIENT_ID     = os.getenv("OKTA_CLIENT_ID")
OKTA_CLIENT_SECRET = os.getenv("OKTA_CLIENT_SECRET")
if not all([OKTA_DOMAIN, OKTA_CLIENT_ID, OKTA_CLIENT_SECRET]):
    raise RuntimeError("Set OKTA_DOMAIN, OKTA_CLIENT_ID, and OKTA_CLIENT_SECRET in .env")

oauth = OAuth(app)
oauth.register(
    name="okta",
    client_id=OKTA_CLIENT_ID,
    client_secret=OKTA_CLIENT_SECRET,
    server_metadata_url=f"{OKTA_DOMAIN}/.well-known/openid-configuration",
    client_kwargs={
        "scope": "openid email profile groups",
    },
)

# --- RocketLane config ---
ROCKETLANE_API_KEY      = os.getenv("ROCKETLANE_API_KEY")
ROCKETLANE_API_BASE_URL = os.getenv("ROCKETLANE_API_BASE_URL", "https://api.rocketlane.com/api")
# Web-app host (not the API host) — used to build project deep-links for the consultant view.
ROCKETLANE_DOMAIN       = os.getenv("ROCKETLANE_DOMAIN", "digicert-ps.rocketlane.com")

# Salesforce Lightning host — used to build deep-links from project rows to the linked
# Rocketlane_Project__c record. Override via env if your SF org uses a different domain.
SALESFORCE_LIGHTNING_DOMAIN = os.getenv("SALESFORCE_LIGHTNING_DOMAIN", "digicert.lightning.force.com")

if not ROCKETLANE_API_KEY:
    raise RuntimeError("Set ROCKETLANE_API_KEY in .env before running")

HEADERS = {
    "api-key": ROCKETLANE_API_KEY,
    "accept": "application/json",
    "content-type": "application/json",
}

PRIORITY_FIELD_ID = 2098590  # SINGLE_CHOICE; values: 1=High, 2=Low, 3=Urgent
PRIORITY_VALUE_BY_LABEL = {"High": 1, "Low": 2, "Urgent": 3}
PRIORITY_LABEL_BY_VALUE = {v: k for k, v in PRIORITY_VALUE_BY_LABEL.items()}

ACV_FIELD_ID = 2595804  # NUMBER; Annual Contract Value (kept for reference, not displayed)
# ARR is a native field returned as item["annualizedRecurringRevenue"], not in fields[].

# Column keys for the manager projects table (order = default layout). "name" is locked first.
PROJECT_COLUMN_KEYS = ("name", "links", "package", "arr", "segment", "delta_price",
                       "sub_region", "close_date", "csm", "sdp", "loe", "priority", "due_date",
                       "weekly_status")
# Each dashboard view keeps its own column layout.
PROJECT_VIEW_KEYS = ("missing", "sdp", "internal", "readyForKickoff", "kickoffScheduled",
                     "onHold", "blockedCustomer", "blockedDigiCert")
# "Weekly Note (AI)" (MULTI_LINE_TEXT, 2630517) — writable field; an automation propagates its
# value to the read-only "Weekly Status Update" Note (1769450) whenever it changes.
CONSULTANT_UPDATE_FIELD_ID = 2630517

# Customer Provisioning Form (PKI/DC1) contact fields
TECH_CONTACT_NAME_FIELD_ID  = 1640037  # TEXT: Technical Contact Name
TECH_CONTACT_EMAIL_FIELD_ID = 1640038  # TEXT: Technical Contact Email Address

# --- Level of Effort (LOE) details ------------------------------------------
# The LOE data is a large block of project custom fields (2599xxx) grouped into
# RocketLane "field sections". "Level of Effort Product" is the discriminator; each
# activity is a TEXT field (value is a count, 0/blank = not in scope) immediately
# followed (by fieldId, within its section) by its "... Details" MULTI_LINE_TEXT field.
LOE_PRODUCT_FIELD_ID = 2599268                       # SINGLE_SELECT: 1=TLM, 2=STM, 3=TLM & STM
LOE_PRODUCT_LABELS   = {1: "TLM", 2: "STM", 3: "TLM & STM"}
LOE_STATUS_FIELD_ID  = 2599267                       # Level of Effort Status
LOE_HOURS_FIELD_IDS  = {"tlm": 2599269, "stm": 2599270, "total": 2599271, "planning": 2599272}
# fieldSectionId -> display name. These two sections hold the activity/detail line items.
LOE_SECTIONS = {192251: "TLM Level of Effort (LOE)", 192812: "STM Level of Effort (LOE)"}
# Which section a product value maps to.
LOE_PRODUCT_SECTIONS = {1: (192251,), 2: (192812,), 3: (192251, 192812)}

# Project assignment constants
PS_CONSULTANT_ROLE_ID        = 133617                 # PS Consultant (unified role)
DEPARTMENT_FIELD_ID          = 1617964                # Department custom field on users
# SDP vs Internal is now determined by Department starting with "SDP" (not by role).
PMO_ROLE_ID                  = 128021                 # PMO
READY_FOR_KICKOFF_STATUS     = 11                     # top-level project status value
# Saba Matiwane is the AMS regional placeholder owner auto-assigned by Salesforce.
# She must always be removed from a project when it is assigned to a real consultant.

# Consultant view: which project statuses a consultant's own work spans, and who may see the view.
# 2=In Progress, 9=On Hold, 11=Ready for Kick-Off, 15=Blocked-Waiting on Customer,
# 16=Blocked-Waiting on DigiCert, 18=Kick-off Meeting Scheduled.
CONSULTANT_VIEW_STATUSES = (2, 9, 11, 15, 16, 18)
CONSULTANT_VIEW_EXCEPTION_EMAIL = "myles.huffman@okta.com"   # testing exception (non-Consultant access)

# (value, display label), ordered for the consultant Status dropdown. The 6 view statuses + Closed.
# Setting Closed (10) drops the project from the view on next refresh (10 not in the fetched set).
CONSULTANT_STATUS_OPTIONS = [
    (2,  "In Progress"),
    (18, "Kick-off Meeting Scheduled"),
    (11, "Ready for Kick-Off"),
    (9,  "On Hold"),
    (15, "Blocked - Waiting on Customer"),
    (16, "Blocked - Waiting on DigiCert"),
    (10, "Closed"),
]
CONSULTANT_STATUS_VALUES = {v for v, _ in CONSULTANT_STATUS_OPTIONS}
CONSULTANT_STATUS_LABELS = {v: lbl for v, lbl in CONSULTANT_STATUS_OPTIONS}

MANAGER_STATUS_OPTIONS = [
    (5,  "Proposed"),
    (13, "Not Required"),
    (11, "Ready for Kick-Off"),
    (18, "Kick-off Meeting Scheduled"),
    (2,  "In Progress"),
]
MANAGER_STATUS_VALUES = {v for v, _ in MANAGER_STATUS_OPTIONS}
MANAGER_STATUS_LABELS = {v: lbl for v, lbl in MANAGER_STATUS_OPTIONS}

DELIVERY_STATUS_FIELD_ID = 2644033
DELIVERY_STATUS_VALUES = {
    "Active":                        1,
    "Blocked - Waiting on Customer": 2,
    "Blocked - Waiting on DigiCert": 3,
    "On-Hold":                       4,
}


def can_view_consultant(user):
    """A user may see the Consultant view if they're a Consultant or the testing exception."""
    return (user.get("role") == "Consultant"
            or (user.get("email") or "").lower() == CONSULTANT_VIEW_EXCEPTION_EMAIL)


def can_impersonate_consultant(user):
    """'View As' is a testing tool — only Managers and the testing exception may use it."""
    return (user.get("role") == "Manager"
            or (user.get("email") or "").lower() == CONSULTANT_VIEW_EXCEPTION_EMAIL)

# Region taxonomy. RocketLane stores the top-level region in field 1611022 (Account Region) and
# the sub-region in field 1611023 (Account Sub-Region). We filter on sub-region values because
# EMEA and APJ sub-region labels don't share a common prefix the way AMS ones do — substring
# matching against the top-level field wouldn't work for them.
REGIONS = ("AMS", "EMEA", "APJ")
DEFAULT_REGION = "AMS"
REGION_SUB_REGIONS = {
    "AMS":  ("AMS Cent", "AMS East", "AMS South East", "AMS West", "LATAM"),
    "EMEA": ("Africa", "Benelux", "DACH", "Eastern Europe",
             "Middle East", "Nordics", "SW Europe", "UKI"),
    "APJ":  ("ANZ", "ASEAN", "India"),   # GCR and Japan are out-of-scope for this team
}


def normalize_region(value):
    """Return a valid region from the input, falling back to the default."""
    if value in REGIONS:
        return value
    return DEFAULT_REGION


# Reverse map of sub-region label -> top-level region, for views that span all regions.
_REGION_BY_SUB_REGION = {sr: region for region, subs in REGION_SUB_REGIONS.items() for sr in subs}


def region_for_sub_region(sub_region):
    """Best-effort top-level region for a sub-region label (prefix match), else '-'."""
    if not sub_region:
        return "-"
    if sub_region in _REGION_BY_SUB_REGION:
        return _REGION_BY_SUB_REGION[sub_region]
    for sr, region in _REGION_BY_SUB_REGION.items():
        if sub_region.startswith(sr):
            return region
    return "-"
SDP_STATUS_FIELD_ID = 2098584
SDP_STATUS_MAP = {
    "ok":  2,  # Approved
    "no":  5,  # Consent Required
}


# ---------------------------------------------------------------------------
# Auth helpers
# ---------------------------------------------------------------------------

def render_unauthorized(email=None):
    """Render the Access Denied screen. Used both at the SSO boundary and by require_role."""
    return render_template("unauthorized.html", user={"email": email or "(unknown)"}), 403


def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if "user" not in session:
            return redirect(url_for("login", next=request.url))
        # Re-validate against the DB each request — kills stale sessions when a Manager
        # deactivates a user or changes their role mid-session. SQLite lookup is cheap.
        u = db.user_by_id(session["user"].get("id"))
        if u is None or not u["active"]:
            email = session["user"].get("email")
            session.clear()
            return render_unauthorized(email)
        session["user"]["role"] = u["role"]
        session["user"]["name"] = u["name"]
        return f(*args, **kwargs)
    return decorated


def require_role(*roles):
    """Decorator: require the session user to have one of the given roles."""
    def deco(f):
        @wraps(f)
        @login_required
        def wrapped(*args, **kwargs):
            if session["user"]["role"] not in roles:
                return render_unauthorized(session["user"].get("email"))
            return f(*args, **kwargs)
        return wrapped
    return deco


# ---------------------------------------------------------------------------
# Auth routes
# ---------------------------------------------------------------------------

@app.route("/login")
def login():
    redirect_uri = url_for("auth_callback", _external=True)
    return oauth.okta.authorize_redirect(redirect_uri)


@app.route("/logged-out")
def logged_out():
    return render_template("logged_out.html")


@app.route("/authorization-code/callback")
def auth_callback():
    token    = oauth.okta.authorize_access_token()
    userinfo = token.get("userinfo") or oauth.okta.userinfo()
    email = (userinfo.get("email") or "").strip()
    if not email:
        return render_unauthorized(None)

    u = db.user_by_email(email)
    if u is None or not u["active"]:
        # Unknown or deactivated user — Access Denied. A Manager must create them first.
        return render_unauthorized(email)

    session["user"] = {
        "id":    u["id"],
        "email": u["email"],
        "name":  u["name"],
        "role":  u["role"],
    }
    session["id_token"] = token.get("id_token")
    if u["role"] == "Consultant":
        return redirect(url_for("consultant_home"))
    next_url = request.args.get("next") or url_for("index")
    return redirect(next_url)


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("logged_out"))


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def format_date(value):
    if not value:
        return "-"
    try:
        dt = datetime.fromisoformat(value.rstrip("Z"))
        return dt.strftime("%Y-%m-%d")
    except ValueError:
        return value


def _weekly_status_parts(raw, preview_limit=100):
    """Parse a RocketLane NOTE field into (preview, full_text, date).

    Converts block-level HTML into newline-separated paragraphs, returns the first paragraph
    as the preview, all paragraphs joined as full_text, and the date extracted from the
    [Author · date]: prefix (empty string if not present).
    """
    if not raw:
        return ("-", "", "")
    text = re.sub(r"</(p|div|li|h[1-6])>", "\n", raw, flags=re.IGNORECASE)
    text = re.sub(r"<br\s*/?>", "\n", text, flags=re.IGNORECASE)
    text = re.sub(r"<[^>]+>", "", text)
    text = html.unescape(text)
    paragraphs = [p.strip() for p in re.split(r"\n{2,}", text) if p.strip()]
    if not paragraphs:
        paragraphs = [ln.strip() for ln in text.split("\n") if ln.strip()]
    if not paragraphs:
        return ("-", "", "")
    preview = paragraphs[0]
    if len(preview) > preview_limit:
        preview = preview[:preview_limit].rstrip() + "…"
    # Extract date from "[Author · 10 Sep 2026]: ..." prefix in the first paragraph.
    date_match = re.match(r"^\[.*?·\s*(.+?)\]:", paragraphs[0])
    ws_date = date_match.group(1).strip() if date_match else ""
    return (preview, "\n\n".join(paragraphs), ws_date)


def _html_to_text(raw, limit=280):
    """Flatten a RocketLane NOTE field's HTML into a trimmed, single-line plain-text preview."""
    if not raw:
        return "-"
    text = re.sub(r"<[^>]+>", " ", raw)      # drop tags
    text = html.unescape(text)               # &nbsp; etc.
    text = re.sub(r"\s+", " ", text).strip()
    if not text:
        return "-"
    return text if len(text) <= limit else text[:limit].rstrip() + "…"


def get_custom_field(item, field_id):
    for field in item.get("fields", []):
        if field.get("fieldId") == field_id:
            return field.get("fieldValue")
    return None


def get_custom_field_label(item, field_id):
    for field in item.get("fields", []):
        if field.get("fieldId") == field_id:
            return field.get("fieldValueLabel")
    return None


_LOE_CATALOG_CACHE = {}


def _loe_catalog():
    """Build (and cache) the LOE activity catalog from the field definitions.

    Returns {section_id: {"activities": [(act_id, act_label, det_id)], "narratives": [(fid, label)]}}.
    Pairing rule (validated against live data): within a section, walking fields in fieldId order,
    a TEXT field immediately followed by a MULTI_LINE_TEXT field is an (Activity, Details) pair;
    anything else (Project Overview, SE POV Activities Performed) is a narrative. RocketLane omits
    empty fields per project, so this full catalog is what gives us reliable pairing/labels.

    Cached after the first call; empty dict on API failure (endpoint degrades to header-only).
    """
    if _LOE_CATALOG_CACHE:
        return _LOE_CATALOG_CACHE
    try:
        r = requests.get(f"{ROCKETLANE_API_BASE_URL}/v1/fields", headers=HEADERS, timeout=30)
        r.raise_for_status()
        fields = r.json()
    except requests.exceptions.RequestException as exc:
        print(f"[loe] field catalog fetch failed: {exc}")
        return {}

    catalog = {}
    for sec_id in LOE_SECTIONS:
        sec = sorted(
            (f for f in fields
             if f.get("objectType") == "PROJECT" and f.get("fieldSectionId") == sec_id),
            key=lambda f: f.get("fieldId", 0),
        )
        activities, narratives = [], []
        i = 0
        while i < len(sec):
            f = sec[i]
            nxt = sec[i + 1] if i + 1 < len(sec) else None
            if (f.get("fieldDataType") == "TEXT" and nxt is not None
                    and nxt.get("fieldDataType") == "MULTI_LINE_TEXT"):
                activities.append((f["fieldId"], f.get("fieldLabel") or "", nxt["fieldId"]))
                i += 2
            else:
                narratives.append((f["fieldId"], f.get("fieldLabel") or ""))
                i += 1
        catalog[sec_id] = {"activities": activities, "narratives": narratives}

    _LOE_CATALOG_CACHE.update(catalog)
    return _LOE_CATALOG_CACHE


# Per-region cache. RocketLane's projects endpoint has been unreliable (30s+ per page); we cache
# successful fetches so reloads don't re-pay the cost. Each region has its own cache slot so users
# can switch between AMS / EMEA / APJ without invalidating each other. The cache is persisted to
# disk as a region-keyed JSON so a Flask restart still warms all regions. Add to .gitignore.
_PROJECTS_CACHE_BY_REGION = {r: {"data": None, "fetched_at": 0.0} for r in REGIONS}
_PROJECTS_CACHE_TTL_SECONDS = 300   # 5 minutes
_PROJECTS_CACHE_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".projects_cache.json")

# Consultant view pulls a broader, all-region set (more statuses, no region split). One in-memory
# slot is enough: we fetch everything once, then filter per consultant in the request.
_CONSULTANT_CACHE = {"data": None, "fetched_at": 0.0}


def _load_disk_cache():
    try:
        with open(_PROJECTS_CACHE_FILE) as f:
            blob = json.load(f)
    except FileNotFoundError:
        return
    except (json.JSONDecodeError, OSError) as exc:
        print(f"[fetch_projects] disk cache load failed: {exc}")
        return
    if not isinstance(blob, dict):
        return
    for region in REGIONS:
        slot = blob.get(region)
        if isinstance(slot, dict) and isinstance(slot.get("data"), list):
            _PROJECTS_CACHE_BY_REGION[region]["data"]       = slot["data"]
            _PROJECTS_CACHE_BY_REGION[region]["fetched_at"] = float(slot.get("fetched_at") or 0)
            print(f"[fetch_projects] loaded {len(slot['data'])} {region} projects from disk cache "
                  f"(fetched_at={slot.get('fetched_at')})")


def _save_disk_cache():
    try:
        with open(_PROJECTS_CACHE_FILE, "w") as f:
            json.dump(_PROJECTS_CACHE_BY_REGION, f)
    except OSError as exc:
        print(f"[fetch_projects] disk cache save failed: {exc}")


_load_disk_cache()


def fetch_projects(region=DEFAULT_REGION, force_refresh=False):
    """Fetch Closed-Won PS projects for the given top-level region (AMS / EMEA / APJ).

    Cached per-region for 5 minutes. On a fetch failure, the most recent cached value for that
    region (even if expired) is returned if available; otherwise the error propagates.

    Strategy under slow RL API (probed: ~28s per limit=25 page, 30s timeouts at limit≥50):
      - small pages (limit=25) — proven to come back inside reasonable time
      - 90s per-page timeout (was 45s — pages legitimately take ~30s right now)
      - 1 retry on transient failure
      - sub-regions for the requested region fetched in parallel
    """
    region = normalize_region(region)
    sub_regions = REGION_SUB_REGIONS[region]

    now = time.time()
    cache_slot = _PROJECTS_CACHE_BY_REGION[region]
    cached = cache_slot["data"]
    age    = now - cache_slot["fetched_at"]
    if not force_refresh and cached is not None and age < _PROJECTS_CACHE_TTL_SECONDS:
        print(f"[fetch_projects] returning cached {region} data (age {int(age)}s)")
        return cached

    # Use the tenant subdomain for the projects list — empirically ~35% faster than the generic
    # api.rocketlane.com host for this query, which matters when RL is slow enough that the generic
    # host's upstream gateway 504s before our request returns.
    endpoint = "https://digicert-ps.api.rocketlane.com/api/1.0/projects"
    base_params_ps = {
        "limit.eq":         25,
        "includeAllFields": "true",  # needed for annualizedRecurringRevenue
    }
    STATUSES_TO_FETCH = (2, 5, 9, 11, 13, 15, 16, 18)  # In Progress + Proposed + Not Required + On Hold + Ready for Kick-Off + Blocked x2 + Kick-off Meeting Scheduled

    def get_page(params):
        import time as _time
        last_exc = None
        for attempt in range(4):
            try:
                resp = requests.get(endpoint, headers=HEADERS, params=params, timeout=90)
                if resp.status_code == 429:
                    wait = int(resp.headers.get("X-Rate-Limit-Minute-Reset", 15)) + 1
                    print(f"[fetch_projects] 429 rate-limited; waiting {wait}s (attempt {attempt + 1})")
                    _time.sleep(wait)
                    last_exc = requests.exceptions.HTTPError(f"429 rate limit", response=resp)
                    continue
                resp.raise_for_status()
                return resp.json()
            except requests.exceptions.RequestException as exc:
                last_exc = exc
                print(f"[fetch_projects] page attempt {attempt + 1} failed: {exc}")
                if attempt < 3:
                    _time.sleep(2 ** attempt)  # 1s, 2s, 4s backoff for non-429 errors
        raise last_exc

    def fetch_slice(status, sub_region, base_params):
        items = []
        page_token = None
        for _ in range(100):
            params = {
                **base_params,
                "status.eq":                       status,
                "project.field.1611023.contains":  sub_region,
            }
            if page_token:
                params["pageToken"] = page_token
            data = get_page(params)
            items.extend(data.get("data", data.get("projects", [])))
            pagination = data.get("pagination") or {}
            if not pagination.get("hasMore") or not pagination.get("nextPageToken"):
                break
            page_token = pagination["nextPageToken"]
        return items

    slices = [(s, sr, base_params_ps) for s in STATUSES_TO_FETCH for sr in sub_regions]
    raw_items = []
    errors    = []
    # Cap at 4 workers — RocketLane's API is rate-limited to ~100 req/min and bursting
    # too many concurrent slices reliably triggers 429s.
    with ThreadPoolExecutor(max_workers=4) as pool:
        futures = {pool.submit(fetch_slice, s, sr, bp): (s, sr) for (s, sr, bp) in slices}
        for fut in as_completed(futures):
            status, sr = futures[fut]
            try:
                raw_items.extend(fut.result())
            except Exception as exc:
                print(f"[fetch_projects] {region} status={status} {sr} fetch failed entirely: {exc}")
                errors.append(f"status={status} {sr}: {exc}")

    if not raw_items:
        if cached is not None:
            print(f"[fetch_projects] all {region} fetches failed; serving stale cache (age {int(age)}s)")
            return cached
        raise requests.exceptions.RequestException("; ".join(errors) or "Unknown fetch failure")

    # Post-filter to this region. Projects with no sub-region set are included in all regions.
    def item_sub_region(item):
        return get_custom_field_label(item, 1611023)

    raw_items = [
        item for item in raw_items
        if not item_sub_region(item) or item_sub_region(item) in sub_regions
    ]

    projects = []

    for item in raw_items:
        project_name      = item.get("projectName") or item.get("name") or "Untitled"
        close_date        = format_date(item.get("dueDate") or item.get("dueDateActual"))
        csm_handover_date = format_date(get_custom_field(item, 1925432))
        package_name      = get_custom_field(item, 1621438) or "-"
        priority          = get_custom_field_label(item, PRIORITY_FIELD_ID) or "-"
        sdp_status        = get_custom_field_label(item, 2098584) or "-"
        sdp_consent_required = get_custom_field_label(item, 2587103)
        # SDP Consent Required supersedes SDP Status when populated:
        # "No" ≡ Approved, "Yes" ≡ Consent Required.
        if sdp_consent_required == "No":
            sdp_status = "Approved"
        elif sdp_consent_required == "Yes":
            sdp_status = "Consent Required"
        loe_status              = get_custom_field_label(item, 2098585) or "-"
        level_of_effort_status  = get_custom_field_label(item, 2599267) or "-"
        delivery_status         = get_custom_field_label(item, DELIVERY_STATUS_FIELD_ID) or ""
        # A Draft Level of Effort is a strong signal that an LOE exists on the opportunity —
        # surface it on the LOE Status column even when the LOE Status field itself is blank
        # or carries an older value. A Draft is not Uploaded, so prereq_ok still fails.
        if level_of_effort_status == "Draft":
            loe_status = "Draft"
        # Three display states: Complete (Uploaded or Complete), Draft, or None (not clickable).
        if (loe_status == "Uploaded") or (level_of_effort_status == "Complete"):
            loe_display = "Complete"
        elif loe_status == "Draft":
            loe_display = "Draft"
        else:
            loe_display = "-"
        status_label            = (item.get("status") or {}).get("label") or "-"
        segment                 = get_custom_field_label(item, 2600949) or "-"
        salesforce_id           = get_custom_field(item, 1614439) or ""
        salesforce_url          = (
            f"https://{SALESFORCE_LIGHTNING_DOMAIN}/lightning/r/Rocketlane_Project__c/{salesforce_id}/view"
            if salesforce_id else ""
        )
        arr_value = item.get("annualizedRecurringRevenue")
        if arr_value is None:
            arr_display = "-"
        else:
            arr_display = f"${int(round(arr_value)):,}"

        # A project is ready to be assigned only when all prerequisites are met.
        loe_ok = loe_display == "Complete"
        csm_ok = bool(csm_handover_date) and csm_handover_date != "-"
        sdp_ok = sdp_status in ("Approved", "Consent Required")
        prerequisites_ok = loe_ok and csm_ok and sdp_ok
        # Internal assignments don't require SDP approval — mirrors the Internal Assignment
        # view's SDP-agnostic filter, so the Assign button appears there too.
        internal_ready = loe_ok and csm_ok
        ws_field = next((f for f in item.get("fields", []) if f.get("fieldId") == 1769450), None)
        ws_raw = ws_field.get("fieldValue", "") if ws_field else ""
        ws_preview, ws_full, ws_date = _weekly_status_parts(ws_raw)

        projects.append({
            "id":                     item.get("projectId"),
            "name":                   project_name,
            "package_name":           package_name,
            "close_date":             close_date,
            "csm_handover_date":      csm_handover_date,
            "priority":               priority,
            "sdp_status":             sdp_status,
            "sdp_consent_required":   sdp_consent_required or "-",
            "loe_status":             loe_status,
            "loe_display":            loe_display,
            "level_of_effort_status": level_of_effort_status,
            "status":                 status_label,
            "status_value":           (item.get("status") or {}).get("value"),
            "delivery_status":        delivery_status,
            "prerequisites_ok":       prerequisites_ok,
            "internal_ready":         internal_ready,
            "archived":               bool(item.get("archived", False)),
            "segment":                segment,
            "delta_price":            get_custom_field(item, 1697223) or "-",       # Delta Price (Hourly), TEXT
            "sub_region":             get_custom_field_label(item, 1611023) or "-",  # Account Sub-Region
            "arr":                    arr_display,
            "arr_value":              arr_value if arr_value is not None else 0,
            "due_date":               close_date,
            "weekly_status":          ws_preview,
            "weekly_status_full":     ws_full,
            "weekly_status_date":     ws_date,
            "salesforce_id":          salesforce_id,
            "salesforce_url":         salesforce_url,
            "rocketlane_url":         f"https://{ROCKETLANE_DOMAIN}/projects/{item.get('projectId')}/overview",
            "has_loe":                bool(get_custom_field(item, LOE_PRODUCT_FIELD_ID)),
        })

    print(f"[fetch_projects] {region} total returned from API: {len(projects)}")
    seen = set()
    deduped = []
    for p in projects:
        if p["id"] not in seen:
            seen.add(p["id"])
            deduped.append(p)
    if len(deduped) != len(projects):
        print(f"[fetch_projects] {region} duplicates removed: {len(projects) - len(deduped)}")
    projects = deduped
    priority_rank = {"Urgent": 0, "High": 1, "Low": 2}
    projects.sort(key=lambda p: (
        priority_rank.get(p["priority"], 3),
        p["csm_handover_date"] if p["csm_handover_date"] != "-" else "9999-99-99",
    ))

    # Only cache a fully-successful fetch — partial data (one sub-region failed) shouldn't overwrite
    # a complete earlier cache. Stale-complete beats fresh-partial for the dashboard.
    if not errors:
        cache_slot["data"]       = projects
        cache_slot["fetched_at"] = time.time()
        _save_disk_cache()
    else:
        print(f"[fetch_projects] {region} partial fetch ({len(projects)} projects, errors: {errors}); cache not updated")
    return projects


def _person(p):
    """Normalize a RocketLane owner/member object to {userId, email, name}."""
    if not p:
        return {"userId": None, "email": "", "name": ""}
    return {
        "userId": p.get("userId"),
        "email":  (p.get("emailId") or "").lower(),
        "name":   f"{p.get('firstName','')} {p.get('lastName','')}".strip(),
    }


def fetch_consultant_projects(force_refresh=False):
    """All PS projects across every region in the consultant-relevant statuses.

    Returns a list of lightweight dicts carrying owner/member identity so the request can filter
    to a single consultant. Cached in-memory for 5 minutes; serves stale on a fetch failure.
    """
    now = time.time()
    cached = _CONSULTANT_CACHE["data"]
    age    = now - _CONSULTANT_CACHE["fetched_at"]
    if not force_refresh and cached is not None and age < _PROJECTS_CACHE_TTL_SECONDS:
        print(f"[fetch_consultant_projects] returning cached data (age {int(age)}s)")
        return cached

    endpoint = "https://digicert-ps.api.rocketlane.com/api/1.0/projects"
    base_params = {
        "limit.eq": 25,
    }

    def get_page(params):
        last_exc = None
        for attempt in range(2):
            try:
                resp = requests.get(endpoint, headers=HEADERS, params=params, timeout=90)
                resp.raise_for_status()
                return resp.json()
            except requests.exceptions.RequestException as exc:
                last_exc = exc
                print(f"[fetch_consultant_projects] page attempt {attempt + 1} failed: {exc}")
        raise last_exc

    def fetch_status(status):
        # No region/sub-region filter — one paginated query per status spans all regions.
        items = []
        page_token = None
        for _ in range(200):
            params = {**base_params, "status.eq": status}
            if page_token:
                params["pageToken"] = page_token
            data = get_page(params)
            items.extend(data.get("data", data.get("projects", [])))
            pagination = data.get("pagination") or {}
            if not pagination.get("hasMore") or not pagination.get("nextPageToken"):
                break
            page_token = pagination["nextPageToken"]
        return items

    raw_items = []
    errors    = []
    with ThreadPoolExecutor(max_workers=len(CONSULTANT_VIEW_STATUSES)) as pool:
        futures = {pool.submit(fetch_status, s): s for s in CONSULTANT_VIEW_STATUSES}
        for fut in as_completed(futures):
            status = futures[fut]
            try:
                raw_items.extend(fut.result())
            except Exception as exc:
                print(f"[fetch_consultant_projects] status={status} fetch failed entirely: {exc}")
                errors.append(f"status={status}: {exc}")

    if not raw_items:
        if cached is not None:
            print(f"[fetch_consultant_projects] all fetches failed; serving stale cache (age {int(age)}s)")
            return cached
        raise requests.exceptions.RequestException("; ".join(errors) or "Unknown fetch failure")

    seen = set()
    projects = []
    for item in raw_items:
        pid = item.get("projectId")
        if pid is None or pid in seen:
            continue
        seen.add(pid)

        owner   = _person(item.get("owner"))
        members = [_person(m) for m in ((item.get("teamMembers") or {}).get("members") or [])]
        salesforce_id  = get_custom_field(item, 1614439) or ""
        salesforce_url = (
            f"https://{SALESFORCE_LIGHTNING_DOMAIN}/lightning/r/Rocketlane_Project__c/{salesforce_id}/view"
            if salesforce_id else ""
        )
        projects.append({
            "id":             pid,
            "name":           item.get("projectName") or item.get("name") or "Untitled",
            "status":         (item.get("status") or {}).get("label") or "-",
            "status_value":   (item.get("status") or {}).get("value"),
            "start_date":     format_date(item.get("startDate")),
            "due_date":       format_date(item.get("dueDate")),
            "weekly_status":  _html_to_text(get_custom_field(item, 1769450)),
            "consultant_update": (get_custom_field(item, CONSULTANT_UPDATE_FIELD_ID) or "").strip(),
            "delivery_status": get_custom_field_label(item, DELIVERY_STATUS_FIELD_ID) or "",
            "remaining_hours": None,   # enriched per-project from /v1 after filtering (see route)
            "shared_docs":    None,    # enriched per-project from /v1 spaceList + /v1/spaces (see route)
            "region":         region_for_sub_region(get_custom_field_label(item, 1611023)),
            "salesforce_url": salesforce_url,
            "rocketlane_url": f"https://{ROCKETLANE_DOMAIN}/projects/{pid}/overview",
            "owner":          owner,
            "owner_name":     owner["name"] or "-",
            "members":        members,
        })

    projects.sort(key=lambda p: p["name"].lower())
    # Only cache a fully-successful fetch — partial data shouldn't overwrite a complete earlier cache.
    if not errors:
        _CONSULTANT_CACHE["data"]       = projects
        _CONSULTANT_CACHE["fetched_at"] = time.time()
    else:
        print(f"[fetch_consultant_projects] partial fetch ({len(projects)} projects, errors: {errors}); cache not updated")
    return projects


def _matches_consultant(person, target):
    """True if a RocketLane person is the target consultant (by userId, else by email)."""
    if target.get("userId") is not None and person.get("userId") is not None:
        return person["userId"] == target["userId"]
    return bool(target.get("email")) and person.get("email") == target["email"]


def _enrich_project_details(projects):
    """Per-project /v1 enrichment: remaining_hours (metrics) and shared_docs ("Shared space" tabs).

    Bounded to the (already filtered) displayed set. Runs in parallel; leaves defaults on failure.
    """
    todo = [p for p in projects if p.get("remaining_hours") is None and p.get("shared_docs") is None]
    if not todo:
        return

    def fetch_one(p):
        hrs, docs = None, []
        try:
            r = requests.get(f"{ROCKETLANE_API_BASE_URL}/v1/projects/{p['id']}",
                             headers=HEADERS, timeout=30)
            r.raise_for_status()
            body = r.json()
            metrics = body.get("metrics") or {}
            hrs = metrics.get("remainingHours")
            if hrs is None and body.get("remainingMinutes") is not None:
                hrs = body["remainingMinutes"] / 60.0
            hrs = round(hrs, 2) if isinstance(hrs, (int, float)) else None
            # Documents live in the "Shared space" tabs — one extra call, isolated so a space
            # failure never blanks the hours.
            shared = next((s for s in (body.get("spaceList") or [])
                           if s.get("spaceName") == "Shared space"), None)
            if shared and shared.get("spaceId"):
                try:
                    sr = requests.get(f"{ROCKETLANE_API_BASE_URL}/v1/spaces/{shared['spaceId']}",
                                      headers=HEADERS, timeout=30)
                    sr.raise_for_status()
                    for t in (sr.json().get("spaceTabs") or []):
                        if t.get("url"):
                            docs.append({"title": t.get("spaceTabTitle") or "Untitled",
                                         "url": f"https://{ROCKETLANE_DOMAIN}/{t['url']}"})
                except requests.exceptions.RequestException as exc:
                    print(f"[consultant] shared-space fetch failed for {p['id']}: {exc}")
        except requests.exceptions.RequestException as exc:
            print(f"[consultant] detail fetch failed for {p['id']}: {exc}")
        return p, hrs, docs

    with ThreadPoolExecutor(max_workers=8) as pool:
        for p, hrs, docs in pool.map(fetch_one, todo):
            p["remaining_hours"] = hrs
            p["shared_docs"] = docs


# ---------------------------------------------------------------------------
# App routes
# ---------------------------------------------------------------------------

def _tile_stats(projects):
    """Per-filter project counts and summed ARR for the KPI tiles.

    Predicates mirror the tile→preset filters. Each project is counted once per tile
    (Missing Prerequisites sums unique projects — a project missing LOE+SDP+CSM counts once).
    """
    keys = ("missing", "sdp", "internal", "ready", "kos", "onhold", "blocked_customer", "blocked_digicert")
    s = {k: {"count": 0, "arr": 0} for k in keys}
    for p in projects:
        arr          = p.get("arr_value") or 0
        status       = p.get("status")
        proposed     = status == "Proposed"
        csm_ok       = bool(p.get("csm_handover_date")) and p.get("csm_handover_date") != "-"
        loe_ok       = p.get("loe_display") == "Complete"
        sdp_approved = (p.get("sdp_consent_required") == "No") or (p.get("sdp_status") == "Approved")
        if proposed and not p.get("prerequisites_ok"):
            s["missing"]["count"] += 1; s["missing"]["arr"] += arr
        if proposed and csm_ok and loe_ok and sdp_approved:
            s["sdp"]["count"] += 1; s["sdp"]["arr"] += arr
        if proposed and csm_ok and loe_ok:
            s["internal"]["count"] += 1; s["internal"]["arr"] += arr
        if status == "Ready for Kick-Off":
            s["ready"]["count"] += 1; s["ready"]["arr"] += arr
        if status == "Kick-off Meeting Scheduled":
            s["kos"]["count"] += 1; s["kos"]["arr"] += arr
        ds = p.get("delivery_status", "")
        if ds == "On-Hold" or status == "On Hold":
            s["onhold"]["count"] += 1; s["onhold"]["arr"] += arr
        if ds == "Blocked - Waiting on Customer" or status == "Blocked - Waiting on Customer":
            s["blocked_customer"]["count"] += 1; s["blocked_customer"]["arr"] += arr
        if ds == "Blocked - Waiting on DigiCert" or status == "Blocked - Waiting on DigiCert":
            s["blocked_digicert"]["count"] += 1; s["blocked_digicert"]["arr"] += arr
    for k in keys:
        s[k]["arr_display"] = f"${int(round(s[k]['arr'])):,}"
    return s


@app.route("/")
@login_required
def index():
    user = session["user"]
    if user["role"] == "Consultant":
        return redirect(url_for("consultant_home"))

    error    = None
    projects = []
    force_refresh = request.args.get("refresh") == "1"
    active_region = normalize_region(request.args.get("region"))
    try:
        projects = fetch_projects(region=active_region, force_refresh=force_refresh)
    except requests.exceptions.RequestException as exc:
        error = f"API request failed: {exc}"

    fetched_at = _PROJECTS_CACHE_BY_REGION[active_region]["fetched_at"] or 0
    raw_layout = (db.user_by_id(user["id"]) or {}).get("column_layout")
    try:
        column_layout = json.loads(raw_layout) if raw_layout else None
    except (ValueError, TypeError):
        column_layout = None
    if not isinstance(column_layout, dict):   # per-view map expected; ignore legacy/blank
        column_layout = None
    weekly_status_map = {
        p["id"]: {"full": p.get("weekly_status_full", ""), "date": p.get("weekly_status_date", "")}
        for p in projects
    }
    return render_template("index.html", projects=projects, error=error,
                           user=user, current_user=user,
                           active_region=active_region, regions=REGIONS,
                           fetched_at=fetched_at, tiles=_tile_stats(projects),
                           column_layout=column_layout,
                           weekly_status_map=weekly_status_map,
                           manager_status_options=MANAGER_STATUS_OPTIONS)


@app.route("/column-layout", methods=["POST"])
@login_required
def column_layout():
    """Persist the current user's column layout for ONE view.

    Body: {"view": <view-key>, "layout": [{"key","visible"}, ...]}. Layouts are stored per view
    as a {view: [...]} map. Unknown column keys are dropped; "name" is forced present, visible,
    and first (locked identifier column).
    """
    body = request.get_json(silent=True) or {}
    view = body.get("view")
    raw  = body.get("layout")
    if view not in PROJECT_VIEW_KEYS:
        return jsonify({"success": False, "error": "invalid view"}), 400
    if not isinstance(raw, list):
        return jsonify({"success": False, "error": "layout must be a list"}), 400

    seen, cleaned = set(), []
    for item in raw:
        if not isinstance(item, dict):
            continue
        key = item.get("key")
        if key in PROJECT_COLUMN_KEYS and key not in seen and key != "name":
            seen.add(key)
            cleaned.append({"key": key, "visible": bool(item.get("visible", True))})
    for key in PROJECT_COLUMN_KEYS:                       # append omitted/new columns
        if key not in seen and key != "name":
            cleaned.append({"key": key, "visible": True})
    cleaned.insert(0, {"key": "name", "visible": True})

    # Merge into the existing per-view map.
    raw_existing = (db.user_by_id(session["user"]["id"]) or {}).get("column_layout")
    try:
        layouts = json.loads(raw_existing) if raw_existing else {}
    except (ValueError, TypeError):
        layouts = {}
    if not isinstance(layouts, dict):
        layouts = {}
    layouts[view] = cleaned

    db.set_column_layout(session["user"]["id"], json.dumps(layouts))
    return jsonify({"success": True})


@app.route("/consultant-home")
@login_required
def consultant_home():
    user = session["user"]
    if not can_view_consultant(user):
        return redirect(url_for("index"))

    impersonate = can_impersonate_consultant(user)
    # Default target: the logged-in user, matched by email (real consultants).
    target = {"userId": None, "email": (user.get("email") or "").lower(), "name": user.get("name")}
    view_as = (request.args.get("view_as") or "").strip()
    if view_as and impersonate:
        try:
            roster = fetch_consultants()
        except requests.exceptions.RequestException:
            roster = []
        picked = next((c for c in roster if str(c["userId"]) == view_as), None)
        if picked:
            target = {"userId": picked["userId"],
                      "email": (picked.get("email") or "").lower(),
                      "name": picked.get("name")}
        else:
            # Unknown id selected — show nothing rather than the viewer's own (empty) projects.
            target = {"userId": None, "email": "", "name": None}

    error, owned, member = None, [], []
    # The exception viewer with no consultant selected has no projects of their own — skip the fetch.
    needs_fetch = bool(target.get("userId") or target.get("email"))
    if needs_fetch:
        try:
            all_projects = fetch_consultant_projects(force_refresh=request.args.get("refresh") == "1")
            for p in all_projects:
                if _matches_consultant(p["owner"], target):
                    owned.append(p)
                elif any(_matches_consultant(m, target) for m in p["members"]):
                    member.append(p)   # owner takes precedence; no double-listing
            # Remaining Hours + Shared-space docs come from the /v1 detail — enrich only what we'll display.
            _enrich_project_details(owned + member)
        except requests.exceptions.RequestException as exc:
            error = f"API request failed: {exc}"

    # Status is editable only for a real Consultant on their own projects (read-only while impersonating).
    status_editable = user.get("role") == "Consultant"
    return render_template("consultant_home.html", user=user, current_user=user,
                           owned=owned, member=member, target=target,
                           can_impersonate=impersonate, error=error,
                           status_editable=status_editable, status_options=CONSULTANT_STATUS_OPTIONS)


@app.route("/set-project-status/<int:project_id>", methods=["POST"])
@login_required
def set_project_status(project_id):
    """Let a consultant change the status of one of their own projects.

    Body: {"status_value": <int>}. Only real Consultants may write (read-only while impersonating),
    and only on projects they own or are a member of.
    """
    user = session["user"]
    if user.get("role") != "Consultant":
        return jsonify({"success": False, "error": "Not permitted"}), 403

    value = (request.get_json(silent=True) or {}).get("status_value")
    if value not in CONSULTANT_STATUS_VALUES:
        return jsonify({"success": False, "error": f"Invalid status {value!r}"}), 400

    # Ownership check — the project must be one of THIS consultant's (owner or member, by email).
    target = {"userId": None, "email": (user.get("email") or "").lower()}
    try:
        mine = {p["id"] for p in fetch_consultant_projects()
                if _matches_consultant(p["owner"], target)
                or any(_matches_consultant(m, target) for m in p["members"])}
    except requests.exceptions.RequestException as exc:
        return jsonify({"success": False, "error": f"Could not verify project: {exc}"}), 502
    if project_id not in mine:
        return jsonify({"success": False, "error": "Project not in your list"}), 403

    endpoint = f"{ROCKETLANE_API_BASE_URL}/1.0/projects/{project_id}"
    resp = None
    try:
        payload = {"status": {"value": value}}
        resp = requests.put(endpoint, headers=HEADERS, json=payload, timeout=30)
        print(f"[set-project-status] {project_id} -> {value} status={resp.status_code} body={resp.text[:300]}")
        resp.raise_for_status()
        # Reflect the new value in the cached project dict so a same-TTL reload stays consistent.
        for p in (_CONSULTANT_CACHE.get("data") or []):
            if p["id"] == project_id:
                p["status_value"] = value
                p["status"] = CONSULTANT_STATUS_LABELS[value]
                break
        return jsonify({"success": True, "status_value": value, "label": CONSULTANT_STATUS_LABELS[value]})
    except requests.exceptions.RequestException as exc:
        detail = resp.text if resp is not None else ""
        status_code = resp.status_code if resp is not None else 0
        print(f"[set-project-status] ERROR: {exc} | {detail}")
        if status_code == 404:
            return jsonify({"success": False, "error": "Project not found in RocketLane — it may have been archived or removed. Try refreshing."}), 404
        return jsonify({"success": False, "error": str(exc), "detail": detail}), 500


@app.route("/manager-set-status/<int:project_id>", methods=["POST"])
@require_role("Manager", "Project Manager")
def manager_set_status(project_id):
    """Let a Manager or Project Manager change the status of any tracked project."""
    value = (request.get_json(silent=True) or {}).get("status_value")
    if value not in MANAGER_STATUS_VALUES:
        return jsonify({"success": False, "error": f"Invalid status {value!r}"}), 400

    endpoint = f"{ROCKETLANE_API_BASE_URL}/1.0/projects/{project_id}"
    resp = None
    try:
        resp = requests.put(endpoint, headers=HEADERS, json={"status": {"value": value}}, timeout=30)
        print(f"[manager-set-status] {project_id} -> {value} http={resp.status_code}")
        resp.raise_for_status()
        return jsonify({"success": True, "status_value": value, "label": MANAGER_STATUS_LABELS[value]})
    except requests.exceptions.RequestException as exc:
        detail = resp.text if resp is not None else ""
        print(f"[manager-set-status] ERROR: {exc} | {detail}")
        return jsonify({"success": False, "error": str(exc), "detail": detail}), 502


@app.route("/manager/consultant-projects/<int:user_id>", methods=["GET"])
@require_role("Manager", "Project Manager")
def manager_consultant_projects(user_id):
    """Return one consultant's projects grouped by status, sourced from the consultant cache."""
    all_projects = fetch_consultant_projects()
    matched = [
        p for p in all_projects
        if (p.get("owner") or {}).get("userId") == user_id
        or any(m.get("userId") == user_id for m in (p.get("members") or []))
    ]
    STATUS_ORDER = [
        "In Progress", "Kick-off Meeting Scheduled", "Ready for Kick-Off",
        "On Hold", "Blocked - Waiting on Customer", "Blocked - Waiting on DigiCert",
    ]
    groups_map = {}
    for p in matched:
        groups_map.setdefault(p["status"], []).append(p)
    groups = []
    for s in STATUS_ORDER:
        if s in groups_map:
            groups.append({"status": s, "projects": groups_map[s]})
    for s, ps in groups_map.items():
        if s not in STATUS_ORDER:
            groups.append({"status": s, "projects": ps})
    return jsonify({"groups": groups})


@app.route("/update-delivery-status/<int:project_id>", methods=["POST"])
@login_required
def update_delivery_status(project_id):
    """Let any authenticated user set the Delivery Status custom field on a project."""
    label = ((request.get_json(silent=True) or {}).get("delivery_status") or "").strip()
    if label not in DELIVERY_STATUS_VALUES and label != "":
        return jsonify({"success": False, "error": f"Invalid delivery status {label!r}"}), 400

    option_value = DELIVERY_STATUS_VALUES.get(label)
    if option_value is None:
        # Clearing the field — RocketLane expects null for a SINGLE_CHOICE clear
        field_payload = {"fieldId": DELIVERY_STATUS_FIELD_ID, "fieldValue": None}
    else:
        field_payload = {"fieldId": DELIVERY_STATUS_FIELD_ID, "fieldValue": option_value}

    endpoint = f"{ROCKETLANE_API_BASE_URL}/1.0/projects/{project_id}"
    resp = None
    try:
        resp = requests.put(endpoint, headers=HEADERS,
                            json={"fields": [field_payload]}, timeout=30)
        print(f"[update-delivery-status] {project_id} -> {label!r} http={resp.status_code}")
        resp.raise_for_status()
        for region_slot in _PROJECTS_CACHE_BY_REGION.values():
            for p in (region_slot.get("data") or []):
                if p["id"] == project_id:
                    p["delivery_status"] = label
                    break
        for p in (_CONSULTANT_CACHE.get("data") or []):
            if p["id"] == project_id:
                p["delivery_status"] = label
                break
        return jsonify({"success": True, "delivery_status": label})
    except requests.exceptions.RequestException as exc:
        detail = resp.text if resp is not None else ""
        print(f"[update-delivery-status] ERROR: {exc} | {detail}")
        return jsonify({"success": False, "error": str(exc), "detail": detail}), 502


@app.route("/update-weekly-status/<int:project_id>", methods=["POST"])
@login_required
def update_weekly_status(project_id):
    """Append a dated, initialed entry to a project's Weekly Note (AI) field.

    Body: {"text": "<plain text>"}. Writes CONSULTANT_UPDATE_FIELD_ID (MULTI_LINE_TEXT) —
    the "Weekly Status Update" Note (1769450) is read-only via API; an automation propagates
    this field's value to it. Any authenticated user may update any project.
    """
    user = session["user"]

    text = ((request.get_json(silent=True) or {}).get("text") or "").strip()
    if not text:
        return jsonify({"success": False, "error": "Update text is required"}), 400

    name_parts = (user.get("name") or "").split()
    if len(name_parts) >= 2:
        author_label = f"{name_parts[0]} {name_parts[-1][0].upper()}"
    elif name_parts:
        author_label = name_parts[0]
    else:
        author_label = "?"

    endpoint = f"{ROCKETLANE_API_BASE_URL}/1.0/projects/{project_id}"
    resp = None
    try:
        today = datetime.now().strftime("%-d %b %Y")
        entry = f"[{author_label} · {today}]: {text}"
        new_value = entry

        payload = {"fields": [{"fieldId": CONSULTANT_UPDATE_FIELD_ID, "fieldValue": new_value}]}
        resp = requests.put(endpoint, headers=HEADERS, json=payload, timeout=30)
        print(f"[update-weekly-status] {project_id} status={resp.status_code} body={resp.text[:300]}")
        resp.raise_for_status()
        # Reflect the new value in all in-memory caches so a same-TTL reload stays consistent.
        for p in (_CONSULTANT_CACHE.get("data") or []):
            if p["id"] == project_id:
                p["consultant_update"] = new_value
                break
        preview = f"[{author_label} · {today}]: {text[:100]}"
        for region_slot in _PROJECTS_CACHE_BY_REGION.values():
            for p in (region_slot.get("data") or []):
                if p["id"] == project_id:
                    p["weekly_status"] = preview
                    p["weekly_status_full"] = new_value
                    p["weekly_status_date"] = today
                    break
        return jsonify({"success": True, "entry": entry})
    except requests.exceptions.RequestException as exc:
        detail = resp.text if resp is not None else ""
        status_code = resp.status_code if resp is not None else 0
        print(f"[update-weekly-status] ERROR: {exc} | {detail}")
        if status_code == 404:
            return jsonify({"success": False, "error": "Project not found in RocketLane — it may have been archived or removed. Try refreshing."}), 404
        return jsonify({"success": False, "error": str(exc), "detail": detail}), 500


@app.route("/project-weekly-status/<int:project_id>", methods=["GET"])
@login_required
def project_weekly_status(project_id):
    """Return the live value of the read-only Weekly Status Update field (1769450)."""
    try:
        r = requests.get(
            f"{ROCKETLANE_API_BASE_URL}/1.0/projects/{project_id}",
            headers=HEADERS,
            params={"includeAllFields": "true"},
            timeout=30,
        )
        r.raise_for_status()
        data = r.json()
        ws_field = next((f for f in data.get("fields", []) if f.get("fieldId") == 1769450), None)
        raw = ws_field.get("fieldValue", "") if ws_field else ""
        preview, full, ws_date = _weekly_status_parts(raw or "")
        return jsonify({"preview": preview, "full": full, "date": ws_date})
    except requests.exceptions.RequestException as exc:
        return jsonify({"error": str(exc)}), 502


@app.route("/set-priority/<int:project_id>", methods=["POST"])
@require_role("Manager", "Project Manager")
def set_priority(project_id):
    """Set or clear the Priority custom field on a project.

    Body: {"priority": "Low" | "High" | "Urgent" | ""}.
    An empty/None priority clears the field.
    """
    body  = request.get_json(silent=True) or {}
    label = (body.get("priority") or "").strip()

    if label and label not in PRIORITY_VALUE_BY_LABEL:
        return jsonify({"success": False, "error": f"Invalid priority {label!r}"}), 400

    field_value = PRIORITY_VALUE_BY_LABEL.get(label)   # None when label is "" → clears the field
    endpoint = f"{ROCKETLANE_API_BASE_URL}/1.0/projects/{project_id}"
    resp = None
    try:
        payload = {"fields": [{"fieldId": PRIORITY_FIELD_ID, "fieldValue": field_value}]}
        resp = requests.put(endpoint, headers=HEADERS, json=payload, timeout=30)
        print(f"[set-priority] {project_id} -> {label!r} status={resp.status_code} body={resp.text[:300]}")
        resp.raise_for_status()
        return jsonify({"success": True, "priority": label})
    except requests.exceptions.RequestException as exc:
        detail = resp.text if resp is not None else ""
        status_code = resp.status_code if resp is not None else 0
        print(f"[set-priority] ERROR: {exc} | {detail}")
        if status_code == 404:
            return jsonify({"success": False, "error": "Project not found in RocketLane — it may have been archived or removed. Try refreshing."}), 404
        return jsonify({"success": False, "error": str(exc), "detail": detail}), 500


@app.route("/set-csm-handover/<int:project_id>", methods=["POST"])
@require_role("Manager", "Project Manager")
def set_csm_handover(project_id):
    """Set the CSM Handover Date (custom field 1925432, DATE) on a project.

    Body: {"date": "YYYY-MM-DD"}. Used to fill in the date for backlog projects missing it.
    """
    date_str = ((request.get_json(silent=True) or {}).get("date") or "").strip()
    try:
        datetime.strptime(date_str, "%Y-%m-%d")
    except ValueError:
        return jsonify({"success": False, "error": "date must be YYYY-MM-DD"}), 400

    endpoint = f"{ROCKETLANE_API_BASE_URL}/1.0/projects/{project_id}"
    resp = None
    try:
        payload = {"fields": [{"fieldId": 1925432, "fieldValue": date_str}]}
        resp = requests.put(endpoint, headers=HEADERS, json=payload, timeout=30)
        print(f"[set-csm-handover] {project_id} -> {date_str} status={resp.status_code} body={resp.text[:300]}")
        resp.raise_for_status()
        return jsonify({"success": True, "date": format_date(date_str)})
    except requests.exceptions.RequestException as exc:
        detail = resp.text if resp is not None else ""
        status_code = resp.status_code if resp is not None else 0
        print(f"[set-csm-handover] ERROR: {exc} | {detail}")
        if status_code == 404:
            return jsonify({"success": False, "error": "Project not found in RocketLane — it may have been archived or removed. Try refreshing."}), 404
        return jsonify({"success": False, "error": str(exc), "detail": detail}), 500


def fetch_consultants():
    """Active TEAM_MEMBER users with the PS Consultant role, sorted by name.

    Each entry includes isSdp=True when the user's Department field starts with "SDP".
    Raises requests.exceptions.RequestException on API failure (callers decide how to surface it).
    """
    endpoint = f"{ROCKETLANE_API_BASE_URL}/1.0/users"
    seen = {}
    page_token = None
    while True:
        params = {"limit.eq": 250, "roleId.eq": PS_CONSULTANT_ROLE_ID,
                  "includeAllFields": "true"}
        if page_token:
            params["pageToken"] = page_token
        resp = requests.get(endpoint, headers=HEADERS, params=params, timeout=30)
        resp.raise_for_status()
        body = resp.json()
        for u in body.get("data", []):
            if u.get("status") != "ACTIVE" or u.get("type") != "TEAM_MEMBER":
                continue
            uid = u.get("userId")
            if uid is None:
                continue
            department = next(
                (f.get("fieldValueLabel") or "" for f in (u.get("fields") or [])
                 if f.get("fieldId") == DEPARTMENT_FIELD_ID),
                ""
            )
            seen[uid] = {
                "userId": uid,
                "name":   f"{u.get('firstName','')} {u.get('lastName','')}".strip(),
                "email":  u.get("email"),
                "isSdp":  department.upper().startswith("SDP"),
            }
        pg = body.get("pagination") or {}
        if not pg.get("hasMore") or not pg.get("nextPageToken"):
            break
        page_token = pg["nextPageToken"]
    return sorted(seen.values(), key=lambda u: u["name"].lower())


@app.route("/consultants", methods=["GET"])
@login_required
def consultants():
    """Active TEAM_MEMBER users with one of the three consultant roles."""
    try:
        result = fetch_consultants()
    except requests.exceptions.RequestException as exc:
        return jsonify({"error": f"Could not fetch consultants: {exc}"}), 502
    return jsonify({"consultants": result})


def fetch_pms():
    """Active TEAM_MEMBER users with the PMO role, sorted by name."""
    endpoint = f"{ROCKETLANE_API_BASE_URL}/1.0/users"
    pms = {}
    page_token = None
    while True:
        params = {"limit.eq": 250, "roleId.eq": PMO_ROLE_ID}
        if page_token:
            params["pageToken"] = page_token
        resp = requests.get(endpoint, headers=HEADERS, params=params, timeout=30)
        resp.raise_for_status()
        body = resp.json()
        for u in body.get("data", []):
            if u.get("status") != "ACTIVE" or u.get("type") != "TEAM_MEMBER":
                continue
            uid = u.get("userId")
            if uid is None or uid in pms:
                continue
            pms[uid] = {
                "userId":  uid,
                "name":    f"{u.get('firstName','')} {u.get('lastName','')}".strip(),
                "email":   u.get("email"),
                "roleIds": [PMO_ROLE_ID],
            }
        pg = body.get("pagination") or {}
        if not pg.get("hasMore") or not pg.get("nextPageToken"):
            break
        page_token = pg["nextPageToken"]
    return sorted(pms.values(), key=lambda u: u["name"].lower())


@app.route("/pms", methods=["GET"])
@login_required
def pms():
    """Active TEAM_MEMBER users with the PMO role."""
    try:
        result = fetch_pms()
    except requests.exceptions.RequestException as exc:
        return jsonify({"error": f"Could not fetch PMs: {exc}"}), 502
    return jsonify({"pms": result})


@app.route("/project-technical-contact/<int:project_id>", methods=["GET"])
@login_required
def project_technical_contact(project_id):
    """Return Technical Contact Name and Email from the Customer Provisioning Form (PKI/DC1)."""
    try:
        r = requests.get(f"{ROCKETLANE_API_BASE_URL}/1.0/projects/{project_id}",
                         headers=HEADERS, params={"includeAllFields": "true"}, timeout=30)
        r.raise_for_status()
        data = r.json()
        name  = (get_custom_field(data, TECH_CONTACT_NAME_FIELD_ID)  or "").strip()
        email = (get_custom_field(data, TECH_CONTACT_EMAIL_FIELD_ID) or "").strip()
        return jsonify({"name": name, "email": email})
    except requests.exceptions.RequestException as exc:
        return jsonify({"error": str(exc)}), 502


@app.route("/loe-details/<int:project_id>", methods=["GET"])
@login_required
def loe_details(project_id):
    """Structured Level of Effort detail for one project, for the LOE pop-up.

    Reads the project's fields, routes to the TLM/STM LOE section(s) by the "Level of Effort
    Product" value, and returns only the activities whose value is non-zero (with their details).
    """
    try:
        r = requests.get(f"{ROCKETLANE_API_BASE_URL}/v1/projects/{project_id}",
                         headers=HEADERS, timeout=30)
        r.raise_for_status()
        body = r.json()
    except requests.exceptions.RequestException as exc:
        print(f"[loe-details] fetch failed for {project_id}: {exc}")
        return jsonify({"error": f"Could not fetch LOE details: {exc}"}), 502

    vals = {f.get("fieldId"): f.get("fieldValue") for f in (body.get("fields") or [])}
    # Resolved choice labels (metaFieldValue.label) for SINGLE_SELECT fields like the status.
    # metaFieldValue is a dict for single-selects (and a list for multi-selects — ignore those).
    labels = {}
    for f in (body.get("fields") or []):
        mv = f.get("metaFieldValue")
        if isinstance(mv, dict):
            labels[f.get("fieldId")] = mv.get("label")

    def has_value(v):
        return v not in (None, "", 0, "0", "0.0", 0.0)

    product_val   = vals.get(LOE_PRODUCT_FIELD_ID)
    product_label = LOE_PRODUCT_LABELS.get(product_val)
    catalog = _loe_catalog()

    # Which section(s) to show: by product if set, else any section with a populated activity.
    if product_val in LOE_PRODUCT_SECTIONS:
        section_ids = LOE_PRODUCT_SECTIONS[product_val]
    else:
        section_ids = tuple(
            sid for sid in LOE_SECTIONS
            if any(has_value(vals.get(act_id))
                   for act_id, _, _ in catalog.get(sid, {}).get("activities", []))
        )

    sections = []
    for sid in section_ids:
        entry = catalog.get(sid, {"activities": [], "narratives": []})
        activities = [
            {"activity": label, "value": vals.get(act_id), "details": vals.get(det_id) or ""}
            for act_id, label, det_id in entry["activities"]
            if has_value(vals.get(act_id))
        ]
        overview = se_pov = ""
        for fid, label in entry["narratives"]:
            low = label.lower()
            if "project overview" in low:
                overview = vals.get(fid) or ""
            elif "se pov" in low:
                se_pov = vals.get(fid) or ""
        sections.append({
            "name": LOE_SECTIONS[sid],
            "overview": overview,
            "se_pov": se_pov,
            "activities": activities,
        })

    return jsonify({
        "name":    body.get("projectName") or "",
        "product": product_label,
        "status":  labels.get(LOE_STATUS_FIELD_ID) or vals.get(LOE_STATUS_FIELD_ID),
        "hours":   {k: vals.get(fid) for k, fid in LOE_HOURS_FIELD_IDS.items()},
        "sections": sections,
    })


@app.route("/assign-project/<int:project_id>", methods=["POST"])
@require_role("Manager", "Project Manager")
def assign_project(project_id):
    """Make the selected consultant the project owner and set status to Ready For Kick-Off,
    then remove the project's previous owner.

    The previous owner (read live before the ownership PUT) is removed unless it is the person
    performing the assignment (a regional PM keeps themselves on their own project) or the new
    consultant. Owner removal uses DELETE /v1/projects/{id}/remove-members — the only endpoint
    that can drop a member.
    """
    body = request.get_json(silent=True) or {}
    try:
        consultant_id = int(body.get("consultantId"))
    except (TypeError, ValueError):
        return jsonify({"success": False, "error": "consultantId is required (integer)"}), 400
    raw_pm_ids = body.get("pmIds") or ([body["pmId"]] if body.get("pmId") else [])
    try:
        pm_ids = [int(x) for x in raw_pm_ids]
    except (TypeError, ValueError):
        pm_ids = []
    if not pm_ids:
        return jsonify({"success": False, "error": "pmIds is required (non-empty list of integers)"}), 400

    warnings = []
    put_url = f"{ROCKETLANE_API_BASE_URL}/1.0/projects/{project_id}"

    # Step 0: read the current project state BEFORE the PUT — capture previous owner and
    # existing team members so we can merge PMs in and remove the right people after.
    # Also look up each PM's full profile so we can pass complete objects in the PUT.
    prev_owner = None
    existing_members = []
    resp0 = None
    try:
        resp0 = requests.get(put_url, headers=HEADERS, timeout=30)
        resp0.raise_for_status()
        data0 = resp0.json() or {}
        prev_owner = _person(data0.get("owner"))
        existing_members = data0.get("teamMembers", {}).get("members", [])
    except requests.exceptions.RequestException as exc:
        print(f"[assign] GET current owner ERROR: {exc}")
        warnings.append("Could not read the project's current owner; the prior owner was not removed.")

    pm_profiles = {}
    for pm_id in pm_ids:
        if pm_id == consultant_id:
            continue
        try:
            ur = requests.get(f"{ROCKETLANE_API_BASE_URL}/1.0/users/{pm_id}", headers=HEADERS, timeout=30)
            ur.raise_for_status()
            u = ur.json()
            pm_profiles[pm_id] = {
                "userId":    pm_id,
                "emailId":   u.get("email") or u.get("emailId") or "",
                "firstName": u.get("firstName") or "",
                "lastName":  u.get("lastName") or "",
            }
            print(f"[assign] PM profile {pm_id}: {pm_profiles[pm_id]}")
        except requests.exceptions.RequestException as exc:
            print(f"[assign] PM profile lookup {pm_id} failed (non-fatal): {exc}")
            pm_profiles[pm_id] = {"userId": pm_id}

    # Step 1: PUT /1.0/projects/{id} — set owner only (teamMembers is silently ignored by this endpoint).
    # Status is updated last so all membership changes are in place first.
    put_payload = {"owner": {"userId": consultant_id}}
    print(f"[assign] PUT /1.0 request payload: {json.dumps(put_payload)[:500]}")
    resp = None
    try:
        resp = requests.put(put_url, headers=HEADERS, json=put_payload, timeout=30)
        print(f"[assign] PUT /1.0 (owner+PMs) status={resp.status_code} body={resp.text[:300]}")
        resp.raise_for_status()
    except requests.exceptions.RequestException as exc:
        detail = resp.text if resp is not None else ""
        print(f"[assign] PUT ERROR: {exc} | {detail}")
        return jsonify({"success": False, "error": str(exc), "detail": detail}), 502

    # Step 2: remove the previous owner — keep the new consultant and any designated PM.
    prev_id = prev_owner.get("userId") if prev_owner else None
    if not prev_id:
        print("[assign] no prior owner to remove")
    elif prev_id == consultant_id:
        print("[assign] prior owner is the new consultant; keeping")
    elif prev_id in pm_ids:
        print(f"[assign] prior owner {prev_id} is a designated PM; keeping")
    else:
        remove_url = f"{ROCKETLANE_API_BASE_URL}/v1/projects/{project_id}/remove-members"
        resp3 = None
        try:
            resp3 = requests.delete(remove_url, headers=HEADERS, json=[{"userId": prev_id}], timeout=30)
            print(f"[assign] DELETE remove-members {prev_id} status={resp3.status_code} body={resp3.text[:300]}")
            resp3.raise_for_status()
        except requests.exceptions.RequestException as exc:
            detail = resp3.text if resp3 is not None else ""
            print(f"[assign] DELETE remove-members ERROR: {exc} | {detail}")
            warnings.append(f"{prev_owner.get('name') or 'The previous owner'} could not be removed automatically: {exc}.")

    # Step 3: add PM(s) as internal team members via POST /1.0/…/add-members (non-fatal).
    # The PUT /1.0 teamMembers field is silently ignored by the API, so we use a separate call.
    for pm_id in pm_ids:
        if pm_id == consultant_id:
            continue  # already set as owner
        pm_email = (pm_profiles.get(pm_id) or {}).get("emailId", "")
        if not pm_email:
            warnings.append(f"Project Manager (id {pm_id}) email not found; could not be added automatically.")
            continue
        pm_add_url  = f"{ROCKETLANE_API_BASE_URL}/1.0/projects/{project_id}/add-members"
        pm_add_body = {"members": [{"emailId": pm_email}]}
        pm_resp = None
        try:
            pm_resp = requests.post(pm_add_url, headers=HEADERS, json=pm_add_body, timeout=30)
            print(f"[assign] add PM {pm_id} ({pm_email}) status={pm_resp.status_code} body={pm_resp.text[:300]}")
            if not pm_resp.ok:
                try:
                    reason = pm_resp.json()["errors"][0]["reason"]
                except Exception:
                    reason = f"HTTP {pm_resp.status_code}"
                warnings.append(f"Project Manager ({pm_email}) could not be automatically added: {reason}. Please add them manually in RocketLane.")
        except requests.exceptions.RequestException as exc:
            print(f"[assign] add PM {pm_id} error (non-fatal): {exc}")
            warnings.append(f"Project Manager ({pm_email}) could not be automatically added. Please add them manually in RocketLane.")

    # Step 4: add technical contact as a customer team member by email.
    # customer_added tracks success — status is only set to Ready for Kick-Off when a
    # customer contact is on the project (an automation depends on this ordering).
    customer_name  = (body.get("customerName")  or "").strip()
    customer_email = (body.get("customerEmail") or "").strip()
    customer_added = False
    if not customer_email:
        warnings.append(
            "No customer contact was provided. The project status will not be set to"
            " Ready for Kick-Off until a customer contact is added in RocketLane."
        )
    else:
        name_parts = customer_name.split(None, 1)
        first_name = name_parts[0] if name_parts else ""
        last_name  = name_parts[1] if len(name_parts) > 1 else ""
        add_url  = f"{ROCKETLANE_API_BASE_URL}/1.0/projects/{project_id}/add-members"
        add_body = {"customers": [{"emailId": customer_email, "firstName": first_name, "lastName": last_name}]}
        try:
            ir = requests.post(add_url, headers=HEADERS, json=add_body, timeout=30)
            print(f"[assign] add-members (customer) {customer_email} status={ir.status_code} body={ir.text[:300]}")
            if not ir.ok:
                try:
                    reason = ir.json()["errors"][0]["reason"]
                except Exception:
                    reason = f"HTTP {ir.status_code}"
                warnings.append(
                    f"Technical contact ({customer_email}) could not be automatically added: {reason}"
                    " The project status has not been set to Ready for Kick-Off."
                    " Please add the contact manually in RocketLane, then update the status."
                )
            else:
                customer_added = True
                if first_name or last_name:
                    # add-members ignores name fields; update the user profile separately.
                    try:
                        ir_data = ir.json()
                        added_customers = ir_data.get("customers") or []
                        customer_user_id = next(
                            (c.get("userId") for c in added_customers if c.get("emailId") == customer_email),
                            None
                        )
                        if customer_user_id:
                            patch_url  = f"{ROCKETLANE_API_BASE_URL}/1.0/users/{customer_user_id}"
                            patch_body = {"firstName": first_name, "lastName": last_name}
                            pr = requests.put(patch_url, headers=HEADERS, json=patch_body, timeout=30)
                            print(f"[assign] update customer profile {customer_user_id} status={pr.status_code} body={pr.text[:200]}")
                    except Exception as exc:
                        print(f"[assign] update customer profile error (non-fatal): {exc}")
        except requests.exceptions.RequestException as exc:
            print(f"[assign] add-members (customer) error (non-fatal): {exc}")
            warnings.append(
                f"Technical contact ({customer_email}) could not be automatically added."
                " The project status has not been set to Ready for Kick-Off."
                " Please add the contact manually in RocketLane, then update the status."
            )

    # Step 5: set status to Ready for Kick-Off — only when a customer contact is on the project,
    # since an automation triggered by this status change depends on that contact being present.
    if customer_added:
        resp7 = None
        try:
            resp7 = requests.put(put_url, headers=HEADERS, json={"status": {"value": READY_FOR_KICKOFF_STATUS}}, timeout=30)
            print(f"[assign] PUT /1.0 (status) status={resp7.status_code} body={resp7.text[:300]}")
            resp7.raise_for_status()
        except requests.exceptions.RequestException as exc:
            detail = resp7.text if resp7 is not None else ""
            print(f"[assign] PUT status ERROR: {exc} | {detail}")
            warnings.append(f"Status could not be set to Ready for Kick-Off: {exc}. Please update it manually in RocketLane.")
    else:
        print(f"[assign] skipping status change — no customer contact added")

    # Step 6: set "PS Pre Kick-off" task in the Kickoff phase to In Progress (non-fatal)
    try:
        tasks_resp = requests.get(
            f"{ROCKETLANE_API_BASE_URL}/1.0/tasks?project.eq={project_id}&pageSize=200",
            headers=HEADERS, timeout=15
        )
        tasks_resp.raise_for_status()
        tasks = tasks_resp.json().get("data") or []
        ps_task = next((t for t in tasks if t.get("taskName") == "PS Pre Kick-off"), None)
        if ps_task:
            task_id = ps_task["taskId"]
            upd = requests.put(
                f"{ROCKETLANE_API_BASE_URL}/1.0/tasks/{task_id}",
                headers=HEADERS, json={"status": {"value": 2}}, timeout=15
            )
            print(f"[assign] PS Pre Kick-off task {task_id} -> In Progress status={upd.status_code} body={upd.text[:200]}")
            if not upd.ok:
                warnings.append("'PS Pre Kick-off' could not be set to In Progress. Please update it manually in RocketLane.")
        else:
            print(f"[assign] PS Pre Kick-off task not found for project {project_id}")
            warnings.append("'PS Pre Kick-off' task not found — please set it to In Progress manually in RocketLane.")
    except Exception as exc:
        print(f"[assign] PS Pre Kick-off update error (non-fatal): {exc}")
        warnings.append("Could not set 'PS Pre Kick-off' to In Progress. Please update it manually in RocketLane.")

    result = {"success": True}
    if warnings:
        result["warning"] = " ".join(warnings) + " Please complete the remaining steps manually in RocketLane."
    return jsonify(result)


BACKLOG_GROUP_TITLES = {
    "loe": "Missing LOE",
    "csm": "Missing CSM Handover",
    "sdp": "SDP Approval",
}


# ---------------------------------------------------------------------------
# Bulk Assign — assign role placeholders to unassigned template tasks
# ---------------------------------------------------------------------------
# Template tasks are assigned to role *placeholders* (teams), which real users fill in when a
# project is created from the template. Template reads/writes use the /v1/templates endpoints:
#   GET /v1/templates            -> list (includes non-project templates: docs, forms)
#   GET /v1/templates/{id}       -> templateContent.projectPhases[].tasks[] (+ nested subTasks)
#   PUT /v1/templates/{id}       -> full-object replace; templateMeta.assignee.teams must list every
#                                   placeholder used on a task (the server does not derive it).
BULK_ASSIGN_PLACEHOLDERS = {
    "PS Consultant": {"teamId": 351688, "teamName": "PS Consultant", "teamType": "NATIVE",
                      "roleId": 133617, "identityType": "TEAM", "resolved": False},
    "Customer":      {"teamId": 160708, "teamName": "Customer", "teamType": "EXTERNAL",
                      "roleId": -1, "identityType": "TEAM", "resolved": False},
    "PMO":           {"teamId": 238413, "teamName": "PMO", "teamType": "NATIVE",
                      "roleId": 128021, "identityType": "TEAM", "resolved": False},
}
BULK_ASSIGN_NAME_PREFIX = "PS - DC1"  # only templates whose name starts with this are listed
BULK_ASSIGN_TEAM_IDS ={p["teamId"] for p in BULK_ASSIGN_PLACEHOLDERS.values()}
# Fields the PUT accepts back (read-only keys like createdBy/account are dropped).
_TEMPLATE_PUT_KEYS = ("templateName", "templateContent", "templateMeta", "state", "objectType",
                      "sourceType", "effortPlanning", "autoBudgetedHours", "autoAllocation")


def _iter_template_tasks(template):
    """Yield (phase_name, task) for every task and subtask in a project template."""
    def walk(tasks, phase_name):
        for t in tasks or []:
            yield phase_name, t
            yield from walk(t.get("subTasks"), phase_name)
    for ph in (template.get("templateContent") or {}).get("projectPhases") or []:
        yield from walk(ph.get("tasks"), ph.get("projectPhaseName") or "")


def _task_has_role_assignee(task):
    teams = (task.get("assignee") or {}).get("teams") or []
    return any(tm.get("teamId") in BULK_ASSIGN_TEAM_IDS for tm in teams)


_TEMPLATE_CACHE = {}          # template_id -> (fetched_at, detail); read-side only
_TEMPLATE_CACHE_TTL = 300     # seconds; the assign route always reads fresh


def _fetch_template(template_id):
    """GET one template, retrying rate limits (429) and 5xx with backoff / Retry-After."""
    for attempt in range(6):
        resp = requests.get(f"{ROCKETLANE_API_BASE_URL}/v1/templates/{template_id}",
                            headers=HEADERS, timeout=60)
        if resp.status_code != 429 and resp.status_code < 500:
            break
        try:
            delay = float(resp.headers.get("Retry-After", ""))
        except ValueError:
            delay = 3 * 2 ** attempt
        time.sleep(min(delay, 30))
    resp.raise_for_status()
    return resp.json()


@app.route("/bulk-assign/templates", methods=["GET"])
@require_role("Manager", "Project Manager")
def bulk_assign_templates():
    """All project templates with the tasks not yet assigned to a Customer, PMO or PS Consultant."""
    try:
        resp = requests.get(f"{ROCKETLANE_API_BASE_URL}/v1/templates", headers=HEADERS, timeout=30)
        resp.raise_for_status()
        listing = [t for t in resp.json() if t.get("objectType") == "PROJECT"
                   and t.get("state") == "PUBLISHED"
                   and (t.get("templateName") or "").startswith(BULK_ASSIGN_NAME_PREFIX)]
    except requests.exceptions.RequestException as exc:
        return jsonify({"error": f"Could not fetch templates: {exc}"}), 502

    def load(t):
        hit = _TEMPLATE_CACHE.get(t["templateId"])
        if hit and time.time() - hit[0] < _TEMPLATE_CACHE_TTL:
            return t, hit[1], None
        try:
            detail = _fetch_template(t["templateId"])
            _TEMPLATE_CACHE[t["templateId"]] = (time.time(), detail)
            return t, detail, None
        except requests.exceptions.RequestException as exc:
            return t, None, str(exc)

    templates, failed = [], []
    with ThreadPoolExecutor(max_workers=6) as ex:
        for t, detail, err in ex.map(load, listing):
            if err:
                app.logger.warning("bulk-assign: template %s failed: %s", t.get("templateId"), err)
                failed.append(f"{t.get('templateName')} ({err[:120]})")
                continue
            tasks = []
            for phase, task in _iter_template_tasks(detail):
                if _task_has_role_assignee(task):
                    continue
                name = _html_to_text(task.get("taskName"), limit=200)
                tasks.append({"taskId": task["taskId"], "taskName": name, "phase": phase})
            if tasks:
                templates.append({"templateId": t["templateId"],
                                  "templateName": t.get("templateName") or "", "tasks": tasks})
    templates.sort(key=lambda t: t["templateName"].lower())
    return jsonify({"templates": templates, "failed": failed,
                    "roles": list(BULK_ASSIGN_PLACEHOLDERS)})


@app.route("/bulk-assign/assign", methods=["POST"])
@require_role("Manager", "Project Manager")
def bulk_assign_assign():
    """Body: {"assignments": [{"templateId": int, "taskId": int, "role": "PS Consultant"|...}]}.

    Adds the role placeholder to each task's existing assignees (nothing is removed). Each
    template is re-read fresh and written once; tasks that gained a Customer/PMO/PS Consultant
    assignee since the list was loaded are skipped.
    """
    items = (request.get_json(silent=True) or {}).get("assignments")
    if not isinstance(items, list) or not items:
        return jsonify({"error": "No assignments provided."}), 400
    by_template = {}
    for it in items:
        role = it.get("role")
        if role not in BULK_ASSIGN_PLACEHOLDERS or not isinstance(it.get("templateId"), int) \
                or not isinstance(it.get("taskId"), int):
            return jsonify({"error": f"Invalid assignment: {it}"}), 400
        by_template.setdefault(it["templateId"], {})[it["taskId"]] = role

    def apply(item):
        template_id, wanted = item
        try:
            detail = _fetch_template(template_id)
            teams_used, assigned, skipped = {}, 0, 0
            for _, task in _iter_template_tasks(detail):
                role = wanted.get(task.get("taskId"))
                if role is None:
                    continue
                if _task_has_role_assignee(task):
                    skipped += 1
                    continue
                ph = dict(BULK_ASSIGN_PLACEHOLDERS[role])
                assignee = task.setdefault("assignee", {"users": [], "teams": []})
                assignee.setdefault("users", [])
                assignee.setdefault("teams", []).append(ph)
                teams_used[ph["teamId"]] = ph
                assigned += 1
            if not assigned:
                return {"templateId": template_id, "name": detail.get("templateName"),
                        "assigned": 0, "skipped": skipped}
            meta = detail.get("templateMeta") or {}
            meta_teams = meta.setdefault("assignee", {"users": [], "teams": []}).setdefault("teams", [])
            known = {tm.get("teamId") for tm in meta_teams}
            meta_teams.extend(ph for tid, ph in teams_used.items() if tid not in known)
            body = {k: detail[k] for k in _TEMPLATE_PUT_KEYS if k in detail}
            body["templateMeta"] = meta
            put = requests.put(f"{ROCKETLANE_API_BASE_URL}/v1/templates/{template_id}",
                               headers=HEADERS, json=body, timeout=90)
            put.raise_for_status()
            _TEMPLATE_CACHE.pop(template_id, None)
            return {"templateId": template_id, "name": detail.get("templateName"),
                    "assigned": assigned, "skipped": skipped}
        except requests.exceptions.RequestException as exc:
            return {"templateId": template_id, "error": str(exc)}

    with ThreadPoolExecutor(max_workers=4) as ex:
        results = list(ex.map(apply, by_template.items()))
    app.logger.info("bulk-assign by %s: %s", session["user"].get("email"), results)
    return jsonify({"results": results,
                    "assigned": sum(r.get("assigned", 0) for r in results),
                    "errors": [r for r in results if r.get("error")]})


@app.route("/export-backlog/<group>")
@login_required
def export_backlog(group):
    """Export the Proposed projects belonging to a single backlog group (loe/csm/sdp)
    as an XLSX file. Classification matches the frontend's first-failing-check rule.
    """
    if group not in BACKLOG_GROUP_TITLES:
        return jsonify({"error": "invalid group"}), 400

    region = normalize_region(request.args.get("region"))
    try:
        projects = fetch_projects(region=region)
    except requests.exceptions.RequestException as exc:
        return jsonify({"error": f"Could not fetch projects: {exc}"}), 502

    # A project goes into the export for `group` whenever THAT specific prereq is missing —
    # independent of whether other prereqs are also missing. Mirrors the multi-card UI behaviour.
    target = []
    for p in projects:
        if p["status"] != "Proposed":
            continue
        loe_ok = p.get("loe_display") == "Complete"
        csm_ok = bool(p["csm_handover_date"]) and p["csm_handover_date"] != "-"
        sdp_ok = p["sdp_status"] in ("Approved", "Consent Required")
        if group == "loe" and not loe_ok:
            target.append(p)
        elif group == "csm" and not csm_ok:
            target.append(p)
        elif group == "sdp" and not sdp_ok:
            target.append(p)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = BACKLOG_GROUP_TITLES[group]
    ws.append(["Project Name", "Package Name", "ARR", "Segment", "Close Date",
               "CSM Handover Date", "SDP Status", "LOE Status", "Priority"])
    for p in target:
        ws.append([
            p["name"], p["package_name"], p["arr"], p["segment"],
            p["close_date"], p["csm_handover_date"],
            p["sdp_status"], p["loe_status"], p["priority"],
        ])

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    filename = f"backlog_{region.lower()}_{group}_{datetime.now().strftime('%Y-%m-%d')}.xlsx"
    return send_file(output,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                     as_attachment=True, download_name=filename)


@app.route("/export-archived")
@login_required
def export_archived():
    # Fetch all archived projects with pagination.
    # The API returns archived=False in the response body even for archived projects,
    # so we trust the filter rather than checking the field.
    region = normalize_region(request.args.get("region"))
    sub_regions = REGION_SUB_REGIONS[region]
    endpoint = f"{ROCKETLANE_API_BASE_URL}/1.0/projects"
    params = {"archived.eq": "true", "limit.eq": 100}

    raw_items = []
    url = endpoint
    while url:
        resp = requests.get(url, headers=HEADERS, params=params, timeout=30)
        resp.raise_for_status()
        data = resp.json()
        raw_items.extend(data.get("data", []))
        next_url = data.get("pagination", {}).get("nextPage")
        url = next_url
        params = {}  # pagination token is embedded in nextPage URL

    # Deduplicate, then filter to the active region by sub-region membership.
    seen = set()
    archived = []
    for item in raw_items:
        pid = item.get("projectId")
        if pid in seen:
            continue
        sub = get_custom_field_label(item, 1611023)
        if sub not in sub_regions:
            continue
        seen.add(pid)
        archived.append(item)

    print(f"[export-archived] {region}: {len(archived)} archived projects found")

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = f"Archived ({region})"

    headers = ["Project Name", "Package Name", "Opportunity Owner", "Close Date",
               "CSM Handover Date", "SDP Status", "LOE Status", "Status"]
    ws.append(headers)

    for item in archived:
        ws.append([
            item.get("projectName") or item.get("name") or "",
            get_custom_field(item, 1621438) or "",
            get_custom_field(item, 2280752) or "",
            format_date(item.get("dueDate") or item.get("dueDateActual")),
            format_date(get_custom_field(item, 1925432)),
            get_custom_field_label(item, 2098584) or "",
            get_custom_field_label(item, 2098585) or "",
            item.get("status", {}).get("label", ""),
        ])

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    filename = f"archived_projects_{region.lower()}_{datetime.now().strftime('%Y-%m-%d')}.xlsx"
    return send_file(output, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                     as_attachment=True, download_name=filename)


# ---------------------------------------------------------------------------
# Settings page (Manager-only) — SSO + Users + API Keys
# ---------------------------------------------------------------------------

def _sso_context():
    """Build the read-only SSO config block shown on the Settings page.

    Source of truth is .env; the UI only displays. Client secret is masked to
    the last 4 characters at most (revealed via a UI toggle, never the full
    plaintext) so a Manager can sanity-check what is configured without
    leaking the secret.
    """
    issuer        = os.getenv("OKTA_DOMAIN", "")
    client_id     = os.getenv("OKTA_CLIENT_ID", "")
    client_secret = os.getenv("OKTA_CLIENT_SECRET", "")
    last4 = client_secret[-4:] if len(client_secret) >= 4 else ""
    return {
        "issuer":               issuer,
        "discovery_url":        f"{issuer}/.well-known/openid-configuration" if issuer else "",
        "client_id":            client_id,
        "client_secret_masked": "•" * 8,
        "client_secret_last4":  last4,
        "scopes":               "openid email profile groups",
        "redirect_uri":         url_for("auth_callback", _external=True),
        "scim_base_url":        url_for("scim.users_list", _external=True).rsplit("/Users", 1)[0],
        "configured":           bool(issuer and client_id and client_secret),
        "source":               ".env (environment variables)",
    }


def _render_settings(banner=None, error=None, new_token=None):
    return render_template(
        "settings.html",
        user=session["user"], current_user=session["user"],
        users=db.list_users(),
        api_keys=db.apikey_list(),
        roles=db.ROLES,
        sso=_sso_context(),
        banner=banner, error=error, new_token=new_token,
    )


@app.route("/settings")
@require_role("Manager")
def settings_page():
    return _render_settings()


@app.route("/users", methods=["POST"])
@require_role("Manager")
def users_create():
    email = (request.form.get("email") or "").strip()
    first = (request.form.get("first_name") or "").strip()
    last  = (request.form.get("last_name") or "").strip()
    role  = (request.form.get("role") or "").strip()
    if not (email and first and last and role):
        return _render_settings(error="Email, first name, last name, and role are all required.")
    if role not in db.ROLES:
        return _render_settings(error=f"Invalid role: {role}")
    if db.user_by_email(email) is not None:
        return _render_settings(error=f"A user with email {email} already exists.")
    try:
        db.create_user(email=email, first_name=first, last_name=last, role=role)
    except (sqlite3.IntegrityError, ValueError) as exc:
        return _render_settings(error=str(exc))
    return _render_settings(banner=f"User {email} created.")


@app.route("/users/<user_id>/update", methods=["POST"])
@require_role("Manager")
def users_update(user_id):
    if db.user_by_id(user_id) is None:
        return _render_settings(error="User not found.")
    fields = {}
    for key in ("first_name", "last_name", "role"):
        val = (request.form.get(key) or "").strip()
        if val:
            fields[key] = val
    try:
        db.update_user(user_id, **fields)
    except ValueError as exc:
        return _render_settings(error=str(exc))
    return _render_settings(banner="User updated.")


@app.route("/users/<user_id>/deactivate", methods=["POST"])
@require_role("Manager")
def users_deactivate(user_id):
    if user_id == session["user"]["id"]:
        return _render_settings(error="You cannot deactivate yourself.")
    db.set_active(user_id, False)
    return _render_settings(banner="User deactivated.")


@app.route("/users/<user_id>/reactivate", methods=["POST"])
@require_role("Manager")
def users_reactivate(user_id):
    db.set_active(user_id, True)
    return _render_settings(banner="User reactivated.")


@app.route("/api-keys", methods=["POST"])
@require_role("Manager")
def api_keys_create():
    label = (request.form.get("label") or "").strip()
    if not label:
        return _render_settings(error="API key label is required.")
    _, raw = db.apikey_create(label, created_by=session["user"]["id"])
    return _render_settings(banner=f"API key '{label}' created.", new_token=raw)


@app.route("/api-keys/<key_id>/deactivate", methods=["POST"])
@require_role("Manager")
def api_keys_deactivate(key_id):
    db.apikey_deactivate(key_id)
    return _render_settings(banner="API key deactivated.")


# ---------------------------------------------------------------------------
# Debug helpers (remove before prod)
# ---------------------------------------------------------------------------


@app.route("/debug/project-owner/<int:project_id>")
@require_role("Manager")
def debug_project_owner(project_id):
    import html as _h
    r = requests.get(f"{ROCKETLANE_API_BASE_URL}/1.0/projects/{project_id}",
                     headers=HEADERS, timeout=30)
    data = r.json()
    owner = data.get("owner") or data.get("data", {}).get("owner")
    members = (data.get("teamMembers") or {}).get("members", [])
    return (f"<pre>status={r.status_code}\n"
            f"owner={_h.escape(str(owner))}\n"
            f"members={_h.escape(str(members))}\n"
            f"name={_h.escape(str(data.get('projectName') or data.get('data',{}).get('projectName','')))}</pre>")


@app.route("/debug/find-project")
@require_role("Manager")
def debug_find_project():
    import html as _h
    name_q = request.args.get("q", "").lower()
    region = normalize_region(request.args.get("region", "AMS"))
    try:
        projects = fetch_projects(region=region)
    except Exception as exc:
        return f"<pre>ERROR: {exc}</pre>"
    matches = [p for p in projects if name_q in p["name"].lower()]
    rows = "".join(
        f"<p><a href='/debug/project-owner/{p['id']}'>{_h.escape(p['name'])}</a> — id={p['id']} status={p['status']}</p>"
        for p in matches
    )
    return f"<html><body>{rows or '<p>No matches</p>'}</body></html>"


# ---------------------------------------------------------------------------
# SCIM 2.0 blueprint
# ---------------------------------------------------------------------------

app.register_blueprint(scim_bp)


if __name__ == "__main__":
    os.environ.setdefault("AUTHLIB_INSECURE_TRANSPORT", "1")
    app.run(debug=True, port=5000)