"""SCIM 2.0 endpoints for User provisioning.

Implements the subset of SCIM 2.0 needed for IdP-driven provisioning (Okta,
Azure AD, etc.):
  GET  /scim/v2/ServiceProviderConfig    (anonymous, discovery)
  GET  /scim/v2/ResourceTypes            (anonymous, discovery)
  GET  /scim/v2/Schemas                  (anonymous, discovery)
  GET  /scim/v2/Users                    (list, with optional filter)
  GET  /scim/v2/Users/<id>               (read)
  POST /scim/v2/Users                    (create)
  PUT  /scim/v2/Users/<id>               (replace)
  PATCH /scim/v2/Users/<id>              (partial update — Okta uses this for active toggles)
  DELETE /scim/v2/Users/<id>             (soft delete — sets active=false)

All /scim/v2/Users* endpoints require Authorization: Bearer <token> against an
api_keys row issued from the Manager UI.

Schema mapping (SCIM → DB):
  userName            → email
  name.givenName      → first_name
  name.familyName     → last_name
  roles[0].value      → role (default "Read Only" if absent/unknown)
  active              → active
  externalId          → scim_external_id
"""

import re
from urllib.parse import urlparse

from flask import Blueprint, jsonify, make_response, request, url_for

import db

scim_bp = Blueprint("scim", __name__, url_prefix="/scim/v2")

SCIM_CONTENT_TYPE   = "application/scim+json"
USER_SCHEMA         = "urn:ietf:params:scim:schemas:core:2.0:User"
LIST_RESP_SCHEMA    = "urn:ietf:params:scim:api:messages:2.0:ListResponse"
ERROR_SCHEMA        = "urn:ietf:params:scim:api:messages:2.0:Error"
PATCHOP_SCHEMA      = "urn:ietf:params:scim:api:messages:2.0:PatchOp"

VALID_ROLES = set(db.ROLES)


# ---------------------------------------------------------------------------
# Auth
# ---------------------------------------------------------------------------

def _bearer_token():
    auth = request.headers.get("Authorization", "")
    if not auth.lower().startswith("bearer "):
        return None
    return auth[7:].strip()


def _require_auth():
    tok = _bearer_token()
    if not tok or db.apikey_lookup_by_token(tok) is None:
        return _error(401, "Unauthorized")
    return None


# ---------------------------------------------------------------------------
# Response helpers
# ---------------------------------------------------------------------------

def _scim_response(payload, status=200):
    resp = make_response(jsonify(payload), status)
    resp.headers["Content-Type"] = SCIM_CONTENT_TYPE
    return resp


def _error(status, detail, scim_type=None):
    body = {
        "schemas": [ERROR_SCHEMA],
        "status":  str(status),
        "detail":  detail,
    }
    if scim_type:
        body["scimType"] = scim_type
    return _scim_response(body, status)


def _user_to_scim(u):
    """Convert a DB user dict into a SCIM User resource."""
    return {
        "schemas": [USER_SCHEMA],
        "id":      u["id"],
        "externalId": u.get("scim_external_id"),
        "userName": u["email"],
        "name": {
            "givenName":  u["first_name"],
            "familyName": u["last_name"],
            "formatted":  u["name"],
        },
        "displayName": u["name"],
        "emails": [{"value": u["email"], "primary": True}],
        "active":  u["active"],
        "roles":   [{"value": u["role"]}],
        "meta": {
            "resourceType": "User",
            "created":      u["created_at"],
            "lastModified": u["updated_at"],
            "location":     url_for("scim.user_detail", user_id=u["id"], _external=True),
        },
    }


def _role_from_scim(payload):
    """Pull a valid role out of either roles[].value or roles[].display. Defaults to Read Only."""
    roles = payload.get("roles") or []
    for r in roles:
        for key in ("value", "display"):
            v = r.get(key) if isinstance(r, dict) else None
            if v in VALID_ROLES:
                return v
    return "Read Only"


def _fields_from_scim(payload):
    """Map a SCIM User body to DB-side update fields. Returns dict with only present keys."""
    out = {}
    if "userName" in payload:
        out["email"] = payload["userName"]
    name = payload.get("name") or {}
    if "givenName" in name:
        out["first_name"] = name["givenName"]
    if "familyName" in name:
        out["last_name"] = name["familyName"]
    if "roles" in payload:
        out["role"] = _role_from_scim(payload)
    if "active" in payload:
        out["active"] = bool(payload["active"])
    if "externalId" in payload:
        out["scim_external_id"] = payload["externalId"]
    return out


# ---------------------------------------------------------------------------
# Discovery (anonymous)
# ---------------------------------------------------------------------------

@scim_bp.route("/ServiceProviderConfig", methods=["GET"])
def service_provider_config():
    return _scim_response({
        "schemas": ["urn:ietf:params:scim:schemas:core:2.0:ServiceProviderConfig"],
        "documentationUri": "https://datatracker.ietf.org/doc/html/rfc7644",
        "patch":         {"supported": True},
        "bulk":          {"supported": False, "maxOperations": 0, "maxPayloadSize": 0},
        "filter":        {"supported": True, "maxResults": 200},
        "changePassword":{"supported": False},
        "sort":          {"supported": False},
        "etag":          {"supported": False},
        "authenticationSchemes": [{
            "type":        "oauthbearertoken",
            "name":        "OAuth Bearer Token",
            "description": "Authenticate with Authorization: Bearer <token> against an API key issued from the Manager UI.",
            "primary":     True,
        }],
    })


@scim_bp.route("/ResourceTypes", methods=["GET"])
def resource_types():
    return _scim_response({
        "schemas":     [LIST_RESP_SCHEMA],
        "totalResults": 1,
        "Resources": [{
            "schemas":  ["urn:ietf:params:scim:schemas:core:2.0:ResourceType"],
            "id":       "User",
            "name":     "User",
            "endpoint": "/Users",
            "schema":   USER_SCHEMA,
        }],
    })


@scim_bp.route("/Schemas", methods=["GET"])
def schemas():
    return _scim_response({
        "schemas":      [LIST_RESP_SCHEMA],
        "totalResults": 1,
        "Resources": [{
            "id":          USER_SCHEMA,
            "name":        "User",
            "description": "User Account",
            "attributes":  [],   # full attribute discovery not required for Okta provisioning
        }],
    })


# ---------------------------------------------------------------------------
# Users CRUD
# ---------------------------------------------------------------------------

_FILTER_EQ_RE = re.compile(r'^\s*(\w+)\s+eq\s+"([^"]*)"\s*$', re.IGNORECASE)


def _filter_users(users, filter_str):
    """Parse only the minimal `<attr> eq "<val>"` form Okta uses for userName lookups."""
    if not filter_str:
        return users
    m = _FILTER_EQ_RE.match(filter_str)
    if not m:
        return users
    attr, val = m.group(1), m.group(2)
    if attr.lower() == "username":
        return [u for u in users if u["email"].lower() == val.lower()]
    if attr.lower() == "id":
        return [u for u in users if u["id"] == val]
    if attr.lower() == "externalid":
        return [u for u in users if (u.get("scim_external_id") or "") == val]
    return users


@scim_bp.route("/Users", methods=["GET"])
def users_list():
    err = _require_auth()
    if err:
        return err
    users = db.list_users()
    users = _filter_users(users, request.args.get("filter"))
    # Pagination per SCIM. Defaults: startIndex=1, count=200.
    try:
        start = max(1, int(request.args.get("startIndex", 1)))
    except ValueError:
        start = 1
    try:
        count = min(200, max(0, int(request.args.get("count", 200))))
    except ValueError:
        count = 200
    sliced = users[start - 1 : start - 1 + count]
    return _scim_response({
        "schemas":      [LIST_RESP_SCHEMA],
        "totalResults": len(users),
        "startIndex":   start,
        "itemsPerPage": len(sliced),
        "Resources":    [_user_to_scim(u) for u in sliced],
    })


@scim_bp.route("/Users/<user_id>", methods=["GET"])
def user_detail(user_id):
    err = _require_auth()
    if err:
        return err
    u = db.user_by_id(user_id)
    if u is None:
        return _error(404, "User not found")
    return _scim_response(_user_to_scim(u))


@scim_bp.route("/Users", methods=["POST"])
def user_create():
    err = _require_auth()
    if err:
        return err
    payload = request.get_json(silent=True) or {}
    email = (payload.get("userName") or "").strip()
    name  = payload.get("name") or {}
    first = (name.get("givenName")  or "").strip()
    last  = (name.get("familyName") or "").strip()
    if not (email and first and last):
        return _error(400, "userName, name.givenName, and name.familyName are required",
                      scim_type="invalidValue")
    if db.user_by_email(email) is not None:
        return _error(409, "User with that userName already exists", scim_type="uniqueness")
    role = _role_from_scim(payload)
    active = payload.get("active", True)
    try:
        u = db.create_user(
            email=email, first_name=first, last_name=last,
            role=role, active=bool(active),
            scim_external_id=payload.get("externalId"),
        )
    except ValueError as exc:
        return _error(400, str(exc), scim_type="invalidValue")
    return _scim_response(_user_to_scim(u), status=201)


@scim_bp.route("/Users/<user_id>", methods=["PUT"])
def user_replace(user_id):
    err = _require_auth()
    if err:
        return err
    if db.user_by_id(user_id) is None:
        return _error(404, "User not found")
    payload = request.get_json(silent=True) or {}
    fields = _fields_from_scim(payload)
    if not fields.get("email") or not fields.get("first_name") or not fields.get("last_name"):
        return _error(400, "userName, name.givenName, and name.familyName are required for PUT",
                      scim_type="invalidValue")
    # Default role to Read Only if omitted on PUT.
    fields.setdefault("role", "Read Only")
    try:
        u = db.update_user(user_id, **fields)
    except ValueError as exc:
        return _error(400, str(exc), scim_type="invalidValue")
    return _scim_response(_user_to_scim(u))


# Map SCIM path tokens to DB columns for PATCH ops.
_PATCH_PATH_MAP = {
    "username":           "email",
    "name.givenname":     "first_name",
    "name.familyname":    "last_name",
    "active":             "active",
    "externalid":         "scim_external_id",
}


def _apply_patch_op(updates, op):
    op_type = (op.get("op") or "").lower()
    path    = (op.get("path") or "").strip()
    value   = op.get("value")
    if op_type == "remove":
        return  # ignore — we only soft-delete via active=false, no field nulling here
    if op_type not in ("replace", "add"):
        return
    # No path → SCIM allows value to be an object that mirrors a User resource.
    if not path and isinstance(value, dict):
        updates.update(_fields_from_scim(value))
        return
    plower = path.lower()
    if plower == "roles" or plower.startswith("roles"):
        if isinstance(value, list):
            updates["role"] = _role_from_scim({"roles": value})
        elif isinstance(value, dict):
            updates["role"] = _role_from_scim({"roles": [value]})
        elif isinstance(value, str) and value in VALID_ROLES:
            updates["role"] = value
        return
    col = _PATCH_PATH_MAP.get(plower)
    if col:
        if col == "active":
            updates[col] = bool(value)
        else:
            updates[col] = value


@scim_bp.route("/Users/<user_id>", methods=["PATCH"])
def user_patch(user_id):
    err = _require_auth()
    if err:
        return err
    if db.user_by_id(user_id) is None:
        return _error(404, "User not found")
    payload = request.get_json(silent=True) or {}
    ops = payload.get("Operations") or []
    if not isinstance(ops, list):
        return _error(400, "Operations must be a list", scim_type="invalidSyntax")
    updates = {}
    for op in ops:
        if isinstance(op, dict):
            _apply_patch_op(updates, op)
    try:
        u = db.update_user(user_id, **updates)
    except ValueError as exc:
        return _error(400, str(exc), scim_type="invalidValue")
    return _scim_response(_user_to_scim(u))


@scim_bp.route("/Users/<user_id>", methods=["DELETE"])
def user_delete(user_id):
    err = _require_auth()
    if err:
        return err
    if db.user_by_id(user_id) is None:
        return _error(404, "User not found")
    db.set_active(user_id, False)
    return ("", 204)
