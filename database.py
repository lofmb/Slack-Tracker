"""
Persistence for the tracker, backed by LMSA instead of a local SQLite file.

The public function names and signatures below are the ones app.py already
calls, so the Slack side is unchanged: same cards, same modals, same buttons,
same copy. What changed is where the data lives. Every call is now a small HTTP
request to LMSA's loopback-only internal API, which owns the Postgres schema,
the transactions, the row locking, the audit trail and the idempotency.

Why not talk to Postgres from here: LMSA keeps a single writer so that one
transaction discipline applies to every mutation, whoever asks. There is also
no database driver available to this process — the API needs only urllib from
the standard library, which is the whole dependency story for this file.

Rows come back shaped like the sqlite3.Row dictionaries app.py has always
indexed into (task["customer_name"], task["field_elapsed"] and so on), so
callers did not have to change. The per-phase elapsed values are summed from
LMSA's timing ledger on read rather than stored, because a stored total drifts
from the segments it claims to summarise.
"""

import contextvars
import hashlib
import json
import os
import urllib.error
import urllib.request

# --- configuration ---------------------------------------------------------
# Loopback only. The tracker runs beside LMSA on the same host and the API is
# not published anywhere else; the token is its own secret, unrelated to any
# Slack credential, and is never logged or included in an error message.

API_BASE_URL = os.environ.get("TRACKER_API_BASE_URL", "http://127.0.0.1:5000/internal/tracker")
API_TOKEN = os.environ.get("TRACKER_INTERNAL_TOKEN", "")
API_TIMEOUT_SECONDS = float(os.environ.get("TRACKER_API_TIMEOUT_SECONDS", "10"))

PHASES = ("field_sheeting", "border_sheeting", "packing")


# --- errors ----------------------------------------------------------------

class TrackerApiError(RuntimeError):
    """The API could not be reached, or answered with something unusable."""


class TrackerRefused(RuntimeError):
    """
    LMSA declined the request for a business reason it named, such as
    already_running or job_not_open. Not a fault: these are the outcomes a real
    person produces by clicking, and callers that care read .reason.
    """

    def __init__(self, reason, detail=None):
        super().__init__(reason)
        self.reason = reason
        self.detail = detail


# --- request identity ------------------------------------------------------
# One Slack delivery can drive several writes (creating a job also stores its
# card), so the key is the delivery plus the operation. A redelivery of the
# same click repeats action_ts and therefore repeats the key, which LMSA
# absorbs; a genuine later click carries a new action_ts and is treated as the
# new action it is. A context variable rather than a global, so two overlapping
# requests cannot read each other's identity.

_KEY_VERSION = "trk1"
_current_identity = contextvars.ContextVar("tracker_request_identity", default=None)


def delivery_identity(body):
    """Identify one Slack delivery, or None when the payload cannot be identified."""
    if not isinstance(body, dict):
        return None
    kind = body.get("type")

    if kind == "block_actions":
        actions = body.get("actions") or []
        if not actions:
            return None
        action = actions[0] or {}
        action_ts = action.get("action_ts")
        if not action_ts:
            return None
        container = body.get("container") or {}
        return ":".join([
            "block_actions",
            str((body.get("team") or {}).get("id") or ""),
            str((body.get("user") or {}).get("id") or ""),
            str(container.get("channel_id") or ""),
            str(container.get("message_ts") or ""),
            str(action.get("action_id") or ""),
            str(action_ts),
        ])

    if kind == "view_submission":
        view = body.get("view") or {}
        if not view.get("id"):
            return None
        return ":".join([
            "view_submission",
            str((body.get("team") or {}).get("id") or ""),
            str((body.get("user") or {}).get("id") or ""),
            str(view.get("callback_id") or ""),
            str(view.get("id")),
            str(view.get("hash") or ""),
        ])

    return None


class slack_request:
    """
    Scope one delivery's identity to the handler that is running.

    Set on entry and reset on exit, including when the handler raises, so a
    later request can never inherit a key that was minted for an earlier one.
    """

    def __init__(self, body):
        self._identity = delivery_identity(body)
        self._token = None

    def __enter__(self):
        self._token = _current_identity.set(self._identity)
        return self

    def __exit__(self, *exc):
        _current_identity.reset(self._token)
        return False


def _idempotency_key(operation):
    """Hashed so the key stays compact and carries no payload verbatim."""
    identity = _current_identity.get()
    if not identity:
        return None
    return hashlib.sha256(
        f"{_KEY_VERSION}:{identity}:{operation}".encode("utf-8")
    ).hexdigest()


# --- HTTP ------------------------------------------------------------------

def _redact(text):
    return text.replace(API_TOKEN, "<redacted>") if API_TOKEN else text


def _call(method, path, payload=None, operation=None):
    """
    One request to the internal API.

    A 2xx is not taken as success on its own. LMSA answers every request with
    JSON carrying an `ok` field precisely so this side can tell a real reply
    from an HTML error page, a proxy response or an app that is not the one we
    meant to reach — a mistake that would otherwise look like a silent success.
    """
    body = dict(payload or {})
    if operation is not None:
        key = _idempotency_key(operation)
        if key:
            body["idempotencyKey"] = key

    data = json.dumps(body).encode("utf-8") if method != "GET" else None
    request = urllib.request.Request(
        f"{API_BASE_URL}{path}",
        data=data,
        method=method,
        headers={
            "Content-Type": "application/json",
            "X-Tracker-Token": API_TOKEN,
        },
    )

    try:
        with urllib.request.urlopen(request, timeout=API_TIMEOUT_SECONDS) as response:
            status = response.status
            content_type = response.headers.get("Content-Type", "")
            raw = response.read().decode("utf-8", "replace")
    except urllib.error.HTTPError as err:  # 4xx/5xx still carry our envelope
        status = err.code
        content_type = err.headers.get("Content-Type", "") if err.headers else ""
        raw = err.read().decode("utf-8", "replace")
    except urllib.error.URLError as err:
        raise TrackerApiError(f"tracker API unreachable at {API_BASE_URL}: {_redact(str(err.reason))}")
    except TimeoutError:
        raise TrackerApiError(f"tracker API timed out after {API_TIMEOUT_SECONDS}s")

    if "application/json" not in content_type.lower():
        raise TrackerApiError(
            f"tracker API returned {content_type or 'no content-type'} instead of JSON "
            f"(HTTP {status}) — check the API base URL: {_redact(raw[:120])}"
        )
    try:
        envelope = json.loads(raw)
    except ValueError:
        raise TrackerApiError(f"tracker API returned malformed JSON (HTTP {status})")
    if not isinstance(envelope, dict) or "ok" not in envelope:
        raise TrackerApiError(f"tracker API response has no ok field (HTTP {status})")

    if envelope["ok"]:
        return envelope.get("data")

    error = envelope.get("error")
    if error == "refused" or envelope.get("reason"):
        raise TrackerRefused(envelope.get("reason") or error, envelope.get("detail"))
    raise TrackerApiError(f"tracker API error: {error} (HTTP {status})")


# --- shape translation -----------------------------------------------------

def _iso_to_sqlite_datetime(value):
    """LMSA sends ISO instants; app.py has always rendered "YYYY-MM-DD HH:MM:SS"."""
    if not value:
        return None
    return str(value).replace("T", " ")[:19]


def _due_date_to_text(job):
    """
    app.py displays whatever due date it was given, or "N/A".

    LMSA keeps a real date plus the explicit "No Set Date?" choice, so the
    stored date is rendered back in the DD/MM/YY the modal asks for.
    """
    if job.get("dueDateNotApplicable"):
        return "N/A"
    iso = job.get("dueDate")
    if not iso:
        return None
    parts = str(iso)[:10].split("-")
    if len(parts) != 3:
        return str(iso)
    return f"{parts[2]}/{parts[1]}/{parts[0][2:]}"


def _text_to_due_date(text):
    """
    The reverse: the modal collects DD/MM/YY (or DD/MM/YYYY) free text.

    Returns (iso_date_or_None, not_applicable). Anything that is not a date —
    "N/A", a blank, or free text the modal did not ask for — is recorded as
    "no set date", which is the same sentinel app.py already substitutes.
    """
    raw = (text or "").strip()
    if not raw or raw.upper() == "N/A":
        return None, True
    for separator in ("/", "-", "."):
        parts = raw.split(separator)
        if len(parts) == 3 and all(p.isdigit() for p in parts):
            day, month, year = parts
            if len(year) == 2:
                year = f"20{year}"
            try:
                if 1 <= int(day) <= 31 and 1 <= int(month) <= 12 and len(year) == 4:
                    return f"{year}-{int(month):02d}-{int(day):02d}", False
            except ValueError:
                break
    return None, True


def _current_phase(phases):
    """
    The phase the card is showing: the first that has not been completed.

    LMSA models running/paused per phase rather than per job, so the single
    "current phase" app.py expects is derived rather than stored — which is
    what lets a completed phase and a live one coexist in the same job.
    """
    for name in PHASES:
        row = next((p for p in phases if p.get("phase") == name), None)
        if row is None or row.get("state") != "complete":
            return name
    return "completed"


def _legacy_status(job, phases):
    """Translate LMSA's job/phase state into the status vocabulary app.py reads."""
    if job.get("status") == "completed":
        return "completed"
    phase_name = _current_phase(phases)
    if phase_name == "completed":
        return "completed"
    row = next((p for p in phases if p.get("phase") == phase_name), None)
    state = (row or {}).get("state")
    if state == "running":
        return "in_progress"
    if state == "paused":
        return "paused"
    return "created"


def _phase_of(phases, name, field):
    row = next((p for p in phases if p.get("phase") == name), None)
    return (row or {}).get(field)


def _row(view, timing):
    """Build the dictionary app.py indexes into, from the API's job view."""
    job = view["job"]
    phases = view.get("phases") or []
    seconds = (timing or {}).get("perPhaseSeconds") or {}
    field = int(seconds.get("field_sheeting", 0) or 0)
    border = int(seconds.get("border_sheeting", 0) or 0)
    packing = int(seconds.get("packing", 0) or 0)

    return {
        "task_id": job["jobNumber"],
        "user_id": job["ownerSlackUserId"],
        "channel_id": job.get("announceChannelId"),
        "customer_name": job["customerName"],
        "invoice_number": job["invoiceNumber"],
        "task_description": job["taskDescription"],
        "due_date": _due_date_to_text(job),
        "is_na_due_date": 1 if job.get("dueDateNotApplicable") else 0,
        "field_design": _phase_of(phases, "field_sheeting", "designName"),
        "difficulty": _phase_of(phases, "field_sheeting", "difficulty"),
        "field_elapsed": field,
        "border_design": _phase_of(phases, "border_sheeting", "designName"),
        "border_difficulty": _phase_of(phases, "border_sheeting", "difficulty"),
        "border_elapsed": border,
        "packing_elapsed": packing,
        "general_notes": job.get("generalNotes"),
        "issues_encountered": job.get("issuesEncountered"),
        "status": _legacy_status(job, phases),
        "current_phase": _current_phase(phases),
        "created_at": _iso_to_sqlite_datetime(job.get("createdAt")),
        "message_ts": job.get("cardMessageTs"),
        "dm_channel_id": job.get("dmChannelId"),
        "total_elapsed": field + border + packing,
    }


def _view_by_number(task_id):
    """
    Resolve the integer on the card's buttons to the job LMSA knows.

    app.py puts this number in every button value and reads it back with
    int(...), so it stays the tracker's identity; LMSA's own uuid never appears
    in anything a maker sees. Resolved through the API on every call rather
    than remembered, so a restart changes nothing.
    """
    try:
        return _call("GET", f"/jobs/by-number/{int(task_id)}")
    except TrackerRefused as refusal:
        if refusal.reason in ("job_not_found", "no_open_job"):
            return None
        raise


def _timing(job_id):
    return _call("GET", f"/jobs/{job_id}/timing")


def _row_for(task_id):
    view = _view_by_number(task_id)
    if view is None:
        return None
    return view, _row(view, _timing(view["job"]["id"]))


# --- the interface app.py uses --------------------------------------------

def setup_database():
    """
    Confirm the tracker can reach LMSA. It no longer creates anything.

    Schema is applied by hand on the LMSA side and this process has no way to
    change it — no driver, no credentials, no privileges. Removing the ability
    is stronger than a rule saying not to, and it means a tracker restart can
    never migrate a production database as a side effect of booting.
    """
    health = _call("GET", "/health")
    if not isinstance(health, dict) or not health.get("ready"):
        raise TrackerApiError("tracker API did not report ready")
    print("Tracker API ready.")


def create_task(user_id, channel_id, customer_name, invoice_number, task_description, due_date, is_na, design, difficulty):
    """Create a job and return its number — the T-number shown on the card."""
    iso_due, not_applicable = _text_to_due_date(due_date)
    if is_na:
        iso_due, not_applicable = None, True
    data = _call("POST", "/jobs", {
        "ownerSlackUserId": user_id,
        "customerName": customer_name,
        "invoiceNumber": invoice_number,
        "taskDescription": task_description,
        "dueDate": iso_due,
        "dueDateNotApplicable": not_applicable,
        "fieldDesignName": design,
        "fieldDifficulty": difficulty,
        "announceChannelId": channel_id,
        "actor": f"slack:{user_id}",
    }, operation="create_task")
    return data["job"]["jobNumber"]


def get_active_task(user_id):
    """The job a maker currently has open, if any."""
    try:
        view = _call("GET", f"/jobs/open/{user_id}")
    except TrackerRefused as refusal:
        if refusal.reason in ("no_open_job", "job_not_found"):
            return None
        raise
    return _row(view, _timing(view["job"]["id"]))


def get_task(task_id):
    """
    One job by its number, or None.

    A cancelled job reads as None: the tracker's Delete removed the row, and
    app.py already tells the maker the task "may have been deleted". LMSA
    retains it instead of destroying the history, but it is gone from the
    workflow either way.
    """
    resolved = _row_for(task_id)
    if resolved is None:
        return None
    view, row = resolved
    if view["job"].get("status") == "cancelled":
        return None
    return row


def start_task(task_id):
    """Start — or resume — timing the phase the card is showing."""
    resolved = _row_for(task_id)
    if resolved is None:
        return
    view, row = resolved
    phase = row["current_phase"]
    if phase == "completed":
        return
    try:
        _call("POST", f"/jobs/{view['job']['id']}/segments/start", {
            "phase": phase,
            "actor": f"slack:{row['user_id']}",
        }, operation="start_task")
    except TrackerRefused as refusal:
        # already_running is a second click on a card that is already going, and
        # already_processed is the same delivery arriving twice. Both mean the
        # timer is running, which is what the caller is about to render.
        if refusal.reason not in ("already_running", "already_processed"):
            raise


def stop_task(task_id):
    """Pause the running phase. Elapsed time is recomputed by LMSA from segments."""
    resolved = _row_for(task_id)
    if resolved is None:
        return
    view, row = resolved
    phase = row["current_phase"]
    if phase == "completed":
        return
    try:
        _call("POST", f"/jobs/{view['job']['id']}/segments/stop", {
            "phase": phase,
            "actor": f"slack:{row['user_id']}",
            "stopReason": "worker_action",
        }, operation="stop_task")
    except TrackerRefused as refusal:
        if refusal.reason not in ("not_running", "already_processed"):
            raise


def complete_task(task_id):
    """Finish the current phase, closing any timer still running on it."""
    resolved = _row_for(task_id)
    if resolved is None:
        return
    view, row = resolved
    phase = row["current_phase"]
    if phase == "completed":
        return
    try:
        _call("POST", f"/jobs/{view['job']['id']}/phases/complete", {
            "phase": phase,
            "actor": f"slack:{row['user_id']}",
        }, operation="complete_task")
    except TrackerRefused as refusal:
        if refusal.reason not in ("phase_already_complete", "already_processed"):
            raise


def move_to_border_phase(task_id, border_design, border_difficulty):
    """
    Store the border design and difficulty the modal collected.

    The move itself needs no write: the field phase was completed a moment
    earlier, and the card's phase follows from which phases are finished.
    """
    resolved = _row_for(task_id)
    if resolved is None:
        return
    view, row = resolved
    try:
        _call("POST", f"/jobs/{view['job']['id']}/phases/details", {
            "phase": "border_sheeting",
            "designName": border_design,
            "difficulty": border_difficulty,
            "actor": f"slack:{row['user_id']}",
        }, operation="move_to_border_phase")
    except TrackerRefused as refusal:
        if refusal.reason != "already_processed":
            raise


def move_to_packing_phase(task_id):
    """
    Packing collects nothing, so there is nothing to store.

    Completing the border phase is what moves the card on; this call is kept
    because app.py makes it, and it confirms the job is still there.
    """
    _row_for(task_id)


def save_notes_and_complete(task_id, general_notes, issues):
    """Save the closing notes and finish the job."""
    resolved = _row_for(task_id)
    if resolved is None:
        return
    view, row = resolved
    try:
        _call("POST", f"/jobs/{view['job']['id']}/complete", {
            "actor": f"slack:{row['user_id']}",
            "generalNotes": general_notes,
            "issuesEncountered": issues,
        }, operation="save_notes_and_complete")
    except TrackerRefused as refusal:
        if refusal.reason not in ("job_not_open", "already_processed"):
            raise


def get_phase_elapsed(task_id):
    """Seconds per phase, for the final summary."""
    resolved = _row_for(task_id)
    if resolved is None:
        return {}
    _, row = resolved
    return {
        "field_elapsed": row["field_elapsed"],
        "border_elapsed": row["border_elapsed"],
        "packing_elapsed": row["packing_elapsed"],
        "total_elapsed": row["total_elapsed"],
    }


def update_message_ts(task_id, dm_channel_id, message_ts):
    """
    Remember where the job's card is.

    Both halves are stored. The timestamp alone is not enough to find a message
    again: LMSA also needs the DM conversation, because the end-of-day sweep
    runs on a schedule with no click to read a channel from. Slack returns both
    in the same chat_postMessage response, so the caller already has them.
    """
    resolved = _row_for(task_id)
    if resolved is None:
        return
    view, row = resolved
    try:
        _call("POST", f"/jobs/{view['job']['id']}/card", {
            "dmChannelId": dm_channel_id,
            "cardMessageTs": message_ts,
            "actor": f"slack:{row['user_id']}",
        }, operation="update_message_ts")
    except TrackerRefused as refusal:
        if refusal.reason != "already_processed":
            raise


def delete_task(task_id):
    """
    Remove the job from the workflow.

    LMSA cancels rather than deletes, so the timings and the history of a job
    somebody spent the morning on survive. To the maker it is gone: the card is
    replaced and the job stops blocking a new one.
    """
    resolved = _row_for(task_id)
    if resolved is None:
        return
    view, row = resolved
    try:
        _call("POST", f"/jobs/{view['job']['id']}/cancel", {
            "actor": f"slack:{row['user_id']}",
            "reason": "deleted from the job card",
        }, operation="delete_task")
    except TrackerRefused as refusal:
        if refusal.reason not in ("job_not_open", "already_processed"):
            raise


def update_task(task_id, customer, invoice, task_desc, design, difficulty, due_date):
    """Apply the Edit modal's values."""
    resolved = _row_for(task_id)
    if resolved is None:
        return
    view, row = resolved
    iso_due, not_applicable = _text_to_due_date(due_date)
    try:
        _call("POST", f"/jobs/{view['job']['id']}/edit", {
            "customerName": customer,
            "invoiceNumber": invoice,
            "taskDescription": task_desc,
            "designName": design,
            "difficulty": difficulty,
            "dueDate": iso_due,
            "dueDateNotApplicable": not_applicable,
            "actor": f"slack:{row['user_id']}",
        }, operation="update_task")
    except TrackerRefused as refusal:
        if refusal.reason not in ("job_not_open", "already_processed"):
            raise


def format_elapsed(seconds):
    # Format elapsed seconds into a readable hours/minutes/seconds string.
    if seconds is None:
        seconds = 0
    hours = seconds // 3600
    minutes = (seconds % 3600) // 60
    secs = seconds % 60
    return f"{hours} h {minutes}m {secs}s"


def get_completed_tasks():
    """Completed jobs, oldest first, for the Excel export."""
    data = _call("GET", "/jobs/completed")
    rows = []
    for entry in (data or {}).get("jobs", []):
        rows.append(_row({"job": entry["job"], "phases": entry.get("phases") or []}, entry.get("timing")))
    return rows


if __name__ == "__main__":
    setup_database()
