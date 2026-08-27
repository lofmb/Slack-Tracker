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
import threading
import urllib.error
import urllib.request
from concurrent.futures import ThreadPoolExecutor

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


# The receiving thread's copy of the current delivery identity, kept past the
# middleware's exit so listener_executor() can read it when Bolt queues the
# handler. Thread-local, so two requests being received at once cannot see
# each other's identity.
_dispatch_identity = threading.local()


class slack_request:
    """
    Scope one delivery's identity to the handler that is running.

    Set on entry and reset on exit, including when the handler raises, so a
    later request can never inherit a key that was minted for an earlier one.

    Entry also leaves a copy of the identity in _dispatch_identity below. Bolt
    runs this middleware on the thread that received the request, but runs the
    handler itself on a worker thread, and it only queues the handler after
    this middleware has already exited — so by then the context variable is
    empty again. The copy is what listener_executor() picks up, on the same
    receiving thread, at the moment the handler is queued. It is overwritten
    by every request (with None when the payload has no identity), never
    trusted across requests, and the worker thread gets the identity through
    the executor's own set/reset, so nothing here weakens the reset guarantee.
    """

    def __init__(self, body):
        self._identity = delivery_identity(body)
        self._token = None

    def __enter__(self):
        self._token = _current_identity.set(self._identity)
        _dispatch_identity.value = self._identity
        return self

    def __exit__(self, *exc):
        _current_identity.reset(self._token)
        # The _dispatch_identity copy is deliberately left in place: the
        # executor reads it after this exit, still on the receiving thread.
        return False


class _IdentityPreservingExecutor(ThreadPoolExecutor):
    """
    The same thread pool Bolt would create for itself, plus one step: carry
    the delivery identity from the receiving thread onto the worker thread.

    submit() still runs on the receiving thread, so it can read the copy the
    middleware left in _dispatch_identity. The handler is wrapped so that the
    identity is placed into the context variable on the worker thread just
    before the handler runs, and always removed again afterwards — even when
    the handler raises — so a reused worker thread starts the next handler
    clean.
    """

    def submit(self, fn, *args, **kwargs):
        identity = getattr(_dispatch_identity, "value", None)

        def run_with_identity():
            token = _current_identity.set(identity)
            try:
                return fn(*args, **kwargs)
            finally:
                _current_identity.reset(token)

        return super().submit(run_with_identity)


def listener_executor():
    """The executor app.py hands to Bolt; five workers, same as Bolt's own."""
    return _IdentityPreservingExecutor(max_workers=5)


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
    Give app.py back the due date exactly as the person typed it.

    The due date box accepts anything, and always has - "01/09/26", "Friday",
    "ASAP", "next week", "01/09/26 (if paint arrives)". Whatever was typed is
    what shows on the card, what the Edit form is filled in with, and what goes
    in the Due Date column of the spreadsheet, so it has to come back word for
    word. LMSA stores that text as dueDateText.

    The date form is only used for a row saved before the text was kept.
    """
    text = job.get("dueDateText")
    if text is not None:
        return text
    if job.get("dueDateNotApplicable"):
        return "N/A"
    iso = job.get("dueDate")
    if not iso:
        return None
    parts = str(iso)[:10].split("-")
    if len(parts) != 3:
        return str(iso)
    return f"{parts[2]}/{parts[1]}/{parts[0][2:]}"


def _due_date_to_iso(text):
    """
    Work out a real date from the text, when the text happens to be one.

    This is only a bonus copy, kept beside the typed text for anything later
    that wants to sort or filter by date. It is never what gets shown - the
    typed text is - so failing to read a date here costs nothing and is the
    normal outcome for "Friday" or "ASAP". The box has never been validated and
    is not being validated now.
    """
    raw = (text or "").strip()
    if not raw or raw.upper() == "N/A":
        return None
    for separator in ("/", "-", "."):
        parts = raw.split(separator)
        if len(parts) == 3 and all(p.isdigit() for p in parts):
            day, month, year = parts
            if len(year) == 2:
                year = f"20{year}"
            try:
                if 1 <= int(day) <= 31 and 1 <= int(month) <= 12 and len(year) == 4:
                    return f"{year}-{int(month):02d}-{int(day):02d}"
            except ValueError:
                break
    return None


def _current_phase(job):
    """
    The phase the maker is on, exactly as LMSA stores it.

    Read, never derived. Deriving it as "the first phase that is not finished"
    looks equivalent and is not: completing a phase would advance it, and
    app.py completes the phase BEFORE asking where it is. Every branch then
    lands a step early — the Border modal never opens, its design and
    difficulty are never collected, and once all three phases are finished no
    branch matches at all, so Complete goes quiet and the job can never be
    finished. The cursor moves only when a modal is submitted, which is also
    what lets a cancelled modal be re-opened by pressing Complete again.
    """
    phase = job.get("currentPhase")
    if not phase:
        # A server without the cursor column is a mismatched deployment. Fail
        # loudly: the derived answer is wrong in a way nobody would notice.
        raise TrackerApiError(
            "tracker API returned no currentPhase — LMSA and the vendored "
            "tracker are not the same version"
        )
    return phase


def _legacy_status(job, phases):
    """Translate LMSA's job/phase state into the status vocabulary app.py reads."""
    if job.get("status") == "completed":
        return "completed"
    phase_name = _current_phase(job)
    if phase_name == "completed":
        return "completed"
    row = next((p for p in phases if p.get("phase") == phase_name), None)
    state = (row or {}).get("state")
    if state == "running":
        return "in_progress"
    if state == "paused":
        return "paused"
    if state == "complete":
        # The phase under the cursor is finished but the job is not: the maker
        # has pressed Complete and owes the modal that follows. The upstream
        # tracker says "completed" here, and app.py picks the card's buttons
        # from it — answering "created" would rebuild the card without a
        # Complete button and strand anyone who edits at that moment.
        return "completed"
    if state == "skipped":
        # A lane that did not happen is finished with, so the card reads the
        # same as a completed one. Unreachable in the normal flow, where the
        # cursor never rests on a skipped border, but a card rebuilt from a
        # stale cursor must not fall through to "created" and offer Start on a
        # lane that can never be timed.
        return "completed"
    return "created"


def _phase_of(phases, name, field):
    row = next((p for p in phases if p.get("phase") == name), None)
    return (row or {}).get(field)


def _packing_begun(phases, packing_seconds):
    """
    Has packing actually started?

    Asked of the lane state rather than the clock, because a packing lane can
    begin and be stopped again inside the same second and still plainly have
    begun. The seconds are kept as a second opinion only.
    """
    state = _phase_of(phases, "packing", "state")
    if state and state != "not_started":
        return True
    return int(packing_seconds or 0) > 0


def _jigs_of(view, phase_name):
    """
    The jig records LMSA holds for one phase, in the order they were recorded.

    A phase normally has one jig, sometimes two or three — a jig swapped
    mid-run, or two deliberately used together. LMSA keeps each as its own
    record so the earlier one survives; here they come back as small dicts the
    card and the Edit modal both read.
    """
    jigs = view.get("jigs") or []
    return [
        {"id": j["id"], "value": j["jigSizeText"]}
        for j in jigs
        if j.get("phase") == phase_name
    ]


def _jig_display(records):
    """One line of jig values for a card: '49.6', or '49.6 / 50' after a swap."""
    return " / ".join(r["value"] for r in records)


def _activity_seconds(timing, phase, activity):
    """
    Setup or sheeting time for one lane.

    A lane's total is these two added together, which is why the card can show
    "Setup 1h, sheeting 5h" without either number being invented.
    """
    per_phase = (timing or {}).get("perPhaseActivitySeconds") or {}
    return int(((per_phase.get(phase) or {}).get(activity) or 0))


def _cutting_seconds(timing, phase=None):
    """
    Cutting time, either for one lane or for the whole job.

    ALREADY INSIDE the lane time above, never added to it. A maker sheeting a
    field who spends twelve minutes cutting worked the field for the whole
    hour; the twelve minutes say how part of that hour was spent.
    """
    if phase is None:
        return int((timing or {}).get("cuttingSeconds") or 0)
    return int(((timing or {}).get("cuttingSecondsByPhase") or {}).get(phase) or 0)


def _row(view, timing):
    """Build the dictionary app.py indexes into, from the API's job view."""
    job = view["job"]
    phases = view.get("phases") or []
    seconds = (timing or {}).get("perPhaseSeconds") or {}
    field = int(seconds.get("field_sheeting", 0) or 0)
    border = int(seconds.get("border_sheeting", 0) or 0)
    packing = int(seconds.get("packing", 0) or 0)
    open_segment = view.get("openSegment") or None
    open_contained = view.get("openContained") or None
    last_segment = view.get("lastSegment") or None

    field_jig_records = _jigs_of(view, "field_sheeting")
    border_jig_records = _jigs_of(view, "border_sheeting")

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
        "field_jigs": _jig_display(field_jig_records),
        "field_jig_records": field_jig_records,
        "border_design": _phase_of(phases, "border_sheeting", "designName"),
        "border_difficulty": _phase_of(phases, "border_sheeting", "difficulty"),
        "border_elapsed": border,
        "border_jigs": _jig_display(border_jig_records),
        "border_jig_records": border_jig_records,
        # A skipped border and a border worked for no measurable time both sum
        # to zero seconds, so nothing else in this dict can tell them apart.
        # The lane state is carried explicitly, and every border rendering
        # keys off it.
        "border_skipped": _phase_of(phases, "border_sheeting", "state") == "skipped",
        # Where each lane stands: not_started, running, paused, complete, or
        # skipped. The card reads these to decide what the maker can still do -
        # a finished lane is not somewhere to switch to, and a lane declared
        # absent is not either.
        "phase_states": {
            name: _phase_of(phases, name, "state")
            for name in ("field_sheeting", "border_sheeting", "packing")
        },
        "packing_begun": _packing_begun(phases, packing),
        # Packing can be worked out of turn, as an interruption of the
        # sheeting, so the cards need to know two more things: is the packing
        # timer the one running right now, and has packing been finished for
        # good. Everything between those two — packing that has some time but
        # is not running — shows up through packing_elapsed.
        "packing_running": _phase_of(phases, "packing", "state") == "running",
        "packing_finished": _phase_of(phases, "packing", "state") == "complete",
        "packing_elapsed": packing,
        # WHAT THE MAKER IS DOING RIGHT NOW, read from the ledger rather than
        # worked out from lane states. Setup and sheeting share a lane, so the
        # lane cannot say which of them is going, and during a packing
        # interruption the lane that is accruing is not the one the job is on.
        "working_on": (
            {"phase": open_segment["phase"], "activity": open_segment["activity"]}
            if open_segment
            else None
        ),
        # Work being measured INSIDE that, with the main timer still running.
        "cutting_now": (
            {"parent_phase": open_contained["parentPhase"]} if open_contained else None
        ),
        # The last stretch of work, running or not — what a paused card's
        # Resume offers. Its lane may since have been finished or declared
        # absent, which the card checks before offering it.
        "last_work": (
            {"phase": last_segment["phase"], "activity": last_segment["activity"]}
            if last_segment
            else None
        ),
        # The two halves of each sheeting lane, and the cutting contained in it.
        "field_setup_elapsed": _activity_seconds(timing, "field_sheeting", "setup"),
        "field_production_elapsed": _activity_seconds(timing, "field_sheeting", "production"),
        "border_setup_elapsed": _activity_seconds(timing, "border_sheeting", "setup"),
        "border_production_elapsed": _activity_seconds(timing, "border_sheeting", "production"),
        "field_cutting_elapsed": _cutting_seconds(timing, "field_sheeting"),
        "border_cutting_elapsed": _cutting_seconds(timing, "border_sheeting"),
        "cutting_elapsed": _cutting_seconds(timing),
        "general_notes": job.get("generalNotes"),
        "issues_encountered": job.get("issuesEncountered"),
        "status": _legacy_status(job, phases),
        "current_phase": _current_phase(job),
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


def create_task(user_id, channel_id, customer_name, invoice_number, task_description, due_date, design, difficulty):
    """
    Create a job and return its number — the T-number shown on the card.

    Creating it also starts its setup timer, in the same moment. Submitting the
    intake form is the maker taking the job on, and the setup — fetching the
    material, reading the drawings, finding the jig — is the first real work.
    There is nothing left for a "Start" button to start.

    No jig is sent. A maker filling this form in normally does not know the jig
    yet; finding and testing it IS the setup, so the card asks for it at the
    point it can actually be answered.
    """
    payload = {
        "ownerSlackUserId": user_id,
        "customerName": customer_name,
        "invoiceNumber": invoice_number,
        "taskDescription": task_description,
        # The typed text is the due date; the read-back date is a bonus for
        # anything later that wants to sort by it. A blank box sends null,
        # which is a date nobody has been given yet.
        #
        # No dueDateNotApplicable, for the same reason update_task has never
        # sent one: there is no control to source it from. The intake form no
        # longer offers "No Set Date?", because a job with no deadline is not a
        # thing this workshop has - every job is done as soon as practicable.
        # Omitting the key leaves LMSA to default it to false for a new job and
        # leaves an existing row's own value alone.
        "dueDateText": due_date,
        "dueDate": _due_date_to_iso(due_date),
        "fieldDesignName": design,
        "fieldDifficulty": difficulty,
        "announceChannelId": channel_id,
        "actor": f"slack:{user_id}",
    }
    data = _call("POST", "/jobs", payload, operation="create_task")
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


def start_work(task_id, phase=None, activity="production"):
    """
    Move the maker onto a piece of work and start timing it.

    One call covers every way that happens: starting the sheeting after the
    setup, resuming after a pause, going to pack for a while, and coming back
    to the sheeting afterwards. Sent with "interrupting", which tells LMSA to
    close whatever is running in the same moment this opens — so there is never
    an instant with two timers, or none the maker did not ask for.

    It never finishes anything. The work being left is paused, with everything
    it has recorded intact.

    `phase` defaults to the lane the job is on, which is what Resume wants.
    Returns "started", or the reason it could not, so the handler can say so.
    A second click, or the same click delivered twice, reports "started": the
    timer is running, which is what the maker asked for.
    """
    resolved = _row_for(task_id)
    if resolved is None:
        return None
    view, row = resolved
    target = phase or row["current_phase"]
    if target == "completed":
        return "job_not_open"
    try:
        _call("POST", f"/jobs/{view['job']['id']}/segments/start", {
            "phase": target,
            "activity": activity,
            "actor": f"slack:{row['user_id']}",
            "interrupting": True,
        }, operation=f"start_work_{target}_{activity}")
    except TrackerRefused as refusal:
        if refusal.reason in ("already_running", "already_processed"):
            return "started"
        return refusal.reason
    return "started"


def start_packing(task_id):
    """
    Go and pack for a while, leaving the sheeting where it is.

    Kept as its own name because a card posted by an earlier version of the
    tracker still carries the button that calls it, and that card must keep
    working after a deployment. New cards say "Switch work" and go through
    start_work.
    """
    return start_work(task_id, "packing")


def stop_work(task_id):
    """
    Pause whatever the maker is doing.

    LMSA is not told which phase: the card's Pause means "stop what I am doing",
    and only the ledger knows what that is — during a packing interruption the
    running timer is the packing one while the job is still on its sheeting,
    and setup and sheeting share a lane. Anything being measured inside the
    work, such as cutting, is closed with it.

    Elapsed time is recomputed by LMSA from the segments.
    """
    resolved = _row_for(task_id)
    if resolved is None:
        return
    view, row = resolved
    if row["current_phase"] == "completed":
        return
    try:
        _call("POST", f"/jobs/{view['job']['id']}/segments/stop", {
            "actor": f"slack:{row['user_id']}",
            "stopReason": "worker_action",
        }, operation="stop_work")
    except TrackerRefused as refusal:
        if refusal.reason not in ("not_running", "already_processed"):
            raise


def start_cutting(task_id):
    """
    Start measuring cutting, WITHOUT stopping the sheeting.

    The maker goes downstairs, cuts tiles, comes back. They were working the
    field the whole time, so the field timer keeps running and this records how
    part of that time was spent. LMSA refuses it when there is no sheeting
    running to be inside.

    Returns "started" or the reason it could not.
    """
    resolved = _row_for(task_id)
    if resolved is None:
        return None
    view, row = resolved
    try:
        _call("POST", f"/jobs/{view['job']['id']}/contained/start", {
            "activity": "cutting",
            "actor": f"slack:{row['user_id']}",
        }, operation="start_cutting")
    except TrackerRefused as refusal:
        if refusal.reason in ("already_cutting", "already_processed"):
            return "started"
        return refusal.reason
    return "started"


def stop_cutting(task_id):
    """
    Stop measuring the cutting. The sheeting timer carries on, because the
    maker never stopped sheeting — they went and cut some tiles for a while.

    Returns "stopped" or the reason it could not.
    """
    resolved = _row_for(task_id)
    if resolved is None:
        return None
    view, row = resolved
    try:
        _call("POST", f"/jobs/{view['job']['id']}/contained/stop", {
            "actor": f"slack:{row['user_id']}",
        }, operation="stop_cutting")
    except TrackerRefused as refusal:
        if refusal.reason in ("not_cutting", "already_processed"):
            return "stopped"
        return refusal.reason
    return "stopped"


def complete_task(task_id):
    """
    Finish the current phase, closing any timer still running on it.

    Returns "completed", or the reason it could not. The one refusal a maker
    can genuinely cause is "another_phase_running" — pressing Complete Phase
    on a sheeting card while the packing timer is going. Completing the phase
    underneath a running timer would leave that timer stranded, so LMSA says
    no and the handler tells the maker to deal with the packing first.
    """
    resolved = _row_for(task_id)
    if resolved is None:
        return None
    view, row = resolved
    phase = row["current_phase"]
    if phase == "completed":
        return "completed"
    try:
        _call("POST", f"/jobs/{view['job']['id']}/phases/complete", {
            "phase": phase,
            "actor": f"slack:{row['user_id']}",
        }, operation="complete_task")
    except TrackerRefused as refusal:
        if refusal.reason in ("phase_already_complete", "already_processed"):
            return "completed"
        if refusal.reason == "another_phase_running":
            return refusal.reason
        raise
    return "completed"


def _advance_cursor(job_id, phase, actor, operation):
    """
    Move the job's workflow cursor onto the phase whose modal was just
    submitted. This is the only thing that advances it.
    """
    try:
        _call("POST", f"/jobs/{job_id}/phase-cursor", {
            "phase": phase,
            "actor": actor,
        }, operation=operation)
    except TrackerRefused as refusal:
        # cursor_regression means the cursor is already at or past this phase,
        # which is the state the caller wanted; job_not_open means the job
        # finished or was cancelled underneath. Neither is a fault.
        if refusal.reason not in ("already_processed", "cursor_regression", "job_not_open"):
            raise


def move_to_border_phase(task_id, border_design, border_difficulty, border_jig=None):
    """
    Record what the border modal collected, and move the cursor onto border.

    Two calls, because they are two different facts: what the maker typed, and
    where the maker now is. Each carries its own idempotency key, so a
    redelivered submission repeats neither. If the second call is the one that
    fails, the cursor stays on field: pressing Complete re-opens this same
    modal, and submitting it again finishes the move.
    """
    resolved = _row_for(task_id)
    if resolved is None:
        return
    view, row = resolved
    job_id = view["job"]["id"]
    actor = f"slack:{row['user_id']}"

    # Filling this form in IS the statement that there is a border after all.
    # A job previously marked "no border" is therefore put back first, as its
    # own audited step: taking the decision back and describing the border are
    # two different facts, and the history should show both. If the border
    # cannot be put back - packing has already started - the refusal is raised
    # rather than swallowed, because writing border details onto a lane that is
    # still recorded as skipped would be refused anyway, silently, and the
    # maker would be left believing the form saved.
    if row.get("border_skipped"):
        outcome = _post_border_skip_revert(job_id, actor, "move_to_border_phase_unskip")
        if outcome != "reverted":
            raise TrackerRefused(outcome)

    details = {
        "phase": "border_sheeting",
        "designName": border_design,
        "difficulty": border_difficulty,
        "actor": actor,
    }
    # The border modal's jig box is optional; sent only when something was
    # typed. LMSA records it as the border phase's first jig — re-submitting
    # this modal corrects that first value, it never stacks up copies.
    if border_jig:
        details["jigSizeText"] = border_jig
    try:
        _call("POST", f"/jobs/{job_id}/phases/details", details, operation="move_to_border_phase")
    except TrackerRefused as refusal:
        if refusal.reason != "already_processed":
            raise
    _advance_cursor(job_id, "border_sheeting", actor, "move_to_border_phase_cursor")


def move_to_packing_phase(task_id):
    """
    Move the cursor onto packing.

    The packing modal collects nothing, so where the maker now is IS the whole
    of what this submission records.

    Returns "moved", or the reason it did not. A maker with Slack open twice can
    take a "no border" back on one surface while this form is still sitting open
    on the other, and a form opened before that correction must not carry the
    job past a border nobody has decided about. LMSA refuses that; handing the
    reason back is what lets the handler say so, instead of the refusal
    disappearing into an unhandled error the maker never sees.
    """
    resolved = _row_for(task_id)
    if resolved is None:
        return None
    view, row = resolved
    try:
        _advance_cursor(
            view["job"]["id"], "packing", f"slack:{row['user_id']}", "move_to_packing_phase"
        )
    except TrackerRefused as refusal:
        return refusal.reason
    return "moved"


def skip_border_phase(task_id):
    """
    Record that this job has no border.

    One call, and deliberately not two: unlike move_to_border_phase this does
    NOT advance the cursor. The cursor moves when the packing modal that
    follows is submitted, exactly as it would have after the border modal. The
    two answers to the same question therefore have the same shape, and while
    the maker has not moved on the decision costs nothing to take back.

    phase_already_skipped is tolerated for the same reason already_processed
    is: a redelivered click must not show an error to a maker who did nothing
    wrong. Nothing else is swallowed. A refusal that means the skip did not
    happen has to reach the handler, because the handler is what tells the
    maker.
    """
    resolved = _row_for(task_id)
    if resolved is None:
        return
    view, row = resolved
    try:
        _call("POST", f"/jobs/{view['job']['id']}/phases/border-skip", {
            "actor": f"slack:{row['user_id']}",
        }, operation="skip_border_phase")
    except TrackerRefused as refusal:
        if refusal.reason not in ("phase_already_skipped", "already_processed"):
            raise


def _post_border_skip_revert(job_id, actor, operation):
    """
    Put a border that was marked "no border" back to undecided.

    Returns "reverted" or the refusal reason. Never swallows: both callers need
    to know, because a correction that quietly did nothing is worse than an
    error - the card would be rebuilt as though the border were back while the
    record still said skipped.
    """
    try:
        _call("POST", f"/jobs/{job_id}/phases/border-skip/revert", {
            "actor": actor,
        }, operation=operation)
    except TrackerRefused as refusal:
        if refusal.reason == "already_processed":
            return "reverted"
        return refusal.reason
    return "reverted"


def revert_border_skip(task_id):
    """
    Take back a "no border" chosen by mistake, and SAY what happened.

    This one hands its refusal back instead of swallowing it. A correction that
    quietly did nothing is worse than an error: the handler would rebuild a
    card claiming the border is back while the record still says skipped, and
    the maker would carry on believing it. The handler turns the reason into
    something readable.

    Returns "reverted", or the refusal reason, or None when the job could not
    be resolved at all. A replayed click reports "reverted", because the first
    delivery of it did exactly that.
    """
    resolved = _row_for(task_id)
    if resolved is None:
        return None
    view, row = resolved
    return _post_border_skip_revert(
        view["job"]["id"], f"slack:{row['user_id']}", "revert_border_skip"
    )


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


def add_jig(task_id, phase, jig_size):
    """
    Record another jig on the Field or Border phase — the Add Jig button.

    Always adds a new record. The jig already on the card genuinely happened,
    so it stays; the card then shows both, oldest first. A redelivered
    submission is absorbed rather than adding the same jig twice.
    """
    resolved = _row_for(task_id)
    if resolved is None:
        return
    view, row = resolved
    try:
        _call("POST", f"/jobs/{view['job']['id']}/jigs", {
            "phase": phase,
            "jigSizeText": jig_size,
            "actor": f"slack:{row['user_id']}",
        }, operation="add_jig")
    except TrackerRefused as refusal:
        if refusal.reason not in ("job_not_open", "already_processed"):
            raise


def correct_jig(task_id, jig_id, jig_size):
    """
    Fix ONE jig value that was mistyped — an Edit, not a new jig.

    LMSA keeps the old value in the audit trail, changes only the record named
    here, and quietly does nothing if the value has not actually changed. The
    operation carries the record's id so one Save that fixes two jigs sends
    two distinct requests, and neither swallows the other.
    """
    resolved = _row_for(task_id)
    if resolved is None:
        return
    view, row = resolved
    try:
        _call("POST", f"/jobs/{view['job']['id']}/jigs/{jig_id}", {
            "jigSizeText": jig_size,
            "actor": f"slack:{row['user_id']}",
        }, operation=f"correct_jig_{jig_id}")
    except TrackerRefused as refusal:
        if refusal.reason not in ("job_not_open", "already_processed"):
            raise


def update_task(task_id, customer, invoice, task_desc, design, difficulty, due_date):
    """Apply the Edit modal's values."""
    resolved = _row_for(task_id)
    if resolved is None:
        return
    view, row = resolved
    try:
        _call("POST", f"/jobs/{view['job']['id']}/edit", {
            "customerName": customer,
            "invoiceNumber": invoice,
            "taskDescription": task_desc,
            "designName": design,
            "difficulty": difficulty,
            # No dueDateNotApplicable here on purpose. The Edit form has no
            # No Set Date tick box, so there is nothing to send, and the old
            # version never wrote that column on an edit either. Leaving the
            # key out tells LMSA to keep whatever was chosen when the job was
            # started, rather than quietly clearing it.
            "dueDateText": due_date,
            "dueDate": _due_date_to_iso(due_date),
            "actor": f"slack:{row['user_id']}",
        }, operation="update_task")
    except TrackerRefused as refusal:
        if refusal.reason not in ("job_not_open", "already_processed"):
            raise


def format_elapsed(seconds):
    # Format elapsed seconds into a readable hours/minutes/seconds string.
    # Every unit always appears. The xlsx export uses this, where a column of
    # the same shape is easier to scan down than one that changes width.
    if seconds is None:
        seconds = 0
    hours = seconds // 3600
    minutes = (seconds % 3600) // 60
    secs = seconds % 60
    return f"{hours} h {minutes}m {secs}s"


def format_duration(seconds):
    """
    A duration as a person would say it: 50s, 3m 42s, 1h 3m 42s.

    Units that would lead with a zero are left out, because "0 h 3m 42s" makes
    a reader take in an hours figure that is not there before reaching the part
    that matters. Nothing below the largest unit is dropped, so a duration is
    still exact and two of them still sort by eye.

    This is the WORKSHOP-FACING format. The export keeps format_elapsed.
    """
    if seconds is None:
        seconds = 0
    seconds = int(seconds)
    hours = seconds // 3600
    minutes = (seconds % 3600) // 60
    secs = seconds % 60
    if hours:
        return f"{hours}h {minutes}m {secs}s"
    if minutes:
        return f"{minutes}m {secs}s"
    return f"{secs}s"


def get_completed_tasks():
    """
    Completed jobs, oldest first, for the Excel export.

    The whole entry is passed through, jigs included. Dropping them here left
    every jig column in the export empty while the same jigs showed correctly
    on the card, which reads as jobs that never had one.
    """
    data = _call("GET", "/jobs/completed")
    rows = []
    for entry in (data or {}).get("jobs", []):
        rows.append(_row({
            "job": entry["job"],
            "phases": entry.get("phases") or [],
            "jigs": entry.get("jigs") or [],
        }, entry.get("timing")))
    return rows


if __name__ == "__main__":
    setup_database()
