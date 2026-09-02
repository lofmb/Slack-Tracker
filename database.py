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

The file has three layers, in this order:

    request identity and thread glue     making one Slack delivery recognisable
    HTTP                                 reaching LMSA, and failing usefully
    shape translation                    LMSA's JSON back into app.py's rows
    the interface app.py uses            one function per workshop action

If you are looking for a workshop operation - start_work, complete_task,
add_jig - it is in the last of those, and the sections above it exist to make
that one safe to call.
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

# The job's three phases, in order. Nothing here reads this: it records the
# vocabulary the rest of the file and app.py both use, in one obvious place.
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


# --- request identity, and the thread glue that carries it -----------------
# One Slack delivery can drive several writes (creating a job also stores its
# card), so the key is the delivery plus the operation. A redelivery of the
# same click repeats action_ts and therefore repeats the key, which LMSA
# absorbs; a genuine later click carries a new action_ts and is treated as the
# new action it is. A context variable rather than a global, so two overlapping
# requests cannot read each other's identity.
#
# The threading machinery below is here for one reason: Bolt runs a handler on
# a worker thread, and it does so AFTER the middleware that recorded the
# identity has already returned. A context variable does not cross that
# boundary by itself, so the executor copies it over. Without that, the handler
# builds its key from nothing, the key is omitted, and a redelivered click
# writes twice.

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

    The stored text can be anything: the form refuses "Friday" and "ASAP" now,
    but rows saved before it did still hold exactly that, and one of them may be
    what a card is about to show. Whatever is stored is what shows on the card,
    what the Edit form is filled in with, and what goes in the Due Date column
    of the spreadsheet, so it has to come back word for word. Validation belongs
    to the form; this returns the record faithfully. LMSA stores that text as
    dueDateText.

    The date form is the fallback for a row whose text was never kept.
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


def _legacy_status(job, phases, open_segment=None):
    """
    Translate LMSA's job/phase state into the status vocabulary the history and
    the completed-job export were written against.

    NOT what the card reads. The card reads working_on, straight from the
    ledger, because one word cannot say which piece of which lane is accruing.
    This survives for the records already written in its vocabulary.
    """
    if job.get("status") == "completed":
        return "completed"
    # Anything being timed means the maker is working, including the job's own
    # opening setup - which has no lane row for the branch below to find.
    if open_segment:
        return "in_progress"
    phase_name = _current_phase(job)
    if phase_name == "completed":
        return "completed"
    cursor_part = job.get("currentPartId")
    row = next(
        (
            p
            for p in phases
            if p.get("phase") == phase_name and p.get("partId") == cursor_part
        ),
        None,
    )
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


def _phase_of(phases, name, field, part_id=None):
    """
    One lane's value.

    `part_id` says WHICH piece's lane. Passing None means the job's own lane -
    packing - which is the only one that belongs to no piece. A job drawn as
    three pieces has three borders, so asking for "the border" without saying
    which would return whichever row came back first.
    """
    row = next(
        (p for p in phases if p.get("phase") == name and p.get("partId") == part_id),
        None,
    )
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


def _jigs_of(view, phase_name, part_number=None):
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
        if j.get("phase") == phase_name and j.get("partNumber") == part_number
    ]


def _jig_display(records):
    """One line of jig values for a card: '49.6', or '49.6 / 50' after a swap."""
    return " / ".join(r["value"] for r in records)


def _part_timing(timing, part_number):
    """One piece's own figures out of the job's timing, or empty if it has none."""
    for part in (timing or {}).get("perPart") or []:
        if part.get("partNumber") == part_number:
            return part
    return {}


def _activity_seconds(timing, phase, activity, part_number=None):
    """
    Setup or sheeting time for one lane.

    A lane's total is these two added together, which is why the card can show
    "Setup 1h, sheeting 5h" without either number being invented.
    """
    source = _part_timing(timing, part_number) if part_number else (timing or {})
    per_phase = source.get("perPhaseActivitySeconds") or {}
    return int(((per_phase.get(phase) or {}).get(activity) or 0))


def _cutting_seconds(timing, phase=None, part_number=None):
    """
    Cutting time, either for one lane or for the whole job.

    ALREADY INSIDE the lane time above, never added to it. A maker sheeting a
    field who spends twelve minutes cutting worked the field for the whole
    hour; the twelve minutes say how part of that hour was spent.
    """
    if phase is None:
        return int((timing or {}).get("cuttingSeconds") or 0)
    source = _part_timing(timing, part_number) if part_number else (timing or {})
    return int((source.get("cuttingSecondsByPhase") or {}).get(phase) or 0)


def _lane(view, timing, part, phase):
    """
    One lane on one piece, as a small dict app.py reads.

    Everything about a Field or a Border is in here, and nothing about it is
    anywhere else. That is the whole correction: the old row had one
    `border_design`, so a job drawn as three pieces had two borders it could
    not describe.
    """
    phases = view.get("phases") or []
    part_id = part["id"]
    number = part["partNumber"]
    state = _phase_of(phases, phase, "state", part_id)
    jig_records = _jigs_of(view, phase, number)
    setup = _activity_seconds(timing, phase, "setup", number)
    production = _activity_seconds(timing, phase, "production", number)
    return {
        "phase": phase,
        "part": number,
        # A lane the diagram does not have is 'skipped' from the moment the job
        # is created - recorded as work that did not happen, never as work that
        # took no time. Both sum to zero seconds, so nothing else could tell
        # them apart.
        "present": state != "skipped",
        "state": state,
        "design": _phase_of(phases, phase, "designName", part_id),
        "difficulty": _phase_of(phases, phase, "difficulty", part_id),
        "setup_elapsed": setup,
        "production_elapsed": production,
        "elapsed": setup + production,
        # Already inside the figures above, never added to them.
        "cutting_elapsed": _cutting_seconds(timing, phase, number),
        "jigs": _jig_display(jig_records),
        "jig_records": jig_records,
    }


def _row(view, timing):
    """Build the dictionary app.py indexes into, from the API's job view."""
    job = view["job"]
    phases = view.get("phases") or []
    parts = view.get("parts") or []
    seconds = (timing or {}).get("perPhaseSeconds") or {}
    packing = int(seconds.get("packing", 0) or 0)
    open_segment = view.get("openSegment") or None
    open_contained = view.get("openContained") or None
    last_segment = view.get("lastSegment") or None

    # THE PIECES, in the order the maker reads the diagram. Each carries its
    # own field and border, whole. Every question about a lane is asked of one
    # of these, which is what makes "which one?" impossible to leave out.
    part_rows = [
        {
            "part": part["partNumber"],
            "field": _lane(view, timing, part, "field_sheeting"),
            "border": _lane(view, timing, part, "border_sheeting"),
        }
        for part in parts
    ]

    # ONE-PIECE COMPATIBILITY. A job drawn as a single piece is what every job
    # was before pieces could be named, and its Part 1 IS that job - so the
    # flat keys below are exactly true for it. They are read by the completed-
    # job export and the history it has already written, which must keep
    # working. New code asks `parts` and never these: on a job with three
    # pieces they describe the first one and say nothing about the others,
    # which is precisely the ambiguity the parts model exists to remove.
    first = part_rows[0] if part_rows else None
    field = (first or {}).get("field") or {}
    border = (first or {}).get("border") or {}

    return {
        "task_id": job["jobNumber"],
        "user_id": job["ownerSlackUserId"],
        "channel_id": job.get("announceChannelId"),
        "customer_name": job["customerName"],
        "invoice_number": job["invoiceNumber"],
        "task_description": job["taskDescription"],
        "due_date": _due_date_to_text(job),
        "is_na_due_date": 1 if job.get("dueDateNotApplicable") else 0,
        "parts": part_rows,
        "part_count": len(part_rows),
        # THE OPENING PREPARATION, which belongs to the job and not to a lane.
        # It used to be recorded against whichever lane came first, because
        # that was the only place to put it - which made a border-only job's
        # preparation look like field work.
        "job_setup_elapsed": int((timing or {}).get("jobSetupSeconds") or 0),
        # Single-piece compatibility, as above.
        "field_design": field.get("design"),
        "difficulty": field.get("difficulty"),
        "field_elapsed": field.get("elapsed") or 0,
        "field_jigs": field.get("jigs") or "",
        "field_jig_records": field.get("jig_records") or [],
        "field_skipped": not field.get("present", True),
        "border_design": border.get("design"),
        "border_difficulty": border.get("difficulty"),
        "border_elapsed": border.get("elapsed") or 0,
        "border_jigs": border.get("jigs") or "",
        "border_jig_records": border.get("jig_records") or [],
        "border_skipped": not border.get("present", True),
        "field_setup_elapsed": field.get("setup_elapsed") or 0,
        "field_production_elapsed": field.get("production_elapsed") or 0,
        "border_setup_elapsed": border.get("setup_elapsed") or 0,
        "border_production_elapsed": border.get("production_elapsed") or 0,
        "field_cutting_elapsed": field.get("cutting_elapsed") or 0,
        "border_cutting_elapsed": border.get("cutting_elapsed") or 0,
        # Packing is the job's: the maker packs the finished job, not each
        # piece separately, so it stays flat and is not inside `parts`.
        "packing_begun": _packing_begun(phases, packing),
        "packing_state": _phase_of(phases, "packing", "state"),
        "packing_running": _phase_of(phases, "packing", "state") == "running",
        "packing_finished": _phase_of(phases, "packing", "state") == "complete",
        "packing_elapsed": packing,
        "cutting_elapsed": _cutting_seconds(timing),
        # WHAT THE MAKER IS DOING RIGHT NOW, read from the ledger rather than
        # worked out from lane states. It names the piece as well as the lane,
        # because "sheeting the border" stopped identifying the work the moment
        # a job could have three of them. `part` is None for the two job-level
        # phases: the opening setup and packing.
        "working_on": (
            {
                "phase": open_segment["phase"],
                "part": open_segment.get("partNumber"),
                "activity": open_segment["activity"],
            }
            if open_segment
            else None
        ),
        # Work being measured INSIDE that, with the main timer still running.
        # Its piece comes from the segment it is inside, never from the caller.
        "cutting_now": (
            {
                "parent_phase": open_contained["parentPhase"],
                "part": open_contained.get("parentPartNumber"),
            }
            if open_contained
            else None
        ),
        # The last stretch of work, running or not - what a paused card's
        # Resume offers. Its lane may since have been finished or declared
        # absent, which the card checks before offering it.
        "last_work": (
            {
                "phase": last_segment["phase"],
                "part": last_segment.get("partNumber"),
                "activity": last_segment["activity"],
            }
            if last_segment
            else None
        ),
        "general_notes": job.get("generalNotes"),
        "issues_encountered": job.get("issuesEncountered"),
        "status": _legacy_status(job, phases, open_segment),
        "current_phase": _current_phase(job),
        "current_part": _current_part(job, parts),
        "created_at": _iso_to_sqlite_datetime(job.get("createdAt")),
        "message_ts": job.get("cardMessageTs"),
        "dm_channel_id": job.get("dmChannelId"),
        "total_elapsed": int((timing or {}).get("totalSeconds") or 0),
    }


def _current_part(job, parts):
    """The piece the cursor is on, as a number. None once the job is done."""
    part_id = job.get("currentPartId")
    if not part_id:
        return None
    for part in parts:
        if part.get("id") == part_id:
            return part["partNumber"]
    return None


def _lane_by_number(row, part_number, which):
    """
    One lane out of a row, by the piece's number.

    `which` is "field" or "border". Returns an empty dict when the piece or the
    lane is not there, so a caller reading a job that changed underneath gets
    an absent lane rather than an exception.
    """
    for part in row.get("parts") or []:
        if part.get("part") == part_number:
            return part.get(which) or {}
    return {}


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
        if refusal.reason == "job_not_found":
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
    Confirm the tracker can reach LMSA. It creates nothing.

    Schema is applied by hand on the LMSA side and this process has no way to
    change it — no driver, no credentials, no privileges. Removing the ability
    is stronger than a rule saying not to, and it means a tracker restart can
    never migrate a production database as a side effect of booting.
    """
    health = _call("GET", "/health")
    if not isinstance(health, dict) or not health.get("ready"):
        raise TrackerApiError("tracker API did not report ready")
    print("Tracker API ready.")


def create_task(user_id, channel_id, customer_name, invoice_number, task_description, due_date,
                design=None, difficulty=None, border_design=None, border_difficulty=None,
                field_present=None, parts=None):
    """
    Create a job and return its number — the T-number shown on the card.

    Creating it also starts THE JOB'S setup timer, in the same moment.
    Submitting the intake form is the maker taking the job on, and the setup —
    reading the drawings, checking what was supplied, fetching the material —
    is the first real work. There is nothing left for a "Start" button to
    start. That setup is the JOB's: it is not charged to a lane, and on a job
    drawn as several pieces it is not charged to the first one.

    `parts` is the shape of the diagram: one entry per piece, each saying
    whether it has a field, a border, or both. That is ALL the intake form
    establishes — no designs, no difficulty, no jig, no cutting. Those are
    asked for when the maker first enters that piece's lane, which is the
    moment they are looking at it.

    Omitting `parts` describes a single-piece job through the flat arguments,
    which is what every caller did before pieces could be named. It builds one
    real part like any other.
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
        # No dueDateNotApplicable: there is no control to source one from. A job
        # with no deadline is not a thing this workshop has - every job is done
        # as soon as practicable - so the intake form asks for a date or nothing.
        # Omitting the key leaves LMSA to default it to false for a new job and
        # leaves an existing row's own value alone.
        "dueDateText": due_date,
        "dueDate": _due_date_to_iso(due_date),
        # The SHAPE is stated, never inferred from whether a design was typed.
        # A caller that does not say has a field, which is every job the
        # tracker made before a border could stand on its own.
        "fieldPresent": True if field_present is None else bool(field_present),
        "fieldDesignName": design,
        "fieldDifficulty": difficulty,
        "borderDesignName": border_design,
        "borderDifficulty": border_difficulty,
        "announceChannelId": channel_id,
        "actor": f"slack:{user_id}",
    }
    if parts:
        payload["parts"] = [
            {
                "fieldPresent": bool(part.get("field")),
                "borderPresent": bool(part.get("border")),
            }
            for part in parts
        ]
    data = _call("POST", "/jobs", payload, operation="create_task")
    return data["job"]["jobNumber"]


def get_active_task(user_id):
    """
    The job this maker is TIMING right now, if any.

    Not "the job they have open": a maker may hold several unfinished jobs,
    paused, and each keeps its own card. Only one may be accruing time, and
    this is that one. It is what /track asks before opening the intake form,
    and what a refusal names when a press on one job is turned down because
    another is running.
    """
    try:
        view = _call("GET", f"/jobs/active/slack:{user_id}")
    except TrackerRefused as refusal:
        if refusal.reason == "no_active_job":
            return None
        raise
    return _row(view, _timing(view["job"]["id"]))


def get_open_tasks(user_id):
    """
    Every unfinished job this maker holds, newest first — timing or paused.

    A paused job is still theirs: it stays here with everything it recorded
    until it is finished or cancelled, and its card in their DM is the way
    back to it.
    """
    data = _call("GET", f"/jobs/open/{user_id}") or {}
    return [_row(view, _timing(view["job"]["id"])) for view in data.get("jobs") or []]


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


def start_work(task_id, phase=None, activity="production", part=None):
    """
    Move the maker onto a piece of work and start timing it.

    One call covers every way that happens: starting the sheeting after the
    setup, resuming after a pause, going to pack for a while, and coming back
    to the sheeting afterwards. Sent with "interrupting", which tells LMSA to
    close whatever is running in the same moment this opens — so there is never
    an instant with two timers, or none the maker did not ask for.

    It never finishes anything. The work being left is paused, with everything
    it has recorded intact.

    `phase` defaults to the lane the job is on, and `part` to the piece the
    cursor is on, which together are what Resume wants. A lane needs a piece;
    the two job-level phases - the opening setup and packing - take none.

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
    # A lane always belongs to a piece. Defaulting to the cursor's piece is
    # what makes Resume mean "carry on where I was" rather than "carry on
    # somewhere on this job".
    if target in ("field_sheeting", "border_sheeting"):
        target_part = part if part is not None else row.get("current_part")
    else:
        target_part = None
    try:
        _call("POST", f"/jobs/{view['job']['id']}/segments/start", {
            "phase": target,
            "partNumber": target_part,
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
        # stopReason says a person chose to stop, as opposed to the segment
        # being closed by something else - the way a cutting record is closed
        # as "parent_stopped" when the work it sat inside ended. The ledger
        # keeps the two apart so a report can tell deliberate pauses from
        # consequences.
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


def complete_task(task_id, phase=None, part=None):
    """
    Finish one lane, closing any timer still running on it.

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
    target = phase or row["current_phase"]
    if target == "completed":
        return "completed"
    if target in ("field_sheeting", "border_sheeting"):
        target_part = part if part is not None else row.get("current_part")
    else:
        target_part = None
    try:
        _call("POST", f"/jobs/{view['job']['id']}/phases/complete", {
            "phase": target,
            "partNumber": target_part,
            "actor": f"slack:{row['user_id']}",
        }, operation="complete_task")
    except TrackerRefused as refusal:
        if refusal.reason in ("phase_already_complete", "already_processed"):
            return "completed"
        if refusal.reason == "another_phase_running":
            return refusal.reason
        raise
    return "completed"


def _advance_cursor(job_id, phase, actor, operation, part=None):
    """
    Move the job's workflow cursor onto the phase whose modal was just
    submitted. This is the only thing that advances it.
    """
    try:
        _call("POST", f"/jobs/{job_id}/phase-cursor", {
            "phase": phase,
            "partNumber": part,
            "actor": actor,
        }, operation=operation)
    except TrackerRefused as refusal:
        # cursor_regression means the cursor is already at or past this phase,
        # which is the state the caller wanted; job_not_open means the job
        # finished or was cancelled underneath. Neither is a fault.
        if refusal.reason not in ("already_processed", "cursor_regression", "job_not_open"):
            raise


def set_lane_details(task_id, phase, design, difficulty, jig=None, part=None):
    """
    Record what one lane on one piece IS - its design, its difficulty, and the
    jig when the maker already knows it.

    Asked on first entry to that lane, which is the moment the maker is looking
    at that piece of the diagram. It used to be asked of the border only, in a
    form reached by finishing the field; every lane is described the same way
    now, and the piece is named because a job drawn as three pieces has three
    borders and they are not the same border.

    Records the details and NOTHING else. It does not start the work and does
    not move the job's cursor: the caller starts what the maker pressed for,
    which is a separate fact with its own audit.

    A lane previously recorded as absent is put back first, as its own audited
    step - describing work that is on the record as not happening would be
    refused anyway, silently, and the maker would be left believing the form
    saved.
    """
    resolved = _row_for(task_id)
    if resolved is None:
        return
    view, row = resolved
    job_id = view["job"]["id"]
    actor = f"slack:{row['user_id']}"
    target_part = part if part is not None else row.get("current_part")

    which = "field" if phase == "field_sheeting" else "border"
    lane = _lane_by_number(row, target_part, which)
    if lane.get("present") is False and phase == "border_sheeting":
        outcome = _post_border_skip_revert(job_id, actor, "set_lane_details_unskip", target_part)
        if outcome != "reverted":
            raise TrackerRefused(outcome)

    details = {
        "phase": phase,
        "partNumber": target_part,
        "designName": design,
        "difficulty": difficulty,
        "actor": actor,
    }
    # The jig box is optional; sent only when something was typed. LMSA records
    # it as that lane's first jig - re-submitting corrects that first value, it
    # never stacks up copies.
    if jig:
        details["jigSizeText"] = jig
    try:
        _call("POST", f"/jobs/{job_id}/phases/details", details, operation="set_lane_details")
    except TrackerRefused as refusal:
        if refusal.reason != "already_processed":
            raise


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


def skip_border_phase(task_id, part=None):
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
            "partNumber": part if part is not None else row.get("current_part"),
            "actor": f"slack:{row['user_id']}",
        }, operation="skip_border_phase")
    except TrackerRefused as refusal:
        if refusal.reason not in ("phase_already_skipped", "already_processed"):
            raise


def _post_border_skip_revert(job_id, actor, operation, part=None):
    """
    Put a border that was marked "no border" back to undecided.

    Returns "reverted" or the refusal reason. Never swallows: both callers need
    to know, because a correction that quietly did nothing is worse than an
    error - the card would be rebuilt as though the border were back while the
    record still said skipped.
    """
    try:
        _call("POST", f"/jobs/{job_id}/phases/border-skip/revert", {
            "partNumber": part,
            "actor": actor,
        }, operation=operation)
    except TrackerRefused as refusal:
        if refusal.reason == "already_processed":
            return "reverted"
        return refusal.reason
    return "reverted"


def revert_border_skip(task_id, part=None):
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
        view["job"]["id"],
        f"slack:{row['user_id']}",
        "revert_border_skip",
        part if part is not None else row.get("current_part"),
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


def add_jig(task_id, phase, jig_size, part=None):
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
            "partNumber": part if part is not None else row.get("current_part"),
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


def update_task(task_id, customer, invoice, task_desc, design, difficulty, due_date, part=None):
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
            # The field design belongs to one piece's lane, so Edit says
            # which. Part 1 by default, which for a job drawn as a single
            # piece is the only answer there is.
            "partNumber": part if part is not None else 1,
            "designName": design,
            "difficulty": difficulty,
            # No dueDateNotApplicable here on purpose. The Edit form has no
            # No Set Date tick box, so there is nothing to send. Omitting the
            # key tells LMSA to keep whatever the row already carries; sending
            # it as null would clear it, silently erasing a historical choice
            # the maker never touched.
            "dueDateText": due_date,
            "dueDate": _due_date_to_iso(due_date),
            "actor": f"slack:{row['user_id']}",
        }, operation="update_task")
    except TrackerRefused as refusal:
        if refusal.reason not in ("job_not_open", "already_processed"):
            raise


def format_elapsed(seconds):
    """
    Elapsed seconds as hours, minutes and seconds, with every unit always
    shown. The xlsx export uses this: a column of the same shape is easier to
    scan down than one that changes width row to row.
    """
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
            # The pieces come through too, so the export reads a multi-piece
            # job as the several pieces it was rather than as its first one.
            "parts": entry.get("parts") or [],
            "phases": entry.get("phases") or [],
            "jigs": entry.get("jigs") or [],
        }, entry.get("timing")))
    return rows


if __name__ == "__main__":
    setup_database()
